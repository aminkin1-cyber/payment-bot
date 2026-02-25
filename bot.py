"""
Payment Tracker Bot v3
- Full context memory (context.txt)
- Confirmation before writing to Excel
- Balance reconciliation
- Context view/edit via Telegram
"""
import os, json, logging
from datetime import datetime, time
from pathlib import Path
import httpx
from openpyxl import load_workbook
import io
try:
    from pypdf import PdfReader
    HAS_PDF = True
except ImportError:
    HAS_PDF = False
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from telegram import Update, Bot, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (Application, MessageHandler, CommandHandler,
                           CallbackQueryHandler, filters, ContextTypes)

BOT_TOKEN     = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_KEY"]
MY_CHAT_ID    = int(os.environ["MY_CHAT_ID"])
MORNING_HOUR  = int(os.environ.get("MORNING_HOUR", "9"))

DATA_FILE    = Path("data/messages.json")
EXCEL_FILE   = Path("Agent_Model_v2.xlsx")
CONTEXT_FILE = Path("data/context.txt")
PENDING_FILE = Path("data/pending_update.json")  # stores parsed data awaiting confirmation

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ── Styles ────────────────────────────────────────────────────────────────────
WHITE  = "FFFFFF"; YELLOW = "FFF2CC"; GREEN  = "E2EFDA"
RED    = "FCE4D6"; ORANGE = "FDEBD0"; LIGHT  = "D6E4F0"; LGRAY  = "F2F2F2"
thin   = Side(style="thin", color="BFBFBF")
def B(): return Border(top=thin, bottom=thin, left=thin, right=thin)
TYPE_BG = {"Deposit": GREEN, "Payment": WHITE, "Cash Out": ORANGE,
           "Cash In": LIGHT, "❓ Unknown": RED}
STAT_BG = {"✅ Paid": GREEN, "⏳ Pending": YELLOW,
           "⚠ Partial/Check": ORANGE, "❓ Clarify": RED}

def sc(cell, bg=WHITE, bold=False, sz=9, fc="000000", num=None,
       align="left", wrap=False):
    cell.font      = Font(name="Arial", bold=bold, size=sz, color=fc)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border    = B()
    if num: cell.number_format = num

# ── Context ───────────────────────────────────────────────────────────────────
def load_context() -> str:
    CONTEXT_FILE.parent.mkdir(exist_ok=True)
    return CONTEXT_FILE.read_text(encoding="utf-8") if CONTEXT_FILE.exists() else ""

def save_context(text: str):
    CONTEXT_FILE.parent.mkdir(exist_ok=True)
    CONTEXT_FILE.write_text(text, encoding="utf-8")

def update_context_after_update(new_info: str):
    """Append new findings to context file."""
    ctx = load_context()
    ts  = datetime.now().strftime("%d.%m.%Y %H:%M")
    ctx += f"\n\n--- ОБНОВЛЕНИЕ {ts} ---\n{new_info}"
    save_context(ctx)


CHAT_HISTORY_FILE = Path("data/chat_history.json")
MAX_HISTORY = 8  # keep last 8 exchanges

def load_chat_history():
    CHAT_HISTORY_FILE.parent.mkdir(exist_ok=True)
    return json.loads(CHAT_HISTORY_FILE.read_text(encoding="utf-8")) if CHAT_HISTORY_FILE.exists() else []

def save_chat_history(history: list):
    # Keep only last MAX_HISTORY exchanges
    if len(history) > MAX_HISTORY * 2:
        history = history[-(MAX_HISTORY * 2):]
    CHAT_HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")

def clear_chat_history():
    if CHAT_HISTORY_FILE.exists(): CHAT_HISTORY_FILE.unlink()

def get_excel_summary() -> str:
    """Get compact Excel state for chat context."""
    lines = []
    bal = get_balance_from_excel()
    if bal:
        lines.append(f"Баланс агента: ${bal[0]:,.2f} USD (на {bal[1]})")
    pending, usd_total, tbc_count = get_pending_invoices()
    if pending:
        lines.append(f"Pending инвойсов: {len(pending)}")
        lines.extend(pending[:5])
        if len(pending) > 5:
            lines.append(f"  ...и ещё {len(pending)-5}")
        tbc_note = f" + {tbc_count} с суммой TBC" if tbc_count else ""
        lines.append(f"  Итого к оплате: ~${usd_total:,.0f} USD{tbc_note}")
    unknown = get_unknown_transactions()
    if unknown:
        lines.append(f"Неизвестных транзакций: {len(unknown)}")
    queue = load_messages()
    if queue:
        lines.append(f"Накоплено сообщений от агента: {len(queue)} (для /update)")
    return "\n".join(lines) if lines else "Excel не найден"

# ── Message store ─────────────────────────────────────────────────────────────
def load_messages():
    DATA_FILE.parent.mkdir(exist_ok=True)
    return json.loads(DATA_FILE.read_text(encoding="utf-8")) if DATA_FILE.exists() else []

def save_message(d):
    msgs = load_messages(); msgs.append(d)
    DATA_FILE.write_text(json.dumps(msgs, ensure_ascii=False, indent=2), encoding="utf-8")

def clear_messages():
    DATA_FILE.write_text("[]", encoding="utf-8")

def _fmt(msgs):
    lines = []
    for m in msgs:
        line = f"[{m['date']}] {m.get('sender','?')}:"
        if m.get("text"): line += f" {m['text']}"
        if m.get("file"): line += f" [файл: {m['file']}]"
        if m.get("pdf_content"):
            line += f"\n  [СОДЕРЖИМОЕ PDF {m['file']}]:\n  {m['pdf_content'][:2000]}"
        lines.append(line)
    return "\n".join(lines)

# ── Pending confirmation store ────────────────────────────────────────────────
def save_pending(data: dict):
    PENDING_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def load_pending() -> dict:
    return json.loads(PENDING_FILE.read_text(encoding="utf-8")) if PENDING_FILE.exists() else {}

def clear_pending():
    if PENDING_FILE.exists(): PENDING_FILE.unlink()

# ── Excel read ────────────────────────────────────────────────────────────────
def get_balance_from_excel():
    if not EXCEL_FILE.exists(): return None
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Transactions"]
        last_bal = last_date = None
        for row in ws.iter_rows(min_row=5, max_col=11, values_only=True):
            if row[10] is not None and isinstance(row[10], (int, float)):
                last_bal  = row[10]
                last_date = row[0]
        return (last_bal, last_date) if last_bal else None
    except Exception as e:
        log.error(f"Excel balance: {e}"); return None

def get_pending_invoices():
    """Returns (lines, usd_total, tbc_count) for all non-paid invoices."""
    if not EXCEL_FILE.exists(): return [], 0.0, 0
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Invoices"]
        out = []
        usd_total = 0.0
        tbc_count = 0
        for row in ws.iter_rows(min_row=5, max_col=10, values_only=True):
            if row[6] and row[6] != "✅ Paid" and row[0]:
                amt     = f"{row[4]:,.2f}" if isinstance(row[4], (int,float)) else str(row[4] or "TBC")
                usd_val = row[5]  # F = USD Equiv (data_only=True reads calculated value)
                if isinstance(usd_val, (int, float)):
                    usd_str = f" ≈ ${usd_val:,.0f}"
                    usd_total += usd_val
                else:
                    usd_str = " (USD TBC)"
                    tbc_count += 1
                out.append(f"- {row[2] or '?'}: {amt} {row[3] or ''}{usd_str}")
        return out, usd_total, tbc_count
    except Exception as e:
        log.error(f"Excel pending: {e}"); return [], 0.0, 0


def get_recent_unconfirmed(days=14):
    """Get recent Cash In / Deposits that might not be confirmed by agent yet."""
    if not EXCEL_FILE.exists(): return []
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Transactions"]
        items = []
        for row in ws.iter_rows(min_row=5, max_col=12, values_only=True):
            if row[1] in ("Cash In", "Deposit") and row[0]:
                note = str(row[11] or "").lower()
                # Flag as potentially unconfirmed if notes say so or if recent
                if "unconfirmed" in note or "follow up" in note or "not confirm" in note:
                    amt = f"{row[5]:,.2f}" if isinstance(row[5],(int,float)) else str(row[5])
                    items.append(f"- {row[0]}: {row[2] or ''} | {amt} {row[4] or ''} | ПРИМЕЧАНИЕ: {row[11]}")
        return items
    except Exception as e:
        log.error(f"get_recent_unconfirmed error: {e}"); return []


def get_existing_invoices_list():
    """Return list of existing invoice IDs and payees for dedup check."""
    if not EXCEL_FILE.exists(): return ""
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        wi = wb["Invoices"]
        lines = []
        for row in wi.iter_rows(min_row=5, values_only=True):
            if row[1] or row[2]:
                lines.append(f"inv={row[1] or '?'} | payee={row[2] or '?'} | ccy={row[3]} | amt={row[4]} | status={row[6]}")
        return "\n".join(lines)
    except Exception as e:
        log.error(f"get_existing_invoices: {e}"); return ""

def get_unknown_transactions():
    if not EXCEL_FILE.exists(): return []
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Transactions"]
        out = []
        for row in ws.iter_rows(min_row=5, max_col=12, values_only=True):
            if row[1] and "Unknown" in str(row[1]):
                amt = f"{row[5]:,.2f}" if isinstance(row[5], (int,float)) else str(row[5] or "?")
                out.append(f"- {row[0]}: {row[2] or '?'} | {amt} {row[4] or ''}")
        return out
    except Exception as e:
        log.error(f"Excel unknown: {e}"); return []

# ── Excel write ───────────────────────────────────────────────────────────────
def find_last_row(ws, start=5):
    last = start - 1
    for row in ws.iter_rows(min_row=start, max_col=1):
        if row[0].value is not None:
            last = row[0].row
    return last

def _get_fx(ws_parent, ccy):
    """Lookup FX rate from Settings sheet."""
    try:
        for row in ws_parent["Settings"].iter_rows(min_row=7, max_row=20, values_only=True):
            if str(row[0]) == ccy and row[1]:
                return float(row[1])
    except Exception: pass
    return 1.0

def _get_comm(tp, ccy):
    """Get commission rate."""
    if ccy == "RUB": return 0.004
    return {"Deposit":0.0,"Cash In":0.0,"Payment":0.005,
            "Cash Out":0.005,"❓ Unknown":0.005}.get(tp, 0.005)

def _prev_balance(ws, r):
    """Get last computed balance before row r."""
    last = 0.0
    for row in ws.iter_rows(min_row=5, max_row=r-1, max_col=11, values_only=True):
        if row[10] is not None and isinstance(row[10], (int,float)):
            last = float(row[10])
    return last

def apply_tx_row(ws, r, tx):
    tp  = tx.get("type", "Payment")
    bg  = TYPE_BG.get(tp, WHITE)
    ccy = tx.get("ccy","USD")

    # Compute values
    try: amt = float(tx.get("amount") or 0)
    except: amt = 0.0

    fx = float(tx.get("fx_rate")) if tx.get("fx_rate") else _get_fx(ws.parent, ccy)
    comm = float(tx.get("comm")) if tx.get("comm") else _get_comm(tp, ccy)
    gross = round(amt / fx, 2) if fx else amt
    net   = round(gross, 2) if tp in ("Deposit","Cash In") else round(-(gross/max(1-comm,0.0001)),2)
    bal   = round(_prev_balance(ws, r) + net, 2)

    # Write base columns A-F, L
    for col_i, val in enumerate([
        tx.get("date",""), tp, tx.get("description",""), tx.get("payee",""),
        ccy, amt, None, None, None, None, None, tx.get("notes","")
    ], 1):
        if col_i in (7,8,9,10,11): continue
        c = ws.cell(r, col_i, val if val is not None else "")
        sc(c, bg=bg, wrap=(col_i in (3,12)), sz=9)

    # G: FX rate
    sc(ws.cell(r, 7, round(fx,5)), bg=YELLOW, fc="0000CC", num="0.00000")
    # H: Gross USD
    sc(ws.cell(r, 8, gross), bg=YELLOW, num="#,##0.00")
    # I: Comm %
    sc(ws.cell(r, 9, comm), bg=YELLOW, fc="0000CC", num="0.0%")
    # J: Net USD
    ws.cell(r,10).value = (f'=IF(H{r}="","",IF(OR(B{r}="Deposit",B{r}="Cash In"),'
                           f'H{r},-(H{r}/MAX(1-I{r},0.0001))))')
    ws.cell(r,10).number_format = '#,##0.00'; sc(ws.cell(r,10), bg=YELLOW)
    # K: Balance — compute directly so bot can read it back
    # Get previous balance
    prev_bal = 0.0
    try:
        for pr in ws.iter_rows(min_row=5, max_row=r-1, max_col=11, values_only=True):
            if pr[10] is not None and isinstance(pr[10], (int, float)):
                prev_bal = float(pr[10])
    except Exception:
        pass
    # Get starting balance from Settings if no previous
    if prev_bal == 0.0:
        try:
            wb2 = ws.parent
            start = wb2["Settings"].cell(4, 3).value
            if start and isinstance(start, (int, float)):
                prev_bal = float(start)
        except Exception:
            pass
    # Compute fx_rate value
    try:
        fx_val = float(tx.get("fx_rate") or 1.0)
        if not tx.get("fx_rate"):
            # lookup from Settings
            try:
                ccy = tx.get("ccy","")
                for srow in ws.parent["Settings"].iter_rows(min_row=7, max_row=16, values_only=True):
                    if srow[0] == ccy:
                        fx_val = float(srow[1]); break
            except Exception:
                fx_val = 1.0
    except Exception:
        fx_val = 1.0
    try:
        amt = float(tx.get("amount") or 0)
        gross = amt / fx_val if fx_val else amt
        comm_val = float(tx.get("comm") or 0)
        if not tx.get("comm"):
            try:
                tp2 = tx.get("type","Payment")
                comm_map = {"Deposit":0,"Cash In":0,"Payment":0.005,"Cash Out":0.005,"❓ Unknown":0.005}
                comm_val = comm_map.get(tp2, 0.005)
            except Exception:
                comm_val = 0.005
        tp2 = tx.get("type","Payment")
        if tp2 in ("Deposit","Cash In"):
            net = gross
        else:
            net = -(gross / max(1 - comm_val, 0.0001))
        new_bal = prev_bal + net
    except Exception as e:
        new_bal = prev_bal

    ws.cell(r,11).value = round(new_bal, 2)
    ws.cell(r,11).number_format = '#,##0.00'; sc(ws.cell(r,11), bg=YELLOW, bold=True, fc="1F3864")
    ws.row_dimensions[r].height = 28

def apply_inv_update(ws, upd):
    inv_no = str(upd.get("invoice_no","")).strip().lower()
    status = upd.get("new_status","✅ Paid")
    bg     = STAT_BG.get(status, YELLOW)
    for row in ws.iter_rows(min_row=5, max_col=10):
        if inv_no and inv_no in str(row[1].value or "").strip().lower():
            row[6].value = status; sc(row[6], bg=bg, bold=True, align="center")
            row[7].value = upd.get("date_paid",""); sc(row[7], bg=bg)
            if upd.get("ref"): row[8].value = upd["ref"]; sc(row[8], bg=bg, sz=8)
            return True
    return False

def add_new_invoice(ws, inv, last_row):
    r  = last_row + 1
    st = inv.get("status","⏳ Pending")
    bg = STAT_BG.get(st, YELLOW)
    for col_i, val in enumerate([
        inv.get("date",""), inv.get("invoice_no",""), inv.get("payee",""),
        inv.get("ccy",""), inv.get("amount"), None, st,
        inv.get("date_paid",""), inv.get("ref",""), inv.get("notes","")
    ], 1):
        c = ws.cell(r, col_i, val if val is not None else "")
        sc(c, bg=bg, wrap=(col_i in (3,10)), sz=9)
    ws.cell(r,6).value = (f'=IF(OR(E{r}="",E{r}="TBC"),"TBC",'
                          f'IFERROR(E{r}/VLOOKUP(D{r},Settings!$A$7:$B$16,2,FALSE),E{r}))')
    ws.cell(r,6).number_format = '#,##0.00'; sc(ws.cell(r,6), bg=bg)
    ws.row_dimensions[r].height = 26


def apply_transaction_update(ws, upd):
    """Update notes/status of existing transaction row by matching description."""
    match_desc = str(upd.get("match_description","")).lower()
    match_date = str(upd.get("match_date","")).strip()
    new_notes  = upd.get("new_notes","")
    confirmed  = upd.get("confirmed", False)

    if not match_desc:
        return False

    for row in ws.iter_rows(min_row=5):
        r = row[0].row
        desc = str(ws.cell(r,3).value or "").lower()
        date = str(ws.cell(r,1).value or "")
        notes = str(ws.cell(r,12).value or "")

        # Match by keywords in description
        keywords = [w for w in match_desc.split() if len(w) > 3]
        matches = sum(1 for kw in keywords if kw in desc or kw in notes.lower())

        if matches >= 1:
            # Also check date if provided
            if match_date and match_date not in date:
                continue

            # Update notes — remove UNCONFIRMED warning, add confirmation
            updated_notes = notes
            updated_notes = updated_notes.replace("⚠ UNCONFIRMED — ", "✅ CONFIRMED — ")
            updated_notes = updated_notes.replace("Agent did NOT confirm", "Agent CONFIRMED")
            updated_notes = updated_notes.replace("FOLLOW UP!", "")
            if new_notes:
                updated_notes = new_notes
            if confirmed:
                ts = datetime.now().strftime("%d.%m.%Y")
                if "CONFIRMED" not in updated_notes.upper():
                    updated_notes += f" | Подтверждено агентом {ts}"

            ws.cell(r,12).value = updated_notes.strip()

            # Also update description to remove warning
            cur_desc = str(ws.cell(r,3).value or "")
            if "⚠ UNCONFIRMED" in cur_desc or "UNCONFIRMED" in cur_desc:
                ws.cell(r,3).value = cur_desc.replace("⚠ UNCONFIRMED — ", "✅ ").replace("UNCONFIRMED", "CONFIRMED")

            log.info(f"Transaction updated row {r}: {ws.cell(r,3).value}")
            return True

    log.warning(f"Transaction not found for update: {match_desc}")
    return False

def write_to_excel(data: dict):
    if not EXCEL_FILE.exists(): return 0,0,0
    wb  = load_workbook(EXCEL_FILE)
    wst = wb["Transactions"]; wsi = wb["Invoices"]
    tx_a = inv_u = inv_a = 0
    tx_upd = 0
    for tu in data.get("transaction_updates", []):
        if apply_transaction_update(wst, tu): tx_upd += 1
    for tx in data.get("new_transactions", []):
        apply_tx_row(wst, find_last_row(wst) + 1, tx); tx_a += 1
    for upd in data.get("invoice_updates", []):
        if apply_inv_update(wsi, upd): inv_u += 1
    for inv in data.get("new_invoices", []):
        add_new_invoice(wsi, inv, find_last_row(wsi)); inv_a += 1
    wb.save(EXCEL_FILE)
    return tx_a, inv_u, inv_a, tx_upd

# ── Claude API ────────────────────────────────────────────────────────────────
async def ask_claude(prompt: str, system: str = None) -> str:
    sys_msg = system or "You are a financial assistant. Respond in Russian."
    async with httpx.AsyncClient(timeout=90) as client:
        r = await client.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": ANTHROPIC_KEY,
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model": "claude-opus-4-6", "max_tokens": 4000,
                  "system": sys_msg,
                  "messages": [{"role": "user", "content": prompt}]},
        )
        return r.json()["content"][0]["text"]

async def parse_messages(msgs_text: str) -> dict:
    context = load_context()
    excel_bal = get_balance_from_excel()
    bal_str = f"${excel_bal[0]:,.2f} (запись: {excel_bal[1]})" if excel_bal else "нет данных"
    unconfirmed = get_recent_unconfirmed()
    unconfirmed_str = "\n".join(unconfirmed) if unconfirmed else "нет"
    existing_inv = get_existing_invoices_list()

    prompt = f"""КОНТЕКСТ ПРОЕКТА (обязательно учитывай):
{context}

ТЕКУЩИЙ БАЛАНС В EXCEL: {bal_str}

НЕПОДТВЕРЖДЁННЫЕ ТРАНЗАКЦИИ (мы отправили, агент ещё не подтвердил):
{unconfirmed_str}

УЖЕ СУЩЕСТВУЮЩИЕ ИНВОЙСЫ В EXCEL (НЕ добавляй их снова!):
{existing_inv}

---
Из новых сообщений от финансового агента извлеки структурированные данные.

Верни ТОЛЬКО валидный JSON без markdown:
{{
  "new_transactions": [
    {{
      "date": "DD.MM.YYYY",
      "type": "Payment|Deposit|Cash Out|Cash In|❓ Unknown",
      "description": "краткое описание",
      "payee": "название получателя",
      "ccy": "AED|CNY|USD|EUR|SGD|RUB|INR",
      "amount": 12345.67,
      "fx_rate": null,
      "comm": null,
      "notes": "доп. инфо"
    }}
  ],
  "invoice_updates": [
    {{
      "invoice_no": "номер инвойса",
      "new_status": "✅ Paid|⏳ Pending|⚠ Partial/Check|❓ Clarify",
      "date_paid": "DD.MM.YYYY",
      "ref": "референс"
    }}
  ],
  "new_invoices": [
    {{
      "date": "DD.MM.YYYY",
      "invoice_no": "номер",
      "payee": "получатель",
      "ccy": "USD",
      "amount": 12345.67,
      "status": "⏳ Pending",
      "notes": ""
    }}
  ],
  "transaction_updates": [
    {{
      "match_description": "ключевые слова из описания существующей транзакции",
      "match_date": "DD.MM.YYYY или пусто",
      "new_notes": "",
      "confirmed": true
    }}
  ],
  "balance_reconciliation": {{
    "agent_stated_balance": null,
    "our_excel_balance": null,
    "difference": null,
    "difference_explained_by": [],
    "unexplained_difference": null
  }},
  "context_update": "краткая запись для контекста — что нового узнали из этих сообщений",
  "summary": "2-3 предложения — что нового произошло"
}}

Правила:
- Сообщение с балансом агента ("Остаток: X") — занеси в balance_reconciliation, не в транзакции
- "ИСПОЛНЕН", "received", "RCVD", "Поступление подтверждаем", "получили", "поступило" = подтверждение → invoice_updates, НЕ новая транзакция
- Если агент подтверждает получение без деталей — ищи в контексте последнюю UNCONFIRMED/FOLLOW UP транзакцию и обновляй её статус на ✅ Paid
- Депозиты от нас агенту = Deposit. Получатель депозита = конечный получатель денег, не агент и не BALKEMY
- BALKEMY, TROVECO, RAWRIMA, ASTENO = плательщики (наша сторона), а не получатели
- Кэш который агент нам доставляет = Cash Out
- Непонятное → ❓ Unknown
- Если нечего добавить — пустые массивы
- ДЕДУПЛИКАЦИЯ: один и тот же инвойс/платёж упомянут несколько раз — добавь ОДИН РАЗ
- ДЕДУПЛИКАЦИЯ: инвойс уже есть в контексте как оплаченный — не добавляй снова
- ДЕДУПЛИКАЦИЯ: в new_invoices — объединяй дубли, одна запись на один инвойс
- Несколько сообщений об одном событии = одна запись
- ПОДТВЕРЖДЕНИЕ ОТПРАВЛЕННЫХ НАМИ ДЕНЕГ: если агент говорит "Поступление подтверждаем", "received", "получили" 
  БЕЗ указания инвойса — это значит агент подтвердил наш ранее отправленный депозит или Cash In.
  В этом случае НЕ добавляй новую транзакцию! Используй transaction_updates чтобы обновить существующую строку.
  match_description = ключевые слова из описания транзакции (например "150k", "150,000 USD", "50,000 USD")
  new_notes = старые заметки + " | ПОДТВЕРЖДЕНО АГЕНТОМ [дата]"
  confirmed = true

ЛОГИКА СВЕРКИ БАЛАНСА (если агент прислал остаток):
1. agent_stated_balance — сумма из сообщения агента в USD
2. our_excel_balance — последний баланс из Excel (из контекста)
3. difference = our_excel_balance - agent_stated_balance (положительное = мы считаем больше чем агент)
4. difference_explained_by — список транзакций из "НЕПОДТВЕРЖДЁННЫЕ ТРАНЗАКЦИИ" которые объясняют разницу.
   Пример: мы отправили $150k, агент ещё не подтвердил → это объясняет $150k разницы.
   Формат: ["$150,000 USD отправлено 24.02 — агент не подтвердил (Pacs.008)"]
5. unexplained_difference = difference минус сумма объяснённых транзакций
   Если unexplained_difference близко к 0 — всё сходится.
   Если большое — есть реальное расхождение которое надо уточнять у агента.

Новые сообщения:
{msgs_text}"""

    raw = await ask_claude(prompt, system=(
        "You are a JSON extraction assistant. "
        "Return ONLY valid JSON, no markdown, no explanation, no backticks."
    ))
    raw = raw.strip().strip("```").strip()
    if raw.startswith("json"): raw = raw[4:].strip()
    return json.loads(raw)

# ── Format confirmation message ───────────────────────────────────────────────
def format_confirmation(data: dict) -> str:
    lines = ["Вот что я нашёл в сообщениях. Проверь и подтверди запись в Excel.\n"]

    txs = data.get("new_transactions", [])
    if txs:
        lines.append(f"ТРАНЗАКЦИИ ({len(txs)}):")
        for tx in txs:
            amt = f"{tx.get('amount',0):,.2f}" if tx.get('amount') else "?"
            lines.append(f"  + {tx.get('date','')} | {tx.get('type','')} | "
                         f"{tx.get('payee','')} | {amt} {tx.get('ccy','')}")

    upds = data.get("invoice_updates", [])
    if upds:
        lines.append(f"\nОБНОВЛЕНИЯ ИНВОЙСОВ ({len(upds)}):")
        for u in upds:
            lines.append(f"  ~ {u.get('invoice_no','')} → {u.get('new_status','')} "
                         f"({u.get('date_paid','')})")

    invs = data.get("new_invoices", [])
    if invs:
        lines.append(f"\nНОВЫЕ ИНВОЙСЫ ({len(invs)}):")
        for inv in invs:
            amt = f"{inv.get('amount',0):,.2f}" if inv.get('amount') else "TBC"
            lines.append(f"  + {inv.get('payee','')} | {amt} {inv.get('ccy','')} | "
                         f"{inv.get('status','')}")

    rec = data.get("balance_reconciliation", {})
    if rec.get("agent_stated_balance"):
        lines.append(f"\nСВЕРКА БАЛАНСА:")
        lines.append(f"  Агент: {rec.get('agent_stated_balance','?')}")
        lines.append(f"  Excel: {rec.get('our_excel_balance','?')}")
        diff = rec.get("difference")
        if diff is not None:
            lines.append(f"  Разница: {diff:+,.2f}" if isinstance(diff,(int,float)) else f"  Разница: {diff}")
        explained = rec.get("difference_explained_by", [])
        if explained:
            lines.append("  Объясняется:")
            for e in explained:
                lines.append(f"    → {e}")
        unexplained = rec.get("unexplained_difference")
        if unexplained is not None:
            if isinstance(unexplained,(int,float)) and abs(float(unexplained)) < 1000:
                lines.append("  Необъяснённый остаток: ~0 ✅ Сходится!")
            else:
                lines.append(f"  Необъяснённый остаток: {unexplained} ⚠ Уточнить у агента!")

    tx_upds = data.get("transaction_updates", [])
    if tx_upds:
        lines.append(f"\nОБНОВЛЕНИЯ ТРАНЗАКЦИЙ ({len(tx_upds)}):")
        for tu in tx_upds:
            lines.append(f"  ~ {tu.get('match_description','')} "
                         f"({'✅ подтверждено' if tu.get('confirmed') else 'обновлено'})")

    if not txs and not upds and not invs and not tx_upds:
        lines.append("Новых транзакций или инвойсов не найдено.")

    lines.append(f"\nИТОГ: {data.get('summary','')}")
    return "\n".join(lines)

# ── Commands ──────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я ассистент по платежам.\n\n"
        "Просто пиши мне — я понимаю обычный текст:\n"
        "  'какой баланс?'\n"
        "  'покажи pending инвойсы дороже $10k'\n"
        "  'добавь транзакцию: получили 10000 EUR наличными'\n"
        "  'исправь получателя в последней строке'\n\n"
        "Пересылай сообщения от агента → /update.\n\n"
        "Команды:\n"
        "/update  — обработать накопленные сообщения\n"
        "/edit    — изменить строку в Excel\n"
        "/delete  — удалить последнюю транзакцию\n"
        "/balance — текущий баланс\n"
        "/pending — неоплаченные инвойсы\n"
        "/unknown — неизвестные транзакции\n"
        "/summary — полный отчёт\n"
        "/excel   — скачать Excel\n"
        "/clear   — очистить очередь сообщений"
    )


async def cmd_update(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msgs = load_messages()
    if not msgs:
        await update.message.reply_text(
            "Нет накопленных сообщений. Перешли что-нибудь от агента.")
        return

    await update.message.reply_text(f"Анализирую {len(msgs)} сообщений...")

    try:
        data = await parse_messages(_fmt(msgs))
    except Exception as e:
        await update.message.reply_text(f"Ошибка анализа: {e}")
        log.error(f"Parse error: {e}"); return

    txs  = data.get("new_transactions", [])
    upds = data.get("invoice_updates", [])
    invs = data.get("new_invoices", [])

    if not txs and not upds and not invs:
        await update.message.reply_text(
            f"Новых транзакций не найдено.\n\n{data.get('summary','')}")
        # Still update context
        if data.get("context_update"):
            update_context_after_update(data["context_update"])
        clear_messages()
        return

    # Save pending and show confirmation
    save_pending(data)
    conf_text = format_confirmation(data)

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Записать в Excel", callback_data="confirm_update"),
         InlineKeyboardButton("❌ Отмена", callback_data="cancel_update")]
    ])
    await update.message.reply_text(conf_text, reply_markup=keyboard)


def apply_edit(data: dict) -> str:
    """Apply an edit command to Excel."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    sheet_name = data.get("sheet", "Transactions")
    action     = data.get("action", "update")
    row_n      = int(data.get("row_number", 0))
    changes    = data.get("changes", {})
    desc       = data.get("description", "")

    if not EXCEL_FILE.exists():
        return "Excel файл не найден."

    wb = load_workbook(EXCEL_FILE)
    ws = wb[sheet_name]

    thin = Side(style="thin", color="BFBFBF")
    def B(): return Border(top=thin, bottom=thin, left=thin, right=thin)

    if action == "delete":
        ws.delete_rows(row_n)
        wb.save(EXCEL_FILE)
        return f"Строка {row_n} удалена.\n{desc}"

    # Map col names to indices
    col_map = {
        "col_A":1,"col_B":2,"col_C":3,"col_D":4,"col_E":5,
        "col_F":6,"col_G":7,"col_H":8,"col_I":9,"col_J":10,
        "col_K":11,"col_L":12
    }

    applied = []
    for col_name, val in changes.items():
        if val is None: continue
        col_idx = col_map.get(col_name)
        if not col_idx: continue
        cell = ws.cell(row_n, col_idx, val)
        cell.font = Font(name="Arial", size=9)
        cell.border = B()
        cell.alignment = Alignment(vertical="center", wrap_text=(col_idx in (3,12)))
        applied.append(f"{col_name}={val}")

    wb.save(EXCEL_FILE)
    return f"Применено к строке {row_n}:\n" + "\n".join(f"  {a}" for a in applied) + f"\n\n{desc}"

async def callback_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "cancel_update":
        clear_pending()
        await query.edit_message_text("Отменено. Сообщения не удалены — можешь /update снова.")
        return

    data = load_pending()
    if not data:
        await query.edit_message_text("Нет данных для записи.")
        return

    # Handle chat action
    if data.get("type") == "chat_action":
        action = data.get("action","")
        params = data.get("params",{})
        try:
            if action == "add_transaction":
                tx_data = {"new_transactions":[params],"invoice_updates":[],"new_invoices":[]}
                write_to_excel(tx_data)
                msg2 = f"Транзакция добавлена."
            elif action == "add_invoice":
                tx_data = {"new_transactions":[],"invoice_updates":[],"new_invoices":[params]}
                write_to_excel(tx_data)
                msg2 = f"Инвойс добавлен."
            elif action in ("edit_transaction","edit_invoice"):
                sheet = "Transactions" if action == "edit_transaction" else "Invoices"
                edit_data = {"type":"edit","sheet":sheet,
                             "action":"update",
                             "row_number":params.get("row_number"),
                             "changes":params.get("changes",{}),
                             "description":data.get("preview","")}
                msg2 = apply_edit(edit_data)
            elif action == "delete_transaction":
                edit_data = {"type":"edit","sheet":"Transactions",
                             "action":"delete",
                             "row_number":params.get("row_number"),
                             "changes":{},
                             "description":data.get("preview","")}
                msg2 = apply_edit(edit_data)
            else:
                msg2 = f"Неизвестное действие: {action}"
        except Exception as e:
            await query.edit_message_text(f"Ошибка: {e}"); return

        clear_pending()
        await query.edit_message_text(msg2)
        if EXCEL_FILE.exists():
            await ctx.bot.send_document(
                chat_id=MY_CHAT_ID,
                document=EXCEL_FILE.open("rb"),
                filename=f"Agent_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                caption="Excel обновлён"
            )
        return

    # Handle /edit command
    if data.get("type") == "edit":
        try:
            result_text = apply_edit(data)
        except Exception as e:
            await query.edit_message_text(f"Ошибка редактирования: {e}")
            log.error(f"Edit error: {e}"); return
        clear_pending()
        await query.edit_message_text(result_text)
        if EXCEL_FILE.exists():
            await ctx.bot.send_document(
                chat_id=MY_CHAT_ID,
                document=EXCEL_FILE.open("rb"),
                filename=f"Agent_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                caption="Excel после редактирования"
            )
        return

    try:
        tx_a, inv_u, inv_a, tx_upd = write_to_excel(data)
    except Exception as e:
        await query.edit_message_text(f"Ошибка записи в Excel: {e}")
        log.error(f"Excel write error: {e}"); return

    # Update context
    if data.get("context_update"):
        update_context_after_update(data["context_update"])

    clear_pending()
    clear_messages()

    result = (f"Excel обновлён!\n\n"
              f"Транзакций добавлено: {tx_a}\n"
              f"Транзакций обновлено: {tx_upd}\n"
              f"Инвойсов обновлено: {inv_u}\n"
              f"Инвойсов добавлено: {inv_a}")
    await query.edit_message_text(result)

    if EXCEL_FILE.exists():
        await ctx.bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename=f"Agent_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            caption="Обновлённый Excel"
        )


async def cmd_edit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Universal Excel editing command. Examples:
      /edit получили 10,000 EUR наличными в Монако
      /edit поменяй получателя в последней строке на Dubai Insurance
      /edit статус Dubai Insurance — оплачен 25.02.2026
      /edit уточни валюту строки 42 — это BHD а не AED
      /edit удали последнюю строку
      /edit Шафранов оплачен $14,280 банковским переводом
    """
    text = " ".join(ctx.args).strip()
    if not text:
        await update.message.reply_text(
            "Примеры команд:\n\n"
            "ДОБАВИТЬ:\n"
            "/edit получили 10,000 EUR наличными в Монако\n"
            "/edit депозит $50,000 подтверждён агентом сегодня\n"
            "/edit Шафранов оплачен $14,280 банковским переводом\n\n"
            "ИЗМЕНИТЬ:\n"
            "/edit поменяй получателя Dubai Insurance на оплаченный статус\n"
            "/edit исправь валюту последней строки на AED сумма 19502\n"
            "/edit статус Шафранов — оплачен 25.02.2026\n\n"
            "УДАЛИТЬ:\n"
            "/edit удали последнюю транзакцию\n"
            "/edit удали последние 2 строки"
        )
        return

    # Check if it's a delete command
    text_lower = text.lower()
    if any(w in text_lower for w in ["удали", "удалить", "delete"]):
        n = 1
        for word in text_lower.split():
            try: n = int(word); break
            except: pass
        n = min(n, 5)
        if not EXCEL_FILE.exists():
            await update.message.reply_text("Excel файл не найден.")
            return
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Transactions"]
        deleted = []
        for _ in range(n):
            last = None
            for row in ws.iter_rows(min_row=5):
                if row[0].value is not None: last = row[0].row
            if last:
                desc = ws.cell(last,3).value or ""
                amt  = ws.cell(last,6).value or ""
                ccy  = ws.cell(last,5).value or ""
                deleted.append(f"{desc} | {amt} {ccy}")
                ws.delete_rows(last)
        wb.save(EXCEL_FILE)
        result = f"Удалено {len(deleted)} строк:\n" + "\n".join(f"- {d}" for d in deleted)
        await update.message.reply_text(result)
        if EXCEL_FILE.exists():
            await ctx.bot.send_document(chat_id=MY_CHAT_ID,
                document=EXCEL_FILE.open("rb"),
                filename="Agent_after_edit.xlsx", caption="Excel после правки")
        return

    # For all other edits — ask Claude to generate JSON
    context = load_context()

    # Get current last 5 transactions for context
    tx_context = ""
    if EXCEL_FILE.exists():
        try:
            wb2 = load_workbook(EXCEL_FILE, data_only=True)
            ws2 = wb2["Transactions"]
            rows = [r for r in ws2.iter_rows(min_row=5, values_only=True) if r[0] is not None]
            tx_context = "Последние транзакции в Excel:\n" + "\n".join(
                f"  [{r[0]}] {r[1]} | {r[3] or '?'} | {r[5]} {r[4]} | bal={r[10]}"
                for r in rows[-8:]
            )
            inv_rows = [r for r in wb2["Invoices"].iter_rows(min_row=5, values_only=True)
                       if r[0] is not None and r[6] and "Pending" in str(r[6])]
            if inv_rows:
                tx_context += "\n\nPending инвойсы:\n" + "\n".join(
                    f"  {r[1]} | {r[2]} | {r[4]} {r[3]} | {r[6]}"
                    for r in inv_rows[:10]
                )
        except Exception as e:
            log.error(f"Excel read for edit: {e}")

    prompt = f"""КОНТЕКСТ ПРОЕКТА:
{context}

{tx_context}

---
Пользователь хочет внести правку в Excel одной командой:
"{text}"

Это может быть:
1. Добавление новой транзакции
2. Изменение статуса инвойса (если упоминается получатель + "оплачен/paid/исполнен")
3. Исправление данных существующей записи (валюта, сумма, получатель)

Верни ТОЛЬКО валидный JSON без markdown:
{{
  "new_transactions": [
    {{
      "date": "DD.MM.YYYY",
      "type": "Payment|Deposit|Cash Out|Cash In|❓ Unknown",
      "description": "краткое описание",
      "payee": "получатель",
      "ccy": "AED|USD|EUR|CNY|SGD|RUB|INR|BHD",
      "amount": 0.0,
      "fx_rate": null,
      "comm": null,
      "notes": "добавлено вручную"
    }}
  ],
  "invoice_updates": [
    {{
      "invoice_no": "номер или название",
      "new_status": "✅ Paid|⏳ Pending|⚠ Partial/Check|❓ Clarify",
      "date_paid": "DD.MM.YYYY",
      "ref": ""
    }}
  ],
  "new_invoices": [],
  "balance_reconciliation": {{}},
  "context_update": "краткая запись для контекста",
  "summary": "одна строка — что изменили"
}}

Правила:
- "получили кэш/наличные" = Cash Out (агент доставил нам)
- "оплатили/заплатили" = Payment
- "депозит/отправили агенту" = Deposit
- "оплачен/paid/исполнен" + название = invoice_updates
- BALKEMY/TROVECO/RAWRIMA = плательщики, не получатели
- Дата не указана → сегодня: {datetime.now().strftime("%d.%m.%Y")}
- Если правка касается существующей записи (не новая) → используй invoice_updates
- Если нужно добавить новую строку → new_transactions"""

    await update.message.reply_text("Анализирую...")

    try:
        raw = await ask_claude(prompt, system=(
            "You are a JSON extraction assistant. "
            "Return ONLY valid JSON, no markdown, no backticks."
        ))
        raw = raw.strip().strip("`").strip()
        if raw.startswith("json"): raw = raw[4:].strip()
        data = json.loads(raw)
    except Exception as e:
        await update.message.reply_text(f"Ошибка анализа: {e}")
        return

    conf_text = format_confirmation(data)
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Записать в Excel", callback_data="confirm_update"),
         InlineKeyboardButton("❌ Отмена", callback_data="cancel_update")]
    ])
    save_pending(data)
    await update.message.reply_text(conf_text, reply_markup=keyboard)

# Keep /add as alias for /edit
async def cmd_add(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await cmd_edit(update, ctx)



async def handle_chat(update: Update, ctx: ContextTypes.DEFAULT_TYPE, user_text: str):
    """Handle free-form user messages as AI conversation."""
    context    = load_context()
    excel_sum  = get_excel_summary()
    history    = load_chat_history()
    existing_inv = get_existing_invoices_list()

    # Build messages for Claude
    system_prompt = f"""Ты финансовый ассистент для трекера платежей через агента.
Отвечай по-русски, кратко и по делу.
Ты можешь:
- Отвечать на вопросы о балансе, инвойсах, транзакциях
- Вносить изменения в Excel (добавить транзакцию, изменить инвойс, удалить строку)
- Анализировать ситуацию и давать рекомендации

ТЕКУЩЕЕ СОСТОЯНИЕ EXCEL:
{excel_sum}

УЖЕ СУЩЕСТВУЮЩИЕ ИНВОЙСЫ:
{existing_inv}

КОНТЕКСТ ПРОЕКТА:
{context}

Если пользователь просит ДЕЙСТВИЕ (добавить/изменить/удалить) — ответь JSON:
{{
  "type": "action",
  "action": "add_transaction|add_invoice|edit_transaction|edit_invoice|delete_transaction",
  "params": {{...}},
  "preview": "одна строка — что именно сделаем",
  "message": "текст ответа пользователю"
}}

Если просто ВОПРОС или разговор — ответь JSON:
{{
  "type": "text",
  "message": "твой ответ"
}}

Параметры для add_transaction:
  date, type(Payment|Deposit|Cash Out|Cash In|❓ Unknown), description, payee, ccy, amount, fx_rate(null), comm(null), notes

Параметры для add_invoice:
  date, invoice_no, payee, ccy, amount, status(⏳ Pending), notes

Параметры для edit_transaction/edit_invoice:
  row_number, changes: {{col_X: value, ...}}
  Transactions: col_A=Date,col_B=Type,col_C=Desc,col_D=Payee,col_E=CCY,col_F=Amt,col_G=FX,col_H=GrossUSD,col_I=Comm%,col_J=NetUSD,col_K=Bal,col_L=Notes
  Invoices: col_A=Date,col_B=InvNo,col_C=Payee,col_D=CCY,col_E=Amt(ЧИСЛО!),col_F=USD,col_G=Status,col_H=DatePaid,col_I=Ref,col_J=Notes

Параметры для delete_transaction:
  row_number"""

    messages = []
    # Add history
    for h in history:
        messages.append(h)
    # Add current message
    messages.append({"role": "user", "content": user_text})

    try:
        async with httpx.AsyncClient(timeout=60) as client:
            r = await client.post(
                "https://api.anthropic.com/v1/messages",
                headers={"x-api-key": ANTHROPIC_KEY,
                         "anthropic-version": "2023-06-01",
                         "content-type": "application/json"},
                json={"model": "claude-opus-4-6", "max_tokens": 1500,
                      "system": system_prompt, "messages": messages},
            )
            raw = r.json()["content"][0]["text"].strip().strip("`").strip()
            if raw.startswith("json"): raw = raw[4:].strip()
            data = json.loads(raw)
    except Exception as e:
        log.error(f"Chat error: {e}")
        await update.message.reply_text(f"Ошибка: {e}")
        return

    msg_text = data.get("message", "")
    response_type = data.get("type", "text")

    if response_type == "action":
        action   = data.get("action","")
        params   = data.get("params", {})
        preview  = data.get("preview","")

        # Save pending action
        pending_data = {"type": "chat_action", "action": action,
                        "params": params, "preview": preview}
        save_pending(pending_data)

        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Применить", callback_data="confirm_update"),
             InlineKeyboardButton("❌ Отмена",    callback_data="cancel_update")]
        ])
        reply = f"{msg_text}\n\n📋 {preview}"
        await update.message.reply_text(reply, reply_markup=keyboard)
    else:
        await update.message.reply_text(msg_text)

    # Save to history
    history.append({"role": "user",      "content": user_text})
    history.append({"role": "assistant", "content": msg_text})
    save_chat_history(history)

async def cmd_delete(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Delete last N rows from Transactions sheet. Usage: /delete or /delete 2"""
    n = 1
    if ctx.args:
        try: n = int(ctx.args[0])
        except: pass
    n = min(n, 5)  # max 5 at once for safety

    if not EXCEL_FILE.exists():
        await update.message.reply_text("Excel файл не найден.")
        return

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Transactions"]

    deleted = []
    for _ in range(n):
        last = None
        for row in ws.iter_rows(min_row=5):
            if row[0].value is not None:
                last = row[0].row
        if last:
            desc  = ws.cell(last, 3).value or ""
            date  = ws.cell(last, 1).value or ""
            amt   = ws.cell(last, 6).value or ""
            ccy   = ws.cell(last, 5).value or ""
            deleted.append(f"[{date}] {desc} | {amt} {ccy}")
            ws.delete_rows(last)

    wb.save(EXCEL_FILE)

    if deleted:
        msg_text = f"Удалено {len(deleted)} строк:\n" + "\n".join(f"- {d}" for d in deleted)
        await update.message.reply_text(msg_text)
        await ctx.bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename="Agent_after_delete.xlsx",
            caption="Excel после удаления"
        )
    else:
        await update.message.reply_text("Нет строк для удаления.")


async def cmd_edit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Natural language Excel editor. Examples:
      /edit поменяй получателя в последней строке на MAXIMUS WAY
      /edit строка с Port of Fujairah — смени статус на оплачен 25.02
      /edit удали последнюю транзакцию
      /edit Dubai Insurance — обновить сумму 197334.74 AED
      /edit $150k строка 41 — добавь ref OR26022400002178 в примечания
    """
    text = " ".join(ctx.args).strip()
    if not text:
        await update.message.reply_text(
            "Укажи что изменить. Примеры:\n\n"
            "/edit поменяй получателя в последней строке на MAXIMUS WAY\n"
            "/edit строка с Port of Fujairah — статус оплачен 25.02\n"
            "/edit удали последнюю транзакцию\n"
            "/edit Dubai Insurance — обновить сумму 197334 AED\n"
            "/edit строка 41 — добавь примечание: подтверждено агентом"
        )
        return

    if not EXCEL_FILE.exists():
        await update.message.reply_text("Excel файл не найден.")
        return

    # Read current Excel state to give Claude context
    wb_read = load_workbook(EXCEL_FILE, data_only=True)
    ws_read = wb_read["Transactions"]
    wi_read = wb_read["Invoices"]

    tx_rows = []
    for row in ws_read.iter_rows(min_row=5, values_only=True):
        if row[0] is not None:
            tx_rows.append(row)

    inv_rows = []
    for row in wi_read.iter_rows(min_row=5, values_only=True):
        if row[0] is not None or row[1] is not None:
            inv_rows.append(row)

    # Format for Claude
    tx_text = "\n".join(
        f"Row {i+5}: [{r[0]}] {r[1]} | {r[3] or '?'} | {r[5]} {r[4]} | bal={r[10]} | notes={r[11] or ''}"
        for i, r in enumerate(tx_rows)
    )
    inv_text = "\n".join(
        f"Row {i+5}: [{r[0]}] inv={r[1]} | {r[2]} | {r[4]} {r[3]} | status={r[6]} | paid={r[7]}"
        for i, r in enumerate(inv_rows)
    )

    context = load_context()

    prompt = f"""КОНТЕКСТ:
{context}

ТЕКУЩИЕ ТРАНЗАКЦИИ (Transactions sheet, строки начиная с 5):
Колонки: A=Date, B=Type, C=Description, D=Payee, E=CCY, F=Amount, G=FX, H=GrossUSD, I=Comm%, J=NetUSD, K=Balance, L=Notes
{tx_text}

ТЕКУЩИЕ ИНВОЙСЫ (Invoices sheet):
Колонки: A=Date, B=InvNo, C=Payee, D=CCY, E=Amount, F=USD, G=Status, H=DatePaid, I=Ref, J=Notes
{inv_text}

КОМАНДА ПОЛЬЗОВАТЕЛЯ: {text}

Верни ТОЛЬКО валидный JSON без markdown:
{{
  "sheet": "Transactions|Invoices",
  "action": "update|delete",
  "row_number": 42,
  "changes": {{
    "col_A": null,
    "col_B": null,
    "col_C": null,
    "col_D": null,
    "col_E": null,
    "col_F": null,
    "col_G": null,
    "col_H": null,
    "col_I": null,
    "col_J": null,
    "col_K": null,
    "col_L": null
  }},
  "description": "одна строка — что именно меняем и почему"
}}

Правила:
- row_number: точный номер строки Excel (начиная с 5)
- changes: только те колонки которые нужно изменить, остальные null
- action=delete: удалить строку целиком
- Если команда непонятна или строка не найдена — верни {{"error": "описание проблемы"}}
- Не пересчитывай баланс — только меняй указанные поля

КОЛОНКИ Transactions: col_A=Date, col_B=Type, col_C=Description, col_D=Payee, col_E=CCY, col_F=Amount(число), col_G=FX, col_H=GrossUSD, col_I=Comm%, col_J=NetUSD, col_K=Balance, col_L=Notes
КОЛОНКИ Invoices: col_A=Date, col_B=InvNo, col_C=Payee, col_D=CCY(валюта), col_E=Amount(ЧИСЛО!), col_F=USD_equiv, col_G=Status, col_H=DatePaid, col_I=Ref, col_J=Notes

ВАЖНО для Invoices:
- col_D = валюта (AED/USD/EUR/etc) — СТРОКА
- col_E = сумма — ЧИСЛО (например 242022.05, не "AED"!)
- Никогда не пиши валюту в col_E — только число"""

    await update.message.reply_text("Анализирую команду...")

    try:
        raw = await ask_claude(prompt, system=(
            "You are a JSON assistant. Return ONLY valid JSON, no markdown, no backticks."
        ))
        raw = raw.strip().strip("`").strip()
        if raw.startswith("json"): raw = raw[4:].strip()
        data = json.loads(raw)
    except Exception as e:
        await update.message.reply_text(f"Ошибка анализа: {e}")
        return

    if "error" in data:
        await update.message.reply_text(f"Не понял команду: {data['error']}")
        return

    # Show confirmation
    desc = data.get("description","")
    row_n = data.get("row_number")
    action = data.get("action","update")
    sheet = data.get("sheet","Transactions")
    changes = data.get("changes",{})

    non_null = {k:v for k,v in changes.items() if v is not None}
    changes_text = "\n".join(f"  {k}: {v}" for k,v in non_null.items()) if non_null else "удаление строки"

    confirm_text = (
        f"Команда: {desc}\n\n"
        f"Лист: {sheet}\n"
        f"Строка: {row_n}\n"
        f"Действие: {action}\n"
        f"Изменения:\n{changes_text}"
    )

    save_pending({"type": "edit", "sheet": sheet, "action": action,
                  "row_number": row_n, "changes": changes, "description": desc})

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Применить", callback_data="confirm_update"),
         InlineKeyboardButton("❌ Отмена",    callback_data="cancel_update")]
    ])
    await update.message.reply_text(confirm_text, reply_markup=keyboard)

async def cmd_balance(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    result = get_balance_from_excel()
    if result:
        bal, date = result
        await update.message.reply_text(
            f"БАЛАНС АГЕНТА (из Excel)\n${bal:,.2f} USD\nПоследняя запись: {date}")
    else:
        await update.message.reply_text("Баланс не найден. Попробуй /update")

async def cmd_pending(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    pending, usd_total, tbc_count = get_pending_invoices()
    tbc_note = f"\n(+ {tbc_count} инвойс(ов) с суммой TBC)" if tbc_count else ""
    text = (f"ОЖИДАЮТ ОПЛАТЫ ({len(pending)}):\n\n" +
            ("\n".join(pending) if pending else "нет") +
            (f"\n\nИТОГО: ~${usd_total:,.0f} USD{tbc_note}" if pending else ""))
    await update.message.reply_text(text)

async def cmd_unknown(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    unknowns = get_unknown_transactions()
    text = f"НЕИЗВЕСТНЫЕ ТРАНЗАКЦИИ ({len(unknowns)}):\n\n" + (
        "\n".join(unknowns) if unknowns else "нет — хорошо!")
    await update.message.reply_text(text)

async def cmd_context(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    context = load_context()
    if not context:
        await update.message.reply_text("Контекст пуст.")
        return
    # Telegram limit 4096 chars
    if len(context) > 3800:
        context = context[-3800:]
        await update.message.reply_text(
            f"(Показаны последние 3800 символов)\n\n{context}")
    else:
        await update.message.reply_text(context)

async def cmd_summary(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await _send_report(ctx.bot, triggered_manually=True)

async def cmd_excel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if EXCEL_FILE.exists():
        await ctx.bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename="Agent_Model.xlsx",
            caption="Актуальный Excel"
        )
    else:
        await update.message.reply_text("Excel файл не найден на сервере.")

async def cmd_clear(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    clear_messages()
    await update.message.reply_text("Накопленные сообщения очищены.")

# ── Message handler ───────────────────────────────────────────────────────────
async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg: return
    if msg.chat_id != MY_CHAT_ID:
        log.info(f"IGNORED chat_id={msg.chat_id} expected={MY_CHAT_ID}")
        return
    text     = msg.text or msg.caption or ""

    # If NOT a forwarded message and NOT a document — treat as chat
    is_forwarded = bool(
        getattr(msg, "forward_origin", None) or
        getattr(msg, "forward_sender_name", None) or
        getattr(msg, "forward_from", None)
    )
    if not is_forwarded and not msg.document and text:
        await handle_chat(update, ctx, text)
        return

    # Handle both old and new telegram-bot API forward attributes
    sender = ""
    try:
        origin = msg.forward_origin
        if origin:
            if hasattr(origin, "sender_name"):
                sender = origin.sender_name or ""
            elif hasattr(origin, "sender_user"):
                u = origin.sender_user
                sender = (u.full_name or u.username or "") if u else ""
            elif hasattr(origin, "chat"):
                c = origin.chat
                sender = (c.title or c.username or "") if c else ""
    except Exception:
        pass
    if not sender:
        try:
            sender = msg.forward_sender_name or ""
        except Exception:
            sender = ""
    date_str = msg.date.strftime("%d.%m.%Y %H:%M")
    try:
        if msg.forward_date:
            date_str = msg.forward_date.strftime("%d.%m.%Y %H:%M")
    except Exception:
        pass
    file_n   = msg.document.file_name if msg.document else ""
    pdf_text = ""

    # Download and read PDF if attached
    if msg.document and file_n.lower().endswith(".pdf") and HAS_PDF:
        try:
            tg_file = await msg.document.get_file()
            buf = io.BytesIO()
            await tg_file.download_to_memory(buf)
            buf.seek(0)
            reader = PdfReader(buf)
            pages_text = []
            for page in reader.pages:
                t = page.extract_text()
                if t: pages_text.append(t.strip())
            pdf_text = "\n".join(pages_text)[:3000]  # limit
            log.info(f"PDF extracted: {len(pdf_text)} chars from {file_n}")
        except Exception as e:
            log.error(f"PDF read error: {e}")
            pdf_text = f"[PDF не удалось прочитать: {e}]"

    entry = {"date": date_str, "sender": sender, "text": text, "file": file_n}
    if pdf_text:
        entry["pdf_content"] = pdf_text
    save_message(entry)

    preview = text[:60] + ("…" if len(text) > 60 else "")
    parts   = [f"от {sender}"] if sender else []
    if file_n:
        pdf_note = " (прочитан)" if pdf_text else ""
        parts.append(f"файл: {file_n}{pdf_note}")
    if preview: parts.append(f'"{preview}"')
    count = len(load_messages())
    await msg.reply_text(
        f"Сохранено | {' | '.join(parts)}\n"
        f"В очереди: {count} сообщений. Когда готов — /update"
    )

# ── Morning report ────────────────────────────────────────────────────────────
async def _send_report(bot: Bot, triggered_manually=False):
    today = datetime.now().strftime("%d.%m.%Y")
    msgs  = load_messages()
    updates_text = ""

    if msgs:
        try:
            data    = await parse_messages(_fmt(msgs))
            tx_a, inv_u, inv_a, tx_upd = write_to_excel(data)
            if data.get("context_update"):
                update_context_after_update(data["context_update"])
            if tx_a + inv_u + inv_a > 0:
                updates_text = (f"Автообновлено: +{tx_a} транзакций, "
                                f"{inv_u} инвойсов обновлено, +{inv_a} новых\n\n")
            clear_messages()
        except Exception as e:
            log.error(f"Auto-update error: {e}")
            updates_text = f"(Ошибка автообновления: {e})\n\n"

    result  = get_balance_from_excel()
    bal_str = f"${result[0]:,.2f} USD (запись: {result[1]})" if result else "нет данных"
    pending, usd_total, tbc_count = get_pending_invoices()
    unknown = get_unknown_transactions()
    tbc_note = f"\n  (+ {tbc_count} инвойс(ов) с суммой TBC — не включены)" if tbc_count else ""
    pending_total_str = f"\nИТОГО К ОПЛАТЕ: ~${usd_total:,.0f} USD{tbc_note}"

    text = (f"ОТЧЁТ — {today}\n\n"
            f"{updates_text}"
            f"БАЛАНС: {bal_str}\n\n"
            f"ОЖИДАЮТ ОПЛАТЫ ({len(pending)}):\n"
            + ("\n".join(pending) if pending else "нет") +
            (pending_total_str if pending else "") +
            f"\n\nНЕИЗВЕСТНЫЕ ({len(unknown)}):\n"
            + ("\n".join(unknown) if unknown else "нет"))

    await bot.send_message(chat_id=MY_CHAT_ID, text=text)
    if EXCEL_FILE.exists():
        await bot.send_document(
            chat_id=MY_CHAT_ID,
            document=EXCEL_FILE.open("rb"),
            filename=f"Agent_{datetime.now().strftime('%Y%m%d')}.xlsx",
            caption="Актуальный Excel"
        )

async def morning_job(ctx: ContextTypes.DEFAULT_TYPE):
    await _send_report(ctx.bot)

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    Path("data").mkdir(exist_ok=True)
    app = Application.builder().token(BOT_TOKEN).build()
    for cmd, fn in [
        ("start", cmd_start), ("update", cmd_update),
        ("edit", cmd_edit),
        ("delete", cmd_delete),
        ("edit", cmd_edit),
        ("balance", cmd_balance), ("pending", cmd_pending),
        ("unknown", cmd_unknown), ("summary", cmd_summary),
        ("excel", cmd_excel), ("context", cmd_context), ("clear", cmd_clear)
    ]:
        app.add_handler(CommandHandler(cmd, fn))
    app.add_handler(CallbackQueryHandler(callback_confirm))
    app.add_handler(MessageHandler(filters.ALL & ~filters.COMMAND, handle_message))
    app.job_queue.run_daily(morning_job, time=time(hour=MORNING_HOUR, minute=0))
    log.info("Bot v3 started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
