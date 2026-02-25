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
    return "\n".join(
        f"[{m['date']}] {m.get('sender','?')}: {m.get('text','')} "
        f"{'[файл: '+m['file']+']' if m.get('file') else ''}"
        for m in msgs
    )

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
    if not EXCEL_FILE.exists(): return []
    try:
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = wb["Invoices"]
        out = []
        for row in ws.iter_rows(min_row=5, max_col=10, values_only=True):
            if row[6] and "Pending" in str(row[6]):
                amt = f"{row[4]:,.2f}" if isinstance(row[4], (int,float)) else str(row[4] or "TBC")
                out.append(f"- {row[2] or '?'}: {amt} {row[3] or ''}")
        return out
    except Exception as e:
        log.error(f"Excel pending: {e}"); return []

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

def apply_tx_row(ws, r, tx):
    tp = tx.get("type", "Payment")
    bg = TYPE_BG.get(tp, WHITE)
    for col_i, val in enumerate([
        tx.get("date",""), tp, tx.get("description",""), tx.get("payee",""),
        tx.get("ccy",""), tx.get("amount"), tx.get("fx_rate"),
        None, tx.get("comm"), None, None, tx.get("notes","")
    ], 1):
        c = ws.cell(r, col_i, val if val is not None else "")
        sc(c, bg=bg, wrap=(col_i in (3,12)), sz=9)
    # G: FX rate
    if not tx.get("fx_rate"):
        ws.cell(r,7).value = f'=IF(E{r}="","",IFERROR(VLOOKUP(E{r},Settings!$A$7:$B$16,2,FALSE),1))'
    ws.cell(r,7).number_format = "0.00000"; sc(ws.cell(r,7), bg=YELLOW, fc="0000CC")
    # I: comm
    if not tx.get("comm"):
        ws.cell(r,9).value = f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},Settings!$E$19:$F$23,2,FALSE),0))'
    ws.cell(r,9).number_format = "0.0%"; sc(ws.cell(r,9), bg=YELLOW, fc="0000CC")
    # H: Gross
    ws.cell(r,8).value = f'=IF(OR(F{r}="",G{r}=""),"",F{r}/G{r})'
    ws.cell(r,8).number_format = '#,##0.00'; sc(ws.cell(r,8), bg=YELLOW)
    # J: Net
    ws.cell(r,10).value = (f'=IF(H{r}="","",IF(OR(B{r}="Deposit",B{r}="Cash In"),'
                           f'H{r},-(H{r}/MAX(1-I{r},0.0001))))')
    ws.cell(r,10).number_format = '#,##0.00'; sc(ws.cell(r,10), bg=YELLOW)
    # K: Balance
    ws.cell(r,11).value = (f'=IF(J{r}="","",IF(ISNUMBER(K{r-1}),K{r-1},Settings!$C$4)+J{r})')
    ws.cell(r,11).number_format = '#,##0.00'; sc(ws.cell(r,11), bg=YELLOW, bold=True)
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

def write_to_excel(data: dict):
    if not EXCEL_FILE.exists(): return 0,0,0
    wb  = load_workbook(EXCEL_FILE)
    wst = wb["Transactions"]; wsi = wb["Invoices"]
    tx_a = inv_u = inv_a = 0
    for tx in data.get("new_transactions", []):
        apply_tx_row(wst, find_last_row(wst) + 1, tx); tx_a += 1
    for upd in data.get("invoice_updates", []):
        if apply_inv_update(wsi, upd): inv_u += 1
    for inv in data.get("new_invoices", []):
        add_new_invoice(wsi, inv, find_last_row(wsi)); inv_a += 1
    wb.save(EXCEL_FILE)
    return tx_a, inv_u, inv_a

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
    prompt = f"""КОНТЕКСТ ПРОЕКТА (обязательно учитывай):
{context}

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
  "balance_reconciliation": {{
    "agent_stated_balance": null,
    "discrepancies": []
  }},
  "context_update": "краткая запись для контекста — что нового узнали из этих сообщений",
  "summary": "2-3 предложения — что нового произошло"
}}

Правила:
- Сообщение с балансом агента ("Остаток: X") — занеси в balance_reconciliation, не в транзакции
- "ИСПОЛНЕН", "received", "RCVD" = подтверждение → invoice_updates
- Депозиты от нас агенту = Deposit, кэш который агент нам доставляет = Cash Out
- Непонятное → ❓ Unknown
- Если нечего добавить — пустые массивы

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
        lines.append(f"\nБАЛАНС ОТ АГЕНТА: {rec['agent_stated_balance']}")
        excel_bal = get_balance_from_excel()
        if excel_bal:
            lines.append(f"БАЛАНС В EXCEL: ${excel_bal[0]:,.2f}")
        if rec.get("discrepancies"):
            lines.append("РАСХОЖДЕНИЯ: " + "; ".join(rec["discrepancies"]))

    if not txs and not upds and not invs:
        lines.append("Новых транзакций или инвойсов не найдено.")

    lines.append(f"\nИТОГ: {data.get('summary','')}")
    return "\n".join(lines)

# ── Commands ──────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я трекер платежей с памятью.\n\n"
        "Пересылай мне сообщения от агента, потом:\n\n"
        "/update  — обработать и показать что нашёл (с подтверждением)\n"
        "/add     — быстро добавить операцию вручную одной строкой\n"
        "/balance — баланс из Excel\n"
        "/pending — что висит\n"
        "/unknown — неизвестные транзакции\n"
        "/summary — полный отчёт\n"
        "/excel   — скачать Excel\n"
        "/context — посмотреть контекст\n"
        "/clear   — очистить накопленные сообщения"
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

    try:
        tx_a, inv_u, inv_a = write_to_excel(data)
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


async def cmd_add(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """
    Quick manual add. Examples:
      /add получили 10,000 EUR наличными в Монако
      /add оплатили инвойс Dubai Insurance 197334.74 AED
      /add депозит 50000 USD подтверждён агентом
    """
    text = " ".join(ctx.args).strip()
    if not text:
        await update.message.reply_text(
            "Укажи операцию после команды. Примеры:\n\n"
            "/add получили 10,000 EUR наличными в Монако\n"
            "/add оплатили инвойс Dubai Insurance 197334.74 AED\n"
            "/add депозит 50000 USD подтверждён агентом\n"
            "/add Шафранов оплачен $14,280 банковским переводом"
        )
        return

    context = load_context()
    prompt = f"""КОНТЕКСТ ПРОЕКТА:
{context}

---
Пользователь вручную добавляет операцию одной строкой:
"{text}"

Преобразуй в JSON. Верни ТОЛЬКО валидный JSON без markdown:
{{
  "new_transactions": [
    {{
      "date": "DD.MM.YYYY",
      "type": "Payment|Deposit|Cash Out|Cash In|❓ Unknown",
      "description": "краткое описание",
      "payee": "получатель или отправитель",
      "ccy": "EUR|AED|USD|CNY|SGD|RUB|INR",
      "amount": 10000.00,
      "fx_rate": null,
      "comm": null,
      "notes": "добавлено вручную через /add"
    }}
  ],
  "invoice_updates": [],
  "new_invoices": [],
  "balance_reconciliation": {{}},
  "context_update": "краткая запись для контекста",
  "summary": "одна строка — что добавили"
}}

Правила определения типа:
- "получили", "кэш", "наличные", "доставили" = Cash Out (агент доставил нам)
- "оплатили", "заплатили", "платёж" = Payment
- "депозит", "отправили агенту", "пополнили" = Deposit
- "получили от нас", "мы отправили" = Cash In
- Дата не указана → используй сегодняшнюю: {datetime.now().strftime("%d.%m.%Y")}"""

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
         InlineKeyboardButton("❌ Отмена",           callback_data="cancel_update")]
    ])
    save_pending(data)
    await update.message.reply_text(conf_text, reply_markup=keyboard)

async def cmd_balance(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    result = get_balance_from_excel()
    if result:
        bal, date = result
        await update.message.reply_text(
            f"БАЛАНС АГЕНТА (из Excel)\n${bal:,.2f} USD\nПоследняя запись: {date}")
    else:
        await update.message.reply_text("Баланс не найден. Попробуй /update")

async def cmd_pending(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    pending = get_pending_invoices()
    text = f"ОЖИДАЮТ ОПЛАТЫ ({len(pending)}):\n\n" + (
        "\n".join(pending) if pending else "нет")
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
    if msg.chat_id != MY_CHAT_ID: return
    text     = msg.text or msg.caption or ""
    sender   = (msg.forward_from.full_name if msg.forward_from
                else msg.forward_sender_name or "")
    date_str = (msg.forward_date or msg.date).strftime("%d.%m.%Y %H:%M")
    file_n   = msg.document.file_name if msg.document else ""
    save_message({"date": date_str, "sender": sender,
                  "text": text, "file": file_n})
    preview = text[:60] + ("…" if len(text) > 60 else "")
    parts   = [f"от {sender}"] if sender else []
    if file_n:  parts.append(f"файл: {file_n}")
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
            tx_a, inv_u, inv_a = write_to_excel(data)
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
    pending = get_pending_invoices()
    unknown = get_unknown_transactions()

    text = (f"ОТЧЁТ — {today}\n\n"
            f"{updates_text}"
            f"БАЛАНС: {bal_str}\n\n"
            f"ОЖИДАЮТ ОПЛАТЫ ({len(pending)}):\n"
            + ("\n".join(pending) if pending else "нет") +
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
        ("add", cmd_add),
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
