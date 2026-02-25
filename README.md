# Payment Tracker Bot

Скопируй файл `Agent_Model_v2.xlsx` в папку `data/` перед деплоем.

## Переменные окружения (Railway → Variables)

| Переменная | Пример | Описание |
|---|---|---|
| BOT_TOKEN | 7123456789:AAF-... | Токен от @BotFather |
| ANTHROPIC_KEY | sk-ant-api03-... | Ключ с console.anthropic.com |
| MY_CHAT_ID | 123456789 | Твой Telegram ID (от @userinfobot) |
| MORNING_HOUR | 9 | Час утреннего отчёта (UTC+4 = UTC+0 минус 4) |

## Как пользоваться

Пересылай боту сообщения от агента → каждое утро получаешь саммари + Excel.

Команды: /balance /pending /summary /excel /unknown /clear
