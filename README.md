# Telegram Contact Converter — Render Free Web Service

No admin, VIP, referral, database, or ZIP output. All converted files are sent as Telegram documents.

## Deploy
1. Upload `bot.py`, `requirements.txt`, `README.md` to the root of a GitHub repository.
2. Render → New → Web Service → connect repository → Free.
3. Build: `pip install -r requirements.txt`
4. Start: `uvicorn bot:app --host 0.0.0.0 --port $PORT`
5. Set environment variables:
   - `BOT_TOKEN` = newly generated BotFather token (NEVER publish it)
   - `WEBHOOK_URL` = `https://YOUR-SERVICE.onrender.com/webhook`
   - `WEBHOOK_SECRET` = Render-generated random string (Telegram accepts A-Z, a-z, 0-9, _, -; length 1-256)
6. Deploy and send `/start` to bot. The app registers the webhook on startup.

Note: Webhook URL must match the actual Render URL. Free services may sleep, and conversation state resets on restart. Do not run another polling bot with the same token. Maximum supported input and output file size: 18 MB each. Maximum 50 output files per conversion.

## Formats
TXT: one phone number per line, optionally `Name,Phone`. CSV: `Mobile` header (or Phone/Number) and optionally Name. VCF: vCard 3.0 (simple FN and TEL records).

## Commands
Use `/help` in Telegram. Sending a TXT file without a command starts CSV conversion and asks output filename and number of CSV files. `/to_vcf` asks contact name, output filename, contacts per VCF, and VCF file count. The requested VCF count must match the number needed to include all contacts.

## Security
Rotate any bot token that was ever posted in a chat or screenshot. Do not upload real contact lists to public GitHub. Files are processed in RAM and are not intentionally stored. Only use contact data you have permission to process.
