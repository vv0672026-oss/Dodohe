import os, re, csv, io, math, asyncio, logging, tempfile
from pathlib import Path
from collections import OrderedDict
from telegram import Update, BotCommand, InputFile, ReplyKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters
from fastapi import FastAPI, Request, HTTPException
from contextlib import asynccontextmanager

logging.basicConfig(level=logging.INFO)
TOKEN = os.environ['BOT_TOKEN']
WEBHOOK_URL = os.environ['WEBHOOK_URL'].rstrip('/')
SECRET = os.environ['WEBHOOK_SECRET']
MAX_BYTES = 18 * 1024 * 1024
MAX_OUTPUT = 50
appbot = Application.builder().token(TOKEN).build()

# All user state is in RAM. It resets if Render restarts.
state = {}
PHONE = re.compile(r'^\+?[0-9][0-9 ()-]{5,22}[0-9]$')

def normalize(s):
    return re.sub(r'[^0-9+]', '', s.strip())

def parse_txt(data):
    text = data.decode('utf-8-sig', errors='replace')
    rows = []
    for line in text.splitlines():
        line=line.strip()
        if not line: continue
        # Supports one phone per line, or name,phone / name|phone.
        parts=re.split(r'[,;|\t]', line)
        if len(parts)>1:
            name=' '.join(parts[:-1]).strip(); raw=parts[-1].strip()
        else:
            name=''; raw=line
        if PHONE.fullmatch(raw): rows.append((name, normalize(raw)))
    return rows

def parse_csv(data):
    s=io.StringIO(data.decode('utf-8-sig', errors='replace'))
    rows=list(csv.reader(s))
    if not rows: return []
    header=[x.strip().lower() for x in rows[0]]
    number_idx=next((i for i,x in enumerate(header) if x in ('mobile','phone','number','telephone','tel')),None)
    name_idx=next((i for i,x in enumerate(header) if x in ('name','contact','contact name')),None)
    start=1 if number_idx is not None else 0
    if number_idx is None: number_idx=0
    out=[]
    for row in rows[start:]:
        if number_idx>=len(row): continue
        raw=row[number_idx].strip()
        if PHONE.fullmatch(raw): out.append((row[name_idx].strip() if name_idx is not None and name_idx<len(row) else '',normalize(raw)))
    return out

def vcf_unescape(s):
    return s.replace('\\n','\n').replace('\\,',',').replace('\\;',';').replace('\\\\','\\')

def parse_vcf(data):
    text=data.decode('utf-8-sig', errors='replace').replace('\r\n','\n')
    # Unfold folded vCard lines.
    text=re.sub(r'\n[ \t]', '', text)
    out=[]
    for card in re.findall(r'BEGIN:VCARD\s*(.*?)\s*END:VCARD',text,re.I|re.S):
        name=''; nums=[]
        for line in card.splitlines():
            k,sep,val=line.partition(':')
            if not sep: continue
            field=k.split(';')[0].upper()
            if field=='FN': name=vcf_unescape(val.strip())
            if field=='TEL' and PHONE.fullmatch(val.strip()): nums.append(normalize(val))
        for num in nums: out.append((name,num))
    return out

def parse_file(filename,data):
    ext=Path(filename).suffix.lower()
    if ext=='.txt': return parse_txt(data)
    if ext=='.csv': return parse_csv(data)
    if ext=='.vcf': return parse_vcf(data)
    raise ValueError('Only .txt, .csv and .vcf files are supported.')

def csv_bytes(rows):
    s=io.StringIO(newline=''); w=csv.writer(s); w.writerow(['Mobile']); w.writerows([[p] for _,p in rows]); return s.getvalue().encode('utf-8-sig')

def txt_bytes(rows): return ('\n'.join(p for _,p in rows)+'\n').encode('utf-8')

def esc_vcf(s): return s.replace('\\','\\\\').replace('\n','\\n').replace(',','\\,').replace(';','\\;')

def vcf_bytes(rows, base='Contact'):
    out=[]
    for i,(name,num) in enumerate(rows,1):
        name=name or f'{base} {i:03d}'
        out.extend(['BEGIN:VCARD','VERSION:3.0',f'FN:{esc_vcf(name)}',f'TEL;TYPE=CELL:{num}','END:VCARD'])
    return ('\r\n'.join(out)+'\r\n').encode('utf-8')

def clean_name(s):
    return re.sub(r'[^\w .-]','_',s,flags=re.UNICODE).strip(' ._')[:60] or 'Contacts'

def chunks(rows,n):
    q,r=divmod(len(rows),n); pos=0
    for i in range(n):
        k=q+(i<r); yield rows[pos:pos+k]; pos+=k

async def sendfile(update,filename,content):
    if len(content)>MAX_BYTES:
        await update.effective_message.reply_text(f'{filename} exceeds 18 MB; split into smaller files.'); return
    bio=io.BytesIO(content); bio.name=filename
    await update.effective_message.reply_document(document=InputFile(bio,filename=filename))

KEYBOARD = ReplyKeyboardMarkup([
    ['/to_vcf', '/to_txt', '/admin', '/manual'],
    ['/add', '/delete', '/renamectc', '/renamefile'],
    ['/merge', '/split', '/count', '/nodup'],
    ['/getname', '/generate', '/getconten', '/setting'],
    ['/status', '/vip', '/referral', '/help'],
    ['/to_csv', '/csv_to_vcf', '/done', '/cancel'],
], resize_keyboard=True, is_persistent=True)

WELCOME = """Bot Admin @Fileagent_21
REDIRECT TO OUR MAIN CHANNEL 🍼👍

https://t.me/+_FCFBVqYDu8zNDY9 🎁
https://t.me/+sAYo43uq1L0wM2M1 😯
https://t.me/+uFnuP2O4BAY5MjY1 🏆
https://t.me/+A42AtiADAKRiNmQ9 🍼
https://t.me/fileagexapi 🌟

JOIN FAST QT'S 🌸🌸

RG-@FILEAGENT_21"""

async def start(update:Update,context:ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(WELCOME, reply_markup=KEYBOARD, disable_web_page_preview=True)
    await update.message.reply_text('📇 Select a tool using the buttons below, or upload a TXT file to convert to CSV. All output files are sent directly, without ZIP.')

async def info_command(update, context):
    cmd = update.message.text.split()[0].split('@')[0].lower()
    st = state.get(update.effective_user.id)
    if cmd == '/admin':
        msg = 'Bot Admin / Support: @Fileagent_21'
    elif cmd == '/setting':
        msg = '⚙️ Settings\nFormats: TXT, CSV, VCF\nMaximum file size: 18 MB\nMaximum output files: 50\nZIP: disabled\nUse /cancel to reset your current operation.'
    elif cmd == '/status':
        msg = '🟢 Bot online\n' + (f"Current operation: {st['mode']} ({st['step']})" if st else 'No active operation.')
    elif cmd == '/vip':
        msg = 'VIP membership is not configured. Contact @Fileagent_21 for details.'
    else:
        me = await context.bot.get_me()
        msg = f'Share the bot with friends:\nhttps://t.me/{me.username}\nReferral rewards are not configured.'
    await update.message.reply_text(msg, reply_markup=KEYBOARD)

HELP='''Commands:
/to_csv — TXT/VCF to CSV (asks name and file count)
/to_vcf — TXT/CSV to VCF (asks 4 questions)
/to_txt — CSV/VCF to TXT
/csv_to_vcf — CSV to VCF
/manual — enter phone numbers manually
/add — combine two files (send one then another)
/delete — remove numbers listed in second file
/renamectc — rename all VCF contacts
/renamefile — change output filename
/merge — merge multiple files (send files, then /done)
/split — split file into N files
/count — count contacts
/nodup — remove duplicate phone numbers
/getname — extract VCF contact names
/generate — generate sequential contact names in VCF
/getconten — extract numbers as TXT
/admin — admin contact
/setting — view settings
/status — current operation
/vip — VIP information
/referral — share bot link
/cancel — cancel current operation

Upload TXT without a command to start TXT→CSV. All results are sent individually, never as ZIP.'''

async def helpcmd(update,context): await update.message.reply_text(HELP, reply_markup=KEYBOARD)

async def setmode(update,context):
    mode=update.message.text.split()[0][1:].split('@')[0].lower()
    state[update.effective_user.id]={'mode':mode,'step':'upload','files':[]}
    prompt={
        'to_csv':'Send TXT or VCF file.', 'to_vcf':'Send TXT or CSV file.',
        'to_txt':'Send CSV or VCF file.', 'csv_to_vcf':'Send CSV file.',
        'manual':'Send phone numbers, one per line.',
        'add':'Send first TXT/CSV/VCF file, then second file.',
        'delete':'Send base file, then a file of numbers to remove.',
        'merge':'Send TXT/CSV/VCF files one at a time; type /done when finished.',
        'split':'Send TXT/CSV/VCF file.', 'count':'Send TXT/CSV/VCF file.',
        'nodup':'Send TXT/CSV/VCF file.', 'getname':'Send VCF file.',
        'getconten':'Send TXT/CSV/VCF file.', 'renamectc':'Send VCF file.',
        'renamefile':'Send TXT/CSV/VCF file.', 'generate':'Send TXT/CSV/VCF file.'}
    await update.message.reply_text(prompt.get(mode,'Send a file.'))

async def cancel(update,context):
    state.pop(update.effective_user.id,None)
    await update.message.reply_text('Cancelled.')

async def done(update,context):
    st=state.get(update.effective_user.id)
    if not st or st['mode']!='merge' or not st['files']:
        await update.message.reply_text('No files to merge. Use /merge first.');return
    rows=[item for f in st['files'] for item in f]
    state.pop(update.effective_user.id,None)
    await sendfile(update,'merged.vcf',vcf_bytes(rows))
    await update.message.reply_text(f'✅ Merged {len(rows)} contacts.')

async def receive_document(update,context):
    doc=update.message.document
    if not doc: return
    filename=doc.file_name or 'file.txt'
    if Path(filename).suffix.lower() not in ('.txt','.csv','.vcf'):
        await update.message.reply_text('Please send .txt, .csv or .vcf.');return
    if doc.file_size and doc.file_size>MAX_BYTES:
        await update.message.reply_text('Maximum input size: 18 MB.');return
    uid=update.effective_user.id
    st=state.get(uid)
    if st is None:
        if filename.lower().endswith('.txt'):
            st={'mode':'to_csv','step':'upload','files':[]};state[uid]=st
        else:
            await update.message.reply_text('Choose /to_csv, /to_txt, /to_vcf or another command first.');return
    if st['step']!='upload':
        await update.message.reply_text('Please answer my current question, or /cancel.');return
    mode=st['mode']; ext=Path(filename).suffix.lower()
    if mode=='to_vcf' and ext not in ('.txt','.csv') or mode=='csv_to_vcf' and ext!='.csv' or mode in ('renamectc','getname') and ext!='.vcf':
        await update.message.reply_text('Wrong file type for this command.');return
    try:
        f=await doc.get_file()
        data=bytes(await f.download_as_bytearray())
        rows=parse_file(filename,data)
    except Exception as exc:
        logging.exception('File read error')
        await update.message.reply_text(f'Could not read file: {type(exc).__name__}');return
    if not rows:
        await update.message.reply_text('No valid phone numbers found. Expected one phone per line, or a Mobile CSV column.');return
    if mode in ('merge','add','delete'):
        st['files'].append(rows)
        if mode=='merge':
            await update.message.reply_text(f'Added {len(rows)} contacts. Send another file or /done.');return
        if len(st['files'])==1:
            await update.message.reply_text(f'First file: {len(rows)} contacts. Send second file.');return
        a,b=st['files'][:2];state.pop(uid,None)
        if mode=='add': result=a+b
        else:
            banned={p for _,p in b};result=[x for x in a if x[1] not in banned]
        await sendfile(update,'result.vcf',vcf_bytes(result))
        await update.message.reply_text(f'✅ {len(result)} contacts.');return
    st['rows']=rows;st['original']=filename
    await update.message.reply_text(f'✅ File received! {len(rows)} valid contacts found.')
    if mode in ('to_csv','to_vcf','csv_to_vcf','renamefile'):
        st['step']='name' if mode!='to_vcf' else 'contact_name'
        await update.message.reply_text('What contact name should I use?' if mode=='to_vcf' else 'What should the output file name be? (without extension)')
    elif mode in ('split',):
        st['step']='file_count';await update.message.reply_text('How many output files do you want? (1–50)')
    elif mode in ('renamectc','generate'):
        st['step']='contact_name';await update.message.reply_text('Enter the contact name prefix (example: Customer).')
    else:
        state.pop(uid,None)
        if mode=='count': await update.message.reply_text(f'📊 Total contacts: {len(rows)}')
        elif mode=='nodup':
            dedup=list(OrderedDict((p,(n,p)) for n,p in rows).values())
            await sendfile(update,'no_duplicates.vcf',vcf_bytes(dedup))
            await update.message.reply_text(f'✅ {len(dedup)} unique; removed {len(rows)-len(dedup)} duplicates.')
        elif mode=='getname':
            names='\n'.join(n for n,p in rows if n)
            await sendfile(update,'contact_names.txt',(names+'\n').encode())
        elif mode in ('to_txt','getconten'):
            await sendfile(update,clean_name(Path(filename).stem)+'.txt',txt_bytes(rows))

async def receive_text(update,context):
    uid=update.effective_user.id;st=state.get(uid)
    if not st: await update.message.reply_text('Send a TXT file for CSV, or use /help.');return
    mode=st['mode'];step=st['step'];value=update.message.text.strip()
    if mode=='manual' and step=='upload':
        rows=parse_txt(value.encode())
        if not rows: await update.message.reply_text('No valid phone numbers.');return
        state.pop(uid,None);await sendfile(update,'manual.vcf',vcf_bytes(rows));return
    if step=='upload': await update.message.reply_text('Please upload a file.');return
    if step=='contact_name':
        st['contact_name']=clean_name(value)
        if mode in ('renamectc','generate'):
            rows=[(f"{st['contact_name']} {i:03d}",p) for i,(_,p) in enumerate(st['rows'],1)]
            state.pop(uid,None);await sendfile(update,'renamed_contacts.vcf',vcf_bytes(rows));return
        st['step']='name';await update.message.reply_text('What should the VCF file name be? (without .vcf)');return
    if step=='name':
        st['name']=clean_name(value)
        if mode=='renamefile':
            ext=Path(st['original']).suffix.lower();rows=st['rows'];state.pop(uid,None)
            data=csv_bytes(rows) if ext=='.csv' else txt_bytes(rows) if ext=='.txt' else vcf_bytes(rows)
            await sendfile(update,st['name']+ext,data);return
        if mode=='to_vcf':
            st['step']='per_file';await update.message.reply_text('How many contacts in EACH VCF file?');return
        st['step']='file_count';await update.message.reply_text('How many CSV files do you want? (1–50)');return
    if step in ('file_count','per_file'):
        try: n=int(value)
        except ValueError: n=0
        if n<1 or (step=='file_count' and n>MAX_OUTPUT):
            await update.message.reply_text('Enter a valid positive number (maximum 50 files).');return
        if step=='per_file':
            st['per_file']=n;st['step']='file_count'
            needed=math.ceil(len(st['rows'])/n)
            await update.message.reply_text(f'How many VCF files do you want? To include all {len(st["rows"])} contacts at {n} per file, enter {needed} (max 50).');return
        rows=st['rows'];total=len(rows)
        if n>total:
            await update.message.reply_text(f'Cannot create {n} nonempty files from {total} contacts. Choose 1–{min(total,MAX_OUTPUT)}.');return
        if mode in ('to_vcf','csv_to_vcf'):
            if mode=='to_vcf':
                per=st['per_file'];needed=math.ceil(total/per)
                if n!=needed:
                    await update.message.reply_text(f'To include all contacts with max {per} per VCF, choose exactly {needed} files.');return
                parts=[rows[i:i+per] for i in range(0,total,per)]
            else: parts=list(chunks(rows,n))
            ext='.vcf'
        elif mode=='to_csv': parts=list(chunks(rows,n));ext='.csv'
        elif mode=='split': parts=list(chunks(rows,n));ext=Path(st['original']).suffix.lower()
        else: await update.message.reply_text('Unexpected mode. /cancel and retry.');return
        base=st.get('name',clean_name(Path(st['original']).stem))
        state.pop(uid,None)
        await update.message.reply_text(f'⏳ Sending {len(parts)} files directly in Telegram...')
        for i,part in enumerate(parts,1):
            fn=f'{base}_{i}{ext}' if len(parts)>1 else f'{base}{ext}'
            content=csv_bytes(part) if ext=='.csv' else txt_bytes(part) if ext=='.txt' else vcf_bytes(part,st.get('contact_name','Contact'))
            await sendfile(update,fn,content)
        await update.message.reply_text(f'✅ Done! {total} contacts, {len(parts)} files. No ZIP.')

COMMANDS=['to_csv','to_vcf','to_txt','csv_to_vcf','manual','add','delete','renamectc','renamefile','merge','split','count','nodup','getname','generate','getconten']
appbot.add_handler(CommandHandler('start',start))
for cmd in ['admin','setting','status','vip','referral']:
    appbot.add_handler(CommandHandler(cmd,info_command))
appbot.add_handler(CommandHandler('help',helpcmd))
appbot.add_handler(CommandHandler('cancel',cancel))
appbot.add_handler(CommandHandler('done',done))
for cmd in COMMANDS: appbot.add_handler(CommandHandler(cmd,setmode))
appbot.add_handler(MessageHandler(filters.Document.ALL,receive_document))
appbot.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,receive_text))

@asynccontextmanager
async def lifespan(app):
    await appbot.initialize()
    await appbot.start()
    await appbot.bot.set_my_commands([BotCommand('start','Start bot'),BotCommand('to_csv','Convert to CSV'),BotCommand('to_vcf','Convert to VCF'),BotCommand('to_txt','Convert to TXT'),BotCommand('csv_to_vcf','CSV to VCF'),BotCommand('merge','Merge files'),BotCommand('split','Split files'),BotCommand('count','Count contacts'),BotCommand('nodup','Remove duplicates'),BotCommand('help','All commands')])
    await appbot.bot.set_webhook(WEBHOOK_URL,secret_token=SECRET,allowed_updates=Update.ALL_TYPES)
    yield
    await appbot.stop()
    await appbot.shutdown()

app=FastAPI(lifespan=lifespan)

@app.get('/')
async def health(): return {'status':'ok','bot':'Contact Converter'}

@app.post('/webhook')
async def webhook(request:Request):
    if request.headers.get('X-Telegram-Bot-Api-Secret-Token') != SECRET:
        raise HTTPException(status_code=403,detail='Forbidden')
    update=Update.de_json(await request.json(),appbot.bot)
    await appbot.update_queue.put(update)
    return {'ok':True}
