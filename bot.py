import httpx
import os, re, csv, io, math, asyncio, logging, tempfile
from openpyxl import Workbook, load_workbook
import xlrd, xlwt
from pathlib import Path
from collections import OrderedDict
from telegram import ReplyKeyboardMarkup, Update, BotCommand, InputFile, InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardRemove, MenuButtonCommands
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters, CallbackQueryHandler, ApplicationHandlerStop
from fastapi import FastAPI, Request, HTTPException
from contextlib import asynccontextmanager

logging.basicConfig(level=logging.WARNING)
TOKEN = os.environ['BOT_TOKEN']
WEBHOOK_URL = os.environ['WEBHOOK_URL'].rstrip('/')
SECRET = os.environ['WEBHOOK_SECRET']
MAX_BYTES = 18 * 1024 * 1024
MAX_OUTPUT = 50
OWNER_ID = int(os.environ.get("OWNER_ID", "0"))
# Comma-separated numeric channel IDs or public @usernames; bot must be an admin.
BROADCAST_CHANNELS = [v.strip() for v in os.environ.get("BROADCAST_CHANNEL_IDS", "").split(",") if v.strip()]
# Required joins are configured only through Render environment variables.
JOIN_TARGETS = [
    (f'Channel {i}', os.environ.get(f'CHANNEL_{i}_ID', '').strip(), os.environ.get(f'CHANNEL_{i}_LINK', '').strip())
    for i in range(1, 4)
] + [('Group', os.environ.get('GROUP_ID', '').strip(), os.environ.get('GROUP_LINK', '').strip())]
JOIN_TARGETS = [(name, chat_id, link) for name, chat_id, link in JOIN_TARGETS if chat_id]
# The same 3 channels + 1 group are also broadcast destinations.
BROADCAST_CHANNELS = [chat_id for _, chat_id, _ in JOIN_TARGETS]

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
    return spreadsheet_rows(list(csv.reader(s)))

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
    if ext=='.xlsx': return parse_xlsx(data)
    if ext=='.xls': return parse_xls(data)
    raise ValueError('Supported files: .txt, .csv, .vcf, .xls, .xlsx.')

def spreadsheet_rows(table):
    if not table: return []
    first=[str(v or '').strip().lower() for v in table[0]]
    number_idx=next((i for i,x in enumerate(first) if x in ('mobile','phone','number','telephone','tel','phone number','mobile number')),None)
    name_idx=next((i for i,x in enumerate(first) if x in ('name','contact','contact name','full name')),None)
    start=1 if number_idx is not None else 0
    if number_idx is None:
        # Prefer the first column with recognizable phone numbers.
        scores=[sum(bool(PHONE.fullmatch(str(row[i]).strip())) for row in table[:30] if i<len(row)) for i in range(max(map(len,table)))]
        number_idx=max(range(len(scores)),key=lambda i:scores[i]) if scores else 0
    out=[]
    for row in table[start:]:
        if number_idx>=len(row): continue
        raw=str(row[number_idx]).strip()
        if raw.endswith('.0') and raw[:-2].isdigit(): raw=raw[:-2]
        if PHONE.fullmatch(raw):
            name=str(row[name_idx]).strip() if name_idx is not None and name_idx<len(row) and row[name_idx] is not None else ''
            out.append((name,normalize(raw)))
    return out

def parse_xlsx(data):
    wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True)
    try:
        return [item for sheet in wb.worksheets for item in spreadsheet_rows([list(row) for row in sheet.iter_rows(values_only=True)])]
    finally: wb.close()

def parse_xls(data):
    wb=xlrd.open_workbook(file_contents=data)
    return [item for sheet in wb.sheets() for item in spreadsheet_rows([sheet.row_values(i) for i in range(sheet.nrows)])]

def xlsx_bytes(rows):
    wb=Workbook(); ws=wb.active; ws.title='Contacts'; ws.append(['Name','Mobile'])
    for name,phone in rows:
        ws.append([name,phone]); ws.cell(ws.max_row,2).number_format='@'
    f=io.BytesIO();wb.save(f);return f.getvalue()

def xls_bytes(rows):
    if len(rows)>65535: raise ValueError('XLS supports at most 65,535 contacts per sheet. Use XLSX or split the file.')
    wb=xlwt.Workbook();ws=wb.add_sheet('Contacts');ws.write(0,0,'Name');ws.write(0,1,'Mobile')
    for i,(name,phone) in enumerate(rows,1):ws.write(i,0,name);ws.write(i,1,phone)
    f=io.BytesIO();wb.save(f);return f.getvalue()

def output_bytes(rows,ext,base='Contact'):
    if ext=='.vcf':return vcf_bytes(rows,base)
    return {'.csv':csv_bytes,'.txt':txt_bytes,'.xlsx':xlsx_bytes,'.xls':xls_bytes}[ext](rows)

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

DB=os.environ.get('SUPABASE_URL','').rstrip('/')+'/rest/v1/'
KEY=os.environ.get('SUPABASE_KEY','')
HEAD={'apikey':KEY,'Content-Type':'application/json','Prefer':'return=representation'}
if not KEY.startswith('sb_secret_'): HEAD['Authorization']='Bearer '+KEY
async def db(method,table,params=None,data=None):
    async with httpx.AsyncClient(timeout=15) as c:
        r=await c.request(method,DB+table,params=params,json=data,headers=HEAD)
        r.raise_for_status(); return r.json() if r.content else []
async def put(table,data):
    async with httpx.AsyncClient(timeout=15) as c:
        r=await c.post(DB+table,json=data,headers={**HEAD,'Prefer':'resolution=merge-duplicates,return=minimal'}); r.raise_for_status()
async def settings(): return {r['key']:r['value'] for r in await db('GET','bot_settings')}
async def setting(k,v): await put('bot_settings',{'key':k,'value':str(v)})
async def user(uid):
    if not KEY or not os.environ.get('SUPABASE_URL'): return {}
    r=await db('GET','bot_users',{'user_id':f'eq.{uid}'}); return r[0] if r else {}
LANGS={'en':'🇬🇧 English','hi':'🇮🇳 हिन्दी','zh':'🇨🇳 简体中文','id':'🇮🇩 Bahasa Indonesia'}
WORDS={
'en':['Welcome','Choose your language','Join Official Channel','Verify & Unlock','Access granted','Choose a tool','Access locked','Send your file','Cancelled'],
'hi':['स्वागत है','अपनी भाषा चुनें','चैनल जॉइन करें','सत्यापित करें','प्रवेश मिल गया','टूल चुनें','प्रवेश बंद है','फाइल भेजें','रद्द किया गया'],
'zh':['欢迎','选择语言','加入官方频道','验证并解锁','访问已授权','选择工具','访问锁定','请发送文件','已取消'],
'id':['Selamat datang','Pilih bahasa','Gabung channel resmi','Verifikasi & buka akses','Akses diberikan','Pilih alat','Akses terkunci','Kirim file','Dibatalkan']}
def words(context): return WORDS.get(context.user_data.get('lang','en'),WORDS['en'])
def keyboard():
    c=['/to_vcf','/to_txt','/to_csv','/to_xls','/to_xlsx','/manual','/add','/delete','/renamectc','/renamefile','/merge','/split','/count','/nodup','/getname','/generate','/getconten','/csv_to_vcf','/language','/tools','/cancel','/help']
    return ReplyKeyboardMarkup([c[i:i+4] for i in range(0,len(c),4)],resize_keyboard=True,one_time_keyboard=True)
async def language(update,context):
    await update.effective_message.reply_text('🌍 '+words(context)[1],reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton(v,callback_data='lang:'+k)] for k,v in LANGS.items()]))
async def langpick(update,context):
    q=update.callback_query
    await q.answer()
    l=q.data.split(':')[1]
    if l not in LANGS:return
    context.user_data['lang']=l
    confirmation={'en':'Language changed to English','hi':'भाषा हिन्दी में बदल दी गई है।','zh':'语言已切换为简体中文。','id':'Bahasa berhasil diubah ke Bahasa Indonesia.'}
    await q.message.reply_text('✅ '+confirmation[l])
    try:
        if KEY and os.environ.get('SUPABASE_URL'):
            await put('bot_users',{'user_id':update.effective_user.id,'language':l})
    except Exception:
        logging.warning('Language persistence failed')
        await q.message.reply_text('⚠️ Language saved for this session only. Database unavailable.')
    await welcome(update,context)
async def missing_joins(update, context):
    if OWNER_ID and update.effective_user.id == OWNER_ID:
        return []
    missing = []
    for label, chat_id, link in JOIN_TARGETS:
        try:
            member = await context.bot.get_chat_member(chat_id=int(chat_id) if chat_id.lstrip('-').isdigit() else chat_id, user_id=update.effective_user.id)
            status = member.status
            joined = status in ('creator', 'administrator', 'member') or (status == 'restricted' and getattr(member, 'is_member', False))
            if not joined:
                missing.append((label, link))
        except Exception:
            logging.exception('Cannot verify membership for %s', label)
            missing.append((label, link))
    return missing

def join_markup(missing):
    buttons = [[InlineKeyboardButton(('👥 ' if label == 'Group' else '📢 ') + 'Join ' + label, url=link)] for label, link in missing if link.startswith(('https://t.me/', 'http://t.me/'))]
    buttons.append([InlineKeyboardButton('✅ Verify All', callback_data='verify')])
    return InlineKeyboardMarkup(buttons)

async def welcome(update,context):
    w=words(context)
    missing=await missing_joins(update,context)
    locked=bool(missing)
    text=(f"👑 FILEAGENT • FILE CONVERTER\n✨ {w[0]}, {update.effective_user.first_name}!\n"
          "━━━━━━━━━━━━━━\n💎 TXT ⇄ CSV ⇄ VCF ⇄ XLS ⇄ XLSX\n"
          "📂 Merge, Split, Rename, Remove Duplicates\n🌍 4 Languages\n"
          + ('🔐 Join all 3 channels and 1 group, then tap Verify All.' if locked else '⚡ Access granted!'))
    markup=join_markup(missing) if locked else None
    image_path = Path(__file__).with_name('welcome.jpg')
    if image_path.is_file():
        try:
            with image_path.open('rb') as photo:
                await update.effective_message.reply_photo(photo=photo, caption=text[:1024], reply_markup=markup)
        except Exception:
            logging.exception('Could not send welcome image')
            await update.effective_message.reply_text(text, reply_markup=markup)
    else:
        await update.effective_message.reply_text(text, reply_markup=markup)
    if not locked: await tools_menu(update,context)

async def tools_menu(update,context):
    if await missing_joins(update,context):
        await welcome(update,context)
        return
    await update.effective_message.reply_text('💎 '+words(context)[5],reply_markup=keyboard())

async def start(update,context):
    state.pop(update.effective_user.id,None)
    if 'lang' not in context.user_data:
        try:
            u=await user(update.effective_user.id)
            if u.get('language') in LANGS: context.user_data['lang']=u['language']
        except Exception: logging.warning('Language lookup failed')
    if KEY and os.environ.get('SUPABASE_URL'):
        try:
            await put('bot_users',{'user_id':update.effective_user.id,'language':context.user_data.get('lang','en')})
        except Exception: logging.exception('Unable to register user for broadcasts')
    if 'lang' not in context.user_data: await language(update,context)
    else: await welcome(update,context)

async def verify(update,context):
    await update.callback_query.answer()
    await welcome(update,context)

async def membership_gate(update, context):
    # Block normal users' tools until all four joins are verified. Owner bypasses.
    if not update.effective_chat or update.effective_chat.type != 'private':
        raise ApplicationHandlerStop
    if OWNER_ID and update.effective_user.id == OWNER_ID:
        return
    if update.callback_query:
        if update.callback_query.data in ('verify',) or update.callback_query.data.startswith('lang:'):
            return
        if await missing_joins(update,context):
            await update.callback_query.answer('Join all required chats first', show_alert=True)
            raise ApplicationHandlerStop
        return
    if not update.message:
        raise ApplicationHandlerStop
    cmd=(update.message.text or '').split(' ')[0].split('@')[0]
    if cmd in ('/start','/language'):
        return
    missing=await missing_joins(update,context)
    if missing:
        await welcome(update,context)
        raise ApplicationHandlerStop

HELP='''Commands:
/to_csv — TXT/VCF/XLS/XLSX to CSV (asks name and file count)
/to_vcf — TXT/CSV/XLS/XLSX to VCF (asks 3 questions; automatic file count)
/to_txt — CSV/VCF/XLS/XLSX to TXT
/to_xls — TXT/CSV/VCF/XLS/XLSX/XLSX to Excel XLS
/to_xlsx — TXT/CSV/VCF/XLS/XLSX/XLS to Excel XLSX
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
/cancel — cancel current operation

Upload TXT without a command to start TXT→CSV.'''

async def helpcmd(update,context): await update.message.reply_text(HELP)

async def setmode(update,context):
    mode=update.message.text.split()[0][1:].split('@')[0].lower()
    # Keep csv_to_vcf separate to enforce CSV input.
    state[update.effective_user.id]={'mode':mode,'step':'upload','files':[]}
    prompt={
        'to_csv':{'en':'Send TXT, VCF, XLS or XLSX file.','hi':'TXT या VCF फाइल भेजें।','zh':'请发送 TXT 或 VCF 文件。','id':'Kirim file TXT atau VCF.'}[context.user_data.get('lang','en')], 'to_vcf':{'en':'Send TXT, CSV, XLS or XLSX file.','hi':'TXT या CSV फाइल भेजें।','zh':'请发送 TXT 或 CSV 文件。','id':'Kirim file TXT atau CSV.'}[context.user_data.get('lang','en')],
        'to_txt':{'en':'Send CSV, VCF, XLS or XLSX file.','hi':'CSV या VCF फाइल भेजें।','zh':'请发送 CSV 或 VCF 文件。','id':'Kirim file CSV atau VCF.'}[context.user_data.get('lang','en')], 'csv_to_vcf':'Send CSV file.',
        'to_xls':'Send TXT, CSV, VCF or XLSX file.',
        'to_xlsx':'Send TXT, CSV, VCF or XLS file.',
        'manual':'Send phone numbers, one per line.',
        'add':'Send first TXT/CSV/VCF/XLS/XLSX file, then second file.',
        'delete':'Send base file, then a file of numbers to remove.',
        'merge':'Send TXT/CSV/VCF/XLS/XLSX files one at a time; type /done when finished.',
        'split':'Send TXT/CSV/VCF/XLS/XLSX file.', 'count':'Send TXT/CSV/VCF/XLS/XLSX file.',
        'nodup':'Send TXT/CSV/VCF/XLS/XLSX file.', 'getname':'Send VCF file.',
        'getconten':'Send TXT/CSV/VCF/XLS/XLSX file.', 'renamectc':'Send VCF file.',
        'renamefile':'Send TXT/CSV/VCF/XLS/XLSX file.', 'generate':'Send TXT/CSV/VCF/XLS/XLSX file.'}
    await update.message.reply_text('💎 '+words(context)[7]+'\n'+prompt.get(mode,'Send a file.'),reply_markup=ReplyKeyboardRemove())

async def cancel(update,context):
    state.pop(update.effective_user.id,None)
    await update.message.reply_text(words(context)[8],reply_markup=keyboard())

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
    if Path(filename).suffix.lower() not in ('.txt','.csv','.vcf','.xls','.xlsx'):
        await update.message.reply_text('Please send .txt, .csv, .vcf, .xls or .xlsx.');return
    if doc.file_size and doc.file_size>MAX_BYTES:
        await update.message.reply_text('Maximum input size: 18 MB.');return
    uid=update.effective_user.id
    st=state.get(uid)
    if st is None:
        if filename.lower().endswith('.txt'):
            st={'mode':'to_csv','step':'upload','files':[]};state[uid]=st
        else:
            await update.message.reply_text('Choose /to_csv, /to_txt, /to_vcf, /to_xls, /to_xlsx or another command first.');return
    if st['step']!='upload':
        await update.message.reply_text('Please answer my current question, or /cancel.');return
    mode=st['mode']; ext=Path(filename).suffix.lower()
    if mode in ('to_vcf','csv_to_vcf') and ext not in (('.csv',) if mode=='csv_to_vcf' else ('.txt','.csv','.xls','.xlsx')) or mode=='csv_to_vcf' and ext!='.csv' or mode in ('renamectc','getname') and ext!='.vcf':
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
    if mode=='renamefile': st['raw']=data
    await update.message.reply_text(f'✅ File received! {len(rows)} valid contacts found.')
    if mode in ('to_csv','to_vcf','to_xls','to_xlsx','csv_to_vcf','renamefile'):
        st['step']='name' if mode not in ('to_vcf','csv_to_vcf') else 'contact_name'
        await update.message.reply_text('What contact name should I use?' if mode in ('to_vcf','csv_to_vcf') else 'What should the output file name be? (without extension)')
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
            data=st['raw']
            await sendfile(update,st['name']+ext,data);return
        if mode in ('to_vcf','csv_to_vcf'):
            st['step']='per_file';await update.message.reply_text('How many contacts in EACH VCF file?');return
        st['step']='file_count'
        kind='XLS' if mode=='to_xls' else 'XLSX' if mode=='to_xlsx' else 'CSV'
        await update.message.reply_text(f'How many {kind} files do you want? (1–50)');return
    if step in ('file_count','per_file'):
        try: n=int(value)
        except ValueError: n=0
        if n<1 or (step=='file_count' and n>MAX_OUTPUT):
            await update.message.reply_text('Enter a valid positive number (maximum 50 files).');return
        rows=st['rows'];total=len(rows)
        if step=='per_file':
            needed=math.ceil(total/n)
            if needed>MAX_OUTPUT:
                await update.message.reply_text(f'Maximum {MAX_OUTPUT} files. Enter at least {math.ceil(total/MAX_OUTPUT)} contacts per file.');return
            rows=[(f"{st['contact_name']} {i:03d}",p) for i,(_,p) in enumerate(rows,1)]
            parts=[rows[i:i+n] for i in range(0,total,n)]
            ext='.vcf'
        elif n>total:
            await update.message.reply_text(f'Cannot create {n} nonempty files from {total} contacts. Choose 1–{min(total,MAX_OUTPUT)}.');return
        elif mode in ('to_csv','to_xls','to_xlsx'):
            parts=list(chunks(rows,n));ext={'to_csv':'.csv','to_xls':'.xls','to_xlsx':'.xlsx'}[mode]
        elif mode=='split': parts=list(chunks(rows,n));ext=Path(st['original']).suffix.lower()
        else: await update.message.reply_text('Unexpected mode. /cancel and retry.');return
        base=st.get('name',clean_name(Path(st['original']).stem))
        state.pop(uid,None)
        await update.message.reply_text(f'⏳ Sending {len(parts)} files directly in Telegram...')
        for i,part in enumerate(parts,1):
            fn=f'{base}_{i}{ext}' if len(parts)>1 else f'{base}{ext}'
            try:
                content=output_bytes(part,ext,st.get('contact_name','Contact'))
            except ValueError as exc:
                await update.message.reply_text(str(exc));return
            await sendfile(update,fn,content)
        await update.message.reply_text(f'✅ Done! {total} contacts, {len(parts)} files.')


# Owner-only broadcasts. Only users who have started the bot can receive messages.
# Channel broadcasts need channel IDs and bot posting permissions; channel joining is NOT required.
broadcast_drafts = {}
broadcast_running = set()

async def broadcast_command(update, context):
    uid = update.effective_user.id
    if not OWNER_ID or uid != OWNER_ID or update.effective_chat.type != 'private':
        await update.effective_message.reply_text('⛔ This command is restricted to the bot owner.')
        return
    if uid in broadcast_running:
        await update.message.reply_text('⏳ A broadcast is already running.')
        return
    broadcast_drafts[uid] = {'step':'target'}
    markup = InlineKeyboardMarkup([
        [InlineKeyboardButton('👥 All bot users', callback_data='bc:target:users')],
        [InlineKeyboardButton('📢 All 3 channels + group', callback_data='bc:target:channels')],
        [InlineKeyboardButton('🎯 One channel / group', callback_data='bc:pick')],
        [InlineKeyboardButton('🌍 Users + channels', callback_data='bc:target:both')],
        [InlineKeyboardButton('❌ Cancel', callback_data='bc:cancel')]
    ])
    await update.message.reply_text('👑 BROADCAST PANEL\nChoose recipients:', reply_markup=markup)

async def broadcast_callback(update, context):
    q = update.callback_query
    uid = update.effective_user.id
    if not OWNER_ID or uid != OWNER_ID or update.effective_chat.type != 'private':
        await q.answer('Owner only', show_alert=True)
        return
    data=q.data.split(':')
    if data[1]=='cancel':
        broadcast_drafts.pop(uid,None)
        await q.answer()
        await q.message.reply_text('❌ Broadcast cancelled.')
        return
    if data[1]=='pick':
        if not BROADCAST_CHANNELS:
            await q.answer('Set channel/group IDs on Render first', show_alert=True); return
        await q.answer()
        await q.message.reply_text('🎯 Choose one destination:', reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton(label, callback_data=f'bc:one:{i}')]
            for i,(label, _, _) in enumerate(JOIN_TARGETS)
        ]))
        return
    if data[1]=='one':
        try:
            i=int(data[2]); label, chat_id, _=JOIN_TARGETS[i]
        except (ValueError, IndexError):
            await q.answer('Invalid destination',show_alert=True); return
        broadcast_drafts[uid]={'step':'content','target':'one','destination':chat_id,'destination_label':label}
        await q.answer()
        await q.message.reply_text(f'✍️ Send a message for {label} (text, photo, video, GIF or document).')
        return
    if data[1]=='target':
        if uid in broadcast_running:
            await q.answer('Broadcast already running', show_alert=True); return
        target=data[2]
        if target in ('channels','both') and not BROADCAST_CHANNELS:
            await q.answer('Set BROADCAST_CHANNEL_IDS on Render first',show_alert=True); return
        broadcast_drafts[uid]={'step':'content','target':target}
        await q.answer()
        await q.message.reply_text('✍️ Send your broadcast message now (text, photo, video, animation or document).\nUse /cancel to stop.')
        return
    if data[1]=='send':
        draft=broadcast_drafts.get(uid)
        if not draft or draft.get('step')!='confirm' or uid in broadcast_running:
            await q.answer('No pending broadcast',show_alert=True);return
        broadcast_drafts.pop(uid,None)
        broadcast_running.add(uid)
        await q.answer()
        await q.message.reply_text('⏳ Broadcasting started. I will send a delivery report when finished.')
        context.application.create_task(run_broadcast(context.application,uid,draft))

async def broadcast_content(update,context):
    uid=update.effective_user.id
    draft=broadcast_drafts.get(uid)
    if not OWNER_ID or uid!=OWNER_ID or not draft or draft.get('step')!='content':return False
    msg=update.effective_message
    if msg.text and msg.text.startswith('/'):
        return False
    if not (msg.text or msg.photo or msg.video or msg.animation or msg.document):
        await msg.reply_text('Please send text, photo, video, GIF or document.')
        return True
    draft.update(step='confirm',source_chat_id=msg.chat_id,message_id=msg.message_id)
    await msg.reply_text('👀 Broadcast preview:')
    await context.bot.copy_message(chat_id=uid,from_chat_id=msg.chat_id,message_id=msg.message_id)
    target={'users':'all registered bot users','channels':'all configured channels','both':'users and channels','one':draft.get('destination_label','one chat')}[draft['target']]
    await msg.reply_text(f'📢 Send this to {target}?',reply_markup=InlineKeyboardMarkup([
        [InlineKeyboardButton('✅ Confirm & Send',callback_data='bc:send')],
        [InlineKeyboardButton('❌ Cancel',callback_data='bc:cancel')]
    ]))
    return True

async def fetch_broadcast_users():
    if not KEY or not os.environ.get('SUPABASE_URL'):
        raise RuntimeError('Supabase is not configured. Cannot load registered bot users.')
    ids=set()
    offset=0
    while True:
        rows=await db('GET','bot_users',{'select':'user_id','order':'user_id.asc','limit':'1000','offset':str(offset)})
        for row in rows:
            try: ids.add(int(row['user_id']))
            except (TypeError,ValueError,KeyError): pass
        if len(rows)<1000:break
        offset+=1000
    return ids

async def run_broadcast(application,uid,draft):
    sent=failed=0
    errors=[]
    try:
        destinations=[]
        if draft['target'] in ('users','both'):
            destinations.extend(sorted(await fetch_broadcast_users()))
        if draft['target']=='one':
            destinations.append(int(draft['destination']))
        if draft['target'] in ('channels','both'):
            destinations.extend(int(v) if v.lstrip('-').isdigit() else v for v in BROADCAST_CHANNELS)
        # No duplicate destination, and don't broadcast to owner twice.
        destinations=list(dict.fromkeys(destinations))
        for dest in destinations:
            try:
                await application.bot.copy_message(chat_id=dest,from_chat_id=draft['source_chat_id'],message_id=draft['message_id'])
                sent+=1
            except Exception as exc:
                # Respect Telegram flood-control RetryAfter where possible.
                from telegram.error import RetryAfter
                if isinstance(exc,RetryAfter):
                    await asyncio.sleep(float(exc.retry_after)+1)
                    try:
                        await application.bot.copy_message(chat_id=dest,from_chat_id=draft['source_chat_id'],message_id=draft['message_id'])
                        sent+=1
                    except Exception as retry_exc:
                        failed+=1
                        if len(errors)<3:errors.append(f'{dest}: {type(retry_exc).__name__}')
                else:
                    failed+=1
                    if len(errors)<3:errors.append(f'{dest}: {type(exc).__name__}')
            await asyncio.sleep(0.06)
        report=f'📊 Broadcast complete\n✅ Delivered: {sent}\n❌ Failed: {failed}\n🎯 Total: {len(destinations)}'
        if errors: report+='\nSample errors: '+'; '.join(errors)
        await application.bot.send_message(uid,report)
    except Exception as exc:
        logging.exception('Broadcast failed')
        await application.bot.send_message(uid,f'❌ Broadcast stopped: {type(exc).__name__}: {str(exc)[:250]}')
    finally:
        broadcast_running.discard(uid)

async def broadcast_message_router(update,context):
    if await broadcast_content(update,context):return
    if update.message.document:
        await receive_document(update,context)
    elif update.message.text:
        await receive_text(update,context)

async def broadcast_cancel(update,context):
    broadcast_drafts.pop(update.effective_user.id,None)
    await cancel(update,context)

COMMANDS=['to_csv','to_vcf','to_txt','to_xls','to_xlsx','csv_to_vcf','manual','add','delete','renamectc','renamefile','merge','split','count','nodup','getname','generate','getconten']
from telegram.ext import TypeHandler
appbot.add_handler(TypeHandler(Update, membership_gate), group=-1)
appbot.add_handler(CommandHandler('start',start))
appbot.add_handler(CallbackQueryHandler(verify,pattern='^verify$'))
appbot.add_handler(CallbackQueryHandler(broadcast_callback,pattern=r'^bc:'))
appbot.add_handler(CommandHandler('broadcast',broadcast_command))
appbot.add_handler(CallbackQueryHandler(langpick,pattern='^lang:'))
appbot.add_handler(CommandHandler('help',helpcmd))
appbot.add_handler(CommandHandler('language',language))
appbot.add_handler(CommandHandler('tools',tools_menu))
appbot.add_handler(CommandHandler('cancel',broadcast_cancel))
appbot.add_handler(CommandHandler('done',done))
for cmd in COMMANDS: appbot.add_handler(CommandHandler(cmd,setmode))
appbot.add_handler(MessageHandler((filters.Document.ALL | filters.TEXT | filters.PHOTO | filters.VIDEO | filters.ANIMATION) & ~filters.COMMAND,broadcast_message_router))

@asynccontextmanager
async def lifespan(app):
    await appbot.initialize()
    await appbot.start()
    await appbot.bot.set_my_commands([BotCommand('start','Start Fileagent'), *[BotCommand(cmd, cmd.replace('_',' ').capitalize()) for cmd in COMMANDS], BotCommand('language','Change language'), BotCommand('tools','Show keyboard'), BotCommand('done','Finish merge'), BotCommand('cancel','Cancel operation'), BotCommand('help','All commands')])
    await appbot.bot.set_chat_menu_button(menu_button=MenuButtonCommands())
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
