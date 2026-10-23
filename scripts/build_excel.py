"""把證交所/櫃買中心公司基本資料 + 永續報告書申報資料，整理成寄送名單 Excel。

用法: python3 build_excel.py <資料夾> <輸出.xlsx>
資料夾內需有 L.json (上市公司基本資料)、O.json (上櫃公司基本資料)、reports.json (永續報告書申報資料)，
可另放 contacts_parsed.tsv（merge_contacts.py 產出的報告書聯絡人部門/職稱/姓名）
"""
import json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from contact_phone import PHONE, FW  # noqa: E402
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

DATA, OUT = sys.argv[1], sys.argv[2]
EMAIL = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}')


def emails(text):
    """抓出文字中的 email；處理全形＠、email 中間被打了空白或控制字元的情況。"""
    if not text:
        return []
    t = text.replace('＠', '@').replace('．', '.')
    found = EMAIL.findall(t)
    # 只壓掉 ASCII 字元之間的空白/控制字元，避免把中文說明黏進去
    squeezed = EMAIL.findall(re.sub(r'(?<=[\x21-\x7e])[\s\x00-\x1f]+(?=[\x21-\x7e])', '', t))
    if not found:
        found = squeezed
    else:
        # "jack .lee@x.com" 原文只抓得到 lee@x.com：改用壓掉空白後較完整的版本（只補前段帳號）
        # 只在原文帳號明顯被截斷（以 . _ - 開頭）時才補，避免把前面的分機號碼黏進來
        found = [next((q for q in squeezed if q.endswith(e) and '@' not in q[:-len(e)]), e)
                 if e[:1] in '._-' else e for e in found]
    out = []
    for e in found:
        e = e.strip('.').lower()
        if e not in out:
            out.append(e)
    return out


def load(name):
    with open(f'{DATA}/{name}', encoding='utf-8') as f:
        return json.load(f)


companies = {}
for r in load('L.json'):
    companies[r['公司代號']] = dict(code=r['公司代號'], short=r['公司簡稱'], name=r['公司名稱'], market='上市',
                                  co_email=r.get('電子郵件信箱', ''), web=r.get('網址', ''), phone=r.get('總機電話', ''))
for r in load('O.json'):
    companies[r['SecuritiesCompanyCode']] = dict(code=r['SecuritiesCompanyCode'], short=r['CompanyAbbreviation'],
                                               name=r['CompanyName'], market='上櫃', co_email=r.get('EmailAddress', ''),
                                               web=r.get('WebAddress', ''), phone=r.get('Telephone', ''))

reports = {}
for key, v in load('reports.json').items():
    for r in v['data']:
        r['market'] = '上市' if key.endswith('_0') else '上櫃'
        reports[r['code']] = r

contacts = {}
if os.path.exists(f'{DATA}/contacts_parsed.tsv'):
    with open(f'{DATA}/contacts_parsed.tsv', encoding='utf-8') as f:
        next(f)
        for line in f:
            code, dept, title, name, pdf_emails, page, phone, notes, sal_sur, sal_title, sal_unsure, sal_note = \
                line.rstrip('\n').split('\t')
            contacts[code] = dict(dept=dept, title=title, name=name, pdf_emails=pdf_emails, page=page, phone=phone,
                                  notes=notes, sal_sur=sal_sur, sal_title=sal_title, sal_unsure=sal_unsure,
                                  sal_note=sal_note)


def filed_phone(text):
    """申報資料「聯絡資訊」欄若只留電話，取出來當備援。"""
    m = PHONE.search((text or '').translate(FW))
    return re.sub(r'\s+', ' ', m.group(0)).strip(' -') if m else ''

rows, notes_by_row = [], []
for code in sorted(set(companies) | set(reports)):
    c = companies.get(code, {})
    r = reports.get(code)
    esg = emails(r.get('contactInfo')) if r else []
    co = emails(c.get('co_email'))
    ct = contacts.get(code, {})
    pdf = emails(ct.get('pdf_emails'))
    # 寄送 email 以「報告書內文」為主，才會和 H~J 欄的聯絡人對得上；內文沒有才用申報資料
    if pdf:
        send = pdf[:1] if ct.get('name') else [next((e for e in pdf if e in esg), pdf[0])]
        source = '報告書內文（與聯絡人對應）'
    elif esg:
        send, source = esg, '申報資料聯絡 email'
    elif co:
        send = co
        source = '公司信箱（報告書未填 email）' if r else '公司信箱（查無本年度永續報告書）'
    else:
        send, source = [], '查無 email，需人工補'
    # W 欄需確認事項：🔴 必須處理、🟡 建議確認（網址無效在做超連結時再補）
    check = []
    if not send:
        check.append('🔴 查無 email，需人工補')
    elif source.startswith('公司信箱'):
        check.append('🟡 改用公司總信箱，建議查有無 ESG 窗口')
    if len(pdf) > 1 and not ct.get('name'):
        # 有聯絡人姓名時，F 欄已取緊接在姓名後的 email，不必再確認
        check.append('🟡 報告書列多個 email，請確認 F 欄')
    if ct.get('sal_unsure'):
        check.append('🟡 四字姓名，請確認稱呼姓')
    rows.append([
        code,
        c.get('short') or (r or {}).get('shortName', ''),
        c.get('name') or (r or {}).get('name', ''),
        c.get('market') or (r or {}).get('market', ''),
        (r or {}).get('sector') or '',
        '; '.join(send),
        source,
        ct.get('dept', ''),
        ct.get('title', ''),
        ct.get('name', ''),
        ct.get('sal_sur', ''),
        ct.get('sal_title', ''),
        '; '.join(esg),
        '; '.join(pdf),
        '; '.join(co),
        ((r or {}).get('contactInfo') or '').strip(),
        int(ct['page']) if ct.get('page') else None,
        (r or {}).get('reportingInterval') or '',
        (r or {}).get('twDocLink') or '',
        (c.get('web') or '').strip().strip('　'),
        ct.get('phone') or filed_phone((r or {}).get('contactInfo')),
        c.get('phone', ''),
        '；'.join(check),
    ])
    notes_by_row.append((ct.get('notes', ''), ct.get('sal_note', '')))

headers = ['公司代號', '公司簡稱', '公司名稱', '市場', '產業', '建議寄送 email（對應聯絡人）', 'email 來源',
           '聯絡人部門', '聯絡人職稱', '聯絡人姓名', '稱呼姓', '稱呼職稱',
           '永續報告書聯絡 email（申報資料）', '報告書內文聯絡 email', '公司信箱 (發言人/IR)', '永續報告書聯絡資訊原文', '聯絡資訊在報告書第幾頁', '報告期間',
           '永續報告書網址', '公司網址', '聯絡人電話（報告書）', '公司總機電話', '需確認事項']

wb = Workbook()
ws = wb.active
ws.title = '寄送名單'
ws.append(headers)
for row in rows:
    ws.append([re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', v) if isinstance(v, str) else v for v in row])
head_fill = PatternFill('solid', fgColor='1F4E78')
for cell in ws[1]:
    cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = head_fill
    cell.alignment = Alignment(vertical='center', wrap_text=True)
warn = PatternFill('solid', fgColor='FFF2CC')
bad = PatternFill('solid', fgColor='F8CBAD')
# Q、R 欄（永續報告書網址、公司網址）做成可點的超連結
for col in (headers.index('永續報告書網址') + 1, headers.index('公司網址') + 1):
    for i in range(2, len(rows) + 2):
        cell = ws.cell(i, col)
        # 修正公司申報時打錯的網址：中間空格、全形冒號、多打 h、少打冒號、瀏覽器外掛前綴
        url = re.sub(r'\s+', '', cell.value or '').replace('：', ':')
        url = re.sub(r'^chrome-extension://[^/]+/', '', url)
        url = re.sub(r'^h+ttp', 'http', url, flags=re.I)
        url = re.sub(r'^(https?)//', r'\1://', url, flags=re.I)
        if not url or '.' not in url or not re.match(r'(https?:)?//|www\.|[\w-]+\.', url, re.I):
            if col == headers.index('永續報告書網址') + 1 and cell.value:
                w = ws.cell(i, headers.index('需確認事項') + 1)
                w.value = '；'.join(x for x in (w.value, '🔴 永續報告書網址無效') if x)
            continue
        cell.value = url
        if not re.match(r'https?://', url, re.I):
            url = 'https://' + url.lstrip('/')
        cell.hyperlink = url
        cell.style = 'Hyperlink'
for i, row in enumerate(rows, start=2):
    if row[6].startswith('公司信箱'):
        ws.cell(i, 7).fill = warn
    elif row[6].startswith('查無'):
        ws.cell(i, 7).fill = bad
# 被移除的聯絡人、職稱用註解標在 I、J 欄
for i, (note, sal_note) in enumerate(notes_by_row, start=2):
    if sal_note:
        ws.cell(i, headers.index('稱呼職稱') + 1).comment = Comment(sal_note, '整理說明', width=240, height=50)
    for part in filter(None, note.split(' / ')):
        col = 10 if part.startswith('報告書另列聯絡人') else 9
        cell = ws.cell(i, col)
        text = (cell.comment.text + '\n' if cell.comment else '') + part
        cell.comment = Comment(text, '整理說明', width=320, height=90)
# 需確認事項依等級上色
wcol = headers.index('需確認事項') + 1
for i in range(2, len(rows) + 2):
    v = ws.cell(i, wcol).value or ''
    if '🔴' in v:
        ws.cell(i, wcol).fill = bad
    elif '🟡' in v:
        ws.cell(i, wcol).fill = warn
widths = [9, 12, 30, 6, 14, 36, 26, 28, 24, 20, 10, 14, 36, 36, 30, 50, 12, 22, 40, 34, 28, 16, 44]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = 'C2'
ws.auto_filter.ref = ws.dimensions

n = len(rows)
cnt = lambda p: sum(1 for r in rows if r[6].startswith(p))
info = wb.create_sheet('說明')
for line in [
    ['上市櫃公司永續報告書聯絡 email 名單'],
    [],
    ['公司總數', n],
    ['寄送 email 用報告書內文（與聯絡人對應）', cnt('報告書內文')],
    ['寄送 email 用申報資料（報告書內文沒寫 email）', cnt('申報資料')],
    ['有聯絡人部門', sum(1 for r in rows if r[7])],
    ['有聯絡人職稱', sum(1 for r in rows if r[8])],
    ['有聯絡人姓名', sum(1 for r in rows if r[9])],
    ['有聯絡人電話', sum(1 for r in rows if r[headers.index('聯絡人電話（報告書）')])],
    ['改用公司信箱（報告書未填 email 或未申報）', cnt('公司信箱')],
    ['查無 email，需人工補', cnt('查無')],
    [],
    ['資料來源'],
    ['永續報告書聯絡資訊', '證交所 ESG InfoHub「永續報告書」申報資料（2025 年度報告，2026 年申報），欄位：聯絡資訊'],
    ['公司信箱', '證交所 / 櫃買中心 Open API「公司基本資料」，欄位：電子郵件信箱'],
    ['聯絡人部門／職稱／姓名', '永續報告書 PDF 內「聯絡資訊／意見回饋」段落，逐份判讀；報告書沒寫的就留空'],
    [],
    ['欄位說明'],
    ['建議寄送 email（對應聯絡人）', '優先用報告書內文聯絡 email（N 欄，和聯絡人同一段）；內文沒有才用申報資料（M 欄），都沒有改用公司信箱'],
    ['永續報告書網址／公司網址', 'S、T 欄可直接點擊開啟；公司申報時打錯的網址格式（空格、全形冒號等）已自動修正'],
    ['聯絡人電話（報告書）', '報告書聯絡段落的電話（含分機，已排除傳真）；報告書沒寫的改用申報資料聯絡資訊中的電話；都沒有就留空，可看 V 欄總機'],
    ['email 來源', '黃色＝改用公司信箱；橘色＝查無 email，需人工補'],
    ['聯絡人（H～J 欄）', '報告書列多位聯絡人時只留職級最高者；一人多個職稱只留最高的（例：發言人暨總經理 → 總經理）。被移除的寫在 I、J 欄的註解（儲存格右上角紅色三角形）'],
    ['職級順序', '董事長 > 總經理/執行長 > 副總經理/財務長/永續長 > 協理/總監/處長 > 經理/副理 > 課長/主任/組長 > 專員/秘書/特助 > 發言人/公司治理主管/執行秘書等角色'],
    ['稱呼姓／稱呼職稱（K、L 欄）', '信件開頭用「K+L」，例：王經理、劉小姐、廖執行秘書；英文名只填名字（信件用「Wendy 您好」）；兩欄都空白時信件用「您好」。只有姓名沒有職稱的，報告書原文寫明或 email 英文名可判斷時才填「先生／小姐」（依 email 判斷的在 L 欄有註解），其餘維持「您好」'],
    ['需確認事項（W 欄）', '🔴 必須處理（紅底）、🟡 建議確認（黃底）；空白代表不需處理'],
    ['報告書內文聯絡 email', '報告書 PDF 聯絡資訊段落裡寫的 email；少數公司申報資料填的 email 和報告書內文不同，部門/職稱/姓名是依報告書內文'],
    ['永續報告書聯絡資訊原文', '公司申報時填寫的原文，有些只留電話或網址，可參考後人工補 email'],
    [],
    ['寄送提醒', '建議用 Mailchimp、Brevo 等大量寄信工具（Gmail 每日約 500 封上限），信中附「取消訂閱」連結與寄件單位資訊'],
]:
    info.append(line)
info['A1'].font = Font(bold=True, size=14)
for row in info.iter_rows(min_row=2):
    if row[0].value in ('資料來源', '欄位說明'):
        row[0].font = Font(bold=True)
info.column_dimensions['A'].width = 40
info.column_dimensions['B'].width = 90

wb.save(OUT)
print('rows', n, 'pdf', cnt('報告書內文'), 'filed', cnt('申報資料'), 'company', cnt('公司信箱'), 'none', cnt('查無'))
