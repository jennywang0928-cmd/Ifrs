"""把證交所/櫃買中心公司基本資料 + 永續報告書申報資料，整理成寄送名單 Excel。

用法: python3 build_excel.py <資料夾> <輸出.xlsx>
資料夾內需有 L.json (上市公司基本資料)、O.json (上櫃公司基本資料)、reports.json (永續報告書申報資料)
"""
import json, re, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
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

rows = []
for code in sorted(set(companies) | set(reports)):
    c = companies.get(code, {})
    r = reports.get(code)
    esg = emails(r.get('contactInfo')) if r else []
    co = emails(c.get('co_email'))
    if esg:
        send, source = esg, '永續報告書聯絡窗口'
    elif co:
        send = co
        source = '公司信箱（報告書未填 email）' if r else '公司信箱（查無本年度永續報告書）'
    else:
        send, source = [], '查無 email，需人工補'
    rows.append([
        code,
        c.get('short') or (r or {}).get('shortName', ''),
        c.get('name') or (r or {}).get('name', ''),
        c.get('market') or (r or {}).get('market', ''),
        (r or {}).get('sector') or '',
        '; '.join(send),
        source,
        '; '.join(esg),
        '; '.join(co),
        ((r or {}).get('contactInfo') or '').strip(),
        (r or {}).get('reportingInterval') or '',
        (r or {}).get('twDocLink') or '',
        (c.get('web') or '').strip().strip('　'),
        c.get('phone', ''),
    ])

headers = ['公司代號', '公司簡稱', '公司名稱', '市場', '產業', '建議寄送 email', 'email 來源',
           '永續報告書聯絡 email', '公司信箱 (發言人/IR)', '永續報告書聯絡資訊原文', '報告期間',
           '永續報告書網址', '公司網址', '總機電話']

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
for i, row in enumerate(rows, start=2):
    if row[6].startswith('公司信箱'):
        ws.cell(i, 7).fill = warn
    elif row[6].startswith('查無'):
        ws.cell(i, 7).fill = bad
widths = [9, 12, 30, 6, 14, 36, 26, 36, 30, 50, 22, 40, 34, 16]
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
    ['永續報告書聯絡窗口 email', cnt('永續報告書')],
    ['改用公司信箱（報告書未填 email 或未申報）', cnt('公司信箱')],
    ['查無 email，需人工補', cnt('查無')],
    [],
    ['資料來源'],
    ['永續報告書聯絡資訊', '證交所 ESG InfoHub「永續報告書」申報資料（2025 年度報告，2026 年申報），欄位：聯絡資訊'],
    ['公司信箱', '證交所 / 櫃買中心 Open API「公司基本資料」，欄位：電子郵件信箱'],
    [],
    ['欄位說明'],
    ['建議寄送 email', '優先用永續報告書上的聯絡 email；沒有的話改用公司信箱'],
    ['email 來源', '黃色＝改用公司信箱；橘色＝查無 email，需人工補'],
    ['永續報告書聯絡資訊原文', '公司申報時填寫的原文，有些只留電話或網址，可參考後人工補 email'],
    [],
    ['寄送提醒', '建議用 Mailchimp、Brevo 等大量寄信工具（Gmail 每日約 500 封上限），信中附「取消訂閱」連結與寄件單位資訊'],
]:
    info.append(line)
info['A1'].font = Font(bold=True, size=14)
for r in (8, 12):
    info.cell(r, 1).font = Font(bold=True)
info.column_dimensions['A'].width = 40
info.column_dimensions['B'].width = 90

wb.save(OUT)
print('rows', n, 'esg', cnt('永續報告書'), 'company', cnt('公司信箱'), 'none', cnt('查無'))
