"""合併人工判讀的「部門/職稱/姓名」與 PDF 擷取的頁碼、email，輸出 contacts_parsed.tsv。

用法: python3 merge_contacts.py <parse 資料夾> <snips.jsonl> <snips_deep.jsonl> <輸出.tsv> <reports.json> <contacts.jsonl> <contacts_deep.jsonl>
parse 資料夾內：manual_*.tsv（第一輪判讀）、deep_*.tsv（第二輪判讀），格式為 代號\t部門\t職稱\t姓名
"""
import glob, json, re, sys
from contact_phone import find_phone

P, SN, SND, OUT, REPORTS, CT, CTD = sys.argv[1:8]
EMAIL = re.compile(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}')


def load_snips(path):
    return {r['code']: r for r in map(json.loads, open(path, encoding='utf-8'))}


def rows(pattern):
    out = {}
    for f in sorted(glob.glob(f'{P}/{pattern}')):
        for line in open(f, encoding='utf-8'):
            if line.strip():
                code, dept, title, name = (line.rstrip('\n').split('\t') + ['', '', ''])[:4]
                out[code] = (dept.strip(), title.strip(), name.strip())
    return out


snips, snips_deep = load_snips(SN), load_snips(SND)
reports = {r['code']: r for v in json.load(open(REPORTS, encoding='utf-8')).values() for r in v['data']}
pages = {}
for path in (CT, CTD):
    pages[path] = {r['code']: r['pages'][0]['text'] for r in map(json.loads, open(path, encoding='utf-8')) if r['pages']}
merged = {}
# 第二輪先放，第一輪有判讀到的覆蓋（第一輪找到的頁面通常就是報告書聯絡資訊頁）
for src, full, data in ((snips_deep, pages[CTD], rows('deep_*.tsv')), (snips, pages[CT], rows('manual_*.tsv'))):
    for code, (dept, title, name) in data.items():
        s = src.get(code, {})
        text = (s.get('snip') or '').replace('＠', '@')
        found = []
        for e in EMAIL.findall(text):
            # 電話號碼和 email 黏在一起時（02-8227-3827abc@x.com），去掉前面的電話
            e = re.sub(r'^(?:\d+-)+\d+(?=[A-Za-z])', '', e)
            if e.lower() not in found:
                found.append(e.lower())
        filed = re.findall(r'[a-z0-9._%+\-]+@[a-z0-9.\-]+',
                           re.sub(r'\s', '', (reports.get(code, {}).get('contactInfo') or '').replace('＠', '@')).lower())
        phone = find_phone(full.get(code, ''), found or filed)
        merged[code] = (dept, title, name, '; '.join(found), str(s.get('page', '')), phone)

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('code\tdept\ttitle\tname\tpdf_emails\tpage\tphone\n')
    for code in sorted(merged):
        f.write(code + '\t' + '\t'.join(merged[code]) + '\n')
print('contacts', len(merged), 'with name', sum(1 for v in merged.values() if v[2]),
      'with dept', sum(1 for v in merged.values() if v[0]), 'with title', sum(1 for v in merged.values() if v[1]),
      'with phone', sum(1 for v in merged.values() if v[5]))
