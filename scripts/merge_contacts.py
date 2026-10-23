"""合併人工判讀的「部門/職稱/姓名」與 PDF 擷取的頁碼、email，輸出 contacts_parsed.tsv。

用法: python3 merge_contacts.py <parse 資料夾> <snips.jsonl> <snips_deep.jsonl> <輸出.tsv> <reports.json> <contacts.jsonl> <contacts_deep.jsonl>
parse 資料夾內：manual_*.tsv（第一輪判讀）、deep_*.tsv（第二輪判讀），格式為 代號\t部門\t職稱\t姓名
"""
import glob, json, re, sys
from contact_phone import find_phone
from contact_rank import pick, salutation

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
    for code, (dept0, title0, name0) in data.items():
        # 多位聯絡人只留職級最高者、一人多職稱只留最高職稱；被移除的寫進 notes（Excel 用註解標示）
        dept, title, name, notes = pick(dept0, title0, name0)
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
        names = [n for n in re.split(r'[、,，/]', name) if n]
        if len(found) > 1 and names:
            # 有多個 email 時，把緊接在聯絡人姓名後面的那個排第一，F 欄寄送 email 才會和姓名對得上
            lines = [re.sub(r'\s', '', l).replace('＠', '@').lower() for l in full.get(code, '').split('\n')]
            hit = next((i for i, l in enumerate(lines) if any(len(n) >= 2 and re.sub(r'\s', '', n).lower() in l for n in names)), None)
            if hit is not None:
                near = next((e for l in lines[hit:hit + 5] for e in found if e in l), None)
                if near:
                    found = [near] + [e for e in found if e != near]
        phone = find_phone(full.get(code, ''), found or filed, names)
        sur, stitle, unsure = salutation(title, name)
        merged[code] = (dept, title, name, '; '.join(found), str(s.get('page', '')), phone,
                        ' / '.join(notes), sur, stitle, '1' if unsure else '')

with open(OUT, 'w', encoding='utf-8') as f:
    f.write('code\tdept\ttitle\tname\tpdf_emails\tpage\tphone\tnotes\tsal_surname\tsal_title\tsal_unsure\n')
    for code in sorted(merged):
        f.write(code + '\t' + '\t'.join(merged[code]) + '\n')
print('contacts', len(merged), 'with name', sum(1 for v in merged.values() if v[2]),
      'with dept', sum(1 for v in merged.values() if v[0]), 'with title', sum(1 for v in merged.values() if v[1]),
      'with phone', sum(1 for v in merged.values() if v[5]))
