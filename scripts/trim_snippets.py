"""把擷取到的聯絡資訊頁文字，裁成 email 附近的幾行，方便判讀部門/職稱/姓名。

用法: python3 trim_snippets.py <reports.json> <contacts.jsonl> <輸出.jsonl>
"""
import json, re, sys

reports = {r['code']: r for v in json.load(open(sys.argv[1], encoding='utf-8')).values() for r in v['data']}
KEY = re.compile(r'聯絡|聯繫|連絡|意見|回饋|窗口|洽詢|聯系')
out = open(sys.argv[3], 'w', encoding='utf-8')
for line in open(sys.argv[2], encoding='utf-8'):
    r = json.loads(line)
    if not r['pages']:
        continue
    text = r['pages'][0]['text']
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    ci = (reports[r['code']].get('contactInfo') or '').replace('＠', '@')
    emails = [e.lower() for e in re.findall(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+', re.sub(r'\s', '', ci))]
    flat = lambda s: re.sub(r'\s', '', s.replace('＠', '@')).lower()
    anchor = next((i for i, l in enumerate(lines) if any(e in flat(l) for e in emails)), None)
    if anchor is None:
        anchor = next((i for i, l in enumerate(lines) if '@' in l or '＠' in l), None)
    if anchor is None:
        anchor = next((i for i, l in enumerate(lines) if KEY.search(l)), 0)
    # 往上找最近的「聯絡/意見」標題，最多 12 行
    start = anchor
    for i in range(anchor, max(-1, anchor - 12), -1):
        if KEY.search(lines[i]):
            start = i
    start = max(0, min(start, anchor - 4))
    snip = ' / '.join(lines[start:anchor + 3])
    out.write(json.dumps({'code': r['code'], 'page': r['pages'][0]['page'], 'snip': snip[:700]}, ensure_ascii=False) + '\n')
