"""從永續報告書聯絡資訊頁的文字，找出聯絡人電話（含分機），排除傳真。"""
import re

KEY = re.compile(r'聯絡|聯繫|連絡|意見|回饋|窗口|洽詢|聯系')
TEL_LABEL = re.compile(r'電\s*話|專\s*線|TEL|Tel|tel|分\s*機|Phone|phone|聯絡方式')
FAX = re.compile(r'傳\s*真|FAX|Fax|fax')
PHONE = re.compile(
    r'(?:\(\+?886[\s-]?0?\d{1,2}\)|\+?\s?886[\s-]?\(?0?\d{1,2}\)?|\+?\s?86[\s-]?\d{2,4}|\(0\d{1,2}\)|0\d{1,2}(?=[\s-]?\d{3,4}[\s-]?\d{3,4}))'
    r'[\s-]?\d{3,4}[\s-]?\d{3,4}'
    r'(?:\s*(?:#|＃|分機|ext\.?|Ext\.?|EXT\.?|轉|\*|-)\s*:?\s*\d{1,6}(?:\s*[、,]\s*#?\s*\d{1,6})*)?'
    r'|09\d{2}[\s-]?\d{3}[\s-]?\d{3}')
FW = str.maketrans('（）＃：－０１２３４５６７８９＋　', '()#:-0123456789+ ')


def _norm(s):
    return s.translate(FW)


def _anchor(lines, emails):
    flat = lambda s: re.sub(r'\s', '', s.replace('＠', '@')).lower()
    for i, l in enumerate(lines):
        if any(e in flat(l) for e in emails):
            return i
    for i, l in enumerate(lines):
        if '@' in l or '＠' in l:
            return i
    for i, l in enumerate(lines):
        if KEY.search(l):
            return i
    return None


def find_phone(text, emails=()):
    """回傳聯絡段落中的電話；優先「電話/專線/TEL」那一行，跳過傳真。"""
    lines = [_norm(l.strip()) for l in (text or '').split('\n') if l.strip()]
    a = _anchor(lines, [e.lower() for e in emails])
    if a is None:
        return ''
    start = a
    for i in range(a, max(-1, a - 12), -1):
        if KEY.search(lines[i]):
            start = i
    # 依原文順序：沒標示「電話/傳真」的表格，通常電話寫在傳真前面
    win = list(range(max(0, start - 1), min(len(lines), a + 8)))
    labelled, plain = [], []
    for i in win:
        l = lines[i]
        if FAX.search(l) and not TEL_LABEL.search(l.split('傳')[0].split('FAX')[0].split('Fax')[0]):
            continue
        seg = re.split(r'傳\s*真|FAX|Fax|fax', l)[0]
        for m in PHONE.finditer(seg):
            p = re.sub(r'\s+', ' ', m.group(0)).strip(' -')
            if len(re.sub(r'\D', '', p)) < 8:
                continue
            (labelled if TEL_LABEL.search(seg) else plain).append(p)
    return (labelled or plain or [''])[0]
