"""從永續報告書聯絡資訊頁的文字，找出聯絡人電話（含分機），排除傳真。"""
import re

KEY = re.compile(r'聯絡|聯繫|連絡|意見|回饋|窗口|洽詢|聯系')
TEL_LABEL = re.compile(r'電\s*話|專\s*線|TEL|Tel|tel|分\s*機|Phone|phone|聯絡方式')
FAX = re.compile(r'傳\s*真|FAX|Fax|fax')
SKIP = re.compile(r'客服|消費者|申訴|檢舉|0800')
PHONE = re.compile(
    r'(?:\(\+?886\)[\s-]?0?\d{1,2}|\(\+?886[\s-]?0?\d{1,2}\)|\+?\s?886[\s-]?\(?0?\d{1,2}\)?|\+?\s?86[\s-]?\d{2,4}|\(0\d{1,2}\)|0\d{1,2}(?=[\s-]?\d{3,4}[\s-]?\d{3,4}))'
    r'[\s-]?\d{3,4}[\s-]?\d{3,4}'
    r'(?:\s*(?:#|＃|分機|ext\.?|Ext\.?|EXT\.?|轉|\*|-)\s*:?\s*\d{1,6}(?:\s*[、,]\s*#?\s*\d{1,6})*)?'
    r'|09\d{2}[\s-]?\d{3}[\s-]?\d{3}')
FW = str.maketrans('（）＃：－０１２３４５６７８９＋　', '()#:-0123456789+ ')


def _norm(s):
    # 「（ 02 ）2798-9169」括號內多了空格
    return re.sub(r'\(\s*(\+?\d{1,4})\s*\)', r'(\1)', s.translate(FW))


def _squeeze_digits(text):
    """PDF 文字裡數字中間被插入空格：「(02)2 202- 0152 分機：20 0 6」→「(02)2202-0152 分機：2006」"""
    text = re.sub(r'(?<=\b\d) (?=\d)|(?<=\d) (?=\d\b)', '', text)
    return re.sub(r'(?<=-) (?=\d)', '', text)


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


def _is_label(l):
    """只有「電話：」「傳真」這類標示、沒有數字的短行。"""
    return len(l) <= 12 and not re.search(r'\d', l) and bool(TEL_LABEL.search(l) or FAX.search(l))


def find_phone(text, emails=(), names=()):
    """先用原文找；找不到時才試「去掉數字間空格」的版本，避免把正常的「+886 2 xxxx」擠在一起。"""
    return _find_phone(text, emails, names) or _find_phone(_squeeze_digits(text or ''), emails, names)


def _find_phone(text, emails=(), names=()):
    """回傳聯絡段落中的電話，排除傳真與客服專線。

    有聯絡人姓名時，取緊接在姓名後面的電話；否則取有「電話」標示、離 email 最近的；
    都沒有標示時依原文順序取第一個（表格多半電話在傳真前面）。
    """
    lines = [_norm(l.strip()) for l in (text or '').split('\n') if l.strip()]
    a = _anchor(lines, [e.lower() for e in emails])
    if a is None:
        return ''
    start = a
    for i in range(a, max(-1, a - 12), -1):
        if KEY.search(lines[i]):
            start = i
    # 往上至少看 6 行：有些表格先列「地址、聯絡人、電話、傳真、email」的值，標題列在後面
    win = list(range(max(0, min(start - 1, a - 6)), min(len(lines), a + 8)))
    label_lines = [i for i in win if _is_label(lines[i])]
    phone_lines = [i for i in win if PHONE.search(lines[i]) and not _is_label(lines[i])]
    # 標示寫在號碼後面（例如「02-xxx / 電話： / 02-yyy / 傳真：」）
    labels_after = bool(label_lines and phone_lines and phone_lines[0] < label_lines[0]
                        and phone_lines[0] + 1 in label_lines)

    cands = []  # (行號, 該行第幾個, 號碼, 是否有電話標示)
    for i in win:
        l = lines[i]
        if _is_label(l) or SKIP.search(l):
            continue
        near = lines[i + 1] if labels_after and i + 1 < len(lines) else (lines[i - 1] if i > 0 else '')
        near = near if _is_label(near) else ''
        seg = re.split(r'傳\s*真|FAX|Fax|fax', l)[0]
        if not seg.strip() or (not TEL_LABEL.search(seg) and FAX.search(near) and not TEL_LABEL.search(near)):
            continue
        for k, m in enumerate(PHONE.finditer(seg)):
            p = re.sub(r'\s+', ' ', m.group(0)).strip(' -')
            if len(re.sub(r'\D', '', p)) >= 8:
                cands.append((i, k, p, bool(TEL_LABEL.search(seg) or TEL_LABEL.search(near))))
    if not cands:
        return ''
    flat = lambda x: re.sub(r'\s', '', x)
    for n in names:
        n = flat(n)
        hit = next((i for i in win if len(n) >= 2 and n in flat(lines[i])), None)
        if hit is not None:
            after = [c for c in cands if hit <= c[0] <= hit + 4]
            if after:
                return min(after)[2]
    labelled = [c for c in cands if c[3]]
    if labelled:
        return min(labelled, key=lambda c: (abs(c[0] - a), c[0], c[1]))[2]
    return min(cands)[2]
