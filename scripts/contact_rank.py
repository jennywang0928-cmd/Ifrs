"""多位聯絡人只留職級最高者、一人多職稱只留最高職稱，並產生信件稱呼（姓 + 職稱）。"""
import re

# 數字越小職級越高；比對順序很重要（「董事長特助」要先判成特助，「副總經理」要先於「總經理」）
RANKS = [
    (7, r'特助|特別助理|助理'),
    (8, r'代理發言人|發言人|公司治理主管|治理主管|主任委員|總召集人|召集人|總幹事|總執行|管理代表|ESG窗口|窗口'),
    (8, r'執行秘書'),
    (1, r'副董事長|董事長'),
    (3, r'副總經理|副總裁|副總|財務長|永續長|營運長|技術長|策略長'),
    (2, r'總經理|執行長|CEO|總裁'),
    (4, r'協理|總監|副處長|處長|部長|廠長'),
    (5, r'副廠長|資深經理|經理|副理'),
    (6, r'課長|主任|襄理|副組長|組長|主管'),
    (7, r'專員|管理師|秘書|幹事|股務|編輯|Administrator|代表'),
]
# 日本姓氏（報告書出現過的）
JAPANESE = ('澄川',)
COMPOUND = JAPANESE + ('歐陽', '司馬', '上官', '諸葛', '張簡', '范姜', '簡林', '東方', '西門', '南宮', '司徒', '夏侯', '公孫', '慕容')
SALUTE = re.compile(r'^(.+?)(先生|小姐|女士)$')
SPLIT = re.compile(r'[、，,／/；;]|暨|兼任|兼|（|）|\(|\)')


def rank_of(token):
    for r, pat in RANKS:
        m = re.search(pat, token)
        if m:
            return r, m.group(0)
    return 9, ''


def best_title(titles):
    """一個人的多個職稱，留職級最高的；回傳 (職級, 職稱原文)。"""
    toks = [t.strip() for t in SPLIT.split(titles or '') if t and t.strip() and t.strip() not in ('任',)]
    toks = [re.sub(r'^兼?任', '', t) for t in toks]
    if not toks:
        return 9, ''
    scored = [(rank_of(t)[0], i, t) for i, t in enumerate(toks)]
    r, _, t = min(scored)
    return r, t


def split_people(dept, title, name):
    """把「名字、名字」「職稱；職稱」拆成每個人的 (部門, 職稱, 姓名)。"""
    names = [n.strip() for n in re.split(r'[、,，]', name or '') if n.strip()] or ['']
    tgroups = [t.strip() for t in re.split(r'[；;]', title or '')]
    if len(tgroups) == 1 and len(names) > 1:
        parts = [t.strip() for t in re.split(r'[、,，]', tgroups[0]) if t.strip()]
        tgroups = parts if len(parts) == len(names) else tgroups * len(names)
    if len(tgroups) != len(names):
        tgroups = (tgroups + [tgroups[-1]] * len(names))[:len(names)]
    dgroups = [d.strip() for d in re.split(r'[；;]', dept or '')]
    if len(dgroups) != len(names):
        dgroups = [dept or ''] * len(names)
    return list(zip(dgroups, tgroups, names))


def pick(dept, title, name, prefer=None):
    """回傳挑出的 (部門, 職稱, 姓名, 被移除的說明)。prefer：指定要留下第幾位（依 email 比對出的本人）。"""
    people = split_people(dept, title, name)
    scored = []
    for i, (d, t, n) in enumerate(people):
        r, bt = best_title(t)
        # 姓名本身帶職稱（謝副理、閔副總經理）時也拿來比
        if r == 9 and n:
            r2, _ = rank_of(n)
            r = min(r, r2)
        scored.append((r, i, d, t, bt, n))
    r, i, d, t, bt, n = scored[prefer] if prefer is not None else min(scored)
    bt = _full_title(bt)
    notes = []
    if len(people) > 1:
        others = [f'{p[2]} {p[1]}'.strip() for j, p in enumerate(people) if j != i]
        notes.append('報告書另列聯絡人：' + '、'.join(o for o in others if o))
    if t and bt != t:
        notes.append(f'原職稱：{t}')
    return d, bt, n, notes


def surname(full):
    for c in COMPOUND:
        if full.startswith(c) and len(full) >= 3:
            return c, False
    # 四字姓名多半是冠夫姓或複姓，不確定姓氏怎麼切
    return full[0], len(full) >= 4


def salutation(title, name):
    """回傳 (稱呼姓, 稱呼職稱, 是否需確認)。依規則：
    劉小姐 → 劉/小姐；王小名 經理 → 王/經理；Wendy Juan → Wendy/空；
    沒有人名、沒有職稱、或只有發言人/公司治理主管這類角色 → 空/空（信件用「您好」）；
    執行秘書 → 王/執行秘書。
    """
    n = (name or '').strip()
    if not n:
        return '', '', False
    if re.match(r'[A-Za-z]', n):
        return n.split()[0], '', False
    m = SALUTE.match(n)
    if m:
        return m.group(1), m.group(2), False
    r, kw = rank_of(n)
    if r != 9 and kw and n.endswith(n[n.find(kw):]) and len(n) - len(n[n.find(kw):]) <= 2:
        # 報告書原文就寫「謝副理」「閔副總經理」：姓 + 職稱
        sur, t = n[:n.find(kw)], n[n.find(kw):]
        prefix = re.search(r'(資深|執行|副)+$', sur)
        if prefix:
            sur, t = sur[:prefix.start()], prefix.group(0) + t
        if 1 <= len(sur) <= 2:
            return sur, _full_title(t), False
    r, kw = rank_of(title or '')
    if r == 9 or (r == 8 and '執行秘書' not in (title or '')):
        return '', '', False
    sur, unsure = surname(n)
    t = '執行秘書' if '執行秘書' in title else _full_title(_core(title, kw))
    if t in NOT_SALUTATION:
        return '', '', False
    return sur, t, unsure


def _core(title, kw):
    """「財務部經理」→「經理」、「資深副總」→「資深副總經理」；保留資深/執行/副這類前綴。"""
    i = title.find(kw)
    pre = re.search(r'(資深|執行|副)*$', title[:i]).group(0)
    return pre + kw


def _full_title(t):
    t = re.sub(r'副總$', '副總經理', t)
    return '特別助理' if t == '特助' else t


# 這些不適合接在姓後面當信件稱呼（王Administrator、張股務），改走先生／小姐判斷或「您好」
NOT_SALUTATION = {'Administrator', '股務', '編輯', '代表', '助理'}
