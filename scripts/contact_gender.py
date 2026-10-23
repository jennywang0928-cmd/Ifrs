"""只有姓名、沒有可用職稱的聯絡人，判斷稱呼「先生／小姐」。

只用兩種可靠來源（不靠中文名字用字猜）：
1. 報告書原文就寫了「陳麗年 小姐」「潘信宏 先生」
2. 聯絡人 email 的英文名（jeffrey.lee → 先生），且 email 的姓氏拼音要對得上中文姓（避免是別人的信箱）
"""
import re
from pypinyin import pinyin, Style

MALE = set('''aaron adam adrian albert alan allen alvin andrew andy angus anson anthony arthur austin ben benjamin benson
bill billy bob boris brian bruce bruno bryan calvin carl carson charles charlie chester colin dale dallas dan daniel danny
darren david dean dennis denny derek dick dominic don douglas duncan dylan eason eddie eddy edward edwin elvis eric ernest
evan felix frank frankie fred frederick gary gavin george gilbert gordon grant greg gregory hank hanson harry harvey henry
herman howard ian isaac ivan jack jacky jackson jacob jake james jason jay jeff jeffery jeffrey jeremy jerome jerry jesse jim
jimmy joe joel john johnny johnson jonas jonathan joseph josh joshua justin keith ken kenneth kenny kent kevin kyle lance larry
lawrence lawson leo leon leonard lewis louis lucas luke marco marcus mark martin marvin mason matt matthew max michael mike
morris nathan neil nelson nick nicholas nolan norman oliver oscar owen patrick paul peter philip phillip ralph randy ray
raymond rex richard rick ricky rob robert rodney roger ronald ronnie ross roy rudy russell ryan samson samuel scott sean
shawn simon spencer stan stanley stephen steve steven stewart stuart ted terry thomas tim timothy toby tom tommy tony
travis tristan troy tyler victor vincent wallace walter warren wayne wesley will william willy wilson winston zack'''.split())
FEMALE = set('''abby ada adela agnes aileen alice alicia alison amanda amber amy ana andrea angel angela angelina angie anita
ann anna anne annie april ariel ashley audrey ava barbara bella belinda betty beverly bonnie brenda bridget candy carina
carol caroline carrie cathy catherine cecilia celia celine charlene charlotte chelsea cherry cheryl chloe christina christine
cindy claire clara claudia coco connie cora crystal cynthia daisy dana daphne debbie deborah diana diane doris dora edith
eileen elaine eleanor elena eliza elizabeth ella ellen elsa emily emma erica esther eunice eva eve evelyn fanny fiona flora
florence frances gigi gillian gina grace gloria hannah hazel heather helen helena ida irene iris ivy jacqueline jade jane
janet janice jasmine jean jeanne jenny jennifer jessica jessie jill joan joanna joanne jocelyn josephine joy joyce judy julia
julie june karen karina kate katherine kathy katie kelly kiki kitty krystal lara laura lauren leah lena lillian lily linda lisa
liz lois lora lorraine louisa lucia lucy luna lydia mabel maggie mandy maria marian marie marina mary maureen maya megan
melinda melissa melody mia michelle mimi miranda miriam molly mona monica nancy natalie nicole nina nora olivia pamela
patricia peggy penny phoebe phyllis polly priscilla queenie rachel rebecca regina rita rosa rose rosie ruby ruth sabrina
sally samantha sammi sandra sandy sara sarah selina serena sharon sheila shelley sherry shirley sophia sophie stacey stacy
stella stephanie sue susan susie sylvia tammy tanya teresa tiffany tina tracy una valerie vanessa vera veronica vicky
victoria viola vivian vivien wendy winnie yvonne zoe'''.split())

# 常見姓氏的非漢語拼音寫法（台灣威妥瑪、粵語、閩南語等）
EXTRA = {
    '李': 'lee lie', '王': 'wong ong', '陳': 'chan tan', '黃': 'wong hwang ng', '張': 'cheung chang', '劉': 'lau lew liou',
    '吳': 'woo ng goh', '林': 'lim lam', '蔡': 'choi chua tsai', '楊': 'yeung young yong', '何': 'ho', '胡': 'hu wu woo',
    '周': 'chow chou', '許': 'hsu hui khoo shu', '郭': 'kwok kuo', '梁': 'leung liang', '羅': 'law lo lou', '鄧': 'teng tang',
    '謝': 'tse hsieh shieh hsie', '葉': 'yeh yip ip', '蘇': 'su so soo', '游': 'yu yew yiu', '簡': 'chien kan jian',
    '余': 'yu yee yue', '徐': 'hsu tsui shu', '呂': 'lu lui leu lyu lue', '邱': 'chiu chiou khoo', '江': 'chiang kong',
    '沈': 'shen sim', '鍾': 'chung', '曾': 'tseng tsang zeng', '莊': 'chuang', '施': 'shih sze', '翁': 'weng ong wong',
    '洪': 'hung hong', '賴': 'lai', '魏': 'wei', '唐': 'tang', '傅': 'fu', '侯': 'hou', '彭': 'peng pang', '廖': 'liao liu',
    '詹': 'chan chan', '康': 'kang', '高': 'kao ko', '馬': 'ma', '朱': 'chu', '連': 'lien', '温': 'wen', '溫': 'wen',
}
# 拼音 → 威妥瑪常見寫法
WG = [('zh', 'ch'), ('ch', 'ch'), ('x', 'hs'), ('q', 'ch'), ('j', 'ch'), ('z', 'ts'), ('c', 'ts'), ('g', 'k'),
      ('k', 'k'), ('d', 't'), ('b', 'p'), ('r', 'j')]


def _wade(py):
    out = {py}
    w = py
    for a, b in WG:
        if w.startswith(a):
            w = b + w[len(a):]
            break
    w = re.sub(r'ong$', 'ung', w)
    w = re.sub(r'ian$', 'ien', w)
    w = re.sub(r'ie$', 'ieh', w)
    w = re.sub(r'uo$', 'o', w)
    w = re.sub(r'^(hs|ch)u$', r'\1u', w)
    out |= {w, w.replace('ung', 'ong'), py.replace('v', 'u'), w.replace('v', 'u')}
    return out


def surname_variants(sur):
    v = set()
    for ch in sur:
        for py in pinyin(ch, style=Style.NORMAL, heteronym=True)[0]:
            v |= _wade(py)
        v |= set(EXTRA.get(ch, '').split())
    if len(sur) == 2:
        v |= {a + b for a in surname_variants(sur[0]) for b in surname_variants(sur[1])}
    return v


ALL_SURNAMES = set()
for _ch in EXTRA:
    ALL_SURNAMES |= surname_variants(_ch)


def from_report(text, name):
    """報告書原文「姓名 + 先生/小姐/女士」。"""
    t = re.sub(r'\s', '', text or '')
    m = re.search(re.escape(name) + r'[^，,。；;：:\n]{0,8}?(先生|小姐|女士)', t)
    return m.group(1) if m else ''


# 只用來「排除矛盾」，不用來猜性別：email 判斷結果和中文名明顯用字衝突時，寧可不填
STRONG_MALE = set('雄豪強傑偉剛勇龍泉輝威武彬斌')
STRONG_FEMALE = set('婷娟芳美淑雯玲惠珍秀麗琪瑩妤嫻媛姿娥蓉鳳萍燕')


def conflicts(name, g):
    given = name[1:]
    if g == '小姐' and any(c in STRONG_MALE for c in given):
        return True
    return g == '先生' and any(c in STRONG_FEMALE for c in given)


def from_email(email, sur):
    """email 英文名判斷先生/小姐；回傳 (稱謂, 英文名)，判斷不了回傳 ('', '')。"""
    local = (email or '').split('@')[0].lower()
    toks = [t for t in re.split(r'[._\-\d]+', local) if t]
    variants = surname_variants(sur)
    for i, t in enumerate(toks):
        first, rest = '', []
        if t in MALE or t in FEMALE:
            first, rest = t, toks[:i] + toks[i + 1:]
        else:
            # 黏在一起的「cherrywu」：英文名 + 姓
            for n in sorted(MALE | FEMALE, key=len, reverse=True):
                if t.startswith(n) and t[len(n):] in variants:
                    first, rest = n, [t[len(n):]] + toks[:i] + toks[i + 1:]
                    break
        if not first:
            continue
        others = [r for r in rest if len(r) > 1]
        if others and not any(r in variants for r in others) and any(r in ALL_SURNAMES for r in others):
            return '', ''  # email 上是別的姓，不是這位聯絡人
        return ('先生' if first in MALE else '小姐'), first
    return '', ''
