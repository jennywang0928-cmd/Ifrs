"""PDF 文字裡 email 常和後面的字黏在一起（info@x.com.twemail、ct100@hinet.netchapter），把多出來的字切掉。"""
import re

# email 最後一段合法的結尾（報告書裡實際出現過的都要列進來）
TLDS = {'tw', 'com', 'net', 'org', 'edu', 'gov', 'cn', 'jp', 'hk', 'sg', 'co', 'asia', 'us', 'biz', 'global', 'info',
        'io', 'my', 'th', 'vn', 'de', 'uk', 'kr', 'holdings', 'ltd', 'money', 'ai', 'tv', 'mo', 'ph', 'id', 'in', 'au',
        'ca', 'nl', 'fr', 'it', 'eu', 'cc', 'me', 'app', 'tech', 'group'}


def trim_tld(e):
    head, _, last = e.rpartition('.')
    if not head or last.lower() in TLDS:
        return e
    m = re.match(r'(com|net|org|tw)', last, re.I)
    return f'{head}.{m.group(1)}' if m else e


def bad_tld(e):
    return e.rpartition('.')[2].lower() not in TLDS
