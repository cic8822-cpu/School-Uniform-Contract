#!/usr/bin/env python3
"""tr[3]의 tc[9](교장 서명/직인 칸으로 추정) 원시 XML 구조 덤프."""
import sys
from lxml import etree

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
}

def local(tag):
    return etree.QName(tag).localname

SECTION_PATH = sys.argv[1]
tree = etree.parse(SECTION_PATH)
root = tree.getroot()
sec_elem = next(e for e in root.iter() if local(e.tag) == "sec")
sec_children = list(sec_elem)
p185 = sec_children[185]
tbl = next(e for e in p185.iter() if local(e.tag) == "tbl")
rows = tbl.findall("hp:tr", namespaces=NS)
tr3 = rows[3]
tcs = tr3.findall("hp:tc", namespaces=NS)
tc9 = tcs[9]
tc3 = tcs[3]

print("=== tc[9] (교장 칸 추정) raw XML ===")
print(etree.tostring(tc9, pretty_print=True).decode()[:3000])
print()
print("=== tc[3] (담당자 이름 '○○○' 칸) raw XML ===")
print(etree.tostring(tc3, pretty_print=True).decode()[:3000])
