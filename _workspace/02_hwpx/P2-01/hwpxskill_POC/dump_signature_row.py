#!/usr/bin/env python3
"""결재란 표(sec_child[185]) tr[3]/tr[4]의 전체 셀(빈 셀 포함) 구조를 덤프."""
import sys
from lxml import etree

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
}

SECTION_PATH = sys.argv[1]

def local(tag):
    return etree.QName(tag).localname

def full_text(elem):
    return "".join(t.text or "" for t in elem.iter() if local(t.tag) == "t")

tree = etree.parse(SECTION_PATH)
root = tree.getroot()
sec_elem = next(e for e in root.iter() if local(e.tag) == "sec")
sec_children = list(sec_elem)

p185 = sec_children[185]
tbl = next(e for e in p185.iter() if local(e.tag) == "tbl")
rows = tbl.findall("hp:tr", namespaces=NS)
print(f"tbl rowCnt={tbl.get('rowCnt')} colCnt={tbl.get('colCnt')}, 실제 tr={len(rows)}")
for ri, tr in enumerate(rows):
    tcs = tr.findall("hp:tc", namespaces=NS)
    print(f"-- tr[{ri}] (tc {len(tcs)}개) --")
    for ci, tc in enumerate(tcs):
        cell_addr = tc.find("hp:cellAddr", namespaces=NS)
        span = tc.find("hp:cellSpan", namespaces=NS)
        addr_str = f"addr={cell_addr.get('colAddr')},{cell_addr.get('rowAddr')}" if cell_addr is not None else "addr=?"
        span_str = f"span={span.get('colSpan')}x{span.get('rowSpan')}" if span is not None else "span=?"
        txt = full_text(tc)
        print(f"   tc[{ci}] {addr_str} {span_str} text={txt!r}")
