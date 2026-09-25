#!/usr/bin/env python3
"""F-007 구간 구조 분석: 표 내부 텍스트(제목표 학교명, 결재란 등) 위치 파악.
읽기 전용 분석만 수행. 원본을 수정하지 않음.
"""
import sys
from pathlib import Path
from lxml import etree

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
}

SECTION_PATH = sys.argv[1] if len(sys.argv) > 1 else "unpacked_ref/Contents/section0.xml"
START_ANCHOR = "[ 6 ] 교복 학교주관구매 구매 요청(예시)"
END_ANCHOR = "[ 7 ] 교복 학교주관구매 기초금액 및 계약방법 결정(예시)"

tree = etree.parse(SECTION_PATH)
root = tree.getroot()

sec_elem = None
for elem in root.iter():
    if etree.QName(elem.tag).localname == "sec":
        sec_elem = elem
        break
assert sec_elem is not None, "sec 요소를 찾지 못함"

sec_children = list(sec_elem)
print(f"sec 직계 자식 수: {len(sec_children)}")

def local(tag):
    return etree.QName(tag).localname

def full_text(elem):
    return "".join(t.text or "" for t in elem.iter() if local(t.tag) == "t")

# 1) 시작/끝 앵커가 들어있는 sec_child 인덱스 찾기
start_idx = None
end_idx = None
for i, child in enumerate(sec_children):
    txt = full_text(child)
    if START_ANCHOR in txt and start_idx is None:
        start_idx = i
    if END_ANCHOR in txt and end_idx is None:
        end_idx = i
        break

print(f"START_ANCHOR found at sec_child[{start_idx}]")
print(f"END_ANCHOR found at sec_child[{end_idx}]")
print()

if start_idx is None or end_idx is None:
    print("앵커를 찾지 못함 - 전체 스캔 결과 일부 출력")
    for i, child in enumerate(sec_children[:20]):
        txt = full_text(child)
        if txt.strip():
            print(f"[{i}] {local(child.tag)}: {txt[:60]!r}")
    sys.exit(1)

print("=== F-007 구간 (start_idx ~ end_idx) 문단/표 목록 ===")
for i in range(start_idx, end_idx + 1):
    child = sec_children[i]
    tag = local(child.tag)
    if tag != "p":
        print(f"[{i}] {tag}")
        continue
    # 이 문단 안에 표(tbl)가 있는지 확인
    tbls = [e for e in child.iter() if local(e.tag) == "tbl"]
    txt = full_text(child)
    marker = f" [표 {len(tbls)}개 포함]" if tbls else ""
    print(f"[{i}] p{marker}: {txt[:80]!r}")
    for ti, tbl in enumerate(tbls):
        rows = tbl.findall(".//hp:tr", namespaces=NS)
        print(f"    tbl[{ti}] rowCnt={tbl.get('rowCnt')} colCnt={tbl.get('colCnt')} 실제 tr수={len(rows)}")
        for ri, tr in enumerate(rows):
            tcs = tr.findall("hp:tc", namespaces=NS)
            for ci, tc in enumerate(tcs):
                cell_txt = full_text(tc)
                if cell_txt.strip():
                    print(f"      tr[{ri}] tc[{ci}]: {cell_txt.strip()[:60]!r}")

print()
print("=== 학교명/서명란/직위 관련 키워드 검색 (표 내부 포함 전체) ===")
KEYWORDS = ["학교", "교장", "담당자", "서명", "직인", "결재", "기안", "검토", "확인", "위원장", "행정실장"]
region_children = sec_children[start_idx:end_idx + 1]
seen = set()
for child in region_children:
    for t in child.iter():
        if local(t.tag) != "t" or not t.text:
            continue
        for kw in KEYWORDS:
            if kw in t.text:
                key = (kw, t.text.strip()[:40])
                if key not in seen:
                    seen.add(key)
                    # 이 t가 표 내부인지 판별
                    anc_tbl = None
                    cur = t.getparent()
                    while cur is not None:
                        if local(cur.tag) == "tbl":
                            anc_tbl = cur
                            break
                        cur = cur.getparent()
                    loc = "표 내부(nested_tbl)" if anc_tbl is not None else "직계"
                    print(f"  [{kw}] {loc}: {t.text.strip()[:50]!r}")
