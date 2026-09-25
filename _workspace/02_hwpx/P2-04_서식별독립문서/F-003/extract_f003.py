"""F-003(교복선정위원회 위원 청렴 및 보안 서약서) 원본 보존형 독립 HWPX 추출.

원본 87쪽 매뉴얼 사본의 section0.xml에서 sec 직계 자식 인덱스 37(F-003 본문,
idx36 "[1-2]" 캡션 제외)만 남기고 독립 문서로 분리한 뒤 허용 공통값 자리
(학교명)에만 토큰을 삽입한다. F-002와 완전히 동일한 구조(단일 문단 안에
고정폭 표 1개가 anchor됨, rowCnt=1 colCnt=1 width=46528 HWPUNIT)임을 원본
XML 직접 확인으로 검증함.

허용 반영 범위(기초자료입력 VBA 수식 대조로 확정, build_excel_v1_structure.ps1
7o절): 학교명(C-01)만 자동 반영하고, 학년도는 이 서식에서 "20○○학년도"로
하드코딩돼 있어(기초자료입력!C5와 연동하지 않음) 토큰화하지 않는다. 성명·
서명·서명일자는 위원 개인 자필 공란으로 유지한다(F-002와 동일 원칙).

F-002에서 실측 확인된 두 가지 함정을 처음부터 반영한다:
1) 매치를 3회 이상 처리할 때 `root_elem.iter()` 순회 중 `remove_own_lineseg`로
   트리를 변형하면 lxml 이터레이터가 이후 매치를 건너뜀 → 매치 리스트를 먼저
   확정한 뒤에만 변형함.
2) 본문이 고정 절대폭 단일 셀 표 안에 있어, 짧은 placeholder("○○학교")보다
   긴 실제 학교명으로 치환하면 해당 줄이 2줄로 줄바꿈되면서 한컴이 줄바꿈된
   두 번째 줄의 가로 위치를 잘못 계산해 표 밖으로 삐져나오는 렌더링 결함이
   있음(2026-09-22 F-002에서 확정, 사용자 결정으로 표 폭을 페이지 여백 한계
   까지 확장하는 완화책 채택). 학교명 8자까지 안전, 10자에서 재현(잔여 한계).

또한 "○○○학교장 귀하"(3개 원)가 "○○학교"(2개 원)의 상위 문자열이므로,
치환 시 3개 원 패턴을 먼저 처리해야 한 개의 원이 남는 사고를 피할 수 있다
(F-003 원문 대조로 확정, F-002와 다른 부분).

사용법: python extract_f003.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree

CARRIER_CHAR_PR_ID = "52"
CARRIER_PARA_PR_ID = "129"


def qn(local):
    return f"{{http://www.hancom.co.kr/hwpml/2011/paragraph}}{local}"


def remove_own_lineseg(p):
    for c in list(p):
        if etree.QName(c.tag).localname == "linesegarray":
            p.remove(c)


def nearest_p(elem):
    cur = elem.getparent()
    while cur is not None and etree.QName(cur.tag).localname != "p":
        cur = cur.getparent()
    return cur


def replace_text_anywhere(root_elem, old_text, new_text):
    targets = [
        t for t in root_elem.iter()
        if etree.QName(t.tag).localname == "t" and t.text and old_text in t.text
    ]
    for t in targets:
        t.text = t.text.replace(old_text, new_text)
        host = nearest_p(t)
        if host is not None:
            remove_own_lineseg(host)
    if not targets:
        raise AssertionError(f"'{old_text}' 를 찾지 못함")
    return len(targets)


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp")
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)

    with zipfile.ZipFile(src_path) as zf:
        zf.extractall(work_dir)

    section_path = os.path.join(work_dir, "Contents", "section0.xml")
    tree = etree.parse(section_path)
    root = tree.getroot()
    sec_elem = next(e for e in root.iter() if etree.QName(e.tag).localname == "sec")
    sec_children = list(sec_elem)

    if len(sec_children) < 38:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[36].iter() if etree.QName(t.tag).localname == "t"
    )
    if "1-2" not in caption_text:
        raise AssertionError(f"idx36이 F-003 캡션이 아님: {caption_text!r}")

    # --- 페이지 설정 문단(idx0)에서 secPr/ctrl만 남기고 로고(pic) 제거 ---
    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    # --- F-003 본문 범위(idx 37, 단일 문단) 깊은 복사 ---
    body_paragraph = copy.deepcopy(sec_children[37])

    # --- 표 폭을 페이지 여백 한계까지 확장(F-002와 동일 결정 재사용) ---
    secpr_for_margin = carrier_run.find(qn("secPr"))
    page_pr = secpr_for_margin.find(qn("pagePr"))
    margin = page_pr.find(qn("margin"))
    usable_width = int(page_pr.get("width")) - int(margin.get("left")) - int(margin.get("right"))
    tbl = next(e for e in body_paragraph.iter() if etree.QName(e.tag).localname == "tbl")
    tbl_sz = tbl.find(qn("sz"))
    original_width = int(tbl_sz.get("width"))
    if usable_width > original_width:
        tbl_sz.set("width", str(usable_width))
        for tc in tbl.iter(qn("tc")):
            cell_sz = tc.find(qn("cellSz"))
            if cell_sz is not None:
                cell_sz.set("width", str(usable_width))
        print(f"INFO: F-003 표 폭 확장 {original_width} -> {usable_width} (줄바꿈 회피 여유 확보)")

    # --- 새 sec 자식 목록으로 교체(맨 앞에 페이지설정 캐리어 삽입) ---
    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    sec_elem.append(body_paragraph)

    # --- 토큰 삽입: 학교명(C-01)만 자동 반영. 3개 원 패턴을 먼저 처리한다 ---
    replace_text_anywhere(body_paragraph, "○○○학교", "{{학교명}}")
    replace_text_anywhere(body_paragraph, "○○학교", "{{학교명}}")

    # --- 잔존 검증: 토큰이 정확히 4회 존재해야 함(본문 3회 + "장 귀하" 1회) ---
    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    if actual != 4:
        raise AssertionError(f"{{{{학교명}}}} 개수 불일치: 기대 4, 실제 {actual}")
    if "○○학교" in all_text or "○○○학교" in all_text:
        raise AssertionError("치환되지 않은 '○○학교' 계열 잔존 문자열 발견")

    tree.write(section_path, xml_declaration=True, encoding=tree.docinfo.encoding or "UTF-8", standalone=tree.docinfo.standalone)

    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w") as zf:
        mimetype_path = os.path.join(work_dir, "mimetype")
        zf.write(mimetype_path, "mimetype", compress_type=zipfile.ZIP_STORED)
        for dirpath, _, filenames in os.walk(work_dir):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                arcname = os.path.relpath(filepath, work_dir).replace(os.sep, "/")
                if arcname == "mimetype":
                    continue
                zf.write(filepath, arcname, compress_type=zipfile.ZIP_DEFLATED)

    shutil.rmtree(work_dir)
    print(f"PASS: F-003 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
