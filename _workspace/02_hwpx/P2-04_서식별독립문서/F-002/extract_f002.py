"""F-002(교복선정위원회 위원 수락 및 확인서) 원본 보존형 독립 HWPX 추출 스크립트.

원본 87쪽 매뉴얼 사본(artifacts/hwpx/P0-02_원본사본.hwpx)의 section0.xml에서
sec 직계 자식 인덱스 35(F-002 본문, idx34 "[1-1]" 캡션은 제외)만 남기고 독립
문서로 분리한 뒤 허용 공통값 자리(학교명)에만 토큰을 삽입한다. F-002 본문은
단일 문단(idx35) 안에 표가 통째로 anchor돼 있는 구조라 F-001과 달리 본문
범위가 문단 1개뿐이다.

주소·성명·서명·서명일자는 위원 개인이 자필로 작성하는 영역이라 토큰화하지
않는다(기초자료입력!C5 학년도의 뒤 2자리로 서명일자 연도를 추정 반영하는
Excel 수식이 있으나, 이는 실제 서명일과 무관할 수 있는 편의 기능이라 원본
보존형 HWPX에서는 재현하지 않고 원문 그대로 공란 placeholder를 유지함 —
2026-09-22 8.2절 F-002 작업 결정).

원본 인덱스 0(secPr 페이지 설정 문단)은 F-001 파일럿에서 확정한 방식대로
표지 로고(pic)만 제거하고 charPr/paraPr를 본문 크기로 교체해 secPr/ctrl은
그대로 유지한다. header.xml/mimetype/META-INF/BinData/Preview는 손대지 않는다.

사용법: python extract_f002.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree

NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"

# F-001 파일럿에서 실측 확정한 "본문 크기, 테두리 없음" 전역 스타일 ID
# (idx17 run의 charPrIDRef=52, idx18 문단의 paraPrIDRef=129). header.xml
# 스타일 레코드는 문서 전체가 공유하므로 모든 서식에서 동일하게 재사용한다.
CARRIER_CHAR_PR_ID = "52"
CARRIER_PARA_PR_ID = "129"


def qn(local):
    return f"{{{NS_HP}}}{local}"


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
    # 매치 대상을 먼저 리스트로 확정한 뒤에만 변형한다. root_elem.iter()로 순회하며 동시에
    # remove_own_lineseg로 트리 구조를 변형하면(예: 3회 이상 매치) lxml 이터레이터가 다음
    # 형제 노드를 건너뛰어 실제로는 1건만 처리되는 결함이 F-002에서 실측 재현됨(기대 3, 실제 1).
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

    if len(sec_children) < 36:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[34].iter() if etree.QName(t.tag).localname == "t"
    )
    if "1-1" not in caption_text:
        raise AssertionError(f"idx34가 F-002 캡션이 아님: {caption_text!r}")

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

    # --- F-002 본문 범위(idx 35, 단일 문단) 깊은 복사 ---
    body_paragraph = copy.deepcopy(sec_children[35])

    # --- 표 폭을 페이지 여백 한계까지 확장(2026-09-22 사용자 결정) ---
    # 원본 표 폭(46528)은 "○○학교"(4자) placeholder에 딱 맞게 짜여 있어, 실제 학교명
    # (예: "검증초등학교" 6자)으로 치환하면 "아울러 본인은 ... 업체와 전혀" 문단이 2줄로
    # 줄바꿈되는데, 이때 한컴이 줄바꿈된 두 번째 줄의 가로 위치를 잘못 계산해 표 테두리
    # 밖으로 텍스트가 삐져나오는 렌더링 결함이 실측 확인됨(2026-09-22, F-002 육안 QA).
    # 표 폭을 페이지 여백이 허용하는 최대치까지 넓혀 이 줄바꿈 자체를 회피한다(원본
    # 페이지 여백 안에서만 넓히므로 페이지 레이아웃을 벗어나지 않음).
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
        print(f"INFO: F-002 표 폭 확장 {original_width} -> {usable_width} (줄바꿈 회피 여유 확보)")

    # --- 새 sec 자식 목록으로 교체(맨 앞에 페이지설정 캐리어 삽입) ---
    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    sec_elem.append(body_paragraph)

    # --- 토큰 삽입: 학교명(C-01)만 자동 반영, 나머지(주소/성명/서명/서명일)는 자필 공란 유지 ---
    replace_text_anywhere(body_paragraph, "○○학교", "{{학교명}}")

    # --- 잔존 검증: 토큰이 정확히 3회 존재해야 함(본문 2회 + "○○학교장 귀하" 1회) ---
    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    if actual != 3:
        raise AssertionError(f"{{{{학교명}}}} 개수 불일치: 기대 3, 실제 {actual}")
    if "○○학교" in all_text:
        raise AssertionError("치환되지 않은 '○○학교' 잔존 문자열 발견")

    tree.write(section_path, xml_declaration=True, encoding=tree.docinfo.encoding or "UTF-8", standalone=tree.docinfo.standalone)

    # --- 재패키징: mimetype 비압축 우선 ---
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
    print(f"PASS: F-002 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
