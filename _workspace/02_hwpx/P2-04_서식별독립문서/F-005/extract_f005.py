"""F-005(교복 학교주관구매 추진 계획안) 원본 보존형 독립 HWPX 추출.

원본 87쪽 매뉴얼 사본의 section0.xml에서 sec 직계 자식 인덱스 62~137(F-005
본문, idx61 "[2-1]" 캡션 제외)만 남기고 독립 문서로 분리한다. Excel 구현
(build_excel_v1_structure.ps1 7q절)과 동일한 반영 범위를 따른다:
"학교명·학년도(C-01·C-02) 공통값만 자동 반영하고, 위원 반복행·상한가격
계산·나머지 정책 설명 본문은 학교마다 달라지지 않는 고정 안내문이므로
원문 그대로 반영한다." 위원 명단(idx89, 8행5열 표)과 교복 상한가격(idx100,
금액·기간)은 Excel의 DB 반복행·계산 셀 연동이 없는 정적 HWPX 단계에서는
아직 채울 데이터 소스가 없어(③ Excel→HWPX 연동은 47종 템플릿이 갖춰진
뒤 진행 예정, 계획.md 8.2절) 이번 범위에서 토큰화하지 않고 원문 placeholder
그대로 유지한다.

원문에서 학교명·학년도 placeholder는 4가지 표기 변형이 섞여 있음을 원본
직접 확인으로 파악함: "○ ○ 학 교"(idx62, 표지 스탬프), "0000학년도"(idx62),
"2OOO학년도"(idx63, 원문자 ○ 대신 라틴 대문자 O 사용), "20OO학년도"(idx67·71,
2회). 이 문서 전체에서 "학교"가 포함된 다른 문장(예: "학교주관구매",
"학교회계", "학교장이")은 특정 학교를 가리키는 placeholder가 아니라 정책
설명에 쓰인 일반 명사이므로 치환 대상에서 제외함(원본 대조로 확인).

idx62의 표지 레이아웃은 고정 절대폭 단일 셀 표(rowCnt=1 colCnt=1
width=50460)에 담겨 있어 F-002/F-003과 동일한 줄바꿈 오버플로 위험이
있으므로 같은 폭 확장(페이지 여백 한계까지)을 적용한다.

사용법: python extract_f005.py <원본.hwpx> <출력_템플릿.hwpx>
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

    if len(sec_children) < 138:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[61].iter() if etree.QName(t.tag).localname == "t"
    )
    if "2-1" not in caption_text:
        raise AssertionError(f"idx61이 F-005 캡션이 아님: {caption_text!r}")

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

    # --- F-005 본문 범위(idx 62~137) 깊은 복사 ---
    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(62, 138)]

    # --- idx62(new idx0)의 표지 고정폭 단일 셀 표 폭 확장(F-002/F-003과 동일 결정) ---
    secpr_for_margin = carrier_run.find(qn("secPr"))
    page_pr = secpr_for_margin.find(qn("pagePr"))
    margin = page_pr.find(qn("margin"))
    usable_width = int(page_pr.get("width")) - int(margin.get("left")) - int(margin.get("right"))
    cover = body_paragraphs[0]
    cover_tbl = next((e for e in cover.iter() if etree.QName(e.tag).localname == "tbl"), None)
    assert cover_tbl is not None, "idx62 표지 표를 찾지 못함"
    tbl_sz = cover_tbl.find(qn("sz"))
    original_width = int(tbl_sz.get("width"))
    if usable_width > original_width:
        tbl_sz.set("width", str(usable_width))
        # 주의: cover_tbl.iter(qn("tc"))는 재귀 탐색이라 표지 표 안에 중첩된
        # 제목 박스(3행2열 표, F-002/F-003에는 없던 F-005 고유 구조)의 셀까지
        # 잡아버려 그 중첩 표의 각 열 폭을 통째로 usable_width로 덮어써
        # 렌더링 폭이 배로 늘어나는 결함이 실측 확인됨(2026-09-22). 바깥 표
        # 자신의 tr 직계 자식 tc만 수정해 중첩 표는 절대 건드리지 않는다.
        outer_tr = cover_tbl.find(qn("tr"))
        for tc in outer_tr.findall(qn("tc")):
            cell_sz = tc.find(qn("cellSz"))
            if cell_sz is not None:
                cell_sz.set("width", str(usable_width))
        print(f"INFO: F-005 표지 표 폭 확장 {original_width} -> {usable_width} (줄바꿈 회피 여유 확보)")

    # --- 새 sec 자식 목록으로 교체(맨 앞에 페이지설정 캐리어 삽입) ---
    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 62)]

    replace_text_anywhere(np(62), "○ ○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(np(62), "0000학년도", "{{학년도}}")
    replace_text_anywhere(np(63), "2OOO학년도", "{{학년도}}")
    replace_text_anywhere(np(67), "20OO학년도", "{{학년도}}")
    replace_text_anywhere(np(71), "20OO학년도", "{{학년도}}")

    # --- 잔존 검증 ---
    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{학년도}}", 4)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["○ ○ ○ 학 교", "0000학년도", "2OOO학년도", "20OO학년도"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"

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
    print(f"PASS: F-005 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
