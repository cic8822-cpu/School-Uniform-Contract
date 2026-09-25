"""F-004(교복 학교주관구매 추진 계획 수립 기안문) 원본 보존형 독립 HWPX 추출.

원본 87쪽 매뉴얼 사본의 section0.xml에서 sec 직계 자식 인덱스 39~60(F-004 본문,
idx38 "[2]" 캡션 제외)만 남기고 독립 문서로 분리한 뒤 허용 공통값 자리(학교명·
학년도·관련문서·문서번호)에만 토큰을 삽입한다. F-001과 동일하게 표지 스탬프
(idx39, 3열 표)·수신/제목 표(idx40, 3행2열 표)·본문 문단(idx41/42/45)·결재란
그리드(idx60, 8행41열 표)로 구성된 "여러 문단 범위" 구조이며, F-002/F-003의
위험한 고정폭 단일 셀(rowCnt=1 colCnt=1) 표가 아님을 원본 XML로 직접 확인함
(idx39 sz={width:50707}, idx40 sz={width:50145}, idx60 sz={width:45978} —
모두 다중 행/열 그리드형 표라 F-002/F-003과 같은 줄바꿈 오버플로 위험이 낮음,
다만 골든결과 PDF 육안 대조로 반드시 재확인함).

허용 반영 범위(기초자료입력 VBA 수식 대조, build_excel_v1_structure.ps1 7p절):
C-01(학교명)·C-02(학년도)·D-03(관련문서, C23 자유 텍스트)·C-06(문서번호, C9).
결재란의 "시행" 값(문서번호)은 F-001 idx33과 동일한 placeholder("○○○○학교-")
이므로 같은 토큰({{문서번호}})을 재사용함. 발행일(C-05)은 이 서식에서 결재란에
별도 표시 자리가 없어(F-001과 달리 idx60에 "접수()"만 있고 시행 옆에 날짜
칸이 없음) 반영하지 않음.

F-002/F-003에서 확립한 안전 패턴을 적용함: 매치 리스트를 먼저 확정한 뒤에만
`linesegarray`를 제거한다(lxml 이터레이터 변형 중 순회 결함 회피).

사용법: python extract_f004.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree

CARRIER_CHAR_PR_ID = "52"
CARRIER_PARA_PR_ID = "129"


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

    if len(sec_children) < 61:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[38].iter() if etree.QName(t.tag).localname == "t"
    )
    if caption_text.strip() != "[ 2 ] 교복 학교주관구매 추진 계획 수립(예시)":
        raise AssertionError(f"idx38이 F-004 캡션이 아님: {caption_text!r}")

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

    # --- F-004 본문 범위(idx 39~60) 깊은 복사 ---
    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(39, 61)]

    # --- 새 sec 자식 목록으로 교체(맨 앞에 페이지설정 캐리어 삽입) ---
    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)
    # body_paragraphs는 new_children[1:]에 대응 (idx39~60 -> new idx1~22)
    def np(orig_idx):
        return new_children[1 + (orig_idx - 39)]

    idx39, idx40, idx41, idx42, idx45, idx60 = (
        np(39), np(40), np(41), np(42), np(45), np(60)
    )

    replace_text_anywhere(idx39, "○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(idx40, "○○학년도 교복 학교주관구매 추진 계획(안)", "{{학년도}} 교복 학교주관구매 추진 계획(안)")
    replace_text_anywhere(idx41, "○○학교-0000(20○○.○.○)", "{{관련문서}}")
    replace_text_anywhere(idx42, "○○학교", "{{학교명}}")
    replace_text_anywhere(idx42, "○○학년도", "{{학년도}}")
    replace_text_anywhere(idx45, "20○○학년도", "{{학년도}}")
    replace_text_anywhere(idx60, "○○○○학교-", "{{문서번호}}")

    # --- 잔존 검증 ---
    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 2), ("{{학년도}}", 3), ("{{관련문서}}", 1), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["○○학교", "○○학년도", "20○○학년도"]:
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
    print(f"PASS: F-004 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
