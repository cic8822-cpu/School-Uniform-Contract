"""F-007(교복 학교주관구매 구매 요청) 원본 보존형 독립 HWPX 추출.
**P2-04 표준 파이프라인 신규 적용** — P2-01에 기존 산출물 3종
(`F007_구매요청_템플릿.hwpx`(9엔트리, placeholder 미토큰화 원문 그대로)·
`F007_구매요청_누름틀_템플릿.hwpx`(필드 컨트롤 방식)·
`교복매뉴얼_F007_원본보존형_템플릿.hwpx`(87쪽 전체 매뉴얼이 통째로 들어있어
F-007 단독 문서가 아님))이 있었으나 전부 P2-04 표준(58엔트리, `{{토큰}}`
텍스트 치환) 구조와 호환되지 않아 재사용하지 않고 원본에서 새로 추출함.

원본 sec 직계 자식 인덱스 165~185(F-007 본문, idx164 "[6]" 캡션 제외,
idx186은 결재란 뒤 트레일링 빈 문단 1개라 범위에서 제외, idx187 "[6-1]"는
다음 서식 F-008 캡션)만 남기고 독립 문서로 분리한다. F-001/F-004/F-032
계열과 동일한 "표지(1x3)+수신제목(3x2)+본문+결재란(8x41)" 구조(idx165 tbl
1x3, idx166 tbl 3x2, idx185 tbl 8x41).

허용 반영 범위: 학교명(C-01, idx165)·학년도(C-02, idx166 제목·idx168
본문·idx169 대상및수량, 3회)·관련문서(D-03, idx167 "○○학교-○○(20○○.○.○.)")·
문서번호(C-06, idx185)만 토큰화한다. idx169의 "○○명"(참여자 수)·idx170
"동복 000,000원, 하복 00,000원"(예상단가)·idx172 납품기한 날짜는 특정
회차의 실제 데이터라 원문 유지한다(F-005/009와 동일 원칙).

사용법: python extract_f007.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree

CARRIER_CHAR_PR_ID = "52"
CARRIER_PARA_PR_ID = "41"


def remove_own_lineseg(p):
    for c in list(p):
        if etree.QName(c.tag).localname == "linesegarray":
            p.remove(c)


def nearest_p(elem):
    cur = elem.getparent()
    while cur is not None and etree.QName(cur.tag).localname != "p":
        cur = cur.getparent()
    return cur


def replace_text_anywhere(root_elem, old_text, new_text, expected_count=None):
    targets = [
        t for t in root_elem.iter()
        if etree.QName(t.tag).localname == "t" and t.text and old_text in t.text
    ]
    total = 0
    for t in targets:
        total += t.text.count(old_text)
        t.text = t.text.replace(old_text, new_text)
        host = nearest_p(t)
        if host is not None:
            remove_own_lineseg(host)
    if not targets:
        raise AssertionError(f"'{old_text}' 를 찾지 못함")
    if expected_count is not None and total != expected_count:
        raise AssertionError(f"'{old_text}' 매치 수 불일치: 기대 {expected_count}, 실제 {total}")
    return total


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f007")
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

    if len(sec_children) < 188:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[164].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 6 ]" not in caption_text:
        raise AssertionError(f"idx164가 F-007 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[187].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 6-1 ]" not in next_caption_text:
        raise AssertionError(f"idx187이 다음(F-008) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(165, 186)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 165)]

    replace_text_anywhere(np(165), "○ ○ 학 교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(166), "○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(167), "○○학교-○○(20○○.○.○.)", "{{관련문서}}", expected_count=1)
    replace_text_anywhere(np(168), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(169), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(185), "○○○○학교-", "{{문서번호}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{학년도}}", 3), ("{{관련문서}}", 1), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["○ ○ 학 교", "○○학교-○○(20○○.○.○.)", "20○○학년도", "○○○○학교-"]:
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
    print(f"PASS: F-007 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
