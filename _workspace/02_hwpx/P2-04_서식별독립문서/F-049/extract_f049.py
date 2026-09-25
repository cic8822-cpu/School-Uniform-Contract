"""F-049(교복 학교주관구매 만족도 설문조사 실시) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 886~909(F-049 본문, idx885 "[19]" 캡션 제외,
idx910 "[19-1]"는 다음 서식 F-050 캡션)만 남기고 독립 문서로 분리한다.
F-044/F-046과 완전히 동일한 "표지 스탬프(1x3표)+수신/제목 표(3x2표)+본문
문단+결재란(8x41표)" 구조임을 원본 XML로 직접 확인함(idx887 tbl 1x3, idx888
tbl 3x2, idx909 tbl 8x41). idx895~908(빈 문단 14개, paraPrIDRef=44)는
F-044/046과 동일하게 "붙임" 문단과 결재란(paraPrIDRef=285) 사이에 있어
"빈 2쪽" 위험 패턴이 아니다.

허용 반영 범위: 학교명(C-01, idx887 "○ ○ 학 교")·문서번호(C-06, idx909
결재란 "○○○○학교-")만 토큰화한다. idx889의 "전북특별자치도교육청
학교안전과-0000(20○○.)"는 F-044~048과 동일하게 도교육청 고정 참조 문서라
원문 유지. idx891 "2027학년도"도 F-044~048과 동일 판단(20○○/○○ 같은
placeholder 문자가 없는 고정 예시 연도라 C-02와 동일시하지 않음, 일관성
유지)으로 원문 유지한다.

사용법: python extract_f049.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f049")
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

    if len(sec_children) < 911:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[885].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 19 ]" not in caption_text:
        raise AssertionError(f"idx885가 F-049 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[910].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 19-1 ]" not in next_caption_text:
        raise AssertionError(f"idx910이 다음(F-050) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(886, 910)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 886)]

    replace_text_anywhere(np(887), "○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(np(909), "○○○○학교-", "{{문서번호}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["○ ○ 학 교", "○○○○학교-"]:
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
    print(f"PASS: F-049 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
