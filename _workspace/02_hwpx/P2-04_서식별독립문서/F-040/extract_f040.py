"""F-040(교복선정위원회 위원 청렴 및 보안 서약서, [13-4]) 원본 보존형 독립
HWPX 추출.

원본 sec 직계 자식 인덱스 764~767(F-040 본문, idx763 "[13-4]" 캡션 제외)만
남기고 독립 문서로 분리한다. 본문 텍스트(idx765)가 F-003 본문(idx37)과
바이트 단위로 완전히 동일함을 직접 대조로 확인했으므로(길이 522자, 문자열
비교 True) F-003 스크립트를 인덱스만 바꿔 그대로 재사용한다: 고정 절대폭
단일 셀 표(rowCnt=1 colCnt=1) 안에 본문이 들어 있어 표 폭을 페이지 여백
한계까지 확장해야 하고, "○○○학교장 귀하"(3개 원)가 "○○학교"(2개 원)의
상위 문자열이므로 3개 원 패턴을 먼저 치환해야 한다(F-002/F-003에서 확정된
함정, `_workspace/02_hwpx/P2-04_서식별독립문서/F-003/extract_f003.py` 참고).

허용 반영 범위: 학교명(C-01)만 자동 반영한다. 학년도는 F-003과 동일하게
이 서식에서 "20○○학년도"로 하드코딩돼 있어(기초자료입력 미연동) 토큰화하지
않는다. 서약자 성명·서명·날짜는 위원 개인 자필 공란으로 유지한다.

사용법: python extract_f040.py <원본.hwpx> <출력_템플릿.hwpx>
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

    if len(sec_children) < 768:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[763].iter() if etree.QName(t.tag).localname == "t"
    )
    if "13-4" not in caption_text:
        raise AssertionError(f"idx763이 F-040 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[768].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 14 ]" not in next_caption_text:
        raise AssertionError(f"idx768이 다음(F-041) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(764, 768)]
    body_with_table = body_paragraphs[1]

    secpr_for_margin = carrier_run.find(qn("secPr"))
    page_pr = secpr_for_margin.find(qn("pagePr"))
    margin = page_pr.find(qn("margin"))
    usable_width = int(page_pr.get("width")) - int(margin.get("left")) - int(margin.get("right"))
    tbl = next(e for e in body_with_table.iter() if etree.QName(e.tag).localname == "tbl")
    tbl_sz = tbl.find(qn("sz"))
    original_width = int(tbl_sz.get("width"))
    if usable_width > original_width:
        tbl_sz.set("width", str(usable_width))
        for tc in tbl.iter(qn("tc")):
            cell_sz = tc.find(qn("cellSz"))
            if cell_sz is not None:
                cell_sz.set("width", str(usable_width))
        print(f"INFO: F-040 표 폭 확장 {original_width} -> {usable_width} (줄바꿈 회피 여유 확보)")

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    replace_text_anywhere(body_with_table, "○○○학교", "{{학교명}}")
    replace_text_anywhere(body_with_table, "○○학교", "{{학교명}}")

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
    print(f"PASS: F-040 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
