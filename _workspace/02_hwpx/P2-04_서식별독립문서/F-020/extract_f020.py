"""F-020(교복 학교주관구매 입찰 참가 신고서, [서식 3]) 원본 보존형 독립
HWPX 추출.

원본 sec 직계 자식 인덱스 608(F-020 본문 1개 문단, idx607 "[9-8]" 캡션
제외, idx609 "[9-9]"는 다음 서식 F-021 캡션)만 남기고 독립 문서로
분리한다. 1x1 CELL 표 2개(안내문+신고 내용) 구조.

허용 반영 범위: 학교명(C-01, "◯◯◯◯학교", 원문자 U+25EF, 2회: 본문+말미)만
토큰화한다. 사업자 상호·번호·대표는 원문에서도 빈칸(업체 정보). 공고일자·
제출일·사업자 대표 성명("○○○", 개인 성명 마스킹)은 특정 회차 데이터·
개인정보라 원문 유지.

사용법: python extract_f020.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree


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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f020")
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

    if len(sec_children) < 610:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[607].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-8 ]" not in caption_text:
        raise AssertionError(f"idx607이 F-020 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[609].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-9 ]" not in next_caption_text:
        raise AssertionError(f"idx609가 다음(F-021) 캡션이 아님: {next_caption_text!r}")

    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body = copy.deepcopy(sec_children[608])
    body_run = next(c for c in body if etree.QName(c.tag).localname == "run")
    body_run.insert(0, ctrl)
    body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(body)

    replace_text_anywhere(body, "◯◯◯◯학교", "{{학교명}}", expected_count=2)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 2, f"{{학교명}} 개수 불일치: 기대 2, 실제 {actual}"
    assert "◯◯◯◯학교" not in all_text, "치환되지 않은 '◯◯◯◯학교' 잔존 문자열 발견"
    assert "사업자 상호 :" in all_text, "사업자 정보 원문이 예상과 달라짐(직접 재확인 필요)"

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
    print(f"PASS: F-020 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
