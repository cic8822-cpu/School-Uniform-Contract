"""F-030(청렴계약 이행서약서) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 633(F-030 본문 1개 문단, idx632 "[9-19]" 캡션
제외, idx634 빈 문단도 제외, idx635 "[9-20]"는 다음 서식 F-031 캡션)만
남기고 독립 문서로 분리한다. 1x1 TABLE 표(서약 조항 4개 항목+서명란)
단일 구조.

원본을 직접 확인한 결과 이 서식은 학년도(C-02) placeholder가 전혀 없음
(순수 서약 조항 텍스트, "202  .   .   ." 제출일만 있음 — 4자리 연도 앞
두 자리만 고정 인쇄되고 나머지가 빈칸인 형태로 학년도 자동 반영 대상이
아님). 학교명(C-01) placeholder만 "○○○학교장 귀하"(원문자 3개, F-003/
F-040과 동일 패턴) 1회 존재.

허용 반영 범위: 학교명(1회)만 토큰화한다. "202  .   .   ."(제출일)·
업체명·대표자(인) 서명란은 원문에서도 이미 빈칸인 업체 정보라 원문 유지.

사용법: python extract_f030.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f030")
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

    if len(sec_children) < 636:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[632].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-19 ]" not in caption_text:
        raise AssertionError(f"idx632가 F-030 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[635].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-20 ]" not in next_caption_text:
        raise AssertionError(f"idx635가 다음(F-031) 캡션이 아님: {next_caption_text!r}")

    # idx633 표는 pageBreak='TABLE'(F-045/047/050류의 'CELL'과 다름)이지만
    # 안전을 위해 동일하게 캐리어 병합 기법을 적용한다(별도 carrier 문단을
    # 추가하지 않고 secPr/ctrl을 본문 표 문단 자신의 run에 직접 삽입).
    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body = copy.deepcopy(sec_children[633])
    body_run = next(c for c in body if etree.QName(c.tag).localname == "run")
    body_run.insert(0, ctrl)
    body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(body)

    replace_text_anywhere(body, "○○○학교장", "{{학교명}}장", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 1, f"{{학교명}} 개수 불일치: 기대 1, 실제 {actual}"
    assert "○○○학교장" not in all_text, "치환되지 않은 '○○○학교장' 잔존 문자열 발견"
    assert "202  .     .     ." in all_text, "제출일 원문이 예상과 달라짐(직접 재확인 필요)"
    assert "업체명 :" in all_text, "업체명 서명란 원문이 예상과 달라짐(직접 재확인 필요)"

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
    print(f"PASS: F-030 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
