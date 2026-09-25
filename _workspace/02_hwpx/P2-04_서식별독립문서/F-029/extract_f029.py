"""F-029(개인정보제공 동의서) 원본 보존형 독립 HWPX 추출.
**개인정보 가능 문서(서식_매핑표.md H등급)** — 동의자 성명·주소·연락처·
서명·식별번호는 원문에서도 이미 빈칸이며 이 파이프라인이 전혀 손대지
않는다.

원본 sec 직계 자식 인덱스 631(F-029 본문 1개 문단, idx630 "[9-18]" 캡션
제외, idx632 "[9-19]"는 다음 서식 F-030 캡션)만 남기고 독립 문서로
분리한다. 1x1 CELL 표(안내문+동의 항목+서명란) 단일 구조.

허용 반영 범위: 학교명(C-01, "○○학교", 원문자 U+25CB(F-027/031의 U+25EF와
다름, 직접 확인) — 2회: 본문 1회+말미 "○○학교장 귀하" 1회)·학년도(C-02,
"20○○학년도", 1회)만 토큰화한다. "20○○. . ."(동의 서명 날짜)·업체명·
사업자등록번호·대표자(인)·동의 여부 체크박스는 원문에서도 이미 빈칸인
업체/개인 정보라 절대 채우지 않는다.

사용법: python extract_f029.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f029")
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

    if len(sec_children) < 633:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[630].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-18 ]" not in caption_text:
        raise AssertionError(f"idx630이 F-029 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[632].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-19 ]" not in next_caption_text:
        raise AssertionError(f"idx632가 다음(F-030) 캡션이 아님: {next_caption_text!r}")

    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body = copy.deepcopy(sec_children[631])
    body_run = next(c for c in body if etree.QName(c.tag).localname == "run")
    body_run.insert(0, ctrl)
    body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(body)

    replace_text_anywhere(body, "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(body, "○○학교", "{{학교명}}", expected_count=2)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 2), ("{{학년도}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["20○○학년도", "○○학교"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"
    for blank in ["업 체 명：", "사업자등록번호：", "대 표 자：", "동의 □", "동의하지 않음 □"]:
        assert blank in all_text, f"개인정보 응답란 원문이 예상과 달라짐(직접 재확인 필요): {blank!r}"

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
    print(f"PASS: F-029 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
