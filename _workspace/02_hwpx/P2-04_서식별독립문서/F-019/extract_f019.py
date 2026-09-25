"""F-019(입찰참가신청서, [서식 2]) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 606(F-019 본문 1개 문단, idx605 "[9-7]" 캡션
제외, idx607 "[9-8]"는 다음 서식 F-020 캡션)만 남기고 독립 문서로
분리한다. 13x9 NONE 표(빈 신청서 양식) 단일 구조. R-03(업체 반복행)
관련 서식군의 시작 — Excel 트랙과 동일하게 "업체명만 반영, 나머지
개인정보는 빈칸" 원칙을 적용한다(서식_매핑표.md 참고).

허용 반영 범위: 학교명(C-01, "◯◯◯◯학교", 원문자 U+25EF, 4회: 공고
지명번호·건명·본문 참가문구·말미 서명)·학년도(C-02, "20○○학년도", 건명
안 1회)만 토큰화한다. "공고 제20○○–00호"(공고번호 연도)·"20○○.  .  ."
(제출일)은 원문 유지. 신청인 상호·법인등록번호·주소·전화번호·대표자·
주민등록번호(전부 원문에서도 빈칸인 업체/개인정보)는 손대지 않는다.

사용법: python extract_f019.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f019")
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

    if len(sec_children) < 608:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[605].iter() if etree.QName(t.tag).localname == "t"
    )
    if "입찰참가신청서" not in caption_text.replace(" ", ""):
        raise AssertionError(f"idx605가 F-019 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[607].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-8 ]" not in next_caption_text:
        raise AssertionError(f"idx607이 다음(F-020) 캡션이 아님: {next_caption_text!r}")

    # idx606은 13x9 표가 거의 전면을 차지하는 단일 문단 - 캐리어 병합 기법 적용.
    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body = copy.deepcopy(sec_children[606])
    body_run = next(c for c in body if etree.QName(c.tag).localname == "run")
    body_run.insert(0, ctrl)
    body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(body)

    replace_text_anywhere(body, "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(body, "◯◯◯◯학교", "{{학교명}}", expected_count=4)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 4), ("{{학년도}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["20○○학년도", "◯◯◯◯학교"]:
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
    print(f"PASS: F-019 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
