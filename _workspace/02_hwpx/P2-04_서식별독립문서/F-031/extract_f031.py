"""F-031(교복 품목별 금액표, 최종 낙찰자만 제출) 원본 보존형 독립 HWPX
추출.

원본 sec 직계 자식 인덱스 636~637(F-031 본문, idx635 "[9-20]" 캡션 제외,
idx638 "[10]"는 다음 서식(F-032, 이미 완료) 캡션)만 남기고 독립 문서로
분리한다. 1x1 CELL 표(안내문+건명)+8x4 CELL 표(품목별 금액 산출 내역,
빈 계산표) 구조.

허용 반영 범위: 학교명(C-01, "◯◯◯◯학교", 원문자 U+25EF, F-011/012와
동일 표기 — 2회: 건명 1회+말미 "◯◯◯◯학교장 귀하" 1회)·학년도(C-02,
"20○○학년도", 건명 안 1회)만 토큰화한다. "3. 낙찰금액 : 000,000원"(K-01
계산값)·8x4 표의 품목별 수량·금액·비율(전부 빈 계산표, K-02 매핑)·
"20○○년 월 일"(제출일)은 원문 유지한다(F-005/009와 동일 원칙 — 최종
낙찰자만 제출하는 특정 계약 건의 실제 계산값이라 정적 단계에서 미반영).

사용법: python extract_f031.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f031")
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

    if len(sec_children) < 639:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[635].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-20 ]" not in caption_text:
        raise AssertionError(f"idx635가 F-031 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[638].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 10 ]" not in next_caption_text:
        raise AssertionError(f"idx638이 다음(F-032, 범위 밖) 캡션이 아님: {next_caption_text!r}")

    # idx636 표는 pageBreak='CELL' 절대 고정 형식이라 F-045/047/050류와
    # 동일하게 캐리어 병합 기법을 적용한다.
    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(636, 638)]
    first_body = body_paragraphs[0]
    first_body_run = next(c for c in first_body if etree.QName(c.tag).localname == "run")
    first_body_run.insert(0, ctrl)
    first_body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[orig_idx - 636]

    replace_text_anywhere(np(636), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(636), "◯◯◯◯학교", "{{학교명}}", expected_count=2)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 2), ("{{학년도}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["20○○학년도", "◯◯◯◯학교"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"
    assert "000,000원" in all_text, "낙찰금액 원문이 예상과 달라짐(직접 재확인 필요)"

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
    print(f"PASS: F-031 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
