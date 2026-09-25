"""F-010(사전규격공개 기안문 작성) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 229~247(F-010 본문, idx228 "[8]" 캡션 제외,
idx248~249는 결재란 뒤 트레일링 빈 문단 2개라 F-032류 "빈 2쪽" 위험을
피하기 위해 범위에서 제외, idx250 "[9]"는 다음 서식 F-011 캡션)만
남기고 독립 문서로 분리한다. F-009와 동일한 "표지(1x3)+수신제목(3x2)+
본문+결재란(8x41)" 구조(idx230 tbl 1x3, idx231 tbl 3x2, idx247 tbl 8x41).

원본에서 직접 확인한 placeholder:
  - idx230 "○ ○ 학 교"(표지, 학교명)
  - idx231 제목 안의 "○○학년도"
  - idx235 본문 "20○○학년도"
  - idx236 "가. 사 업 명：○○학교 학생 교복 구입"의 "○○학교"(학교명, 2개
    원문자 표준 패턴)
  - idx240 "마. 공개사항：20○○학년도 학생 교복사양서"의 "20○○학년도"
  - idx244 "붙임  20○○학년도 학생 교복사양서 1부.  끝."의 "20○○학년도"
  - idx247 결재란 "○○○○학교-"(문서번호)

허용 반영 범위: 학교명(C-01, 2곳)·학년도(C-02, 4곳)·문서번호(C-06)만
토큰화한다. idx237 "나. 기초금액：단가 ○○원(동복 ○○원, 하복 ○○원)"·
idx238 "다. 배정예산：000,000원×000벌=000,000,000원"(B-04 기초금액·계산값,
K-01 매핑이나 정적 단계에서 미계산, F-005/009와 동일 원칙)·idx241
"바. 사전공개기간：20○○.○.○. ~ 20○○.○.○.(5일간)"·idx242 "사. 의견등록
마감일시：20○○.○.○.(○) ○○:○○"(특정 회차의 실제 이벤트 날짜)는 원문
유지한다. "1. 관련 가./나."(idx233~234)는 법령 조문 인용이라 placeholder
자체가 없음(D-03 관련문서 필드가 이 서식엔 없음, 직접 확인).

사용법: python extract_f010.py <원본.hwpx> <출력_템플릿.hwpx>
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


def replace_split_across_runs(root_elem, part_texts, new_text):
    """part_texts에 나열된 문자열이 순서대로 연속된 별도 <t> 노드에 걸쳐
    분할돼 있는 경우(F-010 idx235: "20" + "○○학년도 ...")를 처리한다.
    첫 노드의 text를 new_text로 바꾸고 나머지 노드는 빈 문자열로 비운다."""
    all_t = [t for t in root_elem.iter() if etree.QName(t.tag).localname == "t"]
    n = len(part_texts)
    for i in range(len(all_t) - n + 1):
        window = all_t[i:i + n]
        if all((w.text or "").startswith(part_texts[j]) if j == n - 1 else (w.text or "") == part_texts[j]
               for j, w in enumerate(window)):
            remainder = (window[-1].text or "")[len(part_texts[-1]):]
            window[0].text = new_text
            for w in window[1:-1]:
                w.text = ""
            window[-1].text = remainder
            host = nearest_p(window[0])
            if host is not None:
                remove_own_lineseg(host)
            return
    raise AssertionError(f"분할 매치 실패: {part_texts!r}")


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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f010")
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

    if len(sec_children) < 251:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[228].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 8 ]" not in caption_text:
        raise AssertionError(f"idx228이 F-010 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[250].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9 ]" not in next_caption_text:
        raise AssertionError(f"idx250이 다음(F-011) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(229, 248)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 229)]

    replace_text_anywhere(np(230), "○ ○ 학 교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(231), "○○학년도", "{{학년도}}", expected_count=1)
    replace_split_across_runs(np(235), ["20", "○○학년도"], "{{학년도}}")
    replace_text_anywhere(np(236), "○○학교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(240), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(244), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(247), "○○○○학교-", "{{문서번호}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 2), ("{{학년도}}", 4), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["○ ○ 학 교", "○○학년도", "20○○학년도", "○○○○학교-"]:
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
    print(f"PASS: F-010 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
