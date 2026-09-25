"""F-009(교복 학교주관구매 기초금액 및 계약방법 결정) 원본 보존형 독립
HWPX 추출.

원본 sec 직계 자식 인덱스 209~226(F-009 본문, idx208 "[7]" 캡션 제외,
idx227은 결재란 뒤 트레일링 빈 문단이라 F-032류 "빈 2쪽" 위험을 피하기
위해 범위에서 제외, idx228 "[8]"는 다음 서식 F-010 캡션)만 남기고 독립
문서로 분리한다. F-001/F-004/F-032 계열과 동일한 "표지(1x3)+수신제목(3x2)+
본문+결재란(8x41)" 구조(idx210 tbl 1x3, idx211 tbl 3x2, idx226 tbl 8x41).

원본에서 직접 확인한 placeholder(추측 금지):
  - idx210 "○ ○ 학 교"(표지, 학교명)
  - idx211 제목 안의 "○○학년도"(원문자 U+25CB 2개, 다른 서식과 달리 제목
    자체에 학년도가 있음)
  - idx212 "〇〇학교-000(20○○.○. ○.)"(관련문서 — **원문자 U+3007**(한자권
    0) 사용, F-041 등의 "○○학교-000(...)"(U+25CB)와 다른 문자 종류이므로
    직접 확인 후 그대로 매치)
  - idx213/idx214의 "20○○학년도"(본문·건명, 학년도)
  - idx226 결재란 "○○○○학교-"(문서번호)

허용 반영 범위: 학교명(C-01)·학년도(C-02, 3곳: 제목·본문·건명)·관련문서
(D-03)·문서번호(C-06)만 토큰화한다. idx216 "다. 기초금액：금000,000원"·
idx217 "동복 000,000원, 하복 000,000"(B-04 기초금액, 매핑표상 K-01 계산
매핑이나 이 정적 HWPX 단계에서는 계산하지 않음, F-005/041과 동일 원칙)과
idx218 "라. 납품기한：20○○. ○. ○.(○)"(특정 계약 건의 실제 이벤트 날짜,
매핑표 F-009 허용목록에도 없음)는 원문 유지한다. idx214 "가. 건명：
20○○학년도 교복 학교주관 구매"의 "교복 학교주관 구매"는 매핑표가 B-01을
허용 목록에 포함하지만 원문 자체에 별도 placeholder가 없는 고정 문구라(
F-004/F-023과 같은 유형의 매핑표-원문 불일치) 학년도만 토큰화하고 나머지는
그대로 둔다.

사용법: python extract_f009.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f009")
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

    if len(sec_children) < 229:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[208].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 7 ]" not in caption_text:
        raise AssertionError(f"idx208이 F-009 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[228].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 8 ]" not in next_caption_text:
        raise AssertionError(f"idx228이 다음(F-010) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(209, 227)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 209)]

    replace_text_anywhere(np(210), "○ ○ 학 교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(211), "○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(212), "〇〇학교-000(20○○.○. ○.)", "{{관련문서}}", expected_count=1)
    replace_text_anywhere(np(213), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(214), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(226), "○○○○학교-", "{{문서번호}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{학년도}}", 3), ("{{관련문서}}", 1), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["○ ○ 학 교", "○○학년도", "〇〇학교-000", "20○○학년도", "○○○○학교-"]:
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
    print(f"PASS: F-009 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
