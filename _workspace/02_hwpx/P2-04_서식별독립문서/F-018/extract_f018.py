"""F-018(정량적 평가 자기 평점표, [서식 1_1]) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 602~604(F-018 본문, idx601 "[9-7]" 캡션 제외,
idx605 "[9-7]"는 다음 서식(F-019, 같은 라벨 재사용) 캡션)만 남기고 독립
문서로 분리한다. 16x5 CELL 표(정량평가 자기 평점) 단일 구조.

F-015/016과 동일 패턴: "2026학년도"는 명시적 placeholder 문자가 없는
고정 연도라 원문 유지, 학교명("○○학교", 1회)만 토큰화한다.

사용법: python extract_f018.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f018")
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

    if len(sec_children) < 606:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[601].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-7" not in caption_text:
        raise AssertionError(f"idx601이 F-018 캡션이 아님: {caption_text!r}")
    caption_full = caption_text
    if "정량적 평가 자기 평점표" not in caption_full:
        raise AssertionError(f"idx601 캡션 본문이 예상과 다름: {caption_full!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[605].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-7 ]" not in next_caption_text or "입찰참가신청서" not in next_caption_text.replace(" ", ""):
        raise AssertionError(f"idx605가 다음(F-019) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(602, 605)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 602)]

    replace_text_anywhere(np(604), "○○학교", "{{학교명}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 1, f"{{학교명}} 개수 불일치: 기대 1, 실제 {actual}"
    assert "○○학교" not in all_text, "치환되지 않은 '○○학교' 잔존 문자열 발견"
    assert "2026학년도" in all_text, "고정 연도 '2026학년도'가 예상과 달라짐(직접 재확인 필요)"

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
    print(f"PASS: F-018 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
