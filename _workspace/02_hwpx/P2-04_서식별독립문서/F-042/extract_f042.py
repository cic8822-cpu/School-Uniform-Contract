"""F-042(낙찰자 결정 통보) 원본 보존형 독립 HWPX 추출.

원본 87쪽 매뉴얼 사본의 section0.xml에서 sec 직계 자식 인덱스 787~804(F-042
본문, idx786 "[15]" 캡션 제외)만 남기고 독립 문서로 분리한다. F-041과 동일한
구조(표지+수신/제목 표+본문+결재란). 업체명 등은 U+3007 "〇"(예: "〇〇교복사")
placeholder라 손대지 않고, 학교명·학년도는 U+25CB "○" placeholder만 치환한다.

허용 반영 범위: 학교명(C-01)·학년도(C-02)·제출기한(B-07, 기초자료입력!C18 —
task.md 기록상 F-042가 이 필드의 첫 소비처)·문서번호(C-06)만 토큰화한다.
업체명·대표자·단가·예정수량·계약체결금액 등은 특정 입찰 회차의 실제 데이터라
정적 HWPX 단계에서는 원문 유지(F-005/F-006/F-041과 동일 스코프 결정).

사용법: python extract_f042.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree

CARRIER_CHAR_PR_ID = "52"
CARRIER_PARA_PR_ID = "129"


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

    if len(sec_children) < 805:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[786].iter() if etree.QName(t.tag).localname == "t"
    )
    if caption_text.strip() != "[ 15 ] 낙찰자 결정 통보(예시)":
        raise AssertionError(f"idx786이 F-042 캡션이 아님: {caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(787, 805)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 787)]

    replace_text_anywhere(np(788), "○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(np(789), "20○○년도", "{{학년도}}")
    replace_text_anywhere(np(791), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(791), "20○○.○○.○○까지", "{{제출기한}}까지")
    replace_text_anywhere(np(802), "○○○○학교-", "{{문서번호}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{학년도}}", 2), ("{{제출기한}}", 1), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["20○○년도", "20○○학년도", "○ ○ 학 교", "20○○.○○.○○"]:
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
    print(f"PASS: F-042 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
