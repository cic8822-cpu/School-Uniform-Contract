"""F-043(계약체결) 원본 보존형 독립 HWPX 추출.

원본 87쪽 매뉴얼 사본의 section0.xml에서 sec 직계 자식 인덱스 806~828(F-043
본문, idx805 "[16]" 캡션 제외)만 남기고 독립 문서로 분리한다. F-041/F-042와
동일한 구조(표지+수신/제목 표+본문+결재란).

허용 반영 범위: 학교명(C-01)·학년도(C-02)·관련문서(D-03)·문서번호(C-06)만
토큰화한다. 계약금액·계약상대자(업체명·소재지·대표자·전화번호)·계약보증금·
납품기한은 특정 계약 건의 실제 데이터(K-01/K-02 포함)라 정적 HWPX 단계에서는
원문 유지(F-005/F-006/F-041/F-042와 동일 스코프 결정). "나. 계약금액" 문구의
"○○벌(예정수량)"은 수량 placeholder(학교명·학년도 범주가 아님)이므로 손대지
않는다.

사용법: python extract_f043.py <원본.hwpx> <출력_템플릿.hwpx>
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


def replace_split_across_runs(root_elem, part_texts, new_text):
    """part_texts에 나열된 문자열이 순서대로 연속된 별도 <t> 노드에 걸쳐
    분할돼 있는 경우(F-041에서 실측 확인된 패턴)를 처리한다."""
    all_t = [t for t in root_elem.iter() if etree.QName(t.tag).localname == "t"]
    n = len(part_texts)
    for i in range(len(all_t) - n + 1):
        window = all_t[i:i + n]
        if all((w.text or "") == part_texts[j] for j, w in enumerate(window)):
            window[0].text = new_text
            for w in window[1:]:
                w.text = ""
            host = nearest_p(window[0])
            if host is not None:
                remove_own_lineseg(host)
            return
    raise AssertionError(f"분할 매치 실패: {part_texts!r}")


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

    if len(sec_children) < 829:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[805].iter() if etree.QName(t.tag).localname == "t"
    )
    if caption_text.strip() != "[ 16 ] 계약체결(예시)":
        raise AssertionError(f"idx805가 F-043 캡션이 아님: {caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(806, 829)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 806)]

    replace_text_anywhere(np(807), "○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(np(808), "20○○학년도", "{{학년도}}")
    replace_split_across_runs(np(809), ["○○", "학교-000(20○○.○.○.)"], "{{관련문서}}")
    replace_text_anywhere(np(810), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(811), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(828), "○○○○학교-", "{{문서번호}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{학년도}}", 3), ("{{관련문서}}", 1), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["20○○학년도", "○ ○ 학 교"]:
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
    print(f"PASS: F-043 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
