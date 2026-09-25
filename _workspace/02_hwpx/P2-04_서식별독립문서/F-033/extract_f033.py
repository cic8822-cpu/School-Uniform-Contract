"""F-033(제안서 접수대장) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 660~666(F-033 본문, idx659 "[10_1]" 캡션 제외)만
남기고 독립 문서로 분리한다. idx664의 6행×11열 표는 헤더만 있는 빈 양식
(인스턴스 데이터 없음, 직접 텍스트 덤프로 확인)이다.

**빈 2쪽 결함 회피(F-032에서 처음 확인한 패턴 재적용)**: idx667~678(내용
없는 여백 문단 12개, paraPrIDRef=22)을 포함해 격리하면 원본은 1쪽인데
한컴이 빈 2쪽을 추가로 만들어냄(원본대조본만으로 재현). 본문 범위를
idx666(각주 2번째 줄)까지로 좁혀 1쪽으로 정상화됨을 격리 실험으로 확인함.

허용 반영 범위: 학교명(C-01, idx663 "○○"+"학교" 두 개 별도 `<hp:t>` 노드에
걸쳐 분할됨 — `replace_split_across_runs` 필요, F-041 idx772와 같은 함정)·
학년도(C-02, idx660)만 토큰화한다. idx664 표는 빈 양식이라 데이터 없음.

사용법: python extract_f033.py <원본.hwpx> <출력_템플릿.hwpx>
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


def replace_split_across_runs(root_elem, part_texts, new_text):
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

    if len(sec_children) < 667:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[659].iter() if etree.QName(t.tag).localname == "t"
    )
    if "10_1" not in caption_text:
        raise AssertionError(f"idx659이 F-033 캡션이 아님: {caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(660, 667)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 660)]

    replace_text_anywhere(np(660), "20○○학년도", "{{학년도}}")
    replace_split_across_runs(np(663), ["○○", "학교"], "{{학교명}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학년도}}", 1), ("{{학교명}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"

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
    print(f"PASS: F-033 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
