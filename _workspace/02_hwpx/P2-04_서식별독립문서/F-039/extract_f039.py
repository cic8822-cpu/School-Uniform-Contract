"""F-039(교복 구매 업체별 제안서 평가표) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 756~762(F-039 본문, idx755 "[13-3]" 캡션 제외,
idx763 "[13-4]"는 이미 완료된 F-040 캡션)만 남기고 독립 문서로 분리한다.
F-034~036과 달리 이 서식은 "표지 스탬프+수신제목 표+결재란"이 전혀 없는
문서 전용 평가표다(제목 + 11x9 평가표 + 각주만 존재, 원본 XML로 직접
확인). 학교명·관련문서·문서번호 토큰화 대상 자체가 없다.

학년도는 idx757(제목) 단 1회만 존재("20○○학년도")한다. idx759 표
데이터행의 "○○"(t[15])는 학년도가 아니라 위원명(개인정보 마스킹) 자리라
절대 건드리지 않는다(원본 XML로 직접 확인 — 추측 금지, 8.2절 공통 교훈
4번).

허용 반영 범위: 학년도(C-02)만 토큰화한다. 업체명·위원명·항목별 점수·
총계·평균은 특정 입찰 회차·위원회의 실제 평가 데이터이거나 개인정보
마스킹 대상이라 원문 유지한다.

사용법: python extract_f039.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f039")
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

    if len(sec_children) < 764:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[755].iter() if etree.QName(t.tag).localname == "t"
    )
    if "13-3" not in caption_text:
        raise AssertionError(f"idx755가 F-039 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[763].iter() if etree.QName(t.tag).localname == "t"
    )
    if "13-4" not in next_caption_text:
        raise AssertionError(f"idx763이 다음(F-040) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(756, 763)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 756)]

    replace_text_anywhere(np(757), "20○○학년도", "{{학년도}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학년도}}")
    assert actual == 1, f"{{{{학년도}}}} 개수 불일치: 기대 1, 실제 {actual}"
    assert "20○○학년도" not in all_text, "치환되지 않은 '20○○학년도' 잔존 문자열 발견"

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
    print(f"PASS: F-039 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
