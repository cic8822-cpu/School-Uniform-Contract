"""F-032(교복구매 제안서 접수 결과) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 639~658(F-032 본문, idx638 "[10]" 캡션 제외)만
남기고 독립 문서로 분리한다. F-001/F-004/F-041과 동일한 "표지 스탬프(1x3표)+
수신/제목 표(3x2표)+본문 문단+결재란(8x41표)" 구조이며 F-002/F-003의 위험한
고정폭 단일 셀 표가 아님을 원본 XML로 확인함.

허용 반영 범위(F-041과 동일 원칙): 학교명(C-01)·학년도(C-02)·관련문서(D-03)·
문서번호(C-06)만 토큰화한다. idx646의 접수번호/접수일자/업체명/대표자 표와
idx656 결재란의 우편번호·주소·전화·팩스·이메일은 Excel 미연동 인스턴스/고정
서식 데이터라 원문 유지한다(F-001/F-041과 동일 스코프 결정).

**빈 2쪽 결함 발견·회피(2026-09-22 F-032에서 최초 확인)**: idx657~658(빈
문단, paraPrIDRef=25 — F-037/F-038의 "표제 앞 여백" 스타일과 동일)을 포함해
격리하면 원본은 1쪽인데 한컴 렌더링이 완전히 빈 2쪽을 만들어냄(원본
대조본만으로도 재현, 토큰 치환과 무관). 두 문단 모두 텍스트가 빈 문자열임을
직접 확인했고, 본문 범위를 idx656(결재란 표)까지로 좁혀 제외하니 1쪽으로
정상화됨(격리 실험으로 검증). 원본 87쪽 전체 문서에서는 이 대형 스타일
빈 줄이 다음 서식([10_1])의 여백으로 자연스럽게 흡수되지만, 독립 문서로
잘라내면 그 여백 스타일의 큰 줄높이가 페이지 하단에서 넘쳐 빈 쪽을
만드는 것으로 추정함. 다른 Form ID 추출 시에도 캡션 직전 마지막 문단들의
paraPrIDRef를 확인해 유사 패턴이면 제외 여부를 검토할 것.

사용법: python extract_f032.py <원본.hwpx> <출력_템플릿.hwpx>
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

    if len(sec_children) < 659:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[638].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 10 ]" not in caption_text:
        raise AssertionError(f"idx638이 F-032 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[659].iter() if etree.QName(t.tag).localname == "t"
    )
    if "10_1" not in next_caption_text:
        raise AssertionError(f"idx659이 다음(F-033) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(639, 657)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 639)]

    replace_text_anywhere(np(640), "○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(np(641), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(642), "○○학교-000(20○○.○.○)", "{{관련문서}}")
    replace_text_anywhere(np(643), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(648), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(656), "○○○○학교-", "{{문서번호}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{학년도}}", 3), ("{{관련문서}}", 1), ("{{문서번호}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["20○○학년도", "○ ○ 학 교", "○○○○학교-"]:
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
    print(f"PASS: F-032 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
