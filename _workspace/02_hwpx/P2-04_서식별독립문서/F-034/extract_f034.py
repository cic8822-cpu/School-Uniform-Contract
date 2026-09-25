"""F-034(교복 제안서 평가위원회 개최) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 680~697(F-034 본문, idx679 "[11]" 캡션 제외,
idx698 "[12]"는 다음 서식 F-035 캡션)만 남기고 독립 문서로 분리한다.
F-001/F-004/F-032/F-041과 동일한 "표지 스탬프(1x3표)+수신/제목 표(3x2표)+
본문 문단+결재란(8x41표)" 구조임을 원본 XML로 직접 확인함(idx681 tbl
rowCnt=1 colCnt=3, idx682 tbl rowCnt=3 colCnt=2, idx697 tbl rowCnt=8
colCnt=41). F-002/F-003 계열의 위험한 고정폭 단일 셀 표가 아니다.

idx690~696(빈 문단 7개, paraPrIDRef=44)은 idx697(결재란, 역시 paraPrIDRef=44)
"이전"에 위치하고 본문 범위(idx680~697) 안에 전부 포함되므로, F-032에서
발견된 "캡션 직전 트레일링 빈 문단으로 인한 빈 2쪽" 패턴과는 다르다(그
패턴은 idx698 캡션 직전에 트레일링 빈 문단이 있고 그걸 잘라내는 경우였음).
여기서는 idx697 결재란 자체가 범위의 마지막이라 트레일링 빈 문단이 없음
-- 다만 한컴 육안 확인 단계에서 재차 확인 필요.

허용 반영 범위(F-032/F-041과 동일 원칙): 학교명(C-01)·학년도(C-02)·
관련문서(D-03)·문서번호(C-06)만 토큰화한다. idx685~687(일시·장소·대상)·
idx688(위원 명단 표, 성명은 "○○○" 마스킹 원문 그대로)·idx689(안내방법)·
idx697 결재란의 우편번호·주소·전화·팩스·이메일은 행사 고유 데이터이거나
전북특별자치도교육청 고정 서식 정보라 원문 유지한다.

사용법: python extract_f034.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f034")
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

    if len(sec_children) < 699:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[679].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 11 ]" not in caption_text:
        raise AssertionError(f"idx679이 F-034 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[698].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 12 ]" not in next_caption_text:
        raise AssertionError(f"idx698이 다음(F-035) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(680, 698)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 680)]

    replace_text_anywhere(np(681), "○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(np(682), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(683), "○○학교-000(20○○.○.○)", "{{관련문서}}")
    replace_text_anywhere(np(684), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(697), "○○○○학교-", "{{문서번호}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 1), ("{{학년도}}", 2), ("{{관련문서}}", 1), ("{{문서번호}}", 1)]:
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
    print(f"PASS: F-034 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
