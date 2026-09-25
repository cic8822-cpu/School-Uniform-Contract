"""F-001(교복선정위원회 구성 기안) 원본 보존형 독립 HWPX 추출 스크립트.

원본 87쪽 매뉴얼 사본(artifacts/hwpx/P0-02_원본사본.hwpx)의 section0.xml에서
sec 직계 자식 인덱스 14~33(F-001 본문, [1] 캡션 제외)만 남기고 독립 문서로
분리한 뒤 허용 공통값 자리(학교명/학년도/문서번호/발행일)에만 토큰을 삽입한다.
원본 인덱스 0(secPr 페이지 설정 문단)은 표지 로고(pic)만 제거하고 secPr/ctrl은
그대로 유지해 A4 페이지 설정을 보존한다. header.xml/mimetype/META-INF/BinData/
Preview는 손대지 않는다.

사용법: python extract_f001.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree

NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def qn(local):
    return f"{{{NS_HP}}}{local}"


def remove_own_lineseg(p):
    for c in list(p):
        if etree.QName(c.tag).localname == "linesegarray":
            p.remove(c)


def nearest_p(elem):
    cur = elem.getparent()
    while cur is not None and etree.QName(cur.tag).localname != "p":
        cur = cur.getparent()
    return cur


def replace_text_anywhere(p_elem, old_text, new_text):
    replaced = 0
    for t in p_elem.iter():
        if etree.QName(t.tag).localname == "t" and t.text and old_text in t.text:
            t.text = t.text.replace(old_text, new_text)
            replaced += 1
            host = nearest_p(t)
            if host is not None:
                remove_own_lineseg(host)
    if replaced == 0:
        raise AssertionError(f"'{old_text}' 를 찾지 못함")
    return replaced


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

    if len(sec_children) < 34:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    # --- 페이지 설정 문단(idx0)에서 secPr/ctrl만 남기고 로고(pic) 제거.
    #     표지 전용 charPr/paraPr(테두리·큰 글자)를 그대로 쓰면 A4 페이지 설정은
    #     살지만 빈 줄 하나가 본문 크기로 남아 마지막 표가 다음 쪽으로 밀린다.
    #     실측 검증: charPr를 본문 크기(52)로, paraPr를 본문 문단(idx18 것)으로
    #     바꾸면 원본과 동일하게 1쪽에 들어감(2026-09-22 F-001 파일럿에서 확정). ---
    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    body_char_pr_ref = sec_children[17][0].get("charPrIDRef")  # idx17(본문) run의 charPrIDRef
    body_para_pr_ref = sec_children[18].get("paraPrIDRef")  # idx18(본문) 문단의 paraPrIDRef
    carrier_run.set("charPrIDRef", body_char_pr_ref)
    carrier.set("paraPrIDRef", body_para_pr_ref)

    # --- F-001 본문 범위(idx 14~33) 깊은 복사 ---
    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(14, 34)]

    # --- 새 sec 자식 목록으로 교체(맨 앞에 페이지설정 캐리어 삽입) ---
    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)
    # body_paragraphs는 new_children[1:]에 대응 (idx14~33 -> new idx1~20)
    idx15 = new_children[1 + (15 - 14)]
    idx16 = new_children[1 + (16 - 14)]
    idx17 = new_children[1 + (17 - 14)]
    idx18 = new_children[1 + (18 - 14)]
    idx33 = new_children[1 + (33 - 14)]

    replace_text_anywhere(idx15, "○ ○ 학 교", "{{학교명}}")
    replace_text_anywhere(idx16, "○○학년도 교복선정위원회 구성(안)", "{{학년도}} 교복선정위원회 구성(안)")
    replace_text_anywhere(idx17, "○○학교-000(20○○.○.○)", "{{문서번호}}({{발행일}})")
    replace_text_anywhere(
        idx18,
        "○○학년도 ○○학교 교복선정을 위한 교복선정위원회를 아래와 같이 구성하고자 합니다.",
        "{{학년도}} {{학교명}} 교복선정을 위한 교복선정위원회를 아래와 같이 구성하고자 합니다.",
    )
    replace_text_anywhere(idx33, "○○○○학교-", "{{문서번호}}")

    # --- 잔존 검증: 토큰 4종이 정확히 1회씩만 존재해야 함 ---
    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 2), ("{{학년도}}", 2), ("{{문서번호}}", 2), ("{{발행일}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"

    tree.write(section_path, xml_declaration=True, encoding=tree.docinfo.encoding or "UTF-8", standalone=tree.docinfo.standalone)

    # --- 재패키징: mimetype 비압축 우선 ---
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
    print(f"PASS: F-001 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
