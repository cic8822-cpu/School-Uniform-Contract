"""F-006(교복 학교주관구매 계획 학교운영위원회 심의(안)) 원본 보존형 독립 HWPX 추출.

원본 87쪽 매뉴얼 사본의 section0.xml에서 sec 직계 자식 인덱스 139~163(F-006
본문, idx138 "[3]" 캡션 제외)만 남기고 독립 문서로 분리한다.

Excel 구현(build_excel_v1_structure.ps1 7r절)과 대조한 반영 범위:
- 학교명(C-01, idx140·idx161)·학년도(C-02, idx161의 "20○○학년도")만 토큰화.
- "라. 협의사항"의 위원회 구성 표(idx158, 8행5열)는 F-001과 동일한 위원
  반복행 구조이나 위원 성명은 F-001~F-005와 동일하게 "○○○" 원문 그대로
  유지함(자동 반영 금지, R-02 마스킹 원칙).
- "다. 교복구매 상한가격"(idx155)은 Excel에서 기초자료입력!D29:D34 합계
  수식(K-01/K-02 계산 필드)으로 실제 금액을 계산하지만, 이 정적 HWPX
  단계에는 아직 그 계산값을 가져올 데이터 소스가 없어(③ Excel→HWPX 연동
  전) F-005와 동일한 스코프 결정에 따라 토큰화하지 않고 원문 placeholder
  ("OO,OOO원"/"OOO,OOO원", 라틴 대문자 O)를 그대로 유지함.
- idx146("1. 제안 이유"의 두 번째 문장)은 Excel B13 수식이 "학교명 학년도"를
  덧붙이지만, 원본 HWPX 원문 자체에는 이 문구가 없음(원문 직접 대조로 확인).
  이 파이프라인은 원본 바이트를 보존하고 기존 placeholder만 치환하는 것이
  목적이므로, 원문에 없는 문구를 새로 삽입하지 않고 원문 그대로 둔다
  (Excel의 편집상 보강과 HWPX 원문 사이의 의도적 불일치, 기록만 함).

F-002~F-005에서 확립한 안전 패턴 적용: 매치 리스트 선확정 후 변형.

사용법: python extract_f006.py <원본.hwpx> <출력_템플릿.hwpx>
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

    if len(sec_children) < 164:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[138].iter() if etree.QName(t.tag).localname == "t"
    )
    if caption_text.strip() != "[ 3 ] 교복 학교주관구매 계획 학교운영위원회 심의(안) (예시)":
        raise AssertionError(f"idx138이 F-006 캡션이 아님: {caption_text!r}")

    # --- 페이지 설정 문단(idx0)에서 secPr/ctrl만 남기고 로고(pic) 제거 ---
    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    # --- F-006 본문 범위(idx 139~163) 깊은 복사 ---
    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(139, 164)]

    # --- 새 sec 자식 목록으로 교체(맨 앞에 페이지설정 캐리어 삽입) ---
    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 139)]

    replace_text_anywhere(np(140), "○○학교", "{{학교명}}")
    replace_text_anywhere(np(161), "20○○학년도", "{{학년도}}")
    replace_text_anywhere(np(161), "○○학교", "{{학교명}}")

    # --- 잔존 검증 ---
    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 2), ("{{학년도}}", 1)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["○○학교", "20○○학년도"]:
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
    print(f"PASS: F-006 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
