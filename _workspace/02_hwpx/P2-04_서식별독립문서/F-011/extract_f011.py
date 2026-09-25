"""F-011(교복 학교주관구매 입찰 공고) 원본 보존형 독립 HWPX 추출.
**고난도 조판(167개 문단, 가장 긴 서식)** — 표 1개(idx259)+순수 텍스트
리스트(12개 대항목, 입찰방법·자격·제출서류 등) 구조.

원본 sec 직계 자식 인덱스 251~407(F-011 본문, idx250 "[9]" 캡션 제외,
idx408~417은 결재란류 콘텐츠 없이 10개 연속 빈 문단(paraPrIDRef=1)이라
F-032류 "빈 2쪽" 위험을 피하기 위해 범위에서 제외, idx418 "[9-1]"는 다음
서식 F-012 캡션)만 남기고 독립 문서로 분리한다.

**신규 발견(원문자 종류 3중 혼재, 추측 금지 원칙에 따라 직접 XML로 전수
확인)**: 이 서식은 학교명 placeholder에 **서로 다른 원문자를 섞어
쓴다** — "◯◯◯◯학교"(U+25EF LARGE CIRCLE, idx252/254/256/259(2회)/284/343/
381/382, 총 8회)가 압도적으로 많고, "○○○○학교"(U+25CB WHITE CIRCLE,
idx259 "납품장소○○○○학교" 1회)만 다른 문자를 쓰며, 서명란은 공백 포함
"◯ ◯ ◯ ◯ 학 교 장"(U+25EF, idx380/407, 2회)이다. **idx328 "본교 ◯◯◯◯실"은
학교명이 아니라 회의실 이름 placeholder이므로 절대 학교명으로 오인해
치환하지 않는다**(F-039 위원명 오인 방지 사례와 동일 함정, 문맥으로
직접 확인).

허용 반영 범위: 학교명(C-01, 위 3가지 원문자 변형 총 12회 — "◯◯◯◯학교"
단일 런 6회(idx252/259/284/343/381/382) + "◯◯◯◯"+"학교" 분할 3회
(idx254/256/259) + 서명란 2회(idx380/407) + "○○○○학교" 1회(idx259
납품장소))·학년도(C-02,
"20○○학년도" U+25CB, idx254/256/259/381/382, 총 5회)만 토큰화한다.
idx252 "공고 제20○○ – 00호"(공고번호, 매핑표에 대응 필드 없음)·idx259
표 안의 "000,000원"(기초금액)·"000벌"(구매예정수량)·납품기한 날짜들·
idx282/322/327/343/378/405의 "20○○.00.00.(요일)"류 이벤트 날짜(제출기간·
평가일시·가격개찰일시·게시일자)는 특정 회차의 실제 데이터라 원문 유지한다
(F-005/009/010과 동일 원칙).

사용법: python extract_f011.py <원본.hwpx> <출력_템플릿.hwpx>
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


def replace_exact_node_text_with_suffix_strip(root_elem, old_text, new_text, next_prefix_to_strip):
    """<hp:t> 노드의 text가 old_text와 정확히 일치하는 독립 런을 찾아 new_text로
    바꾸고, 바로 다음 <hp:t> 노드가 next_prefix_to_strip으로 시작하면 그 접두
    문자열을 제거한다(F-048에서 확립한 "학교" 중복 방지 기법 재사용)."""
    all_t = [t for t in root_elem.iter() if etree.QName(t.tag).localname == "t"]
    matched = 0
    for i, t in enumerate(all_t):
        if (t.text or "") != old_text:
            continue
        t.text = new_text
        host = nearest_p(t)
        if host is not None:
            remove_own_lineseg(host)
        matched += 1
        if i + 1 < len(all_t):
            nxt = all_t[i + 1]
            if (nxt.text or "").startswith(next_prefix_to_strip):
                nxt_host = nearest_p(nxt)
                if nxt_host is host:
                    nxt.text = nxt.text[len(next_prefix_to_strip):]
                    remove_own_lineseg(nxt_host)
                else:
                    raise AssertionError(f"'{old_text}' 다음 런이 다른 문단(직접 재확인 필요)")
            else:
                raise AssertionError(f"'{old_text}' 다음 런이 '{next_prefix_to_strip}'로 시작 안 함: {nxt.text!r}")
    if not matched:
        raise AssertionError(f"독립 런 '{old_text}' 를 찾지 못함")
    return matched


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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f011")
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

    if len(sec_children) < 419:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[250].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9 ]" not in caption_text:
        raise AssertionError(f"idx250이 F-011 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[418].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-1 ]" not in next_caption_text:
        raise AssertionError(f"idx418이 다음(F-012) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(251, 408)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 251)]

    # 학교명(◯◯◯◯학교, U+25EF) - 단일 런 안에 통째로 있는 경우(직접 확인)
    replace_text_anywhere(np(252), "◯◯◯◯학교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(284), "◯◯◯◯학교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(343), "◯◯◯◯학교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(381), "◯◯◯◯학교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(382), "◯◯◯◯학교", "{{학교명}}", expected_count=1)
    # idx259 run1 안의 "◯◯◯◯학교"(단일 런, 입찰건명 줄)
    replace_text_anywhere(np(259), "◯◯◯◯학교", "{{학교명}}", expected_count=1)
    # 학교명(◯◯◯◯ + "학교..." 런 분할, F-048에서 확립한 접두 제거 기법)
    replace_exact_node_text_with_suffix_strip(np(254), "◯◯◯◯", "{{학교명}}", "학교")
    replace_exact_node_text_with_suffix_strip(np(256), "◯◯◯◯", "{{학교명}}", "학교")
    replace_exact_node_text_with_suffix_strip(np(259), "◯◯◯◯", "{{학교명}}", "학교")
    # 학교명(○○○○학교, U+25CB) - 1회, idx259 납품장소(단일 런)
    replace_text_anywhere(np(259), "○○○○학교", "{{학교명}}", expected_count=1)
    # 학교명(서명란, 공백 포함 ◯ ◯ ◯ ◯ 학 교 장) - 2회(단일 런)
    replace_text_anywhere(np(380), "◯ ◯ ◯ ◯ 학 교 장", "{{학교명}}장", expected_count=1)
    replace_text_anywhere(np(407), "◯ ◯ ◯ ◯ 학 교 장", "{{학교명}}장", expected_count=1)
    # 학년도(20○○학년도) - 5회(전부 단일 런 안에 있음, 직접 확인)
    replace_text_anywhere(np(254), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(256), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(259), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(381), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(382), "20○○학년도", "{{학년도}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 12), ("{{학년도}}", 5)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["◯◯◯◯학교", "○○○○학교", "◯ ◯ ◯ ◯ 학 교 장", "20○○학년도"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"
    # idx328 회의실명은 절대 손대지 않았어야 함(방어적 재확인)
    assert "본교 ◯◯◯◯실" in all_text, "idx328 회의실 placeholder가 예상과 달라짐(직접 재확인 필요)"

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
    print(f"PASS: F-011 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
