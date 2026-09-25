"""F-012(계약 특수조건, 제1조~제19조) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 419~503(F-012 본문, idx418 "[9-1]" 캡션 제외,
idx504~521은 18개 연속 빈 문단이라 F-032류 "빈 2쪽" 위험을 피하기 위해
범위에서 제외, idx522 "[9-2]"는 다음 서식 F-013 캡션)만 남기고 독립
문서로 분리한다. 표 없이 제1조~제19조 순수 텍스트(계약 일반조항)만 있는
구조.

원본에서 직접 확인한 placeholder: idx419(제목) "20○○학년도 ◯◯◯◯학교",
idx421(전문) "20○○학년도"(단독 런)·"◯◯◯◯"+"학교 교복..."(런 분할)·
"◯◯◯◯"+"학교장(이하..."(런 분할, "학교"만 제거하면 "장(이하..."이 남아
"{{학교명}}장(이하..."이 되도록 처리), idx424(제1조①) "20○○학년도".

허용 반영 범위: 학교명(C-01, 3회)·학년도(C-02, 3회)만 토큰화한다.
idx421의 "계약상대자 ○○교복사 대표 ○○○"(업체명·대표자, 특정 계약 건의
실제 데이터)·idx437/438의 "20○○.OO.OO(요일)까지"(납품기한, 구체 이벤트
날짜)는 F-005/041/042와 동일 원칙(계약상대자 미정·구체 이벤트는 정적
단계에서 원문 유지)에 따라 손대지 않는다. 나머지 제2조~제19조 조항 전체
(치수측정·납품·검사검수·하자보증·사후관리 등)는 계약 표준 문구이며
학교명·학년도 placeholder가 없음을 직접 XML 확인함(추측 금지).

사용법: python extract_f012.py <원본.hwpx> <출력_템플릿.hwpx>
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


def replace_exact_node_text_with_suffix_strip(root_elem, old_text, new_text, next_prefix_to_strip, expected_matches=None):
    """<hp:t> 노드의 text가 old_text와 정확히 일치하는 모든 독립 런을 찾아
    new_text로 바꾸고, 각각 바로 다음 <hp:t> 노드가 next_prefix_to_strip으로
    시작하면 그 접두 문자열만 제거한다(F-048에서 확립한 기법, F-012의
    "학교"/"학교장" 두 접미사를 모두 "학교" 하나만 제거하는 방식으로
    통일 처리 — "학교장"에서 "학교"만 제거하면 "장"이 남아 결과적으로
    "{{학교명}}장"이 되어 올바름)."""
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
    if expected_matches is not None and matched != expected_matches:
        raise AssertionError(f"독립 런 '{old_text}' 매치 수 불일치: 기대 {expected_matches}, 실제 {matched}")
    if not matched:
        raise AssertionError(f"독립 런 '{old_text}' 를 찾지 못함")
    return matched


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f012")
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

    if len(sec_children) < 523:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[418].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-1 ]" not in caption_text:
        raise AssertionError(f"idx418이 F-012 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[522].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-2 ]" not in next_caption_text:
        raise AssertionError(f"idx522가 다음(F-013) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(419, 504)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 419)]

    replace_text_anywhere(np(419), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_text_anywhere(np(419), "◯◯◯◯학교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(421), "20○○학년도", "{{학년도}}", expected_count=1)
    replace_exact_node_text_with_suffix_strip(np(421), "◯◯◯◯", "{{학교명}}", "학교", expected_matches=2)
    replace_text_anywhere(np(424), "20○○학년도", "{{학년도}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    for token, expected in [("{{학교명}}", 3), ("{{학년도}}", 3)]:
        actual = all_text.count(token)
        assert actual == expected, f"{token} 개수 불일치: 기대 {expected}, 실제 {actual}"
    for stray in ["20○○학년도", "◯◯◯◯학교", "◯◯◯◯"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"
    assert "계약상대자 ○○교복사 대표 ○○○" in all_text, "계약상대자 필드가 예상과 달라짐(직접 재확인 필요)"

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
    print(f"PASS: F-012 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
