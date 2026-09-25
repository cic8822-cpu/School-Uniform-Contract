"""F-048(교복 구매 안내(신청 수량 파악) 가정통신문) 원본 보존형 독립 HWPX
추출. 개인정보 응답란(이름·보호자 성명·전화번호)이 포함된 신청서 서식이라
F-047과 동일한 수준으로 신중히 다룬다 — 원문에서도 이미 빈칸인 응답란은
절대 건드리지 않는다.

원본 sec 직계 자식 인덱스 884(F-048 본문 1개 문단, idx883 "[18-3]" 캡션
제외, idx885 "[19]"는 다음 서식 F-049 캡션, 이번 범위 밖)만 남기고 독립
문서로 분리한다. F-047과 동일하게 표지·수신제목·결재란이 없는 1x1 표
(안내문) + 내부 8x8 표(품목별 신청 수량 서식) 구조.

원본에서 직접 확인한 placeholder 혼재 양상(추측 금지):
  - "00학교"(숫자 0 2개, 제목 1회만) — F-047의 "000학교"(0 3개)와도 다른
    또 다른 숫자 표기. 이 문서 안에서도 "○○학교"(원문자 2개, 인사말·서명
    2곳)와 서로 다른 문자를 섞어 쓰고 있음(원본 매뉴얼 자체의 편집 불일치로
    판단, 고치지 않고 그대로 기록만 함).
  - "○○학교"(원문자 2개) — 인사말 1회("○○학교 입학을...")·서명 2회
    ("○○학교장", "○○학교장 귀하") 총 3회. 이것만 학교명(C-01)으로 토큰화.
  - "2027년 O월 O일"(영문 대문자 O), "2027학년도"(구체 연도 고정, F-044~046과
    동일 이유로 C-02와 무관), "OOO(연락처: OOO-OOO-OOO)"(업체명·연락처,
    영문 대문자 O 반복), "000,000원"(숫자 0, 단가), "**2026학년도 충북교육청**
    교복 학교주관구매 권고 상한가"(원문 그대로 **충북교육청**이라고 적혀
    있음 — 이 매뉴얼의 다른 모든 곳은 "전북특별자치도교육청"인데 이 문단만
    다른 도 이름이 그대로 남아있는 원본 자체의 명백한 편집 오류로 추정됨.
    원본 보존 원칙에 따라 고치지 않고 그대로 두되 로그에 기록함) — 전부
    학교마다 달라지는 구체 데이터이거나 원본 자체의 오기이므로 반영하지
    않고 원문 그대로 유지한다.
  - "이름 (           )", "보호자 성명 (        )", "전화번호 (           )"
    — 원문 자체가 이미 빈칸인 개인정보 응답란. 절대 채우지 않는다.

사용법: python extract_f048.py <원본.hwpx> <출력_템플릿.hwpx>
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
    바꾸고, **바로 다음 <hp:t> 노드**가 next_prefix_to_strip으로 시작하면 그
    접두 문자열을 제거한다. F-048에서 실측 확인한 함정: "○○학교"가 "○○"
    (단독 런)과 "학교 입학을..." (다음 런의 접두 부분)으로 분할돼 있어서,
    "○○" 런만 교체하고 다음 런을 그대로 두면 "학교"가 중복 출력된다
    (예: "검증초등학교"+"학교 입학을..." → "검증초등학교학교 입학을...").
    이 함수는 그 중복을 방지하기 위해 다음 런의 접두 "학교"까지 함께 제거한다."""
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
        # 바로 다음 t 노드의 접두 문자열 제거(같은 문단 내에서만 — 문단 경계를
        # 넘어 엉뚱한 텍스트를 지우지 않도록 host가 같은 문단인지 확인).
        if i + 1 < len(all_t):
            nxt = all_t[i + 1]
            if (nxt.text or "").startswith(next_prefix_to_strip):
                nxt_host = nearest_p(nxt)
                if nxt_host is host:
                    nxt.text = nxt.text[len(next_prefix_to_strip):]
                    remove_own_lineseg(nxt_host)
                else:
                    raise AssertionError(
                        f"'{old_text}' 다음 런이 다른 문단에 있어 접두 제거를 건너뜀(직접 재확인 필요)"
                    )
            else:
                raise AssertionError(
                    f"'{old_text}' 다음 런이 '{next_prefix_to_strip}'로 시작하지 않음: {nxt.text!r}"
                )
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f048")
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

    if len(sec_children) < 886:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[883].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 18-3 ]" not in caption_text:
        raise AssertionError(f"idx883이 F-048 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[885].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 19 ]" not in next_caption_text:
        raise AssertionError(f"idx885가 다음(F-049) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[884])]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)
    body = new_children[1]

    # 인사말("○○"+"학교 입학을...")과 서명2("○○"+"학교장 귀하")는 "○○"만
    # 별도 <hp:t> 런으로 분리돼 있고, 다음 런 접두의 "학교"까지 합쳐야 온전한
    # "○○학교" placeholder임(직접 확인) — 다음 런의 "학교" 접두까지 제거해야
    # "학교"가 중복 출력되지 않는다(육안 PDF 확인으로 실제 발견한 결함, 최초
    # 시도에서 "○○"만 바꾸고 다음 런을 그대로 둬 "검증초등학교학교 입학을..."
    # 중복 결함이 발생했었음 — 이 함수로 수정).
    replace_exact_node_text_with_suffix_strip(body, "○○", "{{학교명}}", "학교")
    # 서명1("○○학교장")은 단일 런 안에 붙어있음(직접 확인) — 남은 1건만 매치.
    replace_text_anywhere(body, "○○학교", "{{학교명}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 3, f"{{학교명}} 개수 불일치: 기대 3, 실제 {actual}"
    for stray in ["○○학교", "○○"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"
    # 개인정보 응답란은 원문 그대로 빈칸이어야 함(방어적 재확인)
    for blank in ["이름 (           )", "보호자 성명 (        )", "전화번호 (           )"]:
        assert blank in all_text, f"개인정보 응답란 원문이 예상과 달라짐(직접 재확인 필요): {blank!r}"

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
    print(f"PASS: F-048 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
