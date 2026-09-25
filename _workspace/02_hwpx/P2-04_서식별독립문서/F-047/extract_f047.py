"""F-047(교복구매 수요조사(신청 여부) 가정통신문) 원본 보존형 독립 HWPX 추출.
**개인정보 가능 문서(서식_매핑표.md D-02 등급 H)** — 학생/학부모 성명·연락처·
신청 여부 응답란은 원문에서도 이미 빈칸이며, 이 파이프라인은 그 응답란을
절대 건드리지 않는다(토큰 삽입도, 테스트값 채움도 하지 않음).

원본 sec 직계 자식 인덱스 882(F-047 본문 1개 문단, idx881 "[18-2]" 캡션
제외, idx883 "[18-3]"는 다음 서식 F-048 캡션)만 남기고 독립 문서로
분리한다. 표지·수신제목·결재란이 없는 1x1 표(안내문) + 내부 3x3 표(첨부
신청서 서식) 구조임을 원본 XML로 직접 확인함.

원본에는 서로 다른 여러 placeholder 표기가 섞여 있음을 직접 확인함(추측
금지, 문자 단위로 직접 확인):
  - "000학교"(숫자 0 3개, 제목·인사말 2회) — 학교명 자리로 보이나 다른
    서식들의 표준 "○○학교"(원문자 2개) 패턴과 문자 종류가 다름.
  - "○○○ 학교장"(원문자 3개+공백, 서명란) — 표준 학교장 서명 패턴과
    유사하나 원문자 개수(3개)가 다른 서식의 "○○학교장"(2개)과 다름.
  - "ooo, oo"(영문 소문자 o), "000원"/"000,000원"/"00,000원"(숫자 0),
    "00년 00월 00일"(숫자 0), "20〇〇학년도"(원문자 U+3007, 다른 원문자
    U+25CB와 또 다름) — 품목명·가격·날짜·"예비 신입생 학년도"(F-044/045/046과
    동일 이유로 C-02와 다른 "내년도" 값)이며 전부 학교마다 실제로 달라지는
    구체 데이터이거나 계산이 필요한 값이라 이 정적 단계에서 반영하지 않는다.
  - "이름( )", "보호자 성명( )", "전화번호( )" — 원문 자체가 이미 빈 칸인
    개인정보 응답란. 절대 채우지 않고 원문 그대로 둔다.

**이번 서식에서 실제로 토큰화하는 것은 "000학교"(2회, 학교명)와
"○○○ 학교장"(1회, 학교장 서명)뿐이며, 나머지는 전부 원문 그대로 유지한다.**

**"빈 2쪽" 결함 발견·회피(F-045와 동일 원인·동일 기법)**: 이 표도 rowCnt=1
colCnt=1의 절대 고정 높이(height=66350)라, 별도 carrier 빈 문단을 표
앞에 추가하면 그 미세한 줄 높이만으로 페이지 가용 공간을 넘겨 완전히
빈 2쪽째가 생김(순수 원본대조본만으로 재현). F-045와 동일하게 별도
carrier 문단을 추가하지 않고 secPr·ctrl을 본문 표 문단 자신의 run
맨 앞에 직접 삽입해 병합하는 방식으로 해결함.
"000학교"를 학교명으로 간주하는 것은, 이 서식 안에서 유일하게 "학교" 앞에
붙어 "제목/인사말에서 서술 주어" 역할을 하는 반복 패턴이고 서식_매핑표.md가
F-047의 반영 허용 범위에 C-01(학교명)을 명시하고 있기 때문이며, 개인정보
필드가 절대 아님을 원본 문맥(수요조사 안내 주체 = 학교)으로 재확인했다.

사용법: python extract_f047.py <원본.hwpx> <출력_템플릿.hwpx>
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


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f047")
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

    if len(sec_children) < 884:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[881].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 18-2 ]" not in caption_text:
        raise AssertionError(f"idx881이 F-047 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[883].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 18-3 ]" not in next_caption_text:
        raise AssertionError(f"idx883이 다음(F-048) 캡션이 아님: {next_caption_text!r}")

    # 캐리어 병합 기법(F-045에서 확립) — docstring "빈 2쪽" 결함 회피 설명 참고.
    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body = copy.deepcopy(sec_children[882])
    body_run = next(c for c in body if etree.QName(c.tag).localname == "run")
    body_run.insert(0, ctrl)
    body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(body)

    replace_text_anywhere(body, "000학교", "{{학교명}}", expected_count=2)
    replace_text_anywhere(body, "○○○ 학교장", "{{학교명}} 학교장", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 3, f"{{학교명}} 개수 불일치: 기대 3, 실제 {actual}"
    for stray in ["000학교", "○○○ 학교장"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"
    # 개인정보 응답란은 원문 그대로 빈칸이어야 함(방어적 재확인)
    for blank in ["이름 (           )", "보호자 성명(        )", "전화번호(            )"]:
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
    print(f"PASS: F-047 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
