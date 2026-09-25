"""F-045(교복구매 사전 안내 가정통신문) 원본 보존형 독립 HWPX 추출.

원본 sec 직계 자식 인덱스 855(F-045 본문 1개 문단, idx854 "[17-1]" 캡션
제외, idx856 "[18]"는 다음 서식 F-046 캡션)만 남기고 독립 문서로 분리한다.
F-002/F-003/F-040과 달리 표지·수신제목·결재란이 전혀 없이, 1x1 고정폭
표(width=50444, height=66147, 거의 A4 전면) 안에 전체 가정통신문 본문
(508자)이 통째로 들어있는 "문서 전용 안내장" 구조임을 원본 XML로 직접
확인함.

허용 반영 범위: 학교명(C-01, 본문 말미 서명란 "○○학교"+"장" 접미사, 2개 원)만
토큰화한다. "○○학교"만 {{학교명}}으로 바꾸고 뒤의 "장"(직함 접미사)은
원문 그대로 남겨 "{{학교명}}장"(골든값 반영 시 "검증초등학교장")이 되게 한다.

**"빈 2쪽" 결함 발견·회피(신규 패턴, F-032와 원인이 다름)**: 이 표는
rowCnt=1 colCnt=1의 절대 고정 높이(height=66147 HWPUNIT, 페이지 가용
높이의 대부분을 차지)라, 다른 서식들처럼 secPr/ctrl을 담은 별도 "carrier"
빈 문단을 표 앞에 추가하면 그 빈 문단의 미세한 줄 높이만으로도 페이지
가용 공간을 넘겨 완전히 빈 2쪽째가 생김(치환 여부와 무관하게 순수
원본대조본만으로도 재현, 격리 실험으로 확정). **해결**: 별도 carrier
문단을 추가하지 않고, secPr·ctrl을 본문 표 문단 자신의 run 맨 앞에
직접 삽입해 병합함(캐리어 병합 기법). 이 기법 적용 후 1쪽으로 정상화됨을
확인함. F-047(같은 구조, height=66350)도 동일 원인·동일 기법으로 처리. 본문 중 "2027학년도"는 F-044와 동일하게 20○○/○○ placeholder
문자가 없는 고정 예시 연도이며, 이 서식의 성격("예비 신입생" 대상이라
현재 구매 학년도 C-02보다 1년 뒤를 가리킴)상 C-02와 동일 값이 아니므로
토큰화하지 않고 원문 유지한다(F-044와 동일 판단 근거).

사용법: python extract_f045.py <원본.hwpx> <출력_템플릿.hwpx>
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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f045")
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

    if len(sec_children) < 857:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[854].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 17-1 ]" not in caption_text:
        raise AssertionError(f"idx854가 F-045 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[856].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 18 ]" not in next_caption_text:
        raise AssertionError(f"idx856이 다음(F-046) 캡션이 아님: {next_caption_text!r}")

    # 캐리어 병합 기법: 별도 carrier 빈 문단을 추가하지 않고, secPr/ctrl을
    # 본문 표 문단 자신의 run 맨 앞에 직접 삽입한다(위 docstring "빈 2쪽"
    # 결함 회피 설명 참고).
    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body = copy.deepcopy(sec_children[855])
    body_run = next(c for c in body if etree.QName(c.tag).localname == "run")
    body_run.insert(0, ctrl)
    body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(body)

    replace_text_anywhere(body, "○○학교", "{{학교명}}")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 1, f"{{학교명}} 개수 불일치: 기대 1, 실제 {actual}"
    assert "{{학교명}}장" in all_text, "'장' 접미사가 토큰 뒤에 붙어있지 않음(치환 위치 오류)"
    assert "○○학교" not in all_text, "치환되지 않은 '○○학교' 잔존 문자열 발견"

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
    print(f"PASS: F-045 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
