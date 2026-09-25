"""F-050(만족도 조사 설문지) 원본 보존형 독립 HWPX 추출.
**개인정보 가능 문서(서식_매핑표.md H등급, D-02)** — 응답자 성명·연락처
필드는 원문 자체에 없음(직접 확인, F-047/048과 달리 순수 익명 리커트 척도
설문지+자유서술 "기타 의견"란뿐)이지만, 응답 결과(합계/평균 점수, 기타
의견 서술)는 특정 학교·특정 조사 회차의 실제 데이터이므로 이 정적 단계에서
절대 채우지 않고 원문 그대로 빈칸(공란)으로 둔다.

원본 sec 직계 자식 인덱스 911~919(F-050 본문, idx910 "[19-1]" 캡션 제외,
idx920 "[20]"는 다음 서식 F-051 캡션, 이번 작업 범위 밖)만 남기고 독립
문서로 분리한다. 표지·수신제목·결재란이 없고, 제목(idx912)+안내문 1x1
CELL 표(idx914)+안내(idx916)+15x7 설문 표(idx917, 척도 문항)+안내(idx918)+
"기타 의견" 1x1 CELL 표(idx919) 구조.

허용 반영 범위: 학교명(C-01, idx914 "OOO학교"(영문 대문자 O 3개, 다른
서식들의 원문자 "○○○"·숫자 "000"과 또 다른 표기, 직접 XML로 확인) —
"교복선정위원회" 명의 앞에만 등장)만 토큰화한다. idx917 설문 표의
"합계 ( )점"·"평균 ( )점" 응답 칸과 idx919 "기타 의견" 서술란은 원문에서도
이미 빈칸이며 이 파이프라인이 손대지 않는다(응답 데이터 미반영 원칙).

**"빈 2쪽" 결함 회피(F-045/047과 동일 원인·동일 기법)**: idx914·idx919가
1x1 rowCnt=1 colCnt=1 pageBreak=CELL 절대 고정 높이 표라 별도 carrier 빈
문단을 추가하면 "빈 2쪽" 위험이 있어(F-045/047에서 확립한 패턴), 처음부터
캐리어 병합 기법(secPr·ctrl을 본문 첫 문단 자신의 run에 직접 삽입)을
적용한다.

사용법: python extract_f050.py <원본.hwpx> <출력_템플릿.hwpx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree


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
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f050")
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

    if len(sec_children) < 921:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[910].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 19-1 ]" not in caption_text:
        raise AssertionError(f"idx910이 F-050 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[920].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 20 ]" not in next_caption_text:
        raise AssertionError(f"idx920이 다음(F-051, 범위 밖) 캡션이 아님: {next_caption_text!r}")

    # 캐리어 병합 기법(F-045에서 확립) — docstring "빈 2쪽" 결함 회피 설명 참고.
    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(911, 920)]
    first_body = body_paragraphs[0]
    first_body_run = next(c for c in first_body if etree.QName(c.tag).localname == "run")
    first_body_run.insert(0, ctrl)
    first_body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[orig_idx - 911]

    replace_text_anywhere(np(914), "OOO학교", "{{학교명}}", expected_count=1)

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 1, f"{{학교명}} 개수 불일치: 기대 1, 실제 {actual}"
    assert "OOO학교" not in all_text, "치환되지 않은 'OOO학교' 잔존 문자열 발견"
    # 응답 데이터 미반영 방어적 재확인
    for blank in ["합계 (     )점", "평균 (     )점"]:
        assert blank in all_text, f"설문 응답 칸 원문이 예상과 달라짐(직접 재확인 필요): {blank!r}"

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
    print(f"PASS: F-050 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
