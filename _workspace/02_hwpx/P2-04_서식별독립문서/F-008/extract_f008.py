"""F-008(교복 사양서) 원본 보존형 독립 HWPX 추출.
**고난도 조판(이미지·병합표)** — 계획.md 8.2절 F-013 특별 처리 방침을 동일
적용한다: "이미지가 들어가는 디자인 표 셀은 값 채움 없이 완전 공란(테두리·
라벨만 유지)으로 템플릿을 만들고, '학교별로 디자인 이미지를 직접
삽입·저장'한다는 안내 각주를 템플릿에 포함함. 섬유혼용률·원단색상 등
텍스트 필드는 다른 서식과 동일하게 토큰 처리함."

원본 sec 직계 자식 인덱스 188~207(F-008 본문, idx187 "[6-1]" 캡션 제외,
idx208 "[7]"는 다음 서식 F-009 캡션)만 남기고 독립 문서로 분리한다.
표지·수신제목·결재란 없이 제목(idx188)+섬유혼용률 3x3 표(idx191, 구체
사양 예시 데이터라 원문 유지)+유의사항 문단+디자인 3x4 표(idx204, pic
이미지 5개 중 3개)+사진 2x2 표(idx206, pic 이미지 2개) 구조.

허용 반영 범위: 학교명(C-01, idx188 "○○○○학교", 원문자 4개 — 다른
서식들의 표준 2개 "○○학교"와 다름, 직접 확인)만 토큰화한다. idx191
섬유혼용률 표(모80%/나일론20% 등)는 구체 제품 사양 예시 데이터이므로
F-035/036의 업체 평가점수와 동일한 원칙(인스턴스 데이터는 원문 유지)으로
손대지 않는다.

**이미지 공란 처리**: idx204(3x4 표)·idx206(2x2 표) 안의 모든 <hp:pic>
요소(5개)를 제거한다(같은 run 안에 pic과 t(설명 텍스트)가 함께 있는 구조를
직접 XML로 확인 — pic만 제거하고 t는 그대로 두면 셀 테두리·설명 텍스트는
유지되고 이미지만 빠짐). 안내 각주는 새 문단을 삽입하지 않고(원문에 없는
문단을 새로 만드는 것은 더 큰 구조 변경 위험이 있어 회피), idx203("3. 교복
디자인 및 규격서")·idx205("4. 교복 사진") 제목 문단의 기존 텍스트 끝에
"(학교별로 디자인 이미지를 직접 삽입·저장)" 안내 문구만 추가한다(계획.md
F-013 방침이 명시적으로 승인한 예외 — 새 문단 생성 없이 최소 텍스트 추가로
구현).

사용법: python extract_f008.py <원본.hwpx> <출력_템플릿.hwpx>
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


def remove_pics_in(elem, expected_count):
    pics = [p for p in elem.iter() if etree.QName(p.tag).localname == "pic"]
    assert len(pics) == expected_count, f"pic 개수 불일치: 기대 {expected_count}, 실제 {len(pics)}"
    for pic in pics:
        run = pic.getparent()
        run.remove(pic)
        host = nearest_p(pic)
        if host is not None:
            remove_own_lineseg(host)
    return len(pics)


def append_note_to_paragraph(p, old_exact_text, note):
    """문단 p 안에서 text가 old_exact_text와 정확히 일치하는 <hp:t> 런을 찾아
    끝에 note를 덧붙인다(새 문단·새 런을 만들지 않고 기존 런 텍스트만 확장
    — 계획.md F-013 방침이 승인한 안내 각주 삽입을 최소 침습으로 구현)."""
    matches = [
        t for t in p.iter()
        if etree.QName(t.tag).localname == "t" and (t.text or "") == old_exact_text
    ]
    assert len(matches) == 1, f"'{old_exact_text}' 정확히 일치하는 런을 1개 찾지 못함(실제 {len(matches)}개)"
    matches[0].text = old_exact_text + note
    remove_own_lineseg(p)


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f008")
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

    if len(sec_children) < 209:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[187].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 6-1 ]" not in caption_text:
        raise AssertionError(f"idx187이 F-008 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[208].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 7 ]" not in next_caption_text:
        raise AssertionError(f"idx208이 다음(F-009) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(188, 208)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 188)]

    replace_text_anywhere(np(188), "○○○○학교", "{{학교명}}", expected_count=1)

    total_pics_removed = 0
    total_pics_removed += remove_pics_in(np(204), 3)
    total_pics_removed += remove_pics_in(np(206), 2)
    assert total_pics_removed == 5, f"제거된 pic 총합 불일치: {total_pics_removed}"

    append_note_to_paragraph(np(203), "3. 교복 디자인 및 규격서", "(학교별로 디자인 이미지를 직접 삽입·저장)")
    append_note_to_paragraph(np(205), "4. 교복 사진", "(학교별로 디자인 이미지를 직접 삽입·저장)")

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 1, f"{{학교명}} 개수 불일치: 기대 1, 실제 {actual}"
    assert "○○○○학교" not in all_text, "치환되지 않은 '○○○○학교' 잔존 문자열 발견"
    remaining_pics = [p for p in sec_elem.iter() if etree.QName(p.tag).localname == "pic"]
    assert len(remaining_pics) == 0, f"pic이 아직 남아있음: {len(remaining_pics)}개"
    for note in ["3. 교복 디자인 및 규격서(학교별로 디자인 이미지를 직접 삽입·저장)",
                 "4. 교복 사진(학교별로 디자인 이미지를 직접 삽입·저장)"]:
        assert note in all_text, f"안내 각주 삽입 확인 실패: {note!r}"

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
    print(f"PASS: F-008 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
