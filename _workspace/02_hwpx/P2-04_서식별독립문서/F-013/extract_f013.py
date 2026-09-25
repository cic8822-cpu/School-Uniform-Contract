"""F-013(교복 디자인 및 규격서, 표준 예시) 원본 보존형 독립 HWPX 추출.
**설계상 위임 대상(이미지 학교별 전면 교체) — 계획.md 8.2절 F-013 특별
처리 방침을 그대로 적용**: "이미지가 들어가는 디자인 표 셀은 값 채움 없이
완전 공란(테두리·라벨만 유지)으로 템플릿을 만들고, '학교별로 디자인
이미지를 직접 삽입·저장'한다는 안내 각주를 템플릿에 포함함. 섬유혼용률·
원단색상 등 텍스트 필드는 다른 서식과 동일하게 토큰 처리함."

원본 sec 직계 자식 인덱스 523~578(F-013 본문, idx522 "[9-2]" 캡션 제외,
idx579 "[9-3]"는 다음 서식(F-013 범위 밖) 캡션)만 남기고 독립 문서로
분리한다. F-008(교복 사양서)과 동일 계열이지만 훨씬 큰 규모 — 디자인
이미지 표 9개(가~자 품목별)에 이미지(pic) 총 55개(idx561:7, idx563:6,
idx565:6, idx567:6, idx569:6, idx571:5, idx574:3, idx576:9, idx578:7).

원본에서 직접 확인한 placeholder: idx523(제목) "○○○○학교"(원문자 4개,
F-008과 동일 표기)·idx559("3. OOOO학교 교복 디자인 및 규격서") "OOOO학교"
(**영문 대문자 O 4개** — F-008의 원문자·F-048의 "00학교"(숫자)와 또 다른
표기, 직접 확인). idx527 섬유혼용률 표(20x4, 구체 사양 예시 데이터)는
F-008과 동일 원칙으로 원문 유지.

**이미지 공란 처리**: idx561/563/565/567/569/571/574/576/578의 모든 pic
(총 55개)을 제거한다(같은 run 안에 pic들과 t(설명 텍스트)가 함께 있는
구조, pic만 제거). 안내 각주는 새 문단을 만들지 않고 idx559(전체 이미지
섹션 대제목 "3. OOOO학교 교복 디자인 및 규격서") 한 곳에만 추가한다
(하위 가~자 9개 소제목마다 반복하면 과도하므로 F-008과 달리 대제목
1곳에 통합 — 계획.md 방침이 명시한 "안내 각주 포함"의 취지는 유지하되
최소 침습 설계).

사용법: python extract_f013.py <원본.hwpx> <출력_템플릿.hwpx>
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
    matches = [
        t for t in p.iter()
        if etree.QName(t.tag).localname == "t" and (t.text or "") == old_exact_text
    ]
    assert len(matches) == 1, f"'{old_exact_text}' 정확히 일치하는 런을 1개 찾지 못함(실제 {len(matches)}개)"
    matches[0].text = old_exact_text + note
    remove_own_lineseg(p)


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_f013")
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

    if len(sec_children) < 580:
        raise AssertionError(f"예상보다 문단 수가 적음: {len(sec_children)}")

    caption_text = "".join(
        t.text or "" for t in sec_children[522].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-2 ]" not in caption_text:
        raise AssertionError(f"idx522가 F-013 캡션이 아님: {caption_text!r}")
    next_caption_text = "".join(
        t.text or "" for t in sec_children[579].iter() if etree.QName(t.tag).localname == "t"
    )
    if "[ 9-3 ]" not in next_caption_text:
        raise AssertionError(f"idx579가 다음(F-013 범위 밖) 캡션이 아님: {next_caption_text!r}")

    carrier = copy.deepcopy(sec_children[0])
    carrier_run = next(c for c in carrier if etree.QName(c.tag).localname == "run")
    for pic in [c for c in carrier_run if etree.QName(c.tag).localname == "pic"]:
        carrier_run.remove(pic)
    secpr_present = any(etree.QName(c.tag).localname == "secPr" for c in carrier_run)
    ctrl_present = any(etree.QName(c.tag).localname == "ctrl" for c in carrier_run)
    assert secpr_present and ctrl_present, "secPr/ctrl 유실 - 페이지 설정 파괴 위험"
    carrier_run.set("charPrIDRef", CARRIER_CHAR_PR_ID)
    carrier.set("paraPrIDRef", CARRIER_PARA_PR_ID)

    body_paragraphs = [copy.deepcopy(sec_children[i]) for i in range(523, 579)]

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

    new_children = list(sec_elem)

    def np(orig_idx):
        return new_children[1 + (orig_idx - 523)]

    replace_text_anywhere(np(523), "○○○○학교", "{{학교명}}", expected_count=1)
    replace_text_anywhere(np(559), "OOOO학교", "{{학교명}}", expected_count=1)
    # 위에서 이미 "OOOO학교"→"{{학교명}}"으로 치환됐으므로, 각주는 치환 후 텍스트를 대상으로 추가한다.
    append_note_to_paragraph(np(559), "3. {{학교명}} 교복 디자인 및 규격서", "(학교별로 디자인 이미지를 직접 삽입·저장)")

    pic_counts = {561: 7, 563: 6, 565: 6, 567: 6, 569: 6, 571: 5, 574: 3, 576: 9, 578: 7}
    total_removed = 0
    for idx, count in pic_counts.items():
        total_removed += remove_pics_in(np(idx), count)
    assert total_removed == 55, f"제거된 pic 총합 불일치: {total_removed}"

    all_text = "".join(t.text or "" for t in sec_elem.iter() if etree.QName(t.tag).localname == "t")
    actual = all_text.count("{{학교명}}")
    assert actual == 2, f"{{학교명}} 개수 불일치: 기대 2, 실제 {actual}"
    for stray in ["○○○○학교", "OOOO학교"]:
        assert stray not in all_text, f"치환되지 않은 '{stray}' 잔존 문자열 발견"
    remaining_pics = [p for p in sec_elem.iter() if etree.QName(p.tag).localname == "pic"]
    assert len(remaining_pics) == 0, f"pic이 아직 남아있음: {len(remaining_pics)}개"
    assert "3. {{학교명}} 교복 디자인 및 규격서(학교별로 디자인 이미지를 직접 삽입·저장)" in all_text, "안내 각주 삽입 확인 실패"

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
    print(f"PASS: F-013 독립 템플릿 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
