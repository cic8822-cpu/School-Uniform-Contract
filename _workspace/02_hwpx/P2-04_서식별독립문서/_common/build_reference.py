"""임의 Form ID의 원본 문단 범위를 토큰 치환·캐리어 스타일 변경 없이 그대로
독립 HWPX로 추출한다. page_guard.py로 최종 템플릿/골든 결과와 문단·표·텍스트
길이를 대조하기 위한 순수 원본 대조본을 만드는 범용 스크립트다(47종 모두
동일 로직이라 공용으로 유지함 — extract_f0XX.py의 토큰 치환 로직과 달리
폼별 커스터마이징이 없으므로 DRY 원칙에 따라 공유함).

캐리어(idx0)는 원본 그대로(로고 pic 포함, charPr/paraPr 변경 없음) 사용한다 —
이 대조본은 한컴에서 직접 열어보는 용도가 아니라 page_guard.py의 구조·텍스트
길이 계량 비교 전용이므로, 실제 렌더링 품질에는 영향을 주지 않는다.

사용법: python build_reference.py <원본.hwpx> <출력_원본대조.hwpx> <시작idx> <끝idx제외> [<시작2> <끝2> ...]
"""
import os
import shutil
import sys
import zipfile
import copy

from lxml import etree


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    range_args = [int(x) for x in sys.argv[3:]]
    if len(range_args) % 2 != 0 or not range_args:
        raise SystemExit("인덱스 범위는 <시작 끝> 쌍으로 짝수 개 전달해야 함")
    ranges = [(range_args[i], range_args[i + 1]) for i in range(0, len(range_args), 2)]

    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_ref")
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

    carrier = copy.deepcopy(sec_children[0])
    body_paragraphs = []
    for start, end in ranges:
        body_paragraphs.extend(copy.deepcopy(sec_children[i]) for i in range(start, end))

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(carrier)
    for p in body_paragraphs:
        sec_elem.append(p)

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
    print(f"PASS: 원본 대조본 생성 완료 -> {out_path} (범위: {ranges})")


if __name__ == "__main__":
    main()
