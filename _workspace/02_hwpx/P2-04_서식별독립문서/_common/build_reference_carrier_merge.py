"""캐리어 병합(carrier-merge) 기법을 적용한 독립 문서용 page_guard 원본대조본
빌더.

build_reference.py(별도 캐리어 문단 방식 전용)와 달리, 본문 단일 문단의
run에 secPr/ctrl을 직접 삽입하는 캐리어 병합 구조로 원본대조본을
만든다. 토큰 치환은 전혀 하지 않고 원문 그대로 유지한다(page_guard
비교 기준선).

사용법:
  python build_reference_carrier_merge.py <원본.hwpx> <출력_원본대조.hwpx> <본문idx>
"""
import copy
import os
import shutil
import sys
import zipfile

from lxml import etree


def main():
    src_path, out_path, body_idx = sys.argv[1], sys.argv[2], int(sys.argv[3])
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

    carrier_src = sec_children[0]
    carrier_src_run = next(c for c in carrier_src if etree.QName(c.tag).localname == "run")
    secpr = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "secPr"))
    ctrl = copy.deepcopy(next(c for c in carrier_src_run if etree.QName(c.tag).localname == "ctrl"))

    body = copy.deepcopy(sec_children[body_idx])
    body_run = next(c for c in body if etree.QName(c.tag).localname == "run")
    body_run.insert(0, ctrl)
    body_run.insert(0, secpr)

    for child in list(sec_elem):
        sec_elem.remove(child)
    sec_elem.append(body)

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
    print(f"PASS: 원본대조본 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
