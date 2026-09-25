"""F-012 템플릿에 비식별 테스트값을 채워 골든 결과를 생성한다.
사용법: python fill_f012_golden.py <템플릿.hwpx> <골든결과.hwpx>
"""
import os
import shutil
import sys
import zipfile

from lxml import etree

TOKENS = {
    "{{학교명}}": "검증초등학교",
    "{{학년도}}": "2026학년도",
}


def remove_own_lineseg(p):
    for c in list(p):
        if etree.QName(c.tag).localname == "linesegarray":
            p.remove(c)


def nearest_p(elem):
    cur = elem.getparent()
    while cur is not None and etree.QName(cur.tag).localname != "p":
        cur = cur.getparent()
    return cur


def main():
    src_path, out_path = sys.argv[1], sys.argv[2]
    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_golden_f012")
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)

    with zipfile.ZipFile(src_path) as zf:
        zf.extractall(work_dir)

    section_path = os.path.join(work_dir, "Contents", "section0.xml")
    tree = etree.parse(section_path)
    root = tree.getroot()

    targets = [t for t in root.iter() if etree.QName(t.tag).localname == "t" and t.text]
    replaced_total = {k: 0 for k in TOKENS}
    for t in targets:
        original = t.text
        changed = False
        for token, value in TOKENS.items():
            if token in t.text:
                t.text = t.text.replace(token, value)
                replaced_total[token] += original.count(token)
                changed = True
        if changed:
            host = nearest_p(t)
            if host is not None:
                remove_own_lineseg(host)

    for token, count in replaced_total.items():
        assert count > 0, f"{token} 치환 실패(문서에서 발견 못함)"

    remaining = "".join(
        t.text or "" for t in root.iter() if etree.QName(t.tag).localname == "t"
    )
    for token in TOKENS:
        assert token not in remaining, f"{token} 잔존"

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
    print(f"PASS: F-012 골든결과 생성 완료 -> {out_path}")


if __name__ == "__main__":
    main()
