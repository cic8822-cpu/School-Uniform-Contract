"""P2-04 독립 문서 공통 골든 채움 스크립트.

extract_fXXX.py가 만든 템플릿(허용 토큰 "{{...}}"만 남은 사본)을 입력받아,
JSON으로 지정한 마스킹된 골든 테스트 값으로 텍스트만 치환한다. HWPX
재직렬화 없이 <hp:t> 텍스트 노드 문자열만 바꾸므로 원본 서식 구조(표·
이미지·페이지 설정)는 그대로 보존된다.

허용되지 않은 임의 토큰을 새로 만들지 않도록, JSON의 모든 키가 템플릿
안에 최소 1회 이상 실제로 존재하는지 검증하고, 채움 후에는 템플릿에
있던 모든 "{{...}}" 패턴이 남김없이 치환됐는지도 검증한다.

사용법:
  python fill_tokens_golden.py <템플릿.hwpx> <출력_골든.hwpx> <토큰JSON경로>

토큰JSON 예시:
  {"{{학교명}}": "검증초등학교", "{{학년도}}": "2026학년도"}
"""
import json
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

TOKEN_PATTERN = re.compile(r"\{\{[^{}]+\}\}")


def main():
    template_path, out_path, tokens_json_path = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(tokens_json_path, "r", encoding="utf-8") as f:
        tokens = json.load(f)

    work_dir = os.path.join(os.path.dirname(out_path), "_build_tmp_fill")
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)

    with zipfile.ZipFile(template_path) as zf:
        zf.extractall(work_dir)

    section_path = os.path.join(work_dir, "Contents", "section0.xml")
    tree = etree.parse(section_path)
    root = tree.getroot()

    before_text = "".join(t.text or "" for t in root.iter() if etree.QName(t.tag).localname == "t")
    for token in tokens:
        if token not in before_text:
            raise AssertionError(f"토큰 '{token}'이 템플릿에 존재하지 않음 - JSON을 재확인하세요")

    for t in root.iter():
        if etree.QName(t.tag).localname != "t" or not t.text:
            continue
        original = t.text
        for token, value in tokens.items():
            if token in t.text:
                t.text = t.text.replace(token, value)
        if t.text != original:
            p = t.getparent()
            while p is not None and etree.QName(p.tag).localname != "p":
                p = p.getparent()
            if p is not None:
                for c in list(p):
                    if etree.QName(c.tag).localname == "linesegarray":
                        p.remove(c)

    after_text = "".join(t.text or "" for t in root.iter() if etree.QName(t.tag).localname == "t")
    leftover = TOKEN_PATTERN.findall(after_text)
    if leftover:
        raise AssertionError(f"치환되지 않은 토큰 잔존: {leftover}")

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
    print(f"PASS: 골든 채움 완료 -> {out_path}")


if __name__ == "__main__":
    main()
