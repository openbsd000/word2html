#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
add_page_number.py —— 删掉当前文件夹里所有 Word 文档 (.docx / .doc) 的页眉、页脚，
并在页脚正中间加上页码（PAGE 域）。

用法：cd 到要处理的文件夹，运行
      python3 add_page_number.py

特点：
  * 只用标准库（docx 本身就是 zip + XML），不用装 python-docx
  * 首页页眉页脚、偶数页页眉页脚一律删掉，只留一个居中的页码页脚
  * 页码用 PAGE 域，Word 里打开会自动显示正确页码（纯文本不会写死数字）
  * 默认不备份，直接改原文件（页眉页脚删除后无法还原；要备份把 BACKUP 改成 True）
  * .doc 先用 LibreOffice 转成 docx，处理完再转回 .doc
  * 单个文件失败不影响其它文件

依赖：
  * 处理 .doc 需要 LibreOffice：sudo apt install libreoffice
"""

import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

# ------------------ 可配置项 ------------------
INPUT_EXTENSIONS = {".docx", ".doc"}   # 要处理的扩展名（小写）
BACKUP           = False               # 处理前是否备份成 原名.bak（True 时已有的备份不覆盖）
MIN_FOOTER_TWIPS = 567                 # 页脚距页边至少 1cm，太小页码会被裁掉
# ----------------------------------------------

CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
REL_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

FOOTER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" \
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>
"""

PART_RE = re.compile(r"^word/(?:header|footer)\d+\.xml$")
RELS_RE = re.compile(r"^word/_rels/(?:header|footer)\d+\.xml\.rels$")
REF_RE = re.compile(
    r"<w:(?:header|footer)Reference\b[^>]*/>"
    r"|<w:(?:header|footer)Reference\b.*?</w:(?:header|footer)Reference>", re.S)
CT_RE = re.compile(r'<Override PartName="/word/(?:header|footer)\d+\.xml"[^>]*/>')
REL_ITEM_RE = re.compile(
    r'<Relationship\b[^>]*Type="[^"]*/(?:header|footer)"[^>]*/>')
SECT_RE = re.compile(r"<w:sectPr\b[^>]*>.*?</w:sectPr>", re.S)
SECT_EMPTY_RE = re.compile(r"<w:sectPr\b[^>]*/>")
PGMAR_RE = re.compile(r"<w:pgMar\b[^>]*/>")


def find_soffice() -> str | None:
    """定位 LibreOffice；没有就返回 None（只影响 .doc）。"""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if exe:
        return exe
    for p in ("/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice"):
        if Path(p).is_file():
            return p
    return None


def ensure_footer_distance(doc: str) -> str:
    """页脚距页边太小的话页码会被裁掉，保证至少 MIN_FOOTER_TWIPS。"""
    def bump(m):
        tag = m.group(0)
        cur = re.search(r'\bw:footer="(\d+)"', tag)
        if cur and int(cur.group(1)) >= MIN_FOOTER_TWIPS:
            return tag
        if cur:
            return tag[:cur.start()] + f'w:footer="{MIN_FOOTER_TWIPS}"' + tag[cur.end():]
        return tag.replace("/>", f' w:footer="{MIN_FOOTER_TWIPS}"/>')
    return PGMAR_RE.sub(bump, doc)


def new_rid(rels_xml: str) -> str:
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return "rId" + str((max(used) + 1) if used else 1)


def clean_and_add_page_number(src: Path) -> bool:
    """在 docx 上原地操作：删掉所有页眉页脚，加一个居中页码页脚。"""
    with zipfile.ZipFile(src) as z:
        parts = {n: z.read(n) for n in z.namelist()}

    if "word/document.xml" not in parts:
        print(f"  [失败] {src.name}: 不是有效的 docx（没有 word/document.xml）")
        return False

    # 1) 丢掉所有页眉 / 页脚部件及其关系
    dropped = [n for n in parts if PART_RE.match(n) or RELS_RE.match(n)]
    for n in dropped:
        parts.pop(n)
    if "[Content_Types].xml" in parts:
        parts["[Content_Types].xml"] = CT_RE.sub(
            "", parts["[Content_Types].xml"].decode("utf-8")).encode("utf-8")

    # 2) 正文里去掉引用，并保证页脚留白够
    doc = parts["word/document.xml"].decode("utf-8")
    doc = REF_RE.sub("", doc)
    doc = re.sub(r"<w:titlePg\s*/>", "", doc)          # 首页不同也取消，保证每页都有页码
    doc = ensure_footer_distance(doc)

    # 3) 放一个新的页脚部件
    parts["word/footer1.xml"] = FOOTER_XML.encode("utf-8")
    if "[Content_Types].xml" in parts:
        ct = parts["[Content_Types].xml"].decode("utf-8")
        if "/word/footer1.xml" not in ct:
            ct = ct.replace("</Types>",
                            f'<Override PartName="/word/footer1.xml" '
                            f'ContentType="{CT_FOOTER}"/></Types>')
        parts["[Content_Types].xml"] = ct.encode("utf-8")

    rels_name = "word/_rels/document.xml.rels"
    rels = parts.get(rels_name, b"").decode("utf-8")
    rels = REL_ITEM_RE.sub("", rels)                   # 清掉指向已删除部件的关系
    rid = new_rid(rels)
    rels = rels.replace("</Relationships>",
                        f'<Relationship Id="{rid}" Type="{REL_FOOTER}" '
                        f'Target="footer1.xml"/></Relationships>')
    parts[rels_name] = rels.encode("utf-8")

    # 4) 每一节都挂上这个页脚（sectPr 可能是空标签 <w:sectPr />，两种都要处理）
    ref = f'<w:footerReference w:type="default" r:id="{rid}"/>'
    def add_ref(m):
        sect = m.group(0)
        return sect if "footerReference" in sect else \
            sect.replace("</w:sectPr>", ref + "</w:sectPr>")
    doc, n_sect = re.subn(r"<w:sectPr\b[^>]*/>", f"<w:sectPr>{ref}</w:sectPr>", doc)
    doc, n_full = SECT_RE.subn(add_ref, doc)
    parts["word/document.xml"] = doc.encode("utf-8")
    if n_sect + n_full == 0:
        print(f"  [警告] {src.name}: 没找到节属性(sectPr)，页码可能挂不上")

    tmp = src.with_suffix(src.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    tmp.replace(src)
    return True


def backup(src: Path) -> None:
    if not BACKUP:
        return
    dst = src.with_suffix(src.suffix + ".bak")
    if dst.exists():
        print(f"  [备份] 已有 {dst.name}，保留最早的备份")
        return
    shutil.copy2(src, dst)


def doc_to_docx(src: Path, soffice: str, outdir: Path) -> Path | None:
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, "--headless", "--norestore",
           "-env:UserInstallation=file://" + str(outdir),
           "--convert-to", "docx", "--outdir", str(outdir), str(src)]
    try:
        subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=300)
    except subprocess.TimeoutExpired:
        print(f"  [失败] {src.name}: LibreOffice 转 docx 超时")
        return None
    out = outdir / (src.stem + ".docx")
    return out if out.is_file() else None


def docx_to_doc(src: Path, soffice: str, outdir: Path) -> Path | None:
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, "--headless", "--norestore",
           "-env:UserInstallation=file://" + str(outdir),
           "--convert-to", "doc", "--outdir", str(outdir), str(src)]
    try:
        subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=300)
    except subprocess.TimeoutExpired:
        print(f"  [失败] {src.name}: LibreOffice 转回 doc 超时")
        return None
    out = outdir / (src.stem + ".doc")
    return out if out.is_file() else None


def process_one(src: Path, soffice: str | None) -> bool:
    """处理单个文件：.docx 直接改；.doc 转成 docx 改完再转回。"""
    backup(src)

    if src.suffix.lower() == ".docx":
        return clean_and_add_page_number(src)

    if not soffice:
        print(f"  [失败] {src.name}: .doc 需要 LibreOffice（sudo apt install libreoffice）")
        return False

    with tempfile.TemporaryDirectory(prefix="apn_") as td:
        work = Path(td)
        mid = doc_to_docx(src, soffice, work / "1")
        if not mid:
            print(f"  [失败] {src.name}: 没能转成 docx")
            return False
        if not clean_and_add_page_number(mid):
            return False
        back = docx_to_doc(mid, soffice, work / "2")
        if not back:
            print(f"  [失败] {src.name}: 没能转回 doc")
            return False
        shutil.copy2(back, src)
    return True


def main() -> None:
    work_dir = Path.cwd()
    print(f"[信息] 工作目录: {work_dir}")
    soffice = find_soffice()
    print(f"[信息] LibreOffice: {soffice or '未安装（.doc 无法处理）'}")

    files = sorted(p for p in work_dir.iterdir()
                   if p.is_file() and p.suffix.lower() in INPUT_EXTENSIONS
                   and not p.name.startswith("~$"))
    if not files:
        print(f"[提示] 当前目录没有 {', '.join(sorted(INPUT_EXTENSIONS))} 文件。")
        return
    print(f"[信息] 找到 {len(files)} 个文件\n")

    ok = fail = 0
    for src in files:
        try:
            if process_one(src, soffice):
                ok += 1
                print(f"  [完成] {src.name}（已删页眉页脚并加居中页码）")
            else:
                fail += 1
        except Exception as e:
            fail += 1
            print(f"  [失败] {src.name}: {e}")

    print(f"\n[结束] 成功 {ok} 个，失败 {fail} 个。")
    print("[提示] 页码是 PAGE 域，Word 里打开会自动显示；打印预览可看到效果。")


if __name__ == "__main__":
    main()
