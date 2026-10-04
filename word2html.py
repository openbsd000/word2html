#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
word2html.py —— 将当前文件夹内的所有 Word 文档 (.docx) 转换为单文件 HTML。

特点：
  * 保留原文档的标题、列表、表格、图片等格式
  * 图片以 base64 内嵌，输出为单个 HTML 文件（可离线打开、可发邮件）
  * 数学 / 科学公式使用 MathML，浏览器原生渲染，无需联网、无需 MathJax
  * 老式公式编辑器存的 WMF/EMF 公式图（浏览器显示不了），自动用 LibreOffice
    转成高清 PNG、裁掉白边再回写
  * 单个文件失败不影响其它文件

依赖：
  * pandoc >= 2.x
    安装：sudo apt install pandoc
  * LibreOffice + Pillow（可选）：仅当文档里的公式是 WMF/EMF 图片时需要
    安装：sudo apt install libreoffice python3-pil
"""

import base64
import hashlib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

# ------------------ 可配置项 ------------------
INPUT_EXTENSIONS = {".docx"}   # 要转换的扩展名（小写）
OUTPUT_SUFFIX    = ".html"     # 输出后缀
OVERWRITE        = True        # 已存在同名 HTML 时是否覆盖
ADD_TOC          = True        # 是否生成目录
TOC_DEPTH        = 3           # 目录深度
FIX_FORMULA      = True        # 把 WMF/EMF 公式图转成裁边高清 PNG（需 LibreOffice + Pillow）
FORMULA_DPI      = 200         # 公式导出精度：A4 页面按此 DPI 光栅化，显示时缩回 96dpi 物理尺寸
# ----------------------------------------------

EXTRA_CSS = """
html { -webkit-text-size-adjust: 100%; }
body {
    max-width: 920px;
    margin: 2em auto;
    padding: 0 1.2em 4em;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
                 "Noto Sans CJK SC", "Source Han Sans SC",
                 "PingFang SC", "Microsoft YaHei",
                 Roboto, Helvetica, Arial, sans-serif;
    font-size: 16px;
    line-height: 1.75;
    color: #222;
    word-wrap: break-word;
}
h1, h2, h3, h4, h5 { line-height: 1.3; margin-top: 1.5em; }
h1 { border-bottom: 2px solid #eee; padding-bottom: .3em; }

/* 表格 */
table { border-collapse: collapse; margin: 1em 0; }
th, td { border: 1px solid #ccc; padding: .4em .7em; vertical-align: top; }
th { background: #f5f5f5; }

/* 图片 */
img { max-width: 100%; height: auto; }

/* 行内 / 块级代码 */
code { background: #f5f5f5; padding: .1em .35em; border-radius: 3px; font-size: .95em; }
pre  { background: #f7f7f7; padding: .8em 1em; border-radius: 4px; overflow-x: auto; }
pre code { background: none; padding: 0; }

/* 引用 */
blockquote { border-left: 4px solid #ddd; margin: 1em 0;
             padding: .2em 1em; color: #555; }

/* 公式：MathML 略放大一点更清晰 */
math { font-size: 1.05em; }

/* 目录 */
nav#TOC { background: #fafafa; border: 1px solid #eee;
          border-radius: 6px; padding: .8em 1.2em; margin: 1.5em 0; }
nav#TOC ul { margin: .3em 0; padding-left: 1.4em; }
"""


def find_pandoc() -> str:
    """定位 pandoc 可执行文件。"""
    exe = shutil.which("pandoc")
    if exe:
        return exe
    for p in ("/usr/bin/pandoc", "/usr/local/bin/pandoc",
              "/opt/homebrew/bin/pandoc"):
        if Path(p).is_file():
            return p
    print("[错误] 没有找到 pandoc，无法转换。")
    print("       Ubuntu/Debian:  sudo apt install pandoc")
    print("       macOS:           brew install pandoc")
    print("       Windows:         winget install --id JohnMacFarlane.Pandoc")
    print("       官网：https://pandoc.org/installing.html")
    sys.exit(1)


def find_soffice() -> str | None:
    """定位 LibreOffice；没有就返回 None（只影响 WMF/EMF 公式转换）。"""
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if exe:
        return exe
    for p in ("/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice"):
        if Path(p).is_file():
            return p
    return None


# 公式图：pandoc 对老式公式（MathType/公式编辑器 3.0）只会原样输出 WMF/EMF 图片，
# 浏览器一律显示不出来，必须转格式
FORMULA_RE = re.compile(r'src="(data:image/x-(wmf|emf);base64,([A-Za-z0-9+/=]+))"')
PAGE_W_PX = round(8.27 * FORMULA_DPI)   # A4 竖版按 FORMULA_DPI 光栅化的像素宽
PAGE_H_PX = round(11.69 * FORMULA_DPI)


def docx_image_sizes(src: Path) -> dict:
    """读出每张图片在 Word 里的显示尺寸（pt），按图片字节的 sha1 建索引。

    公式是 OLE 对象，尺寸藏在 VML 的 v:shape style="height:..pt;width:..pt;" 里；
    普通图片在 DrawingML 的 wp:extent（EMU）里。两者都取，之后用 HTML 里
    base64 的内容算 sha1 就能对上号。
    """
    sizes = {}
    with zipfile.ZipFile(src) as z:
        names = set(z.namelist())
        if "word/document.xml" not in names:
            return sizes
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
        rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"',
                               z.read("word/_rels/document.xml.rels").decode("utf-8", "ignore")))

        def full(t):
            return t if t in names else "word/" + t

        def put(name, w_pt, h_pt):
            name = full(name)
            if name in names and w_pt > 0 and h_pt > 0:
                sizes.setdefault(hashlib.sha1(z.read(name)).hexdigest(), (w_pt, h_pt))

        for blk in re.split(r'(?=<v:shape)', xml):            # VML：OLE 公式图
            if not blk.startswith("<v:shape"):
                continue
            style = re.search(r'style="([^"]*)"', blk)
            rid = re.search(r'<v:imagedata[^>]*r:id="(rId\d+)"', blk)
            if not (style and rid and rid.group(1) in rels):
                continue
            w = re.search(r'width:([\d.]+)pt', style.group(1))
            h = re.search(r'height:([\d.]+)pt', style.group(1))
            if w and h:
                put(rels[rid.group(1)], float(w.group(1)), float(h.group(1)))

        for blk in re.findall(r'<wp:(?:inline|anchor)\b.*?</wp:(?:inline|anchor)>', xml, re.S):
            ext = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', blk)   # DrawingML：普通图片
            rid = re.search(r'<a:blip[^>]*r:embed="(rId\d+)"', blk)
            if ext and rid and rid.group(1) in rels:
                put(rels[rid.group(1)], int(ext.group(1)) / 12700, int(ext.group(2)) / 12700)
    return sizes


def fix_formula_images(dst: Path, soffice: str, sizes: dict) -> None:
    """把 HTML 里的 WMF/EMF 公式图批量转成裁掉白边的高清 PNG 并写回。

    LibreOffice 导出的是整张 A4（公式缩在页面中间），所以按 FORMULA_DPI 放大导出、
    Pillow 自动裁白边；显示尺寸优先用 Word 里声明的大小（sizes），墨迹比例和声明
    比例不一致时用 object-fit:contain 缩放而不拉伸，保证和 Word 里一样大。
    """
    html = dst.read_text(encoding="utf-8")
    jobs = {}                                # base64 内容 -> 扩展名（内容去重）
    for m in FORMULA_RE.finditer(html):
        jobs.setdefault(m.group(3), m.group(2))
    if not jobs:
        return
    if not soffice:
        print("  [警告] 文档里有 WMF/EMF 公式图，但没找到 LibreOffice，保持原样"
              "（sudo apt install libreoffice）")
        return

    try:
        from PIL import Image, ImageChops
    except ImportError:
        print("  [警告] 文档里有 WMF/EMF 公式图，但未安装 Pillow，保持原样"
              "（sudo apt install python3-pil）")
        return

    with tempfile.TemporaryDirectory(prefix="w2h_") as td:
        indir, outdir = Path(td) / "in", Path(td) / "out"
        indir.mkdir()
        outdir.mkdir()
        srcs = []
        raws = {}                            # base64 内容 -> 原始字节
        for i, (blob, ext) in enumerate(sorted(jobs.items())):
            raw = base64.b64decode(blob)
            raws[blob] = raw
            p = indir / f"{i:04d}.{ext}"
            p.write_bytes(raw)
            srcs.append(p)

        filt = ('png:draw_png_Export:{"PixelWidth":{"type":"long","value":%d},'
                '"PixelHeight":{"type":"long","value":%d}}' % (PAGE_W_PX, PAGE_H_PX))
        cmd = [soffice, "--headless", "--norestore",
               "-env:UserInstallation=file://" + td,   # 独立配置，避免和开着的 LibreOffice 冲突
               "--convert-to", filt, "--outdir", str(outdir)] + [str(p) for p in srcs]
        try:
            subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=600)
        except subprocess.TimeoutExpired:
            print("  [警告] LibreOffice 批量转公式超时，部分公式保持原样")

        mapping = {}                         # base64 内容 -> (新 data uri, 宽px, 高px|None)
        from_docx = 0
        for i, blob in enumerate(sorted(jobs)):
            out = outdir / f"{i:04d}.png"
            if not out.is_file():
                continue
            im = Image.open(out).convert("RGB")
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bbox = ImageChops.difference(im, bg).getbbox()
            if not bbox:
                continue
            pad = 4
            box = (max(bbox[0] - pad, 0), max(bbox[1] - pad, 0),
                   min(bbox[2] + pad, im.width), min(bbox[3] + pad, im.height))
            im.crop(box).save(out, optimize=True)
            w_px = h_px = None
            declared = sizes.get(hashlib.sha1(raws[blob]).hexdigest())
            if declared:
                w_px = max(1, round(declared[0] * 96 / 72))     # pt → CSS px
                h_px = max(1, round(declared[1] * 96 / 72))
                from_docx += 1
            else:                                              # 拿不到声明尺寸时按导出精度折算
                w_px = max(1, round((box[2] - box[0]) * 96 / FORMULA_DPI))
            mapping[blob] = ("data:image/png;base64,"
                             + base64.b64encode(out.read_bytes()).decode(), w_px, h_px)

    if not mapping:
        print("  [警告] WMF/EMF 公式图转换失败，保持原样")
        return

    def repl(m):
        hit = mapping.get(m.group(3))
        if not hit:
            return m.group(0)
        uri, w_px, h_px = hit
        if h_px:
            # 声明框和墨迹比例常常不一样（Word 会留白），contain 保证只缩放不拉伸
            return f'src="{uri}" width="{w_px}" height="{h_px}" style="object-fit:contain"'
        return f'src="{uri}" width="{w_px}"'

    html = FORMULA_RE.sub(repl, html)
    dst.write_text(html, encoding="utf-8")
    print(f"  [公式] {len(mapping)}/{len(jobs)} 个 WMF/EMF 公式图已转为裁边高清 PNG"
          f"（{from_docx} 个按 Word 原尺寸显示）")


def convert_one(pandoc: str, src: Path, header_file: Path,
                soffice: str | None, sizes: dict) -> bool:
    """把单个 docx 转成单文件 html。成功返回 True。"""
    dst = src.with_suffix(OUTPUT_SUFFIX)

    if dst.exists() and not OVERWRITE:
        print(f"  [跳过] {dst.name} 已存在")
        return True

    cmd = [
        pandoc,
        str(src),
        "-o", str(dst),
        "--standalone",               # 生成完整 HTML（含 head/body）
        "--embed-resources",          # 图片/CSS 全部内嵌 → 真正的单文件
        "--mathml",                   # 公式用 MathML，浏览器原生支持
        "--metadata", f"title={src.stem}",
        "--resource-path", str(src.parent),
        "--include-in-header", str(header_file),
    ]
    if ADD_TOC:
        cmd += ["--toc", f"--toc-depth={TOC_DEPTH}"]

    try:
        r = subprocess.run(cmd, capture_output=True, text=True, check=False)
    except Exception as e:
        print(f"  [失败] {src.name}: {e}")
        return False

    if r.returncode != 0 or not dst.exists():
        print(f"  [失败] {src.name}")
        if r.stderr.strip():
            print("         " + r.stderr.strip().replace("\n", "\n         "))
        return False

    if FIX_FORMULA:
        fix_formula_images(dst, soffice, sizes)

    size_kb = dst.stat().st_size / 1024
    print(f"  [完成] {src.name}  →  {dst.name}  ({size_kb:.1f} KB)")
    return True


def main():
    pandoc = find_pandoc()
    print(f"[信息] 使用 pandoc: {pandoc}")
    soffice = find_soffice()
    if FIX_FORMULA and soffice:
        print(f"[信息] 公式图转换使用 LibreOffice: {soffice}")

    work_dir = Path.cwd()
    print(f"[信息] 工作目录: {work_dir}")

    files = sorted(
        p for p in work_dir.iterdir()
        if p.is_file() and p.suffix.lower() in INPUT_EXTENSIONS
    )
    if not files:
        print(f"[提示] 当前目录没有找到 {', '.join(INPUT_EXTENSIONS)} 文件。")
        return
    print(f"[信息] 找到 {len(files)} 个待转换文件\n")

    # 写一个临时 header 文件（里面是 <style>），供 pandoc 注入到 <head>
    with tempfile.NamedTemporaryFile(
        "w", suffix=".html", delete=False, encoding="utf-8"
    ) as fh:
        fh.write("<style>\n" + EXTRA_CSS + "\n</style>\n")
        header_file = Path(fh.name)

    ok = fail = 0
    try:
        for src in files:
            sizes = docx_image_sizes(src) if FIX_FORMULA else {}
            if convert_one(pandoc, src, header_file, soffice, sizes):
                ok += 1
            else:
                fail += 1
    finally:
        header_file.unlink(missing_ok=True)

    print(f"\n[结束] 成功 {ok} 个，失败 {fail} 个。")


if __name__ == "__main__":
    main()

