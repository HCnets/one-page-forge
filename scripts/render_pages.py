# -*- coding: utf-8 -*-
"""render_pages.py — 把指定页渲染成 PNG,供多模态 AI 视觉转录

用法:
  # 渲染手动指定的页
  python scripts/render_pages.py 课件.pdf --pages 1-5,8,12-15 -o png_out

  # 直接渲染 analyze_pdf.py 判定的低密度页(推荐)
  python scripts/analyze_pdf.py 课件.pdf -o analysis_out
  python scripts/render_pages.py 课件.pdf --pages-file analysis_out/低密度页清单.txt -o png_out

输出: png_out/p001.png, p002.png ...(200 DPI,长页按高度上限整页缩放)
之后把 PNG 批量拖给你自己的多模态 AI,提示词用 prompts/02-图页转录.md。
"""
import argparse
import os
import re
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("缺少依赖: pip install pymupdf")


def parse_pages(spec: str):
    """'1-5,8,12-15' -> [1,2,3,4,5,8,12,13,14,15]"""
    pages = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            pages.extend(range(int(a), int(b) + 1))
        else:
            pages.append(int(part))
    return pages


def main():
    ap = argparse.ArgumentParser(description="把 PDF 指定页渲染成 PNG(供视觉转录)")
    ap.add_argument("pdf")
    ap.add_argument("-p", "--pages", help="页码,如 '1-5,8,12-15'")
    ap.add_argument("-f", "--pages-file", help="从 analyze_pdf 输出的清单文件读取页码范围")
    ap.add_argument("-o", "--out", default="png_out")
    ap.add_argument("--dpi", type=int, default=200, help="渲染分辨率,默认 200")
    ap.add_argument("--max-height", type=int, default=2800,
                    help="PNG 最大高度像素(超长页会整体缩放,防喂不进视觉模型),默认 2800")
    args = ap.parse_args()
    if not args.pages and not args.pages_file:
        ap.error("必须给 --pages 或 --pages-file")

    pages = []
    if args.pages_file:
        for line in open(args.pages_file, encoding="utf-8"):
            m = re.search(r"--pages\s+(\S+)", line)
            if m:
                pages = parse_pages(m.group(1))
                break
        if not pages:
            sys.exit(f"未能从 {args.pages_file} 解析页码(找 '--pages' 行)")
    else:
        pages = parse_pages(args.pages)

    os.makedirs(args.out, exist_ok=True)
    doc = fitz.open(args.pdf)
    zoom = args.dpi / 72.0
    for p in pages:
        if p < 1 or p > len(doc):
            print(f"跳过越界页 {p}")
            continue
        page = doc[p - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        if pix.height > args.max_height:  # 超长页整体缩放
            s = args.max_height / pix.height
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom * s, zoom * s))
        out = os.path.join(args.out, f"p{p:03d}.png")
        pix.save(out)
        print(f"已渲染 第{p}页 -> {out} ({pix.width}x{pix.height})")

    print(f"\n共 {len(pages)} 页。下一步: 把 {args.out} 里的 PNG 拖给多模态 AI,"
          f"提示词见 prompts/02-图页转录.md。")


if __name__ == "__main__":
    main()
