# -*- coding: utf-8 -*-
"""analyze_pdf.py — 课件PDF三合一分析

功能:
  1. 抽取每一页的文字层,输出带页号的全文 txt;
  2. 统计每页字数(pages_report.csv);
  3. 找出"低文字密度页"(图页/公式页/大字标题页的典型特征),
     这类页的文字层常常是空的——公式是图片对象,文字层里根本没有——
     必须走视觉转录(渲染成 PNG 喂给多模态 AI),否则知识点会静默丢失。

已知大坑(实战踩过):
  - PowerPoint 另存的 PDF,公式/图表是"图片对象",文字层提取不到任何内容;
    某 677 页课件有 40+ 页公式因此完全空白。
  - 有的 PDF 文字层存在但顺序乱(文本框错位),以页为单位检查时留意。

用法:
  python scripts/analyze_pdf.py 课件.pdf -o out_dir
  或
  python scripts/analyze_pdf.py 课件.pdf -o out_dir --density-threshold 40

输出(out_dir 下):
  全文.txt            逐页全文,每页以 "===== 第 N 页 =====" 分隔,并标注字数
  pages_report.csv    页号,字数
  低密度页清单.txt     需要视觉转录的页码清单(供 render_pages.py 使用)
"""
import argparse
import csv
import os
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("缺少依赖: pip install pymupdf")


def main():
    ap = argparse.ArgumentParser(description="课件 PDF 文字层提取 + 页级密度分析")
    ap.add_argument("pdf", help="课件 PDF 路径")
    ap.add_argument("-o", "--out", default="analysis_out", help="输出目录(默认 analysis_out)")
    ap.add_argument("-t", "--density-threshold", type=int, default=30,
                    help="字数低于该值的页判定为低密度页,默认 30")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    doc = fitz.open(args.pdf)
    total_pages = len(doc)
    print(f"PDF 共 {total_pages} 页,开始逐页提取...")

    full_lines, low_pages, rows = [], [], []
    for i, page in enumerate(doc, start=1):
        text = page.get_text("text")
        n = len(text.strip())
        rows.append((i, n))
        full_lines.append(f"===== 第 {i} 页 (字数 {n}) =====")
        if text.strip():
            full_lines.append(text.rstrip())
        if n < args.density_threshold:
            low_pages.append(i)

    base = os.path.splitext(os.path.basename(args.pdf))[0]
    txt_path = os.path.join(args.out, base + "_全文.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n\n".join(full_lines))

    csv_path = os.path.join(args.out, "pages_report.csv")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["页码", "文字层字数"])
        w.writerows(rows)

    # 低密度页按连续区间合并,便于批量渲染
    spans = []
    for p in low_pages:
        if spans and p == spans[-1][-1] + 1:
            spans[-1].append(p)
        else:
            spans.append([p])
    span_str = ",".join(f"{s[0]}-{s[-1]}" if len(s) > 1 else str(s[0]) for s in spans)

    low_path = os.path.join(args.out, "低密度页清单.txt")
    with open(low_path, "w", encoding="utf-8") as f:
        f.write(f"低密度阈值: 每页 < {args.density_threshold} 字\n")
        f.write(f"低密度页 {len(low_pages)} 个,共 {sum(len(s) for s in spans)} 页\n")
        f.write("低密度页编号: " + ", ".join(map(str, low_pages)) + "\n")
        f.write("\n渲染参数(直接粘给 render_pages.py):\n")
        f.write("--pages " + span_str + "\n")

    blank = [p for p, n in rows if n == 0]
    print(f"完成: {txt_path}")
    print(f"页数统计: {csv_path}")
    print(f"低密度页(疑似图/公式,需视觉转录): {len(low_pages)} 个 -> {low_path}")
    if blank:
        print(f"⚠ 完全空白页 {len(blank)} 个: {blank}")
    print(f"⚠ 低密度页只是'候选':文字层正常但排版稀疏的页也可能有考点图,"
          f"人工抽翻 PDF 复核一遍更稳(见 README FAQ)。")


if __name__ == "__main__":
    main()
