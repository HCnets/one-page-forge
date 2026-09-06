# -*- coding: utf-8 -*-
"""analyze_pdf.py — 课件PDF三合一分析

功能:
  1. 抽取每一页的文字层,输出带页号的全文 txt;
  2. 统计每页字数 + 嵌入图片情况(pages_report.csv);
  3. 找出"需要视觉转录的候选页",双信号判定,分三档:

    A 必转 · 零文字/近零文字页(字数 < --text-threshold):
      整页公式/整页截图/大字标题。PowerPoint 另存的 PDF 公式是图片对象,
      文字层里一个字都没有,不转 = 静默丢分。

    B 建议转 · 文字层正常但图是内容(经验规则,已在 677 页真实课件上校准,
      可召回人工逐页筛选出的内容图页的 94%):
        单张图片显示面积 ≥ 页面 25%  或 (页内图 ≥ 2 张且最大单图 ≥ 15%)
      PPT 的内容图多是"横幅型"(宽 > 高),单张占页 15%~35%,与装饰图区间
      重叠,没有纯面积阈值能精确分开 —— 见下"图特别多的课件怎么办"。

    C 抽查 · 其余含图页(单图占页 8%~25% 且不满足 B):多数是点缀图,
      若课件信息密度高,把 C 档页也渲染了让 AI 快速筛一遍更稳。

用法:
  python scripts/analyze_pdf.py 课件.pdf -o out_dir
  python scripts/analyze_pdf.py 课件.pdf -o out_dir --text-threshold 30 --img-area-ratio 0.08

输出(out_dir 下):
  全文.txt            逐页全文,每页标注 (字数/图数/最大图占比)
  pages_report.csv    页号, 文字层字数, 嵌入图片数, 最大图片面积占比
  视觉转录候选页.txt   三档清单 + 渲染命令参数

图特别多的课件怎么办(候选页 > 150 时):
  说明课件几乎页页带插图,全转录成本高。工作流改为两轮:
  ① 渲染全部候选页(PNG 成本低);
  ② 每 15~20 张图拖给 AI,先跑"快速筛选"提示词(见 prompts/02 附录),
     让 AI 只回答哪些页的图包含文字层没有的信息,再把命中页做正式转录。
  快速筛选一轮能砍掉大半工作量,且比人工翻 PDF 更不会漏。
"""
import argparse
import csv
import os
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("缺少依赖: pip install pymupdf")

MULTI_STRONG = 0.15  # 经验值: 页内≥2图且最大图≥15% → 内容图页(横幅插图常见形态)


def page_images(page):
    """返回该页嵌入图片的显示面积占比列表(过滤掉未实际显示的图)。"""
    areas = []
    try:
        for img in page.get_images(full=True):
            xref = img[0]
            for r in page.get_image_rects(xref):
                if r.width > 1 and r.height > 1:
                    areas.append(r.width * r.height / (page.rect.width * page.rect.height))
    except Exception:
        pass
    return areas


def span_str(pages):
    spans = []
    for p in pages:
        if spans and p == spans[-1][-1] + 1:
            spans[-1].append(p)
        else:
            spans.append([p])
    return ",".join(f"{s[0]}-{s[-1]}" if len(s) > 1 else str(s[0]) for s in spans)


def main():
    ap = argparse.ArgumentParser(description="课件 PDF 文字层提取 + 图页/公式页三档检测")
    ap.add_argument("pdf", help="课件 PDF 路径")
    ap.add_argument("-o", "--out", default="analysis_out", help="输出目录(默认 analysis_out)")
    ap.add_argument("-t", "--text-threshold", type=int, default=30,
                    help="文字层字数低于该值 → A档,默认 30")
    ap.add_argument("-r", "--img-area-ratio", type=float, default=0.08,
                    help="单图面积占比下限(低于它视为点缀,不进候选),默认 0.08")
    ap.add_argument("-s", "--img-strong-ratio", type=float, default=0.25,
                    help="单图面积占比 ≥ 该值 → B档,默认 0.25")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    doc = fitz.open(args.pdf)
    total_pages = len(doc)
    print(f"PDF 共 {total_pages} 页,开始逐页分析...")

    full_lines, rows = [], []
    class_a, class_b, class_c = [], [], []
    for i, page in enumerate(doc, start=1):
        text = page.get_text("text").strip()
        n = len(text)
        imgs = page_images(page)
        max_ratio = max(imgs) if imgs else 0.0
        rows.append((i, n, len(imgs), round(max_ratio, 4)))

        full_lines.append(f"===== 第 {i} 页 (字数 {n} / 图 {len(imgs)} 张, 最大占页 {max_ratio:.0%}) =====")
        if text:
            full_lines.append(text)

        if n < args.text_threshold:
            class_a.append(i)
        elif max_ratio >= args.img_area_ratio:  # 有值得看的图才进 B/C 候选
            if max_ratio >= args.img_strong_ratio or (len(imgs) >= 2 and max_ratio >= MULTI_STRONG):
                class_b.append(i)
            else:
                class_c.append(i)

    base = os.path.splitext(os.path.basename(args.pdf))[0]
    txt_path = os.path.join(args.out, base + "_全文.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n\n".join(full_lines))

    csv_path = os.path.join(args.out, "pages_report.csv")
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["页码", "文字层字数", "嵌入图片数", "最大图片面积占比"])
        w.writerows(rows)

    def render_cmd(pages, label):
        return f"# {label}: python scripts/render_pages.py 课件.pdf --pages {span_str(pages)} -o png_out"

    list_path = os.path.join(args.out, "视觉转录候选页.txt")
    with open(list_path, "w", encoding="utf-8") as f:
        f.write("判定参数: A=字数<%d | B=单图≥%.0f%%或(图≥2且单图≥%.0f%%) | C=其余含图页(单图≥%.0f%%)\n"
                % (args.text_threshold, args.img_strong_ratio * 100, MULTI_STRONG * 100, args.img_area_ratio * 100))
        f.write(f"共 {total_pages} 页: A档 {len(class_a)} / B档 {len(class_b)} / C档 {len(class_c)}\n\n")
        f.write("A档 · 零文字页(必转,不转必丢分):\n" + ", ".join(map(str, class_a)) + "\n\n")
        f.write("B档 · 文字层正常但图承载内容(建议转):\n" + ", ".join(map(str, class_b)) + "\n\n")
        f.write("C档 · 其余含图页(抽查;图特别多的课件按 README FAQ 两轮工作流处理):\n"
                + ", ".join(map(str, class_c)) + "\n\n")
        f.write("渲染命令:\n")
        f.write(render_cmd(sorted(set(class_a) | set(class_b)), "A+B 档") + "\n")
        f.write(render_cmd(class_c, "C 档(可选)") + "\n")

    print(f"完成: {txt_path}")
    print(f"页级报告: {csv_path}")
    print(f"候选页: A档 {len(class_a)}(必转) / B档 {len(class_b)}(建议转) / C档 {len(class_c)}(抽查)")
    if len(class_b) + len(class_c) > 150:
        print("⚠ 含图候选页偏多(>150): 课件可能每页带插图。渲染后先跑 prompts/02 附录的"
              "'快速筛选'提示词,让 AI 挑出真需要转录的页,再正式转录,别盲目全转。")
    elif len(class_a) > 0:
        print(f"⚠ A档 {len(class_a)} 页文字层为空: 全是公式/图片/大字页,不转录将静默丢分。")


if __name__ == "__main__":
    main()
