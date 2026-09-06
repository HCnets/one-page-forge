# -*- coding: utf-8 -*-
"""make_cheat.py — 极限排版器:把"内容块 txt"排版成 A4 双面小抄 docx

实战规格(近代史 / 具身大模型理论两届验证):A4、边距 0.5cm、4.5pt 等线 Light、
行距 EXACT 4.7pt、段前后 0 → 单页约 172 文本行,双面共约 3.4 万个文本行位。

内容块轻量标记(与 prompts/03 的产出规则配套,AI 直接按此生成):
  【标题】          段首标签 → 加粗(通常作块名,如【3·模型结构】)
  〈关键名词〉      关键名词 → 红色加粗,全文同一词最多标红 red_limit 次
  **重点短语**      考点短语 → 加粗(注意: 是成对英文星号)
  名词: 值          冒号前字段名 → 自动加粗(如 "准确率: 88.7%")
  数字/单位数字     自动蓝色(0-9、亿/万/GB/B/token/%…)
  其余纯文本        黑色普通

用 --mono 可关闭红蓝,只留黑字+加粗(黑白打印时颜色没用)。

用法:
  python scripts/make_cheat.py --blocks-dir blocks --out a4.docx
  python scripts/make_cheat.py "内容块_*.txt" -o a4.docx --font-size 5 --mono
  python scripts/make_cheat.py -o a4.docx --style config/style.json   # 块文件默认取 块*.txt
生成后务必用 verify_pages.py 确认页数(Word 实际折行与估算有出入)。
"""
import argparse
import glob
import json
import math
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

DEFAULT_STYLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "config", "style.json")

# 数字/单位正则 → 蓝色(与实战版一致)
NUM_PAT = re.compile(
    r"(?:19|20)\d{2}(?:\.\d+)?(?:年|月|日)?"          # 年份/日期
    r"|\d+(?:\.\d+)?\s*(?:×10\^\d+|万亿|亿|万|千|百|十|GB|TB|MB|KB|B|token|Token|词元|词|条|个|种|次|轮|天|小时|分钟|月|年|分|秒|层|头|维|位|页|%)"
    r"|\d+(?:\.\d+)?\s*[a-zA-Z]+"                     # 3.0e-4、45TB 兜底(前者已被上一条覆盖)
)

# 冒号前字段名加粗(2~14 字、不以引导词开头、字段内无标点括号)
COLON_PAT = re.compile(r"([^。；;，,、\n（()）【】〈〉*]{1,14}?[:：])")
COLON_SKIP = ("例如", "比如", "如", "注", "注意", "说明", "其中", "即", "指", "表示", "认为",
              "包括", "分为", "如下", "综上", "总结", "详见", "换句话说", "也就是说", "但", "且",
              "并", "而", "或", "若", "如果", "因为", "由于", "所以", "因此", "此外", "另外",
              "同时", "第一", "第二", "首先", "其次", "最后")


def build_pat():
    """交替正则:【标签】| 〈红词〉 | **粗词** | 蓝色数字"""
    return re.compile(
        r"(【[^】]{1,60}】)"
        r"|(〈[^〉]{1,40}〉)"
        r"|(\*\*[^*]{1,60}\*\*)"
        r"|(" + NUM_PAT.pattern + r")"
    )


class Styler:
    def __init__(self, cfg):
        self.font_name = cfg["font"]["name"]
        self.size = cfg["font"]["size_pt"]
        self.line_pt = cfg["font"]["line_spacing_pt"]
        self.red = cfg["color"]["red"]
        self.blue = cfg["color"]["blue"]
        self.red_limit = cfg["color"]["red_limit"]
        self.bold_limit = cfg["color"]["bold_limit"]
        self.red_on = cfg["color"]["enabled"]
        self.counts = {}   # 红词出现计数
        self.bcounts = {}  # 加粗短语出现计数
        self.pat = build_pat()

    def set_font(self, run):
        run.font.size = Pt(self.size)
        run.font.name = self.font_name
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        for attr in ("w:ascii", "w:eastAsia", "w:hAnsi"):
            rFonts.set(qn(attr), self.font_name)

    def add_run(self, par, tok, bold=False, color=None):
        r = par.add_run(tok)
        if bold:
            r.font.bold = True
        if color and self.red_on:
            r.font.color.rgb = RGBColor.from_string(color)
        self.set_font(r)

    # ---- 块级标记 ----
    def add_rich(self, par, text):
        pos = 0
        for m in self.pat.finditer(text):
            s, e = m.start(), m.end()
            if s > pos:
                self.add_plain(par, text[pos:s])          # 普通段(内部切冒号字段)
            tok = text[s:e]
            if m.group(1):        # 【标签】→ 粗
                self.add_run(par, tok, bold=True)
            elif m.group(2):      # 〈红词〉→ 红(限次)
                word = tok[1:-1]
                n = self.counts.get(word, 0)
                if n < self.red_limit:
                    self.counts[word] = n + 1
                    self.add_run(par, word, bold=True, color=self.red)
                else:
                    self.add_run(par, word)               # 超限: 普通黑字
            elif m.group(3):      # **短语** → 粗(限次)
                phrase = tok[2:-2]
                n = self.bcounts.get(phrase, 0)
                if n < self.bold_limit:
                    self.bcounts[phrase] = n + 1
                    self.add_run(par, phrase, bold=True)
                else:
                    self.add_run(par, phrase)
            else:                 # 蓝色数字
                self.add_run(par, tok, color=self.blue)
            pos = e
        if pos < len(text):
            self.add_plain(par, text[pos:])

    # ---- 冒号前字段名 ----
    def add_plain(self, par, seg):
        pos = 0
        for m in COLON_PAT.finditer(seg):
            s, e = m.start(), m.end()
            field = m.group(1)
            name = field[:-1].strip()
            if (2 <= len(name) <= 14 and ":" not in name and "：" not in name
                    and not any(name.startswith(k) for k in COLON_SKIP)
                    and name[-1] not in "。，、；：()（）"):
                if s > pos:
                    self.add_run(par, seg[pos:s])
                self.add_run(par, field, bold=True)
                pos = e
        if pos < len(seg):
            self.add_run(par, seg[pos:])


def estimate_lines(texts, chars_per_line):
    """粗估总文本行数(仅供预警;最终以 verify_pages.py 为准)。"""
    total = sum(len(t) for t in texts)
    lines = math.ceil(total / chars_per_line)
    # 每个段落至少占一行(块文件=段落)
    lines = max(lines, len(texts))
    return lines, total


def main():
    ap = argparse.ArgumentParser(description="内容块 → A4 双面极限小抄 docx")
    ap.add_argument("glob", nargs="?", default=None, help="内容块文件通配符,默认 '块*.txt'")
    ap.add_argument("-o", "--out", default="a4_小抄.docx")
    ap.add_argument("--style", default=DEFAULT_STYLE, help="样式 json,默认 config/style.json")
    ap.add_argument("--font-size", type=float, help="覆盖字号(如 5.0 留更多呼吸感)")
    ap.add_argument("--line-spacing-pt", type=float, help="覆盖行距(pt)")
    ap.add_argument("--mono", action="store_true", help="关闭红/蓝颜色(黑白打印)")
    ap.add_argument("--max-pages", type=int, default=2, help="预计纸面数(正面+背面),默认 2")
    args = ap.parse_args()

    cfg = json.load(open(args.style, encoding="utf-8"))
    if args.font_size:
        cfg["font"]["size_pt"] = args.font_size
    if args.line_spacing_pt:
        cfg["font"]["line_spacing_pt"] = args.line_spacing_pt
    if args.mono:
        cfg["color"]["enabled"] = False
    st = Styler(cfg)

    # 内容块文件(每个 txt = 一段,按文件名排序 = 版面顺序,建议 01_ 前缀)
    pattern = args.glob or "块*.txt"
    files = sorted(glob.glob(pattern))
    if not files:
        sys.exit(f"没有匹配 '{pattern}' 的内容块文件。请在块文件所在目录运行,"
                 f"或指定通配符,如: python scripts/make_cheat.py 'blocks/*.txt'")
    texts = [open(f, encoding="utf-8").read().strip() for f in files]
    # 以"\n"连接: 块文件内若有多行,每行也是一个独立紧凑段落(文本行=段落)
    paras = []
    for t in texts:
        paras.extend(x.strip() for x in t.split("\n") if x.strip())

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(cfg["page"]["width_cm"]), Cm(cfg["page"]["height_cm"])
    m = cfg["page"]["margin_cm"]
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(m)

    for text in paras:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(st.line_pt)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        st.add_rich(p, text)

    doc.save(args.out)

    # 粗估行数预警(有效字符宽按字号 0.9 估,与 Word 实际折行有出入)
    usable_cm = cfg["page"]["width_cm"] - 2 * m
    usable_pt = usable_cm * 28.35
    chars_per_line = max(1, usable_pt / (st.size * 0.9))
    page_h = cfg["page"]["height_cm"] - 2 * m
    lines_per_page = (page_h * 28.35) / st.line_pt
    est_lines, total_chars = estimate_lines(paras, chars_per_line)
    capacity = int(lines_per_page * args.max_pages)
    print(f"已生成 {args.out}")
    print(f"  内容块 {len(files)} 个 / 段落 {len(paras)} 段 / 总字符 {total_chars}")
    print(f"  粗估文本行 {est_lines} 行 ≈ {est_lines/lines_per_page:.2f} 页"
          f"(容量 {args.max_pages} 页 / {capacity} 行)")
    if est_lines > capacity:
        print(f"  ⚠ 粗估超页({est_lines}>{capacity}),需压缩内容或调大字号/行距之外的参数。"
              f"最终以 verify_pages.py 的 PDF 页数为准。")
    else:
        print(f"  粗估在容量内,仍请运行 verify_pages.py 做权威页数校验。")


if __name__ == "__main__":
    main()
