# -*- coding: utf-8 -*-
"""verify_pages.py — 页数权威校验:docx → PDF → 数页

排版器的"粗估"与 Word 实际折行有出入,唯一权威是 Word 真实排版后的页数。
本脚本自动调 Word(Windows)或 LibreOffice 把 docx 转 PDF 再数页:
  页数 <= 你的纸面数 → 通过(打印 1 页 = 单面;打印 2 页 = A4 双面)
  页数 >  你的纸面数 → 退出码 1,提示压缩内容。

用法:
  python scripts/verify_pages.py a4.docx                 # 默认要求 <= 2 页
  python scripts/verify_pages.py a4.docx --max-pages 1   # 老师只许单面时
  python scripts/verify_pages.py a4.docx --keep-pdf      # 保留转换出的 PDF 以便预览/打印

依赖(任一即可):  Windows + 装有 Microsoft Word;  或  LibreOffice(soffice)。
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("缺少依赖: pip install pymupdf")


def to_pdf_via_word(docx_path, pdf_path):
    """Windows + Word: 用 COM 转 PDF(FileFormat=17)。"""
    import win32com.client  # pip install pywin32
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close(False)
    finally:
        word.Quit()


def to_pdf_via_soffice(docx_path, pdf_path):
    """LibreOffice 兜底(跨平台): soffice --headless --convert-to pdf"""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    out_dir = os.path.dirname(pdf_path)
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                   check=True, capture_output=True)
    produced = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if os.path.exists(produced):
        os.replace(produced, pdf_path)
    return True


def main():
    ap = argparse.ArgumentParser(description="docx → pdf 页数权威校验")
    ap.add_argument("docx")
    ap.add_argument("--max-pages", type=int, default=2, help="纸面数上限(双面=2),默认 2")
    ap.add_argument("--keep-pdf", action="store_true", help="保留转换出的 PDF(可预览/打印)")
    args = ap.parse_args()
    if not os.path.exists(args.docx):
        sys.exit(f"找不到 {args.docx}")

    tmp = tempfile.mkdtemp(prefix="cheat_verify_")
    pdf_path = os.path.join(tmp, "verify.pdf")
    try:
        ok = False
        if os.name == "nt":
            try:
                to_pdf_via_word(args.docx, pdf_path)
                ok = True
            except Exception as e:
                print(f"Word 转换失败({e}),尝试 LibreOffice…")
        if not ok:
            ok = to_pdf_via_soffice(args.docx, pdf_path)
        if not ok:
            sys.exit("需要转换器: Windows 请装 Microsoft Word(pip install pywin32),"
                     "或安装 LibreOffice 并加入 PATH。")

        with fitz.open(pdf_path) as doc:
            n = len(doc)
        verdict = "✅ 通过" if n <= args.max_pages else f"❌ 超页: {n} > {args.max_pages}"
        print(f"PDF 共 {n} 页(上限 {args.max_pages}) -> {verdict}")
        if n > args.max_pages:
            print("建议: 压缩内容块(砍冗余/合并字段),或把字号压到 4.5pt,"
                  "或向老师确认能否双面。改完重新运行本脚本,直到通过。")
            if args.keep_pdf:
                shutil.copy(pdf_path, os.path.splitext(args.docx)[0] + ".pdf")
            sys.exit(1)
        if args.keep_pdf:
            out_pdf = os.path.splitext(args.docx)[0] + ".pdf"
            shutil.copy(pdf_path, out_pdf)
            print(f"PDF 已保留: {out_pdf}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
