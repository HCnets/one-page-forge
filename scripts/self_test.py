# -*- coding: utf-8 -*-
"""self_test.py — 冒烟自测:验证脚本在改动后仍可用(本地与 CI 共用)

自动执行:
  1. 用 PyMuPDF 现场生成一份 3 页测试 PDF(第1页纯文字 / 第2页嵌入大图 /
     第3页大字标题几乎无文字),模拟课件三种典型页;
  2. 跑 analyze_pdf.py → 断言 A 档含第 3 页、B 档含第 2 页;
  3. 跑 render_pages.py → 断言 3 张 PNG 全部产出;
  4. 跑 make_cheat.py(用 examples/demo_blocks 演示内容)→ 断言 docx 生成。
verify_pages.py 需要 Word/LibreOffice,不进自测(手动跑)。

用法:  python scripts/self_test.py
失败时退出码非 0 并打印原因。CI(.github/workflows/test.yml)每次 push 自动运行。
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)


def make_test_pdf(path):
    import fitz
    doc = fitz.open()
    # 第1页: 纯文字(文字层正常、无图 → 不进任何候选档)
    p = doc.new_page()
    p.insert_text((72, 100), "First page plain text." * 8, fontsize=12)
    # 第2页: 文字层正常 + 一张占页 >35% 的大图(模拟"图承载内容"页 → B档)
    p = doc.new_page()
    p.insert_text((72, 60), "Second page has text, but the big image below is the content.",
                  fontsize=12)
    import io
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (800, 600), (200, 30, 30)).save(buf, "PNG")
    p.insert_image(fitz.Rect(72, 120, 72 + 500, 120 + 420), stream=buf.getvalue())
    # 第3页: 只有一行大标题(≈零文字 → A档)
    p = doc.new_page()
    p.insert_text((72, 200), "Chapter Three", fontsize=48)
    doc.save(path)
    doc.close()


def run(args, cwd=None):
    r = subprocess.run([sys.executable] + args, capture_output=True, text=True, cwd=cwd)
    if r.returncode != 0:
        print("  命令失败:", " ".join(args))
        print(r.stdout)
        print(r.stderr)
        sys.exit(1)
    return r.stdout


def main():
    tmp = tempfile.mkdtemp(prefix="forge_selftest_")
    fails = []
    try:
        print("[1/5] 生成测试 PDF...")
        pdf = os.path.join(tmp, "test.pdf")
        make_test_pdf(pdf)

        print("[2/5] analyze_pdf.py 三档检测...")
        out = run(["scripts/analyze_pdf.py", pdf, "-o", os.path.join(tmp, "ana")], cwd=REPO)
        with open(os.path.join(tmp, "ana", "视觉转录候选页.txt"), encoding="utf-8") as f:
            report = f.read()
        for marker in ("A档", "B档", "C档"):
            if marker not in report:
                fails.append(f"清单缺 {marker} 段")

        def sec_pages(marker):  # 清单格式: "标题行:\n页码行",取第二行
            part = report.split(marker + " ·")[1]
            lines = [l for l in part.split("\n") if l.strip()]
            return lines[1] if len(lines) > 1 else ""
        # 第3页(近零文字)必须进 A 档;第2页(大图)必须进 B 档
        if "3" not in sec_pages("A档"):
            fails.append("第3页(零文字)未进 A 档 -> " + sec_pages("A档"))
        if "2" not in sec_pages("B档"):
            fails.append("第2页(大图)未进 B 档 -> " + sec_pages("B档"))

        print("[3/5] render_pages.py 渲染...")
        render_dir = os.path.join(tmp, "render")
        run(["scripts/render_pages.py", pdf, "--pages", "1-3", "-o", render_dir], cwd=REPO)
        for pno in (1, 2, 3):
            png = os.path.join(render_dir, f"p{pno:03d}.png")
            if not os.path.exists(png) or os.path.getsize(png) < 1000:
                fails.append(f"渲染缺页或文件过小: {png}")

        print("[4/5] make_cheat.py 演示排版...")
        blocks = os.path.join(REPO, "examples", "demo_blocks")
        out_docx = os.path.join(tmp, "out.docx")
        run(["scripts/make_cheat.py", os.path.join(blocks, "块*.txt"), "-o", out_docx], cwd=REPO)
        if not os.path.exists(out_docx) or os.path.getsize(out_docx) < 5000:
            fails.append("make_cheat 输出 docx 缺失或过小")

        print("[5/5] transcribe_images.py 语法检查...")
        run(["scripts/transcribe_images.py", "--help"], cwd=REPO)

    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    if fails:
        print("❌ 自测失败:")
        for f in fails:
            print("  -", f)
        sys.exit(1)
    print("✅ 自测通过: 三档检测命中 + 排版器正常。"
          "\n   提示: verify_pages.py 需 Word/LibreOffice,请手动跑一次确认页数。")


if __name__ == "__main__":
    main()
