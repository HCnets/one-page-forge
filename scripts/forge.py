# -*- coding: utf-8 -*-
"""forge.py — 全自动小抄锻造炉: 课件 PDF → A4 双面小抄 docx(一键)

v0.3 实验版: 把整个流水线(分析→图页转录→分片压缩自审→全稿终审→排版→校验)
编排成一条命令。视觉转录与文本压缩均调用**同一个 OpenAI 兼容服务**
(选同时提供文本模型与视觉模型的服务商;若只想自动做文本部分,视觉转录可
--vision-tier none 跳过,改用 transcribe_images.py 人工补转录)。

用法:
  set OPENAI_API_KEY=sk-xxx
  python scripts/forge.py 课件.pdf ^
      --base-url https://api.minimaxi.com/v1 --model MiniMax-M3 ^
      --chapters 1-80,81-160,161-250,251-330,331-450,451-560,561-677 ^
      --out-dir forge_out

参数速览:
  --chapters "a-b,c-d"   章节页码范围(强烈建议给,课件首页目录 10 秒能数出来);
                         不给则按每 --chunk-pages 页自动切片
  --vision-model NAME    视觉模型(默认同 --model;给 none 关闭视觉转录)
  --vision-tier auto     auto=A+B必转+C档筛选后转(默认)/ full=候选全直接转(贵)/
                         none=跳过视觉
  --target-chars 50000   小抄总字符预算(≈4.5pt 双面容量)
  --parallel 2           分片压缩并发数
  --resume               断点续跑(中间产物都在 --out-dir/work 下)

流程与中间产物(都在 --out-dir 下):
  work/analyze/           逐页全文 + 三档候选清单
  work/png/               候选图页渲染
  work/vision_*.md        视觉筛选/转录结果
  blocks/块NN_*.txt       内容块(可直接给 make_cheat.py)
  小抄.docx               成品

质量护栏内嵌: 分片考点条数过少/块格式非法 → 自动重试一次;终审补丁无法解析 →
自动重试。护栏拦不住的质量问题,请拿成品对照 prompts/04 做人工终审。
退出码: 0=成功(可能含超页警告) / 1=失败或超页。
"""
import argparse
import base64 as b64
import concurrent.futures as cf
import io as _io
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
CHAR_PER_TOKEN = 0.75        # 中文 ≈0.7~1 token/字,保守取 0.75

# ---------- 协议提示词(与 prompts/ 同步,程序化版本) ----------

COMPRESS_PROMPT = """你是"极限考试小抄"排版师兼审计员。我给你课件某一"片"(页码范围见文末,
每页以 `===== 第 N 页 =====` 标记)。请在一轮内完成压缩+自审,输出终稿。

压缩纪律:
- 只留可考内容(定义/机制/公式/数字/对比/易错点),背景/客套/工程细节全砍;
- 句子语义完整可直接誊抄;公式原样并紧跟符号说明;数字与单位不改写;
- 同一知识点只出现一次,不跨条目重复;
- 课件内部自相矛盾处: 保留主流说法并括注"(课件 p.X 为另一说法,疑笔误)"。

排版语法(重要,终端用户直接打印):
- 块首行 `【编号·片名】`(编号用两位数字,如 01/02);
- 同一自然句写一行,不要手动换行;
- 关键名词用 〈〉 括起(整块≤2个);需强调短语用 **半角星号成对** 括起(≤4个);
- "名词: 内容"字段名自然写;数字照常写(排版自动变蓝)。

立刻自审(不许跳过):
- 基于刚读的原文,列出本片全部可考知识点清单(逐条带页码);
- 逐条核对压缩是否覆盖(语义等价算覆盖;只出现名词没给定义=缺口;并列讲3条只写
  2条=缺口;公式缺符号说明、数字丢失=缺口),缺口就地补进块(每条≤40字);
- 目标块字数约 {target_chars} 字(±25%): 明显超了压冗余,明显少了查漏。

输出格式(严格四段,画蛇添足的部分会被丢弃):
块内容开始
【01·片名】
(块内容,一个自然句一行)
块内容结束
---考点清单---
- p.页码 考点简述(每行一条,至少 {min_points} 条)
"""

FINISH_PROMPT = """你是考试小抄的全稿终审员。材料: ①《考点清单》: 各片压缩时列出的
可考知识点(带页码);②《小抄全稿》: 所有内容块,每块以 【数字·片名】 开头
(块文件名为 块NN_...,NN 为两位数字)。
请检查: 每条考点清单是否在某个块有落点;同一知识点是否在≥2块重复(保留表述更全的,
其余给删除补丁);同一实体数字/口径跨块是否打架(以出现次数多的为准,另一处给改法)。
输出最小补丁列表,每一行一种,只输出补丁行,不要解释:
块NN|ADD|要追加的补丁文本(≤40字,可用标记语法)
块NN|DEL|块内需删除的原文子串(必须逐字出现在该块)
块NN|FIX|原文子串→替换文本(替换文本≤60字)
NN 必须与块文件名的两位数字一致(如 块03)。没有问题要改也要给"检查|维度名|条数"行。
全部结束后输出一行: 终审完成,共X条补丁
"""

# ---------- LLM 调用 ----------

def chat_once(base_url, model, api_key, system, user, max_tokens=6000):
    """单次对话调用: 网络错误重试 3 次;输出被 max_tokens 截断(finish_reason=
    'length')时自动翻倍上限重试(thinking 模型常见,最多翻到 24000)。"""
    url = base_url.rstrip("/") + "/chat/completions"
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    last = None
    for attempt in range(3):
        try:
            while True:
                payload = {"model": model, "max_tokens": max_tokens,
                           "messages": [{"role": "system", "content": system},
                                        {"role": "user", "content": user}]}
                req = urllib.request.Request(url, data=json.dumps(payload).encode(),
                                             headers=headers)
                with urllib.request.urlopen(req, timeout=600) as resp:
                    data = json.loads(resp.read().decode())
                choice = data["choices"][0]
                msg = choice["message"]["content"].strip()
                u = data.get("usage", {})
                if choice.get("finish_reason") == "length" and max_tokens < 24000:
                    max_tokens = min(max_tokens * 2, 24000)
                    print(f"      ⚠ 输出被截断(finish=length),max_tokens 翻倍到 {max_tokens} 重试")
                    continue
                return msg, u.get("prompt_tokens", 0), u.get("completion_tokens", 0)
        except Exception as e:
            last = e
            time.sleep(2 ** attempt)
    raise RuntimeError(f"模型调用 3 次重试均失败: {last}")


def strip_thinking(text):
    if text.lstrip().startswith("<think>"):
        end = text.find("</think>")
        return text[end + len("</think>"):].strip() if end != -1 else ""
    return text.strip()


class Usage:
    def __init__(self):
        self.p = self.c = 0
        self.lock = threading.Lock()

    def add(self, p, c):
        with self.lock:
            self.p += p
            self.c += c
            print(f"    [用量] 累计 {self.p/1e4:.0f}万 in + {self.c/1e4:.0f}万 out", flush=True)


USAGE = Usage()

# ---------- 本地小工具 ----------

def run_script(script, *args):
    r = subprocess.run([sys.executable, os.path.join(HERE, script)] + list(args),
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        raise RuntimeError(f"{script} 失败:\n{r.stdout}\n{r.stderr}")


def est_tokens(s):
    return int(len(s) * CHAR_PER_TOKEN)


def load(path):
    return open(path, encoding="utf-8").read()


def save(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def extract_pages(text):
    """{页号: 页文本},按页标记切。"""
    segs = re.split(r"^===== 第 (\d+) 页", text, flags=re.M)
    pages = {}
    for i in range(1, len(segs) - 1, 2):
        pages[int(segs[i])] = segs[i + 1].strip()
    return pages


def merge_vision(text, vision_md):
    """把图页转录文本注入对应页段(标记行之后)。"""
    if not vision_md or not os.path.exists(vision_md):
        return text
    recs = re.findall(r"【(p?\d+)】\n(.*?)(?=\n【|$)", load(vision_md), re.S)
    for tag, body in recs:
        pno = int(re.sub(r"\D", "", tag))
        marker = f"===== 第 {pno} 页"
        if marker in text:
            inject = "".join("    [图页转录] " + l for l in body.splitlines())
            text = text.replace(marker + " ", marker + " \n" + inject + "\n", 1)
    return text


def parse_blocks(blocks_dir):
    """{块文件名(去.txt): 文本},按字典序。"""
    return {os.path.splitext(f)[0]: load(os.path.join(blocks_dir, f))
            for f in sorted(os.listdir(blocks_dir)) if f.endswith(".txt")}


def read_tier_lists(list_path):
    """从 视觉转录候选页.txt 解析三档页码。"""
    txt = load(list_path)
    out = {}
    for label in ("A档", "B档", "C档"):
        m = re.search(rf"{label} · [^\n]*\n([0-9, ]+)", txt)
        pages = []
        if m:
            pages = [int(x) for x in m.group(1).replace(" ", "").split(",") if x]
        out[label] = pages
    return out


# ---------- 步骤 ----------

def step_analyze(pdf, work):
    ana_dir = os.path.join(work, "analyze")
    list_path = os.path.join(ana_dir, "视觉转录候选页.txt")
    if not os.path.exists(list_path):
        run_script("analyze_pdf.py", pdf, "-o", ana_dir)
    full_path = None
    for f in os.listdir(ana_dir):
        if f.endswith("_全文.txt"):
            full_path = os.path.join(ana_dir, f)
    if not full_path:
        raise RuntimeError("analyze 输出缺失")
    return full_path, ana_dir


def step_render(pdf, ana_dir, png_dir, tiers):
    """渲染 A+B+C 全部候选页(后续 screen 需要 C 档页)。"""
    if os.path.exists(png_dir) and any(f.endswith(".png") for f in os.listdir(png_dir)):
        return png_dir
    all_pages = sorted(set(tiers["A档"] + tiers["B档"] + tiers["C档"]))
    if all_pages:
        run_script("render_pages.py", pdf, "--pages",
                   ",".join(str(p) for p in all_pages), "-o", png_dir)
    else:
        os.makedirs(png_dir, exist_ok=True)
    return png_dir


def transcribe_dir(img_dir, out_md, base_url, model, api_key, screen):
    """对目录内图片逐张调视觉模型;结果写 out_md(think 剥离、逐张落盘)。"""
    if os.path.exists(out_md) and os.path.getsize(out_md) > 0:
        return
    from PIL import Image
    prompt = ("这张图来自课程课件(渲染页)。先把图内所有文字逐字读一遍(标题/标注/数据/"
              "图例),再判断: 图中是否有**文字层没有**的信息——即无法从课件正文文字推断的"
              "内容。只输出一行: 有信息 / 无信息(1~3字理由,注明依据)。不要 think 标签。"
              if screen else
              "你是课件图页转录器(公式/表格/示意图,这些内容在 PDF 文字层不存在,"
              "只有图片)。公式用纯文本: 上标^(如2^n) 下标_(如W_q) 求和Σ(注上下限) "
              "希腊字母写'中文名+字母'(如 伽马 γ);表格转 Markdown;图内文字逐字列出;"
              "数字结论原样保留;看不清写'此处看不清';不要编造;不要 think 标签。")
    max_tokens = 300 if screen else 2500
    imgs = sorted(f for f in os.listdir(img_dir)
                  if f.lower().endswith((".png", ".jpg", ".jpeg")))
    if not imgs:
        save(out_md, "")
        return
    with open(out_md, "w", encoding="utf-8") as out:
        for i, name in enumerate(imgs, 1):
            pno = int(re.sub(r"\D", "", name))
            label = "筛选" if screen else "转录"
            print(f"    [视觉·{label}] {name} ({i}/{len(imgs)})", flush=True)
            im = Image.open(os.path.join(img_dir, name)).convert("RGB")
            if max(im.size) > 1600:
                r = 1600 / max(im.size)
                im = im.resize((int(im.width * r), int(im.height * r)), Image.LANCZOS)
            buf = _io.BytesIO()
            im.save(buf, "JPEG", quality=88)
            payload = {"model": model, "max_tokens": max_tokens,
                       "messages": [{"role": "user", "content": [
                           {"type": "text", "text": prompt},
                           {"type": "image_url",
                            "image_url": {"url": "data:image/jpeg;base64,"
                                                  + b64.b64encode(buf.getvalue()).decode()}}]}]}
            url = base_url.rstrip("/") + "/chat/completions"
            headers = {"Content-Type": "application/json",
                       "Authorization": f"Bearer {api_key}"}
            text = "(转录失败,请用 transcribe_images.py 人工补)"
            pt = ct = 0
            for attempt in range(3):
                try:
                    req = urllib.request.Request(url, data=json.dumps(payload).encode(),
                                                 headers=headers)
                    with urllib.request.urlopen(req, timeout=600) as resp:
                        data = json.loads(resp.read().decode())
                    text = strip_thinking(data["choices"][0]["message"]["content"])
                    u = data.get("usage", {})
                    pt, ct = u.get("prompt_tokens", 0), u.get("completion_tokens", 0)
                    break
                except Exception as e:
                    if attempt == 2:
                        print(f"      ✗ {name} 失败: {e}")
                    else:
                        time.sleep(2 ** attempt)
            USAGE.add(pt, ct)
            out.write(f"\n【p{pno:03d}】\n{text}\n")
            out.flush()


def step_vision(base_url, vision_model, api_key, ana_dir, png_dir, work, tier):
    """返回转录 md 路径(含 A 档 + 命中页)。tier: auto/all/none。"""
    out = os.path.join(work, "vision_transcribe.md")
    if tier == "none":
        save(out, "")
        return out
    if os.path.exists(out) and os.path.getsize(out) > 0:
        return out
    tiers = read_tier_lists(os.path.join(ana_dir, "视觉转录候选页.txt"))
    a, b, c = tiers["A档"], tiers["B档"], tiers["C档"]

    def copy_pages(pages, dst_dir):
        os.makedirs(dst_dir, exist_ok=True)
        for pno in pages:
            src = os.path.join(png_dir, f"p{pno:03d}.png")
            if os.path.exists(src):
                shutil.copy(src, os.path.join(dst_dir, os.path.basename(src)))

    if tier == "full":
        hit_dir = os.path.join(work, "png_hit")
        copy_pages(sorted(set(a + b + c)), hit_dir)
        transcribe_dir(hit_dir, out, base_url, vision_model, api_key, screen=False)
        print(f"  [视觉] full 档: 直接转录 {len(a)+len(b)+len(c)} 页")
        return out

    # auto: A+B 档无条件转录(B 档已是"图承载内容"的高置信信号);
    # C 档(单图 8-25% 占比)先 screen 筛选,命中再转录——C 档可能漏,
    # 渲染出的 PNG 建议肉眼抽查一遍(README FAQ)。
    hit = set(a) | set(b)
    if c:
        scan_dir = os.path.join(work, "png_scan")
        copy_pages(sorted(c), scan_dir)
        scan_md = os.path.join(work, "vision_screen.md")
        transcribe_dir(scan_dir, scan_md, base_url, vision_model, api_key, screen=True)
        if os.path.exists(scan_md):
            for tag, verdict in re.findall(r"【p(\d+)】\n(.+)", load(scan_md), re.M):
                low = verdict.lower()
                if "无信息" not in low and "没有信息" not in low and "no info" not in low:
                    hit.add(int(tag))
    hit = sorted(hit)
    hit_dir = os.path.join(work, "png_hit")
    copy_pages(hit, hit_dir)
    transcribe_dir(hit_dir, out, base_url, vision_model, api_key, screen=False)
    print(f"  [视觉] auto 档: A+B 必转 {len(a)+len(b)} + C 筛选命中 {len(hit)-len(a)-len(b)} 页")
    return out


def build_chunks(full_text, chapters, chunk_pages):
    pages = extract_pages(full_text)
    ordered = sorted(pages)
    if not ordered:
        raise RuntimeError("全文为空——课件是扫描版?请先 OCR(见 README FAQ)")
    ranges = []
    if chapters:
        for part in chapters.split(","):
            a, b = part.split("-")
            ranges.append((int(a), int(b)))
    else:
        for i in range(0, len(ordered), chunk_pages):
            ranges.append((ordered[i], ordered[min(i + chunk_pages - 1, len(ordered) - 1)]))
    chunks = []
    for a, b in ranges:
        body = "\n\n".join(f"===== 第 {p} 页 =====" + pages[p]
                           for p in ordered if a <= p <= b)
        if body.strip():
            chunks.append((a, b, body))
    return chunks


def compress_chunk(idx, a, b, body, model, base_url, api_key, target_chars,
                   blocks_dir, points_dir):
    """单片压缩+自审;护栏: 块首行【、考点条数下限;失败重试一次。"""
    fname = f"块{idx:02d}_片{a}-{b}.txt"
    bpath = os.path.join(blocks_dir, fname)
    ppath = os.path.join(points_dir, f"{idx:02d}.txt")
    if os.path.exists(bpath) and os.path.exists(ppath) and os.path.getsize(bpath) > 200:
        print(f"  [压缩] 片{idx}(p{a}-{b}) 已存在,跳过")
        return
    pages_n = len(re.findall(r"^===== 第 (\d+) 页", body, re.M))
    min_points = max(5, pages_n // 12)
    user = f"课件片 {idx}(页 {a}-{b},共 {pages_n} 页)全文如下:\n\n{body}"
    sys_prompt = COMPRESS_PROMPT.format(target_chars=target_chars, min_points=min_points)
    for attempt in range(2):
        print(f"  [压缩] 片{idx}(p{a}-{b}, ~{est_tokens(body)/1e4:.0f}万tok) 调用模型...",
              flush=True)
        raw, pt, ct = chat_once(base_url, model, api_key, sys_prompt, user, max_tokens=8000)
        USAGE.add(pt, ct)
        raw = strip_thinking(raw)
        m = re.search(r"块内容开始\n(.*?)块内容结束", raw, re.S)
        pm = re.search(r"---考点清单---\n(.*?)$", raw, re.S)
        if not m:
            save(os.path.join(os.path.dirname(blocks_dir), "debug", f"片{idx}_试{attempt+1}.txt"),
                 raw)
            user += "\n\n上一轮缺少'块内容开始/块内容结束'标记,请严格按格式重出。"
            continue
        block = m.group(1).strip()
        points = "\n".join(l.strip() for l in (pm.group(1) if pm else "").splitlines()
                           if re.match(r"^-", l.strip()))
        n_points = len([l for l in points.splitlines() if l.strip()])
        problems = []
        if not block.startswith("【"):
            problems.append("块首行不是【开头")
        if "……" in block[-80:] or "等。" in block[-40:]:
            problems.append("块尾疑似省略敷衍")
        if n_points < min_points:
            problems.append(f"考点清单仅 {n_points} 条 < 要求 {min_points}")
        if problems:
            save(os.path.join(os.path.dirname(blocks_dir), "debug", f"片{idx}_试{attempt+1}.txt"),
                 raw)
            user += "\n\n上一轮输出问题: " + "; ".join(problems) + "。修正后严格按格式重出。"
            continue
        save(bpath, block)
        save(ppath, points)
        print(f"  [压缩] 片{idx} 完成: 块 {len(block)} 字符 / 考点 {n_points} 条")
        return
    save(bpath, f"【{idx:02d}·片{a}-{b}】(自动压缩两次未过护栏,请人工按 prompts/03 补)")
    save(ppath, "")
    print(f"  ⚠ 片{idx} 两次尝试未过护栏,已留占位,请人工处理")


def apply_patches(blocks_dir, blocks, lines):
    """终审补丁落盘。块号匹配按两位数字容错(块1→块01)。"""
    def resolve(key):
        if key in blocks:
            return key
        num = re.search(r"\d+", key)
        if not num:
            return None
        cands = [k for k in blocks if re.search(rf"块0*{num.group()}", k)]
        return cands[0] if cands else None

    n = 0
    for line in lines:
        parts = [x.strip() for x in line.split("|")]
        if len(parts) < 3 or not parts[0].startswith("块"):
            continue
        key = resolve(parts[0])
        if not key:
            continue
        op, arg = parts[1], parts[2]
        text = blocks[key]
        if op == "ADD":
            blocks[key] = text + ("\n" if text else "") + arg
            n += 1
        elif op == "DEL" and arg in text:
            blocks[key] = text.replace(arg, "", 1)
            n += 1
        elif op == "FIX" and "→" in arg:
            src, dst = arg.split("→", 1)
            if src in text:
                blocks[key] = text.replace(src, dst, 1)
                n += 1
    for key, val in blocks.items():
        if key.startswith("块"):
            save(os.path.join(blocks_dir, key + ".txt"), val)
    return n


def finish_audit(blocks_dir, points_dir, model, base_url, api_key):
    blocks = parse_blocks(blocks_dir)
    point_parts = []
    for f in sorted(os.listdir(points_dir)):
        t = load(os.path.join(points_dir, f)).strip()
        if t:
            point_parts.append(f"片{f.split('.')[0]}:\n{t}")
    if not point_parts:
        print("  [终审] 无考点清单,跳过(质量风险: 无法查跨片遗漏)")
        return -1
    full = "\n\n".join(f"{k}: {v}" for k, v in blocks.items())
    user = "《考点清单》:\n" + "\n\n".join(point_parts) + "\n\n《小抄全稿》:\n" + full
    for attempt in range(2):
        print("  [终审] 全稿终审调用模型...", flush=True)
        raw, pt, ct = chat_once(base_url, model, api_key, FINISH_PROMPT, user, max_tokens=4000)
        USAGE.add(pt, ct)
        raw = strip_thinking(raw)
        # 接受两类行: 补丁行(块NN|ADD/DEL/FIX|...)与检查行(检查|维度|条数)
        # 模型输出"检查|覆盖|18条"这类行是正常履职,不能当格式失败
        lines = [l for l in raw.splitlines()
                 if ("|" in l and re.match(r"^(块|检查)", l.strip()))
                 or l.startswith("终审完成")]
        if not lines:
            save(os.path.join(os.path.dirname(points_dir), "debug", f"终审_试{attempt+1}.txt"),
                 raw)
            user += ("\n\n上一轮没有输出任何'块NN|ADD/DEL/FIX|内容'或'检查|维度|条数'行,"
                     "请去掉序号与列表符号,每行直接以块NN或检查开头重出。")
            continue
        n = apply_patches(blocks_dir, blocks, lines)
        print(f"  [终审] 完成,应用补丁 {n} 条")
        return n
    print("  ⚠ 终审两次格式失败,原始输出已存 work/debug/,请手动按 prompts/04 执行")
    return -1


def main():
    ap = argparse.ArgumentParser(description="forge: 课件 PDF → A4 小抄 全自动流水线")
    ap.add_argument("pdf")
    ap.add_argument("--base-url", required=True, help="OpenAI 兼容服务地址")
    ap.add_argument("--model", required=True, help="文本模型名")
    ap.add_argument("--vision-model", default=None,
                    help="视觉模型名(默认同 --model;给 none 关闭视觉转录)")
    ap.add_argument("--api-key", default=None, help="key(默认环境变量 OPENAI_API_KEY)")
    ap.add_argument("--chapters", default=None, help="章页码范围 'a-b,c-d'(强烈建议)")
    ap.add_argument("--chunk-pages", type=int, default=70, help="无 --chapters 时每片页数")
    ap.add_argument("--vision-tier", choices=["auto", "full", "none"], default="auto",
                    help="auto=A+B必转+C筛选后转(默认)/ full=候选全直接转(贵)/ none=跳过")
    ap.add_argument("--target-chars", type=int, default=50000, help="小抄总字符预算")
    ap.add_argument("--font-size", type=float, default=None, help="透传排版字号")
    ap.add_argument("--max-pages", type=int, default=2)
    ap.add_argument("--parallel", type=int, default=2, help="分片压缩并发数")
    ap.add_argument("--out-dir", default="forge_out")
    args = ap.parse_args()

    api_key = args.api_key or os.environ.get("OPENAI_API_KEY")
    if not api_key:
        sys.exit("需要 key: 环境变量 OPENAI_API_KEY 或 --api-key")
    vision_model = args.model if args.vision_model in (None, "none") else args.vision_model
    vision_tier = "none" if not vision_model else args.vision_tier

    os.makedirs(args.out_dir, exist_ok=True)
    work = os.path.join(args.out_dir, "work")
    blocks_dir = os.path.join(args.out_dir, "blocks")
    points_dir = os.path.join(work, "points")
    for d in (work, blocks_dir, points_dir):
        os.makedirs(d, exist_ok=True)

    t0 = time.time()
    print("== [1/6] 分析 PDF ==")
    full_path, ana_dir = step_analyze(args.pdf, work)
    full_text = load(full_path)

    vision_md = os.path.join(work, "vision_transcribe.md")
    if vision_model:
        print("== [2/6] 渲染候选图页 ==")
        tiers = read_tier_lists(os.path.join(ana_dir, "视觉转录候选页.txt"))
        png_dir = os.path.join(work, "png")
        step_render(args.pdf, ana_dir, png_dir, tiers)
        print("== [3/6] 视觉转录 ==")
        vision_md = step_vision(args.base_url, vision_model, api_key,
                                ana_dir, png_dir, work, vision_tier)
    else:
        print("== [2-3/6] 视觉转录关闭(--vision-model none) ==")
    full_text = merge_vision(full_text, vision_md)
    merged_path = os.path.join(work, "全_带转录.txt")
    save(merged_path, full_text)

    print("== [4/6] 分片压缩+自审 ==")
    chunks = build_chunks(full_text, args.chapters, args.chunk_pages)
    if not chunks:
        sys.exit("没有可处理的内容(检查 --chapters 页码范围)")
    total_pages = sum(b - a + 1 for a, b, _ in chunks)
    jobs = []
    for idx, (a, b, body) in enumerate(chunks, 1):
        share = (b - a + 1) / total_pages
        # 预算按页占比分,但不超过输入 40%(防小输入片被要求灌水凑字数)
        target = min(int(args.target_chars * share * 1.08), int(len(body) * 0.4))
        jobs.append((idx, a, b, body, max(800, target)))
    with cf.ThreadPoolExecutor(max_workers=args.parallel) as ex:
        futs = []
        for idx, a, b, body, tgt in jobs:
            futs.append(ex.submit(compress_chunk, idx=idx, a=a, b=b, body=body,
                                  target_chars=tgt, model=args.model,
                                  base_url=args.base_url, api_key=api_key,
                                  blocks_dir=blocks_dir, points_dir=points_dir))
        for f in cf.as_completed(futs):
            f.result()

    print("== [5/6] 全稿终审 ==")
    finish_audit(blocks_dir, points_dir, args.model, args.base_url, api_key)

    print("== [6/6] 排版 + 页数校验 ==")
    docx = os.path.join(args.out_dir, "小抄.docx")
    cmd = [sys.executable, os.path.join(HERE, "make_cheat.py"),
           os.path.join(blocks_dir, "块*.txt"), "-o", docx]
    if args.font_size:
        cmd += ["--font-size", str(args.font_size)]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    print(r.stdout)
    if r.returncode != 0:
        print(r.stderr)
        sys.exit(1)
    r = subprocess.run([sys.executable, os.path.join(HERE, "verify_pages.py"), docx,
                        "--max-pages", str(args.max_pages)],
                       capture_output=True, text=True, encoding="utf-8")
    print(r.stdout or r.stderr)

    print(f"\n锻造完成,耗时 {(time.time()-t0)/60:.1f} 分钟")
    print(f"产物: {docx}\n内容块: {blocks_dir}(可手动编辑后重跑排版/校验)")
    if r.returncode != 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
