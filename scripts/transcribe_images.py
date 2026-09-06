# -*- coding: utf-8 -*-
"""transcribe_images.py — 批量视觉转录:渲染出的图页 PNG → 文本(Markdown)

把"人肉拖图给 AI"这一步自动化:本脚本调用任意 **OpenAI 兼容** 的视觉模型
(base_url + model + key 都由你自己指定),逐张读取 PNG 并转录/筛选,结果
累积写入一个 md 文件,可直接并入复习底稿。

设计原则:
  * 模型无关 —— 只认 OpenAI 兼容 /chat/completions 协议(带 image_url 的
    content 数组)。GLM-4V / MiniMax / 通义 qwen-vl / Kimi / 豆包 / OpenAI
    的视觉模型都兼容,选你本地网络可达、性价比合适的即可;
  * 灰度与密钥只在你自己手里 —— 脚本不含任何厂商私货,key 走环境变量或
    命令行,不落盘;
  * 稳健 —— 失败自动重试 3 次,支持 --resume 断点续传,不会白跑。

两种模式(对应 prompts/02 的两个阶段):
  --mode transcribe  正式转录: 公式/表格/结构 + 图中文字逐字列出
  --mode screen      快速筛选: 只回答"图里是否有文字层没有的信息",
                     用于候选页特别多(>150)时的第一轮,砍转录量

用法:
  # 环境变量方式(key 不留在命令行历史里)
  set OPENAI_API_KEY=sk-xxx            (Windows) / export ...(Linux/mac)
  python scripts/transcribe_images.py --img-dir png_out --mode transcribe ^
      --base-url https://open.bigmodel.cn/api/paas/v1 --model glm-4v-plus

  # 或命令行传 key(注意 shell 历史)
  python scripts/transcribe_images.py --img-dir png_out --mode screen ^
      --base-url https://api.minimaxi.com/v1 --model MiniMax-M3 ^
      --api-key sk-xxx -o 筛选结果.md

依赖: Pillow(发图前自动压缩到 --max-size,适配各家 API 的尺寸/体积限制),
仅用标准库发请求,无其他第三方依赖。
"""
import argparse
import base64
import io
import os
import re
import sys
import time
import urllib.error
import urllib.request

TRANSCRIBE_PROMPT = """你是课件图页转录器。这张图来自课程课件(公式页/示意图/表格页,
这些内容在 PDF 文字层中不存在,只有图片)。请严格转录:
- 含公式: 用纯文本记号,上标用^(如2^n)、下标用_(如W_q)、求和用Σ(注明上下限)、
  希腊字母写"中文名+字母"(如 伽马 γ);矩阵用方括号排版;逐项列出,不省略;
- 含表格: 转成 Markdown 表格,单元格不缩写;
- 概念/流程图: 描述结构(方框/箭头/分支)并把图中所有文字(标题/标注/轴标签/图例)
  逐字列出;
- 含数字结论: 原样保留数字与单位。
看不清的地方老实写"此处看不清",不要编造。
直接输出转录内容,不要思考过程,不要 think/思考标签。"""

SCREEN_PROMPT = """这张图来自课程课件(渲染页)。只做一件事: 判断"图中是否包含文字层
没有的信息"——即图的标题/标注/数据/结构无法从课件正文文字推断。
只输出一行: 有信息 / 无信息(1~3字理由)。不要展开,不要思考过程,不要 think 标签。"""


def compress_to_jpeg_b64(path, max_size):
    """读图→必要时等比缩小→JPEG(q88)编码为 base64(data URI 用)。"""
    from PIL import Image
    im = Image.open(path).convert("RGB")
    if max(im.size) > max_size:
        ratio = max_size / max(im.size)
        im = im.resize((int(im.width * ratio), int(im.height * ratio)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=88)
    return base64.b64encode(buf.getvalue()).decode()


def call_vision(base_url, model, api_key, b64, prompt, max_tokens):
    """单次 /chat/completions 调用,带 3 次重试(指数退避)。"""
    payload = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url",
                 "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
            ],
        }],
    }
    url = base_url.rstrip("/") + "/chat/completions"
    headers = {"Content-Type": "application/json",
               "Authorization": f"Bearer {api_key}"}
    body = __import__("json").dumps(payload).encode()
    last_err = None
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, data=body, headers=headers)
            with urllib.request.urlopen(req, timeout=180) as resp:
                data = __import__("json").loads(resp.read().decode())
            return data["choices"][0]["message"]["content"].strip()
        except Exception as e:  # 网络/限流/格式错误一律重试
            last_err = e
            time.sleep(2 ** attempt)
    raise RuntimeError(f"3 次重试均失败: {last_err}")


def strip_thinking(text):
    """剥离部分模型(如 MiniMax-M3)自带的 <think> 思考段,只留正式回答。"""
    if text.lstrip().startswith("<think>"):
        end = text.find("</think>")
        if end != -1:
            return text[end + len("</think>"):].strip()
        return ""  # 思考段没闭合: 丢弃(宁可少要不可要废)
    return text.strip()


def main():
    ap = argparse.ArgumentParser(description="批量视觉转录(任意 OpenAI 兼容视觉模型)")
    ap.add_argument("--img-dir", required=True, help="PNG 目录(render_pages.py 的输出)")
    ap.add_argument("--mode", choices=["transcribe", "screen"], default="transcribe",
                    help="transcribe=正式转录 / screen=快速筛选(候选页多时先用)")
    ap.add_argument("--base-url", required=True,
                    help="API base,如 https://open.bigmodel.cn/api/paas/v1(通常以 /v1 结尾)")
    ap.add_argument("--model", required=True, help="视觉模型名,如 glm-4v-plus / MiniMax-M3")
    ap.add_argument("--api-key", default=None, help="key;默认读环境变量 OPENAI_API_KEY")
    ap.add_argument("--out", default="图页转录.md", help="输出 md 路径")
    ap.add_argument("--max-size", type=int, default=1600,
                    help="发图前把长边压到该像素(适配 API 体积限制),默认 1600")
    ap.add_argument("--max-tokens", type=int, default=None,
                    help="覆盖单张回复上限(transcribe 默认 3000 / screen 默认 300)")
    ap.add_argument("--resume", action="store_true", help="跳过 out 文件里已转录的图")
    args = ap.parse_args()

    api_key = args.api_key or os.environ.get("OPENAI_API_KEY")
    if not api_key:
        sys.exit("需要 key: 设环境变量 OPENAI_API_KEY 或用 --api-key(选本地可达的模型服务)")

    prompt = SCREEN_PROMPT if args.mode == "screen" else TRANSCRIBE_PROMPT
    max_tokens = args.max_tokens or (300 if args.mode == "screen" else 3000)

    imgs = sorted(f for f in os.listdir(args.img_dir)
                  if f.lower().endswith((".png", ".jpg", ".jpeg")))
    if not imgs:
        sys.exit(f"{args.img_dir} 里没有 PNG/JPG")

    done = set()
    if args.resume and os.path.exists(args.out):
        with open(args.out, encoding="utf-8") as f:
            done = set(re.findall(r"^【([^】]+)】", f.read(), re.M))
    todo = [f for f in imgs if os.path.splitext(f)[0] not in done]

    print(f"共 {len(imgs)} 张,已转录 {len(imgs)-len(todo)},本次处理 {len(todo)} 张 "
          f"(mode={args.mode}, model={args.model})")
    out = open(args.out, "a", encoding="utf-8")
    failed = []
    try:
        for i, name in enumerate(todo, 1):
            tag = os.path.splitext(name)[0]
            print(f"[{i}/{len(todo)}] {name} ...", flush=True)
            try:
                b64 = compress_to_jpeg_b64(os.path.join(args.img_dir, name), args.max_size)
                text = strip_thinking(call_vision(args.base_url, args.model, api_key,
                                                  b64, prompt, max_tokens))
                if text:
                    out.write(f"\n【{tag}】\n{text}\n")
                    out.flush()
                else:
                    raise RuntimeError("模型返回空内容(可能 thinking 段未闭合)")
            except Exception as e:
                print(f"  ✗ 失败: {e}")
                failed.append(name)
    finally:
        out.close()

    if failed:
        print(f"\n⚠ {len(failed)} 张失败(网络/限流): {failed}")
        print("  重跑加 --resume 即可续传。")
    else:
        print(f"\n✅ 全部完成 -> {args.out}")


if __name__ == "__main__":
    main()
