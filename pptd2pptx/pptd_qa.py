#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptd_qa.py — PPTD 本地视觉 QA 流水线(替代走 Kimi 云端的 export_images.py)

链路: .pptd → ①pptd2pptx.py 渲染 PPTX(临时目录) → ②soffice 转 PDF(临时目录)
      → ③pdftoppm 输出逐页 JPEG 到输出目录 → ④写 manifest.json(每页 index+image 路径)

用法:
    python3 pptd_qa.py <deck.pptd> [--output-dir DIR] [--dpi N]
"""
import argparse
import glob
import json
import os
import re
import subprocess
import sys
import tempfile
from datetime import datetime, timezone

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
RENDERER = os.path.join(SCRIPT_DIR, "pptd2pptx.py")


def run(cmd, cwd=None, timeout=600):
    """执行外部命令; 失败时抛出带清晰错误信息的 RuntimeError。"""
    print("$ " + " ".join(cmd))
    try:
        proc = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, timeout=timeout)
    except FileNotFoundError as e:
        raise RuntimeError(f"找不到命令: {cmd[0]} (请确认已安装并加入 PATH)") from e
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"命令超时({timeout}s): {' '.join(cmd)}") from e
    if proc.returncode != 0:
        tail = (proc.stderr or "").strip().splitlines()[-15:]
        if not tail:
            tail = (proc.stdout or "").strip().splitlines()[-15:]
        detail = "\n".join(tail) if tail else "(无输出)"
        raise RuntimeError(
            f"命令执行失败 (exit={proc.returncode}): {' '.join(cmd)}\n"
            f"--- 输出尾部 ---\n{detail}"
        )
    return proc


def main():
    ap = argparse.ArgumentParser(
        description="PPTD 本地视觉 QA: 渲染 PPTX → PDF → 逐页 JPEG + manifest.json"
    )
    ap.add_argument("input", help=".pptd 文件路径")
    ap.add_argument("--output-dir", default=".qa-local", help="页面图输出目录 (默认: .qa-local)")
    ap.add_argument("--dpi", type=int, default=70, help="渲染 DPI (默认: 70)")
    args = ap.parse_args()

    pptd_path = os.path.abspath(args.input)
    if not os.path.isfile(pptd_path):
        sys.exit(f"错误: 找不到输入文件: {pptd_path}")
    if not os.path.isfile(RENDERER):
        sys.exit(f"错误: 找不到渲染器: {RENDERER}")
    if args.dpi <= 0:
        sys.exit("错误: --dpi 必须为正整数")

    out_dir = os.path.abspath(args.output_dir)
    os.makedirs(out_dir, exist_ok=True)

    # 清理旧的 page-*.jpg 与 manifest.json, 避免陈旧文件污染本次结果
    for old in glob.glob(os.path.join(out_dir, "page-*.jpg")) + [os.path.join(out_dir, "manifest.json")]:
        if os.path.isfile(old):
            os.remove(old)

    try:
        with tempfile.TemporaryDirectory(prefix="pptd_qa_") as tmp:
            pptx_path = os.path.join(tmp, "deck.pptx")

            # ① 渲染 PPTX: 在渲染器自身目录运行, 保证 assets 相对路径可用
            run([sys.executable, RENDERER, pptd_path, "--output", pptx_path], cwd=SCRIPT_DIR)
            if not os.path.isfile(pptx_path) or os.path.getsize(pptx_path) == 0:
                raise RuntimeError(f"渲染器未生成有效的 PPTX: {pptx_path}")

            # ② soffice → PDF (隔离 UserInstallation profile, 避免锁冲突/多实例问题)
            lo_profile = os.path.join(tmp, "lo_profile")
            run([
                "soffice", "--headless",
                "-env:UserInstallation=file://" + lo_profile,
                "--convert-to", "pdf", "--outdir", tmp, pptx_path,
            ])
            pdf_path = os.path.join(tmp, "deck.pdf")
            if not os.path.isfile(pdf_path):
                raise RuntimeError(f"soffice 未生成 PDF: {pdf_path}")

            # ③ pdftoppm → 逐页 JPEG
            run(["pdftoppm", "-jpeg", "-r", str(args.dpi), pdf_path, os.path.join(out_dir, "page")])

        # ④ 收集页面并写 manifest.json
        page_re = re.compile(r"^page-(\d+)\.jpg$")
        pages = []
        for f in os.listdir(out_dir):
            m = page_re.match(f)
            if m:
                pages.append((int(m.group(1)), os.path.join(out_dir, f)))
        pages.sort()
        if not pages:
            sys.exit("错误: pdftoppm 未生成任何 page-*.jpg")

        manifest = {
            "source": pptd_path,
            "dpi": args.dpi,
            "page_count": len(pages),
            "created_at": datetime.now(timezone.utc).isoformat(),
            "pages": [{"index": num, "image": path} for num, path in pages],
        }
        manifest_path = os.path.join(out_dir, "manifest.json")
        with open(manifest_path, "w", encoding="utf-8") as fh:
            json.dump(manifest, fh, ensure_ascii=False, indent=2)

        print(f"本地 QA 完成: {len(pages)} 页 → {out_dir}")
        print(f"manifest: {manifest_path}")
    except RuntimeError as e:
        sys.exit(f"QA 失败: {e}")


if __name__ == "__main__":
    main()
