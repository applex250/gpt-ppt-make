#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""下载 Font Awesome 常用图标到本地 (开源库, CC BY 4.0 / SIL OFL)

用法: python3 fetch_fa_icons.py [--count N]
来源: jsdelivr @fortawesome/fontawesome-free (官方开源分发)
"""
import argparse
import urllib.request
from pathlib import Path

FA_VERSION = '6.5.2'
BASE = f'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free@{FA_VERSION}/svgs/solid'
OUT = Path(__file__).parent / 'assets' / 'fa-icons'

# 常用图标 (按使用频率挑选, 可扩展)
ICONS = [
    'house', 'lightbulb', 'chart-bar', 'chart-line', 'chart-pie', 'flask',
    'microscope', 'gears', 'database', 'magnifying-glass', 'robot', 'brain',
    'file', 'folder', 'folder-open', 'download', 'upload', 'trash',
    'pen', 'wand-magic-sparkles', 'bullseye', 'target', 'circle-check',
    'circle-xmark', 'triangle-exclamation', 'circle-info', 'question',
    'graduation-cap', 'book', 'book-open', 'newspaper', 'quote-left',
    'arrow-right', 'arrow-left', 'arrow-up', 'arrow-down', 'arrows-rotate',
    'layer-group', 'sitemap', 'diagram-project', 'network-wired', 'server',
    'cloud', 'code', 'terminal', 'bug', 'rocket', 'cube', 'box',
    'scale-balanced', 'clock', 'calendar', 'user', 'users', 'lock',
    'key', 'paperclip', 'link', 'image', 'table', 'list', 'table-list',
    'thumbs-up', 'thumbs-down', 'heart', 'star', 'award', 'trophy',
    'seedling', 'mountain', 'water', 'fire', 'bolt', 'wave-square',
]


def fetch(name: str) -> bool:
    url = f'{BASE}/{name}.svg'
    out = OUT / f'{name}.svg'
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as r:
            data = r.read()
        if b'<svg' not in data or b'404' in data[:200]:
            return False
        out.write_bytes(data)
        return True
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--count', type=int, default=len(ICONS))
    args = ap.parse_args()
    OUT.mkdir(parents=True, exist_ok=True)
    ok, fail = 0, []
    for name in ICONS[:args.count]:
        if fetch(name):
            ok += 1
        else:
            fail.append(name)
    print(f'图标库: {ok} 个成功, {len(fail)} 个失败')
    if fail:
        print('失败:', ', '.join(fail))
    print(f'位置: {OUT}')


if __name__ == '__main__':
    main()
