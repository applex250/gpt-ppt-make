# -*- coding: utf-8 -*-
"""预生成 facies v3 的 3 张图表 PNG (设计系统配色 + 中文字体 + 合理比例)。
流程原则: 先生成高质量图片, 再用 image 元素加入 PPT (比运行时 chart 渲染稳定)。
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams['font.family'] = 'Noto Sans CJK JP'
plt.rcParams['axes.unicode_minus'] = False

ACCENT = '#2E8B7E'
INFO = '#5B9BD5'
HIGHLIGHT = '#F2B90F'
INK = '#1A1A1A'
MUTED = '#8C8C8C'

# ── 1. 参数量对比柱状图 (840x110, 对数轴) ──
fig, ax = plt.subplots(figsize=(8.4, 1.1), dpi=200)
models = ['seq', 'cnn', 'rnn']
vals = [0.19, 0.06, 2.28]
colors = [ACCENT, ACCENT, ACCENT]
bars = ax.bar(range(3), vals, color=colors, width=0.5)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width()/2, v * 1.12, f'{v:.2f}M',
            ha='center', va='bottom', fontsize=9, color=INK)
ax.set_yscale('log')
ax.set_ylim(0.03, 10)
ax.set_xticks(range(3))
ax.set_xticklabels(models, fontsize=10)
ax.set_yticks([0.1, 1])
ax.set_yticklabels(['0.1', '1'], fontsize=8, color=MUTED)
ax.tick_params(axis='both', length=0)
for spine in ax.spines.values():
    spine.set_color('#BFBFBF')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
fig.tight_layout(pad=0.3)
fig.savefig('/home/user/facies_analysis/facies_ppt_v3/media/chart_params.png', dpi=200, facecolor='white')
plt.close(fig)
print('chart_params.png OK')

# ── 2. 类分布饼图 (420x240, 横向图例) ──
fig, ax = plt.subplots(figsize=(4.2, 2.4), dpi=200)
cats = ['SS', 'CSiS', 'FSiS', 'SiSh', 'MS', 'WS', 'D', 'PS', 'BS']
counts = [259, 738, 615, 353, 217, 300, 98, 498, 93]
palette = ['#2E8B7E', '#5B9BD5', '#F2B90F', '#9AA5B1', '#C0504D',
           '#8E5AA8', '#BFBFBF', '#8C8C8C', '#1A1A1A']
wedges, _, autotexts = ax.pie(
    counts, labels=None, autopct='%1.0f%%', colors=palette[:len(counts)],
    startangle=90, counterclock=False,
    textprops={'fontsize': 7})
for at in autotexts:
    at.set_color('white')
    at.set_fontweight('bold')
    at.set_fontsize(6)
ax.legend(wedges, [f'{c} ({n})' for c, n in zip(cats, counts)],
          loc='center left', bbox_to_anchor=(0.98, 0.5), fontsize=7,
          frameon=False)
ax.set_aspect('equal')
fig.tight_layout(pad=0.3)
fig.savefig('/home/user/facies_analysis/facies_ppt_v3/media/chart_dist.png', dpi=200, facecolor='white')
plt.close(fig)
print('chart_dist.png OK')

# ── 3. CLR 曲线 (320x220) ──
fig, ax = plt.subplots(figsize=(3.2, 2.2), dpi=200)
iters = [0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24]
lrs = [0.000001, 0.000500, 0.001000, 0.001500, 0.001999, 0.001499,
       0.000999, 0.000500, 0.001998, 0.001497, 0.000998, 0.000500, 0.001997]
ax.plot(iters, lrs, marker='o', markersize=3, linewidth=1.6, color=ACCENT)
ax.set_xlabel('iteration', fontsize=8, color=MUTED)
ax.set_ylabel('learning rate', fontsize=8, color=MUTED)
ax.tick_params(axis='both', labelsize=7, length=0)
for spine in ax.spines.values():
    spine.set_color('#BFBFBF')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.grid(axis='y', alpha=0.3)
fig.tight_layout(pad=0.3)
fig.savefig('/home/user/facies_analysis/facies_ppt_v3/media/chart_clr.png', dpi=200, facecolor='white')
plt.close(fig)
print('chart_clr.png OK')
