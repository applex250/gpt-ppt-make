# gpt-ppt-make

**提示词 → PPTD → PPTX 的全本地化生成系统 —— 零 Kimi 云端依赖的 PPT 生成管线。**

由 LLM 根据自然语言需求生成 PPTD(YAML 声明式幻灯片格式)描述文件,再由本地渲染器直接产出 PPTX。图表用 matplotlib 预生成、视觉 QA 全程本地化,输出轻量(约 319KB / 9 页),**不访问 kimi.com / statics.moonshot.cn 任何服务**(实测断网可跑通渲染链路)。

---

## ✨ 特性

| 特性 | 说明 |
|---|---|
| 🚫 **零云端依赖** | 渲染/QA 全程本地,`unshare -n` 断网实测通过,不碰 Kimi/Moonshot |
| 🪶 **轻量输出** | 不嵌入字体,9 页 PPTX 约 319KB(嵌入全量中文字体版 13MB,已弃用) |
| 🖼️ **图表预生成** | matplotlib 先生成 PNG → `image` 元素引用,渲染器不做运行时画图,质量稳定 |
| 🎨 **设计系统驱动** | 内置学术答辩风格(teal-green),配色/字体层级/留白有规范约束 |
| 🧪 **本地视觉 QA** | `pptd_qa.py`:PPTD → 渲染 → PDF → 逐页图,供视觉模型/人工检查 |
| 📦 **格式兼容** | PPTD 为 Moonshot 公开格式,与 Kimi 编辑器通用,本项目为独立本地实现 |

---

## 🧭 背景与动机

Moonshot 的 Kimi PPT 编辑器使用 **PPTD**(YAML 中间格式)抽象 OOXML,每页自包含、所见即所得。但官方 PPTX 导出链路需要:

```
PPTD → 浏览器端 WASM writer(pptd_wasm_bg.wasm) → 签名接口(signatures API) → PPTX
                                             ↑ 未登录返回 401, 导出被锁死
```

对官方 WASM 的逆向分析结论(见下方"技术内幕"):签名是 **Ed25519 公钥验证**,私钥仅在服务端,**数学上不可伪造绕过**。因此本项目走了另一条路:

> **自研本地渲染器,把 PPTD → PPTX 这一步完整本地化。** 不依赖、不绕过任何云端服务,输出完全属于自己的 PPTX。

---

## 🏗️ 架构

```
提示词 (自然语言需求)
   │ ① LLM 生成 PPTD
   ▼
PPTD 项目                         ← 可编辑源: .pptd 主入口 + pages/*.page + media/*
   │ ② pptd2pptx.py 本地渲染器 (python-pptx)
   ▼
PPTX                              ← 轻量, 字体由系统 fallback
   │ ③ pptd_qa.py 本地 QA
   ▼
逐页图 (page-1.jpg ... + manifest.json)
   │ ④ GLM-4.6V / 人工视觉检查
   ▼
交付: deck.pptx + 完整项目目录
```

**设计原则**:
1. **渲染器只放图,不画图** —— `chart` 元素已移除,图表一律预生成 PNG
2. **渲染器只排版,不嵌字** —— 字体不嵌入,轻量化优先
3. **质量责任前置** —— 图表/图片的质量在生成时定死,渲染器只是搬运工

---

## 🚀 快速开始

```bash
# 0. 依赖
pip install python-pptx pillow pyyaml cairosvg   # 渲染
#    系统还需要: soffice(LibreOffice) + pdftoppm(poppler-utils), 仅 QA 环节用

# 1. 渲染 PPTD → PPTX
python3 pptd2pptx/pptd2pptx.py examples/facies/facies_experiment.pptd -o output.pptx --progress

# 2. 本地视觉 QA(生成逐页图)
python3 pptd2pptx/pptd_qa.py examples/facies/facies_experiment.pptd --output-dir .qa-local

# 3. 预览检查
#    .qa-local/page-1.jpg ... page-9.jpg  +  manifest.json
```

---

## 📖 详细使用指南

### CLI 参数

**pptd2pptx.py**(渲染器):

| 参数 | 说明 |
|---|---|
| `input` | `.pptd` 主入口路径(必填) |
| `-o, --output` | 输出 `.pptx` 路径(必填) |
| `--debug` | 打印渲染明细(逐元素) |
| `--progress` | 打印进度百分比(借鉴 WASM onProgress 设计) |

**pptd_qa.py**(本地 QA):

| 参数 | 说明 |
|---|---|
| `input` | `.pptd` 文件路径(必填) |
| `--output-dir` | 页面图输出目录(默认 `.qa-local`) |
| `--dpi` | 渲染 DPI(默认 70) |

### PPTD 格式速览

PPTD 项目结构:

```
deck/
├── deck.pptd              # 主入口: version/title/size/theme/pages
├── pages/
│   └── 1_cover.page       # 每页: background + elements[]
└── media/                 # 图片资源
```

主入口示例(`deck.pptd`):

```yaml
version: v2
title: 演示标题
size: [960, 540]                       # 画布 16:9
theme:
  colors: {primary: "#334047", accent: "#B66A3C"}
  textStyles:
    h1: {fontSize: 28, bold: true, color: "$primary"}
pages:
  - pages/1_cover.page
```

页面示例(`pages/1_cover.page`):

```yaml
pageType: cover
background: {type: solid, color: "#FFFFFF"}
elements:
  - elementId: title
    elementType: text
    bounds: [60, 70, 840, 90]          # [x, y, 宽, 高] 设计像素
    content:
      style: "$h1"
      text: 主标题
```

**支持的元素类型**(实测):

| 类型 | 能力 |
|---|---|
| `text` | rich text:`<p>/<span>/<strong>/<em>/<u>/<s>/<sup>/<sub>/<a>`;样式继承链 inline > para > content > theme > 默认 |
| `shape` | 40+ 种(rect/roundRect/ellipse/箭头/流程图/星形/云/心形等);solid + gradient 填充 |
| `line` | 线段(border 样式) |
| `image` | PNG/JPEG,contain 适配 |
| `table` | 合并单元格、主题样式、三线表 |
| `icon` | Font Awesome 本地 SVG,任意 fill 色(73 个内置) |

**主题能力**:`colors`(含 `$token` 引用)、`textStyles`(字号/粗细/颜色/行高/字体)、`tableStyles`(表头/表体/边框/斑马纹)。

### 图表生成规范(重要)

> ⚠️ **渲染器不支持 `chart` 元素**(已移除)。原因:运行时 matplotlib 画图受画布比例/字体/嵌入适配多变量影响,反复出现比例错乱/留白/挤压问题。

**正确流程**(示例见 `examples/facies/plot_charts.py`):

```python
# 1. matplotlib 预生成 PNG(精确控制尺寸/配色/中文字体/数值标签)
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'Noto Sans CJK JP'   # 中文: 用系统 Noto, 不嵌入
fig, ax = plt.subplots(figsize=(8.4, 1.1), dpi=200)
ax.bar(...)
fig.savefig('media/chart_params.png', dpi=200, facecolor='white')

# 2. PPTD 里用 image 元素引用
#    - elementId: chart-params
#      elementType: image
#      bounds: [60, 382, 840, 110]
#      src: "media/chart_params.png"
#      fit: {mode: contain}
```

**关键细节**:
- 图片尺寸比例 = `bounds` 宽高比(如 840×110 区域配 7.64:1 的图)
- 中文字体用系统 Noto Sans CJK(绘图时注册,不随 PPTX 分发)
- 图表内部标题尽量省掉,页面文字负责标题(避免双标题挤压)

### 字体策略

**不嵌入字体**(轻量化设计决策)。PPTX 内仅写字体名字符串,由打开方系统 fallback:

| 方案 | 体积 | 跨机一致性 |
|---|---|---|
| 不嵌入(当前) | ~319KB | 打开方字体不同则显示不同(可接受) |
| 嵌入全量中文字体 | 13MB+ | 一致,但违背轻量化(已弃用) |

---

## 🎨 设计系统(示例)

`examples/facies` 内置 **teal-green-academic-defense**(学术答辩·青绿)设计系统的落地:

- **配色 tokens**:`bg #FFFFFF` / `ink #1A1A1A` / `accent #2E8B7E` / `line #BFBFBF` 等
- **规范**:纯白底、彩色面积 ≤15%、无卡片(用线条/留白/字号分层)、无三等分布局
- **表格**:三线表(仅顶/底/表头线)
- **列表**:方点标记 ■(accent 色)
- **留白**:主体元素 ≤85%,至少两角留空
- **页脚**:章节导航 + 页码 + 日期署名

在 `.pptd` 的 `theme:` 段落地即可,渲染器按 tokens 解析。

---

## 📁 目录结构

```
gpt-ppt-make/
├── README.md
├── LICENSE                    # MIT
├── .gitignore
├── pptd2pptx/                 # 核心渲染器(967 行)
│   ├── pptd2pptx.py           #   PPTD → PPTX 渲染器(776 行)
│   ├── pptd_qa.py             #   本地视觉 QA(124 行)
│   ├── fetch_fa_icons.py      #   Font Awesome 图标下载器
│   └── assets/fa-icons/       #   73 个本地 SVG 图标
└── examples/
    └── facies/                # 完整示例: 测井岩性相分类实验 PPT
        ├── facies_experiment.pptd
        ├── pages/             #   9 页 .page
        ├── media/             #   图表 PNG + 对比图
        └── plot_charts.py     #   图表预生成脚本
```

---

## 🔬 技术内幕:为什么官方导出不可绕过(逆向分析结论)

对官方 `pptd_wasm_bg.wasm`(Rust wasm-bindgen 模块)的分析:

1. **导出流程**:内容 SHA-256 → POST `signatures` API(需登录 token)→ 服务端 **Ed25519 私钥**签发 → WASM 内嵌**公钥**验证 → 通过才生成 PPTX
2. **不可伪造**:Ed25519 为非对称签名,没有服务端私钥,数学上不可能构造出通过公钥验证的签名
3. **实测证据链**:格式错误 → `must be a base64-encoded 64-byte Ed25519 signature`;伪造合法签名 → `pptd export signature verification failed`
4. **结论**:WASM 公开分发但导出锁在服务端认证后 —— 唯一合法路径是本地自研渲染器(本项目)

> ⚠️ 本项目**不包含也不协助**任何绕过官方签名/访问控制的代码。官方通道在登录态下仍可用作参考。

---

## ⚠️ 已知限制

- **soffice QA 渲染偏差**:LibreOffice 渲染 PDF 与 PowerPoint 实际显示有细微差异(字体 fallback、间距),QA 图用于布局/重叠/内容检查足够,非像素级预览
- **字体 fallback 不可控**:打开方无 PPTD 指定字体时替换显示,跨机不完全一致(轻量化的代价)
- **无可视化编辑器**:改 PPTD 需直接编辑 YAML(或让 LLM 生成)
- **图标依赖本地 SVG 库**:新图标需先 `fetch_fa_icons.py` 下载(或手动放 SVG 进 `assets/fa-icons/`)
- **无动画/过渡**:输出为静态 PPTX(无 fade 等切换效果)

---

## 🔗 与官方 Kimi 生态的关系

- **格式兼容**:PPTD 是 Moonshot 公开定义的 YAML 格式,`examples/facies` 的源文件可直接在 Kimi 编辑器打开
- **独立实现**:渲染器是纯本地 python-pptx 实现,不包含任何 Moonshot 代码/资源(图标为开源 Font Awesome)
- **定位**:官方通道(登录态)适合要原生图表/动画/字体嵌入的场景;本仓库是零依赖、可离线、可复现的替代

---

## 🗺️ 路线图

- [ ] 更多设计系统预设(blue-line / paper-white / dark-data 等)
- [ ] 本地可视化编辑器(拖拽式改 PPTD)
- [ ] CI:自动化渲染测试 + 视觉回归
- [ ] PPTX → PPTD 反向转换(导入现有 PPT)
- [ ] Python 包发布(pip install gpt-ppt-make)

---

## ❓ FAQ

**Q: 和 Kimi PPT 有什么关系?**
A: 共用 PPTD 格式(格式兼容),但本项目是独立的本地渲染实现,不依赖 Kimi 任何服务。

**Q: 为什么不用官方导出?**
A: 官方导出需要登录态(签名 401);且签名机制数学上不可绕过(Ed25519 公钥验证)。本地自研是唯一合法稳定的路径。

**Q: 为什么图表不用 chart 元素?**
A: 运行时画图质量不稳定(比例/留白/字体问题反复出现)。预生成 PNG 把质量责任交给绘图脚本,渲染器只放图,一版通过。

**Q: 支持动画吗?**
A: 不支持,输出静态 PPTX。需要动画可用 Kimi 官方通道(登录态)。

---

## 📄 License

MIT License — Copyright (c) 2026 applex250

Font Awesome 图标:CC BY 4.0 / SIL OFL(开源授权)。示例数据:SEG 2016 相分类竞赛公开数据集。
