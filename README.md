# gpt-ppt-make

**提示词 → PPTD → PPTX 的全本地化生成系统。**

由 LLM 根据需求生成 PPTD(YAML 声明式幻灯片格式)描述文件,再由本地渲染器直接产出 PPTX —— 全程不依赖任何云端 PPT 服务(kimi.com / statics.moonshot.cn),图表预生成、视觉 QA 本地化,输出轻量(约 300KB/9 页)。

## 为什么做这个

Moonshot 的 Kimi PPT 编辑器使用 PPTD(YAML)作为中间格式,但官方 PPTX 导出需要登录态(签名接口 401 阻塞)。本项目把"PPTD → PPTX"这一步完整本地化:

- **零云端依赖**:渲染/QA 全程本地,不访问 kimi.com / statics.moonshot.cn(实测断网可跑通)
- **轻量化**:不嵌入字体,9 页 PPTX 约 319KB(嵌入字体版 13MB,已弃用)
- **图表预生成**:matplotlib 先生成 PNG,PPTD 用 `image` 元素引用(渲染器不做运行时画图,保证质量稳定)
- **本地视觉 QA**:PPTD → 渲染 → PDF → 逐页图,供视觉模型检查布局

## 架构

```
提示词 (自然语言)
   │  LLM 生成
   ▼
PPTD 项目 (YAML: .pptd 主入口 + pages/*.page + media/*)
   │  pptd2pptx.py (本地渲染器, python-pptx)
   ▼
PPTX
   │  pptd_qa.py (本地 QA: 渲染→PDF→逐页图)
   ▼
视觉检查 (GLM-4.6V 或人工)
```

## 快速开始

```bash
# 1. 渲染 PPTD → PPTX
python3 pptd2pptx/pptd2pptx.py examples/facies/facies_experiment.pptd -o output.pptx --progress

# 2. 本地视觉 QA(生成逐页图供检查)
python3 pptd2pptx/pptd_qa.py examples/facies/facies_experiment.pptd --output-dir .qa-local

# 3. 预览 QA 图
#    .qa-local/page-1.jpg ... page-9.jpg + manifest.json
```

### 依赖

- Python 3.9+ : `python-pptx`、`Pillow`、`PyYAML`(渲染);`cairosvg`(图标,可选)
- LibreOffice `soffice` + poppler-utils `pdftoppm`(仅 QA 环节)
- Font Awesome 图标:`pptd2pptx/fetch_fa_icons.py` 自动下载到本地(开源 CC BY 4.0 / SIL OFL)

## 目录结构

```
gpt-ppt-make/
├── pptd2pptx/                  # 核心渲染器
│   ├── pptd2pptx.py            #   PPTD → PPTX 渲染器
│   ├── pptd_qa.py              #   本地视觉 QA(替代云端 export_images)
│   ├── fetch_fa_icons.py       #   Font Awesome 图标本地化
│   └── assets/fa-icons/        #   本地图标库(SVG)
└── examples/
    └── facies/                 # 完整示例:测井岩性相分类实验 PPT
        ├── facies_experiment.pptd
        ├── pages/              #   9 页 .page 文件
        ├── media/              #   图表 PNG + 对比图
        └── plot_charts.py      #   图表预生成脚本
```

## PPTD 格式支持范围

- **文本**:rich text(`<p>/<span>/<strong>/<em>/<u>/<s>/<sup>/<sub>/<a>`),样式继承链 inline > para > content > theme > 默认
- **形状**:40+ 种(rect/roundRect/ellipse/箭头/流程图/星形/云/心形等),solid/gradient 填充
- **表格**:合并单元格、主题样式、三线表(设计系统风格)
- **图片**:PNG/JPEG,contain 适配
- **图标**:Font Awesome 本地 SVG,任意 fill 色
- **主题**:colors/textStyles/tableStyles + `$token` 引用

> ⚠️ `chart` 元素不支持(已移除):图表一律用 matplotlib 预生成 PNG 后以 `image` 元素引用,保证质量稳定。

## 设计系统

内置学术答辩风格(teal-green-academic-defense)示例:纯白底、青绿点缀(彩色面积 ≤15%)、三线表、方点列表、留白 ≥2 角、页脚导航。示例 PPTD 的 theme 即该设计系统的落地。

## 与官方 Kimi 生态的关系

- **格式兼容**:PPTD 是 Moonshot 公开的 YAML 格式(与 Kimi 编辑器通用),本项目是独立的本地实现
- **官方通道**(可选):登录 Kimi 后仍可用官方导出作参考(本仓库不包含,也不依赖)
- **版权**:Font Awesome 图标为开源授权;示例数据为公开数据集(SEG 2016 相分类竞赛)

## License

MIT
