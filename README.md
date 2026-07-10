# PPT 美化技能合集（WorkBuddy）

一键给 PPT「换肤」美化：**保留原版式，只统一视觉风格**，不丢图、不丢字、不溢出。

本仓库含两个版本的 WorkBuddy 技能（Skill）：

| 目录 | 版本 | 说明 |
|------|------|------|
| `ppt-beautify-simple/` | 简单版 | 纯原地换肤，最稳，适合绝大多数 PPT |
| `ppt-beautify-big-cover/` | 大封面版 | 在简单版基础上，新增 `--cover hero` 电影海报式封面（单张大图铺满） |

> 两份技能均已实测可用：深蓝金 / 中国红 / 科技青三主题全跑通，颜色审计零刺眼残留、图片无损。

---

## 效果一览

| 维度 | 美化前 | 美化后 |
|------|--------|--------|
| 背景 | 纯白底 + 堆图 | 主题深底（深蓝 / 中国红 / 科技青） |
| 配色 | 无统一色系 | 主题色 + 金色强调（党政/主旋律气质） |
| 文字 | 黑色普通字、偶有刺眼红橙 | 标题白、正文浅蓝、强调金，零刺眼色 |
| 卡片 | 浅色/杂色块 | 统一主题卡片色 + 金色细边 |
| 图片 | 无边框 | 全部加主题色细边 |

## 三个内置主题

| 主题键 | 风格 | 适用 |
|--------|------|------|
| `blue-gold` | 深海军蓝 + 金 | 庄重、政务、教育（默认） |
| `red-gold` | 中国红 + 金 | 主旋律、节庆、党建 |
| `teal` | 科技青 + 浅蓝 | 科技、互联网、清爽风 |

---

## 安装（三种方式任选）

**方式 A：拖进技能目录（推荐给朋友）**
把 `ppt-beautify-simple/` 或 `ppt-beautify-big-cover/` 文件夹，整体丢进你的
`~/.workbuddy/skills/` 目录（Windows 为 `C:\Users\你的用户名\.workbuddy\skills\`）。
WorkBuddy 会自动发现，对话里说「把这份 PPT 美化一下」就会加载。

**方式 B：从聊天加载**
把对应技能目录里的 `SKILL.md` 直接拖进 WorkBuddy 对话框，按提示安装。

**方式 C：命令行**
```bash
npx skills add iithink88/ppt-beautify@ppt-beautify-simple
# 或
npx skills add iithink88/ppt-beautify@ppt-beautify-big-cover
```

### 首次使用需安装依赖（一次性）
技能脚本用 Python + `python-pptx`。首次使用前建个虚拟环境装好依赖：
```bash
python -m venv .venv          # Windows 用  py -3 -m venv .venv
.venv/Scripts/pip install -i https://pypi.tuna.tsinghua.edu.cn/simple python-pptx
# macOS / Linux 用  .venv/bin/pip install python-pptx
```

---

## 快速上手

```bash
# 简单版：深蓝金主题 + 页脚
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold --footer --footer-text "演示文档"

# 大封面版：电影海报封面 + 深蓝金主题
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold --cover hero --footer --footer-text "电影《超强台风》观影导读"
```

封面不在第 1 页时，用 `--cover-slide 3` 指定页码。

---

## 简单版说明

把这份踩坑换来的方法做成了标准 Skill，已实测可用
（深蓝金 / 中国红 / 科技青三主题全跑通，颜色审计零刺眼残留、图片无损）。

**技能目录结构**
```
ppt-beautify/
├── SKILL.md              # 用法、黄金法则、QA 方法、踩坑记录
└── scripts/
    ├── beautify.py       # 原地换肤核心（递归分组 + 表格 + XML 兜底）
    ├── qa_color_audit.py # 程序化颜色审计（无多模态也能验收）
    └── export_preview.ps1 # 调 PowerPoint / WPS 导出每页截图做视觉 QA
```

**朋友怎么用**
拿到技能文件夹后，丢进自己的 `~/.workbuddy/skills/` 目录，
首次用前按上面的步骤装一下 `python-pptx`（一行 venv 命令），
之后对话里说「把这份 PPT 美化一下」就会自动加载。

**设计说明**
这个技能的「换肤」是**保留原版式、只统一视觉**（不丢图、不丢字、不溢出），所以最稳。
若想要更大幅度的版式重排（如封面大图铺满），请用大封面版或手动在 PowerPoint 里调——这点已在 `SKILL.md` 中写明。

---

## 大封面版说明（新增 `--cover hero`）

在原有「原地换肤」基础上，封面可一键改成**电影海报感**：
单张最大剧照铺满全屏 + 半透明深色蒙版 + 标题置顶。

```bash
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold --cover hero --footer --footer-text "电影《超强台风》观影导读"
```
封面不在第 1 页时用 `--cover-slide 3` 指定。

**实现要点（都是踩坑换来的）**
- **递归找最大图**：封面剧照常藏在 GROUP 组合里，只删顶层图片会留一堆图盖住大图 → 必须递归收集、整组删除。
- **清掉原稿全屏块**：原稿若有不透明背景块会压住大图，一并删除。
- **用原图数据重铺**：取 `hero.image.blob` 经 `add_picture` 重新铺一张全屏图（比缩放原图更稳，不变坐标）。
- **蒙版透明度**：直接写 XML `a:srgbClr/a:alpha`（55%）实现半透明，保证标题可读。
- **标题置顶**：`bring_to_front` 把文字压在蒙版之上。

**验证结果**
- 封面结构：仅 1 张大图 + 1 个蒙版 + 标题，分组 / 拼贴全清（审计 slide1 图片数 = 1、GROUP = 0）。
- 其余页面不受影响，全深蓝主题、无刺眼残留。

**一点提醒**
- `hero` 模式会拉伸最大图铺满，若有轻微变形属正常。
- 若封面标题原本嵌在分组里（少见），可能被一并清掉，需手动补——这点已在 `SKILL.md` 中写明。
- 进阶可扩展：蒙版渐变（下深上浅）、大图居中裁切不变形。

---

## 视觉验收（QA）

没有多模态看图能力时，可用脚本做程序化验收：
```bash
python scripts/qa_color_audit.py 输出.pptx
```
它会输出每页背景色、文字色分布，并标出是否还有刺眼红 / 橙色残留、图片总数，确保换肤真正生效、内容未丢。

也可调用 `export_preview.ps1` 让 PowerPoint / WPS 把每页导出成图片，肉眼核对实际效果。

---

## 黄金法则（来自踩坑经验）

1. **原地换肤，不重建**：直接在原 PPT 上改颜色和字体，不复制 / 重组元素，坐标零误差。
2. **必须递归进 GROUP**：`slide.shapes` 只遍历顶层，GROUP 内部图片和文字需递归进去才能统一。
3. **表格单元格单独处理**：`shape.has_table` 的文字不通过 `has_text_frame` 暴露，要专门迭代单元格。
4. **继承 / 方案色要用 XML 兜底**：python-pptx 对「继承色」直接赋值可能静默失败，必要时直接写 `a:srgbClr`。
5. **空格 run 跳过**：只含空格的隐藏 run 不影响视觉，跳过避免误伤。
6. **Windows 路径坑**：python.exe 传参用 `C:/` 风格，而非 git bash 的 `/c/`。

---

## 许可证

MIT —— 自由使用、修改、分享。
