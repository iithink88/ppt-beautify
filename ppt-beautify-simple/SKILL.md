---
name: ppt-beautify
description: PPT 原地换肤美化 —— 把"白底堆图、配色杂乱"的 .pptx 一键改成统一主题（深蓝金 / 中国红 / 科技青）。当用户说"把这个PPT改漂亮/美化一下/统一配色/换个风格/做成主旋律风/做得大气点"时使用。核心是用 python-pptx 原地遍历重着色、递归处理分组，绝不重建幻灯片，零坐标误差，保证不丢图、不丢字、不溢出。
---

# PPT 原地换肤美化 (In-place Re-skin)

把一份配色杂乱、白底堆图、红字刺眼的 PPT，改成视觉统一、专业大气的演示文稿。
**最大优势：保留原稿全部内容（图片 / 视频 / 文字 / 版式坐标），只统一视觉皮肤**，因此不会丢内容、不会错位。

## ⚠️ 黄金法则（踩过坑才有的经验）

1. **绝不重建幻灯片。** 不要用 `add_group_shape()` / 复制元素再重排。原稿大量使用 GROUP（组合）时，python-pptx 复制子元素会出现坐标系转换错误 → 文字全飞出画面、图片丢失。
   ✅ 正确做法：**打开原文件，遍历每个 shape 原地改色/字体，不复制、不移动任何元素** = 零坐标风险。
2. **必须递归进入 GROUP。** `slide.shapes` 只遍历**顶层**形状。GROUP 内部的文字和卡片要用 `for child in group_shape.shapes` 递归进去处理，否则分组里的红字改不掉。
3. **背景用 `slide.background.fill.solid()`。** 不要依赖"找到一个全屏矩形去改色"——很多原稿背景是 slide background 而非形状，直接改背景填充最稳。

## 环境准备（Windows）

```powershell
# 1) 建隔离 venv（注意 Windows 用 Scripts/，不是 bin/）
python -m venv .venv   # Windows 用 py -3 -m venv .venv
.venv\Scripts\pip install -i https://pypi.tuna.tsinghua.edu.cn/simple python-pptx   # macOS/Linux: .venv/bin/pip
```

> 本机没装 LibreOffice，且 pptx 技能自带的 `scripts/office/soffice.py` 仅支持 Linux 沙箱（Windows 会 `AF_UNIX` 报错）。**预览/QA 走 PowerPoint 或 WPS 演示的 COM 导出**，见下方"视觉 QA"。

## 用法

```powershell
# 美化（默认深蓝金主题）
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold

# 可选主题：red-gold（主旋律中国红）/ teal（科技青）
python scripts/beautify.py 输入.pptx 输出.pptx --theme red-gold

# 加页脚与页码（默认不加，避免压到现有内容；确认不重叠再开）
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold --footer --footer-text "演示文档"
```

`beautify.py` 内部做了什么：
- 每页 `slide.background.fill` → 主题深底
- 递归遍历所有形状：图片加主题色细边；所有纯色形状（卡片/文本框）→ 主题卡片色 + 金/强调细边；全屏形状当背景处理不染色
- 递归遍历所有文字 run（含 GROUP 内、表格单元格）：红/橙等刺眼色 → 收为金色强调；标题（粗体/大字号）→ 白；正文 → 浅蓝；空格 run 不动；零坐标改动

## 视觉 QA（无多模态时也能验证）

本模型读不了图（Read 图片会被过滤）。用两种方式交叉验证：

**A. PowerPoint / WPS COM 导出每页真实截图**（本机已装其一即可，无需管理员）：
```powershell
powershell -File scripts/export_preview.ps1 输出.pptx 预览目录
```
脚本会调 `PowerPoint.Application`（WPS 用 `kwpp.Application`）把每页导成 JPG，你（或用户）可肉眼看效果。

**B. 程序化颜色审计**（模型自己能跑，确认换肤成功、无刺眼残留）：
```powershell
python scripts/qa_color_audit.py 输出.pptx
```
它会递归扫描所有 run 文字色 + 形状填充色，报告每页背景色、文字色分布、卡片填充色，并标出是否还有红/橙残留。

**C. 内容完整性**（确认没丢字）：
```powershell
python -m markitdown 输出.pptx > 新文本.md
# 对比原稿，关键短语应都在
```

## 可调项 / 常见微调

- **封面想用单张大剧照铺满**：`beautify.py` 暂不改版式，可手动在 PowerPoint 里把封面剧照拉满，再叠标题。
- **换配色**：改 `THEMES` 字典即可，或加新主题键。
- **某页想加标题竖条/页眉**：当前默认不加（防重叠）；需要时给那页单独加文本框。
- **页脚压到内容**：关掉 `--footer`，或调 `beautify.py` 里页脚 y 坐标。

## 真实案例

《超级台风》观影导读（6 页，含 13 张剧照 + 1 个嵌入视频）：v1 解包重组导致文字飞出/剧照丢失；改为原地换肤 + 递归分组后 v4 成功。6 页背景全变深蓝 `0E2A47`，红字收为金色，卡片染蓝，零内容丢失，PowerPoint 实拍 + 颜色审计双重确认。

## 安装 / 分享给朋友

技能就是一个文件夹，复制即可用，无需注册：

1. 把 `ppt-beautify/` 整个文件夹复制到：
   - 用户级：`~/.workbuddy/skills/ppt-beautify/`（所有项目可用）
   - 或项目级：`<你的项目>/.workbuddy/skills/ppt-beautify/`（仅该项目）
2. 首次使用需装 python-pptx（见"环境准备"的 venv 命令）。
3. 之后在对话里说"把这份 PPT 美化一下"，WorkBuddy 会自动加载本技能。

> 想打包发送：把文件夹压缩成 `ppt-beautify.zip` 发给朋友，解压到上述目录即可。
