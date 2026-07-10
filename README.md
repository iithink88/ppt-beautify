# PPT 美化技能 (ppt-beautify)

一套用于 **WorkBuddy** 的 PPT 一键美化技能，基于 `python-pptx` **原地换肤**：把「白底堆图、配色杂乱、红字刺眼」的 `.pptx`，改成视觉统一、专业大气的演示文稿。

> 核心优势：**保留原稿全部内容**（图片 / 视频 / 文字 / 版式坐标），只统一视觉皮肤 —— 不丢图、不丢字、不溢出、零坐标误差。

## 两个版本

| 目录 | 版本 | 说明 |
|------|------|------|
| [`ppt-beautify-simple/`](./ppt-beautify-simple) | **简单版** | 纯「原地换肤」：统一背景、卡片、文字配色，最稳最安全 |
| [`ppt-beautify-big-cover/`](./ppt-beautify-big-cover) | **大封面版** | 在换肤基础上多一个 `--cover hero` 模式：封面单张大图铺满 + 半透明蒙版 + 标题置顶，电影海报感 |

三套内置主题：`blue-gold`（深蓝金）/ `red-gold`（中国红）/ `teal`（科技青）。

## 效果一览

- 背景：统一主题深底色（如深海军蓝 `0E2A47`）
- 文字：标题白、正文浅蓝、点缀金；红/橙刺眼色自动收编
- 卡片/形状：统一主题卡片色 + 金色细边
- 图片：统一加主题色细边
- 可选：页脚 + 页码；封面大图（大封面版）

## 朋友怎么安装这个技能

任选一种：

1. **拖文件**：下载对应版本的 `SKILL.md`，直接拖进 WorkBuddy 聊天框。
2. **放目录**：把 `ppt-beautify-simple/` 或 `ppt-beautify-big-cover/` 整个文件夹，放进你的 `~/.workbuddy/skills/`（用户级，所有项目可用）。
3. **命令行**：`npx skills add iithink88/ppt-beautify@ppt-beautify-simple`（或 `@ppt-beautify-big-cover`）。

安装后首次使用前，装一次依赖：

```bash
python -m venv .venv                # Windows: py -3 -m venv .venv
.venv/Scripts/pip install python-pptx   # macOS/Linux: .venv/bin/pip install python-pptx
```

## 用法

```bash
# 基础换肤
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold

# 加页脚页码
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold --footer --footer-text "演示文档"

# 封面大图（仅大封面版）
python scripts/beautify.py 输入.pptx 输出.pptx --theme blue-gold --cover hero --footer --footer-text "演示文档"
```

更多细节见各版本目录下的 `SKILL.md`。

## License

MIT
