---
name: docx-zh-en-translation
description: Translate an existing Chinese Word document (.docx) into English while preserving layout, structure, tables, headers/footers, comments, and footnotes as much as possible. Use this skill whenever the user wants an English version of an existing Word file, asks to translate a Chinese contract/report/proposal/manual in Word format, wants to keep the original formatting, or wants a new translated .docx output rather than a rewritten document. Also use it for requests that mention preserving Word layout, tables, numbering, comments, or page furniture during translation. Do not use this for PDF translation, OCR, creating a new document from scratch, or editing charts/images/formulas.
compatibility:
  tools: Bash, Read, Write, Edit, Glob
---

# DOCX 中文转英文并保留排版

## Purpose

将现有中文 `.docx` 翻译为英文，并输出新文件 `原文件名_译文.docx`。

优先目标：
- 保留原有页面设置、段落结构、表格结构、样式、页眉页脚、批注、脚注/尾注
- 不重建文档
- 不修改图表、公式、图片内文字、目录域代码等高风险对象

## When to use this skill

当用户属于以下任一场景时应使用本 skill：
- 要把现有中文 Word / `.docx` 文档翻译成英文
- 要生成一个保留原排版的英文版 Word 文件
- 明确提到合同、报告、方案、手册、表格文档等已有 `.docx` 要翻译
- 明确要求尽量保留样式、编号、表格、页眉页脚、批注、脚注/尾注
- 虽然没有直接说“docx”，但上下文已经表明用户手里有一个现成的 Word 文档并想输出英文副本

常见触发表达包括但不限于：
- “帮我把这个 Word 文档翻成英文”
- “保留格式翻译这个 docx”
- “给我一个英文版合同，不要破坏原来的排版”
- “这个中文报告要出英文版，表格和编号尽量别乱”

## When not to use this skill

以下情况不要使用本 skill：
- 用户要从零新建一份英文文档，而不是基于现有 `.docx` 翻译
- 用户处理的是 PDF、扫描件、图片或需要 OCR 的材料
- 用户要改写内容、润色文风、总结提炼，而不只是保结构翻译
- 用户重点在图表文本、公式对象、图片内文字、SmartArt 等本 skill 明确不处理的内容

## Preconditions

开始前先确认：
- 输入文件确实是现有 `.docx`
- 运行环境可用 `python`
- 至少满足以下其一：
  - 已设置 `ANTHROPIC_API_KEY`
  - 当前机器上的 Claude Code CLI 已可用，可作为翻译回退路径

如果这些前提不满足，先向用户说明，再继续。

## Output contract

- 输入：一个现有 `.docx`
- 默认输出：同目录下新文件 `原文档名称_译文.docx`
- 也可显式指定输出路径
- 第一版覆盖：
  - 正文
  - 表格单元格
  - 页眉页脚
  - 批注
  - 脚注 / 尾注
- 第一版跳过：
  - 图表
  - 公式
  - 图片内文字
  - field codes / TOC 指令文本
  - 非普通 WordprocessingML 段落承载的 drawing payload text

交付目标是“生成可打开、结构基本保持、主要中文正文已变成英文的新 docx”，而不是逐像素保持分页完全一致。

## Workflow

### 1. Confirm scope and file type

先确认用户要处理的是已有 `.docx`，且需求是“保留 Word 结构的中译英”。如果用户需求更像 PDF 翻译、OCR、新建文档或内容改写，应改用别的方法。

### 2. Run the translator

运行：

```bash
python scripts/translate_docx.py <input.docx> [output.docx]
```

脚本会：
- 自动推断默认输出名为 `*_译文.docx`
- 解包 docx
- 处理以下 XML：
  - `word/document.xml`
  - `word/header*.xml`
  - `word/footer*.xml`
  - `word/comments.xml`
  - `word/footnotes.xml`
  - `word/endnotes.xml`
- 保守提取可见文本
- 跳过高风险节点
- 将中文翻译为英文
- 回写原 XML 结构
- 重新打包并校验输出文件

### 3. Handle failures conservatively

如果脚本失败：
- 先看是否缺少 `ANTHROPIC_API_KEY` 且 `claude` CLI 也不可用
- 再看输入是否是有效 `.docx`
- 再看 `validate` 失败信息和对应 XML part
- 不要在未定位问题时交付损坏文件

## Editing rules

- 保持原有 XML 结构，优先只替换 `<w:t>` 文本内容
- 不重排段落，不重建表格，不改样式定义
- 遇到 `w:instrText`、图表文本、绘图对象、公式对象时跳过
- 遇到复杂 mixed runs 时，优先保留 run 边界；如无法完全对齐，做最小化回写
- 不要为了“更自然的英文版式”而擅自改 margin、font、spacing、table width

## Known limitations

当前实现有以下已知限制，使用时应有预期：
- 优先使用 `ANTHROPIC_API_KEY`；未设置时会回退到本机 `claude` CLI
- 依赖缓存中的 docx office helper 脚本完成 unpack / pack / validate
- 英文膨胀可能导致换行、分页变化
- 高度精细的 run-level 样式不一定能完全保真
- 图表、公式、图片文字、目录域代码不会被翻译

## Verification checklist

交付前至少检查：
- 是否成功生成 `*_译文.docx`
- 输出文件是否能正常打开
- `validate.py` 是否通过
- 页眉页脚、表格、批注、脚注/尾注是否仍存在
- 图表和其他跳过对象是否未损坏
- 中文正文是否已实质翻译为英文

必要时把 docx 转成 PDF 做视觉比对。

## References

如需补充判断，读取：
- `references/xml-scope.md`
- `scripts/translate_docx.py`
- `scripts/office/validate.py`
