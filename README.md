# docx-zh-en-translation

> 将现有中文 `.docx` 文档翻译为英文，并尽可能保留原有版式与结构。

## ✨ 项目简介

这是一个面向 **中文 Word 文档转英文** 场景的 DOCX 翻译工具，核心目标不是“重新生成一份新文档”，而是：

- **基于原始 `.docx` 直接处理**
- **尽量保留 Word 原有排版与结构**
- **输出新的英文版文档**
- **优先保证文档可打开、可用、结构稳定**

它特别适合以下类型的文档：

- 合同
- 报告
- 方案 / 标书
- 手册
- 含表格、编号、页眉页脚、批注、脚注的正式 Word 文档

---

## 🎯 设计目标

本项目遵循一个明确原则：

> **文档结构安全优先于极限翻译覆盖率。**

也就是说，与其冒险修改高风险 Office XML 对象导致 Word 文件损坏，不如只处理那些**可安全翻译、可稳定回写**的文本区域。

因此，这个工具更关注：

- 文档能正常打开
- 表格、段落、页眉页脚等结构尽量不被破坏
- 英文输出尽量自然、专业
- 避免因激进改写 XML 导致格式错乱或文件损坏

---

## ✅ 主要能力

- 保留原始 `.docx` 结构并生成新的英文版文件
- 翻译正文段落文字
- 翻译表格中的文本内容
- 翻译页眉 / 页脚
- 翻译批注（comments）
- 翻译脚注 / 尾注
- 使用保守策略处理 Word Open XML，避免不必要的结构重建
- 回写后重新打包并校验输出文档
- 默认输出文件名为：`原文件名_en.docx`

---

## ⚠️ 当前有意跳过的内容

以下内容**默认不处理**，以降低破坏文档结构的风险：

- 图表（charts）
- 公式 / 数学对象
- 图片中的文字
- 域代码（如 `w:instrText` / 目录 / 域指令）
- 高风险 drawing payload 文本
- 其他不属于普通 WordprocessingML 段落文本的复杂对象

这不是缺陷，而是本项目的刻意设计：

> **宁可少翻一点，也不要把文档搞坏。**

---

## 🧱 工作原理

整体流程如下：

1. **校验输入文件**
   - 检查文件是否存在
   - 检查扩展名是否为 `.docx`

2. **解包 DOCX**
   - DOCX 本质上是一个 ZIP 包
   - 工具会调用 Office helper 脚本将其解包为可编辑 XML

3. **筛选可安全处理的 XML 部分**
   当前会处理：
   - `word/document.xml`
   - `word/styles.xml`
   - `word/fontTable.xml`
   - `word/comments.xml`
   - `word/footnotes.xml`
   - `word/endnotes.xml`
   - `word/header*.xml`
   - `word/footer*.xml`

4. **提取可见文本**
   - 只处理安全、可见的文字节点
   - 跳过删除文本、域代码等高风险内容

5. **批量翻译中文内容**
   - 按段落收集内容
   - 批量送入翻译后端

6. **回写到原始 XML 结构中**
   - 优先只替换 `<w:t>` 文本节点
   - 尽量不破坏原有段落 / run / 表格结构

7. **重新打包并校验**
   - 将修改后的 XML 重新封装为新的 `.docx`
   - 执行校验流程，降低生成损坏文档的概率

---

## 🗂 仓库结构

```text
.
├── README.md
├── CLAUDE.md
├── SKILL.md
├── evals/
│   └── evals.json
├── references/
│   └── xml-scope.md
└── scripts/
    └── translate_docx.py
```

说明：

- `scripts/translate_docx.py`：主脚本，核心逻辑都在这里
- `references/xml-scope.md`：说明哪些 XML / 内容会处理，哪些会跳过
- `evals/evals.json`：评估样例与验收预期
- `SKILL.md`：面向 Claude skill 的说明文档
- `CLAUDE.md`：面向 Claude Code 的仓库指导文件

---

## 🖥 运行要求

### Python

- Python **3.11 及以上**

### 翻译后端

当前实现支持以下两类翻译路径：

#### 方案 A：可配置的兼容接口（默认优先）
脚本通过环境变量读取翻译接口配置，设计上并不绑定某一个固定模型或服务商。只要你的服务能够提供**兼容 Chat Completions 风格**的接口，就可以接入。

当前代码使用以下环境变量：

- `LONGCAT_API_BASE`：兼容接口地址
- `LONGCAT_API_KEY`：接口访问密钥
- `LONGCAT_MODEL`：调用时使用的模型名称

虽然变量名目前保留了 `LONGCAT_` 前缀，但它们本质上只是**当前实现中的配置键名**，并不意味着你只能使用某一个特定服务。你完全可以将其指向其他兼容服务或自定义模型。

#### 方案 B：本地 Claude CLI 回退
如果接口调用失败，且本机 `claude` 命令可用，脚本会自动尝试回退到本地 Claude CLI。

### Office helper 脚本

脚本依赖 Office helper 完成以下步骤：

- unpack
- pack
- validate

当前代码会尝试在以下位置寻找这些 helper：

- `scripts` 附近
- `~/.claude/plugins/cache/anthropic-agent-skills/` 下的 skill 缓存目录

---

## ⚙️ 配置方式

根据当前代码实现，可通过环境变量配置翻译接口。

需要说明的是：

- 当前脚本读取的是 `LONGCAT_API_BASE / LONGCAT_API_KEY / LONGCAT_MODEL`
- 这些名称是**当前实现中的环境变量名**，不是对具体服务商的强绑定
- 只要你的后端提供兼容接口，就可以复用这组配置项

### Windows CMD

```bat
set LONGCAT_API_BASE=https://your-compatible-endpoint
set LONGCAT_API_KEY=your_api_key
set LONGCAT_MODEL=your_model_name
```

### PowerShell

```powershell
$env:LONGCAT_API_BASE="https://your-compatible-endpoint"
$env:LONGCAT_API_KEY="your_api_key"
$env:LONGCAT_MODEL="your_model_name"
```

例如，你可以把它理解为：

- `LONGCAT_API_BASE`：你的兼容接口地址
- `LONGCAT_API_KEY`：你的接口访问凭证
- `LONGCAT_MODEL`：你希望调用的模型名称

如果未正确配置接口，但本机安装并可使用 `claude` CLI，脚本会尝试使用 CLI 作为回退路径。

---

## 🚀 使用方法

### 基本命令

```bash
python scripts/translate_docx.py <input.docx> [output.docx]
```

### 示例

```bash
python scripts/translate_docx.py 中文合同.docx
python scripts/translate_docx.py 中文项目报告.docx 中文项目报告_en.docx
```

### 输出命名规则

- 如果**未指定输出路径**，默认输出：`<原文件名>_en.docx`
- 输出文件默认写到与源文件相同的目录

---

## 🔍 当前覆盖范围

### 已覆盖

- 正文段落文本
- 表格单元格文本
- 页眉 / 页脚
- 批注
- 脚注 / 尾注

### 明确跳过

- 图表
- 公式
- 图片内文字
- 域代码
- 复杂绘图承载文本

---

## 🧠 实现特性

### 1. 保守 XML 回写策略

本项目不会重建整份 Word 文档，而是尽量：

- 保持原始 XML 结构
- 只修改必要的文本节点
- 尽量保持段落和表格不变

### 2. 批量翻译

脚本会将需要翻译的中文文本按批次发送到后端，而不是逐句单独请求，以提升效率并减少重复开销。

### 3. 按段落处理

处理单位以段落为主，再尽量将英文结果重新分配回原有文本节点中，以减少结构性破坏。

### 4. 兼容性修复

XML 写回后，脚本还会修复必要的 namespace / prefix 信息，以提高 Word 打开兼容性。

---

## 🧪 校验与验证

### 语法检查

```bash
python -m py_compile scripts/translate_docx.py
```

### 运行一次真实翻译验证

```bash
python scripts/translate_docx.py 中文合同.docx
```

预期默认输出：

```text
中文合同_en.docx
```

### 人工检查清单

生成文件后，建议至少检查以下内容：

- [ ] 输出文件能否正常被 Word 打开
- [ ] 中文正文是否已翻译为英文
- [ ] 表格是否仍然完整
- [ ] 页眉 / 页脚是否保留
- [ ] 批注、脚注、尾注是否保留
- [ ] 图表、公式等未处理对象是否未被破坏
- [ ] 整体版式是否基本可接受

---

## 📌 已知限制

- 英文通常比中文更长，可能导致换行、分页变化
- 高度精细的 run-level 样式不一定能 100% 保真
- 本项目不做 OCR，因此不会翻译图片中的文字
- 不处理图表、公式、域代码等高风险对象
- 当前仓库未提供完整单元测试体系，验证以真实 DOCX 运行和人工检查为主

---

## ❓常见问题

### 为什么不尝试翻译 DOCX 里的所有内容？

因为很多 Office 对象在 XML 层面非常脆弱。盲目修改虽然可能“覆盖更多内容”，但也更容易导致：

- 文档打不开
- 排版损坏
- 域结构失效
- 图表 / 公式 / 复杂对象异常

所以本项目选择更稳妥的策略。

### 为什么默认输出是 `_en.docx`？

这样更简洁、稳定，也更方便在终端、脚本和跨环境处理中使用。

### 以后可以替换翻译后端吗？

可以。当前实现本身就是后端可配置的，只要接口协议兼容即可替换。

---

## 🔐 安全提示

- 不要提交真实 API Key
- 不要在未经许可的情况下提交客户文档
- 涉及合同、内部资料、敏感文本时，请在受控环境中运行

---

## 📚 参考文件

- `scripts/translate_docx.py`
- `references/xml-scope.md`
- `SKILL.md`
- `evals/evals.json`

---

## 🏁 总结

这是一个强调 **稳定性、结构保真、可交付性** 的 DOCX 翻译工具。

如果你的目标是：

> **把现有中文 Word 文档翻成英文，同时尽量保留原来的排版、表格、页眉页脚、批注和注释结构**

那么这个项目就是为这个场景设计的。