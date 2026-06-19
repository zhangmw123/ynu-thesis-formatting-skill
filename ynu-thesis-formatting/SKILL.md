---
name: ynu-thesis-formatting
description: "将Markdown或未排版Word草稿自动转换为符合中国高校硕士/博士毕业论文格式的Word文档。当用户需要论文排版、设置页边距和装订线、多级标题样式(1/1.1/1.1.1)、奇偶页页眉页脚、黑体/宋体中文排版、三线表、参考文献自动编号、LaTeX公式转换时，应使用此技能。默认基准为云南大学硕士论文格式，所有参数可修改适配其他高校。即使用户未明确说'论文格式化'，只要涉及中国学位论文Word生成或排版，也应触发。"
---

# 中国高校硕博论文 Word 格式化技能（云南大学基准）

## 概述

本技能将 Markdown 文本或未排版 Word 文档自动转换为符合中国高校硕博论文格式的 Word 文档（.docx）。
基于 `docx` (v9.x) 库纯 Node.js 实现，**无需 Pandoc 或 Microsoft Word**。

默认格式以**云南大学硕士论文**为基准，所有参数（校名、页边距、字体、字号、页码格式等）均可通过修改脚本中的常量来适配其他高校。

### 核心能力

1. **页面设置** - A4纸张，上下左右边距 + 装订线
2. **页眉** - 偶数页=学校名称，奇数页=章节标题（多 section 实现各章不同页眉）
3. **页码** - 前置部分（摘要/目录）小写罗马数字；正文从绪论起十进制编号，起始页码=1
4. **多级标题** - `#`→一级（第X章），`##`→二级（1.1），`###`/`####`→三级（1.1.1）
5. **字体样式** - 黑体（标题）/ 宋体（正文），小三号/四号/小四号
6. **正文样式** - 首行缩进两字符（420 DXA），1.5倍行距
7. **三线表** - 学术标准三线表（上下粗线、表头下细线），表题五号黑体居中于表格上方
8. **图片标注** - 图号图题，五号黑体居中于图片下方
9. **公式支持** - LaTeX `$...$` / `$$...$$` → Unicode 纯文本 + Cambria Math 斜体（非 OMML，避免乱码）
10. **参考文献** - Word 自动编号 `[1][2]...`，442 DXA 悬挂缩进，仅在 `# 参考文献` 区域生效
11. **自动目录** - 英文摘要关键词后自动插入 TOC 域
12. **Markdown 列表** - `* item` / `- item` 去掉前缀作为正文段落
13. **标点优化** - 中文段落及表格中英文双引号 `""` 自动转中文双引号 `""`
14. **英文摘要** - Abstract 页面正文和关键字字体统一强制为 Times New Roman

### 适用场景

- 用户说"帮我写一篇关于XXX的论文，用XXX格式"
- 用户说"把这段文字按论文格式排版"
- 用户上传 Markdown 想要转成符合学校格式的 docx
- 用户上传未排版 `.docx`，希望重新转换为论文格式 `.docx`

---

## 适配你的学校

本 Skill 默认以**云南大学硕士论文格式**为基准。如果你的学校格式要求不同，编辑 `scripts/generate_thesis.js` 中的常量即可。详见代码中的 `格式常量` 区块（约第 80-130 行）和 README.md 中的 "适配你的学校" 章节。

可修改的内容包括：
- 学校名称、学科、学位级别（页眉文字）
- 页边距、装订线、页眉页脚距离
- 标题和正文字体（默认黑体+宋体）
- 字号、行距、首行缩进
- 页码格式（罗马数字 vs 十进制）
- 参考文献编号格式

---

## 使用方法

### 作为 Skill（Claude Code）

在对话中直接说明需求：

> "把 `draft/chapter3.md` 按硕士论文格式排版，输出到 `output/chapter3.docx`"

### 命令行

```bash
node ynu-thesis-formatting/scripts/generate_thesis.js input.md output.docx
node ynu-thesis-formatting/scripts/generate_thesis.js input.docx output.docx
```

### 编程调用

```javascript
const { generateThesisDocx, parseMarkdown, extractMarkdownFromDocx } =
  require("./ynu-thesis-formatting/scripts/generate_thesis");

const buffer = await generateThesisDocx(markdownString, "output.docx", {
  schoolName: "云南大学",
  discipline: "专业",
  degreeType: "硕士",
});
```

---

## 格式标准

详见 README.md 中 "格式标准说明" 章节。关键参数：

- 纸张：A4，上/下/左/右边距 2.5/2.0/3.0/2.0 cm
- 一级标题：黑体 15pt 居中，二级：黑体 14pt 左对齐，三级：黑体 12pt 左对齐缩进
- 正文：宋体 12pt，首行缩进两字符，1.5 倍行距
- 页码：前置 i/ii/iii，正文 1/2/3 从绪论第 1 页起
- 参考文献：Word 自动编号 `[%1]`，悬挂缩进 442 DXA

---

## 支持的 Markdown 语法

`#`~`####` 标题 | `**粗体**` / `*斜体*` | `* item` / `- item` 列表 | `` `代码` `` | `$公式$` / `$$公式$$` | `| 表格 |` | `![图](path)` | `[1]` 参考文献

---

## 生成脚本

脚本位于 `scripts/generate_thesis.js`。回归测试：

```bash
node ynu-thesis-formatting/scripts/format_regression_test.js
node ynu-thesis-formatting/scripts/docx_input_regression_test.js
```

---

## 限制与已知问题

### 公式
1. **不使用 OMML**：公式为 Unicode 纯文本 + Cambria Math 斜体，非 Word 原生公式对象。temml→XSLT→OMML 管线在 Word 中文版中会乱码，目前无 Pandoc 可用。
2. **分式线性化**：`\frac{a}{b}` → `(a)/(b)`，上下标保留 `_{...}` `^{...}` 语法。
3. **DOCX 输入**：原 `.docx` 中的 OMML 公式会提取文本重新渲染。纯文本公式需手动加 `$$...$$` 包裹才能被识别为公式段落。

### 参考文献
4. **正文引用不是交叉引用**：`[1]` 为上标纯文本，非可点击的 Word 交叉引用域。
5. **识别范围**：仅 `# 参考文献` 标题之后的内容为参考文献，正文中的编号不误判。

### 其他
6. **图片**：需有效路径，无效则显示占位符。
7. **不支持 .doc** 格式（旧版 Word 97-2003）。
