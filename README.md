# ynu-thesis-formatting

> 一个 Claude Code Skill——把 Markdown / Word 论文草稿一键转为符合中国高校硕士毕业论文格式的 Word 文档。仅针对毕业论文的格式调整，无写作功能。

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

**无需 Pandoc，无需安装 Microsoft Word**。底层基于 `docx` (v9.x) 库纯 Node.js 实现。若需进一步优化公式的问题可以使用pandoc进行嵌入。

默认基准：**云南大学硕士论文格式**。所有参数（校名、页边距、字体、字号等）均可修改，适配你自己的学校。

---

## 目录

- [安装](#安装)
- [使用](#使用)
- [它能做什么](#它能做什么)
- [配置与自定义](#配置与自定义)
- [格式标准](#格式标准)
- [已知限制](#已知限制)
- [适配你的学校](#适配你的学校)
- [文件结构](#文件结构)
- [License](#license)

---

## 安装

### 方式一：npm 全局安装（命令行用）

```bash
# 1. 从 GitHub 安装
npm install -g github:zhangmw123/ynu-thesis-formatting-skill

# 2. 安装后即可在任意目录使用
cd 我的论文目录
ynu-thesis-format 论文草稿.md 论文输出.docx
```

如果 npm 全局安装不可用，也可以本地安装：

```bash
git clone https://github.com/zhangmw123/ynu-thesis-formatting-skill.git
cd ynu-thesis-formatting-skill
npm install
node ynu-thesis-formatting/scripts/generate_thesis.js 论文草稿.md 论文输出.docx
```

### 方式二：给 Claude Code 安装（AI 对话中直接调用）

把仓库地址给 Claude Code，它会自己配置好：

```bash
# 在终端执行
claude skills install https://github.com/zhangmw123/ynu-thesis-formatting-skill.git
```

安装完就可以在 Claude Code 里直接说：

> "帮我把 `draft/论文草稿.md` 按硕士论文格式排版，输出为 `output/论文格式版.docx`"

Skill 会自行调用格式化脚本来完成。

### 依赖

| 包 | 用途 |
|---|------|
| `docx` (^9.x) | 生成 Word .docx 文件 |
| `jszip` (^3.x) | 读取 .docx 输入文件 |
| `@xmldom/xmldom` | 回归测试用（仅开发依赖） |

---

## 使用

### 命令行

```bash
# Markdown 草稿 → 格式化 Word（推荐）
node ynu-thesis-formatting/scripts/generate_thesis.js 我的论文.md 我的论文_格式版.docx

# 已有 Word 草稿 → 重新排版
node ynu-thesis-formatting/scripts/generate_thesis.js 原始草稿.docx 格式版.docx
```

### 在 Claude Code 对话中调用

安装 Skill 后，直接在对话中说类似的话即可触发：

- "调用 ynu-thesis-formatting 对我的论文草稿 `draft/第三章.md` 进行格式调整"
- "帮我把 `毕业论文草稿.docx` 按硕士论文格式重新排版"
- "把这个 markdown 转成学校要求的 Word 格式"

### 编程调用

```javascript
const { generateThesisDocx } = require("ynu-thesis-formatting");

// Markdown 字符串 → DOCX
await generateThesisDocx(markdownContent, "output.docx", {
  schoolName: "云南大学",
  discipline: "工学",
  degreeType: "硕士",
});
```

---

## 它能做什么

| 功能 | 说明 |
|------|------|
| 📄 **页面设置** | A4，上/下/左/右边距 2.5/2.0/3.0/2.0 cm |
| 📑 **页码** | 前置部分用罗马数字（i, ii, iii），正文从第 1 页起十进制编号 |
| 📰 **页眉** | 偶数页显示校名，奇数页显示章节标题，首页无页眉 |
| 🏷️ **多级标题** | `#` → "第一章"，`##` → "1.1"，`###`/`####` → "1.1.1" |
| 🔤 **字体** | 标题=黑体，正文=宋体+Times New Roman（中西文自动分字体） |
| 📝 **正文** | 首行缩进两字符，1.5 倍行距 |
| 📊 **三线表** | 学术标准三线表（上下粗线、表头下细线），表题居中于表格上方 |
| 🖼️ **图片** | `![图1 标题](path)` → 图号图题居中于图片下方 |
| 🧮 **公式** | `$...$`/`$$...$$` → Unicode + Cambria Math 斜体（不依赖 OMML） |
| 📚 **参考文献** | Word 自动编号 `[1]` `[2]`...，悬挂缩进 |
| 📖 **自动目录** | 英文摘要后自动插入 TOC 域 |
| ✏️ **标点修正** | 中文段落 `""` → `""`，表格同理 |
| 🔤 **英文摘要** | Abstract 页统一 Times New Roman |

---

## 配置与自定义

### 改校名（最简单）

编辑 `scripts/generate_thesis.js`，搜索 `schoolName`：

```javascript
const opts = {
  schoolName: '你的学校',   // ← 改这里
  discipline: '工学',       // ← 改这里
  degreeType: '硕士',       // ← 改这里（或'博士'）
};
```

### 改页边距

```javascript
const MARGINS = {
  top: Math.round(2.5 * 567),      // 上边距 (cm × 567 = DXA)
  bottom: Math.round(2.0 * 567),
  left: Math.round(3.0 * 567),
  right: Math.round(2.0 * 567),
};
```

### 改字体

```javascript
const CJK_FONT = {
  heading: '黑体',   // ← 改成 '方正小标宋简体' 等
  body: '宋体',
};
```

---

## 格式标准

### 标题层级

| 级别 | Markdown | 编号 | 字体 | 字号 | 对齐 |
|------|----------|------|------|------|------|
| 题目 | 首行 | 无 | 黑体 | 18pt | 居中 |
| 一级 | `#` | "第X章" | 黑体 | 15pt | 居中 |
| 二级 | `##` | "X.X" | 黑体 | 14pt | 左对齐 |
| 三级 | `###`/`####` | "X.X.X" | 黑体 | 12pt | 左对齐缩进 |

### 正文

- 宋体 + Times New Roman，12pt（小四）
- 首行缩进两字符（420 DXA）
- 1.5 倍行距
- 两端对齐

### 参考文献

- 编号：Word 自动 `[1]` `[2]`...（decimal，从 1 开始）
- 悬挂缩进：~0.78 cm
- **仅识别 `# 参考文献` 标题之后的内容**，正文编号列表不会误判

### 公式处理

LaTeX → Unicode 纯文本 + Cambria Math 斜体：

| 输入 | 输出 |
|------|------|
| `$\theta$` | **θ**（Cambria Math 斜体） |
| `$\sum x_i$` | **∑ x_i** |
| `$$\frac{a+b}{c}$$` | **(a+b)/(c)** 居中 + 编号 |
| `\begin{cases}...\end{cases}` | 无边框表格 |

---

## 已知限制

| 限制 | 说明 |
|------|------|
| 公式不是 OMML | 以 Unicode 纯文本渲染，非 Word 原生公式对象（temml→OMML 在中文 Word 中乱码）。分式显示为 `(a)/(b)` |
| DOCX 中纯文本公式不自动识别 | 需手动在公式前后加 `$$` |
| 引用不可点击 | 正文 `[1]` 是上标文本，不是 Word 交叉引用域 |
| 不支持 `.doc` | 仅支持 `.docx`（Office 2007+） |
| 图片需有效路径 | 路径无效显示占位符 |

---

## 适配你的学校

1. **Fork 本仓库**
2. 改 `scripts/generate_thesis.js` 里的常量和默认值（校名、页边距、字体）
3. 跑测试确认没坏：`node ynu-thesis-formatting/scripts/format_regression_test.js`
4. 把改好的版本推到你自己的仓库，用 `claude skills install <你的仓库地址>` 安装

---

## 文件结构

```
ynu-thesis-formatting-skill/
├── README.md                       # ← 你在看这个
├── package.json
└── ynu-thesis-formatting/          # Skill 主体
    ├── SKILL.md                    # Claude Code Skill 定义
    ├── assets/
    │   └── MML2OMML.XSL            # XSLT 模板（保留供参考）
    ├── evals/
    │   └── evals.json              # 评估用例
    ├── scripts/
    │   ├── generate_thesis.js      # 核心引擎
    │   ├── format_regression_test.js
    │   └── docx_input_regression_test.js
    └── workspace/
        └── regression/             # 测试夹具
```

---

## License

MIT © 2025

欢迎 Fork、修改、适配你的学校格式。提 PR 或自行发布都 OK。
