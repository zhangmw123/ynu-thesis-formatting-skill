/**
 * 中国工科硕士论文 Word 格式化引擎
 *
 * 输入：Markdown (.md) 或未排版 Word (.docx)
 * 输出：符合硕士论文格式的 .docx
 *
 * 用法：
 *   node generate_thesis.js input.md  [output.docx]
 *   node generate_thesis.js input.docx [output.docx]
 */

const {
  Document, Packer, Paragraph, TextRun,
  Header, Footer, AlignmentType, PageNumber,
  BorderStyle, WidthType, Table, TableRow, TableCell,
  ImageRun, HeadingLevel,
  Tab, TabStopType, TableOfContents,
  LevelFormat, NumberFormat,
} = require('docx');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

// ======================== 格式常量 ========================

const FONT_SIZE = {
  '小二': 36,   // 18pt — 论文题目
  '小三号': 30, // 15pt — 一级标题
  '四号': 28,   // 14pt — 二级标题
  '小四号': 24, // 12pt — 正文、三级标题
  '五号': 21,   // 10.5pt — 表题、图题
};

const CJK_FONT = {
  heading: '黑体',
  body: '宋体',
};
const LATIN_FONT = 'Times New Roman';
const BODY_INDENT = 420; // 改变为 420 DXA 以符合 2字符缩进标准
const LINE_SPACING = 360; // 1.5倍行距（半磅）

const MARGINS = {
  top: Math.round(2.5 * 567),      // 2.50 cm = 1418 dxa
  bottom: Math.round(2.0 * 567),   // 2.00 cm = 1134 dxa
  left: Math.round(3.0 * 567),     // 3.00 cm = 1701 dxa
  right: Math.round(2.0 * 567),    // 2.00 cm = 1134 dxa
  gutter: 0,                       // 0 cm
};
const HEADER_DISTANCE = Math.round(1.5 * 567);   // 1.50 cm = 851 dxa
const FOOTER_DISTANCE = Math.round(1.75 * 567); // 1.75 cm = 992 dxa

const BIB_INDENT = 420; // 悬挂缩进两字符 420 DXA
const TWO_CHARS_420 = 420;

// 三线表边框
const TABLE_TOP_BORDER    = { style: BorderStyle.SINGLE, size: 12, color: "000000" };
const TABLE_BOTTOM_BORDER = { style: BorderStyle.SINGLE, size: 12, color: "000000" };
const HEADER_CELL_BORDER  = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const NIL_BORDER          = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };

// ======================== 文本工具 ========================

/** 判断字符是否为 CJK（中文/日文/韩文） */
function isCJK(ch) {
  const cp = ch.codePointAt(0);
  return (cp >= 0x4E00 && cp <= 0x9FFF) ||
         (cp >= 0x3400 && cp <= 0x4DBF) ||
         (cp >= 0xF900 && cp <= 0xFAFF) ||
         (cp >= 0x3040 && cp <= 0x309F) ||
         (cp >= 0x30A0 && cp <= 0x30FF) ||
         (cp >= 0xAC00 && cp <= 0xD7AF);
}

/** 判断字符是否为 Latin/数字/符号 */
function isLatin(ch) {
  const cp = ch.codePointAt(0);
  return (cp >= 0x0020 && cp <= 0x007F) ||
         (cp >= 0x00A0 && cp <= 0x00FF) ||
         (cp >= 0x0100 && cp <= 0x024F);
}

/**
 * 将文本按字体系列拆分为 segments：
 *   [{ text, font: 'cjk'|'latin'|'mixed' }]
 * 连续的同类字符合并为一个 segment。
 */
function splitTextByFont(text) {
  const segments = [];
  if (!text) return segments;
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    let kind;
    if (isCJK(ch)) kind = 'cjk';
    else if (isLatin(ch)) kind = 'latin';
    else kind = 'mixed'; // 标点、空格等归入前一段或新段
    let j = i + 1;
    while (j < text.length) {
      const ch2 = text[j];
      let k2;
      if (isCJK(ch2)) k2 = 'cjk';
      else if (isLatin(ch2)) k2 = 'latin';
      else k2 = 'mixed';
      // mixed 字符跟随前一段；cjk 和 latin 不能合并
      if (k2 === 'mixed') { j++; continue; }
      if (kind === 'mixed') { kind = k2; j++; continue; }
      if (k2 !== kind) break;
      j++;
    }
    segments.push({ text: text.slice(i, j), kind });
    i = j;
  }
  // 合并相邻同类段
  const merged = [];
  for (const seg of segments) {
    if (merged.length > 0 && merged[merged.length - 1].kind === seg.kind) {
      merged[merged.length - 1].text += seg.text;
    } else {
      merged.push(seg);
    }
  }
  return merged;
}

function convertEnglishQuotes(text) {
  if (!text) return text;
  let isOpening = true;
  return text.replace(/"/g, () => {
    const quote = isOpening ? '“' : '”';
    isOpening = !isOpening;
    return quote;
  });
}

function decodeXmlText(text) {
  return text
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function isSpecialOrChapterHeading(text) {
  const clean = text.replace(/^(?:#{1,3})\s+/, '')
                    .replace(/^第?[0-9一二三四五六七八九十]+[章节、.\s]+/, '')
                    .replace(/^[0-9]+(\.[0-9]+)*\s*[、.．]?\s*/, '')
                    .trim();
  if (/^(摘\s*要|Abstract|参考文献|References|结论|致谢|附录|目录|TOC|Table of Contents)$/i.test(clean)) {
    return true;
  }
  if (/^第?[0-9一二三四五六七八九十]+章/.test(text.replace(/^(?:#{1,3})\s+/, '').trim())) {
    return true;
  }
  return false;
}

// ======================== LaTeX → Unicode 转换 ========================

/**
 * 将 LaTeX 公式语法转换为可读的 Unicode 纯文本。
 * 不使用 OMML（Office Math Markup Language），因为 temml/XSLT 管线
 * 生成的 OMML 在 Word 中会乱码，且我们没有 pandoc 可用。
 * 转换后的纯文本使用 Times New Roman 斜体渲染，可读且正确。
 */

function latexToUnicode(text) {
  text = text.replace(/\\tag\{[^}]*\}/g, '');
  text = text.replace(/\btag\{[^}]*\}/g, '');
  text = text.replace(/\\{/g, '{').replace(/\\}/g, '}');
  text = text.replace(/\\\$/g, '$').replace(/\\#/g, '#');
  text = text.replace(/\\%/g, '%').replace(/\\&/g, '&');
  text = text.replace(/\\_/g, '_').replace(/\\~/g, '~');
  text = text.replace(/\\\^/g, '^');

  text = text.replace(/\\text\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  text = text.replace(/\\mathbf\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  text = text.replace(/\\mathrm\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  text = text.replace(/\\mathit\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  text = text.replace(/\\mathcal\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  text = text.replace(/\\mathbb\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  
  text = text.replace(/\\frac\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '($1)/($2)');
  text = text.replace(/\\sqrt\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '√($1)');

  text = text.replace(/\\begin\{cases\}/g, '{ ').replace(/\\end\{cases\}/g, ' }');
  text = text.replace(/\\begin\{matrix\}/g, '[').replace(/\\end\{matrix\}/g, ']');
  text = text.replace(/\\begin\{[^}]*\}/g, '').replace(/\\end\{[^}]*\}/g, '');

  const greekMap = {
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
    'epsilon': 'ε', 'varepsilon': 'ε',
    'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'vartheta': 'ϑ',
    'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ',
    'nu': 'ν', 'xi': 'ξ', 'omicron': 'ο', 'pi': 'π', 'varpi': 'ϖ',
    'rho': 'ρ', 'varrho': 'ϱ', 'sigma': 'σ', 'varsigma': 'ς',
    'tau': 'τ', 'upsilon': 'υ', 'phi': 'φ', 'varphi': 'ϕ',
    'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ',
    'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ', 'Upsilon': 'Υ',
    'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω'
  };
  for (const [key, val] of Object.entries(greekMap)) {
    const regex = new RegExp('\\\\' + key + '(?![a-zA-Z])', 'g');
    text = text.replace(regex, val);
  }

  const symbolMap = {
    'times': '×', 'cdot': '·', 'div': '÷', 'pm': '±', 'mp': '∓',
    'oplus': '⊕', 'ominus': '⊖', 'otimes': '⊗', 'odot': '⊙',
    'star': '⋆', 'ast': '∗', 'circ': '∘', 'bullet': '•',
    'setminus': '∖', 'cap': '∩', 'cup': '∪',
    'sqcap': '⊓', 'sqcup': '⊔', 'wedge': '∧', 'vee': '∨',
    'leq': '≤', 'geq': '≥', 'll': '≪', 'gg': '≫',
    'neq': '≠', 'ne': '≠', 'equiv': '≡', 'approx': '≈',
    'sim': '∼', 'simeq': '≃', 'cong': '≅', 'propto': '∝',
    'subset': '⊂', 'supset': '⊃', 'subseteq': '⊆', 'supseteq': '⊇',
    'in': '∈', 'ni': '∋', 'notin': '∉',
    'mid': '∣', 'parallel': '∥', 'perp': '⊥',
    'forall': '∀', 'exists': '∃', 'nexists': '∄',
    'land': '∧', 'lor': '∨', 'lnot': '¬', 'neg': '¬',
    'implies': '⇒', 'iff': '⇔',
    'rightarrow': '→', 'Rightarrow': '⇒',
    'leftarrow': '←', 'Leftarrow': '⇐',
    'leftrightarrow': '↔', 'Leftrightarrow': '⇔',
    'to': '→', 'mapsto': '↦', 'infty': '∞', 'partial': '∂',
    'nabla': '∇', 'emptyset': '∅', 'langle': '⟨', 'rangle': '⟩',
    // 大型运算符（之前缺失，导致 \sum→sum, \int→int 等问题）
    'sum': '∑', 'int': '∫', 'prod': '∏', 'oint': '∮',
    'bigcup': '⋃', 'bigcap': '⋂', 'bigvee': '⋁', 'bigwedge': '⋀',
    'coprod': '∐', 'iint': '∬', 'iiint': '∭',
    // 省略号
    'ldots': '…', 'cdots': '⋯', 'vdots': '⋮', 'ddots': '⋱',
    // 其他常用符号
    'triangle': '△', 'angle': '∠', 'measuredangle': '∡',
    'aleph': 'ℵ', 'hbar': 'ℏ', 'imath': 'ı', 'jmath': 'ȷ',
    'ell': 'ℓ', 'wp': '℘', 'Re': 'ℜ', 'Im': 'ℑ',
    'surd': '√', 'diamond': '◇', 'clubsuit': '♣',
    'diamondsuit': '♦', 'heartsuit': '♥', 'spadesuit': '♠',
    'Box': '□', 'square': '□',
  };
  for (const [key, val] of Object.entries(symbolMap)) {
    const regex = new RegExp('\\\\' + key + '(?![a-zA-Z])', 'g');
    text = text.replace(regex, val);
  }

  const mathFuncs = [
    'arg', 'max', 'min', 'lim', 'sup', 'inf', 'dim', 'log', 'lg', 'ln', 'exp',
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'coth', 'det', 'gcd'
  ];
  for (const func of mathFuncs) {
    const regex = new RegExp('\\\\' + func + '(?![a-zA-Z])', 'g');
    text = text.replace(regex, func);
  }

  text = text.replace(/\\(?:left|right|big[lr]?|Big[lr]?|bigg[lr]?|Bigg[lr]?|middle)(?:\\{|\\}|\[|\]|\(|\)|\||\\|\.)?/g, '');
  text = text.replace(/\\(?:displaystyle|scriptstyle|textstyle|limits|nolimits)\b/g, '');
  // 重音命令：去掉命令保留参数 (如 \bar{x}→x, \hat{x}→x)
  text = text.replace(/\\(?:bar|hat|tilde|vec|dot|ddot|acute|grave|breve|check|mathring)\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  text = text.replace(/\\(?:overline|underline)\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '$1');
  // 下标的 \sb{...} 和上标的 \sp{...} (少见但存在)
  text = text.replace(/\\sb\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '_{$1}');
  text = text.replace(/\\sp\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/g, '^{$1}');
  text = text.replace(/\\([a-zA-Z]+)/g, '$1');
  text = text.replace(/\s+/g, ' ').trim();
  text = text.replace(/&/g, ' ').replace(/\\\\/g, '; ');

  return text;
}

/**
 * 将 LaTeX 公式转为纯文本 TextRun 数组（Cambria Math 斜体，Word 标准数学字体）。
 * 不再尝试生成 OMML Math 对象，以避免乱码。
 */
function createFormulaRuns(latex, opts = {}) {
  const {
    fontSize = FONT_SIZE['小四号'],
  } = opts;
  const unicodeText = latexToUnicode(latex);
  return [new TextRun({
    text: unicodeText,
    font: 'Cambria Math',  // Word 标准数学字体，公式展示效果更好
    size: fontSize,
    italics: true,
  })];
}

// ======================== 行内文本解析 ========================

/**
 * 将文本转为 TextRun/OfficMath 数组。
 * - 中西文自动分字体（CJK→cjkFont, Latin→latinFont）
 * - 行内 $...$ → OMML
 * - 上标引用 \[1\] 或 [1],[2] → superscript
 * - 支持 Markdown 粗体 ** 和斜体 *
 */
function parseInlineRuns(text, opts = {}) {
  const {
    cjkFont = CJK_FONT.body,
    latinFont = LATIN_FONT,
    fontSize = FONT_SIZE['小四号'],
    bold = false,
    italics = false,
  } = opts;

  const runs = [];
  let remaining = text;

  function pushTextRun(t) {
    if (!t) return;
    const segs = splitTextByFont(t);
    for (const seg of segs) {
      const font = seg.kind === 'cjk' ? cjkFont : latinFont;
      runs.push(new TextRun({ text: seg.text, font, size: fontSize, bold, italics }));
    }
  }

  while (remaining.length > 0) {
    const matches = [];

    const fm = remaining.match(/\$([^$\n]+)\$/);
    if (fm) matches.push({ type: 'formula', match: fm });

    const cm = remaining.match(/\[(\d+(?:[-,]\s*\d+)*)\]/);
    if (cm) matches.push({ type: 'cite', match: cm });

    const bm1 = remaining.match(/\*\*([^*]+)\*\*/);
    if (bm1) matches.push({ type: 'bold', match: bm1 });

    const bm2 = remaining.match(/__([^_]+)__/);
    if (bm2) matches.push({ type: 'bold', match: bm2 });

    const im1 = remaining.match(/\*([^*]+)\*/);
    if (im1) matches.push({ type: 'italic', match: im1 });

    const im2 = remaining.match(/(?<![A-Za-z0-9])_([^_]+)_(?![A-Za-z0-9])/);
    if (im2) matches.push({ type: 'italic', match: im2 });

    if (matches.length === 0) {
      pushTextRun(remaining);
      break;
    }

    matches.sort((a, b) => a.match.index - b.match.index);
    const m = matches[0];
    const matchObj = m.match;

    if (matchObj.index > 0) {
      pushTextRun(remaining.slice(0, matchObj.index));
    }

    if (m.type === 'formula') {
      runs.push(...createFormulaRuns(matchObj[1].trim(), { fontSize }));
    } else if (m.type === 'cite') {
      runs.push(new TextRun({
        text: '[' + matchObj[1] + ']',
        font: latinFont,
        size: fontSize,
        superScript: true,
      }));
    } else if (m.type === 'bold') {
      runs.push(...parseInlineRuns(matchObj[1], { ...opts, bold: true }));
    } else if (m.type === 'italic') {
      runs.push(...parseInlineRuns(matchObj[1], { ...opts, italics: true }));
    }

    remaining = remaining.slice(matchObj.index + matchObj[0].length);
  }

  return runs.length > 0 ? runs : [new TextRun({ text: '', font: latinFont, size: fontSize })];
}

// ======================== Markdown 解析 ========================

function isTableSeparator(line) {
  return /^\|\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$/.test(line.trim());
}

function looksLikeReferenceEntry(text) {
  return /^([\[［【(（]?[0-9]+[\]］】)）]?[.．、]?)\s+/.test(text) || /^\[[0-9]+\]/.test(text);
}

const CN_NUMS = { 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九', 10: '十' };
function toChineseNum(num) {
  return CN_NUMS[num] || String(num);
}

function parseMarkdown(markdown) {
  const lines = markdown.split('\n');
  const sections = [];
  let counters = { 1: 0, 2: 0, 3: 0 };
  let chapterCounter = 0;
  let tableCounter = 0;
  let figureCounter = 0;
  let formulaCounter = 0;
  let currentChapterTitle = '绪论';
  let inReferences = false;
  let inAbstract = false;
  let pendingTableCaption = null;
  let isFirstH1 = true;

  // Pre-scan to identify the thesis title from the first non-empty line
  let titleIndex = -1;
  let titleText = '';
  for (let idx = 0; idx < lines.length; idx++) {
    const line = lines[idx].trim();
    if (!line) continue;
    if (/^(?:-{3,}|\*{3,}|_{3,})$/.test(line) || /^#+$/.test(line) || /^\*?-+\*?$/.test(line) || /^\*+$/.test(line)) continue;
    
    if (!isSpecialOrChapterHeading(line)) {
      titleIndex = idx;
      titleText = line.replace(/^#+\s*/, '').replace(/\*\*|__/g, '').trim();
    }
    break; // Only evaluate the first non-empty, non-separator line
  }

  if (titleIndex !== -1) {
    sections.push({ type: 'title', text: convertEnglishQuotes(titleText) });
    isFirstH1 = false;
  }

  let i = 0;
  while (i < lines.length) {
    if (i === titleIndex) {
      i++;
      continue;
    }
    const trimmed = lines[i].trim();

    // 空行
    if (!trimmed) { i++; continue; }

    // 忽略 markdown 分割线 (如 ---, ***, ___ 等) 以及乱码分割线 / 杂散星号 (如 #, *, *---, --- 等)
    // 也过滤 #*、#——、*—— 等 Markdown 标记残留
    if (/^(?:-{3,}|\*{3,}|_{3,})$/.test(trimmed) || /^#+$/.test(trimmed) || /^\*?-+\*?$/.test(trimmed) || /^\*+$/.test(trimmed)) {
      i++; continue;
    }
    // 过滤以 #* 或 #—— 开头的垃圾行
    if (/^#\*/.test(trimmed) || /^#——/.test(trimmed) || /^\*——/.test(trimmed)) {
      i++; continue;
    }

    // 优先识别特定的特殊标题（如摘要、Abstract、参考文献等，可能带有或不带有 # 或 ** 标记）
    const specialHeadingMatch = trimmed.match(/^(?:#\s*|(?:\*\*|__))?\s*(摘\s*要|Abstract|参考文献|References|结论|致谢|附录)\s*(?:\*\*|__)?\s*$/i);
    if (specialHeadingMatch) {
      const clean = specialHeadingMatch[1].trim();
      const normClean = clean.replace(/\s+/g, '');
      if (normClean === '摘要') {
        inAbstract = true;
        inReferences = false;
        counters[1]++; counters[2] = 0; counters[3] = 0;
        currentChapterTitle = '摘  要';
        sections.push({ type: 'abstractHeading', text: '摘  要', numStr: '' });
        i++; continue;
      }
      if (normClean.toLowerCase() === 'abstract') {
        inAbstract = true;
        inReferences = false;
        counters[1]++; counters[2] = 0; counters[3] = 0;
        currentChapterTitle = 'Abstract';
        sections.push({ type: 'abstractHeading', text: 'Abstract', numStr: '' });
        i++; continue;
      }
      if (normClean === '参考文献' || normClean.toLowerCase() === 'references') {
        inReferences = true; inAbstract = false;
        counters[1]++; counters[2] = 0; counters[3] = 0;
        currentChapterTitle = clean;
        sections.push({ type: 'heading', level: 1, rawText: clean, cleanText: clean,
          numStr: '', chapterTitle: clean, isReferenceHeading: true, headingNumber: counters[1] });
        i++; continue;
      }
      inAbstract = false;
      inReferences = false;
      counters[1]++; counters[2] = 0; counters[3] = 0;
      currentChapterTitle = clean;
      sections.push({ type: 'heading', level: 1, rawText: clean, cleanText: clean,
        numStr: '', chapterTitle: clean, headingNumber: counters[1] });
      i++; continue;
    }

    // 表题（在表格前一行）
    const captionMatch = trimmed.match(/^表\s*([0-9-.]+)?\s*(.+)$/);
    if (captionMatch && i + 1 < lines.length && lines[i + 1].trim().startsWith('|')) {
      const title = captionMatch[2].trim();
      let num;
      if (chapterCounter > 0) {
        tableCounter++;
        num = chapterCounter + '-' + tableCounter;
      } else {
        tableCounter++;
        num = String(tableCounter);
      }
      pendingTableCaption = { num, title };
      i++; continue;
    }

    // 表格
    if (trimmed.startsWith('|')) {
      const tableLines = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        tableLines.push(lines[i].trim()); i++;
      }
      const rows = [];
      for (const tl of tableLines) {
        if (isTableSeparator(tl)) continue;
        const cells = tl.split('|').slice(1, -1).map(c => convertEnglishQuotes(c.trim()));
        if (cells.length > 0) rows.push(cells);
      }
      if (rows.length > 0) {
        let num;
        if (pendingTableCaption) {
          num = pendingTableCaption.num;
        } else {
          if (chapterCounter > 0) {
            tableCounter++;
            num = chapterCounter + '-' + tableCounter;
          } else {
            tableCounter++;
            num = String(tableCounter);
          }
        }
        sections.push({ type: 'table', rows, caption: pendingTableCaption, autoNum: num });
        pendingTableCaption = null;
      }
      continue;
    }

    // 图片
    const imgMatch = trimmed.match(/!\[([^\]]*)\]\(([^)]+)\)/);
    if (imgMatch) {
      const altText = imgMatch[1];
      const imgPath = imgMatch[2];
      const figMatch = altText.match(/^图\s*([0-9-.]+)?\s*(.+)$/);
      let num, title;
      if (figMatch) {
        title = figMatch[2].trim();
      } else {
        title = altText || '图片';
      }
      if (chapterCounter > 0) {
        figureCounter++;
        num = chapterCounter + '-' + figureCounter;
      } else {
        figureCounter++;
        num = String(figureCounter);
      }
      sections.push({ type: 'image', path: imgPath, figureNum: num, figureTitle: title });
      i++; continue;
    }

    // 独立公式 $$...$$（支持单行 and 多行）
    if (trimmed.startsWith('$$')) {
      let formulaText = '';
      if (trimmed.length > 2 && trimmed.endsWith('$$')) {
        formulaText = trimmed.slice(2, -2).trim();
        i++;
      } else {
        const formulaLines = []; i++;
        while (i < lines.length && !lines[i].trim().endsWith('$$')) {
          formulaLines.push(lines[i].trim()); i++;
        }
        if (i < lines.length) {
          const close = lines[i].trim();
          const before = close.slice(0, -2).trim();
          if (before) formulaLines.push(before);
        }
        formulaText = formulaLines.join(' ').trim();
        i++;
      }
      
      // Parse out \tag{...} if present
      const tagMatch = formulaText.match(/\\tag\{([^}]+)\}/);
      let num;
      if (tagMatch) {
        num = tagMatch[1];
        formulaText = formulaText.replace(/\\tag\{[^}]*\}/g, '').trim();
      } else {
        if (chapterCounter > 0) {
          formulaCounter++;
          num = chapterCounter + '-' + formulaCounter;
        } else {
          formulaCounter++;
          num = String(formulaCounter);
        }
      }
      sections.push({ type: 'formula', text: formulaText, formulaNum: num });
      continue;
    }

    // 标题：支持 # ~ #### (h4 作为三级标题等价处理)
    const hMatch = trimmed.match(/^(#{1,4})\s+(.+)/);
    if (hMatch) {
      let level = hMatch[1].length;
      if (level >= 4) level = 3; // h4→h3, h5→h3
      const rawText = hMatch[2].trim();
      const clean = rawText.replace(/^第?[0-9一二三四五六七八九十]+[章节、.\s]+/, '')
        .replace(/^[0-9]+(\.[0-9]+)*\s*[、.．]?\s*/, '').trim();

      const processedClean = convertEnglishQuotes(clean);
      const processedRaw = convertEnglishQuotes(rawText);

      if (level === 1) {
        // 论文题目（第一个 H1，排除章标题或特殊标题）
        const isSpecial = isSpecialOrChapterHeading(trimmed);
        if (isFirstH1 && !isSpecial) {
          isFirstH1 = false;
          inReferences = false;
          sections.push({ type: 'title', text: processedClean });
          i++; continue;
        }
        isFirstH1 = false;

        // 摘要块 (支持空格，如 "摘  要") - 此时已在此前的 specialHeadingMatch 识别，此处仅作为兜底
        if (/^摘\s*要$/.test(clean) || /^Abstract$/i.test(clean)) {
          inAbstract = true;
          inReferences = false;
          counters[1]++; counters[2] = 0; counters[3] = 0;
          currentChapterTitle = /^摘\s*要$/.test(clean) ? '摘  要' : clean;
          sections.push({ type: 'abstractHeading', text: currentChapterTitle, numStr: '' });
          i++; continue;
        }
        // 参考文献 - 同样作为兜底
        if (/^(参考文献|References)$/i.test(clean)) {
          inReferences = true; inAbstract = false;
          counters[1]++; counters[2] = 0; counters[3] = 0;
          currentChapterTitle = clean;
          sections.push({ type: 'heading', level: 1, rawText: processedRaw, cleanText: processedClean,
            numStr: '', chapterTitle: clean, isReferenceHeading: true, headingNumber: counters[1] });
          i++; continue;
        }
        // 普通章标题
        inAbstract = false;
        inReferences = false;
        counters[1]++; counters[2] = 0; counters[3] = 0;
        chapterCounter++;
        tableCounter = 0;
        figureCounter = 0;
        formulaCounter = 0;
        currentChapterTitle = clean || rawText;
        
        const numStr = '第' + toChineseNum(chapterCounter) + '章';
        sections.push({ type: 'heading', level: 1, rawText: processedRaw, cleanText: processedClean,
          numStr: numStr, chapterTitle: currentChapterTitle, headingNumber: counters[1] });
      } else if (level === 2) {
        inReferences = false;
        counters[2]++; counters[3] = 0;
        const numStr = chapterCounter > 0 ? (chapterCounter + '.' + counters[2]) : String(counters[2]);
        sections.push({ type: 'heading', level: 2, rawText: processedRaw, cleanText: processedClean,
          numStr: numStr, headingNumber: counters[1] });
      } else if (level === 3) {
        inReferences = false;
        counters[3]++;
        const numStr = chapterCounter > 0 ? (chapterCounter + '.' + counters[2] + '.' + counters[3]) : (counters[2] + '.' + counters[3]);
        sections.push({ type: 'heading', level: 3, rawText: processedRaw, cleanText: processedClean,
          numStr: numStr, headingNumber: counters[1] });
      }
      i++; continue;
    }

    // Markdown 无序列表项（* item, - item）：去掉前缀作为正文段落
    const bulletMatch = trimmed.match(/^[\*\-]\s+(.+)/);
    if (bulletMatch && !inReferences && !looksLikeReferenceEntry(trimmed)) {
      const bulletText = convertEnglishQuotes(bulletMatch[1].trim());
      const isEnglishAbstract = inAbstract && currentChapterTitle === 'Abstract';
      const textToUse = isEnglishAbstract ? bulletText : bulletText;
      sections.push({ type: 'paragraph', text: textToUse, isEnglishAbstract });
      i++; continue;
    }

    // 关键词行（摘要块内，支持可选粗体标记）
    const kwMatch = trimmed.match(/^(?:\*\*|__)?(关键词\s*[：:]|Keywords\s*[：:])(?:\*\*|__)?\s*(.*)/i);
    if (inAbstract && kwMatch) {
      sections.push({ type: 'keywords', text: trimmed, isEnglishAbstract: currentChapterTitle === 'Abstract' });
      i++; continue;
    }

    // 参考文献条目：仅在已进入参考文献区域（遇到 # 参考文献 标题）后才识别
    // 不再使用 looksLikeReferenceEntry 在正文中自动检测，防止正文中编号列表被误判为参考文献
    if (inReferences) {
      if (looksLikeReferenceEntry(trimmed)) {
        // 新的一条参考文献（以 [N] 等编号开头）
        sections.push({ type: 'reference', text: trimmed });
      } else {
        // 参考文献续行：拼回上一条参考文献
        const lastSec = sections[sections.length - 1];
        if (lastSec && lastSec.type === 'reference') {
          const lastChar = lastSec.text.slice(-1);
          const firstChar = trimmed.charAt(0);
          const needsSpace = /\w/.test(lastChar) && /\w/.test(firstChar);
          lastSec.text += (needsSpace ? ' ' : '') + trimmed;
        } else {
          // 参考文献区域内不以编号开头的内容，当作普通参考文献条目处理
          sections.push({ type: 'reference', text: trimmed });
        }
      }
    } else {
      const isEnglishAbstract = inAbstract && currentChapterTitle === 'Abstract';
      const textToUse = isEnglishAbstract ? trimmed : convertEnglishQuotes(trimmed);
      sections.push({ type: 'paragraph', text: textToUse, isEnglishAbstract });
    }
    i++;
  }

  return sections;
}

// ======================== 创建段落元素 ========================

/** 创建论文题目段落 */
function createTitleParagraph(section) {
  return new Paragraph({
    children: [new TextRun({
      text: section.text,
      font: CJK_FONT.heading,
      size: FONT_SIZE['小二'],
      bold: true,
    })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 360, line: LINE_SPACING, lineRule: 'auto' },
  });
}

/** 创建摘要/Abstract 标题 */
function createAbstractHeading(section) {
  const isEnglish = /abstract/i.test(section.text);
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({
      text: section.text,
      font: isEnglish ? LATIN_FONT : CJK_FONT.heading,
      size: FONT_SIZE['小三号'],
      bold: true,
    })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 360, line: LINE_SPACING, lineRule: 'auto' },
    pageBreakBefore: true,
  });
}

/** 创建关键词段落："关键词："加粗黑体，具体关键词宋体 */
function createKeywordsParagraph(section) {
  const rawText = section.text;
  let cleanText = rawText.replace(/\*\*|__/g, '');
  const match = cleanText.match(/^(关键词\s*[：:]|Keywords\s*[：:])\s*(.*)/i);
  const runs = [];
  if (match) {
    const label = match[1];
    const isEnglishLabel = /Keywords/i.test(label) || section.isEnglishAbstract;
    // 标签部分：加粗
    runs.push(new TextRun({ 
      text: label, 
      font: isEnglishLabel ? LATIN_FONT : CJK_FONT.heading, 
      size: FONT_SIZE['小四号'], 
      bold: true 
    }));
    // 关键词内容：宋体 + Times New Roman
    if (match[2]) {
      const segs = splitTextByFont(match[2]);
      for (const seg of segs) {
        runs.push(new TextRun({ 
          text: seg.text, 
          font: isEnglishLabel ? LATIN_FONT : (seg.kind === 'cjk' ? CJK_FONT.body : LATIN_FONT), 
          size: FONT_SIZE['小四号'] 
        }));
      }
    }
  } else {
    runs.push(...parseInlineRuns(rawText, section.isEnglishAbstract ? { cjkFont: LATIN_FONT, latinFont: LATIN_FONT } : {}));
  }
  return new Paragraph({
    children: runs,
    spacing: { before: 0, after: 120, line: LINE_SPACING, lineRule: 'auto' },
    indent: { firstLine: BODY_INDENT },
  });
}

/** 创建标题段落（使用 Word 内置 Heading 样式以支持导航窗格） */
function createHeadingParagraph(section) {
  const { level, numStr, cleanText } = section;
  const headingMap = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  const styleMap = {
    1: { font: CJK_FONT.heading, size: FONT_SIZE['小三号'], before: 480, after: 360 },
    2: { font: CJK_FONT.heading, size: FONT_SIZE['四号'], before: 360, after: 240 },
    3: { font: CJK_FONT.heading, size: FONT_SIZE['小四号'], before: 240, after: 120 },
  };
  const s = styleMap[level];
  const displayText = numStr ? numStr + ' ' + cleanText : cleanText;

  return new Paragraph({
    heading: headingMap[level],
    children: [new TextRun({
      text: displayText,
      font: s.font,
      size: s.size,
      bold: true,
    })],
    alignment: level === 1 ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { before: s.before, after: s.after, line: LINE_SPACING, lineRule: 'auto' },
    indent: level === 3 ? { left: BODY_INDENT } : {},
  });
}

/** 创建正文段落（中西文分离 + 引用上标） */
function createBodyParagraph(section) {
  const opts = {};
  if (section.isEnglishAbstract) {
    opts.cjkFont = LATIN_FONT;
    opts.latinFont = LATIN_FONT;
  }
  return new Paragraph({
    children: parseInlineRuns(section.text, opts),
    spacing: { before: 0, after: 0, line: LINE_SPACING, lineRule: 'auto' },
    indent: { firstLine: BODY_INDENT },
  });
}

/** 创建参考文献条目（悬挂缩进） */
function createReferenceParagraph(section) {
  let text = section.text;
  // 剥离开头的 [1]、- 、• 等前缀，以及阿拉伯数字加标点
  text = text.replace(/^\s*(?:\[\d+\]|[-–•]|[0-9]+[、.．])\s*/, '').trim();

  const restRuns = parseInlineRuns(text, { fontSize: FONT_SIZE['小四号'] });
  return new Paragraph({
    numbering: {
      reference: "bib-list",
      level: 0,
    },
    children: restRuns,
    spacing: { before: 0, after: 0, line: 440, lineRule: 'exact' },
    indent: { left: 442, hanging: 442 },
  });
}

/** 创建三线表 */
function createThreeLineTable(rows) {
  if (rows.length === 0) return null;
  const tableWidth = 9070;
  const numCols = rows[0].length;
  const colWidth = Math.floor(tableWidth / numCols);

  function makeCell(cellText, isHeader, borders) {
    return new TableCell({
      borders,
      width: { size: colWidth, type: WidthType.DXA },
      children: [new Paragraph({
        children: parseInlineRuns(cellText, { cjkFont: isHeader ? CJK_FONT.heading : CJK_FONT.body, bold: false }),
        alignment: AlignmentType.CENTER,
        spacing: { before: isHeader ? 60 : 40, after: isHeader ? 60 : 40 },
        indent: { firstLine: 0, left: 0 }, // 显式清除表格内缩进
      })],
    });
  }

  const allRows = [
    // 表头行
    new TableRow({ children: rows[0].map(c => makeCell(c, true, {
      top: TABLE_TOP_BORDER, bottom: HEADER_CELL_BORDER, left: NIL_BORDER, right: NIL_BORDER }))
    }),
    // 数据行（除最后一行）
    ...rows.slice(1, -1).map(row => new TableRow({ children: row.map(c => makeCell(c, false, {
      top: NIL_BORDER, bottom: NIL_BORDER, left: NIL_BORDER, right: NIL_BORDER }))
    })),
  ];
  // 最后一行（如果有数据行）
  if (rows.length > 1) {
    allRows.push(new TableRow({ children: rows[rows.length - 1].map(c => makeCell(c, false, {
      top: NIL_BORDER, bottom: TABLE_BOTTOM_BORDER, left: NIL_BORDER, right: NIL_BORDER }))
    }));
  }

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: Array(numCols).fill(colWidth),
    borders: {
      top: NIL_BORDER, bottom: NIL_BORDER, left: NIL_BORDER, right: NIL_BORDER,
      insideHorizontal: NIL_BORDER, insideVertical: NIL_BORDER,
    },
    rows: allRows,
  });
}

// ======================== 页眉页脚 ========================

function createFooter() {
  return new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: '第 ', font: CJK_FONT.body, size: 20 }),
        new TextRun({ children: [PageNumber.CURRENT], font: LATIN_FONT, size: 20 }),
        new TextRun({ text: ' 页', font: CJK_FONT.body, size: 20 }),
      ],
      alignment: AlignmentType.CENTER,
    })],
  });
}

function createEvenHeader(schoolName, discipline, degreeType) {
  return new Header({
    children: [new Paragraph({
      children: [new TextRun({
        text: schoolName + '（' + discipline + '）' + degreeType + '学位论文',
        font: CJK_FONT.body, size: 20,
      })],
      alignment: AlignmentType.CENTER,
    })],
  });
}

function createOddHeader(title) {
  return new Header({
    children: [new Paragraph({
      children: [new TextRun({ text: title || '绪论', font: CJK_FONT.body, size: 20 })],
      alignment: AlignmentType.CENTER,
    })],
  });
}

// ======================== Cases Fallback Table ========================

function parseCasesLatex(latex) {
  const match = latex.match(/\\begin\{cases\}([\s\S]*?)\\end\{cases\}/);
  if (!match) return null;
  const content = match[1].trim();
  const lines = content.split(/\\\\/);
  const rows = [];
  for (const line of lines) {
    const trimmedLine = line.trim();
    if (!trimmedLine) continue;
    const parts = trimmedLine.split('&').map(p => p.trim());
    rows.push({
      val: parts[0] || '',
      cond: parts[1] || ''
    });
  }
  return rows;
}

function cleanLatexText(str) {
  return str
    .replace(/\\text\s*\{([^}]+)\}/g, '$1')
    .replace(/\\mathrm\s*\{([^}]+)\}/g, '$1')
    .replace(/\\mathbf\s*\{([^}]+)\}/g, '$1')
    .replace(/\\ge\b/g, '≥')
    .replace(/\\le\b/g, '≤')
    .replace(/\\exists\b/g, '∃')
    .replace(/\\in\b/g, '∈')
    .replace(/\\ldots\b/g, '…')
    .replace(/\\cdot\b/g, '·')
    .replace(/\\theta\b/g, 'θ')
    .trim();
}

function createCasesFallbackTable(rows) {
  const braceChars = [];
  if (rows.length === 1) {
    braceChars.push('{');
  } else if (rows.length === 2) {
    braceChars.push('⎧', '⎩');
  } else {
    braceChars.push('⎧');
    for (let i = 0; i < rows.length - 2; i++) {
      braceChars.push('⎪');
    }
    braceChars.push('⎩');
  }

  const tableRows = [];
  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];
    const brace = braceChars[idx] || '';
    const valText = cleanLatexText(row.val);
    const condText = cleanLatexText(row.cond);

    const cellL = new TableCell({
      borders: {
        top: NIL_BORDER, bottom: NIL_BORDER, left: NIL_BORDER, right: NIL_BORDER
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({ text: brace + ' ', font: LATIN_FONT, size: FONT_SIZE['小四号'] }),
            ...parseInlineRuns(valText, { fontSize: FONT_SIZE['小四号'] })
          ],
          alignment: AlignmentType.RIGHT,
          spacing: { before: 40, after: 40 }
        })
      ]
    });

    const cellR = new TableCell({
      borders: {
        top: NIL_BORDER, bottom: NIL_BORDER, left: NIL_BORDER, right: NIL_BORDER
      },
      children: [
        new Paragraph({
          children: parseInlineRuns(condText, { fontSize: FONT_SIZE['小四号'] }),
          alignment: AlignmentType.LEFT,
          spacing: { before: 40, after: 40 }
        })
      ]
    });

    tableRows.push(new TableRow({ children: [cellL, cellR] }));
  }

  return new Table({
    width: { size: 9070, type: WidthType.DXA },
    borders: {
      top: NIL_BORDER, bottom: NIL_BORDER, left: NIL_BORDER, right: NIL_BORDER,
      insideHorizontal: NIL_BORDER, insideVertical: NIL_BORDER
    },
    rows: tableRows
  });
}

function buildDocumentChildren(parsedSections) {
  const children = [];
  let keywordsCount = 0;

  for (const sec of parsedSections) {
    switch (sec.type) {
      case 'title':
        children.push(createTitleParagraph(sec));
        break;

      case 'abstractHeading':
        children.push(createAbstractHeading(sec));
        break;

      case 'keywords': {
        children.push(createKeywordsParagraph(sec));
        keywordsCount++;
        if (keywordsCount === 2) {
          children.push(new Paragraph({
            children: [],
            pageBreakBefore: true,
          }));
          children.push(new Paragraph({
            children: [new TextRun({
              text: "目  录",
              font: CJK_FONT.heading,
              size: FONT_SIZE['小三号'],
              bold: true,
            })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 480, after: 360 },
          }));
          children.push(new TableOfContents("目录", {
            hyperlink: true,
            headingStyleRange: "1-3",
          }));
        }
        break;
      }

      case 'heading':
        children.push(createHeadingParagraph(sec));
        break;

      case 'paragraph':
        children.push(createBodyParagraph(sec));
        break;

      case 'reference':
        children.push(createReferenceParagraph(sec));
        break;

      case 'table': {
        // 表题（表格上方，五号黑体居中，数字 Times New Roman）
        const capText = sec.caption
          ? ('表' + sec.caption.num + ' ' + sec.caption.title)
          : ('表' + sec.autoNum);
        children.push(new Paragraph({
          children: parseInlineRuns(capText, {
            cjkFont: CJK_FONT.heading,
            latinFont: LATIN_FONT,
            fontSize: FONT_SIZE['五号'],
            bold: true,
          }),
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 60 },
        }));
        const table = createThreeLineTable(sec.rows);
        if (table) children.push(table);
        children.push(new Paragraph({ children: [new TextRun({ text: '', font: LATIN_FONT, size: FONT_SIZE['小四号'] })] }));
        break;
      }

      case 'image': {
        // 图片
        let imageRun = null;
        if (fs.existsSync(sec.path)) {
          try {
            const data = fs.readFileSync(sec.path);
            imageRun = new ImageRun({
              type: path.extname(sec.path).slice(1).toLowerCase().replace('jpg', 'jpeg'),
              data,
              transformation: { width: 300, height: 200 },
            });
          } catch (e) {}
        }
        if (imageRun) {
          children.push(new Paragraph({
            children: [imageRun],
            alignment: AlignmentType.CENTER,
            spacing: { before: 120, after: 60 },
          }));
        } else {
          children.push(new Paragraph({
            children: [new TextRun({
              text: '[图片 ' + sec.figureNum + ': ' + sec.figureTitle + ']',
              font: CJK_FONT.body, size: FONT_SIZE['小四号'], color: '888888',
            })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 120, after: 60 },
          }));
        }
        // 图题（图下方，五号黑体居中，数字 Times New Roman）
        const capText = '图' + sec.figureNum + ' ' + sec.figureTitle;
        children.push(new Paragraph({
          children: parseInlineRuns(capText, {
            cjkFont: CJK_FONT.heading,
            latinFont: LATIN_FONT,
            fontSize: FONT_SIZE['五号'],
            bold: true,
          }),
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 120 },
        }));
        break;
      }

      case 'formula': {
        const cleanText = sec.text.trim();
        const isCases = /\\begin\{cases\}/.test(cleanText);

        if (isCases) {
          // cases 分段函数：使用无边框表格渲染
          const parsedRows = parseCasesLatex(cleanText);
          if (parsedRows && parsedRows.length > 0) {
            const casesTable = createCasesFallbackTable(parsedRows);
            if (casesTable) {
              children.push(casesTable);
              if (sec.formulaNum) {
                children.push(new Paragraph({
                  children: [
                    new Tab(),
                    new TextRun({
                      text: '(' + sec.formulaNum + ')',
                      font: LATIN_FONT,
                      size: FONT_SIZE['小四号'],
                    })
                  ],
                  tabStops: [{ position: 9070, type: TabStopType.RIGHT }],
                  spacing: { before: 60, after: 120 },
                }));
              } else {
                children.push(new Paragraph({ children: [], spacing: { before: 60, after: 60 } }));
              }
              break;
            }
          }
        }

        // 普通公式：Unicode 纯文本，Times New Roman 斜体
        if (sec.formulaNum) {
          // 带编号的显示公式：居中公式 + 右对齐编号
          children.push(new Paragraph({
            tabStops: [
              { position: 4535, type: TabStopType.CENTER },
              { position: 9070, type: TabStopType.RIGHT }
            ],
            children: [
              new Tab(),
              ...createFormulaRuns(cleanText),
              new Tab(),
              new TextRun({
                text: '(' + sec.formulaNum + ')',
                font: LATIN_FONT,
                size: FONT_SIZE['小四号'],
              })
            ],
            spacing: { before: 120, after: 120, line: LINE_SPACING, lineRule: 'auto' },
          }));
        } else {
          // 无编号显示公式：居中
          children.push(new Paragraph({
            children: createFormulaRuns(cleanText),
            alignment: AlignmentType.CENTER,
            spacing: { before: 120, after: 120, line: LINE_SPACING, lineRule: 'auto' },
          }));
        }
        break;
      }
    }
  }
  return children;
}

// ======================== 多 Section 分组 ========================

function groupSectionsByHeader(parsedSections) {
  const groups = [];
  let currentGroup = { headerTitle: '绪论', sections: [] };
  let firstChapter = true;

  for (const sec of parsedSections) {
    // 标题页、摘要等放在第一个 group
    if (sec.type === 'title' || sec.type === 'abstractHeading' || sec.type === 'keywords') {
      currentGroup.sections.push(sec);
      continue;
    }

    if (sec.type === 'heading' && sec.level === 1) {
      if (currentGroup.sections.length > 0) {
        groups.push(currentGroup);
      }
      currentGroup = {
        headerTitle: sec.cleanText || sec.chapterTitle || '绪论',
        sections: [sec],
        isFirstContent: false,
      };
      firstChapter = false;
    } else {
      if (sec.type === 'heading' && sec.level === 1) {
        currentGroup.headerTitle = sec.cleanText || sec.chapterTitle || currentGroup.headerTitle;
      }
      currentGroup.sections.push(sec);
    }
  }
  if (currentGroup.sections.length > 0) groups.push(currentGroup);
  return groups;
}

/** 将 section group 转为 docx Section 配置 */
function createSectionConfig(group, index, totalGroups, options) {
  const { schoolName, discipline, degreeType } = options;
  const isFirst = index === 0;
  const isFirstChapter = index === 1; // 第二个 section 通常是第一章（绪论）

  // 页面编号策略：
  // - 前置部分（section 0）：小写罗马数字，首页（题目页）不显示页码
  // - 第一章（section 1）：阿拉伯数字，从第 1 页开始
  // - 后续章节：继续阿拉伯数字编号
  const pageNumbers = isFirst
    ? { formatType: NumberFormat.LOWER_ROMAN }
    : isFirstChapter
      ? { start: 1, formatType: NumberFormat.DECIMAL }
      : { formatType: NumberFormat.DECIMAL };

  const sectionConfig = {
    properties: {
      titlePage: isFirst,
      page: {
        size: { width: 11906, height: 16838 },
        margin: {
          top: MARGINS.top, bottom: MARGINS.bottom,
          left: MARGINS.left, right: MARGINS.right,
          gutter: MARGINS.gutter,
          header: HEADER_DISTANCE, footer: FOOTER_DISTANCE,
        },
        pageNumbers,
      },
    },
    headers: {
      default: createOddHeader(group.headerTitle),
      even: createEvenHeader(schoolName, discipline, degreeType),
    },
    footers: { default: createFooter() },
    children: buildDocumentChildren(group.sections),
  };

  // 前置部分首页（题目页）不显示页码
  if (isFirst) {
    sectionConfig.footers.first = createEmptyFooter();
  }

  return sectionConfig;
}

function createEmptyFooter() {
  return new Footer({
    children: [new Paragraph({
      children: [new TextRun({ text: '', font: LATIN_FONT, size: 20 })],
      alignment: AlignmentType.CENTER,
    })],
  });
}

// ======================== 主入口 ========================

/**
 * 从 Markdown 字符串生成格式化论文 .docx
 */
async function generateThesisDocx(markdown, outputPath, options = {}) {
  const opts = {
    schoolName: '云南大学',
    discipline: '专业',
    degreeType: '硕士',
    ...options,
  };
  const sections = parseMarkdown(markdown);
  const grouped = groupSectionsByHeader(sections);

  const doc = new Document({
    features: {
      updateFields: true,
    },
    evenAndOddHeaderAndFooters: true,
    numbering: {
      config: [
        {
          reference: "bib-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: "[%1]",
              start: 1,
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 440, hanging: 440 },
                },
              },
            },
          ],
        },
      ],
    },
    styles: {
      default: {
        document: {
          run: {
            font: CJK_FONT.body,
            size: FONT_SIZE['小四号'],
          },
        },
      },
    },
    sections: grouped.map((g, i) => createSectionConfig(g, i, grouped.length, opts)),
  });

  return Packer.toBuffer(doc).then(buffer => {
    if (outputPath) fs.writeFileSync(outputPath, buffer);
    return outputPath;
  });
}

// ======================== DOCX → DOCX 转换 ========================

/**
 * 从未排版的 .docx 提取结构化内容（Markdown 文本），
 * 然后复用 Markdown 管线重新排版。
 * 相比上一版改进：保留表格结构、保留图片引用。
 */
async function extractMarkdownFromDocx(inputFile) {
  const zip = await JSZip.loadAsync(fs.readFileSync(inputFile));
  const docEntry = zip.file('word/document.xml');
  if (!docEntry) throw new Error('不是有效的 .docx 文件: ' + inputFile);

  const docXml = await docEntry.async('string');
  const lines = [];
  let tableCount = 0;
  let figureCount = 0;

  // 使用简单的状态机遍历 body 的子元素
  const bodyMatch = docXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) throw new Error('无法解析 document.xml');

  const bodyContent = bodyMatch[1];

  // 按 <w:p> 段落和 <w:tbl> 表格分割
  const elements = [];
  const re = /(<w:p[\s>][\s\S]*?<\/w:p>)|(<w:tbl[\s>][\s\S]*?<\/w:tbl>)/g;
  let match;
  while ((match = re.exec(bodyContent)) !== null) {
    if (match[1]) elements.push({ type: 'p', xml: match[1] });
    else if (match[2]) elements.push({ type: 'tbl', xml: match[2] });
  }

  for (const el of elements) {
    if (el.type === 'p') {
      // 提取段落文本 (支持 w:t 文本和 m:oMath 公式对象，按文档顺序)
      // 不再生成 OMML_MATH base64 标记，而是提取公式纯文本并用 $ 包裹
      const parts = [];
      const partRe = /(<w:t\b[^>]*>([\s\S]*?)<\/w:t>)|(<m:oMath\b[^>]*>([\s\S]*?)<\/m:oMath>)|(<m:oMathPara\b[^>]*>([\s\S]*?)<\/m:oMathPara>)/g;
      let pm;
      while ((pm = partRe.exec(el.xml)) !== null) {
        if (pm[1]) {
          parts.push(decodeXmlText(pm[2]));
        } else if (pm[3] || pm[5]) {
          // OMML 公式对象：提取其中的文本并用 $ 包裹
          const ommlXml = pm[3] || pm[5];
          const textMatches = [];
          const tRe = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
          let tm;
          while ((tm = tRe.exec(ommlXml)) !== null) {
            textMatches.push(decodeXmlText(tm[1]));
          }
          if (textMatches.length > 0) {
            parts.push('$' + textMatches.join('') + '$');
          }
        }
      }
      const full = parts.join('').trim();
      if (!full) { lines.push(''); continue; }

      // 检测段落样式
      const styleMatch = el.xml.match(/<w:pStyle[^>]*w:val="([^"]+)"/);
      const style = styleMatch ? styleMatch[1].toLowerCase() : '';

      // 检测是否为图片（包含 w:drawing）
      const hasDrawing = /<w:drawing/.test(el.xml);

      // 判断是否为标题样式 (支持空格，如 "heading 1" 或 "标题 1")
      const isHeading1Style = /heading\s*1|标题\s*1|^1$/.test(style);
      const isHeading2Style = /heading\s*2|标题\s*2|^2$/.test(style);
      const isHeading3Style = /heading\s*3|标题\s*3|^3$/.test(style);

      // 判断是否像章标题（短文本，以"第X章"开头）
      const looksLikeChapterTitle = /^第[一二三四五六七八九十0-9]+章/.test(full) && full.length <= 40;
      // 判断是否像特殊标题（摘要、参考文献等，支持空格）
      const looksLikeSpecialHeading = /^(摘\s*要|Abstract|参考文献|References|结论|致谢|附录)$/i.test(full);

      if (hasDrawing) {
        figureCount++;
        lines.push('![图' + figureCount + '](embedded)');
      } else if (looksLikeSpecialHeading || looksLikeChapterTitle) {
        // 短章标题或特殊标题 → H1
        lines.push('# ' + full);
      } else if (isHeading1Style && full.length <= 40) {
        // Heading 1 样式但文本较短（如"绪论"、"研究区概况"等无"第X章"前缀的标题）
        lines.push('# ' + full);
      } else if (isHeading1Style && full.length > 40) {
        // Heading 1 样式但文本很长 → 这是章节摘要误标为标题，降级为正文
        lines.push(full);
      } else if (isHeading3Style || (/^[0-9]+\.[0-9]+\.[0-9]+(?:\s|[、.．]|$|(?=[^\d]))/.test(full) && !full.endsWith('。') && !/^[0-9]+\.[0-9]+\.[0-9]+\s*节/.test(full) && full.length < 60)) {
        // 优先匹配三级标题
        lines.push('### ' + full);
      } else if (isHeading2Style || (/^[0-9]+\.[0-9]+(?:\s|[、.．]|$|(?=[^\d]))/.test(full) && !full.endsWith('。') && !/^[0-9]+\.[0-9]+\s*节/.test(full) && full.length < 60)) {
        // 再匹配二级标题
        lines.push('## ' + full);
      } else {
        lines.push(full);
      }
    } else if (el.type === 'tbl') {
      // 提取表格为 Markdown 表格格式
      tableCount++;
      const rows = [];
      const rowRe = /<w:tr[\s>]([\s\S]*?)<\/w:tr>/g;
      let rm;
      while ((rm = rowRe.exec(el.xml)) !== null) {
        const cells = [];
        const cellRe = /<w:tc[\s>]([\s\S]*?)<\/w:tc>/g;
        let cm;
        while ((cm = cellRe.exec(rm[1])) !== null) {
          const ct = [];
          const ctRe = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
          let ctm;
          while ((ctm = ctRe.exec(cm[1])) !== null) {
            ct.push(decodeXmlText(ctm[1]));
          }
          cells.push(ct.join('').trim());
        }
        if (cells.length > 0) rows.push(cells);
      }
      if (rows.length > 0) {
        // 不单独插入 "表X" 标题，使其能直接与前一行的表题段落绑定
        const header = '| ' + rows[0].map(c => c || ' ').join(' | ') + ' |';
        const sep = '|' + rows[0].map(() => '------').join('|') + '|';
        lines.push(header);
        lines.push(sep);
        for (let r = 1; r < rows.length; r++) {
          lines.push('| ' + rows[r].map(c => c || ' ').join(' | ') + ' |');
        }
      }
    }
    lines.push('');
  }

  return lines.join('\n');
}

/**
 * 从 .docx 文件生成格式化论文 .docx（便捷封装）
 */
async function reformatDocx(inputPath, outputPath, options = {}) {
  const md = await extractMarkdownFromDocx(inputPath);
  return generateThesisDocx(md, outputPath, options);
}

// ======================== CLI ========================

function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error('用法: node generate_thesis.js <input.md|input.docx> [output.docx]');
    process.exit(1);
  }
  const inputFile = args[0];
  const outputFile = args[1] || inputFile.replace(/\.(md|docx)$/i, '') + '_formatted.docx';

  if (!fs.existsSync(inputFile)) {
    console.error('错误: 文件不存在 "' + inputFile + '"');
    process.exit(1);
  }

  const ext = path.extname(inputFile).toLowerCase();

  const promise = ext === '.docx'
    ? reformatDocx(inputFile, outputFile)
    : Promise.resolve(fs.readFileSync(inputFile, 'utf8')).then(md => generateThesisDocx(md, outputFile));

  promise
    .then(f => console.log('已生成: ' + f))
    .catch(err => { console.error(err); process.exit(1); });
}

if (require.main === module) { main(); }

module.exports = { generateThesisDocx, parseMarkdown, extractMarkdownFromDocx, reformatDocx };
