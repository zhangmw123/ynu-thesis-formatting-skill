const assert = require("assert");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const { generateThesisDocx } = require("./generate_thesis");

const WORK_DIR = path.join(__dirname, "..", "workspace", "regression");
const OUT_DOCX = path.join(WORK_DIR, "format_regression_output.docx");

function paragraphContaining(xml, text) {
  const paragraphs = xml.match(/<w:p[\s\S]*?<\/w:p>/g) || [];
  return paragraphs.find((paragraph) => {
    const plainText = paragraph.replace(/<[^>]+>/g, '');
    return plainText.includes(text);
  }) || "";
}

function assertNoFirstLineIndent(xml, text) {
  const paragraph = paragraphContaining(xml, text);
  assert(paragraph, `Expected paragraph containing "${text}"`);
  const firstLineMatch = paragraph.match(/w:firstLine="([^"]+)"/);
  if (firstLineMatch) {
    assert(
      firstLineMatch[1] === "0",
      `Paragraph containing "${text}" must not have non-zero first-line indentation`,
    );
  }
}

async function readDocxXml(docxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  const entries = {};
  for (const name of Object.keys(zip.files)) {
    if (name.endsWith(".xml")) {
      entries[name] = await zip.files[name].async("string");
    }
  }
  return entries;
}

async function main() {
  fs.mkdirSync(WORK_DIR, { recursive: true });

  const markdown = [
    "# 第一章 绪论",
    "",
    "正文包含行内公式 $E=mc^2$，并继续说明文字。",
    "",
    "$$",
    "\\frac{a+b}{c}",
    "$$",
    "",
    "$$\\theta_{coh} \\geq \\sum_{i=1}^{n} x_i$$",
    "",
    "表1 示例表",
    "| 指标 | 数值 |",
    "|------|------|",
    "| A | 1 |",
    "",
    "![图1 结构图](missing.png)",
    "",
    "# 参考文献",
    "",
    "[1] 张三. 面向知识图谱构建的长标题文献条目，用于检查悬挂缩进而不是首行缩进。",
    "",
    "# 第二章 方法",
    "",
    "第二章正文。",
  ].join("\n");

  await generateThesisDocx(markdown, OUT_DOCX);
  const xml = await readDocxXml(OUT_DOCX);
  const documentXml = xml["word/document.xml"];
  const settingsXml = xml["word/settings.xml"];
  const headerXml = Object.entries(xml)
    .filter(([name]) => /^word\/header\d+\.xml$/.test(name))
    .map(([, content]) => content)
    .join("\n");

  // 公式应转为 Unicode 纯文本（不再使用 OMML，避免乱码）
  assert(documentXml.includes("θ") || documentXml.includes("&"), "Formulas should contain readable Unicode symbols");
  assert(!documentXml.includes("\\theta"), "Raw LaTeX commands should not leak into Word formula text");
  assert(!documentXml.includes("\\frac"), "Raw LaTeX \\frac should not leak into output");

  assert(settingsXml.includes("w:evenAndOddHeaders"), "Document settings must enable odd/even headers");
  assert(documentXml.includes('w:type="even"'), "Section must reference an even-page header");
  assert(documentXml.includes('w:type="default"'), "Section must reference the odd/default-page header");
  assert(headerXml.includes("云南大学"), "Even-page header should contain the school name");
  assert(headerXml.includes("绪论"), "Odd-page header should contain the chapter title");
  assert(headerXml.includes("方法"), "Odd-page headers should update for later chapters");

  const { DOMParser } = require('@xmldom/xmldom');
  const parser = new DOMParser();
  const doc = parser.parseFromString(documentXml, 'text/xml');
  const paragraphs = doc.getElementsByTagName('w:p');

  let refFound = false;
  for (let idx = 0; idx < paragraphs.length; idx++) {
    const p = paragraphs[idx];
    const texts = p.getElementsByTagName('w:t');
    let textContent = '';
    for (let j = 0; j < texts.length; j++) {
      textContent += texts[j].textContent || '';
    }

    if (textContent.includes("张三")) {
      refFound = true;
      const pPrs = p.getElementsByTagName('w:pPr');
      assert(pPrs.length > 0, "Reference entry should have w:pPr");
      const pPr = pPrs[0];

      const inds = pPr.getElementsByTagName('w:ind');
      assert(inds.length > 0, "Reference entry should have w:ind");
      const ind = inds[0];

      const left = ind.getAttribute('w:left');
      const hanging = ind.getAttribute('w:hanging');
      const firstLine = ind.getAttribute('w:firstLine');

      assert.strictEqual(left, '442', `Reference entry w:left should be 442, got ${left}`);
      assert.strictEqual(hanging, '442', `Reference entry w:hanging should be 442, got ${hanging}`);
      assert(firstLine === null || firstLine === '0' || firstLine === '', `Reference entry w:firstLine should be null, empty, or 0, got ${firstLine}`);

      const numPrs = pPr.getElementsByTagName('w:numPr');
      assert.strictEqual(numPrs.length, 1, "Reference entry must have w:numPr (Word automatic numbering)");

      const spacingElements = pPr.getElementsByTagName('w:spacing');
      assert(spacingElements.length > 0, "Reference entry should have w:spacing");
      const spacing = spacingElements[0];
      const line = spacing.getAttribute('w:line');
      const lineRule = spacing.getAttribute('w:lineRule');
      assert.strictEqual(line, '440', `Reference entry w:line spacing should be 440, got ${line}`);
      assert.strictEqual(lineRule, 'exact', `Reference entry w:lineRule spacing should be exact, got ${lineRule}`);
    }
  }
  assert(refFound, "Reference entry containing '张三' should exist");

  assertNoFirstLineIndent(documentXml, "指标");
  assertNoFirstLineIndent(documentXml, "A");

  const figureCaption = paragraphContaining(documentXml, "图1-1 结构图");
  assert(figureCaption, "Figure caption paragraph should exist");
  assert(/<w:sz\b[^>]*w:val="21"/.test(figureCaption), "Figure captions must use five-point Chinese size");

  console.log("format_regression_test passed");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
