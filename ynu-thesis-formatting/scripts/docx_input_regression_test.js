const assert = require("assert");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require("docx");

const thesis = require("./generate_thesis");

const WORK_DIR = path.join(__dirname, "..", "workspace", "regression");
const INPUT_DOCX = path.join(WORK_DIR, "plain_input.docx");
const OUTPUT_DOCX = path.join(WORK_DIR, "plain_input_formatted.docx");

async function readDocxXml(docxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  return {
    document: await zip.file("word/document.xml").async("string"),
    settings: await zip.file("word/settings.xml").async("string"),
  };
}

async function createPlainDocx() {
  fs.mkdirSync(WORK_DIR, { recursive: true });
  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ text: "第一章 绪论", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: "这是未排版 Word 中的正文，包含 $E=mc^2$。" }),
        new Paragraph({ text: "参考文献", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({ text: "[1] 张三. 普通 Word 输入中的参考文献条目。" }),
      ],
    }],
  });
  fs.writeFileSync(INPUT_DOCX, await Packer.toBuffer(doc));
}

async function main() {
  await createPlainDocx();
  assert.strictEqual(typeof thesis.extractMarkdownFromDocx, "function", "Script should export docx extraction");

  const markdown = await thesis.extractMarkdownFromDocx(INPUT_DOCX);
  assert(markdown.includes("# 第一章 绪论"), "Heading 1 should become markdown H1");
  assert(markdown.includes("这是未排版 Word 中的正文"), "Body text should be retained");

  await thesis.generateThesisDocx(markdown, OUTPUT_DOCX);
  const xml = await readDocxXml(OUTPUT_DOCX);
  assert(xml.settings.includes("w:evenAndOddHeaders"), "Formatted docx output should keep thesis headers");
  // 公式转为 Unicode 纯文本：检查公式内容是否存在（不再检查 OMML 标签）
  assert(xml.document.includes("E=mc") || xml.document.includes("E = mc") || xml.document.includes("mc²"), "Inline formula text should be preserved in output");
  const { DOMParser } = require('@xmldom/xmldom');
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml.document, 'text/xml');
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
  assert(refFound, "Reference entry from Word input should exist");

  console.log("docx_input_regression_test passed");
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
