const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, TabStopType } = require('docx');

const FONT = 'Times New Roman';
const PAGE_W = 11906;
const PAGE_H = 16838;
const MARGIN = 1440;
const CONTENT_W = PAGE_W - MARGIN * 2;
const LS = 240;

const SZ_NAME = 28;
const SZ_CONTACT = 20;
const SZ_SECTION = 22;
const SZ_ENTRY = 21;
const SZ_BODY = 21;
const SZ_SMALL = 20;

function r(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: opts.sz || SZ_BODY, bold: opts.bold || false, italics: opts.italic || false });
}
function emptyP() {
  return new Paragraph({ spacing: { line: LS, before: 0, after: 0 }, children: [r('', { sz: 10 })] });
}
function nameP(name) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LS, before: 0, after: 40 }, children: [r(name, { bold: true, sz: SZ_NAME })] });
}
function contactP(text) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: LS, before: 0, after: 80 }, children: [r(text, { sz: SZ_CONTACT })] });
}
function sectionP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 160, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
    children: [r(text.toUpperCase(), { bold: true, sz: SZ_SECTION })]
  });
}
function entryHeaderP(leftText, rightText) {
  const children = [r(leftText, { bold: true, sz: SZ_ENTRY })];
  if (rightText && rightText.trim()) {
    children.push(r('\t', { sz: SZ_ENTRY }));
    children.push(r(rightText.trim(), { italic: true, sz: SZ_SMALL }));
  }
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    spacing: { line: LS, before: 80, after: 0 },
    children
  });
}
function bulletP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    indent: { left: 360, hanging: 360 },
    children: [r('- ' + text, { sz: SZ_BODY })]
  });
}
function skillsP(label, value) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [r(label + ': ', { bold: true, sz: SZ_BODY }), r(value, { sz: SZ_BODY })]
  });
}
function subLineP(text) {
  return new Paragraph({ spacing: { line: LS, before: 0, after: 0 }, children: [r(text, { sz: SZ_BODY })] });
}
function bodyP(text) {
  return new Paragraph({ spacing: { line: LS, before: 0, after: 0 }, children: [r(text, { sz: SZ_BODY })] });
}

// Build docx from the structured HTML elements sent from the browser
// Instead of parsing raw text, we receive pre-classified elements
function buildFromElements(elements) {
  const children = [];

  for (const el of elements) {
    switch (el.type) {
      case 'name':    children.push(nameP(el.text)); break;
      case 'contact': children.push(contactP(el.text)); break;
      case 'section': children.push(sectionP(el.text)); break;
      case 'entry':   children.push(entryHeaderP(el.left, el.right)); break;
      case 'bullet':  children.push(bulletP(el.text)); break;
      case 'skills':  children.push(skillsP(el.label, el.value)); break;
      case 'sub':     children.push(subLineP(el.text)); break;
      case 'gap':     children.push(emptyP()); break;
      default:        if (el.text) children.push(bodyP(el.text)); break;
    }
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN }
        }
      },
      children
    }]
  });
  return Packer.toBuffer(doc);
}

function buildCoverLetterDoc(coverText) {
  const lines = coverText.replace(/```[\w]*\n?/g, '').replace(/\*\*/g, '').replace(/\*/g, '').split('\n');
  const children = [];
  for (const line of lines) {
    const t = line.trim();
    if (!t) { children.push(emptyP()); continue; }
    children.push(new Paragraph({ spacing: { line: 276, before: 0, after: 80 }, children: [r(t, { sz: SZ_BODY })] }));
  }
  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
      children
    }]
  });
  return Packer.toBuffer(doc);
}

module.exports = { buildFromElements, parseCoverLetterToDocx: buildCoverLetterDoc };
