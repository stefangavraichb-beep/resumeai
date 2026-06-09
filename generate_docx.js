const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, TabStopType } = require('docx');

// ── KARINA OROSZ FORMAT ───────────────────────────────────────────────────────
// Font: Calibri
// Name: 14pt bold centered
// Contact: 11pt centered with • separators
// Section headers: 11pt BOLD ALL CAPS + bottom border, space before
// Entry headers: 11pt bold italic left + italic date right (tab stop)
// Sub-lines: 11pt normal
// Bullets: 11pt, dash prefix, left indent
// Skills: 11pt bold label + normal value
// Spacing: generous — 160 before sections, 80 before entries, 40 after entries

const FONT = 'Calibri';
const PAGE_W = 11906;  // A4
const PAGE_H = 16838;
const MARGIN = 1080;   // ~0.75 inch — same as Karina's doc
const CONTENT_W = PAGE_W - MARGIN * 2;  // 9746 DXA

const SZ_NAME    = 28;  // 14pt
const SZ_CONTACT = 22;  // 11pt
const SZ_SECTION = 22;  // 11pt
const SZ_ENTRY   = 22;  // 11pt
const SZ_BODY    = 22;  // 11pt
const SZ_DATE    = 22;  // 11pt

const LS = 240; // single spacing

function r(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: opts.sz || SZ_BODY,
    bold: opts.bold || false,
    italics: opts.italic || false,
  });
}

// Empty line — small height
function emptyP() {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [r('', { sz: 14 })]
  });
}

// NAME — bold, centered, 14pt
function nameP(name) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LS, before: 0, after: 60 },
    children: [r(name, { bold: true, sz: SZ_NAME })]
  });
}

// CONTACT — centered, 11pt
function contactP(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LS, before: 0, after: 60 },
    children: [r(text, { sz: SZ_CONTACT })]
  });
}

// SECTION HEADER — BOLD ALL CAPS + bottom border
function sectionP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 200, after: 80 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 }
    },
    children: [r(text.toUpperCase(), { bold: true, sz: SZ_SECTION })]
  });
}

// ENTRY HEADER — bold italic company name left, italic date right
function entryHeaderP(leftText, rightText) {
  const children = [r(leftText, { bold: true, italic: true, sz: SZ_ENTRY })];
  if (rightText && rightText.trim()) {
    children.push(r('\t', { sz: SZ_ENTRY }));
    children.push(r(rightText.trim(), { italic: true, sz: SZ_DATE }));
  }
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    spacing: { line: LS, before: 100, after: 0 },
    children
  });
}

// SUB LINE — normal text under entry (degree, role description)
function subLineP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [r(text, { sz: SZ_BODY })]
  });
}

// BULLET — dash prefix, indented
function bulletP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    indent: { left: 360, hanging: 360 },
    children: [
      r('- ', { sz: SZ_BODY }),
      r(text, { sz: SZ_BODY })
    ]
  });
}

// SKILLS LINE — bold label + normal value
function skillsP(label, value) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [
      r(label + ': ', { bold: true, sz: SZ_BODY }),
      r(value, { sz: SZ_BODY })
    ]
  });
}

function bodyP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [r(text, { sz: SZ_BODY })]
  });
}

// ── BUILD FROM ELEMENTS (sent from browser DOM) ───────────────────────────────
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
      case 'body':    if (el.text) children.push(bodyP(el.text)); break;
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

// ── COVER LETTER ──────────────────────────────────────────────────────────────
function buildCoverLetterDoc(coverText) {
  const lines = coverText
    .replace(/```[\w]*\n?/g, '')
    .replace(/\*\*/g, '')
    .replace(/\*/g, '')
    .split('\n');

  const children = [];
  for (const line of lines) {
    const t = line.trim();
    if (!t) { children.push(emptyP()); continue; }
    children.push(new Paragraph({
      spacing: { line: 276, before: 0, after: 80 },
      children: [r(t, { sz: SZ_BODY })]
    }));
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

module.exports = { buildFromElements, parseCoverLetterToDocx: buildCoverLetterDoc };
