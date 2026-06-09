const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, TabStopType } = require('docx');

const FONT = 'Times New Roman';
const PAGE_W = 11906;
const PAGE_H = 16838;
const MARGIN = 1440;
const CONTENT_W = PAGE_W - MARGIN * 2;

const SZ_NAME = 28;
const SZ_CONTACT = 20;
const SZ_SECTION = 22;
const SZ_ENTRY = 21;
const SZ_BODY = 21;
const SZ_SMALL = 20;
const LS = 240;

function r(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: opts.sz || SZ_BODY, bold: opts.bold || false, italics: opts.italic || false });
}

function emptyP() {
  return new Paragraph({ spacing: { line: LS, before: 0, after: 0 }, children: [r('')] });
}

function nameP(name) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LS, before: 0, after: 40 },
    children: [r(name, { bold: true, sz: SZ_NAME })]
  });
}

function contactP(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LS, before: 0, after: 80 },
    children: [r(text, { sz: SZ_CONTACT })]
  });
}

function sectionP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 160, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
    children: [r(text, { bold: true, sz: SZ_SECTION })]
  });
}

function entryHeaderP(leftText, rightText) {
  const children = [r(leftText, { bold: true, sz: SZ_ENTRY })];
  if (rightText) {
    children.push(r('\t', { sz: SZ_ENTRY }));
    children.push(r(rightText, { italic: true, sz: SZ_SMALL }));
  }
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    spacing: { line: LS, before: 80, after: 20 },
    children
  });
}

function subLineP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [r(text, { sz: SZ_BODY })]
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

function bodyP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [r(text, { sz: SZ_BODY })]
  });
}

function cleanLine(text) {
  return text
    .replace(/```+/g, '')        // remove backticks
    .replace(/\*\*/g, '')        // remove bold markdown
    .replace(/\*/g, '')          // remove italic markdown
    .replace(/^#{1,4}\s*/g, '')  // remove headers
    .replace(/_{2,}/g, '')       // remove underscores
    .trim();
}

function parseAndBuild(resumeText) {
  // Clean the entire text first
  const rawLines = resumeText
    .replace(/```[\w]*\n?/g, '')   // remove code fences
    .replace(/\r\n/g, '\n')
    .split('\n');

  const children = [];
  let nameAdded = false;
  let contactAdded = false;

  const SECTIONS = ['EDUCATION', 'WORK EXPERIENCE', 'EXPERIENCE', 'EXTRACURRICULAR',
    'ACTIVITIES', 'SKILLS AND INTERESTS', 'SKILLS & INTERESTS', 'SKILLS',
    'INTERESTS', 'LANGUAGES', 'ACHIEVEMENTS', 'CERTIFICATIONS', 'PROJECTS',
    'VOLUNTEERING', 'LEADERSHIP', 'PROFESSIONAL EXPERIENCE'];

  const DATE_RE = /\b(January|February|March|April|May|June|July|August|September|October|November|December|20\d{2}|19\d{2}|Present|Current)\b/i;
  const SKILLS_RE = /^(Languages|Skills|Interests|Certifications|Technical Skills|Soft Skills|Hobbies)\s*:/i;

  for (let i = 0; i < rawLines.length; i++) {
    const trimmed = rawLines[i].trim();
    if (!trimmed) {
      if (nameAdded) children.push(emptyP());
      continue;
    }

    const clean = cleanLine(trimmed);
    if (!clean) continue;

    // 1. Name
    if (!nameAdded) {
      children.push(nameP(clean));
      nameAdded = true;
      continue;
    }

    // 2. Contact line
    if (!contactAdded && (clean.includes('@') || clean.includes('+') || clean.includes('•') || clean.includes('|'))) {
      children.push(contactP(clean));
      contactAdded = true;
      continue;
    }

    // 3. Section headers
    const upper = clean.toUpperCase();
    const isSection = SECTIONS.some(k => upper === k || upper.startsWith(k + ' ') || upper.startsWith(k + '\t')) ||
      (clean === clean.toUpperCase() && clean.length > 3 && !clean.startsWith('-') && /[A-Z]{3}/.test(clean) && !DATE_RE.test(clean));

    if (isSection) {
      children.push(sectionP(clean));
      continue;
    }

    // 4. Bullet points
    if (trimmed.startsWith('-') || trimmed.startsWith('•')) {
      children.push(bulletP(clean.replace(/^[-•]\s*/, '')));
      continue;
    }

    // 5. Skills lines
    if (SKILLS_RE.test(clean)) {
      const idx = clean.indexOf(':');
      children.push(skillsP(clean.substring(0, idx).trim(), clean.substring(idx + 1).trim()));
      continue;
    }

    // 6. Entry headers (have date, have tab/multiple spaces split)
    if (DATE_RE.test(clean) && clean.length < 200) {
      const split = clean.match(/^(.+?)\s{2,}(.+)$/) || clean.match(/^(.+?)\t(.+)$/);
      if (split) {
        children.push(entryHeaderP(split[1].trim(), split[2].trim()));
      } else {
        children.push(entryHeaderP(clean, ''));
      }
      continue;
    }

    // 7. Sub-lines (course, degree, GPA etc)
    const lc = clean.toLowerCase();
    if (lc.startsWith('course') || lc.startsWith('final') || lc.startsWith('predicted') ||
        lc.startsWith('bsc') || lc.startsWith('ba ') || lc.startsWith('msc') ||
        lc.startsWith('gpa') || lc.startsWith('degree') || lc.startsWith('a-level')) {
      children.push(subLineP(clean));
      continue;
    }

    // 8. Short lines before dated lines = entry headers without dates
    const nextLine = rawLines[i + 1]?.trim() || '';
    if (clean.length < 80 && !clean.includes(':') &&
        (DATE_RE.test(nextLine) || nextLine.startsWith('-') || nextLine.toLowerCase().startsWith('course'))) {
      children.push(entryHeaderP(clean, ''));
      continue;
    }

    // 9. Default body text
    children.push(bodyP(clean));
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
  const lines = coverText
    .replace(/```[\w]*\n?/g, '')
    .split('\n');

  const children = [];
  for (const line of lines) {
    const clean = cleanLine(line);
    if (!clean) { children.push(emptyP()); continue; }
    children.push(new Paragraph({
      spacing: { line: 276, before: 0, after: 80 },
      children: [r(clean, { sz: SZ_BODY })]
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

module.exports = { parseResumeToDocx: parseAndBuild, parseCoverLetterToDocx: buildCoverLetterDoc };
