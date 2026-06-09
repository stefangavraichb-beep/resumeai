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

function subLineP(text) {
  return new Paragraph({
    spacing: { line: LS, before: 0, after: 0 },
    children: [r(text, { italic: false, sz: SZ_BODY })]
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

// ── MAIN PARSER ───────────────────────────────────────────────────────────────

function parseAndBuild(resumeText) {
  // Step 1: Aggressive cleaning
  let text = resumeText
    .replace(/```[\w]*\n?/g, '\n')
    .replace(/```/g, '\n')
    .replace(/\*\*/g, '')
    .replace(/\*/g, '')
    .replace(/^#{1,4}\s*/gm, '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');

  // Step 2: Fix missing newlines — insert newline before known section keywords
  const SECTION_KEYWORDS = ['EDUCATION', 'WORK EXPERIENCE', 'PROFESSIONAL EXPERIENCE', 'EXPERIENCE',
    'EXTRACURRICULAR ACTIVITIES', 'EXTRACURRICULAR', 'ACTIVITIES',
    'SKILLS AND INTERESTS', 'SKILLS & INTERESTS', 'SKILLS', 'INTERESTS',
    'LANGUAGES', 'ACHIEVEMENTS', 'CERTIFICATIONS', 'VOLUNTEERING', 'PROJECTS', 'LEADERSHIP'];

  for (const kw of SECTION_KEYWORDS) {
    // Insert newline before keyword if it's not already at start of line
    text = text.replace(new RegExp(`([^\\n])(${kw})`, 'g'), '$1\n$2');
  }

  // Step 3: Insert newlines before bullet points if missing
  text = text.replace(/([^-\n])(- [A-Z])/g, '$1\n$2');

  const lines = text.split('\n').map(l => l.trim()).filter((l, i, arr) => {
    // Remove consecutive empty lines
    if (!l && arr[i - 1] === '') return false;
    return true;
  });

  const children = [];
  let nameAdded = false;
  let contactAdded = false;

  const DATE_RE = /\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?|20\d{2}|19\d{2}|Present|Current)\b/i;
  const SKILLS_RE = /^(Languages|Skills|Interests|Technical|Technical Skills|Soft Skills|Certifications|Hobbies)\s*:/i;
  const SECTIONS_SET = new Set(SECTION_KEYWORDS);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (!line) {
      if (nameAdded) children.push(emptyP());
      continue;
    }

    // 1. Name — first non-empty line
    if (!nameAdded) {
      children.push(nameP(line));
      nameAdded = true;
      continue;
    }

    // 2. Contact line
    if (!contactAdded && (line.includes('@') || line.includes('+44') || line.includes('+4') || line.includes('•') || (line.includes('|') && line.includes('@')))) {
      children.push(contactP(line));
      contactAdded = true;
      continue;
    }

    // 3. Section headers
    const upper = line.toUpperCase().trim();
    const isSection = SECTIONS_SET.has(upper) ||
      SECTION_KEYWORDS.some(k => upper === k) ||
      (line === line.toUpperCase() && line.length > 3 && line.length < 40 && !line.startsWith('-') && /[A-Z]{3}/.test(line) && !DATE_RE.test(line) && !line.includes('@'));

    if (isSection) {
      children.push(sectionP(line));
      continue;
    }

    // 4. Bullet points
    if (line.startsWith('-') || line.startsWith('•')) {
      children.push(bulletP(line.replace(/^[-•]\s*/, '')));
      continue;
    }

    // 5. Skills lines
    if (SKILLS_RE.test(line)) {
      const idx = line.indexOf(':');
      children.push(skillsP(line.substring(0, idx).trim(), line.substring(idx + 1).trim()));
      continue;
    }

    // 6. Entry headers — contain a date AND can be split into left/right
    if (DATE_RE.test(line) && line.length < 250) {
      // Try multiple spaces split first
      const split = line.match(/^(.+?)\s{2,}(.+)$/) || line.match(/^(.+?)\t(.+)$/);
      if (split) {
        children.push(entryHeaderP(split[1].trim(), split[2].trim()));
      } else {
        // Try splitting at the date itself
        const dateMatch = line.match(/^(.+?)((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{4}.*|20\d{2}.*)$/i);
        if (dateMatch && dateMatch[1].length > 3) {
          children.push(entryHeaderP(dateMatch[1].trim(), dateMatch[2].trim()));
        } else {
          children.push(entryHeaderP(line, ''));
        }
      }
      continue;
    }

    // 7. Sub-lines
    const lc = line.toLowerCase();
    if (lc.startsWith('course') || lc.startsWith('final') || lc.startsWith('predicted') ||
        lc.startsWith('bsc') || lc.startsWith('ba ') || lc.startsWith('msc') || lc.startsWith('ma ') ||
        lc.startsWith('gpa') || lc.startsWith('degree') || lc.startsWith('a-level') ||
        lc.startsWith('relevant modules') || lc.startsWith('modules')) {
      children.push(subLineP(line));
      continue;
    }

    // 8. Short lines before dated/bulleted lines = entry headers without dates
    const nextLine = lines[i + 1] || '';
    if (line.length < 80 && !line.includes(':') && !line.includes('@') &&
        (DATE_RE.test(nextLine) || nextLine.startsWith('-') || nextLine.toLowerCase().startsWith('course') || nextLine.toLowerCase().startsWith('predicted'))) {
      children.push(entryHeaderP(line, ''));
      continue;
    }

    // 9. Default body
    children.push(bodyP(line));
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
  const cleaned = coverText
    .replace(/```[\w]*\n?/g, '')
    .replace(/```/g, '')
    .replace(/\*\*/g, '')
    .replace(/\*/g, '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n');

  const lines = cleaned.split('\n');
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

module.exports = { parseResumeToDocx: parseAndBuild, parseCoverLetterToDocx: buildCoverLetterDoc };
