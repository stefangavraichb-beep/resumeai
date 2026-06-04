// Finance-grade CV generator
// Based on research from Mergers & Inquisitions, CFI, Financial Edge, CV Anywhere
// Standards: Times New Roman, 10.5-11pt body, 14pt name, 1 page A4, 1" margins
// Section order: Name > Contact > Education > Work Experience > Extracurriculars > Skills & Interests

const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, TabStopType, TabStopPosition, LevelFormat } = require('docx');

const FONT = 'Times New Roman';
// DXA units: 1 inch = 1440 DXA, 1pt = 20 DXA
// Page: A4 = 11906 x 16838 DXA
// Margins: 1 inch all sides = 1440 DXA each
// Content width: 11906 - 2880 = 9026 DXA

const PAGE_W = 11906;
const PAGE_H = 16838;
const MARGIN = 1440; // 1 inch
const CONTENT_W = PAGE_W - MARGIN * 2; // 9026 DXA

// Font sizes in half-points (docx unit)
const SZ_NAME = 28;     // 14pt - name
const SZ_CONTACT = 20;  // 10pt - contact line
const SZ_SECTION = 24;  // 12pt - section headers
const SZ_ENTRY = 22;    // 11pt - company/institution names
const SZ_BODY = 21;     // 10.5pt - body text, bullets
const SZ_SMALL = 20;    // 10pt - fine print

// Line spacing: single = 240, 1.15 = 276
const LS = 240;

function r(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: opts.sz || SZ_BODY,
    bold: opts.bold || false,
    italics: opts.italic || false,
  });
}

function emptyP() {
  return new Paragraph({
    spacing: { line: LS, lineRule: 'auto', before: 0, after: 0 },
    children: [r('', { sz: SZ_BODY })]
  });
}

// Name - centered, bold, 14pt
function nameP(name) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LS, lineRule: 'auto', before: 0, after: 40 },
    children: [r(name, { bold: true, sz: SZ_NAME })]
  });
}

// Contact - centered, 10pt
function contactP(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: LS, lineRule: 'auto', before: 0, after: 80 },
    children: [r(text, { sz: SZ_CONTACT })]
  });
}

// Section header - bold 12pt, solid bottom border, small space above
function sectionP(text) {
  return new Paragraph({
    spacing: { line: LS, lineRule: 'auto', before: 160, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
    children: [r(text, { bold: true, sz: SZ_SECTION })]
  });
}

// Entry header: bold company/school name left, italic location+date right
// Uses tab stop at right edge of content area
function entryHeaderP(leftText, rightText, italic = false) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    spacing: { line: LS, lineRule: 'auto', before: 80, after: 20 },
    children: [
      r(leftText, { bold: true, sz: SZ_ENTRY, italic }),
      r('\t', { sz: SZ_ENTRY }),
      r(rightText, { italic: true, sz: SZ_SMALL }),
    ]
  });
}

// Sub-line under entry header (degree, role description) - normal body text
function subLineP(text) {
  return new Paragraph({
    spacing: { line: LS, lineRule: 'auto', before: 0, after: 0 },
    children: [r(text, { sz: SZ_BODY })]
  });
}

// Bullet point - left indent 360 DXA, hanging 360 DXA
function bulletP(text) {
  return new Paragraph({
    spacing: { line: LS, lineRule: 'auto', before: 0, after: 0 },
    indent: { left: 360, hanging: 360 },
    children: [
      r('- ', { sz: SZ_BODY }),
      r(text, { sz: SZ_BODY })
    ]
  });
}

// Skills line: bold label, normal value
function skillsP(label, value) {
  return new Paragraph({
    spacing: { line: LS, lineRule: 'auto', before: 0, after: 0 },
    children: [
      r(label + ': ', { bold: true, sz: SZ_BODY }),
      r(value, { sz: SZ_BODY })
    ]
  });
}

// Body paragraph - normal text
function bodyP(text) {
  return new Paragraph({
    spacing: { line: LS, lineRule: 'auto', before: 0, after: 0 },
    children: [r(text, { sz: SZ_BODY })]
  });
}

// Parse plain text CV into structured docx
function parseAndBuild(resumeText) {
  const rawLines = resumeText.split('\n');
  const children = [];

  let nameAdded = false;
  let contactAdded = false;

  const SECTION_KEYWORDS = [
    'EDUCATION', 'WORK EXPERIENCE', 'EXPERIENCE', 'EXTRACURRICULAR',
    'ACTIVITIES', 'SKILLS AND INTERESTS', 'SKILLS & INTERESTS',
    'SKILLS', 'INTERESTS', 'LANGUAGES', 'ACHIEVEMENTS',
    'CERTIFICATIONS', 'PROJECTS', 'VOLUNTEERING', 'LEADERSHIP'
  ];

  const DATE_PATTERN = /\b(January|February|March|April|May|June|July|August|September|October|November|December|20\d{2}|19\d{2}|Present|Current)\b/i;
  const SKILLS_PATTERN = /^(Languages|Skills|Interests|Certifications|Hobbies|Technical Skills|Soft Skills)\s*:/i;

  for (let i = 0; i < rawLines.length; i++) {
    const raw = rawLines[i];
    const trimmed = raw.trim();

    // Skip empty lines but add spacing
    if (!trimmed) {
      if (nameAdded) children.push(emptyP());
      continue;
    }

    // Strip markdown asterisks used for bold/italic
    const clean = trimmed.replace(/\*+/g, '').trim();

    // 1. Name - first non-empty line
    if (!nameAdded) {
      children.push(nameP(clean));
      nameAdded = true;
      continue;
    }

    // 2. Contact line - has email/phone/location indicators
    if (!contactAdded && (
      clean.includes('@') ||
      clean.includes('+') ||
      clean.match(/^\w[\w\s]+,\s+\w/) ||
      clean.includes('•') ||
      clean.includes('|')
    )) {
      children.push(contactP(clean));
      contactAdded = true;
      continue;
    }

    // 3. Section headers - ALL CAPS or known keywords
    const upperClean = clean.toUpperCase();
    const isSection = SECTION_KEYWORDS.some(k => upperClean === k || upperClean.startsWith(k + ' ')) ||
      (clean === clean.toUpperCase() && clean.length > 3 && !clean.startsWith('-') && clean.replace(/[^A-Z\s]/g, '').length > 3);

    if (isSection) {
      children.push(sectionP(clean));
      continue;
    }

    // 4. Bullet points
    if (trimmed.startsWith('-') || trimmed.startsWith('•')) {
      const text = clean.replace(/^[-•]\s*/, '');
      children.push(bulletP(text));
      continue;
    }

    // 5. Skills/Languages/Interests lines
    const skillsMatch = clean.match(SKILLS_PATTERN);
    if (skillsMatch) {
      const colonIdx = clean.indexOf(':');
      const label = clean.substring(0, colonIdx).trim();
      const value = clean.substring(colonIdx + 1).trim();
      children.push(skillsP(label, value));
      continue;
    }

    // 6. Entry header lines - have a date AND are not too long
    // These are lines like "Goldman Sachs - Analyst    London, June 2024 - Present"
    if (DATE_PATTERN.test(clean) && clean.length < 200) {
      // Try to split at multiple spaces or tab
      const splitMatch = clean.match(/^(.+?)\s{2,}(.+)$/) || clean.match(/^(.+?)\t(.+)$/);
      if (splitMatch) {
        const left = splitMatch[1].trim();
        const right = splitMatch[2].trim();
        // Is this italic (company) or bold (institution)?
        const isCompany = left.includes('-') || /internship|analyst|associate|manager|director|founder|engineer|consultant|developer/i.test(left);
        children.push(entryHeaderP(left, right, isCompany));
      } else {
        // Whole line is an entry header - treat as bold italic
        children.push(new Paragraph({
          spacing: { line: LS, lineRule: 'auto', before: 80, after: 20 },
          children: [r(clean, { bold: true, italic: true, sz: SZ_ENTRY })]
        }));
      }
      continue;
    }

    // 7. Sub-lines under entry headers (degree names, course descriptions, etc.)
    // Short lines that come right after an entry header
    const prevLine = rawLines[i - 1]?.trim() || '';
    const nextLine = rawLines[i + 1]?.trim() || '';
    const looksLikeSubLine = clean.toLowerCase().startsWith('course') ||
      clean.toLowerCase().startsWith('final') ||
      clean.toLowerCase().startsWith('bsc') ||
      clean.toLowerCase().startsWith('ba ') ||
      clean.toLowerCase().startsWith('msc') ||
      clean.toLowerCase().startsWith('degree') ||
      clean.toLowerCase().startsWith('a-levels') ||
      clean.toLowerCase().startsWith('gpa') ||
      clean.toLowerCase().startsWith('sat') ||
      clean.toLowerCase().startsWith('high school');

    if (looksLikeSubLine) {
      children.push(subLineP(clean));
      continue;
    }

    // 8. Short lines that look like entry headers without dates
    // (institution names, company names that span one line)
    if (clean.length < 80 && !clean.includes(':') && (
      nextLine.startsWith('-') ||
      nextLine.toLowerCase().startsWith('course') ||
      nextLine.toLowerCase().startsWith('final') ||
      DATE_PATTERN.test(nextLine)
    )) {
      children.push(entryHeaderP(clean, '', false));
      continue;
    }

    // 9. Default - regular body text
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

// Cover letter generator - clean, consistent font throughout
function buildCoverLetterDoc(coverText) {
  const lines = coverText.split('\n');
  const children = [];

  for (const line of lines) {
    const trimmed = line.trim().replace(/\*+/g, '');
    if (!trimmed) {
      children.push(emptyP());
      continue;
    }
    children.push(new Paragraph({
      spacing: { line: 276, lineRule: 'auto', before: 0, after: 80 }, // 1.15 spacing
      children: [r(trimmed, { sz: SZ_BODY })]
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
