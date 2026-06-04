const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, LevelFormat, TabStopType, WidthType } = require('docx');

// Exact replica of Karina's CV format based on parsed XML:
// - Font: Times New Roman throughout
// - Name: sz=40 (20pt), bold, centered
// - Contact: sz=18 (9pt), centered
// - Section headers: sz=28 (14pt), bold, with horizontal rule via paragraph border
// - Institution/Company: sz=23 (11.5pt), bold (italic for companies)
// - Body text: sz=20-21 (10-10.5pt)
// - Bullets: indented left=720, hanging=360
// - Margins: top=990, bottom=270, left=1440, right=900 (twips)

const FONT = 'Times New Roman';
const MARGIN = { top: 990, bottom: 270, left: 1440, right: 900 };

function mkRun(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: opts.size || 20,
    bold: opts.bold || false,
    italics: opts.italics || false,
  });
}

function nameP(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, lineRule: 'auto', after: 0 },
    children: [mkRun(text, { bold: true, size: 40 })]
  });
}

function contactP(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, lineRule: 'auto', after: 0 },
    children: [mkRun(text, { size: 18 })]
  });
}

function sectionHeader(text) {
  return new Paragraph({
    spacing: { line: 240, lineRule: 'auto', before: 120, after: 0 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
    children: [mkRun(text, { bold: true, size: 28 })]
  });
}

function institutionLine(name, location, date, isItalic = false) {
  // Bold name on left, italic location+date on right (spaces used to push right like original)
  const children = [
    mkRun(name, { bold: true, size: 23, italics: isItalic }),
    mkRun('   ', { size: 21 }),
    mkRun(location + ', ' + date, { italics: true, size: 21 }),
  ];
  return new Paragraph({
    spacing: { line: 240, lineRule: 'auto', before: 0, after: 0 },
    children
  });
}

function bodyLine(text, size = 21) {
  return new Paragraph({
    spacing: { line: 240, lineRule: 'auto', after: 0 },
    children: [mkRun(text, { size })]
  });
}

function bulletLine(text) {
  return new Paragraph({
    spacing: { line: 240, lineRule: 'auto', after: 0 },
    indent: { left: 720, hanging: 360 },
    children: [
      mkRun('- ', { size: 20 }),
      mkRun(text, { size: 20 })
    ]
  });
}

function boldLabelLine(label, value) {
  return new Paragraph({
    spacing: { line: 240, lineRule: 'auto', after: 0 },
    children: [
      mkRun(label + ': ', { bold: true, size: 20 }),
      mkRun(value, { size: 20 })
    ]
  });
}

function emptyLine() {
  return new Paragraph({
    spacing: { line: 240, lineRule: 'auto', after: 0 },
    children: [mkRun('', { size: 20 })]
  });
}

function parseAndBuild(resumeText) {
  const lines = resumeText.split('\n');
  const children = [];

  let state = 'start';
  let nameAdded = false;
  let contactAdded = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    if (!trimmed) {
      if (nameAdded) children.push(emptyLine());
      continue;
    }

    // Name (first non-empty line)
    if (!nameAdded) {
      children.push(nameP(trimmed));
      nameAdded = true;
      continue;
    }

    // Contact line (has @ or • or + or city info)
    if (!contactAdded && (trimmed.includes('@') || trimmed.includes('•') || trimmed.includes('+') || trimmed.match(/^\w+,\s+\w+/))) {
      children.push(contactP(trimmed));
      contactAdded = true;
      continue;
    }

    // Section headers - ALL CAPS words or known section keywords
    const upperTrimmed = trimmed.toUpperCase();
    const sectionKeywords = ['EDUCATION', 'WORK EXPERIENCE', 'EXTRACURRICULAR', 'SKILLS AND INTERESTS', 'SKILLS', 'INTERESTS', 'LANGUAGES', 'ACTIVITIES', 'VOLUNTEERING', 'ACHIEVEMENTS', 'CERTIFICATIONS', 'PROJECTS'];
    const isSection = sectionKeywords.some(k => upperTrimmed.startsWith(k)) ||
      (trimmed === upperTrimmed && trimmed.length > 4 && !trimmed.startsWith('-') && trimmed.split(' ').length <= 5);

    if (isSection) {
      children.push(emptyLine());
      children.push(sectionHeader(trimmed));
      children.push(emptyLine());
      continue;
    }

    // Bullet lines
    if (trimmed.startsWith('-') || trimmed.startsWith('•')) {
      const text = trimmed.replace(/^[-•]\s*/, '');
      children.push(bulletLine(text));
      continue;
    }

    // Skills/Languages/Interests lines with bold label
    const labelMatch = trimmed.match(/^(Languages|Skills|Interests|Certifications|Awards):\s*(.*)/i);
    if (labelMatch) {
      children.push(boldLabelLine(labelMatch[1], labelMatch[2]));
      continue;
    }

    // Institution/company lines — detect if they have a date pattern
    const datePattern = /\b(January|February|March|April|May|June|July|August|September|October|November|December|20\d\d|19\d\d)\b/;
    const hasDates = datePattern.test(trimmed);

    if (hasDates && trimmed.length < 150) {
      // Try to split name from date
      const parts = trimmed.split(/\s{3,}|\t+/);
      if (parts.length >= 2) {
        const name = parts[0].trim().replace(/\*/g, '');
        const rest = parts.slice(1).join(' ').trim().replace(/\*/g, '');
        const isCompany = name.includes('Internship') || name.includes('Company') || name.includes('-') || /^[A-Z]/.test(name);
        children.push(institutionLine(name, '', rest, isCompany));
      } else {
        // Single part with dates embedded — bold italic line
        children.push(new Paragraph({
          spacing: { line: 240, lineRule: 'auto', after: 0 },
          children: [mkRun(trimmed.replace(/\*/g, ''), { bold: true, italics: true, size: 23 })]
        }));
      }
      continue;
    }

    // Short non-bullet lines that look like headings (institution names, degree names)
    if (trimmed.length < 80 && !trimmed.includes(':')) {
      const nextLine = lines[i + 1]?.trim() || '';
      const isBoldCandidate = nextLine.startsWith('-') || nextLine.toLowerCase().startsWith('course') || nextLine.toLowerCase().startsWith('final') || nextLine.toLowerCase().startsWith('high school') || nextLine === '';
      if (isBoldCandidate || trimmed.length < 50) {
        children.push(new Paragraph({
          spacing: { line: 240, lineRule: 'auto', after: 0 },
          children: [mkRun(trimmed.replace(/\*/g, ''), { bold: true, size: 23 })]
        }));
        continue;
      }
    }

    // Default body line
    children.push(bodyLine(trimmed.replace(/\*/g, '')));
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 }, // US Letter
          margin: MARGIN
        }
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

function buildCoverLetterDoc(coverText) {
  const lines = coverText.split('\n');
  const children = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      children.push(emptyLine());
      continue;
    }
    children.push(new Paragraph({
      spacing: { line: 240, lineRule: 'auto', after: 80 },
      children: [mkRun(trimmed, { size: 21 })]
    }));
  }

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
        }
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

module.exports = { parseResumeToDocx: parseAndBuild, parseCoverLetterToDocx: buildCoverLetterDoc };
