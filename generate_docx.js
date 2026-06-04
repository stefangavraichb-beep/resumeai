const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, LevelFormat, TabStopType, TabStopPosition,
  UnderlineType
} = require('docx');

// Replicates Karina's CV format exactly:
// - Name: bold, centered, slightly larger
// - Contact: centered, normal
// - Section headers: bold, with right-aligned location via tab stop, underlined bottom border
// - Company/role: bold italic with date on right via tab
// - Bullet points: dash style
// - Skills: bold label + normal text

function buildCVDoc(resumeText) {
  const lines = resumeText.split('\n');
  const children = [];

  const numbering = {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '-',
        alignment: AlignmentType.LEFT,
        style: {
          paragraph: {
            indent: { left: 360, hanging: 360 },
            spacing: { after: 40 }
          }
        }
      }]
    }]
  };

  // Page width A4 with 1 inch margins = 9026 DXA content width
  const PAGE_WIDTH = 9026;

  let nameAdded = false;
  let contactAdded = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    if (!trimmed) {
      children.push(new Paragraph({ children: [new TextRun('')], spacing: { after: 60 } }));
      continue;
    }

    // First non-empty line = Name
    if (!nameAdded) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [new TextRun({ text: trimmed, bold: true, size: 28, font: 'Calibri' })]
      }));
      nameAdded = true;
      continue;
    }

    // Second non-empty line = contact info
    if (!contactAdded && (trimmed.includes('@') || trimmed.includes('•') || trimmed.includes('+'))) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [new TextRun({ text: trimmed, size: 20, font: 'Calibri' })]
      }));
      contactAdded = true;
      continue;
    }

    // Section headers — ALL CAPS or known section names, bold with bottom border
    const sectionNames = ['EDUCATION', 'WORK EXPERIENCE', 'EXTRACURRICULAR', 'SKILLS', 'INTERESTS', 'LANGUAGES', 'SKILLS AND INTERESTS', 'ACTIVITIES'];
    const isSection = sectionNames.some(s => trimmed.toUpperCase().includes(s)) ||
      (trimmed === trimmed.toUpperCase() && trimmed.length > 3 && !trimmed.startsWith('-'));

    if (isSection) {
      children.push(new Paragraph({
        spacing: { before: 200, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
        children: [new TextRun({ text: trimmed, bold: true, size: 22, font: 'Calibri' })]
      }));
      continue;
    }

    // Bullet points
    if (trimmed.startsWith('-') || trimmed.startsWith('•')) {
      const text = trimmed.replace(/^[-•]\s*/, '');
      children.push(new Paragraph({
        numbering: { reference: 'bullets', level: 0 },
        spacing: { after: 40 },
        children: [new TextRun({ text, size: 20, font: 'Calibri' })]
      }));
      continue;
    }

    // Lines with tab or multiple spaces — role/company with date on right
    if (trimmed.includes('\t') || (trimmed.includes('  ') && (trimmed.includes('20') || trimmed.includes('January') || trimmed.includes('August') || trimmed.includes('September')))) {
      const parts = trimmed.split(/\t+|\s{3,}/);
      const left = parts[0]?.trim() || '';
      const right = parts[parts.length - 1]?.trim() || '';

      const isBoldItalic = left.includes('*') || i > 0;

      children.push(new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: PAGE_WIDTH }],
        spacing: { before: 80, after: 40 },
        children: [
          new TextRun({ text: left.replace(/\*/g, ''), bold: true, italics: true, size: 20, font: 'Calibri' }),
          new TextRun({ text: '\t', size: 20 }),
          new TextRun({ text: right.replace(/\*/g, ''), italics: true, size: 20, font: 'Calibri' })
        ]
      }));
      continue;
    }

    // Bold lines (short, likely role/institution names)
    if (trimmed.length < 100 && !trimmed.includes(':') && i < lines.length) {
      const nextLine = lines[i + 1]?.trim() || '';
      const looksLikeHeader = nextLine.startsWith('-') || nextLine.toLowerCase().includes('course') || nextLine.toLowerCase().includes('final');

      if (looksLikeHeader || trimmed.length < 60) {
        children.push(new Paragraph({
          spacing: { before: 80, after: 40 },
          children: [new TextRun({ text: trimmed, bold: true, size: 20, font: 'Calibri' })]
        }));
        continue;
      }
    }

    // Skills/Languages lines with bold label
    if (trimmed.includes(':')) {
      const colonIdx = trimmed.indexOf(':');
      const label = trimmed.substring(0, colonIdx + 1);
      const value = trimmed.substring(colonIdx + 1);
      children.push(new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: label, bold: true, size: 20, font: 'Calibri' }),
          new TextRun({ text: value, size: 20, font: 'Calibri' })
        ]
      }));
      continue;
    }

    // Default paragraph
    children.push(new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: trimmed, size: 20, font: 'Calibri' })]
    }));
  }

  const doc = new Document({
    numbering,
    styles: {
      default: {
        document: { run: { font: 'Calibri', size: 20 } }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
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
      children.push(new Paragraph({ children: [new TextRun('')], spacing: { after: 120 } }));
      continue;
    }
    children.push(new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: trimmed, size: 22, font: 'Calibri' })]
    }));
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

module.exports = { parseResumeToDocx: buildCVDoc, parseCoverLetterToDocx: buildCoverLetterDoc };
