const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, BorderStyle, LevelFormat } = require('docx');
const fs = require('fs');

function parseResumeToDocx(resumeText, filename) {
  const lines = resumeText.split('\n').filter(l => l.trim());
  const children = [];

  // Numbering config for bullets
  const numbering = {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }]
  };

  let nameAdded = false;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      children.push(new Paragraph({ children: [new TextRun('')] }));
      continue;
    }

    // First line = name (big heading)
    if (!nameAdded) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: trimmed, bold: true, size: 36, font: 'Arial' })]
      }));
      nameAdded = true;
      continue;
    }

    // Contact line (contains @ or |)
    if (trimmed.includes('@') || (trimmed.includes('|') && trimmed.length < 120)) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: trimmed, size: 20, font: 'Arial', color: '444444' })]
      }));
      continue;
    }

    // Section headers (ALL CAPS or ends with :)
    if (trimmed === trimmed.toUpperCase() && trimmed.length > 2 && !trimmed.startsWith('-') && !trimmed.startsWith('•')) {
      children.push(new Paragraph({
        spacing: { before: 240, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2563eb', space: 1 } },
        children: [new TextRun({ text: trimmed, bold: true, size: 24, font: 'Arial', color: '2563eb' })]
      }));
      continue;
    }

    // Bullet points
    if (trimmed.startsWith('-') || trimmed.startsWith('•')) {
      const text = trimmed.replace(/^[-•]\s*/, '');
      children.push(new Paragraph({
        numbering: { reference: 'bullets', level: 0 },
        children: [new TextRun({ text, size: 20, font: 'Arial' })]
      }));
      continue;
    }

    // Bold lines (job titles / company names - shorter lines)
    if (trimmed.length < 80 && !trimmed.includes(',')) {
      children.push(new Paragraph({
        spacing: { before: 120 },
        children: [new TextRun({ text: trimmed, bold: true, size: 22, font: 'Arial' })]
      }));
      continue;
    }

    // Regular paragraph
    children.push(new Paragraph({
      children: [new TextRun({ text: trimmed, size: 20, font: 'Arial' })]
    }));
  }

  const doc = new Document({
    numbering,
    styles: {
      default: { document: { run: { font: 'Arial', size: 20 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
        }
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

function parseCoverLetterToDocx(coverText) {
  const lines = coverText.split('\n');
  const children = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      children.push(new Paragraph({ children: [new TextRun('')], spacing: { after: 120 } }));
      continue;
    }
    children.push(new Paragraph({
      children: [new TextRun({ text: trimmed, size: 22, font: 'Arial' })],
      spacing: { after: 80 }
    }));
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
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

module.exports = { parseResumeToDocx, parseCoverLetterToDocx };
