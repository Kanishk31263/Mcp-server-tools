/**
 * PPTX Generator Module
 * Core presentation generation logic with enhanced features
 */

import fs from 'fs';
import path from 'path';
import PptxGenJS from 'pptxgenjs';
import matter from 'gray-matter';
import AdmZip from 'adm-zip';


export function parseMarkdown(content) {
  const { data: frontmatter, content: body } = matter(content);
  const slides = [];
  const sections = body.split(/^## /m).filter(s => s.trim());

  for (const section of sections) {
    const lines = section.trim().split('\n');
    const headerLine = lines[0];
    const slideContent = lines.slice(1).join('\n').trim();

    const typeMatch = headerLine.match(/^\[(\w+)\]\s*(.*)$/);
    let type = 'content';
    let title = headerLine;

    if (typeMatch) {
      type = typeMatch[1].toLowerCase();
      title = typeMatch[2];
    }

    slides.push({ type, title, content: slideContent });
  }

  return { frontmatter, slides };
}

// Parse bullet points from content
export function parseBullets(content) {
  const bullets = [];
  const lines = content.split('\n');

  for (const line of lines) {
    const match = line.match(/^(\s*)[-*•]\s+(.+)$/);
    if (match) {
      const indent = Math.floor(match[1].length / 2);
      bullets.push({ text: match[2], level: indent });
    }
  }

  return bullets;
}

// Parse code blocks from content
export function parseCodeBlocks(content) {
  const codeBlocks = [];
  const regex = /```(\w*)\r?\n([\s\S]*?)```/g;
  let match;

  while ((match = regex.exec(content)) !== null) {
    codeBlocks.push({
      language: match[1] || 'text',
      code: match[2].trim()
    });
  }

  return codeBlocks;
}

// Parse markdown tables
export function parseTables(content) {
  const tables = [];
  const lines = content.split('\n');
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
      const tableLines = [];
      const tableStartIdx = i;

      while (i < lines.length && lines[i].trim().startsWith('|')) {
        tableLines.push(lines[i]);
        i++;
      }

      if (tableLines.length >= 2) {
        const rows = [];
        let isHeader = true;

        for (const tableLine of tableLines) {
          if (/^\|[\s:-]+\|/.test(tableLine) && tableLine.includes('---')) {
            isHeader = false;
            continue;
          }

          const cells = tableLine.split('|').slice(1, -1).map(cell => cell.trim());

          if (cells.length > 0) {
            rows.push({ cells, isHeader: isHeader });
            isHeader = false;
          }
        }

        if (rows.length > 0) {
          tables.push({ rows, startLine: tableStartIdx });
        }
      }
    } else {
      i++;
    }
  }

  return tables;
}

// Parse content into structured elements (text, bullets, code, tables)
export function parseContentElements(content) {
  const elements = [];

  let workingContent = content;
  const codeBlocks = [];
  const tables = [];

  // Extract code blocks
  const codeRegex = /```(\w*)\r?\n([\s\S]*?)```/g;
  let match;
  let placeholderIndex = 0;

  while ((match = codeRegex.exec(content)) !== null) {
    const placeholder = `__CODE_BLOCK_${placeholderIndex}__`;
    codeBlocks.push({
      placeholder,
      language: match[1] || 'text',
      code: match[2].trim()
    });
    workingContent = workingContent.replace(match[0], placeholder);
    placeholderIndex++;
  }

  // Extract tables
  const lines = workingContent.split('\n');
  let i = 0;
  const newLines = [];

  while (i < lines.length) {
    const line = lines[i];

    if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
      const tableLines = [];

      while (i < lines.length && lines[i].trim().startsWith('|')) {
        tableLines.push(lines[i]);
        i++;
      }

      if (tableLines.length >= 2) {
        const placeholder = `__TABLE_${tables.length}__`;
        const rows = [];
        let isHeader = true;

        for (const tableLine of tableLines) {
          if (/^\|[\s:-]+\|/.test(tableLine) && tableLine.includes('---')) {
            isHeader = false;
            continue;
          }

          const cells = tableLine.split('|').slice(1, -1).map(cell => cell.trim());

          if (cells.length > 0) {
            rows.push({ cells, isHeader: isHeader });
            isHeader = false;
          }
        }

        if (rows.length > 0) {
          tables.push({ placeholder, rows });
          newLines.push(placeholder);
        }
      }
    } else {
      newLines.push(line);
      i++;
    }
  }

  workingContent = newLines.join('\n');

  // Parse remaining content
  const finalLines = workingContent.split('\n');
  let currentBulletBlock = [];
  let currentTextBlock = [];

  for (const line of finalLines) {
    const codeMatch = line.match(/__CODE_BLOCK_(\d+)__/);
    if (codeMatch) {
      if (currentTextBlock.length > 0) {
        elements.push({ type: 'text', content: currentTextBlock.join('\n').trim() });
        currentTextBlock = [];
      }
      if (currentBulletBlock.length > 0) {
        elements.push({ type: 'bullets', content: currentBulletBlock.join('\n') });
        currentBulletBlock = [];
      }

      const codeBlock = codeBlocks.find(cb => cb.placeholder === line.trim());
      if (codeBlock) {
        elements.push({
          type: 'code',
          language: codeBlock.language,
          code: codeBlock.code
        });
      }
      continue;
    }

    const tableMatch = line.match(/__TABLE_(\d+)__/);
    if (tableMatch) {
      if (currentTextBlock.length > 0) {
        elements.push({ type: 'text', content: currentTextBlock.join('\n').trim() });
        currentTextBlock = [];
      }
      if (currentBulletBlock.length > 0) {
        elements.push({ type: 'bullets', content: currentBulletBlock.join('\n') });
        currentBulletBlock = [];
      }

      const table = tables.find(t => t.placeholder === line.trim());
      if (table) {
        elements.push({
          type: 'table',
          rows: table.rows
        });
      }
      continue;
    }

    if (/^\s*[-*•]\s+/.test(line)) {
      if (currentTextBlock.length > 0) {
        elements.push({ type: 'text', content: currentTextBlock.join('\n').trim() });
        currentTextBlock = [];
      }
      currentBulletBlock.push(line);
    } else if (line.trim()) {
      if (currentBulletBlock.length > 0) {
        elements.push({ type: 'bullets', content: currentBulletBlock.join('\n') });
        currentBulletBlock = [];
      }
      currentTextBlock.push(line);
    } else {
      if (currentTextBlock.length > 0) {
        elements.push({ type: 'text', content: currentTextBlock.join('\n').trim() });
        currentTextBlock = [];
      }
      if (currentBulletBlock.length > 0) {
        elements.push({ type: 'bullets', content: currentBulletBlock.join('\n') });
        currentBulletBlock = [];
      }
    }
  }

  if (currentTextBlock.length > 0) {
    elements.push({ type: 'text', content: currentTextBlock.join('\n').trim() });
  }
  if (currentBulletBlock.length > 0) {
    elements.push({ type: 'bullets', content: currentBulletBlock.join('\n') });
  }

  return elements;
}

// Fix for Google Slides compatibility
export function fixForGoogleSlides(pptxPath, outputPath) {
  const zip = new AdmZip(pptxPath);
  const entries = zip.getEntries();
  const imageRenames = new Map();
  let imageCounter = 1;

  entries.forEach(entry => {
    if (entry.entryName.startsWith('ppt/media/') && !entry.isDirectory) {
      const oldName = path.basename(entry.entryName);
      const ext = path.extname(oldName);
      const newName = `image${imageCounter}${ext}`;
      imageCounter++;

      if (oldName !== newName) {
        imageRenames.set(oldName, newName);
      }
    }
  });

  if (imageRenames.size === 0) {
    fs.copyFileSync(pptxPath, outputPath);
    return;
  }

  const newZip = new AdmZip();

  entries.forEach(entry => {
    let entryName = entry.entryName;
    let content = entry.getData();

    if (entryName.startsWith('ppt/media/')) {
      const oldName = path.basename(entryName);
      const newName = imageRenames.get(oldName);
      if (newName) entryName = `ppt/media/${newName}`;
    }

    if (entryName.endsWith('.xml') || entryName.endsWith('.rels')) {
      let text = content.toString('utf8');
      imageRenames.forEach((newName, oldName) => {
        text = text.split(oldName).join(newName);
      });
      content = Buffer.from(text, 'utf8');
    }

    newZip.addFile(entryName, content);
  });

  newZip.writeZip(outputPath);
}

// Create presentation instance
function createPresentation(config) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = config.instructor.name;
  pptx.company = config.institution.name[0];
   
  pptx.defineSlideMaster({
    title: 'TITLE_SLIDE',
    // Main red background
    background: {
        color: config.colors.red.replace('#', '') // your red color
    },
    objects: [
        // Left pink vertical strip
        {
            rect: {
                x: 0,           // left edge
                y: 0,
                w: 0.3,         // thickness of strip
                h: '100%',      // full height
                fill: { color: config.colors.pink.replace('#', '') }
            }
        },
        // Bottom pink horizontal strip
        {
            rect: {
                x: 0,
                y: '95%',        // bottom position
                w: '100%',       // full width
                h: '5%',         // height of strip
                fill: { color: config.colors.pink.replace('#', '') }
            }
        }
    ]
});


  pptx.defineSlideMaster({
    title: 'CONTENT_SLIDE',
    background: { color: 'FFFFFF' }
  });

  pptx.defineSlideMaster({
    title: 'DIVIDER_SLIDE',
 // Main red background
    background: {
        color: config.colors.red.replace('#', '') // your red color
    },
    objects: [
        // Left pink vertical strip
        {
            rect: {
                x: 0,           // left edge
                y: 0,
                w: 0.3,         // thickness of strip
                h: '100%',      // full height
                fill: { color: config.colors.pink.replace('#', '') }
            }
        },
        // Bottom pink horizontal strip
        {
            rect: {
                x: 0,
                y: '95%',        // bottom position
                w: '100%',       // full width
                h: '5%',         // height of strip
                fill: { color: config.colors.pink.replace('#', '') }
            }
        }
    ]  });

  return pptx;
}
//  "primary": "#000000",
//         "secondary": "#F15A22",
//         "lightGray": "#EDEDED",
//         "darkGray": "#4D4D4D",
//         "white": "#FFFFFF",
//         "titleText": "#000000",
//         "bodyText": "#333333",
//         "codeBackground": "#F5F5F5"

// Add title slide
function addTitleSlide(pptx, config, frontmatter, baseDir) {
  const slide = pptx.addSlide({ masterName: 'TITLE_SLIDE' });
  const lessonType = config.lessonTypes[frontmatter.type] || config.lessonTypes.lecture;

  slide.addText(config.institution.name.join('\n'), {
    x: 0, y: 0.67, w: '100%', h: 1.2,
    align: 'center',
    fontSize: config.sizes.institutionName,
    fontFace: config.fonts.title,
    bold: true,
    color: 'FFFFFF',
    italic:true
  });

  let infoText = `Discipline: ${frontmatter.discipline || 'Course Name'}\n\n`;
  infoText += `${lessonType}\n\n`;
  if (frontmatter.module) infoText += `Module ${frontmatter.module}\n`;
  infoText += `Lesson ${frontmatter.lesson || '1.1: Lesson Title'}`;

  slide.addText(infoText, {
    x: 0.3, y: 1.85, w: 9.2, h: 2.2,
    align: 'center', fontSize: config.sizes.lessonInfo,
    fontFace: config.fonts.body, color: 'FFFFFF', valign: 'middle'
  });

  const instructorPosition = `${config.instructor.position[0].toUpperCase()}${config.instructor.position.slice(1)}`;
  const instructorText = `${instructorPosition} ${config.institution.department}\n${config.instructor.rank}\t${config.instructor.name}`;
  slide.addText(instructorText, {
    x: 5, y: 4.2, w: 4.5, h: 0.8,
    align: 'left', fontSize: config.sizes.footer,
    fontFace: config.fonts.body, bold: true, color: 'FFFFFF'
  });

  const logoPath = path.join(baseDir, config.logo.path);
  if (fs.existsSync(logoPath)) {
    const logoData = fs.readFileSync(logoPath).toString('base64');
    const pos = config.logo.titleSlide;
    slide.addImage({ data: `image/png;base64,${logoData}`, x: pos.x, y: pos.y, w: pos.w, h: pos.h });
  }
}


// Add plan slide
function addPlanSlide(pptx, config, title, content) {
  const slide = pptx.addSlide({ masterName: 'CONTENT_SLIDE' });
  const bullets = parseBullets(content);

  slide.addText(title || 'План заняття', {
    x: 0.5, y: 0.2, w: 9, h: 0.6,
    fontSize: config.sizes.slideTitle, fontFace: config.fonts.title,
    bold: true, color: config.colors.titleText.replace('#', '')
  });

  const bulletRows = bullets.map(b => ({
    text: b.text,
    options: { bullet: { type: 'bullet', code: '2022' }, indentLevel: b.level, paraSpaceBefore: 8, paraSpaceAfter: 4 }
  }));

  slide.addText(bulletRows, {
    x: 0.5, y: 0.9, w: 9, h: 4,
    fontSize: config.sizes.body, fontFace: config.fonts.body,
    color: config.colors.bodyText.replace('#', ''), valign: 'top'
  });
}

// Add divider slide
function addDividerSlide(pptx, config, title) {
  const slide = pptx.addSlide({ masterName: 'DIVIDER_SLIDE' });

  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: '100%', h: '100%',
    // fill: { color: config.colors.secondary.replace('#', '') },
    shadow: { type: 'outer', blur: 4, offset: 2, angle: 90, color: '000000', opacity: 0.35 }
  });

  slide.addText(title, {
    x: 0.5, y: 2, w: 9, h: 1,
    align: 'center', fontSize: config.sizes.dividerTitle,
    fontFace: config.fonts.title, bold: true, color: 'FFFFFF'
  });
}

// Add content slide with mixed content (bullets, text, code, tables)
function addContentSlide(pptx, config, title, content) {
  const slide = pptx.addSlide({ masterName: 'CONTENT_SLIDE' });
  // HEADER HEIGHT EXACT MATCH (0.85 inches)
  const HEADER_HEIGHT = 0.85;

  // LEFT PwC Orange (#D1402B)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0,
    y: 0,
    w: 8,         // slightly wider left part
    h: HEADER_HEIGHT,
    fill: { color: "D1402B" },
    line: { color: "D1402B" }
  });

  // RIGHT PwC Pink (#C44F6D)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 8,
    y: 0,
    w: 2,         // remainder of the slide
    h: HEADER_HEIGHT,
    fill: { color: "C44F6D" },
    line: { color: "C44F6D" }
  });
  slide.addText(title, {
    x: 0,
    y: 0.15,                // placed INSIDE the bar
    w: 10,
    h: 0.7,
    align: "center",
    fontSize: 24,           // PwC typical title size
    fontFace: "Georgia",
    italic: true,
    bold: false,
    color: "FFFFFF"
  });

  const elements = parseContentElements(content);

  let yPos = 0.9;
  const maxY = 5.0;

  for (const element of elements) {
    if (yPos >= maxY) break;

    switch (element.type) {
      case 'bullets': {
        const bullets = parseBullets(element.content);
        if (bullets.length > 0) {
          const bulletRows = [];

          for (const b of bullets) {
            const text = b.text;

            if (text.includes('**') || text.includes('`')) {
              const parts = [];
              const regex = /(\*\*(.+?)\*\*|`(.+?)`)/g;
              let currentPos = 0;
              let match;

              while ((match = regex.exec(text)) !== null) {
                if (match.index > currentPos) {
                  parts.push({
                    text: text.substring(currentPos, match.index),
                    options: {}
                  });
                }

                if (match[0].startsWith('**')) {
                  parts.push({
                    text: match[2],
                    options: { bold: true }
                  });
                } else if (match[0].startsWith('`')) {
                  parts.push({
                    text: match[3],
                    options: {
                      fontFace: config.fonts.code,
                      color: '172B4D'
                    }
                  });
                }

                currentPos = match.index + match[0].length;
              }

              if (currentPos < text.length) {
                parts.push({
                  text: text.substring(currentPos),
                  options: {}
                });
              }

              // Each bullet item needs separate addText call for proper formatting
              bulletRows.push({
                text: parts,
                options: {
                  bullet: { type: 'bullet', code: '2022' },
                  indentLevel: b.level,
                  paraSpaceBefore: 6,
                  paraSpaceAfter: 3
                }
              });
            } else {
              bulletRows.push({
                text: text,
                options: {
                  bullet: { type: 'bullet', code: '2022' },
                  indentLevel: b.level,
                  paraSpaceBefore: 6,
                  paraSpaceAfter: 3
                }
              });
            }
          }

          const estimatedHeight = Math.min(bullets.length * 0.4, maxY - yPos);

          // Format for PptxGenJS with mixed formatting
          const formattedBullets = bulletRows.map(item => {
            if (Array.isArray(item.text)) {
              // For formatted text, merge base options with item options
              return item.text.map((part, idx) => {
                if (idx === 0) {
                  // First part gets bullet options
                  return {
                    text: part.text,
                    options: {
                      ...item.options,
                      ...part.options
                    }
                  };
                } else {
                  // Subsequent parts don't get bullet
                  return {
                    text: part.text,
                    options: part.options
                  };
                }
              });
            } else {
              // Simple text
              return {
                text: item.text,
                options: item.options
              };
            }
          }).flat();

          slide.addText(formattedBullets, {
            x: 0.5, y: yPos, w: 9, h: estimatedHeight,
            fontSize: config.sizes.body,
            fontFace: config.fonts.body,
            color: config.colors.bodyText.replace('#', ''),
            valign: 'top',
            align: 'left'
          });

          yPos += estimatedHeight + 0.15;
        }
        break;
      }

      case 'text': {
        const textLines = element.content.split('\n');
        let textContent = [];

        for (const line of textLines) {
          if (line.includes('**') || line.includes('`')) {
            const parts = [];
            const regex = /(\*\*(.+?)\*\*|`(.+?)`)/g;
            let currentPos = 0;
            let match;

            while ((match = regex.exec(line)) !== null) {
              if (match.index > currentPos) {
                parts.push({
                  text: line.substring(currentPos, match.index),
                  options: {}
                });
              }

              if (match[0].startsWith('**')) {
                parts.push({
                  text: match[2],
                  options: { bold: true }
                });
              } else if (match[0].startsWith('`')) {
                parts.push({
                  text: match[3],
                  options: {
                    fontFace: config.fonts.code,
                    color: '172B4D'
                  }
                });
              }

              currentPos = match.index + match[0].length;
            }

            if (currentPos < line.length) {
              parts.push({
                text: line.substring(currentPos),
                options: {}
              });
            }

            if (textLines.indexOf(line) < textLines.length - 1) {
              parts.push({ text: '\n', options: {} });
            }

            textContent = textContent.concat(parts);
          } else {
            textContent.push({
              text: line + (textLines.indexOf(line) < textLines.length - 1 ? '\n' : ''),
              options: {}
            });
          }
        }

        const textHeight = Math.min(1.2, maxY - yPos);
        slide.addText(textContent, {
          x: 0.5, y: yPos, w: 9, h: textHeight,
          fontSize: config.sizes.body,
          fontFace: config.fonts.body,
          color: config.colors.bodyText.replace('#', ''),
          valign: 'top'
        });
        yPos += textHeight + 0.1;
        break;
      }

      case 'code': {
        const codeLines = element.code.split('\n').length;
        const codeHeight = Math.min(Math.max(codeLines * 0.28, 1.2), maxY - yPos - 0.3);

        slide.addShape(pptx.shapes.RECTANGLE, {
          x: 0.4, y: yPos - 0.05, w: 9.2, h: codeHeight + 0.1,
          fill: { color: 'EBECF0' },
          line: { color: 'CCCCCC', width: 1 }
        });

        slide.addText(element.code, {
          x: 0.5, y: yPos, w: 9, h: codeHeight,
          fontSize: config.sizes.code,
          fontFace: config.fonts.code,
          color: '172B4D',
          valign: 'top',
          wrap: false
        });

        yPos += codeHeight + 0.25;
        break;
      }

      case 'table': {
        const tableData = element.rows.map((row, idx) => {
          return row.cells.map(cell => {
            if (cell.includes('`')) {
              const parts = [];
              const regex = /`(.+?)`/g;
              let currentPos = 0;
              let match;

              while ((match = regex.exec(cell)) !== null) {
                if (match.index > currentPos) {
                  parts.push({
                    text: cell.substring(currentPos, match.index),
                    options: {}
                  });
                }

                parts.push({
                  text: match[1],
                  options: {
                    fontFace: config.fonts.code
                  }
                });

                currentPos = match.index + match[0].length;
              }

              if (currentPos < cell.length) {
                parts.push({
                  text: cell.substring(currentPos),
                  options: {}
                });
              }

              return {
                text: parts,
                options: {
                  fontSize: row.isHeader ? 16 : 14,
                  bold: row.isHeader,
                  fill: row.isHeader ? config.colors.secondary.replace('#', '') : 'FFFFFF',
                  color: row.isHeader ? 'FFFFFF' : config.colors.bodyText.replace('#', ''),
                  align: 'center',
                  valign: 'middle'
                }
              };
            } else {
              return {
                text: cell,
                options: {
                  fontSize: row.isHeader ? 16 : 14,
                  bold: row.isHeader,
                  fill: row.isHeader ? config.colors.secondary.replace('#', '') : 'FFFFFF',
                  color: row.isHeader ? 'FFFFFF' : config.colors.bodyText.replace('#', ''),
                  align: 'center',
                  valign: 'middle'
                }
              };
            }
          });
        });

        const colCount = element.rows[0].cells.length;
        const rowCount = element.rows.length;
        const tableWidth = 9;
        const colWidth = tableWidth / colCount;
        const rowHeight = 0.4;
        const tableHeight = Math.min(rowCount * rowHeight, maxY - yPos - 0.1);

        slide.addTable(tableData, {
          x: 0.5,
          y: yPos,
          w: tableWidth,
          colW: Array(colCount).fill(colWidth),
          rowH: Array(rowCount).fill(rowHeight),
          border: { pt: 1, color: 'CCCCCC' }
        });

        yPos += tableHeight + 0.2;
        break;
      }
    }
  }

  return slide;
}


// Add closing slide
function addClosingSlide(pptx, config, baseDir) {
  const slide = pptx.addSlide({ masterName: 'TITLE_SLIDE' });

  slide.addText('Thank you for your attention', {
    x: 0, y: 2.2, w: '100%', h: 0.6,
    align: 'center', fontSize: config.sizes.lessonInfo,
    fontFace: config.fonts.body, color: 'FFFFFF'
  });

  const institutionText = `${config.institution.name[0]}\n${config.institution.name.slice(1).join(' ')}`;
  slide.addText(institutionText, {
    x: 0.6, y: 4.35, w: 4, h: 0.8,
    align: 'left', fontSize: 14, fontFace: config.fonts.body,
    bold: true, color: 'FFFFFF'
  });

  const logoPath = path.join(baseDir, config.logo.path);
  if (fs.existsSync(logoPath)) {
    const logoData = fs.readFileSync(logoPath).toString('base64');
    const logoDataUri = `image/png;base64,${logoData}`;

    const mainPos = config.logo.closingSlide;
    slide.addImage({ data: logoDataUri, x: mainPos.x, y: mainPos.y, w: mainPos.w, h: mainPos.h });

    const smallPos = config.logo.closingSmall;
    slide.addImage({ data: logoDataUri, x: smallPos.x, y: smallPos.y, w: smallPos.w, h: smallPos.h });
  }
}

/**
 * Main generation function
 * @param {string} markdown - Markdown content with frontmatter
 * @param {string} outputPath - Full path for output file
 * @param {string} baseDir - Base directory for config and assets
 * @returns {Promise<string>} - Path to generated file
 */
export async function generatePresentation(markdown, outputPath, baseDir) {
  const configPath = path.join(baseDir, 'config.json');
  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));

  const { frontmatter, slides } = parseMarkdown(markdown);
  const pptx = createPresentation(config);

  addTitleSlide(pptx, config, frontmatter, baseDir);

  for (const slide of slides) {
    switch (slide.type) {
      case 'plan':
        addPlanSlide(pptx, config, slide.title, slide.content);
        break;
      case 'divider':
        addDividerSlide(pptx, config, slide.title);
        break;
      case 'code':
        addContentSlide(pptx, config, slide.title, slide.content);
        break;
      case 'content':
      default:
        addContentSlide(pptx, config, slide.title, slide.content);
        break;
    }
  }

  addClosingSlide(pptx, config, baseDir);

  const outputDir = path.dirname(outputPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const tempPath = outputPath.replace('.pptx', '-temp.pptx');
  await pptx.writeFile({ fileName: tempPath });
  fixForGoogleSlides(tempPath, outputPath);
  fs.unlinkSync(tempPath);

  return outputPath;
}
