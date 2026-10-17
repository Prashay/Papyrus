import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy, PDFPageProxy, PageViewport } from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import type { PdfAnnotation } from '../types';

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

export interface TextItemBox {
  str: string;
  x: number;
  y: number;
  width: number;
  height: number;
  start: number;
  end: number;
  fontName?: string;
  fontFamily?: string;
  fontSize?: number;
  serif?: boolean;
  bold?: boolean;
  italic?: boolean;
}

export interface PageText {
  text: string;
  items: TextItemBox[];
}

export { pdfjsLib };
export type { PDFDocumentProxy, PDFPageProxy, PageViewport };

export async function loadPdf(bytes: ArrayBuffer): Promise<PDFDocumentProxy> {
  const data = new Uint8Array(bytes.slice(0));
  const task = pdfjsLib.getDocument({ data, isEvalSupported: false });
  return task.promise;
}

export function renderPage(
  page: PDFPageProxy,
  canvas: HTMLCanvasElement,
  scale: number,
): Promise<PageViewport> {
  const viewport = page.getViewport({ scale });
  const ratio = Math.min(window.devicePixelRatio || 1, 2);
  canvas.width = Math.max(1, Math.floor(viewport.width * ratio));
  canvas.height = Math.max(1, Math.floor(viewport.height * ratio));
  canvas.style.width = `${viewport.width}px`;
  canvas.style.height = `${viewport.height}px`;

  const context = canvas.getContext('2d');
  if (!context) return Promise.resolve(viewport);

  const renderTask = page.render({
    canvasContext: context,
    viewport,
    transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined,
  } as Parameters<PDFPageProxy['render']>[0]);

  return renderTask.promise.then(() => viewport);
}

function viewportText(page: PDFPageProxy, textContent: Awaited<ReturnType<PDFPageProxy['getTextContent']>>): PageText {
  const viewport = page.getViewport({ scale: 1 });
  const items: TextItemBox[] = [];
  let text = '';

  for (const raw of textContent.items) {
    if (!('str' in raw)) continue;
    const item = raw;
    if (!item.str) continue;

    const transform = pdfjsLib.Util.transform(viewport.transform, item.transform);
    const height = Math.hypot(transform[2], transform[3]);
    const x = transform[4];
    const y = transform[5] - height;

    const fontName = item.fontName || '';
    const style = (textContent.styles as Record<string, { fontFamily?: string } | undefined>)?.[fontName];
    const family = (style?.fontFamily || '').toLowerCase();
    const nameLower = fontName.toLowerCase();

    const isSerif =
      /times|georgia|serif|roman|garamond|cambria|palatino|minion|baskerville|didot|caslon|bodoni/i.test(nameLower) ||
      /times|georgia|serif|roman|garamond|cambria|palatino/i.test(family);

    const isBold =
      /bold|black|heavy|700|semibold|medium|dark|demi/i.test(nameLower) ||
      /bold|700/i.test(family);

    const isItalic =
      /italic|oblique|slant/i.test(nameLower) ||
      /italic|oblique/i.test(family);

    if (text.length > 0) {
      const last = text[text.length - 1];
      if (!/\s/.test(last) && !/^\s/.test(item.str)) text += ' ';
    }

    const start = text.length;
    text += item.str;
    items.push({
      str: item.str,
      x,
      y,
      width: Math.max(item.width, 1),
      height: Math.max(height, 1),
      start,
      end: text.length,
      fontName,
      fontFamily: style?.fontFamily,
      fontSize: height,
      serif: isSerif,
      bold: isBold,
      italic: isItalic,
    });

    if (item.hasEOL) text += '\n';
  }

  return { text, items };
}

export async function extractPageText(page: PDFPageProxy): Promise<PageText> {
  const content = await page.getTextContent();
  return viewportText(page, content);
}

export async function extractDocumentText(pdf: PDFDocumentProxy): Promise<PageText[]> {
  const pages: PageText[] = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    pages.push(await extractPageText(page));
  }
  return pages;
}

export async function exportAnnotatedPdf(
  originalBytes: ArrayBuffer,
  annotations: PdfAnnotation[],
  title: string,
): Promise<Uint8Array> {
  const pdf = await PDFDocument.load(originalBytes.slice(0), { ignoreEncryption: true });
  pdf.setTitle(title);
  pdf.setCreator('Papyrus Studio');
  pdf.setProducer('Papyrus Studio');
  pdf.setModificationDate(new Date());

  const helvetica = await pdf.embedFont(StandardFonts.Helvetica);
  const helveticaBold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const helveticaOblique = await pdf.embedFont(StandardFonts.HelveticaOblique);
  const helveticaBoldOblique = await pdf.embedFont(StandardFonts.HelveticaBoldOblique);
  const times = await pdf.embedFont(StandardFonts.TimesRoman);
  const timesBold = await pdf.embedFont(StandardFonts.TimesRomanBold);
  const timesItalic = await pdf.embedFont(StandardFonts.TimesRomanItalic);
  const timesBoldItalic = await pdf.embedFont(StandardFonts.TimesRomanBoldItalic);

  const safeText = (font: typeof helvetica, text: string): string => {
    try {
      font.encodeText(text);
      return text;
    } catch {
      return text.replace(/[^\u0020-\u007e\u00a0-\u00ff]/g, '?');
    }
  };

  const pages = pdf.getPages();
  for (const annotation of annotations) {
    const page = pages[annotation.page - 1];
    if (!page) continue;
    const { width: pageWidth } = page.getSize();
    const color = hexToRgb(annotation.color);

    if (annotation.type === 'whiteout') {
      const x = Math.min(annotation.x, annotation.x + annotation.width);
      const y = Math.min(annotation.y, annotation.y + annotation.height);
      page.drawRectangle({
        x,
        y,
        width: Math.abs(annotation.width),
        height: Math.abs(annotation.height),
        color: hexToRgb(annotation.color || '#ffffff'),
        opacity: 1,
        borderWidth: 0,
      });
    } else if (annotation.type === 'highlight') {
      const x = Math.min(annotation.x, annotation.x + annotation.width);
      const y = Math.min(annotation.y, annotation.y + annotation.height);
      page.drawRectangle({
        x,
        y,
        width: Math.abs(annotation.width),
        height: Math.abs(annotation.height),
        color,
        opacity: annotation.opacity,
        borderWidth: 0,
      });
    } else if (annotation.type === 'ink') {
      for (let index = 1; index < annotation.points.length; index += 1) {
        const previous = annotation.points[index - 1];
        const current = annotation.points[index];
        page.drawLine({
          start: { x: previous.x, y: previous.y },
          end: { x: current.x, y: current.y },
          thickness: Math.max(0.5, annotation.width),
          color,
          opacity: 1,
        });
      }
      if (annotation.points.length === 1) {
        const point = annotation.points[0];
        page.drawCircle({ x: point.x, y: point.y, size: Math.max(1, annotation.width / 2), color });
      }
    } else if (annotation.type === 'text') {
      const isBold = Boolean(annotation.bold);
      const isItalic = Boolean(annotation.italic);
      const font = annotation.serif
        ? (isBold
            ? (isItalic ? timesBoldItalic : timesBold)
            : (isItalic ? timesItalic : times))
        : (isBold
            ? (isItalic ? helveticaBoldOblique : helveticaBold)
            : (isItalic ? helveticaOblique : helvetica));
      const size = Math.max(6, annotation.size);
      const sanitized = safeText(font, annotation.text);
      const lines = wrapText(sanitized, font, size, Math.max(60, pageWidth - annotation.x - 36));
      const lineHeight = size * 1.3;
      const baselineTop = annotation.y - size * 0.86;

      if (annotation.background) {
        const textWidth = Math.max(
          ...lines.map((line) => {
            try {
              return font.widthOfTextAtSize(line, size);
            } catch {
              return line.length * size * 0.6;
            }
          }),
        );
        const textHeight = lines.length * lineHeight;
        page.drawRectangle({
          x: annotation.x - 2,
          y: baselineTop - (lines.length - 1) * lineHeight - 2,
          width: textWidth + 4,
          height: textHeight + 4,
          color: hexToRgb(annotation.background),
          opacity: 1,
          borderWidth: 0,
        });
      }

      lines.forEach((line, index) => {
        page.drawText(line, {
          x: annotation.x,
          y: baselineTop - index * lineHeight,
          size,
          font,
          color,
        });
      });
    }
  }

  return pdf.save();
}

export async function extractPdfAsHtml(pdf: PDFDocumentProxy, _title?: string): Promise<string> {
  let html = '';
  let totalSerifCount = 0;
  let totalItemCount = 0;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const viewport = page.getViewport({ scale: 1 });

    if (pageNumber > 1) {
      html += `<hr class="page-break" style="margin: 28px 0 20px; border: none; border-top: 2px dashed #ded6c7;" />`;
    }

    interface LineItem {
      str: string;
      x: number;
      y: number;
      height: number;
      width?: number;
      bold?: boolean;
      italic?: boolean;
      serif?: boolean;
    }

    const rawLines: { y: number; height: number; items: LineItem[] }[] = [];

    for (const raw of textContent.items) {
      if (!('str' in raw) || !raw.str) continue;
      const transform = pdfjsLib.Util.transform(viewport.transform, raw.transform);
      const height = Math.hypot(transform[2], transform[3]);
      const x = transform[4];
      const y = Math.round(transform[5]);

      const fontName = raw.fontName || '';
      const style = (textContent.styles as Record<string, { fontFamily?: string } | undefined>)?.[fontName];
      const family = (style?.fontFamily || '').toLowerCase();
      const nameLower = fontName.toLowerCase();

      const isSerif =
        /times|georgia|serif|roman|garamond|cambria|palatino|minion|baskerville|didot|caslon|bodoni/i.test(nameLower) ||
        /times|georgia|serif|roman|garamond|cambria|palatino/i.test(family);

      const isBold =
        /bold|black|heavy|700|semibold|medium|dark|demi/i.test(nameLower) ||
        /bold|700/i.test(family);

      const isItalic =
        /italic|oblique|slant/i.test(nameLower) ||
        /italic|oblique/i.test(family);

      totalItemCount += 1;
      if (isSerif) totalSerifCount += 1;

      const rawItem = raw as { width?: number };
      const itemWidth = typeof rawItem.width === 'number' && rawItem.width > 0 ? rawItem.width : raw.str.length * height * 0.52;

      const item: LineItem = {
        str: raw.str,
        x,
        y,
        height,
        width: itemWidth,
        serif: isSerif,
        bold: isBold,
        italic: isItalic,
      };

      const existing = rawLines.find((l) => Math.abs(l.y - y) <= 3);
      if (existing) {
        existing.items.push(item);
        existing.height = Math.max(existing.height, height);
      } else {
        rawLines.push({ y, height, items: [item] });
      }
    }

    rawLines.sort((a, b) => a.y - b.y);

    let activeList = false;
    let activeBulletContent = '';
    let currentParagraph = '';
    let currentAlign = '';
    let inSkillsSection = false;
    let activeSkillsTable = false;
    let lastY = 0;

    const flushParagraph = () => {
      if (currentParagraph) {
        html += `<p${currentAlign ? ` style="${currentAlign}"` : ''}>${currentParagraph}</p>`;
        currentParagraph = '';
        currentAlign = '';
      }
    };

    const flushBullet = () => {
      if (activeBulletContent) {
        html += `<li style="margin-bottom: 2px; font-size: 11px; line-height: 1.34;"><p>${activeBulletContent}</p></li>`;
        activeBulletContent = '';
      }
    };

    const flushList = () => {
      flushBullet();
      if (activeList) {
        html += `</ul>`;
        activeList = false;
      }
    };

    const flushSkillsTable = () => {
      if (activeSkillsTable) {
        html += `</tbody></table>`;
        activeSkillsTable = false;
      }
    };

    for (let lineIndex = 0; lineIndex < rawLines.length; lineIndex += 1) {
      const line = rawLines[lineIndex];
      // Separate non-empty text items from wide whitespace spacer items
      const nonWhitespaceItems = line.items.filter((i) => i.str.trim().length > 0);
      nonWhitespaceItems.sort((a, b) => a.x - b.x);

      const plainLine = nonWhitespaceItems.map((i) => i.str.trim()).join(' ').replace(/\s+/g, ' ').trim();
      if (!plainLine) continue;

      const pageWidth = viewport.width;
      const minX = Math.min(...nonWhitespaceItems.map((i) => i.x));
      const maxX = Math.max(...nonWhitespaceItems.map((i) => i.x + (i.width || 0)));
      const lineWidth = maxX - minX;

      // Format helper with bold & italic
      const formatItems = (items: LineItem[]) =>
        items
          .map((i) => {
            let p = escapeHtml(i.str);
            if (i.bold && i.italic) p = `<strong><em>${p}</em></strong>`;
            else if (i.bold) p = `<strong>${p}</strong>`;
            else if (i.italic) p = `<em>${p}</em>`;
            return p;
          })
          .join(' ')
          .replace(/\s+/g, ' ')
          .trim();

      const renderedLine = formatItems(nonWhitespaceItems);

      // Continuation of skills table column 2 (e.g. "Applications" or "Workflows" wrapped at x > 150)
      if (inSkillsSection && activeSkillsTable && minX > 150 && nonWhitespaceItems.length <= 4) {
        const marker = `</td></tr>`;
        const lastIdx = html.lastIndexOf(marker);
        if (lastIdx !== -1) {
          const before = html.slice(0, lastIdx - 4); // before </p>
          html = `${before} ${renderedLine}</p></td></tr>`;
          lastY = line.y;
          continue;
        }
      }

      // Check Section Header
      const hasSeparators = /[:|,;]/.test(plainLine);
      const isSectionKeyword =
        !hasSeparators &&
        plainLine.length < 35 &&
        /^(PROFESSIONAL SUMMARY|SELECTED IMPACT|PROFESSIONAL EXPERIENCE|EXPERIENCE|WORK EXPERIENCE|EDUCATION|TECHNICAL SKILLS|SKILLS|PROJECTS|CERTIFICATIONS|PUBLICATIONS|SUMMARY|AWARDS|AWARDS & RECOGNITION|AWARDS AND RECOGNITION|CONTACT|PROFILE|EMPLOYMENT HISTORY|CAREER SUMMARY)$/i.test(
          plainLine.trim(),
        );

      if (isSectionKeyword) {
        flushParagraph();
        flushList();
        flushSkillsTable();
        inSkillsSection = /SKILLS/i.test(plainLine);
        html += `<h2 style="color: #1e3a8a; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; margin: 12px 0 1px;">${renderedLine}</h2><hr class="section-rule" style="margin: 1px 0 6px; border: none; border-top: 1.5px solid #2563eb;" />`;
        lastY = line.y;
        continue;
      }

      // If we are in the skills section
      if (inSkillsSection) {
        let skillSplitGapIndex = -1;
        for (let i = 0; i < nonWhitespaceItems.length - 1; i += 1) {
          const itemEnd = nonWhitespaceItems[i].x + (nonWhitespaceItems[i].width || 0);
          const nextStart = nonWhitespaceItems[i + 1].x;
          if (nextStart - itemEnd > 30 && nextStart > 180) {
            skillSplitGapIndex = i;
            break;
          }
        }

        if (skillSplitGapIndex !== -1) {
          if (!activeSkillsTable) {
            html += `<table class="skills-table" style="width: 100%; border: none; margin: 4px 0 8px;"><tbody>`;
            activeSkillsTable = true;
          }
          const leftItems = nonWhitespaceItems.slice(0, skillSplitGapIndex + 1);
          const rightItems = nonWhitespaceItems.slice(skillSplitGapIndex + 1);
          const leftHtml = formatItems(leftItems);
          const rightHtml = formatItems(rightItems);
          html += `<tr><td style="border: none; padding: 2px 8px 2px 0; width: 140px; text-align: left; vertical-align: top;"><p><strong>${leftHtml}</strong></p></td><td style="border: none; padding: 2px 0; text-align: left; vertical-align: top; color: #1f2937;"><p>${rightHtml}</p></td></tr>`;
          lastY = line.y;
          continue;
        }
      }

      // If outside skills section, close any open skills table
      flushSkillsTable();
      inSkillsSection = false;

      // Candidate Name (Top of Page 1)
      const isCandidateHeader =
        pageNumber === 1 &&
        lineIndex <= 2 &&
        minX > 50 &&
        Math.abs(minX + lineWidth / 2 - pageWidth / 2) < 40 &&
        lineWidth < pageWidth * 0.85;

      if (pageNumber === 1 && lineIndex === 0 && (isCandidateHeader || line.height >= 16)) {
        flushParagraph();
        flushList();
        html += `<h1 style="text-align: center; color: #1e3a8a; font-size: 18px; font-weight: 700; margin: 0 0 4px;">${renderedLine}</h1>`;
        lastY = line.y;
        continue;
      }

      if (isCandidateHeader) {
        flushParagraph();
        flushList();
        html += `<p style="text-align: center; margin: 0 0 3px; font-size: 10px; color: #475569;">${renderedLine}</p>`;
        lastY = line.y;
        continue;
      }

      // Check Job Role + Date split (date is on the far right: nextStart > pageWidth * 0.60)
      let splitGapIndex = -1;
      for (let i = 0; i < nonWhitespaceItems.length - 1; i += 1) {
        const itemEnd = nonWhitespaceItems[i].x + (nonWhitespaceItems[i].width || 0);
        const nextStart = nonWhitespaceItems[i + 1].x;
        if (nextStart - itemEnd > 45 && nextStart > pageWidth * 0.60) {
          splitGapIndex = i;
          break;
        }
      }

      if (splitGapIndex !== -1) {
        flushParagraph();
        flushList();
        const leftItems = nonWhitespaceItems.slice(0, splitGapIndex + 1);
        const rightItems = nonWhitespaceItems.slice(splitGapIndex + 1);
        const leftHtml = formatItems(leftItems);
        const rightHtml = formatItems(rightItems);
        html += `<table class="job-split" style="width: 100%; border: none; margin: 6px 0 1px;"><tbody><tr><td style="border: none; padding: 0; text-align: left;"><p><strong>${leftHtml}</strong></p></td><td style="border: none; padding: 0; text-align: right; color: #475569; font-size: 10px;"><p><em>${rightHtml}</em></p></td></tr></tbody></table>`;
        lastY = line.y;
        continue;
      }

      // Project Subline under Job Title (e.g. "Projects: LOCI & THOR | Java 17...")
      const isProjectSubline = /^(projects?|technologies|tech stack)\s*:/i.test(plainLine);
      if (isProjectSubline) {
        flushParagraph();
        flushList();
        html += `<p style="margin: 1px 0 3px; font-size: 10px; color: #374151;"><em>${renderedLine}</em></p>`;
        lastY = line.y;
        continue;
      }

      // Bullet Point Item
      const bulletMatch = /^[•·▪●*–-]\s*(.*)$/.exec(plainLine);
      if (bulletMatch) {
        flushParagraph();
        flushBullet();
        if (!activeList) {
          html += `<ul style="margin: 3px 0 6px; padding-left: 18px;">`;
          activeList = true;
        }
        const cleanedContent = renderedLine.replace(/^[•·▪●*–-]\s*/, '');
        activeBulletContent = cleanedContent;
        lastY = line.y;
        continue;
      }

      // Bullet Continuation Line (wrapped lines of previous bullet)
      if (activeList && activeBulletContent) {
        activeBulletContent += (activeBulletContent.endsWith('-') ? '' : ' ') + renderedLine;
        lastY = line.y;
        continue;
      }

      // Normal Paragraph
      flushList();
      const verticalGap = lastY > 0 ? line.y - lastY : 0;
      if (verticalGap > line.height * 1.6 && currentParagraph) {
        flushParagraph();
        currentParagraph = renderedLine;
      } else {
        if (currentParagraph) {
          currentParagraph += (currentParagraph.endsWith('-') ? '' : ' ') + renderedLine;
        } else {
          currentParagraph = renderedLine;
        }
      }
      lastY = line.y;
    }

    flushParagraph();
    flushList();
    flushSkillsTable();
  }

  const isSans = totalItemCount > 0 ? totalSerifCount < totalItemCount / 2 : true;
  if (isSans) {
    return `<div class="font-sans">${html}</div>`;
  }
  return html || '<p></p>';
}

function escapeHtml(str: string): string {
  return str.replace(/[&<>'"]/g, (tag) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;',
  }[tag] || tag));
}

export async function exportHtmlToPdf(html: string, title: string): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  pdf.setTitle(title);
  pdf.setCreator('Papyrus Studio');
  pdf.setProducer('Papyrus Studio');
  pdf.setModificationDate(new Date());

  const helvetica = await pdf.embedFont(StandardFonts.Helvetica);
  const helveticaBold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const helveticaOblique = await pdf.embedFont(StandardFonts.HelveticaOblique);
  const times = await pdf.embedFont(StandardFonts.TimesRoman);
  const courier = await pdf.embedFont(StandardFonts.Courier);

  const timesBold = await pdf.embedFont(StandardFonts.TimesRomanBold);
  const timesItalic = await pdf.embedFont(StandardFonts.TimesRomanItalic);

  const safe = (font: typeof helvetica, text: string): string => {
    try {
      font.encodeText(text);
      return text;
    } catch {
      return text.replace(/[^\u0020-\u007e\u00a0-\u00ff]/g, '?');
    }
  };

  const isDocSans = html.includes('font-sans');
  const isLetter = isDocSans;
  const pageWidth = isLetter ? 612 : 595.28;
  const pageHeight = isLetter ? 792 : 841.89;
  const marginX = isLetter ? 36 : 54;
  const marginTop = isLetter ? 36 : 54;
  const marginBottom = isLetter ? 36 : 54;
  const contentWidth = pageWidth - marginX * 2;

  let currentPage = pdf.addPage([pageWidth, pageHeight]);
  let cursorY = pageHeight - marginTop;

  const ensureSpace = (neededHeight: number) => {
    if (cursorY - neededHeight < marginBottom) {
      currentPage = pdf.addPage([pageWidth, pageHeight]);
      cursorY = pageHeight - marginTop;
    }
  };

  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  function renderNode(element: Element) {
    const tagName = element.tagName.toLowerCase();
    const text = (element.textContent || '').trim();
    const styleAttr = element.getAttribute('style') || '';
    const isCenter = /text-align:\s*center/i.test(styleAttr) || element.getAttribute('align') === 'center';
    const isRight = /text-align:\s*right/i.test(styleAttr) || element.getAttribute('align') === 'right';

    const getX = (line: string, font: { widthOfTextAtSize: (t: string, s: number) => number }, size: number) => {
      if (isCenter) {
        try {
          const w = font.widthOfTextAtSize(line, size);
          return marginX + Math.max(0, (contentWidth - w) / 2);
        } catch {
          return marginX;
        }
      }
      if (isRight) {
        try {
          const w = font.widthOfTextAtSize(line, size);
          return marginX + Math.max(0, contentWidth - w);
        } catch {
          return marginX;
        }
      }
      return marginX;
    };

    const baseFont = isDocSans ? helvetica : times;
    const baseFontBold = isDocSans ? helveticaBold : timesBold;
    const baseFontItalic = isDocSans ? helveticaOblique : timesItalic;

    if (tagName === 'h1') {
      if (!text) return;
      const size = isDocSans ? 18 : 22;
      const font = helveticaBold;
      const isBlue = /#1e3a8a|#2563eb|#1a365d/i.test(styleAttr) || /#1e3a8a|#2563eb|#1a365d/i.test(element.innerHTML);
      const lines = wrapText(safe(font, text), font, size, contentWidth);
      const blockHeight = lines.length * (size * 1.3) + 14;
      ensureSpace(blockHeight);
      lines.forEach((line) => {
        cursorY -= size * 1.3;
        currentPage.drawText(line, {
          x: getX(line, font, size),
          y: cursorY,
          size,
          font,
          color: isBlue ? rgb(0.12, 0.23, 0.54) : rgb(0.12, 0.11, 0.09),
        });
      });
      cursorY -= 8;
    } else if (tagName === 'h2') {
      if (!text) return;
      const size = isDocSans ? 11 : 13.5;
      const font = helveticaBold;
      const isBlue = /#1e3a8a|#2563eb|#1a365d/i.test(styleAttr) || /#1e3a8a|#2563eb|#1a365d/i.test(element.innerHTML);
      const lines = wrapText(safe(font, text), font, size, contentWidth);
      const blockHeight = lines.length * (size * 1.3) + 12;
      ensureSpace(blockHeight);
      lines.forEach((line) => {
        cursorY -= size * 1.3;
        currentPage.drawText(line, {
          x: getX(line, font, size),
          y: cursorY,
          size,
          font,
          color: isBlue ? rgb(0.12, 0.23, 0.54) : rgb(0.18, 0.16, 0.13),
        });
      });
      cursorY -= 4;
    } else if (tagName === 'h3') {
      if (!text) return;
      const size = 12.5;
      const font = helveticaBold;
      const lines = wrapText(safe(font, text), font, size, contentWidth);
      const blockHeight = lines.length * (size * 1.3) + 8;
      ensureSpace(blockHeight);
      lines.forEach((line) => {
        cursorY -= size * 1.3;
        currentPage.drawText(line, { x: getX(line, font, size), y: cursorY, size, font, color: rgb(0.2, 0.18, 0.15) });
      });
      cursorY -= 6;
    } else if (tagName === 'table') {
      const isJobSplit =
        element.classList.contains('job-split') ||
        (!element.classList.contains('skills-table') && (element.textContent?.includes('–') || element.textContent?.includes('Present')));
      const isSkillsTable = element.classList.contains('skills-table');
      const rows = Array.from(element.querySelectorAll('tr'));

      rows.forEach((row) => {
        const cells = Array.from(row.querySelectorAll('td, th'));
        if (cells.length === 2 && isJobSplit && !isSkillsTable) {
          const leftText = (cells[0].textContent || '').trim();
          const rightText = (cells[1].textContent || '').trim();
          const size = isDocSans ? 10.5 : 11;
          const fontLeft = helveticaBold;
          const fontRight = helveticaOblique;
          ensureSpace(size * 1.5 + 4);
          cursorY -= size * 1.35;
          currentPage.drawText(safe(fontLeft, leftText), {
            x: marginX,
            y: cursorY,
            size,
            font: fontLeft,
            color: rgb(0.08, 0.1, 0.15),
          });
          const rightW = fontRight.widthOfTextAtSize(rightText, size);
          currentPage.drawText(safe(fontRight, rightText), {
            x: marginX + contentWidth - rightW,
            y: cursorY,
            size,
            font: fontRight,
            color: rgb(0.28, 0.33, 0.41),
          });
          cursorY -= 3;
        } else if (cells.length === 2) {
          // 2-column key-value table (e.g. Technical Skills)
          const leftText = (cells[0].textContent || '').trim();
          const rightText = (cells[1].textContent || '').trim();
          const size = isDocSans ? 9.5 : 10;
          const col1Width = 140;
          const col2Width = contentWidth - col1Width;
          const fontLeft = helveticaBold;
          const fontRight = helvetica;
          const rightLines = wrapText(safe(fontRight, rightText), fontRight, size, col2Width);
          const blockHeight = Math.max(1, rightLines.length) * (size * 1.3) + 2;
          ensureSpace(blockHeight);
          currentPage.drawText(safe(fontLeft, leftText), {
            x: marginX,
            y: cursorY - size * 1.25,
            size,
            font: fontLeft,
            color: rgb(0.12, 0.23, 0.54), // navy blue
          });
          let valY = cursorY;
          rightLines.forEach((rLine) => {
            valY -= size * 1.3;
            currentPage.drawText(rLine, {
              x: marginX + col1Width,
              y: valY,
              size,
              font: fontRight,
              color: rgb(0.12, 0.15, 0.2),
            });
          });
          cursorY = valY - 2;
        } else {
          // Generic N-column table
          const numCols = cells.length;
          if (numCols === 0) return;
          const colWidth = contentWidth / numCols;
          const size = isDocSans ? 9.5 : 10;
          const cellLines = cells.map((cell) => {
            const cText = (cell.textContent || '').trim();
            const isTh = cell.tagName.toLowerCase() === 'th';
            const font = isTh ? helveticaBold : helvetica;
            return {
              lines: wrapText(safe(font, cText), font, size, colWidth - 8),
              font,
            };
          });
          const maxLines = Math.max(1, ...cellLines.map((c) => c.lines.length));
          const rowHeight = maxLines * (size * 1.3) + 6;
          ensureSpace(rowHeight);
          currentPage.drawRectangle({
            x: marginX,
            y: cursorY - rowHeight,
            width: contentWidth,
            height: rowHeight,
            borderColor: rgb(0.85, 0.88, 0.92),
            borderWidth: 0.5,
          });
          cellLines.forEach((cData, cIdx) => {
            const cellX = marginX + cIdx * colWidth;
            let cY = cursorY - 4;
            cData.lines.forEach((l) => {
              cY -= size * 1.3;
              currentPage.drawText(l, {
                x: cellX + 4,
                y: cY,
                size,
                font: cData.font,
                color: rgb(0.12, 0.15, 0.2),
              });
            });
          });
          cursorY -= rowHeight;
        }
      });
      cursorY -= 4;
    } else if (tagName === 'p') {
      if (!text) {
        cursorY -= 8;
        return;
      }
      const isSub = /font-size:\s*11px.*Page/i.test(styleAttr);
      const isBold = Boolean(element.querySelector('strong, b')) && !element.querySelector('span, em, i');
      const isItalic = Boolean(element.querySelector('em, i')) && !element.querySelector('strong, b');
      const font = isSub ? helvetica : (isBold ? baseFontBold : isItalic ? baseFontItalic : baseFont);
      const size = isSub ? 9.5 : (isDocSans ? 10 : 11);
      const lines = wrapText(safe(font, text), font, size, contentWidth);
      const blockHeight = lines.length * (size * 1.35) + 6;
      ensureSpace(blockHeight);
      lines.forEach((line) => {
        cursorY -= size * 1.35;
        currentPage.drawText(line, {
          x: getX(line, font, size),
          y: cursorY,
          size,
          font,
          color: isSub ? rgb(0.55, 0.52, 0.46) : rgb(0.15, 0.14, 0.12),
        });
      });
      cursorY -= 4;
    } else if (tagName === 'ul' || tagName === 'ol') {
      const items = Array.from(element.querySelectorAll(':scope > li'));
      items.forEach((item, index) => {
        const itemText = (item.textContent || '').trim();
        if (!itemText) return;
        const size = isDocSans ? 10 : 11;
        const font = baseFont;
        const bullet = tagName === 'ol' ? `${index + 1}. ` : '• ';
        const bulletWidth = 14;
        const lines = wrapText(safe(font, itemText), font, size, contentWidth - bulletWidth);
        const blockHeight = lines.length * (size * 1.35) + 4;
        ensureSpace(blockHeight);
        currentPage.drawText(bullet, { x: marginX + 2, y: cursorY - size * 1.35, size, font: helveticaBold, color: rgb(0.35, 0.35, 0.35) });
        lines.forEach((line) => {
          cursorY -= size * 1.35;
          currentPage.drawText(line, { x: marginX + bulletWidth, y: cursorY, size, font, color: rgb(0.15, 0.14, 0.12) });
        });
        cursorY -= 2;
      });
      cursorY -= 4;
    } else if (tagName === 'blockquote') {
      if (!text) return;
      const size = 10.5;
      const font = helveticaOblique;
      const indent = 16;
      const lines = wrapText(safe(font, text), font, size, contentWidth - indent);
      const blockHeight = lines.length * (size * 1.35) + 8;
      ensureSpace(blockHeight);
      currentPage.drawLine({
        start: { x: marginX + 4, y: cursorY },
        end: { x: marginX + 4, y: cursorY - blockHeight + 6 },
        thickness: 2,
        color: rgb(0.77, 0.32, 0.18),
      });
      lines.forEach((line) => {
        cursorY -= size * 1.35;
        currentPage.drawText(line, { x: marginX + indent, y: cursorY, size, font, color: rgb(0.35, 0.33, 0.29) });
      });
      cursorY -= 8;
    } else if (tagName === 'hr') {
      const prevTag = element.previousElementSibling?.tagName.toLowerCase();
      const isSectionRule =
        element.classList.contains('section-rule') ||
        prevTag === 'h2' ||
        styleAttr.includes('2563eb') ||
        styleAttr.includes('1.5px');
      ensureSpace(12);
      cursorY -= 4;
      currentPage.drawLine({
        start: { x: marginX, y: cursorY },
        end: { x: marginX + contentWidth, y: cursorY },
        thickness: isSectionRule ? 1.5 : 0.75,
        color: isSectionRule ? rgb(0.15, 0.39, 0.92) : rgb(0.85, 0.82, 0.76),
      });
      cursorY -= 6;
    } else if (tagName === 'pre' || tagName === 'code') {
      if (!text) return;
      const size = 9.5;
      const font = courier;
      const lines = text.split('\n');
      const blockHeight = lines.length * (size * 1.3) + 12;
      ensureSpace(blockHeight);
      currentPage.drawRectangle({
        x: marginX,
        y: cursorY - blockHeight + 6,
        width: contentWidth,
        height: blockHeight - 6,
        color: rgb(0.96, 0.94, 0.9),
      });
      lines.forEach((line) => {
        cursorY -= size * 1.3;
        currentPage.drawText(safe(font, line.slice(0, 85)), { x: marginX + 8, y: cursorY, size, font, color: rgb(0.2, 0.2, 0.2) });
      });
      cursorY -= 8;
    } else if (element.children.length > 0) {
      Array.from(element.children).forEach(renderNode);
    } else if (text) {
      const size = 11;
      const font = times;
      const lines = wrapText(safe(font, text), font, size, contentWidth);
      const blockHeight = lines.length * (size * 1.35) + 4;
      ensureSpace(blockHeight);
      lines.forEach((line) => {
        cursorY -= size * 1.35;
        currentPage.drawText(line, { x: marginX, y: cursorY, size, font, color: rgb(0.15, 0.14, 0.12) });
      });
      cursorY -= 4;
    }
  }

  Array.from(doc.body.children).forEach(renderNode);

  // If no children were rendered
  if (doc.body.children.length === 0 && doc.body.textContent) {
    const text = doc.body.textContent.trim();
    const size = 11;
    const font = times;
    const lines = wrapText(safe(font, text), font, size, contentWidth);
    lines.forEach((line) => {
      ensureSpace(size * 1.35);
      cursorY -= size * 1.35;
      currentPage.drawText(line, { x: marginX, y: cursorY, size, font, color: rgb(0.15, 0.14, 0.12) });
    });
  }

  const allPages = pdf.getPages();
  allPages.forEach((p, idx) => {
    const pageNumText = `Page ${idx + 1} of ${allPages.length}`;
    try {
      const textWidth = helvetica.widthOfTextAtSize(pageNumText, 9);
      p.drawText(pageNumText, {
        x: (pageWidth - textWidth) / 2,
        y: 28,
        size: 9,
        font: helvetica,
        color: rgb(0.55, 0.52, 0.46),
      });
    } catch {
      // ignore
    }
  });

  return pdf.save();
}

function wrapText(
  text: string,
  font: { widthOfTextAtSize: (text: string, size: number) => number },
  size: number,
  maxWidth: number,
): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split('\n')) {
    const words = paragraph.split(/(\s+)/);
    let current = '';
    for (const word of words) {
      const next = current + word;
      if (current && font.widthOfTextAtSize(next, size) > maxWidth) {
        lines.push(current.trimEnd());
        current = word.trimStart();
      } else {
        current = next;
      }
    }
    lines.push(current.trimEnd());
  }
  return lines.length > 0 ? lines : [''];
}

export function hexToRgb(hex: string) {
  const normalized = hex.replace('#', '');
  const value = normalized.length === 3 ? normalized.split('').map((char) => char + char).join('') : normalized;
  const int = Number.parseInt(value, 16);
  if (!Number.isFinite(int)) return rgb(0, 0, 0);
  return rgb(((int >> 16) & 255) / 255, ((int >> 8) & 255) / 255, (int & 255) / 255);
}

export async function buildSamplePdf(): Promise<ArrayBuffer> {
  const pdf = await PDFDocument.create();
  const serif = await pdf.embedFont(StandardFonts.TimesRoman);
  const serifItalic = await pdf.embedFont(StandardFonts.TimesRomanItalic);
  const sans = await pdf.embedFont(StandardFonts.Helvetica);
  const sansBold = await pdf.embedFont(StandardFonts.HelveticaBold);

  const ink = rgb(0.12, 0.11, 0.09);
  const soft = rgb(0.35, 0.33, 0.29);
  const accent = rgb(0.77, 0.32, 0.18);
  const cream = rgb(0.965, 0.945, 0.9);

  const pageOne = pdf.addPage([595, 842]);
  pageOne.drawRectangle({ x: 0, y: 0, width: 595, height: 842, color: cream });
  pageOne.drawText('PAPYRUS STUDIO', {
    x: 60,
    y: 770,
    size: 9,
    font: sansBold,
    color: accent,
  });
  pageOne.drawText('Field notes on', { x: 58, y: 690, size: 38, font: serif, color: ink });
  pageOne.drawText('reading & marking', { x: 58, y: 646, size: 38, font: serifItalic, color: ink });
  pageOne.drawText('a short guide to this little workspace', {
    x: 60,
    y: 612,
    size: 12,
    font: sans,
    color: soft,
  });
  pageOne.drawLine({ start: { x: 60, y: 588 }, end: { x: 535, y: 588 }, thickness: 0.8, color: ink });

  const body = [
    'This sample PDF is yours to explore. Flip through the',
    'pages, search for a phrase, or pick up a pen from the',
    'toolbar and leave a mark in the margin.',
    '',
    'Papyrus keeps your documents in this browser. Nothing',
    'is uploaded, and your annotations stay with the file',
    'until you choose to export a flattened copy.',
  ];
  body.forEach((line, index) => {
    pageOne.drawText(line, {
      x: 60,
      y: 548 - index * 22,
      size: 12,
      font: serif,
      color: ink,
    });
  });

  pageOne.drawRectangle({
    x: 60,
    y: 250,
    width: 475,
    height: 82,
    color: rgb(1, 1, 1),
    opacity: 0.55,
    borderColor: rgb(0.85, 0.81, 0.72),
    borderWidth: 1,
  });
  pageOne.drawText('Try this:', { x: 78, y: 306, size: 11, font: sansBold, color: accent });
  pageOne.drawText('Choose the highlighter, then drag across this sentence.', {
    x: 78,
    y: 286,
    size: 11,
    font: sans,
    color: soft,
  });
  pageOne.drawText('1', { x: 512, y: 50, size: 9, font: sans, color: soft });

  const pageTwo = pdf.addPage([595, 842]);
  pageTwo.drawRectangle({ x: 0, y: 0, width: 595, height: 842, color: cream });
  pageTwo.drawText('A few quiet gestures', { x: 58, y: 740, size: 30, font: serif, color: ink });
  pageTwo.drawText('02', { x: 60, y: 706, size: 10, font: sansBold, color: accent });

  const gestures = [
    ['Read', 'Scroll or use the page controls. Zoom with the toolbar;'],
    ['', 'the layout reflows without losing your place.'],
    ['Search', 'Press the search field and type a word. Matches are'],
    ['', 'highlighted on every page, and you can step between them.'],
    ['Mark', 'Highlight passages, draw freehand, or drop a note in'],
    ['', 'the margin. Every mark can be undone.'],
    ['Take it with you', 'Download a flattened PDF with your annotations, or'],
    ['', 'send the extracted text into the Word-style editor.'],
  ];

  let y = 640;
  for (const [label, description] of gestures) {
    if (label) {
      pageTwo.drawText(label, { x: 60, y, size: 11, font: sansBold, color: accent });
    }
    pageTwo.drawText(description, { x: 160, y, size: 11, font: serif, color: ink });
    y -= 34;
  }
  pageTwo.drawLine({ start: { x: 60, y: 420 }, end: { x: 535, y: 420 }, thickness: 0.7, color: soft });
  pageTwo.drawText('“The page is a stage. Mark it up.”', {
    x: 60,
    y: 380,
    size: 16,
    font: serifItalic,
    color: soft,
  });
  pageTwo.drawText('2', { x: 512, y: 50, size: 9, font: sans, color: soft });

  const pageThree = pdf.addPage([595, 842]);
  pageThree.drawRectangle({ x: 0, y: 0, width: 595, height: 842, color: cream });
  pageThree.drawText('The last page', { x: 58, y: 740, size: 30, font: serif, color: ink });
  pageThree.drawText(
    'Everything in this document lives on your device.',
    { x: 60, y: 692, size: 12, font: serif, color: ink },
  );
  pageThree.drawText(
    'Close the tab and come back later — recent documents',
    { x: 60, y: 670, size: 12, font: serif, color: ink },
  );
  pageThree.drawText(
    'will be waiting in the library.',
    { x: 60, y: 648, size: 12, font: serif, color: ink },
  );

  pageThree.drawRectangle({
    x: 60,
    y: 500,
    width: 475,
    height: 94,
    color: rgb(0.12, 0.11, 0.09),
  });
  pageThree.drawText('papyrus  /puh-PY-rus/', {
    x: 80,
    y: 556,
    size: 13,
    font: serifItalic,
    color: rgb(0.98, 0.96, 0.92),
  });
  pageThree.drawText('noun — paper made from the pith of the papyrus plant.', {
    x: 80,
    y: 532,
    size: 9.5,
    font: sans,
    color: rgb(0.8, 0.78, 0.72),
  });
  pageThree.drawText('3', { x: 512, y: 50, size: 9, font: sans, color: soft });

  const bytes = await pdf.save();
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}
