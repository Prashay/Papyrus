import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType,
} from 'docx';
import mammoth from 'mammoth';

interface PMMark {
  type: string;
  attrs?: Record<string, unknown> | null;
}

interface PMNode {
  type?: string;
  attrs?: Record<string, unknown> | null;
  content?: PMNode[];
  marks?: PMMark[];
  text?: string;
}

const FONT_MAP: Record<string, string> = {
  'Newsreader Variable': 'Georgia',
  'Fraunces Variable': 'Georgia',
  'Inter Variable': 'Arial',
  'Helvetica Neue': 'Arial',
  'Iowan Old Style': 'Georgia',
};

const ALIGN_MAP: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

const NUMBERING_REFERENCE = 'papyrus-ordered';

export async function importDocx(bytes: ArrayBuffer): Promise<{ html: string; warnings: string[] }> {
  const result = await mammoth.convertToHtml(
    { arrayBuffer: bytes },
    {
      styleMap: [
        "p[style-name='Title'] => h1.papyrus-title:fresh",
        "p[style-name='Subtitle'] => p.papyrus-subtitle:fresh",
      ],
      includeDefaultStyleMap: true,
    },
  );

  const warnings = result.messages
    .filter((message) => message.type === 'warning')
    .map((message) => message.message)
    .filter((message) => !/image/i.test(message))
    .slice(0, 3);

  const html = result.value.trim() || '<p></p>';
  return { html, warnings };
}

function headingLevel(level: number): (typeof HeadingLevel)[keyof typeof HeadingLevel] {
  if (level <= 1) return HeadingLevel.HEADING_1;
  if (level === 2) return HeadingLevel.HEADING_2;
  if (level === 3) return HeadingLevel.HEADING_3;
  if (level === 4) return HeadingLevel.HEADING_4;
  if (level === 5) return HeadingLevel.HEADING_5;
  return HeadingLevel.HEADING_6;
}

function normalizeLineSpacing(value: unknown): number {
  const parsed = typeof value === 'string' ? Number.parseFloat(value) : Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) return 340;
  return Math.round(240 * Math.min(Math.max(parsed, 1), 2.4));
}

function parseFontSize(value: unknown): number | undefined {
  if (value === undefined || value === null || value === '') return undefined;
  const raw = String(value).trim();
  const amount = Number.parseFloat(raw);
  if (!Number.isFinite(amount) || amount <= 0) return undefined;
  if (raw.endsWith('pt')) return Math.round(amount * 2);
  if (raw.endsWith('em') || raw.endsWith('rem')) return Math.round(amount * 16 * 1.5);
  return Math.round(amount * 1.5);
}

function normalizeAlign(value: unknown): (typeof AlignmentType)[keyof typeof AlignmentType] | undefined {
  if (typeof value !== 'string') return undefined;
  return ALIGN_MAP[value];
}

function normalizeHex(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const hex = value.replace('#', '').trim();
  return /^[0-9a-f]{6}$/i.test(hex) ? hex.toUpperCase() : undefined;
}

function fontFromMark(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const first = value.split(',')[0]?.replace(/['"]/g, '').trim();
  if (!first) return undefined;
  return FONT_MAP[first] ?? first;
}

async function dataUrlToPng(
  dataUrl: string,
): Promise<{ data: ArrayBuffer; width: number; height: number } | null> {
  try {
    const response = await fetch(dataUrl);
    const blob = await response.blob();
    const bitmap = await createImageBitmap(blob);
    const maxWidth = 480;
    const scale = Math.min(1, maxWidth / bitmap.width);
    const width = Math.max(1, Math.round(bitmap.width * scale));
    const height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d');
    if (!context) return null;
    context.drawImage(bitmap, 0, 0, width, height);
    bitmap.close?.();
    const pngBlob = await new Promise<Blob | null>((resolve) =>
      canvas.toBlob((result) => resolve(result), 'image/png'),
    );
    if (!pngBlob) return null;
    return { data: await pngBlob.arrayBuffer(), width, height };
  } catch {
    return null;
  }
}

function textRunsForNode(node: PMNode): TextRun[] {
  if (!node.text) return [];
  const marks = node.marks ?? [];
  const mark = (type: string) => marks.find((item) => item.type === type);

  const bold = Boolean(mark('bold'));
  const italics = Boolean(mark('italic'));
  const underline = Boolean(mark('underline'));
  const strike = Boolean(mark('strike'));
  const code = Boolean(mark('code'));
  const link = mark('link');
  const textStyle = mark('textStyle');
  const highlight = mark('highlight');

  const styleAttrs = (textStyle?.attrs ?? {}) as Record<string, unknown>;
  const fontSize = parseFontSize(styleAttrs.fontSize);
  const fontFamily = fontFromMark(styleAttrs.fontFamily);
  const color = normalizeHex(styleAttrs.color) ?? (link ? '1F5C99' : undefined);
  const highlightColor = normalizeHex(highlight?.attrs?.color);

  const run = new TextRun({
    text: node.text,
    bold,
    italics,
    strike,
    underline: underline ? { type: UnderlineType.SINGLE } : undefined,
    font: code ? 'Courier New' : fontFamily,
    size: code ? 20 : fontSize,
    color: code ? '7A3E2A' : color,
    shading: highlightColor
      ? { fill: highlightColor, type: ShadingType.CLEAR, color: 'auto' }
      : code
        ? { fill: 'F1EDE4', type: ShadingType.CLEAR, color: 'auto' }
        : undefined,
  });

  if (link && typeof link.attrs?.href === 'string') {
    return [];
  }

  return [run];
}

function inlineChildren(children: PMNode[] | undefined): ParagraphChildWithLink[] {
  const result: ParagraphChildWithLink[] = [];
  for (const child of children ?? []) {
    if (child.type === 'text') {
      const link = child.marks?.find((mark) => mark.type === 'link');
      const href = link && typeof link.attrs?.href === 'string' ? link.attrs.href : null;
      if (href) {
        const withoutLink = { ...child, marks: child.marks?.filter((mark) => mark.type !== 'link') };
        result.push({ link: href, runs: textRunsForNode(withoutLink as PMNode) });
      } else {
        result.push(...textRunsForNode(child).map((run) => ({ run })));
      }
    } else if (child.type === 'hardBreak') {
      result.push({ run: new TextRun({ break: 1 }) });
    } else if (child.type === 'image') {
      const src = typeof child.attrs?.src === 'string' ? child.attrs.src : '';
      if (src.startsWith('data:')) {
        result.push({ image: src });
      }
    } else if (child.content) {
      result.push(...inlineChildren(child.content));
    }
  }
  return result;
}

type ParagraphChildWithLink =
  | { run: TextRun; link?: undefined; image?: undefined }
  | { link: string; runs: TextRun[]; run?: undefined; image?: undefined }
  | { image: string; run?: undefined; link?: undefined };

async function buildChildren(children: PMNode[] | undefined) {
  const items = inlineChildren(children);
  const output: (TextRun | ExternalHyperlink | ImageRun)[] = [];

  for (const item of items) {
    if ('run' in item && item.run) {
      output.push(item.run);
    } else if ('link' in item && item.link) {
      const runs = item.runs.length > 0 ? item.runs : [new TextRun({ text: item.link, color: '1F5C99' })];
      output.push(new ExternalHyperlink({ link: item.link, children: runs }));
    } else if ('image' in item && item.image) {
      const prepared = await dataUrlToPng(item.image);
      if (prepared) {
        output.push(
          new ImageRun({
            type: 'png',
            data: prepared.data,
            transformation: { width: prepared.width, height: prepared.height },
          }),
        );
      }
    }
  }

  return output;
}

interface BlockContext {
  listDepth: number;
  inBlockquote: boolean;
}

async function paragraphForNode(node: PMNode, context: BlockContext): Promise<Paragraph[]> {
  const children = await buildChildren(node.content);
  const attrs = node.attrs ?? {};
  const align = normalizeAlign(attrs.textAlign);
  const lineSpacing = normalizeLineSpacing(attrs.lineHeight);
  const heading = node.type === 'heading' ? headingLevel(Number(attrs.level ?? 1)) : undefined;

  if (children.length === 0) {
    children.push(new TextRun({ text: '' }));
  }

  return [
    new Paragraph({
      children,
      heading,
      alignment: align,
      spacing: { line: lineSpacing, after: heading ? 160 : 120 },
      indent: context.inBlockquote ? { left: 720 } : undefined,
      style: context.inBlockquote ? 'PapyrusQuote' : undefined,
    }),
  ];
}

async function listParagraphs(
  node: PMNode,
  depth: number,
  ordered: boolean,
  context: BlockContext,
): Promise<ParallelParagraph[]> {
  const output: ParallelParagraph[] = [];

  for (const item of node.content ?? []) {
    if (item.type !== 'listItem') continue;
    let firstParagraphDone = false;

    for (const child of item.content ?? []) {
      if (child.type === 'paragraph') {
        const children = await buildChildren(child.content);
        if (children.length === 0) children.push(new TextRun({ text: '' }));
        output.push(
          new Paragraph({
            children,
            bullet: ordered ? undefined : { level: depth },
            numbering: ordered ? { reference: NUMBERING_REFERENCE, level: depth } : undefined,
            spacing: { line: 320, after: 80 },
            indent: context.inBlockquote ? { left: 720 } : undefined,
          }),
        );
        firstParagraphDone = true;
      } else if (child.type === 'bulletList' || child.type === 'orderedList') {
        output.push(
          ...(await listParagraphs(child, Math.min(depth + 1, 2), child.type === 'orderedList', context)),
        );
      } else if (child.type === 'table') {
        output.push(await tableForNode(child));
      }
    }

    if (!firstParagraphDone) {
      output.push(
        new Paragraph({
          children: [new TextRun({ text: '' })],
          bullet: ordered ? undefined : { level: depth },
          numbering: ordered ? { reference: NUMBERING_REFERENCE, level: depth } : undefined,
        }),
      );
    }
  }

  return output;
}

type ParallelParagraph = Paragraph | Table;

async function tableForNode(node: PMNode): Promise<Table> {
  const rows: TableRow[] = [];

  for (const rowNode of node.content ?? []) {
    if (rowNode.type !== 'tableRow') continue;
    const cells: TableCell[] = [];

    for (const cellNode of rowNode.content ?? []) {
      if (cellNode.type !== 'tableCell' && cellNode.type !== 'tableHeader') continue;
      const paragraphs: Paragraph[] = [];
      for (const child of cellNode.content ?? []) {
        if (child.type === 'paragraph') {
          paragraphs.push(...(await paragraphForNode(child, { listDepth: 0, inBlockquote: false })));
        }
      }
      if (paragraphs.length === 0) paragraphs.push(new Paragraph({ text: '' }));
      cells.push(
        new TableCell({
          children: paragraphs,
          shading:
            cellNode.type === 'tableHeader'
              ? { fill: 'F1EDE4', type: ShadingType.CLEAR, color: 'auto' }
              : undefined,
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
        }),
      );
    }

    if (cells.length > 0) rows.push(new TableRow({ children: cells }));
  }

  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
}

async function blocksForNode(node: PMNode, context: BlockContext): Promise<ParallelParagraph[]> {
  switch (node.type) {
    case 'paragraph':
      return paragraphForNode(node, context);
    case 'heading':
      return paragraphForNode(node, context);
    case 'blockquote': {
      const nested = (node.content ?? []).flatMap((child) =>
        child.type === 'paragraph' || child.type === 'heading'
          ? [{ ...child, attrs: { ...child.attrs, textAlign: child.attrs?.textAlign } }]
          : [],
      );
      const paragraphs: Paragraph[] = [];
      for (const child of nested) {
        paragraphs.push(...(await paragraphForNode(child, { ...context, inBlockquote: true })));
      }
      return paragraphs;
    }
    case 'bulletList':
      return listParagraphs(node, 0, false, context);
    case 'orderedList':
      return listParagraphs(node, 0, true, context);
    case 'codeBlock': {
      const text = node.content?.map((child) => child.text ?? '').join('') ?? '';
      const lines = text.split('\n');
      const runs: TextRun[] = [];
      lines.forEach((line, index) => {
        if (index > 0) runs.push(new TextRun({ break: 1 }));
        runs.push(new TextRun({ text: line || ' ', font: 'Courier New', size: 20, color: '2A2721' }));
      });
      return [
        new Paragraph({
          children: runs,
          shading: { fill: 'F1EDE4', type: ShadingType.CLEAR, color: 'auto' },
          spacing: { line: 280, after: 160, before: 120 },
          border: {
            top: { style: BorderStyle.SINGLE, size: 4, color: 'DED6C7', space: 8 },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: 'DED6C7', space: 8 },
            left: { style: BorderStyle.SINGLE, size: 4, color: 'DED6C7', space: 8 },
            right: { style: BorderStyle.SINGLE, size: 4, color: 'DED6C7', space: 8 },
          },
        }),
      ];
    }
    case 'horizontalRule':
      return [
        new Paragraph({
          text: '',
          spacing: { before: 120, after: 120 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: 'C9C0AE', space: 1 },
          },
        }),
      ];
    case 'image': {
      const src = typeof node.attrs?.src === 'string' ? node.attrs.src : '';
      if (!src.startsWith('data:')) return [];
      const prepared = await dataUrlToPng(src);
      if (!prepared) return [];
      return [
        new Paragraph({
          children: [
            new ImageRun({
              type: 'png',
              data: prepared.data,
              transformation: { width: prepared.width, height: prepared.height },
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 160 },
        }),
      ];
    }
    case 'table':
      return [await tableForNode(node)];
    default:
      if (node.content) {
        const output: ParallelParagraph[] = [];
        for (const child of node.content) {
          output.push(...(await blocksForNode(child, context)));
        }
        return output;
      }
      return [];
  }
}

export async function exportDocx(content: PMNode, title: string): Promise<Blob> {
  const blocks: ParallelParagraph[] = [];
  for (const node of content.content ?? []) {
    blocks.push(...(await blocksForNode(node, { listDepth: 0, inBlockquote: false })));
  }

  if (blocks.length === 0) {
    blocks.push(new Paragraph({ text: '' }));
  }

  const doc = new Document({
    creator: 'Papyrus',
    title,
    description: 'Edited with Papyrus',
    styles: {
      default: {
        document: {
          run: { font: 'Georgia', size: 22, color: '23211D' },
          paragraph: { spacing: { line: 340, after: 120 } },
        },
        heading1: {
          run: { font: 'Georgia', size: 36, bold: true, color: '1B1915' },
          paragraph: { spacing: { before: 320, after: 140 } },
        },
        heading2: {
          run: { font: 'Georgia', size: 30, bold: true, color: '1B1915' },
          paragraph: { spacing: { before: 280, after: 120 } },
        },
        heading3: {
          run: { font: 'Georgia', size: 26, bold: true, color: '1B1915' },
          paragraph: { spacing: { before: 240, after: 100 } },
        },
        heading4: {
          run: { font: 'Georgia', size: 23, bold: true, color: '1B1915' },
          paragraph: { spacing: { before: 200, after: 80 } },
        },
      },
    },
    numbering: {
      config: [
        {
          reference: NUMBERING_REFERENCE,
          levels: [
            {
              level: 0,
              format: LevelFormat.DECIMAL,
              text: '%1.',
              alignment: AlignmentType.START,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } },
            },
            {
              level: 1,
              format: LevelFormat.LOWER_LETTER,
              text: '%2.',
              alignment: AlignmentType.START,
              style: { paragraph: { indent: { left: 1440, hanging: 360 } } },
            },
            {
              level: 2,
              format: LevelFormat.LOWER_ROMAN,
              text: '%3.',
              alignment: AlignmentType.START,
              style: { paragraph: { indent: { left: 2160, hanging: 360 } } },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        children: blocks,
      },
    ],
  });

  return Packer.toBlob(doc);
}

export function plainTextToHtml(text: string): string {
  const normalized = text.replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ');
  const paragraphs = normalized.split(/\n{2,}/);
  return paragraphs
    .map((paragraph) => {
      const lines = paragraph.split('\n');
      const inner = lines
        .map((line) => line.replace(/[&<>]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[char] ?? char))
        .join('<br />');
      return `<p>${inner || '<br />'}</p>`;
    })
    .join('');
}

export function textToDocxContent(text: string, title?: string): PMNode {
  const lines = text.replace(/\r\n?/g, '\n').split('\n');
  const content: PMNode[] = [];
  if (title) {
    content.push({ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: title }] });
  }
  for (const line of lines) {
    content.push({
      type: 'paragraph',
      content: line ? [{ type: 'text', text: line }] : [],
    });
  }
  return { type: 'doc', content };
}

export function pagesToDocxContent(title: string, pages: string[]): PMNode {
  const content: PMNode[] = [
    { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: title }] },
  ];

  pages.forEach((pageText, index) => {
    content.push({
      type: 'heading',
      attrs: { level: 2 },
      content: [{ type: 'text', text: `Page ${index + 1}` }],
    });
    const lines = pageText.replace(/\r\n?/g, '\n').split('\n');
    for (const line of lines) {
      content.push({
        type: 'paragraph',
        content: line.trim() ? [{ type: 'text', text: line.trim() }] : [],
      });
    }
  });

  return { type: 'doc', content };
}
