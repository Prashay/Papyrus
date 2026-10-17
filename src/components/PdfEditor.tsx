import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type PointerEvent as ReactPointerEvent,
  type ReactNode,
} from 'react';
import {
  AlertTriangle,
  Bold,
  Check,
  ChevronLeft,
  ChevronRight,
  Download,
  Edit3,
  Eraser,
  FileEdit,
  FileOutput,
  FolderOpen,
  Highlighter,
  Italic,
  Loader2,
  Maximize2,
  MousePointer2,
  PenTool,
  Plus,
  Redo2,
  Search,
  Square,
  Trash2,
  Type,
  Undo2,
  X,
} from 'lucide-react';
import { saveAs } from 'file-saver';
import type { OpenDocument, PdfAnnotation, PdfTool, Point, StoredState } from '../types';
import { ACCEPTED_EXTENSIONS } from '../lib/detect';
import { baseName, clamp, uid, withExtension } from '../lib/format';
import {
  extractDocumentText,
  exportAnnotatedPdf,
  extractPageText,
  extractPdfAsHtml,
  loadPdf,
  renderPage,
  type PageText,
  type PageViewport,
  type PDFDocumentProxy,
  type TextItemBox,
} from '../lib/pdf';
import { exportDocx, pagesToDocxContent } from '../lib/docx';
import type { ToastTone } from './ui';

interface SearchMatch {
  page: number;
  start: number;
  rects: { x: number; y: number; width: number; height: number }[];
}

type SaveStatus = 'saved' | 'saving' | 'error';

interface PdfEditorProps {
  doc: OpenDocument;
  notify: (message: string, tone?: ToastTone) => void;
  onBack: () => void;
  onRename: (name: string) => void;
  onPersist: (patch: StoredState) => Promise<void>;
  onOpenFile: (file: File) => void;
  onEditAsDocument?: (html: string) => void;
}

interface TextDraft {
  page: number;
  x: number;
  y: number;
  left: number;
  top: number;
  originalItem?: TextItemBox;
  editingAnnotationId?: string;
  fontSize: number;
  serif: boolean;
  bold: boolean;
  italic: boolean;
  textColor: string;
  bgColor: string;
}

function sampleColorsAtItem(
  canvas: HTMLCanvasElement | null,
  item: TextItemBox,
  zoom: number,
  viewport: PageViewport,
): { bgColor: string; textColor: string } {
  const fallback = { bgColor: '#ffffff', textColor: '#23211d' };
  if (!canvas) return fallback;
  try {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    if (!ctx) return fallback;

    const scaleX = canvas.width / viewport.width;
    const scaleY = canvas.height / viewport.height;
    const left = Math.round(item.x * zoom * scaleX);
    const top = Math.round(item.y * zoom * scaleY);
    const width = Math.round(Math.max(4, item.width * zoom * scaleX));
    const height = Math.round(Math.max(4, item.height * zoom * scaleY));

    const bgSamples: [number, number][] = [
      [left + Math.floor(width / 2), Math.max(0, top - 3)],
      [Math.max(0, left - 3), top + Math.floor(height / 2)],
      [Math.min(canvas.width - 1, left + width + 3), top + Math.floor(height / 2)],
      [left + 2, Math.max(0, top - 2)],
    ];

    let bgR = 255;
    let bgG = 255;
    let bgB = 255;
    let foundBg = false;

    for (const [sx, sy] of bgSamples) {
      if (sx >= 0 && sx < canvas.width && sy >= 0 && sy < canvas.height) {
        const pixel = ctx.getImageData(sx, sy, 1, 1).data;
        if (pixel[3] > 180) {
          bgR = pixel[0];
          bgG = pixel[1];
          bgB = pixel[2];
          foundBg = true;
          break;
        }
      }
    }

    const sampleW = Math.min(width, 120);
    const sampleH = Math.min(height, 60);
    const boxData = ctx.getImageData(
      Math.max(0, Math.min(canvas.width - sampleW, left)),
      Math.max(0, Math.min(canvas.height - sampleH, top)),
      sampleW,
      sampleH,
    ).data;

    let maxDist = 0;
    let textR = 35;
    let textG = 33;
    let textB = 29;

    for (let i = 0; i < boxData.length; i += 4) {
      const alpha = boxData[i + 3];
      if (alpha < 180) continue;
      const r = boxData[i];
      const g = boxData[i + 1];
      const b = boxData[i + 2];
      const dist = (r - bgR) ** 2 + (g - bgG) ** 2 + (b - bgB) ** 2;
      if (dist > maxDist) {
        maxDist = dist;
        textR = r;
        textG = g;
        textB = b;
      }
    }

    const toHex = (c: number) => c.toString(16).padStart(2, '0');
    const sampledBg = foundBg ? `#${toHex(bgR)}${toHex(bgG)}${toHex(bgB)}` : '#ffffff';
    const sampledText = maxDist > 1600 ? `#${toHex(textR)}${toHex(textG)}${toHex(textB)}` : '#23211d';

    return { bgColor: sampledBg, textColor: sampledText };
  } catch {
    return fallback;
  }
}

function sampleBgColorAtPoint(
  canvas: HTMLCanvasElement | null,
  viewportX: number,
  viewportY: number,
  viewport: PageViewport,
): string {
  if (!canvas) return '#ffffff';
  try {
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    if (!ctx) return '#ffffff';
    const scaleX = canvas.width / viewport.width;
    const scaleY = canvas.height / viewport.height;
    const cx = Math.round(viewportX * scaleX);
    const cy = Math.round(viewportY * scaleY);
    if (cx >= 0 && cx < canvas.width && cy >= 0 && cy < canvas.height) {
      const pixel = ctx.getImageData(cx, cy, 1, 1).data;
      if (pixel[3] > 180) {
        const toHex = (c: number) => c.toString(16).padStart(2, '0');
        return `#${toHex(pixel[0])}${toHex(pixel[1])}${toHex(pixel[2])}`;
      }
    }
  } catch {
    // fallback
  }
  return '#ffffff';
}

const HIGHLIGHT_COLORS = ['#f2cf56', '#f0a35a', '#e2664a', '#4c7a5b', '#3c6e9f'];
const TEXT_COLORS = ['#23211d', '#c4512d', '#3c6e9f', '#4c7a5b', '#8a4f9e'];
const INK_COLORS = ['#23211d', '#c4512d', '#3c6e9f', '#4c7a5b', '#8a4f9e'];
const WHITEOUT_COLORS = ['#ffffff', '#fbf9f4', '#f4efe5', '#23211d'];

function ToolButton({
  title,
  active,
  disabled,
  onClick,
  children,
  className,
}: {
  title: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
  className?: string;
}) {
  return (
    <button
      type="button"
      className={`tool-btn ${className ?? ''} ${active ? 'active' : ''}`}
      title={title}
      aria-label={title}
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function distanceToSegment(point: Point, start: Point, end: Point): number {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  if (dx === 0 && dy === 0) return Math.hypot(point.x - start.x, point.y - start.y);
  const t = clamp(((point.x - start.x) * dx + (point.y - start.y) * dy) / (dx * dx + dy * dy), 0, 1);
  return Math.hypot(point.x - (start.x + t * dx), point.y - (start.y + t * dy));
}

function hitTest(annotation: PdfAnnotation, point: Point, zoom: number): boolean {
  if (annotation.type === 'highlight' || annotation.type === 'whiteout') {
    return (
      point.x >= annotation.x &&
      point.x <= annotation.x + annotation.width &&
      point.y >= annotation.y &&
      point.y <= annotation.y + annotation.height
    );
  }
  if (annotation.type === 'text') {
    const lines = annotation.text.split('\n');
    const width =
      Math.max(1, ...lines.map((line) => line.length)) * annotation.size * 0.58;
    const height = Math.max(1, lines.length) * annotation.size * 1.3;
    return (
      point.x >= annotation.x - 4 &&
      point.x <= annotation.x + width + 4 &&
      point.y <= annotation.y &&
      point.y >= annotation.y - height
    );
  }
  if (annotation.type === 'ink') {
    const threshold = annotation.width / 2 + 6 / zoom;
    for (let index = 1; index < annotation.points.length; index += 1) {
      if (distanceToSegment(point, annotation.points[index - 1], annotation.points[index]) <= threshold) {
        return true;
      }
    }
  }
  return false;
}

function Thumbnail({
  pdf,
  pageNumber,
  active,
  onSelect,
}: {
  pdf: PDFDocumentProxy;
  pageNumber: number;
  active: boolean;
  onSelect: () => void;
}) {
  const wrapRef = useRef<HTMLSpanElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [rendered, setRendered] = useState(false);

  useEffect(() => {
    let cancelled = false;
    let started = false;
    const element = wrapRef.current;
    if (!element) return;

    const draw = async () => {
      if (started) return;
      started = true;
      try {
        const page = await pdf.getPage(pageNumber);
        if (cancelled) return;
        const base = page.getViewport({ scale: 1 });
        const scale = 132 / base.width;
        await renderPage(page, canvasRef.current as HTMLCanvasElement, scale);
        if (cancelled) return;
        if (canvasRef.current) {
          canvasRef.current.style.width = '100%';
          canvasRef.current.style.height = 'auto';
        }
        setRendered(true);
      } catch {
        // Keep the placeholder if a thumbnail cannot be rendered.
      }
    };

    const observer = new IntersectionObserver(
      (entries) => {
        if (entries.some((entry) => entry.isIntersecting)) {
          void draw();
          observer.disconnect();
        }
      },
      { rootMargin: '300px 0px' },
    );
    observer.observe(element);

    return () => {
      cancelled = true;
      observer.disconnect();
    };
  }, [pdf, pageNumber]);

  return (
    <button
      type="button"
      className={`thumb ${active ? 'active' : ''}`}
      onClick={onSelect}
      aria-label={`Go to page ${pageNumber}`}
    >
      <span className="thumb-canvas-wrap" ref={wrapRef}>
        {!rendered && <span className="thumb-placeholder">{pageNumber}</span>}
        <canvas ref={canvasRef} style={{ display: rendered ? 'block' : 'none' }} />
      </span>
      <span className="thumb-num">{pageNumber}</span>
    </button>
  );
}

export function PdfEditor({
  doc,
  notify,
  onBack,
  onRename,
  onPersist,
  onOpenFile,
  onEditAsDocument,
}: PdfEditorProps) {
  const [pdf, setPdf] = useState<PDFDocumentProxy | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [numPages, setNumPages] = useState(0);
  const [page, setPage] = useState(doc.page ?? 1);
  const [zoom, setZoom] = useState(doc.zoom && doc.zoom > 0 ? doc.zoom : 1);
  const [viewport, setViewport] = useState<PageViewport | null>(null);
  const [pageTexts, setPageTexts] = useState<PageText[] | null>(null);
  const [indexing, setIndexing] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [matches, setMatches] = useState<SearchMatch[]>([]);
  const [matchIndex, setMatchIndex] = useState(0);
  const [tool, setTool] = useState<PdfTool>('select');
  const [highlightColor, setHighlightColor] = useState('#f2cf56');
  const [inkColor, setInkColor] = useState('#23211d');
  const [textColor, setTextColor] = useState('#23211d');
  const [whiteoutColor, setWhiteoutColor] = useState('#ffffff');
  const [textSerif, setTextSerif] = useState(false);
  const [textBold, setTextBold] = useState(false);
  const [textItalic, setTextItalic] = useState(false);
  const [inkWidth, setInkWidth] = useState(2.5);
  const [textSize, setTextSize] = useState(14);
  const [history, setHistory] = useState<PdfAnnotation[][]>([doc.annotations ?? []]);
  const [historyIndex, setHistoryIndex] = useState(0);
  const [status, setStatus] = useState<SaveStatus>('saved');
  const [rendering, setRendering] = useState(true);
  const [name, setName] = useState(doc.name);
  const [draftHighlight, setDraftHighlight] = useState<{ start: Point; end: Point } | null>(null);
  const [draftWhiteout, setDraftWhiteout] = useState<{ start: Point; end: Point } | null>(null);
  const [draftInk, setDraftInk] = useState<Point[] | null>(null);
  const [textDraft, setTextDraft] = useState<TextDraft | null>(null);
  const [textValue, setTextValue] = useState('');

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const viewerRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);
  const openFileRef = useRef<HTMLInputElement>(null);
  const renderTaskRef = useRef<{ cancel: () => void } | null>(null);
  const saveTimer = useRef<number | null>(null);
  const skipFirstSave = useRef(true);
  const fitted = useRef(false);
  const textDraftRef = useRef<TextDraft | null>(null);
  const textValueRef = useRef('');
  const highlightDraftRef = useRef<{ start: Point; end: Point } | null>(null);
  const whiteoutDraftRef = useRef<{ start: Point; end: Point } | null>(null);
  const inkDraftRef = useRef<Point[] | null>(null);
  const annotationsRef = useRef<PdfAnnotation[]>([]);
  const pageRef = useRef(page);
  const zoomRef = useRef(zoom);
  const persistRef = useRef(onPersist);
  const notifyRef = useRef(notify);

  persistRef.current = onPersist;
  notifyRef.current = notify;
  pageRef.current = page;
  zoomRef.current = zoom;
  annotationsRef.current = history[historyIndex] ?? [];

  const annotations = history[historyIndex] ?? [];
  const pageAnnotations = useMemo(
    () => annotations.filter((annotation) => annotation.page === page),
    [annotations, page],
  );

  useEffect(() => {
    setName(doc.name);
  }, [doc.name]);

  useEffect(() => {
    let cancelled = false;
    let loaded: PDFDocumentProxy | null = null;
    setLoadError(null);
    setIndexing(true);
    setPageTexts(null);

    loadPdf(doc.bytes)
      .then((document) => {
        if (cancelled) {
          void document.destroy();
          return;
        }
        loaded = document;
        setPdf(document);
        setNumPages(document.numPages);
        setPage((current) => clamp(current, 1, document.numPages));
      })
      .catch(() => {
        if (!cancelled) {
          setLoadError('This PDF could not be opened. It may be damaged or password protected.');
          notifyRef.current('This PDF could not be opened.', 'error');
        }
      });

    return () => {
      cancelled = true;
      if (loaded) void loaded.destroy();
    };
  }, [doc.bytes]);

  useEffect(() => {
    if (!pdf || fitted.current) return;
    fitted.current = true;
    if (doc.zoom) {
      setZoom(doc.zoom);
      return;
    }
    let cancelled = false;
    void pdf.getPage(1).then((firstPage) => {
      if (cancelled) return;
      const width = firstPage.getViewport({ scale: 1 }).width;
      const container = viewerRef.current?.clientWidth ?? 900;
      setZoom(clamp((container - 64) / width, 0.4, 1.7));
    });
    return () => {
      cancelled = true;
    };
  }, [pdf, doc.zoom]);

  useEffect(() => {
    if (!pdf || !canvasRef.current) return;
    let cancelled = false;
    const canvas = canvasRef.current;
    setRendering(true);

    void (async () => {
      try {
        const pdfPage = await pdf.getPage(page);
        if (cancelled) return;
        const nextViewport = pdfPage.getViewport({ scale: zoom });
        const ratio = Math.min(window.devicePixelRatio || 1, 2);
        canvas.width = Math.max(1, Math.floor(nextViewport.width * ratio));
        canvas.height = Math.max(1, Math.floor(nextViewport.height * ratio));
        canvas.style.width = `${nextViewport.width}px`;
        canvas.style.height = `${nextViewport.height}px`;
        const context = canvas.getContext('2d');
        if (!context) return;
        const task = pdfPage.render({
          canvasContext: context,
          viewport: nextViewport,
          transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined,
        } as Parameters<typeof pdfPage.render>[0]);
        renderTaskRef.current = task;
        await task.promise;
        if (cancelled) return;
        setViewport(nextViewport);
        setRendering(false);
      } catch (error) {
        if ((error as Error)?.name !== 'RenderingCancelledException' && !cancelled) {
          setRendering(false);
          notifyRef.current('This page could not be rendered.', 'error');
        }
      }
    })();

    return () => {
      cancelled = true;
      renderTaskRef.current?.cancel();
    };
  }, [pdf, page, zoom]);

  useEffect(() => {
    if (!pdf) return;
    let cancelled = false;
    void (async () => {
      const collected: PageText[] = [];
      try {
        for (let index = 1; index <= pdf.numPages; index += 1) {
          const pdfPage = await pdf.getPage(index);
          const text = await extractPageText(pdfPage);
          if (cancelled) return;
          collected.push(text);
          if (index === pdf.numPages || index % 4 === 0) {
            setPageTexts([...collected]);
          }
        }
      } catch {
        if (!cancelled) {
          setPageTexts(collected);
          notifyRef.current('Text search could not be prepared for this PDF.', 'error');
        }
      } finally {
        if (!cancelled) setIndexing(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [pdf]);

  useEffect(() => {
    if (!pageTexts || indexing) {
      setMatches([]);
      return;
    }
    const query = searchQuery.trim();
    if (!query) {
      setMatches([]);
      setMatchIndex(0);
      return;
    }

    const pattern = new RegExp(escapeRegExp(query).replace(/\s+/g, '\\s+'), 'gi');
    const results: SearchMatch[] = [];

    pageTexts.forEach((pageText, pageIndex) => {
      let count = 0;
      let match: RegExpExecArray | null;
      while ((match = pattern.exec(pageText.text)) !== null && count < 200) {
        const start = match.index;
        const end = start + match[0].length;
        const rects = pageText.items
          .filter((item) => item.end > start && item.start < end)
          .map((item) => ({ x: item.x, y: item.y, width: item.width, height: item.height }));
        if (rects.length > 0) {
          results.push({ page: pageIndex + 1, start, rects });
        }
        count += 1;
        if (match[0].length === 0) pattern.lastIndex += 1;
      }
    });

    setMatches(results);
    setMatchIndex(0);
    if (results[0]) setPage(results[0].page);
  }, [pageTexts, indexing, searchQuery]);

  useEffect(() => {
    if (skipFirstSave.current) {
      skipFirstSave.current = false;
      return;
    }
    setStatus('saving');
    if (saveTimer.current) window.clearTimeout(saveTimer.current);
    saveTimer.current = window.setTimeout(() => {
      void persistRef
        .current({ annotations, page, zoom })
        .then(() => setStatus('saved'))
        .catch(() => {
          setStatus('error');
          notifyRef.current('Annotations could not be saved locally.', 'error');
        });
    }, 550);
    return () => {
      if (saveTimer.current) window.clearTimeout(saveTimer.current);
    };
  }, [annotations, page, zoom]);

  useEffect(() => {
    const flush = () => {
      if (saveTimer.current) {
        window.clearTimeout(saveTimer.current);
        saveTimer.current = null;
      }
      void persistRef
        .current({
          annotations: annotationsRef.current,
          page: pageRef.current,
          zoom: zoomRef.current,
        })
        .catch(() => undefined);
    };
    const onVisibility = () => {
      if (document.visibilityState === 'hidden') flush();
    };
    document.addEventListener('visibilitychange', onVisibility);
    window.addEventListener('beforeunload', flush);
    return () => {
      document.removeEventListener('visibilitychange', onVisibility);
      window.removeEventListener('beforeunload', flush);
      flush();
    };
  }, []);

  useEffect(() => {
    viewerRef.current?.scrollTo({ top: 0, left: 0 });
  }, [page]);

  useEffect(() => {
    if (textDraft) {
      const area = textareaRef.current;
      window.setTimeout(() => area?.focus(), 0);
    }
  }, [textDraft]);

  const commitAnnotations = useCallback(
    (next: PdfAnnotation[]) => {
      const sliced = history.slice(0, historyIndex + 1);
      const updated = [...sliced, next];
      setHistory(updated);
      setHistoryIndex(updated.length - 1);
    },
    [history, historyIndex],
  );

  const undo = useCallback(() => {
    setHistoryIndex((current) => Math.max(0, current - 1));
  }, []);

  const redo = useCallback(() => {
    setHistoryIndex((current) => Math.min(history.length - 1, current + 1));
  }, [history.length]);

  useEffect(() => {
    const handler = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      const typing =
        target?.tagName === 'INPUT' || target?.tagName === 'TEXTAREA' || target?.isContentEditable;
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z' && !typing) {
        event.preventDefault();
        if (event.shiftKey) redo();
        else undo();
      }
      if (event.key === 'Escape') {
        textDraftRef.current = null;
        setTextDraft(null);
        setTextValue('');
        textValueRef.current = '';
        whiteoutDraftRef.current = null;
        setDraftWhiteout(null);
        highlightDraftRef.current = null;
        setDraftHighlight(null);
        inkDraftRef.current = null;
        setDraftInk(null);
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [redo, undo]);

  const commitName = () => {
    const trimmed = name.trim();
    if (trimmed && trimmed !== doc.name) onRename(trimmed);
    else setName(doc.name);
  };

  const goToPage = (next: number) => {
    setPage(clamp(next, 1, Math.max(1, numPages)));
  };

  const gotoMatch = (delta: number) => {
    if (matches.length === 0) return;
    const next = (matchIndex + delta + matches.length) % matches.length;
    setMatchIndex(next);
    const target = matches[next];
    if (target && target.page !== page) setPage(target.page);
  };

  const toPdfPoint = (event: ReactPointerEvent<HTMLDivElement>): Point | null => {
    if (!viewport) return null;
    const rect = event.currentTarget.getBoundingClientRect();
    const [x, y] = viewport.convertToPdfPoint(event.clientX - rect.left, event.clientY - rect.top);
    return { x, y };
  };

  const openTextDraft = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!viewport) return;
    const rect = event.currentTarget.getBoundingClientRect();
    const leftOnPage = event.clientX - rect.left;
    const topOnPage = event.clientY - rect.top;
    const left = clamp(leftOnPage, 6, Math.max(6, viewport.width - 200));
    const top = clamp(topOnPage, 6, Math.max(6, viewport.height - 60));
    const [x, y] = viewport.convertToPdfPoint(left, top);

    const sampledBg = sampleBgColorAtPoint(canvasRef.current, leftOnPage, topOnPage, viewport);

    const draft: TextDraft = {
      page,
      x,
      y,
      left,
      top,
      fontSize: textSize,
      serif: textSerif,
      bold: textBold,
      italic: textItalic,
      textColor: textColor,
      bgColor: sampledBg,
    };
    textDraftRef.current = draft;
    setTextDraft(draft);
    setTextValue('');
    textValueRef.current = '';
  };

  const handleEditExistingItem = (item: TextItemBox) => {
    if (!viewport) return;
    const left = item.x * zoom;
    const top = item.y * zoom;
    const [x, y] = viewport.convertToPdfPoint(left, top);

    const colors = sampleColorsAtItem(canvasRef.current, item, zoom, viewport);
    const itemFontSize = Math.round(item.fontSize || item.height || 14);

    const draft: TextDraft = {
      page,
      x,
      y,
      left: clamp(left, 6, Math.max(6, viewport.width - 220)),
      top: clamp(top, 6, Math.max(6, viewport.height - 60)),
      originalItem: item,
      fontSize: itemFontSize,
      serif: item.serif ?? false,
      bold: item.bold ?? false,
      italic: item.italic ?? false,
      textColor: colors.textColor,
      bgColor: colors.bgColor,
    };
    setTextSize(itemFontSize);
    setTextColor(colors.textColor);
    setWhiteoutColor(colors.bgColor);
    setTextSerif(item.serif ?? false);
    setTextBold(item.bold ?? false);
    setTextItalic(item.italic ?? false);

    textDraftRef.current = draft;
    setTextDraft(draft);
    setTextValue(item.str);
    textValueRef.current = item.str;
  };

  const finishTextDraft = () => {
    const draft = textDraftRef.current;
    if (!draft) return;
    textDraftRef.current = null;
    const value = textValueRef.current.trim();
    setTextDraft(null);
    setTextValue('');
    textValueRef.current = '';

    // If editing existing annotation
    if (draft.editingAnnotationId) {
      if (!value) {
        commitAnnotations(annotationsRef.current.filter((a) => a.id !== draft.editingAnnotationId));
      } else {
        commitAnnotations(
          annotationsRef.current.map((a) =>
            a.id === draft.editingAnnotationId && a.type === 'text'
              ? {
                  ...a,
                  text: value,
                  size: draft.fontSize,
                  color: draft.textColor,
                  serif: draft.serif,
                  bold: draft.bold,
                  italic: draft.italic,
                  background: draft.bgColor,
                }
              : a,
          ),
        );
      }
      return;
    }

    // If editing an existing text item from the PDF
    if (draft.originalItem && viewport) {
      const item = draft.originalItem;
      const [pdfX1, pdfY1] = viewport.convertToPdfPoint(item.x * zoom, item.y * zoom);
      const [pdfX2, pdfY2] = viewport.convertToPdfPoint(
        (item.x + item.width) * zoom,
        (item.y + item.height) * zoom,
      );
      const whiteoutX = Math.min(pdfX1, pdfX2) - 1.5;
      const whiteoutY = Math.min(pdfY1, pdfY2) - 1.5;
      const whiteoutWidth = Math.abs(pdfX2 - pdfX1) + 3;
      const whiteoutHeight = Math.abs(pdfY2 - pdfY1) + 3;

      const whiteoutAnno: PdfAnnotation = {
        id: uid(),
        type: 'whiteout',
        page: draft.page,
        x: whiteoutX,
        y: whiteoutY,
        width: whiteoutWidth,
        height: whiteoutHeight,
        color: draft.bgColor || '#ffffff',
      };

      if (!value) {
        commitAnnotations([...annotationsRef.current, whiteoutAnno]);
      } else {
        const textAnno: PdfAnnotation = {
          id: uid(),
          type: 'text',
          page: draft.page,
          x: draft.x,
          y: draft.y,
          text: value,
          size: draft.fontSize,
          color: draft.textColor,
          serif: draft.serif,
          bold: draft.bold,
          italic: draft.italic,
          background: draft.bgColor,
        };
        commitAnnotations([...annotationsRef.current, whiteoutAnno, textAnno]);
      }
      return;
    }

    if (!value) return;
    commitAnnotations([
      ...annotationsRef.current,
      {
        id: uid(),
        type: 'text',
        page: draft.page,
        x: draft.x,
        y: draft.y,
        text: value,
        size: draft.fontSize,
        color: draft.textColor,
        serif: draft.serif,
        bold: draft.bold,
        italic: draft.italic,
        background: draft.bgColor,
      },
    ]);
  };

  const cancelTextDraft = () => {
    textDraftRef.current = null;
    setTextDraft(null);
    setTextValue('');
    textValueRef.current = '';
  };

  const eraseAt = (point: Point) => {
    const target = [...pageAnnotations].reverse().find((annotation) => hitTest(annotation, point, zoom));
    if (!target) return;
    commitAnnotations(annotations.filter((annotation) => annotation.id !== target.id));
  };

  const handlePointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (!viewport || tool === 'select') return;
    const point = toPdfPoint(event);
    if (!point) return;

    if (tool === 'edit-text') {
      const clickedAnno = [...pageAnnotations].reverse().find(
        (a) => a.type === 'text' && hitTest(a, point, zoom),
      );
      if (clickedAnno && clickedAnno.type === 'text') {
        const [left, top] = viewport.convertToViewportPoint(clickedAnno.x, clickedAnno.y);
        const draft: TextDraft = {
          page,
          x: clickedAnno.x,
          y: clickedAnno.y,
          left,
          top,
          editingAnnotationId: clickedAnno.id,
          fontSize: clickedAnno.size,
          serif: clickedAnno.serif ?? false,
          bold: clickedAnno.bold ?? false,
          italic: clickedAnno.italic ?? false,
          textColor: clickedAnno.color,
          bgColor: clickedAnno.background || '#ffffff',
        };
        setTextSize(clickedAnno.size);
        setTextColor(clickedAnno.color);
        setTextSerif(clickedAnno.serif ?? false);
        setTextBold(clickedAnno.bold ?? false);
        setTextItalic(clickedAnno.italic ?? false);
        textDraftRef.current = draft;
        setTextDraft(draft);
        setTextValue(clickedAnno.text);
        textValueRef.current = clickedAnno.text;
        return;
      }
      return;
    }

    if (tool === 'text') {
      openTextDraft(event);
      return;
    }
    if (tool === 'eraser') {
      eraseAt(point);
      return;
    }

    event.currentTarget.setPointerCapture(event.pointerId);
    if (tool === 'whiteout') {
      const rect = event.currentTarget.getBoundingClientRect();
      const leftOnPage = event.clientX - rect.left;
      const topOnPage = event.clientY - rect.top;
      const sampledBg = sampleBgColorAtPoint(canvasRef.current, leftOnPage, topOnPage, viewport);
      if (sampledBg) {
        setWhiteoutColor(sampledBg);
      }
      const draft = { start: point, end: point };
      whiteoutDraftRef.current = draft;
      setDraftWhiteout(draft);
    } else if (tool === 'highlight') {
      const draft = { start: point, end: point };
      highlightDraftRef.current = draft;
      setDraftHighlight(draft);
    } else if (tool === 'ink') {
      const draft = [point];
      inkDraftRef.current = draft;
      setDraftInk(draft);
    }
  };

  const handlePointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const point = toPdfPoint(event);
    if (!point) return;

    if (whiteoutDraftRef.current) {
      const draft = { ...whiteoutDraftRef.current, end: point };
      whiteoutDraftRef.current = draft;
      setDraftWhiteout(draft);
    }
    if (highlightDraftRef.current) {
      const draft = { ...highlightDraftRef.current, end: point };
      highlightDraftRef.current = draft;
      setDraftHighlight(draft);
    }
    if (inkDraftRef.current) {
      const draft = [...inkDraftRef.current, point];
      inkDraftRef.current = draft;
      setDraftInk(draft);
    }
  };

  const handlePointerUp = () => {
    if (whiteoutDraftRef.current) {
      const { start, end } = whiteoutDraftRef.current;
      whiteoutDraftRef.current = null;
      setDraftWhiteout(null);
      const x = Math.min(start.x, end.x);
      const y = Math.min(start.y, end.y);
      const width = Math.abs(end.x - start.x);
      const height = Math.abs(end.y - start.y);
      if (width > 2 && height > 2) {
        commitAnnotations([
          ...annotationsRef.current,
          {
            id: uid(),
            type: 'whiteout',
            page,
            x,
            y,
            width,
            height,
            color: whiteoutColor,
          },
        ]);
      }
    }

    if (highlightDraftRef.current) {
      const { start, end } = highlightDraftRef.current;
      highlightDraftRef.current = null;
      setDraftHighlight(null);
      const x = Math.min(start.x, end.x);
      const y = Math.min(start.y, end.y);
      const width = Math.abs(end.x - start.x);
      const height = Math.abs(end.y - start.y);
      if (width > 2 && height > 2) {
        commitAnnotations([
          ...annotationsRef.current,
          {
            id: uid(),
            type: 'highlight',
            page,
            x,
            y,
            width,
            height,
            color: highlightColor,
            opacity: 0.36,
          },
        ]);
      }
    }

    if (inkDraftRef.current) {
      const points = inkDraftRef.current;
      inkDraftRef.current = null;
      setDraftInk(null);
      if (points.length > 1) {
        commitAnnotations([
          ...annotationsRef.current,
          {
            id: uid(),
            type: 'ink',
            page,
            points,
            width: inkWidth,
            color: inkColor,
          },
        ]);
      }
    }
  };

  const handleDownload = async () => {
    if (!pdf) return;
    setStatus('saving');
    try {
      const bytes = await exportAnnotatedPdf(doc.bytes, annotations, baseName(name));
      saveAs(new Blob([new Uint8Array(bytes)], { type: 'application/pdf' }), withExtension(name, 'pdf'));
      setStatus('saved');
      notify('PDF downloaded with your marks.', 'success');
    } catch {
      setStatus('error');
      notify('This PDF could not be exported. It may be encrypted.', 'error');
    }
  };

  const handleExportWord = async () => {
    if (!pdf) return;
    setStatus('saving');
    try {
      let texts: PageText[];
      if (!pageTexts || pageTexts.length < numPages) {
        texts = await extractDocumentText(pdf);
      } else {
        texts = pageTexts;
      }
      const pages = texts.map((item) => item.text.trim());
      if (!pages.some(Boolean)) {
        setStatus('saved');
        notify('This PDF has no selectable text to export. It may be a scanned document.', 'error');
        return;
      }
      const title = baseName(name);
      const content = pagesToDocxContent(title, pages);
      const blob = await exportDocx(content, title);
      saveAs(blob, `${title}-text.docx`);
      setStatus('saved');
      notify('Text exported to a Word document.', 'success');
    } catch {
      setStatus('error');
      notify('The text could not be exported.', 'error');
    }
  };

  const clearPage = () => {
    if (pageAnnotations.length === 0) return;
    if (!window.confirm(`Remove all ${pageAnnotations.length} mark(s) from page ${page}?`)) return;
    commitAnnotations(annotations.filter((annotation) => annotation.page !== page));
  };

  const fitWidth = () => {
    if (!viewport) return;
    const baseWidth = viewport.width / zoom;
    const available = (viewerRef.current?.clientWidth ?? 900) - 64;
    setZoom(clamp(available / baseWidth, 0.3, 3));
  };

  const handleConvertToDocument = async () => {
    if (!pdf) return;
    setStatus('saving');
    try {
      notify('Converting PDF into editable document…');
      const html = await extractPdfAsHtml(pdf);
      setStatus('saved');
      if (onEditAsDocument) {
        onEditAsDocument(html);
      }
    } catch {
      setStatus('error');
      notify('Could not convert PDF to editable document.', 'error');
    }
  };

  const currentMatch = matches[matchIndex];
  const currentText = pageTexts?.[page - 1];
  const activeColor =
    tool === 'highlight'
      ? highlightColor
      : tool === 'ink'
        ? inkColor
        : tool === 'whiteout'
          ? whiteoutColor
          : textColor;
  const palette =
    tool === 'highlight'
      ? HIGHLIGHT_COLORS
      : tool === 'ink'
        ? INK_COLORS
        : tool === 'whiteout'
          ? WHITEOUT_COLORS
          : TEXT_COLORS;
  const setActiveColor = (value: string) => {
    if (tool === 'highlight') setHighlightColor(value);
    else if (tool === 'ink') setInkColor(value);
    else if (tool === 'whiteout') setWhiteoutColor(value);
    else setTextColor(value);
  };

  const pageWidth = viewport?.width ?? 740;
  const pageHeight = viewport?.height ?? 960;

  const highlightDraftRect = draftHighlight
    ? (() => {
        if (!viewport) return null;
        const [x1, y1] = viewport.convertToViewportPoint(draftHighlight.start.x, draftHighlight.start.y);
        const [x2, y2] = viewport.convertToViewportPoint(draftHighlight.end.x, draftHighlight.end.y);
        return {
          left: Math.min(x1, x2),
          top: Math.min(y1, y2),
          width: Math.abs(x2 - x1),
          height: Math.abs(y2 - y1),
        };
      })()
    : null;

  const whiteoutDraftRect = draftWhiteout
    ? (() => {
        if (!viewport) return null;
        const [x1, y1] = viewport.convertToViewportPoint(draftWhiteout.start.x, draftWhiteout.start.y);
        const [x2, y2] = viewport.convertToViewportPoint(draftWhiteout.end.x, draftWhiteout.end.y);
        return {
          left: Math.min(x1, x2),
          top: Math.min(y1, y2),
          width: Math.abs(x2 - x1),
          height: Math.abs(y2 - y1),
        };
      })()
    : null;

  const draftInkPoints =
    draftInk && viewport
      ? draftInk
          .map((point) => {
            const [x, y] = viewport.convertToViewportPoint(point.x, point.y);
            return `${x},${y}`;
          })
          .join(' ')
      : '';

  return (
    <div className="app">
      <header className="app-header">
        <button type="button" className="icon-btn" title="Back to library" onClick={onBack}>
          <ChevronLeft size={18} />
        </button>
        <div className="brand" style={{ gap: 9 }}>
          <span className="brand-mark" style={{ width: 30, height: 30, fontSize: 18 }}>
            P
          </span>
        </div>
        <div className="header-divider" />
        <div className="doc-title-wrap">
          <input
            className="doc-title-input"
            value={name}
            aria-label="Document name"
            spellCheck={false}
            onChange={(event) => setName(event.target.value)}
            onBlur={commitName}
            onKeyDown={(event) => {
              if (event.key === 'Enter') event.currentTarget.blur();
            }}
          />
          <span className="chip chip-pdf">PDF</span>
        </div>
        <div className="header-actions">
          <span className={`save-indicator ${status === 'saving' ? 'saving' : ''}`}>
            <span className="dot" style={status === 'error' ? { background: '#b03a25' } : undefined} />
            {status === 'saving' ? 'Saving…' : status === 'error' ? 'Not saved' : 'Saved locally'}
          </span>
          <input
            ref={openFileRef}
            type="file"
            accept={ACCEPTED_EXTENSIONS}
            hidden
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) onOpenFile(file);
              event.target.value = '';
            }}
          />
          <button
            type="button"
            className="btn btn-sm"
            title="Open another document"
            onClick={() => openFileRef.current?.click()}
          >
            <FolderOpen size={14} />
            <span>Open</span>
          </button>
          <button
            type="button"
            className="btn btn-sm"
            style={{
              background: 'var(--accent-soft)',
              color: 'var(--accent-deep)',
              borderColor: 'rgba(196, 81, 45, 0.3)',
              fontWeight: 500,
            }}
            onClick={handleConvertToDocument}
            title="Convert and edit complete document text with full formatting tools"
          >
            <FileEdit size={14} />
            <span>Edit as Document</span>
          </button>
          <button type="button" className="btn btn-sm" onClick={handleExportWord} title="Extract text to Word">
            <FileOutput size={14} />
            <span>To Word</span>
          </button>
          <button type="button" className="btn btn-sm btn-primary" onClick={handleDownload}>
            <Download size={14} />
            <span>Download PDF</span>
          </button>
        </div>
      </header>

      <div className="pdf-body">
        <aside className="pdf-sidebar">
          <div className="sidebar-head">
            <span>Pages</span>
            <span>{numPages || '—'}</span>
          </div>
          <div className="thumbs">
            {pdf &&
              Array.from({ length: numPages }, (_, index) => (
                <Thumbnail
                  key={index + 1}
                  pdf={pdf}
                  pageNumber={index + 1}
                  active={page === index + 1}
                  onSelect={() => setPage(index + 1)}
                />
              ))}
          </div>
        </aside>

        <div className="pdf-main">
          <div className="pdf-toolbar" role="toolbar" aria-label="PDF tools">
            <div className="page-nav">
              <button
                type="button"
                className="icon-btn"
                style={{ width: 26, height: 26 }}
                title="Previous page"
                disabled={page <= 1}
                onClick={() => goToPage(page - 1)}
              >
                <ChevronLeft size={14} />
              </button>
              <input
                className="page-input"
                type="number"
                min={1}
                max={numPages || 1}
                value={page}
                aria-label="Current page"
                onChange={(event) => {
                  const value = Number.parseInt(event.target.value, 10);
                  if (Number.isFinite(value)) goToPage(value);
                }}
              />
              <span className="page-total">/ {numPages || '—'}</span>
              <button
                type="button"
                className="icon-btn"
                style={{ width: 26, height: 26 }}
                title="Next page"
                disabled={page >= numPages}
                onClick={() => goToPage(page + 1)}
              >
                <ChevronRight size={14} />
              </button>
            </div>

            <div className="tool-divider" />
            <div className="tool-group">
              <ToolButton title="Select text" active={tool === 'select'} onClick={() => setTool('select')}>
                <MousePointer2 size={15} />
              </ToolButton>
              <ToolButton
                title="Edit PDF text (Click any text on the page to edit or replace it)"
                active={tool === 'edit-text'}
                onClick={() => setTool('edit-text')}
              >
                <Edit3 size={15} />
              </ToolButton>
              <ToolButton
                title="Add text"
                className="tool-text"
                active={tool === 'text'}
                onClick={() => setTool('text')}
              >
                <Type size={15} />
              </ToolButton>
              <ToolButton
                title="Whiteout / Erase (Drag a box to cover text or graphics)"
                active={tool === 'whiteout'}
                onClick={() => setTool('whiteout')}
              >
                <Square size={15} />
              </ToolButton>
              <ToolButton
                title="Highlight"
                active={tool === 'highlight'}
                onClick={() => setTool('highlight')}
              >
                <Highlighter size={15} />
              </ToolButton>
              <ToolButton title="Draw" active={tool === 'ink'} onClick={() => setTool('ink')}>
                <PenTool size={15} />
              </ToolButton>
              <ToolButton title="Erase a mark" active={tool === 'eraser'} onClick={() => setTool('eraser')}>
                <Eraser size={15} />
              </ToolButton>
            </div>

            {(tool === 'highlight' || tool === 'ink' || tool === 'text' || tool === 'edit-text' || tool === 'whiteout') && (
              <>
                <div className="tool-divider" />
                <div className="tool-group" style={{ gap: 5, padding: '0 3px' }}>
                  {palette.map((color) => (
                    <button
                      key={color}
                      type="button"
                      className={`swatch ${activeColor === color ? 'selected' : ''}`}
                      style={{
                        background: color,
                        width: 18,
                        height: 18,
                        boxShadow: color === '#ffffff' ? 'inset 0 0 0 1px #ded6c7' : undefined,
                      }}
                      title={color === '#ffffff' ? 'White' : color}
                      aria-label={`Use color ${color}`}
                      onClick={() => setActiveColor(color)}
                    />
                  ))}
                </div>
                {tool === 'ink' && (
                  <div className="tool-group" style={{ gap: 2, marginLeft: 4 }}>
                    {[1.5, 2.5, 4.5].map((width) => (
                      <button
                        key={width}
                        type="button"
                        className={`icon-btn ${inkWidth === width ? 'active' : ''}`}
                        style={{ width: 28, height: 28 }}
                        title={`Stroke ${width} pt`}
                        onClick={() => setInkWidth(width)}
                      >
                        <span
                          style={{
                            display: 'block',
                            width: width * 2.4 + 3,
                            height: width * 2.4 + 3,
                            borderRadius: '50%',
                            background: 'currentColor',
                          }}
                        />
                      </button>
                    ))}
                  </div>
                )}
                {(tool === 'text' || tool === 'edit-text') && (
                  <div className="tool-group" style={{ marginLeft: 4, gap: 3 }}>
                    <select
                      className="size-select"
                      aria-label="Text size"
                      value={textSize}
                      onChange={(event) => {
                        const val = Number(event.target.value);
                        setTextSize(val);
                        if (textDraftRef.current) {
                          const updated = { ...textDraftRef.current, fontSize: val };
                          textDraftRef.current = updated;
                          setTextDraft(updated);
                        }
                      }}
                      style={{ width: 68 }}
                    >
                      {[9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 38, 48].map((size) => (
                        <option value={size} key={size}>
                          {size} pt
                        </option>
                      ))}
                    </select>
                    <button
                      type="button"
                      className={`icon-btn ${textBold ? 'active' : ''}`}
                      style={{ fontWeight: 700, width: 28, height: 28 }}
                      title="Bold"
                      onClick={() => {
                        const next = !textBold;
                        setTextBold(next);
                        if (textDraftRef.current) {
                          const updated = { ...textDraftRef.current, bold: next };
                          textDraftRef.current = updated;
                          setTextDraft(updated);
                        }
                      }}
                    >
                      B
                    </button>
                    <button
                      type="button"
                      className={`icon-btn ${textItalic ? 'active' : ''}`}
                      style={{ fontStyle: 'italic', width: 28, height: 28 }}
                      title="Italic"
                      onClick={() => {
                        const next = !textItalic;
                        setTextItalic(next);
                        if (textDraftRef.current) {
                          const updated = { ...textDraftRef.current, italic: next };
                          textDraftRef.current = updated;
                          setTextDraft(updated);
                        }
                      }}
                    >
                      I
                    </button>
                    <button
                      type="button"
                      className={`icon-btn ${textSerif ? 'active' : ''}`}
                      style={{ fontFamily: 'Georgia, serif', width: 28, height: 28, fontSize: 13 }}
                      title="Serif font"
                      onClick={() => {
                        const next = !textSerif;
                        setTextSerif(next);
                        if (textDraftRef.current) {
                          const updated = { ...textDraftRef.current, serif: next };
                          textDraftRef.current = updated;
                          setTextDraft(updated);
                        }
                      }}
                    >
                      S
                    </button>
                  </div>
                )}
              </>
            )}

            <div className="tool-divider" />
            <div className="tool-group">
              <ToolButton title="Undo mark (Ctrl+Z)" disabled={historyIndex <= 0} onClick={undo}>
                <Undo2 size={15} />
              </ToolButton>
              <ToolButton
                title="Redo mark (Ctrl+Shift+Z)"
                disabled={historyIndex >= history.length - 1}
                onClick={redo}
              >
                <Redo2 size={15} />
              </ToolButton>
              <ToolButton
                title="Clear marks on this page"
                disabled={pageAnnotations.length === 0}
                onClick={clearPage}
              >
                <Trash2 size={15} />
              </ToolButton>
            </div>

            <div className="tool-divider" />
            <div className="search-box">
              <Search size={13} />
              <input
                ref={searchInputRef}
                type="search"
                placeholder={indexing ? 'Indexing pages…' : 'Search document'}
                value={searchQuery}
                disabled={indexing}
                aria-label="Search document"
                onChange={(event) => setSearchQuery(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    gotoMatch(event.shiftKey ? -1 : 1);
                  }
                  if (event.key === 'Escape') {
                    setSearchQuery('');
                    event.currentTarget.blur();
                  }
                }}
              />
              {indexing ? (
                <Loader2 size={12} className="spin" style={{ marginRight: 6 }} />
              ) : searchQuery ? (
                <>
                  <span className="search-count">
                    {matches.length > 0 ? `${matchIndex + 1}/${matches.length}` : '0/0'}
                  </span>
                  <span className="search-nav">
                    <button type="button" className="icon-btn" title="Previous match" onClick={() => gotoMatch(-1)}>
                      <ChevronLeft size={12} />
                    </button>
                    <button type="button" className="icon-btn" title="Next match" onClick={() => gotoMatch(1)}>
                      <ChevronRight size={12} />
                    </button>
                    <button
                      type="button"
                      className="icon-btn"
                      title="Clear search"
                      onClick={() => {
                        setSearchQuery('');
                        searchInputRef.current?.focus();
                      }}
                    >
                      <X size={12} />
                    </button>
                  </span>
                </>
              ) : null}
            </div>

            <div className="tool-divider" />
            <div className="tool-group">
              <button
                type="button"
                className="icon-btn"
                title="Zoom out"
                onClick={() => setZoom((value) => clamp(Math.round((value - 0.1) * 100) / 100, 0.3, 3))}
              >
                <span style={{ fontSize: 15, lineHeight: 1 }}>−</span>
              </button>
              <span className="zoom-percent">{Math.round(zoom * 100)}%</span>
              <button
                type="button"
                className="icon-btn"
                title="Zoom in"
                onClick={() => setZoom((value) => clamp(Math.round((value + 0.1) * 100) / 100, 0.3, 3))}
              >
                <Plus size={14} />
              </button>
              <button type="button" className="icon-btn" title="Fit width" onClick={fitWidth}>
                <Maximize2 size={14} />
              </button>
            </div>
          </div>

          <div className="pdf-viewer" ref={viewerRef}>
            {loadError ? (
              <div className="pdf-empty" style={{ flexDirection: 'column', gap: 14, textAlign: 'center' }}>
                <AlertTriangle size={26} style={{ color: 'var(--accent)' }} />
                <div style={{ maxWidth: 320 }}>{loadError}</div>
                <button type="button" className="btn btn-sm" onClick={onBack}>
                  Back to library
                </button>
              </div>
            ) : !pdf ? (
              <div className="pdf-empty">
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 9 }}>
                  <Loader2 size={16} className="spin" />
                  Preparing document…
                </span>
              </div>
            ) : (
              <div className="pdf-page-shell">
                <div className="pdf-page" style={{ width: pageWidth, height: pageHeight }}>
                  <canvas ref={canvasRef} />
                  {rendering && (
                    <div
                      style={{
                        position: 'absolute',
                        inset: 0,
                        display: 'grid',
                        placeItems: 'center',
                        background: 'rgba(255,255,255,0.4)',
                        zIndex: 6,
                      }}
                    >
                      <Loader2 size={20} className="spin" style={{ color: 'var(--muted)' }} />
                    </div>
                  )}
                  <div
                    className={`text-layer ${tool === 'select' ? 'selectable' : tool === 'edit-text' ? 'edit-mode' : ''}`}
                  >
                    {currentText?.items.map((item, index) => (
                      <span
                        key={`${item.start}-${index}`}
                        style={{
                          left: item.x * zoom,
                          top: item.y * zoom,
                          fontSize: item.height * zoom,
                          width: item.width * zoom,
                          height: item.height * zoom,
                        }}
                        title={tool === 'edit-text' ? 'Click to edit or replace this text' : undefined}
                        onClick={tool === 'edit-text' ? () => handleEditExistingItem(item) : undefined}
                      >
                        {item.str}
                      </span>
                    ))}
                  </div>

                  <div className="search-layer">
                    {matches
                      .map((match, index) => ({ match, index }))
                      .filter(({ match }) => match.page === page)
                      .flatMap(({ match, index }) =>
                        match.rects.map((rect, rectIndex) => (
                          <div
                            key={`${index}-${rectIndex}`}
                            className={`search-hit ${currentMatch === match ? 'current' : ''}`}
                            style={{
                              left: rect.x * zoom,
                              top: rect.y * zoom,
                              width: rect.width * zoom,
                              height: rect.height * zoom,
                            }}
                          />
                        )),
                      )}
                  </div>

                  <div
                    className={`anno-layer tool-${tool}`}
                    onPointerDown={handlePointerDown}
                    onPointerMove={handlePointerMove}
                    onPointerUp={handlePointerUp}
                    onPointerCancel={handlePointerUp}
                  >
                    {pageAnnotations.map((annotation) => {
                      if (!viewport) return null;
                      if (annotation.type === 'whiteout') {
                        const [x1, y1] = viewport.convertToViewportPoint(annotation.x, annotation.y);
                        const [x2, y2] = viewport.convertToViewportPoint(
                          annotation.x + annotation.width,
                          annotation.y + annotation.height,
                        );
                        return (
                          <div
                            key={annotation.id}
                            className="anno-whiteout"
                            style={{
                              left: Math.min(x1, x2),
                              top: Math.min(y1, y2),
                              width: Math.abs(x2 - x1),
                              height: Math.abs(y2 - y1),
                              background: annotation.color || '#ffffff',
                            }}
                          />
                        );
                      }
                      if (annotation.type === 'highlight') {
                        const [x1, y1] = viewport.convertToViewportPoint(annotation.x, annotation.y);
                        const [x2, y2] = viewport.convertToViewportPoint(
                          annotation.x + annotation.width,
                          annotation.y + annotation.height,
                        );
                        return (
                          <div
                            key={annotation.id}
                            className="anno-highlight"
                            style={{
                              left: Math.min(x1, x2),
                              top: Math.min(y1, y2),
                              width: Math.abs(x2 - x1),
                              height: Math.abs(y2 - y1),
                              background: annotation.color,
                              opacity: annotation.opacity,
                            }}
                          />
                        );
                      }
                      if (annotation.type === 'ink') {
                        const points = annotation.points
                          .map((point) => {
                            const [x, y] = viewport.convertToViewportPoint(point.x, point.y);
                            return `${x},${y}`;
                          })
                          .join(' ');
                        return (
                          <svg key={annotation.id} className="anno-ink">
                            <polyline
                              points={points}
                              fill="none"
                              stroke={annotation.color}
                              strokeWidth={annotation.width * zoom}
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            />
                          </svg>
                        );
                      }
                      if (annotation.type === 'text') {
                        const [x, y] = viewport.convertToViewportPoint(annotation.x, annotation.y);
                        return (
                          <div
                            key={annotation.id}
                            className={`anno-text ${annotation.serif ? 'serif' : 'sans'} ${annotation.bold ? 'bold' : ''} ${annotation.italic ? 'italic' : ''}`}
                            style={{
                              left: x,
                              top: y,
                              fontSize: annotation.size * zoom,
                              color: annotation.color,
                              background: annotation.background || 'transparent',
                              fontWeight: annotation.bold ? 700 : 400,
                              fontStyle: annotation.italic ? 'italic' : 'normal',
                              fontFamily: annotation.serif
                                ? 'Georgia, "Times New Roman", Times, serif'
                                : '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
                            }}
                          >
                            {annotation.text}
                          </div>
                        );
                      }
                      return null;
                    })}

                    {whiteoutDraftRect && (
                      <div
                        className="anno-whiteout"
                        style={{
                          left: whiteoutDraftRect.left,
                          top: whiteoutDraftRect.top,
                          width: whiteoutDraftRect.width,
                          height: whiteoutDraftRect.height,
                          background: whiteoutColor,
                          boxShadow: '0 0 0 1px var(--accent)',
                        }}
                      />
                    )}
                    {highlightDraftRect && (
                      <div
                        className="anno-highlight"
                        style={{
                          left: highlightDraftRect.left,
                          top: highlightDraftRect.top,
                          width: highlightDraftRect.width,
                          height: highlightDraftRect.height,
                          background: highlightColor,
                          opacity: 0.3,
                        }}
                      />
                    )}
                    {draftInk && viewport && (
                      <svg className="anno-ink">
                        <polyline
                          points={draftInkPoints}
                          fill="none"
                          stroke={inkColor}
                          strokeWidth={inkWidth * zoom}
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      </svg>
                    )}
                  </div>

                  {textDraft && (
                    <div
                      className="pdf-inline-editor-wrap"
                      style={{
                        position: 'absolute',
                        left: textDraft.left,
                        top: textDraft.top,
                        zIndex: 20,
                      }}
                    >
                      <div
                        className="pdf-inline-bubble"
                        onMouseDown={(e) => e.preventDefault()}
                      >
                        <button
                          type="button"
                          className={`bubble-btn ${textDraft.bold ? 'active' : ''}`}
                          title="Toggle bold"
                          onClick={() => {
                            const next = !textDraft.bold;
                            const updated = { ...textDraft, bold: next };
                            textDraftRef.current = updated;
                            setTextDraft(updated);
                            setTextBold(next);
                          }}
                        >
                          <Bold size={11} />
                        </button>
                        <button
                          type="button"
                          className={`bubble-btn ${textDraft.italic ? 'active' : ''}`}
                          title="Toggle italic"
                          onClick={() => {
                            const next = !textDraft.italic;
                            const updated = { ...textDraft, italic: next };
                            textDraftRef.current = updated;
                            setTextDraft(updated);
                            setTextItalic(next);
                          }}
                        >
                          <Italic size={11} />
                        </button>
                        <button
                          type="button"
                          className={`bubble-btn ${textDraft.serif ? 'active' : ''}`}
                          title="Toggle serif font"
                          onClick={() => {
                            const next = !textDraft.serif;
                            const updated = { ...textDraft, serif: next };
                            textDraftRef.current = updated;
                            setTextDraft(updated);
                            setTextSerif(next);
                          }}
                        >
                          <span style={{ fontFamily: 'Georgia, serif', fontSize: 11, fontWeight: 600 }}>S</span>
                        </button>
                        <span className="bubble-divider" />
                        <span className="bubble-pill">{Math.round(textDraft.fontSize)} pt</span>
                        <span className="bubble-divider" />
                        <button
                          type="button"
                          className="bubble-btn bubble-btn-save"
                          title="Apply edit (Enter)"
                          onClick={finishTextDraft}
                        >
                          <Check size={12} />
                        </button>
                        <button
                          type="button"
                          className="bubble-btn"
                          title="Cancel (Esc)"
                          onClick={cancelTextDraft}
                        >
                          <X size={12} />
                        </button>
                      </div>

                      <textarea
                        ref={textareaRef}
                        className={`pdf-text-editor ${textDraft.serif ? 'serif' : 'sans'} ${textDraft.bold ? 'bold' : ''} ${textDraft.italic ? 'italic' : ''}`}
                        style={{
                          fontSize: Math.max(10, textDraft.fontSize * zoom),
                          fontWeight: textDraft.bold ? 700 : 400,
                          fontStyle: textDraft.italic ? 'italic' : 'normal',
                          fontFamily: textDraft.serif
                            ? 'Georgia, "Times New Roman", Times, serif'
                            : '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
                          maxWidth: Math.max(220, pageWidth - textDraft.left - 24),
                          color: textDraft.textColor,
                          backgroundColor: textDraft.bgColor,
                        }}
                        value={textValue}
                        placeholder={textDraft.originalItem ? 'Edit text…' : 'Type a note…'}
                        onChange={(event) => {
                          setTextValue(event.target.value);
                          textValueRef.current = event.target.value;
                          event.currentTarget.style.height = 'auto';
                          event.currentTarget.style.height = `${event.currentTarget.scrollHeight}px`;
                        }}
                        onBlur={finishTextDraft}
                        onKeyDown={(event) => {
                          if (event.key === 'Escape') {
                            event.preventDefault();
                            cancelTextDraft();
                          }
                          if (event.key === 'Enter' && !event.shiftKey) {
                            event.preventDefault();
                            finishTextDraft();
                          }
                        }}
                      />
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          <div className="pdf-status">
            <span>
              Page {page} of {numPages || '—'}
            </span>
            <span className="hide-mobile">
              {annotations.length} {annotations.length === 1 ? 'mark' : 'marks'}
            </span>
            <span className="right">
              <span className="tool-hint hide-mobile">
                {tool === 'select'
                  ? 'Select and copy text'
                  : tool === 'edit-text'
                    ? 'Click any text on the page to edit or replace it'
                    : tool === 'text'
                      ? 'Click the page to place new text'
                      : tool === 'whiteout'
                        ? 'Drag a box to erase/whiteout area'
                        : tool === 'highlight'
                          ? 'Drag to highlight'
                          : tool === 'ink'
                            ? 'Draw with your pointer'
                            : 'Click a mark to erase it'}
              </span>
              <span className="status-item">Local annotations</span>
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}
