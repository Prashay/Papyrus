import { useCallback, useEffect, useRef, useState, type ReactNode } from 'react';
import { EditorContent, useEditor, useEditorState } from '@tiptap/react';
import { Extension } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';
import TextAlign from '@tiptap/extension-text-align';
import Link from '@tiptap/extension-link';
import Highlight from '@tiptap/extension-highlight';
import TextStyle from '@tiptap/extension-text-style';
import Color from '@tiptap/extension-color';
import FontFamily from '@tiptap/extension-font-family';
import Image from '@tiptap/extension-image';
import Placeholder from '@tiptap/extension-placeholder';
import Table from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  ChevronLeft,
  Code,
  Download,
  Eraser,
  Eye,
  FileOutput,
  FolderOpen,
  Highlighter,
  ImagePlus,
  Info,
  Italic,
  Link2,
  List,
  ListOrdered,
  Minus,
  Plus,
  Printer,
  Quote,
  Redo2,
  Strikethrough,
  Table2,
  Underline as UnderlineIcon,
  Undo2,
  X,
} from 'lucide-react';
import { saveAs } from 'file-saver';
import type { OpenDocument, StoredState } from '../types';
import { ACCEPTED_EXTENSIONS } from '../lib/detect';
import { exportDocx } from '../lib/docx';
import { exportHtmlToPdf } from '../lib/pdf';
import { baseName, clamp, withExtension } from '../lib/format';
import type { ToastTone } from './ui';

const FontSize = Extension.create({
  name: 'fontSize',
  addOptions() {
    return { types: ['textStyle'] };
  },
  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          fontSize: {
            default: null,
            parseHTML: (element) => element.style.fontSize || null,
            renderHTML: (attributes) =>
              attributes.fontSize ? { style: `font-size: ${attributes.fontSize}` } : {},
          },
        },
      },
    ];
  },
});

const FONT_OPTIONS = [
  { value: '', label: 'Auto (Document)' },
  { value: 'Arial, sans-serif', label: 'Arial' },
  { value: 'Helvetica, Arial, sans-serif', label: 'Helvetica' },
  { value: 'Inter Variable', label: 'Inter' },
  { value: 'Segoe UI, sans-serif', label: 'Segoe UI' },
  { value: 'Newsreader Variable', label: 'Literary (Serif)' },
  { value: 'Georgia', label: 'Georgia' },
  { value: 'Times New Roman', label: 'Times New Roman' },
  { value: 'Courier New', label: 'Courier' },
];

const SIZE_OPTIONS = ['9px', '10px', '11px', '12px', '14px', '16px', '18px', '20px', '24px', '30px', '36px'];

type SaveStatus = 'saved' | 'saving' | 'error';

interface WordEditorProps {
  doc: OpenDocument;
  notify: (message: string, tone?: ToastTone) => void;
  onBack: () => void;
  onRename: (name: string) => void;
  onPersist: (patch: StoredState) => Promise<void>;
  onOpenFile: (file: File) => void;
  onSwitchToPdf?: () => void;
}

function countWords(text: string): number {
  const trimmed = text.trim();
  if (!trimmed) return 0;
  return trimmed.split(/\s+/).length;
}

function toHex(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback;
  if (/^#[0-9a-f]{6}$/i.test(value)) return value;
  if (/^#[0-9a-f]{3}$/i.test(value)) {
    return `#${value
      .slice(1)
      .split('')
      .map((char) => char + char)
      .join('')}`;
  }
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(value);
  if (rgb) {
    const hex = rgb
      .slice(1, 4)
      .map((part) => Number(part).toString(16).padStart(2, '0'))
      .join('');
    return `#${hex}`;
  }
  return fallback;
}

function ToolButton({
  title,
  active,
  disabled,
  onClick,
  children,
}: {
  title: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      className={`tool-btn ${active ? 'active' : ''}`}
      title={title}
      aria-label={title}
      disabled={disabled}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

export function WordEditor({
  doc,
  notify,
  onBack,
  onRename,
  onPersist,
  onOpenFile,
  onSwitchToPdf,
}: WordEditorProps) {
  const [name, setName] = useState(doc.name);
  const [status, setStatus] = useState<SaveStatus>('saved');
  const [zoom, setZoom] = useState(100);
  const [tableMenu, setTableMenu] = useState(false);
  const [showLegacy, setShowLegacy] = useState(Boolean(doc.legacy));
  const [stats, setStats] = useState({ words: 0, characters: 0 });
  const saveTimer = useRef<number | null>(null);
  const latestHtml = useRef(doc.html ?? '');
  const persistRef = useRef(onPersist);
  const notifyRef = useRef(notify);
  const tableMenuRef = useRef<HTMLDivElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const openInputRef = useRef<HTMLInputElement>(null);

  persistRef.current = onPersist;
  notifyRef.current = notify;

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        heading: { levels: [1, 2, 3, 4] },
        codeBlock: { HTMLAttributes: { spellcheck: 'false' } },
      }),
      Underline,
      TextStyle,
      Color,
      FontSize,
      FontFamily,
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Highlight.configure({ multicolor: true }),
      Link.configure({
        openOnClick: false,
        autolink: true,
        linkOnPaste: true,
        HTMLAttributes: { rel: 'noopener noreferrer', target: '_blank' },
      }),
      Image.configure({ allowBase64: true, inline: false }),
      Placeholder.configure({
        placeholder: 'Start writing, or paste something worth keeping…',
      }),
      Table.configure({ resizable: false }),
      TableRow,
      TableHeader,
      TableCell,
    ],
    content: doc.html || '<p></p>',
    editorProps: {
      attributes: {
        spellcheck: 'true',
        'aria-label': 'Document editor',
      },
    },
    onUpdate: ({ editor: instance }) => {
      latestHtml.current = instance.getHTML();
      const text = instance.getText();
      setStats({ words: countWords(text), characters: text.length });
      setStatus('saving');
      if (saveTimer.current) window.clearTimeout(saveTimer.current);
      saveTimer.current = window.setTimeout(() => {
        void persistRef
          .current({ html: latestHtml.current })
          .then(() => setStatus('saved'))
          .catch(() => {
            setStatus('error');
            notifyRef.current('Your changes could not be saved locally.', 'error');
          });
      }, 650);
    },
  });

  useEffect(() => {
    if (!editor) return;
    const text = editor.getText();
    latestHtml.current = editor.getHTML();
    setStats({ words: countWords(text), characters: text.length });
  }, [editor]);

  useEffect(() => {
    setName(doc.name);
  }, [doc.name]);

  useEffect(() => {
    const flush = () => {
      if (saveTimer.current) {
        window.clearTimeout(saveTimer.current);
        saveTimer.current = null;
      }
      void persistRef.current({ html: latestHtml.current }).catch(() => undefined);
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
    if (!tableMenu) return;
    const handler = (event: MouseEvent) => {
      if (!tableMenuRef.current?.contains(event.target as Node)) setTableMenu(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [tableMenu]);

  const ui = useEditorState({
    editor,
    selector: ({ editor: instance }) => {
      if (!instance) return null;
      const textStyle = instance.getAttributes('textStyle');
      return {
        bold: instance.isActive('bold'),
        italic: instance.isActive('italic'),
        underline: instance.isActive('underline'),
        strike: instance.isActive('strike'),
        code: instance.isActive('code'),
        bulletList: instance.isActive('bulletList'),
        orderedList: instance.isActive('orderedList'),
        blockquote: instance.isActive('blockquote'),
        link: instance.isActive('link'),
        align: instance.isActive({ textAlign: 'center' })
          ? 'center'
          : instance.isActive({ textAlign: 'right' })
            ? 'right'
            : instance.isActive({ textAlign: 'justify' })
              ? 'justify'
              : 'left',
        block: instance.isActive('heading', { level: 1 })
          ? 'h1'
          : instance.isActive('heading', { level: 2 })
            ? 'h2'
            : instance.isActive('heading', { level: 3 })
              ? 'h3'
              : instance.isActive('heading', { level: 4 })
                ? 'h4'
                : 'p',
        fontFamily: (textStyle.fontFamily as string | undefined) ?? '',
        fontSize: (textStyle.fontSize as string | undefined) ?? '',
        color: toHex(textStyle.color, '#1e1c18'),
        highlight: instance.isActive('highlight'),
        highlightColor: toHex(instance.getAttributes('highlight').color, '#f2d675'),
        canUndo: instance.can().undo(),
        canRedo: instance.can().redo(),
      };
    },
  });

  const commitName = useCallback(() => {
    const trimmed = name.trim();
    if (trimmed && trimmed !== doc.name) {
      onRename(trimmed);
    } else {
      setName(doc.name);
    }
  }, [doc.name, name, onRename]);

  const handleDownload = async () => {
    if (!editor) return;
    try {
      setStatus('saving');
      const blob = await exportDocx(editor.getJSON(), baseName(name));
      saveAs(blob, withExtension(name, 'docx'));
      setStatus('saved');
      notify('Word document downloaded.', 'success');
    } catch {
      setStatus('error');
      notify('The document could not be exported. Please try again.', 'error');
    }
  };

  const handleDownloadPdf = async () => {
    if (!editor) return;
    try {
      setStatus('saving');
      const html = editor.getHTML();
      const bytes = await exportHtmlToPdf(html, baseName(name));
      saveAs(new Blob([new Uint8Array(bytes)], { type: 'application/pdf' }), withExtension(name, 'pdf'));
      setStatus('saved');
      notify('PDF document downloaded.', 'success');
    } catch {
      setStatus('error');
      notify('The PDF could not be exported. Please try again.', 'error');
    }
  };

  const insertImage = (file: File) => {
    if (!editor) return;
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result !== 'string') return;
      editor.chain().focus().setImage({ src: reader.result }).run();
    };
    reader.readAsDataURL(file);
  };

  const setLink = () => {
    if (!editor) return;
    const previous = editor.getAttributes('link').href as string | undefined;
    const url = window.prompt('Link URL', previous ?? 'https://');
    if (url === null) return;
    if (!url.trim()) {
      editor.chain().focus().extendMarkRange('link').unsetLink().run();
      return;
    }
    editor.chain().focus().extendMarkRange('link').setLink({ href: url.trim() }).run();
  };

  const fontSize = ui?.fontSize || '18px';
  const fontSizeValue = SIZE_OPTIONS.includes(fontSize) ? fontSize : '__current';
  const estimatedPages = Math.max(1, Math.ceil(stats.words / 430));

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
              if (event.key === 'Enter') {
                event.currentTarget.blur();
              }
            }}
          />
          <span className="chip chip-docx">
            {doc.kind === 'pdf' ? 'PDF' : doc.legacy ? 'DOC' : 'DOCX'}
          </span>
        </div>

        <div className="header-actions">
          <span className={`save-indicator ${status === 'saving' ? 'saving' : ''}`}>
            <span className="dot" style={status === 'error' ? { background: '#b03a25' } : undefined} />
            {status === 'saving' ? 'Saving…' : status === 'error' ? 'Not saved' : 'Saved locally'}
          </span>
          <input
            ref={openInputRef}
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
            onClick={() => openInputRef.current?.click()}
          >
            <FolderOpen size={14} />
            <span>Open</span>
          </button>
          <button type="button" className="icon-btn" title="Print" onClick={() => window.print()}>
            <Printer size={16} />
          </button>
          {onSwitchToPdf && (
            <button
              type="button"
              className="btn btn-sm"
              onClick={onSwitchToPdf}
              title="Switch to PDF page markup & visual view"
            >
              <Eye size={14} />
              <span>PDF Page View</span>
            </button>
          )}
          <button
            type="button"
            className={`btn btn-sm ${doc.kind === 'pdf' ? 'btn-primary' : ''}`}
            onClick={handleDownloadPdf}
            title="Download formatted PDF document"
          >
            <Download size={14} />
            <span>Download PDF</span>
          </button>
          <button
            type="button"
            className={`btn btn-sm ${doc.kind !== 'pdf' ? 'btn-primary' : ''}`}
            onClick={handleDownload}
            title="Download Word document"
          >
            <FileOutput size={14} />
            <span>Download .docx</span>
          </button>
        </div>
      </header>

      <div className="word-body">
        {showLegacy && (
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 9,
              padding: '8px 16px',
              background: '#f3ecdc',
              borderBottom: '1px solid var(--line)',
              color: 'var(--ink-soft)',
              fontSize: 12,
            }}
          >
            <Info size={14} style={{ flex: 'none', color: 'var(--accent)' }} />
            <span style={{ flex: 1 }}>
              This legacy <strong>.doc</strong> file was imported as editable text. Complex layout and
              embedded objects are not preserved. Download a .docx to keep your changes.
            </span>
            <button
              type="button"
              className="icon-btn"
              style={{ width: 24, height: 24 }}
              onClick={() => setShowLegacy(false)}
              aria-label="Dismiss notice"
            >
              <X size={13} />
            </button>
          </div>
        )}

        <div className="toolbar" role="toolbar" aria-label="Text formatting">
          <div className="tool-group">
            <ToolButton title="Undo (Ctrl+Z)" disabled={!ui?.canUndo} onClick={() => editor?.chain().focus().undo().run()}>
              <Undo2 size={16} />
            </ToolButton>
            <ToolButton
              title="Redo (Ctrl+Shift+Z)"
              disabled={!ui?.canRedo}
              onClick={() => editor?.chain().focus().redo().run()}
            >
              <Redo2 size={16} />
            </ToolButton>
          </div>
          <div className="tool-divider" />
          <div className="tool-group">
            <select
              className="style-select"
              aria-label="Paragraph style"
              value={ui?.block ?? 'p'}
              onChange={(event) => {
                const value = event.target.value;
                if (!editor) return;
                if (value === 'p') {
                  editor.chain().focus().setParagraph().run();
                } else {
                  editor
                    .chain()
                    .focus()
                    .toggleHeading({ level: Number(value.slice(1)) as 1 | 2 | 3 | 4 })
                    .run();
                }
              }}
            >
              <option value="p">Paragraph</option>
              <option value="h1">Heading 1</option>
              <option value="h2">Heading 2</option>
              <option value="h3">Heading 3</option>
              <option value="h4">Heading 4</option>
            </select>
            <select
              className="font-select"
              aria-label="Font family"
              value={ui?.fontFamily ?? ''}
              onChange={(event) => {
                if (!editor) return;
                const value = event.target.value;
                if (!value) {
                  editor.chain().focus().unsetFontFamily().run();
                } else {
                  editor.chain().focus().setFontFamily(value).run();
                }
              }}
            >
              {FONT_OPTIONS.map((option) => (
                <option value={option.value} key={option.value}>
                  {option.label}
                </option>
              ))}
              {ui?.fontFamily && !FONT_OPTIONS.some((option) => option.value === ui.fontFamily) && (
                <option value={ui.fontFamily}>{ui.fontFamily}</option>
              )}
            </select>
            <select
              className="size-select"
              aria-label="Font size"
              value={fontSizeValue}
              onChange={(event) => {
                if (!editor) return;
                const value = event.target.value;
                if (value === '__current' || value === '18px') {
                  editor.chain().focus().setMark('textStyle', { fontSize: null }).run();
                } else {
                  editor.chain().focus().setMark('textStyle', { fontSize: value }).run();
                }
              }}
            >
              <option value="18px">Auto</option>
              {fontSizeValue === '__current' && <option value="__current">{fontSize}</option>}
              {SIZE_OPTIONS.filter((size) => size !== '18px').map((size) => (
                <option value={size} key={size}>
                  {size.replace('px', '')}
                </option>
              ))}
            </select>
          </div>
          <div className="tool-divider" />
          <div className="tool-group">
            <ToolButton
              title="Bold (Ctrl+B)"
              active={ui?.bold}
              onClick={() => editor?.chain().focus().toggleBold().run()}
            >
              <Bold size={15} />
            </ToolButton>
            <ToolButton
              title="Italic (Ctrl+I)"
              active={ui?.italic}
              onClick={() => editor?.chain().focus().toggleItalic().run()}
            >
              <Italic size={15} />
            </ToolButton>
            <ToolButton
              title="Underline (Ctrl+U)"
              active={ui?.underline}
              onClick={() => editor?.chain().focus().toggleUnderline().run()}
            >
              <UnderlineIcon size={15} />
            </ToolButton>
            <ToolButton
              title="Strikethrough"
              active={ui?.strike}
              onClick={() => editor?.chain().focus().toggleStrike().run()}
            >
              <Strikethrough size={15} />
            </ToolButton>
            <ToolButton
              title="Inline code"
              active={ui?.code}
              onClick={() => editor?.chain().focus().toggleCode().run()}
            >
              <Code size={15} />
            </ToolButton>
            <label className="tool-btn color-tool" title="Text color">
              <span style={{ fontFamily: 'Georgia, serif', fontSize: 15, fontWeight: 600, paddingBottom: 4 }}>
                A
              </span>
              <span className="color-bar" style={{ background: ui?.color ?? '#1e1c18' }} />
              <input
                type="color"
                aria-label="Text color"
                value={ui?.color ?? '#1e1c18'}
                onChange={(event) => editor?.chain().focus().setColor(event.target.value).run()}
              />
            </label>
            <label className="tool-btn color-tool" title="Highlight color">
              <Highlighter size={15} style={{ paddingBottom: 4 }} />
              <span className="color-bar" style={{ background: ui?.highlightColor ?? '#f2d675' }} />
              <input
                type="color"
                aria-label="Highlight color"
                value={ui?.highlightColor ?? '#f2d675'}
                onChange={(event) =>
                  editor?.chain().focus().toggleHighlight({ color: event.target.value }).run()
                }
              />
            </label>
          </div>
          <div className="tool-divider" />
          <div className="tool-group">
            <ToolButton
              title="Align left"
              active={ui?.align === 'left'}
              onClick={() => editor?.chain().focus().setTextAlign('left').run()}
            >
              <AlignLeft size={15} />
            </ToolButton>
            <ToolButton
              title="Align center"
              active={ui?.align === 'center'}
              onClick={() => editor?.chain().focus().setTextAlign('center').run()}
            >
              <AlignCenter size={15} />
            </ToolButton>
            <ToolButton
              title="Align right"
              active={ui?.align === 'right'}
              onClick={() => editor?.chain().focus().setTextAlign('right').run()}
            >
              <AlignRight size={15} />
            </ToolButton>
            <ToolButton
              title="Justify"
              active={ui?.align === 'justify'}
              onClick={() => editor?.chain().focus().setTextAlign('justify').run()}
            >
              <AlignJustify size={15} />
            </ToolButton>
          </div>
          <div className="tool-divider" />
          <div className="tool-group">
            <ToolButton
              title="Bulleted list"
              active={ui?.bulletList}
              onClick={() => editor?.chain().focus().toggleBulletList().run()}
            >
              <List size={15} />
            </ToolButton>
            <ToolButton
              title="Numbered list"
              active={ui?.orderedList}
              onClick={() => editor?.chain().focus().toggleOrderedList().run()}
            >
              <ListOrdered size={15} />
            </ToolButton>
            <ToolButton
              title="Quote"
              active={ui?.blockquote}
              onClick={() => editor?.chain().focus().toggleBlockquote().run()}
            >
              <Quote size={15} />
            </ToolButton>
          </div>
          <div className="tool-divider" />
          <div className="tool-group">
            <ToolButton title="Link" active={ui?.link} onClick={setLink}>
              <Link2 size={15} />
            </ToolButton>
            <ToolButton title="Insert image" onClick={() => imageInputRef.current?.click()}>
              <ImagePlus size={15} />
            </ToolButton>
            <div className="menu-wrap" ref={tableMenuRef}>
              <ToolButton title="Table" active={tableMenu} onClick={() => setTableMenu((open) => !open)}>
                <Table2 size={15} />
              </ToolButton>
              {tableMenu && (
                <div className="popover">
                  <div className="popover-title">Table</div>
                  <button
                    type="button"
                    className="popover-item"
                    onClick={() => {
                      editor?.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
                      setTableMenu(false);
                    }}
                  >
                    Insert 3 × 3 table
                  </button>
                  <div className="popover-sep" />
                  <button
                    type="button"
                    className="popover-item"
                    disabled={!editor?.can().addRowAfter()}
                    onClick={() => editor?.chain().focus().addRowAfter().run()}
                  >
                    Add row below
                  </button>
                  <button
                    type="button"
                    className="popover-item"
                    disabled={!editor?.can().addColumnAfter()}
                    onClick={() => editor?.chain().focus().addColumnAfter().run()}
                  >
                    Add column right
                  </button>
                  <div className="popover-sep" />
                  <button
                    type="button"
                    className="popover-item"
                    disabled={!editor?.can().deleteRow()}
                    onClick={() => editor?.chain().focus().deleteRow().run()}
                  >
                    Delete row
                  </button>
                  <button
                    type="button"
                    className="popover-item"
                    disabled={!editor?.can().deleteColumn()}
                    onClick={() => editor?.chain().focus().deleteColumn().run()}
                  >
                    Delete column
                  </button>
                  <button
                    type="button"
                    className="popover-item"
                    disabled={!editor?.can().deleteTable()}
                    onClick={() => editor?.chain().focus().deleteTable().run()}
                  >
                    Delete table
                  </button>
                </div>
              )}
            </div>
            <ToolButton
              title="Divider"
              onClick={() => editor?.chain().focus().setHorizontalRule().run()}
            >
              <Minus size={15} />
            </ToolButton>
            <ToolButton
              title="Clear formatting"
              onClick={() => editor?.chain().focus().unsetAllMarks().clearNodes().run()}
            >
              <Eraser size={15} />
            </ToolButton>
          </div>
          <input
            ref={imageInputRef}
            type="file"
            accept="image/*"
            hidden
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) insertImage(file);
              event.target.value = '';
            }}
          />
        </div>

        <div className="editor-scroll">
          <div
            className={`paper ${doc.kind === 'pdf' || (doc.html && doc.html.includes('font-sans')) ? 'font-sans' : ''}`}
            style={{ zoom: zoom / 100 }}
          >
            <EditorContent editor={editor} />
          </div>
        </div>

        <div className="statusbar">
          <div className="statusbar-left">
            <span className="status-item">{stats.words} words</span>
            <span className="status-item hide-mobile">{stats.characters} characters</span>
            <span className="status-item hide-mobile">
              {estimatedPages} {estimatedPages === 1 ? 'page' : 'pages'} est.
            </span>
          </div>
          <div className="statusbar-right">
            <span className="status-item hide-mobile">Editing in your browser</span>
            <div className="zoom-control">
              <button
                type="button"
                title="Zoom out"
                onClick={() => setZoom((value) => clamp(value - 10, 50, 150))}
              >
                <Minus size={12} />
              </button>
              <button type="button" className="zoom-label" title="Reset zoom" onClick={() => setZoom(100)}>
                {zoom}%
              </button>
              <button
                type="button"
                title="Zoom in"
                onClick={() => setZoom((value) => clamp(value + 10, 50, 150))}
              >
                <Plus size={12} />
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
