import { useRef, useState } from 'react';
import {
  ArrowRight,
  Eraser,
  FileEdit,
  FileOutput,
  FilePlus2,
  FolderOpen,
  Highlighter,
  Lock,
  Moon,
  ShieldCheck,
  Sparkles,
  Sun,
  UploadCloud,
  X,
} from 'lucide-react';
import type { RecentDoc } from '../types';
import { ACCEPTED_EXTENSIONS } from '../lib/detect';
import { formatBytes, formatRelativeTime } from '../lib/format';
import { FileGlyph } from './ui';

interface HomeProps {
  recent: RecentDoc[];
  onOpenFile: (file: File) => void;
  onOpenRecent: (doc: RecentDoc) => void;
  onRemoveRecent: (doc: RecentDoc) => void;
  onBlank: () => void;
  onSample: () => void;
}

export function Home({ recent, onOpenFile, onOpenRecent, onRemoveRecent, onBlank, onSample }: HomeProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    if (typeof document !== 'undefined') {
      const current = document.documentElement.getAttribute('data-theme');
      if (current === 'light' || current === 'dark') return current;
      if (typeof window !== 'undefined' && window.matchMedia('(prefers-color-scheme: dark)').matches) {
        return 'dark';
      }
    }
    return 'light';
  });

  const toggleTheme = () => {
    const next = theme === 'dark' ? 'light' : 'dark';
    setTheme(next);
    document.documentElement.setAttribute('data-theme', next);
    try {
      localStorage.setItem('papyrus-theme', next);
    } catch {}
  };

  const chooseFile = () => inputRef.current?.click();

  return (
    <div className="home">
      <input
        ref={inputRef}
        type="file"
        accept={ACCEPTED_EXTENSIONS}
        hidden
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onOpenFile(file);
          event.target.value = '';
        }}
      />

      <div className="home-inner">
        <nav className="home-nav">
          <div className="brand">
            <span className="brand-mark">P</span>
            <span>
              <span className="brand-name">Papyrus</span>
              <span className="brand-sub">Document studio</span>
            </span>
          </div>
          <div className="home-nav-right">
            <button
              type="button"
              className="theme-toggle"
              onClick={toggleTheme}
              title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}
              aria-label="Toggle color theme"
            >
              {theme === 'dark' ? <Sun size={17} /> : <Moon size={17} />}
            </button>
            <span className="privacy-pill">
              <ShieldCheck size={13} />
              <span>Local-first · files never leave your browser</span>
            </span>
          </div>
        </nav>

        <section className="hero">
          <div>
            <div className="eyebrow">Word editing &amp; PDF editing, together</div>
            <h1 className="hero-title">
              Read deeply.
              <br />
              <em>Edit &amp; mark up.</em>
            </h1>
            <p className="hero-copy">
              A calm, capable desk for the documents you actually work with. Open Word or PDF files, edit text
              directly, whiteout, annotate, or convert PDFs into fully editable documents. Everything stays on this device.
            </p>
            <div className="hero-actions">
              <button type="button" className="btn btn-primary btn-lg" onClick={chooseFile}>
                <FolderOpen size={16} />
                Open a document
              </button>
              <button type="button" className="btn btn-lg" onClick={onBlank}>
                <FilePlus2 size={16} />
                Blank page
              </button>
            </div>
            <div className="hero-footnote">
              <span>
                <Lock size={12} />
                No uploads, no accounts
              </span>
              <span>
                <Sparkles size={12} />
                DOCX &amp; PDF editing, whiteout &amp; export
              </span>
            </div>
            <div className="hero-badges">
              <span className="feature-badge"><FileEdit size={12} /> Full DOCX editing</span>
              <span className="feature-badge"><Highlighter size={12} /> PDF annotation</span>
              <span className="feature-badge"><Eraser size={12} /> Whiteout tool</span>
              <span className="feature-badge"><FileOutput size={12} /> Export PDF &amp; DOCX</span>
            </div>
          </div>

          <div
            className={`drop-card ${dragOver ? 'dragging' : ''}`}
            onDragOver={(event) => {
              event.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(event) => {
              event.preventDefault();
              event.stopPropagation();
              setDragOver(false);
              const file = event.dataTransfer.files?.[0];
              if (file) onOpenFile(file);
            }}
          >
            <div className="drop-inner">
              <div className="drop-icon">
                <UploadCloud size={26} strokeWidth={1.7} />
              </div>
              <div className="drop-title">Drop a document here</div>
              <p className="drop-copy">or choose a file from your computer</p>
              <div className="drop-actions">
                <button type="button" className="btn btn-primary" onClick={chooseFile}>
                  Choose file
                </button>
                <button type="button" className="btn" onClick={onSample}>
                  Sample PDF
                </button>
              </div>
              <div className="format-row">
                <span className="format-tag docx">
                  <i />
                  DOCX · full editing
                </span>
                <span className="format-tag doc">
                  <i />
                  DOC · text import
                </span>
                <span className="format-tag pdf">
                  <i />
                  PDF · edit, whiteout &amp; markup
                </span>
              </div>
              <div className="drop-note">
                <Lock size={11} />
                Processed in your browser. Original files are untouched.
              </div>
            </div>
          </div>
        </section>

        <section>
          <div className="section-head">
            <div>
              <h2 className="section-title">Your library</h2>
              <p className="section-sub">
                {recent.length > 0
                  ? 'Pick up where you left off. Edits are saved to this browser.'
                  : 'Recently opened documents will appear here.'}
              </p>
            </div>
            {recent.length > 0 && (
              <span className="chip">{recent.length} saved</span>
            )}
          </div>

          {recent.length === 0 ? (
            <div className="empty-recent">
              Nothing here yet. Open a file above to start your library.
            </div>
          ) : (
            <div className="recent-grid">
              {recent.map((doc) => (
                <div className="recent-card" key={doc.id}>
                  <button
                    type="button"
                    className="recent-open"
                    onClick={() => onOpenRecent(doc)}
                    title={`Open ${doc.name}`}
                  >
                    <FileGlyph kind={doc.kind} />
                    <span className="recent-info">
                      <span className="recent-name">{doc.name}</span>
                      <span className="recent-meta">
                        {formatRelativeTime(doc.updatedAt)}
                        <span className="dot" />
                        {formatBytes(doc.size)}
                      </span>
                    </span>
                    <ArrowRight size={14} style={{ marginLeft: 'auto', color: 'var(--muted)', flex: 'none', marginTop: 4 }} />
                  </button>
                  <button
                    type="button"
                    className="icon-btn recent-remove"
                    title="Remove from library"
                    aria-label={`Remove ${doc.name} from library`}
                    onClick={() => onRemoveRecent(doc)}
                  >
                    <X size={13} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </section>

        <footer className="home-footer">
          <p>
            <span className="brand-mark" style={{ width: 22, height: 22, fontSize: 13, borderRadius: 6 }}>
              P
            </span>
            Papyrus · a quiet place for documents
          </p>
          <p>
            <Lock size={11} />
            Your files are stored in this browser only
          </p>
        </footer>
      </div>
    </div>
  );
}
