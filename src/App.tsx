import { useCallback, useEffect, useRef, useState, type DragEvent } from 'react';
import { UploadCloud } from 'lucide-react';
import { Home } from './components/Home';
import { WordEditor } from './components/WordEditor';
import { PdfEditor } from './components/PdfEditor';
import { ProcessingOverlay, ToastStack, type ToastItem, type ToastTone } from './components/ui';
import { detectKind } from './lib/detect';
import { importDocx } from './lib/docx';
import { extractLegacyDoc } from './lib/legacyDoc';
import { buildSamplePdf } from './lib/pdf';
import {
  listRecentDocuments,
  loadDocument,
  removeDocument,
  renameDocument,
  saveDocument,
  updateDocumentState,
} from './lib/storage';
import { uid } from './lib/format';
import type { OpenDocument, RecentDoc, StoredState } from './types';

type Screen = 'home' | 'word' | 'pdf';

export default function App() {
  const [screen, setScreen] = useState<Screen>('home');
  const [active, setActive] = useState<OpenDocument | null>(null);
  const [recent, setRecent] = useState<RecentDoc[]>([]);
  const [busy, setBusy] = useState<string | null>(null);
  const [toasts, setToasts] = useState<ToastItem[]>([]);
  const [dragActive, setDragActive] = useState(false);
  const dragDepth = useRef(0);

  const notify = useCallback((message: string, tone: ToastTone = 'default') => {
    const id = uid();
    setToasts((current) => [...current.slice(-2), { id, message, tone }]);
    window.setTimeout(() => {
      setToasts((current) => current.filter((toast) => toast.id !== id));
    }, 4800);
  }, []);

  const refreshRecent = useCallback(() => {
    void listRecentDocuments().then(setRecent);
  }, []);

  useEffect(() => {
    refreshRecent();
  }, [refreshRecent]);

  const openDocument = useCallback(
    async (file: File) => {
      if (busy) return;
      setBusy('Reading document…');
      try {
        const bytes = await file.arrayBuffer();
        const kind = detectKind(file.name, bytes);
        if (!kind) {
          notify('Unsupported file. Open a DOCX, DOC or PDF document.', 'error');
          return;
        }

        let html: string | undefined;
        let legacy = false;
        let warnings: string[] = [];

        if (kind === 'docx') {
          setBusy('Opening Word document…');
          const result = await importDocx(bytes);
          html = result.html;
          warnings = result.warnings;
        } else if (kind === 'doc') {
          setBusy('Reading legacy .doc file…');
          const result = extractLegacyDoc(bytes);
          html = result.html;
          legacy = true;
        }

        const meta: RecentDoc = {
          id: uid(),
          name: file.name || 'Untitled document',
          kind,
          size: file.size,
          updatedAt: Date.now(),
        };

        await saveDocument(meta, { bytes, html });
        const opened: OpenDocument = { ...meta, bytes, html, legacy };
        setActive(opened);
        setScreen(kind === 'pdf' ? 'pdf' : 'word');
        refreshRecent();

        if (legacy) {
          notify('Legacy .doc imported as editable text. Download a .docx to keep your changes.');
        } else if (warnings.length > 0) {
          notify(`Document opened with ${warnings.length} formatting note${warnings.length === 1 ? '' : 's'}.`);
        }
      } catch (error) {
        const message =
          error instanceof Error && /password|encrypted/i.test(error.message)
            ? 'This document appears to be password protected.'
            : 'The document could not be opened. It may be damaged or use an unsupported format.';
        notify(message, 'error');
      } finally {
        setBusy(null);
      }
    },
    [busy, notify, refreshRecent],
  );

  const openRecent = useCallback(
    async (doc: RecentDoc) => {
      if (busy) return;
      setBusy('Opening document…');
      try {
        const stored = await loadDocument(doc.id);
        if (!stored) {
          notify('This document is no longer available in your library.', 'error');
          refreshRecent();
          return;
        }

        let html = stored.html;
        let legacy = stored.meta.kind === 'doc';
        if (!html && stored.meta.kind === 'docx') {
          const result = await importDocx(stored.bytes);
          html = result.html;
        } else if (!html && stored.meta.kind === 'doc') {
          const result = extractLegacyDoc(stored.bytes);
          html = result.html;
          legacy = true;
        }

        if (stored.meta.kind === 'pdf' && stored.bytes.byteLength === 0) {
          notify('The original PDF data is missing. Please open the file again.', 'error');
          return;
        }

        const opened: OpenDocument = {
          ...stored.meta,
          bytes: stored.bytes,
          html,
          legacy,
          annotations: stored.annotations,
          page: stored.page,
          zoom: stored.zoom,
        };
        setActive(opened);
        setScreen(stored.meta.kind === 'pdf' ? 'pdf' : 'word');
      } catch {
        notify('This document could not be reopened.', 'error');
      } finally {
        setBusy(null);
      }
    },
    [busy, notify, refreshRecent],
  );

  const createBlank = useCallback(async () => {
    if (busy) return;
    const untitledCount = recent.filter((doc) => /^untitled document/i.test(doc.name)).length;
    const name = untitledCount === 0 ? 'Untitled document' : `Untitled document ${untitledCount + 1}`;
    const meta: RecentDoc = {
      id: uid(),
      name,
      kind: 'docx',
      size: 0,
      updatedAt: Date.now(),
    };
    const bytes = new ArrayBuffer(0);
    await saveDocument(meta, { bytes, html: '' });
    setActive({ ...meta, bytes, html: '' });
    setScreen('word');
    refreshRecent();
  }, [busy, recent, refreshRecent]);

  const openSample = useCallback(async () => {
    if (busy) return;
    setBusy('Preparing sample…');
    try {
      const bytes = await buildSamplePdf();
      const meta: RecentDoc = {
        id: uid(),
        name: 'Field notes — Papyrus.pdf',
        kind: 'pdf',
        size: bytes.byteLength,
        updatedAt: Date.now(),
      };
      await saveDocument(meta, { bytes });
      setActive({ ...meta, bytes });
      setScreen('pdf');
      refreshRecent();
    } catch {
      notify('The sample document could not be prepared.', 'error');
    } finally {
      setBusy(null);
    }
  }, [busy, notify, refreshRecent]);

  const persist = useCallback(
    async (patch: StoredState) => {
      if (!active) return;
      await updateDocumentState(active.id, patch, { name: active.name, size: active.size });
      setRecent((current) =>
        current.map((doc) =>
          doc.id === active.id ? { ...doc, name: active.name, updatedAt: Date.now() } : doc,
        ),
      );
    },
    [active],
  );

  const renameActive = useCallback(
    async (name: string) => {
      if (!active) return;
      setActive((current) => (current ? { ...current, name } : current));
      setRecent((current) => current.map((doc) => (doc.id === active.id ? { ...doc, name } : doc)));
      try {
        await renameDocument(active.id, name);
      } catch {
        notify('The new name could not be saved.', 'error');
      }
    },
    [active, notify],
  );

  const removeRecent = useCallback(
    async (doc: RecentDoc) => {
      try {
        await removeDocument(doc.id);
        setRecent((current) => current.filter((item) => item.id !== doc.id));
      } catch {
        notify('The document could not be removed.', 'error');
      }
    },
    [notify],
  );

  const handleEditPdfAsDocument = useCallback(
    async (html: string) => {
      if (!active) return;
      const updated: OpenDocument = { ...active, html };
      setActive(updated);
      await updateDocumentState(active.id, { html }, { name: active.name, size: active.size });
      setScreen('word');
      notify('PDF opened in Document Editor. You can edit all text and export as PDF or DOCX.', 'success');
    },
    [active, notify],
  );

  const goHome = useCallback(() => {
    setScreen('home');
    setActive(null);
    refreshRecent();
    window.setTimeout(refreshRecent, 450);
  }, [refreshRecent]);

  const isFileDrag = (event: DragEvent) =>
    Array.from(event.dataTransfer?.types ?? []).includes('Files');

  const handleDragEnter = (event: DragEvent) => {
    if (!isFileDrag(event)) return;
    event.preventDefault();
    dragDepth.current += 1;
    setDragActive(true);
  };

  const handleDragOver = (event: DragEvent) => {
    if (!isFileDrag(event)) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
  };

  const handleDragLeave = (event: DragEvent) => {
    if (!isFileDrag(event)) return;
    dragDepth.current = Math.max(0, dragDepth.current - 1);
    if (dragDepth.current === 0) setDragActive(false);
  };

  const handleDrop = (event: DragEvent) => {
    if (!isFileDrag(event)) return;
    event.preventDefault();
    dragDepth.current = 0;
    setDragActive(false);
    const file = event.dataTransfer.files?.[0];
    if (file) void openDocument(file);
  };

  return (
    <div
      style={{ height: '100%' }}
      onDragEnter={handleDragEnter}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {screen === 'home' || !active ? (
        <Home
          recent={recent}
          onOpenFile={(file) => void openDocument(file)}
          onOpenRecent={(doc) => void openRecent(doc)}
          onRemoveRecent={(doc) => void removeRecent(doc)}
          onBlank={() => void createBlank()}
          onSample={() => void openSample()}
        />
      ) : screen === 'word' ? (
        <WordEditor
          key={active.id}
          doc={active}
          notify={notify}
          onBack={goHome}
          onRename={(name) => void renameActive(name)}
          onPersist={persist}
          onOpenFile={(file) => void openDocument(file)}
          onSwitchToPdf={active.kind === 'pdf' ? () => setScreen('pdf') : undefined}
        />
      ) : (
        <PdfEditor
          key={active.id}
          doc={active}
          notify={notify}
          onBack={goHome}
          onRename={(name) => void renameActive(name)}
          onPersist={persist}
          onOpenFile={(file) => void openDocument(file)}
          onEditAsDocument={(html) => void handleEditPdfAsDocument(html)}
        />
      )}

      {busy && <ProcessingOverlay label={busy} />}
      <ToastStack toasts={toasts} />
      {dragActive && (
        <div className="drop-overlay">
          <div className="drop-overlay-card">
            <UploadCloud size={34} />
            Drop to open in Papyrus
            <span>DOCX · DOC · PDF</span>
          </div>
        </div>
      )}
    </div>
  );
}
