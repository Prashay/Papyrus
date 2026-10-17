import { AlertTriangle, Check, FileText, Loader2 } from 'lucide-react';
import type { DocKind } from '../types';

export type ToastTone = 'default' | 'success' | 'error';

export interface ToastItem {
  id: string;
  message: string;
  tone: ToastTone;
}

export function ToastStack({ toasts }: { toasts: ToastItem[] }) {
  return (
    <div className="toast-wrap" aria-live="polite">
      {toasts.map((toast) => (
        <div className={`toast ${toast.tone}`} key={toast.id}>
          {toast.tone === 'success' ? (
            <Check size={15} />
          ) : toast.tone === 'error' ? (
            <AlertTriangle size={15} />
          ) : null}
          <span>{toast.message}</span>
        </div>
      ))}
    </div>
  );
}

export function ProcessingOverlay({ label }: { label: string }) {
  return (
    <div className="processing" role="status" aria-live="polite">
      <div className="processing-card">
        <div className="spinner-mark">
          <Loader2 size={20} className="spin" />
        </div>
        <span>{label}</span>
      </div>
    </div>
  );
}

export function FileGlyph({ kind }: { kind: DocKind }) {
  return (
    <span className={`file-glyph ${kind}`}>
      <FileText size={20} strokeWidth={1.6} />
    </span>
  );
}
