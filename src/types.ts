export type DocKind = 'docx' | 'doc' | 'pdf';

export interface RecentDoc {
  id: string;
  name: string;
  kind: DocKind;
  size: number;
  updatedAt: number;
}

export interface Point {
  x: number;
  y: number;
}

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface AnnotationBase {
  id: string;
  page: number;
  color: string;
}

export interface InkAnnotation extends AnnotationBase {
  type: 'ink';
  points: Point[];
  width: number;
}

export interface HighlightAnnotation extends AnnotationBase, Rect {
  type: 'highlight';
  opacity: number;
}

export interface TextAnnotation extends AnnotationBase {
  type: 'text';
  x: number;
  y: number;
  text: string;
  size: number;
  serif?: boolean;
  background?: string;
  bold?: boolean;
  italic?: boolean;
}

export interface WhiteoutAnnotation extends AnnotationBase, Rect {
  type: 'whiteout';
}

export type PdfAnnotation = InkAnnotation | HighlightAnnotation | TextAnnotation | WhiteoutAnnotation;

export type PdfTool = 'select' | 'edit-text' | 'text' | 'whiteout' | 'highlight' | 'ink' | 'eraser';

export interface StoredState {
  html?: string;
  annotations?: PdfAnnotation[];
  page?: number;
  zoom?: number;
}

export interface StoredPayload extends StoredState {
  bytes: ArrayBuffer;
}

export interface OpenDocument extends RecentDoc {
  bytes: ArrayBuffer;
  html?: string;
  annotations?: PdfAnnotation[];
  page?: number;
  zoom?: number;
  legacy?: boolean;
}
