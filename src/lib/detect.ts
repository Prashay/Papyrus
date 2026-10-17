import type { DocKind } from '../types';
import { extensionOf } from './format';

export function detectKind(name: string, bytes: ArrayBuffer): DocKind | null {
  const header = new Uint8Array(bytes.slice(0, 8));
  const isPdf =
    header[0] === 0x25 && header[1] === 0x50 && header[2] === 0x44 && header[3] === 0x46;
  if (isPdf) return 'pdf';

  const isZip = header[0] === 0x50 && header[1] === 0x4b && header[2] === 0x03 && header[3] === 0x04;
  if (isZip) return 'docx';

  const isOle =
    header[0] === 0xd0 &&
    header[1] === 0xcf &&
    header[2] === 0x11 &&
    header[3] === 0xe0 &&
    header[4] === 0xa1 &&
    header[5] === 0xb1 &&
    header[6] === 0x1a &&
    header[7] === 0xe1;
  if (isOle) return 'doc';

  const extension = extensionOf(name);
  if (extension === 'pdf') return 'pdf';
  if (extension === 'docx') return 'docx';
  if (extension === 'doc') return 'doc';

  return null;
}

export const ACCEPTED_EXTENSIONS = '.doc,.docx,.pdf';
