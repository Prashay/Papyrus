import * as CFB from 'cfb';
import type { CFB$Blob } from 'cfb';
import { plainTextToHtml } from './docx';

function asBytes(content: CFB$Blob | undefined): Uint8Array | null {
  if (!content) return null;
  if (content instanceof Uint8Array) return content;
  return Uint8Array.from(content as number[]);
}

function findStream(cfb: CFB.CFB$Container, name: string): Uint8Array | null {
  const entry =
    CFB.find(cfb, name) ??
    CFB.find(cfb, `/${name}`) ??
    cfb.FileIndex.find((item) => item.name === name && item.type === 2) ??
    null;
  if (!entry) return null;
  return asBytes(entry.content);
}

function cleanWordText(raw: string): string {
  let output = '';
  let fieldDepth = 0;
  let inInstruction = 0;

  for (const char of raw) {
    const code = char.charCodeAt(0);

    if (code === 0x13) {
      fieldDepth += 1;
      inInstruction += 1;
      continue;
    }
    if (code === 0x14) {
      if (inInstruction > 0) inInstruction -= 1;
      continue;
    }
    if (code === 0x15) {
      if (fieldDepth > 0) fieldDepth -= 1;
      if (inInstruction > 0 && fieldDepth < inInstruction) inInstruction = fieldDepth;
      continue;
    }
    if (inInstruction > 0) continue;

    if (code === 0x0d || code === 0x0b) {
      output += '\n';
    } else if (code === 0x0c) {
      output += '\n\n';
    } else if (code === 0x07) {
      output += '\t';
    } else if (code === 0x09) {
      output += '\t';
    } else if (code === 0x1e) {
      output += '-';
    } else if (code === 0x1f) {
      // Optional hyphen: omit.
    } else if (code === 0x00) {
      // Padding: omit.
    } else if (code >= 0x20) {
      output += char;
    }
  }

  return output
    .replace(/\t+\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\t{2,}/g, '\t')
    .trim();
}

function decodePiece(
  word: Uint8Array,
  offset: number,
  characters: number,
  compressed: boolean,
  utf16: TextDecoder,
  cp1252: TextDecoder,
): string {
  if (offset < 0 || characters <= 0) return '';
  const length = compressed ? characters : characters * 2;
  if (offset + length > word.length) {
    const available = Math.max(0, word.length - offset);
    if (available <= 0) return '';
    return compressed
      ? cp1252.decode(word.subarray(offset, offset + available))
      : utf16.decode(word.subarray(offset, offset + Math.min(available, characters * 2)));
  }
  return compressed
    ? cp1252.decode(word.subarray(offset, offset + length))
    : utf16.decode(word.subarray(offset, offset + length));
}

function pieceTableText(word: Uint8Array, table: Uint8Array, view: DataView): string {
  const fcClx = view.getUint32(0x01a2, true);
  const lcbClx = view.getUint32(0x01a6, true);
  if (!fcClx || !lcbClx || fcClx + lcbClx > table.length) return '';

  const utf16 = new TextDecoder('utf-16le');
  const cp1252 = new TextDecoder('windows-1252');

  let position = fcClx;
  const end = fcClx + lcbClx;
  while (position < end && table[position] === 0x01) {
    if (position + 3 > end) return '';
    const size = table[position + 1] | (table[position + 2] << 8);
    position += 3 + size;
  }

  if (position >= end || table[position] !== 0x02) return '';
  if (position + 5 > end) return '';
  const plcLength =
    table[position + 1] |
    (table[position + 2] << 8) |
    (table[position + 3] << 16) |
    (table[position + 4] << 24);
  const plcStart = position + 5;
  if (plcStart + plcLength > end || plcLength < 12) return '';

  const pieceCount = Math.floor((plcLength - 4) / 12);
  if (pieceCount <= 0) return '';

  const plc = table.subarray(plcStart, plcStart + plcLength);
  const plcView = new DataView(plc.buffer, plc.byteOffset, plc.byteLength);
  const pieces: { cpStart: number; cpEnd: number; offset: number; compressed: boolean }[] = [];

  for (let index = 0; index < pieceCount; index += 1) {
    const cpStart = plcView.getUint32(index * 4, true);
    const cpEnd = plcView.getUint32((index + 1) * 4, true);
    const pcdOffset = 4 * (pieceCount + 1) + index * 8;
    const fcRaw = plcView.getUint32(pcdOffset + 2, true);
    const compressed = (fcRaw & 0x40000000) !== 0;
    const fileOffset = compressed ? (fcRaw & 0x3fffffff) / 2 : fcRaw & 0x3fffffff;
    if (cpEnd <= cpStart) continue;
    pieces.push({ cpStart, cpEnd, offset: fileOffset, compressed });
  }

  let output = '';
  for (const piece of pieces) {
    output += decodePiece(word, piece.offset, piece.cpEnd - piece.cpStart, piece.compressed, utf16, cp1252);
  }

  return output;
}

function fallbackText(word: Uint8Array, view: DataView): string {
  if (word.length < 0x20) return '';
  const fcMin = view.getUint32(0x0018, true);
  const fcMac = view.getUint32(0x001c, true);
  if (fcMin >= fcMac || fcMac > word.length) return '';
  const utf16 = new TextDecoder('utf-16le');
  const cp1252 = new TextDecoder('windows-1252');
  const utfText = utf16.decode(word.subarray(fcMin, fcMac));
  const cleaned = cleanWordText(utfText);
  if (cleaned.length > 2) return cleaned;
  return cleanWordText(cp1252.decode(word.subarray(fcMin, fcMac)));
}

export interface LegacyDocResult {
  text: string;
  html: string;
}

export function extractLegacyDoc(bytes: ArrayBuffer): LegacyDocResult {
  const cfb = CFB.read(new Uint8Array(bytes), { type: 'array' });
  const word = findStream(cfb, 'WordDocument');
  if (!word) {
    throw new Error('This does not look like a Word 97-2003 document.');
  }

  const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
  const flags = word.length > 0x0c ? view.getUint16(0x000a, true) : 0;
  const primaryTable = (flags & 0x0200) !== 0 ? '1Table' : '0Table';
  const secondaryTable = primaryTable === '1Table' ? '0Table' : '1Table';
  const table = findStream(cfb, primaryTable) ?? findStream(cfb, secondaryTable);

  let text = '';
  if (table) {
    try {
      text = cleanWordText(pieceTableText(word, table, view));
    } catch {
      text = '';
    }
  }

  if (!text) {
    text = fallbackText(word, view);
  }

  text = text.replace(/\u0000/g, '').trim();
  if (!text) {
    throw new Error('No readable text was found in this legacy .doc file.');
  }

  return { text, html: plainTextToHtml(text) };
}
