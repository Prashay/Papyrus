import type { PdfAnnotation, RecentDoc, StoredPayload, StoredState } from '../types';

const DB_NAME = 'papyrus-studio';
const DB_VERSION = 1;
const META_STORE = 'meta';
const FILE_STORE = 'files';

interface FileRecord extends StoredState {
  id: string;
  bytes: ArrayBuffer;
}

let dbPromise: Promise<IDBDatabase | null> | null = null;

function openDatabase(): Promise<IDBDatabase | null> {
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve) => {
    if (typeof indexedDB === 'undefined') {
      resolve(null);
      return;
    }

    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(META_STORE)) {
        db.createObjectStore(META_STORE, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(FILE_STORE)) {
        db.createObjectStore(FILE_STORE, { keyPath: 'id' });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
    request.onblocked = () => resolve(null);
  });

  return dbPromise;
}

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

export async function listRecentDocuments(): Promise<RecentDoc[]> {
  const db = await openDatabase();
  if (!db) return [];
  try {
    const tx = db.transaction(META_STORE, 'readonly');
    const records = await requestToPromise(tx.objectStore(META_STORE).getAll());
    return (records as RecentDoc[]).sort((a, b) => b.updatedAt - a.updatedAt);
  } catch {
    return [];
  }
}

export async function saveDocument(meta: RecentDoc, payload: StoredPayload): Promise<void> {
  const db = await openDatabase();
  if (!db) return;
  const record: FileRecord = { id: meta.id, ...payload };
  const tx = db.transaction([META_STORE, FILE_STORE], 'readwrite');
  tx.objectStore(META_STORE).put(meta);
  tx.objectStore(FILE_STORE).put(record);
  await transactionDone(tx);
}

export async function loadDocument(id: string): Promise<(StoredPayload & { meta: RecentDoc }) | null> {
  const db = await openDatabase();
  if (!db) return null;
  try {
    const tx = db.transaction([META_STORE, FILE_STORE], 'readonly');
    const meta = (await requestToPromise(tx.objectStore(META_STORE).get(id))) as RecentDoc | undefined;
    const file = (await requestToPromise(tx.objectStore(FILE_STORE).get(id))) as FileRecord | undefined;
    if (!meta || !file) return null;
    return {
      meta,
      bytes: file.bytes,
      html: file.html,
      annotations: file.annotations,
      page: file.page,
      zoom: file.zoom,
    };
  } catch {
    return null;
  }
}

export async function updateDocumentState(
  id: string,
  patch: StoredState,
  options?: { touch?: boolean; name?: string; size?: number },
): Promise<void> {
  const db = await openDatabase();
  if (!db) return;
  const tx = db.transaction([META_STORE, FILE_STORE], 'readwrite');
  const metaStore = tx.objectStore(META_STORE);
  const fileStore = tx.objectStore(FILE_STORE);
  const meta = (await requestToPromise(metaStore.get(id))) as RecentDoc | undefined;
  const file = (await requestToPromise(fileStore.get(id))) as FileRecord | undefined;

  if (file) {
    fileStore.put({ ...file, ...patch, id } satisfies FileRecord);
  }

  if (meta) {
    const updated: RecentDoc = {
      ...meta,
      name: options?.name ?? meta.name,
      size: options?.size ?? meta.size,
      updatedAt: options?.touch === false ? meta.updatedAt : Date.now(),
    };
    metaStore.put(updated);
  }

  await transactionDone(tx);
}

export async function renameDocument(id: string, name: string): Promise<void> {
  const db = await openDatabase();
  if (!db) return;
  const tx = db.transaction(META_STORE, 'readwrite');
  const store = tx.objectStore(META_STORE);
  const meta = (await requestToPromise(store.get(id))) as RecentDoc | undefined;
  if (meta) {
    store.put({ ...meta, name });
  }
  await transactionDone(tx);
}

export async function removeDocument(id: string): Promise<void> {
  const db = await openDatabase();
  if (!db) return;
  const tx = db.transaction([META_STORE, FILE_STORE], 'readwrite');
  tx.objectStore(META_STORE).delete(id);
  tx.objectStore(FILE_STORE).delete(id);
  await transactionDone(tx);
}

export async function saveAnnotations(
  id: string,
  annotations: PdfAnnotation[],
  page: number,
  zoom?: number,
): Promise<void> {
  await updateDocumentState(id, { annotations, page, zoom });
}
