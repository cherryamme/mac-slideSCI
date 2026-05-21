import type { SavedCombo } from "./types";

const DB_NAME = "mac-slideSCI";
const STORE_NAME = "combos";
const DB_VERSION = 1;

function openComboDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("当前环境不支持 IndexedDB，无法保存组合。"));
      return;
    }

    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        const store = db.createObjectStore(STORE_NAME, { keyPath: "id" });
        store.createIndex("updatedAt", "updatedAt");
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("打开组合数据库失败。"));
  });
}

async function withStore<T>(mode: IDBTransactionMode, fn: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await openComboDatabase();
  return new Promise<T>((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, mode);
    const store = transaction.objectStore(STORE_NAME);
    const request = fn(store);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("组合数据库操作失败。"));
    transaction.oncomplete = () => db.close();
    transaction.onerror = () => {
      reject(transaction.error ?? new Error("组合数据库事务失败。"));
      db.close();
    };
  });
}

export async function listCombos(): Promise<SavedCombo[]> {
  const items = (await withStore<SavedCombo[]>("readonly", (store) => store.getAll() as IDBRequest<SavedCombo[]>)) ?? [];
  return [...items].sort((a, b) => (b.updatedAt ?? b.createdAt ?? 0) - (a.updatedAt ?? a.createdAt ?? 0));
}

export async function saveCombo(combo: SavedCombo): Promise<void> {
  await withStore("readwrite", (store) => store.put(combo));
}

export async function deleteCombo(id: string): Promise<void> {
  await withStore("readwrite", (store) => store.delete(id));
}

export async function getCombo(id: string): Promise<SavedCombo | null> {
  const item = (await withStore<SavedCombo | undefined>("readonly", (store) => store.get(id) as IDBRequest<SavedCombo | undefined>)) ?? null;
  return item ?? null;
}

export async function renameCombo(id: string, name: string): Promise<SavedCombo | null> {
  const combo = await getCombo(id);
  if (!combo) {
    return null;
  }

  const next: SavedCombo = { ...combo, name: name.trim() || combo.name, updatedAt: Date.now() };
  await saveCombo(next);
  return next;
}
