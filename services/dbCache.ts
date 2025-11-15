
const DB_NAME = 'ExcelProcessorDB';
const STORE_NAME = 'databaseFiles';
const DB_VERSION = 1;

interface CachedFile {
  id: 'singleton';
  name: string;
  data: ArrayBuffer;
}

const getDb = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
  });
};

export const saveDbFile = async (file: File): Promise<void> => {
  const db = await getDb();
  // Read file into buffer BEFORE starting the transaction
  const fileBuffer = await file.arrayBuffer();
  
  const transaction = db.transaction(STORE_NAME, 'readwrite');
  const store = transaction.objectStore(STORE_NAME);
  
  const cachedFile: CachedFile = {
    id: 'singleton', // Use a fixed key to always overwrite the same file
    name: file.name,
    data: fileBuffer,
  };
  
  return new Promise((resolve, reject) => {
    const request = store.put(cachedFile);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

export const getDbFile = async (): Promise<File | null> => {
  const db = await getDb();
  const transaction = db.transaction(STORE_NAME, 'readonly');
  const store = transaction.objectStore(STORE_NAME);
  
  return new Promise((resolve, reject) => {
    const request = store.get('singleton');
    request.onsuccess = () => {
      const result: CachedFile | undefined = request.result;
      if (result) {
        const file = new File([result.data], result.name, {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        resolve(file);
      } else {
        resolve(null);
      }
    };
    request.onerror = () => reject(request.error);
  });
};

export const clearDbFile = async (): Promise<void> => {
    const db = await getDb();
    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);

    return new Promise((resolve, reject) => {
        const request = store.delete('singleton');
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
    });
};
