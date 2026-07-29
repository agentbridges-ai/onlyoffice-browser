export type PersistedDocumentTab = {
  id: string;
  name: string;
  handle: FileSystemFileHandle;
  lastOpenedAt: number;
};

const DATABASE_NAME = 'onlyoffice-browser-pwa';
const DATABASE_VERSION = 1;
const TAB_STORE = 'document-tabs';
type FilePermissionDescriptor = { mode: 'read' | 'readwrite' };

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.addEventListener('success', () => resolve(request.result), { once: true });
    request.addEventListener('error', () => reject(request.error || new Error('IndexedDB request failed.')), {
      once: true,
    });
  });
}

function transactionDone(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.addEventListener('complete', () => resolve(), { once: true });
    transaction.addEventListener(
      'abort',
      () => reject(transaction.error || new Error('IndexedDB transaction aborted.')),
      {
        once: true,
      },
    );
    transaction.addEventListener(
      'error',
      () => reject(transaction.error || new Error('IndexedDB transaction failed.')),
      {
        once: true,
      },
    );
  });
}

export class DocumentTabStore {
  private databasePromise: Promise<IDBDatabase> | null = null;

  async list(): Promise<PersistedDocumentTab[]> {
    const database = await this.open();
    const transaction = database.transaction(TAB_STORE, 'readonly');
    const records = await requestResult(
      transaction.objectStore(TAB_STORE).getAll() as IDBRequest<PersistedDocumentTab[]>,
    );
    await transactionDone(transaction);
    return records.sort((left, right) => left.lastOpenedAt - right.lastOpenedAt);
  }

  async put(tab: PersistedDocumentTab): Promise<void> {
    const database = await this.open();
    const transaction = database.transaction(TAB_STORE, 'readwrite');
    transaction.objectStore(TAB_STORE).put(tab);
    await transactionDone(transaction);
  }

  async remove(id: string): Promise<void> {
    const database = await this.open();
    const transaction = database.transaction(TAB_STORE, 'readwrite');
    transaction.objectStore(TAB_STORE).delete(id);
    await transactionDone(transaction);
  }

  private open(): Promise<IDBDatabase> {
    this.databasePromise ||= new Promise((resolve, reject) => {
      const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
      request.addEventListener(
        'upgradeneeded',
        () => {
          if (!request.result.objectStoreNames.contains(TAB_STORE)) {
            request.result.createObjectStore(TAB_STORE, { keyPath: 'id' });
          }
        },
        { once: true },
      );
      request.addEventListener('success', () => resolve(request.result), { once: true });
      request.addEventListener('error', () => reject(request.error || new Error('Unable to open document tabs.')), {
        once: true,
      });
    });
    return this.databasePromise;
  }
}

export function createDocumentTabId(): string {
  return globalThis.crypto?.randomUUID?.() || `document-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

export function documentTypeForName(name: string): 'word' | 'cell' | 'slide' {
  const extension = name.split('.').pop()?.toLocaleLowerCase();
  if (extension === 'xlsx' || extension === 'xls' || extension === 'csv') return 'cell';
  if (extension === 'pptx' || extension === 'ppt') return 'slide';
  return 'word';
}

export async function hasReadPermission(handle: FileSystemFileHandle): Promise<boolean> {
  const permissionHandle = handle as FileSystemFileHandle & {
    queryPermission?: (options: FilePermissionDescriptor) => Promise<PermissionState>;
  };
  if (!permissionHandle.queryPermission) return true;
  return (await permissionHandle.queryPermission({ mode: 'read' })) === 'granted';
}

export async function requestReadWritePermission(handle: FileSystemFileHandle): Promise<boolean> {
  const permissionHandle = handle as FileSystemFileHandle & {
    queryPermission?: (options: FilePermissionDescriptor) => Promise<PermissionState>;
    requestPermission?: (options: FilePermissionDescriptor) => Promise<PermissionState>;
  };
  if (!permissionHandle.queryPermission || !permissionHandle.requestPermission) return true;
  if ((await permissionHandle.queryPermission({ mode: 'readwrite' })) === 'granted') return true;
  return (await permissionHandle.requestPermission({ mode: 'readwrite' })) === 'granted';
}
