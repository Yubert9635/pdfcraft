/**
 * Tauri Bridge
 * Provides a unified interface for file operations in both Tauri and Browser environments.
 */

export interface FileFilter {
  name: string;
  extensions: string[];
}

const browserFileStore = new Map<string, File>();

export const isTauri = (): boolean => {
  return (
    typeof window !== 'undefined' &&
    ('__TAURI__' in window || '__TAURI_INTERNALS__' in window)
  );
};

/**
 * Dynamically import Tauri invoke to avoid build-time issues in non-Tauri environments
 */
async function getTauriInvoke() {
  const { invoke } = await import('@tauri-apps/api/core');
  return invoke;
}

export async function openFiles(filters: FileFilter[] = []): Promise<string[]> {
  if (isTauri()) {
    const invoke = await getTauriInvoke();
    return invoke<string[]>('open_files', { filters });
  }

  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    if (filters.length > 0) {
      const exts = filters.flatMap(f => f.extensions.map(ext => `.${ext}`));
      input.accept = exts.join(',');
    }

    input.onchange = () => {
      if (input.files) {
        const files = Array.from(input.files).map((file) => {
          const browserPath = `__browser__/${crypto.randomUUID()}/${file.name}`;
          browserFileStore.set(browserPath, file);
          return browserPath;
        });
        resolve(files);
      } else {
        resolve([]);
      }
    };

    input.click();
  });
}

export async function saveFile(suggestedName: string, filters: FileFilter[] = []): Promise<string> {
  if (isTauri()) {
    const invoke = await getTauriInvoke();
    return invoke<string>('save_file', { suggestedName, filters });
  }

  // Browser fallback: no native save path API, use filename as virtual path.
  return suggestedName.replace(/^[/\\]+/, '');
}

export async function readFileBytes(path: string): Promise<Uint8Array> {
  if (isTauri()) {
    const invoke = await getTauriInvoke();
    const data = await invoke<number[]>('read_file', { path });
    return new Uint8Array(data);
  }

  const file = browserFileStore.get(path);
  if (!file) {
    throw new Error(`File not found in browser session: ${path}`);
  }

  const buffer = await file.arrayBuffer();
  return new Uint8Array(buffer);
}

export async function writeFileBytes(path: string, data: Uint8Array): Promise<void> {
  if (isTauri()) {
    const invoke = await getTauriInvoke();
    await invoke('write_file', { path, data: Array.from(data) });
    return;
  }

  const blobData = new Uint8Array(data.byteLength);
  blobData.set(data);
  const blob = new Blob([blobData.buffer]);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const filename = path.split('/').pop() || path.split('\\').pop() || 'download.bin';
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Opens a URL in the system's default web browser.
 * In a desktop Tauri environment, it invokes the native open_url command.
 * In a standard browser environment, it falls back to window.open or anchor click.
 */
export async function openExternalUrl(url: string): Promise<boolean> {
  if (!url) return false;

  if (isTauri()) {
    try {
      const invoke = await getTauriInvoke();
      await invoke('open_url', { url });
      return true;
    } catch (err) {
      console.warn('Tauri native open_url failed, trying browser fallbacks:', err);
    }
  }

  if (typeof window !== 'undefined') {
    try {
      const win = window.open(url, '_blank', 'noopener,noreferrer');
      if (win) return true;
    } catch {
      // ignore
    }

    try {
      const a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
      }, 100);
      return true;
    } catch (e) {
      console.error('Failed to open link via fallback anchor:', e);
    }
  }

  return false;
}

