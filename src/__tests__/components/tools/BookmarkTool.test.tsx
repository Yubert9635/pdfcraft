import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render, screen, fireEvent, waitFor, cleanup, within, createEvent, act } from '@testing-library/react';
import React from 'react';
import { PDFDocument, PDFDict, PDFArray, PDFName, PDFHexString, PDFRef } from 'pdf-lib';

const persistence = vi.hoisted(() => ({ file: null as File | null, realProcessor: false }));
import { NextIntlClientProvider } from 'next-intl';
import messages from '../../../../messages/en.json';
import { BookmarkTool } from '@/components/tools/bookmark/BookmarkTool';
import { processBookmarks } from '@/lib/pdf/processors/bookmark';

vi.mock('@/components/tools/FileUploader', () => ({
  FileUploader: ({ onFilesSelected }: { onFilesSelected: (files: File[]) => void }) => (
    <button
      data-testid="mock-uploader"
      onClick={() => {
        const file = persistence.file ?? new File(['%PDF'], 'book.pdf', { type: 'application/pdf' });
        if (!persistence.file) Object.defineProperty(file, 'arrayBuffer', { value: async () => new ArrayBuffer(4) });
        onFilesSelected([file]);
      }}
    >
      Upload PDF
    </button>
  ),
}));

vi.mock('@/lib/pdf/loader', () => ({
  configurePdfjsWorker: vi.fn(),
  loadPdfLib: () => import('pdf-lib'),
}));

vi.mock('pdfjs-dist', () => ({
  GlobalWorkerOptions: { workerSrc: 'mock-worker.js' },
  getDocument: () => ({
    promise: Promise.resolve({
      numPages: 4,
      getPage: async () => ({}),
      getPageIndex: async (ref: { num: number }) => ref.num - 1,
      getOutline: async () => [
        { title: 'Cover', items: [] },
        { title: 'Page 2', items: [] },
        { title: 'Page 3', items: [] },
        { title: 'Contents', dest: [{ num: 3 }], bold: true, color: [255, 0, 0],
          items: [{ title: 'Chapter 1', dest: [{ num: 4 }], italic: true, items: [] }] },
      ],
    }),
  }),
}));

vi.mock('@/lib/pdf/processors/bookmark', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@/lib/pdf/processors/bookmark')>();
  return { processBookmarks: vi.fn((...args: Parameters<typeof actual.processBookmarks>) =>
    persistence.realProcessor ? actual.processBookmarks(...args) :
      Promise.resolve({ success: true, result: new Blob(['%PDF']) })) };
});

const renderAndLoad = async () => {
  render(
    <NextIntlClientProvider locale="en" messages={messages}>
      <BookmarkTool />
    </NextIntlClientProvider>
  );
  fireEvent.click(screen.getByTestId('mock-uploader'));
  await screen.findByText('Cover');
};

const savedTitles = () => {
  const options = vi.mocked(processBookmarks).mock.calls[0][1];
  const titles: string[] = [];
  const collect = (items: { title: string; children?: unknown[] }[]) => {
    for (const item of items) {
      titles.push(item.title);
      collect((item.children ?? []) as typeof items);
    }
  };
  collect(options.bookmarks ?? []);
  return titles;
};

const originalCreateObjectURL = URL.createObjectURL;

describe('BookmarkTool', () => {
  beforeEach(() => {
    persistence.file = null;
    persistence.realProcessor = false;
    URL.createObjectURL = vi.fn(() => 'blob:mock-url');
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
    vi.spyOn(window, 'confirm').mockReturnValue(true);
    vi.mocked(processBookmarks).mockClear();
  });

  afterEach(() => {
    URL.createObjectURL = originalCreateObjectURL;
    vi.restoreAllMocks();
    cleanup();
  });

  it('removes all bookmarks after confirmation and saves a PDF with none', async () => {
    await renderAndLoad();

    fireEvent.click(screen.getByRole('button', { name: 'Remove all' }));

    expect(window.confirm).toHaveBeenCalledTimes(1);
    expect(screen.queryByText('Cover')).not.toBeInTheDocument();
    expect(screen.getByText('All bookmarks removed. Save to download a PDF with no bookmarks.')).toBeInTheDocument();

    fireEvent.click(screen.getByRole('button', { name: 'Save Bookmarks' }));
    await waitFor(() => expect(processBookmarks).toHaveBeenCalledTimes(1));
    expect(savedTitles()).toEqual([]);
  });

  it('keeps every bookmark when Remove all is cancelled', async () => {
    vi.mocked(window.confirm).mockReturnValue(false);
    await renderAndLoad();

    fireEvent.click(screen.getByRole('button', { name: 'Remove all' }));

    expect(screen.getByText('Cover')).toBeInTheDocument();
    expect(screen.getByText('Chapter 1')).toBeInTheDocument();
  });

  it('Shift+click checks a range and Remove selected deletes only those bookmarks', async () => {
    await renderAndLoad();

    fireEvent.click(screen.getByRole('checkbox', { name: 'Select "Cover"' }));
    fireEvent.click(screen.getByRole('checkbox', { name: 'Select "Page 3"' }), { shiftKey: true });

    expect(screen.getByRole('checkbox', { name: 'Select "Page 2"' })).toBeChecked();
    expect(screen.getByRole('checkbox', { name: 'Select "Contents"' })).not.toBeChecked();

    fireEvent.click(screen.getByRole('button', { name: 'Remove selected (3)' }));

    expect(window.confirm).toHaveBeenCalledWith(
      'Remove 3 selected bookmarks? Their child bookmarks will also be removed.'
    );
    expect(screen.queryByText('Cover')).not.toBeInTheDocument();
    expect(screen.queryByText('Page 2')).not.toBeInTheDocument();
    expect(screen.queryByText('Page 3')).not.toBeInTheDocument();
    expect(screen.queryByRole('button', { name: /Remove selected/ })).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole('button', { name: 'Save Bookmarks' }));
    await waitFor(() => expect(processBookmarks).toHaveBeenCalledTimes(1));
    expect(savedTitles()).toEqual(['Contents', 'Chapter 1']);
  });

  it('removes the children of a checked bookmark and keeps the selection when cancelled', async () => {
    await renderAndLoad();
    fireEvent.click(screen.getByRole('checkbox', { name: 'Select "Contents"' }));

    vi.mocked(window.confirm).mockReturnValueOnce(false);
    fireEvent.click(screen.getByRole('button', { name: 'Remove selected (1)' }));
    expect(screen.getByText('Chapter 1')).toBeInTheDocument();
    expect(screen.getByRole('checkbox', { name: 'Select "Contents"' })).toBeChecked();

    fireEvent.click(screen.getByRole('button', { name: 'Remove selected (1)' }));
    expect(screen.queryByText('Contents')).not.toBeInTheDocument();
    expect(screen.queryByText('Chapter 1')).not.toBeInTheDocument();
    expect(screen.getByText('Cover')).toBeInTheDocument();
  });

  it('still offers Save after the last bookmark is deleted individually', async () => {
    await renderAndLoad();

    for (const title of ['Cover', 'Page 2', 'Page 3', 'Contents']) {
      const row = screen.getByText(title).closest('[draggable="true"]') as HTMLElement;
      fireEvent.click(within(row).getByTitle('Delete'));
    }

    fireEvent.click(screen.getByRole('button', { name: 'Save Bookmarks' }));
    await waitFor(() => expect(processBookmarks).toHaveBeenCalledTimes(1));
    expect(savedTitles()).toEqual([]);
  });
  const dragBookmark = (source: string, target: string, clientY: number) => {
    const sourceRow = screen.getByText(source).closest('[draggable="true"]') as HTMLElement;
    const targetRow = screen.getByText(target).closest('[draggable="true"]') as HTMLElement;
    vi.spyOn(targetRow, 'getBoundingClientRect').mockReturnValue({
      top: 0, height: 100, bottom: 100, left: 0, right: 300, width: 300, x: 0, y: 0,
      toJSON: () => ({}),
    });
    fireEvent.dragStart(sourceRow, { dataTransfer: { setData: vi.fn(), effectAllowed: '' } });
    const over = createEvent.dragOver(targetRow);
    Object.defineProperty(over, 'clientY', { value: clientY });
    fireEvent(targetRow, over);
    fireEvent.drop(targetRow);
    fireEvent.dragEnd(sourceRow);
  };

  it.each([['before', 5], ['inside', 50], ['after', 95]] as const)(
    'preserves the entire tree when dropping a parent %s its descendant', async (_position, y) => {
      await renderAndLoad();
      dragBookmark('Contents', 'Chapter 1', y);
      expect(screen.getByText('Contents')).toBeInTheDocument();
      expect(screen.getByText('Chapter 1')).toBeInTheDocument();
      await act(async () => { fireEvent.click(screen.getByRole('button', { name: 'Save Bookmarks' })); });
      await waitFor(() => expect(processBookmarks).toHaveBeenCalledTimes(1));
      expect(savedTitles()).toEqual(['Cover', 'Page 2', 'Page 3', 'Contents', 'Chapter 1']);
      const items = vi.mocked(processBookmarks).mock.calls[0][1].bookmarks!;
      expect(items[3].children?.map(item => item.title)).toEqual(['Chapter 1']);
    }
  );

  it('keeps a subtree intact when moving its parent before a sibling', async () => {
    await renderAndLoad();
    dragBookmark('Contents', 'Cover', 5);
    await act(async () => { fireEvent.click(screen.getByRole('button', { name: 'Save Bookmarks' })); });
    await waitFor(() => expect(processBookmarks).toHaveBeenCalledTimes(1));
    expect(savedTitles()).toEqual(['Contents', 'Chapter 1', 'Cover', 'Page 2', 'Page 3']);
  });

  it('saves and reloads a valid PDF with the subtree and destinations intact after an invalid drop', async () => {
    const pdf = await PDFDocument.create();
    for (let i = 0; i < 4; i++) pdf.addPage([612, 792]).drawText(`Page ${i + 1}`);
    const bytes = new Uint8Array(await pdf.save());
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    persistence.file = new File([buffer], 'book.pdf', { type: 'application/pdf' });
    Object.defineProperty(persistence.file, 'arrayBuffer', { value: async () => buffer });
    persistence.realProcessor = true;
    await renderAndLoad();
    dragBookmark('Contents', 'Chapter 1', 50);
    await act(async () => { fireEvent.click(screen.getByRole('button', { name: 'Save Bookmarks' })); });
    await waitFor(() => expect(processBookmarks).toHaveBeenCalledTimes(1));
    const output = await vi.mocked(processBookmarks).mock.results[0].value;
    expect(output.success).toBe(true);
    const resultBytes = await new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(output.result as Blob);
    });
    const reloaded = await PDFDocument.load(resultBytes);
    const outlines = reloaded.catalog.lookup(PDFName.of('Outlines'), PDFDict);
    const titles: string[] = [];
    let current = outlines.get(PDFName.of('First'));
    let contents: PDFDict | undefined;
    while (current) {
      const item = reloaded.context.lookup(current, PDFDict);
      titles.push(item.lookup(PDFName.of('Title'), PDFHexString).decodeText());
      if (titles.at(-1) === 'Contents') contents = item;
      current = item.get(PDFName.of('Next'));
    }
    expect(titles).toEqual(['Cover', 'Page 2', 'Page 3', 'Contents']);
    expect(reloaded.getPageCount()).toBe(4);
    expect(contents!.lookup(PDFName.of('Dest'), PDFArray).get(0)).toEqual(reloaded.getPage(2).ref);
    expect(contents!.get(PDFName.of('F'))?.toString()).toBe('2');
    expect(contents!.lookup(PDFName.of('C'), PDFArray).toString()).toBe('[ 1 0 0 ]');
    const childRef = contents!.get(PDFName.of('First')) as PDFRef;
    const child = reloaded.context.lookup(childRef, PDFDict);
    expect(child.lookup(PDFName.of('Title'), PDFHexString).decodeText()).toBe('Chapter 1');
    expect(child.lookup(PDFName.of('Dest'), PDFArray).get(0)).toEqual(reloaded.getPage(3).ref);
    expect(child.get(PDFName.of('F'))?.toString()).toBe('1');
    expect(child.get(PDFName.of('Parent'))).toEqual(reloaded.context.getObjectRef(contents!));
  });

});
