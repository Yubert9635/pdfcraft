import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { render, screen, fireEvent, waitFor, cleanup, act } from '@testing-library/react';
import React from 'react';
import { PDFReaderTool } from '@/components/tools/pdf-reader/PDFReaderTool';
import * as tauriBridge from '@/lib/tauri-bridge';

// Mock next-intl
vi.mock('next-intl', () => ({
  useTranslations: () => (key: string) => {
    const translations: Record<string, string> = {
      'pdfReader.uploadLabel': 'Open PDF File',
      'pdfReader.uploadDescription': 'Open a PDF file to read and view.',
      'pdfReader.print': 'Print (Ctrl+P)',
      'pdfReader.printModalTitle': 'Print Document',
    };
    return translations[key] || key;
  },
}));

// Mock path util
vi.mock('@/lib/utils/path', () => ({
  withBasePath: (path: string) => path,
}));

// Mock FileUploader
vi.mock('@/components/tools/FileUploader', () => ({
  FileUploader: ({ onFilesSelected }: { onFilesSelected: (files: File[]) => void }) => (
    <button
      data-testid="mock-uploader"
      onClick={() => onFilesSelected([new File(['test content'], 'sample.pdf', { type: 'application/pdf' })])}
    >
      Upload PDF
    </button>
  ),
}));

const originalCreateObjectURL = URL.createObjectURL;
const originalRevokeObjectURL = URL.revokeObjectURL;

describe('PDFReaderTool', () => {
  beforeEach(() => {
    URL.createObjectURL = vi.fn(() => 'blob:mock-url');
    URL.revokeObjectURL = vi.fn();
    vi.clearAllMocks();
  });

  afterEach(() => {
    URL.createObjectURL = originalCreateObjectURL;
    URL.revokeObjectURL = originalRevokeObjectURL;
    cleanup();
  });

  it('renders upload component initially', () => {
    render(<PDFReaderTool />);
    expect(screen.getByTestId('mock-uploader')).toBeInTheDocument();
  });

  it('renders viewer and print button after file is selected', async () => {
    render(<PDFReaderTool />);
    fireEvent.click(screen.getByTestId('mock-uploader'));

    await waitFor(() => {
      expect(screen.getByText('sample.pdf')).toBeInTheDocument();
    });

    const printBtn = screen.getByTitle('打印 (Ctrl+P)');
    expect(printBtn).toBeInTheDocument();
  });

  it('falls back to window.print in browser mode when print button is clicked', async () => {
    vi.spyOn(tauriBridge, 'isTauri').mockReturnValue(false);
    const printSpy = vi.fn();
    window.print = printSpy;
    const originalGetter = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'contentWindow');
    Object.defineProperty(HTMLIFrameElement.prototype, 'contentWindow', {
      configurable: true,
      get: () => ({ print: printSpy }),
    });

    try {
      render(<PDFReaderTool />);
      fireEvent.click(screen.getByTestId('mock-uploader'));

      await waitFor(() => {
        expect(screen.getByText('sample.pdf')).toBeInTheDocument();
      });

      const printBtn = screen.getByTitle('打印 (Ctrl+P)');
      fireEvent.click(printBtn);

      expect(printSpy).toHaveBeenCalled();
    } finally {
      if (originalGetter) {
        Object.defineProperty(HTMLIFrameElement.prototype, 'contentWindow', originalGetter);
      }
    }
  });

  it('opens native print modal in Tauri desktop mode when print button is clicked', async () => {
    vi.spyOn(tauriBridge, 'isTauri').mockReturnValue(true);
    vi.spyOn(tauriBridge, 'getSystemPrinters').mockResolvedValue([
      { name: 'HP LaserJet M14', is_default: true, port_name: 'USB001' },
      { name: 'SHARP MX-M3558NV', is_default: false, port_name: '192.168.1.100' },
    ]);

    render(<PDFReaderTool />);
    fireEvent.click(screen.getByTestId('mock-uploader'));

    await waitFor(() => {
      expect(screen.getByText('sample.pdf')).toBeInTheDocument();
    });

    const printBtn = screen.getByTitle('打印 (Ctrl+P)');
    fireEvent.click(printBtn);

    await waitFor(() => {
      expect(screen.getByText('打印文档')).toBeInTheDocument();
      expect(screen.getByText('HP LaserJet M14 (默认)')).toBeInTheDocument();
    });
  });

  it('responds to pdfcraft-print-requested message posted by viewer iframe', async () => {
    vi.spyOn(tauriBridge, 'isTauri').mockReturnValue(true);
    vi.spyOn(tauriBridge, 'getSystemPrinters').mockResolvedValue([
      { name: 'Office Printer', is_default: true, port_name: 'NET' },
    ]);

    render(<PDFReaderTool />);
    fireEvent.click(screen.getByTestId('mock-uploader'));

    await waitFor(() => {
      expect(screen.getByText('sample.pdf')).toBeInTheDocument();
    });

    // Simulate postMessage event from viewer iframe
    act(() => {
      window.dispatchEvent(
        new MessageEvent('message', {
          data: { type: 'pdfcraft-print-requested' },
        })
      );
    });

    await waitFor(() => {
      expect(screen.getByText('打印文档')).toBeInTheDocument();
      expect(screen.getByText('Office Printer (默认)')).toBeInTheDocument();
    });
  });
});
