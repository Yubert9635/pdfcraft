/**
 * Excel to PDF Processor
 * 
 * Converts Excel spreadsheets to PDF using LibreOffice WASM.
 * Uses the shared LibreOfficeConverter singleton (same approach as BentoPDF).
 */

import type {
    ProcessInput,
    ProcessOutput,
    ProgressCallback,
} from '@/types/pdf';
import { PDFErrorCode } from '@/types/pdf';
import { BasePDFProcessor } from '../processor';
import { getLibreOfficeConverter } from '@/lib/libreoffice';

export interface ExcelToPDFOptions {
    /** Reserved for future options */
}

export class ExcelToPDFProcessor extends BasePDFProcessor {
    protected reset(): void {
        super.reset();
    }

    async process(
        input: ProcessInput,
        onProgress?: ProgressCallback
    ): Promise<ProcessOutput> {
        this.reset();
        this.onProgress = onProgress;

        const { files } = input;

        if (files.length !== 1) {
            return this.createErrorOutput(
                PDFErrorCode.INVALID_OPTIONS,
                'Please provide exactly one Excel spreadsheet.',
                `Received ${files.length} file(s).`
            );
        }

        const file = files[0];
        const ext = file.name.split('.').pop()?.toLowerCase() || '';
        const validExts = ['xlsx', 'xls', 'ods', 'csv'];

        if (!validExts.includes(ext)) {
            return this.createErrorOutput(
                PDFErrorCode.FILE_TYPE_INVALID,
                'Invalid file type. Please upload .xlsx, .xls, .ods, or .csv.',
                `Received: ${file.type || file.name}`
            );
        }

        try {
            this.updateProgress(5, 'Loading conversion engine (first time may take 1-2 minutes)...');

            const converter = getLibreOfficeConverter();

            await converter.initialize((progress) => {
                this.updateProgress(Math.min(progress.percent * 0.8, 80), progress.message);
            });

            if (this.checkCancelled()) {
                return this.createErrorOutput(PDFErrorCode.PROCESSING_CANCELLED, 'Processing was cancelled.');
            }

            this.updateProgress(85, 'Converting Excel to PDF...');

            const pdfBlob = await converter.convertToPdf(file);

            if (this.checkCancelled()) {
                return this.createErrorOutput(PDFErrorCode.PROCESSING_CANCELLED, 'Processing was cancelled.');
            }

            this.updateProgress(100, 'Conversion complete!');

            const baseName = file.name.replace(/\.(xlsx?|ods|csv)$/i, '');
            return this.createSuccessOutput(pdfBlob, `${baseName}.pdf`, { format: 'pdf' });

        } catch (error) {
            console.error('Conversion error:', error);
            return this.createErrorOutput(
                PDFErrorCode.PROCESSING_FAILED,
                'Failed to convert Excel to PDF.',
                error instanceof Error ? error.message : 'Unknown error'
            );
        }
    }
}

export function createExcelToPDFProcessor(): ExcelToPDFProcessor {
    return new ExcelToPDFProcessor();
}

export async function excelToPDF(
    file: File,
    options?: Partial<ExcelToPDFOptions>,
    onProgress?: ProgressCallback
): Promise<ProcessOutput> {
    const processor = createExcelToPDFProcessor();
    return processor.process({ files: [file], options: options || {} }, onProgress);
}
