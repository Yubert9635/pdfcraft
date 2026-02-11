/**
 * Word to PDF Processor
 * 
 * Converts Word documents to PDF using LibreOffice WASM.
 */

import type {
    ProcessInput,
    ProcessOutput,
    ProgressCallback,
} from '@/types/pdf';
import { PDFErrorCode } from '@/types/pdf';
import { BasePDFProcessor } from '../processor';

export interface WordToPDFOptions {
    /** Reserved for future options */
}

let converterPromise: Promise<any> | null = null;
let converterInstance: any = null;

async function getConverter(onProgress?: (percent: number, message: string) => void): Promise<any> {
    if (converterInstance?.isReady()) return converterInstance;

    if (converterPromise) {
        await converterPromise;
        return converterInstance;
    }

    converterPromise = (async () => {
        const { getLibreOfficeConverter } = await import('@/lib/libreoffice');
        converterInstance = getLibreOfficeConverter();
        await converterInstance.initialize((progress: any) => {
            onProgress?.(progress.percent, progress.message);
        });
    })();

    await converterPromise;
    return converterInstance;
}

export class WordToPDFProcessor extends BasePDFProcessor {
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
                'Please provide exactly one Word document.',
                `Received ${files.length} file(s).`
            );
        }

        const file = files[0];
        const ext = file.name.split('.').pop()?.toLowerCase() || '';
        const validExts = ['docx', 'doc', 'odt', 'rtf'];

        if (!validExts.includes(ext)) {
            return this.createErrorOutput(
                PDFErrorCode.FILE_TYPE_INVALID,
                'Invalid file type. Please upload .docx, .doc, .odt, or .rtf.',
                `Received: ${file.type || file.name}`
            );
        }

        try {
            this.updateProgress(5, 'Loading conversion engine (first time may take 1-2 minutes)...');

            const converter = await getConverter((percent, message) => {
                this.updateProgress(Math.min(percent * 0.8, 80), message);
            });

            if (this.checkCancelled()) {
                return this.createErrorOutput(PDFErrorCode.PROCESSING_CANCELLED, 'Processing was cancelled.');
            }

            this.updateProgress(85, 'Converting Word document to PDF...');

            const pdfBlob = await converter.convertToPdf(file);

            if (this.checkCancelled()) {
                return this.createErrorOutput(PDFErrorCode.PROCESSING_CANCELLED, 'Processing was cancelled.');
            }

            this.updateProgress(100, 'Conversion complete!');

            const baseName = file.name.replace(/\.(docx?|odt|rtf)$/i, '');
            return this.createSuccessOutput(pdfBlob, `${baseName}.pdf`, { format: 'pdf' });

        } catch (error) {
            console.error('Conversion error:', error);
            return this.createErrorOutput(
                PDFErrorCode.PROCESSING_FAILED,
                'Failed to convert Word document to PDF.',
                error instanceof Error ? error.message : 'Unknown error'
            );
        }
    }
}

export function createWordToPDFProcessor(): WordToPDFProcessor {
    return new WordToPDFProcessor();
}

export async function wordToPDF(
    file: File,
    options?: Partial<WordToPDFOptions>,
    onProgress?: ProgressCallback
): Promise<ProcessOutput> {
    const processor = createWordToPDFProcessor();
    return processor.process({ files: [file], options: options || {} }, onProgress);
}
