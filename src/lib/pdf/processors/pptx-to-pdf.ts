/**
 * PowerPoint to PDF Processor
 * 
 * Converts PowerPoint presentations to PDF using LibreOffice WASM.
 */

import type {
    ProcessInput,
    ProcessOutput,
    ProgressCallback,
} from '@/types/pdf';
import { PDFErrorCode } from '@/types/pdf';
import { BasePDFProcessor } from '../processor';

export interface PPTXToPDFOptions {
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

export class PPTXToPDFProcessor extends BasePDFProcessor {
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
                'Please provide exactly one PowerPoint presentation.',
                `Received ${files.length} file(s).`
            );
        }

        const file = files[0];
        const ext = file.name.split('.').pop()?.toLowerCase() || '';
        const validExts = ['pptx', 'ppt', 'odp'];

        if (!validExts.includes(ext)) {
            return this.createErrorOutput(
                PDFErrorCode.FILE_TYPE_INVALID,
                'Invalid file type. Please upload .pptx, .ppt, or .odp.',
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

            this.updateProgress(85, 'Converting PowerPoint to PDF...');

            const pdfBlob = await converter.convertToPdf(file);

            if (this.checkCancelled()) {
                return this.createErrorOutput(PDFErrorCode.PROCESSING_CANCELLED, 'Processing was cancelled.');
            }

            this.updateProgress(100, 'Conversion complete!');

            const baseName = file.name.replace(/\.(pptx?|odp)$/i, '');
            return this.createSuccessOutput(pdfBlob, `${baseName}.pdf`, { format: 'pdf' });

        } catch (error) {
            console.error('Conversion error:', error);
            return this.createErrorOutput(
                PDFErrorCode.PROCESSING_FAILED,
                'Failed to convert PowerPoint to PDF.',
                error instanceof Error ? error.message : 'Unknown error'
            );
        }
    }
}

export function createPPTXToPDFProcessor(): PPTXToPDFProcessor {
    return new PPTXToPDFProcessor();
}

export async function pptxToPDF(
    file: File,
    options?: Partial<PPTXToPDFOptions>,
    onProgress?: ProgressCallback
): Promise<ProcessOutput> {
    const processor = createPPTXToPDFProcessor();
    return processor.process({ files: [file], options: options || {} }, onProgress);
}
