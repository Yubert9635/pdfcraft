/**
 * LibreOffice WASM Converter
 * 
 * Uses @matbee/libreoffice-converter WorkerBrowserConverter for document conversion.
 * 
 * Key design decisions:
 * 1. Uses WorkerBrowserConverter instead of BrowserConverter — runs WASM in a
 *    dedicated Web Worker, avoiding main-thread blocking and eliminating the need
 *    for fragile loadModule patches / Cloudflare Rocket Loader workarounds
 * 2. Uses uncompressed paths (soffice.wasm / soffice.data) — works natively with
 *    all servers (Next.js dev, Vercel, Netlify, etc.). For Nginx production,
 *    gzip_static automatically serves the .gz version when available.
 * 3. Specifies browserWorkerJs for the library's internal worker communication
 * 4. Checks SharedArrayBuffer support upfront — fails fast with a clear error
 * 
 * Reference: bentopdf project uses the same approach successfully.
 * 
 * Note: CJK font injection is not supported with WorkerBrowserConverter because
 * the WASM filesystem lives inside the Worker and is not directly accessible.
 * To support CJK, fonts must be pre-baked into soffice.data or provided via
 * a custom worker that writes to the FS before LOK initialization.
 */

import { WorkerBrowserConverter } from '@matbee/libreoffice-converter/browser';

const LIBREOFFICE_PATH = '/libreoffice-wasm/';

export interface LoadProgress {
    phase: 'loading' | 'initializing' | 'converting' | 'complete' | 'ready';
    percent: number;
    message: string;
}

export type ProgressCallback = (progress: LoadProgress) => void;

// Singleton for converter instance
let converterInstance: LibreOfficeConverter | null = null;

export class LibreOfficeConverter {
    private converter: WorkerBrowserConverter | null = null;
    private initialized = false;
    private initializing = false;
    private basePath: string;

    constructor(basePath?: string) {
        this.basePath = basePath || LIBREOFFICE_PATH;
    }

    async initialize(onProgress?: ProgressCallback): Promise<void> {
        if (this.initialized) return;

        if (this.initializing) {
            while (this.initializing) {
                await new Promise(r => setTimeout(r, 100));
            }
            return;
        }

        this.initializing = true;
        let progressCallback = onProgress;

        try {
            progressCallback?.({ phase: 'loading', percent: 0, message: 'Checking environment...' });

            // Fail fast if SharedArrayBuffer / COOP+COEP is missing
            await this.checkEnvironment();

            progressCallback?.({ phase: 'loading', percent: 5, message: 'Loading conversion engine...' });

            // Pre-fetch the Worker script and create a Blob URL.
            // This works around issues where the dev server (Turbopack) or static
            // hosting prevents the Worker constructor from loading the script
            // directly via its URL (e.g., wrong MIME type, transformation, or
            // cross-origin isolation enforcement on sub-resources).
            const workerBlobUrl = await this.createWorkerBlobUrl();

            // Use full URLs (with origin) for all WASM paths.
            // Because the Worker runs from a blob: URL, absolute paths like
            // "/libreoffice-wasm/soffice.js" cannot be resolved correctly by
            // importScripts — the blob: URL scheme has no HTTP origin to resolve
            // against. Full URLs (e.g., "http://localhost:3000/...") work everywhere.
            const origin = typeof window !== 'undefined' ? window.location.origin : '';

            this.converter = new WorkerBrowserConverter({
                sofficeJs: `${origin}${this.basePath}soffice.js`,
                sofficeWasm: `${origin}${this.basePath}soffice.wasm`,
                sofficeData: `${origin}${this.basePath}soffice.data`,
                sofficeWorkerJs: `${origin}${this.basePath}soffice.worker.js`,
                browserWorkerJs: workerBlobUrl,
                verbose: false,
                onProgress: (info: { phase: string; percent: number; message: string }) => {
                    if (progressCallback && !this.initialized) {
                        const simplifiedMessage = `Loading conversion engine (${Math.round(info.percent)}%)...`;
                        progressCallback({
                            phase: info.phase as LoadProgress['phase'],
                            percent: info.percent,
                            message: simplifiedMessage
                        });
                    }
                },
                onReady: () => {
                    console.log('[LibreOffice] Ready!');
                },
                onError: (error: Error) => {
                    console.error('[LibreOffice] Error:', error);
                },
            });

            console.log('[LibreOffice] Starting initialization via WorkerBrowserConverter...');
            const initStart = performance.now();
            await this.converter.initialize();
            const initDuration = Math.round(performance.now() - initStart);
            console.log(`[LibreOffice] Initialization completed in ${initDuration}ms`);

            this.initialized = true;

            // Clean up Blob URL after successful initialization
            URL.revokeObjectURL(workerBlobUrl);

            // Signal completion
            progressCallback?.({ phase: 'ready', percent: 100, message: 'Conversion engine ready!' });

            // Null out the callback to prevent any late-firing progress updates
            progressCallback = undefined;
        } finally {
            this.initializing = false;
        }
    }

    /**
     * Pre-fetch the Worker script and return a Blob URL.
     *
     * Why: The `new Worker(url)` call in WorkerBrowserConverter can fail in
     * certain environments (Next.js Turbopack dev server, some static hosts)
     * because the response lacks required headers, has wrong MIME type, or
     * the bundler transforms the Worker URL. By fetching the script ourselves
     * (which the environment check already proved works) and converting it to
     * a same-origin Blob URL, we guarantee the Worker can load the script.
     */
    private async createWorkerBlobUrl(): Promise<string> {
        const workerScriptUrl = `${this.basePath}browser.worker.global.js`;
        console.log(`[LibreOffice] Pre-fetching Worker script: ${workerScriptUrl}`);

        const response = await fetch(workerScriptUrl);
        if (!response.ok) {
            throw new Error(
                `Failed to fetch Worker script: HTTP ${response.status} ${response.statusText}`
            );
        }

        const scriptText = await response.text();
        const blob = new Blob([scriptText], { type: 'application/javascript' });
        const blobUrl = URL.createObjectURL(blob);
        console.log(`[LibreOffice] Worker Blob URL created (${scriptText.length} bytes)`);
        return blobUrl;
    }

    /**
     * Diagnose environment issues — fail fast if SharedArrayBuffer is not available.
     * SharedArrayBuffer requires Cross-Origin Isolation (COOP + COEP headers).
     */
    private async checkEnvironment(): Promise<void> {
        console.warn('[LibreOffice] === Environment Check ===');

        // 1. Check COOP/COEP — this is the #1 cause of WASM timeout
        const isIsolated = window.crossOriginIsolated;
        console.warn(`[LibreOffice] Cross-Origin Isolated: ${isIsolated ? 'YES ✅' : 'NO ❌'}`);

        // 2. Check SharedArrayBuffer directly
        const hasSAB = typeof SharedArrayBuffer !== 'undefined';
        console.warn(`[LibreOffice] SharedArrayBuffer: ${hasSAB ? 'Available ✅' : 'NOT available ❌'}`);

        if (!isIsolated || !hasSAB) {
            const errorMsg = [
                'LibreOffice WASM requires SharedArrayBuffer for multi-threading.',
                '',
                'SharedArrayBuffer is only available in Cross-Origin Isolated contexts.',
                'Your server MUST return these headers on ALL responses:',
                '  Cross-Origin-Opener-Policy: same-origin',
                '  Cross-Origin-Embedder-Policy: require-corp',
                '  Cross-Origin-Resource-Policy: cross-origin',
                '',
                `Current state: crossOriginIsolated=${isIsolated}, SharedArrayBuffer=${hasSAB}`,
            ].join('\n');
            console.error(`[LibreOffice] ${errorMsg}`);
            throw new Error(
                `SharedArrayBuffer is not available (crossOriginIsolated=${isIsolated}). ` +
                'Your server must set Cross-Origin-Opener-Policy and Cross-Origin-Embedder-Policy headers. ' +
                'See browser console for details.'
            );
        }

        // 3. Check file connectivity
        const files = [
            'soffice.wasm',
            'soffice.data',
            'soffice.js',
            'soffice.worker.js',
            'browser.worker.global.js',
        ];
        for (const file of files) {
            const url = `${this.basePath}${file}`;
            try {
                const start = performance.now();
                const res = await fetch(url, { method: 'HEAD' });
                const duration = Math.round(performance.now() - start);

                if (res.ok) {
                    const size = res.headers.get('content-length');
                    const type = res.headers.get('content-type');
                    const sizeMb = size ? (parseInt(size) / 1024 / 1024).toFixed(2) + 'MB' : 'unknown size';
                    console.warn(
                        `[LibreOffice] ${file}: OK (${res.status}) ${duration}ms | ${sizeMb} | type=${type}`
                    );
                } else {
                    console.error(`[LibreOffice] ${file}: FAILED (${res.status} ${res.statusText})`);
                    throw new Error(`Required file ${file} returned HTTP ${res.status}`);
                }
            } catch (e) {
                if (e instanceof Error && e.message.startsWith('Required file')) throw e;
                console.error(`[LibreOffice] ${file}: NETWORK ERROR`, e);
                throw new Error(`Cannot fetch ${file}: ${e}`);
            }
        }

        console.warn('[LibreOffice] === Environment Check Passed ✅ ===');
    }

    isReady(): boolean {
        return this.initialized && this.converter !== null;
    }

    async convert(file: File, outputFormat: string): Promise<Blob> {
        if (!this.converter) {
            throw new Error('Converter not initialized');
        }

        console.log(`[LibreOffice] Converting ${file.name} to ${outputFormat}...`);
        console.log(`[LibreOffice] File type: ${file.type}, Size: ${file.size} bytes`);

        try {
            const arrayBuffer = await file.arrayBuffer();
            const uint8Array = new Uint8Array(arrayBuffer);
            const ext = file.name.split('.').pop()?.toLowerCase() || '';

            console.log(`[LibreOffice] Detected format from extension: ${ext}`);

            const startTime = Date.now();
            const result = await this.converter.convert(uint8Array, {
                outputFormat: outputFormat as any,
                inputFormat: ext as any,
            }, file.name);

            const duration = Date.now() - startTime;
            console.log(`[LibreOffice] Conversion complete! Duration: ${duration}ms, Size: ${result.data.length} bytes`);

            // Create a copy to avoid SharedArrayBuffer type issues
            const data = new Uint8Array(result.data);
            return new Blob([data], { type: result.mimeType });
        } catch (error) {
            console.error(`[LibreOffice] Conversion FAILED for ${file.name}:`, error);
            throw error;
        }
    }

    async convertToPdf(file: File): Promise<Blob> {
        return this.convert(file, 'pdf');
    }

    async wordToPdf(file: File): Promise<Blob> {
        return this.convertToPdf(file);
    }

    async pptToPdf(file: File): Promise<Blob> {
        return this.convertToPdf(file);
    }

    async excelToPdf(file: File): Promise<Blob> {
        return this.convertToPdf(file);
    }

    async destroy(): Promise<void> {
        if (this.converter) {
            await this.converter.destroy();
        }
        this.converter = null;
        this.initialized = false;
    }
}

export function getLibreOfficeConverter(basePath?: string): LibreOfficeConverter {
    if (!converterInstance) {
        converterInstance = new LibreOfficeConverter(basePath);
    }
    return converterInstance;
}
