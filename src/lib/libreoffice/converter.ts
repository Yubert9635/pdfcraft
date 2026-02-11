/**
 * LibreOffice WASM Converter
 * 
 * Uses @matbee/libreoffice-converter for document conversion.
 * 
 * Key customizations:
 * 1. Overrides loadModule() to increase timeout from 60s to 5 minutes
 *    (soffice.wasm ~48MB + soffice.data ~28MB compressed need time to download)
 * 2. Adds data-cfasync="false" to script tag to bypass Cloudflare Rocket Loader
 *    which intercepts/defers JS and breaks WASM initialization
 * 3. Injects CJK fonts into WASM virtual filesystem for Chinese support
 */

import { BrowserConverter } from '@matbee/libreoffice-converter/browser';

const LIBREOFFICE_PATH = '/libreoffice-wasm/';

/** Timeout for WASM loading: 5 minutes (soffice.wasm + soffice.data are ~77MB compressed) */
const WASM_LOAD_TIMEOUT_MS = 5 * 60 * 1000;

/**
 * CJK font files to inject into LibreOffice WASM virtual filesystem.
 * These are fetched from /fonts/ and written to /instdir/share/fonts/truetype/
 * so LibreOffice can render CJK characters correctly.
 */
const CJK_FONTS = [
    { url: '/fonts/NotoSansSC-Regular.ttf', filename: 'NotoSansSC-Regular.ttf' },
];

export interface LoadProgress {
    phase: 'loading' | 'initializing' | 'converting' | 'complete' | 'ready';
    percent: number;
    message: string;
}

export type ProgressCallback = (progress: LoadProgress) => void;

let converterInstance: LibreOfficeConverter | null = null;

export class LibreOfficeConverter {
    private converter: BrowserConverter | null = null;
    private initialized = false;
    private initializing = false;
    private basePath: string;
    private fontsInstalled = false;

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
            progressCallback?.({ phase: 'loading', percent: 0, message: 'Loading conversion engine...' });

            this.converter = new BrowserConverter({
                sofficeJs: `${this.basePath}soffice.js`,
                sofficeWasm: `${this.basePath}soffice.wasm`,
                sofficeData: `${this.basePath}soffice.data`,
                sofficeWorkerJs: `${this.basePath}soffice.worker.js`,
                verbose: false,
                onProgress: (info: { phase: string; percent: number; message: string }) => {
                    if (progressCallback && !this.initialized) {
                        progressCallback({
                            phase: info.phase as LoadProgress['phase'],
                            percent: Math.min(info.percent, 90),
                            message: `Loading conversion engine (${Math.round(info.percent)}%)...`
                        });
                    }
                },
                onReady: () => {
                    console.log('[LibreOffice] WASM ready');
                },
                onError: (error: Error) => {
                    console.error('[LibreOffice] Error:', error);
                },
            });

            // Override loadModule to increase timeout and bypass Cloudflare Rocket Loader
            this.patchLoadModule(this.converter);

            await this.converter.initialize();

            // Install CJK fonts into the WASM virtual filesystem
            progressCallback?.({ phase: 'initializing', percent: 92, message: 'Installing CJK fonts...' });
            await this.installCJKFonts();

            this.initialized = true;
            progressCallback?.({ phase: 'ready', percent: 100, message: 'Conversion engine ready!' });
            progressCallback = undefined;
        } finally {
            this.initializing = false;
        }
    }

    /**
     * Monkey-patch BrowserConverter.loadModule() to:
     * 1. Increase timeout from 60s to 5 minutes
     * 2. Add data-cfasync="false" to bypass Cloudflare Rocket Loader
     */
    private patchLoadModule(converter: BrowserConverter): void {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const instance = converter as any;
        const options = instance.options;

        instance.loadModule = () => {
            const { sofficeJs, sofficeWasm, sofficeData, sofficeWorkerJs } = options;
            const w = window as any;

            return new Promise((resolve, reject) => {
                w.Module = {
                    locateFile: (path: string) => {
                        if (path.endsWith('.wasm')) return sofficeWasm;
                        if (path.endsWith('.data')) return sofficeData;
                        if (path.endsWith('.worker.js') || path.endsWith('.worker.cjs')) return sofficeWorkerJs;
                        return `${sofficeJs.substring(0, sofficeJs.lastIndexOf('/') + 1)}${path}`;
                    },
                    print: options.verbose ? console.log : () => { },
                    printErr: options.verbose ? console.error : () => { },
                    onRuntimeInitialized: () => {
                        console.log('[LibreOffice] WASM runtime initialized');
                        resolve(w.Module);
                    },
                    onAbort: (reason: string) => {
                        reject(new Error(`WASM abort: ${reason}`));
                    },
                };

                const script = document.createElement('script');
                script.src = sofficeJs;

                // Bypass Cloudflare Rocket Loader - prevents it from deferring this script
                script.setAttribute('data-cfasync', 'false');

                script.onerror = () => reject(new Error(`Failed to load ${sofficeJs}`));
                document.head.appendChild(script);

                // Extended timeout: 5 minutes instead of default 60 seconds
                setTimeout(() => {
                    reject(new Error('WASM load timeout (5 min)'));
                }, WASM_LOAD_TIMEOUT_MS);
            });
        };
    }

    /**
     * Install CJK fonts into LibreOffice WASM virtual filesystem.
     * This is necessary because the default soffice.data doesn't include
     * CJK fonts, causing Chinese/Japanese/Korean characters to render
     * as garbled text or empty boxes in converted documents.
     */
    private async installCJKFonts(): Promise<void> {
        if (this.fontsInstalled) return;

        // Access the Emscripten module's virtual filesystem
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const module = (this.converter as any)?.module;
        if (!module?.FS) {
            console.warn('[LibreOffice] Cannot access WASM FS, CJK fonts not installed');
            return;
        }

        const FS = module.FS;

        // Ensure the font directories exist
        const fontDirs = [
            '/instdir/share/fonts',
            '/instdir/share/fonts/truetype',
        ];
        for (const dir of fontDirs) {
            try { FS.mkdir(dir); } catch { /* directory may already exist */ }
        }

        // Fetch and install each CJK font
        for (const font of CJK_FONTS) {
            try {
                console.log(`[LibreOffice] Downloading CJK font: ${font.filename}...`);
                const response = await fetch(font.url);
                if (!response.ok) {
                    console.warn(`[LibreOffice] Failed to fetch font ${font.url}: ${response.status}`);
                    continue;
                }
                const fontBuffer = await response.arrayBuffer();
                const fontData = new Uint8Array(fontBuffer);

                const fontPath = `/instdir/share/fonts/truetype/${font.filename}`;
                FS.writeFile(fontPath, fontData);
                console.log(`[LibreOffice] Installed CJK font: ${fontPath} (${(fontData.length / 1024 / 1024).toFixed(1)}MB)`);
            } catch (err) {
                console.warn(`[LibreOffice] Failed to install font ${font.filename}:`, err);
            }
        }

        this.fontsInstalled = true;
    }

    isReady(): boolean {
        return this.initialized && this.converter !== null;
    }

    async convert(file: File, outputFormat: string): Promise<Blob> {
        if (!this.converter) {
            throw new Error('Converter not initialized');
        }

        const arrayBuffer = await file.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);
        const ext = file.name.split('.').pop()?.toLowerCase() || '';

        const result = await this.converter.convert(uint8Array, {
            outputFormat: outputFormat as any,
            inputFormat: ext as any,
        }, file.name);

        const data = new Uint8Array(result.data);
        return new Blob([data], { type: result.mimeType });
    }

    async convertToPdf(file: File): Promise<Blob> {
        return this.convert(file, 'pdf');
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
