/**
 * PDF Filename Decoder Utility
 * 
 * Accurately decodes attachment filenames extracted from PDF documents.
 * Handles various PDF text encodings:
 * - UTF-8 byte sequences mapped to Latin-1/Windows-1252 (mojibake)
 * - UTF-16BE with BOM (0xFE 0xFF)
 * - UTF-16LE with BOM (0xFF 0xFE)
 * - UTF-8 with BOM (0xEF 0xBB 0xBF)
 * - GB18030 / GBK fallback
 * - PDFDocEncoding character set mapping
 * - Octal escapes (\ddd) and URI percent-encoding (%xx)
 * - Path separator stripping and control characters cleaning
 */

const PDF_DOC_ENCODING_MAP: Record<number, number> = {
  0x18: 0x02D8, // BREVE
  0x19: 0x02C7, // CARON
  0x1A: 0x02C6, // MODIFIER LETTER CIRCUMFLEX ACCENT
  0x1B: 0x02D9, // DOT ABOVE
  0x1C: 0x02DD, // DOUBLE ACUTE ACCENT
  0x1D: 0x02DB, // OGONEK
  0x1E: 0x02DA, // RING ABOVE
  0x1F: 0x02DC, // SMALL TILDE
  0x80: 0x2022, // BULLET
  0x81: 0x2020, // DAGGER
  0x82: 0x2021, // DOUBLE DAGGER
  0x83: 0x2026, // HORIZONTAL ELLIPSIS
  0x84: 0x2014, // EM DASH
  0x85: 0x2013, // EN DASH
  0x86: 0x0192, // LATIN SMALL LETTER F WITH HOOK
  0x87: 0x2044, // FRACTION SLASH
  0x88: 0x2039, // SINGLE LEFT-POINTING ANGLE QUOTATION MARK
  0x89: 0x203A, // SINGLE RIGHT-POINTING ANGLE QUOTATION MARK
  0x8A: 0x2212, // MINUS SIGN
  0x8B: 0x2030, // PER MILLE SIGN
  0x8C: 0x201E, // DOUBLE LOW-9 QUOTATION MARK
  0x8D: 0x201C, // LEFT DOUBLE QUOTATION MARK
  0x8E: 0x201D, // RIGHT DOUBLE QUOTATION MARK
  0x8F: 0x2018, // LEFT SINGLE QUOTATION MARK
  0x90: 0x2019, // RIGHT SINGLE QUOTATION MARK
  0x91: 0x201A, // SINGLE LOW-9 QUOTATION MARK
  0x92: 0x2122, // TRADE MARK SIGN
  0x93: 0xFB01, // LATIN SMALL LIGATURE FI
  0x94: 0xFB02, // LATIN SMALL LIGATURE FL
  0x95: 0x0141, // LATIN CAPITAL LETTER L WITH STROKE
  0x96: 0x0152, // LATIN CAPITAL LIGATURE OE
  0x97: 0x0160, // LATIN CAPITAL LETTER S WITH CARON
  0x98: 0x0178, // LATIN CAPITAL LETTER Y WITH DIAERESIS
  0x99: 0x017D, // LATIN CAPITAL LETTER Z WITH CARON
  0x9A: 0x0131, // LATIN SMALL LETTER DOTLESS I
  0x9B: 0x0142, // LATIN SMALL LETTER L WITH STROKE
  0x9C: 0x0153, // LATIN SMALL LIGATURE OE
  0x9D: 0x0161, // LATIN SMALL LETTER S WITH CARON
  0x9E: 0x017E, // LATIN SMALL LETTER Z WITH CARON
  0xA0: 0x20AC, // EURO SIGN
};

/**
 * Clean filename by stripping control characters and path components
 */
export function cleanFilename(name: string): string {
  if (!name) return '';
  // Remove null bytes and CR/LF
  let cleaned = name.replace(/[\0\r\n]/g, '').trim();
  // Strip any leading directory path (both Windows \ and Unix /)
  cleaned = cleaned.replace(/^.*[/\\]/, '');
  return cleaned;
}

/**
 * Decode PDFDocEncoding byte array to Unicode string
 */
function decodePdfDocEncoding(bytes: Uint8Array): string {
  let result = '';
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i];
    const mapped = PDF_DOC_ENCODING_MAP[b];
    if (mapped !== undefined) {
      result += String.fromCharCode(mapped);
    } else {
      result += String.fromCharCode(b);
    }
  }
  return result;
}

/**
 * Decodes a raw PDF attachment filename string into a proper Unicode string.
 *
 * @param raw The raw filename string obtained from PDF parser/engine
 * @returns Clean, properly decoded Unicode filename string
 */
export function decodePdfFilename(raw: string): string {
  if (!raw || typeof raw !== 'string') return '';

  let str = raw.trim();

  // 1. Unescape PDF octal string escapes if present (e.g. \347\250\213)
  if (/\\([0-7]{1,3})/.test(str)) {
    try {
      str = str.replace(/\\([0-7]{1,3})/g, (_, oct) => String.fromCharCode(parseInt(oct, 8)));
    } catch {
      // Continue with original string if replacement fails
    }
  }

  // 2. Unescape URI percent-encoding if present (e.g. %E7%A8%8B)
  if (/%[0-9A-Fa-f]{2}/.test(str)) {
    try {
      const decodedUri = decodeURIComponent(str);
      if (decodedUri !== str) {
        str = decodedUri;
      }
    } catch {
      // Ignore URI decode errors
    }
  }

  // 3. Check if all characters are single-byte (0..255)
  // If there are characters > 255, the string is already a decoded Unicode string.
  const isByteString = Array.from(str).every(c => c.charCodeAt(0) <= 255);

  if (isByteString) {
    const bytes = new Uint8Array(str.length);
    for (let i = 0; i < str.length; i++) {
      bytes[i] = str.charCodeAt(i) & 0xff;
    }

    // 3a. UTF-16BE with BOM (0xFE 0xFF)
    if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
      try {
        const decoded = new TextDecoder('utf-16be').decode(bytes.slice(2));
        return cleanFilename(decoded);
      } catch {
        // Fallback
      }
    }

    // 3b. UTF-16LE with BOM (0xFF 0xFE)
    if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
      try {
        const decoded = new TextDecoder('utf-16le').decode(bytes.slice(2));
        return cleanFilename(decoded);
      } catch {
        // Fallback
      }
    }

    // 3c. UTF-8 with BOM (0xEF 0xBB 0xBF)
    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
      try {
        const decoded = new TextDecoder('utf-8', { fatal: true }).decode(bytes.slice(3));
        return cleanFilename(decoded);
      } catch {
        // Fallback
      }
    }

    // 3d. Check if any byte is non-ASCII (>= 128)
    const hasHighByte = bytes.some(b => b >= 128);

    if (hasHighByte) {
      // Primary: Try UTF-8 (strict) - handles standard multi-byte UTF-8 mojibake
      try {
        const decoded = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
        return cleanFilename(decoded);
      } catch {
        // Not valid UTF-8
      }

      // Secondary: Try GB18030 / GBK (strict) - handles Chinese PDFs encoded in GBK
      try {
        const decoded = new TextDecoder('gb18030', { fatal: true }).decode(bytes);
        if (decoded && !decoded.includes('\uFFFD')) {
          return cleanFilename(decoded);
        }
      } catch {
        // Not valid GBK
      }

      // Tertiary: Fallback to PDFDocEncoding mapping
      return cleanFilename(decodePdfDocEncoding(bytes));
    }
  }

  return cleanFilename(str);
}
