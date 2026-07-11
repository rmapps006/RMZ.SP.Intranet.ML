import { RALEWAY_WOFF2_BASE64, PUBLIC_SANS_WOFF2_BASE64 } from './fontData';

/**
 * Loads the ORA brand fonts (Raleway + Public Sans) from woff2 data BUNDLED
 * inside the solution — no external CDN dependency.
 *
 * Both families are variable fonts, so a single woff2 per family covers the
 * full weight axis used by the design (Raleway 100–500, Public Sans 300–600).
 * The fonts are embedded as base64 data URIs (see fontData.ts) and injected
 * once per page via a <style> element.
 */
const FONTS_MARKER: string = 'ora-brand-fonts';

function dataUri(base64: string): string {
  return `data:font/woff2;base64,${base64}`;
}

export function ensureBrandFonts(): void {
  if (document.getElementById(FONTS_MARKER)) {
    return;
  }

  const css: string = [
    '@font-face{',
    "font-family:'Raleway';font-style:normal;font-weight:100 500;font-display:swap;",
    `src:url("${dataUri(RALEWAY_WOFF2_BASE64)}") format("woff2");`,
    '}',
    '@font-face{',
    "font-family:'Public Sans';font-style:normal;font-weight:300 600;font-display:swap;",
    `src:url("${dataUri(PUBLIC_SANS_WOFF2_BASE64)}") format("woff2");`,
    '}'
  ].join('');

  const style: HTMLStyleElement = document.createElement('style');
  style.id = FONTS_MARKER;
  style.appendChild(document.createTextNode(css));
  document.head.appendChild(style);
}
