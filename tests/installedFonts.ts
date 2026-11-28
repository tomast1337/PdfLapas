// The installed fonts' advances, for the tests that measure against them and skip themselves without the font.
import type { GlyphAdvances } from "../packages/pptxgenjs/src/defaults";
import { resolveInstalledGlyphAdvances } from "../packages/pptxgenjs/src/node";

/** The advances of the font fontconfig resolves `fontFace` to; `undefined` when it is not installed, which skips the test. */
// eslint-disable-next-line pdflapas/no-nullable-return -- a test fixture: `test.skipIf(!font)` reads undefined as "not installed", the only other outcome.
export function installedAdvances(fontFace: string, bold = false, italic = false): GlyphAdvances | undefined {
    const found = resolveInstalledGlyphAdvances(fontFace, bold, italic);

    return found.kind === "advances" ? found.advances : undefined;
}

// The faces Chromium fell back to for Han when the East Asian boxes were measured against it: MS Gothic (with the
// Office fonts installed) and Droid Sans Fallback (without them). Under Noto Sans CJK, the first CJK face Ubuntu's
// fontconfig offers, Chromium breaks hundreds of those boxes elsewhere (#684), so its tests run only with these.
const MEASURED_HAN_FALLBACKS: ReadonlySet<string> = new Set(["MS Gothic", "Droid Sans", "Droid Sans Fallback"]);

/** Whether fontconfig gives Chromium a Han fallback face the East Asian boxes were measured with. */
export function chromiumHanFallbackMeasured(): boolean {
    const fcMatch = Bun.which("fc-match");

    if (fcMatch === null) {
        return false;
    }

    // U+4E2D (中): the face Chromium's per-character fallback takes first, as fontconfig sorts them for Arial.
    const face = Bun.spawnSync([fcMatch, "-f", "%{family[0]}", "Arial:charset=4e2d"]).stdout.toString();

    return MEASURED_HAN_FALLBACKS.has(face);
}
