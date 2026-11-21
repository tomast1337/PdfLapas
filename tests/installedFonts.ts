// The installed fonts' advances, for the tests that measure against them and skip themselves without the font.
import type { GlyphAdvances } from "../packages/pptxgenjs/src/defaults";
import { resolveInstalledGlyphAdvances } from "../packages/pptxgenjs/src/node";

/** The advances of the font fontconfig resolves `fontFace` to; `undefined` when it is not installed, which skips the test. */
// eslint-disable-next-line pdflapas/no-nullable-return -- a test fixture: `test.skipIf(!font)` reads undefined as "not installed", the only other outcome.
export function installedAdvances(fontFace: string, bold = false, italic = false): GlyphAdvances | undefined {
    const found = resolveInstalledGlyphAdvances(fontFace, bold, italic);

    return found.kind === "advances" ? found.advances : undefined;
}
