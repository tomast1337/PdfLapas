# PuppeteerGen Usage Guide

PuppeteerGen is a library that implements the PptxGenJS interface to generate PDFs using Puppeteer instead of PowerPoint files. It provides the same familiar API with support for positioning elements in inches, styling, and multiple page layouts.

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Page Sizes and Layouts](#page-sizes-and-layouts)
- [Positioning and Sizing](#positioning-and-sizing)
- [Text Styling](#text-styling)
- [Tables](#tables)
- [Multiple Slides](#multiple-slides)
- [Slide Masters](#slide-masters)
- [Custom Layouts](#custom-layouts)
- [API Reference](#api-reference)
- [Running PdfLapas on a server](#running-pdflapas-on-a-server)

## Installation

```bash
bun install
```

## Quick Start

```typescript
import { PAGE_SIZES, PuppeteerGen } from "@pdflapas/pptxgenjs";

// Create a new presentation
const pres = new PuppeteerGen(PAGE_SIZES.SCREEN_16X9.landscape);

// Add a slide
const slide = pres.addSlide();

// Add text to the slide
slide.addText("Hello World!", {
    x: 1, // 1 inch from left
    y: 1, // 1 inch from top
    fontSize: 32,
    bold: true,
    color: "#0066CC",
});

// Generate PDF
await pres.writeFile({ fileName: "presentation.pdf" });
```

## Page Sizes and Layouts

PuppeteerGen supports multiple predefined page sizes in both landscape and portrait orientations:

### Available Page Sizes

```typescript
import { PAGE_SIZES } from "@pdflapas/pptxgenjs";

// Presentation formats
PAGE_SIZES.SCREEN_4X3.landscape; // 10" × 7.5"
PAGE_SIZES.SCREEN_4X3.portrait; // 7.5" × 10"
PAGE_SIZES.SCREEN_16X9.landscape; // 10" × 5.625" (default)
PAGE_SIZES.SCREEN_16X9.portrait; // 5.625" × 10"
PAGE_SIZES.SCREEN_16X10.landscape; // 10" × 6.25"
PAGE_SIZES.SCREEN_16X10.portrait; // 6.25" × 10"

// Document formats
PAGE_SIZES.LETTER.landscape; // 11" × 8.5"
PAGE_SIZES.LETTER.portrait; // 8.5" × 11"
PAGE_SIZES.LEGAL.landscape; // 14" × 8.5"
PAGE_SIZES.LEGAL.portrait; // 8.5" × 14"
PAGE_SIZES.A4.landscape; // 11.69" × 8.27"
PAGE_SIZES.A4.portrait; // 8.27" × 11.69"
PAGE_SIZES.A3.landscape; // 16.54" × 11.69"
PAGE_SIZES.A3.portrait; // 11.69" × 16.54"
PAGE_SIZES.TABLOID.landscape; // 17" × 11"
PAGE_SIZES.TABLOID.portrait; // 11" × 17"
```

### Usage Examples

```typescript
// Create 16:9 presentation (widescreen)
const pres1 = new PuppeteerGen(PAGE_SIZES.SCREEN_16X9.landscape);

// Create A4 document
const pres2 = new PuppeteerGen(PAGE_SIZES.A4.portrait);

// Create Letter size document
const pres3 = new PuppeteerGen(PAGE_SIZES.LETTER.landscape);

// Use default (16:9 landscape)
const pres4 = new PuppeteerGen();
```

## Positioning and Sizing

All positioning and sizing in PuppeteerGen uses **inches** as the unit of measurement, consistent with PptxGenJS.

### Position Properties

- `x`: Horizontal position from left edge (in inches)
- `y`: Vertical position from top edge (in inches)
- `w`: Width (in inches)
- `h`: Height (in inches)

### Example

```typescript
const slide = pres.addSlide();

// Position text at specific coordinates
slide.addText("Top Left", {
    x: 0.5, // 0.5 inches from left
    y: 0.5, // 0.5 inches from top
    w: 3, // 3 inches wide
    h: 0.5, // 0.5 inches tall
});

// Center text horizontally (assuming 10" wide slide)
slide.addText("Centered", {
    x: 2.5, // (10 - 5) / 2 = 2.5
    y: 2,
    w: 5,
    h: 1,
    align: "center",
});
```

## Text Styling

PuppeteerGen supports rich text formatting options:

### Basic Styling

```typescript
slide.addText("Styled Text", {
    x: 1,
    y: 1,
    fontSize: 24, // Font size in points
    fontFace: "Arial", // Font family
    bold: true, // Bold text
    italic: true, // Italic text
    underline: true, // Underlined text
    color: "FF0000", // Text color: 6-digit hex RGB ("#" optional) or a scheme color such as "accent1"; anything else throws
    fill: { color: "#FFFF00" }, // Background color
});
```

### Text Alignment

```typescript
// Horizontal alignment
slide.addText("Left aligned", { x: 1, y: 1, w: 8, align: "left" });
slide.addText("Center aligned", { x: 1, y: 2, w: 8, align: "center" });
slide.addText("Right aligned", { x: 1, y: 3, w: 8, align: "right" });

// Vertical alignment
slide.addText("Top", { x: 1, y: 1, w: 2, h: 2, valign: "top" });
slide.addText("Middle", { x: 4, y: 1, w: 2, h: 2, valign: "middle" });
slide.addText("Bottom", { x: 7, y: 1, w: 2, h: 2, valign: "bottom" });
```

### Multi-part Text

```typescript
slide.addText(
    [
        { text: "Regular ", options: {} },
        { text: "Bold ", options: { bold: true } },
        { text: "Colored", options: { color: "#FF0000" } },
    ],
    {
        x: 1,
        y: 1,
    },
);
```

## Tables

Create tables with custom styling for cells:

### Basic Table

```typescript
slide.addTable(
    [
        ["Header 1", "Header 2", "Header 3"],
        ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
        ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"],
    ],
    {
        x: 1,
        y: 1,
        w: 8,
        h: 3,
    },
);
```

### Styled Table

```typescript
slide.addTable(
    [
        [
            { text: "Product", options: { bold: true, fill: "#4472C4", color: "#FFFFFF" } },
            { text: "Price", options: { bold: true, fill: "#4472C4", color: "#FFFFFF" } },
            { text: "Stock", options: { bold: true, fill: "#4472C4", color: "#FFFFFF" } },
        ],
        [{ text: "Widget A" }, { text: "$19.99", options: { align: "right" } }, { text: "150", options: { align: "center", color: "#00AA00" } }],
        [{ text: "Widget B" }, { text: "$29.99", options: { align: "right" } }, { text: "25", options: { align: "center", color: "#FF0000" } }],
        [
            { text: "Total", options: { bold: true, fill: "#D9E2F3" } },
            { text: "$49.98", options: { bold: true, align: "right", fill: "#D9E2F3" } },
            { text: "175", options: { bold: true, align: "center", fill: "#D9E2F3" } },
        ],
    ],
    {
        x: 1,
        y: 2,
        w: 8,
        h: 3,
    },
);
```

### Tables from HTML (`tableToSlides`)

`tableToSlides()` reads an HTML `<table>` (column widths, colors, fonts, padding and borders from its
computed styles) and adds as many slides as its rows need, as PptxGenJS does. Besides PptxGenJS's
element id, PdfLapas takes the `<table>` element or HTML markup holding one:

```typescript
// PptxGenJS form: the id of a table in the global `document` (the page; in Node, a jsdom you install).
pres.tableToSlides("orders", { autoPageRepeatHeader: true });
// The element itself, from any document with a window (the page, jsdom).
pres.tableToSlides(document.querySelector("table#orders") as HTMLTableElement);
// HTML markup: a string starting with `<`, parsed with `createDocument` (jsdom in the Node entry).
pres.tableToSlides(`<table><tr><td style="width:2in">Widget</td><td>$19.99</td></tr></table>`);
```

Table options such as `h`, `rowH`, `fontSize` or `margin` only decide how many rows fit on each slide, as in
PptxGenJS 4.0.1: each slide's table gets just its position and column widths, so its rows size to their content
and its cells keep the styles read from the HTML.

The markup form needs a document with a window to compute styles: the Node entry's jsdom has one, a
browser's default does not, so in a browser insert the table into the page and pass the element or its id.

Parsing markup is synchronous, and its time grows with the markup's size (about a millisecond per cell) and
faster than linearly with its element nesting. So markup over 1 MiB, or nesting elements more than 512 deep
(where Chromium's parser stops nesting), throws before it is parsed (`TABLE_HTML_LIMITS`). Raise either limit
per call: `pres.tableToSlides(html, { htmlLimits: { maxLength: 4 * 1024 * 1024, maxDepth: 1024 } })`. When
the markup holds user data, also cap the row count your service accepts.

**`pres.page` is removed** (#112). `tableToSlides(id)` looks the id up in the global `document` only, as
PptxGenJS does; a table put into `pres.page` before now goes in the global `document`, or pass the element or
its HTML.

## Multiple Slides

Add multiple slides to create a complete presentation:

```typescript
const pres = new PuppeteerGen(PAGE_SIZES.SCREEN_16X9.landscape);

// Title slide
const slide1 = pres.addSlide();
slide1.addText("My Presentation", {
    x: 1,
    y: 2,
    w: 8,
    fontSize: 48,
    bold: true,
    align: "center",
});

// Content slide
const slide2 = pres.addSlide();
slide2.addText("Content", {
    x: 1,
    y: 0.5,
    fontSize: 36,
    bold: true,
});
slide2.addText("• Point 1\n• Point 2\n• Point 3", {
    x: 1.5,
    y: 1.5,
    fontSize: 24,
});

// Generate PDF with all slides
await pres.writeFile({ fileName: "multi-slide.pdf" });
```

## Slide Masters

`defineSlideMaster()` defines what several slides share, as in PptxGenJS 4.0.1; `addSlide({ masterName })`
adds a slide on it:

```typescript
pres.defineSlideMaster({
    title: "REPORT",
    background: { color: "F4F7FB" },
    margin: [1, 0.5, 0.75, 0.5], // inches: where autoPage tables on its slides start and stop
    slideNumber: { x: 9, y: 5.1, fontFace: "Arial", fontSize: 10, color: "1F3A5F" },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.7, fill: { color: "1F3A5F" } } },
        { text: { text: "Quarterly report", options: { x: 0.4, y: 0.1, w: 6, h: 0.5, color: "FFFFFF", fontSize: 18 } } },
        { image: { path: "logo.png", x: 8.8, y: 0.1, w: 0.8, h: 0.5 } },
        { line: { x: 0.4, y: 5, w: 9.2, h: 0, line: { color: "7F8FA6", width: 1 } } },
    ],
});

const slide = pres.addSlide({ masterName: "REPORT" });

slide.addText("Drawn on top of the master's objects", { x: 0.5, y: 1.2, w: 9, h: 0.6 });
```

- The master's `objects` (`rect`, `line`, `text`, `image`, `chart`) are drawn in order under each
  slide's own objects. In `toHtml()` they are ordinary elements at the start of each slide.
- A slide shows the master's `background` (or the deprecated `bkgd`) until it sets its own
  `background`, which replaces the master's color and picture.
- Each slide starts with the master's `slideNumber`; setting `slide.slideNumber` replaces it.
- `margin` places autoPage tables on the master's slides (and `tableToSlides({ masterSlideName })`
  pages), whose continuation slides are added on the same master.
- A `placeholder` object draws nothing, as its prompt text is not shown on a slide.
  `addText(text, { placeholder: name })` fills it as PptxGenJS does: the placeholder's options (its
  `x`, `y`, `w`, `h` and text style) win over the text's own, and what neither sets comes from the
  master's text styles: a `title` placeholder is 44pt and centered in `theme.headFontFace`, any other
  type 32pt and left-aligned with 20% space before, of each paragraph's largest run, unless a
  `paraSpaceBefore` above 0 sets it; both are top-anchored and black.
  `addImage({ placeholder: name })` moves the picture to the placeholder's `x` and `y`. A name the
  master does not have draws the object from its own options and adds a `warnings()` line.
- An unknown `masterName` adds a slide without a master, as PptxGenJS does, and logs a warning.
- PptxGenJS's built-in `masterName: "DEFAULT"` is a blank master: its slides start with the `slideNumber`
  a slide set last, and autoPage tables on them keep 0.5in margins whatever their `slideMargin`.

## Custom Layouts

PptxGenJS's layout names work: `pres.layout = "LAYOUT_WIDE"` (13⅓" × 7.5"), `LAYOUT_4x3`,
`LAYOUT_16x10` or `LAYOUT_16x9` (the default). Define custom page dimensions:

```typescript
const pres = new PuppeteerGen();

// Set custom dimensions (12" wide × 6" tall); `pres.layout = "Custom Wide"` selects it again later
pres.defineLayout({
    name: "Custom Wide",
    width: 12,
    height: 6,
});

const slide = pres.addSlide();
slide.addText("Custom Layout", { x: 1, y: 2, fontSize: 36 });

await pres.writeFile({ fileName: "custom.pdf" });
```

Or set the page size directly:

```typescript
const pres = new PuppeteerGen();

pres.setPageSize({
    width: 8,
    height: 8,
    name: "Square",
});
```

## API Reference

### PuppeteerGen Class

#### Constructor

```typescript
new PuppeteerGen(pageSize?: PageSize, options?: PuppeteerGenOptions)
```

Creates a new PuppeteerGen instance with the specified page size (default: 16:9 landscape).
`options` injects the `renderer`, `resolveImage` and `createDocument`; see [BACKENDS.md](./BACKENDS.md).
`options.target` picks the rendering target, `"libreoffice"` (the default) or `"powerpoint"`.

**Image paths and URLs from untrusted input.** Like PptxGenJS, the default resolver reads any image path
the process can read and fetches any http(s) URL, following redirects. It stops a fetch after 30 s and
rejects an image (remote body or file) over 50 MiB (`IMAGE_SOURCE_LIMITS`), and `toHtml()` resolves at
most 8 sources at once. A service that renders decks built from user input should confine it with
`options.imageSources`:

```typescript
new PuppeteerGen(undefined, {
    imageSources: {
        baseDir: "/srv/tenants/acme/images", // relative paths resolve here; paths outside it (symlinks followed) throw
        allowRemote: false, // http(s) URLs throw; data URIs still work
        timeoutMs: 10_000,
        maxBytes: 10 * 1024 * 1024,
    },
});
```

| `imageSources` key | Default                   | Effect                                                                                |
| ------------------ | ------------------------- | ------------------------------------------------------------------------------------- |
| `baseDir`          | none (any readable path)  | relative paths resolve here; a path outside it, symlinks followed, throws             |
| `allowLocalPaths`  | `true`                    | `false` rejects file paths and `file:` URLs                                           |
| `allowRemote`      | `true`                    | `false` rejects http(s) URLs                                                          |
| `allowedHosts`     | none (any host)           | fetch only from these host names (exact, case-insensitive), checked on every redirect |
| `maxRedirects`     | 20 (the Fetch standard's) | redirects one image follows; `0` follows none                                         |
| `timeoutMs`        | 30 000                    | per fetch, headers and body                                                           |
| `maxBytes`         | 50 MiB                    | per image, remote body or file                                                        |

A file path must name an image (`.png`, `.jpg`, `.jpeg`, `.gif`, `.svg`, `.webp`), and so must the file
it resolves to when it is a symlink. Errors name the path or URL as the deck gave it, not where it
resolved.

`imageSources` configures the default resolver only, so it cannot be combined with `resolveImage`. The
default resolver checks host names, not addresses: `allowedHosts` keeps fetches to hosts you trust, but
without it a URL can name an internal host or a cloud metadata address. To fetch from arbitrary hosts
while blocking internal addresses, pass your own `resolveImage` that resolves and checks the address
before it connects.

`PuppeteerGenLibreOffice` and `PuppeteerGenPowerPoint` are `PuppeteerGen` subclasses with the target set,
taking the same arguments: use them to match LibreOffice or PowerPoint rendering. They can be extended
and wrapped like `PuppeteerGen`. See "Choosing a rendering target" in [MIGRATION.md](./MIGRATION.md).

**Fallback fonts for the PowerPoint target.** PowerPoint draws a character the run's font lacks in the
first of Arial, Cambria Math and Segoe UI Symbol that has it, then of East Asian fonts that depend on the run's
font: MS Gothic, Yu Gothic, Microsoft JhengHei, Microsoft YaHei and Malgun Gothic after Arial and Calibri, Yu
Gothic, PMingLiU, SimSun and Malgun Gothic after any other font. An East Asian character keeps the font of the
one before it while that font has it, and an emoji takes Segoe UI Emoji. The PowerPoint target writes that list
after each run's face in the HTML, and cuts a run where a character keeps or changes its font so the renderer
draws each stretch from the same font. These are Microsoft fonts that PdfLapas does not ship: install them
where you render (copy the `.ttf`/`.ttc` files from a licensed Windows or Office installation into a fontconfig
directory such as `~/.local/share/fonts` and run `fc-cache`) to get PowerPoint's glyphs and advances for arrows,
check marks, enclosed numbers, emoji or East Asian text in a Latin run. Where one is missing, both renderers
skip it and fall back as on the LibreOffice target, and the run is not cut for it. To see which of these
faces a machine has, see [Fonts for the PowerPoint target](#fonts-for-the-powerpoint-target). It draws U+2011 (non-breaking
hyphen) with the font's own hyphen, as PowerPoint does.

**Font metrics for the PowerPoint target.** PowerPoint places a single-spaced line's baseline from its
font's `head.unitsPerEm` and `OS/2` `usWinAscent` and `usWinDescent`, read from the installed font the
face name resolves to. The Node entry does the same by default: for each face a deck's text uses, it asks
fontconfig (`fc-match`) for the face's regular font, as Chromium does on Linux, and reads the three values
from that file. A face fontconfig answers with a different family for (a substitute) is not installed and
does not count. PdfLapas also knows the values for Arial, Calibri, Cambria, Courier New and Times New Roman
(`FONT_WIN_METRICS` in `defaults.ts`); text in a face that is neither installed nor listed keeps the
font-independent baseline and sits about 0.03 to 0.05em low against PowerPoint. A face's metrics come from,
in order:

1. `options.fontMetrics`, keyed by face name (case-insensitive);
2. `options.resolveFontMetrics(face)`, by default the installed font on Node;
3. `FONT_WIN_METRICS`;
4. the font-independent baseline, 1em below the line top.

Pass the metrics of faces that are not installed where you render, or to override what is installed, with
`options.fontMetrics`:

```typescript
import { readFileSync } from "node:fs";
import { PuppeteerGenPowerPoint, readFontWinMetrics } from "@pdflapas/pptxgenjs";

const pres = new PuppeteerGenPowerPoint(undefined, {
    fontMetrics: {
        "My Brand Sans": { unitsPerEm: 1000, winAscent: 950, winDescent: 250 },
        // or read them from the font file (TrueType, OpenType or a .ttc's face, the first by default; not WOFF):
        "My Brand Serif": readFontWinMetrics(readFileSync("fonts/MyBrandSerif-Regular.ttf")),
    },
});
```

An entry may also carry `pairKerning: boolean`: whether PowerPoint pair-kerns the face. PowerPoint
kerns a font with a `kern` table from that table alone, so only when it is Microsoft-format (version 0)
with a horizontal format 0 subtable; a face whose only pair kerning is an Apple-format `kern` table plus
GPOS is drawn unkerned at every size, and so is a font with no `kern` table: PowerPoint never kerns from
GPOS. PowerPoint
kerns text from 12pt only in a face it kerns, and the PowerPoint target draws a face with
`pairKerning: false` unkerned at every size. Left out, the face is kerned from 12pt. `readFontWinMetrics`
(and so the default resolver) reads it from the font: true only for a version 0 `kern` table as above.
`FONT_WIN_METRICS` sets it for Arial and Calibri. A face whose glyph advances are read
(`resolveGlyphAdvances`) with no `kern` pairs is drawn unkerned too.

An entry may also carry `underline: { position, thickness }`: the underline band the PowerPoint target draws
for the face, in place of the one in its `post` table. The PowerPoint target draws an underline from the
face's `post` table (`underlinePosition`, `underlineThickness`), which is what PowerPoint draws for Arial,
Calibri and Times New Roman, but for some faces PowerPoint draws another band by a rule that is not known yet
(#391). Both numbers are in the entry's font units (`unitsPerEm`):

- `position` is the band's **top**, measured from the baseline, with the `post` table's sign: negative is
  below the baseline. It is the top of the band, not its centre, as the PowerPoint target reads
  `post.underlinePosition`, so the face's own `post` values give the same underline as leaving the field out
  (the same band in the WASM and PDFKit renderers; Chromium may round the explicit band to a different pixel).
- `thickness` is the band's height, greater than 0.

```typescript
const pres = new PuppeteerGenPowerPoint(undefined, {
    fontMetrics: {
        // Measured in PowerPoint's PDF export of a deck using the face: at 10pt, the band's top is 0.45pt below
        // the baseline and it is 0.72pt tall, so -45 and 72 of 1000 units.
        "My Brand Sans": { unitsPerEm: 1000, winAscent: 950, winDescent: 250, underline: { position: -45, thickness: 72 } },
    },
});
```

Take the numbers from PowerPoint, not from the font: underline a run in a deck using the face, export it to
PDF from PowerPoint, and measure the band's top below the baseline and its height at a known size; then
`position = -top / size * unitsPerEm` and `thickness = height / size * unitsPerEm`. The band applies to the
face in every weight and style, and PowerPoint's double, wavy and heavy underlines are drawn around it as
around the `post` band. `toHtml()` states it as `text-decoration-thickness` and `text-underline-offset` in em
on the run, as it states every band it knows (docs/HTML_PROFILE.md). `readFontWinMetrics` (and so the default resolver) never sets it.

The constructor checks each entry: `unitsPerEm` a positive integer, `winAscent` and `winDescent`
non-negative integers with a positive sum, `pairKerning` a boolean if present, `underline` if present an
object with a finite `position` and a positive finite `thickness` and nothing else, and no other field. An
invalid entry throws a `TypeError` naming the face and field. `readFontWinMetrics(bytes, faceIndex = 0)` is browser-safe (it takes bytes,
not a path) and throws `FontFileError` on bytes that are not a readable TrueType or OpenType font. The
LibreOffice target's line box ignores `fontMetrics` and `resolveFontMetrics`, as LibreOffice's line box does not
depend on the font; it reads a face's `winDescent` and `unitsPerEm` (an entry's, the resolver's or
`FONT_WIN_METRICS`') only for the underline and strike, which LibreOffice sizes and places from the face's descent,
and never an entry's `underline`.

`resolveFontMetrics` takes a face name and returns `{ kind: "metrics", metrics }`, or `{ kind: "unknown" }` to
fall through to `FONT_WIN_METRICS`, or `{ kind: "unreadable", face, path, reason }` for a font file it found
but could not read: the face falls through as an unknown one and `pres.warnings()` lists a line naming the
face, file and reason. It is called at most once per face per presentation, only for the PowerPoint target and
only for faces `fontMetrics` does not list; a result that is not one of those three, or carries invalid metrics,
throws a `TypeError` naming the face. Pass your own to look fonts up elsewhere, or `null` to turn the lookup off:

```typescript
import { readFileSync } from "node:fs";
import { PuppeteerGenPowerPoint, readFontWinMetrics } from "@pdflapas/pptxgenjs";

const brand = readFontWinMetrics(readFileSync("fonts/MyBrandSans-Regular.ttf"));
const pres = new PuppeteerGenPowerPoint(undefined, {
    resolveFontMetrics: face => (face === "My Brand Sans" ? { kind: "metrics", metrics: brand } : { kind: "unknown" }),
});
```

```typescript
// The same HTML on every machine: no installed font is read.
const pres = new PuppeteerGenPowerPoint(undefined, { resolveFontMetrics: null });
```

The default resolver lists the installed families once per process (`fc-list`), so a face that is not
installed costs a lookup in that list and is not cached; it runs `fc-match` only for an installed face and
caches that face's result for the life of the process (a font installed later is seen by a new process). The
lookups are synchronous; on a server, the first deck after start pays for the list. It needs fontconfig's
`fc-list` and `fc-match` on the `PATH`: Linux has them; on macOS and Windows
without fontconfig every face falls through to `FONT_WIN_METRICS`. A font file it cannot read (truncated,
WOFF) is `unreadable`: the face falls through the same way, and `pres.warnings()` names the face and file,
once per presentation (the file is parsed once per process):

```typescript
const { html } = await pres.toHtml();

for (const warning of pres.warnings()) {
    console.warn(warning); // cannot read the metrics of "My Brand Sans" from /usr/share/fonts/… (…); …
}
```

The `@pdflapas/pptxgenjs/html` entry has no default resolver.

**Glyph advances for the PowerPoint target.** PowerPoint lays text out with each glyph's advance rounded to
0.125pt, so a long line can be up to about 0.5% wider or narrower than the font's natural advances and
break elsewhere. The PowerPoint target adds that rounding to each run's `letter-spacing`, from the advance
widths of the font the run's face, weight and style resolve to. `options.resolveGlyphAdvances(face, bold,
italic)` finds them: by default on Node the installed font fontconfig matches (when its family is the face),
read with `readGlyphAdvances(bytes, faceIndex = 0)`, which is browser-safe and throws `FontFileError` on a
font it cannot read. It is called at most once per face, weight and style per presentation, never for the
LibreOffice target. It returns `{ kind: "advances", advances }`, `{ kind: "unknown" }` or `{ kind: "unreadable",
face, path, reason }` (natural advances, and a `warnings()` line, as for `resolveFontMetrics`); anything else, or
advances that are not `{ unitsPerEm, advances: Map<codePoint, fontUnits> }` with an optional
`kerning: Map<leftCodePoint, Map<rightCodePoint, fontUnits>>` (the `kern` table's pairs, which kerned runs add), throws a
`TypeError` naming the face. A face it returns `{ kind: "unknown" }` for, every face with `resolveGlyphAdvances: null`,
and every face in the `@pdflapas/pptxgenjs/html` entry without a resolver keep their natural advances:

```typescript
import { readFileSync } from "node:fs";
import { PuppeteerGenPowerPoint, readGlyphAdvances } from "@pdflapas/pptxgenjs";

const brand = readGlyphAdvances(readFileSync("fonts/MyBrandSans-Regular.ttf"));
const pres = new PuppeteerGenPowerPoint(undefined, {
    // Only the regular face here: bold and italic runs keep their natural advances.
    resolveGlyphAdvances: (face, bold, italic) => (face === "My Brand Sans" && !bold && !italic ? { kind: "advances", advances: brand } : { kind: "unknown" }),
});
```

Like `resolveFontMetrics`, the default makes the PowerPoint target's HTML depend on the installed fonts;
pass `resolveGlyphAdvances: null` too when it must be identical across machines.

**Determinism.** With the default resolver, the PowerPoint target's `toHtml()` and `getModel()` depend on
the fonts installed where they run, as the PDF already does: the same deck gives the same bytes on machines
with the same fonts. Pass `resolveFontMetrics: null` (with `fontMetrics` for the faces you need) when the
HTML must be identical across machines, for example in snapshot tests. The LibreOffice target does not
read installed fonts.

**PDF document properties.** `pres.title`, `pres.author`, `pres.subject`, `pres.company` and
`pres.revision` work as in PptxGenJS and reach the PDF's Info dictionary. The PDF properties PptxGenJS has
no field for are a PdfLapas option, `options.pdfMetadata`, so the PptxGenJS-compatible surface stays as
upstream:

```typescript
const pres = new PuppeteerGen(undefined, {
    pdfMetadata: {
        creator: "Acme Reports", // PDF /Creator, default "PdfLapas"
        producer: "Acme Reports 4.2", // PDF /Producer, default "PdfLapas"
        keywords: "q3, sales", // PDF /Keywords, written as given
        creationDate: new Date("2026-11-19T12:00:00Z"), // PDF /CreationDate
        modificationDate: new Date(), // PDF /ModDate
    },
});
```

Both backends write the same values. Dates are written in UTC to the second (`D:20261119120000Z`), so the
output does not depend on the machine's time zone; with only one of the two dates, it is used for both, as a
document saved once was created and last modified then. Without dates the PDF has none, so the same deck
gives the same bytes; with them, the same deck and options still do. An empty string leaves a property at
its default. The constructor throws a `TypeError` for a value of the wrong type or an unknown key, and a
`RangeError` for an `Invalid Date` or a year outside 0 to 9999. The values travel through `getModel()`
(dates as ISO strings) and `toHtml()` as meta tags ([HTML profile](./HTML_PROFILE.md#document-properties)).
WASM also writes them to the XMP packet; Puppeteer writes no XMP.

#### Methods

- `addSlide(props?: AddSlideProps): PptxSlide`
    - Adds a new slide to the presentation

- `defineLayout(layout: PresLayout): void`
    - Defines a custom page layout with width and height

- `defineSlideMaster(props: SlideMasterProps): void`
    - Defines a slide master that `addSlide({ masterName })` adds slides on; see [Slide Masters](#slide-masters)

- `setPageSize(pageSize: PageSize): void`
    - Changes the page size of the presentation

- `tableToSlides(source: string | HTMLTableElement, props?: TableToSlidesProps): void`
    - Adds slides holding an HTML table: its id, the element or HTML markup; see
      [Tables from HTML](#tables-from-html-tabletoslides)

- `writeFile(props: WriteFileProps): Promise<string>`
    - Writes the PDF to disk (Node) or downloads it (browser) and returns the file name. A trailing
      `.pptx` becomes `.pdf`, a `.pdf` name is kept and any other name gets `.pdf` appended. `fileName`
      is required: without it `writeFile()` throws a `TypeError` (PptxGenJS would use `Presentation.pptx`)

### PuppeteerSlide Class

#### Methods

- `addText(text: string | TextProps[], options?: TextPropsOptions): Slide`
    - Adds text to the slide with optional styling

- `addTable(rows: TableRow[], options?: TableProps): Slide`
    - Adds a table to the slide

- `addImage(options: ImageProps): Slide`
    - Adds an image to the slide

- `addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide`
    - Adds a shape to the slide

#### Properties

- `background: BackgroundProps`
    - `{ color }`, or a picture `{ path }` / `{ data }` stretched over the whole slide under every object
- `hidden: boolean`
    - `true` leaves the slide out of the PDF; the other slides keep their numbers
- `slideNumber: SlideNumberProps`
    - Prints the slide's number (hidden slides counted) in a text box at `{ x, y, w?, h? }`. As in
      PptxGenJS, size, color, font, `valign` and `margin` (points) left unset come from the
      `slideNumber` set last on any slide; see [DEFAULTS.md](./DEFAULTS.md#slide-numbers)

### Common Options

#### TextPropsOptions

- `x`: number - X position in inches
- `y`: number - Y position in inches
- `w`: number - Width in inches
- `h`: number - Height in inches
- `fontSize`: number - Font size in points
- `fontFace`: string - Font family name
- `bold`: boolean - Bold text
- `italic`: boolean - Italic text
- `underline`: `{ style?, color? }` - Underlined text; `style` is an OOXML underline type (`"sng"`,
  `"dbl"`, `"heavy"`, `"dash"`, ...). The deprecated string form (`underline: "heavy"`) and `true`
  (`"sng"`) work as in PptxGenJS; in a table cell PptxGenJS writes nothing for `true`, so it draws none.
  All 17 OOXML types are drawn, with the rendering target's geometry: `words` is one unbroken line under the
  words and the spaces on the PowerPoint target, as PowerPoint draws it, and nothing on the LibreOffice target, as
  LibreOffice 26.2 does. A style that is not an OOXML type throws a `TypeError` naming it (PptxGenJS 4.0.1's types list
  `"dotDashHeave"`, a misspelling PowerPoint rejects; use `"dotDashHeavy"`).
- `color`: string - Text color
- `fill`: { color: string } - Background color
- `align`: "left" | "center" | "right" | "justify" - Horizontal alignment
- `valign`: "top" | "middle" | "bottom" - Vertical alignment

#### TableProps

- `x`: number - X position in inches
- `y`: number - Y position in inches
- `w`: number - Width in inches
- `h`: number - Height in inches

#### TableCellProps

- `text`: string - Cell text content
- `options`: object - Cell styling options
    - `bold`: boolean
    - `color`: string
    - `fill`: string
    - `fontSize`: number
    - `align`: "left" | "center" | "right"

## CSS and Print Layout

PuppeteerGen automatically generates CSS that:

- Sets the correct page dimensions using `@page` CSS
- Positions elements absolutely using the slide-container class
- Handles page breaks between slides
- Ensures proper print output

The CSS is automatically injected into the document and handles:

- Slide containers with proper dimensions
- Absolute positioning of elements
- Text styling and alignment
- Table formatting
- Print media queries for PDF generation

## Examples

See the `examples/comprehensive-example.ts` file for complete working examples of all features.

Run the examples:

```bash
bun run examples/comprehensive-example.ts
```

This will generate several PDF files demonstrating different features:

- `example-16x9-landscape.pdf` - 16:9 presentation format
- `example-a4-portrait.pdf` - A4 document format
- `example-letter-landscape.pdf` - Letter document format
- `example-positioning.pdf` - Positioning examples
- `example-text-styling.pdf` - Text styling examples
- `example-tables.pdf` - Table formatting examples
- `example-multiple-slides.pdf` - Multi-slide presentation
- `example-custom-layout.pdf` - Custom page dimensions

## Running PdfLapas on a server

PdfLapas keeps PptxGenJS's defaults, which suit a script that renders its own decks. When the deck's text,
image sources, options or component props come from users, set these at the service's boundary:

- **Image sources.** Confine the default resolver with `imageSources` (above): `baseDir` or
  `allowLocalPaths: false` for files, `allowRemote: false` or `allowedHosts` for URLs, and lower
  `maxBytes` and `timeoutMs` to what your decks need. Or pass a `resolveImage` that only looks images up
  in your own storage.
- **Renderer.** The Puppeteer renderer prints offline with scripts off; run Chromium with its sandbox
  where you can (see the Puppeteer backend section of [BACKENDS.md](./BACKENDS.md)).
- **Components.** Component renderers are trusted code: escape user data in their markup
  ([COMPONENTS.md](./COMPONENTS.md)).
- **Input size.** PdfLapas, like PptxGenJS, has no global limit on slides, objects or text length, and
  rendering time grows with them. Cap the deck size your service accepts before you build the deck.
- **Errors.** Error messages name the input that failed (a path, a URL, an object name). Log them, and
  return a generic message to users who should not see your inputs.

### Load testing

`bun run bench:load` (in a checkout) keeps N renders in flight for a fixed time, each building a deck and calling
`toPdf()`, as a service does. The decks rotate through text, tables with automatic pagination, images, emoji and East
Asian text, and preset shapes, plus one request in six whose PNG is cut after its header, which must fail inside the
renderer without stalling the others. Puppeteer uses the default renderer (one shared Chromium through the Node
entry); WASM and PDFKit use one long-lived renderer for every request. Every sample interval it prints per-render
latency (p50, p95, p99), throughput, the process's RSS and heap, and the number and total RSS of its child processes
(Chromium, read from `/proc`, so Linux only); after each level it waits for the idle browser to close and prints the
same after a full collection.

| Variable                    | Default          | Meaning                                                  |
| --------------------------- | ---------------- | -------------------------------------------------------- |
| `BENCH_LOAD_ENGINES`        | `puppeteer,wasm` | renderers to load, from `puppeteer`, `wasm` and `pdfkit` |
| `BENCH_LOAD_CONCURRENCY`    | `1,4,16`         | concurrent requests, one level after another             |
| `BENCH_LOAD_SECONDS`        | `60`             | seconds per level                                        |
| `BENCH_LOAD_SAMPLE_SECONDS` | `10`             | seconds between samples                                  |

```bash
BENCH_LOAD_ENGINES=puppeteer BENCH_LOAD_CONCURRENCY=16 BENCH_LOAD_SECONDS=600 BENCH_LOAD_SAMPLE_SECONDS=60 bun run bench:load
```

It is not part of CI: run it on the machine, fonts and decks your service uses.

Measured on one machine (Linux x64, 32 cores, Bun 1.4.4), 120 s per level, the mixed decks above:

| Renderer  | Concurrency | Renders/s | p50     | p95     | Memory under load                            |
| --------- | ----------- | --------- | ------- | ------- | -------------------------------------------- |
| Puppeteer | 1           | 9.4       | 97 ms   | 153 ms  | process 200-260 MB, Chromium 7 procs 0.5 GB  |
| Puppeteer | 4           | 27.7      | 124 ms  | 201 ms  | process 250-280 MB, Chromium 10 procs 0.8 GB |
| Puppeteer | 16          | 40.4      | 353 ms  | 432 ms  | process 300-370 MB, Chromium 20 procs 2 GB   |
| Puppeteer | 64          | 41.7      | 1358 ms | 1566 ms |                                              |
| WASM      | 1           | 47.6      | 15 ms   | 25 ms   | process 0.7-1.5 GB (system fonts loaded)     |
| WASM      | 16          | 47.2      | 20 ms   | 1690 ms |                                              |
| PDFKit    | 1           | 18.9      | 44 ms   | 69 ms   | process about 0.7 GB                         |

What that means for a service:

- **Puppeteer** scales with cores until about 16 renders in flight on 32 cores; past that requests only queue.
  Cap concurrency near half the core count and budget about 120 MB of Chromium per render in flight. The shared
  browser closes one second after the last render and relaunches on the next; after every level the process was
  back at 170-185 MB with no Chromium left, and 10 minutes at 16 (22,272 renders) showed no growth.
- **WASM and PDFKit** render on the calling thread, so concurrency adds latency, not throughput, and a render
  blocks the event loop for its duration (#578). Run them in a few worker threads, each with its own renderer,
  when one core is not enough. Memory does not grow with renders: 10 minutes at 16 (27,851 WASM renders) ended at
  the same 475 MB RSS as two minutes.
- A failing render (the broken image) fails alone: no other request failed or slowed in any run.

### Fonts for the PowerPoint target

The PowerPoint target matches PowerPoint only with the faces PowerPoint draws with installed: the faces the
deck names, and the faces below, which it uses for text with no `fontFace` and for characters a run's face
lacks. When one is missing, fontconfig substitutes another face without a warning and text moves or wraps
differently. The list is exported as `POWERPOINT_FACES`; this table is the same list
(`tests/powerPointFonts.test.ts` checks it).

| Face               | Role                | Without it                                                                                                                             |
| ------------------ | ------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| Calibri            | theme default       | text and chart text with no fontFace are drawn in a substitute with other advances and line heights, so lines wrap and sit differently |
| Calibri Light      | theme default       | text and chart text with no fontFace are drawn in a substitute with other advances and line heights, so lines wrap and sit differently |
| Arial              | theme default       | text and chart text with no fontFace are drawn in a substitute with other advances and line heights, so lines wrap and sit differently |
| Cambria Math       | symbol fallback     | arrows, check marks, enclosed numbers and math symbols the run's face lacks come from another face, with other widths                  |
| Segoe UI Symbol    | symbol fallback     | arrows, check marks, enclosed numbers and math symbols the run's face lacks come from another face, with other widths                  |
| MS Gothic          | East Asian fallback | Han, kana and Hangul in a Latin run come from another East Asian face, with other widths, so lines break differently                   |
| Yu Gothic          | East Asian fallback | Han, kana and Hangul in a Latin run come from another East Asian face, with other widths, so lines break differently                   |
| Microsoft JhengHei | East Asian fallback | Han, kana and Hangul in a Latin run come from another East Asian face, with other widths, so lines break differently                   |
| Microsoft YaHei    | East Asian fallback | Han, kana and Hangul in a Latin run come from another East Asian face, with other widths, so lines break differently                   |
| Malgun Gothic      | East Asian fallback | Han, kana and Hangul in a Latin run come from another East Asian face, with other widths, so lines break differently                   |
| PMingLiU           | East Asian fallback | Han, kana and Hangul in a Latin run come from another East Asian face, with other widths, so lines break differently                   |
| SimSun             | East Asian fallback | Han, kana and Hangul in a Latin run come from another East Asian face, with other widths, so lines break differently                   |
| Segoe UI Emoji     | emoji               | emoji the run's face lacks come from another face, at another advance                                                                  |

**Where to get them.** All of them are Microsoft fonts, licensed with Windows or Microsoft Office, and none of
them may be redistributed: PdfLapas ships none, and a container image you publish must not either. Copy the
`.ttf`/`.ttc` files from a licensed Windows installation (`C:\Windows\Fonts`; some East Asian faces come with
the "Supplemental Fonts" optional feature of their language) or Office installation into a fontconfig
directory on the server, such as `/usr/local/share/fonts` or `~/.local/share/fonts`, and run `fc-cache -f`.
Check that your license covers the server; Microsoft's [font redistribution FAQ](https://learn.microsoft.com/typography/fonts/font-faq)
explains the terms. Metric-compatible open faces (Carlito for Calibri, Liberation Sans for Arial) keep line
breaks close but are not the same glyphs, and the check below reports them as missing.

**Checking a server.** `checkPowerPointFonts()` (Node entry) lists the faces above, then the deck's own, by
whether fontconfig has a font of that family (compared without case, as the renderers resolve them). It never
throws; decide in your startup code whether a missing face is a warning or an error:

```typescript
import { checkPowerPointFonts } from "@pdflapas/pptxgenjs";

const { present, missing } = checkPowerPointFonts({ faces: ["My Brand Sans"] });
if (missing.length) console.warn(`PowerPoint target: missing fonts ${missing.join(", ")}`);
```

Without fontconfig (macOS or Windows without it) every face reports missing. In a checkout,
`bun run fonts:check "My Brand Sans"` prints the same result, with what goes wrong for each missing face, and
exits with status 1 when one is missing.

**The visual Docker image** (`tests/visual/Dockerfile`) has none of these faces: it installs Carlito, Caladea,
Liberation, DejaVu, Noto Color Emoji and Noto Sans CJK, which serve the LibreOffice references. The tests
that compare against PowerPoint's measured line breaks and need a Microsoft face (Arial, Times New Roman,
Calibri, SimSun or Microsoft JhengHei, in `tests/eastAsianLineBreak.test.ts`, `tests/eastAsianJustify.test.ts`,
`tests/fullWidthFit.test.ts` and `tests/advanceGrid.test.ts`) skip where it is not installed, so they run only
on a machine that has it.

## Tips and Best Practices

1. **Choose the right page size**: Use presentation formats (SCREEN_*) for slides and document formats (LETTER, A4) for reports.

2. **Use consistent units**: All measurements are in inches for compatibility with PptxGenJS.

3. **Plan your layout**: Know your page dimensions and plan element positions accordingly.

4. **Test positioning**: Start with simple layouts and gradually add complexity.

5. **Use tables for structured data**: Tables automatically handle borders and cell formatting.

6. **Leverage multi-part text**: Use TextProps arrays for inline styling within a single text block.

7. **Check alignment**: Use alignment properties instead of manual positioning when possible.

## Troubleshooting

### Text is cut off

Make sure the `w` and `h` properties are large enough for your content.

### Elements overlap

Check your positioning values - elements use absolute positioning and won't automatically avoid each other.

### PDF looks different than expected

Verify that you're using the correct page size and that your position/size values are appropriate for that page size.

### Puppeteer won't launch

Make sure all dependencies are installed and that you have the necessary system libraries for Puppeteer.
Chromium is launched without its sandbox by default; to change the launch options (for example to keep the sandbox,
or to point `executablePath` at another Chromium), see the Puppeteer backend section of `docs/BACKENDS.md`.
