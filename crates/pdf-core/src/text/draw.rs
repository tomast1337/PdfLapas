//! Drawing laid-out text boxes: glyph runs, outlines, highlights and decorations.

use std::sync::Arc;

use krilla::Data;
use krilla::color::rgb;
use krilla::geom::Quadrilateral;
use krilla::geom::{Point, Transform};
use krilla::paint::Stroke;
use krilla::surface::Surface;
use krilla::text::{Font, GlyphId, KrillaGlyph};
use kurbo::{BezPath, Rect};
use parley::{GlyphRun, Layout, Line, PositionedLayoutItem};

use crate::Renderer;
use crate::color::{Rgba, fill_of, unit};
use crate::effect::Effect;
use crate::effect::EffectSource;
use crate::link::page_quad;
use crate::model::{DecorationPiece, DecorationShape, TextBox, TextDirection, VerticalAlign, ZigzagLine};
use crate::path::{rect_path, stroke_width, to_path};
use crate::text::layout::{BoxLayout, Brush, BulletMarker, RunPaint, RunRef, TextStroke};
use crate::text::outline::glyph_outlines;

/// Most distinct glows drawn under one text box.
const MAX_GLOWS_PER_BOX: usize = 16;

/// Thinnest band a decoration is drawn around, px (docs/HTML_PROFILE.md).
const MIN_DECORATION_PX: f32 = 0.5;
/// Widest gap, px, between two glyph runs of one run that still continue one decoration segment: parley sums their
/// offsets, so a run split by a style change meets itself again up to float rounding.
const SEGMENT_JOIN_PX: f32 = 1e-3;
/// Widest a font's box reaches from a glyph's pen position, in em: the full i16 range of `head`'s
/// box over the smallest units per em (1), so the text placement check holds for any font.
const FONT_BOX_EM: f64 = 65536.0;
/// Dashes, half-waves and zigzag halves one render draws for underlines and strikes: `MAX_RENDER_DECORATION_PIECES` in
/// packages/pdfkit/src/text/draw.ts. Their count grows with a run's width over its dash length, not with the HTML:
/// four letters spaced 1e7 px with a dashed underline took 16.6 s and +1.4 GB. A slide of dotted lines spends a few
/// thousand.
pub(crate) const MAX_RENDER_DECORATION_PIECES: usize = 1_000_000;

/// A decoration band in px: its top `offset` above the baseline (below it when negative), `size` tall.
#[derive(Clone, Copy, PartialEq, Debug)]
pub(crate) struct Band {
    pub(crate) offset: f32,
    pub(crate) size: f32,
}

/// What one decoration is drawn along: a run's part on one line.
#[derive(Clone, Copy, Debug)]
pub(crate) struct Segment {
    /// The run's paint, its parley brush.
    pub(crate) paint: usize,
    /// The strike's segment, else the underline's.
    pub(crate) strike: bool,
    pub(crate) start: f32,
    pub(crate) end: f32,
    pub(crate) baseline: f32,
    /// The face's own band for this decoration: `post` for the underline, OS/2 for the strike.
    pub(crate) face: Band,
    pub(crate) font_size: f32,
}

impl Segment {
    /// The band `shape` states, in em, else the face's.
    fn band(&self, shape: &DecorationShape) -> Band {
        match shape.band {
            Some(band) => Band {
                offset: band.offset * self.font_size,
                size: band.size * self.font_size,
            },
            None => self.face,
        }
    }
}

/// Fills `out` with the decoration segments of `line`, drawn with its top at `oy` and its start at `ox`: each run's
/// glyph runs on the line joined into one.
pub(crate) fn line_segments(line: &Line<'_, Brush>, paints: &[RunPaint], (ox, oy): (f32, f32), out: &mut Vec<Segment>) {
    out.clear();
    for item in line.items() {
        let PositionedLayoutItem::GlyphRun(glyph_run) = item else { continue };
        let index = glyph_run.style().brush.0 as usize;
        let Some(paint) = paints.get(index) else { continue };
        let decoration = &paint.decoration;
        if decoration.underline.is_none() && decoration.strike.is_none() {
            continue;
        }
        let run = glyph_run.run();
        let metrics = run.font_metrics();
        let start = ox + glyph_run.offset();
        let segment = |strike: bool, start: f32, end: f32| Segment {
            paint: index,
            strike,
            start,
            end,
            baseline: oy + glyph_run.baseline() - paint.baseline_offset,
            face: if strike {
                Band {
                    offset: metrics.strikethrough_offset,
                    size: metrics.strikethrough_size,
                }
            } else {
                Band {
                    offset: metrics.underline_offset,
                    size: metrics.underline_size,
                }
            },
            font_size: run.font_size(),
        };
        if decoration.strike.is_some() {
            join_segment(out, segment(true, start, start + glyph_run.advance()));
        }
        if decoration.underline.is_some() {
            join_segment(out, segment(false, start, start + glyph_run.advance()));
        }
    }
}

/// Adds `segment`, or continues the last segment of its kind when that one is its run's and ends where it starts.
fn join_segment(out: &mut Vec<Segment>, segment: Segment) {
    match out.iter_mut().rev().find(|s| s.strike == segment.strike) {
        Some(last) if last.paint == segment.paint && (segment.start - last.end).abs() < SEGMENT_JOIN_PX => last.end = segment.end,
        _ => out.push(segment),
    }
}

/// Joins each `data-underline-line-band` underline segment of a line to the one before it when both have one shape and
/// one colour, sit on one baseline and touch, so a dash pattern or zigzag runs on across runs, as PowerPoint draws
/// them (a dash underline on a 24pt then a 12pt run of a synthetic deck is one dash pattern of their shared band;
/// a colour change starts a new one). Run after `line_band`, which weighs each run by its own segment.
pub(crate) fn join_line_band(segments: &mut Vec<Segment>, paints: &[RunPaint]) {
    let key = |segment: &Segment| {
        let paint = paints.get(segment.paint).filter(|p| !segment.strike && p.decoration.underline_line_band)?;
        let shape = paint.decoration.underline.as_ref()?;
        Some((&shape.pieces, paint.decoration.underline_color.unwrap_or(paint.color)))
    };
    let mut kept = 0;
    for index in 0..segments.len() {
        let Some(&segment) = segments.get(index) else { break };
        let previous = segments.get_mut(..kept).and_then(|done| done.iter_mut().rev().find(|s| !s.strike));
        match previous {
            Some(last)
                if key(&segment).is_some()
                    && key(last) == key(&segment)
                    && (last.baseline - segment.baseline).abs() < SEGMENT_JOIN_PX
                    && (segment.start - last.end).abs() < SEGMENT_JOIN_PX =>
            {
                last.end = segment.end;
            }
            _ => {
                if let Some(slot) = segments.get_mut(kept) {
                    *slot = segment;
                }
                kept += 1;
            }
        }
    }
    segments.truncate(kept);
}

/// The band of a line's `data-underline-line-band` underlines: their own bands averaged by their segments' widths (a
/// 12pt then a 24pt Arial run on the PowerPoint target: 1.44pt thick, its top 2.06pt below the baseline, where 12pt
/// alone is 0.84/1.27 and 24pt alone 1.80/2.42, measured against its PDF export of a synthetic deck). None when the
/// line has no such underline.
pub(crate) fn line_band(segments: &[Segment], paints: &[RunPaint]) -> Option<Band> {
    let (mut width, mut offset, mut size) = (0.0, 0.0, 0.0);
    for segment in segments.iter().filter(|s| !s.strike) {
        let decoration = paints.get(segment.paint).map(|paint| &paint.decoration);
        if let Some(shape) = decoration.filter(|d| d.underline_line_band).and_then(|d| d.underline.as_ref()) {
            let (advance, band) = (segment.end - segment.start, segment.band(shape));
            width += advance;
            offset += advance * band.offset;
            size += advance * band.size;
        }
    }
    (width > 0.0).then(|| Band {
        offset: offset / width,
        size: size / width,
    })
}

/// `piece` from `start` to `end` around the band `t` px thick from `top` (y down): the path, and the width it is stroked
/// with, or None when it is filled.
fn piece_path(piece: &DecorationPiece, start: f32, end: f32, top: f32, t: f32) -> (BezPath, Option<f32>) {
    match piece {
        DecorationPiece::Line { top: at, height, dashes } => {
            let (y, h) = (top + at * t, height * t);
            if dashes.is_empty() {
                return (rect_path(start, y, end - start, h, 0.0), None);
            }
            // The pattern starts again with each segment; the last dash is cut at its end.
            let mut path = BezPath::new();
            let mut x = start;
            for (i, length) in dashes.iter().cycle().enumerate() {
                if x >= end {
                    break;
                }
                let length = length * t;
                if i.is_multiple_of(2) {
                    path.extend(rect_path(x, y, length.min(end - x), h, 0.0));
                }
                x += length;
            }
            (path, None)
        }
        DecorationPiece::Wave {
            centre,
            amplitude,
            half,
            stroke,
        } => (wave(start, end, top + centre * t, amplitude * t, half * t), Some(*stroke)),
        DecorationPiece::Zigzag { top: at, height, line } => {
            let line = match *line {
                ZigzagLine::Px(px) => px,
                ZigzagLine::Thicknesses(n) => n * t,
            };
            (zigzag(start, top + at * t, end - start, height * t, line), None)
        }
    }
}

/// The dashes, half-waves or zigzag halves `piece_path` cuts a `width` px segment into around a band `t` px thick, as
/// the PDFKit renderer counts them; not finite for a width that is not.
fn piece_count(piece: &DecorationPiece, width: f32, t: f32) -> f64 {
    let width = f64::from(width);
    match piece {
        DecorationPiece::Line { dashes, .. } => {
            let period: f64 = dashes.iter().map(|&length| f64::from(length * t)).sum();
            if period > 0.0 {
                (width / period).ceil() * (dashes.len() / 2) as f64
            } else {
                1.0
            }
        }
        DecorationPiece::Wave { half, .. } => (width / f64::from(half * t)).ceil(),
        DecorationPiece::Zigzag { height, line, .. } => {
            let line = match *line {
                ZigzagLine::Px(px) => px,
                ZigzagLine::Thicknesses(n) => n * t,
            };
            2.0 * (width / f64::from((height * t - line).max(line))).ceil()
        }
    }
}

/// Cubic half-waves `half` long about `centre`, `amplitude` off it, the first up (toward the baseline) and the next
/// ones alternating, from `start` until they cover `end` (the last one runs past it): each half-wave's first control
/// point on its start, its second at its middle, as LibreOffice's PDF export draws them.
fn wave(start: f32, end: f32, centre: f32, amplitude: f32, half: f32) -> BezPath {
    // A finite length over a half-wave of at least MIN_DECORATION_LENGTH thicknesses: a whole count of at least 1.
    let halves = ((end - start) / half).ceil().max(1.0) as u32;
    let point = |x: f32, y: f32| (f64::from(x), f64::from(y));
    let mut path = BezPath::new();
    path.move_to(point(start, centre));
    for i in 0..halves {
        let x = start + half * i as f32;
        let peak = if i.is_multiple_of(2) { centre - amplitude } else { centre + amplitude };
        path.curve_to(point(x, centre), point(x + half / 2.0, peak), point(x + half, centre));
    }
    path
}

/// A zigzag `line` px tall filling `band` px from `top`, 45-degree slopes, from `start` down first, over whole
/// half-waves covering `advance`.
fn zigzag(start: f32, top: f32, advance: f32, band: f32, line: f32) -> BezPath {
    let reach = (band - line).max(line);
    // A finite advance over a positive reach: a whole count of at least 1.
    let halves = (advance / reach).ceil().max(1.0) as u32;
    let point = |i: u32, y: f32| (f64::from(start + reach * i as f32), f64::from(y));
    let edge = |i: u32| if i.is_multiple_of(2) { top } else { top + reach };
    let mut path = BezPath::new();
    path.move_to(point(0, top));
    for i in 0..=halves {
        path.line_to(point(i, edge(i) + line));
    }
    for i in (0..=halves).rev() {
        path.line_to(point(i, edge(i)));
    }
    path.close_path();
    path
}

/// Shears what is drawn on `baseline` (y down) by `shear` to the right above it and to the left below it.
fn oblique(baseline: f32, shear: f32) -> Transform {
    Transform::from_row(1.0, 0.0, -shear, 1.0, shear * baseline, 0.0)
}

/// A text link's clickable area on one line, and the run whose link it is.
pub(crate) struct LinkQuad {
    pub(crate) run: RunRef,
    pub(crate) quad: Quadrilateral,
}

/// `None` for a width pdf-writer cannot write (stroke_width).
fn stroke_of(Rgba(r, g, b, a): Rgba, width: f32) -> Option<Stroke> {
    Some(Stroke {
        paint: rgb::Color::new(r, g, b).into(),
        width: stroke_width(width)?,
        opacity: unit(a),
        ..Stroke::default()
    })
}

fn outline_stroke(outline: TextStroke) -> Option<Stroke> {
    Some(Stroke {
        miter_limit: outline.miter_limit,
        ..stroke_of(outline.color, outline.width)?
    })
}

impl Renderer {
    /// Draws the box and queues its runs' link annotations; a link the PDF cannot carry is an error.
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn draw_text(&mut self, surface: &mut Surface, text: &TextBox, layout: &BoxLayout, x: f32, y: f32, w: f32, h: f32) -> Result<(), String> {
        let [top, right, bottom, left] = text.margin;
        let (ix, iy, iw, ih) = (x + left, y + top, w - left - right, h - top - bottom);
        // Local frame: u along the line, v along the block axis.
        let (frame, block) = match text.direction {
            TextDirection::Vertical => (Transform::from_row(0.0, 1.0, -1.0, 0.0, ix + iw, iy), iw),
            TextDirection::Vertical270 => (Transform::from_row(0.0, -1.0, 1.0, 0.0, ix, iy + ih), iw),
            TextDirection::Horizontal | TextDirection::Stacked => (Transform::from_translate(ix, iy), ih),
        };
        // `.text-content` `transform: scaleX(-1)`: the frame mirrored about the text area's middle, x to 2cx - x.
        let frame = if text.mirrored {
            let cx = ix + iw / 2.0;
            Transform::from_row(-frame.sx(), frame.ky(), -frame.kx(), frame.sy(), 2.0 * cx - frame.tx(), frame.ty())
        } else {
            frame
        };
        let v0 = match text.vertical_align {
            VerticalAlign::Top => 0.0,
            VerticalAlign::Middle => (block - layout.height) / 2.0,
            VerticalAlign::Bottom => block - layout.height,
        };
        let mut lines = |r: &mut Self, surface: &mut Surface| {
            surface.push_transform(&frame);
            for para in &layout.paras {
                if let Some(BulletMarker { text, layout: marker, x }) = &para.marker {
                    let first = layout.baselines.get(para.baselines.start..para.baselines.start + 1).unwrap_or_default();
                    r.draw_layout(surface, marker, text, &layout.paints, (*x, v0 + para.y), first);
                }
                let baselines = layout.baselines.get(para.baselines.clone()).unwrap_or_default();
                r.draw_layout(surface, &para.layout, &para.text, &layout.paints, (para.x, v0 + para.y), baselines);
            }
            surface.pop();
            Ok(())
        };
        // Each distinct glow of the box goes under all of its text, as LibreOffice draws it.
        let mut glows: Vec<Effect> = Vec::new();
        for glow in layout.paints.iter().filter_map(|p| p.glow) {
            if !glows.contains(&glow) {
                glows.push(glow);
            }
        }
        // Each glow nests one more sub-stream (and call frame) and holds its own mask, so the
        // count is bounded: a box with thousands of distinct glows overflowed the wasm stack.
        // shortcut: glows past the limit are left out, not merged; capture them in one pass when
        // a real deck needs more distinct glows in one box.
        if glows.len() > MAX_GLOWS_PER_BOX {
            glows.truncate(MAX_GLOWS_PER_BOX);
            self.warn(format!(
                "text glows past the first {MAX_GLOWS_PER_BOX} distinct ones in a text box not rendered"
            ));
        }
        self.under_glows(surface, &glows, &mut lines)?;
        let mut quads = std::mem::take(&mut self.link_quads);
        let linked = quads.drain(..).try_for_each(|LinkQuad { run, quad }| {
            let run = text.paragraphs.get(run.paragraph as usize).and_then(|para| para.runs.get(run.run as usize));
            match run.and_then(|run| run.link.as_ref()) {
                Some(link) => self.add_link(link, quad),
                None => Ok(()),
            }
        });
        self.link_quads = quads;
        linked
    }

    fn under_glows(
        &mut self,
        surface: &mut Surface,
        glows: &[Effect],
        draw: &mut dyn FnMut(&mut Self, &mut Surface) -> Result<(), String>,
    ) -> Result<(), String> {
        match glows.split_first() {
            Some((glow, rest)) => self.under_effect(surface, glow, EffectSource::GlowingRuns, |r, surface| r.under_glows(surface, rest, draw)),
            None => draw(self, surface),
        }
    }

    fn pdf_font(&mut self, font: &parley::FontData) -> Option<Font> {
        let key = (font.data.id(), font.index);
        if let Some(f) = self.pdf_fonts.get(&key) {
            return Some(f.clone());
        }
        let bytes: Arc<dyn AsRef<[u8]> + Send + Sync> = Arc::new(font.data.clone());
        let f = Font::new(Data::from(bytes), font.index)?;
        self.pdf_fonts.insert(key, f.clone());
        Some(f)
    }

    /// Draws `layout` at `(ox, oy)` with line `i`'s baseline `baselines[i]` below `oy` (from
    /// `stack_lines`, where LibreOffice puts it), instead of where parley's font metrics would.
    pub(crate) fn draw_layout(
        &mut self,
        surface: &mut Surface,
        layout: &Layout<Brush>,
        text: &str,
        paints: &[RunPaint],
        (ox, oy): (f32, f32),
        baselines: &[f32],
    ) {
        let mut segments = std::mem::take(&mut self.decoration_segments);
        for (line, baseline) in layout.lines().zip(baselines) {
            let oy = oy + baseline - line.metrics().baseline;
            for item in line.items() {
                let PositionedLayoutItem::GlyphRun(glyph_run) = item else { continue };
                #[allow(clippy::expect_used)] // internal invariant, stated in the message
                let paint = paints
                    .get(glyph_run.style().brush.0 as usize)
                    .expect("layout_text sets every brush to the index of the paint it pushes alongside");
                self.draw_glyph_run(surface, &glyph_run, paint, text, (ox, oy));
            }
            // Decorations go along each run's whole part of the line, so a dash pattern or a wave does not restart where
            // parley splits the run.
            line_segments(&line, paints, (ox, oy), &mut segments);
            let band = line_band(&segments, paints);
            join_line_band(&mut segments, paints);
            for segment in &segments {
                self.draw_decoration(surface, paints, segment, band);
            }
        }
        self.decoration_segments = segments;
    }

    /// One segment's underline or strike: its pieces around its band (the line's `band` for a line-band underline), the
    /// underline in its colour, the strike in the run's.
    fn draw_decoration(&mut self, surface: &mut Surface, paints: &[RunPaint], segment: &Segment, band: Option<Band>) {
        let Some(paint) = paints.get(segment.paint) else { return };
        let decoration = &paint.decoration;
        let (shape, color, band) = match (segment.strike, &decoration.strike, &decoration.underline) {
            (true, Some(shape), _) => (shape, paint.color, segment.band(shape)),
            (false, _, Some(shape)) => {
                let own = segment.band(shape);
                let band = if decoration.underline_line_band { band.unwrap_or(own) } else { own };
                (shape, decoration.underline_color.unwrap_or(paint.color), band)
            }
            _ => return,
        };
        if segment.end <= segment.start {
            return;
        }
        let (top, t) = (segment.baseline - band.offset, band.size.max(MIN_DECORATION_PX));
        let count: f64 = shape.pieces.iter().map(|piece| piece_count(piece, segment.end - segment.start, t)).sum();
        // Not a finite count within the budget (a huge letter-spacing makes millions): reported, not drawn.
        let total = self.decoration_pieces as f64 + count;
        if !total.is_finite() || total > MAX_RENDER_DECORATION_PIECES as f64 {
            self.warn(format!(
                "underlines and strikes past {MAX_RENDER_DECORATION_PIECES} dashes, waves and zigzag steps per render are not drawn"
            ));
            return;
        }
        // A whole count within the budget, so the cast is exact.
        self.decoration_pieces += count as usize;
        for piece in &shape.pieces {
            match piece_path(piece, segment.start, segment.end, top, t) {
                (path, None) => self.fill(surface, &path, color),
                (path, Some(width)) => {
                    if let (Some(stroke), Some(pdf)) = (stroke_of(color, width), to_path(&path))
                        && self.placed(surface, &path, Some(&stroke))
                    {
                        self.effects.capture_stroke(surface.ctm(), &path, &stroke);
                        surface.set_fill(None);
                        surface.set_stroke(Some(stroke));
                        surface.draw_path(&pdf);
                        surface.set_stroke(None);
                    }
                }
            }
        }
    }

    /// One glyph run of a line whose top is `oy`: its link area and highlight, and its glyphs.
    fn draw_glyph_run(&mut self, surface: &mut Surface, glyph_run: &GlyphRun<'_, Brush>, paint: &RunPaint, text: &str, (ox, oy): (f32, f32)) {
        let run = glyph_run.run();
        let metrics = *run.font_metrics();
        let baseline = oy + glyph_run.baseline() - paint.baseline_offset;
        let start = ox + glyph_run.offset();
        let advance = glyph_run.advance();
        if let Some(link) = paint.link {
            // One annotation per line fragment, over the run's line box, as Chromium does for an inline <a>.
            if let Some(quad) = page_quad(surface, start, baseline - metrics.ascent, advance, metrics.ascent + metrics.descent) {
                self.link_quads.push(LinkQuad { run: link, quad });
            }
        }
        if let Some(hl) = paint.highlight {
            self.fill(
                surface,
                &rect_path(start, baseline - metrics.ascent, advance, metrics.ascent + metrics.descent, 0.0),
                hl,
            );
        }
        let Some(font) = self.pdf_font(run.font()) else {
            self.warn("font could not be embedded");
            return;
        };
        self.draw_glyphs(surface, glyph_run, paint, font, text, (ox, oy));
    }

    /// The run's glyphs in `font`: an outline under them, the fill (synthetic bold and italic
    /// as the run's `synthesis` says), their silhouette for a capture, an outline over them.
    fn draw_glyphs(&mut self, surface: &mut Surface, glyph_run: &GlyphRun<'_, Brush>, paint: &RunPaint, font: Font, text: &str, (ox, oy): (f32, f32)) {
        let run = glyph_run.run();
        let size = run.font_size();
        let baseline = oy + glyph_run.baseline() - paint.baseline_offset;
        let mut ranges = std::mem::take(&mut self.glyph_ranges);
        let mut positioned = std::mem::take(&mut self.positioned);
        let mut glyphs = std::mem::take(&mut self.glyphs);

        // Glyph -> source text cluster ranges, for the PDF ToUnicode map.
        ranges.clear();
        for cluster in run.visual_clusters() {
            if cluster.style_index() == glyph_run.style_index() {
                for _ in cluster.glyphs() {
                    ranges.push(cluster.text_range());
                }
            }
        }
        positioned.clear();
        positioned.extend(glyph_run.positioned_glyphs());
        if ranges.len() != positioned.len() {
            // shortcut: a run whose style repeats non-contiguously maps all glyphs to the run's text; fix if text extraction matters.
            ranges.clear();
            ranges.resize(positioned.len(), run.text_range());
        }
        let mut pen = glyph_run.offset();
        glyphs.clear();
        glyphs.extend(positioned.iter().zip(ranges.drain(..)).map(|(g, range)| {
            let glyph = KrillaGlyph::new(
                GlyphId::new(g.id),
                g.advance / size,
                (g.x - pen) / size,
                -(g.y - glyph_run.baseline()) / size,
                0.0,
                range,
                None,
            );
            pen += g.advance;
            glyph
        }));
        let origin = Point::from_xy(ox + glyph_run.offset(), baseline);
        // krilla unwraps the run's box: its pen positions grown by the font's box at this size.
        let pens = positioned.iter().flat_map(|g| {
            let y = f64::from(baseline + g.y - glyph_run.baseline());
            [(f64::from(ox + g.x), y), (f64::from(ox + g.x + g.advance), y)]
        });
        let reach = f64::from(size) * FONT_BOX_EM;
        let bbox = pens.fold(Rect::from_points((origin.x, origin.y), (origin.x, origin.y)), |b, p| b.union_pt(p));
        if !self.placed_box(surface, bbox.inflate(reach, reach), "text") {
            self.glyph_ranges = ranges;
            self.positioned = positioned;
            self.glyphs = glyphs;
            return;
        }
        // fontique asks for synthesis when the face it matched lacks the requested style.
        let synthesis = run.synthesis();
        let skewed = synthesis.skew().is_some();
        if skewed {
            surface.push_transform(&oblique(baseline, paint.synthesis.italic_shear));
        }
        if let Some(outline) = paint.outline.filter(|o| o.under) {
            surface.set_fill(None);
            surface.set_stroke(outline_stroke(outline));
            surface.draw_glyphs(origin, &glyphs, font.clone(), text, size, false);
        }
        surface.set_stroke(
            synthesis
                .embolden()
                .then(|| stroke_of(paint.color, size * paint.synthesis.bold_stroke_em))
                .flatten(),
        );
        surface.set_fill(Some(fill_of(paint.color)));
        surface.draw_glyphs(origin, &glyphs, font.clone(), text, size, false);
        if self.effects.capturing() {
            // Pen positions as krilla places them: x from the line start, y from the line's top.
            let at = positioned.iter().map(|g| (ox + g.x, oy + g.y - paint.baseline_offset, g.id));
            let outline = glyph_outlines(run.font(), size, at);
            self.effects.capture(surface.ctm(), &outline, paint.color.3, paint.glow.as_ref());
        }
        if let Some(outline) = paint.outline.filter(|o| !o.under) {
            surface.set_fill(None);
            surface.set_stroke(outline_stroke(outline));
            surface.draw_glyphs(origin, &glyphs, font, text, size, false);
        }
        surface.set_stroke(None);
        if skewed {
            surface.pop();
        }
        self.glyph_ranges = ranges;
        self.positioned = positioned;
        self.glyphs = glyphs;
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Decoration, Synthesis};
    use kurbo::Shape;

    #[test]
    fn obliques_lean_right_above_the_baseline_and_keep_it_in_place() {
        let t = oblique(100.0, 0.25);
        let map = |x: f32, y: f32| (t.sx() * x + t.kx() * y + t.tx(), t.ky() * x + t.sy() * y + t.ty());
        let close = |(x, y): (f32, f32), (ex, ey): (f32, f32)| assert!((x - ex).abs() < 1e-4 && (y - ey).abs() < 1e-4, "({x}, {y}) != ({ex}, {ey})");
        close(map(5.0, 100.0), (5.0, 100.0));
        close(map(5.0, 90.0), (7.5, 90.0));
        close(map(5.0, 110.0), (2.5, 110.0));
    }

    /// The path's subpaths as (min x, min y, max x, max y) boxes.
    fn pieces(path: &BezPath) -> Vec<(f64, f64, f64, f64)> {
        use kurbo::Shape as _;
        let mut boxes = Vec::new();
        let mut current = BezPath::new();
        for el in path.elements() {
            if matches!(el, kurbo::PathEl::MoveTo(_)) && !current.elements().is_empty() {
                let b = current.bounding_box();
                boxes.push((b.x0, b.y0, b.x1, b.y1));
                current = BezPath::new();
            }
            current.push(*el);
        }
        let b = current.bounding_box();
        boxes.push((b.x0, b.y0, b.x1, b.y1));
        boxes
    }

    fn line(top: f32, height: f32, dashes: &[f32]) -> DecorationPiece {
        DecorationPiece::Line {
            top,
            height,
            dashes: dashes.to_vec(),
        }
    }

    /// `piece` from 10 to 110 around a band 2px thick from 20, as filled boxes.
    fn filled(piece: &DecorationPiece) -> Vec<(f64, f64, f64, f64)> {
        let (path, stroke) = piece_path(piece, 10.0, 110.0, 20.0, 2.0);
        assert_eq!(stroke, None);
        pieces(&path)
    }

    #[test]
    fn lines_are_stripes_in_thicknesses_from_the_band_top() {
        assert_eq!(filled(&line(0.0, 1.0, &[])), [(10.0, 20.0, 110.0, 22.0)]);
        // Two half-thickness lines, one above the band and one below it.
        assert_eq!(filled(&line(-0.5, 0.5, &[])), [(10.0, 19.0, 110.0, 20.0)]);
        assert_eq!(filled(&line(1.5, 0.5, &[])), [(10.0, 23.0, 110.0, 24.0)]);
    }

    #[test]
    fn dash_patterns_cycle_from_the_segment_start_and_cut_the_last_dash() {
        // dotDash on a 1px band: dash 4, gap 2, dot 1, gap 2, then the dash again, cut at 12.
        let (path, _) = piece_path(&line(0.0, 1.0, &[4.0, 2.0, 1.0, 2.0]), 0.0, 12.0, 0.0, 1.0);
        assert_eq!(pieces(&path), [(0.0, 0.0, 4.0, 1.0), (6.0, 0.0, 7.0, 1.0), (9.0, 0.0, 12.0, 1.0)]);
    }

    #[test]
    fn waves_are_stroked_cubic_half_waves_starting_up_and_covering_the_segment() {
        let wave = DecorationPiece::Wave {
            centre: 0.5,
            amplitude: 2.0,
            half: 4.0,
            stroke: 0.25,
        };
        let (path, stroke) = piece_path(&wave, 0.0, 10.0, 20.0, 1.0);
        assert_eq!(stroke, Some(0.25));
        // Three half-waves cover 10px of 4px halves; the last runs past the end.
        let elements = path.elements();
        assert_eq!(elements.len(), 4);
        let curve = |i: usize| match elements.get(i) {
            Some(kurbo::PathEl::CurveTo(a, b, c)) => (*a, *b, *c),
            other => panic!("{other:?} is not a cubic"),
        };
        let (a, b, c) = curve(1);
        assert_eq!((a.x, a.y, b.x, b.y, c.x, c.y), (0.0, 20.5, 2.0, 18.5, 4.0, 20.5));
        let (_, b, c) = curve(2);
        assert_eq!((b.x, b.y, c.x), (6.0, 22.5, 8.0));
        assert_eq!(curve(3).2.x, 12.0);
    }

    #[test]
    fn zigzags_fill_their_height_with_whole_half_waves() {
        let zigzag = DecorationPiece::Zigzag {
            top: -0.5,
            height: 2.0,
            line: ZigzagLine::Px(0.25),
        };
        let [(x0, y0, x1, y1)] = filled(&zigzag)[..] else { panic!("one zigzag") };
        // From half a thickness above the band, 2 thicknesses tall; whole half-waves, so it may run past the text.
        let reach = 4.0 - 0.25;
        assert!((y0 - 19.0).abs() < 1e-4 && (y1 - 23.0).abs() < 1e-4, "{y0}..{y1}");
        assert!(x0 == 10.0 && (x1 - (10.0 + reach * (100.0_f64 / reach).ceil())).abs() < 1e-3, "{x0}..{x1}");
    }

    #[test]
    fn piece_counts_are_the_pieces_drawn_and_a_huge_width_passes_the_budget() {
        let wave = DecorationPiece::Wave {
            centre: 0.5,
            amplitude: 2.0,
            half: 4.0,
            stroke: 0.25,
        };
        let zigzag = DecorationPiece::Zigzag {
            top: -0.5,
            height: 2.0,
            line: ZigzagLine::Thicknesses(0.125),
        };
        // 100px around a 2px band: 6 dot-dash periods of 18px, 2 dashes each; the last period's dot falls past the end.
        let dot_dash = line(0.0, 1.0, &[4.0, 2.0, 1.0, 2.0]);
        assert_eq!(piece_count(&dot_dash, 100.0, 2.0), filled(&dot_dash).len() as f64 + 1.0);
        assert_eq!(piece_count(&line(0.0, 1.0, &[]), 100.0, 2.0), 1.0);
        assert_eq!(piece_count(&wave, 100.0, 2.0), 13.0);
        // Halves of 4 - 0.25 px over 100px, counted twice as the PDFKit renderer counts them.
        assert_eq!(piece_count(&zigzag, 100.0, 2.0), 2.0 * (100.0_f64 / 3.75).ceil());
        // A run 1e7 px wide under 4/2 dashes passes the per-render budget; an infinite one counts as not finite.
        assert!(piece_count(&line(0.0, 1.0, &[4.0, 2.0]), 1e7, 1.0) > MAX_RENDER_DECORATION_PIECES as f64);
        assert!(!piece_count(&wave, f32::INFINITY, 1.0).is_finite());
    }

    #[test]
    fn segments_join_a_runs_glyph_runs_but_not_other_runs_or_gaps() {
        let segment = |paint: usize, strike: bool, start: f32, end: f32| Segment {
            paint,
            strike,
            start,
            end,
            baseline: 0.0,
            face: Band { offset: 0.0, size: 1.0 },
            font_size: 10.0,
        };
        let mut out = Vec::new();
        join_segment(&mut out, segment(0, false, 0.0, 10.0));
        join_segment(&mut out, segment(0, true, 0.0, 10.0));
        // The underline continues past the strike that came between, then a gap (a space) and another run split it.
        join_segment(&mut out, segment(0, false, 10.0, 20.0));
        join_segment(&mut out, segment(0, false, 25.0, 30.0));
        join_segment(&mut out, segment(1, false, 30.0, 40.0));
        let spans: Vec<_> = out.iter().map(|s| (s.paint, s.strike, s.start, s.end)).collect();
        assert_eq!(
            spans,
            [(0, false, 0.0, 20.0), (0, true, 0.0, 10.0), (0, false, 25.0, 30.0), (1, false, 30.0, 40.0)]
        );
    }

    #[test]
    fn line_band_underlines_of_one_shape_and_colour_run_on_across_runs() {
        let dashed = DecorationShape {
            band: None,
            pieces: vec![DecorationPiece::Line {
                top: 0.0,
                height: 1.0,
                dashes: vec![3.0, 2.0],
            }],
        };
        let paint = |color: Rgba, shape: &DecorationShape, line_band: bool| RunPaint {
            color,
            highlight: None,
            outline: None,
            decoration: Decoration {
                underline: Some(shape.clone()),
                underline_line_band: line_band,
                ..Decoration::default()
            },
            baseline_offset: 0.0,
            synthesis: Synthesis::default(),
            link: None,
            glow: None,
        };
        let black = Rgba(0, 0, 0, 1.0);
        // 0 and 1 alike, 2 another colour, 3 like 2 but solid, 4 like 3 without the line band.
        let paints = [
            paint(black, &dashed, true),
            paint(black, &dashed, true),
            paint(Rgba(192, 0, 0, 1.0), &dashed, true),
            paint(Rgba(192, 0, 0, 1.0), &DecorationShape::solid(), true),
            paint(Rgba(192, 0, 0, 1.0), &DecorationShape::solid(), false),
            paint(Rgba(192, 0, 0, 1.0), &DecorationShape::solid(), false),
        ];
        let segment = |paint: usize, strike: bool, start: f32, end: f32| Segment {
            paint,
            strike,
            start,
            end,
            baseline: 0.0,
            face: Band { offset: 0.0, size: 1.0 },
            font_size: 10.0,
        };
        let mut segments = vec![
            segment(0, false, 0.0, 10.0),
            segment(0, true, 0.0, 10.0),
            segment(1, false, 10.0, 30.0),
            segment(2, false, 30.0, 40.0),
            segment(3, false, 40.0, 50.0),
            segment(4, false, 50.0, 60.0),
            segment(5, false, 60.0, 70.0),
        ];
        join_line_band(&mut segments, &paints);
        let spans: Vec<_> = segments.iter().map(|s| (s.paint, s.strike, s.start, s.end)).collect();
        assert_eq!(
            spans,
            [
                (0, false, 0.0, 30.0),
                (0, true, 0.0, 10.0),
                (2, false, 30.0, 40.0),
                (3, false, 40.0, 50.0),
                (4, false, 50.0, 60.0),
                (5, false, 60.0, 70.0)
            ]
        );
    }

    #[test]
    fn a_zigzag_line_in_thicknesses_scales_with_its_band() {
        let zigzag = |line: ZigzagLine| {
            piece_path(&DecorationPiece::Zigzag { top: 0.0, height: 2.0, line }, 0.0, 10.0, 0.0, 2.0)
                .0
                .bounding_box()
        };
        assert_eq!(zigzag(ZigzagLine::Thicknesses(0.5)), zigzag(ZigzagLine::Px(1.0)));
    }
}
