// The shrink rule of a `data-text-fit="shrink"` box (docs/HTML_PROFILE.md "Text"). Each step lays the whole box out
// again, so a reader bounds the steps a caller's HTML may ask for: 100 000 steps of 20 KB of text took minutes (#448).

/**
 * Most shrink steps every renderer takes (pdf-core's copy in crates/pdf-core/src/model.rs); `toHtml()` writes exactly
 * this many (PPTX_DEFAULTS.text.shrink.maxSteps).
 */
export const MAX_TEXT_FIT_STEPS = 20;
