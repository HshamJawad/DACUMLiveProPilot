// ============================================================
// /pdf_arabic.js
// ------------------------------------------------------------
// Arabic support layer for the jsPDF exporter.
//
// WHY THIS FILE EXISTS
// --------------------
// exports_pdf.js contains ~145 pdf.text() calls, every one of them
// positioned from the LEFT edge with absolute millimetre coordinates.
// Making that file Arabic-aware call-by-call would mean touching all
// 145 sites, recomputing every x from the right margin, and keeping
// two parallel coordinate systems correct forever.
//
// Instead this module WRAPS THE DOCUMENT. installArabicRTL(pdf)
// replaces text/rect/line/addImage/getTextWidth/splitTextToSize on a
// single jsPDF instance with mirrored, shaped, bidi-ordered versions.
// The layout code above it keeps writing left-to-right coordinates and
// never learns that anything changed — the mirror is exact, so cell
// borders and the text inside them move together, and the four task
// columns reverse (A1 on the right) as a consequence of the geometry
// rather than as a special case.
//
// The mirror is INSTANCE-LEVEL, not prototype-level: nothing here
// leaks to another document, and the English and French paths never
// enter this file at all.
//
//
// THE TWO jsPDF PROCESSORS
// ------------------------
// jsPDF 2.5.1 runs TWO Arabic processors on text, not one, and both
// act after the string leaves our code:
//
//   1. processArabic — a built-in shaper on the `preProcessText`
//      event, called a SECOND time directly from getStringUnitWidth
//      while measuring. Text we have already shaped gets re-shaped
//      against its neighbours AFTER reversal, producing joined
//      nonsense.
//
//   2. bidiEngineFunction — a full BiDi engine on `postProcessText`.
//      It defaults to isInputVisual = true with isOutputVisual unset,
//      i.e. "this text is visual, hand it back logical" — so it
//      re-reverses our finished visual string and the whole line
//      comes out backwards.
//
// Disabling the first alone is not enough. This module handles both:
// the shaper is suspended (suspendJsPdfArabicParser), and the BiDi
// engine is neutralised through its own documented options rather
// than detached — `utf8EscapeFunction` shares the postProcessText
// topic and is required for encoding, and the minified CDN build has
// mangled function names, so the two cannot be told apart safely.
// ============================================================

import {
    loadArabicFont,
    isArabicFontLoaded,
    getArabicFontName,
    shapeArabic,
    bidiVisual
} from './arabic-font.js';


/* Re-exported so exports_pdf.js has one import source for the whole
   Arabic story rather than two. */
export { loadArabicFont, isArabicFontLoaded, getArabicFontName };


// ─────────────────────────────────────────────────────────────
//  suspendJsPdfArabicParser(pdf, jsPDFClass)
//  Silences processArabic for the lifetime of one export.
//  Returns a restore function — ALWAYS call it in a `finally`.
// ─────────────────────────────────────────────────────────────
export function suspendJsPdfArabicParser(pdf, jsPDFClass) {
    const restores = [];

    // 1. The preProcessText subscription on this document instance.
    try {
        const events = pdf.internal && pdf.internal.events;
        const topics = events && typeof events.getTopics === 'function'
            ? events.getTopics()
            : null;

        if (topics && topics.preProcessText) {
            Object.keys(topics.preProcessText).forEach(token => {
                const fn = topics.preProcessText[token][0];
                if (fn === jsPDFClass.API.processArabic ||
                    (jsPDFClass.API.__arabicParser__ &&
                     fn === jsPDFClass.API.__arabicParser__.processArabic)) {
                    events.unsubscribe(token);
                    restores.push(() => events.subscribe('preProcessText', fn));
                }
            });
        }
    } catch (e) {
        console.warn('[PDF] preProcessText detach failed.', e);
    }

    // 2. The direct call from inside getStringUnitWidth, which the
    //    event system knows nothing about. This is the one that made
    //    the first attempt at this fix look like it had failed.
    try {
        const original = jsPDFClass.API.processArabic;
        if (original) {
            jsPDFClass.API.processArabic = undefined;
            restores.push(() => { jsPDFClass.API.processArabic = original; });
        }
    } catch (e) {
        console.warn('[PDF] processArabic suspend failed.', e);
    }

    return () => restores.forEach(fn => { try { fn(); } catch (_) {} });
}


// ─────────────────────────────────────────────────────────────
//  BiDi engine neutraliser
//
//  Visual in, visual out, same direction on both sides. This
//  combination matches no transform branch in doBidiReorder(), so
//  the string is returned untouched. These are documented library
//  options, not a hack.
//
//  A FRESH object every call: jsPDF writes into the options object
//  it is handed, so a shared one would be poisoned after the first
//  line of the document.
// ─────────────────────────────────────────────────────────────
function _bidiOff(extra) {
    return Object.assign({
        isInputVisual:  true,
        isOutputVisual: true,
        isInputRtl:     false,
        isOutputRtl:    false
    }, extra || {});
}


// ─────────────────────────────────────────────────────────────
//  installArabicRTL(pdf)
//
//  Mirrors one jsPDF instance about the vertical centre line and
//  routes every string through the shaper + bidi reorderer.
//  Returns a restore function (used by the tests; the export path
//  discards the document afterwards anyway).
// ─────────────────────────────────────────────────────────────
export function installArabicRTL(pdf) {
    /* The mirror width is captured ONCE, from the first page, and is
       deliberately NOT re-read per page.

       exports_pdf.js computes pageWidth a single time from the opening
       landscape page and uses that value for its portrait pages too.
       Mirroring against a live per-page width would therefore disagree
       with the layout arithmetic exactly where that arithmetic is
       already wrong, putting text and its cell borders on different
       sides of the sheet. A fixed W is the exact inverse of what the
       layout code actually does, so the two stay in agreement. */
    const W = pdf.internal.pageSize.getWidth();

    const family = getArabicFontName();

    const orig = {
        text:            pdf.text.bind(pdf),
        splitTextToSize: pdf.splitTextToSize.bind(pdf),
        rect:            pdf.rect.bind(pdf),
        line:            pdf.line.bind(pdf),
        addImage:        pdf.addImage.bind(pdf),
        getTextWidth:    pdf.getTextWidth.bind(pdf),
        setFont:         pdf.setFont.bind(pdf)
    };

    /* Our code does the reversing. setR2L(true) would reverse it a
       second time and hand back the logical order. */
    pdf.setR2L(false);

    /* Marker for the few drawings that are DIRECTIONAL rather than
       positional — a checkmark reads wrong when reflected, unlike a
       cell border. See drawTick(). */
    pdf.__rtlMirrored = true;

    // ── Font ──────────────────────────────────────────────────
    /* The layout calls setFont(undefined, 'bold') in 60-odd places.
       With `undefined` jsPDF keeps the CURRENT family, so after the
       Arabic face is selected those calls ask for "Cairo bold". Only
       a regular weight is embedded; arabic-font.js registers the same
       face under both styles so the lookup cannot fail, and italic —
       which is not registered at all — is folded onto normal here
       rather than being allowed to fall back to Helvetica and drop
       every Arabic glyph on the line. */
    if (family) {
        pdf.setFont = (name, style) => {
            const s = (style === 'bold' || style === 'bolditalic') ? 'bold' : 'normal';
            return orig.setFont(family, s);
        };
        pdf.setFont(family, 'normal');
    }

    // ── Measurement ───────────────────────────────────────────
    /* Presentation forms are narrower than the base letters they
       replace. Measuring the unshaped string overestimates every
       width, which shows up as premature wrapping and short lines. */
    pdf.getTextWidth = (text) => orig.getTextWidth(shapeArabic(String(text ?? '')));

    // ── Wrapping ──────────────────────────────────────────────
    /* Shape BEFORE splitting, and return the lines still in logical
       order — pdf.text() reverses each one at draw time. Reversing
       here instead would let jsPDF wrap a visual string, breaking it
       at the wrong end and scattering words between lines.

       shapeArabic() is idempotent: presentation forms are not keys in
       its table, so a line that comes back through pdf.text() is not
       shaped twice. */
    pdf.splitTextToSize = (text, maxWidth, options) =>
        orig.splitTextToSize(shapeArabic(String(text ?? '')), maxWidth, options);

    // ── Drawing text ──────────────────────────────────────────
    pdf.text = (text, x, y, options, ...rest) => {
        const lines = Array.isArray(text)
            ? text
            : String(text ?? '').split('\n');

        const visual = lines.map(l => bidiVisual(shapeArabic(String(l ?? ''))));

        /* Mirror the anchor and flip the alignment with it. A string
           drawn left-aligned at x now hangs right-aligned from W - x,
           so it occupies the mirror image of its old box — including
           when x was derived from a cell edge plus padding. */
        const given = (options && options.align) || 'left';
        const align = given === 'center' ? 'center'
                    : given === 'right'  ? 'left'
                    : 'right';

        return orig.text(visual, W - x, y, _bidiOff(Object.assign({}, options, { align })), ...rest);
    };

    // ── Drawing geometry ──────────────────────────────────────
    /* Borders must mirror too, or the text lands outside its cell.
       Mirroring a rect means mirroring its RIGHT edge to become the
       new left edge: x → W - x - w. */
    pdf.rect = (x, y, w, h, style) => orig.rect(W - x - w, y, w, h, style);

    pdf.line = (x1, y1, x2, y2, style) => orig.line(W - x1, y1, W - x2, y2, style);

    pdf.addImage = (img, format, x, y, w, h, ...rest) =>
        orig.addImage(img, format, W - x - w, y, w, h, ...rest);

    return () => Object.assign(pdf, orig);
}


// ─────────────────────────────────────────────────────────────
//  ensureArabicFont(jsPDFClass)
//
//  Warms the module-level font cache in arabic-font.js using a
//  throwaway document.
//
//  This exists so exportToPDF() can stay SYNCHRONOUS. Font loading is
//  a fetch, and making the two exported entry points async would
//  change their contract for every caller in the app for the sake of
//  one network round trip that happens once per session. The export
//  path instead bails out, warms the cache, and re-enters itself.
// ─────────────────────────────────────────────────────────────
export async function ensureArabicFont(jsPDFClass) {
    if (isArabicFontLoaded()) return getArabicFontName();
    const probe = new jsPDFClass({ unit: 'mm', format: 'a4' });
    return loadArabicFont(probe);
}


// ─────────────────────────────────────────────────────────────
//  drawTick(pdf, cx, cy)
//
//  A checkmark drawn as two strokes instead of the character U+2713.
//
//  The matrix used pdf.text('✓'). jsPDF's standard-14 Helvetica is
//  Latin-1 only and Cairo omits the dingbat range, so that glyph is
//  missing in BOTH scripts — silently, because jsPDF drops unmapped
//  characters without complaint. Two lines are always drawable and
//  fix the English and French exports at the same time.
//
//  The mirror would reflect it into a backwards tick, so the x
//  offsets are pre-negated under RTL: the reflection then lands it
//  the right way round, still inside its mirrored cell. Cell borders
//  need no such treatment — a reflected rectangle is the same
//  rectangle — which is why this is the only exception in the file.
// ─────────────────────────────────────────────────────────────
export function drawTick(pdf, cx, cy) {
    const d = pdf.__rtlMirrored ? -1 : 1;
    const prev = pdf.getLineWidth ? pdf.getLineWidth() : 0.2;
    pdf.setLineWidth(0.5);
    pdf.line(cx - 1.5 * d, cy + 0.2, cx - 0.4 * d, cy + 1.4);
    pdf.line(cx - 0.4 * d, cy + 1.4, cx + 1.7 * d, cy - 1.6);
    pdf.setLineWidth(prev);
}
