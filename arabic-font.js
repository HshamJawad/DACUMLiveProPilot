// ============================================================
// arabic-font.js — Arabic Font Loader + Shaping Engine for jsPDF
// DACUM Lite
//
// WHY THIS FILE EXISTS
// ────────────────────
// jsPDF writes glyphs in the exact code-point order it is given.
// It performs NO Arabic letter joining (shaping) and NO bidi
// reordering, and `pdf.setR2L(true)` only reverses the run — it
// does not join letters. PDF viewers do not apply the font's
// GSUB tables either, because the PDF already carries a final
// glyph stream. That is why Arabic came out as disconnected,
// backwards letters.
//
// The fix is to do both jobs in JavaScript before the string ever
// reaches pdf.text():
//
//   1. shapeArabic(str)  — maps every Arabic letter to its correct
//      contextual presentation form (U+FE70–U+FEFF / U+FB50–U+FBFF)
//      and builds the four lam-alef ligatures.
//   2. bidiVisual(str)   — reorders the shaped string from logical
//      order into visual order (RTL base direction), keeping Latin
//      words and numbers running left-to-right, and mirroring
//      brackets.
//   3. arabicVisual(str) — the two steps combined. This is what
//      events.js calls.
//
// GLYPH-COVERAGE FALLBACK — the second half of the bug
// ─────────────────────────────────────────────────────
// Shaping alone is not enough. Cairo (and most modern Arabic
// webfonts) deliberately omit the ISOLATED presentation forms —
// U+FE8D alef, U+FE8F beh, U+FEDD lam and 30-odd others simply do
// not exist in the font's cmap, because the isolated shape is
// already the base code point. jsPDF silently drops any character
// it cannot map, which is why whole letters were vanishing from
// the Arabic sheet while the joined ones survived.
//
// So the loader now reads the cmap out of the TTF it just fetched
// and records exactly which code points the font can draw.
// shapeArabic() consults that set and degrades gracefully:
//   medial  → final   → base
//   initial → isolated → base
//   final   → base
//   isolated → base
// The base code point is always the isolated shape, so nothing is
// lost visually and no character can ever disappear again.
//
// With this in place `pdf.setR2L()` MUST stay false, otherwise the
// text is reversed twice.
//
// FONT FILES — PRIORITY ORDER
// ────────────────────────────
// jsPDF needs a TTF (not woff2). Place ONE of these next to
// index.html, or inside a ./fonts/ folder:
//
//   1. Cairo-Regular.ttf   ← current choice. Cairo omits the
//      isolated presentation forms, but it does carry the isolated
//      shape on the base code point, and the coverage fallback
//      below routes to it, so nothing is lost.
//   2. fonts/Amiri-Regular.ttf   ← the safety net. Amiri carries
//      the full U+FE70–U+FEFF range, isolated forms included, so
//      the fallback never has to fire.
//   3. fonts/Tajawal-Regular.ttf
//
// The loader tries every path in order and caches the first hit,
// so there is exactly one network request per session.
// ============================================================

// ── Font candidates (tried in order) ─────────────────────────
const FONT_CANDIDATES = [
    { file: './Cairo-Regular.ttf',         name: 'Cairo'   },
    { file: './fonts/Cairo-Regular.ttf',   name: 'Cairo'   },
    { file: './fonts/Amiri-Regular.ttf',   name: 'Amiri'   },
    { file: './Amiri-Regular.ttf',         name: 'Amiri'   },
    { file: './fonts/Tajawal-Regular.ttf', name: 'Tajawal' },
    { file: './Tajawal-Regular.ttf',       name: 'Tajawal' },
];

// ── Module-level cache ────────────────────────────────────────
let _cachedFont = null;  // { name, b64, coverage } | null
let _coverage   = null;  // Set<number> of drawable code points | null

// ─────────────────────────────────────────────────────────────
//  loadArabicFont(pdf)
//  Fetch a candidate font, convert to base64, register with the
//  jsPDF instance. Returns the font-family name or null.
// ─────────────────────────────────────────────────────────────
export async function loadArabicFont(pdf) {
    if (_cachedFont) {
        _registerFont(pdf, _cachedFont.name, _cachedFont.b64);
        _coverage = _cachedFont.coverage;
        return _cachedFont.name;
    }

    for (const candidate of FONT_CANDIDATES) {
        try {
            const { b64, coverage } = await _fetchFont(candidate.file);
            _cachedFont = { name: candidate.name, b64, coverage };
            _coverage   = coverage;
            _registerFont(pdf, candidate.name, b64);
            console.info(`[ArabicFont] Loaded "${candidate.file}" → family "${candidate.name}"`);
            return candidate.name;
        } catch {
            /* try the next path */
        }
    }

    console.error(
        '[ArabicFont] No Arabic TTF found. Place one of:\n' +
        FONT_CANDIDATES.map(c => '  • ' + c.file).join('\n')
    );
    return null;
}

export function isArabicFontLoaded() { return _cachedFont !== null; }
export function getArabicFontName()  { return _cachedFont ? _cachedFont.name : null; }

/**
 * Override the glyph-coverage set used by the shaper. Only needed
 * for tests — loadArabicFont() sets this automatically.
 * @param {Set<number>|null} set
 */
export function setFontCoverage(set) { _coverage = set || null; }

// ── Private: register font with jsPDF ────────────────────────
function _registerFont(pdf, name, b64) {
    const vfsName = name + '-Regular.ttf';
    pdf.addFileToVFS(vfsName, b64);
    pdf.addFont(vfsName, name, 'normal');
    pdf.addFont(vfsName, name, 'bold');   // same face — keeps setFont(name,'bold') safe
}

// ── Private: fetch font file → base64 + cmap coverage ────────
async function _fetchFont(filename) {
    const res = await fetch(filename, { cache: 'force-cache' });
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${filename}`);

    const buffer = await res.arrayBuffer();
    const bytes  = new Uint8Array(buffer);

    let binary  = '';
    const chunk = 8192;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }

    let coverage = null;
    try {
        coverage = parseFontCoverage(buffer);
    } catch (e) {
        console.warn('[ArabicFont] cmap unreadable — falling back to base forms only.', e);
    }

    return { b64: btoa(binary), coverage };
}

// ══════════════════════════════════════════════════════════════
//  cmap READER
//  Returns a Set of every code point the font can actually draw,
//  or null when the table cannot be read (in which case the
//  shaper assumes nothing and stays on the safe base forms).
// ══════════════════════════════════════════════════════════════
export function parseFontCoverage(buffer) {
    const dv = new DataView(buffer);
    if (buffer.byteLength < 12) return null;

    // ── Locate the cmap table in the sfnt directory ───────────
    const numTables = dv.getUint16(4);
    let cmapOff = 0;
    for (let i = 0; i < numTables; i++) {
        const rec = 12 + i * 16;
        if (rec + 16 > buffer.byteLength) break;
        const tag = String.fromCharCode(
            dv.getUint8(rec), dv.getUint8(rec + 1),
            dv.getUint8(rec + 2), dv.getUint8(rec + 3)
        );
        if (tag === 'cmap') { cmapOff = dv.getUint32(rec + 8); break; }
    }
    if (!cmapOff || cmapOff + 4 > buffer.byteLength) return null;

    // ── Pick the best subtable: (3,10) > (3,1) > (0,x) ───────
    const nSub = dv.getUint16(cmapOff + 2);
    let best = 0, bestScore = -1;
    for (let i = 0; i < nSub; i++) {
        const rec = cmapOff + 4 + i * 8;
        if (rec + 8 > buffer.byteLength) break;
        const pid = dv.getUint16(rec);
        const eid = dv.getUint16(rec + 2);
        const off = dv.getUint32(rec + 4);
        let score = -1;
        if (pid === 3 && eid === 10)     score = 4;
        else if (pid === 3 && eid === 1) score = 3;
        else if (pid === 0)              score = 2;
        if (score > bestScore) { bestScore = score; best = cmapOff + off; }
    }
    if (!best || best + 4 > buffer.byteLength) return null;

    const set = new Set();
    const fmt = dv.getUint16(best);

    if (fmt === 4) {
        const segX2  = dv.getUint16(best + 6);
        const seg    = segX2 / 2;
        const endO   = best + 14;
        const startO = endO + segX2 + 2;
        const deltaO = startO + segX2;
        const rangeO = deltaO + segX2;
        for (let s = 0; s < seg; s++) {
            const end   = dv.getUint16(endO + s * 2);
            const start = dv.getUint16(startO + s * 2);
            if (start > end || start === 0xFFFF) continue;
            const delta = dv.getInt16(deltaO + s * 2);
            const ro    = dv.getUint16(rangeO + s * 2);
            for (let c = start; c <= end; c++) {
                let g;
                if (ro === 0) {
                    g = (c + delta) & 0xFFFF;
                } else {
                    const gi = rangeO + s * 2 + ro + (c - start) * 2;
                    if (gi + 2 > buffer.byteLength) continue;
                    g = dv.getUint16(gi);
                    if (g !== 0) g = (g + delta) & 0xFFFF;
                }
                if (g) set.add(c);
            }
        }
    } else if (fmt === 12) {
        const nGroups = dv.getUint32(best + 12);
        for (let g = 0; g < nGroups; g++) {
            const o = best + 16 + g * 12;
            if (o + 12 > buffer.byteLength) break;
            const s = dv.getUint32(o);
            const e = dv.getUint32(o + 4);
            for (let c = s; c <= e && c - s < 0x10000; c++) set.add(c);
        }
    } else if (fmt === 6) {
        const first = dv.getUint16(best + 6);
        const count = dv.getUint16(best + 8);
        for (let i = 0; i < count; i++) {
            if (dv.getUint16(best + 10 + i * 2)) set.add(first + i);
        }
    } else {
        return null;
    }

    return set.size ? set : null;
}

// ══════════════════════════════════════════════════════════════
//  PART 2 — ARABIC SHAPING ENGINE
// ══════════════════════════════════════════════════════════════

/* Contextual forms table.
   Key   = base code point
   Value = [ isolated, final, initial, medial ]
   A two-entry array marks a right-joining letter — it never
   connects to the letter that follows it. */
const FORMS = {
    0x0621: [0xFE80],
    0x0622: [0xFE81, 0xFE82],
    0x0623: [0xFE83, 0xFE84],
    0x0624: [0xFE85, 0xFE86],
    0x0625: [0xFE87, 0xFE88],
    0x0626: [0xFE89, 0xFE8A, 0xFE8B, 0xFE8C],
    0x0627: [0xFE8D, 0xFE8E],
    0x0628: [0xFE8F, 0xFE90, 0xFE91, 0xFE92],
    0x0629: [0xFE93, 0xFE94],
    0x062A: [0xFE95, 0xFE96, 0xFE97, 0xFE98],
    0x062B: [0xFE99, 0xFE9A, 0xFE9B, 0xFE9C],
    0x062C: [0xFE9D, 0xFE9E, 0xFE9F, 0xFEA0],
    0x062D: [0xFEA1, 0xFEA2, 0xFEA3, 0xFEA4],
    0x062E: [0xFEA5, 0xFEA6, 0xFEA7, 0xFEA8],
    0x062F: [0xFEA9, 0xFEAA],
    0x0630: [0xFEAB, 0xFEAC],
    0x0631: [0xFEAD, 0xFEAE],
    0x0632: [0xFEAF, 0xFEB0],
    0x0633: [0xFEB1, 0xFEB2, 0xFEB3, 0xFEB4],
    0x0634: [0xFEB5, 0xFEB6, 0xFEB7, 0xFEB8],
    0x0635: [0xFEB9, 0xFEBA, 0xFEBB, 0xFEBC],
    0x0636: [0xFEBD, 0xFEBE, 0xFEBF, 0xFEC0],
    0x0637: [0xFEC1, 0xFEC2, 0xFEC3, 0xFEC4],
    0x0638: [0xFEC5, 0xFEC6, 0xFEC7, 0xFEC8],
    0x0639: [0xFEC9, 0xFECA, 0xFECB, 0xFECC],
    0x063A: [0xFECD, 0xFECE, 0xFECF, 0xFED0],
    0x0640: [0x0640, 0x0640, 0x0640, 0x0640],   // tatweel
    0x0641: [0xFED1, 0xFED2, 0xFED3, 0xFED4],
    0x0642: [0xFED5, 0xFED6, 0xFED7, 0xFED8],
    0x0643: [0xFED9, 0xFEDA, 0xFEDB, 0xFEDC],
    0x0644: [0xFEDD, 0xFEDE, 0xFEDF, 0xFEE0],
    0x0645: [0xFEE1, 0xFEE2, 0xFEE3, 0xFEE4],
    0x0646: [0xFEE5, 0xFEE6, 0xFEE7, 0xFEE8],
    0x0647: [0xFEE9, 0xFEEA, 0xFEEB, 0xFEEC],
    0x0648: [0xFEED, 0xFEEE],
    0x0649: [0xFEEF, 0xFEF0],
    0x064A: [0xFEF1, 0xFEF2, 0xFEF3, 0xFEF4],
    0x0671: [0xFB50, 0xFB51],
    0x0679: [0xFB66, 0xFB67, 0xFB68, 0xFB69],
    0x067E: [0xFB56, 0xFB57, 0xFB58, 0xFB59],
    0x0686: [0xFB7A, 0xFB7B, 0xFB7C, 0xFB7D],
    0x0688: [0xFB88, 0xFB89],
    0x0691: [0xFB8C, 0xFB8D],
    0x0698: [0xFB8A, 0xFB8B],
    0x06A9: [0xFB8E, 0xFB8F, 0xFB90, 0xFB91],
    0x06AF: [0xFB92, 0xFB93, 0xFB94, 0xFB95],
    0x06BA: [0xFB9E, 0xFB9F],
    0x06BE: [0xFBAA, 0xFBAB, 0xFBAC, 0xFBAD],
    0x06C1: [0xFBA6, 0xFBA7, 0xFBA8, 0xFBA9],
    0x06CC: [0xFBFC, 0xFBFD, 0xFBFE, 0xFBFF],
    0x06D2: [0xFBAE, 0xFBAF],
};

/* Lam + Alef → one ligature glyph. [isolated, final] */
const LAM_ALEF = {
    0x0622: [0xFEF5, 0xFEF6],
    0x0623: [0xFEF7, 0xFEF8],
    0x0625: [0xFEF9, 0xFEFA],
    0x0627: [0xFEFB, 0xFEFC],
};

/* Harakat and other combining marks — transparent for joining. */
function _isMark(cp) {
    return (cp >= 0x064B && cp <= 0x065F) ||
           (cp >= 0x0610 && cp <= 0x061A) ||
           (cp >= 0x06D6 && cp <= 0x06ED) ||
            cp === 0x0670;
}

const _canStart = (cp) => !!FORMS[cp] && FORMS[cp].length === 4;  // joins to the next letter
const _canEnd   = (cp) => !!FORMS[cp];                            // joins to the previous letter

/* Can the loaded font actually draw this code point? When no cmap
   was read, assume nothing and let _pick() fall through to the
   base letter, which every Arabic font carries. */
const _has = (cp) => _coverage ? _coverage.has(cp) : false;

/* Return the first candidate the font can draw, otherwise the base
   code point — whose glyph IS the isolated shape. Nothing is ever
   dropped. */
function _pick(candidates, base) {
    for (const c of candidates) if (_has(c)) return c;
    return base;
}

/**
 * Convert Arabic text into contextual presentation forms.
 * Logical order is preserved — call bidiVisual() afterwards.
 *
 * @param {string} input
 * @returns {string}
 */
export function shapeArabic(input) {
    const src = String(input ?? '');
    if (!src) return '';

    const cps = Array.from(src).map(c => c.codePointAt(0));
    const out = [];

    for (let i = 0; i < cps.length; i++) {
        const cp = cps[i];

        if (_isMark(cp) || !FORMS[cp]) { out.push(cp); continue; }

        // Previous non-mark letter — does it join forward?
        let p = i - 1;
        while (p >= 0 && _isMark(cps[p])) p--;
        const prevJoins = p >= 0 && _canStart(cps[p]);

        // Next non-mark letter — does it join backward?
        let n = i + 1;
        while (n < cps.length && _isMark(cps[n])) n++;
        const nextCp    = n < cps.length ? cps[n] : -1;
        const nextJoins = nextCp !== -1 && _canEnd(nextCp);

        const f  = FORMS[cp];
        const fw = _canStart(cp) && nextJoins;   // joins forward
        const bw = prevJoins;                    // joins backward

        // ── Lam-Alef ligature ─────────────────────────────────
        // Falls back to two separate letters when the font has no
        // ligature glyph, rather than dropping the pair.
        if (cp === 0x0644 && LAM_ALEF[nextCp]) {
            const lig  = LAM_ALEF[nextCp];
            const want = bw ? lig[1] : lig[0];
            if (_has(want)) {
                out.push(want);
            } else {
                out.push(_pick([bw ? f[1] : f[0]], cp));           // lam
                out.push(_pick([FORMS[nextCp][1]], nextCp));       // alef, final
            }
            i = n;                     // consume the alef
            continue;
        }

        // ── Standard contextual form, with graceful degradation ─
        if (f.length === 4) {
            if (bw && fw)  out.push(_pick([f[3], f[1]], cp));      // medial  → final  → base
            else if (bw)   out.push(_pick([f[1]], cp));            // final   → base
            else if (fw)   out.push(_pick([f[2], f[0]], cp));      // initial → isolated → base
            else           out.push(_pick([f[0]], cp));            // isolated → base
        } else {
            out.push(bw ? _pick([f[1]], cp) : _pick([f[0]], cp));
        }
    }

    return out.map(c => String.fromCodePoint(c)).join('');
}

// ══════════════════════════════════════════════════════════════
//  PART 3 — BIDI REORDERING (RTL base direction)
// ══════════════════════════════════════════════════════════════

const _MIRROR = {
    '(': ')', ')': '(', '[': ']', ']': '[', '{': '}', '}': '{',
    '<': '>', '>': '<', '«': '»', '»': '«', '‹': '›', '›': '‹',
};

/* Strong right-to-left: Arabic, Supplement/Extended, Presentation
   Forms A and B, plus Hebrew and Syriac/Thaana ranges. */
const _RE_R   = /[\u0590-\u05FF\u0600-\u06FF\u0700-\u074F\u0750-\u077F\u08A0-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;
/* Strong left-to-right: Latin, Greek, Cyrillic. */
const _RE_L   = /[A-Za-z\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF]/;
/* Digits form their own left-to-right sub-run. */
const _RE_NUM = /[0-9]/;

function _dirOf(ch) {
    if (_RE_R.test(ch))   return 'R';
    if (_RE_L.test(ch))   return 'L';
    if (_RE_NUM.test(ch)) return 'N';
    return '?';                      // neutral
}

/* Split into clusters so combining marks travel with their base
   letter when a run is reversed. */
function _clusters(str) {
    const out = [];
    for (const ch of Array.from(str)) {
        if (out.length && _isMark(ch.codePointAt(0))) out[out.length - 1] += ch;
        else out.push(ch);
    }
    return out;
}

/**
 * Reorder a logical-order string into visual order for an RTL
 * paragraph. Latin words and numbers keep their own left-to-right
 * order, everything else is reversed, brackets are mirrored.
 *
 * @param {string} input
 * @returns {string}
 */
export function bidiVisual(input) {
    const src = String(input ?? '');
    if (!src) return '';

    const cl  = _clusters(src);
    const dir = cl.map(c => _dirOf(c[0]));

    // Resolve neutrals: they take the surrounding direction when both
    // sides agree, otherwise the base direction (R).
    for (let i = 0; i < dir.length; i++) {
        if (dir[i] !== '?') continue;
        let a = i - 1; while (a >= 0 && dir[a] === '?') a--;
        let b = i + 1; while (b < dir.length && dir[b] === '?') b++;
        const before = a >= 0         ? (dir[a] === 'R' ? 'R' : 'L') : 'R';
        const after  = b < dir.length ? (dir[b] === 'R' ? 'R' : 'L') : 'R';
        dir[i] = (before === 'L' && after === 'L') ? 'L' : 'R';
    }
    for (let i = 0; i < dir.length; i++) if (dir[i] === 'N') dir[i] = 'L';

    // Build directional runs.
    const runs = [];
    let cur = null;
    for (let i = 0; i < cl.length; i++) {
        if (!cur || cur.dir !== dir[i]) { cur = { dir: dir[i], parts: [] }; runs.push(cur); }
        cur.parts.push(cl[i]);
    }

    // Base is RTL, so runs are emitted in reverse order.
    let out = '';
    for (let i = runs.length - 1; i >= 0; i--) {
        const r = runs[i];
        if (r.dir === 'L') {
            out += r.parts.join('');
        } else {
            for (let k = r.parts.length - 1; k >= 0; k--) {
                const p = r.parts[k];
                out += (p.length === 1 && _MIRROR[p]) ? _MIRROR[p] : p;
            }
        }
    }
    return out;
}

/**
 * Shape + reorder in one call — use for any short string handed
 * straight to pdf.text().
 *
 * For text that must wrap: shape first, run the shaped string
 * through pdf.splitTextToSize(), then pass each returned line
 * through bidiVisual(). That keeps the width measurement honest,
 * because the presentation forms are what actually get drawn.
 *
 * @param {string} input
 * @returns {string}
 */
export function arabicVisual(input) {
    return bidiVisual(shapeArabic(input));
}
