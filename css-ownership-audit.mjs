#!/usr/bin/env node
/* ============================================================
   css-ownership-audit.mjs
   ------------------------------------------------------------
   Reports every CSS selector that is styled in more than one
   stylesheet — "split ownership", the condition that caused the
   oval-button, toolbar-scroll and rating-row bugs. In each case a
   rule in one file silently overrode a rule in another, and the
   only symptom was the wrong appearance.

   Run it before a release and compare the count to the last run.
   The number should fall through Phases 2-4 and never rise.

       node tools/css-ownership-audit.mjs
       node tools/css-ownership-audit.mjs --list      (show all)
       node tools/css-ownership-audit.mjs --max 60    (fail above 60)

   Exits non-zero when the count exceeds --max, so it can gate a
   release if you ever add one.
   ============================================================ */

import fs from 'fs';
import path from 'path';

// Declared in cascade order — the same order as the <link> tags.
const FILES = [
  'dacum-styles.css',
  'dacum-responsive.css',
  'dacum-fixes.css',
  'dacum-typography.css',
  'dacum-components.css',
];

const args    = process.argv.slice(2);
const showAll = args.includes('--list');
const maxIdx  = args.indexOf('--max');
const maxAllowed = maxIdx > -1 ? Number(args[maxIdx + 1]) : Infinity;
const root    = args.find(a => !a.startsWith('--') && !/^\d+$/.test(a)) || '.';

// Selectors are extracted with a brace-depth walk rather than a regex,
// so nested at-rules (@media, @supports) are handled and their inner
// selectors are attributed to the file, not skipped.
// Returns { base, inMedia }. The distinction matters: a selector styled
// at the top level of two different files is genuine split ownership —
// two files both claim to define what the component looks like. A
// selector styled at base level in one file and overridden inside an
// @media block in another is the responsive layer doing its job, and
// flagging it would make the count impossible to drive toward zero and
// therefore useless as a signal.
function selectorsOf(css) {
  css = css.replace(/\/\*[\s\S]*?\*\//g, '');
  const base = new Set(), inMedia = new Set(), typeLayer = new Set();
  let buf = '';
  const stack = [];          // true = we are inside an at-rule block

  for (let i = 0; i < css.length; i++) {
    const ch = css[i];
    if (ch === '{') {
      const head = buf.trim();
      buf = '';
      if (head.startsWith('@')) { stack.push(true); continue; }
      const nested = stack.some(Boolean);
      stack.push(false);
      // Read this block's body without consuming it, to classify it.
      let d = 1, j = i + 1;
      while (j < css.length && d) { if (css[j] === '{') d++; else if (css[j] === '}') d--; j++; }
      const pendingBody = css.slice(i + 1, j - 1);
      // A declaration block that sets ONLY type properties is a
      // legitimate cross-cutting layer, not a competing owner: a
      // typography sheet is *supposed* to set font-size on a component
      // the layout sheet positions. Flagging that would make the count
      // impossible to reach zero and therefore useless as a signal.
      // A block that also sets padding, background or layout is
      // re-skinning the component and IS a competing owner.
      // `color` counts as type: text colour is part of a typographic
      // scale, and a tokens sheet setting it is normal. The line is drawn
      // at anything that changes SHAPE or LAYOUT — padding, background,
      // border-radius, box-shadow, display. That is what .tab was doing
      // from the typography sheet, and it is why it was correctly flagged.
      const TYPE_ONLY = /^(font(-|$)|color$|letter-spacing|line-height|text-|word-|white-space|-webkit-font)/;
      const propsHere = pendingBody
        .split(';').map(d => d.split(':')[0].trim().toLowerCase()).filter(Boolean);
      const typeOnly = propsHere.length > 0 && propsHere.every(pr => TYPE_ONLY.test(pr));

      head.split(',').forEach(sel => {
        sel = sel.trim().replace(/\s+/g, ' ');
        // Keyframe stops (from / to / 0% / 100%) are not selectors.
        if (!sel || /^(from|to|\d+%)$/.test(sel)) return;
        if (typeOnly) { typeLayer.add(sel); return; }
        (nested ? inMedia : base).add(sel);
      });
    } else if (ch === '}') {
      stack.pop(); buf = '';
    } else {
      buf += ch;
    }
  }
  return { base, inMedia, typeLayer };
}

let typeOnlyCount = 0;
const owners  = new Map();     // selector -> Set(file)  [base-level only]
const overrides = new Map();   // selector -> Set(file)  [inside @media]
const present = [];

for (const f of FILES) {
  const p = path.join(root, f);
  if (!fs.existsSync(p)) continue;
  present.push(f);
  const { base, inMedia, typeLayer } = selectorsOf(fs.readFileSync(p, 'utf8'));
  typeOnlyCount += typeLayer.size;
  for (const sel of base) {
    if (!owners.has(sel)) owners.set(sel, new Set());
    owners.get(sel).add(f);
  }
  for (const sel of inMedia) {
    if (!overrides.has(sel)) overrides.set(sel, new Set());
    overrides.get(sel).add(f);
  }
}

const split = [...owners.entries()]
  .filter(([, fs_]) => fs_.size > 1)
  .sort((a, b) => b[1].size - a[1].size || a[0].localeCompare(b[0]));

const short = f => f.replace(/^dacum-|\.css$/g, '');

console.log(`Stylesheets audited : ${present.length}  (${present.map(short).join(', ')})`);
console.log(`Base-level selectors: ${owners.size}`);
console.log(`SPLIT OWNERSHIP     : ${split.length}   <- drive this to 0`);
console.log(`Media overrides     : ${overrides.size}   (normal, not counted)`);
console.log(`Type-layer rules    : ${typeOnlyCount}   (font-only, not counted)`);

if (split.length) {
  const worst = split.filter(([, s]) => s.size >= 3);
  if (worst.length) {
    console.log(`\nStyled in 3+ files — fix these first:`);
    for (const [sel, fset] of worst) {
      console.log(`  ${fset.size}x  ${sel.padEnd(48)} ${[...fset].map(short).join(', ')}`);
    }
  }
  const rest = split.filter(([, s]) => s.size === 2);
  console.log(`\nStyled in 2 files   : ${rest.length}`);
  if (showAll) {
    for (const [sel, fset] of rest) {
      console.log(`      ${sel.padEnd(48)} ${[...fset].map(short).join(', ')}`);
    }
  } else if (rest.length) {
    console.log(`      (run with --list to see them)`);
  }
}

if (split.length > maxAllowed) {
  console.error(`\nFAIL: ${split.length} split selectors exceeds --max ${maxAllowed}`);
  process.exit(1);
}
