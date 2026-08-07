#!/usr/bin/env node
/* ============================================================
   preflight.mjs
   ------------------------------------------------------------
   Catches the two classes of failure in this project that produce
   NO error message — the ones you only discover from a user report.

   1. OFFLINE BREAKAGE. Every asset index.html references, plus
      every module those entry points import, must appear in
      PRECACHE_URLS in sw.js. Miss one and the app keeps working
      perfectly online and silently fails when the network drops —
      which is the entire point of shipping it as a PWA.

   2. STALE-STYLE FLASH. index.html carries no <style> block of its
      own any more, so it depends completely on its stylesheets. If
      the HTML updates and a stylesheet does not, the page paints
      unstyled until the next fetch cycle. Stylesheets therefore
      have to be in the service worker's criticalUrls warm list too,
      not just in PRECACHE_URLS.

   3. VERSION DRIFT. A code change that does not bump CACHE_VERSION
      never reaches anyone who has installed the PWA. They keep
      running the old build and report bugs you already fixed.

   Usage:
     node tools/preflight.mjs              check
     node tools/preflight.mjs --strict     also fail on warnings
   ============================================================ */

import fs from 'fs';
import path from 'path';

const root   = process.argv.find(a => !a.startsWith('--') && !a.endsWith('.mjs') && !a.endsWith('node')) || '.';
const strict = process.argv.includes('--strict');
const read   = f => fs.readFileSync(path.join(root, f), 'utf8');
const exists = f => fs.existsSync(path.join(root, f));

const problems = [];
const warnings = [];
const fail = m => problems.push(m);
const warn = m => warnings.push(m);

if (!exists('index.html') || !exists('sw.js')) {
  console.error('preflight: run this from the project root (index.html and sw.js not found)');
  process.exit(2);
}

const html = read('index.html');
const sw   = read('sw.js');

// ── Collect everything the page pulls in ─────────────────────
const referenced = new Set();
for (const m of html.matchAll(/(?:src|href)="([^"]+\.(?:js|css))"/g)) {
  const url = m[1];
  // Skip anything not served from our own directory: absolute URLs and
  // root-relative paths are third-party (e.g. Cloudflare's injected
  // email-decode script) and are not ours to precache.
  if (/^(https?:)?\/\//.test(url) || url.startsWith('/')) continue;
  referenced.add(url.replace(/^\.\//, ''));
}

// ── Follow ES module imports from each entry point ───────────
// A module imported by app.js is just as fatal to miss as one in a
// <script> tag, and it is far easier to overlook.
const seen = new Set();
const queue = [...referenced].filter(f => f.endsWith('.js'));
while (queue.length) {
  const f = queue.shift();
  if (seen.has(f) || !exists(f)) continue;
  seen.add(f);
  for (const m of read(f).matchAll(/from\s+['"]\.\/([A-Za-z0-9_.-]+\.js)['"]/g)) {
    referenced.add(m[1]);
    if (!seen.has(m[1])) queue.push(m[1]);
  }
}

// ── 1. Precache coverage ─────────────────────────────────────
const precached = new Set(
  [...sw.matchAll(/BASE\s*\+\s*'([^']+)'/g)].map(m => m[1])
);

for (const f of [...referenced].sort()) {
  if (!exists(f)) { warn(`referenced but missing on disk: ${f}`); continue; }
  if (!precached.has(f)) fail(`NOT PRECACHED: ${f} — the app will break offline`);
}

// ── 2. Precache entries that no longer exist ─────────────────
for (const f of [...precached].sort()) {
  if (f === '' || f === 'index.html' || f.endsWith('/')) continue;
  // Deliberately an error, not a warning: sw.js uses cache.addAll, which
  // rejects atomically. ONE 404 in this list means the whole install
  // fails, the new service worker never activates, and every installed
  // user silently keeps running the previous build forever.
  if (!exists(f)) fail(`PRECACHED BUT MISSING: ${f} — cache.addAll is atomic, so NO update reaches installed users`);
}

// ── 3. Stylesheets must also be in the warm list ─────────────
const criticalBlock = (sw.match(/criticalUrls\s*=\s*\[([\s\S]*?)\]/) || [, ''])[1];
const critical = new Set([...criticalBlock.matchAll(/BASE\s*\+\s*'([^']+)'/g)].map(m => m[1]));
for (const f of [...referenced].filter(f => f.endsWith('.css')).sort()) {
  if (!critical.has(f)) {
    warn(`not in criticalUrls: ${f} — an updated index.html may paint unstyled for one cycle`);
  }
}

// ── 4. Version bump discipline ───────────────────────────────
const appV   = (html.match(/APP_VERSION\s*=\s*'([^']+)'/) || [])[1];
const cacheV = (sw.match(/CACHE_VERSION\s*=\s*'([^']+)'/) || [])[1];
if (!appV)   fail('APP_VERSION not found in index.html');
if (!cacheV) fail('CACHE_VERSION not found in sw.js');

// ── 5. Version drift against the last release ────────────────
// Optional but high-value: a code change that does not bump
// CACHE_VERSION never reaches anyone who installed the PWA. They keep
// running the old build and report bugs that are already fixed. If a
// .preflight-last file is present, compare against it.
const stampPath = path.join(root, '.preflight-last');
if (fs.existsSync(stampPath)) {
  const last = JSON.parse(fs.readFileSync(stampPath, 'utf8'));
  const codeFiles = [...referenced].filter(exists).sort();
  const hash = codeFiles.map(f => f + ':' + read(f).length).join('|');
  if (last.hash !== hash && last.cacheVersion === cacheV) {
    fail(`code changed but CACHE_VERSION is still ${cacheV} — installed users will not receive this build`);
  }
  if (process.argv.includes('--stamp')) {
    fs.writeFileSync(stampPath, JSON.stringify({ hash, cacheVersion: cacheV, appVersion: appV }, null, 2));
  }
} else if (process.argv.includes('--stamp')) {
  const codeFiles = [...referenced].filter(exists).sort();
  const hash = codeFiles.map(f => f + ':' + read(f).length).join('|');
  fs.writeFileSync(stampPath, JSON.stringify({ hash, cacheVersion: cacheV, appVersion: appV }, null, 2));
  console.log('created .preflight-last — future runs will detect an unbumped CACHE_VERSION');
}

// ── Report ───────────────────────────────────────────────────
console.log(`Referenced assets : ${referenced.size}  (including transitive imports)`);
console.log(`Precached         : ${precached.size}`);
console.log(`APP_VERSION       : ${appV}`);
console.log(`CACHE_VERSION     : ${cacheV}`);

if (warnings.length) {
  console.log(`\nWarnings (${warnings.length}):`);
  warnings.forEach(w => console.log('  ! ' + w));
}
if (problems.length) {
  console.log(`\nErrors (${problems.length}):`);
  problems.forEach(p => console.log('  x ' + p));
  console.log('\nPREFLIGHT FAILED');
  process.exit(1);
}
if (strict && warnings.length) {
  console.log('\nPREFLIGHT FAILED (--strict: warnings are errors)');
  process.exit(1);
}
console.log('\nPREFLIGHT PASSED');
