/* ============================================================
   dacum-fixes.js  v3.2
   FIX 1: AI card guard (old badge stays hidden)
   FIX 3: PWA Install prompt + button styling
   FIX 5: Version tag
   NOTE:  Hamburger/resize logic is fully owned by dacum-mobile.js
          to avoid double-init conflicts.
   ============================================================ */
(function () {
  'use strict';

  /* ── FIX 5: Version tag ─────────────────────────────────── */
  console.log('%c[DACUM] Running version 3.3', 'color:#667eea;font-weight:700;');

  /* ── FIX 1: AI card guard ───────────────────────────────── */
  window.__USE_NEW_AI_CARD__ = true;

  function _hideOldBadge() {
    var badge = document.getElementById('usageBadge');
    if (badge) badge.style.display = 'none';
  }

  /* ── PWA Install button styling ──────────────────────────
     Deliberately NOT injected here any more. These rules used to be
     written into a runtime <style> block with !important, which
     silently overrode the identical block in dacum-fixes.css and made
     that file's rules dead code. The styling now lives ONLY in
     dacum-fixes.css (same visual result); this file just toggles the
     .dacum-install-visible class. */

  /* ── FIX 3: PWA Install prompt ────────────────────────────
     The event itself is captured by the early stub in index.html's
     <head> (beforeinstallprompt fires once and is never replayed, and
     this file loads at the very end of <body> — too late to be sure of
     catching it). We read it from window.__dacumInstallPrompt and also
     listen for the stub's 'dacum-install-available' relay, so it works
     whichever order the two run in. */
  function _deferred() { return window.__dacumInstallPrompt || null; }

  function _injectInstallButton() {
    if (document.getElementById('dacumInstallBtn')) return;

    var btn = document.createElement('button');
    btn.id        = 'dacumInstallBtn';
    btn.title     = 'Install DACUM Live Pro as an app';
    btn.innerHTML =
      '<svg width="14" height="14" viewBox="0 0 24 24" fill="none"' +
      ' stroke="currentColor" stroke-width="2.5"' +
      ' stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M12 3v13M6 11l6 6 6-6"/>' +
      '<path d="M5 21h14"/></svg>' +
      '<span class="install-label"> Install App</span>';

    /* Safe insertion: try .dtb-right first, then toolbar, then body */
    var target =
      document.querySelector('.dtb-right') ||
      document.getElementById('dacumTopToolbar') ||
      document.body;

    target.appendChild(btn);
    console.log('[PWA] Install button injected into:', target.id || target.className || 'body');

    btn.addEventListener('click', async function () {
      var evt = _deferred();
      if (!evt) return;
      evt.prompt();
      var result = await evt.userChoice;
      console.log('[PWA] Install choice:', result.outcome);
      window.__dacumInstallPrompt = null;
      btn.classList.remove('dacum-install-visible');
    });
  }

  function _showInstallButton() {
    if (!_deferred()) return;                 /* nothing to prompt with */
    _injectInstallButton();
    var btn = document.getElementById('dacumInstallBtn');
    if (btn) btn.classList.add('dacum-install-visible');
  }

  /* Case A: the stub already captured it before this file ran. */
  if (_deferred()) setTimeout(_showInstallButton, 0);

  /* Case B: it arrives later — the stub relays this custom event. */
  window.addEventListener('dacum-install-available', function () {
    setTimeout(_showInstallButton, 0);
  });

  window.addEventListener('appinstalled', function () {
    window.__dacumInstallPrompt = null;
    var btn = document.getElementById('dacumInstallBtn');
    if (btn) btn.classList.remove('dacum-install-visible');
    console.log('[PWA] App installed.');
  });

  /* ── SW Version Guardian — INTENTIONALLY NOT HERE ────────
     Service-Worker update handling is owned entirely by the
     registration block inline at the top of index.html, which already
     does all of this and more:
       • reg.update() on load, then every 30 s
       • SKIP_WAITING on both reg.waiting and updatefound→installed
       • GET_VERSION / VERSION_REPLY check against EXPECTED_SW
       • SW_UPDATED + controllerchange → "Updating…" toast, then reload
       • a per-version sessionStorage throttle that prevents reload loops
     A second copy of that logic in this file (added briefly in an
     earlier revision) duplicated the skip-waiting calls and, worse,
     ran its own controllerchange→reload that bypassed the loop guard
     and the toast. Removed. Do not re-add it here — extend the block
     in index.html instead, and keep EXPECTED_SW in sync with
     CACHE_VERSION in sw.js. */

  /* ── Bootstrap ─────────────────────────────────────────── */
  function _init() {
    _hideOldBadge();
    /* Sidebar/hamburger handled entirely by dacum-mobile.js */
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', _init);
  } else {
    _init();
  }

})();
