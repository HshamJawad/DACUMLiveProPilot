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
  console.log('%c[DACUM] Running version 3.2', 'color:#667eea;font-weight:700;');

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

  /* ── FIX 3: PWA Install prompt ──────────────────────────── */
  var _deferredPrompt = null;

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
      if (!_deferredPrompt) return;
      _deferredPrompt.prompt();
      var result = await _deferredPrompt.userChoice;
      console.log('[PWA] Install choice:', result.outcome);
      _deferredPrompt = null;
      btn.classList.remove('dacum-install-visible');
    });
  }

  function _showInstallButton() {
    _injectInstallButton();
    var btn = document.getElementById('dacumInstallBtn');
    if (btn) btn.classList.add('dacum-install-visible');
  }

  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    _deferredPrompt = e;
    /* Button may not exist yet if DOM is still loading — defer slightly */
    setTimeout(_showInstallButton, 0);
    console.log('[PWA] Install prompt captured.');
  });

  window.addEventListener('appinstalled', function () {
    _deferredPrompt = null;
    var btn = document.getElementById('dacumInstallBtn');
    if (btn) btn.classList.remove('dacum-install-visible');
    console.log('[PWA] App installed.');
  });

  /* ── SW Version Guardian ───────────────────────────────── */
  // Immediately trigger a SW update check on every page load.
  // This catches cases where the old SW is serving stale files.
  //
  // Previously this only asked the waiting worker to skip waiting —
  // but the page you were already looking at kept being served by the
  // OLD worker, so fresh CSS/JS appeared only on the NEXT manual
  // reload. Now a `controllerchange` listener reloads once, the
  // moment the new worker takes over, so a single refresh is enough
  // after every deploy.
  var _swReloading = false;

  function _watchForControllerChange() {
    if (!('serviceWorker' in navigator)) return;

    navigator.serviceWorker.addEventListener('controllerchange', function () {
      // Guard against a reload loop: fire at most once per page load.
      if (_swReloading) return;
      _swReloading = true;
      console.log('[DACUM] New SW took control — reloading for fresh assets.');
      window.location.reload();
    });
  }

  function _activateWaiting(reg) {
    if (reg && reg.waiting) {
      console.log('[DACUM] SW waiting found — activating');
      reg.waiting.postMessage({ type: 'SKIP_WAITING' });
    }
  }

  function _forceSWUpdate() {
    if (!('serviceWorker' in navigator)) return;

    // Only auto-reload when a controller already exists. On a very
    // first visit the SW installs and takes control with nothing
    // stale on screen — reloading there would be a pointless flash.
    if (navigator.serviceWorker.controller) _watchForControllerChange();

    navigator.serviceWorker.getRegistration().then(function (reg) {
      if (!reg) return;

      // A new worker may finish installing slightly after update()
      // resolves — catch that case too.
      reg.addEventListener('updatefound', function () {
        var incoming = reg.installing;
        if (!incoming) return;
        incoming.addEventListener('statechange', function () {
          if (incoming.state === 'installed') _activateWaiting(reg);
        });
      });

      reg.update().then(function () {
        console.log('[DACUM] SW update check complete');
        _activateWaiting(reg);
      }).catch(function () {});
    }).catch(function () {});
  }

  /* ── Bootstrap ─────────────────────────────────────────── */
  function _init() {
    _hideOldBadge();
    _forceSWUpdate();
    /* Sidebar/hamburger handled entirely by dacum-mobile.js */
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', _init);
  } else {
    _init();
  }

})();
