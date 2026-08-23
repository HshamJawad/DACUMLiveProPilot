/**
 * ─────────────────────────────────────────────────────────────────────────────
 *  Module Builder — Error Handler  v1.0
 *  Lightweight, non-intrusive error reporting layer.
 *
 *  Public API:
 *    initErrorHandler()          — wire up all global listeners
 *    triggerTestError()          — test a blocking modal  (dev only)
 *    triggerTestError('toast')   — test a toast           (dev only)
 * ─────────────────────────────────────────────────────────────────────────────
 */
(function (global) {
  'use strict';

  /* ── Per-category error counters ─────────────────────────────────────── */
  const _counters = {
    UI      : 0,
    EXPORT  : 0,
    TABLE   : 0,
    SAVE    : 0,
    LOAD    : 0,
    RENDER  : 0,
    PROMISE : 0,
    GENERAL : 0,
  };

  /* ── Detect error category from message / stack ──────────────────────── */
  function _detectCategory(msg, stack) {
    const haystack = ((msg || '') + ' ' + (stack || '')).toLowerCase();
    if (/export|docx|blob|download|fileSaver/i.test(haystack))       return 'EXPORT';
    if (/table|coverrow|row|grid|cell/i.test(haystack))              return 'TABLE';
    if (/save|localstorage|json|stringify|serialize/i.test(haystack)) return 'SAVE';
    if (/load|restore|import|parse|read/i.test(haystack))            return 'LOAD';
    if (/render|innerHTML|dom|element|node/i.test(haystack))         return 'RENDER';
    if (/click|button|input|ui|event/i.test(haystack))               return 'UI';
    if (/promise|async|await|fetch|then/i.test(haystack))            return 'PROMISE';
    return 'GENERAL';
  }

  /* ── Generate a structured, auto-incrementing error code ─────────────── */
  function _generateCode(category) {
    _counters[category] = (_counters[category] || 0) + 1;
    const seq = String(_counters[category]).padStart(3, '0');
    return 'ERR-' + category + '-' + seq;
  }

  /* ── Persist last error to localStorage ──────────────────────────────── */
  function _storeError(payload) {
    try {
      localStorage.setItem('app_last_error', JSON.stringify({
        code    : payload.code,
        message : payload.message,
        time    : payload.time,
        url     : payload.url,
        line    : payload.line,
      }));
    } catch (_) { /* storage full or unavailable — silently skip */ }
  }

  /* ── Format full details for clipboard ───────────────────────────────── */
  function _buildClipboardText(payload) {
    return [
      '╔═══════════════════════════════════════╗',
      '║   Module Builder — Error Report       ║',
      '╚═══════════════════════════════════════╝',
      '',
      'Error Code : ' + payload.code,
      'Message    : ' + payload.message,
      'Time       : ' + payload.time,
      'URL        : ' + payload.url,
      'Source     : ' + (payload.source || '—'),
      'Line / Col : ' + payload.line + ' / ' + payload.col,
      '',
      'Stack Trace:',
      '─────────────────────────────────────────',
      payload.stack || '(not available)',
    ].join('\n');
  }

  /* ── Inject styles once ───────────────────────────────────────────────── */
  function _injectCSS() {
    if (document.getElementById('eh-styles')) return;
    const s = document.createElement('style');
    s.id = 'eh-styles';
    s.textContent = `
      /* ── Error Handler — Modal overlay ─────────────────────────── */
      #eh-overlay {
        position: fixed; inset: 0;
        background: rgba(15, 23, 42, 0.60);
        backdrop-filter: blur(4px);
        -webkit-backdrop-filter: blur(4px);
        z-index: 2147483640;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 20px;
        animation: eh-bg-in 0.2s ease;
      }
      @keyframes eh-bg-in { from { opacity: 0 } to { opacity: 1 } }

      #eh-modal {
        background: #ffffff;
        border-radius: 18px;
        box-shadow: 0 32px 80px rgba(0, 0, 0, 0.28), 0 0 0 1px rgba(0,0,0,0.04);
        width: 100%;
        max-width: 430px;
        padding: 30px 30px 24px;
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
        animation: eh-modal-in 0.25s cubic-bezier(0.22, 0.68, 0, 1.2);
        position: relative;
      }
      @keyframes eh-modal-in {
        from { opacity: 0; transform: translateY(20px) scale(0.96) }
        to   { opacity: 1; transform: translateY(0)    scale(1) }
      }

      .eh-icon-ring {
        width: 52px; height: 52px; border-radius: 14px;
        background: linear-gradient(135deg, #fff1f2, #fee2e2);
        border: 1.5px solid #fecaca;
        display: flex; align-items: center; justify-content: center;
        margin-bottom: 18px;
        flex-shrink: 0;
      }

      #eh-modal h2 {
        font-size: 1.12rem; font-weight: 700;
        color: #111827; margin: 0 0 6px;
        line-height: 1.3;
      }
      #eh-modal .eh-subtitle {
        font-size: 0.87rem; color: #6b7280;
        line-height: 1.55; margin: 0 0 20px;
      }

      .eh-code-pill {
        background: #f8fafc;
        border: 1.5px solid #e2e8f0;
        border-radius: 10px;
        padding: 11px 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 22px;
      }
      .eh-code-label {
        font-size: 0.76rem; font-weight: 600;
        color: #94a3b8;
        letter-spacing: 0.07em;
        text-transform: uppercase;
      }
      .eh-code-value {
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 0.95rem; font-weight: 700;
        color: #dc2626;
        letter-spacing: 0.06em;
        background: #fff5f5;
        padding: 3px 10px;
        border-radius: 6px;
        border: 1px solid #fecaca;
      }

      .eh-btn-row {
        display: flex; gap: 10px;
      }
      .eh-btn {
        flex: 1; padding: 10px 16px;
        border-radius: 9px; border: none;
        font-size: 0.88rem; font-weight: 600;
        cursor: pointer;
        display: inline-flex; align-items: center;
        justify-content: center; gap: 7px;
        transition: background 0.15s, transform 0.1s, box-shadow 0.15s;
        line-height: 1;
      }
      .eh-btn:active { transform: scale(0.97); }

      .eh-btn-copy {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: #fff;
        box-shadow: 0 3px 10px rgba(102,126,234,0.35);
      }
      .eh-btn-copy:hover {
        box-shadow: 0 5px 14px rgba(102,126,234,0.45);
      }
      .eh-btn-copy.eh-copied {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        box-shadow: 0 3px 10px rgba(16,185,129,0.3);
      }
      .eh-btn-close {
        background: #f3f4f6;
        color: #374151;
      }
      .eh-btn-close:hover { background: #e5e7eb; }

      /* ── Toast (non-blocking, bottom-center) ────────────────────── */
      #eh-toast {
        position: fixed;
        bottom: 28px;
        left: 50%;
        transform: translateX(-50%) translateY(110px);
        background: #1e293b;
        color: #e2e8f0;
        padding: 13px 18px;
        border-radius: 13px;
        box-shadow: 0 10px 40px rgba(0,0,0,0.28), 0 0 0 1px rgba(255,255,255,0.06);
        font-family: 'Segoe UI', system-ui, sans-serif;
        font-size: 0.87rem;
        display: flex;
        align-items: center;
        gap: 12px;
        z-index: 2147483641;
        transition: transform 0.38s cubic-bezier(0.22, 0.68, 0, 1.2);
        max-width: 480px;
        width: calc(100% - 40px);
        box-sizing: border-box;
      }
      #eh-toast.eh-visible {
        transform: translateX(-50%) translateY(0);
      }
      .eh-toast-label {
        color: #94a3b8;
        font-size: 0.78rem;
        font-weight: 500;
        white-space: nowrap;
      }
      .eh-toast-code {
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 0.84rem;
        font-weight: 700;
        color: #f87171;
        white-space: nowrap;
      }
      .eh-toast-text {
        flex: 1;
        color: #cbd5e1;
        font-size: 0.84rem;
      }
      .eh-toast-x {
        background: none; border: none;
        color: #64748b; cursor: pointer;
        font-size: 1rem; padding: 2px 4px;
        line-height: 1; flex-shrink: 0;
        border-radius: 4px;
        transition: color 0.15s, background 0.15s;
      }
      .eh-toast-x:hover { color: #e2e8f0; background: rgba(255,255,255,0.08); }
    `;
    document.head.appendChild(s);
  }

  /* ── Clipboard helper ─────────────────────────────────────────────────── */
  function _copyToClipboard(text, btn) {
    const markCopied = function () {
      if (!btn) return;
      const orig = btn.innerHTML;
      btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg> Copied!';
      btn.classList.add('eh-copied');
      setTimeout(function () {
        btn.innerHTML = orig;
        btn.classList.remove('eh-copied');
      }, 2200);
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(markCopied).catch(function () {
        _execCmdCopy(text, btn, markCopied);
      });
    } else {
      _execCmdCopy(text, btn, markCopied);
    }
  }

  function _execCmdCopy(text, btn, onDone) {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.cssText = 'position:fixed;left:-9999px;top:-9999px;opacity:0;pointer-events:none';
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      if (onDone) onDone();
    } catch (_) { /* copy unavailable — silent */ }
  }

  /* ── Build & show blocking modal ─────────────────────────────────────── */
  function _showModal(payload) {
    _injectCSS();

    // Remove any stale overlay first
    const stale = document.getElementById('eh-overlay');
    if (stale) stale.remove();

    const overlay = document.createElement('div');
    overlay.id = 'eh-overlay';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-labelledby', 'eh-modal-title');

    overlay.innerHTML =
      '<div id="eh-modal">' +
        '<div class="eh-icon-ring">' +
          '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
            '<circle cx="12" cy="12" r="10"/>' +
            '<line x1="12" y1="8" x2="12" y2="12"/>' +
            '<line x1="12" y1="16" x2="12.01" y2="16"/>' +
          '</svg>' +
        '</div>' +
        '<h2 id="eh-modal-title">Something went wrong</h2>' +
        '<p class="eh-subtitle">An unexpected error occurred. Your work may have been auto-saved. Please try again or reload the page.</p>' +
        '<div class="eh-code-pill">' +
          '<span class="eh-code-label">Error Code</span>' +
          '<span class="eh-code-value">' + _esc(payload.code) + '</span>' +
        '</div>' +
        '<div class="eh-btn-row">' +
          '<button class="eh-btn eh-btn-copy" id="eh-copy-btn">' +
            '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">' +
              '<rect x="9" y="9" width="13" height="13" rx="2"/>' +
              '<path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/>' +
            '</svg>' +
            'Copy Error' +
          '</button>' +
          '<button class="eh-btn eh-btn-close" id="eh-close-btn">Close</button>' +
        '</div>' +
      '</div>';

    document.body.appendChild(overlay);

    /* Wire buttons */
    const copyBtn  = document.getElementById('eh-copy-btn');
    const closeBtn = document.getElementById('eh-close-btn');
    const modal    = document.getElementById('eh-modal');

    copyBtn.addEventListener('click', function () {
      _copyToClipboard(_buildClipboardText(payload), copyBtn);
    });

    function _close() {
      overlay.style.animation = 'eh-bg-in 0.15s ease reverse forwards';
      setTimeout(function () { if (overlay.parentNode) overlay.remove(); }, 150);
    }

    closeBtn.addEventListener('click', _close);
    overlay.addEventListener('click', function (e) {
      if (e.target === overlay) _close();
    });
    modal.addEventListener('click', function (e) { e.stopPropagation(); });

    /* Escape key */
    function _onKey(e) {
      if (e.key === 'Escape') {
        _close();
        document.removeEventListener('keydown', _onKey);
      }
    }
    document.addEventListener('keydown', _onKey);

    /* Focus trap on first button */
    setTimeout(function () { if (copyBtn) copyBtn.focus(); }, 60);
  }

  /* ── Build & show non-blocking toast ─────────────────────────────────── */
  function _showToast(payload) {
    _injectCSS();

    const stale = document.getElementById('eh-toast');
    if (stale) stale.remove();

    const toast = document.createElement('div');
    toast.id = 'eh-toast';
    toast.setAttribute('role', 'alert');
    toast.innerHTML =
      '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#f87171" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0">' +
        '<circle cx="12" cy="12" r="10"/>' +
        '<line x1="12" y1="8" x2="12" y2="12"/>' +
        '<line x1="12" y1="16" x2="12.01" y2="16"/>' +
      '</svg>' +
      '<span class="eh-toast-text">Something went wrong</span>' +
      '<span class="eh-toast-label">Code:</span>' +
      '<span class="eh-toast-code">' + _esc(payload.code) + '</span>' +
      '<button class="eh-toast-x" title="Dismiss">✕</button>';

    document.body.appendChild(toast);

    /* Slide in (double rAF ensures transition fires) */
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        toast.classList.add('eh-visible');
      });
    });

    function _dismiss() {
      toast.classList.remove('eh-visible');
      setTimeout(function () { if (toast.parentNode) toast.remove(); }, 420);
    }

    toast.querySelector('.eh-toast-x').addEventListener('click', _dismiss);

    /* Auto-dismiss after 7 s */
    setTimeout(_dismiss, 7000);
  }

  /* ── HTML escape helper (prevent XSS in error messages) ──────────────── */
  function _esc(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /* ── Core error processing ────────────────────────────────────────────── */
  function _handle(message, source, lineno, colno, errObj, eventType) {
    /* Skip if nothing meaningful to report */
    if (!message && !errObj) return;

    /* Ignore noisy third-party / injected scripts */
    if (source && /(cdn-cgi|cloudflare|gtag|analytics|hotjar|intercom)/i.test(source)) {
      return;
    }

    const stack    = (errObj && errObj.stack) ? errObj.stack : '(stack not available)';
    const category = _detectCategory(String(message || ''), stack);
    const code     = _generateCode(category);

    const payload = {
      code    : code,
      message : String(message || 'Unknown error'),
      stack   : stack,
      source  : source  || global.location.href,
      line    : lineno  || '?',
      col     : colno   || '?',
      type    : eventType || 'error',
      time    : new Date().toISOString(),
      url     : global.location.href,
    };

    /* ── Console output (full technical detail) ─────────────────── */
    console.error(
      '%c[ErrorHandler] ' + payload.code,
      'color:#dc2626;font-weight:700;font-size:12px',
      '\n' + payload.message,
      '\n─ source:', payload.source,
      '\n─ line:'  , payload.line + ':' + payload.col,
      '\n─ time:'  , payload.time,
      '\n─ stack:\n', payload.stack
    );

    /* ── Persist ─────────────────────────────────────────────────── */
    _storeError(payload);

    /* ── UI ──────────────────────────────────────────────────────── */
    if (eventType === 'unhandledrejection') {
      _showToast(payload);   /* Promises → non-blocking toast */
    } else {
      _showModal(payload);   /* Sync errors → blocking modal  */
    }
  }

  /* ═══════════════════════════════════════════════════════════════════════
   *  PUBLIC API
   * ═══════════════════════════════════════════════════════════════════════ */

  /**
   * initErrorHandler()
   * Call once on page load to activate global error capturing.
   */
  global.initErrorHandler = function () {
    _injectCSS();

    /* 1 — Synchronous JS errors */
    global.onerror = function (message, source, lineno, colno, errObj) {
      _handle(message, source, lineno, colno, errObj, 'error');
      return false; /* let browser continue default handling */
    };

    /* 2 — Unhandled promise rejections */
    global.addEventListener('unhandledrejection', function (ev) {
      const reason  = ev.reason;
      const message = (reason instanceof Error) ? reason.message : String(reason);
      const fakeErr = { stack: (reason instanceof Error) ? reason.stack : '(async rejection)' };
      _handle(message, null, null, null, fakeErr, 'unhandledrejection');
    });

    console.info(
      '%c[ErrorHandler] ✓ Active — Module Builder Error Reporting v1.0',
      'color:#10b981;font-weight:600'
    );
  };

  /**
   * triggerTestError(type?)
   * Developer helper — open the browser console and call:
   *   triggerTestError()          → blocking modal
   *   triggerTestError('toast')   → non-blocking toast
   */
  global.triggerTestError = function (type) {
    const fakeErr = new Error('This is a test error from triggerTestError()');
    if (type === 'toast') {
      _handle('Test async rejection', global.location.href, 0, 0, fakeErr, 'unhandledrejection');
    } else {
      _handle('Test synchronous error', global.location.href, 0, 0, fakeErr, 'error');
    }
  };

}(window));
