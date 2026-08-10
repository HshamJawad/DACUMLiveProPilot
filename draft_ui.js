// ============================================================
//  draft_ui.js — Full Draft generator: card + modal
//  DACUM Live Pro
// ------------------------------------------------------------
//  Everything here is built at runtime, so none of it is reachable
//  by applyTranslations(). A dacum:langchange listener at the foot
//  of this file rebuilds the modal from state — the same pattern
//  used by duties.js, tasks.js, modules.js and dacum_projects.js.
// ============================================================

import { STAGES, runDraft, cancelDraft, resumeDraft,
         onDraftProgress, isDraftRunning,
         missingPrerequisites, stagesWithExistingContent,
         estimatedCalls, quotaCheck,
         scopeIsMissing }                    from './draft_agent.js';
import { switchTab }                         from './projects.js';
import { showStatus }                        from './renderer.js';

const _t  = (k)    => (window.i18n ? window.i18n.t(k)     : k);
const _tf = (k, v) => (window.i18n ? window.i18n.tf(k, v) : k);

/* Depth into the chain, 1..STAGES-on-chain. Held here rather than in
   appState: it is a dialog setting, not project content. */
const CHAIN   = STAGES.filter(s => !s.optional);
const OPTIONAL = STAGES.filter(s => s.optional);

/* Copy for the off-chain options. The draft-ratings entry is marked
   `caution` because it is the only checkbox in this dialog whose output
   could be mistaken for evidence — its styling says so before the
   banner in the tab has a chance to. */
const OPTIONAL_COPY = {
  additional: { title: 'dgOptAdditional', hint: 'dgOptAdditionalHint' },
  draftTV:    { title: 'dgOptDraftTV',    hint: 'dgOptDraftTVHint', caution: true },
};

let _depth   = CHAIN.length;          // default: generate everything
let _extras  = new Set();             // ids of chosen optional stages
let _phase   = 'setup';               // setup | running | done | error
let _current = -1;                    // index of the stage in flight
let _doneIds = new Set();
let _failed  = null;
let _unsub   = null;

// ── Selection ────────────────────────────────────────────────

function selectedIds() {
  const chain = CHAIN.slice(0, _depth).map(s => s.id);
  // Optional stages are interleaved at their declared position so the
  // progress bar shows the real execution order, not chain-then-extras.
  return STAGES.filter(s => chain.includes(s.id) || _extras.has(s.id))
               .map(s => s.id);
}

// ── Card (Chart Info tab) ────────────────────────────────────

export function renderDraftCard() {
  const host = document.getElementById('draftGeneratorCard');
  if (!host) return;

  host.innerHTML = `
    <div class="dg-card">
      <div class="dg-card-head">
        <span class="dg-card-icon">\u2728</span>
        <div class="dg-card-text">
          <h3 class="dg-card-title">${_esc(_t('dgCardTitle'))}</h3>
          <p class="dg-card-desc">${_esc(_t('dgCardDesc'))}</p>
        </div>
      </div>
      <button type="button" class="dg-card-btn" id="btnOpenDraftModal">
        \u2728 ${_esc(_t('dgCardBtn'))}
      </button>
      <p class="dg-card-hint">\U0001f4a1 ${_esc(_t('dgCardHint'))}</p>
    </div>`;

  host.querySelector('#btnOpenDraftModal')
      .addEventListener('click', openDraftModal);
}

// ── Modal ────────────────────────────────────────────────────

export function openDraftModal() {
  if (document.getElementById('dgOverlay')) return;

  const overlay = document.createElement('div');
  overlay.id = 'dgOverlay';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('dir', (window.i18n && window.i18n.isRTL()) ? 'rtl' : 'ltr');
  document.body.appendChild(overlay);

  /* ESCAPE ROUTES FIRST.
     These were attached after renderModal(). If renderModal() threw,
     openDraftModal() aborted before reaching them — leaving a
     full-screen blurred backdrop with no dialog, no close button and
     no way out but a page reload. Whatever else breaks, the user must
     always be able to dismiss this. */
  overlay.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && _phase !== 'running') closeDraftModal();
  });
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay && _phase !== 'running') closeDraftModal();
  });

  _phase = 'setup';
  renderModal();

  _unsub = onDraftProgress(_onProgress);
  overlay.querySelector('.dg-dialog')?.focus();
}

export function closeDraftModal() {
  if (_unsub) { _unsub(); _unsub = null; }
  document.getElementById('dgOverlay')?.remove();
}

function renderModal() {
  const overlay = document.getElementById('dgOverlay');
  if (!overlay) return;

  try {
    _renderModalInner(overlay);
  } catch (err) {
    /* A blank blurred screen tells the user nothing and offers no way
       back. Show the failure, and always show a way out. */
    console.error('[draft] modal render failed', err);
    overlay.innerHTML = `
      <div class="dg-dialog" tabindex="-1">
        <div class="dg-head">
          <span class="dg-head-icon">\u26A0\uFE0F</span>
          <h2 class="dg-head-title">${_esc(_t('dgRenderFailed'))}</h2>
        </div>
        <div class="dg-body">
          <p class="dg-intro">${_esc(_t('dgRenderFailedBody'))}</p>
          <pre class="dg-errbox">${_esc(err && err.message ? err.message : String(err))}</pre>
        </div>
        <div class="dg-foot">
          <button type="button" class="dg-btn dg-btn-go" id="dgFail">${_esc(_t('dgBtnClose'))}</button>
        </div>
      </div>`;
    overlay.querySelector('#dgFail')?.addEventListener('click', closeDraftModal);
  }
}

function _renderModalInner(overlay) {
  overlay.innerHTML = `
    <div class="dg-dialog" tabindex="-1">
      <div class="dg-head">
        <span class="dg-head-icon">\u2728</span>
        <h2 class="dg-head-title">${_esc(_t('dgModalTitle'))}</h2>
        ${_phase === 'running' ? '' :
          `<button type="button" class="dg-x" id="dgClose"
                   aria-label="${_esc(_t('dgBtnClose'))}">\u2715</button>`}
      </div>
      <div class="dg-body">${_phase === 'setup' ? _setupBody() : _runBody()}</div>
      <div class="dg-foot">${_footButtons()}</div>
    </div>`;

  /* Wiring failures must not leave a rendered-but-dead dialog: the
     close button is attached first inside _wire(), and any later
     failure is logged rather than thrown on. */
  try {
    _wire();
  } catch (err) {
    console.error('[draft] modal wiring failed', err);
    showStatus(_t('dgWireFailed'), 'error');
  }
}

// ── Setup view ───────────────────────────────────────────────

function _setupBody() {
  const missing = missingPrerequisites();
  if (missing.length) return _prereqBody(missing);

  const ids       = selectedIds();
  const clashes   = stagesWithExistingContent(ids);
  const clashTabs = [...new Set(
    STAGES.filter(s => clashes.includes(s.id)).map(s => _t(s.labelKey))
  )].join('\u060C ');

  return `
    <p class="dg-intro">${_esc(_t('dgModalIntro'))}</p>

    <p class="dg-label">${_esc(_t('dgDepthLabel'))}</p>
    <ol class="dg-chain">
      ${CHAIN.map((s, i) => `
        <li class="dg-chain-item ${i < _depth ? 'is-on' : ''}"
            data-depth="${i + 1}" role="button" tabindex="0">
          <span class="dg-chain-num">${i + 1}</span>
          <span class="dg-chain-label">${_esc(_t(s.labelKey))}</span>
          <span class="dg-chain-check">${i < _depth ? '\u2713' : ''}</span>
        </li>`).join('')}
    </ol>

    <p class="dg-label">${_esc(_t('dgExtrasLabel'))}</p>
    ${OPTIONAL.map(s => {
      const copy = OPTIONAL_COPY[s.id] || { title: s.labelKey, hint: null };
      return `
      <label class="dg-extra ${copy.caution ? 'is-caution' : ''}">
        <input type="checkbox" data-extra="${s.id}"
               ${_extras.has(s.id) ? 'checked' : ''}>
        <span>
          <strong>${_esc(_t(copy.title))}</strong>
          ${copy.hint ? `<small>${_esc(_t(copy.hint))}</small>` : ''}
        </span>
      </label>`;
    }).join('')}

    ${scopeIsMissing() ? `
      <div class="dg-note dg-note-warn">
        <strong>\u26A0\uFE0F ${_esc(_t('dgScopeSoftTitle'))}</strong>
        <p>${_esc(_t('dgScopeSoftBody'))}</p>
        <button type="button" class="dg-inline-btn" id="dgAddScope">
          ${_esc(_t('dgScopeAddNow'))}
        </button>
      </div>` : ''}

    <div class="dg-note dg-note-info">
      <strong>\u{1F465} ${_esc(_t('dgVerifExcludedTitle'))}</strong>
      <p>${_esc(_t('dgVerifExcludedBody'))}</p>
    </div>

    ${clashes.length ? `
      <div class="dg-note dg-note-warn">
        <strong>\u26A0\uFE0F ${_esc(_t('dgOverwriteTitle'))}</strong>
        <p>${_esc(_tf('dgOverwriteBody', { tabs: clashTabs }))}</p>
      </div>` : ''}

    ${_quotaBlock(ids)}`;
}

/* The quota line is shown BEFORE the run, not discovered during it.
   quotaCheck() counts the whole chain, so a run either has room for
   every stage or does not start — spending four generations and dying
   at the fifth is the one outcome worth engineering against. */
function _quotaBlock(ids) {
  const q = quotaCheck(ids);
  if (!q.ok) {
    return `
      <div class="dg-note dg-note-warn">
        <strong>\u26A0\uFE0F ${_esc(_t('dgQuotaTitle'))}</strong>
        <p>${_esc(_tf('dgQuotaBody',
          { need: q.need, left: q.left, max: q.max }))}</p>
      </div>`;
  }
  return `
    <p class="dg-cost">
      ${_esc(_tf('dgCostLabel', { n: q.need, max: q.max }))}
      <span class="dg-cost-left">${_esc(_tf('dgQuotaRemaining',
        { left: q.left, max: q.max }))}</span>
    </p>`;
}

/* Missing prerequisites are COLLECTED here rather than reported as an
   error. Opening a dialog only to be told to go elsewhere and come
   back is the worst of both: it interrupts and it does not help. */
function _prereqBody(missing) {
  const field = (id, labelKey, tag) => `
    <label class="dg-prereq-field">
      <span>${_esc(_t(labelKey))}</span>
      ${tag === 'textarea'
        ? `<textarea id="dgFix_${id}" rows="3"></textarea>`
        : `<input type="text" id="dgFix_${id}">`}
    </label>`;

  return `
    <div class="dg-note dg-note-warn">
      <strong>\u26A0\uFE0F ${_esc(_t('dgPrereqTitleV2'))}</strong>
      <p>${_esc(_t('dgPrereqBodyV2'))}</p>
    </div>
    ${missing.includes('occupationTitle')
      ? field('occupationTitle', 'dgNeedOccupation', 'input') : ''}
    ${missing.includes('scopeOfWork')
      ? field('scopeOfWork', 'dgNeedScope', 'textarea') : ''}`;
}

// ── Running / done view ──────────────────────────────────────

function _runBody() {
  const ids = selectedIds();
  const run = STAGES.filter(s => ids.includes(s.id));

  const title = _phase === 'done'  ? _t('dgDoneTitle')
              : _phase === 'error' ? _tf('dgStageFailed',
                  { stage: _t(STAGES.find(s => s.id === _failed)?.labelKey || '') })
              : _t('dgRunningTitle');

  const hint  = _phase === 'done' ? _t('dgDoneHint')
              : _phase === 'error' ? '' : _t('dgRunningHint');

  return `
    <p class="dg-run-title">${_esc(title)}</p>

    <ol class="dg-progress">
      ${run.map((s, i) => {
        const state = _doneIds.has(s.id) ? 'done'
                    : _failed === s.id   ? 'error'
                    : i === _current     ? 'active' : 'idle';
        const mark  = state === 'done'  ? '\u2713'
                    : state === 'error' ? '\u2715'
                    : state === 'active' ? '\u2026' : '';
        return `
          <li class="dg-prog-item is-${state}">
            <span class="dg-prog-mark">${mark}</span>
            <span class="dg-prog-label">${_esc(_t(s.labelKey))}</span>
          </li>`;
      }).join('')}
    </ol>

    ${hint ? `<p class="dg-run-hint">${_esc(hint)}</p>` : ''}`;
}

function _footButtons() {
  if (_phase === 'setup') {
    const missing = missingPrerequisites();
    return missing.length
      ? `<button type="button" class="dg-btn dg-btn-ghost" id="dgCancel">${_esc(_t('dgBtnCancel'))}</button>
         <button type="button" class="dg-btn dg-btn-go" id="dgSaveFix">${_esc(_t('dgSaveAndContinue'))}</button>`
      : `<button type="button" class="dg-btn dg-btn-ghost" id="dgCancel">${_esc(_t('dgBtnCancel'))}</button>
         <button type="button" class="dg-btn dg-btn-go" id="dgStart"
                 ${quotaCheck(selectedIds()).ok ? '' : 'disabled'}>\u2728 ${_esc(_t('dgBtnStart'))}</button>`;
  }
  if (_phase === 'running') {
    return `<button type="button" class="dg-btn dg-btn-ghost" id="dgStop">${_esc(_t('dgBtnStop'))}</button>`;
  }
  if (_phase === 'error') {
    return `<button type="button" class="dg-btn dg-btn-ghost" id="dgClose2">${_esc(_t('dgBtnClose'))}</button>
            <button type="button" class="dg-btn dg-btn-go" id="dgRetry">${_esc(_t('dgBtnRetry'))}</button>`;
  }
  return `<button type="button" class="dg-btn dg-btn-go" id="dgReview">${_esc(_t('dgBtnReview'))}</button>`;
}

// ── Wiring ───────────────────────────────────────────────────

function _wire() {
  const q = (sel) => document.querySelector('#dgOverlay ' + sel);

  // Exits first: if anything below fails, these are already live.
  q('#dgClose')?.addEventListener('click', closeDraftModal);
  q('#dgClose2')?.addEventListener('click', closeDraftModal);
  q('#dgCancel')?.addEventListener('click', closeDraftModal);

  // Depth: clicking stage N selects stages 1..N. The dependency is
  // expressed by the interaction itself, so there is no invalid state
  // to warn about.
  document.querySelectorAll('#dgOverlay .dg-chain-item').forEach(el => {
    const pick = () => {
      _depth = parseInt(el.getAttribute('data-depth'), 10);
      renderModal();
    };
    el.addEventListener('click', pick);
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); }
    });
  });

  document.querySelectorAll('#dgOverlay [data-extra]').forEach(cb => {
    cb.addEventListener('change', () => {
      const id = cb.getAttribute('data-extra');
      cb.checked ? _extras.add(id) : _extras.delete(id);
      renderModal();
    });
  });

  q('#dgSaveFix')?.addEventListener('click', () => {
    /* Write straight into the real Chart Info fields so the values are
       saved with the project, not held only by this dialog.

       Each field is wrapped separately: those inputs carry autosave and
       history listeners, and an exception thrown by ANY of them would
       otherwise abort this handler before renderModal() — the button
       would appear to do nothing at all, which is exactly the symptom
       this replaced. */
    let wrote = 0;
    ['occupationTitle', 'scopeOfWork'].forEach(id => {
      try {
        const src = document.getElementById('dgFix_' + id);
        const dst = document.getElementById(id);
        if (!src || !dst || !src.value.trim()) return;
        dst.value = src.value.trim();
        try { dst.dispatchEvent(new Event('input',  { bubbles: true })); } catch (_) {}
        try { dst.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {}
        wrote++;
      } catch (err) {
        console.error('[draft] could not write ' + id, err);
      }
    });

    console.log('[draft] save & continue: wrote', wrote, 'field(s)');

    if (!wrote && missingPrerequisites().length) {
      // Still blocked, and nothing was written: say so rather than
      // re-rendering the same panel and looking inert.
      showStatus(_t('dgSaveFailed'), 'error');
      return;
    }
    renderModal();
  });

  q('#dgAddScope')?.addEventListener('click', () => {
    // Jump to the field rather than duplicating it here: Scope is a
    // paragraph, and a cramped textarea in a dialog is a worse place to
    // write one than the tab it belongs to.
    closeDraftModal();
    switchTab('info-tab');
    const el = document.getElementById('scopeOfWork');
    if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); el.focus(); }
  });

  q('#dgStart')?.addEventListener('click', async () => {
    _phase = 'running'; _current = 0; _doneIds = new Set(); _failed = null;
    renderModal();
    const res = await runDraft(selectedIds());
    /* runDraft can refuse before emitting any progress event — quota,
       or a run already in flight. Without this the dialog would sit on
       "generating" forever with nothing happening behind it. */
    if (res && !res.ok && (res.reason === 'quota' || res.reason === 'already-running')) {
      _phase = 'setup';
      renderModal();
    }
  });

  q('#dgStop')?.addEventListener('click', () => {
    cancelDraft();
    const el = document.querySelector('#dgOverlay .dg-run-hint');
    if (el) el.textContent = _t('dgStopping');
  });

  q('#dgRetry')?.addEventListener('click', () => {
    _phase = 'running'; _failed = null;
    renderModal();
    resumeDraft(selectedIds());
  });

  q('#dgReview')?.addEventListener('click', () => {
    closeDraftModal();
    switchTab('duties-tab');
  });
}

// ── Progress ─────────────────────────────────────────────────

function _onProgress(ev) {
  switch (ev.type) {
    case 'stage-start':   _current = ev.index; break;
    case 'stage-done':    _doneIds.add(ev.id); break;
    case 'stage-skipped': _doneIds.add(ev.id); break;
    case 'stage-error':   _failed = ev.id; _phase = 'error'; break;
    case 'cancelled':     _phase = 'done'; break;
    case 'complete':      _phase = 'done'; _current = -1; break;
    default: return;
  }
  if (document.getElementById('dgOverlay')) renderModal();
}

// ── Helpers ──────────────────────────────────────────────────

function _esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ── Language change ──────────────────────────────────────────
//
// Both the card and the modal are innerHTML, so applyTranslations()
// cannot reach them. The modal is only rebuilt while idle: during a
// run the switcher is locked anyway, and a rebuild mid-stage would
// reset the progress list the user is watching.
window.addEventListener('dacum:langchange', () => {
  renderDraftCard();
  const overlay = document.getElementById('dgOverlay');
  if (overlay && !isDraftRunning()) {
    overlay.setAttribute('dir', (window.i18n && window.i18n.isRTL()) ? 'rtl' : 'ltr');
    renderModal();
  }
});
