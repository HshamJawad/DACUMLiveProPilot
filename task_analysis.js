// ============================================================
// /task_analysis.js
// Task Analysis tab — sits between Task Verification and
// Competency Clusters in the DACUM workflow.
//
// Data lives in appState.taskAnalysisData, a flat dictionary keyed by
// the task's existing inputId (same key already used by
// verificationRatings / taskMetadata in tasks.js). This is deliberate:
// dutiesData[].tasks[] is rebuilt on every add/remove/reorder, so a
// flat key is what survives that — nesting the record inside the task
// object itself would silently detach it the next time the list is
// rebuilt from state.
//
// Records are created lazily on first edit, not on render. Rendering
// a task that has never been touched must NOT write an empty record
// into appState — otherwise every task in the chart would pick up a
// blank Task Analysis entry just by being displayed once, bloating
// every saved project and JSON export with noise.
// ============================================================

import { appState }        from './state.js';
import { showStatus, escapeHtml } from './renderer.js';
import { getDutyLetter }   from './codes.js';
import { syncAllFromDOM }  from './duties.js';

/* i18n access — resolved lazily; see duties.js for why. */
const _t  = (k)    => (window.i18n ? window.i18n.t(k)     : k);

/* Duty letters and task codes are Latin and must never be reordered by
   the bidi algorithm inside Arabic text — same convention as
   duties.js's _bdi(). */
const _bdi = (code) => `<bdi>${code}</bdi>`;

// ── Field definitions ─────────────────────────────────────────
// Single source of truth for which fields exist, their i18n keys, and
// whether they render as a reorderable list or a plain textarea.
// Adding a field to a future revision means adding one line here.
const LIST_FIELDS = [
  { key: 'performanceSteps',            labelKey: 'taLblSteps',      phKey: 'taPhSteps' },
  { key: 'requiredKnowledge',           labelKey: 'taLblKnowledge',  phKey: 'taPhKnowledge' },
  { key: 'requiredSkills',              labelKey: 'taLblSkills',     phKey: 'taPhSkills' },
  { key: 'toolsEquipmentMaterials',     labelKey: 'taLblTools',      phKey: 'taPhTools' },
  { key: 'safetyOSH',                   labelKey: 'taLblSafety',     phKey: 'taPhSafety' },
  { key: 'decisionsCriticalPoints',     labelKey: 'taLblDecisions',  phKey: 'taPhDecisions' },
  { key: 'performanceCriteria',         labelKey: 'taLblCriteria',   phKey: 'taPhCriteria' },
  { key: 'commonErrorsTroubleshooting', labelKey: 'taLblErrors',     phKey: 'taPhErrors' },
];

const TEXT_FIELDS = [
  { key: 'conditionsWorkEnvironment', labelKey: 'taLblConditions', phKey: 'taPhConditions' },
  { key: 'performanceStandard',       labelKey: 'taLblStandard',   phKey: 'taPhStandard' },
];

// Fields that must ALL have at least one entry/value for a task to be
// considered "Completed" rather than "In Progress" — a heuristic for
// the status dot only. Nothing in the app is blocked by this; per the
// spec, no field is ever required to proceed.
const _CORE_COMPLETE_KEYS = [
  'performanceSteps', 'requiredKnowledge', 'requiredSkills',
  'toolsEquipmentMaterials', 'safetyOSH', 'performanceCriteria'
];

function _blankRecord() {
  return {
    performanceSteps: [], requiredKnowledge: [], requiredSkills: [],
    toolsEquipmentMaterials: [], safetyOSH: [], conditionsWorkEnvironment: '',
    decisionsCriticalPoints: [], performanceCriteria: [], performanceStandard: '',
    commonErrorsTroubleshooting: []
  };
}

// Read-only accessor — returns the stored record or null. Never writes.
function _record(taskKey) {
  return (appState.taskAnalysisData || {})[taskKey] || null;
}

// Read accessor for rendering — never writes a blank record into
// appState just because the UI displayed it.
function _view(taskKey) {
  return _record(taskKey) || _blankRecord();
}

// Get-or-create — the ONLY place a record is written into appState.
function _ensureRecord(taskKey) {
  if (!appState.taskAnalysisData) appState.taskAnalysisData = {};
  if (!appState.taskAnalysisData[taskKey]) appState.taskAnalysisData[taskKey] = _blankRecord();
  return appState.taskAnalysisData[taskKey];
}

function _isRecordEmpty(r) {
  if (!r) return true;
  return LIST_FIELDS.every(f => !(r[f.key] && r[f.key].length)) &&
         !(r.conditionsWorkEnvironment || '').trim() &&
         !(r.performanceStandard || '').trim();
}

function _status(taskKey) {
  const r = _record(taskKey);
  if (_isRecordEmpty(r)) return 'not-started';
  const coreFilled = _CORE_COMPLETE_KEYS.every(k => r[k] && r[k].length > 0) &&
                      (r.performanceStandard || '').trim();
  return coreFilled ? 'completed' : 'in-progress';
}

function _statusLabel(status) {
  if (status === 'completed')   return _t('taStatusCompleted');
  if (status === 'in-progress') return _t('taStatusInProgress');
  return _t('taStatusNotStarted');
}

function _statusDot(status) {
  if (status === 'completed')   return '✓';
  if (status === 'in-progress') return '◐';
  return '○';
}

/** True when at least one task in the whole chart has any analysis
 *  content — used by clearCurrentTab/_isTabEmpty (projects.js) and by
 *  the PDF/DOCX export functions to decide whether to include the
 *  Task Analysis appendix at all. */
export function hasAnyTaskAnalysis() {
  return Object.values(appState.taskAnalysisData || {}).some(r => !_isRecordEmpty(r));
}

/** Count of tasks that have any actual analysis content — used for the
 *  "you're about to lose N records" warning in clearCurrentTab. */
export function countTaskAnalysisRecords() {
  return Object.values(appState.taskAnalysisData || {}).filter(r => !_isRecordEmpty(r)).length;
}

/** Full export dataset: every task that has ANY analysis content, in
 *  duty/task order, with duty letter + task code already resolved.
 *  Consumed by exports_pdf.js and exports_docx.js. */
export function getTaskAnalysisExportData() {
  syncAllFromDOM();
  return _allTasksFlat()
    .filter(entry => !_isRecordEmpty(_record(entry.taskKey)))
    .map(entry => ({
      dutyLetter: getDutyLetter(entry.dutyIndex),
      dutyTitle:  entry.dutyTitle,
      taskCode:   `${getDutyLetter(entry.dutyIndex)}${entry.taskNum}`,
      taskText:   entry.task.text,
      record:     _record(entry.taskKey)
    }));
}

// ── Selection state (module-local, not persisted) ──────────────
let _selectedTaskKey = null;
let _lastSignature    = null;

function _signature() {
  return (appState.dutiesData || []).map(d =>
    'D:' + d.id + '|' + (d.title || '').trim() + '|' +
    (d.tasks || []).map(t => t.inputId + '|' + (t.text || '').trim()).join(',')
  ).join('\n');
}

function _allTasksFlat() {
  const out = [];
  (appState.dutiesData || []).forEach((duty, dutyIndex) => {
    const dutyTitle = (duty.title || '').trim();
    let taskNum = 0;
    (duty.tasks || []).forEach(task => {
      const text = (task.text || '').trim();
      if (!text) return;
      taskNum++;
      out.push({
        dutyId: duty.id, dutyIndex, dutyTitle: dutyTitle || _t('lblUntitledDuty'),
        task, taskKey: task.inputId, taskNum
      });
    });
  });
  return out;
}

// ── Entry point: called on EVERY entry to this tab, from both
//    tabs.js's click listener and switchTab() in projects.js — the
//    two must stay in step, same rule as syncVerificationTab(). ──
export function syncTaskAnalysisTab() {
  const nav = document.getElementById('taskAnalysisNav');
  if (!nav) return;
  const sig = _signature();
  if (sig === _lastSignature) return;   // nothing changed — leave as-is
  _lastSignature = sig;
  renderTaskAnalysisTab();
}

export function renderTaskAnalysisTab() {
  syncAllFromDOM();
  _renderNav();
  _renderFormPanel();
}

// ── Task Navigator (left/top panel, grouped by duty) ────────────
function _renderNav() {
  const nav = document.getElementById('taskAnalysisNav');
  if (!nav) return;
  const flat = _allTasksFlat();

  if (!flat.length) {
    nav.innerHTML = `
      <div class="no-duties-message">
        <h3>⚠️ ${_t('msgNoDutiesFoundTitleTA')}</h3>
        <p>${_t('msgNoDutiesForAnalysis')}</p>
      </div>`;
    _selectedTaskKey = null;
    return;
  }

  if (!_selectedTaskKey || !flat.some(f => f.taskKey === _selectedTaskKey)) {
    _selectedTaskKey = flat[0].taskKey;
  }

  const byDuty = [];
  let current = null;
  flat.forEach(f => {
    if (!current || current.dutyId !== f.dutyId) {
      current = { dutyId: f.dutyId, dutyIndex: f.dutyIndex, dutyTitle: f.dutyTitle, items: [] };
      byDuty.push(current);
    }
    current.items.push(f);
  });

  nav.innerHTML = byDuty.map(d => {
    const letter = getDutyLetter(d.dutyIndex);
    const rows = d.items.map(f => {
      const status = _status(f.taskKey);
      const active = f.taskKey === _selectedTaskKey ? ' ta-nav-task-active' : '';
      return `
        <button type="button" class="ta-nav-task${active} ta-status-${status}"
                data-action="ta-select-task" data-task-key="${f.taskKey}"
                title="${escapeHtml(f.task.text)}">
          <span class="ta-nav-dot">${_statusDot(status)}</span>
          <span class="ta-nav-code">${_bdi(letter + f.taskNum)}</span>
          <span class="ta-nav-text">${escapeHtml(f.task.text)}</span>
        </button>`;
    }).join('');
    return `
      <div class="ta-nav-duty">
        <div class="ta-nav-duty-title">${_bdi(letter)}: ${escapeHtml(d.dutyTitle)}</div>
        <div class="ta-nav-duty-tasks">${rows}</div>
      </div>`;
  }).join('');
}

// Lightweight status refresh — touches only the affected DOM nodes so
// typing in a field never triggers a full nav/form rebuild (which
// would steal focus mid-keystroke).
function _touchStatus(taskKey) {
  const status = _status(taskKey);
  const navBtn = document.querySelector(`.ta-nav-task[data-task-key="${taskKey}"]`);
  if (navBtn) {
    navBtn.classList.remove('ta-status-not-started', 'ta-status-in-progress', 'ta-status-completed');
    navBtn.classList.add('ta-status-' + status);
    const dot = navBtn.querySelector('.ta-nav-dot');
    if (dot) dot.textContent = _statusDot(status);
  }
  if (taskKey === _selectedTaskKey) {
    const pill = document.getElementById('taskAnalysisStatusPill');
    if (pill) {
      pill.className = `completion-indicator ${status === 'completed' ? 'complete' : 'incomplete'} ta-status-pill ta-status-pill-${status}`;
      pill.textContent = _statusLabel(status);
    }
  }
}

// ── Form panel (right/main panel — selected task's analysis) ───
function _renderListField(taskKey, field, items) {
  const rows = (items || []).map((val, idx) => `
    <div class="ta-list-row">
      <span class="ta-list-num">${_bdi(String(idx + 1) + '.')}</span>
      <input type="text" class="ta-list-input" value="${escapeHtml(val)}"
             data-action="ta-edit-item" data-field="${field.key}" data-index="${idx}"
             placeholder="${escapeHtml(_t(field.phKey))}">
      <button type="button" class="ta-icon-btn" data-action="ta-move-up"
              data-field="${field.key}" data-index="${idx}" ${idx === 0 ? 'disabled' : ''}
              title="${escapeHtml(_t('ttMoveItemUp'))}">↑</button>
      <button type="button" class="ta-icon-btn" data-action="ta-move-down"
              data-field="${field.key}" data-index="${idx}" ${idx === items.length - 1 ? 'disabled' : ''}
              title="${escapeHtml(_t('ttMoveItemDown'))}">↓</button>
      <button type="button" class="ta-icon-btn ta-icon-btn-danger" data-action="ta-remove-item"
              data-field="${field.key}" data-index="${idx}"
              title="${escapeHtml(_t('ttRemoveItem'))}">✕</button>
    </div>`).join('');

  return `
    <div class="ta-field-block" data-field-block="${field.key}">
      <div class="ta-field-header">
        <h4>${_t(field.labelKey)}</h4>
        <button type="button" class="ta-add-btn" data-action="ta-add-item" data-field="${field.key}">
          ➕ ${_t('btnTaAddItem')}
        </button>
      </div>
      <div class="ta-list-items">${rows || `<p class="ta-empty-hint">${_t('taEmptyListHint')}</p>`}</div>
    </div>`;
}

function _renderTextField(field, value) {
  return `
    <div class="ta-field-block">
      <div class="ta-field-header"><h4>${_t(field.labelKey)}</h4></div>
      <textarea class="ta-textarea" data-action="ta-edit-text" data-field="${field.key}"
                placeholder="${escapeHtml(_t(field.phKey))}">${escapeHtml(value || '')}</textarea>
    </div>`;
}

function _renderFormPanel() {
  const panel = document.getElementById('taskAnalysisFormPanel');
  if (!panel) return;

  const flat  = _allTasksFlat();
  const entry = _selectedTaskKey ? flat.find(f => f.taskKey === _selectedTaskKey) : null;

  if (!entry) {
    panel.innerHTML = `<div class="no-tasks-message">${_t('msgSelectTaskForAnalysis')}</div>`;
    return;
  }

  const r      = _view(entry.taskKey);
  const letter = getDutyLetter(entry.dutyIndex);
  const status = _status(entry.taskKey);
  const byField = Object.fromEntries(LIST_FIELDS.map(f => [f.key, f]));

  panel.innerHTML = `
    <div class="ta-form-header">
      <div class="ta-form-header-row">
        <div>
          <div class="ta-form-duty">${_bdi(letter)}: ${escapeHtml(entry.dutyTitle)}</div>
          <div class="ta-form-task">${_bdi(letter + entry.taskNum)}. ${escapeHtml(entry.task.text)}</div>
        </div>
        <span id="taskAnalysisStatusPill"
              class="completion-indicator ${status === 'completed' ? 'complete' : 'incomplete'} ta-status-pill ta-status-pill-${status}">
          ${_statusLabel(status)}
        </span>
      </div>
      <button type="button" class="ta-clear-btn" data-action="ta-clear-analysis" data-task-key="${entry.taskKey}">
        🗑️ ${_t('btnClearAnalysis')}
      </button>
    </div>

    ${_renderListField(entry.taskKey, byField.performanceSteps,            r.performanceSteps)}
    ${_renderListField(entry.taskKey, byField.requiredKnowledge,           r.requiredKnowledge)}
    ${_renderListField(entry.taskKey, byField.requiredSkills,              r.requiredSkills)}
    ${_renderListField(entry.taskKey, byField.toolsEquipmentMaterials,     r.toolsEquipmentMaterials)}
    ${_renderListField(entry.taskKey, byField.safetyOSH,                  r.safetyOSH)}
    ${_renderTextField(TEXT_FIELDS[0], r.conditionsWorkEnvironment)}
    ${_renderListField(entry.taskKey, byField.decisionsCriticalPoints,     r.decisionsCriticalPoints)}
    ${_renderListField(entry.taskKey, byField.performanceCriteria,         r.performanceCriteria)}
    ${_renderTextField(TEXT_FIELDS[1], r.performanceStandard)}
    ${_renderListField(entry.taskKey, byField.commonErrorsTroubleshooting, r.commonErrorsTroubleshooting)}
  `;
}

// ── Clear (per-task and whole-tab) ──────────────────────────────

function _clearOneTaskAnalysis(taskKey) {
  const r = _record(taskKey);
  if (_isRecordEmpty(r)) {
    showStatus(_t('msgTaAlreadyEmpty'), 'success');
    return;
  }
  if (!confirm(_t('confirmClearTaskAnalysis'))) return;
  delete appState.taskAnalysisData[taskKey];
  _renderNav();
  _renderFormPanel();
  showStatus(_t('msgTaskAnalysisCleared') + ' ✓', 'success');
}

/** Called by clearCurrentTab('task-analysis-tab') in projects.js. */
export function clearAllTaskAnalysis() {
  appState.taskAnalysisData = {};
  _selectedTaskKey = null;
  renderTaskAnalysisTab();
}

// ── Event wiring ─────────────────────────────────────────────
// Scoped to the tab's own container, same delegation pattern events.js
// uses for every other tab (see e.g. the dutiesCont listener there).
export function setupTaskAnalysisEvents() {
  const root = document.getElementById('task-analysis-tab');
  if (!root || root.__taWired) return;
  root.__taWired = true;

  root.addEventListener('click', function (e) {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const action = btn.getAttribute('data-action');

    if (action === 'ta-select-task') {
      _selectedTaskKey = btn.getAttribute('data-task-key');
      _renderNav();
      _renderFormPanel();
      return;
    }

    if (action === 'ta-add-item') {
      const field = btn.getAttribute('data-field');
      const r = _ensureRecord(_selectedTaskKey);
      r[field].push('');
      _renderFormPanel();
      _touchStatus(_selectedTaskKey);
      const inputs = document.querySelectorAll(`[data-field-block="${field}"] .ta-list-input`);
      const last = inputs[inputs.length - 1];
      if (last) last.focus();
      return;
    }

    if (action === 'ta-remove-item') {
      const field = btn.getAttribute('data-field');
      const idx   = parseInt(btn.getAttribute('data-index'), 10);
      const r = _ensureRecord(_selectedTaskKey);
      r[field].splice(idx, 1);
      _renderFormPanel();
      _touchStatus(_selectedTaskKey);
      return;
    }

    if (action === 'ta-move-up' || action === 'ta-move-down') {
      const field = btn.getAttribute('data-field');
      const idx   = parseInt(btn.getAttribute('data-index'), 10);
      const r = _ensureRecord(_selectedTaskKey);
      const swapWith = action === 'ta-move-up' ? idx - 1 : idx + 1;
      if (swapWith < 0 || swapWith >= r[field].length) return;
      const tmp = r[field][idx];
      r[field][idx] = r[field][swapWith];
      r[field][swapWith] = tmp;
      _renderFormPanel();
      return;
    }

    if (action === 'ta-clear-analysis') {
      _clearOneTaskAnalysis(btn.getAttribute('data-task-key'));
      return;
    }
  });

  // Typing: write into appState immediately but do NOT re-render —
  // re-rendering on every keystroke would rebuild the input and steal
  // the caret mid-word, exactly the bug renderer.js's list-editors
  // avoid elsewhere in this app.
  root.addEventListener('input', function (e) {
    const el     = e.target;
    const action = el.getAttribute && el.getAttribute('data-action');
    if (!_selectedTaskKey) return;

    if (action === 'ta-edit-item') {
      const field = el.getAttribute('data-field');
      const idx   = parseInt(el.getAttribute('data-index'), 10);
      const r = _ensureRecord(_selectedTaskKey);
      if (Array.isArray(r[field])) r[field][idx] = el.value;
      return;
    }
    if (action === 'ta-edit-text') {
      const field = el.getAttribute('data-field');
      const r = _ensureRecord(_selectedTaskKey);
      r[field] = el.value;
      return;
    }
  });

  // Status dot / pill catch up once the field loses focus. Capture
  // phase is required — blur/focusout on inputs doesn't bubble.
  root.addEventListener('blur', function (e) {
    const el = e.target;
    if (el.matches && el.matches('.ta-list-input, .ta-textarea') && _selectedTaskKey) {
      _touchStatus(_selectedTaskKey);
    }
  }, true);
}
