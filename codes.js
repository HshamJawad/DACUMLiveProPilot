// ============================================================
// /codes.js
// Single source of truth for DACUM display codes (Duty A, Task A1).
//
// Design principles
// ─────────────────
// • IDs (duty.id, task.inputId) are IMMUTABLE — used by verification
//   ratings, taskMetadata, clustering, learning outcomes, module
//   mapping. They must never change across a session.
//
// • Codes (A, B, A1, A2 …) are DISPLAY-ONLY and are computed from
//   the *current* position of each duty/task inside appState.dutiesData.
//   They recompute automatically after any drag / delete / refine.
//
// • All rendering modules (duties.js, tasks.js, modules.js, exports.js)
//   import from here — zero duplication.
// ============================================================

import { appState } from './state.js';

/* i18n access — resolved lazily; see duties.js for why. */
const _t  = (k)    => (window.i18n ? window.i18n.t(k)     : k);
const _tf = (k, v) => (window.i18n ? window.i18n.tf(k, v) : k);

// ── Alphabets ────────────────────────────────────────────────
//
// The Latin sequence is A, B, C … the ordinal alphabet everyone reads
// as "first, second, third". Arabic has two orderings and only one of
// them carries that meaning:
//
//   • hija'i  (أ ب ت ث ج ح …) — the DICTIONARY order, grouped by
//     letter shape. Familiar, but it is not how Arabic numbers things.
//   • abjadi  (أ ب ج د ه و ز …) — the historical ORDINAL sequence,
//     still used exactly where English uses (a), (b), (c): legal
//     clauses, exam questions, outline levels.
//
// abjadi is therefore the correct choice: «الواجب ج» reads as "Duty
// three" to an Arabic reader in a way «الواجب ت» does not. 28 letters
// against Latin's 26, so an Arabic chart reaches two-letter codes
// slightly later.
const ABJAD = [
  'أ','ب','ج','د','ه','و','ز','ح','ط','ي',
  'ك','ل','م','ن','س','ع','ف','ص','ق','ر',
  'ش','ت','ث','خ','ذ','ض','ظ','غ'
];

// ── Letter generator ─────────────────────────────────────────
//
// 0 → A, 1 → B, … 25 → Z, 26 → AA, 27 → AB … 701 → ZZ.
// In Arabic: 0 → أ, 1 → ب … 27 → غ, 28 → أأ …
//
// DISPLAY ONLY. Codes are recomputed from each item's CURRENT position
// on every render and are never stored, so switching language changes
// what is drawn and nothing else — a project created in Arabic still
// opens with Latin codes in English, and no saved file is rewritten.
// This is the same rule applied to cluster and module names, and it is
// what makes the change safe: there is no data to migrate.

function _alphabet() {
  return (window.i18n && window.i18n.getLang() === 'ar') ? ABJAD : null;
}

export function getDutyLetter(index) {
  if (typeof index !== 'number' || index < 0 || !isFinite(index)) return '';
  const alpha = _alphabet();
  const base  = alpha ? alpha.length : 26;
  let n = Math.floor(index);
  let s = '';
  do {
    const d = n % base;
    s = (alpha ? alpha[d] : String.fromCharCode(65 + d)) + s;
    n = Math.floor(n / base) - 1;
  } while (n >= 0);
  return s;
}

// ── Position lookup inside live state ────────────────────────

/**
 * Find the current position of a duty inside appState.dutiesData.
 * Returns -1 if the duty is not present (e.g. deleted).
 */
export function getDutyIndex(dutyId) {
  const arr = appState.dutiesData || [];
  for (let i = 0; i < arr.length; i++) {
    if (arr[i] && arr[i].id === dutyId) return i;
  }
  return -1;
}

/**
 * Find the current position of a task within its duty.
 * Returns { dutyIndex, taskIndex } or null if not found.
 * dutyId can be omitted — we'll scan all duties.
 */
export function getTaskPosition(taskInputId, dutyId) {
  const arr = appState.dutiesData || [];
  for (let d = 0; d < arr.length; d++) {
    const duty = arr[d];
    if (!duty) continue;
    if (dutyId && duty.id !== dutyId) continue;
    const tasks = duty.tasks || [];
    for (let t = 0; t < tasks.length; t++) {
      if (tasks[t] && tasks[t].inputId === taskInputId) {
        return { dutyIndex: d, taskIndex: t };
      }
    }
  }
  return null;
}

// ── Public display helpers ───────────────────────────────────

/**
 * Letter code for a duty, e.g. "A", "B", "AA".
 * Returns '' if the duty is not in state.
 */
export function getDutyCode(dutyId) {
  const idx = getDutyIndex(dutyId);
  if (idx < 0) return '';
  return getDutyLetter(idx);
}

/**
 * Bare letter-number code for a task, e.g. "A1", "B3", "AA12".
 * Returns '' if the task is not in state.
 */
export function getTaskCodeShort(taskInputId) {
  const pos = getTaskPosition(taskInputId);
  if (!pos) return '';
  return `${getDutyLetter(pos.dutyIndex)}${pos.taskIndex + 1}`;
}

/**
 * Full task code with "Task " prefix, e.g. "Task A1".
 *
 * Note: this matches the historical signature used by modules.js and
 * exports.js, which render it as `<strong>${taskCode}:</strong> ...`.
 * Keep this form stable — callers depend on it.
 */
export function getTaskCode(taskInputId) {
  const short = getTaskCodeShort(taskInputId);
  return short ? _tf('lblTask', { code: short }) : '';
}

/**
 * Full task label — kept as an alias for getTaskCode for clarity at
 * call sites that prefer the word "label" over "code".
 */
export function getTaskLabel(taskInputId) {
  return getTaskCode(taskInputId);
}

/**
 * Full duty label with title, e.g. "Duty A: Planning".
 * Falls back gracefully to just "Duty A" if title is empty.
 */
export function getDutyLabel(dutyId, title) {
  const letter = getDutyCode(dutyId);
  if (!letter) return title || '';
  const t = (title || '').trim();
  const label = _tf('lblDuty', { code: letter });
  return t ? `${label}: ${t}` : label;
}
