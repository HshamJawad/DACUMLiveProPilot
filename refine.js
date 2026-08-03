// ============================================================
// /refine.js
// AI Result Refinement — non-destructive post-processing.
//
// Design rules:
//   • Only activates AFTER AI generation (markAiGenerated flag).
//   • Always pushes to history FIRST → fully undo-able with Ctrl+Z.
//   • Never silently discards user edits made after generation.
//   • Operates on appState.dutiesData directly, then re-renders once.
//
// Post-refine UX:
//   • Persistent summary card appended inside #refineResultsSection
//     that stays visible until user dismisses, re-runs refine, or
//     clears/loads a project.
//   • Temporary green highlight on any duty/task whose text was
//     mutated in-place (2500 ms), so the user can see exactly what
//     the refinement touched.  Deletions (duplicates/fragments) can
//     only be reported via the summary card since they're gone from
//     the DOM.
//
// ── Alignment with the AI generation prompt (projects.js) ─────
// The generation prompt (_runAIGeneration) enforces TASK RULES:
//   1. ONE clear occupational action verb per task (Verb + Object)
//   2. Only observable, hands-on work actions
//   3. NO outcomes / intentions ("to ensure", "in order to", ...)
//   4. NO learning/cognitive verbs (understand, learn, know, recognize)
//   5. NO administrative/managerial/policy verbs (comply, adhere,
//      manage, coordinate, supervise, report, ...)
//   6. NO tools/equipment/materials/knowledge as task content
//
// Refine.js is a deterministic, local (non-AI) pass — it must never
// invent or rewrite meaning.  So the rules above are handled in two
// different ways, matched to how safe an automated fix is:
//   • Rule 3 (outcome/intention clauses) → SAFE TO AUTO-FIX. Stripping
//     a trailing "to ensure X" clause never changes the verb+object
//     core of the task, so this remains an automatic transformation
//     (see _CLAUSE_PATTERNS, expanded below to match the prompt).
//   • Rules 1, 4, 5 (compound verbs, cognitive verbs, admin verbs) →
//     FLAGGED FOR HUMAN REVIEW ONLY. Rewriting these safely requires
//     judgment (e.g. splitting "Install and test the valve" into two
//     tasks, or deciding whether "Report daily readings" should stay
//     or be dropped). Auto-editing them risks silently changing the
//     trainer's intended meaning, which violates the "never silently
//     discard/alter user content" design rule. Instead they get a
//     distinct amber highlight + a "Needs review" section in the
//     summary card, so the trainer makes the final call.
// ============================================================

import { appState }                          from './state.js';
import { syncAllFromDOM, renderDutiesFromState } from './duties.js';
import { pushHistoryState }                  from './history.js';
import { showStatus }                        from './renderer.js';

// ── Module-level flag (pure UI state, not saved to appState) ──

let _aiGenerated = false;

/** Call immediately after a successful AI generation. */
export function markAiGenerated() {
  _aiGenerated = true;
  _showRefineSection();
}

/** Call on app init / clear-all / load-project to reset. */
export function clearAiGeneratedFlag() {
  _aiGenerated = false;
  _hideRefineSection();
  _removeSummaryCard();     // also clear any stale summary from previous run
}

// ── Visibility helpers ────────────────────────────────────────

function _showRefineSection() {
  const el = document.getElementById('refineResultsSection');
  if (el) {
    el.style.display = 'block';
    el.style.animation = 'refine-fade-in 0.35s ease';
  }
}

function _hideRefineSection() {
  const el = document.getElementById('refineResultsSection');
  if (el) el.style.display = 'none';
}

// ── Public entry point ────────────────────────────────────────

/**
 * refineResults()
 * Applies a set of soft, well-defined transformations to the
 * current dutiesData.  The whole operation is pushed onto the
 * undo stack before any mutation → one Ctrl+Z reverts it all.
 */
export function refineResults() {
  if (!_aiGenerated) return;

  // Flush any pending DOM edits into state first
  syncAllFromDOM();

  // Push BEFORE mutation → undo restores this exact snapshot
  pushHistoryState();

  // Capture text snapshot BEFORE mutation — used to compute highlights
  const beforeSnap = _captureDiffSnapshot();

  const stats = {
    trimmed:          0,
    periods:          0,
    clauses:          0,
    normalized:       0,
    duplicates:       0,
    fragments:        0,
    // Advisory-only counters (nothing below this line auto-edits text)
    flaggedCompound:  0,
    flaggedCognitive: 0,
    flaggedAdmin:     0,
  };

  // Advisory items collected for the "Needs review" section + amber highlight
  const reviewItems = [];

  appState.dutiesData = (appState.dutiesData || []).map(duty => {
    // ── Duty title ─────────────────────────────────────────
    const cleanTitle = _normalizeTitle(duty.title, stats);

    // ── Tasks ──────────────────────────────────────────────
    const seen = new Set();

    const cleanedTasks = duty.tasks
      .map(task => _cleanTask(task, stats))
      .filter(task => {
        // Remove fragments (fewer than 2 words after cleaning)
        const wordCount = task.text.trim().split(/\s+/).filter(Boolean).length;
        if (wordCount < 2) {
          stats.fragments++;
          return false;
        }
        // Remove exact duplicates within the same duty (case-insensitive)
        const key = task.text.trim().toLowerCase();
        if (seen.has(key)) {
          stats.duplicates++;
          return false;
        }
        seen.add(key);
        return true;
      })
      // Re-number tasks sequentially after filtering
      .map((task, i) => ({ ...task, num: i + 1 }));

    // ── Rule-alignment check (advisory, non-mutating) ───────
    // Runs on the final surviving task text and flags anything that
    // still violates the generation prompt's TASK RULES but can't be
    // safely auto-corrected (see header note).
    cleanedTasks.forEach(task => {
      const reasons = _detectRuleFlags(task.text);
      if (reasons.length) {
        reasons.forEach(r => {
          if (r.code === 'compound')  stats.flaggedCompound++;
          if (r.code === 'cognitive') stats.flaggedCognitive++;
          if (r.code === 'admin')     stats.flaggedAdmin++;
        });
        reviewItems.push({ inputId: task.inputId, dutyId: duty.id, text: task.text, reasons });
      }
    });

    return { ...duty, title: cleanTitle, tasks: cleanedTasks };
  });

  renderDutiesFromState();

  // Diff the after-state against the captured snapshot and apply
  // a temporary highlight to every changed element still on screen
  const changed = _computeChanges(beforeSnap);
  _applyRefinedHighlights(changed);

  // Amber "needs review" highlight — separate from the green fix highlight
  _applyReviewFlags(reviewItems);

  // Build the persistent summary card (replaces any previous one)
  _renderSummaryCard(stats, changed, reviewItems);

  // Legacy toast — kept for continuity; the card is the primary signal
  _reportStats(stats, reviewItems);
}

// ── Task-level cleaner ────────────────────────────────────────

function _cleanTask(task, stats) {
  let text = task.text;

  // 1. Trim surrounding whitespace
  const trimmed = text.trim();
  if (trimmed !== text) { stats.trimmed++; text = trimmed; }

  // 2. Remove trailing period (DACUM standard: no sentence-end punctuation)
  if (/[.。]$/.test(text)) {
    text = text.replace(/[.。]\s*$/, '').trimEnd();
    stats.periods++;
  }

  // 3. Strip result/purpose/intention clauses appended to tasks
  //    e.g. "Install valve to ensure flow" → "Install valve"
  //    Matches generation-prompt rule: "NO outcomes, NO intentions"
  const stripped = _stripResultClauses(text);
  if (stripped !== text) { stats.clauses++; text = stripped; }

  // 4. Capitalize first letter
  if (text.length > 0) {
    const cap = text[0].toUpperCase() + text.slice(1);
    if (cap !== text) { stats.normalized++; text = cap; }
  }

  // 5. Remove double spaces
  text = text.replace(/  +/g, ' ').trim();

  return { ...task, text };
}

// ── Title normalizer ──────────────────────────────────────────

function _normalizeTitle(title, stats) {
  if (!title) return title;
  let t = title.trim();
  // Remove trailing period from duty titles
  if (/[.。]$/.test(t)) { t = t.replace(/[.。]\s*$/, '').trimEnd(); stats.periods++; }
  // Capitalize first letter
  if (t.length > 0) {
    const cap = t[0].toUpperCase() + t.slice(1);
    if (cap !== t) { stats.normalized++; t = cap; }
  }
  return t;
}

// ── Result-clause patterns ────────────────────────────────────
//
// Strips purpose/result/intention phrases that DACUM convention and
// the generation prompt forbid on task statements (tasks describe
// WHAT is done, never WHY it's done). Kept in sync with the prompt's
// "NO outcomes, NO intentions (avoid 'to ensure', 'in order to', etc.)"
// rule — expanded beyond the original short list to cover the common
// intention/purpose connectors the prompt's "etc." implies.

const _CLAUSE_PATTERNS = [
  /,?\s+to ensure\b.*$/i,
  /,?\s+in order to\b.*$/i,
  /,?\s+in order that\b.*$/i,
  /,?\s+so that\b.*$/i,
  /,?\s+so as to\b.*$/i,
  /,?\s+for the purpose of\b.*$/i,
  /,?\s+with the aim of\b.*$/i,
  /,?\s+with a view to\b.*$/i,
  /,?\s+aiming to\b.*$/i,
  /,?\s+to prevent\b.*$/i,
  /,?\s+to maintain\b.*$/i,
  /,?\s+to achieve\b.*$/i,
  /,?\s+to verify\b.*$/i,
  /,?\s+to confirm\b.*$/i,
  /,?\s+to support\b.*$/i,
  /,?\s+to facilitate\b.*$/i,
  /,?\s+to improve\b.*$/i,
  /,?\s+to reduce\b.*$/i,
  /,?\s+to avoid\b.*$/i,
  /,?\s+to enable\b.*$/i,
  /,?\s+to allow\b.*$/i,
  /,?\s+to guarantee\b.*$/i,
  /,?\s+to help\b.*$/i,
  /,?\s+to meet\b.*$/i,
  /,?\s+to satisfy\b.*$/i,
  /,?\s+to comply with\b.*$/i,
];

function _stripResultClauses(text) {
  let result = text;
  for (const pattern of _CLAUSE_PATTERNS) {
    result = result.replace(pattern, '');
  }
  return result.trim();
}

// ── Rule-alignment detectors (advisory only — never mutate text) ──
//
// These mirror the three generation-prompt TASK RULES that cannot be
// safely auto-corrected without risking a meaning change:
//   • ONE verb only per task            → "compound" flag
//   • NO learning/cognitive verbs        → "cognitive" flag
//   • NO administrative/managerial verbs → "admin" flag

// A representative set of hands-on, observable DACUM action verbs used
// only to recognize a LIKELY second verb after "and" / "&" — i.e. a
// compound task. Heuristic, not exhaustive; false negatives are fine
// (the check is advisory), false positives are harmless (review-only).
const _ACTION_VERBS = new Set([
  'install','remove','replace','inspect','check','test','verify','measure',
  'cut','drill','grind','weld','solder','wire','connect','disconnect',
  'assemble','disassemble','clean','lubricate','tighten','loosen','adjust',
  'calibrate','align','mix','pour','load','unload','operate','drive','lift',
  'move','transport','store','label','mark','sort','pack','package','wrap',
  'stack','prepare','cook','bake','fry','chop','slice','peel','wash',
  'sterilize','sample','weigh','record','document','fill','apply','spray',
  'paint','coat','sand','polish','sew','stitch','fold','iron','plant',
  'harvest','water','fertilize','prune','feed','dig','excavate','lay',
  'build','construct','demolish','repair','service','troubleshoot',
  'diagnose','program','configure','update','backup','restore','scan',
  'print','bind','laminate','deliver','collect','count','attach','detach',
  'fasten','unfasten','position','secure',
]);

// Administrative / managerial / policy-oriented verbs the prompt bans
// as task-leading verbs (explicit examples from the prompt + close
// synonyms trainers commonly produce).
const _ADMIN_VERBS = [
  'comply','adhere','manage','coordinate','supervise','report','oversee',
  'administer','direct','delegate','authorize','audit','regulate','enforce',
];

// Learning / cognitive verbs the prompt bans as task-leading verbs
// (tasks must be observable actions, not internal mental states).
const _COGNITIVE_VERBS = [
  'understand','learn','know','recognize','recognise','comprehend',
  'appreciate','realize','realise','familiarize','familiarise',
  'acknowledge','memorize','memorise',
];

const _ADMIN_RE     = new RegExp(`^(${_ADMIN_VERBS.join('|')})\\b`, 'i');
const _COGNITIVE_RE = new RegExp(`^(${_COGNITIVE_VERBS.join('|')})\\b`, 'i');

/** Returns an array of { code, verb } flags for a single (already-cleaned) task text. */
function _detectRuleFlags(text) {
  const flags = [];
  if (!text) return flags;

  const adminMatch = text.match(_ADMIN_RE);
  if (adminMatch) flags.push({ code: 'admin', verb: adminMatch[1] });

  const cogMatch = text.match(_COGNITIVE_RE);
  if (cogMatch) flags.push({ code: 'cognitive', verb: cogMatch[1] });

  const andMatch = text.match(/\b(?:and|&)\s+([A-Za-z]+)\b/i);
  if (andMatch && _ACTION_VERBS.has(andMatch[1].toLowerCase())) {
    flags.push({ code: 'compound', verb: andMatch[1] });
  }

  return flags;
}

const _FLAG_LABELS = {
  admin:     verb => `Leads with a managerial/administrative verb ("${verb}") — DACUM tasks describe hands-on execution, not administration`,
  cognitive: verb => `Leads with a cognitive/learning verb ("${verb}") — task statements must be observable actions, not knowledge or understanding`,
  compound:  verb => `May contain more than one action verb ("…and ${verb}…") — DACUM tasks should use ONE verb only`,
};

// ── Diff snapshot + highlight plumbing ────────────────────────

function _captureDiffSnapshot() {
  const snap = { duties: {}, tasks: {} };
  (appState.dutiesData || []).forEach(duty => {
    snap.duties[duty.id] = duty.title || '';
    (duty.tasks || []).forEach(task => {
      snap.tasks[task.inputId] = task.text || '';
    });
  });
  return snap;
}

function _computeChanges(beforeSnap) {
  const changedDuties = [];
  const changedTasks  = [];
  (appState.dutiesData || []).forEach(duty => {
    const before = beforeSnap.duties[duty.id];
    if (before !== undefined && before !== (duty.title || '')) {
      changedDuties.push(duty.id);
    }
    (duty.tasks || []).forEach(task => {
      const b = beforeSnap.tasks[task.inputId];
      if (b !== undefined && b !== (task.text || '')) {
        changedTasks.push(task.inputId);
      }
    });
  });
  return { changedDuties, changedTasks };
}

function _applyRefinedHighlights({ changedDuties, changedTasks }) {
  const HL       = 'refined-highlight';
  const DURATION = 2500;
  const esc      = (typeof CSS !== 'undefined' && CSS.escape) ? CSS.escape : (s => s);

  changedTasks.forEach(inputId => {
    const el = document.querySelector(`[data-task-id="${esc(inputId)}"]`);
    if (!el) return;
    const target = el.closest('.dcv-task-card') || el.closest('.task-item') || el;
    target.classList.add(HL);
    setTimeout(() => {
      const still   = document.querySelector(`[data-task-id="${esc(inputId)}"]`);
      const stillTg = still ? (still.closest('.dcv-task-card') || still.closest('.task-item') || still) : null;
      if (stillTg) stillTg.classList.remove(HL);
    }, DURATION);
  });

  changedDuties.forEach(dutyId => {
    const el = document.querySelector(`[data-duty-id="${esc(dutyId)}"]`);
    if (!el) return;
    const target = el.closest('.dcv-duty-card') || el;
    target.classList.add(HL);
    setTimeout(() => {
      const still   = document.querySelector(`[data-duty-id="${esc(dutyId)}"]`);
      const stillTg = still ? (still.closest('.dcv-duty-card') || still) : null;
      if (stillTg) stillTg.classList.remove(HL);
    }, DURATION);
  });
}

/**
 * Amber "needs review" highlight for tasks that still violate a
 * generation-prompt TASK RULE but were NOT auto-edited (see header
 * note). Unlike the green fix highlight, this one does not expire on
 * a timer — it clears itself the moment the trainer edits that exact
 * field (self-contained listener), or on the next full re-render.
 */
function _applyReviewFlags(reviewItems) {
  const FLAG = 'refine-flag-review';
  const esc  = (typeof CSS !== 'undefined' && CSS.escape) ? CSS.escape : (s => s);

  reviewItems.forEach(({ inputId }) => {
    const el = document.querySelector(`[data-task-id="${esc(inputId)}"]`);
    if (!el) return;
    const target = el.closest('.dcv-task-card') || el.closest('.task-item') || el;
    target.classList.add(FLAG);

    const clear = () => target.classList.remove(FLAG);
    el.addEventListener('input', clear, { once: true });
  });
}

// ── Persistent summary card ───────────────────────────────────

function _removeSummaryCard() {
  const existing = document.getElementById('refineSummaryCard');
  if (existing) existing.remove();
}

function _renderSummaryCard(stats, changed, reviewItems) {
  const parent = document.getElementById('refineResultsSection');
  if (!parent) return;

  _removeSummaryCard();     // always replace, never stack

  const total = stats.trimmed + stats.periods + stats.clauses +
                stats.normalized + stats.duplicates + stats.fragments;
  const highlightCount = changed.changedDuties.length + changed.changedTasks.length;

  const card = document.createElement('div');
  card.id = 'refineSummaryCard';
  card.className = 'refine-summary-card';

  let fixedHtml = '';
  if (total === 0) {
    fixedHtml = `
      <div class="refine-summary-icon">✓</div>
      <div class="refine-summary-title">No automatic fixes needed</div>
      <div class="refine-summary-body">
        AI output already follows DACUM conventions — no trailing periods,
        no result clauses, no duplicate tasks, all statements well-formed.
      </div>
    `;
  } else {
    const bullets = [];
    if (stats.normalized) bullets.push(`${stats.normalized} capitalisation fix${stats.normalized > 1 ? 'es' : ''}`);
    if (stats.periods)    bullets.push(`${stats.periods} trailing period${stats.periods > 1 ? 's' : ''} removed`);
    if (stats.clauses)    bullets.push(`${stats.clauses} result/intention clause${stats.clauses > 1 ? 's' : ''} stripped`);
    if (stats.duplicates) bullets.push(`${stats.duplicates} duplicate task${stats.duplicates > 1 ? 's' : ''} removed`);
    if (stats.fragments)  bullets.push(`${stats.fragments} fragment task${stats.fragments > 1 ? 's' : ''} removed`);
    if (stats.trimmed)    bullets.push(`${stats.trimmed} whitespace fix${stats.trimmed > 1 ? 'es' : ''}`);

    const listHtml = bullets.map(b => `<li>${b}</li>`).join('');
    const hintHtml = highlightCount > 0
      ? `<div class="refine-summary-hint">💡 ${highlightCount} item${highlightCount > 1 ? 's' : ''} highlighted below (fades in ~2.5s)</div>`
      : '';

    fixedHtml = `
      <div class="refine-summary-icon">✨</div>
      <div class="refine-summary-title">Refinement applied</div>
      <ul class="refine-summary-list">${listHtml}</ul>
      ${hintHtml}
      <div class="refine-summary-undo">
        Press <kbd>Ctrl</kbd>+<kbd>Z</kbd> to undo all changes.
      </div>
    `;
  }

  // ── Advisory "needs review" section ─────────────────────────
  // Only rendered when at least one task still violates a TASK RULE
  // that can't be safely auto-fixed (compound verb / cognitive verb /
  // admin verb). Nothing here was changed on disk — it's guidance only.
  let reviewHtml = '';
  if (reviewItems.length > 0) {
    const MAX_SHOWN = 8;
    const shown = reviewItems.slice(0, MAX_SHOWN);
    const extra = reviewItems.length - shown.length;

    const itemsHtml = shown.map(item => {
      const reasonsHtml = item.reasons
        .map(r => `<li>${_FLAG_LABELS[r.code](r.verb)}</li>`)
        .join('');
      const preview = item.text.length > 90 ? item.text.slice(0, 90) + '…' : item.text;
      return `
        <div class="refine-review-item">
          <div class="refine-review-text">"${preview}"</div>
          <ul class="refine-review-reasons">${reasonsHtml}</ul>
        </div>`;
    }).join('');

    const moreHtml = extra > 0
      ? `<div class="refine-review-more">+ ${extra} more task${extra > 1 ? 's' : ''} flagged (amber highlight in the chart)</div>`
      : '';

    reviewHtml = `
      <div class="refine-review-block">
        <div class="refine-review-title">⚠️ ${reviewItems.length} task${reviewItems.length > 1 ? 's' : ''} need${reviewItems.length > 1 ? '' : 's'} manual review</div>
        <div class="refine-review-body">
          These weren't auto-changed to avoid altering meaning — please review
          and edit them directly (amber highlight in the chart).
        </div>
        ${itemsHtml}
        ${moreHtml}
      </div>`;
  }

  card.innerHTML = `
    <button class="refine-summary-close" title="Dismiss" aria-label="Dismiss">×</button>
    ${fixedHtml}
    ${reviewHtml}
  `;

  // Wire the × button — self-contained, no events.js change needed
  const closeBtn = card.querySelector('.refine-summary-close');
  if (closeBtn) closeBtn.addEventListener('click', () => card.remove());

  parent.appendChild(card);
}

// ── Status reporter ───────────────────────────────────────────

function _reportStats(stats, reviewItems) {
  const total = stats.trimmed + stats.periods + stats.clauses +
                stats.normalized + stats.duplicates + stats.fragments;

  const flaggedCount = reviewItems ? reviewItems.length : 0;
  const flagSuffix = flaggedCount > 0
    ? ` — ⚠ ${flaggedCount} task${flaggedCount > 1 ? 's' : ''} flagged for manual review.`
    : '';

  if (total === 0) {
    showStatus(
      flaggedCount > 0
        ? `✓ No automatic fixes needed.${flagSuffix}`
        : '✓ Refinement complete — tasks are already clean and well-formed!',
      'success'
    );
    return;
  }

  const parts = [];
  if (stats.normalized) parts.push(`${stats.normalized} capitalisation fix${stats.normalized > 1 ? 'es' : ''}`);
  if (stats.periods)    parts.push(`${stats.periods} trailing period${stats.periods > 1 ? 's' : ''} removed`);
  if (stats.clauses)    parts.push(`${stats.clauses} result clause${stats.clauses > 1 ? 's' : ''} stripped`);
  if (stats.duplicates) parts.push(`${stats.duplicates} duplicate task${stats.duplicates > 1 ? 's' : ''} removed`);
  if (stats.fragments)  parts.push(`${stats.fragments} fragment task${stats.fragments > 1 ? 's' : ''} removed`);
  if (stats.trimmed)    parts.push(`${stats.trimmed} whitespace fix${stats.trimmed > 1 ? 'es' : ''}`);

  showStatus(`✓ Refined: ${parts.join(' · ')}.${flagSuffix} — Use Ctrl+Z to undo all changes.`, 'success');
}
