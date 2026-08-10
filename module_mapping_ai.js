// ============================================================
// /module_mapping_ai.js
// Module generation for the Module Mapping tab.
//
// Two generators, deliberately different in kind:
//
//   1. generateOneModulePerOutcome() — LOCAL, instant, no network.
//      "One Learning Outcome -> one Module" is pure arithmetic: there
//      is no judgement to make, so routing it through a language model
//      would only add latency, consume the daily AI quota, break
//      offline, and risk a slightly different answer each run. It runs
//      in the browser and is always correct.
//
//   2. generateModulesAI() — grouping + naming + SEQUENCING.
//      "Which outcomes belong together in one training module?" is a
//      pedagogical judgement that needs to read the content, which is
//      exactly what the model is good at. It also names each module
//      properly (instead of "Module 1") and orders the modules from
//      foundational to advanced so the result is a teachable sequence,
//      not just a set of buckets.
//
// Manual grouping via the existing controls is untouched by both.
//
// Hard guarantees enforced in code, not just asked for in the prompt:
//   • Only EXISTING learning outcomes are used — any id the model
//     invents is discarded (_resolveOutcomeIds).
//   • Every outcome lands in exactly one module; anything the model
//     forgot is swept into a final "Additional Outcomes" module rather
//     than silently vanishing from the curriculum.
//   • Module sizes are clamped to MIN/MAX_LOS_PER_MODULE.
// ============================================================

import { appState }        from './state.js';
import { showStatus }      from './renderer.js';
import { renderModules, renderModuleLoList } from './modules.js';
import { checkUsageLimit, incrementUsage,
         showLoadingModal, hideLoadingModal } from './storage.js';
import { isBatchRun } from './draft_mode.js';


/* i18n access — resolved lazily; see duties.js for why. */
const _t  = (k)    => (window.i18n ? window.i18n.t(k)     : k);
const _tf = (k, v) => (window.i18n ? window.i18n.tf(k, v) : k);

/* Output-language directive for the generation backend. Appended at the
   ONE place this module builds a request, so any prompt added later is
   covered without having to remember. Empty string in English. */
const _aiDir = () => (window.i18n ? window.i18n.aiDirective() : '');


const BACKEND_URL = 'https://dacum-ai-backend-production.up.railway.app';

// Grouping bounds. Without these the model drifts to one of two
// useless extremes: a module per outcome (which is Mode 1 in disguise)
// or one giant module containing everything.
const MIN_LOS_PER_MODULE = 2;
const MAX_LOS_PER_MODULE = 6;

// ── Shared helpers ────────────────────────────────────────────

function _outcomes() {
  return (appState.learningOutcomesData?.outcomes) || [];
}

/** Human-readable text for an outcome: its own statement, or its criteria. */
function _outcomeText(o) {
  const stated = (o.statement || '').trim();
  if (stated) return stated;
  const crit = (o.linkedCriteria || [])
    .map(c => (c.text || '').trim())
    .filter(Boolean);
  return crit.length ? crit.join('; ') : '(no statement yet)';
}

/** Replace all modules with a new set, then refresh both panels. */
function _commitModules(modules) {
  const mm = appState.moduleMappingData;
  mm.modules = modules;
  mm.moduleCounter = modules.length;
  renderModuleLoList();
  renderModules();
}

/**
 * Confirm before discarding existing modules. Names the count so the
 * user knows exactly what is at stake rather than facing a generic
 * "are you sure?".
 */
function _confirmOverwrite() {
  const existing = appState.moduleMappingData?.modules || [];
  if (existing.length === 0) return true;
  /* The Full Draft run asks about overwriting ONCE, up front, naming
     every tab at stake. Re-asking here would mean four or five
     dialogs during a run the user has already authorised — and each
     one silently stalls the pipeline until someone notices. */
  if (isBatchRun()) return true;
  return confirm(
    `⚠️ This will replace the ${existing.length} module` +
    `${existing.length > 1 ? 's' : ''} you already have.\n\n` +
    `Learning Outcomes and Performance Criteria are NOT affected — ` +
    `only the module grouping is rebuilt.\n\n` +
    `Click OK to continue, or Cancel to keep your current modules.`
  );
}

// ── Mode 1: one module per outcome (local, instant) ───────────

export function generateOneModulePerOutcome() {
  const outcomes = _outcomes();

  if (outcomes.length === 0) {
    showStatus(_t('msgNoLOsYet'), 'error');
    return false;
  }
  if (!_confirmOverwrite()) {
    showStatus(_t('msgCancelModules'), 'error');
    return false;
  }

  const modules = outcomes.map((o, i) => ({
    id:    `module_${i + 1}`,
    title: `Module ${i + 1}: ${_shortTitle(o)}`,
    learningOutcomes: [o],
  }));

  _commitModules(modules);
  showStatus('✓ ' + _tf('msgModulesOnePerLO', { n: modules.length }), 'success');
  return true;
}

/** A compact module title derived from an outcome, for Mode 1. */
function _shortTitle(o) {
  const text = _outcomeText(o);
  if (text === '(no statement yet)') return o.number || 'Untitled';
  const words = text.split(/\s+/).slice(0, 7).join(' ');
  return words.length < text.length ? `${words}…` : words;
}

// ── Mode 2: AI grouping + naming + sequencing ─────────────────

function _buildPrompt(outcomes) {
  const occupation = (document.getElementById('occupationTitle')?.value || '').trim();
  const jobTitle   = (document.getElementById('jobTitle')?.value || '').trim();
  const scope      = (document.getElementById('scopeOfWork')?.value || '').trim();

  // Each outcome is listed with its cluster and criteria so the model
  // can group on substance rather than on wording alone.
  const list = outcomes.map(o => {
    const crit = (o.linkedCriteria || [])
      .map(c => `      · [Cluster ${c.clusterNumber}] ${(c.text || '').trim()}`)
      .filter(l => l.trim().length > 20);
    return `  - id: ${o.id}\n    outcome: ${_outcomeText(o)}` +
           (crit.length ? `\n    performance criteria:\n${crit.join('\n')}` : '');
  }).join('\n');

  const suggested = Math.max(2, Math.ceil(outcomes.length / MAX_LOS_PER_MODULE));

  return `You are a curriculum design engine specialized in competency-based training (CBT) derived from DACUM analysis.

OCCUPATION: ${occupation || '(not specified)'}${jobTitle ? `
JOB / ROLE: ${jobTitle}` : ''}${scope ? `
SCOPE OF WORK: ${scope}` : ''}

LEARNING OUTCOMES TO ORGANISE (${outcomes.length} total):
${list}

TASK:
Group these Learning Outcomes into training modules (units of competency),
name each module, and ORDER THE MODULES AS A TEACHING SEQUENCE.

GROUPING RULES:
- Group outcomes that share a common workflow, phase of work, or body of
  underpinning knowledge and skill.
- Each module must contain between ${MIN_LOS_PER_MODULE} and ${MAX_LOS_PER_MODULE} outcomes.
- Aim for roughly ${suggested} modules, adjusting where the content clearly justifies it.
- EVERY outcome id listed above must appear in exactly ONE module.
- Use ONLY the ids given above. Do NOT invent, merge, split or reword outcomes.

MODULE TITLE RULES:
- Name the module for the COMPETENCE it develops, not "Module 1".
- Format: Action-oriented noun phrase, e.g. "Preparing and Processing Materials".
- 3-8 words, specific to this occupation, no numbering (numbering is added by the app).

SEQUENCING RULES (IMPORTANT):
- Order the modules from FOUNDATIONAL to ADVANCED — the order they would
  be delivered in a training programme.
- Earlier modules should build the knowledge and skills that later
  modules assume: safety and preparation first, core production work
  next, then finishing, quality assurance, and complex or supervisory
  work last.
- Where two modules are independent, place the one with broader
  transferable value first.
- Return them ALREADY IN THAT ORDER — position in the array IS the sequence.
- For each module, give a one-sentence "rationale" explaining its place
  in the sequence (what it builds on, or what it prepares for).

OUTPUT FORMAT (STRICT — NO EXTRA TEXT, NO MARKDOWN):
{
  "modules": [
    {
      "title": "Preparing and Processing Materials",
      "rationale": "Establishes the material handling skills every later module depends on.",
      "outcomeIds": ["lo_1", "lo_4"]
    }
  ]
}

Return ONLY that JSON object.`;
}

/**
 * Map model-supplied ids back onto real outcome objects.
 * Unknown ids are dropped and duplicates ignored, so a hallucinated or
 * repeated id can never corrupt the curriculum structure.
 */
function _resolveOutcomeIds(ids, byId, alreadyUsed) {
  const resolved = [];
  (ids || []).forEach(rawId => {
    const id = String(rawId || '').trim();
    const outcome = byId[id];
    if (!outcome || alreadyUsed.has(id)) return;
    alreadyUsed.add(id);
    resolved.push(outcome);
  });
  return resolved;
}

export async function generateModulesAI() {
  const outcomes = _outcomes();

  if (outcomes.length === 0) {
    showStatus(_t('msgNoLOsYet'), 'error');
    return false;
  }
  if (outcomes.length < 2) {
    showStatus(_t('msgNeedTwoLOs'), 'error');
    return false;
  }

  const usageStatus = checkUsageLimit();
  if (!usageStatus.allowed) {
    showStatus('❌ ' + _tf('msgDailyLimit', { n: usageStatus.count }), 'error');
    return false;
  }

  if (!_confirmOverwrite()) {
    showStatus(_t('msgCancelModules'), 'error');
    return false;
  }

  showLoadingModal();
  await new Promise(r => setTimeout(r, 100));

  try {
    const response = await fetch(`${BACKEND_URL}/api/generate-dacum`, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ prompt: _buildPrompt(outcomes) + _aiDir() }),
    });

    if (!response.ok) {
      throw new Error(`Backend request failed: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    if (!data.content?.[0]?.text) {
      throw new Error('Invalid response from backend - no content found');
    }

    const jsonText = data.content[0].text.trim()
      .replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

    let parsed;
    try { parsed = JSON.parse(jsonText); }
    catch (e) { throw new Error('Failed to parse AI response as JSON'); }

    if (!Array.isArray(parsed.modules) || parsed.modules.length === 0) {
      throw new Error('AI response contained no modules');
    }

    // ── Rebuild against real data ──────────────────────────────
    const byId = {};
    outcomes.forEach(o => { byId[o.id] = o; });

    const used     = new Set();
    const modules  = [];
    let   trimmed  = 0;

    parsed.modules.forEach(m => {
      let members = _resolveOutcomeIds(m.outcomeIds, byId, used);

      // Hard cap — release the surplus so it can be picked up by the
      // sweep below instead of bloating one module.
      if (members.length > MAX_LOS_PER_MODULE) {
        members.slice(MAX_LOS_PER_MODULE).forEach(o => used.delete(o.id));
        members = members.slice(0, MAX_LOS_PER_MODULE);
        trimmed++;
      }
      if (members.length === 0) return;

      const title = String(m.title || '').trim() || 'Untitled Module';
      modules.push({
        id:    `module_${modules.length + 1}`,
        title: `Module ${modules.length + 1}: ${title}`,
        rationale: String(m.rationale || '').trim(),
        learningOutcomes: members,
      });
    });

    // ── Safety net: nothing may be lost ────────────────────────
    // An outcome the model skipped would otherwise disappear from the
    // curriculum without any warning — the most damaging failure mode
    // here, since the omission is invisible in the UI.
    const orphans = outcomes.filter(o => !used.has(o.id));
    if (orphans.length) {
      modules.push({
        id:    `module_${modules.length + 1}`,
        title: `Module ${modules.length + 1}: Additional Outcomes`,
        rationale: 'Outcomes not placed by the grouping — review and reassign as needed.',
        learningOutcomes: orphans,
      });
    }

    if (modules.length === 0) throw new Error('No valid modules could be built from the response');

    _commitModules(modules);
    hideLoadingModal();
    incrementUsage();

    const notes = [];
    if (orphans.length) notes.push(`${orphans.length} outcome${orphans.length > 1 ? 's' : ''} placed in a review module`);
    if (trimmed)        notes.push(`${trimmed} oversized module${trimmed > 1 ? 's' : ''} trimmed`);

    showStatus(
      '✓ ' + _tf('msgModulesSequenced', { n: modules.length }) +
      (notes.length ? ' ' + _tf('msgNotesSuffix', { notes: notes.join('; ') }) : ''),
      'success'
    );
    return true;

  } catch (error) {
    hideLoadingModal();
    console.error('Error generating modules:', error);
    showStatus(_t('msgAIFailed'), 'error');
    _showAIErrorModal(error.message || String(error));
    return false;
  }
}

// ── Error modal (same design as the other generators) ─────────

function _showAIErrorModal(errorMessage) {
  const existing = document.getElementById('mmAiErrorModal');
  if (existing) existing.remove();

  const isOffline = /Failed to fetch|NetworkError|network|ECONNREFUSED|ERR_CONNECTION|ERR_NAME_NOT_RESOLVED|503|502/i.test(errorMessage);

  const modal = document.createElement('div');
  modal.id = 'mmAiErrorModal';
  modal.setAttribute('role', 'alertdialog');
  modal.setAttribute('aria-modal', 'true');
  modal.style.cssText =
    'position:fixed;inset:0;z-index:999999;display:flex;align-items:center;' +
    'justify-content:center;padding:20px;background:rgba(0,0,0,0.55);' +
    'backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px);' +
    'animation:aiErrFadeIn 0.2s ease';

  const icon   = isOffline ? '\uD83D\uDD0C' : '\u26A0\uFE0F';
  const title  = isOffline ? 'AI Service Unavailable' : 'Module Generation Failed';
  const sub    = isOffline ? 'Backend server unreachable' : 'Check connection and try again';
  const hdrBg  = isOffline ? 'linear-gradient(135deg,#fff7ed,#ffedd5)'
                           : 'linear-gradient(135deg,#fef2f2,#fee2e2)';
  const hdrBdr = isOffline ? '#fed7aa' : '#fecaca';
  const hdrClr = isOffline ? '#9a3412' : '#991b1b';
  const subClr = isOffline ? '#c2410c' : '#b91c1c';

  const bodyText = isOffline
    ? 'The AI backend server is currently offline or unreachable.<br><br>' +
      'Your Learning Outcomes and existing modules have not been changed.'
    : 'An error occurred while grouping the Learning Outcomes:<br><br>' +
      '<code style="font-size:0.82em;background:#f1f5f9;padding:4px 8px;' +
      'border-radius:4px;word-break:break-all;">' +
      (errorMessage || '').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</code>' +
      '<br><br>Your existing modules have not been changed.';

  const tips =
    '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;' +
    'padding:12px 14px;margin-bottom:16px;">' +
    '<p style="margin:0;font-size:0.82em;color:#15803d;font-weight:600;">' +
    '\u2705 What you can do instead:</p>' +
    '<ul style="margin:6px 0 0;padding-left:18px;font-size:0.82em;color:#166534;line-height:1.8;">' +
    '<li>Use <strong>One Module per Outcome</strong> — it works offline</li>' +
    '<li>Select outcomes and group them manually below</li>' +
    '</ul></div>';

  modal.innerHTML =
    '<div style="background:#fff;border-radius:16px;max-width:420px;width:100%;' +
    'box-shadow:0 24px 60px rgba(0,0,0,0.35);overflow:hidden;' +
    'font-family:\'Segoe UI\',system-ui,sans-serif;animation:aiErrSlideIn 0.22s ease;">' +
      '<div style="padding:20px 22px 16px;display:flex;align-items:center;gap:12px;' +
      'background:' + hdrBg + ';border-bottom:1px solid ' + hdrBdr + ';">' +
        '<span style="font-size:1.8em;line-height:1;">' + icon + '</span>' +
        '<div>' +
          '<p style="margin:0;font-size:1em;font-weight:800;color:' + hdrClr + ';">' + title + '</p>' +
          '<p style="margin:2px 0 0;font-size:0.78em;color:' + subClr + ';">' + sub + '</p>' +
        '</div>' +
      '</div>' +
      '<div style="padding:18px 22px 20px;">' +
        '<p style="margin:0 0 16px;font-size:0.88em;color:#374151;line-height:1.6;">' + bodyText + '</p>' +
        tips +
        '<div style="display:flex;justify-content:flex-end;">' +
          '<button id="mmAiErrorClose" style="padding:9px 22px;background:#667eea;' +
          'color:#fff;border:none;border-radius:8px;font-size:0.9em;font-weight:700;' +
          'cursor:pointer;font-family:inherit;">Got it</button>' +
        '</div>' +
      '</div>' +
    '</div>';

  if (!document.getElementById('aiErrStyles')) {
    const st = document.createElement('style');
    st.id = 'aiErrStyles';
    st.textContent =
      '@keyframes aiErrFadeIn  { from{opacity:0} to{opacity:1} }' +
      '@keyframes aiErrSlideIn { from{transform:translateY(-14px);opacity:0}' +
      ' to{transform:translateY(0);opacity:1} }';
    document.head.appendChild(st);
  }

  document.body.appendChild(modal);
  const close = () => modal.remove();
  modal.querySelector('#mmAiErrorClose').addEventListener('click', close);
  modal.addEventListener('click', e => { if (e.target === modal) close(); });
  document.addEventListener('keydown', function esc(e) {
    if (e.key === 'Escape') { close(); document.removeEventListener('keydown', esc); }
  });
}
