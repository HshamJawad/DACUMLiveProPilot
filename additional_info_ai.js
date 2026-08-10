// ============================================================
// /additional_info_ai.js
// AI generation for the Additional Information tab.
//
// Mirrors the duties/tasks generator in projects.js:
//   • Same Railway backend + /api/generate-dacum endpoint (that
//     route is a generic "run this prompt" proxy — it isn't
//     duties-specific, so no backend change is needed here).
//   • Same usage-limit accounting (checkUsageLimit / incrementUsage).
//   • Same loading modal and status-toast conventions.
//   • Same grounding inputs: Occupation Title (required), plus
//     Job Title / Scope of Work / Sector / Context when present.
//
// One thing it does that the duties generator can't: when duties and
// tasks already exist, they are fed into the prompt as the primary
// evidence base. Knowledge, Skills, Tools and Behaviors are supposed
// to be *derived from* the tasks a worker actually performs — that's
// the DACUM logic — so generating them from the chart rather than
// from the occupation title alone produces far more defensible
// output. The chart is truncated (see _summariseChart) to keep the
// request within a sane size on large charts.
//
// Output is written straight into the seven fixed textareas of the
// Additional Information tab. Custom sections added by the user are
// deliberately NOT touched — their headings are user-defined and the
// model has no reliable way to know what belongs in them.
// ============================================================

import { appState }   from './state.js';
import { showStatus } from './renderer.js';
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

// Maps the JSON keys the model returns → the textarea that receives
// them. Order here is also the order used in the overwrite warning.
//
// `max` is a HARD cap enforced in code, not just in the prompt.
// Language models routinely overshoot soft counts, and this output is
// only ever a first draft for the facilitator — an over-long list
// costs review time in the workshop and buries the items that
// actually matter. The prompt asks for min–max; this backstop
// guarantees the ceiling. `min` is prompt-side only (you cannot
// invent missing items in code).
const _FIELD_MAP = [
  { key: 'knowledge',  inputId: 'knowledgeInput',  label: 'Knowledge Requirements',                 min: 10, max: 15 },
  { key: 'skills',     inputId: 'skillsInput',     label: 'Skills Requirements',                    min: 10, max: 15 },
  { key: 'behaviors',  inputId: 'behaviorsInput',  label: 'Worker Behaviors/Traits',                min:  8, max: 12 },
  { key: 'tools',      inputId: 'toolsInput',      label: 'Tools, Equipment, Supplies and Materials', min: 12, max: 18 },
  { key: 'trends',     inputId: 'trendsInput',     label: 'Future Trends and Concerns',             min:  6, max: 10 },
  { key: 'acronyms',   inputId: 'acronymsInput',   label: 'Acronyms',                               min:  6, max: 12 },
  { key: 'careerPath', inputId: 'careerPathInput', label: 'Career Path',                            min:  4, max:  6 },
];

/** Look up the configured range for a field key. */
function _range(key) {
  const f = _FIELD_MAP.find(x => x.key === key);
  return f ? { min: f.min, max: f.max } : { min: 5, max: 15 };
}

// ── Input readers ─────────────────────────────────────────────

function _readAIInputs() {
  return {
    occupationTitle: (document.getElementById('occupationTitle')?.value || '').trim(),
    jobTitle:        (document.getElementById('jobTitle')?.value        || '').trim(),
    scopeOfWork:     (document.getElementById('scopeOfWork')?.value     || '').trim(),
    sector:          (document.getElementById('sector')?.value          || '').trim(),
    context:         (document.getElementById('context')?.value         || '').trim(),
  };
}

/**
 * Condense the current chart into a compact text block for the prompt.
 * Returns '' when there is no usable chart yet, in which case the
 * generator falls back to occupation-level reasoning.
 */
function _summariseChart() {
  const duties = (appState.dutiesData || []).filter(
    d => (d.title || '').trim() || (d.tasks || []).some(t => (t.text || '').trim())
  );
  if (!duties.length) return '';

  // Caps keep very large charts from blowing up the request body.
  const MAX_DUTIES        = 14;
  const MAX_TASKS_PER_DUTY = 12;

  const lines = duties.slice(0, MAX_DUTIES).map(duty => {
    const tasks = (duty.tasks || [])
      .map(t => (t.text || '').trim())
      .filter(Boolean)
      .slice(0, MAX_TASKS_PER_DUTY);
    const title = (duty.title || '').trim() || 'Untitled duty';
    return tasks.length
      ? `- ${title}\n${tasks.map(t => `    · ${t}`).join('\n')}`
      : `- ${title}`;
  });

  return lines.join('\n');
}

/** True when at least one of the seven fields already has text. */
function _collectFilledFields() {
  return _FIELD_MAP.filter(f => {
    const el = document.getElementById(f.inputId);
    return el && el.value.trim().length > 0;
  });
}

// ── Prompt builder ────────────────────────────────────────────

function _buildPrompt(inputs, chartSummary) {
  const { occupationTitle, jobTitle, scopeOfWork, sector, context } = inputs;

  return `You are an occupational analysis engine specialized in DACUM methodology.
Your task is to generate the SUPPORTING INFORMATION sections of a DACUM chart.
The output will be injected directly into a DACUM chart UI.

INPUT:
Occupation Title (BASE CONTEXT): ${occupationTitle}${jobTitle ? `
Job / Role (PRIMARY FOCUS): ${jobTitle}` : ''}${scopeOfWork ? `
Scope of Work (CRITICAL BOUNDARY): ${scopeOfWork}` : ''}${sector ? `
Sector: ${sector}` : ''}${context ? `
Country / Context: ${context}` : ''}
${chartSummary ? `
EXISTING DACUM CHART (PRIMARY EVIDENCE BASE):
${chartSummary}
` : ''}
SCOPE INTERPRETATION RULE (VERY IMPORTANT):
- If Scope of Work is provided → it DEFINES and LIMITS the analysis.
- If Job Title is provided → generate for that specific job within the occupation.
- If Job Title is NOT provided → assume a generic role within the occupation,
  but STRICTLY guided by the Scope if available.
${chartSummary ? `- The chart above lists the REAL WORK already agreed for this job.
  Every item you generate must be traceable to those duties and tasks.
  Do NOT introduce knowledge, skills or tools for work that is not in the chart.
` : ''}
TASK:
Generate the following seven sections.

1. knowledge — Knowledge Requirements
   - WHAT THE WORKER MUST KNOW (cognitive, theoretical, regulatory)
   - Noun phrases, e.g. "Principles of hydraulic pressure", "Local electrical code"
   - NOT actions, NOT tasks
   - COUNT: minimum ${_range('knowledge').min}, maximum ${_range('knowledge').max} items

2. skills — Skills Requirements
   - TRANSFERABLE ABILITIES the work demands (technical + employability)
   - Short ability statements, e.g. "Interpret technical drawings"
   - Distinct from tasks: a skill is an underlying capability, a task is a
     discrete unit of work. Do NOT simply restate the chart's tasks here.
   - COUNT: minimum ${_range('skills').min}, maximum ${_range('skills').max} items

3. behaviors — Worker Behaviors/Traits
   - Personal attributes and work habits expected on the job
   - Short trait phrases, e.g. "Attention to detail", "Punctuality"
   - COUNT: minimum ${_range('behaviors').min}, maximum ${_range('behaviors').max} items

4. tools — Tools, Equipment, Supplies and Materials
   - MAIN CATEGORIES AND KEY ITEMS ONLY — this is a facilitator's draft,
     NOT a procurement inventory. A real workplace may use hundreds of
     items; list only what is characteristic of this occupation.
   - Group related consumables rather than listing them one by one
     (e.g. "Fasteners: screws, bolts, anchors" as ONE item, not three)
   - Prefer items that appear in, or are clearly implied by, the tasks
   - Include category when useful, e.g. "Digital multimeter", "PPE: safety goggles"
   - COUNT: minimum ${_range('tools').min}, maximum ${_range('tools').max} items

5. trends — Future Trends and Concerns
   - Realistic developments affecting this occupation in the next 3–7 years
   - Technology, regulation, market, workforce, sustainability
   - Reflect the Country/Context and Sector when given
   - COUNT: minimum ${_range('trends').min}, maximum ${_range('trends').max} items

6. acronyms — Acronyms
   - Abbreviations that genuinely appear in this occupation
   - STRICT FORMAT: "ABC - Full Expansion" (one per item)
   - Only include acronyms you are confident are real and in use
   - COUNT: minimum ${_range('acronyms').min}, maximum ${_range('acronyms').max} items

7. careerPath — Career Path
   - Realistic progression for this occupation, entry level upward
   - STRICT FORMAT: "Level: Role title" e.g. "Entry Level: Apprentice Technician"
   - Order from entry to most senior
   - COUNT: minimum ${_range('careerPath').min}, maximum ${_range('careerPath').max} items

GENERAL RULES:
- COUNT LIMITS ARE MANDATORY, not suggestions. Never exceed a maximum.
  This output is a STARTING DRAFT for a DACUM facilitator to review with
  an expert panel — not an exhaustive reference. A shorter, sharper list
  of the items that genuinely characterise the occupation is far more
  useful than a long list padded with generic or marginal entries.
- If you cannot reach a minimum with genuinely relevant items, return
  fewer rather than padding with filler.
- Every item is a SINGLE LINE of plain text.
- Do NOT prefix items with bullets, dashes, or numbers — the UI applies
  its own formatting.
- Keep each item concise (under about 12 words).
- No duplicates within a section.
- Stay INSIDE the defined scope; prefer specificity over completeness.
- Be data-informed and realistic for the given sector and country context.
- Use the same language as the Occupation Title input.

OUTPUT FORMAT (STRICT – NO EXTRA TEXT):
Return ONLY valid JSON using the following structure:

{
  "knowledge":  ["item", "item"],
  "skills":     ["item", "item"],
  "behaviors":  ["item", "item"],
  "tools":      ["item", "item"],
  "trends":     ["item", "item"],
  "acronyms":   ["ABC - Full Expansion"],
  "careerPath": ["Entry Level: Role title"]
}

Generate the supporting information now in valid JSON format only.`;
}

// ── Public entry point ────────────────────────────────────────

/**
 * generateAdditionalInfoAI()
 * Validates, prompts, calls the backend, and fills the seven
 * Additional Information textareas. Returns true on success.
 */
export async function generateAdditionalInfoAI() {
  // ── Usage limit (shared budget with the duties generator) ──
  const usageStatus = checkUsageLimit();
  if (!usageStatus.allowed) {
    showStatus('❌ ' + _tf('msgDailyLimit', { n: usageStatus.count }), 'error');
    return false;
  }

  const inputs = _readAIInputs();

  if (!inputs.occupationTitle) {
    alert('Please enter an Occupation Title in Chart Info to generate the supporting information.');
    showStatus(_t('msgOccupationRequired'), 'error');
    return false;
  }

  // ── Overwrite guard — name exactly which sections are at risk ──
  const filled = _collectFilledFields();
  /* The Full Draft run asks about overwriting ONCE, up front, naming
     every tab at stake. Re-asking here would mean four or five
     dialogs during a run the user has already authorised — and each
     one silently stalls the pipeline until someone notices. */
  if (!isBatchRun() && filled.length) {
    const names = filled.map(f => `  • ${f.label}`).join('\n');
    if (!confirm(
      '⚠️ AI GENERATION WILL REPLACE THE CONTENT OF THESE SECTIONS:\n\n' +
      names +
      '\n\nCustom sections you added yourself are NOT affected.\n\n' +
      'Click OK to continue, or Cancel to keep your current work.'
    )) {
      showStatus(_t('msgCancelAddInfo'), 'error');
      return false;
    }
  }

  const chartSummary = _summariseChart();

  showLoadingModal();
  await new Promise(resolve => setTimeout(resolve, 100));

  const prompt = _buildPrompt(inputs, chartSummary);

  try {
    const response = await fetch(`${BACKEND_URL}/api/generate-dacum`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ prompt: prompt + _aiDir() })
    });

    if (!response.ok) {
      throw new Error(`Backend request failed: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    if (!data.content || !data.content[0] || !data.content[0].text) {
      throw new Error('Invalid response from backend - no content found');
    }

    const jsonText = data.content[0].text.trim()
      .replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

    let info;
    try { info = JSON.parse(jsonText); }
    catch (e) { throw new Error('Failed to parse AI response as JSON'); }

    // ── Write into the textareas ────────────────────────────
    // Partial responses are tolerated: a section the model omitted
    // is left untouched rather than being blanked out.
    let filledCount = 0;
    let itemCount   = 0;
    let trimmedAny  = false;

    _FIELD_MAP.forEach(({ key, inputId, max }) => {
      const items = info[key];
      if (!Array.isArray(items) || items.length === 0) return;

      let lines = items
        .map(v => String(v == null ? '' : v).trim())
        // Strip any bullet/number the model added despite instructions
        .map(v => v.replace(/^[\s]*[•\-\*○●]\s*/, '').replace(/^[\s]*\d+[.)]\s*/, '').trim())
        .filter(Boolean);

      // Drop case-insensitive duplicates before applying the cap, so a
      // repeated item never consumes one of the allotted slots.
      const seen = new Set();
      lines = lines.filter(v => {
        const k = v.toLowerCase();
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
      });

      // HARD CAP — the prompt states the maximum, this enforces it.
      // Items are kept in the model's own order, which puts the most
      // characteristic entries first.
      if (max && lines.length > max) {
        lines = lines.slice(0, max);
        trimmedAny = true;
      }

      if (!lines.length) return;

      const el = document.getElementById(inputId);
      if (!el) return;

      el.value = lines.join('\n');
      filledCount++;
      itemCount += lines.length;
    });

    if (filledCount === 0) {
      throw new Error('AI response contained no usable sections');
    }

    hideLoadingModal();
    incrementUsage();

    const basis = chartSummary
      ? 'derived from your duties & tasks'
      : 'based on the occupation details';
    const trimNote = trimmedAny
      ? ' Long lists were trimmed to the top items — expand them with your panel.'
      : '';
    showStatus(
      '✓ ' + _tf('msgAddInfoGenerated',
        { basis: basis, sections: filledCount, items: itemCount }) + trimNote,
      'success'
    );
    return true;

  } catch (error) {
    hideLoadingModal();
    console.error('Error generating Additional Information:', error);
    showStatus(_t('msgAIFailed'), 'error');
    _showAIErrorModal(error.message || String(error));
    return false;
  }
}

// ── Error modal ───────────────────────────────────────────────
// Same structure, colours and animations as the duties-tab modal in
// projects.js (_showAIErrorModal), so both generators fail in a way
// the user already recognises. Only the copy differs: the offline
// fallback advice points at the Additional Information workflow
// instead of "add duties and tasks manually".

function _showAIErrorModal(errorMessage) {
  const existing = document.getElementById('aiInfoErrorModal');
  if (existing) existing.remove();

  const isOffline = /Failed to fetch|NetworkError|network|ECONNREFUSED|ERR_CONNECTION|ERR_NAME_NOT_RESOLVED|503|502/i.test(errorMessage);

  const modal = document.createElement('div');
  modal.id = 'aiInfoErrorModal';
  modal.setAttribute('role', 'alertdialog');
  modal.setAttribute('aria-modal', 'true');
  /* Appended to <body> with inline styles only, so the RTL sheet
     cannot reach it — direction is set here or Arabic renders
     left-aligned with its punctuation on the wrong side. */
  modal.setAttribute('dir', (window.i18n && window.i18n.isRTL()) ? 'rtl' : 'ltr');

  modal.style.cssText =
    'position:fixed;inset:0;z-index:999999;display:flex;align-items:center;' +
    'justify-content:center;padding:20px;background:rgba(0,0,0,0.55);' +
    'backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px);' +
    'animation:aiErrFadeIn 0.2s ease';

  const icon   = isOffline ? '\uD83D\uDD0C' : '\u26A0\uFE0F';
  const title  = _t(isOffline ? 'aiErrOfflineTitle' : 'aiErrInfoTitle');
  const sub    = _t(isOffline ? 'aiErrOfflineSub'   : 'aiErrFailedSub');
  const hdrBg  = isOffline
    ? 'linear-gradient(135deg,#fff7ed,#ffedd5)'
    : 'linear-gradient(135deg,#fef2f2,#fee2e2)';
  const hdrBdr = isOffline ? '#fed7aa' : '#fecaca';
  const hdrClr = isOffline ? '#9a3412' : '#991b1b';
  const subClr = isOffline ? '#c2410c' : '#b91c1c';

  /* Two branches, both translated. The raw error string stays as
     it came from the browser: it is a diagnostic for whoever is
     debugging, and translating an exception message would make it
     unsearchable. It is isolated LTR by the stylesheet. */
  const bodyText = isOffline
    ? _t('aiErrOfflineBody') + '<br><br>' + _t('aiErrSafeInfo')
    : _t('aiErrOccurred') + '<br><br>' +
      '<code style="font-size:0.82em;background:#f1f5f9;padding:4px 8px;' +
      'border-radius:4px;word-break:break-all;direction:ltr;' +
      'unicode-bidi:isolate;display:inline-block;">' +
      (errorMessage || '').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</code>' +
      '<br><br>' + _t('aiErrSafeInfo');

  const offlineTips = isOffline
    ? '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;' +
      'padding:12px 14px;margin-bottom:16px;">' +
      '<p style="margin:0;font-size:0.82em;color:#15803d;font-weight:600;">' +
      '\u2705 ' + _t('aiErrWhatInstead') + '</p>' +
      '<ul style="margin:6px 0 0;padding-left:18px;font-size:0.82em;color:#166534;line-height:1.8;">' +
      '<li>' + _t('aiTipInfo1') + '</li>' +
      '<li>' + _t('aiTipInfo2') + '</li>' +
      '<li>' + _t('aiTipInfo3') + '</li></ul></div>'
    : '';

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
        offlineTips +
        '<div style="display:flex;justify-content:flex-end;">' +
          '<button id="aiInfoErrorModalClose" style="padding:9px 22px;background:#667eea;' +
          'color:#fff;border:none;border-radius:8px;font-size:0.9em;font-weight:700;' +
          'cursor:pointer;transition:background 0.15s;"' +
          ' onmouseover="this.style.background=\'#5a67d8\'"' +
          ' onmouseout="this.style.background=\'#667eea\'">' + _t('btnGotIt') + '</button>' +
        '</div>' +
      '</div>' +
    '</div>';

  // Reuses the keyframes injected by either generator — whichever
  // runs first creates them; the id guard prevents duplicates.
  if (!document.getElementById('aiErrStyles')) {
    const s = document.createElement('style');
    s.id = 'aiErrStyles';
    s.textContent =
      '@keyframes aiErrFadeIn  { from{opacity:0} to{opacity:1} }' +
      '@keyframes aiErrSlideIn { from{transform:translateY(-14px);opacity:0}' +
      ' to{transform:translateY(0);opacity:1} }';
    document.head.appendChild(s);
  }

  document.body.appendChild(modal);

  function _close() { modal.remove(); }
  document.getElementById('aiInfoErrorModalClose').addEventListener('click', _close);
  modal.addEventListener('click', function (e) { if (e.target === modal) _close(); });
  document.addEventListener('keydown', function _esc(e) {
    if (e.key === 'Escape') { _close(); document.removeEventListener('keydown', _esc); }
  });
}
