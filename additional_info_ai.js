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

const BACKEND_URL = 'https://dacum-ai-backend-production.up.railway.app';

// Maps the JSON keys the model returns → the textarea that receives
// them. Order here is also the order used in the overwrite warning.
const _FIELD_MAP = [
  { key: 'knowledge',  inputId: 'knowledgeInput',  label: 'Knowledge Requirements' },
  { key: 'skills',     inputId: 'skillsInput',     label: 'Skills Requirements' },
  { key: 'behaviors',  inputId: 'behaviorsInput',  label: 'Worker Behaviors/Traits' },
  { key: 'tools',      inputId: 'toolsInput',      label: 'Tools, Equipment, Supplies and Materials' },
  { key: 'trends',     inputId: 'trendsInput',     label: 'Future Trends and Concerns' },
  { key: 'acronyms',   inputId: 'acronymsInput',   label: 'Acronyms' },
  { key: 'careerPath', inputId: 'careerPathInput', label: 'Career Path' },
];

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
   - Typically 10–20 items

2. skills — Skills Requirements
   - TRANSFERABLE ABILITIES the work demands (technical + employability)
   - Short ability statements, e.g. "Interpret technical drawings"
   - Distinct from tasks: a skill is an underlying capability, a task is a
     discrete unit of work. Do NOT simply restate the chart's tasks here.
   - Typically 10–20 items

3. behaviors — Worker Behaviors/Traits
   - Personal attributes and work habits expected on the job
   - Short trait phrases, e.g. "Attention to detail", "Punctuality"
   - Typically 8–15 items

4. tools — Tools, Equipment, Supplies and Materials
   - Concrete, nameable items used to perform the tasks
   - Include category when useful, e.g. "Digital multimeter", "PPE: safety goggles"
   - Typically 12–25 items

5. trends — Future Trends and Concerns
   - Realistic developments affecting this occupation in the next 3–7 years
   - Technology, regulation, market, workforce, sustainability
   - Reflect the Country/Context and Sector when given
   - Typically 8–15 items

6. acronyms — Acronyms
   - Abbreviations that genuinely appear in this occupation
   - STRICT FORMAT: "ABC - Full Expansion" (one per item)
   - Only include acronyms you are confident are real and in use
   - Typically 6–15 items

7. careerPath — Career Path
   - Realistic progression for this occupation, entry level upward
   - STRICT FORMAT: "Level: Role title" e.g. "Entry Level: Apprentice Technician"
   - Order from entry to most senior
   - Typically 4–7 items

GENERAL RULES:
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
    showStatus(`❌ Daily limit reached (${usageStatus.count} generations). Try again tomorrow!`, 'error');
    return false;
  }

  const inputs = _readAIInputs();

  if (!inputs.occupationTitle) {
    alert('Please enter an Occupation Title in Chart Info to generate the supporting information.');
    showStatus('Occupation Title is required for AI generation.', 'error');
    return false;
  }

  // ── Overwrite guard — name exactly which sections are at risk ──
  const filled = _collectFilledFields();
  if (filled.length) {
    const names = filled.map(f => `  • ${f.label}`).join('\n');
    if (!confirm(
      '⚠️ AI GENERATION WILL REPLACE THE CONTENT OF THESE SECTIONS:\n\n' +
      names +
      '\n\nCustom sections you added yourself are NOT affected.\n\n' +
      'Click OK to continue, or Cancel to keep your current work.'
    )) {
      showStatus('AI generation cancelled. Your existing content is preserved.', 'error');
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
      body: JSON.stringify({ prompt })
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

    _FIELD_MAP.forEach(({ key, inputId }) => {
      const items = info[key];
      if (!Array.isArray(items) || items.length === 0) return;

      const lines = items
        .map(v => String(v == null ? '' : v).trim())
        // Strip any bullet/number the model added despite instructions
        .map(v => v.replace(/^[\s]*[•\-\*○●]\s*/, '').replace(/^[\s]*\d+[.)]\s*/, '').trim())
        .filter(Boolean);

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
    showStatus(
      `✓ Supporting information generated (${basis}) — ${filledCount} sections, ${itemCount} items.`,
      'success'
    );
    return true;

  } catch (error) {
    hideLoadingModal();
    console.error('Error generating Additional Information:', error);
    showStatus('AI generation failed. See the error dialog for details.', 'error');
    _showAIErrorModal(error.message || String(error));
    return false;
  }
}

// ── Error modal (self-contained, same look as the duties one) ──

function _showAIErrorModal(errorMessage) {
  const existing = document.getElementById('aiInfoErrorModal');
  if (existing) existing.remove();

  const overlay = document.createElement('div');
  overlay.id = 'aiInfoErrorModal';
  overlay.style.cssText =
    'position:fixed;inset:0;background:rgba(15,23,42,0.55);z-index:2147483000;' +
    'display:flex;align-items:center;justify-content:center;padding:20px;';

  const box = document.createElement('div');
  box.style.cssText =
    'background:#fff;border-radius:14px;max-width:520px;width:100%;padding:24px;' +
    'box-shadow:0 20px 60px rgba(0,0,0,0.3);font-family:inherit;';
  box.innerHTML = `
    <h3 style="margin:0 0 10px;color:#b91c1c;font-size:1.15em;">⚠️ AI generation failed</h3>
    <p style="margin:0 0 12px;color:#334155;font-size:0.92em;line-height:1.6;">
      The supporting information could not be generated. Your existing content
      has not been changed.
    </p>
    <pre style="margin:0 0 16px;padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;
                border-radius:8px;font-size:0.8em;color:#475569;white-space:pre-wrap;
                word-break:break-word;max-height:160px;overflow:auto;">${
      String(errorMessage).replace(/</g, '&lt;').replace(/>/g, '&gt;')
    }</pre>
    <div style="text-align:right;">
      <button id="aiInfoErrorClose"
              style="background:#667eea;color:#fff;border:none;border-radius:8px;
                     padding:9px 20px;font-weight:600;cursor:pointer;font-family:inherit;">
        Close
      </button>
    </div>
  `;

  overlay.appendChild(box);
  document.body.appendChild(overlay);

  const close = () => overlay.remove();
  box.querySelector('#aiInfoErrorClose').addEventListener('click', close);
  overlay.addEventListener('click', e => { if (e.target === overlay) close(); });
}
