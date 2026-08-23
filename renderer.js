// ============================================================
// /renderer.js
// Utility functions and Skills Level matrix renderer.
// Also: addCustomSection, toggleEditHeading, clearSection, formatList.
// ============================================================

import { appState, defaultSkillsLevelData, skillsLevelIsEmpty } from './state.js';

/* i18n access — resolved lazily, see duties.js for the reasoning. */
const _t  = (k)    => (window.i18n ? window.i18n.t(k)     : k);
const _tf = (k, v) => (window.i18n ? window.i18n.tf(k, v) : k);


// ── Status / Utility ──────────────────────────────────────────

export function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  if (!statusDiv) return;
  statusDiv.textContent = message;
  statusDiv.className = `status ${type}`;
  statusDiv.style.display = 'block';
  if (type === 'success') {
    setTimeout(() => { statusDiv.style.display = 'none'; }, 3000);
  }
}

export function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/* toggleInfoBox() was removed in 3.20.1.

   It drove an info box that no longer exists: #infoBoxContent,
   .btn-toggle-info and #btnToggleInfoBox are all absent from
   index.html. The only reason it never threw is that _on() in
   events.js skips a missing element, so the click handler was never
   attached and the function was never reached.

   Had it ever been reached it would have thrown immediately on
   `infoBoxContent.style` — null. Translating its hard-coded 'Hide' /
   'Show' would have meant adding two i18n keys to keep unreachable
   code tidy; deleting it is the actual fix. */

/* ── Section action icons ─────────────────────────────────────
   Inline SVG rather than emoji: 🔢 and ✏️ render as a different
   picture on every platform, cannot take the button's colour, and
   sit on the text baseline instead of centring. These use
   fill="currentColor", so they follow the button through hover,
   focus and the pressed state for free.

   Declared here so the custom sections built by addCustomSection()
   are identical to the seven static ones in index.html. */
const _ICON_LINES =
  '<rect x="5.8" y="3.1" width="9.2" height="1.5" rx=".75"/>' +
  '<rect x="5.8" y="7.25" width="9.2" height="1.5" rx=".75"/>' +
  '<rect x="5.8" y="11.4" width="9.2" height="1.5" rx=".75"/>';

const ICON_BULLET =
  '<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true" focusable="false">' +
  '<rect x="1" y="2.6" width="2.8" height="2.8" rx=".6"/>' +
  '<rect x="1" y="6.75" width="2.8" height="2.8" rx=".6"/>' +
  '<rect x="1" y="10.9" width="2.8" height="2.8" rx=".6"/>' + _ICON_LINES + '</svg>';

const ICON_NUMBER =
  '<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true" focusable="false">' +
  '<text x="0.4" y="5.35" font-size="4.7" font-weight="700" font-family="sans-serif">1</text>' +
  '<text x="0.4" y="9.5" font-size="4.7" font-weight="700" font-family="sans-serif">2</text>' +
  '<text x="0.4" y="13.65" font-size="4.7" font-weight="700" font-family="sans-serif">3</text>' + _ICON_LINES + '</svg>';

/* An I-beam text cursor, not a pencil. A pencil reads as "edit the
   content"; this button edits the HEADING, and the I-beam is the
   established affordance for entering text-edit mode. */
const ICON_RENAME =
  '<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true" focusable="false">' +
  '<rect x="7.1" y="2" width="1.8" height="12" rx=".4"/>' +
  '<rect x="4.2" y="1.6" width="7.6" height="1.7" rx=".6"/>' +
  '<rect x="4.2" y="12.7" width="7.6" height="1.7" rx=".6"/></svg>';

// ── Skills Level Matrix ───────────────────────────────────────

/* Checkbox labels reuse the SAME keys as the explanatory list in the
   info box above the matrix. Duplicating the wording would let the
   legend and the checkboxes drift apart in translation. */
const LEVEL_KEYS = {
  craftsman:   'lvlCraftsman',
  skilled:     'lvlSkilled',
  semiSkilled: 'lvlSemiSkilled',
  foundation:  'lvlFoundation'
};


export function toggleSkillsLevelSection() {
  const header  = document.querySelector('.skills-level-header');
  const content = document.getElementById('skillsLevelContent');
  header.classList.toggle('active');
  content.classList.toggle('active');
}

export function addSkillsCategory() {
  const newId = appState.skillsLevelData.length + 1;
  appState.skillsLevelData.push({
    id: newId, category: '',
    competencies: [
      { id: `${newId}.1`, text: '', levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]
  });
  renderSkillsLevel();
}

export function removeSkillsCategory(categoryIndex) {
  if (appState.skillsLevelData.length <= 1) {
    alert(_t('msgMinOneCategory'));
    return;
  }
  if (confirm(_t('confirmRemoveCategory'))) {
    appState.skillsLevelData.splice(categoryIndex, 1);
    renderSkillsLevel();
  }
}

export function updateSkillsCategoryName(categoryIndex, name) {
  appState.skillsLevelData[categoryIndex].category = name;
}

export function addSkillsCompetency(categoryIndex) {
  const category    = appState.skillsLevelData[categoryIndex];
  const categoryId  = category.id;
  const newNum      = category.competencies.length + 1;
  category.competencies.push({
    id: `${categoryId}.${newNum}`, text: '',
    levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false }
  });
  renderSkillsLevel();
}

export function removeSkillsCompetency(categoryIndex, competencyIndex) {
  const category = appState.skillsLevelData[categoryIndex];
  if (category.competencies.length <= 1) {
    alert(_t('msgMinOneCompetency'));
    return;
  }
  category.competencies.splice(competencyIndex, 1);
  category.competencies.forEach((comp, index) => {
    comp.id = `${category.id}.${index + 1}`;
  });
  renderSkillsLevel();
}

export function updateSkillsCompetencyText(categoryIndex, competencyIndex, text) {
  appState.skillsLevelData[categoryIndex].competencies[competencyIndex].text = text;
}

export function handleSkillsLevelChange(categoryIndex, competencyIndex, level, isChecked) {
  appState.skillsLevelData[categoryIndex].competencies[competencyIndex].levels[level] = isChecked;
}

export function resetSkillsLevel(withConfirm = true) {
  // "Already at defaults" means no tick anywhere and no user-added rows.
  // Resetting that changes nothing, so it needs no warning.
  if (withConfirm) {
    const untouched = !(appState.skillsLevelData || []).some(cat =>
      (cat.competencies || []).some(c => Object.values(c.levels || {}).some(Boolean))
    );
    if (untouched) {
      showStatus(_t('msgSkillsAtDefaults'), 'success');
      return;
    }
    if (!confirm(_t('confirmResetSkills'))) return;
  }

  /* The 33 default strings used to be repeated here as English
     literals, duplicating state.js. They now come from one generator,
     which also means Reset regenerates the matrix in whatever language
     the interface is in RIGHT NOW — the one moment where re-resolving
     the wording is what the user actually asked for. */
  appState.skillsLevelData.length = 0;
  defaultSkillsLevelData().forEach(cat => appState.skillsLevelData.push(cat));
  renderSkillsLevel();
}

export function renderSkillsLevel() {
  const container = document.getElementById('skillsLevelContainer');
  if (!container) return;

  /* Seed on first render rather than at module load. state.js is
     evaluated as part of app.js's module graph, and although
     translations.js is a classic script that runs before it, seeding
     here keeps the matrix independent of that ordering AND lets
     storage.js load a saved project first without being overwritten.
     An empty array means a genuinely new matrix. */
  if (skillsLevelIsEmpty()) {
    defaultSkillsLevelData().forEach(cat => appState.skillsLevelData.push(cat));
  }

  let html = '';

  appState.skillsLevelData.forEach((category, categoryIndex) => {
    html += `
      <div class="skills-level-category">
        <div class="skills-level-category-header">
          <h4>${escapeHtml(_tf('expCategoryN', { n: category.id }))}</h4>
          <button class="btn-remove-category"
            data-action="remove-skills-category" data-cat-index="${categoryIndex}">${escapeHtml(_t('btnRemoveCategory'))}</button>
        </div>
        <input type="text" class="skills-level-category-name"
          placeholder="${escapeHtml(_t('phCategoryName'))}"
          value="${escapeHtml(category.category)}"
          data-action="update-skills-category-name" data-cat-index="${categoryIndex}">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;">
          <h5 style="margin:0;">${escapeHtml(_t('lblCompetencies'))}</h5>
          <button class="btn-add-competency"
            data-action="add-skills-competency" data-cat-index="${categoryIndex}">+ ${escapeHtml(_t('btnAddCompetency'))}</button>
        </div>
        <div>`;

    category.competencies.forEach((competency, competencyIndex) => {
      html += `
        <div class="skills-competency-row">
          <div class="skills-competency-input-row">
            <div class="skills-competency-id">${competency.id}:</div>
            <input type="text" class="skills-competency-text"
              placeholder="${escapeHtml(_t('phCompetencyText'))}"
              value="${escapeHtml(competency.text)}"
              data-action="update-skills-competency-text"
              data-cat-index="${categoryIndex}" data-comp-index="${competencyIndex}">
            <button class="btn-remove-competency"
              title="${escapeHtml(_t('ttRemoveCompetency'))}"
              data-action="remove-skills-competency"
              data-cat-index="${categoryIndex}" data-comp-index="${competencyIndex}">×</button>
          </div>
          <div class="skills-level-checkboxes">
            ${['craftsman','skilled','semiSkilled','foundation'].map(level => `
              <label class="skills-level-checkbox-label">
                <input type="checkbox" ${competency.levels[level] ? 'checked' : ''}
                  data-action="handle-skills-level-change"
                  data-cat-index="${categoryIndex}" data-comp-index="${competencyIndex}"
                  data-level="${level}">
                <span>${escapeHtml(_t(LEVEL_KEYS[level]))}</span>
              </label>`).join('')}
          </div>
        </div>`;
    });

    html += `</div></div>`;
  });

  container.innerHTML = html;
}

// ── Additional Info Helpers ────────────────────────────────────

export function toggleEditHeading(headingId) {
  const heading = document.getElementById(headingId);
  const isEditable = heading.getAttribute('contenteditable') === 'true';

  /* The rename button is a TOGGLE, and it is now icon-only — there is
     no text left to change, so the state has to be carried by
     aria-pressed. That drives the pressed styling in CSS and is also
     what a screen reader announces, which is the whole reason the
     button can afford to lose its label. */
  const btn = document.querySelector(
    `[data-action="toggle-edit-heading"][data-heading-id="${headingId}"]`
  );
  if (btn) btn.setAttribute('aria-pressed', isEditable ? 'false' : 'true');

  if (isEditable) {
    heading.setAttribute('contenteditable', 'false');
    heading.style.cursor = '';
    showStatus(_t('msgHeadingUpdated'), 'success');
  } else {
    heading.setAttribute('contenteditable', 'true');
    heading.focus();
    const range = document.createRange();
    range.selectNodeContents(heading);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }
}

export function clearSection(inputId, headingId, defaultHeading, headingKey) {
  /* headingKey is optional: markup added it as data-default-heading-key.
     When present the reset restores the heading in the CURRENT language;
     the English attribute remains the fallback for custom sections and
     for any caller that does not pass a key. */
  if (headingKey && window.i18n && window.i18n.has(headingKey)) {
    defaultHeading = _t(headingKey);
  }
  const current = (document.getElementById(inputId)?.value || '').trim();
  const heading = document.getElementById(headingId)?.textContent?.trim();
  const isDefaultHeading = !heading || heading === defaultHeading ||
    (headingKey && window.i18n && window.i18n.has(headingKey) &&
     heading === window.i18n.t(headingKey));
  if (!current && isDefaultHeading) {
    showStatus(_t('msgSectionAlreadyEmpty'), 'success');
    return;
  }
  if (confirm(_t('confirmClearSection'))) {
    document.getElementById(inputId).value = '';
    document.getElementById(headingId).textContent = defaultHeading;
    document.getElementById(headingId).setAttribute('contenteditable', 'false');
    showStatus(_t('msgSectionCleared') + ' ✓', 'success');
  }
}

export function formatList(inputId, formatType) {
  const textarea = document.getElementById(inputId);
  const text = textarea.value.trim();
  if (!text) { showStatus(_t('msgNothingToFormat'), 'error'); return; }

  let lines = text.split('\n').filter(l => l.trim());
  lines = lines.map(line => {
    line = line.replace(/^[\s]*[•\-\*○●]\s*/, '');
    line = line.replace(/^[\s]*\d+[\.\)]\s*/, '');
    return line.trim();
  });

  let formatted = [];
  if (formatType === 'number') {
    lines.forEach((line, i) => formatted.push(`${i + 1}. ${line}`));
  } else if (formatType === 'bullet') {
    lines.forEach(line => formatted.push(`• ${line}`));
  }

  textarea.value = formatted.join('\n');
  showStatus(_t(formatType === 'number' ? 'msgFormattedNumbering' : 'msgFormattedBullets'), 'success');
}

export function addCustomSection() {
  appState.customSectionCounter++;
  const sectionId = `customSection${appState.customSectionCounter}`;
  const headingId = `${sectionId}Heading`;
  const inputId   = `${sectionId}Input`;

  const container = document.getElementById('customSectionsContainer');
  const sectionDiv = document.createElement('div');
  sectionDiv.className = 'section-container';
  sectionDiv.id = sectionId;
  sectionDiv.innerHTML = `
    <div class="section-header-editable">
      <h3 id="${headingId}" contenteditable="false">${_tf('lblCustomSection', { n: appState.customSectionCounter })}</h3>
      <div style="display:flex;gap:10px;align-items:center;">
        <button class="btn-rename btn-icon" data-action="toggle-edit-heading"
          data-heading-id="${headingId}" aria-pressed="false"
          title="${escapeHtml(_t('ttRenameHeading'))}"
          aria-label="${escapeHtml(_t('ttRenameHeading'))}">${ICON_RENAME}</button>
        <button class="btn-clear-section has-ico" data-action="clear-section"
          data-input-id="${inputId}" data-heading-id="${headingId}"
          data-default-heading="${_tf('lblCustomSection', { n: appState.customSectionCounter })}"><svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"><path d="M4 7h16"/><path d="M9.5 7V5.6A1.6 1.6 0 0 1 11.1 4h1.8a1.6 1.6 0 0 1 1.6 1.6V7"/><path d="M6.6 7l.75 11.6A1.7 1.7 0 0 0 9.05 20.2h5.9a1.7 1.7 0 0 0 1.7-1.6L17.4 7"/><path d="M10.3 11v5.4M13.7 11v5.4"/></svg> ${_t('btnClear')}</button>
        <button class="btn-remove-section has-ico" data-action="remove-custom-section" data-section-id="${sectionId}">
          <svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"><path d="M4 7h16"/><path d="M9.5 7V5.6A1.6 1.6 0 0 1 11.1 4h1.8a1.6 1.6 0 0 1 1.6 1.6V7"/><path d="M6.6 7l.75 11.6A1.7 1.7 0 0 0 9.05 20.2h5.9a1.7 1.7 0 0 0 1.7-1.6L17.4 7"/><path d="M10.3 11v5.4M13.7 11v5.4"/></svg> ${_t('btnRemove')}
        </button>
      </div>
    </div>
    <textarea id="${inputId}" placeholder="${_t('phCustomSection')}"></textarea>`;

  container.appendChild(sectionDiv);
  showStatus(_t('msgCustomSectionAdded') + ' ✓', 'success');
}

export function removeCustomSection(sectionId) {
  if (confirm(_t('confirmRemoveSection'))) {
    const section = document.getElementById(sectionId);
    if (section) { section.remove(); showStatus(_t('msgSectionRemoved'), 'success'); }
  }
}


/* ── Re-render on language change ────────────────────────────────────
   Custom sections added by the facilitator are generated here as
   innerHTML, so their Rename/Clear/Remove buttons and placeholder are
   outside applyTranslations()' reach — the same gap that froze the
   Add Duty button and the verification accordion.

   renderSkillsLevel() is the safe re-entry point: it rebuilds from
   appState, so selections are preserved. Custom SECTION headings are
   not rebuilt on purpose — they may carry a name the user typed, and
   their default text is handled by data-i18n-once in the markup. */
window.addEventListener('dacum:langchange', () => {
  if (document.getElementById('skillsLevelContainer')) renderSkillsLevel();
});
