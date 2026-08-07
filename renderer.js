// ============================================================
// /renderer.js
// Utility functions and Skills Level matrix renderer.
// Also: addCustomSection, toggleEditHeading, clearSection, formatList.
// ============================================================

import { appState } from './state.js';

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

export function toggleInfoBox() {
  const infoBoxContent = document.getElementById('infoBoxContent');
  const toggleButton   = document.querySelector('.btn-toggle-info');
  if (infoBoxContent.style.display === 'none') {
    infoBoxContent.style.display = 'block';
    toggleButton.textContent = 'Hide';
  } else {
    infoBoxContent.style.display = 'none';
    toggleButton.textContent = 'Show';
  }
}

// ── Skills Level Matrix ───────────────────────────────────────

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
    alert('At least one category is required.');
    return;
  }
  if (confirm('Are you sure you want to remove this category?')) {
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
    alert('At least one competency is required per category.');
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
      showStatus('Skills Level is already at its defaults — nothing to reset', 'success');
      return;
    }
    if (!confirm('Are you sure you want to reset all Skills Level data?')) return;
  }

  appState.skillsLevelData.length = 0;
  const defaults = [
    { id: 1, category: 'Communication', competencies: [
      { id: '1.1', text: 'Verbally communicate with others',  levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '1.2', text: 'Communicate with others in writing', levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 2, category: 'Teamwork', competencies: [
      { id: '2.1', text: 'Work within a team',                          levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '2.2', text: 'Solve disputes and negotiate with others',    levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '2.3', text: 'Defend rights at work',                       levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '2.4', text: 'Time and resource management',                levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '2.5', text: 'Make decisions',                              levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 3, category: 'Self-marketing', competencies: [
      { id: '3.1', text: 'CV writing',           levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '3.2', text: 'Job interviews',        levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '3.3', text: 'Presentation skills',  levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 4, category: 'Problem Solving', competencies: [
      { id: '4.1', text: 'Identify and analyse work problems',  levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '4.2', text: 'Solve problems at a work site',       levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '4.3', text: 'Evaluate results and make decisions', levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 5, category: 'Entrepreneurship', competencies: [
      { id: '5.1', text: 'Critical thinking',                                                     levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '5.2', text: 'Find/create small business idea project',                                levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '5.3', text: 'Prepare simple feasibility studies for their projects',                  levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '5.4', text: 'Prepare business plan of project to present to loans institutions',      levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '5.5', text: 'Managing, improving and developing their project',                       levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 6, category: 'Computer/ICT skills', competencies: [
      { id: '6.1', text: 'Use a computer', levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '6.2', text: 'Use internet',   levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 7, category: 'Foreign Languages', competencies: [
      { id: '7.1', text: 'Basic communication skills',                           levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '7.2', text: 'Use English technical terms related to construction',  levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 8, category: 'Mathematical Skills', competencies: [
      { id: '8.1', text: 'Perform basic measurement operations', levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '8.2', text: 'Perform mathematical operations',      levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]},
    { id: 9, category: '', competencies: [
      { id: '9.1', text: '', levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } },
      { id: '9.2', text: '', levels: { craftsman: false, skilled: false, semiSkilled: false, foundation: false } }
    ]}
  ];
  defaults.forEach(d => appState.skillsLevelData.push(d));
  renderSkillsLevel();
}

export function renderSkillsLevel() {
  const container = document.getElementById('skillsLevelContainer');
  if (!container) return;
  let html = '';

  appState.skillsLevelData.forEach((category, categoryIndex) => {
    html += `
      <div class="skills-level-category">
        <div class="skills-level-category-header">
          <h4>Category ${category.id}</h4>
          <button class="btn-remove-category"
            data-action="remove-skills-category" data-cat-index="${categoryIndex}">Remove Category</button>
        </div>
        <input type="text" class="skills-level-category-name"
          placeholder="e.g., Communication, Problem Solving, etc."
          value="${escapeHtml(category.category)}"
          data-action="update-skills-category-name" data-cat-index="${categoryIndex}">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;">
          <h5 style="margin:0;">Competencies</h5>
          <button class="btn-add-competency"
            data-action="add-skills-competency" data-cat-index="${categoryIndex}">+ Add Competency</button>
        </div>
        <div>`;

    category.competencies.forEach((competency, competencyIndex) => {
      html += `
        <div class="skills-competency-row">
          <div class="skills-competency-input-row">
            <div class="skills-competency-id">${competency.id}:</div>
            <input type="text" class="skills-competency-text"
              placeholder="Enter competency description"
              value="${escapeHtml(competency.text)}"
              data-action="update-skills-competency-text"
              data-cat-index="${categoryIndex}" data-comp-index="${competencyIndex}">
            <button class="btn-remove-competency"
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
                <span>${level === 'craftsman' ? 'Craftsman/Supervisor' : level === 'semiSkilled' ? 'Semi-skilled' : level === 'foundation' ? 'Foundation skills' : 'Skilled'}</span>
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

  if (isEditable) {
    heading.setAttribute('contenteditable', 'false');
    heading.style.cursor = '';
    showStatus('Heading updated! ✓', 'success');
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

export function clearSection(inputId, headingId, defaultHeading) {
  const current = (document.getElementById(inputId)?.value || '').trim();
  const heading = document.getElementById(headingId)?.textContent?.trim();
  if (!current && (!heading || heading === defaultHeading)) {
    showStatus('This section is already empty — nothing to clear', 'success');
    return;
  }
  if (confirm('Are you sure you want to clear this section?')) {
    document.getElementById(inputId).value = '';
    document.getElementById(headingId).textContent = defaultHeading;
    document.getElementById(headingId).setAttribute('contenteditable', 'false');
    showStatus('Section cleared! ✓', 'success');
  }
}

export function formatList(inputId, formatType) {
  const textarea = document.getElementById(inputId);
  const text = textarea.value.trim();
  if (!text) { showStatus('Nothing to format! Add some content first.', 'error'); return; }

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
  showStatus(`✓ Formatted with ${formatType === 'number' ? 'numbering' : 'bullets'}!`, 'success');
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
      <h3 id="${headingId}" contenteditable="false">Custom Section ${appState.customSectionCounter}</h3>
      <div style="display:flex;gap:10px;">
        <button class="btn-rename" data-action="toggle-edit-heading" data-heading-id="${headingId}">✏️ Rename</button>
        <button class="btn-clear-section" data-action="clear-section"
          data-input-id="${inputId}" data-heading-id="${headingId}"
          data-default-heading="Custom Section ${appState.customSectionCounter}">🗑️ Clear</button>
        <button class="btn-remove-section" data-action="remove-custom-section" data-section-id="${sectionId}"
          style="background:linear-gradient(135deg,#ef4444 0%,#dc2626 100%);color:white;padding:8px 16px;font-size:0.95em;border:none;border-radius:8px;cursor:pointer;">
          ❌ Remove
        </button>
      </div>
    </div>
    <textarea id="${inputId}" placeholder="Enter information for this custom section on separate lines"></textarea>`;

  container.appendChild(sectionDiv);
  showStatus('Custom section added! ✓', 'success');
}

export function removeCustomSection(sectionId) {
  if (confirm('Are you sure you want to remove this section? This cannot be undone!')) {
    const section = document.getElementById(sectionId);
    if (section) { section.remove(); showStatus('Section removed! ✓', 'success'); }
  }
}
