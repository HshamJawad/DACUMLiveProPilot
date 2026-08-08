// ============================================================
// /modules.js
// Competency Clustering, Learning Outcomes, and Module Mapping
// ============================================================

import { appState } from './state.js';
import { showStatus } from './renderer.js';
import { lwExtractDutiesAndTasks } from './workshop.js';
import { getTaskCode, getDutyLabel } from './codes.js';

/* i18n access — resolved lazily; see duties.js for why. */
const _t  = (k)    => (window.i18n ? window.i18n.t(k)     : k);
const _tf = (k, v) => (window.i18n ? window.i18n.tf(k, v) : k);


// switchTab is exposed on window by app.js to avoid circular deps
function switchTab(tabId) { window.switchTab(tabId); }

// ── Helpers (getTaskCode now sourced from codes.js) ─────────
// The old local implementation parsed the immutable task ID
// (e.g. "duty_3_2" → "Task C2") which was *wrong* after a drag
// reorder.  The imported version reads the live position from
// appState.dutiesData and always returns the correct letter.

// ── Clustering ────────────────────────────────────────────────

export function bypassToClusteringTab() {
  appState.verificationDecisionMade = true;
  appState.clusteringAllowed = true;
  document.getElementById('btnLWFinalize').disabled = true;
  document.getElementById('btnBypassToClustering').disabled = true;
  document.getElementById('btnResetDecision').style.display = 'inline-block';
  initializeClusteringFromTasks();
  switchTab('clustering-tab');
}

export function resetVerificationDecision() {
  appState.verificationDecisionMade = false;
  appState.clusteringAllowed = false;
  document.getElementById('btnLWFinalize').disabled = false;
  document.getElementById('btnBypassToClustering').disabled = false;
  document.getElementById('btnResetDecision').style.display = 'none';
}

export function initializeClusteringFromTasks() {
  const cd = appState.clusteringData;
  cd.availableTasks = [];
  cd.clusters = [];
  cd.clusterCounter = 0;

  if (appState.lwAggregatedResults && appState.lwAggregatedResults.taskResults) {
    const taskResults = appState.lwAggregatedResults.taskResults;
    const allTasks = [];
    Object.keys(taskResults).forEach(taskId => {
      const voteData = taskResults[taskId];
      allTasks.push({
        id: taskId,
        text: voteData.taskText,
        dutyTitle: voteData.dutyTitle,
        priorityIndex: voteData.priorityIndex || 0
      });
    });
    allTasks.sort((a, b) => b.priorityIndex - a.priorityIndex);
    cd.availableTasks = allTasks;
  } else {
    const duties = lwExtractDutiesAndTasks();
    const allTasks = [];
    Object.keys(duties).forEach(dutyId => {
      const duty = duties[dutyId];
      duty.tasks.forEach(task => {
        allTasks.push({ id: task.id, text: task.text, dutyTitle: duty.title, priorityIndex: null });
      });
    });
    cd.availableTasks = allTasks;
  }

  renderAvailableTasks();
  renderClusters();
}

export function renderAvailableTasks() {
  const cd = appState.clusteringData;
  const container = document.getElementById('availableTasksList');

  if (cd.availableTasks.length === 0) {
    container.innerHTML = `<div class="no-tasks-message">${_t('msgAllTasksAssigned')}</div>`;
    document.getElementById('btnCreateCluster').disabled = true;
    return;
  }

  let html = '';
  cd.availableTasks.forEach((task, index) => {
    const taskCode = getTaskCode(task.id);
    let clusterOptions = `<option value="">${_t('optSelectCluster')}</option>`;
    cd.clusters.forEach(cluster => {
      clusterOptions += `<option value="${cluster.id}">${cluster.name}</option>`;
    });

    html += `
      <div class="task-checkbox-item">
        <input type="checkbox" id="task_${index}" data-action="update-cluster-button">
        <label for="task_${index}" class="task-checkbox-label">
          <strong>${taskCode}:</strong> ${task.text}
        </label>
        ${task.priorityIndex !== null ? `<span class="task-priority-badge">PI: ${task.priorityIndex.toFixed(2)}</span>` : ''}
        ${cd.clusters.length > 0 ? `
        <div class="task-dropdown-container">
          <span class="task-dropdown-label">${_t('lblAddTo')}</span>
          <select class="task-reassign-dropdown" data-action="add-task-to-cluster-dropdown" data-task-index="${index}">
            ${clusterOptions}
          </select>
        </div>` : ''}
      </div>`;
  });

  container.innerHTML = html;
  updateCreateClusterButton();
}

export function updateCreateClusterButton() {
  const checkboxes = document.querySelectorAll('#availableTasksList input[type="checkbox"]');
  const anyChecked = Array.from(checkboxes).some(cb => cb.checked);
  document.getElementById('btnCreateCluster').disabled = !anyChecked;
}

export function createCluster() {
  const cd = appState.clusteringData;
  const checkboxes = document.querySelectorAll('#availableTasksList input[type="checkbox"]');
  const selectedIndices = [];
  checkboxes.forEach((cb, index) => { if (cb.checked) selectedIndices.push(index); });
  if (selectedIndices.length === 0) return;

  cd.clusterCounter++;
  const newCluster = {
    id: `cluster_${cd.clusterCounter}`,
    name: _tf('lblClusterN', { n: cd.clusterCounter }),
    tasks: [],
    range: '',
    performanceCriteria: []
  };

  selectedIndices.sort((a, b) => b - a);
  selectedIndices.forEach(index => {
    newCluster.tasks.push(cd.availableTasks[index]);
    cd.availableTasks.splice(index, 1);
  });

  cd.clusters.push(newCluster);
  renderAvailableTasks();
  renderClusters();
}

export function renderClusters() {
  const cd = appState.clusteringData;
  const container = document.getElementById('clustersContainer');

  if (cd.clusters.length === 0) {
    container.innerHTML = `<div class="no-clusters-message">${_t('msgNoClusters')}</div>`;
    return;
  }

  let html = '';
  cd.clusters.forEach((cluster, clusterIndex) => {
    const clusterNumber = clusterIndex + 1;
    let displayValue = '';
    if (cluster.performanceCriteria && cluster.performanceCriteria.length > 0) {
      displayValue = cluster.performanceCriteria
        .map((criterion, idx) => `${clusterNumber}-${idx + 1} ${criterion}`)
        .join('\n');
    }

    html += `
      <div class="cluster-item">
        <div class="cluster-header">
          <div class="cluster-title">${cluster.name}</div>
          <div class="cluster-actions">
            <button class="btn-rename-cluster" data-action="regen-cluster-criteria" data-cluster-id="${cluster.id}"
                    title="${_t('ttRegenCriteria')}">🤖 ${_t('btnAICriteria')}</button>
            <button class="btn-rename-cluster" data-action="rename-cluster" data-cluster-id="${cluster.id}">✏️ ${_t('btnRename')}</button>
            <button class="btn-delete-cluster" data-action="delete-cluster" data-cluster-id="${cluster.id}">🗑️ ${_t('btnDelete')}</button>
          </div>
        </div>

        <div class="cluster-section">
          <h4>📋 ${_t('lblRelatedTasks')}</h4>
          <div class="related-tasks-list">
            ${cluster.tasks.map((task, taskIndex) => {
              const taskCode = getTaskCode(task.id);
              return `
                <div class="related-task-item" style="display:flex;justify-content:space-between;align-items:center;">
                  <div style="flex:1"><strong>${taskCode}:</strong> ${task.text}</div>
                  <button class="btn-remove-task" data-action="remove-task-from-cluster"
                    data-cluster-id="${cluster.id}" data-task-index="${taskIndex}" style="margin-left:10px;">✕</button>
                </div>`;
            }).join('') || `<div style="color:#999;font-style:italic;">${_t('msgNoTasksAssigned')}</div>`}
          </div>
        </div>

        <div class="cluster-section">
          <div class="cluster-section-header">
            <h4>🎯 ${_t('lblRange')}</h4>
            <button type="button" class="tab-help-btn" data-action="show-pc-range-help" title="${_t('ttPCRangeHelp')}" aria-label="${_t('ttPCRangeHelp')}" aria-haspopup="dialog">?</button>
          </div>
          <div class="cluster-helper-text">${_t('hintRange')}</div>
          <textarea id="range_${cluster.id}" data-action="update-cluster-range" data-cluster-id="${cluster.id}">${cluster.range || ''}</textarea>
        </div>

        <div class="cluster-section">
          <div class="cluster-section-header">
            <h4>✅ ${_t('lblPerformanceCriteria')}</h4>
            <button type="button" class="tab-help-btn" data-action="show-pc-range-help" title="${_t('ttPCRangeHelp')}" aria-label="${_t('ttPCRangeHelp')}" aria-haspopup="dialog">?</button>
          </div>
          <div class="cluster-helper-text">${_t('hintCriteria')}</div>
          <textarea id="criteria_${cluster.id}"
            data-cluster-number="${clusterNumber}"
            data-cluster-id="${cluster.id}"
            data-action-focus="init-criteria-number"
            data-action-keydown="handle-criteria-keydown"
            data-action-blur="update-cluster-criteria-numbered"
            placeholder="${_t('phFirstCriterion')}"
            style="min-height:120px;">${displayValue}</textarea>
        </div>
      </div>`;
  });

  container.innerHTML = html;
}

export function renameCluster(clusterId) {
  const cluster = appState.clusteringData.clusters.find(c => c.id === clusterId);
  if (!cluster) return;
  const newName = prompt(_t('promptRenameCluster'), cluster.name);
  if (newName && newName.trim()) {
    cluster.name = newName.trim();
    renderClusters();
  }
}

export function deleteCluster(clusterId) {
  const cd = appState.clusteringData;
  const idx = cd.clusters.findIndex(c => c.id === clusterId);
  if (idx === -1) return;
  const cluster = cd.clusters[idx];
  cd.availableTasks.push(...cluster.tasks);
  if (cd.availableTasks.length > 0 && cd.availableTasks[0].priorityIndex !== null) {
    cd.availableTasks.sort((a, b) => b.priorityIndex - a.priorityIndex);
  }
  cd.clusters.splice(idx, 1);
  renderAvailableTasks();
  renderClusters();
}

export function removeTaskFromCluster(clusterId, taskIndex) {
  const cd = appState.clusteringData;
  const cluster = cd.clusters.find(c => c.id === clusterId);
  if (!cluster) return;
  const task = cluster.tasks[taskIndex];
  cluster.tasks.splice(taskIndex, 1);
  cd.availableTasks.push(task);
  if (cd.availableTasks.length > 0 && cd.availableTasks[0].priorityIndex !== null) {
    cd.availableTasks.sort((a, b) => b.priorityIndex - a.priorityIndex);
  }
  renderAvailableTasks();
  renderClusters();
}

export function addTaskToClusterFromDropdown(taskIndex, clusterId) {
  if (!clusterId) return;
  const cd = appState.clusteringData;
  const cluster = cd.clusters.find(c => c.id === clusterId);
  if (!cluster) return;
  const task = cd.availableTasks[taskIndex];
  if (!task) return;
  cluster.tasks.push(task);
  cd.availableTasks.splice(taskIndex, 1);
  renderAvailableTasks();
  renderClusters();
}

export function updateClusterRange(clusterId, value) {
  const cluster = appState.clusteringData.clusters.find(c => c.id === clusterId);
  if (cluster) cluster.range = value;
}

export function updateClusterCriteria(clusterId, value) {
  const cluster = appState.clusteringData.clusters.find(c => c.id === clusterId);
  if (cluster) {
    cluster.performanceCriteria = value.split('\n').map(l => l.trim()).filter(l => l);
    renderClusters();
  }
}

export function updateClusterCriteriaFromNumbered(clusterId, value) {
  const cluster = appState.clusteringData.clusters.find(c => c.id === clusterId);
  if (!cluster) return;
  const lines = value.split('\n');
  const stripped = lines.map(line => {
    const match = line.match(/^\d+-\d+\s+(.*)$/);
    return match ? match[1].trim() : line.trim();
  }).filter(line => line);
  cluster.performanceCriteria = stripped;
}

export function handleCriteriaKeydown(event, clusterId) {
  if (event.key === 'Enter') {
    const textarea = event.target;
    const clusterNumber = textarea.getAttribute('data-cluster-number');
    const cursorPos = textarea.selectionStart;
    const value = textarea.value;
    const lines = value.substring(0, cursorPos).split('\n');
    const nextNumber = lines.length + 1;
    event.preventDefault();
    const before = value.substring(0, cursorPos);
    const after = value.substring(cursorPos);
    const newText = before + '\n' + clusterNumber + '-' + nextNumber + ' ' + after;
    textarea.value = newText;
    const newCursorPos = cursorPos + 1 + clusterNumber.length + 1 + String(nextNumber).length + 1;
    textarea.setSelectionRange(newCursorPos, newCursorPos);
  }
}

export function initCriteriaNumber(event, clusterId) {
  const textarea = event.target;
  const clusterNumber = textarea.getAttribute('data-cluster-number');
  if (!textarea.value.trim()) {
    textarea.value = clusterNumber + '-1 ';
    textarea.setSelectionRange(textarea.value.length, textarea.value.length);
  }
}

export function proceedToClusteringFromVerification() {
  if (appState.clusteringAllowed !== true) {
    alert('Please choose one option above (Live Voting or Without Verification) first.');
    return;
  }
  initializeClusteringFromTasks();
  switchTab('clustering-tab');
}

// ── Learning Outcomes ─────────────────────────────────────────

export function renderPCSourceList() {
  const container = document.getElementById('pcSourceList');
  if (!container) return;

  const cd = appState.clusteringData;
  if (!cd.clusters || cd.clusters.length === 0) {
    container.innerHTML = `<div class="no-tasks-message">${_t('msgNoPCAvailable')}</div>`;
    return;
  }

  const lo = appState.learningOutcomesData;
  const usedPCIds = new Set();
  lo.outcomes.forEach(outcome => {
    if (outcome.linkedCriteria) outcome.linkedCriteria.forEach(pc => usedPCIds.add(pc.id));
  });

  let html = '';
  let hasAnyCriteria = false;

  cd.clusters.forEach((cluster, clusterIndex) => {
    const clusterNumber = clusterIndex + 1;
    if (!cluster.performanceCriteria || !Array.isArray(cluster.performanceCriteria) || cluster.performanceCriteria.length === 0) return;

    hasAnyCriteria = true;
    html += `<div class="pc-cluster-group"><h4>${cluster.name}</h4>`;

    cluster.performanceCriteria.forEach((criterion, criterionIndex) => {
      if (!criterion || !criterion.trim()) return;
      const pcId = `C${clusterNumber}-PC${criterionIndex + 1}`;
      const isUsed = usedPCIds.has(pcId);

      let loOptions = `<option value="">${_t('optAssignToLO')}</option>`;
      lo.outcomes.forEach(outcome => {
        loOptions += `<option value="${outcome.id}">${outcome.number}</option>`;
      });

      html += `
        <div class="pc-checkbox-item ${isUsed ? 'used' : ''}" id="pc_${pcId}">
          <input type="checkbox" id="cb_${pcId}"
            data-pc-id="${pcId}" data-cluster="${clusterNumber}" data-criterion="${criterionIndex}"
            ${isUsed ? 'disabled' : ''} data-action="update-lo-button">
          <label for="cb_${pcId}" class="pc-label">
            <span class="pc-number">${pcId}:</span> ${criterion}
          </label>
          ${isUsed ? `<span class="pc-used-badge">${_t('lblUsed')}</span>` : ''}
          ${lo.outcomes.length > 0 ? `
          <div class="task-dropdown-container" style="margin-left:10px;">
            <select class="task-reassign-dropdown"
              data-action="reassign-pc-to-lo"
              data-pc-id="${pcId}" data-cluster="${clusterNumber}" data-criterion="${criterionIndex}">
              ${loOptions}
            </select>
          </div>` : ''}
        </div>`;
    });

    html += '</div>';
  });

  if (!hasAnyCriteria || !html) {
    container.innerHTML = `<div class="no-tasks-message">${_t('msgNoPCForModules')}</div>`;
  } else {
    container.innerHTML = html;
  }

  updateCreateLOButton();
}

export function updateCreateLOButton() {
  const checkboxes = document.querySelectorAll('#pcSourceList input[type="checkbox"]:not([disabled])');
  const anyChecked = Array.from(checkboxes).some(cb => cb.checked);
  document.getElementById('btnCreateLO').disabled = !anyChecked;
}

export function createLearningOutcome() {
  const checkboxes = document.querySelectorAll('#pcSourceList input[type="checkbox"]:checked');
  if (checkboxes.length === 0) return;

  const linkedCriteria = [];
  checkboxes.forEach(cb => {
    const pcId = cb.getAttribute('data-pc-id');
    const clusterNum = parseInt(cb.getAttribute('data-cluster'));
    const criterionIdx = parseInt(cb.getAttribute('data-criterion'));
    const cluster = appState.clusteringData.clusters[clusterNum - 1];
    const criterionText = cluster.performanceCriteria[criterionIdx];
    linkedCriteria.push({ id: pcId, text: criterionText, clusterNumber: clusterNum });
  });

  const lo = appState.learningOutcomesData;
  lo.outcomeCounter++;
  lo.outcomes.push({
    id: `lo_${lo.outcomeCounter}`,
    number: `LO-${lo.outcomeCounter}`,
    statement: '',
    linkedCriteria
  });

  renderPCSourceList();
  renderLearningOutcomes();
}

export function renderLearningOutcomes() {
  const container = document.getElementById('loBlocksContainer');
  const lo = appState.learningOutcomesData;

  if (lo.outcomes.length === 0) {
    container.innerHTML = `<div class="no-clusters-message">${_t('msgNoLOs')}</div>`;
    return;
  }

  let html = '';
  lo.outcomes.forEach(outcome => {
    const isEditing = outcome.editing || false;
    html += `
      <div class="lo-block" id="${outcome.id}">
        <div class="lo-block-header">
          <div class="lo-number">${outcome.number}</div>
          <div class="lo-actions">
            <button class="btn-edit-lo" data-action="toggle-edit-lo" data-lo-id="${outcome.id}">
              ${isEditing ? '💾 ' + _t('btnSave') : '✏️ ' + _t('btnEdit')}
            </button>
            <button class="btn-delete-lo" data-action="delete-lo" data-lo-id="${outcome.id}">❌ ${_t('btnDelete')}</button>
          </div>
        </div>
        <div class="lo-statement" id="statement_${outcome.id}">
          ${isEditing
            ? `<textarea id="textarea_${outcome.id}" data-action-blur="save-lo-statement" data-lo-id="${outcome.id}">${outcome.statement}</textarea>`
            : `${outcome.statement || `<em style="color:#999;">${_t('phLOStatement')}</em>`}`
          }
        </div>
        <div class="lo-linked-criteria">
          <h5>📎 ${_t('lblMappedPC')}</h5>
          ${outcome.linkedCriteria.map(pc => `
            <div class="lo-linked-item">
              <div style="flex:1"><strong>${pc.id}:</strong> ${pc.text}</div>
              <button class="btn-remove-task" data-action="unassign-pc-from-lo"
                data-lo-id="${outcome.id}" data-pc-id="${pc.id}" style="margin-left:10px;">✕</button>
            </div>`).join('')}
        </div>
      </div>`;
  });

  container.innerHTML = html;
}

export function toggleEditLO(loId) {
  const lo = appState.learningOutcomesData.outcomes.find(o => o.id === loId);
  if (!lo) return;
  if (lo.editing) {
    saveLOStatement(loId);
    lo.editing = false;
  } else {
    lo.editing = true;
  }
  renderLearningOutcomes();
  if (lo.editing) {
    setTimeout(() => {
      const ta = document.getElementById(`textarea_${loId}`);
      if (ta) { ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length); }
    }, 50);
  }
}

export function saveLOStatement(loId) {
  const ta = document.getElementById(`textarea_${loId}`);
  if (!ta) return;
  const lo = appState.learningOutcomesData.outcomes.find(o => o.id === loId);
  if (lo) lo.statement = ta.value.trim();
}

export function deleteLearningOutcome(loId) {
  if (!confirm(_t('confirmDeleteLO'))) return;
  const data = appState.learningOutcomesData;
  const idx = data.outcomes.findIndex(o => o.id === loId);
  if (idx !== -1) data.outcomes.splice(idx, 1);
  renderPCSourceList();
  renderLearningOutcomes();
}

export function reassignPCToLO(pcId, clusterNumber, criterionIndex, targetLoId) {
  if (!targetLoId) return;
  const lo = appState.learningOutcomesData;
  const targetLO = lo.outcomes.find(o => o.id === targetLoId);
  if (!targetLO) return;

  const alreadyInTarget = targetLO.linkedCriteria.some(pc => pc.id === pcId);
  if (!alreadyInTarget) {
    lo.outcomes.forEach(outcome => {
      const idx = outcome.linkedCriteria.findIndex(pc => pc.id === pcId);
      if (idx !== -1) outcome.linkedCriteria.splice(idx, 1);
    });
    const cluster = appState.clusteringData.clusters[clusterNumber - 1];
    const criterionText = cluster.performanceCriteria[criterionIndex];
    targetLO.linkedCriteria.push({ id: pcId, text: criterionText, clusterNumber });
  }

  renderPCSourceList();
  renderLearningOutcomes();
}

export function unassignPCFromLO(loId, pcId) {
  const lo = appState.learningOutcomesData.outcomes.find(o => o.id === loId);
  if (!lo) return;
  const idx = lo.linkedCriteria.findIndex(pc => pc.id === pcId);
  if (idx !== -1) lo.linkedCriteria.splice(idx, 1);
  renderPCSourceList();
  renderLearningOutcomes();
}

// ── Module Mapping ────────────────────────────────────────────

export function renderModuleLoList() {
  const container = document.getElementById('moduleLoList');
  const lo = appState.learningOutcomesData;
  const mm = appState.moduleMappingData;

  if (!lo.outcomes || lo.outcomes.length === 0) {
    container.innerHTML = '<div class="no-tasks-message">No Learning Outcomes available. Please create Learning Outcomes first.</div>';
    document.getElementById('btnCreateModule').disabled = true;
    return;
  }

  const assignedLoIds = new Set();
  mm.modules.forEach(module => module.learningOutcomes.forEach(o => assignedLoIds.add(o.id)));
  const availableLos = lo.outcomes.filter(o => !assignedLoIds.has(o.id));

  if (availableLos.length === 0) {
    container.innerHTML = '<div class="no-tasks-message">All Learning Outcomes have been assigned to modules.</div>';
    document.getElementById('btnCreateModule').disabled = true;
    return;
  }

  let html = '';
  availableLos.forEach(outcome => {
    const criteriaText = outcome.linkedCriteria.map(pc => pc.id).join(', ');
    let moduleOptions = '<option value="">Select Module</option>';
    mm.modules.forEach(m => { moduleOptions += `<option value="${m.id}">${m.title}</option>`; });

    html += `
      <div class="module-lo-item">
        <input type="checkbox" id="mlo_${outcome.id}" data-lo-id="${outcome.id}" data-action="update-module-button">
        <div class="module-lo-content">
          <div class="module-lo-number">${outcome.number}</div>
          <div class="module-lo-statement">${outcome.statement || '<em>No statement provided</em>'}</div>
          <div class="module-lo-criteria">Mapped PC: ${criteriaText}</div>
        </div>
        ${mm.modules.length > 0 ? `
        <div class="task-dropdown-container">
          <span class="task-dropdown-label">${_t('lblAddTo')}</span>
          <select class="task-reassign-dropdown" data-action="add-lo-to-module-dropdown" data-lo-id="${outcome.id}">
            ${moduleOptions}
          </select>
        </div>` : ''}
      </div>`;
  });

  container.innerHTML = html;
  updateCreateModuleButton();
}

export function updateCreateModuleButton() {
  const checkboxes = document.querySelectorAll('#moduleLoList input[type="checkbox"]');
  const anyChecked = Array.from(checkboxes).some(cb => cb.checked);
  document.getElementById('btnCreateModule').disabled = !anyChecked;
}

export function createModule() {
  const mm = appState.moduleMappingData;
  const lo = appState.learningOutcomesData;
  const checkboxes = document.querySelectorAll('#moduleLoList input[type="checkbox"]');
  const selectedLoIds = [];
  checkboxes.forEach(cb => { if (cb.checked) selectedLoIds.push(cb.getAttribute('data-lo-id')); });
  if (selectedLoIds.length === 0) return;

  mm.moduleCounter++;
  const newModule = { id: `module_${mm.moduleCounter}`, title: `Module ${mm.moduleCounter}`, learningOutcomes: [] };
  selectedLoIds.forEach(loId => {
    const outcome = lo.outcomes.find(o => o.id === loId);
    if (outcome) newModule.learningOutcomes.push(outcome);
  });
  mm.modules.push(newModule);

  renderModuleLoList();
  renderModules();
}

export function renderModules() {
  const container = document.getElementById('modulesContainer');
  const mm = appState.moduleMappingData;

  if (mm.modules.length === 0) {
    container.innerHTML = '<div class="no-clusters-message">No modules created yet.</div>';
    return;
  }

  let html = '';
  mm.modules.forEach(module => {
    html += `
      <div class="module-item">
        <div class="module-header">
          <div class="module-title">${module.title}</div>
          <div class="module-actions">
            <button class="btn-rename-module" data-action="rename-module" data-module-id="${module.id}">✏️ Rename</button>
            <button class="btn-delete-module" data-action="delete-module" data-module-id="${module.id}">🗑️ Delete Module</button>
          </div>
        </div>
        <div class="module-los-list">
          ${module.learningOutcomes.map(outcome => {
            const criteriaText = outcome.linkedCriteria.map(pc => `${pc.id}: ${pc.text}`).join(' • ');
            return `
              <div class="module-lo-assigned">
                <div class="module-lo-assigned-content">
                  <div class="module-lo-assigned-number">${outcome.number}</div>
                  <div class="module-lo-assigned-statement">${outcome.statement || '<em>No statement</em>'}</div>
                  <div class="module-lo-assigned-criteria">Mapped PC: ${criteriaText}</div>
                </div>
                <button class="btn-remove-lo" data-action="remove-lo-from-module"
                  data-module-id="${module.id}" data-lo-id="${outcome.id}">✕ Remove</button>
              </div>`;
          }).join('')}
        </div>
      </div>`;
  });

  container.innerHTML = html;
}

export function renameModule(moduleId) {
  const module = appState.moduleMappingData.modules.find(m => m.id === moduleId);
  if (!module) return;
  const newTitle = prompt('Enter new module title:', module.title);
  if (newTitle && newTitle.trim()) {
    module.title = newTitle.trim();
    renderModules();
  }
}

export function deleteModule(moduleId) {
  const mm = appState.moduleMappingData;
  const idx = mm.modules.findIndex(m => m.id === moduleId);
  if (idx === -1) return;
  if (!confirm('Delete this module? All Learning Outcomes will return to the available list.')) return;
  mm.modules.splice(idx, 1);
  renderModuleLoList();
  renderModules();
}

export function removeLoFromModule(moduleId, loId) {
  const module = appState.moduleMappingData.modules.find(m => m.id === moduleId);
  if (!module) return;
  const idx = module.learningOutcomes.findIndex(o => o.id === loId);
  if (idx !== -1) module.learningOutcomes.splice(idx, 1);
  renderModuleLoList();
  renderModules();
}

export function addLoToModuleFromDropdown(loId, moduleId) {
  if (!moduleId) return;
  const mm = appState.moduleMappingData;
  const lo = appState.learningOutcomesData;
  const module = mm.modules.find(m => m.id === moduleId);
  if (!module) return;
  const outcome = lo.outcomes.find(o => o.id === loId);
  if (!outcome) return;
  module.learningOutcomes.push(outcome);
  renderModuleLoList();
  renderModules();
}

export function openModuleBuilderFromMapping() {
  const occupationTitle = document.getElementById('occupationTitle')?.value || '';
  const jobTitle = document.getElementById('jobTitle')?.value || '';
  const occupation = occupationTitle || jobTitle || 'Unknown Occupation';
  const mm = appState.moduleMappingData;

  const exportObject = {
    source: 'DACUM Live Pro v1.0',
    exportDate: new Date().toISOString(),
    occupation,
    modules: mm.modules.map(module => ({
      moduleId: module.id,
      moduleTitle: module.title,
      learningOutcomes: module.learningOutcomes.map(o => ({
        number: o.number,
        statement: o.statement,
        performanceCriteria: o.linkedCriteria.map(pc => ({ id: pc.id, text: pc.text }))
      }))
    }))
  };

  try {
    localStorage.setItem('dacum_modules_export', JSON.stringify(exportObject));
    window.open('Module_Builder.html', '_blank');
    showStatus('Module data exported! Opening Module Builder...', 'success');
  } catch (error) {
    console.error('Error exporting to Module Builder:', error);
    showStatus('Error exporting data: ' + error.message, 'error');
  }
}

export function exportModuleMappingJSON() {
  const mm = appState.moduleMappingData;
  if (!mm.modules || mm.modules.length === 0) {
    showStatus('No modules to export. Please create modules first.', 'error');
    return;
  }

  const occupationTitle = document.getElementById('occupationTitle')?.value || '';
  const jobTitle = document.getElementById('jobTitle')?.value || '';
  const occupation = occupationTitle || jobTitle || 'Unknown Occupation';

  const exportData = {
    metadata: {
      toolName: 'DACUM Live Pro', toolVersion: '1.0',
      exportDate: new Date().toISOString(), exportType: 'Module Mapping', occupation
    },
    modules: mm.modules.map(module => ({
      moduleId: module.id, moduleTitle: module.title,
      learningOutcomes: module.learningOutcomes.map(o => ({
        number: o.number, statement: o.statement,
        performanceCriteria: o.linkedCriteria.map(pc => ({ id: pc.id, description: pc.text })),
        sourceTaskIds: o.linkedCriteria.map(pc => pc.taskId).filter(Boolean)
      }))
    })),
    summary: {
      totalModules: mm.modules.length,
      totalLearningOutcomes: mm.modules.reduce((s, m) => s + m.learningOutcomes.length, 0),
      totalPerformanceCriteria: mm.modules.reduce((s, m) =>
        s + m.learningOutcomes.reduce((ls, o) => ls + o.linkedCriteria.length, 0), 0)
    }
  };

  const dateStr = new Date().toISOString().split('T')[0];
  const filename = `module-mapping-export_${dateStr}.json`;

  try {
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url; link.download = filename;
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showStatus(`Module Mapping exported successfully: ${filename}`, 'success');
  } catch (error) {
    console.error('Error exporting module mapping:', error);
    showStatus('Error exporting module mapping: ' + error.message, 'error');
  }
}


/* ── Re-render on language change ────────────────────────────────────
   Both lists are innerHTML-generated and survive tab switches, so they
   are outside applyTranslations()' reach. Rendering is pure from
   appState — no user input is held in the DOM alone — so a rebuild is
   lossless. Guarded on the containers existing so a language switch
   never constructs a tab the user has not opened. */
window.addEventListener('dacum:langchange', () => {
  if (document.getElementById('availableTasksList')) renderAvailableTasks();
  if (document.getElementById('clustersContainer'))  renderClusters();
  if (document.getElementById('pcSourceList'))       renderPCSourceList();

  /* An outcome being edited holds its text in an unsaved <textarea>.
     Clicking the language button blurs it first, which fires the blur
     handler and commits the text — so the rebuild below is safe. Doing
     it in the other order would silently discard whatever the user had
     just typed. */
  if (document.getElementById('loBlocksContainer'))  renderLearningOutcomes();
});
