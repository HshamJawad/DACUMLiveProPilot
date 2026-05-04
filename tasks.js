// ============================================================
// /tasks.js
// Task Verification, Workshop Aggregated Counts, Dashboard.
// ============================================================

import { appState } from './state.js';
import { showStatus, escapeHtml } from './renderer.js';
import { getDutyLetter, getDutyCode, getTaskCodeShort, getDutyLabel } from './codes.js';

// ── Mode Controls ─────────────────────────────────────────────

export function updateCollectionMode() {
  const workshopRadio = document.getElementById('mode-workshop');
  const surveyRadio   = document.getElementById('mode-survey');
  if (workshopRadio.checked)    appState.collectionMode = 'workshop';
  else if (surveyRadio.checked) appState.collectionMode = 'survey';

  const workshopSection = document.getElementById('workshopParticipantsSection');
  const dashboardSection = document.getElementById('resultsDashboard');
  if (appState.collectionMode === 'workshop') {
    workshopSection.style.display = 'block';
    dashboardSection.style.display = 'block';
  } else {
    workshopSection.style.display = 'none';
    dashboardSection.style.display = 'none';
  }
  loadDutiesForVerification();
  showStatus(`Data collection mode: ${appState.collectionMode === 'workshop' ? 'Workshop (Facilitated)' : 'Individual / Survey'}`, 'success');
}

export function updateWorkflowMode() {
  const standardRadio  = document.getElementById('workflow-standard');
  const extendedRadio  = document.getElementById('workflow-extended');
  if (standardRadio.checked)   appState.workflowMode = 'standard';
  else if (extendedRadio.checked) appState.workflowMode = 'extended';

  const container = document.getElementById('verificationAccordionContainer');
  if (appState.workflowMode === 'extended') {
    container.classList.add('workflow-extended');
  } else {
    container.classList.remove('workflow-extended');
  }
  loadDutiesForVerification();

  const priorityFormulaSection = document.getElementById('priorityFormulaSection');
  if (priorityFormulaSection) {
    priorityFormulaSection.style.display =
      (appState.workflowMode === 'standard' && appState.collectionMode === 'workshop') ? 'block' : 'none';
  }
  showStatus(`Workflow mode: ${appState.workflowMode === 'standard' ? 'Standard (DACUM)' : 'Extended (DACUM)'}`, 'success');
}

export function updateParticipantCount() {
  const input = document.getElementById('workshopParticipants');
  appState.workshopParticipants = parseInt(input.value) || 10;
  validateAndComputeWorkshopResults();
  showStatus(`Participants set to ${appState.workshopParticipants}. Re-validating all tasks...`, 'success');
}

export function updatePriorityFormula() {
  const ifRadio  = document.getElementById('formula-if');
  const ifdRadio = document.getElementById('formula-ifd');
  if (ifRadio.checked)       appState.priorityFormula = 'if';
  else if (ifdRadio.checked) appState.priorityFormula = 'ifd';
  validateAndComputeWorkshopResults();
  refreshDashboard();
  showStatus(`Priority formula: ${appState.priorityFormula === 'if' ? 'I × F' : 'I × F × D'}`, 'success');
}

export function updateTVExportMode() {
  const appendixRadio   = document.getElementById('tvExportAppendix');
  const standaloneRadio = document.getElementById('tvExportStandalone');
  if (appendixRadio && appendixRadio.checked)       appState.tvExportMode = 'appendix';
  else if (standaloneRadio && standaloneRadio.checked) appState.tvExportMode = 'standalone';
  showStatus(`Export mode: ${appState.tvExportMode === 'appendix' ? 'Include as Appendix' : 'Standalone Report'}`, 'success');
}

export function updateTrainingLoadMethod() {
  const advanced = document.querySelector('input[name="trainingLoadMethod"][value="advanced"]');
  appState.trainingLoadMethod = advanced && advanced.checked ? 'advanced' : 'simple';
  const label = document.getElementById('trainingLoadMethodLabel');
  if (label) {
    label.innerHTML = `Current Method: <strong style="color:#667eea;">${appState.trainingLoadMethod === 'advanced' ? 'Advanced' : 'Simple'}</strong>`;
  }
  updateDutyLevelSummary();
}

// ── Load Duties for Verification ─────────────────────────────

export function loadDutiesForVerification() {
  const container   = document.getElementById('verificationAccordionContainer');
  // Restrict to real text fields only — buttons in Card View also carry
  // data-duty-id (for remove-duty actions) and would throw on .value.trim()
  const dutyInputs  = document.querySelectorAll('input[data-duty-id], textarea[data-duty-id]');

  if (dutyInputs.length === 0) {
    container.innerHTML = `
      <div class="no-duties-message">
        <h3>⚠️ No Duties Found</h3>
        <p>Please go to the "Duties & Tasks" tab and create duties with tasks first.</p>
        <p style="margin-top:10px;">Once you've added duties and tasks, click the "Refresh Duties & Tasks" button above.</p>
      </div>`;
    return;
  }

  container.innerHTML = '';
  let totalDuties = 0, totalTasks = 0;

  dutyInputs.forEach((dutyInput, dutyIndex) => {
    const dutyId   = dutyInput.getAttribute('data-duty-id');
    const dutyText = dutyInput.value.trim();
    if (!dutyText) return;
    totalDuties++;

    const tasks = [];
    // Only query real input/textarea fields — Card View also renders
    // buttons with data-task-div-id / data-action attributes, plus the
    // remove button on each task card which must not be treated as data
    document.querySelectorAll(
      `input[data-task-id^="${dutyId}_"], textarea[data-task-id^="${dutyId}_"]`
    ).forEach(taskInput => {
      const taskText = taskInput.value.trim();
      if (taskText) {
        const taskId = taskInput.getAttribute('data-task-id');
        appState.taskMetadata[taskId] = { dutyId, dutyTitle: dutyText, taskTitle: taskText };
        tasks.push({ id: taskId, text: taskText, dutyId, dutyTitle: dutyText });
        totalTasks++;
      }
    });

    if (tasks.length === 0) return;
    // dutyIndex here is the DOM-query index which matches the
    // appState.dutiesData order because renderDutiesFromState renders
    // duties in that order — so getDutyLetter(dutyIndex) is correct.
    container.insertAdjacentHTML('beforeend', createDutyAccordion(dutyId, dutyText, tasks, dutyIndex));
  });

  attachAccordionListeners();
  showStatus(`✓ Loaded ${totalDuties} duties with ${totalTasks} tasks for verification`, 'success');
}

// ── Accordion HTML Builder ────────────────────────────────────

function createDutyAccordion(dutyId, dutyText, tasks, dutyIndex = 0) {
  const isExtended  = appState.workflowMode === 'extended';
  const isWorkshop  = appState.collectionMode === 'workshop';
  const dutyLetter  = getDutyLetter(dutyIndex);

  const tasksTableRows = tasks.map((task, index) => {
    const taskKey = task.id;
    const ratings = appState.verificationRatings[taskKey] || {
      importance: null, frequency: null, difficulty: null,
      performsTask: false, criticality: null, comments: ''
    };

    const isComplete = ratings.importance !== null && ratings.frequency !== null && ratings.difficulty !== null;
    const taskScore  = isComplete ? (ratings.importance + ratings.frequency + ratings.difficulty) : '-';

    let weightedScore = '-', priorityLevel = 'low';
    if (isExtended && ratings.importance !== null && ratings.frequency !== null &&
        ratings.difficulty !== null && ratings.criticality !== null) {
      weightedScore = (ratings.importance * ratings.frequency) + ratings.difficulty + ratings.criticality;
      if      (weightedScore >= 10) priorityLevel = 'high';
      else if (weightedScore >= 6)  priorityLevel = 'medium';
    }

    const completionHtml = `<span class="completion-indicator ${isComplete ? 'complete' : 'incomplete'}">${isComplete ? '✓ Complete' : '○ Incomplete'}</span>`;
    const priorityHtml   = `<span class="priority-badge ${priorityLevel}">${priorityLevel.toUpperCase()}</span>`;

    let workshopMeans = '';
    if (isWorkshop && !isExtended) {
      const res = appState.workshopResults[taskKey];
      if (res && res.valid) {
        workshopMeans = `
          <td style="width:8%;text-align:center;"><span class="weighted-mean" id="mean_imp_${taskKey}">${res.meanImportance.toFixed(2)}</span></td>
          <td style="width:8%;text-align:center;"><span class="weighted-mean" id="mean_freq_${taskKey}">${res.meanFrequency.toFixed(2)}</span></td>
          <td style="width:8%;text-align:center;"><span class="weighted-mean" id="mean_diff_${taskKey}">${res.meanDifficulty.toFixed(2)}</span></td>
          <td style="width:8%;text-align:center;"><span class="priority-index" id="priority_${taskKey}">${res.priorityIndex.toFixed(2)}</span></td>`;
      } else {
        workshopMeans = `
          <td style="width:8%;text-align:center;"><span class="weighted-mean" id="mean_imp_${taskKey}">-</span></td>
          <td style="width:8%;text-align:center;"><span class="weighted-mean" id="mean_freq_${taskKey}">-</span></td>
          <td style="width:8%;text-align:center;"><span class="weighted-mean" id="mean_diff_${taskKey}">-</span></td>
          <td style="width:8%;text-align:center;"><span class="priority-index" id="priority_${taskKey}">-</span></td>`;
      }
    }

    return `
      <tr data-task-key="${taskKey}">
        <td style="width:${isExtended ? '25%' : (isWorkshop ? '28%' : '40%')};">
          <div class="task-text">${dutyLetter}${index + 1}. ${escapeHtml(task.text)}</div>
        </td>
        <td style="width:${isExtended ? '12%' : '15%'};">
          ${isWorkshop ? createCountInputs(taskKey, 'importance') : createRatingScale(taskKey, 'importance', ratings.importance)}
        </td>
        <td style="width:${isExtended ? '12%' : '15%'};">
          ${isWorkshop ? createCountInputs(taskKey, 'frequency') : createRatingScale(taskKey, 'frequency', ratings.frequency)}
        </td>
        <td style="width:${isExtended ? '12%' : '15%'};">
          ${isWorkshop ? createCountInputs(taskKey, 'difficulty') : createRatingScale(taskKey, 'difficulty', ratings.difficulty)}
        </td>
        ${!isExtended && !isWorkshop ? `
          <td style="width:8%;text-align:center;"><span class="score-display" id="score_${taskKey}">${taskScore}</span></td>
          <td style="width:12%;text-align:center;"><span id="completion_${taskKey}">${completionHtml}</span></td>
        ` : ''}
        ${!isExtended && isWorkshop ? workshopMeans : ''}
        ${isExtended ? `
          <td class="extended-only" style="width:10%;text-align:center;">
            <div class="performs-task-toggle">
              <input type="checkbox" id="performs_${taskKey}" data-action="performs-task" data-task-key="${taskKey}" ${ratings.performsTask ? 'checked' : ''}>
              <label for="performs_${taskKey}">Yes</label>
            </div>
          </td>
          <td class="extended-only" style="width:12%;">
            ${isWorkshop ? createCountInputs(taskKey, 'criticality') : createRatingScale(taskKey, 'criticality', ratings.criticality)}
          </td>
          <td class="extended-only" style="width:8%;text-align:center;"><span class="score-display" id="weighted_${taskKey}">${weightedScore}</span></td>
          <td class="extended-only" style="width:10%;text-align:center;"><span id="priority_${taskKey}">${priorityHtml}</span></td>
          <td class="extended-only" style="width:15%;">
            <textarea class="task-comments" id="comments_${taskKey}" data-action="update-comments" data-task-key="${taskKey}" placeholder="Optional comments...">${escapeHtml(ratings.comments || '')}</textarea>
          </td>
        ` : ''}
      </tr>`;
  }).join('');

  const tableHeader = !isExtended && !isWorkshop ? `<tr>
    <th>Task</th><th>Importance<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>Frequency<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>Learning Difficulty<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>Task Score</th><th>Completion</th></tr>`
  : !isExtended && isWorkshop ? `<tr>
    <th>Task</th><th>Importance Counts<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>Frequency Counts<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>Difficulty Counts<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>Mean<br>Importance</th><th>Mean<br>Frequency</th><th>Mean<br>Difficulty</th><th>Priority<br>Index</th></tr>`
  : `<tr>
    <th>Task</th>
    <th>${isWorkshop ? 'Importance Counts' : 'Importance'}<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>${isWorkshop ? 'Frequency Counts' : 'Frequency'}<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th>${isWorkshop ? 'Difficulty Counts' : 'Difficulty'}<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th class="extended-only">Performs?</th>
    <th class="extended-only">${isWorkshop ? 'Criticality Counts' : 'Criticality'}<br><span style="font-weight:400;font-size:.85em;">(0-3)</span></th>
    <th class="extended-only">Weighted Score</th>
    <th class="extended-only">Priority</th>
    <th class="extended-only">Comments</th></tr>`;

  return `
    <div class="duty-accordion">
      <div class="duty-accordion-header" data-duty="${dutyId}">
        <div class="duty-title">Duty ${dutyLetter}: ${escapeHtml(dutyText)}</div>
        <div class="duty-toggle">▼</div>
      </div>
      <div class="duty-accordion-content">
        <div style="overflow-x:auto;width:100%;">
          <table class="verification-table">
            <thead>${tableHeader}</thead>
            <tbody>${tasksTableRows}</tbody>
          </table>
        </div>
      </div>
    </div>`;
}

// ── Rating / Count Input Builders ─────────────────────────────

function createRatingScale(taskKey, dimension, currentValue) {
  return `<div class="rating-scale">
    ${[0,1,2,3].map(value => {
      const inputId  = `${taskKey}_${dimension}_${value}`;
      const isChecked = currentValue === value ? 'checked' : '';
      return `<div class="rating-option">
        <input type="radio" id="${inputId}" name="${taskKey}_${dimension}" value="${value}"
               ${isChecked} data-action="update-rating" data-task-key="${taskKey}"
               data-dimension="${dimension}" data-value="${value}">
        <label for="${inputId}">${value}</label>
      </div>`;
    }).join('')}
  </div>`;
}

function createCountInputs(taskKey, dimension) {
  const counts   = appState.workshopCounts[taskKey] || {};
  const dimCounts = counts[`${dimension}Counts`] || {0:0,1:0,2:0,3:0};
  return `<div class="count-input-grid">
    ${[0,1,2,3].map(value => {
      const inputId = `${taskKey}_${dimension}_count_${value}`;
      const cur     = dimCounts[value] || 0;
      return `<div class="count-input-item">
        <label for="${inputId}">${value}</label>
        <input type="number" id="${inputId}" min="0" max="${appState.workshopParticipants}"
               value="${cur}" data-action="update-workshop-count"
               data-task-key="${taskKey}" data-dimension="${dimension}" data-scale="${value}"
               oninput="this.value=Math.max(0,Math.min(${appState.workshopParticipants},parseInt(this.value)||0))">
      </div>`;
    }).join('')}
  </div>
  <div class="validation-warning" id="warning_${taskKey}_${dimension}"></div>`;
}

// ── Rating Update Handlers ────────────────────────────────────

export function updateRating(taskKey, dimension, value) {
  if (!appState.verificationRatings[taskKey]) {
    const meta = appState.taskMetadata[taskKey] || {};
    appState.verificationRatings[taskKey] = {
      dutyId: meta.dutyId, dutyTitle: meta.dutyTitle, taskTitle: meta.taskTitle,
      importance: null, frequency: null, difficulty: null,
      performsTask: false, criticality: null, comments: ''
    };
  }
  appState.verificationRatings[taskKey][dimension] = parseInt(value);
  updateComputedValues(taskKey);
}

export function updatePerformsTask(taskKey, value) {
  if (!appState.verificationRatings[taskKey]) {
    const meta = appState.taskMetadata[taskKey] || {};
    appState.verificationRatings[taskKey] = {
      dutyId: meta.dutyId, dutyTitle: meta.dutyTitle, taskTitle: meta.taskTitle,
      importance: null, frequency: null, difficulty: null,
      performsTask: false, criticality: null, comments: ''
    };
  }
  appState.verificationRatings[taskKey].performsTask = value;
}

export function updateComments(taskKey, value) {
  if (!appState.verificationRatings[taskKey]) {
    const meta = appState.taskMetadata[taskKey] || {};
    appState.verificationRatings[taskKey] = {
      dutyId: meta.dutyId, dutyTitle: meta.dutyTitle, taskTitle: meta.taskTitle,
      importance: null, frequency: null, difficulty: null,
      performsTask: false, criticality: null, comments: ''
    };
  }
  appState.verificationRatings[taskKey].comments = value;
}

function updateComputedValues(taskKey) {
  const ratings     = appState.verificationRatings[taskKey];
  if (!ratings) return;
  const isExtended  = appState.workflowMode === 'extended';
  if (!isExtended) {
    const scoreEl      = document.getElementById(`score_${taskKey}`);
    const completionEl = document.getElementById(`completion_${taskKey}`);
    if (scoreEl && completionEl) {
      const isComplete = ratings.importance !== null && ratings.frequency !== null && ratings.difficulty !== null;
      scoreEl.textContent = isComplete ? (ratings.importance + ratings.frequency + ratings.difficulty) : '-';
      completionEl.innerHTML = `<span class="completion-indicator ${isComplete ? 'complete' : 'incomplete'}">${isComplete ? '✓ Complete' : '○ Incomplete'}</span>`;
    }
  } else {
    const weightedEl = document.getElementById(`weighted_${taskKey}`);
    const priorityEl = document.getElementById(`priority_${taskKey}`);
    if (weightedEl && priorityEl) {
      if (ratings.importance !== null && ratings.frequency !== null &&
          ratings.difficulty !== null && ratings.criticality !== null) {
        const ws = (ratings.importance * ratings.frequency) + ratings.difficulty + ratings.criticality;
        weightedEl.textContent = ws;
        let pl = ws >= 10 ? 'high' : ws >= 6 ? 'medium' : 'low';
        priorityEl.innerHTML = `<span class="priority-badge ${pl}">${pl.toUpperCase()}</span>`;
      } else {
        weightedEl.textContent = '-';
        priorityEl.innerHTML = `<span class="priority-badge low">LOW</span>`;
      }
    }
  }
}

// ── Workshop Count Handler ────────────────────────────────────

export function updateWorkshopCount(taskKey, dimension, value, count) {
  if (!appState.workshopCounts[taskKey]) {
    const meta = appState.taskMetadata[taskKey] || {};
    let dutyId = meta.dutyId, dutyTitle = meta.dutyTitle, taskTitle = meta.taskTitle;
    if (!dutyTitle) {
      const parts   = taskKey.split('_task_');
      dutyId        = parts[0];
      const dutyEl  = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`);
      dutyTitle     = dutyEl ? dutyEl.value.trim() : '';
      const taskEl  = document.querySelector(`input[data-task-id="${taskKey}"], textarea[data-task-id="${taskKey}"]`);
      taskTitle     = taskEl ? taskEl.value.trim() : '';
    }
    appState.workshopCounts[taskKey] = {
      dutyId, dutyTitle, taskTitle,
      importanceCounts:   {0:0,1:0,2:0,3:0},
      frequencyCounts:    {0:0,1:0,2:0,3:0},
      difficultyCounts:   {0:0,1:0,2:0,3:0},
      criticalityCounts:  {0:0,1:0,2:0,3:0}
    };
  }
  appState.workshopCounts[taskKey][`${dimension}Counts`][value] = parseInt(count) || 0;
  validateAndComputeTask(taskKey);
}

// ── Validation & Computation ──────────────────────────────────

export function validateAndComputeTask(taskKey) {
  if (!appState.workshopCounts[taskKey]) return;
  const counts      = appState.workshopCounts[taskKey];
  const isExtended  = appState.workflowMode === 'extended';

  const iSum = Object.values(counts.importanceCounts).reduce((a,b) => a+b, 0);
  const fSum = Object.values(counts.frequencyCounts).reduce((a,b) => a+b, 0);
  const dSum = Object.values(counts.difficultyCounts).reduce((a,b) => a+b, 0);

  const iErr = iSum > appState.workshopParticipants;
  const fErr = fSum > appState.workshopParticipants;
  const dErr = dSum > appState.workshopParticipants;
  const iWarn = iSum < appState.workshopParticipants && iSum > 0;
  const fWarn = fSum < appState.workshopParticipants && fSum > 0;
  const dWarn = dSum < appState.workshopParticipants && dSum > 0;

  let cSum = 0, cErr = false, cWarn = false;
  if (isExtended) {
    cSum  = Object.values(counts.criticalityCounts).reduce((a,b) => a+b, 0);
    cErr  = cSum > appState.workshopParticipants;
    cWarn = cSum < appState.workshopParticipants && cSum > 0;
  }

  showValidationMessage(taskKey, 'importance',  iErr, iWarn, iSum);
  showValidationMessage(taskKey, 'frequency',   fErr, fWarn, fSum);
  showValidationMessage(taskKey, 'difficulty',  dErr, dWarn, dSum);
  if (isExtended) showValidationMessage(taskKey, 'criticality', cErr, cWarn, cSum);

  const hasErrors   = iErr || fErr || dErr || (isExtended ? cErr : false);
  const canCalc     = !hasErrors && iSum > 0 && fSum > 0 && dSum > 0 && (isExtended ? cSum > 0 : true);

  if (canCalc) {
    const mI = calculateWeightedMean(counts.importanceCounts);
    const mF = calculateWeightedMean(counts.frequencyCounts);
    const mD = calculateWeightedMean(counts.difficultyCounts);
    const mC = isExtended ? calculateWeightedMean(counts.criticalityCounts) : null;

    if (mI !== null && mF !== null && mD !== null && (isExtended ? mC !== null : true)) {
      const pi = appState.priorityFormula === 'ifd' ? mI * mF * mD : mI * mF;
      appState.workshopResults[taskKey] = {
        dutyId: counts.dutyId, dutyTitle: counts.dutyTitle, taskTitle: counts.taskTitle,
        meanImportance: mI, meanFrequency: mF, meanDifficulty: mD,
        meanCriticality: mC || 0, priorityIndex: pi, valid: true,
        responseCount: { importance: iSum, frequency: fSum, difficulty: dSum, criticality: cSum }
      };
      if (appState.workflowMode === 'standard') updateWorkshopTaskDisplay(taskKey);
      refreshDashboard();
    } else {
      if (appState.workshopResults[taskKey]) appState.workshopResults[taskKey].valid = false;
    }
  } else {
    if (appState.workshopResults[taskKey]) appState.workshopResults[taskKey].valid = false;
  }
}

function calculateWeightedMean(counts) {
  let wSum = 0, total = 0;
  for (let v = 0; v <= 3; v++) {
    const c = counts[v] || 0;
    wSum  += c * v;
    total += c;
  }
  return total > 0 ? wSum / total : null;
}

function showValidationMessage(taskKey, dimension, isError, isWarning, sum) {
  const el = document.getElementById(`warning_${taskKey}_${dimension}`);
  if (!el) return;
  if (isError) {
    el.innerHTML = `<p>❌ ERROR: Total responses (${sum}) exceeds ${appState.workshopParticipants} participants. Cannot calculate.</p>`;
    el.className = 'validation-warning show error';
  } else if (isWarning) {
    el.innerHTML = `<p>⚠️ WARNING: Only ${sum} of ${appState.workshopParticipants} participants responded. Calculation will use available responses.</p>`;
    el.className = 'validation-warning show warning';
  } else {
    el.className = 'validation-warning';
  }
}

function updateWorkshopTaskDisplay(taskKey) {
  const res = appState.workshopResults[taskKey];
  const mI = document.getElementById(`mean_imp_${taskKey}`);
  const mF = document.getElementById(`mean_freq_${taskKey}`);
  const mD = document.getElementById(`mean_diff_${taskKey}`);
  const pI = document.getElementById(`priority_${taskKey}`);
  if (res && res.valid) {
    if (mI) mI.textContent = res.meanImportance  !== null ? res.meanImportance.toFixed(2)  : 'N/A';
    if (mF) mF.textContent = res.meanFrequency   !== null ? res.meanFrequency.toFixed(2)   : 'N/A';
    if (mD) mD.textContent = res.meanDifficulty  !== null ? res.meanDifficulty.toFixed(2)  : 'N/A';
    if (pI) pI.textContent = res.priorityIndex   !== null ? res.priorityIndex.toFixed(2)   : 'N/A';
  } else {
    [mI, mF, mD, pI].forEach(el => { if (el) el.textContent = '-'; });
  }
}

export function validateAndComputeWorkshopResults() {
  if (appState.collectionMode !== 'workshop') return;
  Object.keys(appState.workshopCounts).forEach(k => validateAndComputeTask(k));
}

// ── Dashboard ─────────────────────────────────────────────────

// ── Project selector for dashboard ───────────────────────────

/** Re-populate the project dropdown above the dashboard */
export function renderDashboardProjectSelector() {
  const sel = document.getElementById('dashboardProjectSelector');
  if (!sel) return;

  let projects = [];
  try { projects = JSON.parse(localStorage.getItem('dacum_projects') || '[]'); } catch(e) {}
  const activeId = localStorage.getItem('dacum_active_project') || '';

  // Show ALL projects (user may want to switch even without prior results)
  // Sort newest first (projects are stored oldest-first, so reverse)
  const sorted = projects.slice().reverse();

  sel.innerHTML = '';

  sorted.forEach(p => {
    const opt        = document.createElement('option');
    opt.value        = p.id;
    const dCount     = (p.state?.dutiesData || []).length;
    const hasResults = p.state?.workshopResults && Object.keys(p.state.workshopResults).length > 0;
    const tag        = hasResults ? ' 🗳️' : '';
    opt.textContent  = `${p.name}${tag}`;
    sel.appendChild(opt);
  });

  // Always pre-select the active project
  if (activeId) sel.value = activeId;
  else if (sorted.length > 0) sel.value = sorted[0].id;

  // Show selector only when there is more than one project
  const wrap = document.getElementById('dashboardProjectSelectorWrap');
  if (wrap) wrap.style.display = sorted.length > 1 ? 'flex' : 'none';
}

export function refreshDashboard() {
  if (appState.collectionMode !== 'workshop') return;

  // NOTE: renderDashboardProjectSelector() is NOT called here —
  // it's only called from the project-loaded event and explicit selector renders,
  // so user's manual dropdown selection is never overwritten.

  // Determine which results to display from the selected project
  const sel        = document.getElementById('dashboardProjectSelector');
  const selectedId = sel?.value || '';
  const activeId   = localStorage.getItem('dacum_active_project') || '';

  let resultsSource = appState.workshopResults;
  let formulaSource = appState.priorityFormula;
  let projectLabel  = '';
  // Context for letter-code lookup: the dutiesData array that defines
  // the canonical order for this snapshot.  Active project → live
  // appState.dutiesData; foreign project → that project's saved copy.
  let dutiesContext = appState.dutiesData || [];

  // If selector points to a different project than the active one, read from storage
  if (selectedId && selectedId !== activeId) {
    try {
      const projects = JSON.parse(localStorage.getItem('dacum_projects') || '[]');
      const proj = projects.find(p => p.id === selectedId);
      if (proj?.state?.workshopResults) {
        resultsSource = proj.state.workshopResults;
        formulaSource = proj.state.priorityFormula || 'if';
        projectLabel  = proj.name;
        dutiesContext = proj.state.dutiesData || [];
      }
    } catch(e) {}
  }

  // Build a position lookup from the chosen context — one pass,
  // O(1) retrieval later inside the map loop below.
  const codeIndex = _buildCodeIndex(dutiesContext);

  const validResults = [];
  Object.keys(resultsSource || {}).forEach(taskKey => {
    const result = resultsSource[taskKey];
    if (result && result.valid) {
      let dutyText = result.dutyTitle, taskText = result.taskTitle;
      if (!dutyText || !taskText) {
        const parts  = taskKey.split('_task_');
        const dutyId = parts[0];
        if (!dutyText) { const el = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`); dutyText = el ? el.value.trim() : 'Unassigned Duty'; }
        if (!taskText) { const el = document.querySelector(`input[data-task-id="${taskKey}"], textarea[data-task-id="${taskKey}"]`); taskText = el ? el.value.trim() : 'Unassigned Task'; }
      }
      const dutyId    = result.dutyId || taskKey.split('_task_')[0];
      const dutyCode  = codeIndex.dutyLetters[dutyId] || '';
      const taskCode  = codeIndex.taskCodes[taskKey]   || '';
      const dutyLabel = dutyCode ? `Duty ${dutyCode}: ${dutyText}` : dutyText;
      const taskLabel = taskCode ? `${taskCode} — ${taskText}`    : taskText;
      validResults.push({
        duty: dutyLabel, task: taskLabel,
        dutyCode, taskCode,
        rawDuty: dutyText, rawTask: taskText,
        meanImportance: result.meanImportance, meanFrequency: result.meanFrequency,
        meanDifficulty: result.meanDifficulty, priorityIndex: result.priorityIndex
      });
    }
  });
  validResults.sort((a,b) => b.priorityIndex - a.priorityIndex);

  const summaryEl = document.getElementById('dashboardSummary');
  if (summaryEl) {
    if (validResults.length > 0) {
      const avgPriority       = validResults.reduce((s,r) => s + r.priorityIndex, 0) / validResults.length;
      const threshold         = validResults[Math.floor(validResults.length * 0.3)]?.priorityIndex || 0;
      const highPriorityCount = validResults.filter(r => r.priorityIndex >= threshold).length;
      const labelHtml = projectLabel
        ? `<div style="grid-column:1/-1;text-align:center;margin-bottom:8px;font-size:0.82em;color:#667eea;font-weight:600;">Showing results for: ${escapeHtml(projectLabel)}</div>`
        : '';
      summaryEl.innerHTML = `${labelHtml}
        <div class="summary-card"><h4>Tasks Verified</h4><p>${validResults.length}</p></div>
        <div class="summary-card"><h4>Avg Priority Index</h4><p>${avgPriority.toFixed(2)}</p></div>
        <div class="summary-card"><h4>High Priority Tasks</h4><p>${highPriorityCount} tasks</p></div>
        <div class="summary-card"><h4>Formula Used</h4><p>${formulaSource === 'if' ? 'I × F' : 'I × F × D'}</p></div>`;
    } else {
      summaryEl.innerHTML = '';
    }
  }
  updateDashboardTable(validResults);
  updateDutyLevelSummaryFromSource(resultsSource);
}

function updateDashboardTable(results) {
  const tableBody = document.getElementById('dashboardTableBody');
  if (!tableBody) return;
  if (results.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:40px;color:#999;">No valid task data available.</td></tr>`;
    return;
  }
  const topThreshold = results[Math.floor(results.length * 0.3)]?.priorityIndex || 0;
  tableBody.innerHTML = results.map((r, index) => {
    const rank  = index + 1;
    const isTop = r.priorityIndex >= topThreshold && rank <= Math.ceil(results.length * 0.3);
    return `<tr class="${isTop ? 'high-priority' : ''}">
      <td><span class="rank-badge ${isTop ? 'top' : ''}">#${rank}</span></td>
      <td>${escapeHtml(r.duty)}</td><td>${escapeHtml(r.task)}</td>
      <td><span class="mean-value">${r.meanImportance.toFixed(2)}</span></td>
      <td><span class="mean-value">${r.meanFrequency.toFixed(2)}</span></td>
      <td><span class="mean-value">${r.meanDifficulty.toFixed(2)}</span></td>
      <td><span class="priority-index">${r.priorityIndex.toFixed(2)}</span></td>
    </tr>`;
  }).join('');
}

export function toggleDashboard() {
  document.querySelector('.results-dashboard-header').classList.toggle('active');
  document.getElementById('dashboardContent').classList.toggle('active');
}

export function updateDutyLevelSummary() {
  updateDutyLevelSummaryFromSource(appState.workshopResults);
}

export function toggleDutyLevelSummary() {
  const content = document.getElementById('dutyLevelContent');
  const toggle  = document.getElementById('dutyLevelToggle');
  if (content.style.display === 'none') { content.style.display = 'block'; toggle.textContent = '▼'; }
  else { content.style.display = 'none'; toggle.textContent = '▶'; }
}

function updateDutyLevelSummaryFromSource(resultsSource) {
  if (!resultsSource) resultsSource = appState.workshopResults;
  const tableBody = document.getElementById('dutyLevelTableBody');
  if (!tableBody) return;

  // Build a code index from whichever context matches the source:
  //   – live workshopResults → live appState.dutiesData
  //   – foreign project results → that project's stored dutiesData
  // We detect foreign source by identity compare.  If source is NOT
  // the live workshopResults, walk the stored projects to find the one
  // whose state.workshopResults === resultsSource, then use its duties.
  let dutiesContext = appState.dutiesData || [];
  if (resultsSource !== appState.workshopResults) {
    try {
      const sel = document.getElementById('dashboardProjectSelector');
      const selectedId = sel?.value;
      if (selectedId) {
        const projects = JSON.parse(localStorage.getItem('dacum_projects') || '[]');
        const proj = projects.find(p => p.id === selectedId);
        if (proj?.state?.dutiesData) dutiesContext = proj.state.dutiesData;
      }
    } catch (e) {}
  }
  const codeIndex = _buildCodeIndex(dutiesContext);

  const dutyMap = {};
  Object.keys(resultsSource).forEach(taskKey => {
    const result = resultsSource[taskKey];
    if (result && result.valid) {
      let dutyId    = result.dutyId || taskKey.split('_task_')[0];
      let dutyTitle = result.dutyTitle;
      if (!dutyTitle) { const el = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`); dutyTitle = el ? el.value.trim() : 'Unassigned Duty'; }
      if (!dutyMap[dutyId]) dutyMap[dutyId] = { dutyTitle, dutyCode: codeIndex.dutyLetters[dutyId] || '', totalTasks:0, validTasks:0, importanceSum:0, frequencySum:0, difficultySum:0, prioritySum:0, tasks:[] };
      const d = dutyMap[dutyId];
      d.totalTasks++; d.validTasks++;
      d.importanceSum += result.meanImportance; d.frequencySum += result.meanFrequency;
      d.difficultySum += result.meanDifficulty; d.prioritySum  += result.priorityIndex;
      d.tasks.push({ priorityIndex: result.priorityIndex, meanDifficulty: result.meanDifficulty });
    }
  });

  const dutyResults = [];
  Object.keys(dutyMap).forEach(dutyId => {
    const duty = dutyMap[dutyId];
    const vc   = duty.validTasks;
    if (vc > 0) {
      const avgI = duty.importanceSum / vc, avgF = duty.frequencySum / vc;
      const avgD = duty.difficultySum / vc, avgP = duty.prioritySum  / vc;
      let trainingLoad = appState.trainingLoadMethod === 'advanced'
        ? duty.tasks.reduce((s,t) => s + t.priorityIndex * t.meanDifficulty, 0)
        : avgP * vc;
      const sorted   = duty.tasks.sort((a,b) => b.priorityIndex - a.priorityIndex);
      const thr      = sorted[Math.floor(sorted.length * 0.3)]?.priorityIndex || 0;
      const highPC   = sorted.filter(t => t.priorityIndex >= thr).length;
      dutyResults.push({ dutyTitle: duty.dutyTitle, dutyCode: duty.dutyCode, totalTasks: duty.totalTasks, validTasks: vc, avgImportance: avgI, avgFrequency: avgF, avgDifficulty: avgD, avgPriority: avgP, highPriorityCount: highPC, trainingLoad });
    }
  });

  // Sort: if sortBy is 'code' (letter order), use dutyCode; else priority/trainingLoad as before
  const sortBy = document.getElementById('dutySortSelector')?.value || 'priority';
  dutyResults.sort((a,b) => sortBy === 'trainingLoad' ? b.trainingLoad - a.trainingLoad : b.avgPriority - a.avgPriority);

  if (dutyResults.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;">No valid duty data available.</td></tr>`;
    return;
  }
  tableBody.innerHTML = dutyResults.map(duty => `<tr>
    <td><strong>${duty.dutyCode ? `Duty ${duty.dutyCode}: ` : ''}${escapeHtml(duty.dutyTitle)}</strong></td>
    <td style="text-align:center;">${duty.totalTasks}</td>
    <td style="text-align:center;">${duty.validTasks}</td>
    <td style="text-align:center;"><span class="mean-value">${duty.avgImportance.toFixed(2)}</span></td>
    <td style="text-align:center;"><span class="mean-value">${duty.avgFrequency.toFixed(2)}</span></td>
    <td style="text-align:center;"><span class="mean-value">${duty.avgDifficulty.toFixed(2)}</span></td>
    <td style="text-align:center;"><span class="priority-index">${duty.avgPriority.toFixed(2)}</span></td>
    <td style="text-align:center;">${duty.highPriorityCount}</td>
    <td style="text-align:center;"><strong style="color:#667eea;">${duty.trainingLoad.toFixed(2)}</strong></td>
  </tr>`).join('');
}

export function exportDashboard() {
  if (appState.collectionMode !== 'workshop') { showStatus('Dashboard export only available in Workshop mode', 'error'); return; }
  const codeIndex = _buildCodeIndex(appState.dutiesData || []);
  const validResults = [];
  Object.keys(appState.workshopResults).forEach(taskKey => {
    const r = appState.workshopResults[taskKey];
    if (r && r.valid) {
      let dutyText = r.dutyTitle, taskText = r.taskTitle;
      if (!dutyText || !taskText) {
        const parts  = taskKey.split('_task_');
        const dutyId = parts[0];
        if (!dutyText) { const el = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`); dutyText = el ? el.value.trim() : 'Unassigned Duty'; }
        if (!taskText) { const el = document.querySelector(`input[data-task-id="${taskKey}"], textarea[data-task-id="${taskKey}"]`); taskText = el ? el.value.trim() : 'Unassigned Task'; }
      }
      const dutyId   = r.dutyId || taskKey.split('_task_')[0];
      const dutyCode = codeIndex.dutyLetters[dutyId] || '';
      const taskCode = codeIndex.taskCodes[taskKey]   || '';
      validResults.push({ dutyCode, taskCode, duty: dutyText, task: taskText, meanImportance: r.meanImportance, meanFrequency: r.meanFrequency, meanDifficulty: r.meanDifficulty, priorityIndex: r.priorityIndex });
    }
  });
  validResults.sort((a,b) => b.priorityIndex - a.priorityIndex);
  let csv = 'Rank,Duty Code,Duty,Task Code,Task,Mean Importance,Mean Frequency,Mean Difficulty,Priority Index\n';
  validResults.forEach((r,i) => {
    csv += `${i+1},"${r.dutyCode}","${r.duty.replace(/"/g,'""')}","${r.taskCode}","${r.task.replace(/"/g,'""')}",${r.meanImportance.toFixed(2)},${r.meanFrequency.toFixed(2)},${r.meanDifficulty.toFixed(2)},${r.priorityIndex.toFixed(2)}\n`;
  });
  const blob = new Blob([csv], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href  = url;
  const occ  = document.getElementById('occupationTitle').value || 'DACUM';
  link.download = `${occ.replace(/[^a-z0-9]/gi,'_')}_Workshop_Results_${new Date().toISOString().split('T')[0]}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  showStatus('Dashboard exported as CSV successfully! ✓', 'success');
}

// ── Accordion Listeners ───────────────────────────────────────

export function attachAccordionListeners() {
  document.querySelectorAll('.duty-accordion-header').forEach(header => {
    header.addEventListener('click', function () {
      const isActive = this.classList.contains('active');
      document.querySelectorAll('.duty-accordion-header').forEach(h => h.classList.remove('active'));
      document.querySelectorAll('.duty-accordion-content').forEach(c => c.classList.remove('active'));
      if (!isActive) {
        this.classList.add('active');
        this.nextElementSibling.classList.add('active');
      }
    });
  });
}

// ── Auto-refresh dashboard when project switches ──────────────
// Listens for the custom event fired by dacum_projects.js loadProject()
document.addEventListener('dacum:project-loaded', function (e) {
  setTimeout(function () {
    const projectId = e.detail?.projectId;
    const activeId  = localStorage.getItem('dacum_active_project') || '';
    const target    = projectId || activeId;

    // Rebuild options list
    renderDashboardProjectSelector();

    // Force dropdown to the newly loaded project
    const sel = document.getElementById('dashboardProjectSelector');
    if (sel && target) {
      const opt = sel.querySelector(`option[value="${target}"]`);
      if (opt) sel.value = target;
    }

    // Refresh dashboard content
    refreshDashboard();
  }, 50);
});

// ── Internal helper ───────────────────────────────────────────
//
// Build an O(1) lookup of duty & task codes from a given dutiesData
// array.  Works for both the live appState.dutiesData and foreign
// project snapshots read from localStorage.
//
// Returns: { dutyLetters: { dutyId -> 'A' }, taskCodes: { taskInputId -> 'A1' } }
function _buildCodeIndex(dutiesData) {
  const dutyLetters = {};
  const taskCodes   = {};
  (dutiesData || []).forEach((duty, dutyIndex) => {
    if (!duty) return;
    const letter = getDutyLetter(dutyIndex);
    if (duty.id) dutyLetters[duty.id] = letter;
    (duty.tasks || []).forEach((task, taskIndex) => {
      if (task && task.inputId) {
        taskCodes[task.inputId] = `${letter}${taskIndex + 1}`;
      }
    });
  });
  return { dutyLetters, taskCodes };
}
