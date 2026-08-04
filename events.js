// ============================================================
// /events.js
// All addEventListener calls and delegated event handling.
// Replaces all inline onclick="..." attributes.
// ============================================================

import { appState }             from './state.js';
import { generateAdditionalInfoAI } from './additional_info_ai.js';
import { addDuty, addTask, removeDuty, removeTask, clearDuty,
         syncAllFromDOM, syncDutyTitle, syncTaskText,
         toggleViewMode, switchToViewMode }            from './duties.js';
import { pushHistoryState, undo, redo,
         resetHistoryToCurrentState }                  from './history.js';
import { updateCollectionMode, updateWorkflowMode, updateParticipantCount,
  updatePriorityFormula, updateTVExportMode, updateTrainingLoadMethod,
  loadDutiesForVerification, updateRating, updatePerformsTask, updateComments,
  updateWorkshopCount, validateAndComputeTask, validateAndComputeWorkshopResults,
  toggleDashboard, refreshDashboard, toggleDutyLevelSummary, exportDashboard,
  renderDashboardProjectSelector,
  attachAccordionListeners }    from './tasks.js';
import { bypassToClusteringTab, resetVerificationDecision, initializeClusteringFromTasks,
  updateCreateClusterButton, createCluster, renameCluster, deleteCluster,
  removeTaskFromCluster, addTaskToClusterFromDropdown, updateClusterRange,
  updateClusterCriteria, updateClusterCriteriaFromNumbered,
  handleCriteriaKeydown, initCriteriaNumber, proceedToClusteringFromVerification,
  updateCreateLOButton, createLearningOutcome, toggleEditLO, deleteLearningOutcome,
  reassignPCToLO, unassignPCFromLO,
  updateCreateModuleButton, createModule, renameModule, deleteModule,
  removeLoFromModule, addLoToModuleFromDropdown,
  openModuleBuilderFromMapping, exportModuleMappingJSON }  from './modules.js';
import { showStatus, toggleInfoBox, escapeHtml,
  toggleSkillsLevelSection, addSkillsCategory, removeSkillsCategory,
  updateSkillsCategoryName, addSkillsCompetency, removeSkillsCompetency,
  updateSkillsCompetencyText, handleSkillsLevelChange, resetSkillsLevel,
  toggleEditHeading, clearSection, formatList,
  addCustomSection, removeCustomSection }                  from './renderer.js';
import { exportToPDF, exportToWord,
  exportTaskVerificationPDF, exportTaskVerificationWord }  from './exports.js';
import { clearAll, clearAllSilent, clearCurrentTab, generateAIDacum } from './projects.js';
import { handleImageUpload, removeImage }                  from './storage.js';
import { saveToJSON, loadFromJSON }                        from './snapshots.js';
import { saveSnapshot, restoreSnapshot,
         deleteSnapshot,
         renderSnapshotPanel }                             from './workshop_snapshots.js';
import { saveCurrentProject,
         loadProject,
         renderProjectsSidebar,
         deleteActiveProject }                              from './dacum_projects.js';
import { lwFinalizeAndCreateSession, lwCopyLink, lwShowQRCode,
  lwCloseQRModal, lwDownloadQRPNG, lwFetchResults,
  lwExportJSON, lwExportCSV, lwExportSnapshot,
  lwCloseVoting, lwExportVerifiedPDF, lwExportVerifiedDOCX }  from './workshop.js';
import { markAiGenerated, refineResults,
         clearAiGeneratedFlag }                               from './refine.js';

// ── Delegation helper ─────────────────────────────────────────

function delegate(container, selector, eventType, handler) {
  if (!container) return;
  container.addEventListener(eventType, function (e) {
    const target = e.target.closest(selector);
    if (target && container.contains(target)) {
      handler(e, target);
    }
  });
}

// ── Setup all events ──────────────────────────────────────────

export function setupEvents() {
  // ── Static buttons ────────────────────────────────────────

  _on('btnAddDuty', 'click', () => { syncAllFromDOM(); pushHistoryState(); addDuty(); });
  _on('btnSaveJSON',  'click', () => saveToJSON());
  _on('btnClearAll',  'click', () => {
    // clearAll() shows the confirm dialog and returns false if cancelled
    const cleared = clearAll();
    if (!cleared) return;

    // Remove the active project card from sidebar —
    // the state is now empty so there is nothing to save
    deleteActiveProject();

    // Reset undo/redo stack to the empty state
    resetHistoryToCurrentState();

    // Hide the Refine Results card — there is nothing AI-generated to refine
    clearAiGeneratedFlag();
  });

  // Undo / Redo buttons
  _on('btnUndo', 'click', () => undo());
  _on('btnRedo', 'click', () => redo());

  // Keyboard shortcuts (active only on duties tab)
  document.addEventListener('keydown', function (e) {
    const dutiesTab = document.getElementById('duties-tab');
    if (!dutiesTab || !dutiesTab.classList.contains('active')) return;
    if (e.ctrlKey && !e.shiftKey && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); }
    if (e.ctrlKey && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) { e.preventDefault(); redo(); }
  });

  // Save Snapshot button (toolbar)
  _on('btnSaveSnapshot', 'click', () => {
    const name = prompt('Enter a name for this snapshot:', _defaultSnapName());
    if (name) saveSnapshot(name);
  });

  // Snapshot lens button — open modal
  _on('btnOpenSnapshots', 'click', () => {
    renderSnapshotPanel();
    const m = document.getElementById('snapModal');
    if (m) m.style.display = 'block';
  });
  // Modal save button
  _on('snapModalSave', 'click', () => {
    const name = prompt('Snapshot name:', _defaultSnapName());
    if (name) { saveSnapshot(name); renderSnapshotPanel(); }
  });
  // Close modal
  _on('snapModalClose', 'click', () => {
    const m = document.getElementById('snapModal');
    if (m) m.style.display = 'none';
  });
  _on('snapModalOverlay', 'click', () => {
    const m = document.getElementById('snapModal');
    if (m) m.style.display = 'none';
  });
  _on('aiGenerateBtn', 'click', () => {
    generateAIDacum()
      .then((ok) => {
        resetHistoryToCurrentState();
        saveCurrentProject();
        renderProjectsSidebar();
        // Reveal the Refine Results card only if generation actually ran
        if (ok) markAiGenerated();
      })
      .catch(() => {});
  });

  // ── Tab "?" help modals ─────────────────────────────────────
  _on('addInfoHelpBtn',    'click', () => _showAdditionalInfoHelp());
  _on('clusteringHelpBtn',    'click', () => _showClusteringHelp());
  _on('clusterNamingHelpBtn', 'click', () => _showClusterNamingHelp());

  // ── Additional Information tab: AI supporting-info generation ──
  // No markAiGenerated() here — the Refine Results card operates on
  // duties/tasks only, so surfacing it after this run would be wrong.
  // The generated text lives in plain textareas that saveCurrentProject()
  // already reads straight from the DOM, so a save is all that's needed.
  _on('aiGenerateInfoBtn', 'click', () => {
    generateAdditionalInfoAI()
      .then((ok) => {
        if (!ok) return;
        saveCurrentProject();
        renderProjectsSidebar();
      })
      .catch(() => {});
  });

  // ── "Generate Anyway" path (from the yellow Scope warning card) ──
  // projects.js dispatches this event after a successful generation
  // that bypassed the scope gate.  We mirror the main Generate button's
  // post-success behaviour so Refine Results + history reset + project
  // save + sidebar render all happen identically.
  document.addEventListener('dacum:ai-generated', () => {
    resetHistoryToCurrentState();
    saveCurrentProject();
    renderProjectsSidebar();
    markAiGenerated();
  });
  // Refine Results — runs soft cleanup on AI-generated duties/tasks
  _on('btnRefineResults', 'click', () => refineResults());
  _on('btnExportPDF',          'click', () => exportToPDF());
  _on('btnExportWord',         'click', () => exportToWord());

  // Task Verification controls
  _onRadioGroup('collectionMode',  () => { updateCollectionMode(); });
  _onRadioGroup('workflowMode',    () => updateWorkflowMode());
  _onRadioGroup('priorityFormula', () => updatePriorityFormula());
  _onRadioGroup('tvExportMode',    () => updateTVExportMode());
  _onRadioGroup('trainingLoadMethod', () => updateTrainingLoadMethod());
  _on('workshopParticipants', 'change', () => updateParticipantCount());
  _on('btnLoadDutiesForVerification', 'click', () => loadDutiesForVerification());
  _on('btnValidateAll',        'click', () => validateAndComputeWorkshopResults());
  _on('btnToggleDashboard',    'click', () => toggleDashboard());
  _on('btnRefreshDashboard',   'click', () => refreshDashboard());
  // Project selector — load the selected project (updates sidebar, session card, dashboard)
  _on('dashboardProjectSelector', 'change', () => {
    const sel = document.getElementById('dashboardProjectSelector');
    if (sel && sel.value) loadProject(sel.value);
  });
  _on('btnToggleDutyLevelSummary', 'click', () => toggleDutyLevelSummary());
  _on('btnExportDashboard',    'click', () => exportDashboard());
  _on('btnExportTVPDF',        'click', () => exportTaskVerificationPDF());
  _on('btnExportTVWord',       'click', () => exportTaskVerificationWord());

  // Clustering
  _on('btnBypassToClustering',      'click', () => bypassToClusteringTab());
  _on('btnResetDecision',           'click', () => resetVerificationDecision());
  _on('btnProceedToClustering',     'click', () => proceedToClusteringFromVerification());
  _on('btnCreateCluster',           'click', () => createCluster());

  // Learning Outcomes
  _on('btnCreateLO',                'click', () => createLearningOutcome());

  // Module Mapping
  _on('btnCreateModule',            'click', () => createModule());
  _on('btnOpenModuleBuilder',       'click', () => openModuleBuilderFromMapping());
  _on('btnExportModuleMapping',     'click', () => exportModuleMappingJSON());

  // Skills Level
  _on('btnToggleSkillsLevel',       'click', () => toggleSkillsLevelSection());
  _on('btnAddSkillsCategory',       'click', () => addSkillsCategory());
  _on('btnResetSkillsLevel',        'click', () => resetSkillsLevel());

  // Additional Info
  _on('btnAddCustomSection',        'click', () => addCustomSection());
  _on('btnToggleInfoBox',           'click', () => toggleInfoBox());

  // Image upload
  _on('producedForImageInput',      'change', (e) => handleImageUpload(e, 'producedFor'));
  _on('producedByImageInput',       'change', (e) => handleImageUpload(e, 'producedBy'));
  _on('removeProducedForImage',     'click',  () => removeImage('producedFor'));
  _on('removeProducedByImage',      'click',  () => removeImage('producedBy'));

  // JSON Save/Load
  _on('loadFileInput', 'change', (e) => {
    loadFromJSON(e);
    // Project creation, sidebar render, and resetHistory are now
    // handled inside loadFromJSON via importProjectFromData + loadProject
  });

  // Live Workshop
  _on('btnLWFinalize',              'click', () => lwFinalizeAndCreateSession());
  _on('btnLWCopyLink',              'click', () => lwCopyLink());
  _on('btnLWShowQR',                'click', () => lwShowQRCode());
  _on('btnLWFetchResults',          'click', () => lwFetchResults());
  _on('btnLWExportJSON',            'click', () => lwExportJSON());
  _on('btnLWExportCSV',             'click', () => lwExportCSV());
  _on('btnLWExportSnapshot',        'click', () => lwExportSnapshot());
  _on('btnLWCloseVoting',           'click', () => lwCloseVoting());
  _on('btnLWExportVerifiedPDF',     'click', () => lwExportVerifiedPDF());
  _on('btnLWExportVerifiedDOCX',    'click', () => lwExportVerifiedDOCX());
  _on('btnLWCloseQR',               'click', () => lwCloseQRModal());

  // Register refreshDashboard callback so workshop.js can trigger it after live votes arrive
  appState._onResultsRefreshed = () => refreshDashboard();
  _on('btnLWDownloadQR',            'click', () => lwDownloadQRPNG());

  // Per-tab clear buttons (may use data-tab-id)
  document.querySelectorAll('[data-action="clear-tab"]').forEach(btn => {
    btn.addEventListener('click', function () {
      clearCurrentTab(this.getAttribute('data-tab-id'));
    });
  });

  // ── Delegated events ──────────────────────────────────────

  // Duties container
  const dutiesCont = document.getElementById('dutiesContainer');
  if (dutiesCont) {

    // Structural clicks — sync → push history → mutate
    dutiesCont.addEventListener('click', function (e) {
      const target = e.target;
      if (target.matches('[data-action="add-task"]')) {
        syncAllFromDOM(); pushHistoryState();
        addTask(target.getAttribute('data-duty-id'));
      } else if (target.matches('[data-action="add-duty"]')) {
        // Wall View duty cards carry their own ＋ button that adds a
        // whole new duty row (task cards' ＋ adds a task instead).
        syncAllFromDOM(); pushHistoryState();
        addDuty();
      } else if (target.matches('[data-action="remove-duty"]')) {
        pushHistoryState();           // syncAllFromDOM called inside removeDuty
        removeDuty(target.getAttribute('data-duty-id'));
      } else if (target.matches('[data-action="remove-task"]')) {
        pushHistoryState();           // syncAllFromDOM called inside removeTask
        removeTask(target.getAttribute('data-task-div-id'));
      } else if (target.matches('[data-action="clear-duty"]')) {
        pushHistoryState();           // syncAllFromDOM called inside clearDuty
        clearDuty(target.getAttribute('data-duty-id'));
      }
    });

    // View toggle button (legacy 2-state Card ↔ Table)
    _on('btnToggleDutiesView', 'click', toggleViewMode);

    // Segmented view switch (3-state: Card / Table / Wall) — delegated
    // so it works even if the buttons are re-rendered.  Any element
    // with data-view-switch="card|table|wall" triggers switchToViewMode.
    document.addEventListener('click', function (e) {
      const btn = e.target.closest('[data-view-switch]');
      if (!btn) return;
      const mode = btn.getAttribute('data-view-switch');
      if (mode === 'card' || mode === 'table' || mode === 'wall') {
        switchToViewMode(mode);
      }
    });

    // Word-level text undo (burst model) — works for input (table) and textarea (card)
    const _dutyOrTask = 'input[data-duty-id], input[data-task-id], textarea[data-duty-id], textarea[data-task-id]';
    let _burstPushed = false;
    let _burstTimer  = null;

    dutiesCont.addEventListener('focusin', function (e) {
      if (e.target.matches(_dutyOrTask)) {
        clearTimeout(_burstTimer); _burstPushed = false;
      }
    });

    dutiesCont.addEventListener('input', function (e) {
      const t = e.target;
      if      (t.matches('input[data-duty-id], textarea[data-duty-id]'))
        syncDutyTitle(t.getAttribute('data-duty-id'), t.value);
      else if (t.matches('input[data-task-id], textarea[data-task-id]'))
        syncTaskText(t.getAttribute('data-task-id'), t.value);
      else return;
      if (!_burstPushed) { pushHistoryState(); _burstPushed = true; }
      clearTimeout(_burstTimer);
      _burstTimer = setTimeout(() => { _burstPushed = false; }, 120);
    });

    dutiesCont.addEventListener('focusout', function (e) {
      if (e.target.matches(_dutyOrTask)) {
        clearTimeout(_burstTimer); _burstPushed = false;
      }
    });
  }

  // Verification accordion
  const verCont = document.getElementById('verificationAccordionContainer');
  if (verCont) {
    verCont.addEventListener('click', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      const taskKey = target.getAttribute('data-task-key');
      const dimension = target.getAttribute('data-dimension');
      if (action === 'update-rating') {
        updateRating(taskKey, dimension, target.value);
      } else if (action === 'update-performs-task') {
        updatePerformsTask(taskKey, target.value);
      } else if (action === 'update-comments') {
        updateComments(taskKey, target.value);
      } else if (action === 'update-workshop-count') {
        updateWorkshopCount(taskKey, dimension, target.getAttribute('data-scale'), parseInt(target.value));
      } else if (action === 'validate-compute-task') {
        validateAndComputeTask(taskKey);
      }
    });
    verCont.addEventListener('change', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      const taskKey = target.getAttribute('data-task-key');
      const dimension = target.getAttribute('data-dimension');
      if (action === 'update-rating') {
        updateRating(taskKey, dimension, target.value);
      } else if (action === 'update-performs-task') {
        updatePerformsTask(taskKey, target.value);
      } else if (action === 'update-comments') {
        updateComments(taskKey, target.value);
      }
    });

    // ── FIX: workshop count inputs use 'input' event (not click) ──
    // The count inputs have data-scale (not data-value) as the scale key.
    // Listening on 'input' ensures real-time validation warnings and
    // live mean calculations update as the user types or uses spinners.
    verCont.addEventListener('input', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      if (target.getAttribute('data-action') !== 'update-workshop-count') return;
      const taskKey   = target.getAttribute('data-task-key');
      const dimension = target.getAttribute('data-dimension');
      const scale     = target.getAttribute('data-scale');   // 0 | 1 | 2 | 3
      updateWorkshopCount(taskKey, dimension, scale, parseInt(target.value) || 0);
    });
  }

  // Clusters container
  const clustersCont = document.getElementById('clustersContainer');
  if (clustersCont) {
    clustersCont.addEventListener('click', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      if (action === 'rename-cluster')          renameCluster(target.getAttribute('data-cluster-id'));
      else if (action === 'delete-cluster')     deleteCluster(target.getAttribute('data-cluster-id'));
      else if (action === 'show-pc-range-help') _showPCRangeHelp();
      else if (action === 'remove-task-from-cluster') {
        removeTaskFromCluster(target.getAttribute('data-cluster-id'), parseInt(target.getAttribute('data-task-index')));
      }
    });
    clustersCont.addEventListener('change', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      if (action === 'update-cluster-range') {
        updateClusterRange(target.getAttribute('data-cluster-id'), target.value);
      }
    });
    clustersCont.addEventListener('blur', function (e) {
      const target = e.target.closest('[data-action-blur]');
      if (!target) return;
      if (target.getAttribute('data-action-blur') === 'update-cluster-criteria-numbered') {
        updateClusterCriteriaFromNumbered(target.getAttribute('data-cluster-id'), target.value);
      }
    }, true);
    clustersCont.addEventListener('focus', function (e) {
      const target = e.target.closest('[data-action-focus]');
      if (!target) return;
      if (target.getAttribute('data-action-focus') === 'init-criteria-number') {
        initCriteriaNumber(e, target.getAttribute('data-cluster-id'));
      }
    }, true);
    clustersCont.addEventListener('keydown', function (e) {
      const target = e.target.closest('[data-action-keydown]');
      if (!target) return;
      if (target.getAttribute('data-action-keydown') === 'handle-criteria-keydown') {
        handleCriteriaKeydown(e, target.getAttribute('data-cluster-id'));
      }
    });
  }

  // Available tasks list (checkboxes + dropdowns)
  const availableTasksList = document.getElementById('availableTasksList');
  if (availableTasksList) {
    availableTasksList.addEventListener('change', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      if (target.getAttribute('data-action') === 'update-cluster-button') {
        updateCreateClusterButton();
      } else if (target.getAttribute('data-action') === 'add-task-to-cluster-dropdown') {
        addTaskToClusterFromDropdown(parseInt(target.getAttribute('data-task-index')), target.value);
        target.value = '';
      }
    });
  }

  // PC Source List (LO)
  const pcSourceList = document.getElementById('pcSourceList');
  if (pcSourceList) {
    pcSourceList.addEventListener('change', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      if (target.getAttribute('data-action') === 'update-lo-button') {
        updateCreateLOButton();
      } else if (target.getAttribute('data-action') === 'reassign-pc-to-lo') {
        reassignPCToLO(
          target.getAttribute('data-pc-id'),
          parseInt(target.getAttribute('data-cluster')),
          parseInt(target.getAttribute('data-criterion')),
          target.value
        );
        target.value = '';
      }
    });
  }

  // LO blocks container
  const loBlocksCont = document.getElementById('loBlocksContainer');
  if (loBlocksCont) {
    loBlocksCont.addEventListener('click', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      if (action === 'toggle-edit-lo')      toggleEditLO(target.getAttribute('data-lo-id'));
      else if (action === 'delete-lo')      deleteLearningOutcome(target.getAttribute('data-lo-id'));
      else if (action === 'unassign-pc-from-lo') {
        unassignPCFromLO(target.getAttribute('data-lo-id'), target.getAttribute('data-pc-id'));
      }
    });
    loBlocksCont.addEventListener('blur', function (e) {
      const target = e.target.closest('[data-action-blur]');
      if (!target) return;
      // saveLOStatement handled via toggleEditLO save path
    }, true);
  }

  // Module LO list
  const moduleLoList = document.getElementById('moduleLoList');
  if (moduleLoList) {
    moduleLoList.addEventListener('change', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      if (target.getAttribute('data-action') === 'update-module-button') {
        updateCreateModuleButton();
      } else if (target.getAttribute('data-action') === 'add-lo-to-module-dropdown') {
        addLoToModuleFromDropdown(target.getAttribute('data-lo-id'), target.value);
        target.value = '';
      }
    });
  }

  // Modules container
  const modulesCont = document.getElementById('modulesContainer');
  if (modulesCont) {
    modulesCont.addEventListener('click', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      if (action === 'rename-module')          renameModule(target.getAttribute('data-module-id'));
      else if (action === 'delete-module')     deleteModule(target.getAttribute('data-module-id'));
      else if (action === 'remove-lo-from-module') {
        removeLoFromModule(target.getAttribute('data-module-id'), target.getAttribute('data-lo-id'));
      }
    });
  }

  // Custom sections container
  const customSectCont = document.getElementById('customSectionsContainer');
  if (customSectCont) {
    customSectCont.addEventListener('click', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      if (action === 'toggle-edit-heading') {
        toggleEditHeading(target.getAttribute('data-heading-id'));
      } else if (action === 'clear-section') {
        clearSection(
          target.getAttribute('data-input-id'),
          target.getAttribute('data-heading-id'),
          target.getAttribute('data-default-heading')
        );
      } else if (action === 'remove-custom-section') {
        removeCustomSection(target.getAttribute('data-section-id'));
      }
    });
  }

  // Skills Level container
  const skillsCont = document.getElementById('skillsLevelContainer');
  if (skillsCont) {
    skillsCont.addEventListener('click', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      if (action === 'remove-skills-category') {
        removeSkillsCategory(parseInt(target.getAttribute('data-cat-index')));
      } else if (action === 'add-skills-competency') {
        addSkillsCompetency(parseInt(target.getAttribute('data-cat-index')));
      } else if (action === 'remove-skills-competency') {
        removeSkillsCompetency(
          parseInt(target.getAttribute('data-cat-index')),
          parseInt(target.getAttribute('data-comp-index'))
        );
      }
    });
    skillsCont.addEventListener('change', function (e) {
      const target = e.target.closest('[data-action]');
      if (!target) return;
      const action = target.getAttribute('data-action');
      if (action === 'update-skills-category-name') {
        updateSkillsCategoryName(parseInt(target.getAttribute('data-cat-index')), target.value);
      } else if (action === 'update-skills-competency-text') {
        updateSkillsCompetencyText(
          parseInt(target.getAttribute('data-cat-index')),
          parseInt(target.getAttribute('data-comp-index')),
          target.value
        );
      } else if (action === 'handle-skills-level-change') {
        handleSkillsLevelChange(
          parseInt(target.getAttribute('data-cat-index')),
          parseInt(target.getAttribute('data-comp-index')),
          target.getAttribute('data-level'),
          target.checked
        );
      }
    });
  }

  // Additional Info static buttons (formatList / toggleEditHeading / clearSection)
  _onStaticInfoButtons();

  // When the last project is deleted via the sidebar bin button,
  // dacum_projects.js dispatches this event so we silently reset the DOM
  document.addEventListener('dacum:last-project-deleted', () => {
    clearAllSilent();
    resetHistoryToCurrentState();
  });

  // ── Snapshot panel delegation ────────────────────────────────
  const snapList = document.getElementById('snapshotList');
  if (snapList) {
    snapList.addEventListener('click', function (e) {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;
      const action = btn.getAttribute('data-action');
      const id     = btn.getAttribute('data-snap-id');
      if (action === 'restore-snapshot') restoreSnapshot(id);
      else if (action === 'delete-snapshot')  deleteSnapshot(id);
    });
  }
}

// ── Private helpers ───────────────────────────────────────────

function _on(id, event, handler) {
  const el = document.getElementById(id);
  if (el) el.addEventListener(event, handler);
}

function _defaultSnapName() {
  const occ = (document.getElementById('occupationTitle')?.value || '').trim();
  const now = new Date();
  const time = now.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
  return occ ? `${occ} — ${time}` : `Workshop checkpoint — ${time}`;
}

function _onRadioGroup(name, handler) {
  document.querySelectorAll(`input[name="${name}"]`).forEach(radio => {
    radio.addEventListener('change', handler);
  });
}

function _onStaticInfoButtons() {
  // Heading rename buttons and format buttons are in the static HTML
  document.querySelectorAll('[data-action="toggle-edit-heading"]').forEach(btn => {
    btn.addEventListener('click', function () {
      toggleEditHeading(this.getAttribute('data-heading-id'));
    });
  });
  document.querySelectorAll('[data-action="clear-section"]').forEach(btn => {
    btn.addEventListener('click', function () {
      clearSection(
        this.getAttribute('data-input-id'),
        this.getAttribute('data-heading-id'),
        this.getAttribute('data-default-heading')
      );
    });
  });
  document.querySelectorAll('[data-action="format-list"]').forEach(btn => {
    btn.addEventListener('click', function () {
      formatList(this.getAttribute('data-input-id'), this.getAttribute('data-format-type'));
    });
  });
}

// ── Tab help modals ───────────────────────────────────────────
// Small "?" buttons next to a tab heading open a guidance modal.
// These replace the permanent hint paragraphs that used to sit under
// each heading — the guidance stays available without spending
// vertical space above the fields the facilitator actually works in.
//
// _showHelpModal is generic so each tab only supplies its content;
// the chrome, animations, and close behaviour are defined once.

function _showHelpModal({ id, icon, title, intro, items, note, bodyHtml, maxWidth }) {
  const existing = document.getElementById(id);
  if (existing) { existing.remove(); return; }   // second click closes it

  const overlay = document.createElement('div');
  overlay.id = id;
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', title);
  overlay.style.cssText =
    'position:fixed;inset:0;z-index:999999;display:flex;align-items:center;' +
    'justify-content:center;padding:20px;background:rgba(0,0,0,0.55);' +
    'backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px);' +
    'animation:aiErrFadeIn 0.2s ease;';

  const rows = (items || []).map(([bullet, strong, rest]) =>
    '<li style="display:flex;gap:10px;align-items:flex-start;padding:9px 0;' +
    'border-bottom:1px solid #f1f5f9;">' +
      '<span style="font-size:1.05em;line-height:1.4;flex-shrink:0;">' + bullet + '</span>' +
      '<span style="font-size:0.87em;line-height:1.6;color:#334155;">' +
        (strong ? '<strong style="color:#1e293b;">' + strong + '</strong>' + (rest ? ' — ' : '') : '') +
        (rest || '') +
      '</span>' +
    '</li>'
  ).join('');

  const noteHtml = note
    ? '<div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;' +
      'padding:11px 13px;margin-bottom:18px;font-size:0.84em;line-height:1.6;color:#92400e;">' +
      note + '</div>'
    : '';

  overlay.innerHTML =
    '<div style="background:#fff;border-radius:16px;max-width:' + (maxWidth || '480px') + ';width:100%;' +
    'box-shadow:0 24px 60px rgba(0,0,0,0.35);overflow:hidden;' +
    'font-family:\'Segoe UI\',system-ui,sans-serif;animation:aiErrSlideIn 0.22s ease;' +
    'max-height:86vh;display:flex;flex-direction:column;">' +
      '<div style="padding:20px 22px 16px;display:flex;align-items:center;gap:12px;' +
      'background:linear-gradient(135deg,#eef2ff,#e0e7ff);border-bottom:1px solid #c7d2fe;' +
      'flex-shrink:0;">' +
        '<span style="font-size:1.7em;line-height:1;">' + icon + '</span>' +
        '<p style="margin:0;font-size:1em;font-weight:800;color:#3730a3;">' + title + '</p>' +
      '</div>' +
      '<div style="padding:18px 22px 20px;overflow-y:auto;">' +
        (intro
          ? '<p style="margin:0 0 8px;font-size:0.88em;color:#475569;line-height:1.6;">' +
            intro + '</p>'
          : '') +
        (rows ? '<ul style="list-style:none;margin:0 0 18px;padding:0;">' + rows + '</ul>' : '') +
        (bodyHtml || '') +
        noteHtml +
        '<div style="display:flex;justify-content:flex-end;">' +
          '<button data-help-close style="padding:9px 22px;background:#667eea;' +
          'color:#fff;border:none;border-radius:8px;font-size:0.9em;font-weight:700;' +
          'cursor:pointer;font-family:inherit;">Got it</button>' +
        '</div>' +
      '</div>' +
    '</div>';

  // Reuses the keyframes the AI error modals inject; create them only
  // if nothing has needed them yet this session.
  if (!document.getElementById('aiErrStyles')) {
    const st = document.createElement('style');
    st.id = 'aiErrStyles';
    st.textContent =
      '@keyframes aiErrFadeIn  { from{opacity:0} to{opacity:1} }' +
      '@keyframes aiErrSlideIn { from{transform:translateY(-14px);opacity:0}' +
      ' to{transform:translateY(0);opacity:1} }';
    document.head.appendChild(st);
  }

  document.body.appendChild(overlay);

  const close = () => overlay.remove();
  overlay.querySelector('[data-help-close]').addEventListener('click', close);
  overlay.addEventListener('click', e => { if (e.target === overlay) close(); });
  document.addEventListener('keydown', function esc(e) {
    if (e.key === 'Escape') { close(); document.removeEventListener('keydown', esc); }
  });
}

function _showAdditionalInfoHelp() {
  _showHelpModal({
    id:    'addInfoHelpModal',
    icon:  '💡',
    title: 'Why Additional Information Matters',
    intro: 'This section captures important context that supports the DACUM chart:',
    items: [
      ['📚', 'Skills &amp; Knowledge', 'Core competencies needed across multiple duties'],
      ['🧭', 'Behaviors',              'Work attitudes and traits that lead to success'],
      ['🛠️', 'Tools &amp; Equipment',  'Resources used to perform the occupation'],
      ['📈', 'Future Trends',          'Emerging factors that may impact the occupation'],
    ],
  });
}

function _showClusteringHelp() {
  _showHelpModal({
    id:    'clusteringHelpModal',
    icon:  '🎯',
    title: 'Create Competence Clusters',
    intro: 'Group related tasks into competence clusters based on:',
    items: [
      ['🎯', '', 'Common purpose or industry objective'],
      ['🔄', '', 'Similar workflow or process'],
      ['🧠', '', 'Shared knowledge and skills required'],
    ],
    note: '<strong>Important:</strong> Tasks from different duties can be grouped ' +
          'together if they are related by purpose, process, or required skills!',
  });
}

function _showClusterNamingHelp() {
  _showHelpModal({
    id:    'clusterNamingHelpModal',
    icon:  '✍️',
    title: 'Competence Statement Format',
    intro: 'Clusters are created as "Cluster 1", "Cluster 2"… Rename each one to a ' +
           'competence statement so the cluster says what the worker can DO, not just ' +
           'which group it is.',
    items: [
      ['🧩', 'Structure',
       'Action Verb + Task/Activity (What) + Context (Where/How/Why if relevant)'],
      ['💬', 'Example',
       '"Calibrate testing equipment according to manufacturer specifications"'],
    ],
    note: 'Good competence statements are clear, measurable, and include the context ' +
          'or standards that guide performance. Keep the standard (<em>"according to…", ' +
          '"within ±0.5 mm"</em>) and leave out the purpose (<em>"to ensure…"</em>) — ' +
          'the standard is what can be assessed.',
  });
}

// Performance Criteria + Range guidance.
// One modal, opened from BOTH section headings inside every cluster
// card — the two concepts are defined in relation to each other
// (criteria = the standard, range = the conditions the standard
// applies across), so splitting them into two modals would force the
// facilitator to open both to understand either.
function _showPCRangeHelp() {
  const P  = 'margin:0 0 8px;font-size:0.88em;color:#475569;line-height:1.6;';
  const LI = 'font-size:0.87em;line-height:1.7;color:#334155;';

  const component = (label, colour, desc) =>
    '<div style="flex:1 1 150px;min-width:0;background:#fff;border:1px solid #e2e8f0;' +
    'border-radius:8px;padding:10px 12px;">' +
      '<p style="margin:0 0 3px;font-size:0.84em;font-weight:800;color:' + colour + ';">' +
        label + '</p>' +
      '<p style="margin:0;font-size:0.79em;color:#64748b;line-height:1.5;">' + desc + '</p>' +
    '</div>';

  const chip = (text, bg, fg) =>
    '<span style="display:inline-block;background:' + bg + ';color:' + fg + ';' +
    'border-radius:5px;padding:1px 7px;font-size:0.88em;font-weight:700;' +
    'font-style:italic;white-space:nowrap;">(' + text + ')</span>';

  const bodyHtml =
    '<p style="' + P + '"><strong style="color:#1e293b;">Performance Criteria</strong> ' +
      'define the standards to which a competency must be performed. They should be:</p>' +
    '<ol style="margin:0 0 16px;padding-left:20px;">' +
      '<li style="' + LI + '"><strong>Observable</strong> — can be seen or detected during assessment</li>' +
      '<li style="' + LI + '"><strong>Measurable</strong> — can be evaluated against a standard</li>' +
      '<li style="' + LI + '"><strong>Learner-focused</strong> — describes what the learner must demonstrate</li>' +
    '</ol>' +

    '<p style="' + P + '"><strong style="color:#1e293b;">Range</strong> defines the scope ' +
      'and context in which the competency is applied, including:</p>' +
    '<ul style="margin:0 0 16px;padding-left:20px;">' +
      '<li style="' + LI + '">Different situations, environments, or conditions</li>' +
      '<li style="' + LI + '">Types of equipment, tools, or materials used</li>' +
      '<li style="' + LI + '">Variable contexts that may affect performance</li>' +
    '</ul>' +

    '<p style="margin:0 0 10px;font-size:0.88em;font-weight:700;color:#1e293b;">' +
      'Components of an effective performance criterion:</p>' +
    '<div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:16px;">' +
      component('What (Object)',  '#2563eb', 'The thing being acted upon') +
      component('Action (Verb)',  '#16a34a', 'The precise action being performed') +
      component('Qualifier',      '#9333ea', 'Specific conditions, standards, or requirements') +
    '</div>' +

    '<div style="background:linear-gradient(135deg,#3b82f6,#2563eb);border-radius:10px;' +
    'padding:14px 16px;margin-bottom:4px;">' +
      '<p style="margin:0 0 8px;font-size:0.86em;font-weight:800;color:#fff;">Example Structure:</p>' +
      '<p style="margin:0;font-size:0.87em;color:#fff;line-height:2;">' +
        '"Equipment calibration ' + chip('What', '#bfdbfe', '#1e3a8a') +
        ' is verified ' + chip('Action', '#bbf7d0', '#14532d') +
        ' to be within manufacturer\'s tolerance ranges ' + chip('Qualifier', '#e9d5ff', '#581c87') +
        '"' +
      '</p>' +
    '</div>';

  _showHelpModal({
    id:       'pcRangeHelpModal',
    icon:     '📐',
    title:    'What are Performance Criteria and Range?',
    maxWidth: '620px',
    bodyHtml,
  });
}
