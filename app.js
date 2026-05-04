// ============================================================
// /app.js
// Application entry point — wires everything together on DOMContentLoaded.
// ============================================================

import { appState }          from './state.js';
import { renderSkillsLevel } from './renderer.js';
import { updateUsageBadge }  from './storage.js';
import { setupTabs }         from './tabs.js';
import { setupEvents }       from './events.js';
import { switchTab }         from './projects.js';
import { addDuty, addTask }  from './duties.js';
import { updateCollectionMode, updateWorkflowMode, updateDutyLevelSummary } from './tasks.js';
import { lwCheckAndShowSection } from './workshop.js';
import { setBaseline }       from './history.js';
import { renderSnapshotPanel } from './workshop_snapshots.js';
import { initProjectsSidebar, saveCurrentProject,
         createProject, getActiveProjectId } from './dacum_projects.js';
import { startAutoSave, checkCrashRecovery } from './autosave.js';
import { clearAiGeneratedFlag } from './refine.js';
import { initDragDrop }        from './drag_drop.js';

// Expose switchTab globally (called from HTML onclick and live workshop guards)
window.switchTab = switchTab;
window.updateDutyLevelSummary = updateDutyLevelSummary;

document.addEventListener('DOMContentLoaded', function () {
  // Initialize Skills Level Matrix
  renderSkillsLevel();

  // Ensure Refine Results button is hidden until AI runs
  clearAiGeneratedFlag();

  // Initialize usage badge
  updateUsageBadge();

  // Wire tabs
  setupTabs();

  // Wire all event listeners
  setupEvents();

  // Add initial duty + task if duties container is empty
  const dutiesContainer = document.getElementById('dutiesContainer');
  if (dutiesContainer && dutiesContainer.children.length === 0) {
    addDuty();
    addTask(`duty_${appState.dutyCount}`);
  }

  // Anchor the history baseline
  setBaseline();

  // Render saved snapshots panel
  renderSnapshotPanel();

  // Initialize multi-project sidebar
  initProjectsSidebar();

  // If no active project yet, create one automatically from the initial state
  if (!getActiveProjectId()) {
    const occ = document.getElementById('occupationTitle')?.value?.trim();
    createProject(occ || 'My First DACUM Project');
  }

  // Auto-save active project when user leaves the page
  window.addEventListener('beforeunload', () => saveCurrentProject());

  // Start auto-save observer (debounced, 800 ms)
  startAutoSave();

  // Initialize drag & drop for task cards (Card View only)
  initDragDrop();

  // Check for unsaved work from a previous crashed session
  checkCrashRecovery();

  // Initialize Task Verification controls
  updateCollectionMode();
  updateWorkflowMode();

  // Check Live Workshop section visibility
  const urlParams = new URLSearchParams(window.location.search);
  const sessionParam = urlParams.get('lwsession');
  if (sessionParam) {
    // Participant mode – redirect
    const currentPath = window.location.pathname;
    const directory   = currentPath.substring(0, currentPath.lastIndexOf('/') + 1);
    const participantFileUrl = window.location.origin + directory + 'DACUM_LiveWorkshop_Participant.html';
    window.location.href = `${participantFileUrl}?lwsession=${sessionParam}`;
  } else {
    setTimeout(lwCheckAndShowSection, 100);
  }
});
