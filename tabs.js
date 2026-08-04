// ============================================================
// /tabs.js
// Tab click setup (non-switchTab tabs, i.e. all except the
// ones gated by clusteringAllowed which go via switchTab).
// ============================================================

import { appState } from './state.js';
import { addDuty } from './duties.js';
import { initializeClusteringFromTasks } from './modules.js';
import { renderPCSourceList, renderLearningOutcomes,
  renderModuleLoList, renderModules } from './modules.js';

export function setupTabs() {
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', function () {
      const tabId = this.getAttribute('data-tab');

      // clustering-tab is guarded by switchTab() – do not wire it here directly
      // (The HTML button for clustering tab should call window.switchTab)
      if (tabId === 'clustering-tab') return;

      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

      this.classList.add('active');
      const content = document.getElementById(tabId);
      if (content) content.classList.add('active');

      // Side-effects per tab
      if (tabId === 'duties-tab') {
        const dutiesContainer = document.getElementById('dutiesContainer');
        if (dutiesContainer && dutiesContainer.children.length === 0) {
          addDuty();   // seeds its own first task — see duties.js
        }
      }

      if (tabId === 'learning-outcomes-tab') {
        renderPCSourceList();
        renderLearningOutcomes();
      }

      if (tabId === 'module-mapping-tab') {
        renderModuleLoList();
        renderModules();
      }
    });
  });
}
