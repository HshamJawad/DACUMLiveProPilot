// ============================================================
// /task_charts.js
// Task Verification analytics — Bar + Scatter charts in a modal.
//
// Public API:
//   initTaskCharts()           — wire modal, tabs, filters, export
//   openTaskChartsModal()      — open modal + render default chart
//   refreshChartButtonState()  — enable/disable the View Charts button
//                                based on whether any task has been rated
//
// Data sources (read-only):
//   - Individual mode: appState.verificationRatings[taskKey]
//                       { importance, frequency, difficulty, criticality, ... }
//   - Workshop mode:   appState.workshopCounts[taskKey]
//                       { importanceCounts:{0,1,2,3}, frequencyCounts:..., ... }
//
// Workshop counts are reduced to a weighted average per dimension:
//     avg = Σ(value × count) / Σ(count)
//
// Library: Chart.js v4 loaded LAZILY on first modal open via CDN.
// No global Chart import; we reference window.Chart after load.
// ============================================================

import { appState }      from './state.js';
import { getDutyLetter } from './codes.js';

const CDN_URL =
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js';

// Active Chart.js instance (destroyed before re-render to avoid leaks)
let _chart       = null;
let _chartLoadP  = null;       // promise that resolves when Chart.js is ready
let _currentType = 'bar';      // 'bar' | 'scatter'
let _wired       = false;

// ── Public API ──────────────────────────────────────────────

/** One-time wiring of modal controls.  Safe to call multiple times. */
export function initTaskCharts() {
  if (_wired) return;
  _wired = true;

  const modal = document.getElementById('taskChartsModal');
  if (!modal) return;

  // Close handlers
  modal.querySelectorAll('[data-tc-close]').forEach(el =>
    el.addEventListener('click', closeTaskChartsModal)
  );
  // Click on backdrop (the modal element itself, not its children) closes
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeTaskChartsModal();
  });
  // ESC closes
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modal.classList.contains('is-open')) {
      closeTaskChartsModal();
    }
  });

  // Chart-type tabs
  modal.querySelectorAll('[data-tc-type]').forEach(btn =>
    btn.addEventListener('click', () => {
      _currentType = btn.getAttribute('data-tc-type');
      _syncTypeTabs();
      _render();
    })
  );

  // Filters
  const dutyFilter = document.getElementById('tcDutyFilter');
  const sortFilter = document.getElementById('tcSortFilter');
  if (dutyFilter) dutyFilter.addEventListener('change', _render);
  if (sortFilter) sortFilter.addEventListener('change', _render);

  // PNG export
  const exportBtn = document.getElementById('btnTCExportPNG');
  if (exportBtn) exportBtn.addEventListener('click', _exportPNG);

  // Refresh button state whenever ratings change anywhere in the app.
  // tasks.js dispatches this debounced event after updateRating /
  // updateWorkshopCount.  If the modal is open, we also re-render so
  // the user sees live updates while interacting in another window/tab.
  document.addEventListener('dacum:ratings-changed', () => {
    refreshChartButtonState();
    const modal = document.getElementById('taskChartsModal');
    if (modal && modal.classList.contains('is-open') && window.Chart) {
      _render();
    }
  });

  // Initial button state (in case data is already present at boot)
  refreshChartButtonState();
}

/**
 * Open the modal and render the current chart type with current data.
 * Lazy-loads Chart.js if not already loaded.
 */
export function openTaskChartsModal() {
  const modal = document.getElementById('taskChartsModal');
  if (!modal) return;

  // Refresh duty filter options every open (duties may have changed)
  _populateDutyFilter();
  _syncTypeTabs();

  modal.classList.add('is-open');
  document.body.style.overflow = 'hidden';

  // Lazy-load Chart.js, then render
  _ensureChartJs()
    .then(() => _render())
    .catch((err) => {
      console.error('Chart.js failed to load:', err);
      _showError(
        'Chart library failed to load. Check your internet connection and try again.'
      );
    });
}

/** Close the modal and free the chart instance. */
export function closeTaskChartsModal() {
  const modal = document.getElementById('taskChartsModal');
  if (!modal) return;
  modal.classList.remove('is-open');
  document.body.style.overflow = '';
  if (_chart) { try { _chart.destroy(); } catch (_) {} _chart = null; }
}

/**
 * Update the disabled state of the View Charts button.  Called after any
 * rating change, on tab switches, after Live Workshop fetches, etc.
 */
export function refreshChartButtonState() {
  const btn = document.getElementById('btnViewTaskCharts');
  if (!btn) return;
  const hasData = _hasAnyRatedTask();
  btn.disabled = !hasData;
  btn.title = hasData
    ? 'Open task analysis charts'
    : 'Rate at least one task to enable charts';
}

// ── Lazy CDN loader ─────────────────────────────────────────

function _ensureChartJs() {
  if (window.Chart) return Promise.resolve(window.Chart);
  if (_chartLoadP) return _chartLoadP;
  _chartLoadP = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = CDN_URL;
    s.async = true;
    s.onload  = () => resolve(window.Chart);
    s.onerror = () => reject(new Error('Network error loading Chart.js'));
    document.head.appendChild(s);
  });
  return _chartLoadP;
}

// ── Data extraction ─────────────────────────────────────────

/**
 * Build a normalized list of rated tasks.  Each entry:
 *   { taskId, taskCode, taskText, dutyId, dutyLetter, dutyTitle,
 *     importance, frequency, difficulty, score }
 * Tasks without ANY rating data are skipped.
 *
 * Score is the simple sum (I + F + D) — what the existing code uses.
 */
function _readChartData() {
  const isWorkshop = appState.collectionMode === 'workshop';
  const duties     = appState.dutiesData || [];

  // Build duty lookup: dutyId → { index, letter, title }
  const dutyInfo = {};
  duties.forEach((d, i) => {
    dutyInfo[d.id] = { index: i, letter: getDutyLetter(i), title: d.title || '' };
  });

  const out = [];

  duties.forEach((duty) => {
    const di = dutyInfo[duty.id];
    (duty.tasks || []).forEach((task, taskIndex) => {
      const taskId = task.id || task.inputId;
      if (!taskId) return;

      const i = _readDimension(taskId, 'importance', isWorkshop);
      const f = _readDimension(taskId, 'frequency',  isWorkshop);
      const d = _readDimension(taskId, 'difficulty', isWorkshop);

      // Skip tasks where ALL three dimensions are unrated
      if (i === null && f === null && d === null) return;

      // Treat null as 0 for charting purposes (so a partial rating still appears)
      const I = i === null ? 0 : i;
      const F = f === null ? 0 : f;
      const D = d === null ? 0 : d;

      out.push({
        taskId,
        taskCode: `Task ${di.letter}${taskIndex + 1}`,
        taskText: task.text || '',
        dutyId:   duty.id,
        dutyLetter: di.letter,
        dutyTitle:  di.title,
        importance: I,
        frequency:  F,
        difficulty: D,
        score:      I + F + D,
        partial:    (i === null) || (f === null) || (d === null),
      });
    });
  });

  return out;
}

/**
 * Read one dimension's value for a task.  Returns:
 *   - workshop mode → weighted average across counts (or null if no votes)
 *   - individual mode → the explicit rating (or null if not rated)
 */
function _readDimension(taskKey, dim, isWorkshop) {
  if (isWorkshop) {
    const counts = appState.workshopCounts?.[taskKey];
    if (!counts) return null;
    const c = counts[`${dim}Counts`];
    if (!c) return null;
    let total = 0, weighted = 0;
    for (const v of [0, 1, 2, 3]) {
      const n = c[v] || 0;
      total    += n;
      weighted += v * n;
    }
    return total > 0 ? +(weighted / total).toFixed(2) : null;
  }
  const r = appState.verificationRatings?.[taskKey];
  if (!r) return null;
  return (r[dim] === null || r[dim] === undefined) ? null : Number(r[dim]);
}

/** True if any task in the project has at least one dimension rated. */
function _hasAnyRatedTask() {
  return _readChartData().length > 0;
}

// ── UI helpers ──────────────────────────────────────────────

function _syncTypeTabs() {
  document.querySelectorAll('[data-tc-type]').forEach(btn => {
    btn.classList.toggle('is-active',
      btn.getAttribute('data-tc-type') === _currentType);
  });
}

function _populateDutyFilter() {
  const sel = document.getElementById('tcDutyFilter');
  if (!sel) return;
  const prev = sel.value;
  const duties = appState.dutiesData || [];
  sel.innerHTML = '<option value="__all__">All duties</option>' +
    duties.map((d, i) => {
      const letter = getDutyLetter(i);
      const title  = (d.title || '').replace(/</g, '&lt;');
      return `<option value="${d.id}">Duty ${letter} — ${title || 'Untitled'}</option>`;
    }).join('');
  // Restore selection if still valid
  if (prev && [...sel.options].some(o => o.value === prev)) {
    sel.value = prev;
  } else {
    sel.value = '__all__';
  }
}

function _showError(msg) {
  const wrap = document.getElementById('tcChartWrap');
  const empty = document.getElementById('tcEmptyState');
  if (wrap)  wrap.style.display  = 'none';
  if (empty) {
    empty.style.display = 'block';
    empty.innerHTML = `<div class="tc-empty-icon">⚠️</div><h3>Error</h3><p>${msg}</p>`;
  }
}

function _showEmpty(msg) {
  const wrap = document.getElementById('tcChartWrap');
  const empty = document.getElementById('tcEmptyState');
  if (wrap)  wrap.style.display  = 'none';
  if (empty) {
    empty.style.display = 'block';
    empty.innerHTML = `<div class="tc-empty-icon">📊</div><h3>No data yet</h3><p>${msg}</p>`;
  }
}

function _showChart() {
  const wrap = document.getElementById('tcChartWrap');
  const empty = document.getElementById('tcEmptyState');
  if (wrap)  wrap.style.display  = 'block';
  if (empty) empty.style.display = 'none';
}

function _updateInfoBar(visibleCount, totalRated) {
  const el = document.getElementById('tcInfoBar');
  if (!el) return;
  if (visibleCount === totalRated) {
    el.textContent = `Showing all ${totalRated} rated task${totalRated === 1 ? '' : 's'}.`;
  } else {
    el.textContent = `Showing ${visibleCount} of ${totalRated} rated tasks.`;
  }
}

// ── Filtering & sorting ─────────────────────────────────────

function _applyFilters(rows) {
  const dutyFilter = document.getElementById('tcDutyFilter')?.value || '__all__';
  const sort       = document.getElementById('tcSortFilter')?.value || 'order';

  let filtered = (dutyFilter === '__all__')
    ? rows.slice()
    : rows.filter(r => r.dutyId === dutyFilter);

  switch (sort) {
    case 'score':       filtered.sort((a, b) => b.score - a.score); break;
    case 'importance':  filtered.sort((a, b) => b.importance - a.importance); break;
    case 'frequency':   filtered.sort((a, b) => b.frequency  - a.frequency);  break;
    case 'difficulty':  filtered.sort((a, b) => b.difficulty - a.difficulty); break;
    case 'order':
    default:            /* keep insertion order */ break;
  }
  return filtered;
}

// ── Rendering ───────────────────────────────────────────────

function _render() {
  if (!window.Chart) return;   // not loaded yet
  if (_chart) { try { _chart.destroy(); } catch (_) {} _chart = null; }

  const rows     = _readChartData();
  const filtered = _applyFilters(rows);

  if (rows.length === 0) {
    _showEmpty('Rate at least one task in the Task Verification tab to see charts.');
    _updateInfoBar(0, 0);
    return;
  }
  if (filtered.length === 0) {
    _showEmpty('No rated tasks match the current filters. Try selecting "All duties".');
    _updateInfoBar(0, rows.length);
    return;
  }

  _showChart();
  _updateInfoBar(filtered.length, rows.length);

  if (_currentType === 'bar')          _renderBar(filtered);
  else if (_currentType === 'scatter') _renderScatter(filtered);
}

function _renderBar(rows) {
  const canvas = document.getElementById('tcCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const labels = rows.map(r => r.taskCode);

  _chart = new window.Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Importance',
          data:  rows.map(r => r.importance),
          backgroundColor: '#3b82f6',
          borderRadius: 3,
          maxBarThickness: 22,
        },
        {
          label: 'Frequency',
          data:  rows.map(r => r.frequency),
          backgroundColor: '#10b981',
          borderRadius: 3,
          maxBarThickness: 22,
        },
        {
          label: 'Difficulty',
          data:  rows.map(r => r.difficulty),
          backgroundColor: '#f59e0b',
          borderRadius: 3,
          maxBarThickness: 22,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        tooltip: {
          callbacks: {
            afterTitle: (items) => {
              if (!items || !items[0]) return '';
              const row = rows[items[0].dataIndex];
              return row ? row.taskText.slice(0, 80) + (row.taskText.length > 80 ? '…' : '') : '';
            },
          },
        },
        title: { display: false },
      },
      scales: {
        x: {
          ticks: { autoSkip: false, maxRotation: 70, minRotation: 0, font: { size: 10 } },
          grid:  { display: false },
        },
        y: {
          beginAtZero: true,
          max: 3,
          ticks: { stepSize: 0.5 },
          title: { display: true, text: 'Rating (0 – 3)' },
        },
      },
    },
  });
}

function _renderScatter(rows) {
  const canvas = document.getElementById('tcCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  // Each task = one bubble: x=importance, y=difficulty, r=size from frequency
  // Color encodes difficulty band for quick visual sorting.
  const data = rows.map(r => ({
    x: r.importance,
    y: r.difficulty,
    r: 4 + (r.frequency * 3.5),    // 4 – 14.5 px
    _row: r,
  }));

  _chart = new window.Chart(ctx, {
    type: 'bubble',
    data: {
      datasets: [{
        label: 'Tasks',
        data,
        backgroundColor: data.map(d => _bubbleColor(d._row.difficulty, 0.65)),
        borderColor:     data.map(d => _bubbleColor(d._row.difficulty, 1)),
        borderWidth: 1.5,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx2) => {
              const r = ctx2.raw && ctx2.raw._row;
              if (!r) return '';
              return [
                r.taskCode,
                `${r.taskText.slice(0, 80)}${r.taskText.length > 80 ? '…' : ''}`,
                `Importance: ${r.importance}`,
                `Difficulty: ${r.difficulty}`,
                `Frequency:  ${r.frequency}`,
              ];
            },
          },
        },
        title: { display: false },
      },
      scales: {
        x: {
          min: -0.2, max: 3.2,
          ticks: { stepSize: 0.5 },
          title: { display: true, text: 'Importance →' },
          grid:  { color: (ctx2) => ctx2.tick.value === 1.5 ? '#94a3b8' : '#e5e7eb',
                   lineWidth: (ctx2) => ctx2.tick.value === 1.5 ? 1.5 : 1 },
        },
        y: {
          min: -0.2, max: 3.2,
          ticks: { stepSize: 0.5 },
          title: { display: true, text: 'Difficulty →' },
          grid:  { color: (ctx2) => ctx2.tick.value === 1.5 ? '#94a3b8' : '#e5e7eb',
                   lineWidth: (ctx2) => ctx2.tick.value === 1.5 ? 1.5 : 1 },
        },
      },
    },
  });
}

function _bubbleColor(difficulty, alpha) {
  // Difficulty 0–3 → green → yellow → orange → red gradient
  const palette = [
    [16, 185, 129],   // green-500
    [234, 179, 8],    // yellow-500
    [249, 115, 22],   // orange-500
    [239, 68, 68],    // red-500
  ];
  const idx = Math.max(0, Math.min(3, Math.round(difficulty)));
  const [r, g, b] = palette[idx];
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// ── PNG export ──────────────────────────────────────────────

function _exportPNG() {
  if (!_chart) return;
  try {
    const url = _chart.toBase64Image('image/png', 1);
    const a = document.createElement('a');
    const occ = (document.getElementById('occupationTitle')?.value || 'DACUM').trim()
      .replace(/[^a-z0-9]/gi, '_').slice(0, 40) || 'DACUM';
    const date = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `${occ}_TaskAnalysis_${_currentType}_${date}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (err) {
    console.error('PNG export failed:', err);
    alert('Export failed. Please try again.');
  }
}
