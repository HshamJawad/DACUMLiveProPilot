// ============================================================
// /state.js
// Single source of truth for all mutable application state.
// Import appState in any module that needs to read or write state.
// ============================================================

export const appState = {

  // ── Chart Info Images ──────────────────────────────────────
  producedForImage: null,
  producedByImage: null,

  // ── Duties & Tasks ─────────────────────────────────────────
  dutyCount: 0,
  taskCounts: {},          // { dutyId: number }

  // ── Additional Info ────────────────────────────────────────
  customSectionCounter: 0,

  // ── Skills Level Matrix ────────────────────────────────────
  // Seeded from i18n at first render, not here: see
  // defaultSkillsLevelData() at the foot of this file.
  skillsLevelData: [],

  // ── Task Verification ──────────────────────────────────────
  verificationRatings: {},   // { taskKey: { importance, frequency, difficulty, ... } }
  taskMetadata: {},           // { taskKey: { dutyId, dutyTitle } }
  collectionMode: 'workshop',
  workflowMode: 'standard',
  verificationDecisionMade: false,
  clusteringAllowed: false,

  // ── Workshop Aggregated Counts ─────────────────────────────
  workshopParticipants: 10,
  priorityFormula: 'if',
  workshopCounts: {},
  workshopResults: {},

  // ── Export Modes ───────────────────────────────────────────
  tvExportMode: 'appendix',
  trainingLoadMethod: 'advanced',

  // ── Live Workshop ──────────────────────────────────────────
  lwSessionId: null,
  lwIsFinalized: false,
  lwFinalizedData: null,
  lwAggregatedResults: null,

  // ── Competency Clustering ──────────────────────────────────
  clusteringData: {
    availableTasks: [],
    clusters: [],
    clusterCounter: 0
  },

  // ── Learning Outcomes ──────────────────────────────────────
  learningOutcomesData: {
    outcomes: [],
    outcomeCounter: 0
  },

  // ── Module Mapping ─────────────────────────────────────────
  moduleMappingData: {
    modules: [],
    moduleCounter: 0
  }
};

/* ── Skills Level Matrix defaults ─────────────────────────────
   The seed used to be written out twice as English literals: once in
   this file and once again, verbatim, inside resetSkillsLevel() in
   renderer.js. Two copies of the same 33 strings is one copy too many
   — the second had already drifted in whitespace — so both are now
   generated from this single spec.

   Only the STRUCTURE lives here. The wording is resolved from i18n at
   call time, which is what makes the matrix appear in Arabic or French
   instead of English.

   Category 9 is deliberately empty: it is the blank row a facilitator
   fills in themselves, so it has no key and stays empty in every
   language. */
const SKILLS_SEED = [
  { id: 1, key: 'slCat1', comps: ['slComp1_1', 'slComp1_2'] },
  { id: 2, key: 'slCat2', comps: ['slComp2_1', 'slComp2_2', 'slComp2_3', 'slComp2_4', 'slComp2_5'] },
  { id: 3, key: 'slCat3', comps: ['slComp3_1', 'slComp3_2', 'slComp3_3'] },
  { id: 4, key: 'slCat4', comps: ['slComp4_1', 'slComp4_2', 'slComp4_3'] },
  { id: 5, key: 'slCat5', comps: ['slComp5_1', 'slComp5_2', 'slComp5_3', 'slComp5_4', 'slComp5_5'] },
  { id: 6, key: 'slCat6', comps: ['slComp6_1', 'slComp6_2'] },
  { id: 7, key: 'slCat7', comps: ['slComp7_1', 'slComp7_2'] },
  { id: 8, key: 'slCat8', comps: ['slComp8_1', 'slComp8_2'] },
  { id: 9, key: null,     comps: [null, null] },
];

const _t = (k) => (k && window.i18n ? window.i18n.t(k) : '');

const _levels = () => ({ craftsman: false, skilled: false, semiSkilled: false, foundation: false });

/**
 * Build a fresh default matrix IN THE CURRENT INTERFACE LANGUAGE.
 *
 * Called when a new matrix is generated: first render of an empty
 * project, or an explicit Reset. It is deliberately NOT called on
 * language switch — once these strings are in appState they are data,
 * and the rows are user-editable. Re-translating them later would
 * overwrite wording a facilitator had adjusted for their own sector,
 * which is a worse failure than an English row in an Arabic chart.
 *
 * The old version of this function deep-cloned appState.skillsLevelData
 * — the LIVE array — so calling it after any edit returned the edited
 * data, not the defaults. It had no callers, which is the only reason
 * that never surfaced as a bug.
 */
export function defaultSkillsLevelData() {
  return SKILLS_SEED.map(cat => ({
    id: cat.id,
    category: _t(cat.key),
    competencies: cat.comps.map((compKey, i) => ({
      id: `${cat.id}.${i + 1}`,
      text: _t(compKey),
      levels: _levels()
    }))
  }));
}

/** True when the matrix has never been populated (fresh project). */
export function skillsLevelIsEmpty() {
  return !Array.isArray(appState.skillsLevelData) || appState.skillsLevelData.length === 0;
}
