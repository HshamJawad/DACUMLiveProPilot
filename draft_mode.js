// ============================================================
//  draft_mode.js — "a pipeline is running" flag
//  DACUM Live Pro
// ------------------------------------------------------------
//  WHY THIS IS ITS OWN FILE
//
//  Every AI module needs to know whether it is being run on its own
//  or as one stage of a Full Draft, because the answer changes two
//  behaviours: whether it may open a confirm() dialog, and whether
//  it shows its own loading modal.
//
//  draft_agent.js already imports all of them, so putting the flag
//  there would make each module import its own importer — a cycle.
//  ES modules tolerate cycles, but the resolution order becomes
//  load-order dependent and breaks in ways that are miserable to
//  debug. A leaf module with no imports of its own cannot cycle.
//
//  It is deliberately NOT on appState: that object is serialised
//  into project files, and a stuck `batchMode: true` saved into a
//  .json would silently suppress every confirmation dialog for
//  whoever opened it next.
// ============================================================

let _batch = false;

/** True while a Full Draft run is in flight. */
export function isBatchRun() { return _batch; }

export function setBatchRun(on) { _batch = !!on; }
