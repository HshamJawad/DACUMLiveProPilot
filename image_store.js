// ============================================================
// /image_store.js
// IndexedDB-backed storage for chart logos.
//
// WHY:
//   Logos were stored as base64 data URLs inside each project's state
//   in `dacum_projects` (localStorage). base64 inflates binary by ~33%,
//   localStorage tops out around 5 MB for the whole origin, and every
//   autosave re-serialised the images along with everything else. Moving
//   them to IndexedDB gives a quota measured in hundreds of megabytes
//   and takes the largest payload out of the hot save path entirely.
//
// THE DESIGN PROBLEM, AND HOW IT IS SOLVED:
//   IndexedDB is asynchronous. The project system is not:
//   _captureState() and _applyState() in dacum_projects.js are ordinary
//   synchronous functions called from many places, and exports.js reads
//   appState.producedForImage synchronously while building a PDF.
//   Rewriting all of that to be async would be a large, risky change.
//
//   So IndexedDB is used as the PERSISTENCE layer only, with an
//   in-memory Map as the ACCESS layer:
//     • init()      — called once at boot, loads every stored image
//                     into memory (a handful of small entries).
//     • getSync()   — synchronous reads, straight from memory.
//     • set()       — updates memory immediately, writes to IndexedDB
//                     in the background.
//   Callers keep their synchronous shape; only the durable write is
//   deferred. If IndexedDB is unavailable (private browsing, old
//   browser, blocked storage), everything still works from memory for
//   the session and falls back to inline storage on save.
//
// STORAGE FORMAT:
//   A project's state holds a reference string, `idb:<key>`, instead of
//   the image itself. Legacy projects still holding a raw `data:` URL
//   keep working and are migrated on first load.
// ============================================================

const DB_NAME    = 'dacum_images';
const DB_VERSION = 1;
const STORE      = 'logos';

export const IMAGE_REF_PREFIX = 'idb:';

let _db        = null;
let _available = false;
const _cache   = new Map();   // key -> data URL

// ── Low-level IndexedDB ───────────────────────────────────────

function _openDB() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('IndexedDB unavailable'));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error || new Error('IndexedDB open failed'));
    req.onblocked = () => reject(new Error('IndexedDB blocked'));
  });
}

function _tx(mode) {
  if (!_db) throw new Error('IndexedDB not ready');
  return _db.transaction(STORE, mode).objectStore(STORE);
}

// ── Public API ────────────────────────────────────────────────

/**
 * Open the database and pull every stored image into memory.
 * Safe to call more than once; resolves even when IndexedDB is
 * unavailable, in which case the store degrades to memory-only.
 */
export async function initImageStore() {
  if (_db) return true;
  try {
    _db = await _openDB();
    _available = true;
  } catch (err) {
    console.warn('[image_store] IndexedDB unavailable, using memory only:', err.message);
    _available = false;
    return false;
  }

  // Warm the cache so every later read can be synchronous.
  try {
    await new Promise((resolve) => {
      const store = _tx('readonly');
      const req   = store.openCursor();
      req.onsuccess = () => {
        const cursor = req.result;
        if (!cursor) { resolve(); return; }
        if (typeof cursor.value === 'string') _cache.set(cursor.key, cursor.value);
        cursor.continue();
      };
      req.onerror = () => resolve();
    });
  } catch (err) {
    console.warn('[image_store] cache warm failed:', err);
  }

  return true;
}

/** True when a value is a reference rather than an inline image. */
export function isImageRef(value) {
  return typeof value === 'string' && value.startsWith(IMAGE_REF_PREFIX);
}

/** Build the storage key for one of a project's two logo slots. */
export function imageKey(projectId, slot) {
  return `${projectId || 'unassigned'}::${slot}`;
}

/**
 * Read an image synchronously.
 * Accepts either a bare key or a full `idb:` reference, and passes a
 * raw data URL straight through so callers never need to know which
 * form a given project uses.
 */
export function getImageSync(refOrKey) {
  if (!refOrKey) return null;
  if (typeof refOrKey !== 'string') return null;
  if (refOrKey.startsWith('data:')) return refOrKey;      // legacy inline
  const key = isImageRef(refOrKey) ? refOrKey.slice(IMAGE_REF_PREFIX.length) : refOrKey;
  return _cache.get(key) || null;
}

/**
 * Store an image and return the reference to persist in project state.
 * The memory cache is updated immediately so a read on the very next
 * line sees the new value; the durable write happens in the background.
 */
export function setImage(key, dataUrl) {
  if (!key || !dataUrl) return null;
  _cache.set(key, dataUrl);

  if (_available) {
    try {
      _tx('readwrite').put(dataUrl, key);
    } catch (err) {
      console.warn('[image_store] write failed:', err);
    }
  }
  return IMAGE_REF_PREFIX + key;
}

/** Remove an image from both memory and IndexedDB. */
export function removeImage(refOrKey) {
  if (!refOrKey) return;
  const key = isImageRef(refOrKey) ? refOrKey.slice(IMAGE_REF_PREFIX.length) : refOrKey;
  _cache.delete(key);
  if (_available) {
    try { _tx('readwrite').delete(key); }
    catch (err) { console.warn('[image_store] delete failed:', err); }
  }
}

/**
 * Delete every image belonging to a project — called when the project
 * itself is deleted, so images cannot outlive their owner and quietly
 * accumulate in IndexedDB forever.
 */
export function removeProjectImages(projectId) {
  if (!projectId) return;
  ['producedFor', 'producedBy'].forEach(slot => removeImage(imageKey(projectId, slot)));
}

/** True when durable storage is actually working. */
export function imageStoreAvailable() {
  return _available;
}

/** Approximate bytes held in the image store (for diagnostics). */
export function imageStoreSize() {
  let total = 0;
  _cache.forEach(v => { total += v.length; });
  return total;
}
