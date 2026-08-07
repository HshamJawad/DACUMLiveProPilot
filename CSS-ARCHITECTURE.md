# CSS Architecture — DACUM Live Pro

The rules below exist because this project repeatedly shipped bugs that
produced **no error message**: buttons rendered as vertical ovals on
Android, the toolbar stopped scrolling so "Clear All" became unreachable,
the rating scale split into two rows, and three mobile breakpoints for the
tab font size silently stopped applying. Every one of them was a cascade
conflict between two stylesheets that both claimed the same component.

None was catchable by looking at the page in a desktop browser.

---

## The load order

```html
<link rel="stylesheet" href="dacum-styles.css">      <!-- 1 -->
<link rel="stylesheet" href="dacum-responsive.css">  <!-- 2 -->
<link rel="stylesheet" href="dacum-typography.css">  <!-- 3 -->
<link rel="stylesheet" href="dacum-components.css">  <!-- 4 -->
```

This order is load-bearing and nothing in the code enforces it. Later
files win ties at equal specificity, so reordering changes which rule
applies with no error to warn you.

## What each file may contain

| File | May contain | Must never contain |
|---|---|---|
| `dacum-styles.css` | Base, layout, and the canonical definition of every component | — |
| `dacum-responsive.css` | `@media` blocks **only** | Any rule outside a media query |
| `dacum-typography.css` | `font-*`, `color`, `letter-spacing`, `line-height`, `text-*` | Padding, background, border-radius, box-shadow, layout |
| `dacum-components.css` | Overrides that genuinely need to load last | A component's primary definition |

## The one rule

**One component, one owner.** If two files both describe what something
looks like, the appearance depends on load order, and load order is
invisible at the point of editing.

A media-query override of a base rule is not split ownership — that is
the responsive layer doing its job. Neither is the typography layer
setting type and colour. The line is drawn at **shape and layout**.

## Before adding a rule

1. Search all four files for the selector first. If it already exists,
   extend the existing rule instead of writing a new one elsewhere.
2. Fix the **source**, not the symptom. If a base rule is wrong, change
   it. Adding a later rule with `!important` to cancel it produces the
   three-file argument that Phase 4 dismantled.
3. Pin every axis on fixed-shape elements. `dacum-responsive.css` applies
   `min-height: 44px` to every button on touch devices; declaring only
   `height` leaves the inherited 44px in force and turns a circle into a
   vertical oval. Declare `width`, `height`, `min-*` and `max-*` together.
4. Never use `:not()` chains to exempt elements from a global rule.
   `:not()` inherits its argument's specificity, so a chain of seven
   raises the rule above every class selector and inverts the intent —
   this is exactly how the oval-button bug came back after being fixed.

## Before every release

```bash
node tools/css-ownership-audit.mjs . --max 0
node tools/preflight.mjs .
```

Both run automatically in CI (`.github/workflows/preflight.yml`).

`css-ownership-audit` reports selectors defined in more than one file.
The count is **0**; keep it there.

`preflight` verifies that every referenced asset — including transitively
imported ES modules — is in `PRECACHE_URLS` in `sw.js`. This matters more
than it looks: `cache.addAll` rejects atomically, so a single missing
entry means the service worker never installs and **no update reaches
anyone who installed the PWA**. They keep running the old build and
report bugs that were already fixed.

## When adding or renaming any file

1. Add it to `PRECACHE_URLS` in `sw.js`.
2. If it is a stylesheet, add it to `criticalUrls` as well — `index.html`
   carries no inline `<style>`, so it depends entirely on its linked
   sheets and would paint unstyled for one fetch cycle without this.
3. Bump `CACHE_VERSION` in `sw.js` and `APP_VERSION` in `index.html`.

---

## History

| Phase | Change | Split ownership |
|---|---|---|
| — | Starting point | 22 |
| 1 | 1,860-line inline `<style>` moved out of `index.html` | 22 |
| 2 | `dacum-fixes.css` dissolved into the owning sheets | — |
| 3 | Non-media rules removed from `dacum-responsive.css` | 15 |
| 4 | Duplicate definitions merged at source | **0** |
| 5 | Audit, preflight, CI, this document | 0 |
