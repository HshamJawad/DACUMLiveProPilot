// ============================================================
// /src/module_model.js
// One plain-JS description of a module, in ONE language, with every
// "should this appear" rule already applied.
//
// ── WHY THIS FILE EXISTS ────────────────────────────────────
// exports_docx.js does two jobs at once: it decides WHAT belongs in the
// document, and it builds docx objects. The first job is not about docx
// at all. An outcome with no titled sheets is skipped, an untitled sheet
// is skipped, an assessment form nobody filled is skipped — those are
// facts about the module, and any exporter that disagrees with them
// produces a different document from the Word one and quietly makes the
// tool look unreliable.
//
// So the rules live here, once, and exporters render the result.
// exports_html.js is the first consumer. A future PPTX exporter is the
// second, and will need no new rules — only new rendering.
//
// ── WHAT THIS FILE DOES NOT DO ──────────────────────────────
// No DOM. No docx. No HTML. No storage. It reads window.mbState and
// returns data, so it can be tested by reading the return value rather
// than by unzipping a file.
//
// ── ON THE DUPLICATION, HONESTLY ────────────────────────────
// These predicates are MIRRORED from exports_docx.js, not extracted
// from it. Extracting them would have meant editing the Word exporter,
// and the Word exporter currently reproduces its previous output byte
// for byte — a guarantee worth more than avoiding a duplicated `if`.
//
// The cost is real: two places now encode the same rule, and they can
// drift. The mitigation is a test that builds a module both ways and
// asserts the two agree on which sections exist. If that test ever
// fails, they have drifted, and the fix is to unify them properly
// rather than to patch one side.
// ============================================================

/* ── Inclusion rules ────────────────────────────────────────
   Named, exported, and commented individually so that when a future
   exporter asks "does this outcome count?" there is exactly one answer
   to point at. */

/** An outcome earns a heading only if it has at least one TITLED sheet.
    Presence is not content: the tool creates an empty sheet the moment
    you open a tab, so counting sheets would give every outcome a page. */
function mmOutcomeHasContent(lo) {
    if (!lo) return false;
    return (lo.infoSheets || []).some(function (s) { return _mmStr(s && s.title).trim(); }) ||
           (lo.activitySheets || []).some(function (s) { return _mmStr(s && s.title).trim(); });
}

/** A sheet without a title is a sheet the author started and abandoned. */
function mmSheetHasContent(sheet) {
    return !!(sheet && _mmStr(sheet.title).trim());
}

/** Delegates to assessment.js when it is loaded, because that function
    is the definition of "filled" and this file should not invent a
    second one. The fallback is deliberately strict rather than
    permissive: exporting a blank grid is worse than omitting a form. */
function mmAssessmentFilled(form) {
    if (typeof mbAssessmentFormFilled === 'function') return mbAssessmentFormFilled(form);
    if (!form) return false;
    return (form.rows || []).some(function (r) {
        return r && ['criteria', 'activities', 'outcomes', 'verification', 'date']
            .some(function (k) { return _mmStr(r[k]).trim(); });
    });
}

/* ── Helpers ────────────────────────────────────────────────
   The model is built AFTER biFlattenDeep, so every value should already
   be a plain string. `_mmStr` exists because "should" is not "is": a
   legacy project, a half-migrated field, or a null in a cover row all
   arrive here eventually, and an exporter must not crash on one. */
function _mmStr(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return String(v);
    /* A bilingual object that escaped flattening — take any side rather
       than printing [object Object] into someone's training manual. */
    if (typeof v === 'object') {
        for (const k of ['en', 'fr', 'ar']) {
            if (typeof v[k] === 'string' && v[k].trim()) return v[k];
        }
    }
    return '';
}

function _mmArr(v) { return Array.isArray(v) ? v : (v ? [v] : []); }

function _mmLabel(type, fallbackKey) {
    if (!type) return fallbackKey;
    const key = type.label || type.key;
    try {
        if (typeof window !== 'undefined' && window.i18n && typeof window.i18n.t === 'function') {
            const txt = window.i18n.t(key);
            if (txt && txt !== key) return txt;
        }
    } catch (e) { /* i18n absent — the key is a poor caption but not a crash */ }
    return key;
}

/** Sheet numbers follow the tool's own auto-numbering when the author
    has not overridden them, so the HTML and the Word file agree on what
    "Information Sheet 2-1" means. */
function _mmSheetNumber(sheet, loIndex, sheetIndex) {
    const explicit = _mmStr(sheet && sheet.sheetNumber).trim();
    if (explicit) return explicit;
    if (typeof getAutoSheetNumber === 'function') return getAutoSheetNumber(loIndex, sheetIndex);
    return (loIndex + 1) + '-' + (sheetIndex + 1);
}

function _mmMarks(list) {
    return _mmArr(list)
        .filter(function (m) { return m && _mmStr(m.text).trim(); })
        .map(function (m) {
            /* The mark's own theme, looked up from resources.js. Colour
               and label belong to the mark type, not to this model, and
               not to whichever exporter happens to be running. */
            let type = null;
            if (typeof MARK_TYPES !== 'undefined' && Array.isArray(MARK_TYPES)) {
                type = MARK_TYPES.find(function (x) { return x.key === m.key; }) || null;
            }
            return {
                key: m.key,
                text: _mmStr(m.text),
                /* type.label is an i18n KEY ('mkAttention'), not a
                   caption. Resolved here so exporters receive text they
                   can print. */
                label:      _mmLabel(type, m.key),
                icon:       type ? (type.icon || '') : '',
                header:     type ? type.header : '#4f46e5',
                bg:         type ? type.bg : '#eef2ff',
                border:     type ? type.border : '#4f46e5',
                headerText: type ? type.headerText : '#ffffff'
            };
        });
}

function _mmTables(list) {
    return _mmArr(list).map(function (t) {
        const cells = _mmArr(t && t.cells).map(function (row) {
            return _mmArr(row).map(_mmStr);
        });
        return { rows: cells.length, cols: cells.length ? cells[0].length : 0, cells: cells };
    }).filter(function (t) {
        /* A table the author added and never typed into is furniture,
           not content. */
        return t.cells.some(function (r) { return r.some(function (c) { return c.trim(); }); });
    });
}

function _mmImages(bag, id) {
    if (!bag || !id) return [];
    return _mmArr(bag[id]).filter(function (im) { return im; });
}

/* ── The builder ────────────────────────────────────────────── */

/**
 * @param {string} lang  'en' | 'fr' | 'ar'
 * @param {object} [state]  defaults to window.mbState
 * @returns {object} a plain, serialisable description of the module
 */
function mbBuildModuleModel(lang, state) {
    const raw = state || (typeof window !== 'undefined' ? window.mbState : null);
    if (!raw) return null;

    /* Flatten ONCE, here. Every exporter downstream then works with
       plain strings and never has to think about bilingual pairs. */
    const st = (typeof biFlattenDeep === 'function') ? biFlattenDeep(raw, lang) : raw;

    const modules = _mmArr(st.modulesData);
    const currentModule = modules.find(function (m) { return m.id === st.currentModuleId; }) || modules[0] || null;

    const model = {
        lang: lang,
        rtl: (typeof biIsRtl === 'function') ? biIsRtl(lang) : (lang === 'ar'),
        title: currentModule ? _mmStr(currentModule.title) : '',
        generatedAt: new Date().toISOString(),

        cover: {
            frontImage: st.frontCoverImage || null,
            backImage:  st.backCoverImage || null,
            /* Rows whose VALUE is empty are dropped: a label with nothing
               beside it is a question the author never answered. */
            rows: _mmArr(st.coverRows)
                .map(function (r) { return { label: _mmStr(r && r.label), value: _mmStr(r && r.value) }; })
                .filter(function (r) { return r.value.trim(); }),
            /* Field names taken from mb_state.js, not guessed:
               coversAdditionalInfo sits above the table, and
               coversAdditionalNotes below it. */
            notesAbove: _mmStr(st.coversAdditionalInfo),
            notesBelow: _mmStr(st.coversAdditionalNotes)
        },

        intro: {
            team: _mmArr(st.teamMembers)
                .map(function (m) {
                    return {
                        name: _mmStr(m && m.name),
                        task: _mmStr(m && m.task),
                        location: _mmStr(m && m.workLocation)
                    };
                })
                .filter(function (m) { return m.name.trim() || m.task.trim() || m.location.trim(); }),
            blocks: _mmArr(st.introBlocks)
                .map(function (b) { return { title: _mmStr(b && b.title), body: _mmStr(b && b.body) }; })
                .filter(function (b) { return b.title.trim() || b.body.trim(); }),
            additional: _mmStr(st.introAdditionalDetails)
        },

        outcomes: [],
        assessment: [],

        references: {
            title: _mmStr(st.referencesTitle),
            items: _mmArr(st.referencesData)
                .map(function (r) { return _mmStr(r && r.value); })
                .filter(function (v) { return v.trim(); })
        }
    };

    _mmArr(st.learningOutcomesData).forEach(function (lo, loIndex) {
        /* Same gate as the Word export. An outcome that fails it does
           not appear at all — not as an empty heading, not as a stub. */
        if (!mmOutcomeHasContent(lo)) return;

        const out = {
            index: loIndex + 1,
            id: lo.id,
            title: _mmStr(lo.title),
            description: _mmStr(lo.description),
            criteria: _mmArr(lo.performanceCriteria)
                .map(function (c) { return _mmStr(c && (c.text !== undefined ? c.text : c)); })
                .filter(function (t) { return t.trim(); }),
            blocks: _mmArr(lo.blocks)
                .map(function (b) { return { title: _mmStr(b && b.title), body: _mmStr(b && b.body) }; })
                .filter(function (b) { return b.title.trim() || b.body.trim(); }),
            infoSheets: [],
            activitySheets: []
        };

        _mmArr(lo.infoSheets).forEach(function (sh, si) {
            if (!mmSheetHasContent(sh)) return;
            const number = _mmSheetNumber(sh, loIndex, si);
            out.infoSheets.push({
                number: number,
                title: _mmStr(sh.title),
                objectiveLead: _mmStr(sh.objectiveLead),
                objective: _mmStr(sh.objective),
                sections: _mmArr(sh.contentSections).map(function (cs) {
                    return {
                        heading: _mmStr(cs && cs.heading),
                        text: _mmStr(cs && cs.text),
                        images: _mmImages(sh.contentSectionImages, cs && cs.contentId),
                        marks: _mmMarks(cs && cs.marks),
                        tables: _mmTables(cs && cs.tables)
                    };
                }).filter(function (cs) {
                    return cs.heading.trim() || cs.text.trim() ||
                           cs.images.length || cs.marks.length || cs.tables.length;
                }),
                /* Numbered from the sheet, exactly as the Word export
                   does, and omitted entirely when the author wrote no
                   questions — an empty self-check tells the trainee to
                   answer questions that are not in the document. */
                selfCheck: _mmStr(sh.selfCheckContent).trim()
                    ? { number: _mmStr(sh.selfCheckNumber) || number, content: _mmStr(sh.selfCheckContent) }
                    : null,
                answersKey: _mmStr(sh.answersKeyContent).trim()
                    ? { number: _mmStr(sh.answersKeyNumber) || number, content: _mmStr(sh.answersKeyContent) }
                    : null,
                /* The QR IMAGE is global to the tool (mbState.infoQRImage),
                   while the link text belongs to the individual sheet.
                   That asymmetry is the tool's, and it is reproduced
                   rather than tidied. */
                qr: (st.infoQRImage || _mmStr(sh.linkSubject).trim() || _mmStr(sh.linkUrl).trim())
                    ? { image: st.infoQRImage || null,
                        url: _mmStr(sh.linkUrl),
                        subject: _mmStr(sh.linkSubject) }
                    : null
            });
        });

        _mmArr(lo.activitySheets).forEach(function (sh, si) {
            if (!mmSheetHasContent(sh)) return;
            const number = _mmSheetNumber(sh, loIndex, si);
            out.activitySheets.push({
                number: number,
                title: _mmStr(sh.title),
                objectiveLead: _mmStr(sh.objectiveLead),
                objective: _mmStr(sh.objective),
                duration: _mmStr(sh.duration),
                resources: _mmArr(sh.resources)
                    .map(function (r) { return { name: _mmStr(r && r.name), quantity: _mmStr(r && r.quantity) }; })
                    .filter(function (r) { return r.name.trim(); }),
                steps: _mmArr(sh.steps).map(function (stp, idx) {
                    return {
                        index: idx + 1,
                        text: _mmStr(stp && stp.text),
                        images: _mmImages(sh.stepImages, stp && stp.stepId),
                        marks: _mmMarks(stp && stp.marks)
                    };
                }).filter(function (s) { return s.text.trim() || s.images.length || s.marks.length; }),
                criteria: {
                    title: _mmStr(sh.criteriaTitle),
                    instruction: _mmStr(sh.criteriaInstruction),
                    footer: _mmStr(sh.criteriaFooter),
                    rows: _mmArr(sh.criteria)
                        .map(function (c) { return _mmStr(c && (c.text !== undefined ? c.text : c)); })
                        .filter(function (t) { return t.trim(); })
                },
                qr: (st.activityQRImage || _mmStr(sh.linkSubject).trim() || _mmStr(sh.linkUrl).trim())
                    ? { image: st.activityQRImage || null,
                        url: _mmStr(sh.linkUrl),
                        subject: _mmStr(sh.linkSubject) }
                    : null
            });
        });

        model.outcomes.push(out);
    });

    /* Assessment forms follow the outcomes that survived the gate, and
       only when actually filled. */
    const forms = st.assessmentFormsData || {};
    model.outcomes.forEach(function (out) {
        const form = forms[out.id];
        if (!mmAssessmentFilled(form)) return;
        model.assessment.push({
            outcomeIndex: out.index,
            outcomeTitle: out.title,
            rows: _mmArr(form.rows)
                .map(function (r) {
                    return {
                        criteria: _mmStr(r && r.criteria),
                        activities: _mmStr(r && r.activities),
                        outcomes: _mmStr(r && r.outcomes),
                        verification: _mmStr(r && r.verification),
                        date: _mmStr(r && r.date)
                    };
                })
                .filter(function (r) {
                    return r.criteria.trim() || r.activities.trim() || r.outcomes.trim() ||
                           r.verification.trim() || r.date.trim();
                }),
            competent: !!form.competent,
            notYetCompetent: !!form.notYetCompetent,
            teacherName: _mmStr(form.teacherName),
            teacherDate: _mmStr(form.teacherDate),
            learnerName: _mmStr(form.learnerName),
            learnerDate: _mmStr(form.learnerDate)
        });
    });

    return model;
}
