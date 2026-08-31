/* ============================================================
   exports_os_docx.js — Occupational Profile + Occupational Standard
   ============================================================

   A SECOND LAYOUT over data that already exists. Nothing here reads a
   field the application does not already collect, and nothing here is
   generated: the duties, the additional-information sections, the
   employability matrix and the competencies are the same objects the
   DACUM chart export prints, re-ordered into the two-part document that
   national qualification bodies actually endorse.

   Deliberately a separate module rather than a branch inside
   exports_docx.js. exportToWord() is the artefact every existing
   project has been validated against; a layout switch inside it would
   put that guarantee at risk for no gain. Here the existing exporter is
   untouched by construction, and the only edit to it was making seven
   private helpers exportable so the two documents cannot drift apart in
   how they handle Arabic.

   WHAT IS DELIBERATELY NOT COLLECTED
   The reference code (ASCO/ISCO), the organisational structure and the
   endorsement block are printed as EMPTY LABELLED LINES. They belong to
   a qualifications framework, not to a DACUM workshop, and adding three
   input fields for them would push framework bureaucracy into a tool
   whose job is to run a panel. They are filled in institutionally,
   after the document leaves this application.

   ARABIC
   Every paragraph — including the ones inside table cells — carries
   bidirectional:_rtl() and start-edge alignment, and every table
   carries visuallyRightToLeft. A cell paragraph that omits w:bidi still
   renders its text right-to-left (Unicode does that on its own) but
   anchors it to the LEFT edge of the cell and puts trailing punctuation
   on the wrong side, which is the specific defect this file has to
   avoid. _cellPara() below is the single place that gets it right, and
   no cell in this module builds a Paragraph any other way.
   ============================================================ */

import { appState } from './state.js';
import { showStatus } from './renderer.js';
import { getTaskCode } from './codes.js';
import {
    _rtl,
    _start,
    _font,
    _withArabicLang,
    _withArabicLangParagraph,
    _applyDocDefaultsLang,
    _safeFilename,
} from './exports_docx.js';

const _t  = (k)    => (window.i18n ? window.i18n.t(k)     : k);
const _tf = (k, v) => (window.i18n ? window.i18n.tf(k, v) : k);

/* Same grey the rest of the suite fills header bars with, and the same
   ShadingType.CLEAR rule: val="solid" paints the cell in the PATTERN
   colour and ignores w:fill, which renders as solid black. */
const HEADER_FILL = 'DCDCDC';
const LABEL_FILL  = 'F2F2F2';

const _val = (id) => {
    const el = document.getElementById(id);
    return el ? String(el.value || '').trim() : '';
};

const _lines = (text) =>
    String(text || '')
        .split('\n')
        .map(l => l.trim())
        .filter(Boolean);

export async function exportOccupationalStandardWord() {
    try {
        if (typeof window.docx === 'undefined') {
            showStatus(_t('msgDocxMissing') || 'docx library not loaded', 'error');
            return;
        }

        const {
            Document, Paragraph: _Paragraph, TextRun: _TextRun,
            Table, TableRow, TableCell, WidthType, AlignmentType,
            Packer, PageBreak, ShadingType,
        } = window.docx;

        const TextRun   = _withArabicLang(_TextRun);
        const Paragraph = _withArabicLangParagraph(_Paragraph, TextRun);

        /* ---- building blocks -------------------------------------- */

        // THE cell-paragraph constructor. See the Arabic note at the top
        // of this file: bidirectional + start alignment together are what
        // make a table cell read correctly in Arabic, and both are set
        // here so no call site can forget one of them.
        const _cellPara = (text, opts = {}) => new Paragraph({
            children: [new TextRun({
                text: String(text == null ? '' : text),
                bold: !!opts.bold,
                size: opts.size || 22,
            })],
            alignment: opts.center ? AlignmentType.CENTER : _start(AlignmentType),
            bidirectional: _rtl(),
            spacing: { before: 40, after: 40 },
        });

        const _cell = (text, opts = {}) => new TableCell({
            children: Array.isArray(text)
                ? (text.length ? text.map(t => _cellPara(t, opts)) : [_cellPara('', opts)])
                : [_cellPara(text, opts)],
            width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
            shading: opts.fill
                ? { fill: opts.fill, type: ShadingType.CLEAR, color: 'auto' }
                : undefined,
            columnSpan: opts.span || undefined,
        });

        const _table = (rows) => new Table({
            visuallyRightToLeft: _rtl(),
            width: { size: 9071, type: WidthType.DXA },
            layout: 'fixed',
            rows,
        });

        // Two-column label/value row — the shape of every header block in
        // the standard.
        const _kvRow = (label, value) => new TableRow({
            children: [
                _cell(label, { bold: true, width: 2600, fill: LABEL_FILL }),
                _cell(value, { width: 6471 }),
            ],
        });

        const _h1 = (text) => new Paragraph({
            children: [new TextRun({ text, bold: true, size: 32 })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 300, after: 300 },
            bidirectional: _rtl(),
        });

        const _h2 = (text) => new Paragraph({
            children: [new TextRun({ text, bold: true, size: 26 })],
            spacing: { before: 300, after: 150 },
            alignment: _start(AlignmentType),
            bidirectional: _rtl(),
        });

        const _body = (text, indent) => new Paragraph({
            children: [new TextRun({ text, size: 22 })],
            spacing: { after: 100 },
            indent: indent ? { left: 720 } : undefined,
            alignment: _start(AlignmentType),
            bidirectional: _rtl(),
        });

        const _spacer = () => new Paragraph({ text: '', bidirectional: _rtl() });
        const _break  = () => new Paragraph({ children: [new PageBreak()], bidirectional: _rtl() });

        const children = [];

        /* ============================================================
           PART 1 — OCCUPATIONAL PROFILE
           ============================================================ */

        const occupation = _val('occupationTitle');
        const job        = _val('jobTitle');

        children.push(_h1(_t('osPart1Title')));
        children.push(new Paragraph({
            children: [new TextRun({ text: occupation, bold: true, size: 28 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            bidirectional: _rtl(),
        }));

        // Panel block. The workshop provenance is what makes a profile
        // defensible, so it leads the document rather than trailing it.
        const panelRows = [
            _kvRow(_t('osFieldOccupation'), occupation),
        ];
        if (job) panelRows.push(_kvRow(_t('osFieldJob'), job));
        panelRows.push(
            _kvRow(_t('osFieldSector'),      _val('sector')),
            _kvRow(_t('osFieldContext'),     _val('context')),
            _kvRow(_t('osFieldProducedFor'), _val('producedFor')),
            _kvRow(_t('osFieldProducedBy'),  _val('producedBy')),
            _kvRow(_t('osFieldFacilitators'), _lines(_val('facilitators'))),
            _kvRow(_t('osFieldPanel'),        _lines(_val('panelMembers'))),
            _kvRow(_t('osFieldObservers'),    _lines(_val('observers'))),
            _kvRow(_t('osFieldVenueDate'),
                [_val('venue'), _val('workshopDate') || _val('date')].filter(Boolean).join(' — ')),
        );
        children.push(_table(panelRows));

        // ---- Duties and Tasks -------------------------------------
        // Read from the live DOM, exactly as exportToWord does, so the
        // two documents can never disagree about what the chart says.
        const duties = [];
        document.querySelectorAll('input[data-duty-id], textarea[data-duty-id]').forEach(dutyInput => {
            const dutyText = String(dutyInput.value || '').trim();
            if (!dutyText) return;
            const dutyId = dutyInput.getAttribute('data-duty-id');
            const tasks = [];
            document.querySelectorAll(
                `input[data-task-id^="${dutyId}_"], textarea[data-task-id^="${dutyId}_"]`
            ).forEach(taskInput => {
                const t = String(taskInput.value || '').trim();
                if (t) tasks.push(t);
            });
            duties.push({ duty: dutyText, tasks });
        });

        if (duties.length) {
            children.push(_break());
            children.push(_h2(_t('osDutiesTasks')));

            const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
            duties.forEach((d, di) => {
                const letter = letters[di] || String(di + 1);
                const rows = [
                    new TableRow({
                        children: [_cell(`${letter}. ${d.duty}`, {
                            bold: true, fill: HEADER_FILL, span: 2,
                        })],
                    }),
                ];
                d.tasks.forEach((task, ti) => {
                    rows.push(new TableRow({
                        children: [
                            _cell(`${letter}${ti + 1}`, { bold: true, width: 1000, center: true }),
                            _cell(task, { width: 8071 }),
                        ],
                    }));
                });
                children.push(_table(rows));
                children.push(_spacer());
            });
        }

        // ---- Profile narrative sections ----------------------------
        // Concerns has no dedicated field; a facilitator who wants it
        // separate from Trends adds a custom section, which is picked up
        // by the custom-section loop below at no code cost.
        const profileSections = [
            ['behaviorsHeading',  'behaviorsInput'],
            ['knowledgeHeading',  'knowledgeInput'],
            ['skillsHeading',     'skillsInput'],
            ['trendsHeading',     'trendsInput'],
            ['careerPathHeading', 'careerPathInput'],
            ['acronymsHeading',   'acronymsInput'],
        ];

        const narrative = [];
        profileSections.forEach(([headId, inputId]) => {
            const head = document.getElementById(headId);
            const body = _val(inputId);
            if (!body) return;
            narrative.push({ head: head ? head.textContent.trim() : inputId, body });
        });

        const customContainer = document.getElementById('customSectionsContainer');
        if (customContainer) {
            customContainer.querySelectorAll('.section-container').forEach(div => {
                const h = div.querySelector('input[type="text"], .section-heading');
                const t = div.querySelector('textarea');
                const head = h ? String(h.value || h.textContent || '').trim() : '';
                const body = t ? String(t.value || '').trim() : '';
                if (head && body) narrative.push({ head, body });
            });
        }

        if (narrative.length) {
            children.push(_break());
            narrative.forEach(sec => {
                children.push(_table([
                    new TableRow({ children: [_cell(sec.head, { bold: true, fill: HEADER_FILL })] }),
                    new TableRow({ children: [_cell(_lines(sec.body))] }),
                ]));
                children.push(_spacer());
            });
        }

        /* ============================================================
           PART 2 — OCCUPATIONAL STANDARD
           ============================================================ */

        children.push(_break());
        children.push(_h1(_t('osPart2Title')));

        // Header block. The four blank rows are the endorsement chain —
        // see the note at the top of this file for why they are printed
        // empty rather than collected in the UI.
        children.push(_table([
            _kvRow(_t('osFieldStandardTitle'), occupation),
            _kvRow(_t('osFieldSector'),        _val('sector')),
            _kvRow(_t('osFieldRefCode'),       ''),
            _kvRow(_t('osFieldScope'),         _lines(_val('scopeOfWork'))),
            _kvRow(_t('osFieldDevelopedBy'),   _val('producedBy')),
            _kvRow(_t('osFieldEndorsedBy'),    ''),
            _kvRow(_t('osFieldApprovedBy'),    ''),
            _kvRow(_t('osFieldApprovalDate'),  ''),
            _kvRow(_t('osFieldReviewDate'),    ''),
        ]));

        // ---- Employability competencies by occupational level -------
        // skillsLevelData is an ARRAY of categories, each holding
        // competencies whose `levels` object carries the four booleans —
        // not a flat list. Getting this shape wrong produces a silently
        // empty matrix rather than an error.
        const sl = appState.skillsLevelData;
        if (Array.isArray(sl) && sl.length) {
            children.push(_h2(_t('expEmployability')));

            const levelKeys = ['craftsman', 'skilled', 'semiSkilled', 'foundation'];
            const levelLabels = [
                _t('expCraftsman'), _t('expSkilled'),
                _t('expSemiSkilled'), _t('expFoundation'),
            ];

            const rows = [new TableRow({
                tableHeader: true,
                children: [
                    _cell(_t('osColCompetency'), { bold: true, fill: HEADER_FILL, width: 4271 }),
                    ...levelLabels.map(l => _cell(l, { bold: true, fill: HEADER_FILL, center: true, width: 1200 })),
                ],
            })];

            let printed = 0;
            sl.forEach(cat => {
                const catName = cat.category || '';
                const comps = Array.isArray(cat.competencies) ? cat.competencies : [];
                // Seed row 9 is a blank spare for the facilitator; skip it
                // rather than printing an empty banner.
                if (!catName && !comps.some(c => c.text)) return;

                rows.push(new TableRow({
                    children: [_cell(catName, { bold: true, fill: LABEL_FILL, span: 5 })],
                }));

                comps.forEach(comp => {
                    if (!comp.text) return;
                    const lv = comp.levels || {};
                    rows.push(new TableRow({
                        children: [
                            _cell(comp.text, { width: 4271 }),
                            ...levelKeys.map(k => _cell(lv[k] ? 'X' : '', { center: true, width: 1200 })),
                        ],
                    }));
                    printed++;
                });
            });

            if (printed) {
                children.push(_table(rows));
                children.push(_spacer());
            } else {
                children.pop(); // drop the heading we just pushed
            }
        }

        // ---- Competencies ------------------------------------------
        const clusters = (appState.clusteringData && appState.clusteringData.clusters) || [];
        clusters.forEach((cluster, i) => {
            const n = i + 1;
            children.push(_break());
            children.push(_h2(_tf('expCompetencyN', { n, name: cluster.name })));

            const rows = [];

            if (cluster.range && cluster.range.trim()) {
                rows.push(new TableRow({
                    children: [
                        _cell(_t('expRangeLabel'), { bold: true, width: 2600, fill: LABEL_FILL }),
                        _cell(_lines(cluster.range), { width: 6471 }),
                    ],
                }));
            }

            if (Array.isArray(cluster.tasks) && cluster.tasks.length) {
                rows.push(new TableRow({
                    children: [
                        _cell(_t('osRelatedTasksFromProfile'), { bold: true, width: 2600, fill: LABEL_FILL }),
                        _cell(cluster.tasks.map(task => `${getTaskCode(task.id)}: ${task.text}`), { width: 6471 }),
                    ],
                }));
            }

            if (Array.isArray(cluster.performanceCriteria) && cluster.performanceCriteria.length) {
                rows.push(new TableRow({
                    children: [
                        _cell(_t('expPCLabel'), { bold: true, width: 2600, fill: LABEL_FILL }),
                        _cell(cluster.performanceCriteria.map((c, ci) => `${n}.${ci + 1}  ${c}`), { width: 6471 }),
                    ],
                }));
            }

            /* Performance criteria are numbered n.1, n.2 … here and
               nowhere else in the suite. The UNESCO-style module
               descriptor cites them as "5.6; 5.1" — a reference that
               cannot be checked unless the standard actually prints
               those numbers. Without them the citation is unverifiable,
               which is how a curriculum quietly detaches from its
               standard. */

            if (rows.length) children.push(_table(rows));
        });

        // ---- Tools, equipment and materials -------------------------
        const tools = _val('toolsInput');
        if (tools) {
            children.push(_spacer());
            children.push(_table([
                new TableRow({
                    children: [_cell(
                        (document.getElementById('toolsHeading')?.textContent || '').trim() || _t('osToolsEquipment'),
                        { bold: true, fill: HEADER_FILL }
                    )],
                }),
                new TableRow({ children: [_cell(_lines(tools))] }),
            ]));
        }

        /* ---- assemble --------------------------------------------- */

        const doc = new Document({
            styles: { default: { document: { run: { font: _font() } } } },
            sections: [{
                properties: {
                    page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } },
                },
                children,
            }],
        });
        _applyDocDefaultsLang(doc);

        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = _safeFilename(occupation + (job ? ` ${job}` : ''), '_OS.docx');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

        showStatus(_t('msgOSExported') + ' ✓', 'success');

    } catch (error) {
        console.error('Error generating Occupational Standard document:', error);
        showStatus(_tf('msgWordError', { msg: error.message }), 'error');
    }
}
