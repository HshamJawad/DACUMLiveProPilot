// ============================================================
// /exports_docx.js
// ------------------------------------------------------------
// Word (.docx) generation: the standalone Task Verification report
// and the full DACUM chart report.
//
// Split from exports.js. Depends on the `docx` library being loaded
// globally by index.html.
// ============================================================

import { appState } from './state.js';
import { showStatus } from './renderer.js';
import { getTaskCode, getDutyLetter } from './codes.js';
import { buildVerificationDataset, getVerificationCoverage } from './exports_shared.js';
import { noteExportExclusion } from './draft_unverified.js';


/* ── i18n + direction helpers ────────────────────────────────────────
   Every paragraph in this file used to carry `bidirectional: false`
   with the comment "Force LTR" — 152 of them. That was correct while
   the app was English-only and it is the single thing that made an
   Arabic Word export unusable: with an LTR base direction, an Arabic
   paragraph renders right-aligned text with its punctuation, brackets
   and any Latin fragments resolved against the wrong base, so a line
   ending in a full stop puts the stop on the left.

   What `bidirectional` sets is the paragraph's BASE direction, not the
   script. Word still runs the Unicode bidi algorithm inside the
   paragraph, so a Latin tool name or an ISO code inside Arabic prose
   comes out correctly either way. Basing it on the export language is
   therefore both sufficient and correct. */
const _t   = (k)    => (window.i18n ? window.i18n.t(k)     : k);
const _tf  = (k, v) => (window.i18n ? window.i18n.tf(k, v) : k);
const _rtl = ()     => (window.i18n ? window.i18n.isRTL()  : false);

/* Start/end alignment. AlignmentType has no logical START, so it is
   resolved here rather than at 48 call sites. */
const _start = (A) => (_rtl() ? A.RIGHT : A.LEFT);

/* Arabic needs a face that actually carries the glyphs. Word falls back
   silently when it cannot find one, which is how a document ends up
   full of boxes on a machine without the original font. Amiri and
   Cairo are common on Arabic systems; Arial ships everywhere and has
   full Arabic coverage, so it is the safe default rather than the
   pretty one. */
const _font = () => (_rtl() ? 'Arial' : 'Calibri');

/* Dates in the exported document follow the EXPORT language, not the
   browser's locale — those are frequently different, and a report is a
   deliverable that must be internally consistent. */
const _today = () => new Date().toLocaleDateString(
  window.i18n ? window.i18n.getLang() : undefined
);


/* A filename built with /[^a-z0-9]/gi turns an Arabic occupation title
   into a row of underscores — every Arabic export arrived as
   "________.docx". Keep Unicode letters and digits; strip only what a
   filesystem actually objects to. */
function _safeFilename(title, suffix) {
  const base = String(title || '')
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, '')   // illegal on Windows
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 80);
  return (base || _t('fileUntitled')) + suffix;
}


export async function exportTaskVerificationWord() {
    // Tell the user WHY the appendix is missing rather than
    // shipping a report that is quietly short a section.
    noteExportExclusion();

            try {
                // Works in BOTH collection modes. The old guard rejected
                // anything that was not workshop mode before inspecting
                // the data at all, so Individual/Survey could never be
                // exported however complete it was.
                const tvResults = buildVerificationDataset();

                const validResults = Object.keys(tvResults).filter(key =>
                    tvResults[key] && tvResults[key].valid
                );

                // One fully-rated task is enough. A partial export is a
                // normal thing to want mid-analysis; the coverage line
                // below states plainly how partial it is.
                if (validResults.length === 0) {
                    alert('No task has been fully rated yet.\n\n' +
                          'Rate Importance, Frequency and Learning Difficulty for at least ' +
                          'one task in the Task Verification tab, then export again.');
                    return;
                }

                const tvCoverage = getVerificationCoverage(tvResults);
                
                if (typeof window.docx === 'undefined') {
                    showStatus(_t('msgDocxLibMissing'), 'error');
                    return;
                }

                const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, ShadingType, Packer } = window.docx;

                showStatus(_t('msgGeneratingTVWord'), 'success');

                const children = [];
                
                const occupationTitleInput = document.getElementById('occupationTitle');
                const occupationTitle = occupationTitleInput ? occupationTitleInput.value : 'Unknown Occupation';
                
                // Title
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _t('expTaskVerification'),
                            bold: true,
                            size: 32,
                        }),
                    ],
                    spacing: { after: 300 },
                    bidirectional: _rtl(),
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expOccupation', { v: occupationTitle }),
                            bold: true,
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));
                
                const today = _today();
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expDateOfAnalysis', { v: today }),
                            size: 24,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expBasedOn', { v: occupationTitle }),
                            italics: true,
                            size: 20,
                        }),
                    ],
                    spacing: { after: 400 },
                    bidirectional: _rtl(),
                }));
                
                // Methodology Summary
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _t('expMethodologySummary'),
                            bold: true,
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expCollectionMode', { v: _t(appState.collectionMode === 'workshop' ? 'modeWorkshop' : 'expIndividualSurvey') }),
                            size: 22,
                        }),
                    ],
                    spacing: { after: 100 },
                    bidirectional: _rtl(),
                }));
                
                // Participant count only exists when there was a panel.
                // Printing appState.workshopParticipants in Individual /
                // Survey mode would state an evidence base that does not
                // exist for these numbers.
                if (appState.collectionMode === 'workshop') {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _tf('expParticipants', { v: appState.workshopParticipants }),
                                size: 22,
                            }),
                        ],
                        spacing: { after: 100 },
                        bidirectional: _rtl(),
                    }));
                }

                // Coverage. Bold and red when partial, so a work-in-progress
                // export can never be mistaken for a completed verification.
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expCoverage', { v: tvCoverage.label }),
                            size: 22,
                            bold: !tvCoverage.complete,
                            color: tvCoverage.complete ? '000000' : 'B91C1C',
                        }),
                    ],
                    spacing: { after: 100 },
                    bidirectional: _rtl(),
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expWorkflowMode', { v: _t(appState.workflowMode === 'standard' ? 'modeStandard' : 'modeExtended') }),
                            size: 22,
                        }),
                    ],
                    spacing: { after: 100 },
                    bidirectional: _rtl(),
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expPriorityFormula', { v: _t(appState.priorityFormula === 'if' ? 'formulaIF' : 'formulaIFD') }),
                            size: 22,
                        }),
                    ],
                    spacing: { after: 400 },
                    bidirectional: _rtl(),
                }));
                
                // Priority Rankings
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _t('expPriorityRankings'),
                            bold: true,
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));
                
                // Get and sort results
                const sortedResults = [];
                validResults.forEach(taskKey => {
                    const result = tvResults[taskKey];
                    
                    // Use stored duty and task titles (with backward compatibility)
                    let dutyText = result.dutyTitle;
                    let taskText = result.taskTitle;
                    
                    // Backward compatibility: if not stored, look up from DOM
                    if (!dutyText || !taskText) {
                        const taskParts = taskKey.split('_task_');
                        const dutyId = taskParts[0];
                        
                        if (!dutyText) {
                            const dutyInput = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`);
                            dutyText = dutyInput ? dutyInput.value.trim() : 'Unassigned';
                        }
                        
                        if (!taskText) {
                            const taskInput = document.querySelector(`input[data-task-id="${taskKey}"], textarea[data-task-id="${taskKey}"]`);
                            taskText = taskInput ? taskInput.value.trim() : 'Unassigned';
                        }
                    }
                    
                    sortedResults.push({
                        duty: dutyText,
                        task: taskText,
                        meanI: result.meanImportance,
                        meanF: result.meanFrequency,
                        meanD: result.meanDifficulty,
                        priority: result.priorityIndex
                    });
                });
                
                sortedResults.sort((a, b) => b.priority - a.priority);
                
                // Create table
                const tableRows = [];
                
                // Header row
                tableRows.push(new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: _t('expRank'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: _t('expDutyLabel'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: _t('expTaskLabel'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: _t('expMeanI'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: _t('expMeanF'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: _t('expMeanD'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: _t('expPriority'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                    ],
                }));
                
                // Data rows
                sortedResults.forEach((row, index) => {
                    tableRows.push(new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `#${index + 1}` })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ text: row.duty, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ text: row.task, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanI !== null ? row.meanI.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanF !== null ? row.meanF.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanD !== null ? row.meanD.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.priority !== null ? row.priority.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                        ],
                    }));
                });
                
                children.push(new Table({
                    visuallyRightToLeft: _rtl(),
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: tableRows,
                }));
                
                // Duty-Level Summary section
                children.push(new Paragraph({ spacing: { after: 400 } }));
                
                children.push(new Paragraph({
                    children: [new TextRun({ text: _t('expDutyLevelSummary'), bold: true, size: 28 })],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));
                
                children.push(new Paragraph({
                    children: [new TextRun({ text: _tf('expTrainingLoadMethod', { v: _t(appState.trainingLoadMethod === 'advanced' ? 'expAdvancedMethod' : 'expSimpleMethod') }), size: 20, italics: true })],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));
                
                // Aggregate duty-level data
                const dutyMap = {};
                Object.keys(tvResults).forEach(taskKey => {
                    const result = tvResults[taskKey];
                    if (result && result.valid) {
                        let dutyId = result.dutyId || taskKey.split('_task_')[0];
                        let dutyTitle = result.dutyTitle;
                        
                        if (!dutyTitle) {
                            const dutyInput = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`);
                            dutyTitle = dutyInput ? dutyInput.value.trim() : 'Unassigned';
                        }
                        
                        if (!dutyMap[dutyId]) {
                            dutyMap[dutyId] = { dutyTitle: dutyTitle, validTasks: 0, prioritySum: 0, tasks: [] };
                        }
                        
                        dutyMap[dutyId].validTasks++;
                        dutyMap[dutyId].prioritySum += result.priorityIndex;
                        dutyMap[dutyId].tasks.push({ priorityIndex: result.priorityIndex, meanDifficulty: result.meanDifficulty });
                    }
                });
                
                const dutyResults = [];
                Object.keys(dutyMap).forEach(dutyId => {
                    const duty = dutyMap[dutyId];
                    const avgPriority = duty.prioritySum / duty.validTasks;
                    let trainingLoad = 0;
                    if (appState.trainingLoadMethod === 'advanced') {
                        trainingLoad = duty.tasks.reduce((sum, t) => sum + (t.priorityIndex * t.meanDifficulty), 0);
                    } else {
                        trainingLoad = avgPriority * duty.validTasks;
                    }
                    dutyResults.push({ dutyTitle: duty.dutyTitle, validTasks: duty.validTasks, avgPriority: avgPriority, trainingLoad: trainingLoad });
                });
                
                dutyResults.sort((a, b) => b.avgPriority - a.avgPriority);
                
                // Duty table
                const dutyTableRows = [
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expDutyTitle'), bold: true })], alignment: _start(AlignmentType), bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expTasks'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expAvgPriority'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expTrainingLoad'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                        ],
                    })
                ];
                
                dutyResults.forEach(duty => {
                    dutyTableRows.push(new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: duty.dutyTitle, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.validTasks.toString() })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.avgPriority.toFixed(2) })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.trainingLoad.toFixed(2), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                        ],
                    }));
                });
                
                children.push(new Table({
                    visuallyRightToLeft: _rtl(),
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: dutyTableRows,
                }));
                
                // Notes section
                children.push(new Paragraph({ spacing: { after: 400 } }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _t('expNotesMethodology'),
                            bold: true,
                            size: 24,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));
                
                const notes = [
                    'Weighted Mean = Σ(value × count) ÷ total responses',
                    'Importance scale: 0=Not Important, 1=Somewhat, 2=Important, 3=Critical',
                    'Frequency scale: 0=Rarely, 1=Sometimes, 2=Often, 3=Daily',
                    'Difficulty scale: 0=Easy, 1=Moderate, 2=Challenging, 3=Very Difficult',
                    `Priority Index = ${appState.priorityFormula === 'if' ? 'Mean Importance × Mean Frequency' : 'Mean Importance × Mean Frequency × Mean Difficulty'}`,
                    'Higher priority values indicate greater training importance',
                    'Results follow DACUM (Developing A Curriculum) methodology'
                ];
                
                notes.forEach(note => {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `• ${note}`,
                                size: 20,
                            }),
                        ],
                        spacing: { after: 100 },
                        bidirectional: _rtl(),
                    }));
                });
                
                // Create document
                const doc = new Document({
                    styles: {
                        default: {
                            document: { run: { font: _font() } },
                        },
                    },
                    sections: [{
                        properties: {
                            /* bidi on the SECTION makes Word treat the whole
                               section as right-to-left: table columns run
                               right-to-left, and list markers and numbering
                               sit on the correct edge. Paragraph-level
                               bidirectional alone does not do either. */
                            bidi: _rtl(),
                            page: {
                                margin: {
                                    top: 1440,
                                    right: 1440,
                                    bottom: 1440,
                                    left: 1440,
                                },
                            },
                        },
                        children: children,
                    }],
                });

                // Generate and download
                const blob = await Packer.toBlob(doc);
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = _safeFilename(occupationTitle, '_Task_Verification.docx');
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                
                showStatus(_t('msgWordExported') + ' ✓', 'success');

            } catch (error) {
                console.error('Error generating Task Verification Word document:', error);
                showStatus(_tf('msgTVWordError', { msg: error.message }), 'error');
            }
        }

export async function exportToWord() {
    // Tell the user WHY the appendix is missing rather than
    // shipping a report that is quietly short a section.
    noteExportExclusion();

    // ── TABLE SHADING ──────────────────────────────────────────
    // Every shaded cell in this document uses DCDCDC = RGB(220,220,220),
    // the same grey the PDF exporter fills duty bars with. This export
    // previously mixed four tints (667eea purple headers, E8E8E8,
    // F5F5F5 and DCDCDC), so the Word and PDF versions of the same
    // chart read as two different documents.
    //
    // IMPORTANT — always use ShadingType.CLEAR here, never SOLID.
    // In OOXML, w:shd carries BOTH a background (w:fill) and a pattern
    // foreground (w:color). val="solid" means "paint the cell 100% in
    // the PATTERN colour", so w:fill is ignored entirely — and where no
    // colour was given it defaulted to "auto", i.e. black. That is why
    // these bars rendered as solid black blocks regardless of the fill
    // value set. val="clear" means "no pattern", which lets w:fill show
    // through as an ordinary background. Cell text stays black.
            // ============ CHECK FOR VERIFIED LIVE WORKSHOP RESULTS ============
            const hasVerifiedResults = typeof appState.lwFinalizedData !== 'undefined' && appState.lwFinalizedData && 
                                        typeof appState.lwAggregatedResults !== 'undefined' && appState.lwAggregatedResults;
            
            // ============ VERIFIED LIVE WORKSHOP STANDALONE EXPORT ============
            if (hasVerifiedResults && appState.tvExportMode === 'standalone') {
                await lwExportVerifiedDOCX();
                return;
            }
            
            // ============ REGULAR TASK VERIFICATION STANDALONE EXPORT ============
            if (!hasVerifiedResults && appState.tvExportMode === 'standalone') {
                await exportTaskVerificationWord();
                return;
            }
            
            // ============ NORMAL DACUM EXPORT (with optional appendix) ============
            try {
                if (typeof window.docx === 'undefined') {
                    showStatus(_t('msgDocxLibMissing'), 'error');
                    return;
                }

                const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, Packer, PageBreak, convertInchesToTwip, ShadingType, TextDirection, ImageRun } = window.docx;

                // Get all input values
                const dacumDateValue = document.getElementById('dacumDate').value;
                let dacumDate = '';
                if (dacumDateValue) {
                    const dateObj = new Date(dacumDateValue + 'T00:00:00');
                    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                    const day = String(dateObj.getDate()).padStart(2, '0');
                    const year = dateObj.getFullYear();
                    dacumDate = `${month}/${day}/${year}`;
                }
                const producedFor = document.getElementById('producedFor').value;
                const producedBy = document.getElementById('producedBy').value;
                const occupationTitle = document.getElementById('occupationTitle').value;
                const jobTitle = document.getElementById('jobTitle').value;

                if (!occupationTitle) {
                    alert('Please enter an Occupation Title before exporting.');
                    showStatus(_t('msgOccupationRequiredExport'), 'error');
                    return;
                }

                showStatus(_t('msgGeneratingWord'), 'success');

                const children = [];

                // ============ TITLE PAGE ============
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: _tf('expOccupationTitle', { v: occupationTitle }),
                            bold: true,
                            size: 28, // 14pt
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: _rtl(),
                }));

                // Scope of Work / Occupational Definition (optional)
                const scopeOfWorkEl    = document.getElementById('scopeOfWork');
                const scopeOfWorkValue = (scopeOfWorkEl ? scopeOfWorkEl.value : '').trim();
                if (scopeOfWorkValue) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _t('expScopeOfWork'),
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { before: 80, after: 80 },
                        bidirectional: _rtl(),
                    }));
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: scopeOfWorkValue,
                                size: 22, // 11pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: _rtl(),
                    }));
                }

                // Job Title is optional — skip the paragraph entirely when empty
                if (jobTitle && jobTitle.trim()) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _tf('expJobTitle', { v: jobTitle }),
                                bold: true,
                                size: 28, // 14pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: _rtl(),
                    }));
                }

                // Add DACUM Date if exists
                if (dacumDate) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _tf('expDacumDate', { v: dacumDate }),
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: _rtl(),
                    }));
                }
                
                // Add Venue if exists
                const venueValue = document.getElementById('venue')?.value;
                if (venueValue) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _tf('expVenue', { v: venueValue }),
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: _rtl(),
                    }));
                }

                // Add Produced For if exists
                if (producedFor) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _tf('expProducedFor', { v: producedFor }),
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: _rtl(),
                    }));
                    
                    // Add Produced For logo if exists
                    if (appState.producedForImage) {
                        try {
                            const base64Data = appState.producedForImage.split(',')[1];
                            
                            children.push(new Paragraph({
                                children: [
                                    new ImageRun({
                                        data: Uint8Array.from(atob(base64Data), c => c.charCodeAt(0)),
                                        transformation: {
                                            width: 94, // 2.5cm = 94 points approximately
                                            height: 94,
                                        },
                                    }),
                                ],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 200 },
                            }));
                        } catch (imgError) {
                            console.error('Error adding Produced For image:', imgError);
                        }
                    }
                }

                // Add Produced By if exists
                if (producedBy) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _tf('expProducedBy', { v: producedBy }),
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: _rtl(),
                    }));
                    
                    // Add Produced By logo if exists
                    if (appState.producedByImage) {
                        try {
                            const base64Data = appState.producedByImage.split(',')[1];
                            
                            children.push(new Paragraph({
                                children: [
                                    new ImageRun({
                                        data: Uint8Array.from(atob(base64Data), c => c.charCodeAt(0)),
                                        transformation: {
                                            width: 94, // 2.5cm = 94 points approximately
                                            height: 94,
                                        },
                                    }),
                                ],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 400 },
                            }));
                        } catch (imgError) {
                            console.error('Error adding Produced By image:', imgError);
                        }
                    }
                } else {
                    // Add extra spacing if no Produced By section
                    children.push(new Paragraph({ spacing: { after: 200 } }));
                }

                // Workshop Roles Section
                const facilitatorsText = document.getElementById('facilitators')?.value.trim();
                const observersText = document.getElementById('observers')?.value.trim();
                const panelMembersText = document.getElementById('panelMembers')?.value.trim();
                
                if (facilitatorsText) {
                    const facilitatorNames = facilitatorsText.split('\n').map(s => s.trim()).filter(s => s);
                    if (facilitatorNames.length > 0) {
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expFacilitators'),
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: _rtl(),
                        }));
                        
                        const facilitatorRows = facilitatorNames.map(name => 
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: name,
                                                        size: 22, // 11pt
                                                    }),
                                                ],
                                                bidirectional: _rtl(),
                                            }),
                                        ],
                                    }),
                                ],
                            })
                        );
                        
                        children.push(new Table({
                            visuallyRightToLeft: _rtl(),
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: facilitatorRows,
                        }));
                    }
                }
                
                if (observersText) {
                    const observerNames = observersText.split('\n').map(s => s.trim()).filter(s => s);
                    if (observerNames.length > 0) {
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expObservers'),
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: _rtl(),
                        }));
                        
                        const observerRows = observerNames.map(name => 
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: name,
                                                        size: 22, // 11pt
                                                    }),
                                                ],
                                                bidirectional: _rtl(),
                                            }),
                                        ],
                                    }),
                                ],
                            })
                        );
                        
                        children.push(new Table({
                            visuallyRightToLeft: _rtl(),
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: observerRows,
                        }));
                    }
                }
                
                if (panelMembersText) {
                    const panelMemberNames = panelMembersText.split('\n').map(s => s.trim()).filter(s => s);
                    if (panelMemberNames.length > 0) {
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expPanelMembers'),
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: _rtl(),
                        }));
                        
                        const panelMemberRows = panelMemberNames.map(name => 
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: name,
                                                        size: 22, // 11pt
                                                    }),
                                                ],
                                                bidirectional: _rtl(),
                                            }),
                                        ],
                                    }),
                                ],
                            })
                        );
                        
                        children.push(new Table({
                            visuallyRightToLeft: _rtl(),
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: panelMemberRows,
                        }));
                    }
                }

                // ============ DUTIES AND TASKS (NEW PAGE) ============
                children.push(new Paragraph({
                    children: [
                        new PageBreak(),
                        new TextRun({
                            text: _t('expDutiesAndTasks'),
                            bold: true,
                            size: 28, // 14pt
                        }),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 300 },
                    bidirectional: _rtl(),
                }));

                // Collect duties and tasks
                const dutyInputs = document.querySelectorAll('input[data-duty-id], textarea[data-duty-id]');
                const duties = [];
                
                dutyInputs.forEach(dutyInput => {
                    const dutyText = dutyInput.value.trim();
                    if (dutyText) {
                        const dutyId = dutyInput.getAttribute('data-duty-id');
                        const taskInputs = document.querySelectorAll(`input[data-task-id^="${dutyId}_"], textarea[data-task-id^="${dutyId}_"]`);
                        const tasks = [];
                        
                        taskInputs.forEach(taskInput => {
                            const taskText = taskInput.value.trim();
                            if (taskText) {
                                tasks.push(taskText);
                            }
                        });
                        
                        duties.push({
                            duty: dutyText,
                            tasks: tasks
                        });
                    }
                });

                // Create a table for each duty
                duties.forEach((dutyData, dutyIndex) => {
                    const dutyLetter = getDutyLetter(dutyIndex);
                    const dutyLabel = `${_tf('lblDuty', { code: dutyLetter })}: ${dutyData.duty}`;
                    
                    // Calculate number of rows needed (header + task rows)
                    const tasksPerRow = 4;
                    const numTaskRows = Math.ceil(dutyData.tasks.length / tasksPerRow);
                    const tableRows = [];
                    
                    // Header row (duty description spans all 4 columns)
                    tableRows.push(
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: dutyLabel,
                                                    bold: true,
                                                    size: 24, // 12pt
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    columnSpan: 4,
                                    shading: {
                                        // DCDCDC = RGB(220,220,220) — the exact grey the
                                        // PDF exporter fills duty bars with, so the Word
                                        // and PDF versions of the same chart match.
                                        fill: "DCDCDC",
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                    width: {
                                        size: 100,
                                        type: WidthType.PERCENTAGE,
                                    },
                                }),
                            ],
                        })
                    );
                    
                    // Task rows (4 tasks per row)
                    for (let row = 0; row < numTaskRows; row++) {
                        const rowCells = [];
                        
                        for (let col = 0; col < tasksPerRow; col++) {
                            const taskIndex = row * tasksPerRow + col;
                            
                            if (taskIndex < dutyData.tasks.length) {
                                const taskLabel = _tf('lblTask', { code: `${dutyLetter}${taskIndex + 1}` });
                                const taskText = `${taskLabel}: ${dutyData.tasks[taskIndex]}`;
                                
                                rowCells.push(
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: taskText,
                                                        size: 24, // 12pt
                                                    }),
                                                ],
                                                bidirectional: _rtl(),
                                            }),
                                        ],
                                        width: {
                                            size: 25,
                                            type: WidthType.PERCENTAGE,
                                        },
                                    })
                                );
                            } else {
                                // Empty cell
                                rowCells.push(
                                    new TableCell({
                                        children: [new Paragraph('')],
                                        width: {
                                            size: 25,
                                            type: WidthType.PERCENTAGE,
                                        },
                                    })
                                );
                            }
                        }
                        
                        tableRows.push(new TableRow({ children: rowCells }));
                    }
                    
                    // Create the table with 16cm width
                    children.push(
                        new Table({
                            visuallyRightToLeft: _rtl(),
                            width: {
                                size: 9071, // 16cm in twips (16 * 567.05 ≈ 9071)
                                type: WidthType.DXA,
                            },
                            layout: "fixed", // Fixed table layout for consistent width
                            rows: tableRows,
                        })
                    );
                    
                    // Add spacing after table
                    children.push(new Paragraph({ spacing: { after: 200 } }));
                });

                // ============ ADDITIONAL INFORMATION (NEW PAGE) ============
                children.push(new Paragraph({
                    children: [
                        new PageBreak(),
                        new TextRun({
                            text: _t('expAdditionalInfo'),
                            bold: true,
                            size: 24, // 12pt
                        }),
                    ],
                    spacing: { after: 300 },
                    bidirectional: _rtl(),
                }));

                // Create 2-column tables for additional info
                const additionalInfoSections = [
                    {
                        heading1: document.getElementById('knowledgeHeading').textContent,
                        content1: document.getElementById('knowledgeInput').value.trim(),
                        heading2: document.getElementById('behaviorsHeading').textContent,
                        content2: document.getElementById('behaviorsInput').value.trim(),
                    },
                    {
                        heading1: document.getElementById('skillsHeading').textContent,
                        content1: document.getElementById('skillsInput').value.trim(),
                        heading2: '', // Empty for single column
                        content2: '',
                    },
                    {
                        heading1: document.getElementById('toolsHeading').textContent,
                        content1: document.getElementById('toolsInput').value.trim(),
                        heading2: document.getElementById('trendsHeading').textContent,
                        content2: document.getElementById('trendsInput').value.trim(),
                    },
                    {
                        heading1: document.getElementById('acronymsHeading').textContent,
                        content1: document.getElementById('acronymsInput').value.trim(),
                        heading2: document.getElementById('careerPathHeading').textContent,
                        content2: document.getElementById('careerPathInput').value.trim(),
                    },
                ];

                additionalInfoSections.forEach((section, index) => {
                    // Special handling for Acronyms (index 3, content1) - separate table with heading in first cell
                    if (index === 3 && section.content1) {
                        const row = new TableRow({
                            children: [
                                // First cell: Heading only with gray background
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: section.heading1,
                                                    bold: true,
                                                    size: 24, // 12pt
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    shading: {
                                        fill: "DCDCDC", // RGB(220,220,220) — matches the duty bar
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                    width: {
                                        size: 30,
                                        type: WidthType.PERCENTAGE,
                                    },
                                }),
                                // Second cell: Content only
                                new TableCell({
                                    children: section.content1.split('\n').filter(line => line.trim()).map(line => 
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: line.trim().replace(/^[•\-*]\s*/, '• '),
                                                    size: 24, // 12pt
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        })
                                    ),
                                    width: {
                                        size: 70,
                                        type: WidthType.PERCENTAGE,
                                    },
                                }),
                            ],
                        });
                        
                        children.push(
                            new Table({
                                visuallyRightToLeft: _rtl(),
                                width: {
                                    size: 9071, // 16cm in twips
                                    type: WidthType.DXA,
                                },
                                layout: "fixed",
                                rows: [row],
                            })
                        );
                        
                        children.push(new Paragraph({ spacing: { after: 200 } }));
                    }
                    // Regular format for all other sections (heading + content together)
                    else if (section.content1 || section.content2) {
                        const row = new TableRow({
                            children: [
                                // Left column
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: section.heading1,
                                                    bold: true,
                                                    size: 24, // 12pt
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                        ...section.content1.split('\n').filter(line => line.trim()).map(line => 
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: line.trim().replace(/^[•\-*]\s*/, '• '),
                                                        size: 24, // 12pt
                                                    }),
                                                ],
                                                bidirectional: _rtl(),
                                            })
                                        ),
                                    ],
                                    width: {
                                        size: 50,
                                        type: WidthType.PERCENTAGE,
                                    },
                                }),
                                // Right column
                                new TableCell({
                                    children: section.content2 ? [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: section.heading2,
                                                    bold: true,
                                                    size: 24, // 12pt
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                        ...section.content2.split('\n').filter(line => line.trim()).map(line => 
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: line.trim().replace(/^[•\-*]\s*/, '• '),
                                                        size: 24, // 12pt
                                                    }),
                                                ],
                                                bidirectional: _rtl(),
                                            })
                                        ),
                                    ] : [new Paragraph('')],
                                    width: {
                                        size: 50,
                                        type: WidthType.PERCENTAGE,
                                    },
                                }),
                            ],
                        });
                        
                        children.push(
                            new Table({
                                visuallyRightToLeft: _rtl(),
                                width: {
                                    size: 9071, // 16cm in twips
                                    type: WidthType.DXA,
                                },
                                layout: "fixed",
                                rows: [row],
                            })
                        );
                        
                        children.push(new Paragraph({ spacing: { after: 200 } }));
                    }
                });

                // Add custom sections
                const customSectionsContainer = document.getElementById('customSectionsContainer');
                const customSectionDivs = customSectionsContainer.querySelectorAll('.section-container');
                
                customSectionDivs.forEach(sectionDiv => {
                    const headingElement = sectionDiv.querySelector('h3');
                    const textareaElement = sectionDiv.querySelector('textarea');
                    
                    if (headingElement && textareaElement && textareaElement.value.trim()) {
                        const row = new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: headingElement.textContent,
                                                    bold: true,
                                                    size: 24, // 12pt
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                        ...textareaElement.value.split('\n').filter(line => line.trim()).map(line => 
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: line.trim().replace(/^[•\-*]\s*/, '• '),
                                                        size: 24, // 12pt
                                                    }),
                                                ],
                                                bidirectional: _rtl(),
                                            })
                                        ),
                                    ],
                                    columnSpan: 2,
                                    width: {
                                        size: 100,
                                        type: WidthType.PERCENTAGE,
                                    },
                                }),
                            ],
                        });
                        
                        children.push(
                            new Table({
                                visuallyRightToLeft: _rtl(),
                                width: {
                                    size: 9071, // 16cm in twips
                                    type: WidthType.DXA,
                                },
                                layout: "fixed", // Fixed table layout for consistent width
                                rows: [row],
                            })
                        );
                        
                        children.push(new Paragraph({ spacing: { after: 200 } }));
                    }
                });

                // ============ SKILLS LEVEL MATRIX EXPORT ============
                // Check if there's any meaningful data in Skills Level Matrix
                const hasSkillsLevelData = appState.skillsLevelData?.some(category =>
                    category.competencies.some(comp =>
                        Object.values(comp.levels).some(v => v === true)
                    )
                );

                if (hasSkillsLevelData) {
                    // Add Skills Level Matrix heading
                    children.push(new Paragraph({ spacing: { after: 200 } }));
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _t('expEmployability'),
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: _rtl(),
                    }));

                    // Create Skills Level Matrix table
                    appState.skillsLevelData.forEach(category => {
                        // Skip empty categories
                        if (category.category.trim() === '' && category.competencies.every(c => c.text.trim() === '')) {
                            return;
                        }

                        // Category header row
                        const headerRow = new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: category.category || `Category ${category.id}`,
                                                    bold: true,
                                                    size: 24,
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    columnSpan: 5,
                                    shading: {
                                        fill: "DCDCDC",
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                }),
                            ],
                        });

                        // Column headers row
                        const columnHeaderRow = new TableRow({
                            children: [
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: _t('expCompetency'),
                                                    bold: true,
                                                    size: 22,
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    width: { size: 40, type: WidthType.PERCENTAGE },
                                    shading: {
                                        fill: "DCDCDC",
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: _t('expCraftsman'),
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    width: { size: 15, type: WidthType.PERCENTAGE },
                                    shading: {
                                        fill: "DCDCDC",
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: _t('expSkilled'),
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    width: { size: 15, type: WidthType.PERCENTAGE },
                                    shading: {
                                        fill: "DCDCDC",
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: _t('expSemiSkilled'),
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    width: { size: 15, type: WidthType.PERCENTAGE },
                                    shading: {
                                        fill: "DCDCDC",
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                }),
                                new TableCell({
                                    children: [
                                        new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: _t('expFoundation'),
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: _rtl(),
                                        }),
                                    ],
                                    width: { size: 15, type: WidthType.PERCENTAGE },
                                    shading: {
                                        fill: "DCDCDC",
                                        type: ShadingType.CLEAR,
                                        color: "auto",
                                    },
                                }),
                            ],
                        });

                        // Competency rows
                        const competencyRows = category.competencies
                            .filter(comp => comp.text.trim() !== '')
                            .map(competency => {
                                return new TableRow({
                                    children: [
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({
                                                            text: `${competency.id}. ${competency.text}`,
                                                            size: 22,
                                                        }),
                                                    ],
                                                    bidirectional: _rtl(),
                                                }),
                                            ],
                                            width: { size: 40, type: WidthType.PERCENTAGE },
                                        }),
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({
                                                            text: competency.levels.craftsman ? '✓' : '',
                                                            size: 22,
                                                        }),
                                                    ],
                                                    alignment: AlignmentType.CENTER,
                                                    bidirectional: _rtl(),
                                                }),
                                            ],
                                            width: { size: 15, type: WidthType.PERCENTAGE },
                                        }),
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({
                                                            text: competency.levels.skilled ? '✓' : '',
                                                            size: 22,
                                                        }),
                                                    ],
                                                    alignment: AlignmentType.CENTER,
                                                    bidirectional: _rtl(),
                                                }),
                                            ],
                                            width: { size: 15, type: WidthType.PERCENTAGE },
                                        }),
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({
                                                            text: competency.levels.semiSkilled ? '✓' : '',
                                                            size: 22,
                                                        }),
                                                    ],
                                                    alignment: AlignmentType.CENTER,
                                                    bidirectional: _rtl(),
                                                }),
                                            ],
                                            width: { size: 15, type: WidthType.PERCENTAGE },
                                        }),
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({
                                                            text: competency.levels.foundation ? '✓' : '',
                                                            size: 22,
                                                        }),
                                                    ],
                                                    alignment: AlignmentType.CENTER,
                                                    bidirectional: _rtl(),
                                                }),
                                            ],
                                            width: { size: 15, type: WidthType.PERCENTAGE },
                                        }),
                                    ],
                                });
                            });

                        // Add table for this category
                        children.push(
                            new Table({
                                visuallyRightToLeft: _rtl(),
                                width: {
                                    size: 9071, // 16cm in twips
                                    type: WidthType.DXA,
                                },
                                layout: "fixed",
                                rows: [headerRow, columnHeaderRow, ...competencyRows],
                            })
                        );
                        
                        children.push(new Paragraph({ spacing: { after: 200 } }));
                    });
                }

                // ============ TASK VERIFICATION APPENDIX (if mode = 'appendix') ============
                if (appState.tvExportMode === 'appendix' && appState.collectionMode === 'workshop') {
                    const validResults = Object.keys(appState.workshopResults).filter(key => 
                        appState.workshopResults[key] && appState.workshopResults[key].valid
                    );
                    
                    if (validResults.length > 0) {
                        // Page break before appendix
                        children.push(new Paragraph({
                            children: [new PageBreak()],
                        }));
                        
                        // Appendix title
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expTaskVerifAppendix'),
                                    bold: true,
                                    size: 32, // 16pt
                                }),
                            ],
                            spacing: { after: 300 },
                            bidirectional: _rtl(), // direction of the Task Verification section
                        }));
                        
                        // Methodology Summary heading
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expMethodologySummary'),
                                    bold: true,
                                    size: 28, // 14pt
                                }),
                            ],
                            spacing: { after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Methodology details
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _tf('expCollectionMode', { v: _t(appState.collectionMode === 'workshop' ? 'modeWorkshop' : 'expIndividualSurvey') }),
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 100 },
                            bidirectional: _rtl(),
                        }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _tf('expParticipants', { v: appState.workshopParticipants }),
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 100 },
                            bidirectional: _rtl(),
                        }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _tf('expWorkflowMode', { v: _t(appState.workflowMode === 'standard' ? 'modeStandard' : 'modeExtended') }),
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 100 },
                            bidirectional: _rtl(),
                        }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _tf('expPriorityFormula', { v: _t(appState.priorityFormula === 'if' ? 'formulaIF' : 'formulaIFD') }),
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 300 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Priority Rankings heading
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expPriorityRankings'),
                                    bold: true,
                                    size: 28,
                                }),
                            ],
                            spacing: { after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Get and sort results
                        const sortedResults = [];
                        validResults.forEach(taskKey => {
                            const result = appState.workshopResults[taskKey];
                            
                            // Use stored duty and task titles (with backward compatibility)
                            let dutyText = result.dutyTitle;
                            let taskText = result.taskTitle;
                            
                            // Backward compatibility: if not stored, look up from DOM
                            if (!dutyText || !taskText) {
                                const taskParts = taskKey.split('_task_');
                                const dutyId = taskParts[0];
                                
                                if (!dutyText) {
                                    const dutyInput = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`);
                                    dutyText = dutyInput ? dutyInput.value.trim() : 'Unassigned';
                                }
                                
                                if (!taskText) {
                                    const taskInput = document.querySelector(`input[data-task-id="${taskKey}"], textarea[data-task-id="${taskKey}"]`);
                                    taskText = taskInput ? taskInput.value.trim() : 'Unassigned';
                                }
                            }
                            
                            sortedResults.push({
                                duty: dutyText,
                                task: taskText,
                                meanI: result.meanImportance,
                                meanF: result.meanFrequency,
                                meanD: result.meanDifficulty,
                                priority: result.priorityIndex
                            });
                        });
                        
                        sortedResults.sort((a, b) => b.priority - a.priority);
                        
                        // Create table
                        const tableRows = [];
                        
                        // Header row
                        tableRows.push(new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: _t('expRank'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: _t('expDutyLabel'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: _t('expTaskLabel'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: _t('expMeanI'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: _t('expMeanF'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: _t('expMeanD'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: _t('expPriority'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                            ],
                        }));
                        
                        // Data rows
                        sortedResults.forEach((row, index) => {
                            tableRows.push(new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `#${index + 1}` })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: row.duty, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: row.task, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanI !== null ? row.meanI.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanF !== null ? row.meanF.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanD !== null ? row.meanD.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.priority !== null ? row.priority.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                ],
                            }));
                        });
                        
                        children.push(new Table({
                            visuallyRightToLeft: _rtl(),
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: tableRows,
                        }));
                        
                        // Duty-Level Summary section
                        children.push(new Paragraph({ spacing: { after: 400 } }));
                        
                        children.push(new Paragraph({
                            children: [new TextRun({ text: _t('expDutyLevelSummary'), bold: true, size: 28 })],
                            spacing: { after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        children.push(new Paragraph({
                            children: [new TextRun({ text: _tf('expTrainingLoadMethod', { v: _t(appState.trainingLoadMethod === 'advanced' ? 'expAdvancedMethod' : 'expSimpleMethod') }), size: 20, italics: true })],
                            spacing: { after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Aggregate duty-level data
                        const appendixDutyMap = {};
                        Object.keys(appState.workshopResults).forEach(taskKey => {
                            const result = appState.workshopResults[taskKey];
                            if (result && result.valid) {
                                let dutyId = result.dutyId || taskKey.split('_task_')[0];
                                let dutyTitle = result.dutyTitle;
                                
                                if (!dutyTitle) {
                                    const dutyInput = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`);
                                    dutyTitle = dutyInput ? dutyInput.value.trim() : 'Unassigned';
                                }
                                
                                if (!appendixDutyMap[dutyId]) {
                                    appendixDutyMap[dutyId] = { dutyTitle: dutyTitle, validTasks: 0, prioritySum: 0, tasks: [] };
                                }
                                
                                appendixDutyMap[dutyId].validTasks++;
                                appendixDutyMap[dutyId].prioritySum += result.priorityIndex;
                                appendixDutyMap[dutyId].tasks.push({ priorityIndex: result.priorityIndex, meanDifficulty: result.meanDifficulty });
                            }
                        });
                        
                        const appendixDutyResults = [];
                        Object.keys(appendixDutyMap).forEach(dutyId => {
                            const duty = appendixDutyMap[dutyId];
                            const avgPriority = duty.prioritySum / duty.validTasks;
                            let trainingLoad = 0;
                            if (appState.trainingLoadMethod === 'advanced') {
                                trainingLoad = duty.tasks.reduce((sum, t) => sum + (t.priorityIndex * t.meanDifficulty), 0);
                            } else {
                                trainingLoad = avgPriority * duty.validTasks;
                            }
                            appendixDutyResults.push({ dutyTitle: duty.dutyTitle, validTasks: duty.validTasks, avgPriority: avgPriority, trainingLoad: trainingLoad });
                        });
                        
                        appendixDutyResults.sort((a, b) => b.avgPriority - a.avgPriority);
                        
                        // Duty table
                        const dutyTableRows = [
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expDutyTitle'), bold: true })], alignment: _start(AlignmentType), bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expTasks'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expAvgPriority'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: _t('expTrainingLoad'), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })], shading: { fill: 'DCDCDC' } }),
                                ],
                            })
                        ];
                        
                        appendixDutyResults.forEach(duty => {
                            dutyTableRows.push(new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: duty.dutyTitle, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.validTasks.toString() })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.avgPriority.toFixed(2) })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.trainingLoad.toFixed(2), bold: true })], alignment: AlignmentType.CENTER, bidirectional: _rtl() })] }),
                                ],
                            }));
                        });
                        
                        children.push(new Table({
                            visuallyRightToLeft: _rtl(),
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: dutyTableRows,
                        }));
                        
                        // Notes section
                        children.push(new Paragraph({ spacing: { after: 300 } }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expNotes'),
                                    bold: true,
                                    size: 24,
                                }),
                            ],
                            spacing: { after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        const notes = [
                            'Weighted Mean = Σ(value × count) ÷ total responses',
                            'Priority Index calculated using selected formula',
                            'Higher priority values indicate greater training importance',
                            'Results based on DACUM methodology'
                        ];
                        
                        notes.forEach(note => {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `• ${note}`,
                                        size: 20,
                                    }),
                                ],
                                spacing: { after: 100 },
                                bidirectional: _rtl(),
                            }));
                        });
                    }
                }

                // ============ VERIFIED LIVE WORKSHOP RESULTS APPENDIX ============
                if (appState.tvExportMode === 'appendix' && hasVerifiedResults) {
                    // Page break before verified results appendix
                    children.push(new Paragraph({ children: [new PageBreak()] }));
                    
                    // Appendix title
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _t('expPostVoteResults'),
                                bold: true,
                                size: 32,
                            }),
                        ],
                        spacing: { after: 300 },
                        bidirectional: _rtl(),
                    }));
                    
                    // Metadata
                    children.push(new Paragraph({
                        children: [new TextRun({ text: _tf('expOccupation', { v: appState.lwFinalizedData.occupation }), size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: _rtl(),
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: _tf('expJobTitle', { v: appState.lwFinalizedData.jobTitle }), size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: _rtl(),
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: _tf('expDate', { v: _today() }), size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: _rtl(),
                    }));
                    const vFormula = appState.lwFinalizedData.appState.priorityFormula || 'if';
                    const vFormulaText = vFormula === 'ifd' ? 'Importance × Frequency × Difficulty' : 'Importance × Frequency';
                    children.push(new Paragraph({
                        children: [new TextRun({ text: _tf('expPriorityFormula', { v: vFormulaText }), size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: _rtl(),
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: _tf('expTotalParticipants', { v: appState.lwAggregatedResults.totalVotes }), size: 22 })],
                        spacing: { after: 300 },
                        bidirectional: _rtl(),
                    }));
                    
                    // Collect all verified tasks with metrics
                    const verifiedTasks = [];
                    Object.keys(appState.lwFinalizedData.duties).forEach(dutyId => {
                        const duty = appState.lwFinalizedData.duties[dutyId];
                        duty.tasks.forEach(task => {
                            if (task.priorityIndex !== undefined) {
                                verifiedTasks.push({
                                    dutyTitle: duty.title,
                                    taskText: task.text,
                                    meanImportance: task.meanImportance,
                                    meanFrequency: task.meanFrequency,
                                    meanDifficulty: task.meanDifficulty,
                                    priorityIndex: task.priorityIndex,
                                    rank: task.rank
                                });
                            }
                        });
                    });
                    
                    verifiedTasks.sort((a, b) => a.rank - b.rank);
                    
                    // Create table
                    const verifiedTableRows = [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ text: _t('expRank'), bold: true, bidirectional: _rtl() })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: _t('expDutyLabel'), bold: true, bidirectional: _rtl() })], width: { size: 22, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: _t('expTaskLabel'), bold: true, bidirectional: _rtl() })], width: { size: 35, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: _t('expInitialI'), bold: true, bidirectional: _rtl() })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: _t('expInitialF'), bold: true, bidirectional: _rtl() })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: _t('expInitialD'), bold: true, bidirectional: _rtl() })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: _t('expPI'), bold: true, bidirectional: _rtl() })], width: { size: 11, type: WidthType.PERCENTAGE } })
                            ]
                        })
                    ];
                    
                    verifiedTasks.forEach(task => {
                        verifiedTableRows.push(
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: String(task.rank), bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.dutyTitle, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.taskText, bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.meanImportance.toFixed(2), bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.meanFrequency.toFixed(2), bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.meanDifficulty.toFixed(2), bidirectional: _rtl() })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.priorityIndex.toFixed(2), bidirectional: _rtl() })] })
                                ]
                            })
                        );
                    });
                    
                    children.push(new Table({
                        visuallyRightToLeft: _rtl(),
                        rows: verifiedTableRows,
                        width: { size: 100, type: WidthType.PERCENTAGE }
                    }));
                }

                // ============ COMPETENCY CLUSTERS SECTION ============
                if (appState.clusteringData.clusters && appState.clusteringData.clusters.length > 0) {
                    children.push(new Paragraph({ children: [new PageBreak()], bidirectional: _rtl() }));
                    
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _t('expClusters'),
                                bold: true,
                                size: 32, // 16pt
                            }),
                        ],
                        spacing: { before: 400, after: 400 },
                        alignment: AlignmentType.CENTER,
                        bidirectional: _rtl(),
                    }));
                    
                    appState.clusteringData.clusters.forEach((cluster, clusterIndex) => {
                        const clusterNumber = clusterIndex + 1;
                        
                        // Cluster header
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _tf('expCompetencyN', { n: clusterNumber, name: cluster.name }),
                                    bold: true,
                                    size: 28, // 14pt
                                }),
                            ],
                            spacing: { before: 300, after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Range section
                        if (cluster.range && cluster.range.trim()) {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: _t('expRangeLabel'),
                                        bold: true,
                                        size: 24, // 12pt
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                                bidirectional: _rtl(),
                            }));
                            
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: cluster.range,
                                        size: 22, // 11pt
                                    }),
                                ],
                                spacing: { after: 200 },
                                indent: { left: 720 },
                                bidirectional: _rtl(),
                            }));
                        }
                        
                        // Related Tasks section
                        if (cluster.tasks && cluster.tasks.length > 0) {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: _t('expRelatedTasks'),
                                        bold: true,
                                        size: 24, // 12pt
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                                bidirectional: _rtl(),
                            }));
                            
                            cluster.tasks.forEach(task => {
                                const taskCode = getTaskCode(task.id);
                                children.push(new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `- ${taskCode}: ${task.text}`,
                                            size: 22, // 11pt
                                        }),
                                    ],
                                    spacing: { after: 100 },
                                    indent: { left: 720 },
                                    bidirectional: _rtl(),
                                }));
                            });
                        }
                        
                        // Performance Criteria section
                        if (cluster.performanceCriteria && cluster.performanceCriteria.length > 0) {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: _t('expPCLabel'),
                                        bold: true,
                                        size: 24, // 12pt
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                                bidirectional: _rtl(),
                            }));
                            
                            cluster.performanceCriteria.forEach((criterion, idx) => {
                                children.push(new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `${clusterNumber}-${idx + 1} ${criterion}`,
                                            size: 22, // 11pt
                                        }),
                                    ],
                                    spacing: { after: 100 },
                                    indent: { left: 720 },
                                    bidirectional: _rtl(),
                                }));
                            });
                        }
                    });
                }

                // ============ LEARNING OUTCOMES SECTION ============
                if (appState.learningOutcomesData.outcomes && appState.learningOutcomesData.outcomes.length > 0) {
                    children.push(new Paragraph({ children: [new PageBreak()], bidirectional: _rtl() }));
                    
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _t('expLearningOutcomes'),
                                bold: true,
                                size: 32, // 16pt
                            }),
                        ],
                        spacing: { before: 400, after: 400 },
                        alignment: AlignmentType.CENTER,
                        bidirectional: _rtl(),
                    }));
                    
                    // Group LOs by cluster
                    const losByCluster = {};
                    appState.learningOutcomesData.outcomes.forEach(lo => {
                        lo.linkedCriteria.forEach(pc => {
                            if (!losByCluster[pc.clusterNumber]) {
                                losByCluster[pc.clusterNumber] = [];
                            }
                            if (!losByCluster[pc.clusterNumber].includes(lo)) {
                                losByCluster[pc.clusterNumber].push(lo);
                            }
                        });
                    });
                    
                    // Sort cluster numbers
                    const clusterNumbers = Object.keys(losByCluster).sort((a, b) => parseInt(a) - parseInt(b));
                    
                    clusterNumbers.forEach(clusterNum => {
                        const clusterIndex = parseInt(clusterNum) - 1;
                        const cluster = appState.clusteringData.clusters[clusterIndex];
                        const los = losByCluster[clusterNum];
                        
                        // Cluster header
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: cluster.name,
                                    bold: true,
                                    size: 28, // 14pt
                                }),
                            ],
                            spacing: { before: 300, after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Learning Outcomes for this cluster
                        los.forEach(lo => {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `${lo.number}:`,
                                        bold: true,
                                        size: 24, // 12pt
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                                bidirectional: _rtl(),
                            }));
                            
                            if (lo.statement && lo.statement.trim()) {
                                children.push(new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: lo.statement,
                                            size: 22, // 11pt
                                        }),
                                    ],
                                    spacing: { after: 100 },
                                    indent: { left: 720 },
                                    bidirectional: _rtl(),
                                }));
                            }
                            
                            // Mapped Performance Criteria
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: _t('expMappedPC'),
                                        italic: true,
                                        size: 20, // 10pt
                                    }),
                                ],
                                spacing: { before: 100, after: 50 },
                                indent: { left: 720 },
                                bidirectional: _rtl(),
                            }));
                            
                            lo.linkedCriteria.forEach(pc => {
                                children.push(new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `- ${pc.id}: ${pc.text}`,
                                            size: 18, // 9pt
                                        }),
                                    ],
                                    spacing: { after: 50 },
                                    indent: { left: 1440 },
                                    bidirectional: _rtl(),
                                }));
                            });
                        });
                    });
                }

                // ============ MODULE MAPPING SECTION ============
                if (appState.moduleMappingData.modules && appState.moduleMappingData.modules.length > 0) {
                    children.push(new Paragraph({ children: [new PageBreak()], bidirectional: _rtl() }));
                    
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: _t('expModuleMapping'),
                                bold: true,
                                size: 32, // 16pt
                            }),
                        ],
                        spacing: { before: 400, after: 400 },
                        alignment: AlignmentType.CENTER,
                        bidirectional: _rtl(),
                    }));
                    
                    appState.moduleMappingData.modules.forEach(module => {
                        // Module title
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: module.title,
                                    bold: true,
                                    size: 28, // 14pt
                                }),
                            ],
                            spacing: { before: 300, after: 200 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Learning Outcomes header
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: _t('expLOsLabel'),
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: _rtl(),
                        }));
                        
                        // Learning Outcomes in this module
                        module.learningOutcomes.forEach(lo => {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `${lo.number}:`,
                                        bold: true,
                                        size: 22, // 11pt
                                    }),
                                ],
                                spacing: { before: 150, after: 50 },
                                indent: { left: 720 },
                                bidirectional: _rtl(),
                            }));
                            
                            if (lo.statement && lo.statement.trim()) {
                                children.push(new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: lo.statement,
                                            size: 20, // 10pt
                                        }),
                                    ],
                                    spacing: { after: 50 },
                                    indent: { left: 1440 },
                                    bidirectional: _rtl(),
                                }));
                            }
                            
                            // Referenced Performance Criteria
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: _t('expReferencedPC'),
                                        italic: true,
                                        size: 18, // 9pt
                                    }),
                                ],
                                spacing: { before: 50, after: 30 },
                                indent: { left: 1440 },
                                bidirectional: _rtl(),
                            }));
                            
                            lo.linkedCriteria.forEach(pc => {
                                children.push(new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: `- ${pc.id}: ${pc.text}`,
                                            size: 16, // 8pt
                                        }),
                                    ],
                                    spacing: { after: 30 },
                                    indent: { left: 2160 },
                                    bidirectional: _rtl(),
                                }));
                            });
                        });
                    });
                }

                // Create document
                const doc = new Document({
                    styles: {
                        default: {
                            document: { run: { font: _font() } },
                        },
                    },
                    sections: [{
                        properties: {
                            /* bidi on the SECTION makes Word treat the whole
                               section as right-to-left: table columns run
                               right-to-left, and list markers and numbering
                               sit on the correct edge. Paragraph-level
                               bidirectional alone does not do either. */
                            bidi: _rtl(),
                            page: {
                                margin: {
                                    top: 1440,
                                    right: 1440,
                                    bottom: 1440,
                                    left: 1440,
                                },
                            },
                        },
                        children: children,
                    }],
                });

                // Generate and download
                const blob = await Packer.toBlob(doc);
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                /* Was /[^a-z0-9]/gi, which erased every Arabic character —
                   the file arrived as "______.docx". */
                const _jobPart = (jobTitle && jobTitle.trim()) ? ` ${jobTitle}` : '';
                link.download = _safeFilename(occupationTitle + _jobPart, '_DACUM_Chart.docx');
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                
                showStatus(_t('msgWordExported') + ' ✓', 'success');

            } catch (error) {
                console.error('Error generating Word document:', error);
                showStatus(_tf('msgWordError', { msg: error.message }), 'error');
            }
        }
