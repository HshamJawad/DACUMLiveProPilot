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
import { getTaskCode } from './codes.js';
import { buildVerificationDataset, getVerificationCoverage } from './exports_shared.js';

export async function exportTaskVerificationWord() {
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
                    showStatus('Error: Word export library not loaded. Please refresh the page.', 'error');
                    return;
                }

                const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, ShadingType, Packer } = window.docx;

                showStatus('Generating Task Verification Word document...', 'success');

                const children = [];
                
                const occupationTitleInput = document.getElementById('occupationTitle');
                const occupationTitle = occupationTitleInput ? occupationTitleInput.value : 'Unknown Occupation';
                
                // Title
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Task Verification & Training Priority Analysis',
                            bold: true,
                            size: 32,
                        }),
                    ],
                    spacing: { after: 300 },
                    bidirectional: false, // Force LTR
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `Occupation: ${occupationTitle}`,
                            bold: true,
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: false, // Force LTR
                }));
                
                const today = new Date().toLocaleDateString();
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `Date of Analysis: ${today}`,
                            size: 24,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: false, // Force LTR
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `This Task Verification is based on the DACUM Chart for ${occupationTitle}.`,
                            italics: true,
                            size: 20,
                        }),
                    ],
                    spacing: { after: 400 },
                    bidirectional: false, // Force LTR
                }));
                
                // Methodology Summary
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Methodology Summary',
                            bold: true,
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: false, // Force LTR
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `Data Collection Mode: ${appState.collectionMode === 'workshop' ? 'Workshop (Facilitated)' : 'Individual/Survey'}`,
                            size: 22,
                        }),
                    ],
                    spacing: { after: 100 },
                    bidirectional: false, // Force LTR
                }));
                
                // Participant count only exists when there was a panel.
                // Printing appState.workshopParticipants in Individual /
                // Survey mode would state an evidence base that does not
                // exist for these numbers.
                if (appState.collectionMode === 'workshop') {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `Number of Participants: ${appState.workshopParticipants}`,
                                size: 22,
                            }),
                        ],
                        spacing: { after: 100 },
                        bidirectional: false, // Force LTR
                    }));
                }

                // Coverage. Bold and red when partial, so a work-in-progress
                // export can never be mistaken for a completed verification.
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `Coverage: ${tvCoverage.label}`,
                            size: 22,
                            bold: !tvCoverage.complete,
                            color: tvCoverage.complete ? '000000' : 'B91C1C',
                        }),
                    ],
                    spacing: { after: 100 },
                    bidirectional: false, // Force LTR
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `Workflow Mode: ${appState.workflowMode === 'standard' ? 'Standard (DACUM)' : 'Extended (DACUM)'}`,
                            size: 22,
                        }),
                    ],
                    spacing: { after: 100 },
                    bidirectional: false, // Force LTR
                }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `Priority Formula: ${appState.priorityFormula === 'if' ? 'Importance × Frequency' : 'Importance × Frequency × Difficulty'}`,
                            size: 22,
                        }),
                    ],
                    spacing: { after: 400 },
                    bidirectional: false, // Force LTR
                }));
                
                // Priority Rankings
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Priority Rankings',
                            bold: true,
                            size: 28,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: false, // Force LTR
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
                            children: [new Paragraph({ children: [new TextRun({ text: 'Rank', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Duty', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Task', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Mean I', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Mean F', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Mean D', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Priority', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                        }),
                    ],
                }));
                
                // Data rows
                sortedResults.forEach((row, index) => {
                    tableRows.push(new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `#${index + 1}` })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ text: row.duty, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ text: row.task, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanI !== null ? row.meanI.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanF !== null ? row.meanF.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanD !== null ? row.meanD.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.priority !== null ? row.priority.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                        ],
                    }));
                });
                
                children.push(new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: tableRows,
                }));
                
                // Duty-Level Summary section
                children.push(new Paragraph({ spacing: { after: 400 } }));
                
                children.push(new Paragraph({
                    children: [new TextRun({ text: 'Duty-Level Summary', bold: true, size: 28 })],
                    spacing: { after: 200 },
                    bidirectional: false,
                }));
                
                children.push(new Paragraph({
                    children: [new TextRun({ text: `Training Load Method: ${appState.trainingLoadMethod === 'advanced' ? 'Advanced (Σ Priority × Difficulty)' : 'Simple (Avg Priority × Tasks)'}`, size: 20, italics: true })],
                    spacing: { after: 200 },
                    bidirectional: false,
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
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Duty Title', bold: true })], alignment: AlignmentType.LEFT, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Tasks', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Avg Priority', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Training Load', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                        ],
                    })
                ];
                
                dutyResults.forEach(duty => {
                    dutyTableRows.push(new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: duty.dutyTitle, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.validTasks.toString() })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.avgPriority.toFixed(2) })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.trainingLoad.toFixed(2), bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                        ],
                    }));
                });
                
                children.push(new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    rows: dutyTableRows,
                }));
                
                // Notes section
                children.push(new Paragraph({ spacing: { after: 400 } }));
                
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Notes & Methodology',
                            bold: true,
                            size: 24,
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: false, // Force LTR
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
                        bidirectional: false, // Force LTR
                    }));
                });
                
                // Create document
                const doc = new Document({
                    sections: [{
                        properties: {
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
                link.download = `${occupationTitle.replace(/[^a-z0-9]/gi, '_')}_Task_Verification.docx`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                
                showStatus('Task Verification Word document exported successfully! ✓', 'success');

            } catch (error) {
                console.error('Error generating Task Verification Word document:', error);
                showStatus('Error generating Task Verification Word document: ' + error.message, 'error');
            }
        }

export async function exportToWord() {
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
                    showStatus('Error: Word export library not loaded. Please refresh the page.', 'error');
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
                    showStatus('Occupation Title is required for export.', 'error');
                    return;
                }

                showStatus('Generating Word document...', 'success');

                const children = [];

                // ============ TITLE PAGE ============
                children.push(new Paragraph({
                    children: [
                        new TextRun({
                            text: `Occupation Title: ${occupationTitle}`,
                            bold: true,
                            size: 28, // 14pt
                        }),
                    ],
                    spacing: { after: 200 },
                    bidirectional: false,
                }));

                // Scope of Work / Occupational Definition (optional)
                const scopeOfWorkEl    = document.getElementById('scopeOfWork');
                const scopeOfWorkValue = (scopeOfWorkEl ? scopeOfWorkEl.value : '').trim();
                if (scopeOfWorkValue) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: 'Scope of Work / Occupational Definition:',
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { before: 80, after: 80 },
                        bidirectional: false,
                    }));
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: scopeOfWorkValue,
                                size: 22, // 11pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: false,
                    }));
                }

                // Job Title is optional — skip the paragraph entirely when empty
                if (jobTitle && jobTitle.trim()) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `Job Title: ${jobTitle}`,
                                bold: true,
                                size: 28, // 14pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: false,
                    }));
                }

                // Add DACUM Date if exists
                if (dacumDate) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `DACUM Date: ${dacumDate}`,
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: false,
                    }));
                }
                
                // Add Venue if exists
                const venueValue = document.getElementById('venue')?.value;
                if (venueValue) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `Venue: ${venueValue}`,
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: false,
                    }));
                }

                // Add Produced For if exists
                if (producedFor) {
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `Produced For: ${producedFor}`,
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: false,
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
                                text: `Produced By: ${producedBy}`,
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: false,
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
                                    text: 'Facilitators',
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: false,
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
                                                bidirectional: false,
                                            }),
                                        ],
                                    }),
                                ],
                            })
                        );
                        
                        children.push(new Table({
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
                                    text: 'Observers',
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: false,
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
                                                bidirectional: false,
                                            }),
                                        ],
                                    }),
                                ],
                            })
                        );
                        
                        children.push(new Table({
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
                                    text: 'Panel Members',
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: false,
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
                                                bidirectional: false,
                                            }),
                                        ],
                                    }),
                                ],
                            })
                        );
                        
                        children.push(new Table({
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
                            text: 'Duties and Tasks',
                            bold: true,
                            size: 28, // 14pt
                        }),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 300 },
                    bidirectional: false,
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
                    const dutyLetter = String.fromCharCode(65 + dutyIndex); // A, B, C...
                    const dutyLabel = `DUTY ${dutyLetter}: ${dutyData.duty}`;
                    
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
                                            bidirectional: false,
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
                                const taskLabel = `Task ${dutyLetter}${taskIndex + 1}`;
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
                                                bidirectional: false,
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
                            text: 'Additional Information',
                            bold: true,
                            size: 24, // 12pt
                        }),
                    ],
                    spacing: { after: 300 },
                    bidirectional: false,
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
                                            bidirectional: false,
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
                                            bidirectional: false,
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
                                            bidirectional: false,
                                        }),
                                        ...section.content1.split('\n').filter(line => line.trim()).map(line => 
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: line.trim().replace(/^[•\-*]\s*/, '• '),
                                                        size: 24, // 12pt
                                                    }),
                                                ],
                                                bidirectional: false,
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
                                            bidirectional: false,
                                        }),
                                        ...section.content2.split('\n').filter(line => line.trim()).map(line => 
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: line.trim().replace(/^[•\-*]\s*/, '• '),
                                                        size: 24, // 12pt
                                                    }),
                                                ],
                                                bidirectional: false,
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
                                            bidirectional: false,
                                        }),
                                        ...textareaElement.value.split('\n').filter(line => line.trim()).map(line => 
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: line.trim().replace(/^[•\-*]\s*/, '• '),
                                                        size: 24, // 12pt
                                                    }),
                                                ],
                                                bidirectional: false,
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
                                text: 'Employability Competencies by Occupational Level',
                                bold: true,
                                size: 24, // 12pt
                            }),
                        ],
                        spacing: { after: 200 },
                        bidirectional: false,
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
                                            bidirectional: false,
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
                                                    text: 'Competency',
                                                    bold: true,
                                                    size: 22,
                                                }),
                                            ],
                                            bidirectional: false,
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
                                                    text: 'Craftsman/\nSupervisor',
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: false,
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
                                                    text: 'Skilled',
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: false,
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
                                                    text: 'Semi-skilled',
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: false,
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
                                                    text: 'Foundation\nskills',
                                                    bold: true,
                                                    size: 20,
                                                }),
                                            ],
                                            bidirectional: false,
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
                                                    bidirectional: false,
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
                                                    bidirectional: false,
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
                                                    bidirectional: false,
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
                                                    bidirectional: false,
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
                                                    bidirectional: false,
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
                                    text: 'Task Verification & Training Priority Analysis (Appendix)',
                                    bold: true,
                                    size: 32, // 16pt
                                }),
                            ],
                            spacing: { after: 300 },
                            bidirectional: false, // Force LTR for Task Verification section
                        }));
                        
                        // Methodology Summary heading
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: 'Methodology Summary',
                                    bold: true,
                                    size: 28, // 14pt
                                }),
                            ],
                            spacing: { after: 200 },
                            bidirectional: false, // Force LTR
                        }));
                        
                        // Methodology details
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Data Collection Mode: ${appState.collectionMode === 'workshop' ? 'Workshop (Facilitated)' : 'Individual/Survey'}`,
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 100 },
                            bidirectional: false, // Force LTR
                        }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Number of Participants: ${appState.workshopParticipants}`,
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 100 },
                            bidirectional: false, // Force LTR
                        }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Workflow Mode: ${appState.workflowMode === 'standard' ? 'Standard (DACUM)' : 'Extended (DACUM)'}`,
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 100 },
                            bidirectional: false, // Force LTR
                        }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Priority Formula: ${appState.priorityFormula === 'if' ? 'Importance × Frequency' : 'Importance × Frequency × Difficulty'}`,
                                    size: 22,
                                }),
                            ],
                            spacing: { after: 300 },
                            bidirectional: false, // Force LTR
                        }));
                        
                        // Priority Rankings heading
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: 'Priority Rankings',
                                    bold: true,
                                    size: 28,
                                }),
                            ],
                            spacing: { after: 200 },
                            bidirectional: false, // Force LTR
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
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Rank', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Duty', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Task', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Mean I', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Mean F', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Mean D', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Priority', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.CLEAR, color: 'auto' },
                                }),
                            ],
                        }));
                        
                        // Data rows
                        sortedResults.forEach((row, index) => {
                            tableRows.push(new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `#${index + 1}` })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: row.duty, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: row.task, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanI !== null ? row.meanI.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanF !== null ? row.meanF.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.meanD !== null ? row.meanD.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: row.priority !== null ? row.priority.toFixed(2) : 'N/A' })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                ],
                            }));
                        });
                        
                        children.push(new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: tableRows,
                        }));
                        
                        // Duty-Level Summary section
                        children.push(new Paragraph({ spacing: { after: 400 } }));
                        
                        children.push(new Paragraph({
                            children: [new TextRun({ text: 'Duty-Level Summary', bold: true, size: 28 })],
                            spacing: { after: 200 },
                            bidirectional: false,
                        }));
                        
                        children.push(new Paragraph({
                            children: [new TextRun({ text: `Training Load Method: ${appState.trainingLoadMethod === 'advanced' ? 'Advanced (Σ Priority × Difficulty)' : 'Simple (Avg Priority × Tasks)'}`, size: 20, italics: true })],
                            spacing: { after: 200 },
                            bidirectional: false,
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
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Duty Title', bold: true })], alignment: AlignmentType.LEFT, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Tasks', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Avg Priority', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Training Load', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })], shading: { fill: 'DCDCDC' } }),
                                ],
                            })
                        ];
                        
                        appendixDutyResults.forEach(duty => {
                            dutyTableRows.push(new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: duty.dutyTitle, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.validTasks.toString() })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.avgPriority.toFixed(2) })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: duty.trainingLoad.toFixed(2), bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })] }),
                                ],
                            }));
                        });
                        
                        children.push(new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            rows: dutyTableRows,
                        }));
                        
                        // Notes section
                        children.push(new Paragraph({ spacing: { after: 300 } }));
                        
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: 'Notes',
                                    bold: true,
                                    size: 24,
                                }),
                            ],
                            spacing: { after: 200 },
                            bidirectional: false, // Force LTR
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
                                bidirectional: false, // Force LTR
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
                                text: 'DACUM Live Pro - Verified (Post-Vote) Results (Appendix)',
                                bold: true,
                                size: 32,
                            }),
                        ],
                        spacing: { after: 300 },
                        bidirectional: false,
                    }));
                    
                    // Metadata
                    children.push(new Paragraph({
                        children: [new TextRun({ text: `Occupation: ${appState.lwFinalizedData.occupation}`, size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: false,
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: `Job Title: ${appState.lwFinalizedData.jobTitle}`, size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: false,
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: `Date: ${new Date().toLocaleDateString()}`, size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: false,
                    }));
                    const vFormula = appState.lwFinalizedData.appState.priorityFormula || 'if';
                    const vFormulaText = vFormula === 'ifd' ? 'Importance × Frequency × Difficulty' : 'Importance × Frequency';
                    children.push(new Paragraph({
                        children: [new TextRun({ text: `Priority Formula: ${vFormulaText}`, size: 22 })],
                        spacing: { after: 100 },
                        bidirectional: false,
                    }));
                    children.push(new Paragraph({
                        children: [new TextRun({ text: `Total Participants: ${appState.lwAggregatedResults.totalVotes}`, size: 22 })],
                        spacing: { after: 300 },
                        bidirectional: false,
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
                                new TableCell({ children: [new Paragraph({ text: 'Rank', bold: true, bidirectional: false })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: 'Duty', bold: true, bidirectional: false })], width: { size: 22, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: 'Task', bold: true, bidirectional: false })], width: { size: 35, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: 'I', bold: true, bidirectional: false })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: 'F', bold: true, bidirectional: false })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: 'D', bold: true, bidirectional: false })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                                new TableCell({ children: [new Paragraph({ text: 'PI', bold: true, bidirectional: false })], width: { size: 11, type: WidthType.PERCENTAGE } })
                            ]
                        })
                    ];
                    
                    verifiedTasks.forEach(task => {
                        verifiedTableRows.push(
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: String(task.rank), bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.dutyTitle, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.taskText, bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.meanImportance.toFixed(2), bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.meanFrequency.toFixed(2), bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.meanDifficulty.toFixed(2), bidirectional: false })] }),
                                    new TableCell({ children: [new Paragraph({ text: task.priorityIndex.toFixed(2), bidirectional: false })] })
                                ]
                            })
                        );
                    });
                    
                    children.push(new Table({
                        rows: verifiedTableRows,
                        width: { size: 100, type: WidthType.PERCENTAGE }
                    }));
                }

                // ============ COMPETENCY CLUSTERS SECTION ============
                if (appState.clusteringData.clusters && appState.clusteringData.clusters.length > 0) {
                    children.push(new Paragraph({ children: [new PageBreak()], bidirectional: false }));
                    
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: 'Competency Clusters',
                                bold: true,
                                size: 32, // 16pt
                            }),
                        ],
                        spacing: { before: 400, after: 400 },
                        alignment: AlignmentType.CENTER,
                        bidirectional: false,
                    }));
                    
                    appState.clusteringData.clusters.forEach((cluster, clusterIndex) => {
                        const clusterNumber = clusterIndex + 1;
                        
                        // Cluster header
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Competency ${clusterNumber}: ${cluster.name}`,
                                    bold: true,
                                    size: 28, // 14pt
                                }),
                            ],
                            spacing: { before: 300, after: 200 },
                            bidirectional: false,
                        }));
                        
                        // Range section
                        if (cluster.range && cluster.range.trim()) {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: 'Range:',
                                        bold: true,
                                        size: 24, // 12pt
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                                bidirectional: false,
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
                                bidirectional: false,
                            }));
                        }
                        
                        // Related Tasks section
                        if (cluster.tasks && cluster.tasks.length > 0) {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: 'Related Tasks:',
                                        bold: true,
                                        size: 24, // 12pt
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                                bidirectional: false,
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
                                    bidirectional: false,
                                }));
                            });
                        }
                        
                        // Performance Criteria section
                        if (cluster.performanceCriteria && cluster.performanceCriteria.length > 0) {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: 'Performance Criteria:',
                                        bold: true,
                                        size: 24, // 12pt
                                    }),
                                ],
                                spacing: { before: 200, after: 100 },
                                bidirectional: false,
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
                                    bidirectional: false,
                                }));
                            });
                        }
                    });
                }

                // ============ LEARNING OUTCOMES SECTION ============
                if (appState.learningOutcomesData.outcomes && appState.learningOutcomesData.outcomes.length > 0) {
                    children.push(new Paragraph({ children: [new PageBreak()], bidirectional: false }));
                    
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: 'Learning Outcomes',
                                bold: true,
                                size: 32, // 16pt
                            }),
                        ],
                        spacing: { before: 400, after: 400 },
                        alignment: AlignmentType.CENTER,
                        bidirectional: false,
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
                            bidirectional: false,
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
                                bidirectional: false,
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
                                    bidirectional: false,
                                }));
                            }
                            
                            // Mapped Performance Criteria
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: 'Mapped Performance Criteria:',
                                        italic: true,
                                        size: 20, // 10pt
                                    }),
                                ],
                                spacing: { before: 100, after: 50 },
                                indent: { left: 720 },
                                bidirectional: false,
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
                                    bidirectional: false,
                                }));
                            });
                        });
                    });
                }

                // ============ MODULE MAPPING SECTION ============
                if (appState.moduleMappingData.modules && appState.moduleMappingData.modules.length > 0) {
                    children.push(new Paragraph({ children: [new PageBreak()], bidirectional: false }));
                    
                    children.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: 'Module Mapping',
                                bold: true,
                                size: 32, // 16pt
                            }),
                        ],
                        spacing: { before: 400, after: 400 },
                        alignment: AlignmentType.CENTER,
                        bidirectional: false,
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
                            bidirectional: false,
                        }));
                        
                        // Learning Outcomes header
                        children.push(new Paragraph({
                            children: [
                                new TextRun({
                                    text: 'Learning Outcomes:',
                                    bold: true,
                                    size: 24, // 12pt
                                }),
                            ],
                            spacing: { before: 200, after: 100 },
                            bidirectional: false,
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
                                bidirectional: false,
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
                                    bidirectional: false,
                                }));
                            }
                            
                            // Referenced Performance Criteria
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({
                                        text: 'Referenced PC:',
                                        italic: true,
                                        size: 18, // 9pt
                                    }),
                                ],
                                spacing: { before: 50, after: 30 },
                                indent: { left: 1440 },
                                bidirectional: false,
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
                                    bidirectional: false,
                                }));
                            });
                        });
                    });
                }

                // Create document
                const doc = new Document({
                    sections: [{
                        properties: {
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
                const _occSlug = occupationTitle.replace(/[^a-z0-9]/gi, '_');
                const _jobSlug = (jobTitle && jobTitle.trim()) ? `_${jobTitle.replace(/[^a-z0-9]/gi, '_')}` : '';
                link.download = `${_occSlug}${_jobSlug}_DACUM_Chart.docx`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                
                showStatus('Word document exported successfully! ✓', 'success');

            } catch (error) {
                console.error('Error generating Word document:', error);
                showStatus('Error generating Word document: ' + error.message, 'error');
            }
        }
