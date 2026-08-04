// ============================================================
// /exports.js
// PDF and DOCX export functions (TV standalone + full chart)
// ============================================================

import { appState } from './state.js';
import { showStatus, escapeHtml } from './renderer.js';
import { getTaskCode, getDutyLetter } from './codes.js';

export async function exportTaskVerificationWord() {
            try {
                // Check if we have Task Verification data
                if (appState.collectionMode !== 'workshop' || !appState.workshopResults || Object.keys(appState.workshopResults).length === 0) {
                    alert('No Task Verification data available. Please complete workshop counts in the Task Verification tab first.');
                    return;
                }
                
                const validResults = Object.keys(appState.workshopResults).filter(key => 
                    appState.workshopResults[key] && appState.workshopResults[key].valid
                );
                
                if (validResults.length === 0) {
                    alert('No valid Task Verification results. Please ensure all required fields are completed.');
                    return;
                }
                
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
                            shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Duty', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Task', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Mean I', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Mean F', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Mean D', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: 'Priority', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                            shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
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
                Object.keys(appState.workshopResults).forEach(taskKey => {
                    const result = appState.workshopResults[taskKey];
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
    // chart read as two different documents. Cell text is black, which
    // also has far better contrast on light grey than on the old purple.
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
                                        type: ShadingType.SOLID,
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
                                        type: ShadingType.SOLID,
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
                                        type: ShadingType.SOLID,
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
                                        type: ShadingType.SOLID,
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
                                        type: ShadingType.SOLID,
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
                                        type: ShadingType.SOLID,
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
                                        type: ShadingType.SOLID,
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
                                        type: ShadingType.SOLID,
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
                                    shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Duty', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Task', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Mean I', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Mean F', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Mean D', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
                                }),
                                new TableCell({
                                    children: [new Paragraph({ children: [new TextRun({ text: 'Priority', bold: true })], alignment: AlignmentType.CENTER, bidirectional: false })],
                                    shading: { fill: 'DCDCDC', type: ShadingType.SOLID, color: 'ffffff' },
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

export function exportTaskVerificationPDF() {
    try {
        // Check if we have Task Verification data
        if (appState.collectionMode !== 'workshop' || !appState.workshopResults || Object.keys(appState.workshopResults).length === 0) {
            alert('No Task Verification data available. Please complete workshop counts in the Task Verification tab first.');
            return;
        }
        
        const validResults = Object.keys(appState.workshopResults).filter(key => 
            appState.workshopResults[key] && appState.workshopResults[key].valid
        );
        
        if (validResults.length === 0) {
            alert('No valid Task Verification results. Please ensure all required fields are completed.');
            return;
        }
        
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
            orientation: 'landscape',
            unit: 'mm',
            format: 'a4'
        });
        
        const margin = 10;
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        let yPos = margin + 10;
        
        // Get occupation title
        const occupationTitleInput = document.getElementById('occupationTitle');
        const occupationTitle = occupationTitleInput ? occupationTitleInput.value : 'Unknown Occupation';
        
        // Title Page
        pdf.setFontSize(18);
        pdf.setFont(undefined, 'bold');
        pdf.text('Task Verification & Training Priority Analysis', pageWidth / 2, yPos, { align: 'center' });
        yPos += 12;
        
        pdf.setFontSize(14);
        pdf.setFont(undefined, 'bold');
        pdf.text(`Occupation: ${occupationTitle}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 10;
        
        pdf.setFontSize(12);
        pdf.setFont(undefined, 'normal');
        const today = new Date().toLocaleDateString();
        pdf.text(`Date of Analysis: ${today}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 8;
        
        pdf.setFontSize(10);
        pdf.setFont(undefined, 'italic');
        pdf.text(`This Task Verification is based on the DACUM Chart for ${occupationTitle}.`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 15;
        
        // Methodology Summary
        pdf.setFontSize(14);
        pdf.setFont(undefined, 'bold');
        pdf.text('Methodology Summary', margin, yPos);
        yPos += 8;
        
        pdf.setFontSize(11);
        pdf.setFont(undefined, 'normal');
        pdf.text(`Data Collection Mode: ${appState.collectionMode === 'workshop' ? 'Workshop (Facilitated)' : 'Individual/Survey'}`, margin, yPos);
        yPos += 6;
        pdf.text(`Number of Participants: ${appState.workshopParticipants}`, margin, yPos);
        yPos += 6;
        pdf.text(`Workflow Mode: ${appState.workflowMode === 'standard' ? 'Standard (DACUM)' : 'Extended (DACUM)'}`, margin, yPos);
        yPos += 6;
        pdf.text(`Priority Formula: ${appState.priorityFormula === 'if' ? 'Importance × Frequency' : 'Importance × Frequency × Difficulty'}`, margin, yPos);
        yPos += 12;
        
        // Priority Rankings Table
        pdf.setFontSize(14);
        pdf.setFont(undefined, 'bold');
        pdf.text('Priority Rankings', margin, yPos);
        yPos += 8;
        
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
        
        // Table headers
        const colWidths = [15, 50, 75, 25, 25, 25, 25];
        const headers = ['Rank', 'Duty', 'Task', 'Mean I', 'Mean F', 'Mean D', 'Priority'];
        
        pdf.setFontSize(10);
        pdf.setFont(undefined, 'bold');
        let xPos = margin;
        headers.forEach((header, i) => {
            pdf.text(header, xPos, yPos);
            xPos += colWidths[i];
        });
        yPos += 6;
        
        // Table rows
        pdf.setFont(undefined, 'normal');
        sortedResults.forEach((row, index) => {
            if (yPos > pageHeight - 20) {
                pdf.addPage();
                yPos = margin + 10;
            }
            
            xPos = margin;
            pdf.text(`#${index + 1}`, xPos, yPos);
            xPos += colWidths[0];
            
            const dutyTrunc = row.duty.length > 20 ? row.duty.substring(0, 17) + '...' : row.duty;
            pdf.text(dutyTrunc, xPos, yPos);
            xPos += colWidths[1];
            
            const taskTrunc = row.task.length > 40 ? row.task.substring(0, 37) + '...' : row.task;
            pdf.text(taskTrunc, xPos, yPos);
            xPos += colWidths[2];
            
            pdf.text(row.meanI !== null ? row.meanI.toFixed(2) : 'N/A', xPos, yPos);
            xPos += colWidths[3];
            pdf.text(row.meanF !== null ? row.meanF.toFixed(2) : 'N/A', xPos, yPos);
            xPos += colWidths[4];
            pdf.text(row.meanD !== null ? row.meanD.toFixed(2) : 'N/A', xPos, yPos);
            xPos += colWidths[5];
            pdf.text(row.priority !== null ? row.priority.toFixed(2) : 'N/A', xPos, yPos);
            
            yPos += 5;
        });
        
        yPos += 10;
        
        // Duty-Level Summary Section
        if (yPos > pageHeight - 30) {
            pdf.addPage();
            yPos = margin + 10;
        }
        
        pdf.setFontSize(14);
        pdf.setFont(undefined, 'bold');
        pdf.text('Duty-Level Summary', margin, yPos);
        yPos += 5;
        
        pdf.setFontSize(9);
        pdf.setFont(undefined, 'italic');
        pdf.text(`Training Load Method: ${appState.trainingLoadMethod === 'advanced' ? 'Advanced (Σ Priority × Difficulty)' : 'Simple (Avg Priority × Tasks)'}`, margin, yPos);
        yPos += 8;
        
        // Aggregate duty-level data
        const dutyMap = {};
        Object.keys(appState.workshopResults).forEach(taskKey => {
            const result = appState.workshopResults[taskKey];
            if (result && result.valid) {
                let dutyId = result.dutyId || taskKey.split('_task_')[0];
                let dutyTitle = result.dutyTitle;
                
                if (!dutyTitle) {
                    const dutyInput = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`);
                    dutyTitle = dutyInput ? dutyInput.value.trim() : 'Unassigned';
                }
                
                if (!dutyMap[dutyId]) {
                    dutyMap[dutyId] = {
                        dutyTitle: dutyTitle,
                        validTasks: 0,
                        prioritySum: 0,
                        difficultySum: 0,
                        tasks: []
                    };
                }
                
                dutyMap[dutyId].validTasks++;
                dutyMap[dutyId].prioritySum += result.priorityIndex;
                dutyMap[dutyId].difficultySum += result.meanDifficulty;
                dutyMap[dutyId].tasks.push({
                    priorityIndex: result.priorityIndex,
                    meanDifficulty: result.meanDifficulty
                });
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
            
            dutyResults.push({
                dutyTitle: duty.dutyTitle,
                validTasks: duty.validTasks,
                avgPriority: avgPriority,
                trainingLoad: trainingLoad
            });
        });
        
        dutyResults.sort((a, b) => b.avgPriority - a.avgPriority);
        
        // Duty table headers
        const dutyColWidths = [80, 30, 40, 45];
        const dutyHeaders = ['Duty Title', 'Tasks', 'Avg Priority', 'Training Load'];
        
        pdf.setFontSize(9);
        pdf.setFont(undefined, 'bold');
        let dutyXPos = margin;
        dutyHeaders.forEach((header, i) => {
            pdf.text(header, dutyXPos, yPos);
            dutyXPos += dutyColWidths[i];
        });
        yPos += 6;
        
        // Duty table rows
        pdf.setFont(undefined, 'normal');
        dutyResults.forEach((duty) => {
            if (yPos > pageHeight - 20) {
                pdf.addPage();
                yPos = margin + 10;
            }
            
            dutyXPos = margin;
            const dutyTitleTrunc = duty.dutyTitle.length > 35 ? duty.dutyTitle.substring(0, 32) + '...' : duty.dutyTitle;
            pdf.text(dutyTitleTrunc, dutyXPos, yPos);
            dutyXPos += dutyColWidths[0];
            
            pdf.text(duty.validTasks.toString(), dutyXPos, yPos);
            dutyXPos += dutyColWidths[1];
            
            pdf.text(duty.avgPriority.toFixed(2), dutyXPos, yPos);
            dutyXPos += dutyColWidths[2];
            
            pdf.text(duty.trainingLoad.toFixed(2), dutyXPos, yPos);
            
            yPos += 5;
        });
        
        yPos += 10;
        
        // Notes section
        pdf.setFontSize(12);
        pdf.setFont(undefined, 'bold');
        pdf.text('Notes & Methodology', margin, yPos);
        yPos += 7;
        
        pdf.setFontSize(10);
        pdf.setFont(undefined, 'normal');
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
            if (yPos > pageHeight - 15) {
                pdf.addPage();
                yPos = margin + 10;
            }
            pdf.text(`• ${note}`, margin, yPos);
            yPos += 5;
        });
        
        // Save PDF
        pdf.save(`${occupationTitle.replace(/[^a-z0-9]/gi, '_')}_Task_Verification.pdf`);
        showStatus('Task Verification PDF exported successfully! ✓', 'success');
        
    } catch (error) {
        console.error('Error generating Task Verification PDF:', error);
        showStatus('Error generating Task Verification PDF: ' + error.message, 'error');
    }
}

export function exportToPDF() {
    // ============ CHECK FOR VERIFIED LIVE WORKSHOP RESULTS ============
    const hasVerifiedResults = typeof appState.lwFinalizedData !== 'undefined' && appState.lwFinalizedData && 
                                typeof appState.lwAggregatedResults !== 'undefined' && appState.lwAggregatedResults;
    
    // ============ VERIFIED LIVE WORKSHOP STANDALONE EXPORT ============
    if (hasVerifiedResults && appState.tvExportMode === 'standalone') {
        lwExportVerifiedPDF();
        return;
    }
    
    // ============ REGULAR TASK VERIFICATION STANDALONE EXPORT ============
    if (!hasVerifiedResults && appState.tvExportMode === 'standalone') {
        exportTaskVerificationPDF();
        return;
    }
    
    // ============ NORMAL DACUM EXPORT (with optional appendix) ============
    try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
            orientation: 'landscape',
            unit: 'mm',
            format: 'a4'
        });
        
        // Get input values
        const dacumDateInput = document.getElementById('dacumDate');
        let dacumDateFormatted = '';
        if (dacumDateInput.value) {
            const dateObj = new Date(dacumDateInput.value + 'T00:00:00');
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const day = String(dateObj.getDate()).padStart(2, '0');
            const year = dateObj.getFullYear();
            dacumDateFormatted = `${month}-${day}-${year}`;
        }
        const producedForInput = document.getElementById('producedFor');
        const producedByInput = document.getElementById('producedBy');
        const occupationTitleInput = document.getElementById('occupationTitle');
        const jobTitleInput = document.getElementById('jobTitle');
        const toolsInput = document.getElementById('toolsInput');
        const trendsInput = document.getElementById('trendsInput');
        const acronymsInput = document.getElementById('acronymsInput');
        
        // Validation — Occupation Title required; Job Title optional
        if (!occupationTitleInput.value) {
            alert('Please enter an Occupation Title before exporting.');
            showStatus('Occupation Title is required for export.', 'error');
            return;
        }
        
        const margin = 10;
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        let yPos = margin + 10;
        
        // ============ TITLE PAGE ============
        pdf.setFontSize(18); // 18pt for main title
        pdf.setFont(undefined, 'bold');
        pdf.text(`DACUM Research Chart for ${occupationTitleInput.value}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 15;
        
        // Two column layout for title page
        const leftColX = margin + 10;
        const rightColX = pageWidth / 2 + 10;
        let leftY = yPos;
        let rightY = yPos;
        
        // Left column - Produced For/By
        if (producedForInput.value) {
            pdf.setFontSize(16); // 16pt for labels
            pdf.setFont(undefined, 'bold');
            pdf.text('Produced for', leftColX, leftY);
            leftY += 7;
            
            // Add logo if exists
            if (appState.producedForImage) {
                try {
                    const imgWidth = 30;
                    const imgHeight = 20;
                    pdf.addImage(appState.producedForImage, 'JPEG', leftColX, leftY, imgWidth, imgHeight);
                    leftY += imgHeight + 5;
                } catch (e) {
                    console.error('Error adding Produced For image:', e);
                }
            }
            
            pdf.setFont(undefined, 'normal');
            pdf.setFontSize(14); // 14pt for content
            pdf.text(producedForInput.value, leftColX, leftY);
            leftY += 15;
        }
        
        if (producedByInput.value) {
            pdf.setFontSize(16); // 16pt for labels
            pdf.setFont(undefined, 'bold');
            pdf.text('Produced by', leftColX, leftY);
            leftY += 7;
            
            // Add logo if exists
            if (appState.producedByImage) {
                try {
                    const imgWidth = 30;
                    const imgHeight = 20;
                    pdf.addImage(appState.producedByImage, 'JPEG', leftColX, leftY, imgWidth, imgHeight);
                    leftY += imgHeight + 5;
                } catch (e) {
                    console.error('Error adding Produced By image:', e);
                }
            }
            
            pdf.setFont(undefined, 'normal');
            pdf.setFontSize(14); // 14pt for content
            pdf.text(producedByInput.value, leftColX, leftY);
            leftY += 10;
        }
        
        if (dacumDateFormatted) {
            pdf.setFontSize(14); // 14pt for date
            pdf.setFont(undefined, 'bold');
            pdf.text(dacumDateFormatted, leftColX, leftY);
            leftY += 7;
        }
        
        // Add venue if exists
        const venueInput = document.getElementById('venue');
        if (venueInput && venueInput.value) {
            pdf.setFontSize(14);
            pdf.setFont(undefined, 'bold');
            const venueLabel = 'Venue: ';
            pdf.text(venueLabel, leftColX, leftY);
            const venueLabelW = pdf.getTextWidth(venueLabel) + 1;
            pdf.setFont(undefined, 'normal');
            pdf.text(venueInput.value, leftColX + venueLabelW, leftY);
        }
        
        // Right column - Job info
        //
        // Label offsets are MEASURED, not hard-coded. The previous
        // fixed +30mm / +15mm offsets were narrower than the bold 16pt
        // labels themselves, so the value printed on top of the label
        // ("Occupation:Furniture Carpenter"). getTextWidth returns the
        // real width at the current font+size, so the gap is always
        // correct whatever the label or font.
        //
        // Values also wrap now: a long occupation title used to run off
        // the right edge of the page instead of flowing to a new line.
        const rightColMaxW = pageWidth - margin - rightColX;
        const LABEL_GAP    = 3;   // mm between label and value

        const drawLabelledValue = (label, value) => {
            pdf.setFontSize(16);
            pdf.setFont(undefined, 'bold');
            pdf.text(label, rightColX, rightY);
            const labelW = pdf.getTextWidth(label) + LABEL_GAP;

            pdf.setFontSize(14);
            pdf.setFont(undefined, 'normal');
            const valueLines = pdf.splitTextToSize(value, Math.max(20, rightColMaxW - labelW));
            pdf.text(valueLines, rightColX + labelW, rightY);
            rightY += Math.max(7, valueLines.length * 5.7);
        };

        // NOTE: these two were previously crossed over — the
        // "Occupation:" line printed jobTitleInput.value and the "Job:"
        // line printed occupationTitleInput.value. Corrected here.
        drawLabelledValue('Occupation:', occupationTitleInput.value);

        // Job line — only if filled (skips orphan "Job:" label with empty value)
        if (jobTitleInput.value && jobTitleInput.value.trim()) {
            drawLabelledValue('Job:', jobTitleInput.value.trim());
        }
        
        // Workshop Roles Section
        const facilitatorsInput = document.getElementById('facilitators');
        const observersInput = document.getElementById('observers');
        const panelMembersInput = document.getElementById('panelMembers');
        
        let workshopY = Math.max(leftY, rightY) + 15;
        const tableWidth = pageWidth - (2 * margin) - 20;
        const tableX = margin + 10;

        // ─── Scope of Work / Occupational Definition (full-width, optional) ───
        // Rendered below the two-column Produced For/By + Job block, before
        // the Facilitators section.  Wrapped with splitTextToSize so long
        // definitions flow across lines cleanly.  Paginates if the paragraph
        // would overflow the current page.
        const scopeOfWorkInput    = document.getElementById('scopeOfWork');
        const scopeOfWorkValuePDF = scopeOfWorkInput ? scopeOfWorkInput.value.trim() : '';
        if (scopeOfWorkValuePDF) {
            if (workshopY + 20 > pageHeight - margin) {
                pdf.addPage('a4', 'portrait');
                workshopY = margin + 10;
            }
            pdf.setFontSize(14);
            pdf.setFont(undefined, 'bold');
            pdf.text('Scope of Work / Occupational Definition', tableX, workshopY);
            workshopY += 6;

            pdf.setFontSize(11);
            pdf.setFont(undefined, 'normal');
            const scopeLines = pdf.splitTextToSize(scopeOfWorkValuePDF, tableWidth);
            scopeLines.forEach(line => {
                if (workshopY + 5 > pageHeight - margin) {
                    pdf.addPage('a4', 'portrait');
                    workshopY = margin + 10;
                }
                pdf.text(line, tableX, workshopY);
                workshopY += 5;
            });
            workshopY += 6;
        }
        
        if (facilitatorsInput && facilitatorsInput.value.trim()) {
            const facilitatorNames = facilitatorsInput.value.split('\n').map(s => s.trim()).filter(s => s);
            if (facilitatorNames.length > 0) {
                if (workshopY + 20 > pageHeight - margin) {
                    pdf.addPage('a4', 'portrait');
                    workshopY = margin + 10;
                }
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text('Facilitators', tableX, workshopY);
                workshopY += 5;
                pdf.setFont(undefined, 'normal');
                pdf.setFontSize(12);
                
                facilitatorNames.forEach(name => {
                    if (workshopY + 7 > pageHeight - margin) {
                        pdf.addPage('a4', 'portrait');
                        workshopY = margin + 10;
                    }
                    pdf.rect(tableX, workshopY, tableWidth, 6, 'S');
                    pdf.text(name, tableX + 2, workshopY + 4);
                    workshopY += 6;
                });
                workshopY += 4;
            }
        }
        
        if (observersInput && observersInput.value.trim()) {
            const observerNames = observersInput.value.split('\n').map(s => s.trim()).filter(s => s);
            if (observerNames.length > 0) {
                if (workshopY + 20 > pageHeight - margin) {
                    pdf.addPage('a4', 'portrait');
                    workshopY = margin + 10;
                }
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text('Observers', tableX, workshopY);
                workshopY += 5;
                pdf.setFont(undefined, 'normal');
                pdf.setFontSize(12);
                
                observerNames.forEach(name => {
                    if (workshopY + 7 > pageHeight - margin) {
                        pdf.addPage('a4', 'portrait');
                        workshopY = margin + 10;
                    }
                    pdf.rect(tableX, workshopY, tableWidth, 6, 'S');
                    pdf.text(name, tableX + 2, workshopY + 4);
                    workshopY += 6;
                });
                workshopY += 4;
            }
        }
        
        if (panelMembersInput && panelMembersInput.value.trim()) {
            const panelMemberNames = panelMembersInput.value.split('\n').map(s => s.trim()).filter(s => s);
            if (panelMemberNames.length > 0) {
                if (workshopY + 20 > pageHeight - margin) {
                    pdf.addPage('a4', 'portrait');
                    workshopY = margin + 10;
                }
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text('Panel Members', tableX, workshopY);
                workshopY += 5;
                pdf.setFont(undefined, 'normal');
                pdf.setFontSize(12);
                
                panelMemberNames.forEach(name => {
                    if (workshopY + 7 > pageHeight - margin) {
                        pdf.addPage('a4', 'portrait');
                        workshopY = margin + 10;
                    }
                    pdf.rect(tableX, workshopY, tableWidth, 6, 'S');
                    pdf.text(name, tableX + 2, workshopY + 4);
                    workshopY += 6;
                });
            }
        }
        
        // ============ DACUM CHART GRID ============
        pdf.addPage('a4', 'landscape');
        yPos = margin + 5;
        
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
        
        if (duties.length === 0) {
            showStatus('Please add at least one duty with tasks', 'error');
            return;
        }
        
        // DUTIES AND TASKS header
        pdf.setFillColor(200, 200, 200);
        pdf.rect(margin, yPos, pageWidth - (margin * 2), 8, 'FD');
        pdf.setFontSize(14); // 14pt for heading
        pdf.setFont(undefined, 'bold');
        pdf.text('DUTIES AND TASKS', pageWidth / 2, yPos + 5.5, { align: 'center' });
        yPos += 8;
        
        // ── DUTIES AND TASKS grid — horizontal (Word-style) layout ──
        //
        // Each duty gets its OWN full-width table: a shaded title bar
        // spanning the page, with that duty's tasks flowing left-to-
        // right in rows of four beneath it (A1 A2 A3 A4 / A5 A6 A7 A8 …).
        //
        // This replaces the previous layout, which placed four DUTIES
        // side by side as columns and ran their tasks vertically down
        // each column. The horizontal arrangement is the conventional
        // DACUM chart form and matches the Word master document, so the
        // two exports of the same chart no longer disagree on structure.
        // Colours are unchanged: 220-grey duty bars, outlined task cells.
        const TASK_COLS  = 4;
        const chartWidth = pageWidth - (margin * 2);
        const colWidth   = chartWidth / TASK_COLS;

        // ── Cell metrics ───────────────────────────────────────────
        // jsPDF renders multi-line text at fontSize x lineHeightFactor
        // (1.15 by default): 12pt -> 4.87mm per line, 14pt -> 5.68mm.
        const LINE_H_TASK = 4.9;   // 12pt
        const LINE_H_DUTY = 5.7;   // 14pt
        const CELL_PAD_X  = 2.5;   // left/right inset
        const CELL_PAD_T  = 4.6;   // top inset to the FIRST BASELINE
        const CELL_PAD_B  = 2.5;   // breathing room under the last line
        const DUTY_GAP    = 5;     // vertical space between duty tables

        const bottomEdge = pageHeight - margin;

        // Repeat the section banner after every page break so a reader
        // landing mid-chart still knows what they are looking at.
        const drawSectionBanner = (text) => {
            pdf.setFillColor(200, 200, 200);
            pdf.rect(margin, yPos, chartWidth, 8, 'FD');
            pdf.setFontSize(14);
            pdf.setFont(undefined, 'bold');
            pdf.text(text, pageWidth / 2, yPos + 5.5, { align: 'center' });
            yPos += 8;
        };

        const newChartPage = () => {
            pdf.addPage('a4', 'landscape');
            yPos = margin + 5;
            drawSectionBanner('DUTIES AND TASKS (continued)');
        };

        duties.forEach((duty, dutyIdx) => {
            const letter = String.fromCharCode(65 + dutyIdx);

            // ── Duty title bar (full width) ───────────────────────
            pdf.setFontSize(14);
            pdf.setFont(undefined, 'bold');
            const headerText  = `DUTY ${letter}: ${duty.duty}`;
            const headerLines = pdf.splitTextToSize(headerText, chartWidth - (CELL_PAD_X * 2));
            const headerH     = Math.max(
                10,
                headerLines.length * LINE_H_DUTY + CELL_PAD_T + CELL_PAD_B
            );

            // Keep the bar with at least one task row — a title stranded
            // alone at the foot of a page reads as an error.
            if (yPos + headerH + 15 > bottomEdge) newChartPage();

            pdf.setFillColor(220, 220, 220);
            pdf.rect(margin, yPos, chartWidth, headerH, 'FD');
            pdf.text(headerLines, margin + CELL_PAD_X, yPos + CELL_PAD_T + 0.6);
            yPos += headerH;

            // ── Task rows, four per row ───────────────────────────
            const tasks = duty.tasks || [];
            for (let i = 0; i < tasks.length; i += TASK_COLS) {
                const rowTasks = tasks.slice(i, i + TASK_COLS);

                // Row height = tallest cell in the row
                pdf.setFontSize(12);
                pdf.setFont(undefined, 'normal');
                let rowHeight = 15;
                const cellLines = rowTasks.map((taskText, c) => {
                    const label = `Task ${letter}${i + c + 1}:`;
                    const lines = pdf.splitTextToSize(
                        `${label}\n${taskText}`, colWidth - (CELL_PAD_X * 2)
                    );
                    rowHeight = Math.max(
                        rowHeight,
                        lines.length * LINE_H_TASK + CELL_PAD_T + CELL_PAD_B
                    );
                    return lines;
                });

                if (yPos + rowHeight > bottomEdge) {
                    newChartPage();
                    // Restate which duty these rows belong to
                    pdf.setFillColor(220, 220, 220);
                    pdf.rect(margin, yPos, chartWidth, 8, 'FD');
                    pdf.setFontSize(12);
                    pdf.setFont(undefined, 'bold');
                    pdf.text(`DUTY ${letter} (continued)`, margin + CELL_PAD_X, yPos + 5.5);
                    yPos += 8;
                }

                // Draw all four cells — empty trailing cells are still
                // outlined so the table keeps a clean rectangular edge,
                // exactly as in the Word master.
                pdf.setFontSize(12);
                pdf.setFont(undefined, 'normal');
                for (let c = 0; c < TASK_COLS; c++) {
                    const x = margin + (c * colWidth);
                    pdf.rect(x, yPos, colWidth, rowHeight, 'S');
                    if (cellLines[c]) {
                        pdf.text(cellLines[c], x + CELL_PAD_X, yPos + CELL_PAD_T);
                    }
                }

                yPos += rowHeight;
            }

            yPos += DUTY_GAP;
        });
        
        // ============ KNOWLEDGE, SKILLS, BEHAVIORS ============
        const knowledgeText = document.getElementById('knowledgeInput').value.trim();
        const skillsText = document.getElementById('skillsInput').value.trim();
        const behaviorsText = document.getElementById('behaviorsInput').value.trim();
        
        if (knowledgeText || skillsText || behaviorsText) {
            pdf.addPage('a4', 'landscape');
            yPos = margin + 5;
            
            pdf.setFontSize(14); // 14pt for main heading
            pdf.setFont(undefined, 'bold');
            pdf.text('General Knowledge and Skills', pageWidth / 2, yPos, { align: 'center' });
            yPos += 8;
            
            // ── Three-column layout ────────────────────────────────
            // Each item is now WRAPPED to its own column width. The old
            // code drew every line at full length with no width limit,
            // so any item longer than a third of the page ran straight
            // across the neighbouring column — the overlap seen in the
            // exported chart. A gutter keeps adjacent columns from
            // touching even when both are full.
            const COL_GUTTER  = 6;                                   // mm between columns
            const thirdWidth  = (pageWidth - (margin * 2)) / 3;
            const colTextW    = thirdWidth - COL_GUTTER;
            const LINE_H_ITEM = 4.9;                                 // 12pt
            const bottomLimit = pageHeight - margin;

            const drawColumn = (text, headingId, colX) => {
                if (!text) return;
                let y = yPos;

                const headingEl = document.getElementById(headingId);
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text(
                    pdf.splitTextToSize(headingEl ? headingEl.textContent : '', colTextW),
                    colX, y
                );
                y += 7;

                pdf.setFontSize(12);
                pdf.setFont(undefined, 'normal');

                text.split('\n').filter(line => line.trim()).forEach(item => {
                    const clean = item.trim().replace(/^[•\-*]\s*/, '');
                    const lines = pdf.splitTextToSize(clean, colTextW);
                    lines.forEach(line => {
                        if (y > bottomLimit) return;   // clip rather than spill off-page
                        pdf.text(line, colX, y);
                        y += LINE_H_ITEM;
                    });
                    y += 1;                            // small gap between items
                });
            };

            drawColumn(knowledgeText, 'knowledgeHeading', margin);
            drawColumn(skillsText,    'skillsHeading',    margin + thirdWidth);
            drawColumn(behaviorsText, 'behaviorsHeading', margin + (thirdWidth * 2));
        }
        
        // ============ TOOLS AND TRENDS ============
        const tools = toolsInput.value.trim() ? toolsInput.value.split('\n').filter(line => line.trim()) : [];
        const trends = trendsInput.value.trim() ? trendsInput.value.split('\n').filter(line => line.trim()) : [];
        
        if (tools.length > 0 || trends.length > 0) {
            pdf.addPage('a4', 'landscape');
            yPos = margin + 5;
            
            // Same wrapping treatment as the Knowledge/Skills page —
            // tool and material names are often long enough to cross
            // into the neighbouring column when drawn unconstrained.
            const halfWidth   = (pageWidth - (margin * 2) - 5) / 2;
            const colTextW2   = halfWidth - 6;
            const LINE_H_ITEM2 = 4.9;
            const bottomLimit2 = pageHeight - margin;

            const drawList = (items, headingId, colX) => {
                if (!items.length) return;
                let y = yPos;

                const headingEl = document.getElementById(headingId);
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text(
                    pdf.splitTextToSize(headingEl ? headingEl.textContent : '', colTextW2),
                    colX, y
                );
                y += 7;

                pdf.setFontSize(12);
                pdf.setFont(undefined, 'normal');
                items.forEach(item => {
                    const clean = item.trim().replace(/^[•\-*]\s*/, '');
                    pdf.splitTextToSize(clean, colTextW2).forEach(line => {
                        if (y > bottomLimit2) return;
                        pdf.text(line, colX, y);
                        y += LINE_H_ITEM2;
                    });
                    y += 1;
                });
            };

            drawList(tools,  'toolsHeading',  margin);
            drawList(trends, 'trendsHeading', margin + halfWidth + 5);
        }
        
        // ============ ACRONYMS ============
        if (acronymsInput.value.trim()) {
            pdf.addPage('a4', 'landscape');
            yPos = margin + 5;
            
            const heading = document.getElementById('acronymsHeading').textContent;
            pdf.setFontSize(14); // 14pt for section heading
            pdf.setFont(undefined, 'bold');
            pdf.text(heading, margin, yPos);
            yPos += 6;
            
            pdf.setFontSize(12); // 12pt for content
            pdf.setFont(undefined, 'normal');
            const acronyms = acronymsInput.value.split('\n').filter(line => line.trim());
            const fullTextW = pageWidth - (margin * 2);
            acronyms.forEach(acronym => {
                const clean = acronym.trim().replace(/^[•\-*]\s*/, '');
                pdf.splitTextToSize(clean, fullTextW).forEach(line => {
                    if (yPos > pageHeight - margin) {
                        pdf.addPage('a4', 'landscape');
                        yPos = margin + 5;
                    }
                    pdf.text(line, margin, yPos);
                    yPos += 4.9;
                });
            });
        }
        
        // ============ CAREER PATH ============
        const careerPathInput = document.getElementById('careerPathInput');
        if (careerPathInput && careerPathInput.value.trim()) {
            pdf.addPage('a4', 'landscape');
            yPos = margin + 5;
            
            const heading = document.getElementById('careerPathHeading').textContent;
            pdf.setFontSize(14); // 14pt for section heading
            pdf.setFont(undefined, 'bold');
            pdf.text(heading, margin, yPos);
            yPos += 6;
            
            pdf.setFontSize(12); // 12pt for content
            pdf.setFont(undefined, 'normal');
            const careerPathItems = careerPathInput.value.split('\n').filter(line => line.trim());
            const fullTextW = pageWidth - (margin * 2);
            careerPathItems.forEach(item => {
                const clean = item.trim().replace(/^[•\-*]\s*/, '');
                pdf.splitTextToSize(clean, fullTextW).forEach(line => {
                    if (yPos > pageHeight - margin) {
                        pdf.addPage('a4', 'landscape');
                        yPos = margin + 5;
                    }
                    pdf.text(line, margin, yPos);
                    yPos += 4.9;
                });
            });
        }
        
        // ============ CUSTOM SECTIONS ============
        const customSectionsContainer = document.getElementById('customSectionsContainer');
        const customSectionDivs = customSectionsContainer.querySelectorAll('.section-container');
        customSectionDivs.forEach(sectionDiv => {
            const headingElement = sectionDiv.querySelector('h3');
            const textareaElement = sectionDiv.querySelector('textarea');
            
            if (headingElement && textareaElement && textareaElement.value.trim()) {
                pdf.addPage('a4', 'landscape');
                yPos = margin + 5;
                
                pdf.setFontSize(14); // 14pt for section heading
                pdf.setFont(undefined, 'bold');
                pdf.text(headingElement.textContent, margin, yPos);
                yPos += 6;
                
                pdf.setFontSize(12); // 12pt for content
                pdf.setFont(undefined, 'normal');
                const items = textareaElement.value.split('\n').filter(line => line.trim());
                const customTextW = pageWidth - (margin * 2);
                items.forEach(item => {
                    const clean = item.trim().replace(/^[•\-*]\s*/, '');
                    pdf.splitTextToSize(clean, customTextW).forEach(line => {
                        if (yPos > pageHeight - margin) {
                            pdf.addPage('a4', 'landscape');
                            yPos = margin + 5;
                        }
                        pdf.text(line, margin, yPos);
                        yPos += 4.9;
                    });
                });
            }
        });
        
        // ============ SKILLS LEVEL MATRIX (PDF EXPORT) ============
        const hasSkillsLevelData = appState.skillsLevelData?.some(category =>
            category.competencies.some(comp =>
                Object.values(comp.levels).some(v => v === true)
            )
        );

        if (hasSkillsLevelData) {
            // Add new page for Skills Level Matrix
            pdf.addPage('a4', 'landscape');
            yPos = margin + 5;
            
            // Main heading
            pdf.setFontSize(14);
            pdf.setFont(undefined, 'bold');
            pdf.text('Employability Competencies by Occupational Level', pageWidth / 2, yPos, { align: 'center' });
            yPos += 10;

            // Process each category
            appState.skillsLevelData.forEach(category => {
                // Skip empty categories
                if (category.category.trim() === '' && category.competencies.every(c => c.text.trim() === '')) {
                    return;
                }

                // Check if we need a new page
                if (yPos > pageHeight - 40) {
                    pdf.addPage('a4', 'landscape');
                    yPos = margin + 5;
                }

                // Category header
                pdf.setFontSize(12);
                pdf.setFont(undefined, 'bold');
                pdf.setFillColor(232, 232, 232);
                pdf.rect(margin, yPos - 4, pageWidth - (margin * 2), 6, 'F');
                pdf.text(category.category || `Category ${category.id}`, margin + 2, yPos);
                yPos += 8;

                // Column headers
                pdf.setFontSize(10);
                const colWidth = (pageWidth - (margin * 2)) / 5;
                pdf.setFillColor(245, 245, 245);
                pdf.rect(margin, yPos - 4, pageWidth - (margin * 2), 6, 'F');
                pdf.text('Competency', margin + 2, yPos);
                pdf.text('Craftsman', margin + colWidth * 1 + 2, yPos);
                pdf.text('Skilled', margin + colWidth * 2 + 2, yPos);
                pdf.text('Semi-skilled', margin + colWidth * 3 + 2, yPos);
                pdf.text('Foundation', margin + colWidth * 4 + 2, yPos);
                yPos += 8;

                // Competency rows
                pdf.setFont(undefined, 'normal');
                category.competencies
                    .filter(comp => comp.text.trim() !== '')
                    .forEach(competency => {
                        // Check if we need a new page
                        if (yPos > pageHeight - 20) {
                            pdf.addPage('a4', 'landscape');
                            yPos = margin + 5;
                            
                            // Repeat column headers on new page
                            pdf.setFontSize(10);
                            pdf.setFont(undefined, 'bold');
                            pdf.setFillColor(245, 245, 245);
                            pdf.rect(margin, yPos - 4, pageWidth - (margin * 2), 6, 'F');
                            pdf.text('Competency', margin + 2, yPos);
                            pdf.text('Craftsman', margin + colWidth * 1 + 2, yPos);
                            pdf.text('Skilled', margin + colWidth * 2 + 2, yPos);
                            pdf.text('Semi-skilled', margin + colWidth * 3 + 2, yPos);
                            pdf.text('Foundation', margin + colWidth * 4 + 2, yPos);
                            yPos += 8;
                            pdf.setFont(undefined, 'normal');
                        }

                        // Competency text
                        const competencyText = `${competency.id}. ${competency.text}`;
                        const textLines = pdf.splitTextToSize(competencyText, colWidth - 4);
                        const lineHeight = 5;
                        const cellHeight = Math.max(lineHeight * textLines.length, 6);

                        // Draw cell borders
                        pdf.rect(margin, yPos - 4, colWidth, cellHeight);
                        pdf.rect(margin + colWidth, yPos - 4, colWidth, cellHeight);
                        pdf.rect(margin + colWidth * 2, yPos - 4, colWidth, cellHeight);
                        pdf.rect(margin + colWidth * 3, yPos - 4, colWidth, cellHeight);
                        pdf.rect(margin + colWidth * 4, yPos - 4, colWidth, cellHeight);

                        // Competency text
                        textLines.forEach((line, idx) => {
                            pdf.text(line, margin + 2, yPos + (idx * lineHeight));
                        });

                        // Checkmarks (centered in cells)
                        const checkY = yPos + (cellHeight / 2) - 2;
                        if (competency.levels.craftsman) {
                            pdf.text('✓', margin + colWidth * 1 + (colWidth / 2), checkY, { align: 'center' });
                        }
                        if (competency.levels.skilled) {
                            pdf.text('✓', margin + colWidth * 2 + (colWidth / 2), checkY, { align: 'center' });
                        }
                        if (competency.levels.semiSkilled) {
                            pdf.text('✓', margin + colWidth * 3 + (colWidth / 2), checkY, { align: 'center' });
                        }
                        if (competency.levels.foundation) {
                            pdf.text('✓', margin + colWidth * 4 + (colWidth / 2), checkY, { align: 'center' });
                        }

                        yPos += cellHeight + 2;
                    });

                yPos += 5; // Extra space after category
            });
        }
        
        // ============ TASK VERIFICATION APPENDIX (if mode = 'appendix') ============
        if (appState.tvExportMode === 'appendix' && appState.collectionMode === 'workshop') {
            // Check if we have valid results to include
            const validResults = Object.keys(appState.workshopResults).filter(key => 
                appState.workshopResults[key] && appState.workshopResults[key].valid
            );
            
            if (validResults.length > 0) {
                // Start new page for appendix
                pdf.addPage();
                yPos = margin + 10;
                
                // Appendix title
                pdf.setFontSize(16);
                pdf.setFont(undefined, 'bold');
                pdf.text('Task Verification & Training Priority Analysis (Appendix)', pageWidth / 2, yPos, { align: 'center' });
                yPos += 12;
                
                // Methodology Summary
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text('Methodology Summary', margin, yPos);
                yPos += 8;
                
                pdf.setFontSize(11);
                pdf.setFont(undefined, 'normal');
                pdf.text(`Data Collection Mode: ${appState.collectionMode === 'workshop' ? 'Workshop (Facilitated)' : 'Individual/Survey'}`, margin, yPos);
                yPos += 6;
                pdf.text(`Number of Participants: ${appState.workshopParticipants}`, margin, yPos);
                yPos += 6;
                pdf.text(`Workflow Mode: ${appState.workflowMode === 'standard' ? 'Standard (DACUM)' : 'Extended (DACUM)'}`, margin, yPos);
                yPos += 6;
                pdf.text(`Priority Formula: ${appState.priorityFormula === 'if' ? 'Importance × Frequency' : 'Importance × Frequency × Difficulty'}`, margin, yPos);
                yPos += 12;
                
                // Priority Rankings Table
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text('Priority Rankings', margin, yPos);
                yPos += 8;
                
                // Get sorted results
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
                
                // Sort by priority descending
                sortedResults.sort((a, b) => b.priority - a.priority);
                
                // Table headers
                const colWidths = [15, 50, 75, 25, 25, 25, 25];
                const headers = ['Rank', 'Duty', 'Task', 'Mean I', 'Mean F', 'Mean D', 'Priority'];
                
                pdf.setFontSize(10);
                pdf.setFont(undefined, 'bold');
                let xPos = margin;
                headers.forEach((header, i) => {
                    pdf.text(header, xPos, yPos);
                    xPos += colWidths[i];
                });
                yPos += 6;
                
                // Table rows
                pdf.setFont(undefined, 'normal');
                sortedResults.forEach((row, index) => {
                    if (yPos > pageHeight - 20) {
                        pdf.addPage();
                        yPos = margin + 10;
                    }
                    
                    xPos = margin;
                    pdf.text(`#${index + 1}`, xPos, yPos);
                    xPos += colWidths[0];
                    
                    // Truncate long text
                    const dutyTrunc = row.duty.length > 20 ? row.duty.substring(0, 17) + '...' : row.duty;
                    pdf.text(dutyTrunc, xPos, yPos);
                    xPos += colWidths[1];
                    
                    const taskTrunc = row.task.length > 40 ? row.task.substring(0, 37) + '...' : row.task;
                    pdf.text(taskTrunc, xPos, yPos);
                    xPos += colWidths[2];
                    
                    pdf.text(row.meanI !== null ? row.meanI.toFixed(2) : 'N/A', xPos, yPos);
                    xPos += colWidths[3];
                    pdf.text(row.meanF !== null ? row.meanF.toFixed(2) : 'N/A', xPos, yPos);
                    xPos += colWidths[4];
                    pdf.text(row.meanD !== null ? row.meanD.toFixed(2) : 'N/A', xPos, yPos);
                    xPos += colWidths[5];
                    pdf.text(row.priority !== null ? row.priority.toFixed(2) : 'N/A', xPos, yPos);
                    
                    yPos += 5;
                });
                
                yPos += 8;
                
                // Duty-Level Summary Section
                if (yPos > pageHeight - 30) {
                    pdf.addPage();
                    yPos = margin + 10;
                }
                
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text('Duty-Level Summary', margin, yPos);
                yPos += 5;
                
                pdf.setFontSize(9);
                pdf.setFont(undefined, 'italic');
                pdf.text(`Training Load Method: ${appState.trainingLoadMethod === 'advanced' ? 'Advanced (Σ Priority × Difficulty)' : 'Simple (Avg Priority × Tasks)'}`, margin, yPos);
                yPos += 8;
                
                // Aggregate duty-level data
                const dutyMap = {};
                Object.keys(appState.workshopResults).forEach(taskKey => {
                    const result = appState.workshopResults[taskKey];
                    if (result && result.valid) {
                        let dutyId = result.dutyId || taskKey.split('_task_')[0];
                        let dutyTitle = result.dutyTitle;
                        
                        if (!dutyTitle) {
                            const dutyInput = document.querySelector(`input[data-duty-id="${dutyId}"], textarea[data-duty-id="${dutyId}"]`);
                            dutyTitle = dutyInput ? dutyInput.value.trim() : 'Unassigned';
                        }
                        
                        if (!dutyMap[dutyId]) {
                            dutyMap[dutyId] = {
                                dutyTitle: dutyTitle,
                                validTasks: 0,
                                prioritySum: 0,
                                difficultySum: 0,
                                tasks: []
                            };
                        }
                        
                        dutyMap[dutyId].validTasks++;
                        dutyMap[dutyId].prioritySum += result.priorityIndex;
                        dutyMap[dutyId].difficultySum += result.meanDifficulty;
                        dutyMap[dutyId].tasks.push({
                            priorityIndex: result.priorityIndex,
                            meanDifficulty: result.meanDifficulty
                        });
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
                    
                    dutyResults.push({
                        dutyTitle: duty.dutyTitle,
                        validTasks: duty.validTasks,
                        avgPriority: avgPriority,
                        trainingLoad: trainingLoad
                    });
                });
                
                dutyResults.sort((a, b) => b.avgPriority - a.avgPriority);
                
                // Duty table headers
                const dutyColWidths = [80, 30, 40, 45];
                const dutyHeaders = ['Duty Title', 'Tasks', 'Avg Priority', 'Training Load'];
                
                pdf.setFontSize(9);
                pdf.setFont(undefined, 'bold');
                let dutyXPos = margin;
                dutyHeaders.forEach((header, i) => {
                    pdf.text(header, dutyXPos, yPos);
                    dutyXPos += dutyColWidths[i];
                });
                yPos += 6;
                
                // Duty table rows
                pdf.setFont(undefined, 'normal');
                dutyResults.forEach((duty) => {
                    if (yPos > pageHeight - 20) {
                        pdf.addPage();
                        yPos = margin + 10;
                    }
                    
                    dutyXPos = margin;
                    const dutyTitleTrunc = duty.dutyTitle.length > 35 ? duty.dutyTitle.substring(0, 32) + '...' : duty.dutyTitle;
                    pdf.text(dutyTitleTrunc, dutyXPos, yPos);
                    dutyXPos += dutyColWidths[0];
                    
                    pdf.text(duty.validTasks.toString(), dutyXPos, yPos);
                    dutyXPos += dutyColWidths[1];
                    
                    pdf.text(duty.avgPriority.toFixed(2), dutyXPos, yPos);
                    dutyXPos += dutyColWidths[2];
                    
                    pdf.text(duty.trainingLoad.toFixed(2), dutyXPos, yPos);
                    
                    yPos += 5;
                });
                
                yPos += 8;
                
                // Notes
                pdf.setFontSize(12);
                pdf.setFont(undefined, 'bold');
                pdf.text('Notes', margin, yPos);
                yPos += 6;
                
                pdf.setFontSize(10);
                pdf.setFont(undefined, 'normal');
                const notes = [
                    'Weighted Mean = Σ(value × count) ÷ total responses',
                    'Priority Index calculated using selected formula',
                    'Higher priority values indicate greater training importance',
                    'Results based on DACUM methodology'
                ];
                notes.forEach(note => {
                    if (yPos > pageHeight - 15) {
                        pdf.addPage();
                        yPos = margin + 10;
                    }
                    pdf.text(`• ${note}`, margin, yPos);
                    yPos += 5;
                });
            }
        }
        
        // ============ VERIFIED LIVE WORKSHOP RESULTS APPENDIX ============
        if (appState.tvExportMode === 'appendix' && hasVerifiedResults) {
            // Start new page for verified results appendix
            pdf.addPage();
            yPos = margin + 10;
            
            // Appendix title
            pdf.setFontSize(16);
            pdf.setFont(undefined, 'bold');
            pdf.text('DACUM Live Pro - Verified (Post-Vote) Results (Appendix)', pageWidth / 2, yPos, { align: 'center' });
            yPos += 12;
            
            // Metadata
            pdf.setFontSize(11);
            pdf.setFont(undefined, 'normal');
            pdf.text(`Occupation: ${appState.lwFinalizedData.occupation}`, margin, yPos);
            yPos += 6;
            pdf.text(`Job Title: ${appState.lwFinalizedData.jobTitle}`, margin, yPos);
            yPos += 6;
            pdf.text(`Date: ${new Date().toLocaleDateString()}`, margin, yPos);
            yPos += 6;
            const vFormula = appState.lwFinalizedData.appState.priorityFormula || 'if';
            const vFormulaText = vFormula === 'ifd' ? 'Importance × Frequency × Difficulty' : 'Importance × Frequency';
            pdf.text(`Priority Formula: ${vFormulaText}`, margin, yPos);
            yPos += 6;
            pdf.text(`Total Participants: ${appState.lwAggregatedResults.totalVotes}`, margin, yPos);
            yPos += 12;
            
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
            
            // Table header
            pdf.setFontSize(10);
            pdf.setFont(undefined, 'bold');
            pdf.text('Rank', margin, yPos);
            pdf.text('Duty', margin + 15, yPos);
            pdf.text('Task', margin + 60, yPos);
            pdf.text('I', margin + 140, yPos);
            pdf.text('F', margin + 150, yPos);
            pdf.text('D', margin + 160, yPos);
            pdf.text('PI', margin + 170, yPos);
            yPos += 5;
            pdf.line(margin, yPos, pageWidth - margin, yPos);
            yPos += 3;
            
            // Table rows
            pdf.setFont(undefined, 'normal');
            pdf.setFontSize(8);
            
            verifiedTasks.forEach(task => {
                // Check if need new page
                if (yPos + 8 > pageHeight - margin) {
                    pdf.addPage();
                    yPos = margin;
                }
                
                pdf.text(String(task.rank), margin, yPos);
                const dutyLines = pdf.splitTextToSize(task.dutyTitle, 40);
                pdf.text(dutyLines[0] || '', margin + 15, yPos);
                const taskLines = pdf.splitTextToSize(task.taskText, 75);
                pdf.text(taskLines[0] || '', margin + 60, yPos);
                pdf.text(task.meanImportance.toFixed(2), margin + 140, yPos);
                pdf.text(task.meanFrequency.toFixed(2), margin + 150, yPos);
                pdf.text(task.meanDifficulty.toFixed(2), margin + 160, yPos);
                pdf.text(task.priorityIndex.toFixed(2), margin + 170, yPos);
                yPos += 6;
            });
        }
        
        // ============ COMPETENCY CLUSTERS SECTION ============
        if (appState.clusteringData.clusters && appState.clusteringData.clusters.length > 0) {
            pdf.addPage();
            yPos = margin + 5;
            
            pdf.setFontSize(16);
            pdf.setFont(undefined, 'bold');
            pdf.text('Competency Clusters', pageWidth / 2, yPos, { align: 'center' });
            yPos += 10;
            
            appState.clusteringData.clusters.forEach((cluster, clusterIndex) => {
                const clusterNumber = clusterIndex + 1;
                
                // Check if need new page
                if (yPos + 20 > pageHeight - margin) {
                    pdf.addPage();
                    yPos = margin + 5;
                }
                
                // Cluster header
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text(`Competency ${clusterNumber}: ${cluster.name}`, margin, yPos);
                yPos += 7;
                
                // Range section
                if (cluster.range && cluster.range.trim()) {
                    pdf.setFontSize(12);
                    pdf.setFont(undefined, 'bold');
                    pdf.text('Range:', margin, yPos);
                    yPos += 5;
                    
                    pdf.setFontSize(10);
                    pdf.setFont(undefined, 'normal');
                    const rangeLines = pdf.splitTextToSize(cluster.range, pageWidth - 2 * margin - 5);
                    rangeLines.forEach(line => {
                        if (yPos + 5 > pageHeight - margin) {
                            pdf.addPage();
                            yPos = margin + 5;
                        }
                        pdf.text(line, margin + 5, yPos);
                        yPos += 5;
                    });
                    yPos += 3;
                }
                
                // Related Tasks section
                if (cluster.tasks && cluster.tasks.length > 0) {
                    if (yPos + 10 > pageHeight - margin) {
                        pdf.addPage();
                        yPos = margin + 5;
                    }
                    
                    pdf.setFontSize(12);
                    pdf.setFont(undefined, 'bold');
                    pdf.text('Related Tasks:', margin, yPos);
                    yPos += 5;
                    
                    pdf.setFontSize(10);
                    pdf.setFont(undefined, 'normal');
                    
                    cluster.tasks.forEach(task => {
                        if (yPos + 6 > pageHeight - margin) {
                            pdf.addPage();
                            yPos = margin + 5;
                        }
                        
                        const taskCode = getTaskCode(task.id);
                        const taskText = `- ${taskCode}: ${task.text}`;
                        const lines = pdf.splitTextToSize(taskText, pageWidth - 2 * margin - 5);
                        
                        lines.forEach(line => {
                            if (yPos + 5 > pageHeight - margin) {
                                pdf.addPage();
                                yPos = margin + 5;
                            }
                            pdf.text(line, margin + 5, yPos);
                            yPos += 5;
                        });
                    });
                    yPos += 3;
                }
                
                // Performance Criteria section
                if (cluster.performanceCriteria && cluster.performanceCriteria.length > 0) {
                    if (yPos + 10 > pageHeight - margin) {
                        pdf.addPage();
                        yPos = margin + 5;
                    }
                    
                    pdf.setFontSize(12);
                    pdf.setFont(undefined, 'bold');
                    pdf.text('Performance Criteria:', margin, yPos);
                    yPos += 5;
                    
                    pdf.setFontSize(10);
                    pdf.setFont(undefined, 'normal');
                    
                    cluster.performanceCriteria.forEach((criterion, idx) => {
                        if (yPos + 6 > pageHeight - margin) {
                            pdf.addPage();
                            yPos = margin + 5;
                        }
                        
                        const criterionText = `${clusterNumber}-${idx + 1} ${criterion}`;
                        const lines = pdf.splitTextToSize(criterionText, pageWidth - 2 * margin - 5);
                        
                        lines.forEach(line => {
                            if (yPos + 5 > pageHeight - margin) {
                                pdf.addPage();
                                yPos = margin + 5;
                            }
                            pdf.text(line, margin + 5, yPos);
                            yPos += 5;
                        });
                    });
                }
                
                yPos += 8;
            });
        }
        
        // ============ LEARNING OUTCOMES SECTION ============
        if (appState.learningOutcomesData.outcomes && appState.learningOutcomesData.outcomes.length > 0) {
            pdf.addPage();
            yPos = margin + 5;
            
            pdf.setFontSize(16);
            pdf.setFont(undefined, 'bold');
            pdf.text('Learning Outcomes', pageWidth / 2, yPos, { align: 'center' });
            yPos += 10;
            
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
                
                if (yPos + 20 > pageHeight - margin) {
                    pdf.addPage();
                    yPos = margin + 5;
                }
                
                // Cluster header
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text(`${cluster.name}`, margin, yPos);
                yPos += 7;
                
                // Learning Outcomes for this cluster
                los.forEach(lo => {
                    if (yPos + 15 > pageHeight - margin) {
                        pdf.addPage();
                        yPos = margin + 5;
                    }
                    
                    pdf.setFontSize(12);
                    pdf.setFont(undefined, 'bold');
                    pdf.text(`${lo.number}:`, margin + 5, yPos);
                    yPos += 5;
                    
                    if (lo.statement && lo.statement.trim()) {
                        pdf.setFontSize(10);
                        pdf.setFont(undefined, 'normal');
                        const statementLines = pdf.splitTextToSize(lo.statement, pageWidth - 2 * margin - 10);
                        statementLines.forEach(line => {
                            if (yPos + 5 > pageHeight - margin) {
                                pdf.addPage();
                                yPos = margin + 5;
                            }
                            pdf.text(line, margin + 10, yPos);
                            yPos += 5;
                        });
                    }
                    
                    yPos += 2;
                    
                    // Mapped Performance Criteria
                    pdf.setFontSize(10);
                    pdf.setFont(undefined, 'italic');
                    pdf.text('Mapped Performance Criteria:', margin + 10, yPos);
                    yPos += 5;
                    
                    pdf.setFont(undefined, 'normal');
                    pdf.setFontSize(9);
                    lo.linkedCriteria.forEach(pc => {
                        if (yPos + 5 > pageHeight - margin) {
                            pdf.addPage();
                            yPos = margin + 5;
                        }
                        const pcText = `- ${pc.id}: ${pc.text}`;
                        const pcLines = pdf.splitTextToSize(pcText, pageWidth - 2 * margin - 15);
                        pcLines.forEach(line => {
                            if (yPos + 5 > pageHeight - margin) {
                                pdf.addPage();
                                yPos = margin + 5;
                            }
                            pdf.text(line, margin + 15, yPos);
                            yPos += 5;
                        });
                    });
                    
                    yPos += 5;
                });
                
                yPos += 3;
            });
        }
        
        // ============ MODULE MAPPING SECTION ============
        if (appState.moduleMappingData.modules && appState.moduleMappingData.modules.length > 0) {
            pdf.addPage();
            yPos = margin + 5;
            
            pdf.setFontSize(16);
            pdf.setFont(undefined, 'bold');
            pdf.text('Module Mapping', pageWidth / 2, yPos, { align: 'center' });
            yPos += 10;
            
            appState.moduleMappingData.modules.forEach(module => {
                if (yPos + 20 > pageHeight - margin) {
                    pdf.addPage();
                    yPos = margin + 5;
                }
                
                // Module title
                pdf.setFontSize(14);
                pdf.setFont(undefined, 'bold');
                pdf.text(module.title, margin, yPos);
                yPos += 7;
                
                // Learning Outcomes in this module
                pdf.setFontSize(12);
                pdf.setFont(undefined, 'bold');
                pdf.text('Learning Outcomes:', margin + 5, yPos);
                yPos += 5;
                
                module.learningOutcomes.forEach(lo => {
                    if (yPos + 15 > pageHeight - margin) {
                        pdf.addPage();
                        yPos = margin + 5;
                    }
                    
                    pdf.setFontSize(11);
                    pdf.setFont(undefined, 'bold');
                    pdf.text(`${lo.number}:`, margin + 10, yPos);
                    yPos += 5;
                    
                    if (lo.statement && lo.statement.trim()) {
                        pdf.setFontSize(10);
                        pdf.setFont(undefined, 'normal');
                        const statementLines = pdf.splitTextToSize(lo.statement, pageWidth - 2 * margin - 15);
                        statementLines.forEach(line => {
                            if (yPos + 5 > pageHeight - margin) {
                                pdf.addPage();
                                yPos = margin + 5;
                            }
                            pdf.text(line, margin + 15, yPos);
                            yPos += 5;
                        });
                    }
                    
                    yPos += 2;
                    
                    // Referenced Performance Criteria
                    pdf.setFontSize(9);
                    pdf.setFont(undefined, 'italic');
                    pdf.text('Referenced PC:', margin + 15, yPos);
                    yPos += 4;
                    
                    pdf.setFont(undefined, 'normal');
                    pdf.setFontSize(8);
                    lo.linkedCriteria.forEach(pc => {
                        if (yPos + 4 > pageHeight - margin) {
                            pdf.addPage();
                            yPos = margin + 5;
                        }
                        const pcText = `- ${pc.id}: ${pc.text}`;
                        const pcLines = pdf.splitTextToSize(pcText, pageWidth - 2 * margin - 20);
                        pcLines.forEach(line => {
                            if (yPos + 4 > pageHeight - margin) {
                                pdf.addPage();
                                yPos = margin + 5;
                            }
                            pdf.text(line, margin + 20, yPos);
                            yPos += 4;
                        });
                    });
                    
                    yPos += 3;
                });
                
                yPos += 5;
            });
        }
        
        const _pdfOccSlug = occupationTitleInput.value.replace(/[^a-z0-9]/gi, '_');
        const _pdfJobVal  = (jobTitleInput.value || '').trim();
        const _pdfJobSlug = _pdfJobVal ? `_${_pdfJobVal.replace(/[^a-z0-9]/gi, '_')}` : '';
        pdf.save(`${_pdfOccSlug}${_pdfJobSlug}_DACUM_Chart.pdf`);
        showStatus('PDF exported successfully! ✓', 'success');
        
    } catch (error) {
        console.error('Error generating PDF:', error);
        showStatus('Error generating PDF: ' + error.message, 'error');
    }
}

