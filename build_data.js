/**
 * Build script to precompile Excel sheets into curriculum-data.js
 * Run whenever the Excel spreadsheets are updated:
 *   node build_data.js
 */

const XLSX = require('./xlsx.full.min.js');
const fs = require('fs');
const path = require('path');

function parseProjectNameAndDuration(rawStr) {
    if (!rawStr) return { name: '', duration: '', full: '' };
    const str = rawStr.toString().trim();
    const match = str.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
    if (match) {
        return {
            name: match[1].trim(),
            duration: match[2].trim(),
            full: str
        };
    }
    return { name: str, duration: '', full: str };
}

function build() {
    console.log('Building precompiled curriculum dataset...');

    const parsedActivities = [];
    const excel1Path = path.join(__dirname, 'Adobe Express_Implementation plan.xlsx');

    if (fs.existsSync(excel1Path)) {
        const buf1 = fs.readFileSync(excel1Path);
        const workbook = XLSX.read(buf1, { type: 'buffer' });

        for (let s_idx = 0; s_idx < workbook.SheetNames.length; s_idx++) {
            const sheetName = workbook.SheetNames[s_idx];
            const cleanGradeName = sheetName.trim().replace(/\s+/g, ' ');
            if (!cleanGradeName || cleanGradeName === '') continue;

            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

            let activityNameRowIdx = -1;
            let skillsRowIdx = -1;
            let instructionsRowIdx = -1;
            let bookScreenshotRowIdx = -1;
            let templateLinkRowIdx = -1;
            let videoLinkRowIdx = -1;
            let descriptionRowIdx = -1;

            for (let r = 0; r < Math.min(rows.length, 20); r++) {
                const label = rows[r] && rows[r][0] ? rows[r][0].toString().trim().toLowerCase() : '';
                if (!label) continue;

                if (label.includes('activity name') || label.includes('गतिविधि का नाम') || label === 'गतिविधि') {
                    activityNameRowIdx = r;
                } else if (label.includes('skill') || label.includes('कौशल')) {
                    skillsRowIdx = r;
                } else if (label.includes('instruction') || label.includes('निर्देश')) {
                    instructionsRowIdx = r;
                } else if (label.includes('screenshot') || label.includes('book') || label.includes('स्क्रीनशॉट') || label.includes('पुस्तक')) {
                    bookScreenshotRowIdx = r;
                } else if (label.includes('activity link') || label.includes('project link') || label.includes('template link') || label.includes('गतिविधि का लिंक')) {
                    templateLinkRowIdx = r;
                } else if (label.includes('video') || label.includes('वीडियो')) {
                    videoLinkRowIdx = r;
                } else if (label.includes('description') || label.includes('विवरण')) {
                    descriptionRowIdx = r;
                }
            }

            if (activityNameRowIdx === -1) activityNameRowIdx = 5;
            if (skillsRowIdx === -1) skillsRowIdx = 6;
            if (instructionsRowIdx === -1) instructionsRowIdx = 7;
            if (templateLinkRowIdx === -1) {
                templateLinkRowIdx = (bookScreenshotRowIdx !== -1) ? 9 : 8;
            }

            const numCols = rows[activityNameRowIdx] ? rows[activityNameRowIdx].length : 0;
            const monthRow = rows[2] || [];
            const monthMap = {};
            let lastMonth = '';
            for (let c = 1; c < numCols; c++) {
                const val = monthRow[c] ? monthRow[c].toString().trim() : null;
                if (val && val !== '' &&
                    !val.toUpperCase().startsWith('GRADE') &&
                    !val.toUpperCase().startsWith('कक्षा')) {
                    lastMonth = val;
                }
                if (lastMonth) {
                    monthMap[c] = lastMonth;
                }
            }

            const weekRow = rows[4] || [];

            for (let c = 1; c < numCols; c++) {
                const activityName = rows[activityNameRowIdx] ? rows[activityNameRowIdx][c] : null;
                if (!activityName || activityName.toString().trim() === '' || 
                    activityName.toString().trim() === 'Activity Name' ||
                    activityName.toString().trim() === 'गतिविधि का नाम') {
                    continue;
                }

                const currentMonth = monthMap[c] || 'July';
                let currentWeek = '';
                const weekVal = weekRow[c] ? weekRow[c].toString().trim() : null;
                if (weekVal && weekVal !== '') {
                    currentWeek = weekVal;
                } else {
                    currentWeek = `Week ${((c - 1) % 4) + 1}`;
                }

                const skills = (rows[skillsRowIdx] && rows[skillsRowIdx][c]) ? rows[skillsRowIdx][c].toString().trim() : '';
                const instructions = (rows[instructionsRowIdx] && rows[instructionsRowIdx][c]) ? rows[instructionsRowIdx][c].toString().trim() : '';
                const bookScreenshot = (bookScreenshotRowIdx !== -1 && rows[bookScreenshotRowIdx] && rows[bookScreenshotRowIdx][c]) ? rows[bookScreenshotRowIdx][c].toString().trim() : '';
                const templateLink = (rows[templateLinkRowIdx] && rows[templateLinkRowIdx][c]) ? rows[templateLinkRowIdx][c].toString().trim() : '';
                const videoLink = (videoLinkRowIdx !== -1 && rows[videoLinkRowIdx] && rows[videoLinkRowIdx][c]) ? rows[videoLinkRowIdx][c].toString().trim() : '';
                const description = (descriptionRowIdx !== -1 && rows[descriptionRowIdx] && rows[descriptionRowIdx][c]) ? rows[descriptionRowIdx][c].toString().trim() : '';

                parsedActivities.push({
                    grade: cleanGradeName,
                    activity_name: activityName.toString().trim(),
                    skills: skills,
                    instructions: instructions,
                    book_screenshot: bookScreenshot,
                    template_link: templateLink,
                    video_link: videoLink,
                    description: description,
                    month: currentMonth,
                    week: currentWeek
                });
            }
        }
        console.log(`Parsed ${parsedActivities.length} activities from Adobe Express_Implementation plan.xlsx`);
    } else {
        console.warn('Adobe Express_Implementation plan.xlsx not found.');
    }

    const kbActivities = [];
    const excel2Path = path.join(__dirname, 'final of Kaushal Bodh Implementation Table.xlsx');

    if (fs.existsSync(excel2Path)) {
        const buf2 = fs.readFileSync(excel2Path);
        const workbook2 = XLSX.read(buf2, { type: 'buffer' });

        for (let s_idx = 0; s_idx < workbook2.SheetNames.length; s_idx++) {
            const rawSheetName = workbook2.SheetNames[s_idx];
            if (!rawSheetName || !rawSheetName.trim()) continue;

            const sheetName = rawSheetName.trim();
            let gradeName = sheetName;
            const matchGrade = sheetName.match(/grade\s*(\d+)/i);
            if (matchGrade) {
                gradeName = `Grade ${matchGrade[1]}`;
            }

            const sheet = workbook2.Sheets[rawSheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

            rows.forEach((row) => {
                const formOfWork = (row['Form of Work'] || row['Form Of Work'] || '').toString().trim();
                const projectNameDuration = (row['Project Name & Duration'] || row['Project Name and Duration'] || '').toString().trim();
                const keyActivities = (row['Key Activities'] || '').toString().trim();
                const keyLearningOutcomes = (row['Key Learning Outcomes'] || '').toString().trim();
                const adobeExpressIntegration = (row['Adobe Express Integration for Activity Books & Portfolios'] || row['Adobe Express Integration'] || '').toString().trim();
                const adobeExpressActivity = (row['Adobe Express Activity'] || '').toString().trim();
                const finalStudentDeliverable = (row['Final Student Deliverable'] || '').toString().trim();
                const templateLink = (row['Template Link'] || '').toString().trim();
                const learningJournalLink = (row['Learning Journal Link'] || '').toString().trim();

                if (!projectNameDuration && !formOfWork) return;

                const parsed = parseProjectNameAndDuration(projectNameDuration);

                kbActivities.push({
                    isKaushalBodh: true,
                    grade: gradeName,
                    form_of_work: formOfWork,
                    project_name_duration: parsed.full || projectNameDuration,
                    activity_name: parsed.name,
                    duration: parsed.duration,
                    skills: keyLearningOutcomes,
                    instructions: keyActivities,
                    key_activities: keyActivities,
                    key_learning_outcomes: keyLearningOutcomes,
                    adobe_express_integration: adobeExpressIntegration,
                    adobe_express_activity: adobeExpressActivity,
                    final_student_deliverable: finalStudentDeliverable,
                    template_link: templateLink,
                    learning_journal_link: learningJournalLink,
                    month: formOfWork || 'Kaushal Bodh',
                    week: ''
                });
            });
        }
        console.log(`Parsed ${kbActivities.length} activities from final of Kaushal Bodh Implementation Table.xlsx`);
    } else {
        console.warn('final of Kaushal Bodh Implementation Table.xlsx not found.');
    }

    const outputContent = `/** Pre-compiled curriculum data for instant zero-latency loading and offline resilience **/\nwindow.AIM_DEFAULT_DATA = ${JSON.stringify({ activities: parsedActivities, kaushalBodh: kbActivities })};\n`;
    const outputPath = path.join(__dirname, 'curriculum-data.js');
    fs.writeFileSync(outputPath, outputContent, 'utf8');
    console.log(`Successfully generated ${outputPath}`);
}

build();
