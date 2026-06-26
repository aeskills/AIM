/**
 * @fileoverview Main application logic for AIM - Adobe Implementation Plane Student Tutorial Platform.
 * Parses Excel files using SheetJS (XLSX) and manages the step routing, filtering,
 * metadata rendering, modal project submissions, and CSV download capabilities.
 */

'use strict';

/**
 * @typedef {Object} Activity
 * @property {string} grade
 * @property {string} activity_name
 * @property {string} skills
 * @property {string} instructions
 * @property {string} template_link
 * @property {string} video_link
 * @property {string} description
 * @property {string} month
 * @property {string} week
 */

/**
 * Central application state object.
 * @type {{
 *   allActivities: Activity[],
 *   selectedGrade: string|null,
 *   selectedActivity: Activity|null,
 *   currentStep: 'grade'|'activities'|'detail',
 *   filterSubject: string, // Represents filtered Month in Excel edition
 *   submissions: Object[]
 * }}
 */
const state = {
    allActivities: [],
    selectedGrade: null,
    selectedActivity: null,
    currentStep: 'grade',
    filterSubject: 'All', // Mapped to Month filtering
    submissions: [],
    language: 'en', // 'en' or 'hi'
    tutorialMode: 'choose' // 'choose' | 'hindi' | 'english'
};

const translations = {
    en: {
        readyToExplore: "Let's begin!",
        heroSubtitle: "Choose your grade level and proceed with the activities",
        selectYourGrade: "Select Your Grade",
        searchGradesPlaceholder: "Search for your grade...",
        noGradesMatch: "No grades match your search query.",
        exploreGrade: "Explore Grade",
        exploreActivities: "Explore Activities",
        showingActivitiesFor: "Showing activities for {grade}",
        monthLabel: "Month",
        allMonthsOption: "All Months",
        noActivitiesFound: "No Activities Found",
        emptyActivitiesDesc: "Try clearing your filters or choosing another grade level to see what is available.",
        resetFilters: "Reset Filters",
        backToHome: "Back to Home",
        backToWelcome: "Back to Language Selection",
        backToActivities: "Back to activities",
        descriptionTitle: "Description",
        skillsTitle: "Skills in Focus",
        instructionsTitle: "Instructions",
        submitYourWork: "Submit Your Work",
        submitWorkPrompt: "Completed this activity? Paste your project link to submit your work!",
        submitProjectBtn: "Submit Project Link",
        submitProjectTitle: "Submit Your Project",
        submitProjectSubtitle: "Fill in the details below to submit your project link.",
        fullNameLabel: "Full Name",
        schoolNameLabel: "School Name",
        projectLinkLabel: "Project/Work Link",
        submitBtnText: "Submit Project",
        cancelBtnText: "Cancel",
        confirmTitle: "Do you want to submit the work?",
        confirmSubtitle: "You will be guided to the submission form to enter your details and project link.",
        confirmYes: "Yes",
        confirmNo: "No",
        successTitle: "Project Submitted!",
        successSubtitle: "Your work has been submitted successfully. Thank you for sharing your project!",
        successOk: "Awesome",
        enterAdminPassword: "Enter Admin Password",
        passwordPrompt: "Please enter the password to download the student submissions CSV.",
        unlockAndDownload: "Unlock & Download",
        incorrectPassword: "Incorrect password. Please try again.",
        downloadSubmissions: "Download Submissions ({count})",
        studentNamePlaceholder: "e.g. Ashish Dokwal",
        schoolNamePlaceholder: "e.g. Delhi Public School",
        projectLinkPlaceholder: "e.g. https://new.express.adobe.com/...",
        passwordInputPlaceholder: "Password",
        footerText: "© 2026 AIM - Adobe Implementation Plane. Crafted for young explorers.",
        failedLoadActivities: "Failed to Load Activities",
        failedLoadDesc: "The activities database could not be loaded or is in an invalid format.",
        tryAgain: "Try Again",
        homeBreadcrumb: "Home",
        gradeBreadcrumb: "Grades",
        activitiesBreadcrumb: "Activities",
        detailBreadcrumb: "Detail",
        activitySingle: "1 activity",
        activitiesPlural: "{count} activities",
        activityTemplate: "📝 Activity Template",
        tutorialVideo: "🎥 Activity Tutorial video",
        templateUnavailable: "🔒 Template link unavailable",
        videoUnavailable: "🔒 Tutorial video link unavailable",
        noDescription: "No description available.",
        notSpecified: "Not specified.",
        noInstructions: "No instructions provided.",
        viewActivity: "View Activity",
        curriculum: "Curriculum",
        monthNames: {
            "All": "All Months",
            "April": "April",
            "May": "May",
            "June": "June",
            "July": "July",
            "August": "August",
            "September": "September",
            "October": "October",
            "November": "November",
            "December": "December",
            "January": "January",
            "February": "February",
            "March": "March"
        }
    },
    hi: {
        readyToExplore: "आइए शुरू करें!",
        heroSubtitle: "अपनी कक्षा चुनें और गतिविधियों के साथ आगे बढ़ें",
        selectYourGrade: "अपनी कक्षा चुनें",
        searchGradesPlaceholder: "अपनी कक्षा खोजें...",
        noGradesMatch: "आपकी खोज से मेल खाती कोई कक्षा नहीं मिली।",
        exploreGrade: "कक्षा देखें",
        exploreActivities: "गतिविधियां देखें",
        showingActivitiesFor: "{grade} के लिए गतिविधियां",
        monthLabel: "महीना",
        allMonthsOption: "सभी महीने",
        noActivitiesFound: "कोई गतिविधि नहीं मिली",
        emptyActivitiesDesc: "उपलब्ध गतिविधियों को देखने के लिए फ़िल्टर बदलें या कोई अन्य कक्षा चुनें।",
        resetFilters: "फ़िल्टर रीसेट करें",
        backToHome: "मुख्य पृष्ठ पर जाएं",
        backToWelcome: "भाषा चयन पर वापस जाएं",
        backToActivities: "गतिविधियों पर वापस जाएं",
        descriptionTitle: "विवरण",
        skillsTitle: "मुख्य कौशल",
        instructionsTitle: "निर्देश",
        submitYourWork: "अपना काम जमा करें",
        submitWorkPrompt: "क्या आपने यह गतिविधि पूरी कर ली है? अपना प्रोजेक्ट जमा करने के लिए लिंक पेस्ट करें!",
        submitProjectBtn: "प्रोजेक्ट लिंक जमा करें",
        submitProjectTitle: "अपना प्रोजेक्ट जमा करें",
        submitProjectSubtitle: "अपना प्रोजेक्ट लिंक जमा करने के लिए नीचे विवरण भरें।",
        fullNameLabel: "पूरा नाम",
        schoolNameLabel: "स्कूल का नाम",
        projectLinkLabel: "प्रोजेक्ट / कार्य का लिंक",
        submitBtnText: "प्रोजेक्ट जमा करें",
        cancelBtnText: "रद्द करें",
        confirmTitle: "क्या आप काम जमा करना चाहते हैं?",
        confirmSubtitle: "प्रोजेक्ट लिंक और विवरण दर्ज करने के लिए आपको सबमिशन फॉर्म पर ले जाया जाएगा।",
        confirmYes: "हाँ",
        confirmNo: "नहीं",
        successTitle: "प्रोजेक्ट जमा हो गया!",
        successSubtitle: "आपका कार्य सफलतापूर्वक जमा कर दिया गया है। अपना प्रोजेक्ट साझा करने के लिए धन्यवाद!",
        successOk: "बहुत बढ़िया",
        enterAdminPassword: "एडमिन पासवर्ड दर्ज करें",
        passwordPrompt: "छात्र सबमिशन CSV डाउनलोड करने के लिए कृपया पासवर्ड दर्ज करें।",
        unlockAndDownload: "अनलॉक और डाउनलोड करें",
        incorrectPassword: "गलत पासवर्ड। कृपया पुनः प्रयास करें।",
        downloadSubmissions: "सबमिशन डाउनलोड करें ({count})",
        studentNamePlaceholder: "जैसे: आशीष डोकवाल",
        schoolNamePlaceholder: "जैसे: दिल्ली पब्लिक स्कूल",
        projectLinkPlaceholder: "जैसे: https://new.express.adobe.com/...",
        passwordInputPlaceholder: "पासवर्ड",
        footerText: "© 2026 AIM - Adobe Implementation Plane. युवा खोजकर्ताओं के लिए निर्मित।",
        failedLoadActivities: "गतिविधियां लोड करने में विफल",
        failedLoadDesc: "गतिविधि डेटाबेस लोड नहीं किया जा सका या यह अमान्य प्रारूप में है।",
        tryAgain: "पुनः प्रयास करें",
        homeBreadcrumb: "होम",
        gradeBreadcrumb: "कक्षाएं",
        activitiesBreadcrumb: "गतिविधियां",
        detailBreadcrumb: "विवरण",
        activitySingle: "1 गतिविधि",
        activitiesPlural: "{count} गतिविधियां",
        activityTemplate: "📝 गतिविधि टेम्पलेट",
        tutorialVideo: "🎥 गतिविधि ट्यूटोरियल वीडियो",
        templateUnavailable: "🔒 टेम्पलेट लिंक अनुपलब्ध",
        videoUnavailable: "🔒 ट्यूटोरियल वीडियो लिंक अनुपलब्ध",
        noDescription: "कोई विवरण उपलब्ध नहीं है।",
        notSpecified: "निर्दिष्ट नहीं है।",
        noInstructions: "कोई निर्देश प्रदान नहीं किया गया है।",
        viewActivity: "गतिविधि देखें",
        curriculum: "पाठ्यक्रम",
        monthNames: {
            "All": "सभी महीने",
            "April": "अप्रैल",
            "May": "मई",
            "June": "जून",
            "July": "जुलाई",
            "August": "अगस्त",
            "September": "सितंबर",
            "October": "अक्टूबर",
            "November": "नवंबर",
            "December": "दिसंबर",
            "January": "जनवरी",
            "February": "फरवरी",
            "March": "मार्च"
        }
    }
};

/**
 * Initializes the application. Loads submissions from LocalStorage,
 * binds events, and fetches the default curriculum spreadsheet.
 * @returns {Promise<void>}
 */
async function init() {
    // Wipe all existing submissions unconditionally for a clean slate
    localStorage.setItem('aim_submissions', JSON.stringify([]));
    state.submissions = [];
    
    detectMode();
    setupEventListeners();
    updateDownloadButtonState();
    await loadDefaultExcel();
}

/**
 * Detects whether we are in Hindi or English subfolder, hash, or choose welcome path.
 * @returns {void}
 */
function detectMode() {
    state.tutorialMode = 'choose';
    state.currentStep = 'welcome';
    if (window.location.hash) {
        window.location.hash = '';
    }
}

/**
 * Cleans the grade sheet name (e.g. "Grade 3 (hindi)") and localizes it.
 * @param {string} grade - The raw sheet name.
 * @returns {string} The formatted grade display name.
 */
function getDisplayGradeName(grade) {
    if (!grade) return "";
    let clean = grade.replace(/\(hindi\)/i, "").replace(/\(english\)/i, "").trim().replace(/\s+/g, " ");
    if (state.language === 'hi') {
        clean = clean.replace(/Grade/i, "कक्षा");
    }
    return clean;
}

/**
 * Binds DOM event listeners for navigation, filters, modals, 
 * submissions form, and downloads CSV button.
 * @returns {void}
 */
function setupEventListeners() {
    // Logo Click Navigation
    const logoLink = document.getElementById('header-logo');
    if (logoLink) {
        logoLink.addEventListener('click', (e) => {
            e.preventDefault();
            state.tutorialMode = 'choose';
            resetState();
        });
    }

    // Welcome Choices Click Events (SPA fallback for instant routing)
    const choiceEnglish = document.getElementById('choice-english');
    if (choiceEnglish) {
        choiceEnglish.addEventListener('click', (e) => {
            e.preventDefault();
            state.tutorialMode = 'english';
            state.language = 'en';
            state.currentStep = 'grade';
            render();
        });
    }

    const choiceHindi = document.getElementById('choice-hindi');
    if (choiceHindi) {
        choiceHindi.addEventListener('click', (e) => {
            e.preventDefault();
            state.tutorialMode = 'hindi';
            state.language = 'hi';
            state.currentStep = 'grade';
            render();
        });
    }

    // Back to Welcome Choice Link
    const backToWelcomeLink = document.getElementById('back-to-welcome-link');
    if (backToWelcomeLink) {
        backToWelcomeLink.addEventListener('click', (e) => {
            e.preventDefault();
            state.tutorialMode = 'choose';
            state.currentStep = 'welcome';
            render();
        });
    }

    // Retry Button Click
    const retryBtn = document.getElementById('retry-btn');
    if (retryBtn) {
        retryBtn.addEventListener('click', () => {
            loadDefaultExcel();
        });
    }

    // Month filter dropdown selection
    const subjectFilter = document.getElementById('subject-filter');
    if (subjectFilter) {
        subjectFilter.addEventListener('change', (e) => {
            state.filterSubject = e.target.value;
            renderFilteredActivities();
        });
    }

    // Back to Activities link
    const backLink = document.getElementById('back-to-activities-link');
    if (backLink) {
        backLink.addEventListener('click', (e) => {
            e.preventDefault();
            state.selectedActivity = null;
            state.currentStep = 'activities';
            render();
        });
    }

    // Back to Home (Grades selection) link
    const backToGradesLink = document.getElementById('back-to-grades-link');
    if (backToGradesLink) {
        backToGradesLink.addEventListener('click', (e) => {
            e.preventDefault();
            resetState();
        });
    }

    // Submissions popup modal controls
    const showSubmissionFormBtn = document.getElementById('show-submission-form-btn');
    if (showSubmissionFormBtn) {
        showSubmissionFormBtn.addEventListener('click', () => {
            openConfirmModal();
        });
    }

    // Custom Confirm Modal button triggers
    const confirmYesBtn = document.getElementById('confirm-yes-btn');
    if (confirmYesBtn) {
        confirmYesBtn.addEventListener('click', () => {
            closeConfirmModal();
            openSubmissionModal();
        });
    }

    const confirmNoBtn = document.getElementById('confirm-no-btn');
    if (confirmNoBtn) {
        confirmNoBtn.addEventListener('click', () => {
            closeConfirmModal();
        });
    }

    const closeConfirmModalBtn = document.getElementById('close-confirm-modal-btn');
    if (closeConfirmModalBtn) {
        closeConfirmModalBtn.addEventListener('click', () => {
            closeConfirmModal();
        });
    }

    // Click on overlay background closes confirm modal
    const confirmModal = document.getElementById('confirm-modal');
    if (confirmModal) {
        confirmModal.addEventListener('click', (e) => {
            if (e.target === confirmModal) {
                closeConfirmModal();
            }
        });
    }

    // Close Modal button click
    const closeModalBtn = document.getElementById('close-modal-btn');
    if (closeModalBtn) {
        closeModalBtn.addEventListener('click', () => {
            closeSubmissionModal();
        });
    }

    // Cancel submission click inside modal
    const cancelSubmissionBtn = document.getElementById('cancel-submission-btn');
    if (cancelSubmissionBtn) {
        cancelSubmissionBtn.addEventListener('click', () => {
            closeSubmissionModal();
        });
    }

    // Click on overlay background closes modal
    const modalOverlay = document.getElementById('submission-modal');
    if (modalOverlay) {
        modalOverlay.addEventListener('click', (e) => {
            if (e.target === modalOverlay) {
                closeSubmissionModal();
            }
        });
    }

    // Custom Success Modal button triggers
    const successOkBtn = document.getElementById('success-ok-btn');
    if (successOkBtn) {
        successOkBtn.addEventListener('click', () => {
            closeSuccessModal();
        });
    }

    const successModal = document.getElementById('success-modal');
    if (successModal) {
        successModal.addEventListener('click', (e) => {
            if (e.target === successModal) {
                closeSuccessModal();
            }
        });
    }

    // Handle student submission form submit event inside modal
    const submissionForm = document.getElementById('submission-form');
    if (submissionForm) {
        submissionForm.addEventListener('submit', (e) => {
            e.preventDefault();
            handleStudentSubmission();
        });
    }

    // Download Submissions CSV button (unlocked via Password Modal)
    const downloadBtn = document.getElementById('download-submissions-btn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', () => {
            openPasswordModal();
        });
    }

    // Custom Password Modal handlers
    const passwordForm = document.getElementById('password-form');
    if (passwordForm) {
        passwordForm.addEventListener('submit', (e) => {
            e.preventDefault();
            handlePasswordVerification();
        });
    }

    const cancelPasswordBtn = document.getElementById('cancel-password-btn');
    if (cancelPasswordBtn) {
        cancelPasswordBtn.addEventListener('click', () => {
            closePasswordModal();
        });
    }

    const closePasswordModalBtn = document.getElementById('close-password-modal-btn');
    if (closePasswordModalBtn) {
        closePasswordModalBtn.addEventListener('click', () => {
            closePasswordModal();
        });
    }

    const passwordModal = document.getElementById('password-modal');
    if (passwordModal) {
        passwordModal.addEventListener('click', (e) => {
            if (e.target === passwordModal) {
                closePasswordModal();
            }
        });
    }

    // Grade Search Input listeners
    const gradeSearchInput = document.getElementById('grade-search');
    const clearSearchBtn = document.getElementById('clear-search-btn');
    if (gradeSearchInput) {
        gradeSearchInput.addEventListener('input', (e) => {
            const query = e.target.value.trim().toLowerCase();
            if (clearSearchBtn) {
                if (query.length > 0) {
                    clearSearchBtn.classList.remove('hidden');
                } else {
                    clearSearchBtn.classList.add('hidden');
                }
            }
            renderGradeSelection();
        });
    }

    if (clearSearchBtn && gradeSearchInput) {
        clearSearchBtn.addEventListener('click', () => {
            gradeSearchInput.value = '';
            clearSearchBtn.classList.add('hidden');
            renderGradeSelection();
            gradeSearchInput.focus();
        });
    }



    // Listen for hash changes to support back/forward browser navigation
    window.addEventListener('hashchange', () => {
        if (window.location.hash) {
            window.location.hash = '';
        }
    });

    // Activity Card Tooltip Handlers
    document.addEventListener('mouseover', (e) => {
        const card = e.target.closest('.activity-card');
        if (!card) return;

        const tooltip = document.getElementById('activity-tooltip');
        if (!tooltip) return;

        const title = decodeURIComponent(card.getAttribute('data-tooltip-title') || '');

        tooltip.querySelector('.tooltip-title').innerText = title;

        tooltip.classList.remove('hidden');
        // Force reflow
        tooltip.offsetHeight;
        tooltip.classList.add('visible');

        positionTooltip(e, tooltip);
    });

    document.addEventListener('mouseout', (e) => {
        const card = e.target.closest('.activity-card');
        if (!card) return;

        const tooltip = document.getElementById('activity-tooltip');
        if (tooltip) {
            tooltip.classList.remove('visible');
            setTimeout(() => {
                if (!tooltip.classList.contains('visible')) {
                    tooltip.classList.add('hidden');
                }
            }, 150);
        }
    });

    document.addEventListener('mousemove', (e) => {
        const card = e.target.closest('.activity-card');
        if (!card) return;

        const tooltip = document.getElementById('activity-tooltip');
        if (tooltip && tooltip.classList.contains('visible')) {
            positionTooltip(e, tooltip);
        }
    });
}

/**
 * Positions the floating activity tooltip relative to the cursor with screen boundary safety.
 * @param {MouseEvent} e - Mouse event.
 * @param {HTMLElement} tooltip - Tooltip DOM element.
 * @returns {void}
 */
function positionTooltip(e, tooltip) {
    const mouseX = e.pageX;
    const mouseY = e.pageY;
    const tooltipWidth = tooltip.offsetWidth;
    const tooltipHeight = tooltip.offsetHeight;

    // Offset from cursor
    let left = mouseX + 15;
    let top = mouseY + 15;

    // Boundary checks (right and bottom edges)
    if (left + tooltipWidth > window.innerWidth + window.pageXOffset - 20) {
        left = mouseX - tooltipWidth - 15;
    }
    if (top + tooltipHeight > window.innerHeight + window.pageYOffset - 20) {
        top = mouseY - tooltipHeight - 15;
    }

    // Boundary checks (left and top edges)
    if (left < window.pageXOffset + 10) {
        left = window.pageXOffset + 10;
    }
    if (top < window.pageYOffset + 10) {
        top = window.pageYOffset + 10;
    }

    tooltip.style.left = left + 'px';
    tooltip.style.top = top + 'px';
}

/**
 * Fetches the default Excel curriculum sheet 'implemetation Plane.xlsx' from the server.
 * @returns {Promise<void>}
 */
async function loadDefaultExcel() {
    showLoader();
    try {
        const fetchUrl = 'implemetation Plane.xlsx?v=' + Date.now();
        const response = await fetch(fetchUrl);
        if (!response.ok) {
            throw new Error(`Server returned status: ${response.status}`);
        }
        const arrayBuffer = await response.arrayBuffer();
        parseExcel(arrayBuffer);
    } catch (error) {
        showError('Could not auto-load the curriculum excel file. Please make sure implemetation Plane.xlsx is in the root folder.');
    }
}

/**
 * Parses Excel workbook contents using SheetJS (XLSX), extracts columns,
 * fills merged values, maps row entries, and builds the state object.
 * @param {ArrayBuffer} arrayBuffer - Raw XLSX array buffer.
 * @returns {void}
 */
function parseExcel(arrayBuffer) {
    if (typeof XLSX === 'undefined') {
        showError('XLSX library failed to load. Please check your internet connection.');
        return;
    }

    try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const parsedActivities = [];

        // Parse sheets dynamically, skipping empty name sheets
        for (let s_idx = 0; s_idx < workbook.SheetNames.length; s_idx++) {
            const sheetName = workbook.SheetNames[s_idx];
            const cleanGradeName = sheetName.trim().replace(/\s+/g, ' ');

            // If the sheet name is empty, skip it.
            // This filters out the duplicate sheet at index 1 and keeps only Grades 5, 7, 8
            if (!cleanGradeName || cleanGradeName === "") {
                continue;
            }
            
            const sheet = workbook.Sheets[sheetName];
            
            // Format sheet cells into a 2D grid matrix
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
            
            // Dynamically locate rows based on the labels in Column A (index 0)
            let activityNameRowIdx = -1;
            let skillsRowIdx = -1;
            let instructionsRowIdx = -1;
            let templateLinkRowIdx = -1;
            let videoLinkRowIdx = -1;
            let descriptionRowIdx = -1;

            for (let r = 0; r < rows.length; r++) {
                const label = rows[r] && rows[r][0] ? rows[r][0].toString().trim().toLowerCase() : "";
                if (!label) continue;

                if (label.includes("activity name")) {
                    activityNameRowIdx = r;
                } else if (label.includes("skill")) {
                    skillsRowIdx = r;
                } else if (label.includes("instruction")) {
                    instructionsRowIdx = r;
                } else if (label.includes("activity link") || label.includes("project link") || label.includes("template link")) {
                    templateLinkRowIdx = r;
                } else if (label.includes("video")) {
                    videoLinkRowIdx = r;
                } else if (label.includes("description")) {
                    descriptionRowIdx = r;
                }
            }

            // Fallbacks to previous default offsets if label is not explicitly found
            if (activityNameRowIdx === -1) activityNameRowIdx = 5;
            if (skillsRowIdx === -1) skillsRowIdx = 6;
            if (instructionsRowIdx === -1) instructionsRowIdx = 7;
            if (templateLinkRowIdx === -1) {
                templateLinkRowIdx = (s_idx > 0) ? 9 : 8;
            }
            if (videoLinkRowIdx === -1) {
                videoLinkRowIdx = (s_idx > 0) ? 11 : 10;
            }
            if (descriptionRowIdx === -1) {
                descriptionRowIdx = (s_idx > 0) ? 12 : 11;
            }

            const numCols = rows[activityNameRowIdx] ? rows[activityNameRowIdx].length : 0;
            
            let currentMonth = "";
            let currentWeek = "";

            for (let c = 1; c < numCols; c++) {
                const activityName = rows[activityNameRowIdx] ? rows[activityNameRowIdx][c] : null;
                // Exclude empty and header label cells
                if (!activityName || activityName.trim() === "" || activityName.trim() === "Activity Name") {
                    continue;
                }

                // Month header (Row index 2) - Fill down merged cell value
                const monthVal = (rows[2] && rows[2][c]) ? rows[2][c].toString().trim() : null;
                if (monthVal && monthVal !== "" && !monthVal.toUpperCase().startsWith("GRADE")) {
                    currentMonth = monthVal;
                }

                // Week header (Row index 4) - Fill down cell value
                const weekVal = (rows[4] && rows[4][c]) ? rows[4][c].toString().trim() : null;
                if (weekVal && weekVal !== "") {
                    currentWeek = weekVal;
                } else {
                    currentWeek = `Week ${((c - 1) % 4) + 1}`;
                }

                const skills = (rows[skillsRowIdx] && rows[skillsRowIdx][c]) ? rows[skillsRowIdx][c].toString().trim() : "";
                const instructions = (rows[instructionsRowIdx] && rows[instructionsRowIdx][c]) ? rows[instructionsRowIdx][c].toString().trim() : "";
                const templateLink = (rows[templateLinkRowIdx] && rows[templateLinkRowIdx][c]) ? rows[templateLinkRowIdx][c].toString().trim() : "";
                const videoLink = (rows[videoLinkRowIdx] && rows[videoLinkRowIdx][c]) ? rows[videoLinkRowIdx][c].toString().trim() : "";
                const description = (rows[descriptionRowIdx] && rows[descriptionRowIdx][c]) ? rows[descriptionRowIdx][c].toString().trim() : "";

                parsedActivities.push({
                    grade: cleanGradeName, // Grade 5, Grade 7, Grade 8
                    activity_name: activityName.trim(),
                    skills: skills,
                    instructions: instructions,
                    template_link: templateLink,
                    video_link: videoLink,
                    description: description,
                    month: currentMonth || "April",
                    week: currentWeek
                });
            }
        }

        if (parsedActivities.length === 0) {
            showError('Could not read any valid activity rows from spreadsheet.');
            return;
        }

        state.allActivities = parsedActivities;
        hideError();
        hideLoader();
        resetState();
    } catch (err) {
        console.error(err);
        showError('Error reading the Excel workbook. Please upload a valid sheet.');
    }
}

/**
 * Extracts, de-duplicates, and naturally sorts all unique grade labels from activities list.
 * @param {Activity[]} activities - List of parsed activities.
 * @returns {string[]} Naturally sorted array of grade strings.
 */
function extractGrades(activities) {
    const gradesSet = new Set(activities.map(a => a.grade).filter(Boolean));
    let gradesArray = Array.from(gradesSet);

    // Filter grades based on the active tutorialMode
    if (state.tutorialMode === 'hindi') {
        gradesArray = gradesArray.filter(g => g.toLowerCase().includes('(hindi)'));
    } else if (state.tutorialMode === 'english') {
        gradesArray = gradesArray.filter(g => g.toLowerCase().includes('(english)'));
    }

    return gradesArray.sort((a, b) => {
        const numA = parseInt(a.replace(/\D/g, ''), 10);
        const numB = parseInt(b.replace(/\D/g, ''), 10);
        if (!isNaN(numA) && !isNaN(numB)) {
            return numA - numB;
        }
        return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    });
}

/**
 * Clears current selections, resets filters to default, and updates active layout.
 * @returns {void}
 */
function resetState() {
    state.selectedGrade = null;
    state.selectedActivity = null;
    state.filterSubject = 'All';
    
    // Clear grade search input and clear button if they exist
    const gradeSearchInput = document.getElementById('grade-search');
    const clearSearchBtn = document.getElementById('clear-search-btn');
    if (gradeSearchInput) {
        gradeSearchInput.value = '';
    }
    if (clearSearchBtn) {
        clearSearchBtn.classList.add('hidden');
    }

    if (state.tutorialMode === 'choose') {
        state.currentStep = 'welcome';
    } else {
        state.currentStep = 'grade';
    }
    render();
}

/**
 * Resets Month filter dropdown.
 * @returns {void}
 */
function resetFilters() {
    state.filterSubject = 'All';

    const subjectFilter = document.getElementById('subject-filter');
    if (subjectFilter) {
        subjectFilter.value = 'All';
    }

    renderFilteredActivities();
}

/**
 * Deterministically returns color helper class based on string hash.
 * Mapped to Months in the Excel curriculum edition.
 * @param {string} subject - Name of the month.
 * @returns {string} The CSS color mapping class name.
 */
function getSubjectColorClass(subject) {
    if (!subject) return 'subject-color-default';
    let hash = 0;
    const cleanSubject = subject.trim().toLowerCase();
    for (let i = 0; i < cleanSubject.length; i++) {
        hash = cleanSubject.charCodeAt(i) + ((hash << 5) - hash);
    }
    const index = (Math.abs(hash) % 6) + 1;
    return `subject-color-${index}`;
}

/**
 * Coordinates current step layout configurations.
 * Handles DOM visibility, filters populating, and keyboard focus changes.
 * @returns {void}
 */
function render() {
    document.getElementById('step-welcome').classList.add('hidden');
    document.getElementById('step-grade').classList.add('hidden');
    document.getElementById('step-activities').classList.add('hidden');
    document.getElementById('step-detail').classList.add('hidden');

    let focusId = '';

    const isChoose = state.tutorialMode === 'choose';
    const breadcrumbNav = document.querySelector('.breadcrumb-nav');

    if (breadcrumbNav) {
        if (isChoose) {
            breadcrumbNav.classList.add('hidden');
        } else {
            breadcrumbNav.classList.remove('hidden');
        }
    }

    if (state.currentStep === 'welcome') {
        document.getElementById('step-welcome').classList.remove('hidden');
        focusId = '';
    } else if (state.currentStep === 'grade') {
        document.getElementById('step-grade').classList.remove('hidden');
        renderGradeSelection();
        focusId = 'grade-title';
    } else if (state.currentStep === 'activities') {
        document.getElementById('step-activities').classList.remove('hidden');

        const gradeActivities = state.allActivities.filter(a => a.grade === state.selectedGrade);
        populateMonthDropdown(gradeActivities);

        // Synchronize month selection
        const subjectFilter = document.getElementById('subject-filter');
        if (subjectFilter) {
            subjectFilter.value = state.filterSubject;
        }

        renderFilteredActivities();
        const subtitleTemplate = translations[state.language || 'en'].showingActivitiesFor;
        document.getElementById('activities-subtitle').innerText = subtitleTemplate.replace('{grade}', getDisplayGradeName(state.selectedGrade));
        focusId = 'activities-title';
    } else if (state.currentStep === 'detail') {
        document.getElementById('step-detail').classList.remove('hidden');
        renderActivityDetail();
        focusId = 'detail-title';
    }

    renderBreadcrumbs();

    // Translate all strings in the document DOM
    translateUIInternal();

    // Fire focus updates for screen readers
    if (focusId) {
        const focusEl = document.getElementById(focusId);
        if (focusEl) {
            setTimeout(() => {
                focusEl.focus();
            }, 50);
        }
    }
}

/**
 * Populates navigation links inside header.
 * @returns {void}
 */
function renderBreadcrumbs() {
    const breadcrumbsList = document.getElementById('breadcrumbs-list');
    if (!breadcrumbsList) return;

    if (state.allActivities.length === 0) {
        breadcrumbsList.innerHTML = '';
        return;
    }

    let html = '';

    // Level 1: Home
    const homeText = translations[state.language || 'en'].homeBreadcrumb;
    if (state.currentStep === 'grade') {
        html += `<li><span class="current">${homeText}</span></li>`;
    } else {
        html += `<li><a href="#" id="breadcrumb-home">${homeText}</a></li>`;
    }

    // Level 2: Grade Selected
    if (state.selectedGrade) {
        const displayGrade = getDisplayGradeName(state.selectedGrade);
        if (state.currentStep === 'activities') {
            html += `<li><span class="current">${displayGrade}</span></li>`;
        } else {
            html += `<li><a href="#" id="breadcrumb-grade">${displayGrade}</a></li>`;
        }
    }

    // Level 3: Activity Detail
    if (state.selectedActivity && state.currentStep === 'detail') {
        html += `<li><span class="current">${state.selectedActivity.activity_name}</span></li>`;
    }

    breadcrumbsList.innerHTML = html;

    // Click triggers
    const homeBtn = document.getElementById('breadcrumb-home');
    if (homeBtn) {
        homeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            resetState();
        });
    }

    const gradeBtn = document.getElementById('breadcrumb-grade');
    if (gradeBtn) {
        gradeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            state.selectedActivity = null;
            state.currentStep = 'activities';
            render();
        });
    }
}

/**
 * Creates Grade selection buttons in Step 1.
 * @returns {void}
 */
function renderGradeSelection() {
    const gradesGrid = document.getElementById('grades-grid');
    if (!gradesGrid) return;

    const searchInput = document.getElementById('grade-search');
    const query = searchInput ? searchInput.value.trim().toLowerCase() : '';

    let grades = extractGrades(state.allActivities);
    
    // Filter grades based on query
    if (query) {
        grades = grades.filter(grade => grade.toLowerCase().includes(query));
    }

    let html = '';

    if (grades.length === 0) {
        html = `
            <div class="no-results">
                <span class="no-results-icon" aria-hidden="true">🔍</span>
                <p data-i18n="noGradesMatch">No grades match your search query.</p>
            </div>
        `;
        gradesGrid.innerHTML = html;
        return;
    }

    grades.forEach(grade => {
        const count = state.allActivities.filter(a => a.grade === grade).length;
        const template = count === 1 ? translations[state.language || 'en'].activitySingle : translations[state.language || 'en'].activitiesPlural;
        const labelText = template.replace('{count}', count);

        const numStr = grade.replace(/\D/g, '');
        let icon = '🎓'; // Single unique graduation cap icon for all grades
        let accentClass = 'indigo';
        if (numStr === '5') {
            accentClass = 'pink';
        } else if (numStr === '7') {
            accentClass = 'green';
        } else if (numStr === '8') {
            accentClass = 'indigo';
        }

        html += `
            <button class="grade-card card-accent-${accentClass}" type="button" data-grade="${encodeURIComponent(grade)}" aria-label="${getDisplayGradeName(grade)}, ${labelText}">
                <div class="grade-icon-wrapper icon-bg-${accentClass}">
                    <span class="grade-icon-emoji">${icon}</span>
                </div>
                <h3 class="grade-card-title">${getDisplayGradeName(grade)}</h3>
                <p class="grade-card-subtitle">${labelText}</p>
                <div class="grade-card-footer">
                    <span class="explore-grade-text" data-i18n="exploreGrade">Explore Grade</span>
                    <span class="arrow-icon">→</span>
                </div>
            </button>
        `;
    });

    gradesGrid.innerHTML = html;

    const cards = gradesGrid.querySelectorAll('.grade-card');
    cards.forEach(card => {
        card.addEventListener('click', () => {
            const grade = decodeURIComponent(card.getAttribute('data-grade'));
            state.selectedGrade = grade;
            state.filterSubject = 'All';
            state.currentStep = 'activities';
            render();
        });
    });
}

/**
 * Populates Month dropdown filter dynamically based on Month values inside current grade.
 * @param {Activity[]} activities - Filtered curriculum list for active grade.
 * @returns {void}
 */
function populateMonthDropdown(activities) {
    const subjectFilter = document.getElementById('subject-filter');
    if (!subjectFilter) return;

    const currentSel = state.filterSubject;
    const months = Array.from(new Set(activities.map(a => a.month).filter(Boolean)));
    
    // Sort months chronologically
    const monthOrder = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    months.sort((a, b) => {
        const idxA = monthOrder.indexOf(a);
        const idxB = monthOrder.indexOf(b);
        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
        return a.localeCompare(b);
    });

    const lang = state.language || 'en';
    const dict = translations[lang];
    const allMonthsText = dict.allMonthsOption;

    let html = `<option value="All" data-i18n="allMonthsOption">${allMonthsText}</option>`;
    months.forEach(m => {
        const translatedName = dict.monthNames[m] || m;
        html += `<option value="${m}">${translatedName}</option>`;
    });

    subjectFilter.innerHTML = html;

    if (months.includes(currentSel)) {
        subjectFilter.value = currentSel;
        state.filterSubject = currentSel;
    } else {
        subjectFilter.value = 'All';
        state.filterSubject = 'All';
    }
}

/**
 * Filters and mounts Activity cards on Step 2.
 * Hides duration, difficulty dots, and renders rectangular highlighted description blocks.
 * @returns {void}
 */
function renderFilteredActivities() {
    const activitiesGrid = document.getElementById('activities-grid');
    const emptyView = document.getElementById('empty-activities-view');
    if (!activitiesGrid || !emptyView) return;

    let list = state.allActivities.filter(a => a.grade === state.selectedGrade);

    // Filter by month
    if (state.filterSubject !== 'All') {
        list = list.filter(a => a.month === state.filterSubject);
    }

    if (list.length === 0) {
        activitiesGrid.classList.add('hidden');
        emptyView.classList.remove('hidden');
    } else {
        emptyView.classList.add('hidden');
        activitiesGrid.classList.remove('hidden');

        let html = '';
        list.forEach(activity => {
            const colorClass = getSubjectColorClass(activity.month);
            const dict = translations[state.language || 'en'];
            const monthText = dict.monthNames[activity.month] || activity.month || dict.curriculum;

            html += `
                <button class="activity-card card-accent-${colorClass}" type="button" data-activity-name="${encodeURIComponent(activity.activity_name)}" data-tooltip-title="${encodeURIComponent(activity.activity_name)}" aria-label="${activity.activity_name}, ${monthText}">
                    <div class="activity-card-header">
                        <span class="subject-pill ${colorClass}">${monthText}</span>
                    </div>
                    <h2 class="activity-name">${activity.activity_name}</h2>
                    <div class="activity-card-footer">
                        <span class="view-activity-text" data-i18n="viewActivity">View Activity</span>
                        <span class="arrow-icon">→</span>
                    </div>
                </button>
            `;
        });

        activitiesGrid.innerHTML = html;

        const cards = activitiesGrid.querySelectorAll('.activity-card');
        cards.forEach(card => {
            card.addEventListener('click', () => {
                const name = decodeURIComponent(card.getAttribute('data-activity-name'));
                const selected = state.allActivities.find(
                    a => a.grade === state.selectedGrade && a.activity_name === name
                );
                if (selected) {
                    state.selectedActivity = selected;
                    state.currentStep = 'detail';
                    render();
                }
            });
        });
    }
}

/**
 * Populates detail page contents on Step 3.
 * Configures dual CTA buttons, skills text, instructions text, and hides duration/difficulty/details.
 * @returns {void}
 */
function renderActivityDetail() {
    const activity = state.selectedActivity;
    if (!activity) return;

    const dict = translations[state.language || 'en'];

    // Header pills
    const detailTags = document.getElementById('detail-tags');
    const colorClass = getSubjectColorClass(activity.month);
    const translatedMonth = dict.monthNames[activity.month] || activity.month || dict.curriculum;
    detailTags.innerHTML = `
        <span class="subject-pill ${colorClass}">${translatedMonth}</span>
        <span class="grade-badge">${getDisplayGradeName(activity.grade)}</span>
    `;

    // Title mapping
    document.getElementById('detail-title').innerText = activity.activity_name;

    // Description, skills, instructions text blocks
    document.getElementById('detail-description').innerText = activity.description || dict.noDescription;
    document.getElementById('detail-skills').innerText = activity.skills || dict.notSpecified;
    document.getElementById('detail-instructions').innerText = activity.instructions || dict.noInstructions;

    // CTA section: Activity Template (1) and Video Tutorial (2)
    const ctaContainer = document.getElementById('tutorial-cta-container');
    const templateLink = activity.template_link ? activity.template_link.trim() : '';
    const videoLink = activity.video_link ? activity.video_link.trim() : '';

    const isTemplateValid = templateLink && (templateLink.startsWith('http://') || templateLink.startsWith('https://'));
    const isVideoValid = videoLink && (videoLink.startsWith('http://') || videoLink.startsWith('https://'));

    let ctaHtml = '';

    // Button 1: Activity Template
    if (isTemplateValid) {
        ctaHtml += `
            <button id="template-cta-btn" class="btn btn-primary" type="button">
                ${dict.activityTemplate}
            </button>
        `;
    } else {
        ctaHtml += `
            <div class="no-tutorial-notice" role="status">
                ${dict.templateUnavailable}
            </div>
        `;
    }

    // Button 2: Activity Tutorial video
    if (isVideoValid) {
        ctaHtml += `
            <button id="video-cta-btn" class="btn btn-primary" type="button" style="background-color: var(--color-accent-success); box-shadow: 0 4px 12px rgba(16, 185, 129, 0.2);">
                ${dict.tutorialVideo}
            </button>
        `;
    } else {
        ctaHtml += `
            <div class="no-tutorial-notice" role="status">
                ${dict.videoUnavailable}
            </div>
        `;
    }

    ctaContainer.innerHTML = ctaHtml;

    // Attach click listeners to detail CTA buttons
    if (isTemplateValid) {
        document.getElementById('template-cta-btn').addEventListener('click', () => {
            window.open(templateLink, '_blank', 'noopener,noreferrer');
        });
    }
    if (isVideoValid) {
        document.getElementById('video-cta-btn').addEventListener('click', () => {
            window.open(videoLink, '_blank', 'noopener,noreferrer');
        });
    }
}

/**
 * Handles processing of a student work submission.
 * Saves values into LocalStorage state, closes the modal popup, and resets controls.
 * @returns {void}
 */
function handleStudentSubmission() {
    const nameInput = document.getElementById('submit-student-name');
    const schoolInput = document.getElementById('submit-school-name');
    const linkInput = document.getElementById('submit-work-link');

    const link = linkInput.value.trim();
    if (!link) {
        return; // Handled by HTML5 input required validation
    }

    const activity = state.selectedActivity;
    if (!activity) return;

    const newSubmission = {
        Timestamp: new Date().toLocaleString(),
        Grade: activity.grade,
        "Activity Name": activity.activity_name,
        "Student Name": nameInput.value.trim() || "Anonymous",
        "School Name": schoolInput.value.trim() || "Not Provided",
        "Work Link": link
    };

    state.submissions.push(newSubmission);
    localStorage.setItem('aim_submissions', JSON.stringify(state.submissions));

    updateDownloadButtonState();
    closeSubmissionModal();
    openSuccessModal();
}

/**
 * Shows the submissions popup overlay modal and clears inputs.
 * @returns {void}
 */
function openSubmissionModal() {
    const modal = document.getElementById('submission-modal');
    if (modal) {
        modal.classList.remove('hidden');
        
        // Reset and focus
        document.getElementById('submit-student-name').value = '';
        document.getElementById('submit-school-name').value = '';
        document.getElementById('submit-work-link').value = '';
        document.getElementById('submit-student-name').focus();
    }
}

/**
 * Dismisses the submissions popup overlay modal.
 * @returns {void}
 */
function closeSubmissionModal() {
    const modal = document.getElementById('submission-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

/**
 * Shows the custom confirmation modal.
 * @returns {void}
 */
function openConfirmModal() {
    const modal = document.getElementById('confirm-modal');
    if (modal) {
        modal.classList.remove('hidden');
        const yesBtn = document.getElementById('confirm-yes-btn');
        if (yesBtn) yesBtn.focus();
    }
}

/**
 * Dismisses the custom confirmation modal.
 * @returns {void}
 */
function closeConfirmModal() {
    const modal = document.getElementById('confirm-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

/**
 * Shows the custom success modal.
 * @returns {void}
 */
function openSuccessModal() {
    const modal = document.getElementById('success-modal');
    if (modal) {
        modal.classList.remove('hidden');
        const okBtn = document.getElementById('success-ok-btn');
        if (okBtn) okBtn.focus();
    }
}

/**
 * Dismisses the custom success modal.
 * @returns {void}
 */
function closeSuccessModal() {
    const modal = document.getElementById('success-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

/**
 * Shows the custom password authentication modal and clears inputs.
 * @returns {void}
 */
function openPasswordModal() {
    const modal = document.getElementById('password-modal');
    if (modal) {
        modal.classList.remove('hidden');
        document.getElementById('admin-password-input').value = '';
        document.getElementById('password-error-message').classList.add('hidden');
        document.getElementById('admin-password-input').focus();
    }
}

/**
 * Dismisses the custom password authentication modal.
 * @returns {void}
 */
function closePasswordModal() {
    const modal = document.getElementById('password-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

/**
 * Handles verifying the entered password against admin token '7777'.
 * On success, closes the modal and executes downloadSubmissionsCSV.
 * @returns {void}
 */
function handlePasswordVerification() {
    const passwordInput = document.getElementById('admin-password-input');
    const errMsg = document.getElementById('password-error-message');
    
    if (passwordInput.value.trim() === '7777') {
        closePasswordModal();
        downloadSubmissionsCSV();
    } else {
        if (errMsg) {
            errMsg.classList.remove('hidden');
        }
        passwordInput.focus();
        passwordInput.select();
    }
}

/**
 * Resets student project submission triggers.
 * @returns {void}
 */
function resetSubmissionUI() {
    closeSubmissionModal();
}

/**
 * Updates download button text and toggles visibility based on submissions count.
 * @returns {void}
 */
function updateDownloadButtonState() {
    const downloadBtn = document.getElementById('download-submissions-btn');
    if (!downloadBtn) return;

    if (state.submissions && state.submissions.length > 0) {
        downloadBtn.classList.remove('hidden');
        const template = translations[state.language || 'en'].downloadSubmissions;
        downloadBtn.innerHTML = `📥 ${template.replace('{count}', state.submissions.length)}`;
    } else {
        downloadBtn.classList.add('hidden');
    }
}

/**
 * Compiles stored submissions into CSV payloads and triggers browser save.
 * @returns {void}
 */
function downloadSubmissionsCSV() {
    const lang = state.language || 'en';
    if (!state.submissions || state.submissions.length === 0) {
        alert(lang === 'en' ? 'No submissions collected yet.' : 'अभी तक कोई सबमिशन एकत्र नहीं हुआ है।');
        return;
    }

    if (typeof Papa === 'undefined') {
        alert(lang === 'en' ? 'PapaParse failed to load. Please check your internet connection.' : 'PapaParse लोड होने में विफल रहा। कृपया अपना इंटरनेट कनेक्शन जांचें।');
        return;
    }

    const csvContent = Papa.unparse(state.submissions);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const blobUrl = URL.createObjectURL(blob);

    const downloadLink = document.createElement('a');
    downloadLink.setAttribute('href', blobUrl);
    downloadLink.setAttribute('download', `aim_student_submissions_${new Date().toISOString().split('T')[0]}.csv`);
    downloadLink.style.visibility = 'hidden';

    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

/**
 * Reveals loading skeleton loaders.
 * @returns {void}
 */
function showLoader() {
    document.getElementById('loader-view').classList.remove('hidden');
    document.getElementById('error-view').classList.add('hidden');
    document.getElementById('step-grade').classList.add('hidden');
    document.getElementById('step-activities').classList.add('hidden');
    document.getElementById('step-detail').classList.add('hidden');
}

/**
 * Dismisses loader skeleton loaders.
 * @returns {void}
 */
function hideLoader() {
    document.getElementById('loader-view').classList.add('hidden');
}

/**
 * Reveals error banner.
 * @param {string} msg - The text to display.
 * @returns {void}
 */
function showError(msg) {
    hideLoader();
    document.getElementById('step-grade').classList.add('hidden');
    document.getElementById('step-activities').classList.add('hidden');
    document.getElementById('step-detail').classList.add('hidden');

    const errorView = document.getElementById('error-view');
    errorView.classList.remove('hidden');

    const errorMsg = document.getElementById('error-message');
    if (errorMsg) {
        errorMsg.innerText = msg;
    }
}

/**
 * Dismisses error banner.
 * @returns {void}
 */
function hideError() {
    document.getElementById('error-view').classList.add('hidden');
}

/**
 * Translates static and metadata elements in the DOM based on active state.language.
 * @returns {void}
 */
function translateUIInternal() {
    const lang = state.language || 'en';
    const dict = translations[lang];

    // Translate all standard data-i18n elements
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (dict[key]) {
            el.innerText = dict[key];
        }
    });

    // Translate placeholder attributes
    document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
        const key = el.getAttribute('data-i18n-placeholder');
        if (dict[key]) {
            el.setAttribute('placeholder', dict[key]);
        }
    });

    // Translate aria-label attributes
    document.querySelectorAll('[data-i18n-aria]').forEach(el => {
        const key = el.getAttribute('data-i18n-aria');
        if (dict[key]) {
            el.setAttribute('aria-label', dict[key]);
        }
    });



    // Refresh dynamic download button text count
    updateDownloadButtonState();
}

// Kickoff
document.addEventListener('DOMContentLoaded', init);
