/**
 * @fileoverview Main application logic for AIM - Adobe Implementation Plan Student Tutorial Platform.
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

// Google Sheets Integration Web App URL (Paste your Google Script webapp executable URL here)
// Example: "https://script.google.com/macros/s/AKfycbw.../exec"
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyOETqGyMYJ7phV0LN2mD9vQykJqYWfcsReF1hdmKFHRA_HlM74FaGsVMLgNhbVhdi3cQ/exec";

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
    kaushalBodhActivities: [],
    selectedGrade: null,
    selectedActivity: null,
    currentStep: 'grade',
    filterSubject: 'All', // Mapped to Month or Form of Work filtering
    submissions: [],
    language: 'en', // 'en' or 'hi'
    userExplicitLanguage: 'en', // Manually set via language toggle
    tutorialMode: 'choose' // 'choose' | 'hindi' | 'english' | 'kaushal-bodh'
};

const translations = {
    en: {
        welcomeTitle: "Welcome to AIM",
        welcomeDesc: "AIM is a creative learning space for students in India. Every week you get fresh activities built on Adobe Express. Design posters, generate images, create videos, and build things that are yours. Start simple. Level up every week.",
        activityOfMonthTag: "Activity of the Month",
        independenceTitle: "Independence Day Special",
        independenceDesc: "Celebrate India's Independence Day by creating patriotic posters and videos using Adobe Express for Education!",
        exploreActivityBtn: "Activity Template",
        submitMonthlyBtn: "Submit Activity Link",
        monthlyTutorialsTitle: "Monthly Activity Tutorials",
        englishTutorial: "English Tutorial",
        hindiTutorial: "Hindi Tutorial",
        englishTutorials: "DCAIS Activities / English",
        hindiTutorials: "DCAIS Activities / Hindi",
        kaushalBodhTab: "Kaushal Bodh",
        formOfWorkLabel: "Form of Work",
        allFormsOfWorkOption: "All Forms of Work",
        workWithLifeForms: "Work with Life Forms",
        workWithMachinesMaterials: "Work with Machines & Materials",
        workInHumanServices: "Work in Human Services",
        keyActivitiesLabel: "Key Activities",
        keyLearningOutcomesLabel: "Key Learning Outcomes",
        adobeExpressIntegrationLabel: "Adobe Express Integration for Activity Books & Portfolios",
        adobeExpressActivityLabel: "Adobe Express Activity",
        finalStudentDeliverableLabel: "Final Student Deliverable",
        openLearningJournal: "Open Learning Journal",
        openTemplateLink: "Open Activity Template",
        learningJournalBtn: "📘 Open Learning Journal Link",
        kaushalHeroSubtitle: "Choose your grade level from Kaushal Bodh curriculum",
        kaushalActivitiesSubtitle: "Showing Kaushal Bodh projects for {grade}",
        importantNoteHeading: "Important Note",
        importantNoteText: "Earn accredited DCAIS certification upon completion of minimum one activity per month and submit the creative assignments.",
        certNotice: "To earn your certificate, please complete at least one activity per month",
        quickSearchTitle: "Quick Activity Search",
        quickSearchDesc: "Search activities across all grades directly by name",
        searchActivityByName: "Search activity by name",
        startBtn: "Start",
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
        gradeLabel: "Grade",
        activityNameLabel: "Activity Name",
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
        gradePlaceholder: "e.g. Grade 3",
        activityNamePlaceholder: "e.g. Summer Vacation Memories",
        projectLinkPlaceholder: "e.g. https://new.express.adobe.com/...",
        passwordInputPlaceholder: "Password",
        footerText: "© 2026 AIM - Activity Implementation Plan. All rights reserved.",
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
        videoUnavailable: "🔒 Tutorial Link Coming Soon",
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
        welcomeTitle: "AIM में आपका स्वागत है",
        welcomeDesc: "AIM भारत में छात्रों के लिए एक रचनात्मक शिक्षण स्थल है। हर हफ्ते आपको Adobe Express पर बनी नई गतिविधियाँ मिलती हैं। पोस्टर डिज़ाइन करें, इमेजेस जनरेट करें, वीडियो बनाएँ, और अपनी चीज़ें बनाएँ। सरल से शुरुआत करें और हर हफ्ते आगे बढ़ें।",
        activityOfMonthTag: "महीने की गतिविधि",
        independenceTitle: "स्वतंत्रता दिवस विशेष",
        independenceDesc: "Adobe Express for Education का उपयोग करके देशभक्ति के पोस्टर और वीडियो बनाकर भारत का स्वतंत्रता दिवस मनाएँ!",
        exploreActivityBtn: "गतिविधि टेम्पलेट",
        submitMonthlyBtn: "गतिविधि लिंक जमा करें",
        monthlyTutorialsTitle: "मासिक गतिविधि ट्यूटोरियल",
        englishTutorial: "अंग्रेज़ी ट्यूटोरियल",
        hindiTutorial: "हिंदी ट्यूटोरियल",
        englishTutorials: "DCAIS एक्टिविटीज़ / अंग्रेज़ी",
        hindiTutorials: "DCAIS एक्टिविटीज़ / हिंदी",
        kaushalBodhTab: "कौशल बोध",
        formOfWorkLabel: "कार्य का प्रकार (Form of Work)",
        allFormsOfWorkOption: "सभी कार्य प्रकार",
        workWithLifeForms: "सजीव रूपों के साथ कार्य (Work with Life Forms)",
        workWithMachinesMaterials: "मशीनों और सामग्रियों के साथ कार्य (Work with Machines & Materials)",
        workInHumanServices: "मानव सेवाओं में कार्य (Work in Human Services)",
        keyActivitiesLabel: "मुख्य गतिविधियाँ (Key Activities)",
        keyLearningOutcomesLabel: "मुख्य सीखने के परिणाम (Key Learning Outcomes)",
        adobeExpressIntegrationLabel: "एक्टिविटी बुक्स और पोर्टफोलियो के लिए एडोब एक्सप्रेस एकीकरण",
        adobeExpressActivityLabel: "एडोब एक्सप्रेस गतिविधि (Adobe Express Activity)",
        finalStudentDeliverableLabel: "अंतिम छात्र आउटपुट (Final Student Deliverable)",
        openLearningJournal: "लर्निंग जर्नल खोलें",
        openTemplateLink: "एक्टिविटी टेम्पलेट खोलें",
        learningJournalBtn: "📘 लर्निंग जर्नल लिंक खोलें",
        kaushalHeroSubtitle: "कौशल बोध प्रोजेक्ट्स और लर्निंग जर्नल देखने के लिए अपनी कक्षा चुनें",
        kaushalActivitiesSubtitle: "{grade} के कौशल बोध प्रोजेक्ट्स",
        importantNoteHeading: "महत्वपूर्ण नोट",
        importantNoteText: "मासिक न्यूनतम एक गतिविधि पूरी करने और रचनात्मक कार्य जमा करने पर मान्यता प्राप्त DCAIS प्रमाण पत्र प्राप्त करें।",
        certNotice: "प्रमाण पत्र प्राप्त करने के लिए, कृपया सुनिश्चित करें कि आप प्रति माह एक गतिविधि अवश्य पूरी करें",
        quickSearchTitle: "त्वरित गतिविधि खोज",
        quickSearchDesc: "नाम से सभी कक्षाओं की गतिविधियों को सीधे खोजें",
        searchActivityByName: "गतिविधि के नाम से खोजें...",
        startBtn: "प्रारंभ करें",
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
        gradeLabel: "कक्षा",
        activityNameLabel: "गतिविधि का नाम",
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
        gradePlaceholder: "जैसे: कक्षा 3",
        activityNamePlaceholder: "जैसे: गर्मी की छुट्टियों की यादें",
        projectLinkPlaceholder: "जैसे: https://new.express.adobe.com/...",
        passwordInputPlaceholder: "पासवर्ड",
        footerText: "© 2026 AIM - Activity Implementation Plan. सर्वाधिकार सुरक्षित (All rights reserved).",
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
        videoUnavailable: "🔒 ट्यूटोरियल लिंक जल्द आ रहा है",
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

    // Dedicated Landing Page Language Toggle Buttons
    const langToggleEn = document.getElementById('lang-toggle-en');
    const langToggleHi = document.getElementById('lang-toggle-hi');

    if (langToggleEn) {
        langToggleEn.addEventListener('click', () => {
            state.userExplicitLanguage = 'en';
            state.language = 'en';
            render();
        });
    }

    if (langToggleHi) {
        langToggleHi.addEventListener('click', () => {
            state.userExplicitLanguage = 'hi';
            state.language = 'hi';
            render();
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

    const choiceKaushalBodh = document.getElementById('choice-kaushal-bodh');
    if (choiceKaushalBodh) {
        choiceKaushalBodh.addEventListener('click', (e) => {
            e.preventDefault();
            state.tutorialMode = 'kaushal-bodh';
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
            if (state.tutorialMode === 'kaushal-bodh') {
                state.currentStep = 'form-of-work';
            } else {
                state.currentStep = 'activities';
            }
            render();
        });
    }

    // Back to Home (Grades selection) link
    const backToGradesLink = document.getElementById('back-to-grades-link');
    if (backToGradesLink) {
        backToGradesLink.addEventListener('click', (e) => {
            e.preventDefault();
            // In kaushal-bodh, go back to Form of Work step
            if (state.tutorialMode === 'kaushal-bodh') {
                state.filterSubject = '';
                state.currentStep = 'form-of-work';
                render();
            } else {
                resetState();
            }
        });
    }

    // Back to Grade Selection from Form of Work step
    const backToGradesFromFow = document.getElementById('back-to-grades-from-fow');
    if (backToGradesFromFow) {
        backToGradesFromFow.addEventListener('click', (e) => {
            e.preventDefault();
            state.filterSubject = '';
            state.currentStep = 'grade';
            render();
        });
    }

    // Submissions Google Form redirect for activities
    const showSubmissionFormBtn = document.getElementById('show-submission-form-btn');
    if (showSubmissionFormBtn) {
        showSubmissionFormBtn.addEventListener('click', (e) => {
            e.preventDefault();
            window.open('https://forms.gle/ZkzcFSFpQQUyxeWC8', '_blank', 'noopener,noreferrer');
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

    // Clear validation error when student edits the work link
    const submitWorkLinkInput = document.getElementById('submit-work-link');
    if (submitWorkLinkInput) {
        submitWorkLinkInput.addEventListener('input', () => {
            submitWorkLinkInput.setCustomValidity('');
            submitWorkLinkInput.classList.remove('input-error');
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

    // Monthly Activity submit button on Welcome Screen (Independence Day Special)
    const submitMonthlyActivityBtn = document.getElementById('submit-monthly-activity-btn');
    if (submitMonthlyActivityBtn) {
        submitMonthlyActivityBtn.addEventListener('click', (e) => {
            e.preventDefault();
            window.open('https://forms.gle/FZjYhFcQ3S8SuLKZA', '_blank', 'noopener,noreferrer');
        });
    }

    const exploreActivityBtn = document.getElementById('explore-activity-btn');
    if (exploreActivityBtn) {
        exploreActivityBtn.addEventListener('click', (e) => {
            const templateUri = "https://new.express.adobe.com/design/template/urn:aaid:sc:VA6C2:2313db0a-a09f-5f71-8b5f-d07c4cc6df42?category=text&taskID=digital-activity";
            window.open(templateUri, "_blank", "noopener,noreferrer");
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



    // Quick Activity Search Input listeners
    const quickSearchInput = document.getElementById('quick-activity-search');
    const clearQuickSearchBtn = document.getElementById('clear-quick-search-btn');

    if (quickSearchInput) {
        quickSearchInput.addEventListener('input', (e) => {
            const val = e.target.value;
            if (clearQuickSearchBtn) {
                if (val.trim().length > 0) {
                    clearQuickSearchBtn.classList.remove('hidden');
                } else {
                    clearQuickSearchBtn.classList.add('hidden');
                }
            }
            renderQuickSearchResults(val);
        });
    }

    if (clearQuickSearchBtn && quickSearchInput) {
        clearQuickSearchBtn.addEventListener('click', () => {
            quickSearchInput.value = '';
            clearQuickSearchBtn.classList.add('hidden');
            renderQuickSearchResults('');
            quickSearchInput.focus();
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
 * Parses project title and duration string from Kaushal Bodh excel.
 */
function parseProjectNameAndDuration(str) {
    if (!str) return { name: "Project", duration: "", full: "Project" };
    
    let clean = str.replace(/\r\n/g, "\n").trim();
    if (clean.includes("Duration:")) {
        const parts = clean.split(/Duration:/i);
        const name = parts[0].trim().replace(/\n+/g, " ");
        const duration = parts[1] ? parts[1].trim() : "";
        return { name, duration, full: `${name} (${duration})` };
    }
    
    const match = clean.match(/^(.*?)\s*\((.*?Hours.*?)\)$/i);
    if (match) {
        return { name: match[1].trim(), duration: match[2].trim(), full: clean };
    }
    
    return { name: clean.split('\n')[0].trim(), duration: "", full: clean };
}

/**
 * Formats multi-line text with bullets or paragraphs into structured HTML.
 */
function formatFormattedText(text) {
    if (!text) return '';
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    let html = '';
    let inList = false;

    lines.forEach(line => {
        if (line.startsWith('•') || line.startsWith('-')) {
            if (!inList) {
                html += '<ul class="custom-bullet-list" style="margin: 0.5rem 0; padding-left: 1.25rem; display: flex; flex-direction: column; gap: 0.4rem;">';
                inList = true;
            }
            const cleanLine = line.replace(/^[•\-]\s*/, '');
            html += `<li style="line-height: 1.5; color: var(--color-text-primary); font-size: 0.925rem;">${escapeHTML(cleanLine)}</li>`;
        } else {
            if (inList) {
                html += '</ul>';
                inList = false;
            }
            html += `<p style="margin: 0.4rem 0; line-height: 1.5; color: var(--color-text-primary); font-size: 0.925rem;">${escapeHTML(line)}</p>`;
        }
    });

    if (inList) {
        html += '</ul>';
    }

    return html;
}

/**
 * Returns metadata and CSS theme styling for Kaushal Bodh Form of Work.
 */
function getFormOfWorkInfo(formOfWork) {
    const f = (formOfWork || "").toLowerCase();
    if (f.includes('life')) {
        return {
            name: "Work with Life Forms",
            icon: "🌿",
            colorClass: "form-life-forms",
            color: "#10B981"
        };
    } else if (f.includes('machine') || f.includes('material')) {
        return {
            name: "Work with Machines & Materials",
            icon: "⚙️",
            colorClass: "form-machines-materials",
            color: "#2563eb"
        };
    } else if (f.includes('human') || f.includes('service')) {
        return {
            name: "Work in Human Services",
            icon: "🤝",
            colorClass: "form-human-services",
            color: "#8b5cf6"
        };
    }
    return {
        name: formOfWork || "General",
        icon: "📋",
        colorClass: "form-general",
        color: "#6b7280"
    };
}

/**
 * Renders multiple Template buttons if cell contains multiple template links.
 */
function renderTemplateButtons(templateLinkStr) {
    if (!templateLinkStr) return '';
    const urls = templateLinkStr.split(/\s+/).map(u => u.trim()).filter(u => u.startsWith('http://') || u.startsWith('https://'));
    if (urls.length === 0) return '';

    const dict = translations[state.language || 'en'];
    let html = '';

    urls.forEach((url, i) => {
        const label = urls.length > 1 ? `${dict.openTemplateLink} (${i + 1})` : dict.openTemplateLink;
        html += `
            <a href="${url}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary kb-template-btn" style="background: #ffffff; border: 2px solid #eb1000; color: #eb1000; padding: 0.85rem 1.5rem; font-size: 0.95rem; font-weight: 700; border-radius: var(--radius-full); text-decoration: none; display: inline-flex; align-items: center; gap: 0.5rem; box-shadow: 0 2px 8px rgba(235, 16, 0, 0.1); transition: all 0.2s ease;">
                <span>📝</span> ${label} <span style="font-size: 1rem;">↗</span>
            </a>
        `;
    });

    return html;
}

/**
 * Parses Kaushal Bodh implementation excel file.
 */
function parseKaushalBodhExcel(arrayBuffer) {
    if (typeof XLSX === 'undefined') return;

    try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const kbActivities = [];

        for (let s_idx = 0; s_idx < workbook.SheetNames.length; s_idx++) {
            const rawSheetName = workbook.SheetNames[s_idx];
            if (!rawSheetName || !rawSheetName.trim()) continue;

            const sheetName = rawSheetName.trim();
            let gradeName = sheetName;
            const matchGrade = sheetName.match(/grade\s*(\d+)/i);
            if (matchGrade) {
                gradeName = `Grade ${matchGrade[1]}`;
            }

            const sheet = workbook.Sheets[rawSheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            rows.forEach((row, rIdx) => {
                const formOfWork = (row['Form of Work'] || row['Form Of Work'] || "").toString().trim();
                const projectNameDuration = (row['Project Name & Duration'] || row['Project Name and Duration'] || "").toString().trim();
                const keyActivities = (row['Key Activities'] || "").toString().trim();
                const keyLearningOutcomes = (row['Key Learning Outcomes'] || "").toString().trim();
                const adobeExpressIntegration = (row['Adobe Express Integration for Activity Books & Portfolios'] || row['Adobe Express Integration'] || "").toString().trim();
                const adobeExpressActivity = (row['Adobe Express Activity'] || "").toString().trim();
                const finalStudentDeliverable = (row['Final Student Deliverable'] || "").toString().trim();
                const templateLink = (row['Template Link'] || "").toString().trim();
                const learningJournalLink = (row['Learning Journal Link'] || "").toString().trim();

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
                    month: formOfWork || "Kaushal Bodh",
                    week: ""
                });
            });
        }

        state.kaushalBodhActivities = kbActivities;
    } catch (err) {
        console.error("Error parsing Kaushal Bodh Excel:", err);
    }
}

/**
 * Fetches the default Excel curriculum sheets from the server.
 * @returns {Promise<void>}
 */
async function loadDefaultExcel() {
    showLoader();
    try {
        const fetchUrl1 = 'Implementation plan.xlsx?v=' + Date.now();
        const fetchUrl2 = 'final of Kaushal Bodh Implementation Table.xlsx?v=' + Date.now();

        const [res1, res2] = await Promise.allSettled([
            fetch(fetchUrl1),
            fetch(fetchUrl2)
        ]);

        if (res1.status === 'fulfilled' && res1.value.ok) {
            const buf1 = await res1.value.arrayBuffer();
            parseExcel(buf1);
        } else {
            console.error('Could not load Implementation plan.xlsx');
        }

        if (res2.status === 'fulfilled' && res2.value.ok) {
            const buf2 = await res2.value.arrayBuffer();
            parseKaushalBodhExcel(buf2);
        } else {
            console.warn('Could not load final of Kaushal Bodh Implementation Table.xlsx');
        }

        hideError();
        hideLoader();
    } catch (error) {
        showError('Could not auto-load the curriculum excel files. Please make sure Implementation plan.xlsx is in the root folder.');
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
            let bookScreenshotRowIdx = -1;
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
                } else if (label.includes("screenshot") || label.includes("book")) {
                    bookScreenshotRowIdx = r;
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
            if (bookScreenshotRowIdx === -1) bookScreenshotRowIdx = 8;
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

                // Month header (Search rows 3, 4, 5 / index 2, 3, 4) - Fill down merged cell value
                let monthVal = null;
                for (let rIdx of [2, 3, 4]) {
                    const val = (rows[rIdx] && rows[rIdx][c]) ? rows[rIdx][c].toString().trim() : null;
                    if (val && val !== "" && 
                        !val.toUpperCase().startsWith("GRADE") && 
                        !val.toUpperCase().startsWith("कक्षा") && 
                        !val.toUpperCase().startsWith("WEEK") && 
                        !val.toUpperCase().startsWith("सप्ताह")) {
                        monthVal = val;
                        break;
                    }
                }
                if (monthVal) {
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
                const bookScreenshot = (bookScreenshotRowIdx !== -1 && rows[bookScreenshotRowIdx] && rows[bookScreenshotRowIdx][c]) ? rows[bookScreenshotRowIdx][c].toString().trim() : "";
                const templateLink = (rows[templateLinkRowIdx] && rows[templateLinkRowIdx][c]) ? rows[templateLinkRowIdx][c].toString().trim() : "";
                const videoLink = (rows[videoLinkRowIdx] && rows[videoLinkRowIdx][c]) ? rows[videoLinkRowIdx][c].toString().trim() : "";
                const description = (rows[descriptionRowIdx] && rows[descriptionRowIdx][c]) ? rows[descriptionRowIdx][c].toString().trim() : "";

                parsedActivities.push({
                    grade: cleanGradeName, // Grade 5, Grade 7, Grade 8
                    activity_name: activityName.trim(),
                    skills: skills,
                    instructions: instructions,
                    book_screenshot: bookScreenshot,
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
    if (state.tutorialMode === 'kaushal-bodh') {
        const kbSet = new Set((state.kaushalBodhActivities || []).map(a => a.grade).filter(Boolean));
        let kbGrades = Array.from(kbSet);
        return kbGrades.sort((a, b) => {
            const numA = parseInt(a.replace(/\D/g, ''), 10);
            const numB = parseInt(b.replace(/\D/g, ''), 10);
            return numA - numB;
        });
    }

    const gradesSet = new Set(activities.map(a => a.grade).filter(Boolean));
    let gradesArray = Array.from(gradesSet);

    // Filter grades based on the active tutorialMode
    if (state.tutorialMode === 'hindi') {
        gradesArray = gradesArray.filter(g => g.toLowerCase().includes('(hindi)'));
    } else if (state.tutorialMode === 'english') {
        gradesArray = gradesArray.filter(g => !g.toLowerCase().includes('(hindi)'));
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
    document.getElementById('step-form-of-work').classList.add('hidden');
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

    const langToggleContainer = document.getElementById('global-lang-toggle');
    if (langToggleContainer) {
        if (state.currentStep === 'welcome') {
            langToggleContainer.classList.remove('hidden');
        } else {
            langToggleContainer.classList.add('hidden');
        }
    }

    if (state.currentStep === 'welcome') {
        state.language = state.userExplicitLanguage || 'en';
        document.getElementById('step-welcome').classList.remove('hidden');
        focusId = '';
    } else if (state.currentStep === 'grade') {
        document.getElementById('step-grade').classList.remove('hidden');
        
        // Hide Quick Activity Search in Kaushal Bodh mode
        const quickSearchSec = document.getElementById('quick-search-section');
        if (quickSearchSec) {
            if (state.tutorialMode === 'kaushal-bodh') {
                quickSearchSec.classList.add('hidden');
            } else {
                quickSearchSec.classList.remove('hidden');
            }
        }

        renderGradeSelection();
        focusId = 'grade-title';
    } else if (state.currentStep === 'form-of-work') {
        document.getElementById('step-form-of-work').classList.remove('hidden');
        renderFormOfWorkCards();

        const lang = state.language || 'en';
        document.getElementById('fow-title').innerText = lang === 'hi' ? 'कार्य का प्रकार चुनें' : 'Select Form of Work';
        document.getElementById('fow-subtitle').innerText = lang === 'hi'
            ? getDisplayGradeName(state.selectedGrade) + ' के लिए कार्य का प्रकार चुनें'
            : 'Choose a form of work for ' + getDisplayGradeName(state.selectedGrade);
        document.getElementById('fow-section-title').innerText = lang === 'hi' ? 'कार्य का प्रकार' : 'Form of Work';
        focusId = 'fow-title';
    } else if (state.currentStep === 'activities') {
        document.getElementById('step-activities').classList.remove('hidden');

        const stdFilterBar = document.getElementById('standard-filter-bar');
        const kbTabsContainer = document.getElementById('kaushal-tabs-container');

        if (state.tutorialMode === 'kaushal-bodh') {
            if (stdFilterBar) stdFilterBar.classList.add('hidden');
            if (kbTabsContainer) kbTabsContainer.classList.add('hidden');
        } else {
            if (stdFilterBar) stdFilterBar.classList.remove('hidden');
            if (kbTabsContainer) kbTabsContainer.classList.add('hidden');
            
            const gradeActivities = state.allActivities.filter(a => a.grade === state.selectedGrade);
            populateMonthDropdown(gradeActivities);

            const subjectFilter = document.getElementById('subject-filter');
            if (subjectFilter) {
                subjectFilter.value = state.filterSubject;
            }
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
        if (state.currentStep === 'form-of-work') {
            html += `<li><span class="current">${displayGrade}</span></li>`;
        } else if (state.currentStep === 'activities' || state.currentStep === 'detail') {
            html += `<li><a href="#" id="breadcrumb-grade">${displayGrade}</a></li>`;
        }
    }

    // Level 3: Form of Work / Activities
    if (state.tutorialMode === 'kaushal-bodh' && state.filterSubject && state.filterSubject !== 'All') {
        const fowName = state.filterSubject;
        if (state.currentStep === 'activities') {
            html += `<li><span class="current">${fowName}</span></li>`;
        } else if (state.currentStep === 'detail') {
            html += `<li><a href="#" id="breadcrumb-fow">${fowName}</a></li>`;
        }
    }

    // Level 4: Activity Detail
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
            if (state.tutorialMode === 'kaushal-bodh') {
                state.currentStep = 'form-of-work';
            } else {
                state.currentStep = 'activities';
            }
            render();
        });
    }

    const fowBtn = document.getElementById('breadcrumb-fow');
    if (fowBtn) {
        fowBtn.addEventListener('click', (e) => {
            e.preventDefault();
            state.selectedActivity = null;
            if (state.tutorialMode === 'kaushal-bodh') {
                state.currentStep = 'form-of-work';
            } else {
                state.currentStep = 'activities';
            }
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

    const isKB = state.tutorialMode === 'kaushal-bodh';
    const activitiesList = isKB ? state.kaushalBodhActivities : state.allActivities;
    let grades = extractGrades(activitiesList);
    
    // Update subtitle for Kaushal Bodh mode if applicable
    const heroSubtitleEl = document.getElementById('hero-subtitle');
    if (heroSubtitleEl) {
        const dict = translations[state.language || 'en'];
        heroSubtitleEl.innerText = isKB ? (dict.kaushalHeroSubtitle || dict.heroSubtitle) : dict.heroSubtitle;
    }

    // Render Quick Activity Search results
    const quickSearchInput = document.getElementById('quick-activity-search');
    renderQuickSearchResults(quickSearchInput ? quickSearchInput.value : '');

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
        const count = activitiesList.filter(a => a.grade === grade).length;
        const template = count === 1 ? translations[state.language || 'en'].activitySingle : translations[state.language || 'en'].activitiesPlural;
        const labelText = template.replace('{count}', count);

        const numStr = grade.replace(/\D/g, '');
        let icon = '🎓';
        let accentClass = 'indigo';
        if (numStr === '6' || numStr === '5') {
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
            if (state.tutorialMode === 'kaushal-bodh') {
                state.filterSubject = '';
                state.currentStep = 'form-of-work';
            } else {
                state.filterSubject = 'All';
                state.currentStep = 'activities';
            }
            render();
        });
    });
}

/**
 * Renders real-time activity search results inside Quick Activity Search on Step 1.
 * Matches both DCAIS activities and Kaushal Bodh projects.
 * @param {string} query
 */
function renderQuickSearchResults(query = '') {
    const container = document.getElementById('quick-activity-results');
    if (!container) return;

    const cleanQuery = (query || '').trim().toLowerCase();
    const lang = state.language || 'en';

    if (!cleanQuery) {
        container.innerHTML = `
            <div style="text-align: center; color: var(--color-text-secondary); padding: 1.5rem 1rem; font-size: 0.875rem;">
                <span style="font-size: 1.75rem; display: block; margin-bottom: 0.35rem;">🔎</span>
                <span>${lang === 'hi' ? 'गतिविधि खोजने के लिए नाम दर्ज करें' : 'Type an activity name on the left for instant results!'}</span>
            </div>
        `;
        return;
    }

    const allCombined = [...state.allActivities, ...state.kaushalBodhActivities];
    const matches = allCombined.filter(a =>
        a.activity_name && a.activity_name.toLowerCase().includes(cleanQuery)
    );

    if (matches.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; color: var(--color-text-secondary); padding: 1.5rem 1rem; font-size: 0.875rem;">
                <span style="font-size: 1.75rem; display: block; margin-bottom: 0.35rem;">🤔</span>
                <span>${lang === 'hi' ? 'कोई मेल खाती गतिविधि नहीं मिली' : 'No activities matching "' + escapeHTML(cleanQuery) + '"'}</span>
            </div>
        `;
        return;
    }

    let html = '';
    matches.slice(0, 15).forEach(act => {
        const isKB = act.isKaushalBodh;
        const tagText = isKB ? (act.form_of_work || 'Kaushal Bodh') : (act.month || '');

        html += `
            <div class="quick-result-item" data-activity-name="${encodeURIComponent(act.activity_name)}" data-activity-grade="${encodeURIComponent(act.grade)}" data-is-kb="${isKB ? 'true' : 'false'}" style="display: flex; align-items: center; justify-content: space-between; padding: 0.65rem 0.85rem; background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius-md); cursor: pointer; transition: all 0.2s ease; box-shadow: var(--shadow-sm);">
                <div style="display: flex; flex-direction: column; gap: 0.15rem; flex: 1; padding-right: 0.5rem; text-align: left;">
                    <span style="font-weight: 700; color: var(--color-text-primary); font-size: 0.9rem;">${escapeHTML(act.activity_name)}</span>
                    <span style="font-size: 0.775rem; color: var(--color-text-secondary);">${escapeHTML(tagText)} ${act.skills ? '• ' + escapeHTML(act.skills.substring(0, 60)) + '...' : ''}</span>
                </div>
                <span class="grade-badge" style="font-size: 0.75rem; padding: 0.2rem 0.55rem; border-radius: var(--radius-full); background: ${isKB ? 'rgba(37, 99, 235, 0.1)' : 'var(--color-subject-1-bg)'}; color: ${isKB ? '#2563eb' : 'var(--color-subject-1)'}; font-weight: 700; white-space: nowrap;">${getDisplayGradeName(act.grade)}</span>
            </div>
        `;
    });

    container.innerHTML = html;

    container.querySelectorAll('.quick-result-item').forEach(item => {
        item.addEventListener('click', () => {
            const actName = decodeURIComponent(item.getAttribute('data-activity-name'));
            const actGrade = decodeURIComponent(item.getAttribute('data-activity-grade'));
            const isKB = item.getAttribute('data-is-kb') === 'true';

            const pool = isKB ? state.kaushalBodhActivities : state.allActivities;
            const selected = pool.find(a => a.grade === actGrade && a.activity_name === actName);

            if (selected) {
                if (isKB) {
                    state.tutorialMode = 'kaushal-bodh';
                }
                state.selectedGrade = actGrade;
                state.selectedActivity = selected;
                state.currentStep = 'detail';
                render();
            }
        });
    });
}

function escapeHTML(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

/**
 * Populates Month dropdown filter dynamically based on Month values inside current grade.
 * @param {Activity[]} activities - Filtered curriculum list for active grade.
 * @returns {void}
 */
function populateMonthDropdown(activities) {
    const subjectFilter = document.getElementById('subject-filter');
    const filterLabelText = document.getElementById('filter-label-text');
    const wrapper = document.getElementById('subject-filter-wrapper');
    if (!subjectFilter) return;

    const isKB = state.tutorialMode === 'kaushal-bodh';
    const lang = state.language || 'en';
    const dict = translations[lang];

    if (isKB) {
        if (filterLabelText) filterLabelText.innerText = dict.formOfWorkLabel || "Form of Work";
        if (wrapper) wrapper.style.width = '240px';

        const formsSet = new Set(activities.map(a => a.form_of_work).filter(Boolean));
        const forms = Array.from(formsSet);

        let html = `<option value="All">${dict.allFormsOfWorkOption || "All Forms of Work"}</option>`;
        forms.forEach(f => {
            let label = f;
            if (lang === 'hi') {
                if (f.toLowerCase().includes('life')) label = dict.workWithLifeForms || f;
                else if (f.toLowerCase().includes('machine')) label = dict.workWithMachinesMaterials || f;
                else if (f.toLowerCase().includes('human')) label = dict.workInHumanServices || f;
            }
            html += `<option value="${escapeHTML(f)}">${escapeHTML(label)}</option>`;
        });

        subjectFilter.innerHTML = html;
        subjectFilter.value = forms.includes(state.filterSubject) ? state.filterSubject : 'All';
        state.filterSubject = subjectFilter.value;
        return;
    }

    if (filterLabelText) filterLabelText.innerText = dict.monthLabel || "Month";
    if (wrapper) wrapper.style.width = '180px';

    const currentSel = state.filterSubject;
    const months = Array.from(new Set(activities.map(a => a.month).filter(Boolean)));
    
    // Sort months chronologically starting from July (academic session)
    const monthMap = {
        "जुलाई": "July", "अगस्त": "August", "सितम्बर": "September", "अक्टूबर": "October", 
        "नवंबर": "November", "दिसंबर": "December", "जनवरी": "January", "फरवरी": "February", 
        "मार्च": "March", "अप्रैल": "April", "मई": "May", "जून": "June",
        "July": "July", "August": "August", "September": "September", "October": "October",
        "November": "November", "December": "December", "January": "January", "February": "February",
        "March": "March", "April": "April", "May": "May", "June": "June"
    };
    const monthOrder = ["July", "August", "September", "October", "November", "December", "January", "February", "March", "April", "May", "June"];
    months.sort((a, b) => {
        const engA = monthMap[a] || a;
        const engB = monthMap[b] || b;
        const idxA = monthOrder.indexOf(engA);
        const idxB = monthOrder.indexOf(engB);
        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
        return a.localeCompare(b);
    });

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
 * Checks if a form_of_work string matches a given category key.
 * @param {string} formValue - The form_of_work from the activity data
 * @param {string} categoryKey - One of 'life', 'machines', 'human'
 * @returns {boolean}
 */
function matchesFormCategory(formValue, categoryKey) {
    const f = (formValue || '').toLowerCase();
    if (categoryKey === 'life') return f.includes('life');
    if (categoryKey === 'machines') return f.includes('machine') || f.includes('material');
    if (categoryKey === 'human') return f.includes('human') || f.includes('service');
    return false;
}

/**
 * Returns the internal category key ('life', 'machines', 'human') for a filterSubject string.
 */
function getFormCategory(filterSubject) {
    const s = (filterSubject || '').toLowerCase();
    if (s.includes('life')) return 'life';
    if (s.includes('machine') || s.includes('material')) return 'machines';
    if (s.includes('human') || s.includes('service')) return 'human';
    return 'life'; // default
}

/**
 * Renders the 3 Form of Work selection cards for Kaushal Bodh mode.
 * Matches the grade-card visual pattern from Step 1.
 */
function renderKaushalTabs() {
    const container = document.getElementById('kaushal-tabs-container');
    if (!container) return;

    const lang = state.language || 'en';

    if (!state.filterSubject || state.filterSubject === 'All') {
        state.filterSubject = 'Work with Life Forms';
    }

    const activeCategory = getFormCategory(state.filterSubject);
    const gradeList = state.kaushalBodhActivities.filter(a => a.grade === state.selectedGrade);

    const formsData = [
        {
            key: 'Work with Life Forms',
            category: 'life',
            title: lang === 'hi' ? 'सजीव रूपों के साथ कार्य' : 'Work with Life Forms',
            subtitle: lang === 'hi' ? 'पौधे, जीव विविधता और पर्यावरण' : 'Plants, Animals & Ecology',
            icon: '🌿',
            accentClass: 'form-card-life'
        },
        {
            key: 'Work with Machines & Materials',
            category: 'machines',
            title: lang === 'hi' ? 'मशीनों और सामग्रियों के साथ कार्य' : 'Work with Machines & Materials',
            subtitle: lang === 'hi' ? 'उपकरण, निर्माण और तकनीक' : 'Tools, Fabrication & Tech',
            icon: '⚙️',
            accentClass: 'form-card-machines'
        },
        {
            key: 'Work in Human Services',
            category: 'human',
            title: lang === 'hi' ? 'मानव सेवाओं में कार्य' : 'Work in Human Services',
            subtitle: lang === 'hi' ? 'स्वास्थ्य, समाज और संचार' : 'Health, Society & Service',
            icon: '🤝',
            accentClass: 'form-card-human'
        }
    ];

    let html = `
        <div class="kb-section-box">
            <div class="kb-section-header">
                <span class="kb-section-label">${lang === 'hi' ? 'खंड 1' : 'SECTION 1'}</span>
                <h2 class="kb-section-heading">${lang === 'hi' ? 'कार्य का प्रकार चुनें' : 'Select Form of Work'}</h2>
                <p class="kb-section-subtext">${lang === 'hi' ? 'नीचे दी गई किसी भी श्रेणी पर क्लिक करके उसके प्रोजेक्ट्स देखें' : 'Click on a category below to view its projects'}</p>
            </div>
            <div class="kb-form-row">
    `;

    formsData.forEach(form => {
        const isActive = form.category === activeCategory;
        const count = gradeList.filter(a => matchesFormCategory(a.form_of_work, form.category)).length;
        const countLabel = count === 1 ? (lang === 'hi' ? '1 प्रोजेक्ट' : '1 project') : (lang === 'hi' ? `${count} प्रोजेक्ट्स` : `${count} projects`);

        let themeColor = '#10b981';
        let themeBg = 'rgba(16, 185, 129, 0.08)';
        let themeBorder = 'rgba(16, 185, 129, 0.35)';
        let themeIconBg = '#ecfdf5';
        if (form.category === 'machines') {
            themeColor = '#2563eb';
            themeBg = 'rgba(37, 99, 235, 0.08)';
            themeBorder = 'rgba(37, 99, 235, 0.35)';
            themeIconBg = '#eff6ff';
        }
        if (form.category === 'human') {
            themeColor = '#8b5cf6';
            themeBg = 'rgba(139, 92, 246, 0.08)';
            themeBorder = 'rgba(139, 92, 246, 0.35)';
            themeIconBg = '#f5f3ff';
        }

        html += `
            <button type="button" class="kb-form-rect-card ${isActive ? 'kb-rect-active' : ''}" data-form-key="${escapeHTML(form.key)}"
                style="--kb-theme: ${themeColor}; --kb-theme-bg: ${themeBg}; --kb-theme-border: ${themeBorder}; --kb-theme-icon-bg: ${themeIconBg};">
                <div class="kb-rect-icon">${form.icon}</div>
                <div class="kb-rect-body">
                    <h3 class="kb-rect-title">${escapeHTML(form.title)}</h3>
                    <p class="kb-rect-subtitle">${escapeHTML(form.subtitle)}</p>
                </div>
                <div class="kb-rect-meta">
                    <span class="kb-rect-count">${countLabel}</span>
                    <span class="kb-rect-status">${isActive ? '✓' : '→'}</span>
                </div>
            </button>
        `;
    });

    // Get active form name for sub-section heading
    const activeForm = formsData.find(f => f.category === activeCategory);
    const activeTitle = activeForm ? activeForm.title : '';

    html += `
            </div>
        </div>

        <div class="kb-projects-header">
            <h2 class="kb-section-heading">${lang === 'hi' ? 'प्रोजेक्ट्स' : 'Projects'} — ${escapeHTML(activeTitle)}</h2>
            <p class="kb-section-subtext">${lang === 'hi' ? 'नीचे दिए गए प्रोजेक्ट कार्ड पर क्लिक करें' : 'Click on any project card below for full details'}</p>
        </div>
    `;

    container.innerHTML = html;

    container.querySelectorAll('.kb-form-rect-card').forEach(card => {
        card.addEventListener('click', () => {
            const formKey = card.getAttribute('data-form-key');
            state.filterSubject = formKey;
            renderKaushalTabs();
            renderFilteredActivities();
        });
    });
}

/**
 * Renders complete Kaushal Bodh Project Card with all Excel Column Details directly inside the box.
 * Includes Key Activities, Key Learning Outcomes, Adobe Express Integration, Adobe Express Activity, Deliverable,
 * and ends with the Learning Journal CTA button at the bottom.
 */
function renderKaushalBodhCardContent(activity, dict) {
    const info = getFormOfWorkInfo(activity.form_of_work);

    return `
        <div class="activity-card card-accent-indigo kb-project-tile" data-activity-name="${encodeURIComponent(activity.activity_name)}" style="display: flex; flex-direction: column; justify-content: space-between; padding: 1.5rem; border-radius: var(--radius-lg); background: var(--color-surface); border: 1px solid var(--color-border); box-shadow: var(--shadow-sm); transition: all 0.2s ease; cursor: pointer;">
            <div>
                <!-- Top Header Pill & Duration -->
                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 0.85rem;">
                    <span class="kb-form-pill ${info.colorClass}">
                        ${info.icon} ${activity.form_of_work}
                    </span>
                    ${activity.duration ? `<span style="font-size: 0.775rem; color: var(--color-text-secondary); font-weight: 600;">⏱️ ${activity.duration}</span>` : ''}
                </div>

                <!-- Project Title -->
                <h2 class="activity-name" style="font-size: 1.25rem; font-weight: 800; color: var(--color-text-primary); line-height: 1.35; margin: 0;">
                    ${escapeHTML(activity.activity_name)}
                </h2>
            </div>
        </div>
    `;
}

/**
 * Renders the 3 Form of Work cards on the dedicated step-form-of-work page.
 * Uses the exact same grade-card CSS classes for visual consistency.
 */
function renderFormOfWorkCards() {
    const fowGrid = document.getElementById('fow-grid');
    if (!fowGrid) return;

    const lang = state.language || 'en';
    const dict = translations[lang];
    const gradeList = state.kaushalBodhActivities.filter(a => a.grade === state.selectedGrade);

    // Default selection to 1st Form of Work if not selected
    if (!state.filterSubject || state.filterSubject === 'All') {
        state.filterSubject = 'Work with Life Forms';
    }

    const formsData = [
        {
            key: 'Work with Life Forms',
            category: 'life',
            title: lang === 'hi' ? 'सजीव रूपों के साथ कार्य' : 'Work with Life Forms',
            subtitle: lang === 'hi' ? 'पौधे, जीव विविधता और पर्यावरण' : 'Plants, Animals & Ecology',
            icon: '🌿',
            accentClass: 'pink'
        },
        {
            key: 'Work with Machines & Materials',
            category: 'machines',
            title: lang === 'hi' ? 'मशीनों और सामग्रियों के साथ कार्य' : 'Work with Machines & Materials',
            subtitle: lang === 'hi' ? 'उपकरण, निर्माण और तकनीक' : 'Tools, Fabrication & Tech',
            icon: '⚙️',
            accentClass: 'indigo'
        },
        {
            key: 'Work in Human Services',
            category: 'human',
            title: lang === 'hi' ? 'मानव सेवाओं में कार्य' : 'Work in Human Services',
            subtitle: lang === 'hi' ? 'स्वास्थ्य, समाज और संचार' : 'Health, Society & Service',
            icon: '🤝',
            accentClass: 'green'
        }
    ];

    // 1. Render Top Form of Work Choice Cards
    let html = '';
    formsData.forEach(form => {
        const count = gradeList.filter(a => matchesFormCategory(a.form_of_work, form.category)).length;
        const template = count === 1
            ? (lang === 'hi' ? '1 प्रोजेक्ट' : '1 project')
            : (lang === 'hi' ? `${count} प्रोजेक्ट्स` : `${count} projects`);
        const isActive = matchesFormCategory(state.filterSubject, form.category);

        html += `
            <button class="grade-card card-accent-${form.accentClass} ${isActive ? 'active' : ''}" type="button" data-form-key="${escapeHTML(form.key)}" aria-label="${form.title}, ${template}">
                <div class="grade-icon-wrapper icon-bg-${form.accentClass}">
                    <span class="grade-icon-emoji">${form.icon}</span>
                </div>
                <h3 class="grade-card-title">${escapeHTML(form.title)}</h3>
                <p class="grade-card-subtitle">${template}</p>
                <div class="grade-card-footer">
                    <span class="explore-grade-text">${isActive ? (lang === 'hi' ? '✓ चयनित' : '✓ SELECTED') : (lang === 'hi' ? 'प्रोजेक्ट्स देखें' : 'EXPLORE PROJECTS')}</span>
                    <span class="arrow-icon">→</span>
                </div>
            </button>
        `;
    });

    fowGrid.innerHTML = html;

    // Attach click listeners to Top Form of Work Cards
    fowGrid.querySelectorAll('.grade-card').forEach(card => {
        card.addEventListener('click', () => {
            const formKey = card.getAttribute('data-form-key');
            state.filterSubject = formKey;
            renderFormOfWorkCards();
        });
    });

    // 2. Render Sub-Breadcrumb Below Border Line
    const activeCategory = getFormCategory(state.filterSubject);
    const activeForm = formsData.find(f => f.category === activeCategory) || formsData[0];
    const subBreadcrumb = document.getElementById('fow-sub-breadcrumb');
    if (subBreadcrumb) {
        const homeText = dict.homeBreadcrumb || 'Home';
        const displayGrade = getDisplayGradeName(state.selectedGrade);

        subBreadcrumb.innerHTML = `
            <span>${escapeHTML(homeText)}</span>
            <span style="opacity: 0.5; font-size: 0.8rem;">›</span>
            <span>${escapeHTML(displayGrade)}</span>
            <span style="opacity: 0.5; font-size: 0.8rem;">›</span>
            <span style="color: #eb1000; font-weight: 700;">${activeForm.icon} ${escapeHTML(activeForm.title)}</span>
        `;
    }

    // 3. Render Project KPI Cards Grid Below Border Line
    const fowActivitiesGrid = document.getElementById('fow-activities-grid');
    if (fowActivitiesGrid) {
        const matchingProjects = gradeList.filter(a => matchesFormCategory(a.form_of_work, activeCategory));

        if (matchingProjects.length === 0) {
            fowActivitiesGrid.innerHTML = `
                <div class="empty-card" style="grid-column: 1 / -1; margin: 1.5rem 0;">
                    <span class="empty-icon">🔍</span>
                    <h2>${dict.noActivitiesFound || 'No Projects Found'}</h2>
                </div>
            `;
        } else {
            let projectsHtml = '';
            matchingProjects.forEach(activity => {
                projectsHtml += renderKaushalBodhCardContent(activity, dict);
            });

            fowActivitiesGrid.innerHTML = projectsHtml;

            // Attach click listener directly to project tiles
            fowActivitiesGrid.querySelectorAll('.kb-project-tile').forEach(card => {
                card.addEventListener('click', () => {
                    const name = decodeURIComponent(card.getAttribute('data-activity-name'));
                    const selected = state.kaushalBodhActivities.find(
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

    const isKB = state.tutorialMode === 'kaushal-bodh';
    const sourceList = isKB ? state.kaushalBodhActivities : state.allActivities;
    let list = sourceList.filter(a => a.grade === state.selectedGrade);

    if (state.filterSubject !== 'All') {
        if (isKB) {
            const cat = getFormCategory(state.filterSubject);
            list = list.filter(a => matchesFormCategory(a.form_of_work, cat));
        } else {
            list = list.filter(a => a.month === state.filterSubject);
        }
    }

    if (list.length === 0) {
        activitiesGrid.classList.add('hidden');
        emptyView.classList.remove('hidden');
    } else {
        emptyView.classList.add('hidden');
        activitiesGrid.classList.remove('hidden');

        let html = '';
        const dict = translations[state.language || 'en'];

        list.forEach(activity => {
            if (isKB) {
                html += renderKaushalBodhCardContent(activity, dict);
            } else {
                const colorClass = getSubjectColorClass(activity.month);
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
            }
        });

        activitiesGrid.innerHTML = html;

        if (isKB) {
            activitiesGrid.querySelectorAll('.kb-project-tile').forEach(card => {
                card.addEventListener('click', () => {
                    const name = decodeURIComponent(card.getAttribute('data-activity-name'));
                    const selected = state.kaushalBodhActivities.find(
                        a => a.grade === state.selectedGrade && a.activity_name === name
                    );
                    if (selected) {
                        state.selectedActivity = selected;
                        state.currentStep = 'detail';
                        render();
                    }
                });
            });
        } else {
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

    if (activity.isKaushalBodh) {
        const detailContainer = document.getElementById('step-detail');
        if (!detailContainer) return;

        const info = getFormOfWorkInfo(activity.form_of_work);
        const journalLink = activity.learning_journal_link ? activity.learning_journal_link.trim() : '';

        // Update back link text
        const backLinkText = document.querySelector('#back-to-activities-link span:not(.arrow)');
        if (backLinkText) {
            const lang = state.language || 'en';
            backLinkText.innerText = lang === 'hi' ? 'कार्य का प्रकार चुनें' : 'Select Form of Work';
        }

        // Header pills
        const detailTags = document.getElementById('detail-tags');
        if (detailTags) {
            detailTags.innerHTML = `
                <span class="kb-form-pill ${info.colorClass}">${info.icon} ${activity.form_of_work}</span>
                <span class="grade-badge">${getDisplayGradeName(activity.grade)}</span>
                ${activity.duration ? `<span style="font-size: 0.85rem; color: var(--color-text-secondary); font-weight: 600; background: #f1f5f9; padding: 0.3rem 0.75rem; border-radius: var(--radius-full);">⏱️ ${activity.duration}</span>` : ''}
            `;
        }

        document.getElementById('detail-title').innerText = activity.activity_name;

        // Custom Layout for Kaushal Bodh Columns
        const descElem = document.getElementById('detail-description');
        if (descElem) {
            descElem.innerHTML = `
                <div style="display: flex; flex-direction: column; gap: 1.25rem; margin-top: 1.25rem;">
                    <!-- Row 1: Box 1 (Key Activities) & Box 2 (Key Learning Outcomes) -->
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 1.25rem;">
                        <!-- Box 1: Key Activities (Indigo Theme) -->
                        <div style="background: linear-gradient(135deg, rgba(79, 70, 229, 0.05), rgba(79, 70, 229, 0.12)); border: 1.5px solid rgba(79, 70, 229, 0.25); border-radius: var(--radius-lg); padding: 1.35rem; display: flex; flex-direction: column;">
                            <span style="font-size: 0.775rem; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 800; color: #4f46e5; display: flex; align-items: center; gap: 0.35rem; margin-bottom: 0.65rem;">
                                📌 ${dict.keyActivitiesLabel || "Key Activities"}
                            </span>
                            <div style="font-size: 0.95rem; color: var(--color-text-primary); line-height: 1.5;">
                                ${formatFormattedText(activity.key_activities)}
                            </div>
                        </div>

                        <!-- Box 2: Key Learning Outcomes (Emerald Theme) -->
                        <div style="background: linear-gradient(135deg, rgba(16, 185, 129, 0.05), rgba(16, 185, 129, 0.12)); border: 1.5px solid rgba(16, 185, 129, 0.25); border-radius: var(--radius-lg); padding: 1.35rem; display: flex; flex-direction: column;">
                            <span style="font-size: 0.775rem; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 800; color: #059669; display: flex; align-items: center; gap: 0.35rem; margin-bottom: 0.65rem;">
                                🎯 ${dict.keyLearningOutcomesLabel || "Key Learning Outcomes"}
                            </span>
                            <div style="font-size: 0.95rem; color: var(--color-text-primary); line-height: 1.5;">
                                ${formatFormattedText(activity.key_learning_outcomes)}
                            </div>
                        </div>
                    </div>

                    <!-- Row 2: Box 3 (Adobe Express Integration) & Box 4 (Adobe Express Activity) -->
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 1.25rem;">
                        <!-- Box 3: Adobe Express Integration (Purple Theme) -->
                        <div style="background: linear-gradient(135deg, rgba(147, 51, 234, 0.05), rgba(147, 51, 234, 0.12)); border: 1.5px solid rgba(147, 51, 234, 0.25); border-radius: var(--radius-lg); padding: 1.35rem; display: flex; flex-direction: column;">
                            <span style="font-size: 0.775rem; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 800; color: #7c3aed; display: flex; align-items: center; gap: 0.35rem; margin-bottom: 0.65rem;">
                                🎨 ${dict.adobeExpressIntegrationLabel || "Adobe Express Integration for Activity Books & Portfolios"}
                            </span>
                            <div style="font-size: 0.95rem; color: var(--color-text-primary); line-height: 1.5;">
                                ${formatFormattedText(activity.adobe_express_integration)}
                            </div>
                        </div>

                        <!-- Box 4: Adobe Express Activity (Red/Coral Theme) -->
                        <div style="background: linear-gradient(135deg, rgba(235, 16, 0, 0.05), rgba(235, 16, 0, 0.12)); border: 1.5px solid rgba(235, 16, 0, 0.25); border-radius: var(--radius-lg); padding: 1.35rem; display: flex; flex-direction: column;">
                            <span style="font-size: 0.775rem; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 800; color: #eb1000; display: flex; align-items: center; gap: 0.35rem; margin-bottom: 0.65rem;">
                                💡 ${dict.adobeExpressActivityLabel || "Adobe Express Activity"}
                            </span>
                            <div style="font-size: 1.05rem; font-weight: 700; color: var(--color-text-primary); line-height: 1.4;">
                                ${escapeHTML(activity.adobe_express_activity)}
                            </div>
                        </div>
                    </div>

                    <!-- Row 3: Box 5 (Final Student Deliverable) Full Width -->
                    <div style="background: linear-gradient(135deg, rgba(59, 130, 246, 0.05), rgba(59, 130, 246, 0.12)); border: 1.5px solid rgba(59, 130, 246, 0.25); border-radius: var(--radius-lg); padding: 1.35rem;">
                        <span style="font-size: 0.775rem; text-transform: uppercase; letter-spacing: 0.06em; font-weight: 800; color: #2563eb; display: flex; align-items: center; gap: 0.35rem; margin-bottom: 0.65rem;">
                            📦 ${dict.finalStudentDeliverableLabel || "Final Student Deliverable"}
                        </span>
                        <div style="font-size: 1.05rem; font-weight: 700; color: var(--color-text-primary); line-height: 1.4;">
                            ${escapeHTML(activity.final_student_deliverable)}
                        </div>
                    </div>
                </div>
            `;
        }

        // Hide default skills/instructions section & book screenshot container for Kaushal Bodh
        const skillsInstSection = document.getElementById('detail-skills-instructions-section');
        if (skillsInstSection) skillsInstSection.style.display = 'none';
        const bookScreenshotContainer = document.getElementById('detail-book-screenshot-container');
        if (bookScreenshotContainer) bookScreenshotContainer.style.display = 'none';

        // CTA Container for Learning Journal Link
        const ctaContainer = document.getElementById('tutorial-cta-container');
        if (ctaContainer) {
            let ctaHtml = `
                <div style="width: 100%; padding: 1.5rem; background: radial-gradient(circle at top left, #eff6ff, #dbeafe); border-radius: var(--radius-xl); border: 2px solid #3b82f6; display: flex; flex-direction: column; align-items: center; text-align: center; gap: 1rem; box-shadow: 0 8px 24px rgba(59, 130, 246, 0.15);">
                    <div style="display: flex; align-items: center; gap: 0.5rem; color: #1d4ed8; font-weight: 800; font-size: 1.1rem; text-transform: uppercase; letter-spacing: 0.03em;">
                        <span style="font-size: 1.5rem;">📘</span> Learning Journal Link
                    </div>
                    <p style="margin: 0; font-size: 0.95rem; color: #1e3a8a; max-width: 600px; line-height: 1.5;">
                        Open your dedicated Adobe Express Remix link for this project's Learning Journal to start documenting your work!
                    </p>
                    <div style="display: flex; gap: 1rem; flex-wrap: wrap; justify-content: center; width: 100%;">
                        ${journalLink ? `
                            <a href="${journalLink}" target="_blank" rel="noopener noreferrer" class="btn btn-primary kb-journal-btn" style="background: linear-gradient(135deg, #2563eb, #1d4ed8); border: none; padding: 0.85rem 1.75rem; font-size: 1.05rem; font-weight: 800; border-radius: var(--radius-full); box-shadow: 0 4px 16px rgba(37, 99, 235, 0.35); text-decoration: none; color: white; display: inline-flex; align-items: center; gap: 0.6rem;">
                                <span>📘</span> ${dict.learningJournalBtn || "Open Learning Journal Link"} <span style="font-size: 1.1rem;">↗</span>
                            </a>
                        ` : ''}
                    </div>
                </div>
            `;
            ctaContainer.innerHTML = ctaHtml;
        }
        return;
    }

    // Ensure standard containers are visible for normal DCAIS activities
    const skillsInstSection = document.getElementById('detail-skills-instructions-section');
    if (skillsInstSection) skillsInstSection.style.display = 'block';

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

    // Skills and instructions text blocks
    const descElem = document.getElementById('detail-description');
    if (descElem) {
        descElem.innerText = activity.description || dict.noDescription;
    }
    document.getElementById('detail-skills').innerText = activity.skills || dict.notSpecified;
    document.getElementById('detail-instructions').innerText = activity.instructions || dict.noInstructions;

    // Book Screenshot rendering (right below Instructions)
    const bookScreenshotContainer = document.getElementById('detail-book-screenshot-container');
    if (bookScreenshotContainer) {
        const bookScreenshot = activity.book_screenshot ? activity.book_screenshot.trim() : '';
        if (bookScreenshot && (bookScreenshot.startsWith('http://') || bookScreenshot.startsWith('https://'))) {
            bookScreenshotContainer.innerHTML = `
                <h2 class="detail-section-title" style="margin-top: 1.5rem; margin-bottom: 0.75rem;">Book Screenshot</h2>
                <a href="${bookScreenshot}" target="_blank" rel="noopener noreferrer" class="btn btn-secondary" style="display: inline-flex; align-items: center; gap: 0.5rem; background: #ffffff; border: 1.5px solid #eb1000; color: #eb1000; padding: 0.6rem 1.25rem; border-radius: var(--radius-md); font-weight: 600; font-size: 0.95rem; text-decoration: none; transition: all 0.2s ease; box-shadow: 0 2px 8px rgba(235, 16, 0, 0.08);">
                    <span style="font-size: 1.1rem;">📖</span> View Book Screenshot
                </a>
            `;
            bookScreenshotContainer.style.display = 'block';
        } else {
            bookScreenshotContainer.innerHTML = '';
            bookScreenshotContainer.style.display = 'none';
        }
    }

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
 * Converts a Google Drive share/view URL into a direct image preview URL.
 * @param {string} url - Google Drive URL or direct image URL.
 * @returns {string} Direct image display URL.
 */
function getGoogleDriveImageUrl(url) {
    if (!url) return '';
    const cleanUrl = url.trim();
    const match = cleanUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || cleanUrl.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
        return `https://lh3.googleusercontent.com/d/${match[1]}`;
    }
    return cleanUrl;
}

/**
 * Handles processing of a student work submission.
 * Saves values into LocalStorage state, closes the modal popup, and resets controls.
 * @returns {void}
 */
async function handleStudentSubmission() {
    const nameInput = document.getElementById('submit-student-name');
    const gradeInput = document.getElementById('submit-student-grade');
    const activityNameInput = document.getElementById('submit-activity-name');
    const schoolInput = document.getElementById('submit-school-name');
    const linkInput = document.getElementById('submit-work-link');

    const link = linkInput.value.trim();
    if (!link) {
        return; // Handled by HTML5 input required validation
    }

    // Validate if it is a proper Adobe Express / Adobe link
    const adobeLinkRegex = /^https?:\/\/(?:[a-zA-Z0-9-]+\.)*(?:express\.adobe\.com|adobe\.com|adobe\.ly|adobeexpress\.com|adobeexpress)(?:\/.*)?$/i;
    if (!adobeLinkRegex.test(link)) {
        const errorMsg = state.language === 'hi' 
            ? "कृपया एक मान्य एडोब एक्सप्रेस या एडोब प्रोजेक्ट लिंक दर्ज करें (जैसे: https://new.express.adobe.com/...)"
            : "Please enter a valid Adobe Express or Adobe project link (e.g. https://new.express.adobe.com/...)";
        linkInput.setCustomValidity(errorMsg);
        linkInput.classList.add('input-error');
        linkInput.reportValidity();
        return;
    }

    // Local Uniqueness check: verify if the link has already been submitted
    const normLink = link.toLowerCase();
    const storedSubmissions = JSON.parse(localStorage.getItem('aim_submissions') || '[]');
    const allLocalSubmissions = [...state.submissions, ...storedSubmissions];
    const isLocalDuplicate = allLocalSubmissions.some(s => {
        const existing = (s["Work Link"] || s.workLink || "").trim().toLowerCase();
        return existing === normLink;
    });

    if (isLocalDuplicate) {
        const dupMsg = state.language === 'hi'
            ? "यह प्रोजेक्ट लिंक पहले ही जमा किया जा चुका है! कृपया एक नया (यूनिक) प्रोजेक्ट लिंक दर्ज करें।"
            : "This project link has already been submitted! Please enter a unique project link.";
        linkInput.setCustomValidity(dupMsg);
        linkInput.classList.add('input-error');
        linkInput.reportValidity();
        return;
    }

    linkInput.setCustomValidity("");
    linkInput.classList.remove('input-error');

    const activity = state.selectedActivity;
    if (!activity) return;

    const studentName = nameInput.value.trim() || "Anonymous";
    const gradeValue = gradeInput.value.trim() || activity.grade;
    const activityNameValue = activityNameInput ? activityNameInput.value.trim() : activity.activity_name;
    const schoolName = schoolInput.value.trim() || "Not Provided";

    // Setup visual loading state on submit button
    const submitBtn = document.querySelector('#submission-form button[type="submit"]');
    const cancelBtn = document.getElementById('cancel-submission-btn');
    let originalBtnText = 'Submit Project';
    if (submitBtn) {
        originalBtnText = submitBtn.textContent;
        submitBtn.disabled = true;
        submitBtn.textContent = state.language === 'hi' ? 'जमा किया जा रहा है...' : 'Submitting...';
    }
    if (cancelBtn) {
        cancelBtn.disabled = true;
    }

    // Remote Google Sheets Uniqueness Check & Submission
    if (GOOGLE_SCRIPT_URL && GOOGLE_SCRIPT_URL.trim() !== "") {
        try {
            // Check for duplicates in Google Sheets remotely before submitting
            const checkUrl = `${GOOGLE_SCRIPT_URL}?action=checkDuplicateLink&workLink=${encodeURIComponent(link)}`;
            const checkResponse = await fetchJSONP(checkUrl);
            if (checkResponse && checkResponse.isDuplicate) {
                const sheetDupMsg = state.language === 'hi'
                    ? "यह प्रोजेक्ट लिंक पहले ही Google Sheets में जमा है! कृपया एक नया प्रोजेक्ट लिंक दर्ज करें।"
                    : "This project link has already been submitted in Google Sheets! Please enter a unique project link.";
                linkInput.setCustomValidity(sheetDupMsg);
                linkInput.classList.add('input-error');
                linkInput.reportValidity();

                if (submitBtn) {
                    submitBtn.disabled = false;
                    submitBtn.textContent = originalBtnText;
                }
                if (cancelBtn) cancelBtn.disabled = false;
                return;
            }
        } catch (chkErr) {
            console.warn("Remote duplicate check skipped or failed, proceeding to submit:", chkErr);
        }
    }

    const newSubmission = {
        Timestamp: new Date().toLocaleString(),
        Grade: gradeValue,
        "Activity Name": activityNameValue,
        "Student Name": studentName,
        "School Name": schoolName,
        "Work Link": link
    };

    let apiSuccess = false;

    // Send to Google Sheets if configured
    if (GOOGLE_SCRIPT_URL && GOOGLE_SCRIPT_URL.trim() !== "") {
        try {
            const queryUrl = `${GOOGLE_SCRIPT_URL}?action=submitStudentWork` +
                `&grade=${encodeURIComponent(gradeValue)}` +
                `&activityName=${encodeURIComponent(activityNameValue)}` +
                `&studentName=${encodeURIComponent(studentName)}` +
                `&schoolName=${encodeURIComponent(schoolName)}` +
                `&workLink=${encodeURIComponent(link)}`;

            const response = await fetchJSONP(queryUrl);
            if (response && response.isDuplicate) {
                const sheetDupMsg = state.language === 'hi'
                    ? "यह प्रोजेक्ट लिंक पहले ही जमा किया जा चुका है! कृपया एक नया (यूनिक) प्रोजेक्ट लिंक दर्ज करें।"
                    : "This project link has already been submitted! Please enter a unique project link.";
                linkInput.setCustomValidity(sheetDupMsg);
                linkInput.classList.add('input-error');
                linkInput.reportValidity();

                if (submitBtn) {
                    submitBtn.disabled = false;
                    submitBtn.textContent = originalBtnText;
                }
                if (cancelBtn) cancelBtn.disabled = false;
                return;
            }

            if (response && response.success) {
                apiSuccess = true;
            } else {
                console.error("Google Sheets submit failed: ", response);
            }
        } catch (apiError) {
            console.error("Google Sheets submit network error: ", apiError);
        }
    }

    // Resilient Fallback to LocalStorage
    state.submissions.push(newSubmission);
    localStorage.setItem('aim_submissions', JSON.stringify(state.submissions));

    // Reset button states
    if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = originalBtnText;
    }
    if (cancelBtn) {
        cancelBtn.disabled = false;
    }

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
        document.getElementById('submit-student-grade').value = getDisplayGradeName(state.selectedGrade);
        document.getElementById('submit-activity-name').value = state.selectedActivity ? state.selectedActivity.activity_name : '';
        document.getElementById('submit-school-name').value = '';
        
        const linkInput = document.getElementById('submit-work-link');
        if (linkInput) {
            linkInput.value = '';
            linkInput.setCustomValidity('');
            linkInput.classList.remove('input-error');
        }
        
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
function resetSubmissionUI() {
    closeSubmissionModal();
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
 * Updates active class styling on header language toggle buttons.
 * @returns {void}
 */
function updateLangToggleUI() {
    const langToggleEn = document.getElementById('lang-toggle-en');
    const langToggleHi = document.getElementById('lang-toggle-hi');
    if (!langToggleEn || !langToggleHi) return;

    if (state.language === 'hi') {
        langToggleHi.classList.add('active-lang');
        langToggleEn.classList.remove('active-lang');
    } else {
        langToggleEn.classList.add('active-lang');
        langToggleHi.classList.remove('active-lang');
    }
}

/**
 * Translates static and metadata elements in the DOM based on active state.language.
 * @returns {void}
 */
function translateUIInternal() {
    const lang = state.language || 'en';
    const dict = translations[lang] || translations.en;

    // Synchronize language toggle UI buttons
    updateLangToggleUI();

    // Translate all standard data-i18n elements
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (dict && dict[key]) {
            el.innerText = dict[key];
        }
    });

    // Translate placeholder attributes
    document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
        const key = el.getAttribute('data-i18n-placeholder');
        if (dict && dict[key]) {
            el.setAttribute('placeholder', dict[key]);
        }
    });

    // Translate aria-label attributes
    document.querySelectorAll('[data-i18n-aria]').forEach(el => {
        const key = el.getAttribute('data-i18n-aria');
        if (dict && dict[key]) {
            el.setAttribute('aria-label', dict[key]);
        }
    });
}

/**
 * JSONP Fetch Helper to bypass CORS redirection blocks from Google script webapps.
 * @param {string} url - Target URL.
 * @returns {Promise<any>} Response payload.
 */
function fetchJSONP(url) {
    return new Promise((resolve, reject) => {
        const callbackName = 'jsonp_cb_' + Math.round(100000 * Math.random());
        const separator = url.includes('?') ? '&' : '?';
        const scriptUrl = `${url}${separator}prefix=${callbackName}`;
        
        const timeoutId = setTimeout(() => {
            cleanup();
            reject(new Error('Network request timed out'));
        }, 10000);

        window[callbackName] = (data) => {
            clearTimeout(timeoutId);
            resolve(data);
            cleanup();
        };
        
        const script = document.createElement('script');
        script.src = scriptUrl;
        script.id = callbackName;
        script.async = true;
        
        script.onerror = () => {
            clearTimeout(timeoutId);
            reject(new Error('Failed to load resource from server'));
            cleanup();
        };
        
        document.head.appendChild(script);
        
        function cleanup() {
            const element = document.getElementById(callbackName);
            if (element) {
                element.remove();
            }
            delete window[callbackName];
        }
    });
}

// Kickoff
document.addEventListener('DOMContentLoaded', init);
