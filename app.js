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
    currentStep: 'welcome',
    filterSubject: 'All', // Mapped to Month or Form of Work filtering
    submissions: [],
    language: 'en', // 'en' or 'hi'
    userExplicitLanguage: 'en', // Manually set via language toggle
    tutorialMode: 'choose', // 'choose' | 'hindi' | 'english' | 'kaushal-bodh' | 'ict-english' | 'ict-hindi'
    curriculumType: 'dcais', // 'dcais' | 'ict' | 'kaushal-bodh'
    aotmLanguage: 'en'
};

const MONTH_TRANSLATION_MAP = {
    // English
    'january': { en: 'January', hi: 'जनवरी', altHi: ['जनवरी'], index: 1 },
    'jan': { en: 'January', hi: 'जनवरी', altHi: ['जनवरी'], index: 1 },
    'february': { en: 'February', hi: 'फ़रवरी', altHi: ['फरवरी', 'फ़रवरी'], index: 2 },
    'feb': { en: 'February', hi: 'फ़रवरी', altHi: ['फरवरी', 'फ़रवरी'], index: 2 },
    'march': { en: 'March', hi: 'मार्च', altHi: ['मार्च'], index: 3 },
    'mar': { en: 'March', hi: 'मार्च', altHi: ['मार्च'], index: 3 },
    'april': { en: 'April', hi: 'अप्रैल', altHi: ['अप्रैल', 'अप्रेल'], index: 4 },
    'apr': { en: 'April', hi: 'अप्रैल', altHi: ['अप्रैल', 'अप्रेल'], index: 4 },
    'may': { en: 'May', hi: 'मई', altHi: ['मई'], index: 5 },
    'june': { en: 'June', hi: 'जून', altHi: ['जून'], index: 6 },
    'jun': { en: 'June', hi: 'जून', altHi: ['जून'], index: 6 },
    'july': { en: 'July', hi: 'जुलाई', altHi: ['जुलाई'], index: 7 },
    'jul': { en: 'July', hi: 'जुलाई', altHi: ['जुलाई'], index: 7 },
    'august': { en: 'August', hi: 'अगस्त', altHi: ['अगस्त'], index: 8 },
    'aug': { en: 'August', hi: 'अगस्त', altHi: ['अगस्त'], index: 8 },
    'september': { en: 'September', hi: 'सितम्बर', altHi: ['सितंबर', 'सितम्बर'], index: 9 },
    'sep': { en: 'September', hi: 'सितम्बर', altHi: ['सितंबर', 'सितम्बर'], index: 9 },
    'sept': { en: 'September', hi: 'सितम्बर', altHi: ['सितंबर', 'सितम्बर'], index: 9 },
    'october': { en: 'October', hi: 'अक्टूबर', altHi: ['अक्टूबर', 'अक्तूबर'], index: 10 },
    'oct': { en: 'October', hi: 'अक्टूबर', altHi: ['अक्टूबर', 'अक्तूबर'], index: 10 },
    'november': { en: 'November', hi: 'नवंबर', altHi: ['नवंबर', 'नवम्बर'], index: 11 },
    'nov': { en: 'November', hi: 'नवंबर', altHi: ['नवंबर', 'नवम्बर'], index: 11 },
    'december': { en: 'December', hi: 'दिसंबर', altHi: ['दिसंबर', 'दिसम्बर'], index: 12 },
    'dec': { en: 'December', hi: 'दिसंबर', altHi: ['दिसंबर', 'दिसम्बर'], index: 12 },

    // Hindi
    'जनवरी': { en: 'January', hi: 'जनवरी', altHi: ['जनवरी'], index: 1 },
    'फरवरी': { en: 'February', hi: 'फ़रवरी', altHi: ['फरवरी', 'फ़रवरी'], index: 2 },
    'फ़रवरी': { en: 'February', hi: 'फ़रवरी', altHi: ['फरवरी', 'फ़रवरी'], index: 2 },
    'मार्च': { en: 'March', hi: 'मार्च', altHi: ['मार्च'], index: 3 },
    'अप्रैल': { en: 'April', hi: 'अप्रैल', altHi: ['अप्रैल', 'अप्रेल'], index: 4 },
    'अप्रेल': { en: 'April', hi: 'अप्रैल', altHi: ['अप्रैल', 'अप्रेल'], index: 4 },
    'मई': { en: 'May', hi: 'मई', altHi: ['मई'], index: 5 },
    'जून': { en: 'June', hi: 'जून', altHi: ['जून'], index: 6 },
    'जुलाई': { en: 'July', hi: 'जुलाई', altHi: ['जुलाई'], index: 7 },
    'अगस्त': { en: 'August', hi: 'अगस्त', altHi: ['अगस्त'], index: 8 },
    'सितंबर': { en: 'September', hi: 'सितम्बर', altHi: ['सितंबर', 'सितम्बर'], index: 9 },
    'सितम्बर': { en: 'September', hi: 'सितम्बर', altHi: ['सितंबर', 'सितम्बर'], index: 9 },
    'अक्टूबर': { en: 'October', hi: 'अक्टूबर', altHi: ['अक्टूबर', 'अक्तूबर'], index: 10 },
    'अक्तूबर': { en: 'October', hi: 'अक्टूबर', altHi: ['अक्टूबर', 'अक्तूबर'], index: 10 },
    'नवंबर': { en: 'November', hi: 'नवंबर', altHi: ['नवंबर', 'नवम्बर'], index: 11 },
    'नवम्बर': { en: 'November', hi: 'नवंबर', altHi: ['नवंबर', 'नवम्बर'], index: 11 },
    'दिसंबर': { en: 'December', hi: 'दिसंबर', altHi: ['दिसंबर', 'दिसम्बर'], index: 12 },
    'दिसम्बर': { en: 'December', hi: 'दिसंबर', altHi: ['दिसंबर', 'दिसम्बर'], index: 12 }
};

const MONTHLY_ACTIVITIES = {
    en: {
        title: "Happy Children's Day! ✨",
        desc: "<p style=\"margin: 0 0 0.25rem 0; line-height: 1.38;\">Think of a magical childhood memory that still makes you smile—celebrating Diwali with family, playing with friends, or summer holidays at grandparents’ home. Write about what happened, what made the moment special, how you felt, and why you will always remember it. Then watch the tutorial to transform your memory into a beautiful visual.</p><div style=\"color: #eb1000; font-weight: 700; font-size: 0.82rem;\">Imagine. Create. Express your story!</div>",
        templateUrl: "https://new.express.adobe.com/design/template/urn:aaid:sc:VA6C2:fcf64e48-8456-5141-9d68-55b0d404d468?category=text&entryPoint=template&taskID=poster",
        videoUrl: "https://youtu.be/eRA7DjSb5gs",
        videoId: "eRA7DjSb5gs",
        videoTitle: "Tutorial",
        videoSubtitle: "Watch before you design",
        btnText: "Activity Template",
        watchBtnText: "▶ Watch"
    },
    hi: {
        title: "बाल दिवस की शुभकामनाएँ! ✨",
        desc: "<p style=\"margin: 0 0 0.25rem 0; line-height: 1.38;\">Adobe Express के साथ अपने बचपन की यादों को जीवंत बनाएँ! अपने बचपन की किसी जादुई और प्यारी याद के बारे में सोचिए, जिसे याद करके आज भी आपके चेहरे पर मुस्कान आ जाती है। यह परिवार के साथ दिवाली मनाने की याद हो सकती है, दोस्तों के साथ खेलने की या नाना-नानी अथवा दादा-दादी के घर गर्मियों की छुट्टियाँ बिताने की। लिखिए कि उस दिन क्या हुआ था, वह पल आपके लिए खास क्यों था, आपको कैसा महसूस हुआ और आप उसे हमेशा क्यों याद रखना चाहेंगे। इसके बाद ट्यूटोरियल देखिए और सीखिए कि Adobe Express की मदद से अपनी बचपन की याद को एक सुंदर और रचनात्मक चित्र में कैसे बदलें।</p><div style=\"color: #eb1000; font-weight: 700; font-size: 0.82rem;\">कल्पना करें। रचना करें। अपनी कहानी को रचनात्मक रंगों में सजाएँ!</div>",
        templateUrl: "https://new.express.adobe.com/design/template/urn:aaid:sc:VA6C2:3fda717c-85e0-57ed-8cab-bf340edd76d9?category=text&entryPoint=template&taskID=poster",
        videoUrl: "https://youtu.be/M1s8vTaJLrY",
        videoId: "M1s8vTaJLrY",
        videoTitle: "ट्यूटोरियल",
        videoSubtitle: "डिज़ाइन करने से पहले देखें",
        btnText: "गतिविधि टेम्पलेट",
        watchBtnText: "▶ देखें"
    }
};

const translations = {
    en: {
        welcomeTitle: "Welcome to AIM",
        welcomeDesc: "AIM is a creative learning space for students in India. Every week you get fresh activities built on Adobe Express. Design posters, generate images, create videos, and build things that are yours.",
        activityOfMonthTag: "Activity of the Month",
        activityOfMonthTitle: "Happy Children's Day! ✨",
        independenceTitle: "Happy Children's Day! ✨",
        independenceDesc: "<p style=\"margin: 0 0 0.25rem 0; line-height: 1.38;\">Think of a magical childhood memory that still makes you smile—celebrating Diwali with family, playing with friends, or summer holidays at grandparents’ home. Write about what happened, what made the moment special, how you felt, and why you will always remember it. Then watch the tutorial to transform your memory into a beautiful visual.</p><div style=\"color: #eb1000; font-weight: 700; font-size: 0.82rem;\">Imagine. Create. Express your story!</div>",
        englishTab: "English",
        hindiTab: "Hindi",
        exploreActivityBtn: "Activity Template",
        tutorialVideoBtn: "🎥 Tutorial Video",
        tutorialVideoLabel: "Tutorial Video",
        aotmTutorialTitle: "Tutorial",
        aotmTutorialSubtitle: "Watch before you design",
        watchVideoBtn: "▶ Watch",
        submitMonthlyBtn: "Submit Activity Link",
        monthlyTutorialsTitle: "Monthly Activity Tutorials",
        englishTutorial: "English Tutorial",
        hindiTutorial: "Hindi Tutorial",
        englishTutorials: "DCAIS Activities / English",
        hindiTutorials: "DCAIS Activities / Hindi",
        kaushalBodhTab: "Kaushal Bodh Activities",
        formOfWorkLabel: "Form of Work",
        allFormsOfWorkOption: "All Forms of Work",
        workWithLifeForms: "Work with Life Forms",
        workWithMachinesMaterials: "Work with Machines & Materials",
        workInHumanServices: "Work with Human Services",
        keyActivitiesLabel: "Key Activities",
        keyLearningOutcomesLabel: "Key Learning Outcomes",
        adobeExpressIntegrationLabel: "Adobe Express Integration for Activity Books & Portfolios",
        adobeExpressActivityLabel: "Adobe Express Activity",
        finalStudentDeliverableLabel: "Final Student Deliverable",
        openLearningJournal: "Open Learning Journal",
        openTemplateLink: "Open Activity Template",
        learningJournalBtn: "Open Learning Journal Link",
        kaushalHeroSubtitle: "Choose your grade level from Kaushal Bodh curriculum",
        kaushalActivitiesSubtitle: "Showing Kaushal Bodh projects for {grade}",
        importantNoteHeading: "Earn DCAIS Certificate",
        importantNoteText: "All students of Grades 6, 7 and 8, please submit your Creative Assessment by 31st December to earn your Digital Creativity and AI Skills certificate. Note: DCAIS is exclusively for students of Grades 6, 7 and 8.",
        certNotice: "To earn your certificate, please complete at least one activity per month",
        quickSearchTitle: "Search Activity by Month",
        quickSearchDesc: "Filter activities across months in English and Hindi",
        searchActivityByName: "Search Activity by Month (e.g. July, August)...",
        searchActivityByMonth: "Search Activity by Month (e.g. July, August)...",
        ictTab: "ICT-Based Activities",
        dcaisTab: "DCAIS Activities",
        clickHereBtn: "Click Here ↗",
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
        backToWelcome: "Back to Home",
        backToGrades: "Back to Grade Selection",
        backToActivities: "Back to activities",
        descriptionTitle: "Description",
        skillsTitle: "Skills in Focus",
        instructionsTitle: "Instructions",
        submitYourWork: "Submit Your Work",
        submitWorkPrompt: "Completed this activity? Paste your project link to submit your work!",
        submissionDisclaimer: "Disclaimer: The submission of student name and school details is completely voluntary. These fields are exclusively utilized for the issuance of personalized certificates.",
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
            "March": "March",
            "जुलाई": "July",
            "अगस्त": "August",
            "सितम्बर": "September",
            "सितंबर": "September",
            "अक्टूबर": "October",
            "नवंबर": "November",
            "दिसंबर": "December",
            "जनवरी": "January",
            "फरवरी": "February",
            "मार्च": "March",
            "अप्रैल": "April",
            "मई": "May",
            "जून": "June"
        }
    },
    hi: {
        welcomeTitle: "AIM में आपका स्वागत है",
        welcomeDesc: "AIM भारत में छात्रों के लिए एक रचनात्मक शिक्षण स्थल है। हर हफ्ते आपको Adobe Express पर बनी नई गतिविधियाँ मिलती हैं। पोस्टर डिज़ाइन करें, इमेजेस जनरेट करें, वीडियो बनाएँ, और अपनी चीज़ें बनाएँ।",
        activityOfMonthTag: "महीने की गतिविधि",
        activityOfMonthTitle: "बाल दिवस की शुभकामनाएँ! ✨",
        independenceTitle: "बाल दिवस की शुभकामनाएँ! ✨",
        independenceDesc: "<p style=\"margin: 0 0 0.25rem 0; line-height: 1.38;\">Adobe Express के साथ अपने बचपन की यादों को जीवंत बनाएँ! अपने बचपन की किसी जादुई और प्यारी याद के बारे में सोचिए, जिसे याद करके आज भी आपके चेहरे पर मुस्कान आ जाती है। यह परिवार के साथ दिवाली मनाने की याद हो सकती है, दोस्तों के साथ खेलने की या नाना-नानी अथवा दादा-दादी के घर गर्मियों की छुट्टियाँ बिताने की। लिखिए कि उस दिन क्या हुआ था, वह पल आपके लिए खास क्यों था, आपको कैसा महसूस हुआ और आप उसे हमेशा क्यों याद रखना चाहेंगे। इसके बाद ट्यूटोरियल देखिए और सीखिए कि Adobe Express की मदद से अपनी बचपन की याद को एक सुंदर और रचनात्मक चित्र में कैसे बदलें।</p><div style=\"color: #eb1000; font-weight: 700; font-size: 0.82rem;\">कल्पना करें। रचना करें। अपनी कहानी को रचनात्मक रंगों में सजाएँ!</div>",
        englishTab: "English",
        hindiTab: "Hindi",
        exploreActivityBtn: "गतिविधि टेम्पलेट",
        tutorialVideoBtn: "🎥 ट्यूटोरियल वीडियो",
        tutorialVideoLabel: "ट्यूटोरियल वीडियो",
        aotmTutorialTitle: "ट्यूटोरियल",
        aotmTutorialSubtitle: "डिज़ाइन करने से पहले देखें",
        watchVideoBtn: "▶ देखें",
        submitMonthlyBtn: "गतिविधि लिंक जमा करें",
        monthlyTutorialsTitle: "मासिक गतिविधि ट्यूटोरियल",
        englishTutorial: "अंग्रेज़ी ट्यूटोरियल",
        hindiTutorial: "हिंदी ट्यूटोरियल",
        englishTutorials: "DCAIS एक्टिविटीज़ / अंग्रेज़ी",
        hindiTutorials: "DCAIS एक्टिविटीज़ / हिंदी",
        kaushalBodhTab: "कौशल बोध गतिविधियाँ",
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
        learningJournalBtn: "लर्निंग जर्नल लिंक खोलें",
        kaushalHeroSubtitle: "कौशल बोध प्रोजेक्ट्स और लर्निंग जर्नल देखने के लिए अपनी कक्षा चुनें",
        kaushalActivitiesSubtitle: "{grade} के कौशल बोध प्रोजेक्ट्स",
        importantNoteHeading: "DCAIS प्रमाण पत्र",
        importantNoteText: "कक्षा 6, 7 और 8 के सभी छात्र, कृपया अपना डिजिटल क्रिएटिविटी और एआई स्किल्स (DCAIS) प्रमाण पत्र प्राप्त करने के लिए 31 दिसंबर तक अपना रचनात्मक मूल्यांकन (Creative Assessment) जमा करें। (नोट: DCAIS केवल कक्षा 6, 7 और 8 के छात्रों के लिए है)",
        certNotice: "प्रमाण पत्र प्राप्त करने के लिए, कृपया सुनिश्चित करें कि आप प्रति माह एक गतिविधि अवश्य पूरी करें",
        quickSearchTitle: "महीने के अनुसार गतिविधि खोजें",
        quickSearchDesc: "अंग्रेजी और हिंदी दोनों में महीने के अनुसार गतिविधियों को खोजें",
        searchActivityByName: "महीने के अनुसार गतिविधि खोजें (जैसे: जुलाई, अगस्त)...",
        searchActivityByMonth: "महीने के अनुसार गतिविधि खोजें (जैसे: जुलाई, अगस्त / July, August)...",
        ictTab: "आईसीटी-आधारित गतिविधियाँ",
        dcaisTab: "DCAIS गतिविधियाँ",
        clickHereBtn: "यहाँ क्लिक करें ↗",
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
        backToWelcome: "होम पेज पर वापस जाएं",
        backToGrades: "कक्षा चयन पर वापस जाएं",
        backToActivities: "गतिविधियों पर वापस जाएं",
        descriptionTitle: "विवरण",
        skillsTitle: "मुख्य कौशल",
        instructionsTitle: "निर्देश",
        submitYourWork: "अपना काम जमा करें",
        submitWorkPrompt: "क्या आपने यह गतिविधि पूरी कर ली है? अपना प्रोजेक्ट जमा करने के लिए लिंक पेस्ट करें!",
        submissionDisclaimer: "अस्वीकरण: छात्र का नाम और स्कूल विवरण प्रस्तुत करना पूरी तरह से स्वैच्छिक है। इन फ़ील्ड्स का उपयोग विशेष रूप से व्यक्तिगत प्रमाण पत्र जारी करने के लिए किया जाता है।",
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
            "March": "मार्च",
            "जुलाई": "जुलाई",
            "अगस्त": "अगस्त",
            "सितम्बर": "सितंबर",
            "सितंबर": "सितंबर",
            "अक्टूबर": "अक्टूबर",
            "नवंबर": "नवंबर",
            "दिसंबर": "दिसंबर",
            "जनवरी": "जनवरी",
            "फरवरी": "फरवरी",
            "मार्च": "मार्च",
            "अप्रैल": "अप्रैल",
            "मई": "मई",
            "जून": "जून"
        }
    }
};

/**
 * Initializes the application. Loads submissions from LocalStorage,
 * binds events, immediately renders the homepage, and fetches the default curriculum spreadsheet.
 * @returns {Promise<void>}
 */
async function init() {
    // Disable automatic browser scroll restoration on reload
    if (typeof history !== 'undefined' && 'scrollRestoration' in history) {
        history.scrollRestoration = 'manual';
    }
    window.scrollTo(0, 0);

    // Wipe all existing submissions unconditionally for a clean slate
    localStorage.setItem('aim_submissions', JSON.stringify([]));
    state.submissions = [];

    // Instant data hydration from pre-compiled curriculum data for lightning-fast first paint
    hydrateFromPrecompiledData();
    
    detectMode();
    setupEventListeners();
    // Render Homepage (step-welcome) IMMEDIATELY so visitors never see an inner step
    render();
    
    await loadDefaultExcel();
}

/**
 * Hydrates state with precompiled curriculum data for instant, zero-latency rendering and offline capability.
 * @returns {boolean} True if precompiled data was present and loaded.
 */
function hydrateFromPrecompiledData() {
    if (typeof window !== 'undefined' && window.AIM_DEFAULT_DATA) {
        if (Array.isArray(window.AIM_DEFAULT_DATA.activities) && window.AIM_DEFAULT_DATA.activities.length > 0) {
            state.allActivities = window.AIM_DEFAULT_DATA.activities;
        }
        if (Array.isArray(window.AIM_DEFAULT_DATA.kaushalBodh) && window.AIM_DEFAULT_DATA.kaushalBodh.length > 0) {
            state.kaushalBodhActivities = window.AIM_DEFAULT_DATA.kaushalBodh;
        }
        if (state.allActivities.length > 0) {
            hideError();
            hideLoader();
            return true;
        }
    }
    return false;
}

/**
 * Reliably navigates to the AIM Welcome Homepage from anywhere in the application.
 * Cleans selections, resets tutorial mode, updates history, and renders.
 * @returns {void}
 */
function navigateToHome() {
    resetInstructionPopupState();
    state.tutorialMode = 'choose';
    state.selectedGrade = null;
    state.selectedActivity = null;
    state.filterSubject = 'All';
    state.currentStep = 'welcome';
    if (typeof history !== 'undefined' && history.pushState) {
        history.pushState({ step: 'welcome' }, '', window.location.pathname);
    }
    window.scrollTo(0, 0);
    render();
}

/**
 * Pushes state to the browser history stack for backward/forward navigation.
 * @param {string} step - The target step.
 * @returns {void}
 */
function pushNavigationState(step) {
    if (typeof history !== 'undefined' && history.pushState) {
        const payload = {
            step: step,
            tutorialMode: state.tutorialMode,
            curriculumType: state.curriculumType,
            language: state.language,
            selectedGrade: state.selectedGrade,
            filterSubject: state.filterSubject,
            activityName: state.selectedActivity ? state.selectedActivity.activity_name : null
        };
        history.pushState(payload, '', window.location.pathname);
    }
}

/**
 * Detects whether we are in Hindi or English subfolder, hash, or choose welcome path.
 * Guarantees that opening the site always starts on the Homepage.
 * @returns {void}
 */
function detectMode() {
    state.tutorialMode = 'choose';
    state.currentStep = 'welcome';
    state.selectedGrade = null;
    state.selectedActivity = null;
    state.filterSubject = 'All';

    // Clear any leftover hash cleanly from URL without page jump or history pollution
    if (window.location.hash) {
        try {
            history.replaceState({ step: 'welcome' }, '', window.location.pathname);
        } catch (e) {
            window.location.hash = '';
        }
    } else if (typeof history !== 'undefined' && history.replaceState) {
        history.replaceState({ step: 'welcome' }, '', window.location.pathname);
    }

    window.scrollTo(0, 0);
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
            navigateToHome();
        });
    }

    // Dedicated Landing Page Language Toggle Buttons
    const langToggleEn = document.getElementById('lang-toggle-en');
    const langToggleHi = document.getElementById('lang-toggle-hi');

    if (langToggleEn) {
        langToggleEn.addEventListener('click', () => {
            state.userExplicitLanguage = 'en';
            state.language = 'en';
            state.aotmLanguage = 'en';
            render();
        });
    }

    if (langToggleHi) {
        langToggleHi.addEventListener('click', () => {
            state.userExplicitLanguage = 'hi';
            state.language = 'hi';
            state.aotmLanguage = 'hi';
            render();
        });
    }

    // Welcome Choices Click Events (SPA fallback for instant routing)
    const choiceEnglish = document.getElementById('choice-english');
    if (choiceEnglish) {
        choiceEnglish.addEventListener('click', (e) => {
            e.preventDefault();
            resetInstructionPopupState();
            state.tutorialMode = 'english';
            state.language = 'en';
            state.selectedGrade = null;
            state.selectedActivity = null;
            state.currentStep = 'grade';
            pushNavigationState('grade');
            render();
        });
    }

    const choiceHindi = document.getElementById('choice-hindi');
    if (choiceHindi) {
        choiceHindi.addEventListener('click', (e) => {
            e.preventDefault();
            resetInstructionPopupState();
            state.tutorialMode = 'hindi';
            state.language = 'hi';
            state.selectedGrade = null;
            state.selectedActivity = null;
            state.currentStep = 'grade';
            pushNavigationState('grade');
            render();
        });
    }

    const choiceKaushalBodh = document.getElementById('choice-kaushal-bodh');
    if (choiceKaushalBodh) {
        choiceKaushalBodh.addEventListener('click', (e) => {
            e.preventDefault();
            resetInstructionPopupState();
            state.tutorialMode = 'kaushal-bodh';
            state.selectedGrade = null;
            state.selectedActivity = null;
            state.currentStep = 'grade';
            pushNavigationState('grade');
            render();
        });
    }

    // =========================================================================
    // ADOBE EXPRESS THEME INTERACTION HANDLERS (Image 1 & 2)
    // =========================================================================
    let tile1Lang = 'en';
    let tile2Lang = 'en';
    let tile3Lang = 'en';

    // Helper: Open AOTM Modal
    window.openAotmModal = function(lang) {
        const targetLang = lang || tile1Lang || state.language || 'en';
        tile1Lang = targetLang;
        updateAotmModalDisplay(targetLang);
        const modal = document.getElementById('aotm-selection-modal');
        if (modal) modal.classList.remove('hidden');
    };

    // Helper: Close AOTM Modal
    window.closeAotmModal = function() {
        const modal = document.getElementById('aotm-selection-modal');
        if (modal) modal.classList.add('hidden');
    };

    // Helper: Update AOTM Modal Display
    function updateAotmModalDisplay(lang) {
        const targetLang = (lang === 'hi') ? 'hi' : 'en';
        const aotmData = MONTHLY_ACTIVITIES[targetLang] || MONTHLY_ACTIVITIES.en;
        
        const titleEl = document.getElementById('aotm-modal-title');
        if (titleEl) titleEl.textContent = aotmData.title;

        const descEl = document.getElementById('aotm-modal-desc');
        if (descEl) descEl.innerHTML = aotmData.desc;

        const labelEl = document.getElementById('aotm-modal-tutorial-label');
        if (labelEl) labelEl.textContent = targetLang === 'hi' ? '🎬 ट्यूटोरियल वीडियो' : '🎬 Step-by-Step Tutorial';

        const videoLink = document.getElementById('aotm-modal-video-link');
        if (videoLink) videoLink.href = aotmData.videoUrl;

        const videoThumb = document.getElementById('aotm-modal-video-thumb');
        if (videoThumb) {
            videoThumb.src = `https://img.youtube.com/vi/${aotmData.videoId}/hqdefault.jpg`;
            videoThumb.alt = aotmData.videoTitle;
        }

        const templateBtn = document.getElementById('aotm-modal-template-btn');
        if (templateBtn) {
            templateBtn.href = aotmData.templateUrl;
            templateBtn.textContent = aotmData.btnText;
        }

        const submitBtn = document.getElementById('aotm-modal-submit-btn');
        if (submitBtn) {
            submitBtn.textContent = targetLang === 'hi' ? 'लिंक जमा करें' : 'Submit Link';
        }

        const btnEn = document.getElementById('aotm-modal-lang-en');
        const btnHi = document.getElementById('aotm-modal-lang-hi');
        if (btnEn && btnHi) {
            if (targetLang === 'hi') {
                btnEn.classList.remove('active');
                btnHi.classList.add('active');
            } else {
                btnHi.classList.remove('active');
                btnEn.classList.add('active');
            }
        }
    }

    // Helper: Open Certificate Modal
    window.openCertModal = function() {
        const modal = document.getElementById('dcais-cert-modal');
        if (modal) modal.classList.remove('hidden');
    };

    // Helper: Close Certificate Modal
    window.closeCertModal = function() {
        const modal = document.getElementById('dcais-cert-modal');
        if (modal) modal.classList.add('hidden');
    };

    // Modal Close Buttons
    const closeAotmBtn = document.getElementById('close-aotm-modal-btn');
    if (closeAotmBtn) {
        closeAotmBtn.addEventListener('click', () => closeAotmModal());
    }

    const aotmModal = document.getElementById('aotm-selection-modal');
    if (aotmModal) {
        aotmModal.addEventListener('click', (e) => {
            if (e.target === aotmModal) closeAotmModal();
        });
    }

    const modalLangEn = document.getElementById('aotm-modal-lang-en');
    const modalLangHi = document.getElementById('aotm-modal-lang-hi');
    if (modalLangEn) {
        modalLangEn.addEventListener('click', () => {
            tile1Lang = 'en';
            updateAotmModalDisplay('en');
            updateAotmDisplay('en');
        });
    }
    if (modalLangHi) {
        modalLangHi.addEventListener('click', () => {
            tile1Lang = 'hi';
            updateAotmModalDisplay('hi');
            updateAotmDisplay('hi');
        });
    }

    const closeCertBtn = document.getElementById('close-cert-modal-btn');
    if (closeCertBtn) {
        closeCertBtn.addEventListener('click', () => closeCertModal());
    }

    const certModalStartBtn = document.getElementById('cert-modal-start-btn');
    if (certModalStartBtn) {
        certModalStartBtn.addEventListener('click', () => {
            closeCertModal();
            const grid = document.getElementById('ae-tiles-grid');
            if (grid) grid.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    }

    const dcaisCertModal = document.getElementById('dcais-cert-modal');
    if (dcaisCertModal) {
        dcaisCertModal.addEventListener('click', (e) => {
            if (e.target === dcaisCertModal) closeCertModal();
        });
    }

    // Highlighted Rectangular Banner: Assignment Link
    const certLearnMoreBtn = document.getElementById('cert-learn-more-btn');
    if (certLearnMoreBtn) {
        certLearnMoreBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            window.open("https://new.express.adobe.com/webpage/Cgir6LQ7u0AIb/remix", "_blank", "noopener,noreferrer");
        });
    }
    const certBanner = document.getElementById('cert-banner');
    if (certBanner) {
        certBanner.addEventListener('click', (e) => {
            if (e.target.closest('#cert-learn-more-btn')) return;
            window.open("https://new.express.adobe.com/webpage/Cgir6LQ7u0AIb/remix", "_blank", "noopener,noreferrer");
        });
    }

    // Hero Action Buttons
    const heroWatchTutorialBtn = document.getElementById('hero-watch-tutorial-btn');
    if (heroWatchTutorialBtn) {
        heroWatchTutorialBtn.addEventListener('click', () => openAotmModal(tile1Lang));
    }

    const heroGetStartedBtn = document.getElementById('hero-get-started-btn');
    if (heroGetStartedBtn) {
        heroGetStartedBtn.addEventListener('click', () => {
            const grid = document.getElementById('ae-tiles-grid');
            if (grid) grid.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
    }

    // Tile 1: Language Toggle & Click inside
    const tile1LangToggle = document.getElementById('tile-1-lang-toggle');
    if (tile1LangToggle) {
        tile1LangToggle.querySelectorAll('.ae-toggle-lang-opt').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const chosenLang = btn.getAttribute('data-lang') || 'en';
                tile1Lang = chosenLang;
                updateAotmDisplay(chosenLang);
            });
        });
    }

    const tileAotm = document.getElementById('tile-aotm');
    const tile1StartBtn = document.getElementById('tile-1-start-btn');
    if (tileAotm) {
        tileAotm.addEventListener('click', () => openAotmModal(tile1Lang));
    }
    if (tile1StartBtn) {
        tile1StartBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            openAotmModal(tile1Lang);
        });
    }

    // =========================================================================
    // =========================================================================
    // DCAIS ACTIVITIES LANGUAGE SELECTION & ROUTING (Grades 6, 7 & 8)
    // =========================================================================
    function startDCAIS(lang) {
        resetInstructionPopupState();
        state.curriculumType = 'dcais';
        if (lang === 'hi') {
            state.tutorialMode = 'hindi';
            state.language = 'hi';
        } else {
            state.tutorialMode = 'english';
            state.language = 'en';
        }
        state.selectedGrade = null;
        state.selectedActivity = null;
        state.currentStep = 'grade';
        pushNavigationState('grade');
        render();
    }

    window.openDCAISLangModal = function() {
        const modal = document.getElementById('dcais-lang-modal');
        if (modal) modal.classList.remove('hidden');
    };

    window.closeDCAISLangModal = function() {
        const modal = document.getElementById('dcais-lang-modal');
        if (modal) modal.classList.add('hidden');
    };

    const closeDCAISLangBtn = document.getElementById('close-dcais-lang-modal-btn');
    if (closeDCAISLangBtn) {
        closeDCAISLangBtn.addEventListener('click', () => closeDCAISLangModal());
    }

    const dcaisLangModal = document.getElementById('dcais-lang-modal');
    if (dcaisLangModal) {
        dcaisLangModal.addEventListener('click', (e) => {
            if (e.target === dcaisLangModal) closeDCAISLangModal();
        });
    }

    const dcaisSelectEnBtn = document.getElementById('dcais-select-en-btn');
    if (dcaisSelectEnBtn) {
        dcaisSelectEnBtn.addEventListener('click', () => {
            closeDCAISLangModal();
            startDCAIS('en');
        });
    }

    const dcaisSelectHiBtn = document.getElementById('dcais-select-hi-btn');
    if (dcaisSelectHiBtn) {
        dcaisSelectHiBtn.addEventListener('click', () => {
            closeDCAISLangModal();
            startDCAIS('hi');
        });
    }

    const tileDcais = document.getElementById('tile-dcais');
    if (tileDcais) {
        tileDcais.addEventListener('click', () => openDCAISLangModal());
    }

    // =========================================================================
    // ICT-BASED ACTIVITIES LANGUAGE SELECTION & ROUTING (Grades 3, 4 & 5)
    // =========================================================================
    function startICT(lang) {
        resetInstructionPopupState();
        state.curriculumType = 'ict';
        state.tutorialMode = (lang === 'hi') ? 'ict-hindi' : 'ict-english';
        state.language = lang;
        state.selectedGrade = null;
        state.selectedActivity = null;
        state.currentStep = 'grade';
        pushNavigationState('grade');
        render();
    }

    window.openICTLangModal = function() {
        const modal = document.getElementById('ict-lang-modal');
        if (modal) modal.classList.remove('hidden');
    };

    window.closeICTLangModal = function() {
        const modal = document.getElementById('ict-lang-modal');
        if (modal) modal.classList.add('hidden');
    };

    const closeICTLangBtn = document.getElementById('close-ict-lang-modal-btn');
    if (closeICTLangBtn) {
        closeICTLangBtn.addEventListener('click', () => closeICTLangModal());
    }

    const ictLangModal = document.getElementById('ict-lang-modal');
    if (ictLangModal) {
        ictLangModal.addEventListener('click', (e) => {
            if (e.target === ictLangModal) closeICTLangModal();
        });
    }

    const ictSelectEnBtn = document.getElementById('ict-select-en-btn');
    if (ictSelectEnBtn) {
        ictSelectEnBtn.addEventListener('click', () => {
            closeICTLangModal();
            startICT('en');
        });
    }

    const ictSelectHiBtn = document.getElementById('ict-select-hi-btn');
    if (ictSelectHiBtn) {
        ictSelectHiBtn.addEventListener('click', () => {
            closeICTLangModal();
            startICT('hi');
        });
    }

    const tileIct = document.getElementById('tile-ict');
    if (tileIct) {
        tileIct.addEventListener('click', () => openICTLangModal());
    }

    // =========================================================================
    // KAUSHAL BODH ACTIVITIES (Vocational Skills)
    // =========================================================================
    function startTile3Kaushal() {
        resetInstructionPopupState();
        state.curriculumType = 'kaushal-bodh';
        state.tutorialMode = 'kaushal-bodh';
        state.selectedGrade = null;
        state.selectedActivity = null;
        state.currentStep = 'grade';
        pushNavigationState('grade');
        render();
    }

    const tileKaushal = document.getElementById('tile-kaushal');
    if (tileKaushal) {
        tileKaushal.addEventListener('click', () => startTile3Kaushal());
    }

    // =========================================================================
    // HEADER NAVIGATION LINKS
    // =========================================================================
    const navItemHome = document.getElementById('nav-item-home');
    if (navItemHome) {
        navItemHome.addEventListener('click', (e) => {
            e.preventDefault();
            navigateToHome();
        });
    }

    const navItemAotm = document.getElementById('nav-item-aotm');
    if (navItemAotm) {
        navItemAotm.addEventListener('click', (e) => {
            e.preventDefault();
            openAotmModal(tile1Lang);
        });
    }

    const navItemIct = document.getElementById('nav-item-ict');
    if (navItemIct) {
        navItemIct.addEventListener('click', (e) => {
            e.preventDefault();
            openICTLangModal();
        });
    }

    const navItemDcais = document.getElementById('nav-item-dcais');
    if (navItemDcais) {
        navItemDcais.addEventListener('click', (e) => {
            e.preventDefault();
            openDCAISLangModal();
        });
    }

    const navItemKaushal = document.getElementById('nav-item-kaushal');
    if (navItemKaushal) {
        navItemKaushal.addEventListener('click', (e) => {
            e.preventDefault();
            startTile3Kaushal();
        });
    }

    // Back to Welcome Choice Link
    const backToWelcomeLink = document.getElementById('back-to-welcome-link');
    if (backToWelcomeLink) {
        backToWelcomeLink.addEventListener('click', (e) => {
            e.preventDefault();
            navigateToHome();
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
            pushNavigationState(state.currentStep);
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
                pushNavigationState('form-of-work');
                render();
            } else {
                state.selectedActivity = null;
                state.filterSubject = 'All';
                state.currentStep = 'grade';
                pushNavigationState('grade');
                render();
            }
        });
    }

    // Back to Grade Selection from Form of Work step
    const backToGradesFromFow = document.getElementById('back-to-grades-from-fow');
    if (backToGradesFromFow) {
        backToGradesFromFow.addEventListener('click', (e) => {
            e.preventDefault();
            state.filterSubject = '';
            state.selectedGrade = null;
            state.currentStep = 'grade';
            pushNavigationState('grade');
            render();
        });
    }

    // Submissions Google Form redirect for activities
    const showSubmissionFormBtn = document.getElementById('show-submission-form-btn');
    if (showSubmissionFormBtn) {
        showSubmissionFormBtn.addEventListener('click', (e) => {
            e.preventDefault();
            const isKaushalBodh = state.tutorialMode === 'kaushal-bodh' || (state.selectedActivity && state.selectedActivity.isKaushalBodh);
            const formUrl = isKaushalBodh ? 'https://forms.gle/Rdp59kNG5EpRcdedA' : 'https://forms.gle/ZkzcFSFpQQUyxeWC8';
            window.open(formUrl, '_blank', 'noopener,noreferrer');
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

    // Monthly Activity submit button on Welcome Screen
    const submitMonthlyActivityBtn = document.getElementById('submit-monthly-activity-btn');
    if (submitMonthlyActivityBtn) {
        submitMonthlyActivityBtn.addEventListener('click', (e) => {
            e.preventDefault();
            const targetUrl = submitMonthlyActivityBtn.href || 'https://docs.google.com/forms/d/e/1FAIpQLScu7n1TAs4qke7E8vrL-iyzpVzewa9ChQkpMB9wpOyv2cpbAQ/viewform';
            window.open(targetUrl, '_blank', 'noopener,noreferrer');
        });
    }

    // Activity of the Month buttons for English & Hindi Longitudinal Boxes
    const exploreActivityBtnEn = document.getElementById('explore-activity-btn-en');
    if (exploreActivityBtnEn) {
        exploreActivityBtnEn.addEventListener('click', (e) => {
            e.preventDefault();
            const uri = exploreActivityBtnEn.href || MONTHLY_ACTIVITIES.en.templateUrl;
            window.open(uri, "_blank", "noopener,noreferrer");
        });
    }

    const exploreActivityBtnHi = document.getElementById('explore-activity-btn-hi');
    if (exploreActivityBtnHi) {
        exploreActivityBtnHi.addEventListener('click', (e) => {
            e.preventDefault();
            const uri = exploreActivityBtnHi.href || MONTHLY_ACTIVITIES.hi.templateUrl;
            window.open(uri, "_blank", "noopener,noreferrer");
        });
    }

    const aotmTutorialCardLinkEn = document.getElementById('aotm-tutorial-card-link-en');
    if (aotmTutorialCardLinkEn) {
        aotmTutorialCardLinkEn.addEventListener('click', (e) => {
            e.preventDefault();
            const uri = aotmTutorialCardLinkEn.href || MONTHLY_ACTIVITIES.en.videoUrl;
            window.open(uri, "_blank", "noopener,noreferrer");
        });
    }

    const aotmTutorialCardLinkHi = document.getElementById('aotm-tutorial-card-link-hi');
    if (aotmTutorialCardLinkHi) {
        aotmTutorialCardLinkHi.addEventListener('click', (e) => {
            e.preventDefault();
            const uri = aotmTutorialCardLinkHi.href || MONTHLY_ACTIVITIES.hi.videoUrl;
            window.open(uri, "_blank", "noopener,noreferrer");
        });
    }

    // Fallback for previous single button IDs if present
    const exploreActivityBtn = document.getElementById('explore-activity-btn');
    if (exploreActivityBtn) {
        exploreActivityBtn.addEventListener('click', (e) => {
            e.preventDefault();
            const currentLang = state.aotmLanguage || state.language || 'en';
            const defaultUrl = (MONTHLY_ACTIVITIES[currentLang] && MONTHLY_ACTIVITIES[currentLang].templateUrl) || MONTHLY_ACTIVITIES.en.templateUrl;
            const templateUri = exploreActivityBtn.href || defaultUrl;
            window.open(templateUri, "_blank", "noopener,noreferrer");
        });
    }

    const aotmTutorialCardLink = document.getElementById('aotm-tutorial-card-link');
    if (aotmTutorialCardLink) {
        aotmTutorialCardLink.addEventListener('click', (e) => {
            e.preventDefault();
            const currentLang = state.aotmLanguage || state.language || 'en';
            const defaultUrl = (MONTHLY_ACTIVITIES[currentLang] && MONTHLY_ACTIVITIES[currentLang].videoUrl) || MONTHLY_ACTIVITIES.en.videoUrl;
            const targetUri = aotmTutorialCardLink.href || defaultUrl;
            window.open(targetUri, "_blank", "noopener,noreferrer");
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

    // Single Line Month Search Input listeners (Image 2)
    const monthSearchBar = document.getElementById('month-search-bar');
    const clearMonthSearchBtn = document.getElementById('clear-month-search-btn');
    const monthSearchResults = document.getElementById('month-search-results');

    if (monthSearchBar) {
        monthSearchBar.addEventListener('input', (e) => {
            const val = e.target.value;
            if (clearMonthSearchBtn) {
                if (val.trim().length > 0) {
                    clearMonthSearchBtn.classList.remove('hidden');
                } else {
                    clearMonthSearchBtn.classList.add('hidden');
                }
            }
            renderMonthSearchResults(val);
        });

        monthSearchBar.addEventListener('focus', () => {
            if (monthSearchBar.value.trim().length > 0) {
                renderMonthSearchResults(monthSearchBar.value);
            }
        });

        monthSearchBar.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                if (monthSearchResults) monthSearchResults.classList.add('hidden');
            }
        });
    }

    if (clearMonthSearchBtn && monthSearchBar) {
        clearMonthSearchBtn.addEventListener('click', () => {
            monthSearchBar.value = '';
            clearMonthSearchBtn.classList.add('hidden');
            if (monthSearchResults) monthSearchResults.classList.add('hidden');
            renderMonthSearchResults('');
            monthSearchBar.focus();
        });
    }

    // Close month search dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (monthSearchResults && !e.target.closest('#month-search-container')) {
            monthSearchResults.classList.add('hidden');
        }
    });

    // Listen for browser back/forward buttons (HTML5 History API)
    window.addEventListener('popstate', (e) => {
        const pop = e.state;
        if (!pop || pop.step === 'welcome') {
            state.tutorialMode = 'choose';
            state.curriculumType = 'dcais';
            state.selectedGrade = null;
            state.selectedActivity = null;
            state.filterSubject = 'All';
            state.currentStep = 'welcome';
            render();
            return;
        }

        state.tutorialMode = pop.tutorialMode || state.tutorialMode;
        state.curriculumType = pop.curriculumType || (state.tutorialMode === 'kaushal-bodh' ? 'kaushal-bodh' : (state.tutorialMode && state.tutorialMode.startsWith('ict') ? 'ict' : 'dcais'));
        state.language = pop.language || state.language;
        state.selectedGrade = pop.selectedGrade || null;
        state.filterSubject = pop.filterSubject || 'All';

        if (pop.step === 'detail' && pop.activityName) {
            const pool = (state.allActivities || []).concat(state.kaushalBodhActivities || []);
            state.selectedActivity = pool.find(a => a.activity_name === pop.activityName) || null;
        } else {
            state.selectedActivity = null;
        }

        state.currentStep = pop.step;
        render();
    });

    // Handle back/forward cache (bfcache) restore
    window.addEventListener('pageshow', (event) => {
        if (event.persisted) {
            navigateToHome();
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
            name: "Work with Human Services",
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
 * Normalizes raw English or Hindi month strings to a standardized month key.
 * @param {string} mStr - Month name string.
 * @returns {string} Standardized month name (e.g. "July").
 */
function normalizeMonth(mStr) {
    if (!mStr) return "July";
    const s = mStr.trim();
    const lower = s.toLowerCase();
    if (lower.includes("july") || lower.includes("जुलाई")) return "July";
    if (lower.includes("august") || lower.includes("अगस्त")) return "August";
    if (lower.includes("september") || lower.includes("सितंबर") || lower.includes("सितम्बर")) return "September";
    if (lower.includes("october") || lower.includes("अक्टूबर")) return "October";
    if (lower.includes("november") || lower.includes("नवंबर") || lower.includes("नवम्बर")) return "November";
    if (lower.includes("december") || lower.includes("दिसंबर") || lower.includes("दिसम्बर")) return "December";
    if (lower.includes("january") || lower.includes("जनवरी")) return "January";
    if (lower.includes("february") || lower.includes("फरवरी")) return "February";
    if (lower.includes("march") || lower.includes("मार्च")) return "March";
    if (lower.includes("april") || lower.includes("अप्रैल")) return "April";
    if (lower.includes("may") || lower.includes("मई")) return "May";
    if (lower.includes("june") || lower.includes("जून")) return "June";
    return s;
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
 * Fetches the default Excel curriculum sheets from the server or loads precompiled data.
 * Fully offline-resilient, cached, and handles URI-encoding properly.
 * @returns {Promise<void>}
 */
async function loadDefaultExcel() {
    // 1. Try instant hydration first from precompiled bundle
    const hasPrecompiledData = hydrateFromPrecompiledData();

    // Only show loader if we have zero activities loaded yet
    if (!hasPrecompiledData && state.allActivities.length === 0) {
        showLoader();
    }

    // If opened directly as a local file (file:///), fetch is blocked by browser CORS security.
    // If precompiled data is loaded, gracefully complete without error.
    if (typeof window !== 'undefined' && window.location && window.location.protocol === 'file:') {
        if (state.allActivities.length > 0) {
            hideError();
            hideLoader();
            return;
        }
    }

    try {
        // Use properly URI-encoded filenames and stable versioning for browser caching
        const fetchUrl1 = encodeURI('Adobe Express_Implementation plan.xlsx') + '?v=3.0';
        const fetchUrl2 = encodeURI('final of Kaushal Bodh Implementation Table.xlsx') + '?v=3.0';

        const [res1, res2] = await Promise.allSettled([
            fetch(fetchUrl1),
            fetch(fetchUrl2)
        ]);

        let updated = false;

        if (res1.status === 'fulfilled' && res1.value && res1.value.ok) {
            const buf1 = await res1.value.arrayBuffer();
            parseExcel(buf1);
            updated = true;
        } else {
            console.warn('Could not fetch live Adobe Express_Implementation plan.xlsx. Using pre-bundled dataset.');
        }

        if (res2.status === 'fulfilled' && res2.value && res2.value.ok) {
            const buf2 = await res2.value.arrayBuffer();
            parseKaushalBodhExcel(buf2);
            updated = true;
        } else {
            console.warn('Could not fetch live final of Kaushal Bodh Implementation Table.xlsx. Using pre-bundled dataset.');
        }

        hideError();
        hideLoader();

        // If updated dynamically while not on welcome step, re-render smoothly
        if (updated && state.currentStep !== 'welcome') {
            render();
        }
    } catch (error) {
        console.warn('Network sync encountered an error:', error);
        // If we already have activities loaded from precompiled data, never display error screen
        if (state.allActivities && state.allActivities.length > 0) {
            hideError();
            hideLoader();
        } else {
            showError('Could not auto-load the curriculum excel files. Please make sure Adobe Express_Implementation plan.xlsx is in the root folder.');
        }
    }
}

/**
 * Parses Excel workbook contents using SheetJS (XLSX), extracts columns,
 * fills merged values, maps row entries, and builds the state object.
 * Supports the new "Adobe Express_Implementation plan.xlsx" format with:
 *   - 9 sheets: Grade 6/7/8 (Hindi), Grade 3/4/5/6/7/8 (English)
 *   - Grades 3-5: 32 activities (July-February), no Book Screenshot row
 *   - Grades 6-8: 18 activities (July-December), with Book Screenshot row
 *   - Hindi sheets: labels in Hindi (गतिविधि का नाम, मुख्य कौशल, etc.)
 *   - No Video Link or Description rows in any sheet
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
            if (!cleanGradeName || cleanGradeName === "") {
                continue;
            }
            
            const sheet = workbook.Sheets[sheetName];
            
            // Format sheet cells into a 2D grid matrix
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
            
            // Dynamically locate rows based on the labels in Column A (index 0)
            // Supports both English and Hindi labels from the new Excel file
            let activityNameRowIdx = -1;
            let skillsRowIdx = -1;
            let instructionsRowIdx = -1;
            let bookScreenshotRowIdx = -1;
            let templateLinkRowIdx = -1;
            let videoLinkRowIdx = -1;
            let descriptionRowIdx = -1;

            for (let r = 0; r < Math.min(rows.length, 20); r++) {
                const label = rows[r] && rows[r][0] ? rows[r][0].toString().trim().toLowerCase() : "";
                if (!label) continue;

                if (label.includes("activity name") || label.includes("गतिविधि का नाम") || label === "गतिविधि") {
                    activityNameRowIdx = r;
                } else if (label.includes("skill") || label.includes("कौशल")) {
                    skillsRowIdx = r;
                } else if (label.includes("instruction") || label.includes("निर्देश")) {
                    instructionsRowIdx = r;
                } else if (label.includes("screenshot") || label.includes("book") || label.includes("स्क्रीनशॉट") || label.includes("पुस्तक")) {
                    bookScreenshotRowIdx = r;
                } else if (label.includes("activity link") || label.includes("project link") || label.includes("template link") || label.includes("गतिविधि का लिंक")) {
                    templateLinkRowIdx = r;
                } else if (label.includes("video") || label.includes("वीडियो")) {
                    videoLinkRowIdx = r;
                } else if (label.includes("description") || label.includes("विवरण")) {
                    descriptionRowIdx = r;
                }
            }

            // Fallbacks to default offsets if label is not explicitly found
            if (activityNameRowIdx === -1) activityNameRowIdx = 5;
            if (skillsRowIdx === -1) skillsRowIdx = 6;
            if (instructionsRowIdx === -1) instructionsRowIdx = 7;
            // bookScreenshotRowIdx remains -1 if not found (Grades 3-5 don't have it)
            if (templateLinkRowIdx === -1) {
                // For sheets with Book Screenshot, Activity Link is at row 9; otherwise row 8
                templateLinkRowIdx = (bookScreenshotRowIdx !== -1) ? 9 : 8;
            }
            // videoLinkRowIdx and descriptionRowIdx stay -1 if not found (new format doesn't have them)

            const numCols = rows[activityNameRowIdx] ? rows[activityNameRowIdx].length : 0;
            
            // Build month mapping from Row 2 (months row) by filling forward merged cells
            // Row 2 has months at specific columns with gaps where cells are merged
            const monthRow = rows[2] || [];
            const monthMap = {}; // colIndex -> month name
            let lastMonth = "";
            for (let c = 1; c < numCols; c++) {
                const val = monthRow[c] ? monthRow[c].toString().trim() : null;
                if (val && val !== "" &&
                    !val.toUpperCase().startsWith("GRADE") &&
                    !val.toUpperCase().startsWith("कक्षा")) {
                    lastMonth = val;
                }
                if (lastMonth) {
                    monthMap[c] = lastMonth;
                }
            }

            // Build week mapping from Row 4 (weeks row)
            const weekRow = rows[4] || [];

            for (let c = 1; c < numCols; c++) {
                const activityName = rows[activityNameRowIdx] ? rows[activityNameRowIdx][c] : null;
                // Exclude empty and header label cells
                if (!activityName || activityName.toString().trim() === "" || 
                    activityName.toString().trim() === "Activity Name" ||
                    activityName.toString().trim() === "गतिविधि का नाम") {
                    continue;
                }

                // Get month from the pre-built month map
                const currentMonth = monthMap[c] || "July";

                // Get week from Row 4, fallback to calculated week
                let currentWeek = "";
                const weekVal = weekRow[c] ? weekRow[c].toString().trim() : null;
                if (weekVal && weekVal !== "") {
                    currentWeek = weekVal;
                } else {
                    currentWeek = `Week ${((c - 1) % 4) + 1}`;
                }

                const skills = (rows[skillsRowIdx] && rows[skillsRowIdx][c]) ? rows[skillsRowIdx][c].toString().trim() : "";
                const instructions = (rows[instructionsRowIdx] && rows[instructionsRowIdx][c]) ? rows[instructionsRowIdx][c].toString().trim() : "";
                const bookScreenshot = (bookScreenshotRowIdx !== -1 && rows[bookScreenshotRowIdx] && rows[bookScreenshotRowIdx][c]) ? rows[bookScreenshotRowIdx][c].toString().trim() : "";
                const templateLink = (rows[templateLinkRowIdx] && rows[templateLinkRowIdx][c]) ? rows[templateLinkRowIdx][c].toString().trim() : "";
                const videoLink = (videoLinkRowIdx !== -1 && rows[videoLinkRowIdx] && rows[videoLinkRowIdx][c]) ? rows[videoLinkRowIdx][c].toString().trim() : "";
                const description = (descriptionRowIdx !== -1 && rows[descriptionRowIdx] && rows[descriptionRowIdx][c]) ? rows[descriptionRowIdx][c].toString().trim() : "";

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

        if (parsedActivities.length === 0) {
            if (!state.allActivities || state.allActivities.length === 0) {
                showError('Could not read any valid activity rows from spreadsheet.');
            }
            return;
        }

        state.allActivities = parsedActivities;
        hideError();
        hideLoader();
        if (state.currentStep !== 'welcome' && !state.selectedActivity) {
            resetState();
        }
    } catch (err) {
        console.error(err);
        if (!state.allActivities || state.allActivities.length === 0) {
            showError('Error reading the Excel workbook. Please upload a valid sheet.');
        }
    }
}

/**
 * Extracts, de-duplicates, and naturally sorts all unique grade labels from activities list.
 * @param {Activity[]} activities - List of parsed activities.
 * @returns {string[]} Naturally sorted array of grade strings.
 */
function extractGrades(activities) {
    if (state.curriculumType === 'kaushal-bodh' || state.tutorialMode === 'kaushal-bodh') {
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

    if (state.curriculumType === 'ict') {
        // ICT-Based Activities: strictly Grades 3, 4, 5
        const hasExplicitHindi = gradesArray.some(g => /(3|4|5)/.test(g) && g.toLowerCase().includes('(hindi)'));
        if (state.language === 'hi' && hasExplicitHindi) {
            gradesArray = gradesArray.filter(g => /(3|4|5)/.test(g) && g.toLowerCase().includes('(hindi)'));
        } else {
            gradesArray = gradesArray.filter(g => /(3|4|5)/.test(g) && !g.toLowerCase().includes('(hindi)'));
        }
    } else {
        // DCAIS Activities: strictly Grades 6, 7, 8
        if (state.language === 'hi' || state.tutorialMode === 'hindi') {
            gradesArray = gradesArray.filter(g => /(6|7|8)/.test(g) && g.toLowerCase().includes('(hindi)'));
        } else {
            gradesArray = gradesArray.filter(g => /(6|7|8)/.test(g) && !g.toLowerCase().includes('(hindi)'));
        }
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
 * Clears all stored popup dismissal flags from sessionStorage and cancels active timers
 * so the AIM Panda Mascot popup appears fresh like a new user arrival when entering any tab/mode.
 */
function resetInstructionPopupState() {
    if (window.popupAutoCloseTimer) {
        clearTimeout(window.popupAutoCloseTimer);
        window.popupAutoCloseTimer = null;
    }
    try {
        Object.keys(sessionStorage).forEach(key => {
            if (key.includes('_popup_closed_')) {
                sessionStorage.removeItem(key);
            }
        });
    } catch (err) {
        console.warn('Failed clearing popup sessionStorage keys:', err);
    }
}

/**
 * Renders or updates the instruction popup banner for current step page in both DCAIS and Kaushal Bodh modes.
 * Displays page-specific instructions in English or Hindi with the AIM Panda Mascot.
 * @param {string} stepName - 'grade', 'form-of-work', 'activities', or 'detail'
 */
function renderDCAISInstructionPopup(stepName) {
    // Clear existing auto-close timer if active
    if (window.popupAutoCloseTimer) {
        clearTimeout(window.popupAutoCloseTimer);
        window.popupAutoCloseTimer = null;
    }

    // Remove existing popover wrappers and inline banners
    document.querySelectorAll('.dcais-instruction-popover-wrapper, .dcais-inline-instruction-banner').forEach(el => el.remove());

    if (state.currentStep === 'welcome') {
        resetInstructionPopupState();
        return;
    }

    const isKaushalBodh = state.tutorialMode === 'kaushal-bodh';
    const lang = state.language || 'en';
    let messageText = '';

    if (stepName === 'grade') {
        if (isKaushalBodh) {
            messageText = lang === 'hi'
                ? 'अपने कौशल बोध गतिविधियों के लिए एडोब एक्सप्रेस डिजिटल परिसंपत्तियों तक पहुंचें। आप अपने कार्य को बेहतर बनाने के लिए इनका उपयोग कर सकते हैं। शुरू करने के लिए अपनी कक्षा चुनें।'
                : 'Access Adobe Express digital assets for your Kaushal Bodh activities. You can use them to elevate your work. Choose your grade to get started.';
        } else {
            messageText = lang === 'hi'
                ? 'अपनी कक्षा चुनें और अपनी पसंद की गतिविधि का चयन करें! यद्यपि यह योजना केवल सुझाव है, सभी विद्यार्थियों को सलाह दी जाती है कि वे प्रति माह कम से कम एक गतिविधि अवश्य पूरी करें।'
                : 'Pick your grade and select an activity that interests you! While this plan is suggestive, all students are advised to finish a minimum of one activity per month.';
        }
    } else if (stepName === 'form-of-work') {
        if (isKaushalBodh) {
            messageText = lang === 'hi'
                ? 'नीचे प्रोजेक्ट सूची का विस्तार करने के लिए कार्य का प्रकार चुनें। अपनी गतिविधि शुरू करने के लिए किसी प्रोजेक्ट पर क्लिक करें।'
                : 'Select a form of work to expand the project list below. Click on a project to begin your activity.';
        }
    } else if (stepName === 'activities') {
        if (!isKaushalBodh) {
            messageText = lang === 'hi'
                ? "गतिविधि टेम्पलेट देखने के लिए अपनी पसंदीदा गतिविधि के नीचे 'व्यू गतिविधि' (View Activity) बटन पर क्लिक करें!"
                : "Click on the ‘view activity’ button under your desired activity to explore the activity template!";
        }
    } else if (stepName === 'detail') {
        messageText = lang === 'hi'
            ? 'नीचे दिए गए गतिविधि टेम्पलेट लिंक पर क्लिक करने के बाद, आपको गतिविधि पूरी करने के लिए एडोब एक्सप्रेस पेज पर निर्देशित किया जाएगा। एक बार जब आप पूरा कर लें, तो कृपया अपनी गतिविधि का शेयर योग्य लिंक बनाएं और दिए गए गूगल फॉर्म के माध्यम से जमा लिंक बटन पर क्लिक करके जमा करें।'
            : 'After clicking on the activity template link given below, you will be redirected to the Adobe Express page to complete the activity. Once you have finished, please generate a shareable link to your activity and submit it by clicking the submit link button through the Google Form provided.';
    }

    if (!messageText) return;

    const modeKey = isKaushalBodh ? 'kb' : 'dcais';
    const storageKey = `${modeKey}_popup_closed_${stepName}`;
    const isClosed = sessionStorage.getItem(storageKey) === 'true';
    const guideBtnText = lang === 'hi' ? 'दिशा-निर्देश' : 'Instructions';

    const wrapper = document.createElement('div');
    wrapper.className = 'dcais-instruction-popover-wrapper';

    wrapper.innerHTML = `
        <div class="dcais-instruction-popover ${isClosed ? 'hidden' : ''}" id="popover-${stepName}">
            <div class="dcais-instruction-popover-content">
                <button type="button" class="dcais-popover-close-btn" aria-label="Close instructions">&times;</button>
                <div class="popover-mascot-left">
                    <img src="aim_mascot.png" alt="AIM Panda Mascot" class="popover-mascot-img" />
                </div>
                <div class="popover-message-right">
                    <p class="dcais-popover-body-text">${messageText}</p>
                </div>
            </div>
        </div>
        ${isClosed ? `
            <button type="button" class="dcais-popover-trigger-badge" id="trigger-${stepName}" aria-label="Show instructions">
                <div class="trigger-mascot-icon">
                    <img src="aim_mascot.png" alt="AIM Panda Mascot" class="trigger-mascot-img" />
                </div>
                <span class="trigger-text">${guideBtnText}</span>
                <span class="trigger-pulse-dot"></span>
            </button>
        ` : ''}
    `;

    document.body.appendChild(wrapper);

    const closeBtn = wrapper.querySelector('.dcais-popover-close-btn');
    if (closeBtn) {
        closeBtn.addEventListener('click', () => {
            if (window.popupAutoCloseTimer) {
                clearTimeout(window.popupAutoCloseTimer);
                window.popupAutoCloseTimer = null;
            }
            sessionStorage.setItem(storageKey, 'true');
            renderDCAISInstructionPopup(stepName);
        });
    }

    const triggerBtn = wrapper.querySelector(`#trigger-${stepName}`);
    if (triggerBtn) {
        triggerBtn.addEventListener('click', () => {
            if (window.popupAutoCloseTimer) {
                clearTimeout(window.popupAutoCloseTimer);
                window.popupAutoCloseTimer = null;
            }
            sessionStorage.removeItem(storageKey);
            renderDCAISInstructionPopup(stepName);
        });
    }

    // Set 6-second auto-collapse timer if popup is open
    if (!isClosed) {
        window.popupAutoCloseTimer = setTimeout(() => {
            sessionStorage.setItem(storageKey, 'true');
            renderDCAISInstructionPopup(stepName);
        }, 6000);
    }
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

    document.body.setAttribute('data-step', state.currentStep);
    document.body.setAttribute('data-curriculum', state.curriculumType || '');

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
        
        // Toggle Month Search Bar visibility
        const monthSearchSec = document.getElementById('month-search-container');
        if (monthSearchSec) {
            if (state.tutorialMode === 'kaushal-bodh') {
                monthSearchSec.classList.add('hidden');
            } else {
                monthSearchSec.classList.remove('hidden');
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
            ? 'अपनी कक्षा के लिए कार्य का प्रकार चुनें और अपनी गतिविधि के साथ आगे बढ़ने के लिए प्रोजेक्ट का चयन करें। प्रोजेक्ट चुनने के बाद, "ओपन लर्निंग जर्नल" टैब पर क्लिक करें जो आपको आपकी गतिविधि पूरी करने के लिए एडोब एक्सप्रेस पेज पर निर्देशित करेगा।'
            : 'Choose the form of work for your grade and select the Project to continue with your activity. After Selecting the project, click on "Open Learning Journal " tab which will redirect you to Adobe Express Page to complete your activity.';
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

    // Update active navigation links in header
    const navItemHome = document.getElementById('nav-item-home');
    const navItemAotm = document.getElementById('nav-item-aotm');
    const navItemIct = document.getElementById('nav-item-ict');
    const navItemDcais = document.getElementById('nav-item-dcais');
    const navItemKaushal = document.getElementById('nav-item-kaushal');

    [navItemHome, navItemAotm, navItemIct, navItemDcais, navItemKaushal].forEach(el => {
        if (el) el.classList.remove('active');
    });

    if (state.currentStep === 'welcome') {
        if (navItemHome) navItemHome.classList.add('active');
    } else if (state.curriculumType === 'ict') {
        if (navItemIct) navItemIct.classList.add('active');
    } else if (state.curriculumType === 'kaushal-bodh') {
        if (navItemKaushal) navItemKaushal.classList.add('active');
    } else {
        if (navItemDcais) navItemDcais.classList.add('active');
    }

    renderBreadcrumbs();
    renderDCAISInstructionPopup(state.currentStep);

    // Translate all strings in the document DOM
    translateUIInternal();

    // Fire focus updates for screen readers without scrolling the viewport
    if (focusId) {
        const focusEl = document.getElementById(focusId);
        if (focusEl) {
            setTimeout(() => {
                focusEl.focus({ preventScroll: true });
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

    // Level 1: Home (Always navigates back to AIM Welcome Homepage)
    const homeText = (translations[state.language || 'en'] && translations[state.language || 'en'].homeBreadcrumb) || 'Home';
    html += `<li><a href="#" id="breadcrumb-home" title="${escapeHTML(homeText)}">🏠 ${escapeHTML(homeText)}</a></li>`;

    // Level 2: Section / Mode
    let modeLabel = '';
    if (state.curriculumType === 'kaushal-bodh' || state.tutorialMode === 'kaushal-bodh') {
        modeLabel = (translations[state.language || 'en'] && translations[state.language || 'en'].kaushalBodhTab) || 'Kaushal Bodh';
    } else if (state.curriculumType === 'ict' || (state.tutorialMode && state.tutorialMode.startsWith('ict'))) {
        modeLabel = state.language === 'hi' ? 'आईसीटी गतिविधियाँ (कक्षा 3-5)' : 'ICT Activities (Grades 3-5)';
    } else if (state.tutorialMode === 'hindi' || state.language === 'hi') {
        modeLabel = (translations[state.language || 'en'] && translations[state.language || 'en'].hindiTutorials) || 'DCAIS Activities / Hindi';
    } else {
        modeLabel = (translations[state.language || 'en'] && translations[state.language || 'en'].englishTutorials) || 'DCAIS Activities / English';
    }

    if (state.currentStep === 'grade') {
        html += `<li><span class="current">${escapeHTML(modeLabel)}</span></li>`;
    } else {
        html += `<li><a href="#" id="breadcrumb-mode">${escapeHTML(modeLabel)}</a></li>`;
    }

    // Level 3: Grade Selected
    if (state.selectedGrade) {
        const displayGrade = getDisplayGradeName(state.selectedGrade);
        if (state.currentStep === 'form-of-work') {
            html += `<li><span class="current">${escapeHTML(displayGrade)}</span></li>`;
        } else if (state.currentStep === 'activities' || state.currentStep === 'detail') {
            html += `<li><a href="#" id="breadcrumb-grade">${escapeHTML(displayGrade)}</a></li>`;
        }
    }

    // Level 4: Form of Work / Activities
    if (state.tutorialMode === 'kaushal-bodh' && state.filterSubject && state.filterSubject !== 'All') {
        const fowInfo = getFormOfWorkInfo(state.filterSubject);
        const fowName = fowInfo.name;
        if (state.currentStep === 'activities') {
            html += `<li><span class="current">${fowInfo.icon} ${escapeHTML(fowName)}</span></li>`;
        } else if (state.currentStep === 'detail') {
            html += `<li><a href="#" id="breadcrumb-fow">${fowInfo.icon} ${escapeHTML(fowName)}</a></li>`;
        }
    }

    // Level 5: Activity Detail
    if (state.selectedActivity && state.currentStep === 'detail') {
        html += `<li><span class="current">${escapeHTML(state.selectedActivity.activity_name)}</span></li>`;
    }

    breadcrumbsList.innerHTML = html;

    // Click triggers
    const homeBtn = document.getElementById('breadcrumb-home');
    if (homeBtn) {
        homeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            navigateToHome();
        });
    }

    const modeBtn = document.getElementById('breadcrumb-mode');
    if (modeBtn) {
        modeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            resetInstructionPopupState();
            state.selectedGrade = null;
            state.selectedActivity = null;
            state.filterSubject = 'All';
            state.currentStep = 'grade';
            pushNavigationState('grade');
            render();
        });
    }

    const gradeBtn = document.getElementById('breadcrumb-grade');
    if (gradeBtn) {
        gradeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            resetInstructionPopupState();
            state.selectedActivity = null;
            if (state.tutorialMode === 'kaushal-bodh') {
                state.currentStep = 'form-of-work';
            } else {
                state.currentStep = 'activities';
            }
            pushNavigationState(state.currentStep);
            render();
        });
    }

    const fowBtn = document.getElementById('breadcrumb-fow');
    if (fowBtn) {
        fowBtn.addEventListener('click', (e) => {
            e.preventDefault();
            resetInstructionPopupState();
            state.selectedActivity = null;
            if (state.tutorialMode === 'kaushal-bodh') {
                state.currentStep = 'form-of-work';
            } else {
                state.currentStep = 'activities';
            }
            pushNavigationState(state.currentStep);
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
    
    // Update subtitle dynamically for ICT vs DCAIS vs Kaushal Bodh
    const heroSubtitleEl = document.getElementById('hero-subtitle');
    if (heroSubtitleEl) {
        const dict = translations[state.language || 'en'];
        if (isKB) {
            heroSubtitleEl.innerText = dict.kaushalHeroSubtitle || dict.heroSubtitle;
        } else if (state.curriculumType === 'ict') {
            heroSubtitleEl.innerText = state.language === 'hi' 
                ? "कक्षा 3, 4 और 5 की गतिविधियाँ देखने के लिए अपनी कक्षा चुनें"
                : "Choose your grade level for ICT-Based Activities (Grades 3, 4 & 5)";
        } else {
            heroSubtitleEl.innerText = state.language === 'hi'
                ? "कक्षा 6, 7 और 8 की DCAIS गतिविधियाँ देखने के लिए अपनी कक्षा चुनें"
                : "Choose your grade level for DCAIS Activities (Grades 6, 7 & 8)";
        }
    }

    // Refresh month search placeholder
    const monthSearchBar = document.getElementById('month-search-bar');
    if (monthSearchBar && !monthSearchBar.value) {
        monthSearchBar.placeholder = (translations[state.language || 'en'] && translations[state.language || 'en'].searchActivityByMonth) 
            || "Search Activity by Month (e.g. July, August / जुलाई, अगस्त)...";
    }

    // Render Month Search results if input has active text
    if (monthSearchBar && monthSearchBar.value.trim().length > 0) {
        renderMonthSearchResults(monthSearchBar.value);
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
        const count = activitiesList.filter(a => a.grade === grade).length;
        const template = count === 1 ? translations[state.language || 'en'].activitySingle : translations[state.language || 'en'].activitiesPlural;
        const labelText = template.replace('{count}', count);

        const numStr = grade.replace(/\D/g, '');
        let icon = '🎓';
        let accentClass = 'indigo';
        if (numStr === '3') {
            accentClass = 'teal';
            icon = '🎨';
        } else if (numStr === '4') {
            accentClass = 'cyan';
            icon = '💻';
        } else if (numStr === '5') {
            accentClass = 'blue';
            icon = '🚀';
        } else if (numStr === '6') {
            accentClass = 'pink';
            icon = '💡';
        } else if (numStr === '7') {
            accentClass = 'green';
            icon = '📐';
        } else if (numStr === '8') {
            accentClass = 'indigo';
            icon = '🤖';
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
            pushNavigationState(state.currentStep);
            render();
        });
    });
}

/**
 * Renders month-based search results with bi-directional English & Hindi matching,
 * strictly scoped to the student's active language medium and curriculum.
 * - If student is in English: shows ONLY English activities (no Hindi activities).
 * - If student is in Hindi: allows searching in English ("july") or Hindi ("जुलाई"), and shows ONLY Hindi activities.
 * - Works strictly for both ICT (Grades 3, 4, 5) and DCAIS (Grades 6, 7, 8).
 * @param {string} query
 */
function renderMonthSearchResults(query = '') {
    const resultsContainer = document.getElementById('month-search-results');
    if (!resultsContainer) return;

    const cleanQuery = (query || '').trim().toLowerCase();
    const isHindiMedium = (state.language === 'hi' || state.tutorialMode === 'hindi' || state.tutorialMode === 'ict-hindi');
    const lang = isHindiMedium ? 'hi' : 'en';
    const dict = translations[lang] || translations.en;

    if (!cleanQuery) {
        resultsContainer.innerHTML = '';
        resultsContainer.classList.add('hidden');
        return;
    }

    // 1. Identify target months from query bi-directionally
    // E.g. 'july', 'jul', 'जुलाई' all resolve to month index 7 (July / जुलाई)
    let matchedMonthObjects = [];
    for (const [key, val] of Object.entries(MONTH_TRANSLATION_MAP)) {
        if (key.toLowerCase().startsWith(cleanQuery) || key.toLowerCase().includes(cleanQuery)) {
            if (!matchedMonthObjects.some(m => m.index === val.index)) {
                matchedMonthObjects.push(val);
            }
        }
    }

    // 2. Candidate activity pool strictly based on active curriculumType and student language medium
    let pool = [];
    if (state.curriculumType === 'ict') {
        // ICT-Based Activities: strictly Grades 3, 4, 5
        const hasExplicitIctHindi = (state.allActivities || []).some(a => /(3|4|5)/.test(a.grade) && a.grade.toLowerCase().includes('(hindi)'));
        if (isHindiMedium && hasExplicitIctHindi) {
            pool = (state.allActivities || []).filter(a => /(3|4|5)/.test(a.grade) && a.grade.toLowerCase().includes('(hindi)'));
        } else {
            pool = (state.allActivities || []).filter(a => /(3|4|5)/.test(a.grade) && !a.grade.toLowerCase().includes('(hindi)'));
        }
    } else if (state.curriculumType === 'kaushal-bodh') {
        pool = state.kaushalBodhActivities || [];
    } else {
        // DCAIS Activities: strictly Grades 6, 7, 8
        if (isHindiMedium) {
            // Hindi medium: ONLY show Hindi curriculum activities (Grade 6 (Hindi), etc.)
            pool = (state.allActivities || []).filter(a => /(6|7|8)/.test(a.grade) && a.grade.toLowerCase().includes('(hindi)'));
        } else {
            // English medium: ONLY show English curriculum activities (Grade 6, 7, 8)
            pool = (state.allActivities || []).filter(a => /(6|7|8)/.test(a.grade) && !a.grade.toLowerCase().includes('(hindi)'));
        }
    }

    // Fallback if specific pool is empty
    if (pool.length === 0) {
        if (isHindiMedium) {
            pool = (state.allActivities || []).filter(a => a.grade.toLowerCase().includes('(hindi)'));
        } else {
            pool = (state.allActivities || []).filter(a => !a.grade.toLowerCase().includes('(hindi)'));
        }
    }

    // 3. Filter pool for activities that match the resolved month or name
    const matches = pool.filter(act => {
        const actMonth = (act.month || '').trim().toLowerCase();
        const actName = (act.activity_name || '').trim().toLowerCase();

        // Check if activity month matches any resolved target month
        if (matchedMonthObjects.length > 0) {
            for (const m of matchedMonthObjects) {
                const enMonth = m.en.toLowerCase();
                const hiMonth = m.hi.toLowerCase();
                const alts = (m.altHi || []).map(a => a.toLowerCase());
                if (actMonth === enMonth || actMonth === hiMonth || alts.includes(actMonth) ||
                    actMonth.includes(enMonth) || actMonth.includes(hiMonth)) {
                    return true;
                }
            }
        }

        // Direct substring match on month or activity name
        if (actMonth && actMonth.includes(cleanQuery)) return true;
        if (actName && actName.includes(cleanQuery)) return true;

        return false;
    });

    resultsContainer.classList.remove('hidden');

    if (matches.length === 0) {
        resultsContainer.innerHTML = `
            <div style="text-align: center; color: var(--color-text-secondary); padding: 1.25rem 1rem; font-size: 0.88rem;">
                <span style="font-size: 1.5rem; display: block; margin-bottom: 0.35rem;">🔍</span>
                <span>${isHindiMedium 
                    ? `"${escapeHTML(cleanQuery)}" के लिए कोई गतिविधि नहीं मिली। कृपया जुलाई, अगस्त, सितंबर जैसे महीने खोजें।` 
                    : `No activities found for "${escapeHTML(cleanQuery)}". Try searching for July, August, September...`}</span>
            </div>
        `;
        return;
    }

    let html = '';
    // Show up to 25 relevant matches strictly matching the student's language
    matches.slice(0, 25).forEach(act => {
        const isKB = act.isKaushalBodh;
        const rawMonth = act.month || '';
        // Localize month label to student language (e.g. English -> July; Hindi -> जुलाई)
        const monthLabel = dict.monthNames[rawMonth] || rawMonth;
        const gradeDisplay = getDisplayGradeName(act.grade);

        html += `
            <button type="button" class="month-search-item" 
                data-activity-name="${encodeURIComponent(act.activity_name)}" 
                data-activity-grade="${encodeURIComponent(act.grade)}" 
                data-is-kb="${isKB ? 'true' : 'false'}"
                aria-label="${escapeHTML(act.activity_name)}, ${escapeHTML(gradeDisplay)}, ${escapeHTML(monthLabel)}">
                <div style="display: flex; flex-direction: column; gap: 0.2rem; flex: 1; text-align: left;">
                    <span class="month-search-item-title">${escapeHTML(act.activity_name)}</span>
                    <span style="font-size: 0.82rem; color: var(--color-text-secondary); font-weight: 500; line-height: 1.35;">
                        ${act.skills ? escapeHTML(act.skills.substring(0, 70)) + '...' : ''}
                    </span>
                </div>
                <div class="month-search-item-meta">
                    <span class="month-search-item-badge">📅 ${escapeHTML(monthLabel)}</span>
                    <span class="grade-badge" style="font-size: 0.72rem; padding: 0.15rem 0.55rem; border-radius: 9999px; background: #f1f5f9; color: #334155; font-weight: 700; white-space: nowrap;">
                        ${escapeHTML(gradeDisplay)}
                    </span>
                </div>
            </button>
        `;
    });

    resultsContainer.innerHTML = html;

    // Bind click events on search results
    resultsContainer.querySelectorAll('.month-search-item').forEach(item => {
        item.addEventListener('click', () => {
            const actName = decodeURIComponent(item.getAttribute('data-activity-name'));
            const actGrade = decodeURIComponent(item.getAttribute('data-activity-grade'));
            const isKB = item.getAttribute('data-is-kb') === 'true';

            const poolToSearch = isKB ? state.kaushalBodhActivities : state.allActivities;
            const selected = poolToSearch.find(a => a.grade === actGrade && a.activity_name === actName);

            if (selected) {
                if (isKB) {
                    state.curriculumType = 'kaushal-bodh';
                    state.tutorialMode = 'kaushal-bodh';
                } else if (/(3|4|5)/.test(actGrade)) {
                    state.curriculumType = 'ict';
                    if (state.language === 'hi') {
                        state.tutorialMode = 'ict-hindi';
                    } else {
                        state.tutorialMode = 'ict-english';
                    }
                } else {
                    state.curriculumType = 'dcais';
                    if (actGrade.toLowerCase().includes('(hindi)') || state.language === 'hi') {
                        state.tutorialMode = 'hindi';
                        state.language = 'hi';
                    } else {
                        state.tutorialMode = 'english';
                        state.language = 'en';
                    }
                }

                state.selectedGrade = actGrade;
                state.selectedActivity = selected;
                state.currentStep = 'detail';
                resultsContainer.classList.add('hidden');
                pushNavigationState('detail');
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
            key: 'Work with Human Services',
            category: 'human',
            title: lang === 'hi' ? 'मानव सेवाओं में कार्य' : 'Work with Human Services',
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
            key: 'Work with Human Services',
            category: 'human',
            title: lang === 'hi' ? 'मानव सेवाओं में कार्य' : 'Work with Human Services',
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
                        pushNavigationState('detail');
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
                        pushNavigationState('detail');
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
                        pushNavigationState('detail');
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

    // Update submission form button URL dynamically
    const showSubmissionFormBtn = document.getElementById('show-submission-form-btn');
    if (showSubmissionFormBtn) {
        const isKaushalBodh = activity.isKaushalBodh || state.tutorialMode === 'kaushal-bodh';
        showSubmissionFormBtn.href = isKaushalBodh ? 'https://forms.gle/Rdp59kNG5EpRcdedA' : 'https://forms.gle/ZkzcFSFpQQUyxeWC8';
    }

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
            descElem.style.display = 'block';
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

    // Description text block (completely hide if no description)
    const descElem = document.getElementById('detail-description');
    if (descElem) {
        const descText = (activity.description || '').trim();
        if (descText) {
            descElem.innerText = descText;
            descElem.style.display = 'block';
        } else {
            descElem.innerText = '';
            descElem.style.display = 'none';
        }
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
    document.getElementById('error-view').classList.add('hidden');
    if (state.currentStep !== 'welcome') {
        document.getElementById('loader-view').classList.remove('hidden');
        document.getElementById('step-grade').classList.add('hidden');
        document.getElementById('step-activities').classList.add('hidden');
        document.getElementById('step-detail').classList.add('hidden');
    } else {
        document.getElementById('loader-view').classList.add('hidden');
    }
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

    // Set document root language attribute for CSS :lang() and typography hooks
    document.documentElement.lang = lang;

    // Synchronize language toggle UI buttons
    updateLangToggleUI();

    // Translate all standard data-i18n elements
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        if (dict && dict[key]) {
            if (el.getAttribute('data-i18n-html') === 'true') {
                el.innerHTML = dict[key];
            } else {
                el.innerText = dict[key];
            }
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

    // Synchronize month search bar placeholder
    const monthSearchBar = document.getElementById('month-search-bar');
    if (monthSearchBar) {
        monthSearchBar.placeholder = dict.searchActivityByMonth || "Search Activity by Month (e.g. July, August / जुलाई, अगस्त)...";
    }

    // Synchronize Tile titles and descriptions
    const tileIctTitle = document.getElementById('tile-ict-title');
    if (tileIctTitle) {
        tileIctTitle.textContent = dict.ictTab || "ICT-Based Activities";
    }
    const tileIctDesc = document.getElementById('tile-ict-desc');
    if (tileIctDesc) {
        tileIctDesc.textContent = (lang === 'hi')
            ? "कक्षा 3, 4 और 5 के लिए डिजिटल रचनात्मकता, कोडिंग ब्लॉक्स, विज़ुअल आर्ट टेम्पलेट्स और ग्राफिक डिज़ाइन गतिविधियाँ।"
            : "Digital creativity, coding blocks, visual art templates and graphic design activities for Grades 3, 4 and 5.";
    }
    const tileDcaisTitle = document.getElementById('tile-2-title');
    if (tileDcaisTitle) {
        tileDcaisTitle.textContent = dict.dcaisTab || "DCAIS Activities";
    }
    const tileDcaisDesc = document.getElementById('tile-2-desc');
    if (tileDcaisDesc) {
        tileDcaisDesc.textContent = (lang === 'hi')
            ? "रचनात्मक पोस्टर टेम्पलेट्स, ग्राफिक कला और दृश्य लेआउट के साथ डिजिटल क्रिएटिविटी और एआई स्किल्स पाठ्यक्रम।"
            : "Digital Creativity & AI Skills curriculum with creative poster templates, graphic art, and visual layouts.";
    }

    // Synchronize Certificate button text
    const certLearnMoreBtn = document.getElementById('cert-learn-more-btn');
    if (certLearnMoreBtn) {
        certLearnMoreBtn.innerHTML = `<span>${dict.clickHereBtn || 'Click Here ↗'}</span>`;
    }

    // Synchronize Activity of the Month card (English vs Hindi)
    const aotmLang = state.aotmLanguage || state.language || 'en';
    updateAotmDisplay(aotmLang);
}

/**
 * Updates the Activity of the Month card content and template link based on selected language ('en' | 'hi').
 * @param {'en'|'hi'} lang - Language code.
 */
function updateAotmDisplay(lang) {
    const targetLang = (lang === 'hi') ? 'hi' : 'en';
    state.aotmLanguage = targetLang;
    const aotmData = MONTHLY_ACTIVITIES[targetLang] || MONTHLY_ACTIVITIES.en;

    // Tile 1 elements in new Adobe Express layout
    const tile1Title = document.getElementById('tile-1-title');
    if (tile1Title) {
        tile1Title.textContent = aotmData.title;
    }

    const tile1Desc = document.getElementById('tile-1-desc');
    if (tile1Desc) {
        tile1Desc.innerHTML = targetLang === 'hi'
            ? 'अपने बचपन की किसी जादुई याद के बारे में सोचिए। लिखिए कि वह पल खास क्यों था और ट्यूटोरियल देखकर इसे Adobe Express से सुंदर चित्र में बदलें! कल्पना करें। रचना करें।'
            : 'Think of a magical childhood memory that still makes you smile. Write about what made it special, then transform it into an Adobe Express visual! Imagine. Create. Express your story!';
    }

    const tile1LangToggle = document.getElementById('tile-1-lang-toggle');
    if (tile1LangToggle) {
        tile1LangToggle.querySelectorAll('.ae-toggle-lang-opt').forEach(btn => {
            if (btn.getAttribute('data-lang') === targetLang) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });
    }

    // Modal elements
    const modalTitle = document.getElementById('aotm-modal-title');
    if (modalTitle) modalTitle.textContent = aotmData.title;

    const modalDesc = document.getElementById('aotm-modal-desc');
    if (modalDesc) modalDesc.innerHTML = aotmData.desc;

    const modalVideoThumb = document.getElementById('aotm-modal-video-thumb');
    if (modalVideoThumb) {
        modalVideoThumb.src = `https://img.youtube.com/vi/${aotmData.videoId}/hqdefault.jpg`;
        modalVideoThumb.alt = aotmData.videoTitle;
    }

    const modalVideoLink = document.getElementById('aotm-modal-video-link');
    if (modalVideoLink) modalVideoLink.href = aotmData.videoUrl;

    const modalTemplateBtn = document.getElementById('aotm-modal-template-btn');
    if (modalTemplateBtn) {
        modalTemplateBtn.href = aotmData.templateUrl;
        modalTemplateBtn.textContent = aotmData.btnText;
    }

    const modalSubmitBtn = document.getElementById('aotm-modal-submit-btn');
    if (modalSubmitBtn) {
        modalSubmitBtn.textContent = targetLang === 'hi' ? 'लिंक जमा करें' : 'Submit Link';
    }

    const modalBtnEn = document.getElementById('aotm-modal-lang-en');
    const modalBtnHi = document.getElementById('aotm-modal-lang-hi');
    if (modalBtnEn && modalBtnHi) {
        if (targetLang === 'hi') {
            modalBtnEn.classList.remove('active');
            modalBtnHi.classList.add('active');
        } else {
            modalBtnHi.classList.remove('active');
            modalBtnEn.classList.add('active');
        }
    }

    // Legacy fallbacks
    const titleEl = document.getElementById('aotm-title');
    if (titleEl) titleEl.textContent = aotmData.title;

    const descEl = document.getElementById('aotm-desc');
    if (descEl) descEl.innerHTML = aotmData.desc;

    const templateBtn = document.getElementById('explore-activity-btn');
    if (templateBtn) {
        templateBtn.href = aotmData.templateUrl;
        templateBtn.textContent = aotmData.btnText;
    }
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
