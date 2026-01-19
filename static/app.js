let currentData = null;
let allEmployees = [];
const employeeById = new Map();
let root = null;
let svg = null;
let g = null;
let linkLayer = null;
let nodeLayer = null;
let zoom = null;
let appSettings = {};
let currentLayout = 'vertical'; // Default layout
const hiddenNodeIds = new Set(JSON.parse(localStorage.getItem('hiddenNodeIds') || '[]'));
let isAuthenticated = false;
const COMPACT_PREFERENCE_KEY = 'orgChart.compactLargeTeams';
let userCompactPreference = null;
const PROFILE_IMAGE_PREFERENCE_KEY = 'orgChart.showProfileImages';
let userProfileImagesPreference = null;
let serverShowProfileImages = null;
const SHOW_EMPLOYEE_COUNT_PREFERENCE_KEY = 'orgChart.showEmployeeCount';
let userShowEmployeeCountPreference = null;
let serverShowEmployeeCount = null;
const SHOW_DEPARTMENTS_PREFERENCE_KEY = 'orgChart.showDepartments';
let userShowDepartmentsPreference = null;
let serverShowDepartments = null;
const SHOW_JOB_TITLES_PREFERENCE_KEY = 'orgChart.showJobTitles';
let userShowJobTitlesPreference = null;
let serverShowJobTitles = null;
const SHOW_OFFICE_PREFERENCE_KEY = 'orgChart.showOffice';
let userShowOfficePreference = null;
let serverShowOffice = null;
const SHOW_NAMES_PREFERENCE_KEY = 'orgChart.showNames';
let userShowNamesPreference = null;
let serverShowNames = null;
const HIGHLIGHT_NEW_EMPLOYEES_SESSION_KEY = 'orgChart.highlightNewEmployees';
let sessionHighlightNewEmployeesPreference = null;
let currentDetailEmployeeId = null;
const TITLE_OVERRIDE_SESSION_KEY = 'orgChart.sessionTitleOverrides';
const DEPARTMENT_OVERRIDE_SESSION_KEY = 'orgChart.sessionDepartmentOverrides';
const titleOverrides = loadTitleOverrides();
const departmentOverrides = loadDepartmentOverrides();
let currentTitleEditEmployeeId = null;
let lastFocusBeforeTitleModal = null;
let currentOverrideFocusField = 'title';

function loadTitleOverrides() {
    try {
        if (!window.sessionStorage) {
            return new Map();
        }
        const raw = window.sessionStorage.getItem(TITLE_OVERRIDE_SESSION_KEY);
        if (!raw) {
            return new Map();
        }
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== 'object') {
            return new Map();
        }
        const entries = Object.entries(parsed).filter(entry => typeof entry[0] === 'string' && typeof entry[1] === 'string');
        return new Map(entries);
    } catch (error) {
        console.warn('Unable to load title overrides from session storage:', error);
        return new Map();
    }
}

function persistTitleOverrides() {
    try {
        if (!window.sessionStorage) {
            return;
        }
        if (titleOverrides.size === 0) {
            window.sessionStorage.removeItem(TITLE_OVERRIDE_SESSION_KEY);
        } else {
            const payload = Object.fromEntries(titleOverrides.entries());
            window.sessionStorage.setItem(TITLE_OVERRIDE_SESSION_KEY, JSON.stringify(payload));
        }
    } catch (error) {
        console.warn('Unable to persist title overrides to session storage:', error);
    }
}

function loadDepartmentOverrides() {
    try {
        if (!window.sessionStorage) {
            return new Map();
        }
        const raw = window.sessionStorage.getItem(DEPARTMENT_OVERRIDE_SESSION_KEY);
        if (!raw) {
            return new Map();
        }
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== 'object') {
            return new Map();
        }
        const entries = Object.entries(parsed).filter(entry => typeof entry[0] === 'string' && typeof entry[1] === 'string');
        return new Map(entries);
    } catch (error) {
        console.warn('Unable to load department overrides from session storage:', error);
        return new Map();
    }
}

function persistDepartmentOverrides() {
    try {
        if (!window.sessionStorage) {
            return;
        }
        if (departmentOverrides.size === 0) {
            window.sessionStorage.removeItem(DEPARTMENT_OVERRIDE_SESSION_KEY);
        } else {
            const payload = Object.fromEntries(departmentOverrides.entries());
            window.sessionStorage.setItem(DEPARTMENT_OVERRIDE_SESSION_KEY, JSON.stringify(payload));
        }
    } catch (error) {
        console.warn('Unable to persist department overrides to session storage:', error);
    }
}

function getTitleOverride(employeeId) {
    if (!employeeId) return undefined;
    return titleOverrides.get(employeeId);
}

function isTitleOverridden(employeeId) {
    return employeeId ? titleOverrides.has(employeeId) : false;
}

function getDepartmentOverride(employeeId) {
    if (!employeeId) return undefined;
    return departmentOverrides.get(employeeId);
}

function isDepartmentOverridden(employeeId) {
    return employeeId ? departmentOverrides.has(employeeId) : false;
}

function getEmployeeOriginalField(employeeId, fieldName) {
    if (!employeeId || !fieldName) {
        return '';
    }
    const record = employeeById.get(employeeId);
    if (!record) {
        return '';
    }
    const value = record[fieldName];
    return typeof value === 'string' ? value.trim() : '';
}

function normalizeOverrideValue(value) {
    return typeof value === 'string' ? value.trim() : '';
}

function setResetButtonVisibility(button, available) {
    if (!button) return;
    if (available) {
        button.classList.remove('is-hidden');
        button.removeAttribute('aria-hidden');
        button.removeAttribute('tabindex');
        button.disabled = false;
    } else {
        button.classList.add('is-hidden');
        button.setAttribute('aria-hidden', 'true');
        button.setAttribute('tabindex', '-1');
        button.disabled = true;
    }
}

function setTitleOverride(employeeId, value) {
    if (!employeeId) return;
    const trimmed = normalizeOverrideValue(value);
    const originalValue = getEmployeeOriginalField(employeeId, 'title');
    if (trimmed && trimmed === originalValue) {
        titleOverrides.delete(employeeId);
    } else if (trimmed) {
        titleOverrides.set(employeeId, trimmed);
    } else {
        titleOverrides.delete(employeeId);
    }
    persistTitleOverrides();
    updateOverrideResetButtons();
}

function clearAllTitleOverrides() {
    if (titleOverrides.size === 0) {
        return false;
    }
    titleOverrides.clear();
    persistTitleOverrides();
    updateOverrideResetButtons();
    return true;
}

function pruneTitleOverrides(validIds = []) {
    if (!Array.isArray(validIds) || validIds.length === 0 || titleOverrides.size === 0) {
        return;
    }
    const validSet = new Set(validIds);
    let changed = false;
    titleOverrides.forEach((storedValue, key) => {
        const normalizedValue = normalizeOverrideValue(storedValue);
        const originalValue = getEmployeeOriginalField(key, 'title');
        if (!validSet.has(key) || normalizedValue === originalValue) {
            titleOverrides.delete(key);
            changed = true;
        }
    });
    if (changed) {
        persistTitleOverrides();
        updateOverrideResetButtons();
    }
}

function updateTitleResetButtonState() {
    const resetButton = document.querySelector('[data-control="reset-titles"]');
    setResetButtonVisibility(resetButton, titleOverrides.size > 0);
}

function setDepartmentOverride(employeeId, value) {
    if (!employeeId) return;
    const trimmed = normalizeOverrideValue(value);
    const originalValue = getEmployeeOriginalField(employeeId, 'department');
    if (trimmed && trimmed === originalValue) {
        departmentOverrides.delete(employeeId);
    } else if (trimmed) {
        departmentOverrides.set(employeeId, trimmed);
    } else {
        departmentOverrides.delete(employeeId);
    }
    persistDepartmentOverrides();
    updateOverrideResetButtons();
}

function clearAllDepartmentOverrides() {
    if (departmentOverrides.size === 0) {
        return false;
    }
    departmentOverrides.clear();
    persistDepartmentOverrides();
    updateOverrideResetButtons();
    return true;
}

function pruneDepartmentOverrides(validIds = []) {
    if (!Array.isArray(validIds) || validIds.length === 0 || departmentOverrides.size === 0) {
        return;
    }
    const validSet = new Set(validIds);
    let changed = false;
    departmentOverrides.forEach((storedValue, key) => {
        const normalizedValue = normalizeOverrideValue(storedValue);
        const originalValue = getEmployeeOriginalField(key, 'department');
        if (!validSet.has(key) || normalizedValue === originalValue) {
            departmentOverrides.delete(key);
            changed = true;
        }
    });
    if (changed) {
        persistDepartmentOverrides();
        updateOverrideResetButtons();
    }
}

function updateDepartmentResetButtonState() {
    const resetButton = document.querySelector('[data-control="reset-departments"]');
    setResetButtonVisibility(resetButton, departmentOverrides.size > 0);
}

function updateHiddenResetButtonState() {
    const resetButton = document.querySelector('[data-control="reset-hidden"]');
    setResetButtonVisibility(resetButton, hiddenNodeIds.size > 0);
}

function updateOverrideResetButtons() {
    updateTitleResetButtonState();
    updateDepartmentResetButtonState();
    updateHiddenResetButtonState();
}

function refreshAfterOverrideChange() {
    if (root) {
        update(root);
    }
    refreshSearchResultsPresentation();
    refreshEmployeeDetailPanel();
}

async function waitForTranslations() {
    if (window.i18n && window.i18n.ready && typeof window.i18n.ready.then === 'function') {
        try {
            await window.i18n.ready;
        } catch (error) {
            console.error('[i18n] Failed to load translations', error);
        }
    }
}

function t(key, params) {
    if (window.i18n && typeof window.i18n.t === 'function') {
        return window.i18n.t(key, params);
    }
    return key;
}

function loadStoredCompactPreference() {
    userCompactPreference = null;
    try {
        const stored = localStorage.getItem(COMPACT_PREFERENCE_KEY);
        if (stored === 'true') {
            userCompactPreference = true;
        } else if (stored === 'false') {
            userCompactPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access compact layout preference storage:', error);
        userCompactPreference = null;
    }
}

function storeCompactPreference(value) {
    userCompactPreference = value;
    try {
        localStorage.setItem(COMPACT_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist compact layout preference:', error);
    }
}

function clearCompactPreferenceStorage() {
    userCompactPreference = null;
    try {
        localStorage.removeItem(COMPACT_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear compact layout preference storage:', error);
    }
}

function getEffectiveCompactEnabled() {
    const serverEnabled = !appSettings || appSettings.multiLineChildrenEnabled !== false;
    if (!isAuthenticated && userCompactPreference !== null) {
        return userCompactPreference;
    }
    return serverEnabled;
}

function loadStoredProfileImagePreference() {
    userProfileImagesPreference = null;
    try {
        const stored = localStorage.getItem(PROFILE_IMAGE_PREFERENCE_KEY);
        if (stored === 'true') {
            userProfileImagesPreference = true;
        } else if (stored === 'false') {
            userProfileImagesPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access profile image preference storage:', error);
        userProfileImagesPreference = null;
    }
}

function storeProfileImagePreference(value) {
    userProfileImagesPreference = value;
    try {
        localStorage.setItem(PROFILE_IMAGE_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist profile image preference:', error);
    }
}

function clearProfileImagePreference() {
    userProfileImagesPreference = null;
    try {
        localStorage.removeItem(PROFILE_IMAGE_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear profile image preference storage:', error);
    }
}

function getEffectiveProfileImagesEnabled() {
    const serverEnabled = (serverShowProfileImages != null)
        ? serverShowProfileImages
        : (!appSettings || appSettings.showProfileImages !== false);
    if (userProfileImagesPreference !== null) {
        return userProfileImagesPreference;
    }
    return serverEnabled;
}

function loadStoredEmployeeCountPreference() {
    userShowEmployeeCountPreference = null;
    try {
        const stored = localStorage.getItem(SHOW_EMPLOYEE_COUNT_PREFERENCE_KEY);
        if (stored === 'true') {
            userShowEmployeeCountPreference = true;
        } else if (stored === 'false') {
            userShowEmployeeCountPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access employee count preference storage:', error);
        userShowEmployeeCountPreference = null;
    }
}

function storeEmployeeCountPreference(value) {
    userShowEmployeeCountPreference = value;
    try {
        localStorage.setItem(SHOW_EMPLOYEE_COUNT_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist employee count preference:', error);
    }
}

function clearEmployeeCountPreference() {
    userShowEmployeeCountPreference = null;
    try {
        localStorage.removeItem(SHOW_EMPLOYEE_COUNT_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear employee count preference storage:', error);
    }
}

function getEffectiveEmployeeCountEnabled() {
    const serverEnabled = (serverShowEmployeeCount != null)
        ? serverShowEmployeeCount
        : (!appSettings || appSettings.showEmployeeCount !== false);
    if (userShowEmployeeCountPreference !== null) {
        return userShowEmployeeCountPreference;
    }
    return serverEnabled;
}

function loadSessionHighlightPreference() {
    sessionHighlightNewEmployeesPreference = null;
    try {
        if (!window.sessionStorage) {
            return;
        }
        const stored = window.sessionStorage.getItem(HIGHLIGHT_NEW_EMPLOYEES_SESSION_KEY);
        if (stored === 'true') {
            sessionHighlightNewEmployeesPreference = true;
        } else if (stored === 'false') {
            sessionHighlightNewEmployeesPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access highlight preference storage:', error);
        sessionHighlightNewEmployeesPreference = null;
    }
}

function storeSessionHighlightPreference(value) {
    sessionHighlightNewEmployeesPreference = value;
    try {
        if (!window.sessionStorage) {
            return;
        }
        window.sessionStorage.setItem(HIGHLIGHT_NEW_EMPLOYEES_SESSION_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist highlight preference:', error);
    }
}

function clearSessionHighlightPreference() {
    sessionHighlightNewEmployeesPreference = null;
    try {
        if (!window.sessionStorage) {
            return;
        }
        window.sessionStorage.removeItem(HIGHLIGHT_NEW_EMPLOYEES_SESSION_KEY);
    } catch (error) {
        console.warn('Unable to clear highlight preference storage:', error);
    }
}

function getEffectiveHighlightNewEmployees() {
    if (sessionHighlightNewEmployeesPreference !== null) {
        return sessionHighlightNewEmployeesPreference;
    }
    return false;
}

function loadStoredDepartmentPreference() {
    userShowDepartmentsPreference = null;
    try {
        const stored = localStorage.getItem(SHOW_DEPARTMENTS_PREFERENCE_KEY);
        if (stored === 'true') {
            userShowDepartmentsPreference = true;
        } else if (stored === 'false') {
            userShowDepartmentsPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access department visibility preference storage:', error);
        userShowDepartmentsPreference = null;
    }
}

function storeDepartmentPreference(value) {
    userShowDepartmentsPreference = value;
    try {
        localStorage.setItem(SHOW_DEPARTMENTS_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist department visibility preference:', error);
    }
}

function clearDepartmentPreference() {
    userShowDepartmentsPreference = null;
    try {
        localStorage.removeItem(SHOW_DEPARTMENTS_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear department visibility preference storage:', error);
    }
}

function getEffectiveDepartmentsEnabled() {
    const serverEnabled = (serverShowDepartments != null)
        ? serverShowDepartments
        : (!appSettings || appSettings.showDepartments !== false);
    if (userShowDepartmentsPreference !== null) {
        return userShowDepartmentsPreference;
    }
    return serverEnabled;
}

function loadStoredJobTitlePreference() {
    userShowJobTitlesPreference = null;
    try {
        const stored = localStorage.getItem(SHOW_JOB_TITLES_PREFERENCE_KEY);
        if (stored === 'true') {
            userShowJobTitlesPreference = true;
        } else if (stored === 'false') {
            userShowJobTitlesPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access job title visibility preference storage:', error);
        userShowJobTitlesPreference = null;
    }
}

function storeJobTitlePreference(value) {
    userShowJobTitlesPreference = value;
    try {
        localStorage.setItem(SHOW_JOB_TITLES_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist job title visibility preference:', error);
    }
}

function clearJobTitlePreference() {
    userShowJobTitlesPreference = null;
    try {
        localStorage.removeItem(SHOW_JOB_TITLES_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear job title visibility preference storage:', error);
    }
}

function loadStoredOfficePreference() {
    userShowOfficePreference = null;
    try {
        const stored = localStorage.getItem(SHOW_OFFICE_PREFERENCE_KEY);
        if (stored === 'true') {
            userShowOfficePreference = true;
        } else if (stored === 'false') {
            userShowOfficePreference = false;
        }
    } catch (error) {
        console.warn('Unable to access office visibility preference storage:', error);
        userShowOfficePreference = null;
    }
}

function storeOfficePreference(value) {
    userShowOfficePreference = value;
    try {
        localStorage.setItem(SHOW_OFFICE_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist office visibility preference:', error);
    }
}

function clearOfficePreference() {
    userShowOfficePreference = null;
    try {
        localStorage.removeItem(SHOW_OFFICE_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear office visibility preference storage:', error);
    }
}

function loadStoredNamePreference() {
    userShowNamesPreference = null;
    try {
        const stored = localStorage.getItem(SHOW_NAMES_PREFERENCE_KEY);
        if (stored === 'true') {
            userShowNamesPreference = true;
        } else if (stored === 'false') {
            userShowNamesPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access name visibility preference storage:', error);
        userShowNamesPreference = null;
    }
}

function storeNamePreference(value) {
    userShowNamesPreference = value;
    try {
        localStorage.setItem(SHOW_NAMES_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist name visibility preference:', error);
    }
}

function clearNamePreference() {
    userShowNamesPreference = null;
    try {
        localStorage.removeItem(SHOW_NAMES_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear name visibility preference storage:', error);
    }
}

function getEffectiveJobTitlesEnabled() {
    const serverEnabled = (serverShowJobTitles != null)
        ? serverShowJobTitles
        : (!appSettings || appSettings.showJobTitles !== false);
    if (userShowJobTitlesPreference !== null) {
        return userShowJobTitlesPreference;
    }
    return serverEnabled;
}

function getEffectiveOfficeEnabled() {
    const serverEnabled = (serverShowOffice != null)
        ? serverShowOffice
        : Boolean(appSettings && appSettings.showOffice === true);
    if (userShowOfficePreference !== null) {
        return userShowOfficePreference;
    }
    return serverEnabled;
}

function getEffectiveNamesEnabled() {
    const serverEnabled = (serverShowNames != null)
        ? serverShowNames
        : (!appSettings || appSettings.showNames !== false);
    if (userShowNamesPreference !== null) {
        return userShowNamesPreference;
    }
    return serverEnabled;
}

function getVisibleNameText(person, { includeFallback = true, fallback } = {}) {
    const rawName = (person && typeof person.name === 'string') ? person.name.trim() : '';
    if (!isNameVisible()) {
        if (!includeFallback) {
            return '';
        }
        if (fallback) {
            return fallback;
        }
        const translation = t('index.employee.detail.nameHidden');
        return translation === 'index.employee.detail.nameHidden' ? 'Name hidden' : translation;
    }
    if (rawName) {
        return person.name;
    }
    if (!includeFallback) {
        return '';
    }
    if (fallback) {
        return fallback;
    }
    const translation = t('index.employee.detail.nameUnknown');
    return translation === 'index.employee.detail.nameUnknown' ? 'Unknown name' : translation;
}

function isNameVisible() {
    return !appSettings || appSettings.showNames !== false;
}

function isJobTitleVisible() {
    return !appSettings || appSettings.showJobTitles !== false;
}

function isDepartmentVisible() {
    return !appSettings || appSettings.showDepartments !== false;
}

function isOfficeVisible() {
    return !!(appSettings && appSettings.showOffice);
}

function getVisibleJobTitleText(person, { includeFallback = true, useOverrides = true } = {}) {
    if (!isJobTitleVisible()) return '';
    const employeeId = person && (person.id || person.employeeId || person.data?.id);
    let title = '';
    if (useOverrides && employeeId && isTitleOverridden(employeeId)) {
        title = titleOverrides.get(employeeId);
    }
    if (!title && person && typeof person.title === 'string') {
        title = person.title.trim();
    }
    if (title) {
        return title;
    }
    if (!includeFallback) {
        return '';
    }
    const translation = t('index.employee.noTitle');
    return translation === 'index.employee.noTitle' ? 'No Title' : translation;
}

function getVisibleDepartmentText(person, { includeFallback = false, fallback, useOverrides = true } = {}) {
    if (!isDepartmentVisible()) return '';
    const employeeId = person && (person.id || person.employeeId || person.data?.id);
    let department = '';
    if (useOverrides && employeeId && isDepartmentOverridden(employeeId)) {
        department = departmentOverrides.get(employeeId);
    }
    if (!department && person && typeof person.department === 'string') {
        department = person.department.trim();
    }
    if (department) {
        return department;
    }
    if (!includeFallback) {
        return '';
    }
    if (fallback) {
        return fallback;
    }
    const translation = t('index.employee.detail.departmentUnknown');
    return translation === 'index.employee.detail.departmentUnknown' ? 'Unknown department' : translation;
}

function getVisibleOfficeText(person) {
    if (!isOfficeVisible()) {
        return '';
    }
    if (!person || typeof person !== 'object') {
        return '';
    }
    const rawLocation = typeof person.location === 'string' ? person.location.trim() : '';
    if (rawLocation) {
        return rawLocation;
    }
    const rawOffice = typeof person.officeLocation === 'string' ? person.officeLocation.trim() : '';
    return rawOffice;
}

function getDepartmentDisplayText(person, options = {}) {
    const departmentsVisible = isDepartmentVisible();
    if (!departmentsVisible) {
        return '';
    }
    const baseDepartment = getVisibleDepartmentText(person, options);
    if (!isOfficeVisible()) {
        return baseDepartment;
    }
    const officeText = getVisibleOfficeText(person);
    if (baseDepartment && officeText) {
        return `${baseDepartment} (${officeText})`;
    }
    if (!baseDepartment && officeText) {
        return officeText;
    }
    return baseDepartment;
}

function persistHiddenIds() {
    localStorage.setItem('hiddenNodeIds', JSON.stringify(Array.from(hiddenNodeIds)));
}

function isHiddenNode(node) {
    // A node is hidden if itself or any ancestor is marked
    let cur = node;
    while (cur) {
        if (hiddenNodeIds.has(cur.data.id)) return true;
        cur = cur.parent;
    }
    return false;
}

function toggleHideNode(d) {
    if (!d || !d.data || !d.data.id) return;
    if (hiddenNodeIds.has(d.data.id)) {
        hiddenNodeIds.delete(d.data.id);
    } else {
        hiddenNodeIds.add(d.data.id);
    }
    persistHiddenIds();
    updateHiddenResetButtonState();
    update(d);
}

function resetHiddenSubtrees() {
    if (hiddenNodeIds.size === 0) return;
    hiddenNodeIds.clear();
    persistHiddenIds();
    updateHiddenResetButtonState();
    update(root);
}

const API_BASE_URL = window.location.origin;
const nodeWidth = 220;
const nodeHeight = 80;
const levelHeight = 102;
const HORIZONTAL_LEVEL_HEIGHT = 130;
const VERTICAL_NODE_SPACING = nodeWidth + 8;
const HORIZONTAL_NODE_SPACING = nodeWidth + 26;
const VERTICAL_MULTILINE_SPACING = nodeWidth + 10;
const HORIZONTAL_MULTILINE_SPACING = nodeWidth + 36;
const VERTICAL_COMPACT_HORIZONTAL = nodeWidth + 11;
const HORIZONTAL_COMPACT_HORIZONTAL = nodeWidth + 40;
const VERTICAL_COMPACT_VERTICAL = levelHeight + 6;
const HORIZONTAL_COMPACT_VERTICAL = HORIZONTAL_LEVEL_HEIGHT + 20;

// Zoom tracking and helpers
let userAdjustedZoom = false;
let programmaticZoomActive = false;
let resizeTimer = null;
const RESIZE_DEBOUNCE_MS = 180;

function applyZoomTransform(transform, { duration = 750, resetUser = false } = {}) {
    if (!svg || !zoom) return;
    programmaticZoomActive = true;

    const finalize = () => {
        programmaticZoomActive = false;
        if (resetUser) {
            userAdjustedZoom = false;
        }
    };

    if (duration > 0) {
        const transition = svg.transition().duration(duration).call(zoom.transform, transform);
        transition.on('end', finalize);
        transition.on('interrupt', finalize);
    } else {
        svg.call(zoom.transform, transform);
        finalize();
    }
}

function updateSvgSize() {
    if (!svg) return;
    const container = document.getElementById('orgChart');
    if (!container) return;
    const width = container.clientWidth;
    const height = container.clientHeight || 800;
    svg.attr('width', width).attr('height', height);
}

function createTreeLayout() {
    const layout = d3.tree()
        .nodeSize(currentLayout === 'vertical'
            ? [VERTICAL_NODE_SPACING, levelHeight]
            : [HORIZONTAL_LEVEL_HEIGHT, HORIZONTAL_NODE_SPACING])
        .separation((a, b) => {
            const sameParent = a.parent && b.parent && a.parent === b.parent;
            // Keep siblings at full spacing to avoid overlap even for large teams
            return sameParent ? 1.0 : 1.2;
        });
    return layout;
}
const userIconUrl = window.location.origin + '/static/usericon.png';


// Security: HTML escaping function to prevent XSS
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function renderAvatar({ imageUrl, name, initials, imageClass, fallbackClass }) {
    const safeInitials = escapeHtml(initials || '');
    if (imageUrl && appSettings.showProfileImages !== false) {
        const safeUrl = escapeHtml(imageUrl);
        return `
            <img class="${imageClass}" src="${safeUrl}" alt="${escapeHtml(name || '')}" data-role="avatar-image">
            <div class="${fallbackClass}" data-role="avatar-fallback" hidden>${safeInitials}</div>
        `;
    }
    return `<div class="${fallbackClass}" data-role="avatar-fallback">${safeInitials}</div>`;
}

// Date formatting function to format hire dates to yyyy-MM-dd
function formatHireDate(dateString) {
    if (!dateString) return '';
    try {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString; // Return original if invalid
        return date.toISOString().split('T')[0]; // Returns yyyy-MM-dd format
    } catch (e) {
        return dateString; // Return original if parsing fails
    }
}

// Dynamic font scaling based on text length
function calculateFontSize(text, baseSize, maxLength, minSize = 9) {
    if (!text) return baseSize;
    const length = text.length;
    if (length <= maxLength * 0.7) return baseSize; // Normal size for short text
    if (length <= maxLength) return Math.max(baseSize * 0.9, minSize); // Slightly smaller for medium text
    return Math.max(baseSize * 0.75, minSize); // Smaller for long text
}

function getLabelOffsetX() {
    return appSettings.showProfileImages !== false ? -nodeWidth / 2 + 50 : 0;
}

function getLabelAnchor() {
    return appSettings.showProfileImages !== false ? 'start' : 'middle';
}

function getNameFontSizePx(name) {
    const showImages = appSettings.showProfileImages !== false;
    const maxLength = showImages ? 25 : 32;
    const baseSize = showImages ? 14 : 16;
    return calculateFontSize(name, baseSize, maxLength) + 'px';
}

function getTitleFontSizePx(title) {
    const showImages = appSettings.showProfileImages !== false;
    const maxLength = showImages ? 25 : 32;
    const baseSize = showImages ? 11 : 13;
    const minSize = showImages ? 8 : 10;
    return calculateFontSize(title, baseSize, maxLength, minSize) + 'px';
}

function getDepartmentFontSizePx(dept) {
    const showImages = appSettings.showProfileImages !== false;
    const maxLength = showImages ? 25 : 38;
    const baseSize = showImages ? 9 : 11;
    const minSize = showImages ? 7 : 9;
    return calculateFontSize(dept, baseSize, maxLength, minSize) + 'px';
}

function getTrimmedTitle(title = '') {
    const charLimit = appSettings.showProfileImages !== false ? 45 : 50;
    return title.length > charLimit ? title.substring(0, charLimit) + '...' : title;
}

function getDirectReportCount(node) {
    if (!node) {
        return 0;
    }
    const directReports = node._children?.length || node.children?.length || 0;
    return directReports;
}

function shouldShowCountBadge(node) {
    return appSettings.showEmployeeCount !== false && getDirectReportCount(node) > 0;
}

function formatDirectReportCount(node) {
    const count = getDirectReportCount(node);
    return count > 99 ? '99+' : String(count);
}

// Security: Safely set innerHTML with escaped content
function safeInnerHTML(element, htmlContent) {
    element.innerHTML = htmlContent;
}

function applyProfileImageAttributes(selection) {
    selection
        .attr('class', 'profile-image')
        .attr('xlink:href', userIconUrl)
        .attr('x', -nodeWidth / 2 + 8)
        .attr('y', -18)
        .attr('width', 36)
        .attr('height', 36)
        .attr('clip-path', 'circle(18px at 18px 18px)')
        .attr('preserveAspectRatio', 'xMidYMid slice')
        .each(function(d) {
            if (d.data.photoUrl && d.data.photoUrl.includes('/api/photo/')) {
                const element = d3.select(this);
                const img = new Image();

                img.onload = function() {
                    element.attr('xlink:href', d.data.photoUrl);
                    console.log(`Photo loaded for ${d.data.name}`);
                };

                img.onerror = function() {
                    console.log(`Photo failed for ${d.data.name}, keeping default icon`);
                };

                img.src = d.data.photoUrl;
            }
        });
}

function formatLastUpdatedTime(value) {
    if (!value) return null;
    const text = typeof value === 'string' ? value.trim() : `${value}`;
    if (!text) return null;
    
    try {
        const date = new Date(text);
        if (!isNaN(date.getTime())) {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const day = String(date.getDate()).padStart(2, '0');
            const hours = String(date.getHours()).padStart(2, '0');
            const minutes = String(date.getMinutes()).padStart(2, '0');
            return `${year}-${month}-${day} ${hours}:${minutes}`;
        }
    } catch (e) {
        // Return null if parsing fails
    }
    return null;
}

function updateHeaderSubtitle(syncing = false) {
    const headerP = document.querySelector('.header-text p');
    if (!headerP) return;
    
    if (syncing) {
        headerP.textContent = t('index.header.autoUpdate.syncing', { defaultValue: 'Syncing...' });
        return;
    }
    
    if (!appSettings.updateTime) return;
    
    let displayTime = appSettings.updateTime;
    // Convert UTC time to user's local timezone for display
    try {
        const [hours, minutes] = appSettings.updateTime.split(':').map(Number);
        const now = new Date();
        const utcDate = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), hours, minutes));
        displayTime = utcDate.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit', hour12: false });
    } catch (e) {
        console.warn('Failed to convert update time to local timezone', e);
    }
    
    // Format last updated time
    const lastUpdated = formatLastUpdatedTime(appSettings.dataLastUpdatedAt);
    
    let timeText;
    if (lastUpdated) {
        timeText = appSettings.autoUpdateEnabled
            ? t('index.header.autoUpdate.enabledWithLastUpdate', { time: displayTime, lastUpdate: lastUpdated })
            : t('index.header.autoUpdate.disabledWithLastUpdate', { lastUpdate: lastUpdated });
    } else {
        timeText = appSettings.autoUpdateEnabled
            ? t('index.header.autoUpdate.enabled', { time: displayTime })
            : t('index.header.autoUpdate.disabled');
    }
    headerP.textContent = timeText;
}

async function loadSettings() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/settings`);
        if (response.ok) {
            appSettings = await response.json();
            if (!Object.prototype.hasOwnProperty.call(appSettings, 'showNames')) {
                appSettings.showNames = true;
            }
            if (!Object.prototype.hasOwnProperty.call(appSettings, 'showOffice')) {
                appSettings.showOffice = false;
            }
            serverShowEmployeeCount = appSettings.showEmployeeCount !== false;
            serverShowProfileImages = appSettings.showProfileImages !== false;
            serverShowDepartments = appSettings.showDepartments !== false;
            serverShowJobTitles = appSettings.showJobTitles !== false;
            serverShowOffice = appSettings.showOffice === true;
            serverShowNames = appSettings.showNames !== false;
            await applySettings();
        } else {
            // If settings fail to load, still show header content with defaults
            showHeaderContent();
        }
    } catch (error) {
        console.error('Error loading settings:', error);
        // If settings fail to load, still show header content with defaults
        showHeaderContent();
    }
}

function showHeaderContent() {
    // Show header content even if settings failed to load
    const headerContent = document.querySelector('.header-content');
    if (headerContent) {
        headerContent.classList.remove('loading');
    }
    const header = document.querySelector('.header');
    if (header) {
        header.classList.remove('is-loading');
    }
    
    // Also show default logo if settings failed
    const logo = document.querySelector('.header-logo');
    ensureLogoFallback(logo);
    if (logo && !logo.src) {
        logo.src = logo.dataset.defaultSrc || '/static/icon.png';
        logo.classList.remove('loading');
        logo.style.display = '';
    }
}

function ensureLogoFallback(logo) {
    if (!logo || logo.dataset.fallbackBound === 'true') {
        return;
    }

    const handleLoad = () => {
        logo.style.display = '';
    };

    const handleError = () => {
        logo.style.display = 'none';
    };

    logo.addEventListener('load', handleLoad);
    logo.addEventListener('error', handleError);
    logo.dataset.fallbackBound = 'true';
}

function setupStaticEventListeners() {
    const configBtn = document.getElementById('configBtn');
    if (configBtn) {
        configBtn.addEventListener('click', () => {
            window.location.href = '/configure';
        });
    }

    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
        loginBtn.addEventListener('click', () => {
            const currentPath = window.location.pathname.replace(/^\/+/, '');
            const loginUrl = currentPath ? `/login?next=${encodeURIComponent(currentPath)}` : '/login';
            window.location.href = loginUrl;
        });
    }

    const reportsBtn = document.getElementById('reportsBtn');
    if (reportsBtn) {
        reportsBtn.addEventListener('click', () => {
            window.location.href = '/reports';
        });
    }

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', () => {
            logout();
        });
    }

    const syncBtn = document.getElementById('syncBtn');
    if (syncBtn) {
        syncBtn.addEventListener('click', () => {
            triggerDataSync();
        });
    }

    document.querySelectorAll('[data-layout]').forEach(button => {
        button.addEventListener('click', () => {
            if (button.dataset.layout) {
                setLayoutOrientation(button.dataset.layout);
            }
        });
    });

    const controls = document.querySelector('.controls');
    if (controls) {
        controls.addEventListener('click', event => {
            const button = event.target.closest('[data-control]');
            if (!button) return;
            handleControlAction(button.dataset.control);
        });
    }

    const saveTopUserBtn = document.getElementById('saveTopUserBtn');
    if (saveTopUserBtn) {
        saveTopUserBtn.addEventListener('click', saveTopUser);
    }

    const resetTopUserBtn = document.getElementById('resetTopUserBtn');
    if (resetTopUserBtn) {
        resetTopUserBtn.addEventListener('click', resetTopUser);
    }

    const compactBtn = document.getElementById('compactToggleBtn');
    if (compactBtn) {
        compactBtn.addEventListener('click', toggleCompactLargeTeams);
    }

    const employeeCountBtn = document.getElementById('employeeCountToggleBtn');
    if (employeeCountBtn) {
        employeeCountBtn.addEventListener('click', toggleEmployeeCountVisibility);
    }

    const highlightBtn = document.getElementById('highlightNewEmployeesToggleBtn');
    if (highlightBtn) {
        highlightBtn.addEventListener('click', toggleHighlightNewEmployees);
    }

    const profileBtn = document.getElementById('profileImageToggleBtn');
    if (profileBtn) {
        profileBtn.addEventListener('click', toggleProfileImages);
    }

    const nameBtn = document.getElementById('nameToggleBtn');
    if (nameBtn) {
        nameBtn.addEventListener('click', toggleNameVisibility);
    }

    const departmentBtn = document.getElementById('departmentToggleBtn');
    if (departmentBtn) {
        departmentBtn.addEventListener('click', toggleDepartmentVisibility);
    }

    const officeBtn = document.getElementById('officeToggleBtn');
    if (officeBtn) {
        officeBtn.addEventListener('click', toggleOfficeVisibility);
    }

    const jobTitleBtn = document.getElementById('jobTitleToggleBtn');
    if (jobTitleBtn) {
        jobTitleBtn.addEventListener('click', toggleJobTitleVisibility);
    }

    const closeDetailBtn = document.getElementById('employeeDetailCloseBtn');
    if (closeDetailBtn) {
        closeDetailBtn.addEventListener('click', closeEmployeeDetail);
    }

    const resultsContainer = document.getElementById('searchResults');
    if (resultsContainer) {
        resultsContainer.addEventListener('click', event => {
            const item = event.target.closest('.search-result-item');
            if (!item) return;
            const employeeId = item.dataset.employeeId;
            if (employeeId) {
                selectSearchResult(employeeId);
            }
        });
    }

    const infoPanel = document.getElementById('employeeInfo');
    if (infoPanel) {
        infoPanel.addEventListener('click', event => {
            const target = event.target.closest('[data-employee-id]');
            if (!target) return;
            showEmployeeDetailById(target.dataset.employeeId);
        });
    }

    const logo = document.querySelector('.header-logo');
    if (logo) {
        ensureLogoFallback(logo);
    }

    setLayoutOrientation(currentLayout);

    document.addEventListener('click', event => {
        const editTrigger = event.target.closest('[data-action="edit-title"], [data-action="edit-department"]');
        if (!editTrigger) return;
        const employeeId = editTrigger.dataset.employeeId;
        if (employeeId && employeeById.has(employeeId)) {
            const focusField = editTrigger.dataset.focusField
                || (editTrigger.dataset.action === 'edit-department' ? 'department' : 'title');
            openTitleEditModal(employeeById.get(employeeId), { focusField });
        }
    });

    updateOverrideResetButtons();
}

function handleControlAction(action) {
    switch (action) {
        case 'zoom-in':
            zoomIn();
            break;
        case 'zoom-out':
            zoomOut();
            break;
        case 'reset-zoom':
            resetZoom();
            break;
        case 'fit':
            fitToScreen();
            break;
        case 'expand':
            expandAll();
            break;
        case 'collapse':
            collapseAll();
            break;
        case 'reset-hidden':
            resetHiddenSubtrees();
            break;
        case 'reset-titles':
            resetTitleOverrides();
            break;
        case 'reset-departments':
            resetDepartmentOverrides();
            break;
        case 'print':
            printChart();
            break;
        case 'export-visible-svg':
            exportToImage('svg', false);
            break;
        case 'export-visible-png':
            exportToImage('png', false);
            break;
        case 'export-visible-pdf':
            exportToPDF(false);
            break;
        case 'export-xlsx':
            exportToXLSX();
            break;
        default:
            break;
    }
}

function resetTitleOverrides() {
    const cleared = clearAllTitleOverrides();
    if (!cleared) {
        return;
    }
    refreshAfterOverrideChange();
}

function resetDepartmentOverrides() {
    const cleared = clearAllDepartmentOverrides();
    if (!cleared) {
        return;
    }
    refreshAfterOverrideChange();
}

async function applySettings() {
    await waitForTranslations();
    if (appSettings.chartTitle) {
        document.querySelector('.header-text h1').textContent = appSettings.chartTitle;
        // Update the browser tab title to match the custom title
        document.title = appSettings.chartTitle;
    } else {
        // Fallback to default title if no custom title is set
    document.title = 'SimpleOrgChart';
    }

    if (appSettings.headerColor) {
        const header = document.querySelector('.header');
        const darker = adjustColor(appSettings.headerColor, -30);
        header.style.background = `linear-gradient(135deg, ${appSettings.headerColor} 0%, ${darker} 100%)`;
    }

    // Handle logo loading
    const logo = document.querySelector('.header-logo');
    ensureLogoFallback(logo);
    if (appSettings.logoPath) {
        logo.src = appSettings.logoPath + '?t=' + Date.now();
    } else {
        // Use default logo if no custom logo is set
        logo.src = logo.dataset.defaultSrc || '/static/icon.png';
    }
    logo.classList.remove('loading'); // Show logo after src is set
    logo.style.display = '';

    updateHeaderSubtitle();
    
    // Show header content after settings are applied to prevent flash of default content
    const headerContent = document.querySelector('.header-content');
    if (headerContent) {
        headerContent.classList.remove('loading');
    }
    const header = document.querySelector('.header');
    if (header) {
        header.classList.remove('is-loading');
    }

    // Reflect Compact Teams toggle state
    try {
        const btn = document.getElementById('compactToggleBtn');
        if (btn && appSettings) {
            const enabled = getEffectiveCompactEnabled();
            appSettings.multiLineChildrenEnabled = enabled;
            btn.classList.toggle('active', enabled);

            // Update label to include threshold number
            const threshold = (appSettings.multiLineChildrenThreshold != null)
                ? appSettings.multiLineChildrenThreshold
                : 20;
            const labelSpan = btn.querySelector('[data-i18n="index.toolbar.layout.compactLabel"]')
                || btn.querySelector('.layout-label');
            const compactLabel = t('index.toolbar.layout.compactLabelWithThreshold', { count: threshold });
            if (labelSpan) {
                labelSpan.textContent = compactLabel;
            } else {
                btn.innerHTML = `<span class="layout-icon">▦</span> ${compactLabel}`;
            }
            btn.setAttribute('aria-label', compactLabel);
        }
    } catch (e) { /* no-op */ }

    const showEmployeeCount = getEffectiveEmployeeCountEnabled();
    appSettings.showEmployeeCount = showEmployeeCount;
    const employeeCountBtn = document.getElementById('employeeCountToggleBtn');
    if (employeeCountBtn) {
        employeeCountBtn.classList.toggle('active', showEmployeeCount);
        employeeCountBtn.setAttribute('aria-pressed', String(showEmployeeCount));
        const countHide = t('index.toolbar.layout.employeeCountHide', { defaultValue: 'Hide employee count badges' });
        const countShow = t('index.toolbar.layout.employeeCountShow', { defaultValue: 'Show employee count badges' });
        const countTitle = showEmployeeCount ? countHide : countShow;
        employeeCountBtn.title = countTitle;
        employeeCountBtn.setAttribute('aria-label', countTitle);
    }

    const highlightNewEmployeesEnabled = getEffectiveHighlightNewEmployees();
    appSettings.highlightNewEmployees = highlightNewEmployeesEnabled;
    const highlightBtn = document.getElementById('highlightNewEmployeesToggleBtn');
    if (highlightBtn) {
        highlightBtn.classList.toggle('active', highlightNewEmployeesEnabled);
        highlightBtn.setAttribute('aria-pressed', String(highlightNewEmployeesEnabled));
        const highlightHide = t('index.toolbar.layout.newEmployeeHighlightHide', { defaultValue: 'Hide new employee highlights' });
        const highlightShow = t('index.toolbar.layout.newEmployeeHighlightShow', { defaultValue: 'Show new employee highlights' });
        const highlightTitle = highlightNewEmployeesEnabled ? highlightHide : highlightShow;
        highlightBtn.title = highlightTitle;
        highlightBtn.setAttribute('aria-label', highlightTitle);
    }

    const showProfileImages = getEffectiveProfileImagesEnabled();
    appSettings.showProfileImages = showProfileImages;
    const profileBtn = document.getElementById('profileImageToggleBtn');
    if (profileBtn) {
        profileBtn.classList.toggle('active', showProfileImages);
        profileBtn.setAttribute('aria-pressed', String(showProfileImages));
        const profileTitle = showProfileImages
            ? t('index.toolbar.layout.profileHide')
            : t('index.toolbar.layout.profileShow');
        profileBtn.title = profileTitle;
        profileBtn.setAttribute('aria-label', profileTitle);
    }

    const showNames = getEffectiveNamesEnabled();
    appSettings.showNames = showNames;
    const nameBtn = document.getElementById('nameToggleBtn');
    if (nameBtn) {
        nameBtn.classList.toggle('active', showNames);
        nameBtn.setAttribute('aria-pressed', String(showNames));
        const nameHide = t('index.toolbar.layout.nameHide', { defaultValue: 'Hide Names' });
        const nameShow = t('index.toolbar.layout.nameShow', { defaultValue: 'Show Names' });
        const nameTitle = showNames ? nameHide : nameShow;
        nameBtn.title = nameTitle;
        nameBtn.setAttribute('aria-label', nameTitle);
    }

    const showDepartments = getEffectiveDepartmentsEnabled();
    appSettings.showDepartments = showDepartments;
    const departmentBtn = document.getElementById('departmentToggleBtn');
    if (departmentBtn) {
        departmentBtn.classList.toggle('active', showDepartments);
        departmentBtn.setAttribute('aria-pressed', String(showDepartments));
        const deptHide = t('index.toolbar.layout.departmentHide', { defaultValue: 'Hide Departments' });
        const deptShow = t('index.toolbar.layout.departmentShow', { defaultValue: 'Show Departments' });
        const departmentTitle = showDepartments ? deptHide : deptShow;
        departmentBtn.title = departmentTitle;
        departmentBtn.setAttribute('aria-label', departmentTitle);
    }

    const showOffice = getEffectiveOfficeEnabled();
    appSettings.showOffice = showOffice;
    const officeBtn = document.getElementById('officeToggleBtn');
    if (officeBtn) {
        officeBtn.classList.toggle('active', showOffice);
        officeBtn.setAttribute('aria-pressed', String(showOffice));
        const officeHide = t('index.toolbar.layout.officeHide', { defaultValue: 'Hide office locations' });
        const officeShow = t('index.toolbar.layout.officeShow', { defaultValue: 'Show office locations' });
        const officeTitle = showOffice ? officeHide : officeShow;
        officeBtn.title = officeTitle;
        officeBtn.setAttribute('aria-label', officeTitle);
    }

    const showJobTitles = getEffectiveJobTitlesEnabled();
    appSettings.showJobTitles = showJobTitles;
    const titleBtn = document.getElementById('jobTitleToggleBtn');
    if (titleBtn) {
        titleBtn.classList.toggle('active', showJobTitles);
        titleBtn.setAttribute('aria-pressed', String(showJobTitles));
        const titleHide = t('index.toolbar.layout.jobTitleHide', { defaultValue: 'Hide Job Titles' });
        const titleShow = t('index.toolbar.layout.jobTitleShow', { defaultValue: 'Show Job Titles' });
        const titleToggle = showJobTitles ? titleHide : titleShow;
        titleBtn.title = titleToggle;
        titleBtn.setAttribute('aria-label', titleToggle);
    }

    await ensureIdentityFieldMinimum({ source: 'apply-settings', skipUpdateAuth: true });

    await updateAuthDependentUI();

    // Check if a sync is already in progress and show the syncing state
    const updateStatus = appSettings.dataUpdateStatus || {};
    if (updateStatus.state === 'running' && isAuthenticated) {
        setSyncButtonState(true);
        updateHeaderSubtitle(true); // Show syncing in header
        startSyncPolling();
    }
}

function adjustColor(color, amount) {
    const num = parseInt(color.replace('#', ''), 16);
    const r = Math.max(0, Math.min(255, (num >> 16) + amount));
    const g = Math.max(0, Math.min(255, ((num >> 8) & 0x00FF) + amount));
    const b = Math.max(0, Math.min(255, (num & 0x0000FF) + amount));
    return '#' + ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0');
}

// No conversion needed; we display and store updateTime in 24-hour HH:MM alongside timezone info

function setLayoutOrientation(orientation) {
    currentLayout = orientation;

    document.querySelectorAll('[data-layout]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.layout === orientation);
    });

    if (root) {
        update(root);
        fitToScreen();
    }
}

function updateAdminActions() {
    const configBtn = document.getElementById('configBtn');
    if (configBtn) {
        configBtn.classList.toggle('is-hidden', !isAuthenticated);
    }

    const loginBtn = document.getElementById('loginBtn');
    if (loginBtn) {
        loginBtn.classList.toggle('is-hidden', isAuthenticated);
    }

    const syncBtn = document.getElementById('syncBtn');
    if (syncBtn) {
        syncBtn.classList.toggle('is-hidden', !isAuthenticated);
    }

    const reportsBtn = document.getElementById('reportsBtn');
    if (reportsBtn) {
        reportsBtn.classList.toggle('is-hidden', !isAuthenticated);
    }

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
        logoutBtn.classList.toggle('is-hidden', !isAuthenticated);
    }
}

async function updateAuthDependentUI() {
    await waitForTranslations();
    updateAdminActions();

    const exportXlsxBtn = document.querySelector('[data-control="export-xlsx"]');
    if (exportXlsxBtn) {
        const baseLabel = t('index.toolbar.controls.exportXlsx');
        const adminLabel = t('index.toolbar.controls.exportXlsxAdmin');
        const label = isAuthenticated ? adminLabel : baseLabel;
        exportXlsxBtn.textContent = label;
        exportXlsxBtn.setAttribute('aria-label', label);
        exportXlsxBtn.title = label;
    }

    const compactBtn = document.getElementById('compactToggleBtn');
    if (compactBtn) {
        compactBtn.disabled = false;
        compactBtn.removeAttribute('aria-disabled');
        const enabled = getEffectiveCompactEnabled();
        compactBtn.classList.toggle('active', enabled);
        const compactTitle = isAuthenticated
            ? t('index.toolbar.layout.compactTitleAdmin')
            : t('index.toolbar.layout.compactTitleGuest');
        compactBtn.title = compactTitle;
        compactBtn.setAttribute('aria-label', compactTitle);
    }

    const employeeCountBtn = document.getElementById('employeeCountToggleBtn');
    if (employeeCountBtn) {
        const showCount = getEffectiveEmployeeCountEnabled();
        employeeCountBtn.classList.toggle('active', showCount);
        employeeCountBtn.setAttribute('aria-pressed', String(showCount));
        const countHide = t('index.toolbar.layout.employeeCountHide', { defaultValue: 'Hide employee count badges' });
        const countShow = t('index.toolbar.layout.employeeCountShow', { defaultValue: 'Show employee count badges' });
        const countTitle = showCount ? countHide : countShow;
        employeeCountBtn.title = countTitle;
        employeeCountBtn.setAttribute('aria-label', countTitle);
    }

    const highlightBtn = document.getElementById('highlightNewEmployeesToggleBtn');
    if (highlightBtn) {
        const highlightEnabled = getEffectiveHighlightNewEmployees();
        highlightBtn.classList.toggle('active', highlightEnabled);
        highlightBtn.setAttribute('aria-pressed', String(highlightEnabled));
        const highlightHide = t('index.toolbar.layout.newEmployeeHighlightHide', { defaultValue: 'Hide new employee highlights' });
        const highlightShow = t('index.toolbar.layout.newEmployeeHighlightShow', { defaultValue: 'Show new employee highlights' });
        const highlightTitle = highlightEnabled ? highlightHide : highlightShow;
        highlightBtn.title = highlightTitle;
        highlightBtn.setAttribute('aria-label', highlightTitle);
    }

    const profileBtn = document.getElementById('profileImageToggleBtn');
    if (profileBtn) {
        const showImages = getEffectiveProfileImagesEnabled();
        profileBtn.classList.toggle('active', showImages);
        profileBtn.setAttribute('aria-pressed', String(showImages));
        const profileTitle = showImages
            ? t('index.toolbar.layout.profileHide')
            : t('index.toolbar.layout.profileShow');
        profileBtn.title = profileTitle;
        profileBtn.setAttribute('aria-label', profileTitle);
    }

    const nameBtn = document.getElementById('nameToggleBtn');
    if (nameBtn) {
        const showNames = getEffectiveNamesEnabled();
        nameBtn.classList.toggle('active', showNames);
        nameBtn.setAttribute('aria-pressed', String(showNames));
        const nameHide = t('index.toolbar.layout.nameHide', { defaultValue: 'Hide Names' });
        const nameShow = t('index.toolbar.layout.nameShow', { defaultValue: 'Show Names' });
        const nameTitle = showNames ? nameHide : nameShow;
        nameBtn.title = nameTitle;
        nameBtn.setAttribute('aria-label', nameTitle);
    }

    const departmentBtn = document.getElementById('departmentToggleBtn');
    if (departmentBtn) {
        const showDepartments = getEffectiveDepartmentsEnabled();
        departmentBtn.classList.toggle('active', showDepartments);
        departmentBtn.setAttribute('aria-pressed', String(showDepartments));
        const deptHide = t('index.toolbar.layout.departmentHide', { defaultValue: 'Hide Departments' });
        const deptShow = t('index.toolbar.layout.departmentShow', { defaultValue: 'Show Departments' });
        const departmentTitle = showDepartments ? deptHide : deptShow;
        departmentBtn.title = departmentTitle;
        departmentBtn.setAttribute('aria-label', departmentTitle);
    }

    const officeBtn = document.getElementById('officeToggleBtn');
    if (officeBtn) {
        const showOffice = getEffectiveOfficeEnabled();
        officeBtn.classList.toggle('active', showOffice);
        officeBtn.setAttribute('aria-pressed', String(showOffice));
        const officeHide = t('index.toolbar.layout.officeHide', { defaultValue: 'Hide office locations' });
        const officeShow = t('index.toolbar.layout.officeShow', { defaultValue: 'Show office locations' });
        const officeTitle = showOffice ? officeHide : officeShow;
        officeBtn.title = officeTitle;
        officeBtn.setAttribute('aria-label', officeTitle);
    }

    const jobTitleBtn = document.getElementById('jobTitleToggleBtn');
    if (jobTitleBtn) {
        const showTitles = getEffectiveJobTitlesEnabled();
        jobTitleBtn.classList.toggle('active', showTitles);
        jobTitleBtn.setAttribute('aria-pressed', String(showTitles));
        const titleHide = t('index.toolbar.layout.jobTitleHide', { defaultValue: 'Hide Job Titles' });
        const titleShow = t('index.toolbar.layout.jobTitleShow', { defaultValue: 'Show Job Titles' });
        const jobTitle = showTitles ? titleHide : titleShow;
        jobTitleBtn.title = jobTitle;
        jobTitleBtn.setAttribute('aria-label', jobTitle);
    }
}

async function checkAuthentication() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/auth-check`, {
            credentials: 'same-origin'
        });
        return response.ok;
    } catch (error) {
        console.error('Authentication check failed:', error);
        return false;
    }
}

async function init() {
    const htmlElement = document.documentElement;

    try {
        isAuthenticated = await checkAuthentication();
        if (isAuthenticated) {
            userCompactPreference = null;
        } else {
            loadStoredCompactPreference();
        }
        loadStoredEmployeeCountPreference();
        loadStoredProfileImagePreference();
        loadStoredNamePreference();
        loadStoredDepartmentPreference();
        loadStoredOfficePreference();
        loadStoredJobTitlePreference();
        loadSessionHighlightPreference();

        await waitForTranslations();
        htmlElement.classList.remove('i18n-loading');

        await updateAuthDependentUI();
        await loadSettings();

        const response = await fetch(`${API_BASE_URL}/api/employees`);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        currentData = await response.json();
        window.currentOrgData = currentData; // Store globally for manager lookup
        
        if (currentData) {
            employeeById.clear();
            allEmployees = flattenTree(currentData);
            const validIds = allEmployees.map(emp => emp.id).filter(Boolean);
            pruneTitleOverrides(validIds);
            pruneDepartmentOverrides(validIds);
            initializeTopUserSearch();
            preloadEmployeeImages(allEmployees);
            renderOrgChart(currentData);
        } else {
            throw new Error('No data received from server');
        }
    } catch (error) {
        console.error('Error loading employee data:', error);
        const container = document.getElementById('orgChart');
        if (container) {
            const loading = container.querySelector('.loading');
            if (loading) {
                const spinner = loading.querySelector('.spinner');
                if (spinner) {
                    spinner.style.display = 'none';
                }
                const message = loading.querySelector('p');
                if (message) {
                    message.textContent = t('index.status.errorLoading');
                } else {
                    loading.textContent = t('index.status.errorLoading');
                }
                loading.style.display = '';
            }
        }
    } finally {
        htmlElement.classList.remove('i18n-loading');
    }
}

async function refreshOrgChart() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/employees`);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        currentData = await response.json();
        window.currentOrgData = currentData;
        
        if (currentData) {
            employeeById.clear();
            allEmployees = flattenTree(currentData);
            const validIds = allEmployees.map(emp => emp.id).filter(Boolean);
            pruneTitleOverrides(validIds);
            pruneDepartmentOverrides(validIds);
            initializeTopUserSearch();
            preloadEmployeeImages(allEmployees);
            renderOrgChart(currentData);
        }
    } catch (error) {
        console.error('Error refreshing org chart:', error);
    }
}

// Toggle Compact Teams from main page
async function toggleCompactLargeTeams() {
    await waitForTranslations();
    const btn = document.getElementById('compactToggleBtn');
    const previousValue = getEffectiveCompactEnabled();
    const newValue = !previousValue;

    if (btn) {
        btn.classList.toggle('active', newValue);
    }

    if (!appSettings) {
        appSettings = {};
    }

    if (!isAuthenticated) {
        storeCompactPreference(newValue);
        appSettings.multiLineChildrenEnabled = newValue;
        if (root) {
            update(root);
            fitToScreen();
        }
        await updateAuthDependentUI();
        return;
    }

    if (btn) {
        btn.disabled = true;
    }

    try {
        const res = await fetch(`${API_BASE_URL}/api/set-multiline-enabled`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ multiLineChildrenEnabled: newValue })
        });
        if (res.status === 401) {
            if (btn) {
                btn.classList.toggle('active', previousValue);
                btn.disabled = false;
            }
            isAuthenticated = false;
            await updateAuthDependentUI();
            alert(t('index.alerts.adminLoginExpired'));
            return;
        }
        if (!res.ok) {
            throw new Error('Failed to save Compact Teams');
        }

        clearCompactPreferenceStorage();
        appSettings.multiLineChildrenEnabled = newValue;
        if (btn) {
            btn.disabled = false;
        }
        await updateAuthDependentUI();
        if (root) {
            update(root);
            fitToScreen();
        }
    } catch (err) {
    console.error('Error toggling Compact Teams:', err);
        if (btn) {
            btn.classList.toggle('active', previousValue);
            btn.disabled = false;
        }
        appSettings.multiLineChildrenEnabled = previousValue;
        await updateAuthDependentUI();
    }
}

async function toggleEmployeeCountVisibility() {
    await waitForTranslations();
    const btn = document.getElementById('employeeCountToggleBtn');
    const currentValue = getEffectiveEmployeeCountEnabled();
    const newValue = !currentValue;

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.showEmployeeCount = newValue;

    if (serverShowEmployeeCount != null && newValue === serverShowEmployeeCount) {
        clearEmployeeCountPreference();
    } else {
        storeEmployeeCountPreference(newValue);
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        const countHide = t('index.toolbar.layout.employeeCountHide', { defaultValue: 'Hide employee count badges' });
        const countShow = t('index.toolbar.layout.employeeCountShow', { defaultValue: 'Show employee count badges' });
        const countTitle = newValue ? countHide : countShow;
        btn.title = countTitle;
        btn.setAttribute('aria-label', countTitle);
    }

    if (root) {
        update(root);
    }

    await updateAuthDependentUI();
}

async function toggleHighlightNewEmployees() {
    await waitForTranslations();
    const btn = document.getElementById('highlightNewEmployeesToggleBtn');
    const currentValue = getEffectiveHighlightNewEmployees();
    const newValue = !currentValue;

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.highlightNewEmployees = newValue;

    if (newValue) {
        storeSessionHighlightPreference(true);
    } else {
        clearSessionHighlightPreference();
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        const highlightHide = t('index.toolbar.layout.newEmployeeHighlightHide', { defaultValue: 'Hide new employee highlights' });
        const highlightShow = t('index.toolbar.layout.newEmployeeHighlightShow', { defaultValue: 'Show new employee highlights' });
        const highlightTitle = newValue ? highlightHide : highlightShow;
        btn.title = highlightTitle;
        btn.setAttribute('aria-label', highlightTitle);
    }

    if (root) {
        update(root);
    }

    await updateAuthDependentUI();
}

async function toggleProfileImages() {
    await waitForTranslations();
    const btn = document.getElementById('profileImageToggleBtn');
    const currentValue = getEffectiveProfileImagesEnabled();
    const newValue = !currentValue;

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.showProfileImages = newValue;

    if (serverShowProfileImages != null && newValue === serverShowProfileImages) {
        clearProfileImagePreference();
    } else {
        storeProfileImagePreference(newValue);
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        const profileTitle = newValue
            ? t('index.toolbar.layout.profileHide')
            : t('index.toolbar.layout.profileShow');
        btn.title = profileTitle;
        btn.setAttribute('aria-label', profileTitle);
    }

    if (root) {
        update(root);
    }

    await updateAuthDependentUI();
}

function getRawIdentityVisibilityState() {
    if (!appSettings) {
        return {
            names: true,
            titles: true,
            departments: true
        };
    }
    return {
        names: appSettings.showNames !== false,
        titles: appSettings.showJobTitles !== false,
        departments: appSettings.showDepartments !== false
    };
}

function getIdentityVisibilityState() {
    const baseState = getRawIdentityVisibilityState();
    const enforcer = window.identityVisibility && typeof window.identityVisibility.enforceIdentityVisibility === 'function'
        ? window.identityVisibility.enforceIdentityVisibility
        : null;

    if (!enforcer) {
        return baseState;
    }

    try {
        const enforcedState = enforcer({ ...baseState });
        if (enforcedState && typeof enforcedState === 'object') {
            return {
                names: enforcedState.names !== false,
                titles: enforcedState.titles !== false,
                departments: enforcedState.departments !== false
            };
        }
    } catch (error) {
        console.warn('identityVisibility enforcement failed:', error);
    }

    return baseState;
}

async function ensureIdentityFieldMinimum({ source, skipUpdateAuth = false } = {}) {
    const rawState = getRawIdentityVisibilityState();
    const enforcedState = getIdentityVisibilityState();
    let changed = false;

    if (enforcedState.names !== rawState.names) {
        await setNameVisibility(enforcedState.names, {
            enforceMinimum: false,
            reason: source || 'enforce',
            skipUpdateAuth: true
        });
        changed = true;
    }

    if (!skipUpdateAuth && changed) {
        await updateAuthDependentUI();
    }

    return changed;
}

async function setNameVisibility(newValue, { enforceMinimum = true, reason = 'user', skipUpdateAuth = false } = {}) {
    await waitForTranslations();
    const btn = document.getElementById('nameToggleBtn');

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.showNames = newValue;

    if (serverShowNames != null && newValue === serverShowNames) {
        clearNamePreference();
    } else {
        storeNamePreference(newValue);
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        const nameHide = t('index.toolbar.layout.nameHide', { defaultValue: 'Hide Names' });
        const nameShow = t('index.toolbar.layout.nameShow', { defaultValue: 'Show Names' });
        const label = newValue ? nameHide : nameShow;
        btn.title = label;
        btn.setAttribute('aria-label', label);
    }

    if (root) {
        update(root);
    }

    refreshSearchResultsPresentation();
    refreshEmployeeDetailPanel();

    if (enforceMinimum) {
        await ensureIdentityFieldMinimum({ source: reason, skipUpdateAuth: true });
    }

    if (!skipUpdateAuth) {
        await updateAuthDependentUI();
    }
}

async function toggleNameVisibility() {
    const currentValue = getEffectiveNamesEnabled();
    await setNameVisibility(!currentValue, { reason: 'toggle' });
}

async function toggleDepartmentVisibility() {
    await waitForTranslations();
    const btn = document.getElementById('departmentToggleBtn');
    const currentValue = getEffectiveDepartmentsEnabled();
    const newValue = !currentValue;

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.showDepartments = newValue;

    if (serverShowDepartments != null && newValue === serverShowDepartments) {
        clearDepartmentPreference();
    } else {
        storeDepartmentPreference(newValue);
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        const deptHide = t('index.toolbar.layout.departmentHide', { defaultValue: 'Hide Departments' });
        const deptShow = t('index.toolbar.layout.departmentShow', { defaultValue: 'Show Departments' });
        const departmentTitle = newValue ? deptHide : deptShow;
        btn.title = departmentTitle;
        btn.setAttribute('aria-label', departmentTitle);
    }

    if (root) {
        update(root);
    }

    refreshSearchResultsPresentation();
    refreshEmployeeDetailPanel();

    await ensureIdentityFieldMinimum({ source: 'department-toggle', skipUpdateAuth: true });

    await updateAuthDependentUI();
}

async function toggleOfficeVisibility() {
    await waitForTranslations();
    const btn = document.getElementById('officeToggleBtn');
    const currentValue = getEffectiveOfficeEnabled();
    const newValue = !currentValue;

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.showOffice = newValue;

    if (serverShowOffice != null && newValue === serverShowOffice) {
        clearOfficePreference();
    } else {
        storeOfficePreference(newValue);
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        const officeHide = t('index.toolbar.layout.officeHide', { defaultValue: 'Hide office locations' });
        const officeShow = t('index.toolbar.layout.officeShow', { defaultValue: 'Show office locations' });
        const officeTitle = newValue ? officeHide : officeShow;
        btn.title = officeTitle;
        btn.setAttribute('aria-label', officeTitle);
    }

    if (root) {
        update(root);
    }

    refreshSearchResultsPresentation();
    refreshEmployeeDetailPanel();

    await updateAuthDependentUI();
}

async function toggleJobTitleVisibility() {
    await waitForTranslations();
    const btn = document.getElementById('jobTitleToggleBtn');
    const currentValue = getEffectiveJobTitlesEnabled();
    const newValue = !currentValue;

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.showJobTitles = newValue;

    if (serverShowJobTitles != null && newValue === serverShowJobTitles) {
        clearJobTitlePreference();
    } else {
        storeJobTitlePreference(newValue);
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        const titleHide = t('index.toolbar.layout.jobTitleHide', { defaultValue: 'Hide Job Titles' });
        const titleShow = t('index.toolbar.layout.jobTitleShow', { defaultValue: 'Show Job Titles' });
        const jobTitle = newValue ? titleHide : titleShow;
        btn.title = jobTitle;
        btn.setAttribute('aria-label', jobTitle);
    }

    if (root) {
        update(root);
    }

    refreshSearchResultsPresentation();
    refreshEmployeeDetailPanel();

    await ensureIdentityFieldMinimum({ source: 'job-title-toggle', skipUpdateAuth: true });

    await updateAuthDependentUI();
}

function preloadEmployeeImages(employees) {
    // Preload employee images to improve loading performance
    if (appSettings.showProfileImages !== false) {
        employees.forEach(employee => {
            if (employee.photoUrl && employee.photoUrl.includes('/api/photo/')) {
                const img = new Image();
                img.onload = () => {
                    console.log(`Preloaded photo for ${employee.name}`);
                };
                img.onerror = () => {
                    console.log(`No photo available for ${employee.name} - will use default icon`);
                };
                // Load the photo URL without cache-busting for preload
                img.src = employee.photoUrl;
            }
        });
    }
}

function flattenTree(node, list = []) {
    if (!node) return list;
    list.push(node);
    if (node.id) {
        employeeById.set(node.id, node);
    }
    if (node.children && Array.isArray(node.children)) {
        node.children.forEach(child => flattenTree(child, list));
    }
    return list;
}

// Initialize top-level user search functionality
function initializeTopUserSearch() {
    const searchInput = document.getElementById('topUserSearch');
    const resultsContainer = document.getElementById('topUserResults');
    
    if (!searchInput || !resultsContainer) return;
    
    let selectedUser = null;
    
    // Set initial value if there's a configured top user
    if (appSettings.topUserEmail) {
        const currentUser = allEmployees.find(emp => emp.email === appSettings.topUserEmail);
        if (currentUser) {
            const displayName = getVisibleNameText(currentUser, { includeFallback: true });
            searchInput.value = displayName;
            selectedUser = currentUser;
        }
    }
    
    // Search functionality
    searchInput.addEventListener('input', function() {
        const query = this.value.trim().toLowerCase();
        
        if (query.length < 2) {
            resultsContainer.classList.remove('active');
            selectedUser = null;
            return;
        }
        
        const matches = allEmployees.filter(employee => {
            if (!employee.name || !employee.email) return false;
            
            const name = employee.name.toLowerCase();
            const title = (employee.title || '').toLowerCase();
            const department = (employee.department || '').toLowerCase();
            
            return name.includes(query) || title.includes(query) || department.includes(query);
        }).slice(0, 10); // Limit to 10 results
        
        displayTopUserResults(matches, resultsContainer, searchInput);
    });
    
    // Handle clicking outside to close results
    document.addEventListener('click', function(e) {
        if (!searchInput.contains(e.target) && !resultsContainer.contains(e.target)) {
            resultsContainer.classList.remove('active');
        }
    });
    
    // Handle escape key
    searchInput.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            resultsContainer.classList.remove('active');
        }
    });
    
    // Store selected user reference
    searchInput._selectedUser = selectedUser;
}

function displayTopUserResults(employees, container, input) {
    container.innerHTML = '';
    
    if (employees.length === 0) {
        container.classList.remove('active');
        return;
    }
    
    employees.forEach(employee => {
        const item = document.createElement('div');
        item.className = 'search-result-item';
        item.dataset.employeeId = employee.id || '';
        item.dataset.name = employee.name || '';
        item.dataset.title = employee.title || '';
        item.dataset.department = employee.department || '';
        item.dataset.location = employee.location || employee.officeLocation || '';

        const nameDiv = document.createElement('div');
        nameDiv.className = 'search-result-name';
        const nameText = getVisibleNameText(employee, { includeFallback: true });
        nameDiv.textContent = nameText;
        nameDiv.hidden = !nameText;

        const metaDiv = document.createElement('div');
        metaDiv.className = 'search-result-title';
        populateResultMeta(metaDiv, employee);

        item.appendChild(nameDiv);
        item.appendChild(metaDiv);
        
        item.addEventListener('click', function() {
            const displayName = getVisibleNameText(employee, { includeFallback: true });
            input.value = displayName;
            input._selectedUser = employee;
            container.classList.remove('active');
        });
        
        container.appendChild(item);
    });
    
    container.classList.add('active');
}

function populateResultMeta(element, data, { includeTitleFallback = true, includeDepartmentFallback = false } = {}) {
    if (!element) return;
    const titleText = getVisibleJobTitleText(data, { includeFallback: includeTitleFallback });
    const departmentText = getDepartmentDisplayText(data, { includeFallback: includeDepartmentFallback });
    const segments = [];
    if (titleText) segments.push(titleText);
    if (departmentText) segments.push(departmentText);
    const employeeId = data && (data.id || data.employeeId);
    if (employeeId && isTitleOverridden(employeeId) && segments.length) {
        const editedLabel = t('index.employee.titleEditedBadge');
        segments[0] = `${segments[0]} • ${editedLabel}`;
    }
    if (segments.length) {
        element.textContent = segments.length === 2 ? `${segments[0]} – ${segments[1]}` : segments[0];
        element.hidden = false;
    } else {
        element.textContent = '';
        element.hidden = true;
    }
}

// Save the selected top-level user
async function saveTopUser() {
    await waitForTranslations();
    const searchInput = document.getElementById('topUserSearch');
    const saveBtn = document.getElementById('saveTopUserBtn');
    
    if (!searchInput) return;

    const selectedUser = searchInput._selectedUser;
    const inputValue = searchInput.value.trim();
    
    // If input is empty, save as auto-detect
    const emailToSave = inputValue === '' ? '' : (selectedUser ? selectedUser.email : '');
    
    // Debug logging
    console.log('SaveTopUser Debug:');
    console.log('- inputValue:', inputValue);
    console.log('- selectedUser:', selectedUser);
    console.log('- emailToSave:', emailToSave);
    
    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.textContent = t('index.topUser.saving');
        }
        
        // Update the setting using the public endpoint
        const response = await fetch(`${API_BASE_URL}/api/set-top-user`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                topUserEmail: emailToSave
            })
        });
        
        if (response.ok) {
            let payload = {};
            try {
                payload = await response.json();
            } catch (parseError) {
                payload = {};
            }

            const resolvedEmail = (payload && typeof payload.topUserEmail === 'string') ? payload.topUserEmail : emailToSave;

            // Update app settings
            appSettings.topUserEmail = resolvedEmail;
            
            // Show success feedback and refresh chart data
            if (saveBtn) {
                saveBtn.textContent = t('index.topUser.saved');
            }

            try {
                await reloadEmployeeData();
            } catch (refreshError) {
                console.error('Failed to reload employee data after saving top user:', refreshError);
            }

            if (saveBtn) {
                saveBtn.disabled = false;
                setTimeout(() => {
                    saveBtn.textContent = t('buttons.save');
                }, 1500);
            }
        } else {
            throw new Error('Failed to save setting');
        }
    } catch (error) {
        console.error('Error saving top user:', error);
        if (saveBtn) {
            saveBtn.textContent = t('index.topUser.error');
            setTimeout(() => {
                saveBtn.textContent = t('buttons.save');
                saveBtn.disabled = false;
            }, 2000);
        }
    }
}

// Reset top-level user to auto-detect
async function resetTopUser() {
    await waitForTranslations();
    const searchInput = document.getElementById('topUserSearch');
    const resultsContainer = document.getElementById('topUserResults');
    const resetBtn = document.getElementById('resetTopUserBtn');
    
    try {
        if (resetBtn) {
            resetBtn.disabled = true;
            resetBtn.textContent = t('index.topUser.resetting');
        }
        
        // Update the setting to empty string (auto-detect) using the public endpoint
        const response = await fetch(`${API_BASE_URL}/api/set-top-user`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                topUserEmail: ''
            })
        });
        
        if (response.ok) {
            let payload = {};
            try {
                payload = await response.json();
            } catch (parseError) {
                payload = {};
            }

            const resolvedEmail = (payload && typeof payload.topUserEmail === 'string') ? payload.topUserEmail : '';

            // Update app settings
            appSettings.topUserEmail = resolvedEmail;
            
            // Clear the search input
            searchInput.value = '';
            searchInput._selectedUser = null;
            resultsContainer.classList.remove('active');
            
            if (resetBtn) {
                resetBtn.textContent = t('index.topUser.resetDone');
            }

            try {
                await reloadEmployeeData();
            } catch (refreshError) {
                console.error('Failed to reload employee data after resetting top user:', refreshError);
            }

            if (resetBtn) {
                resetBtn.disabled = false;
                setTimeout(() => {
                    resetBtn.textContent = t('buttons.reset');
                }, 1500);
            }
        } else {
            throw new Error('Failed to reset setting');
        }
    } catch (error) {
        console.error('Error resetting top user:', error);
        if (resetBtn) {
            resetBtn.textContent = t('index.topUser.error');
            setTimeout(() => {
                resetBtn.textContent = t('buttons.reset');
                resetBtn.disabled = false;
            }, 2000);
        }
    }
}

// Reload employee data and re-render chart
async function reloadEmployeeData() {
    await waitForTranslations();
    try {
        // Show loading state
        const container = document.getElementById('orgChart');
        if (container) {
            const loading = container.querySelector('.loading');
            if (loading) {
                loading.style.display = '';
                const message = loading.querySelector('p');
                if (message) {
                    message.textContent = t('index.status.updating');
                }
                const spinner = loading.querySelector('.spinner');
                if (spinner) {
                    spinner.style.display = '';
                } else {
                    loading.insertAdjacentHTML('afterbegin', '<div class="spinner"></div>');
                }
            }

            const existingSvg = container.querySelector('svg');
            if (existingSvg) {
                existingSvg.remove();
            }
        }
        
        const response = await fetch(`${API_BASE_URL}/api/employees`);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        currentData = await response.json();
        window.currentOrgData = currentData; // Store globally for manager lookup
        
        if (currentData) {
            employeeById.clear();
            allEmployees = flattenTree(currentData);
            const validIds = allEmployees.map(emp => emp.id).filter(Boolean);
            pruneTitleOverrides(validIds);
            pruneDepartmentOverrides(validIds);
            preloadEmployeeImages(allEmployees);
            renderOrgChart(currentData);
        } else {
            throw new Error('No data received from server');
        }
    } catch (error) {
        console.error('Error reloading employee data:', error);
        const container = document.getElementById('orgChart');
        if (container) {
            const loading = container.querySelector('.loading');
            if (loading) {
                const spinner = loading.querySelector('.spinner');
                if (spinner) {
                    spinner.style.display = 'none';
                }
                const message = loading.querySelector('p');
                if (message) {
                    message.textContent = t('index.status.errorLoading');
                } else {
                    loading.textContent = t('index.status.errorLoading');
                }
                loading.style.display = '';
            }
        }
    }
}

function renderOrgChart(data) {
    if (!data) {
        console.error('No data to render');
        return;
    }

    const container = document.getElementById('orgChart');
    container.querySelector('.loading').style.display = 'none';

    const width = container.clientWidth;
    const height = container.clientHeight || 800;

    svg = d3.select('#orgChart')
        .append('svg')
        .attr('width', width)
        .attr('height', height);

    updateSvgSize();

    zoom = d3.zoom()
        .scaleExtent([0.1, 3])
        .on('zoom', (event) => {
            g.attr('transform', event.transform);
            if (!programmaticZoomActive && event.sourceEvent) {
                userAdjustedZoom = true;
            }
        });

    svg.call(zoom);

    g = svg.append('g');
    linkLayer = g.append('g').attr('class', 'links');
    nodeLayer = g.append('g').attr('class', 'nodes');

    const initialTransform = d3.zoomIdentity.translate(width/2, 100);
    applyZoomTransform(initialTransform, { duration: 0, resetUser: true });

    root = d3.hierarchy(data);

    root.x0 = 0;
    root.y0 = 0;

    const treeLayout = createTreeLayout();

    const collapseLevel = appSettings.collapseLevel || '2';
    if (collapseLevel !== 'all') {
        const level = parseInt(collapseLevel);
        root.each(d => {
            if (d.depth >= level - 1 && d.children) {
                d._children = d.children;
                d.children = null;
            }
        });
    }

    update(root);
    fitToScreen({ duration: 0 });
}

function update(source) {
    const treeLayout = createTreeLayout();

    const treeData = treeLayout(root);
    const nodes = treeData.descendants();
    const links = treeData.links();
    const highlightEnabled = !!appSettings.highlightNewEmployees;

    // Swap x and y coordinates for horizontal layout
    if (currentLayout === 'horizontal') {
        nodes.forEach(d => {
            const temp = d.x;
            d.x = d.y;
            d.y = temp;
        });
    }

    // Apply multi-line wrap for large children groups (client-side layout tweak)
    applyMultiLineChildrenLayout(nodes);

    // Identify multi-line parents to render bus-style connectors (include root)
    const enabled = appSettings.multiLineChildrenEnabled !== false;
    const threshold = appSettings.multiLineChildrenThreshold || 20;
    const mlParents = enabled
        ? nodes.filter(p => (p.children || []).length >= threshold)
        : [];
    const excludedTargets = new Set();
    mlParents.forEach(p => (p.children || []).forEach(c => excludedTargets.add(c.data.id)));

    const stdLinks = links.filter(d => !excludedTargets.has(d.target.data.id));

    const link = linkLayer.selectAll('.std-link')
        .data(stdLinks, d => d.target.data.id);

    const linkEnter = link.enter()
        .append('path')
        .attr('class', 'link std-link')
        .attr('d', d => {
            const o = {x: source.x0 || source.x, y: source.y0 || source.y};
            return diagonal(o, o);
        });

    link.merge(linkEnter)
        .transition()
        .duration(500)
        .attr('d', d => diagonal(d.source, d.target));

    link.exit()
        .transition()
        .duration(500)
        .attr('d', d => {
            const o = {x: source.x, y: source.y};
            return diagonal(o, o);
        })
        .remove();

    // Render bus-style connectors for multi-line parents
    function buildBusPath(parent) {
        const children = (parent.children || []).slice().sort((a, b) => a.x - b.x);
        if (!children.length) return '';
        const rowsMap = new Map();
        children.forEach(ch => {
            const key = currentLayout === 'vertical' ? Math.round(ch.y) : Math.round(ch.x);
            if (!rowsMap.has(key)) rowsMap.set(key, []);
            rowsMap.get(key).push(ch);
        });
        const rows = Array.from(rowsMap.entries()).sort((a, b) => a[0] - b[0]).map(e => e[1]);
        let d = '';
        if (currentLayout === 'vertical') {
            const spineYs = rows.map(r => Math.min(...r.map(ch => ch.y - nodeHeight / 2)) - 12);
            const topSpineY = Math.min(...spineYs);
            const bottomSpineY = Math.max(...spineYs);
            d += `M ${parent.x} ${parent.y + nodeHeight/2} L ${parent.x} ${bottomSpineY}`;
            rows.forEach((row, i) => {
                const spineY = spineYs[i];
                const xs = row.map(ch => ch.x);
                const left = Math.min(...xs);
                const right = Math.max(...xs);
                d += ` M ${left} ${spineY} L ${right} ${spineY}`;
                row.forEach(ch => {
                    const childTop = ch.y - nodeHeight/2;
                    d += ` M ${ch.x} ${spineY} L ${ch.x} ${childTop}`;
                });
            });
        } else {
            const spineXs = rows.map(r => Math.min(...r.map(ch => ch.x - nodeWidth / 2)) - 12);
            const leftMostSpineX = Math.min(...spineXs);
            const rightMostSpineX = Math.max(...spineXs);
            d += `M ${parent.x + nodeWidth/2} ${parent.y} L ${rightMostSpineX} ${parent.y}`;
            rows.forEach((row, i) => {
                const spineX = spineXs[i];
                const ys = row.map(ch => ch.y);
                const top = Math.min(...ys);
                const bottom = Math.max(...ys);
                d += ` M ${spineX} ${top} L ${spineX} ${bottom}`;
                row.forEach(ch => {
                    const childLeft = ch.x - nodeWidth/2;
                    d += ` M ${spineX} ${ch.y} L ${childLeft} ${ch.y}`;
                });
            });
        }
        return d;
    }

    const bus = linkLayer.selectAll('path.bus-group')
        .data(mlParents, d => d.data.id);
    bus.enter()
        .append('path')
        .attr('class', 'link bus-group')
        .merge(bus)
        .transition()
        .duration(500)
        .attr('d', d => buildBusPath(d));
    bus.exit().remove();

    const node = nodeLayer.selectAll('.node')
        .data(nodes, d => d.data.id);

    const nodeEnter = node.enter()
        .append('g')
        .attr('class', d => {
            let cls = d.depth === 0 ? 'node ceo' : 'node';
            if (isHiddenNode(d)) cls += ' hidden-subtree';
            return cls;
        })
        .attr('transform', d => `translate(${source.x0 || source.x}, ${source.y0 || source.y})`)
        .on('click', (event, d) => {
            event.stopPropagation();
            showEmployeeDetail(d.data);
        });

    nodeEnter.append('rect')
        .attr('class', d => {
            let classes = 'node-rect';
            if (highlightEnabled && d.data.isNewEmployee) {
                classes += ' new-employee';
            }
            return classes;
        })
        .attr('x', -nodeWidth/2)
        .attr('y', -nodeHeight/2)
        .attr('width', nodeWidth)
        .attr('height', nodeHeight)
        .style('fill', d => {
            const nodeColors = appSettings.nodeColors || {};
            switch(d.depth) {
                case 0: return nodeColors.level0 || '#90EE90';
                case 1: return nodeColors.level1 || '#FFFFE0';
                case 2: return nodeColors.level2 || '#E0F2FF';
                case 3: return nodeColors.level3 || '#FFE4E1';
                case 4: return nodeColors.level4 || '#E8DFF5';
                case 5: return nodeColors.level5 || '#FFEAA7';
                case 6: return nodeColors.level6 || '#FAD7FF';
                case 7: return nodeColors.level7 || '#D7F8FF';
                default: return '#F0F0F0'; 
            }
        })
        .style('stroke', d => {
            if (highlightEnabled && d.data.isNewEmployee) {
                return null;
            }
            const nodeColors = appSettings.nodeColors || {};
            let fillColor;
            switch(d.depth) {
                case 0: fillColor = nodeColors.level0 || '#90EE90'; break;
                case 1: fillColor = nodeColors.level1 || '#FFFFE0'; break;
                case 2: fillColor = nodeColors.level2 || '#E0F2FF'; break;
                case 3: fillColor = nodeColors.level3 || '#FFE4E1'; break;
                case 4: fillColor = nodeColors.level4 || '#E8DFF5'; break;
                case 5: fillColor = nodeColors.level5 || '#FFEAA7'; break;
                case 6: fillColor = nodeColors.level6 || '#FAD7FF'; break;
                case 7: fillColor = nodeColors.level7 || '#D7F8FF'; break;
                default: fillColor = '#F0F0F0';
            }
            return adjustColor(fillColor, -50);
        })
        .style('stroke-width', '2px');

    if (appSettings.showProfileImages !== false) {
        applyProfileImageAttributes(nodeEnter.append('image'));
    }

    const namesInitiallyVisible = isNameVisible();
    const titlesInitiallyVisible = isJobTitleVisible();
    const departmentsInitiallyVisible = isDepartmentVisible();
    const initialTitleY = namesInitiallyVisible ? 3 : -12;
    const initialDepartmentY = namesInitiallyVisible
        ? (titlesInitiallyVisible ? 21 : 6)
        : (titlesInitiallyVisible ? 6 : -10);

    nodeEnter.append('text')
        .attr('class', 'node-text')
        .attr('x', getLabelOffsetX())
        .attr('y', -10)
        .attr('text-anchor', getLabelAnchor())
        .style('font-weight', 'bold')
        .style('font-size', d => getNameFontSizePx(d.data.name))
        .text(d => d.data.name)
        .style('display', namesInitiallyVisible ? null : 'none');

    if (titlesInitiallyVisible) {
        nodeEnter.append('text')
            .attr('class', 'node-title')
            .attr('x', getLabelOffsetX())
            .attr('y', initialTitleY)
            .attr('text-anchor', getLabelAnchor())
            .style('font-size', d => getTitleFontSizePx(getVisibleJobTitleText(d.data, { includeFallback: true })))
            .text(d => {
                const title = getVisibleJobTitleText(d.data, { includeFallback: true });
                return title ? getTrimmedTitle(title) : '';
            });
    }

    if (departmentsInitiallyVisible) {
        nodeEnter.append('text')
            .attr('class', 'node-department')
            .attr('x', getLabelOffsetX())
            .attr('y', initialDepartmentY)
            .attr('text-anchor', getLabelAnchor())
            .style('font-size', d => getDepartmentFontSizePx(getDepartmentDisplayText(d.data, { includeFallback: true, fallback: 'Not specified' })))
            .style('font-style', 'italic')
            .style('fill', '#666')
            .text(d => getDepartmentDisplayText(d.data, { includeFallback: true, fallback: 'Not specified' }));
    }

    const countGroup = nodeEnter.append('g')
        .attr('class', 'count-badge')
        .style('display', d => shouldShowCountBadge(d) ? 'block' : 'none');

    countGroup.append('circle')
        .attr('cx', -nodeWidth/2 + 15)
        .attr('cy', -nodeHeight/2 + 15)
        .attr('r', 12)
        .style('fill', '#ff6b6b')
        .style('stroke', 'white')
        .style('stroke-width', '2px');

    countGroup.append('text')
        .attr('x', -nodeWidth/2 + 15)
        .attr('y', -nodeHeight/2 + 19)
        .attr('text-anchor', 'middle')
        .style('fill', 'white')
        .style('font-size', '11px')
        .style('font-weight', 'bold')
        .text(d => formatDirectReportCount(d));

    const expandBtn = nodeEnter.append('g')
        .attr('class', 'expand-group')
        .style('display', d => (d._children?.length || d.children?.length) ? 'block' : 'none')
        .on('click', (event, d) => {
            event.stopPropagation();
            toggle(d);
        });

    expandBtn.append('circle')
        .attr('class', 'expand-btn')
        .attr('cy', currentLayout === 'vertical' ? nodeHeight/2 + 10 : 0)
        .attr('cx', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0)
        .attr('r', 10);

    expandBtn.append('text')
        .attr('class', 'expand-text')
        .attr('y', currentLayout === 'vertical' ? nodeHeight/2 + 15 : 4)
        .attr('x', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0)
        .attr('text-anchor', 'middle')
        .text(d => d._children?.length ? '+' : '-');

    // Eye icon toggle (placed top-right inside node)
    nodeEnter.append('text')
        .attr('class', 'hide-toggle')
        .attr('x', nodeWidth/2 - 14)
        .attr('y', -nodeHeight/2 + 14)
        .attr('text-anchor', 'middle')
        .text(d => hiddenNodeIds.has(d.data.id) ? '🙈' : '👁')
        .on('click', (event, d) => {
            event.stopPropagation();
            toggleHideNode(d);
        })
        .append('title')
        .text(d => hiddenNodeIds.has(d.data.id) ? t('index.tree.toggleShow') : t('index.tree.toggleHide'));

    const newBadgeGroup = nodeEnter.append('g')
        .attr('class', 'new-employee-badge')
        .style('display', d => (highlightEnabled && d.data.isNewEmployee) ? 'block' : 'none');

    newBadgeGroup.append('rect')
        .attr('class', 'new-badge')
        .attr('x', nodeWidth/2 - 45)
        .attr('y', -nodeHeight/2 - 10)
        .attr('width', 35)
        .attr('height', 18)
        .attr('rx', 9)
        .attr('ry', 9);

    newBadgeGroup.append('text')
        .attr('class', 'new-badge-text')
        .attr('x', nodeWidth/2 - 27)
        .attr('y', -nodeHeight/2 + 2)
        .attr('text-anchor', 'middle')
        .text(t('index.badges.new'));


    const nodeMerge = node.merge(nodeEnter);

    nodeMerge.selectAll('rect.node-rect')
        .classed('new-employee', d => highlightEnabled && d.data.isNewEmployee)
        .style('stroke', d => {
            if (highlightEnabled && d.data.isNewEmployee) {
                return null;
            }
            const nodeColors = appSettings.nodeColors || {};
            let fillColor;
            switch (d.depth) {
                case 0: fillColor = nodeColors.level0 || '#90EE90'; break;
                case 1: fillColor = nodeColors.level1 || '#FFFFE0'; break;
                case 2: fillColor = nodeColors.level2 || '#E0F2FF'; break;
                case 3: fillColor = nodeColors.level3 || '#FFE4E1'; break;
                case 4: fillColor = nodeColors.level4 || '#E8DFF5'; break;
                case 5: fillColor = nodeColors.level5 || '#FFEAA7'; break;
                case 6: fillColor = nodeColors.level6 || '#FAD7FF'; break;
                case 7: fillColor = nodeColors.level7 || '#D7F8FF'; break;
                default: fillColor = '#F0F0F0';
            }
            return adjustColor(fillColor, -50);
        });

    nodeMerge.selectAll('.new-employee-badge')
        .style('display', d => (highlightEnabled && d.data.isNewEmployee) ? 'block' : 'none');

    const nodeUpdate = nodeMerge
        .attr('class', d => {
            let cls = d.depth === 0 ? 'node ceo' : 'node';
            if (isHiddenNode(d)) cls += ' hidden-subtree';
            return cls;
        })
        .transition()
        .duration(500)
        .attr('transform', d => `translate(${d.x}, ${d.y})`);

    // Update eye icons and tooltips on merged selection (after transition start)
    nodeMerge.selectAll('text.hide-toggle')
        .text(d => hiddenNodeIds.has(d.data.id) ? '🙈' : '👁')
        .each(function(d){
            const titleEl = this.querySelector('title');
            if (titleEl) titleEl.textContent = hiddenNodeIds.has(d.data.id)
                ? t('index.tree.toggleShow')
                : t('index.tree.toggleHide');
        });

    nodeUpdate.select('.expand-text')
        .text(d => d._children?.length ? '+' : '-')
        .attr('y', currentLayout === 'vertical' ? nodeHeight/2 + 15 : 4)
        .attr('x', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0);

    nodeUpdate.select('.expand-btn')
        .attr('cy', currentLayout === 'vertical' ? nodeHeight/2 + 10 : 0)
        .attr('cx', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0);

    nodeUpdate.select('.expand-group')
        .style('display', d => (d._children?.length || d.children?.length) ? 'block' : 'none');

    nodeMerge.selectAll('.count-badge')
        .style('display', d => shouldShowCountBadge(d) ? 'block' : 'none');

    nodeMerge.selectAll('.count-badge text')
        .text(d => formatDirectReportCount(d));

    if (appSettings.showProfileImages !== false) {
        nodeMerge.each(function(d) {
            const nodeSel = d3.select(this);
            let img = nodeSel.select('image.profile-image');
            if (img.empty()) {
                img = nodeSel.insert('image', 'text');
            }
            applyProfileImageAttributes(img);
        });
    } else {
        nodeMerge.selectAll('image.profile-image').remove();
    }

    const namesVisible = isNameVisible();
    const titlesVisible = isJobTitleVisible();
    const departmentsVisible = isDepartmentVisible();
    const titleY = namesVisible ? 3 : -12;
    const departmentY = namesVisible
        ? (titlesVisible ? 21 : 6)
        : (titlesVisible ? 6 : -10);

    nodeMerge.select('.node-text')
        .attr('x', getLabelOffsetX())
        .attr('y', -10)
        .attr('text-anchor', getLabelAnchor())
        .style('font-size', d => getNameFontSizePx(d.data.name))
        .text(d => d.data.name)
        .style('display', namesVisible ? null : 'none');

    nodeMerge.each(function(d) {
        const nodeSelection = d3.select(this);

        let titleSelection = nodeSelection.select('text.node-title');
        if (titlesVisible) {
            if (titleSelection.empty()) {
                titleSelection = nodeSelection.append('text')
                    .attr('class', 'node-title');
            }
            const rawTitle = getVisibleJobTitleText(d.data, { includeFallback: true });
            const displayTitle = rawTitle ? getTrimmedTitle(rawTitle) : '';
            titleSelection
                .attr('x', getLabelOffsetX())
                .attr('y', titleY)
                .attr('text-anchor', getLabelAnchor())
                .style('font-size', getTitleFontSizePx(rawTitle || ''))
                .text(displayTitle)
                .style('display', displayTitle ? null : 'none')
                .classed('node-title--edited', isTitleOverridden(d.data.id));
        } else if (!titleSelection.empty()) {
            titleSelection.remove();
        }

        let departmentSelection = nodeSelection.select('text.node-department');
        if (departmentsVisible) {
            if (departmentSelection.empty()) {
                departmentSelection = nodeSelection.append('text')
                    .attr('class', 'node-department')
                    .style('font-style', 'italic')
                    .style('fill', '#666');
            }
            const departmentText = getDepartmentDisplayText(d.data, { includeFallback: true, fallback: 'Not specified' });
            departmentSelection
                .attr('x', getLabelOffsetX())
                .attr('y', departmentY)
                .attr('text-anchor', getLabelAnchor())
                .style('font-size', getDepartmentFontSizePx(departmentText || 'Not specified'))
                .text(departmentText)
                .style('display', departmentText ? null : 'none');
        } else if (!departmentSelection.empty()) {
            departmentSelection.remove();
        }
    });

    node.exit()
        .transition()
        .duration(500)
        .attr('transform', d => `translate(${source.x}, ${source.y})`)
        .remove();

    nodes.forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
    });
}

// Arrange many direct reports in multiple rows to avoid overlap
function applyMultiLineChildrenLayout(nodes) {
    const enabled = appSettings.multiLineChildrenEnabled !== false;
    const threshold = appSettings.multiLineChildrenThreshold || 20;
    if (!enabled) return;

    // Helper: shift an entire subtree rooted at node by dx, dy
    function shiftSubtree(rootNode, dx, dy) {
        if ((dx === 0 && dy === 0) || !rootNode) return;
        nodes.forEach(n => {
            let cur = n;
            while (cur) {
                if (cur === rootNode) {
                    n.x += dx;
                    n.y += dy;
                    break;
                }
                cur = cur.parent;
            }
        });
    }

    // Helper: check if anc is an ancestor of node
    function isAncestor(anc, node) {
        let cur = node;
        while (cur) {
            if (cur === anc) return true;
            cur = cur.parent;
        }
        return false;
    }

    // Helper: compute subtree bounds for a node across provided nodes
    function getSubtreeBounds(rootNode) {
        let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
        nodes.forEach(n => {
            if (isAncestor(rootNode, n)) {
                const left = n.x - nodeWidth / 2;
                const right = n.x + nodeWidth / 2;
                const top = n.y - nodeHeight / 2;
                const bottom = n.y + nodeHeight / 2;
                if (left < minX) minX = left;
                if (right > maxX) maxX = right;
                if (top < minY) minY = top;
                if (bottom > maxY) maxY = bottom;
            }
        });
        return { minX, maxX, minY, maxY };
    }


    nodes.forEach(parent => {
        const kids = parent.children || [];
        if (!kids.length) return;
        if (kids.length < threshold) return;

        // Determine row/column layout and minimize empty slots in last row
        // Preserve D3's left-to-right ordering to reduce crossing
        const orderedKids = kids.slice().sort((a, b) => a.x - b.x);
        const n = orderedKids.length;
        let columns = Math.ceil(Math.sqrt(n));
        let rows = Math.ceil(n / columns);
        columns = Math.ceil(n / rows); // refine to reduce underfill
        const hSpacing = currentLayout === 'horizontal' ? HORIZONTAL_MULTILINE_SPACING : VERTICAL_MULTILINE_SPACING;
        const vSpacing = currentLayout === 'horizontal' ? HORIZONTAL_LEVEL_HEIGHT : levelHeight; // keep consistent per-level spacing
        const totalHeight = (rows - 1) * vSpacing;
        const horizontalBaseOffset = currentLayout === 'horizontal'
            ? (appSettings.multiLineHorizontalBaseOffset ?? Math.max(nodeWidth * 0.5, HORIZONTAL_LEVEL_HEIGHT * 0.45))
            : 0;
        const horizontalRowOffset = currentLayout === 'horizontal'
            ? (appSettings.multiLineHorizontalRowOffset ?? Math.min(HORIZONTAL_LEVEL_HEIGHT * 0.10, nodeWidth * 0.15))
            : 0;

        orderedKids.forEach((child, idx) => {
            const col = idx % columns;
            const row = Math.floor(idx / columns);

            let targetX, targetY;
            if (currentLayout === 'vertical') {
                // Center based on actual items in this row
                let itemsInRow = (row < rows - 1) ? columns : (n - (rows - 1) * columns);
                if (itemsInRow <= 0) itemsInRow = columns;
                const rowWidth = (itemsInRow - 1) * hSpacing;
                const colInRow = col % itemsInRow;
                targetX = parent.x - rowWidth / 2 + colInRow * hSpacing;
                targetY = parent.y + (row + 1) * vSpacing;
            } else {
                let itemsInRow = (row < rows - 1) ? columns : (n - (rows - 1) * columns);
                if (itemsInRow <= 0) itemsInRow = columns;
                const rowWidth = (itemsInRow - 1) * hSpacing;
                const colInRow = col % itemsInRow;
                const baseStep = vSpacing + horizontalBaseOffset;
                targetX = parent.x + (row + 1) * baseStep + row * horizontalRowOffset;
                targetY = parent.y - rowWidth / 2 + colInRow * hSpacing;
            }

            const dx = targetX - child.x;
            const dy = targetY - child.y;
            shiftSubtree(child, dx, dy);
        });

        // Keep compaction logic light for stability
    });

    // Pass 2: Compress ancestors around multi-lined groups, do not alter the group itself
    const configuredGap = appSettings.multiLineCompactGap;
    const gapBetweenSubtrees = (typeof configuredGap === 'number')
        ? configuredGap
        : (currentLayout === 'horizontal' ? 36 : 10); // edge-to-edge gap
    // Identify multi-lined parents (where wrapping occurred)
    const mlParents = nodes.filter(p => (p.children || []).length >= threshold);
    const ancestorSet = new Set();
    mlParents.forEach(mlp => {
        let anc = mlp.parent;
        while (anc) {
            ancestorSet.add(anc);
            anc = anc.parent;
        }
    });

    // Process ancestors deep-to-shallow so higher levels can adapt after recentering
    const ancestors = Array.from(ancestorSet).sort((a, b) => b.depth - a.depth);
    ancestors.forEach(parent => {
        const kids = parent.children || [];
        if (kids.length < 2) return;
        // Build intervals for each child's subtree as fixed blocks
        const intervals = kids.map(child => {
            const b = getSubtreeBounds(child);
            if (currentLayout === 'vertical') {
                const width = Math.max(b.maxX - b.minX, nodeWidth);
                const center = (b.maxX + b.minX) / 2;
                return { child, width, center };
            } else {
                const width = Math.max(b.maxY - b.minY, nodeWidth);
                const center = (b.maxY + b.minY) / 2;
                return { child, width, center };
            }
        });
        // Preserve order by current center
        intervals.sort((a, b) => a.center - b.center);
        const totalWidth = intervals.reduce((sum, it) => sum + it.width, 0) + gapBetweenSubtrees * (intervals.length - 1);
        const groupCenter = currentLayout === 'vertical' ? parent.x : parent.y;
        const start = groupCenter - totalWidth / 2;
        let cursor = start;
        intervals.forEach(it => {
            const targetCenter = cursor + it.width / 2;
            const delta = targetCenter - it.center;
            if (currentLayout === 'vertical') {
                shiftSubtree(it.child, delta, 0);
            } else {
                shiftSubtree(it.child, 0, delta);
            }
            cursor += it.width + gapBetweenSubtrees;
        });

        // After packing children, recenter parent directly above/between them to avoid one-sided gaps
        const childBounds = kids.map(ch => getSubtreeBounds(ch));
        if (currentLayout === 'vertical') {
            const left = Math.min(...childBounds.map(b => b.minX));
            const right = Math.max(...childBounds.map(b => b.maxX));
            const desired = (left + right) / 2;
            const deltaParent = desired - parent.x;
            if (Math.abs(deltaParent) > 0.1) shiftSubtree(parent, deltaParent, 0);
        } else {
            const top = Math.min(...childBounds.map(b => b.minY));
            const bottom = Math.max(...childBounds.map(b => b.maxY));
            const desired = (top + bottom) / 2;
            const deltaParent = desired - parent.y;
            if (Math.abs(deltaParent) > 0.1) shiftSubtree(parent, 0, deltaParent);
        }
    });
}

function diagonal(s, d) {
    if (currentLayout === 'vertical') {
        const midY = (s.y + d.y) / 2;
        return `M ${s.x} ${s.y + nodeHeight/2}
                L ${s.x} ${midY}
                L ${d.x} ${midY}
                L ${d.x} ${d.y - nodeHeight/2}`;
    } else {
        const midX = (s.x + d.x) / 2;
        return `M ${s.x + nodeWidth/2} ${s.y}
                L ${midX} ${s.y}
                L ${midX} ${d.y}
                L ${d.x - nodeWidth/2} ${d.y}`;
    }
}

function toggle(d) {
    if (d.children) {
        d._children = d.children;
        d.children = null;
    } else {
        d.children = d._children;
        d._children = null;
        
        if (d.children) {
            d.children.forEach(child => {
                if (child.depth >= 2 && child.children) {
                    child._children = child.children;
                    child.children = null;
                }
            });
        }
    }
    update(d);
}

function expandAll() {
    root.each(d => {
        if (d._children) {
            d.children = d._children;
            d._children = null;
        }
    });
    update(root);
}

function collapseAll() {
    root.each(d => {
        if (d.depth >= 1 && d.children) {
            d._children = d.children;
            d.children = null;
        }
    });
    update(root);
}

function resetZoom() {
    fitToScreen({ duration: 500, resetUser: true });
}

function fitToScreen(options = {}) {
    if (!root || !svg) return;
    updateSvgSize();

    const { duration = 750, resetUser = true } = options;
    const treeLayout = createTreeLayout();
    const treeData = treeLayout(root);
    const nodes = treeData.descendants();
    if (currentLayout === 'horizontal') {
        nodes.forEach(d => { const t = d.x; d.x = d.y; d.y = t; });
    }
    applyMultiLineChildrenLayout(nodes);
    
    if (nodes.length === 0) return;
    
    const minX = d3.min(nodes, d => d.x) - nodeWidth / 2;
    const maxX = d3.max(nodes, d => d.x) + nodeWidth / 2;
    const minY = d3.min(nodes, d => d.y) - nodeHeight / 2;
    const maxY = d3.max(nodes, d => d.y) + nodeHeight / 2;
    
    const width = maxX - minX;
    const height = maxY - minY;
    
    const container = document.getElementById('orgChart');
    const containerWidth = Math.max(container.clientWidth, 1);
    const containerHeight = Math.max(container.clientHeight || 0, 1);
    
    const scale = Math.min(
        width === 0 ? 1 : (containerWidth * 0.9) / width,
        height === 0 ? 1 : (containerHeight * 0.9) / height,
        1 
    );
    
    const centerX = (minX + maxX) / 2;
    const centerY = (minY + maxY) / 2;
    
    const transform = d3.zoomIdentity
        .translate(containerWidth / 2, containerHeight / 2)
        .scale(scale)
        .translate(-centerX, -centerY);

    applyZoomTransform(transform, { duration, resetUser });
}

function zoomIn() {
    if (!svg) return;
    const transition = svg.transition().call(zoom.scaleBy, 1.2);
    transition.on('end', () => { userAdjustedZoom = true; });
    transition.on('interrupt', () => { userAdjustedZoom = true; });
}

function zoomOut() {
    if (!svg) return;
    const transition = svg.transition().call(zoom.scaleBy, 0.8);
    transition.on('end', () => { userAdjustedZoom = true; });
    transition.on('interrupt', () => { userAdjustedZoom = true; });
}

function getBounds(printRoot) {
    const treeLayout = createTreeLayout();
    const treeData = treeLayout(printRoot);
    const nodes = treeData.descendants();
    applyMultiLineChildrenLayout(nodes);
    const minX = d3.min(nodes, d => d.x) - nodeWidth / 2 - 20;
    const maxX = d3.max(nodes, d => d.x) + nodeWidth / 2 + 20;
    const minY = d3.min(nodes, d => d.y) - nodeHeight / 2 - 20;
    const maxY = d3.max(nodes, d => d.y) + nodeHeight / 2 + 50;
    return { minX, maxX, minY, maxY };
}

function buildExpandedData(node) {
    const copy = { data: node.data, depth: node.depth };
    const allKids = [];
    if (node.children) allKids.push(...node.children);
    if (node._children) allKids.push(...node._children);
    if (allKids.length) {
        copy.children = allKids.map(child => buildExpandedData(child));
    }
    copy.hasCollapsedChildren = !!(node._children && node._children.length);
    return copy;
}

function printChart() {
    createExportSVG(false).then(svgElement => {
        const printWin = window.open('', '_blank');
        printWin.document.write('<html><head><title>Org Chart Print</title>');
        printWin.document.write('<style>@page { margin: 0.5cm; } body { margin:0; padding:0; }</style>');
        printWin.document.write('</head><body>');
        // Clone so we don't mutate original
        const clone = svgElement.cloneNode(true);
        // Fit to page via CSS width 100%
        clone.removeAttribute('width');
        clone.removeAttribute('height');
        clone.style.width = '100%';
        clone.style.height = 'auto';
        printWin.document.body.appendChild(clone);
        printWin.document.write('</body></html>');
        printWin.document.close();
        printWin.focus();
        printWin.print();
    }).catch(err => console.error('Print failed:', err));
}

async function exportToImage(format = 'svg', exportFullChart = false) {
    await waitForTranslations();
    const svgElement = await createExportSVG(exportFullChart);
    const svgString = new XMLSerializer().serializeToString(svgElement);
    
    if (format === 'svg') {
        // Export as SVG
        const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(svgBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `org-chart-${new Date().toISOString().split('T')[0]}.svg`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } else if (format === 'png') {
        // Convert to PNG using HTML5 Canvas approach
        console.log('Starting PNG export...');
        
        try {
            // Parse the SVG to get dimensions
            const parser = new DOMParser();
            const svgDoc = parser.parseFromString(svgString, 'image/svg+xml');
            const svgElement = svgDoc.documentElement;
            
            // Extract dimensions from SVG
            const svgWidth = parseFloat(svgElement.getAttribute('width')) || 800;
            const svgHeight = parseFloat(svgElement.getAttribute('height')) || 600;
            
            console.log('SVG dimensions extracted:', svgWidth, 'x', svgHeight);
            
            // Create canvas with better scaling
            const scale = window.devicePixelRatio || 2;
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            canvas.width = svgWidth * scale;
            canvas.height = svgHeight * scale;
            
            // Set CSS size for proper scaling
            canvas.style.width = svgWidth + 'px';
            canvas.style.height = svgHeight + 'px';
            
            // Scale context
            ctx.scale(scale, scale);
            
            // White background
            ctx.fillStyle = 'white';
            ctx.fillRect(0, 0, svgWidth, svgHeight);
            
            console.log('Canvas created with dimensions:', canvas.width, 'x', canvas.height);
            
            // Try multiple SVG loading methods for better compatibility
            const tryMethod1 = () => {
                const svgDataUrl = `data:image/svg+xml;base64,${btoa(svgString)}`;
                const img = new Image();
                
                img.onload = function() {
                    console.log('Method 1 - SVG loaded successfully');
                    ctx.drawImage(img, 0, 0, svgWidth, svgHeight);
                    downloadCanvas();
                };
                
                img.onerror = function(error) {
                    console.warn('Method 1 failed, trying method 2:', error);
                    tryMethod2();
                };
                
                img.crossOrigin = 'anonymous';
                img.src = svgDataUrl;
            };
            
            const tryMethod2 = () => {
                const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgString)}`;
                const img = new Image();
                
                img.onload = function() {
                    console.log('Method 2 - SVG loaded successfully');
                    ctx.drawImage(img, 0, 0, svgWidth, svgHeight);
                    downloadCanvas();
                };
                
                img.onerror = function(error) {
                    console.warn('Method 2 failed, trying fallback method:', error);
                    tryFallback();
                };
                
                img.crossOrigin = 'anonymous';
                img.src = svgDataUrl;
            };
            
            const tryFallback = () => {
                console.log('Using fallback: rendering SVG directly to canvas');
                try {
                    // Fallback: Create a simplified version without external resources
                    const simplifiedSvg = svgString
                        .replace(/xlink:href="[^"]*api\/photo[^"]*"/g, 'xlink:href=""')
                        .replace(/<image[^>]*api\/photo[^>]*\/>/g, '');
                    
                    const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(simplifiedSvg)}`;
                    const img = new Image();
                    
                    img.onload = function() {
                        console.log('Fallback - SVG loaded successfully');
                        ctx.drawImage(img, 0, 0, svgWidth, svgHeight);
                        downloadCanvas();
                    };
                    
                    img.onerror = function(error) {
                        console.error('All methods failed:', error);
                        alert(t('index.alerts.pngLoadError'));
                    };
                    
                    img.crossOrigin = 'anonymous';
                    img.src = svgDataUrl;
                } catch (error) {
                    console.error('Fallback method failed:', error);
                    alert(t('index.alerts.pngExportError'));
                }
            };
            
            const downloadCanvas = () => {
                canvas.toBlob(function(blob) {
                    if (blob) {
                        console.log('PNG blob created successfully, size:', blob.size);
                        const pngUrl = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = pngUrl;
                        a.download = `org-chart-${new Date().toISOString().split('T')[0]}.png`;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(pngUrl);
                    } else {
                        console.error('Failed to create PNG blob');
                        alert(t('index.alerts.pngBlobError'));
                    }
                }, 'image/png', 0.95);
            };
            
            // Start with method 1
            tryMethod1();
            
        } catch (error) {
            console.error('Error in PNG export setup:', error);
            alert(t('index.alerts.pngSetupError', { message: error.message }));
        }
    }
}

async function exportToPDF(exportFullChart = false) {
    await waitForTranslations();
    try {
        console.log('Starting PDF export...');
        
        // Check if data is loaded
        if (!currentData) {
            alert(t('index.alerts.pdfNoData'));
            return;
        }
        
        // Check if root is available for visible chart export
        if (!exportFullChart && !root) {
            alert(t('index.alerts.pdfNoChartVisible'));
            return;
        }
        
        if (typeof window.jspdf === 'undefined') {
            console.error('jsPDF library not loaded');
            alert(t('index.alerts.pdfLibraryMissing'));
            return;
        }
        
        console.log('Libraries loaded successfully, currentData available:', !!currentData, 'root available:', !!root);

        // Use the existing SVG creation function and scale it to PDF
        const svgElement = await createExportSVG(exportFullChart);
        console.log('SVG created successfully');
        
        // Get SVG dimensions
        const svgWidth = parseFloat(svgElement.getAttribute('width'));
        const svgHeight = parseFloat(svgElement.getAttribute('height'));
        
        console.log(`SVG dimensions: ${svgWidth} x ${svgHeight}`);
        
        // Create PDF with appropriate orientation
        const { jsPDF } = window.jspdf;
        const isLandscape = svgWidth > svgHeight;
        const pdf = new jsPDF(isLandscape ? 'l' : 'p', 'mm', 'a4');
        
        // Get PDF page dimensions
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 10;
        const availableWidth = pageWidth - (2 * margin);
        const availableHeight = pageHeight - (2 * margin);
        
        // Calculate scale to fit SVG in PDF page
        const scaleX = availableWidth / (svgWidth * 0.264583); // Convert px to mm
        const scaleY = availableHeight / (svgHeight * 0.264583);
        const scale = Math.min(scaleX, scaleY);
        
        // Calculate final dimensions and position
        const finalWidth = svgWidth * 0.264583 * scale;
        const finalHeight = svgHeight * 0.264583 * scale;
        const x = margin + (availableWidth - finalWidth) / 2;
        const y = margin + (availableHeight - finalHeight) / 2;
        
        console.log(`PDF: ${pageWidth}x${pageHeight}mm, Final: ${finalWidth.toFixed(1)}x${finalHeight.toFixed(1)}mm, Scale: ${scale.toFixed(3)}`);
        
        // Convert SVG to data URL
        const svgString = new XMLSerializer().serializeToString(svgElement);
        const svgDataUrl = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgString)));
        
        // Add SVG as image to PDF
        try {
            pdf.addImage(svgDataUrl, 'SVG', x, y, finalWidth, finalHeight);
        } catch (error) {
            console.warn('SVG addImage failed, trying PNG conversion:', error);
            
            // Fallback: convert SVG to canvas first
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const img = new Image();
            
            await new Promise((resolve, reject) => {
                img.onload = () => {
                    canvas.width = svgWidth;
                    canvas.height = svgHeight;
                    ctx.fillStyle = 'white';
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    ctx.drawImage(img, 0, 0);
                    resolve();
                };
                img.onerror = reject;
                img.src = svgDataUrl;
            });
            
            const pngDataUrl = canvas.toDataURL('image/png', 0.9);
            pdf.addImage(pngDataUrl, 'PNG', x, y, finalWidth, finalHeight);
        }

        // Add timestamp in the bottom-right corner
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const timestamp = `${year}-${month}-${day} ${hours}:${minutes}`;
        pdf.setFontSize(8);
        pdf.setTextColor(128, 128, 128);
        const timestampText = `Generated: ${timestamp}`;
        const textWidth = pdf.getTextWidth(timestampText);
        pdf.text(timestampText, pageWidth - margin - textWidth, pageHeight - 5);

        const fileName = `org-chart-${new Date().toISOString().split('T')[0]}.pdf`;
        pdf.save(fileName);
        console.log('PDF exported successfully:', fileName);

    } catch (error) {
        console.error('Error in exportToPDF:', error);
        // Ensure cleanup of any temporary elements
        try {
            const orphanSvg = document.querySelector('svg[style*="position: absolute"]');
            if (orphanSvg) {
                document.body.removeChild(orphanSvg);
            }
            const orphanImg = document.querySelector('img[style*="position: absolute"]');
            if (orphanImg) {
                document.body.removeChild(orphanImg);
            }
        } catch (cleanupError) {
            console.warn('Cleanup error:', cleanupError);
        }
        alert(t('index.alerts.pdfGenericError', { message: error.message || error || t('index.alerts.unknownError') }));
    }
}

async function imageToDataUrl(url) {
    try {
        const response = await fetch(url);
        if (!response.ok) {
            console.warn(`Failed to fetch image: ${url}, status: ${response.status}`);
            return null;
        }
        const blob = await response.blob();
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = (err) => {
                console.error('FileReader error:', err);
                reject(err);
            };
            reader.readAsDataURL(blob);
        });
    } catch (error) {
        console.error('Error converting image to data URL:', url, error);
        return null;
    }
}

function cloneNodeDataForExport(nodeData) {
    if (!nodeData || typeof nodeData !== 'object') {
        return {};
    }
    const clone = { ...nodeData };
    delete clone.children;
    delete clone._children;
    return clone;
}

function collectChildSources(node, includeCollapsed) {
    const sources = [];
    if (node && Array.isArray(node.children)) {
        sources.push(...node.children);
    }
    if (includeCollapsed && node && Array.isArray(node._children)) {
        sources.push(...node._children);
    }
    return sources;
}

function buildExportDataTree(node, includeCollapsed = false) {
    if (!node || !node.data) {
        return null;
    }

    const nodeId = node.data.id;
    if (nodeId != null && hiddenNodeIds.has(nodeId)) {
        return null;
    }

    const clone = cloneNodeDataForExport(node.data);
    const childSources = collectChildSources(node, includeCollapsed);
    const childClones = childSources
        .map(child => buildExportDataTree(child, includeCollapsed))
        .filter(Boolean);

    if (childClones.length) {
        clone.children = childClones;
    }

    return clone;
}

function buildExportHierarchy({ exportFullChart = false } = {}) {
    if (!root) {
        return null;
    }

    const includeCollapsed = exportFullChart;
    const exportData = buildExportDataTree(root, includeCollapsed);
    if (!exportData) {
        return null;
    }

    return d3.hierarchy(exportData);
}

async function createExportSVG(exportFullChart = false) {
    if (!currentData) {
        throw new Error('No organizational chart data available');
    }

    const hierarchyRoot = buildExportHierarchy({ exportFullChart });
    if (!hierarchyRoot) {
        throw new Error('No chart nodes available for export');
    }

    const treeLayout = createTreeLayout();
    const treeData = treeLayout(hierarchyRoot);
    let nodesToExport = treeData.descendants();
    let linksToExport = treeData.links();

    if (!nodesToExport || nodesToExport.length === 0) {
        throw new Error('No chart nodes available for export');
    }

    const visibleIdSet = new Set(nodesToExport.map(n => n.data.id));
    linksToExport = linksToExport.filter(l => visibleIdSet.has(l.source.data.id) && visibleIdSet.has(l.target.data.id));
    
    // Adjust for horizontal layout
    if (currentLayout === 'horizontal') {
        nodesToExport.forEach(d => { [d.x, d.y] = [d.y, d.x]; });
    }

    // Apply our layout adjustments before bounds and export
    applyMultiLineChildrenLayout(nodesToExport);
    // No extra compaction
    
    // Calculate bounds
    const padding = 10;
    const minX = d3.min(nodesToExport, d => d.x) - nodeWidth/2 - padding;
    const maxX = d3.max(nodesToExport, d => d.x) + nodeWidth/2 + padding;
    const minY = d3.min(nodesToExport, d => d.y) - nodeHeight/2 - padding;
    const maxY = d3.max(nodesToExport, d => d.y) + nodeHeight/2 + padding;
    const width = maxX - minX;
    const height = maxY - minY;
    
    const showImages = appSettings.showProfileImages !== false;
    const baseNameFontSize = showImages ? '14px' : '16px';
    const baseTitleFontSize = showImages ? '11px' : '13px';
    const baseDepartmentFontSize = showImages ? '9px' : '11px';

    // Create SVG element
    const exportSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    Object.assign(exportSvg.style, { fontFamily: 'Arial, sans-serif' });
    exportSvg.setAttribute('width', width);
    exportSvg.setAttribute('height', height);
    exportSvg.setAttribute('viewBox', `${minX} ${minY} ${width} ${height}`);
    exportSvg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    exportSvg.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
    
    // Add styles
    const style = document.createElementNS('http://www.w3.org/2000/svg', 'style');
    style.textContent = `
        .link { fill: none; stroke: #999; stroke-width: 2px; }
        .node-rect { rx: 4; ry: 4; }
        .node-text { font-size: ${baseNameFontSize}; fill: #333; font-weight: 600; }
        .node-title { font-size: ${baseTitleFontSize}; fill: #555; }
        .node-department { font-size: ${baseDepartmentFontSize}; fill: #666; font-style: italic; }
    `;
    exportSvg.appendChild(style);

    // Add white background
    const bg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    bg.setAttribute('x', minX);
    bg.setAttribute('y', minY);
    bg.setAttribute('width', width);
    bg.setAttribute('height', height);
    bg.setAttribute('fill', 'white');
    exportSvg.appendChild(bg);
    
    const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    const linksGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    linksGroup.setAttribute('class', 'links');
    const nodesGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    nodesGroup.setAttribute('class', 'nodes');
    
    // Identify multi-line parents for bus connectors
    const enabled = appSettings.multiLineChildrenEnabled !== false;
    const threshold = appSettings.multiLineChildrenThreshold || 20;
    const mlParents = enabled
        ? nodesToExport.filter(p => (p.children || []).length >= threshold)
        : [];
    const excludedTargets = new Set();
    mlParents.forEach(p => (p.children || []).forEach(c => excludedTargets.add(c.data.id)));
    const stdLinks = linksToExport.filter(d => !excludedTargets.has(d.target.data.id));

    // Draw standard links
    stdLinks.forEach(link => {
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('class', 'link');
        path.setAttribute('d', diagonal(link.source, link.target));
        linksGroup.appendChild(path);
    });

    // Draw bus connectors
    function buildBusPath(parent) {
        const children = (parent.children || []).slice().sort((a, b) => a.x - b.x);
        if (!children.length) return '';
        const rowsMap = new Map();
        children.forEach(ch => {
            const key = currentLayout === 'vertical' ? Math.round(ch.y) : Math.round(ch.x);
            if (!rowsMap.has(key)) rowsMap.set(key, []);
            rowsMap.get(key).push(ch);
        });
        const rows = Array.from(rowsMap.entries()).sort((a, b) => a[0] - b[0]).map(e => e[1]);
        let d = '';
        if (currentLayout === 'vertical') {
            const spineYs = rows.map(r => Math.min(...r.map(ch => ch.y - nodeHeight / 2)) - 12);
            const bottomSpineY = Math.max(...spineYs);
            d += `M ${parent.x} ${parent.y + nodeHeight/2} L ${parent.x} ${bottomSpineY}`;
            rows.forEach((row, i) => {
                const spineY = spineYs[i];
                const xs = row.map(ch => ch.x);
                const left = Math.min(...xs);
                const right = Math.max(...xs);
                d += ` M ${left} ${spineY} L ${right} ${spineY}`;
                row.forEach(ch => {
                    const childTop = ch.y - nodeHeight/2;
                    d += ` M ${ch.x} ${spineY} L ${ch.x} ${childTop}`;
                });
            });
        } else {
            const spineXs = rows.map(r => Math.min(...r.map(ch => ch.x - nodeWidth / 2)) - 12);
            const rightMostSpineX = Math.max(...spineXs);
            d += `M ${parent.x + nodeWidth/2} ${parent.y} L ${rightMostSpineX} ${parent.y}`;
            rows.forEach((row, i) => {
                const spineX = spineXs[i];
                const ys = row.map(ch => ch.y);
                const top = Math.min(...ys);
                const bottom = Math.max(...ys);
                d += ` M ${spineX} ${top} L ${spineX} ${bottom}`;
                row.forEach(ch => {
                    const childLeft = ch.x - nodeWidth/2;
                    d += ` M ${spineX} ${ch.y} L ${childLeft} ${ch.y}`;
                });
            });
        }
        return d;
    }

    mlParents.forEach(parent => {
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('class', 'link');
        path.setAttribute('d', buildBusPath(parent));
        linksGroup.appendChild(path);
    });
    
    // Pre-fetch all images and convert to data URLs
    const imageCache = new Map();
    const defaultIconDataUrl = await imageToDataUrl(userIconUrl);
    const imagePromises = nodesToExport.map(async (d) => {
        if (appSettings.showProfileImages !== false && d.data.photoUrl && d.data.photoUrl.includes('/api/photo/')) {
            const dataUrl = await imageToDataUrl(window.location.origin + d.data.photoUrl);
            if (dataUrl) imageCache.set(d.data.id, dataUrl);
        }
    });
    await Promise.all(imagePromises);

    const namesVisible = isNameVisible();
    const titlesVisible = isJobTitleVisible();
    const departmentsVisible = isDepartmentVisible();

    // Draw nodes
    for (const d of nodesToExport) {
        const nodeG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        nodeG.setAttribute('transform', `translate(${d.x}, ${d.y})`);
        
        // Node rectangle
        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('class', 'node-rect');
        rect.setAttribute('x', -nodeWidth/2);
        rect.setAttribute('y', -nodeHeight/2);
        rect.setAttribute('width', nodeWidth);
        rect.setAttribute('height', nodeHeight);
        
        const nodeColors = appSettings.nodeColors || {};
        let fillColor;
        switch(d.depth) {
            case 0: fillColor = nodeColors.level0 || '#90EE90'; break;
            case 1: fillColor = nodeColors.level1 || '#FFFFE0'; break;
            case 2: fillColor = nodeColors.level2 || '#E0F2FF'; break;
            case 3: fillColor = nodeColors.level3 || '#FFE4E1'; break;
            case 4: fillColor = nodeColors.level4 || '#E8DFF5'; break;
            case 5: fillColor = nodeColors.level5 || '#FFEAA7'; break;
            case 6: fillColor = nodeColors.level6 || '#FAD7FF'; break;
            case 7: fillColor = nodeColors.level7 || '#D7F8FF'; break;
            default: fillColor = '#F0F0F0';
        }
        rect.setAttribute('fill', fillColor);
        rect.setAttribute('stroke', adjustColor(fillColor, -50));
        rect.setAttribute('stroke-width', '2');
        nodeG.appendChild(rect);
        
        // Profile image with circular clipping
        if (appSettings.showProfileImages !== false) {
            // Create unique clip path ID for this image
            const clipId = `clip-${d.data.id || Math.random().toString(36).substr(2, 9)}`;
            
            // Create clip path definition
            const clipPath = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
            clipPath.setAttribute('id', clipId);
            const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
            circle.setAttribute('cx', -nodeWidth/2 + 26); // Center of image (8 + 18)
            circle.setAttribute('cy', 0); // Center vertically (-18 + 18)
            circle.setAttribute('r', 18); // Radius for circular crop
            clipPath.appendChild(circle);
            
            // Add clip path to defs (create defs if it doesn't exist)
            let defs = exportSvg.querySelector('defs');
            if (!defs) {
                defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
                exportSvg.insertBefore(defs, exportSvg.firstChild);
            }
            defs.appendChild(clipPath);
            
            // Create the image with clipping applied
            const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
            const imageUrl = imageCache.get(d.data.id) || defaultIconDataUrl;
            if (imageUrl) {
                image.setAttributeNS('http://www.w3.org/1999/xlink', 'href', imageUrl);
            }
            image.setAttribute('x', -nodeWidth/2 + 8);
            image.setAttribute('y', -18);
            image.setAttribute('width', 36);
            image.setAttribute('height', 36);
            image.setAttribute('clip-path', `url(#${clipId})`);
            nodeG.appendChild(image);
        }

        const textX = showImages ? -nodeWidth/2 + 50 : 0;
        const textAnchor = showImages ? 'start' : 'middle';
        const textWidth = showImages ? nodeWidth - 58 : nodeWidth - 20;
        const titleBaseY = namesVisible ? 3 : -12;
        const fallbackDepartmentBase = namesVisible ? 6 : -10;
        let renderedTitleLines = 0;
        let titleFontSizeValue = showImages ? 11 : 13;
        
        // Name
        if (namesVisible) {
            const nameText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            nameText.setAttribute('class', 'node-text');
            nameText.setAttribute('x', textX);
            nameText.setAttribute('y', -10);
            nameText.setAttribute('text-anchor', textAnchor);
            const nameValue = getVisibleNameText(d.data, { includeFallback: true });
            nameText.textContent = nameValue;
            nameText.setAttribute('font-size', getNameFontSizePx(nameValue));
            nodeG.appendChild(nameText);
        }
        
        // Title
        const titleText = getVisibleJobTitleText(d.data, { includeFallback: true });
        if (titlesVisible && titleText) {
            const titleElement = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            titleElement.setAttribute('class', 'node-title');
            titleElement.setAttribute('x', textX);
            titleElement.setAttribute('y', titleBaseY);
            titleElement.setAttribute('text-anchor', textAnchor);
            const titleFontSize = getTitleFontSizePx(titleText);
            titleElement.setAttribute('font-size', titleFontSize);
            titleFontSizeValue = parseFloat(titleFontSize) || titleFontSizeValue;

            const words = titleText.split(/\s+/).filter(Boolean);
            let currentLine = '';
            let lineCount = 0;
            const maxLines = 2;
            const wrapThreshold = textWidth / 6;

            if (words.length === 0) {
                const tspan = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
                tspan.setAttribute('x', textX);
                tspan.textContent = titleText;
                titleElement.appendChild(tspan);
                lineCount = 1;
            } else {
                for (const word of words) {
                    const testLine = currentLine ? `${currentLine} ${word}` : word;
                    if (testLine.length > wrapThreshold && lineCount < maxLines - 1) {
                        const tspan = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
                        tspan.setAttribute('x', textX);
                        tspan.setAttribute('dy', `${lineCount === 0 ? 0 : 1.2}em`);
                        tspan.textContent = currentLine || word;
                        titleElement.appendChild(tspan);
                        currentLine = word;
                        lineCount++;
                    } else {
                        currentLine = testLine;
                    }
                }

                if (currentLine) {
                    const lastTspan = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
                    lastTspan.setAttribute('x', textX);
                    lastTspan.setAttribute('dy', `${lineCount === 0 ? 0 : 1.2}em`);
                    lastTspan.textContent = currentLine;
                    titleElement.appendChild(lastTspan);
                    lineCount++;
                }
            }
            renderedTitleLines = Math.max(lineCount, 1);

            nodeG.appendChild(titleElement);
        }

        // Department
        const departmentText = getDepartmentDisplayText(d.data, { includeFallback: true, fallback: 'Not specified' });
        if (departmentsVisible && departmentText) {
            const deptText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            deptText.setAttribute('class', 'node-department');
            deptText.setAttribute('x', textX);
            let departmentY;
            if (titlesVisible && titleText) {
                const additionalLines = Math.max(renderedTitleLines - 1, 0);
                const lineSpacing = titleFontSizeValue * 1.2;
                const baseGap = namesVisible ? 18 : 18;
                departmentY = titleBaseY + baseGap + additionalLines * lineSpacing;
            } else {
                departmentY = fallbackDepartmentBase;
            }
            deptText.setAttribute('y', departmentY);
            deptText.setAttribute('text-anchor', textAnchor);
            deptText.textContent = departmentText;
            deptText.setAttribute('font-size', getDepartmentFontSizePx(departmentText));
            nodeG.appendChild(deptText);
        }
        
        nodesGroup.appendChild(nodeG);
    }
    
    g.appendChild(linksGroup);
    g.appendChild(nodesGroup);
    exportSvg.appendChild(g);
    return exportSvg;
}

function applyCompactLayout(nodes) {
    const threshold = appSettings.compactLayoutThreshold || 20;
    nodes.forEach(node => {
        const childCount = (node.children || node._children || []).length;

        if (childCount >= threshold) {
            node.data.hasCompactChildren = true; 
            
            const children = node.children;
            if (children) {
                const columns = Math.ceil(Math.sqrt(children.length));
                
                const horizontalSpacing = currentLayout === 'horizontal'
                    ? HORIZONTAL_COMPACT_HORIZONTAL
                    : VERTICAL_COMPACT_HORIZONTAL;
                const verticalSpacing = currentLayout === 'horizontal'
                    ? HORIZONTAL_COMPACT_VERTICAL
                    : VERTICAL_COMPACT_VERTICAL;

                const totalWidth = (columns - 1) * horizontalSpacing;

                children.forEach((child, i) => {
                    child.data.isCompact = true;
                    const col = i % columns;
                    const row = Math.floor(i / columns);

                    if (currentLayout === 'vertical') {
                        child.y = node.y + (row + 1) * verticalSpacing;
                        child.x = node.x - totalWidth / 2 + col * horizontalSpacing;
                    } else { // horizontal
                        child.x = node.x + (row + 1) * verticalSpacing;
                        child.y = node.y - totalWidth / 2 + col * horizontalSpacing;
                    }
                });
            }
        } else if (node.data.hasCompactChildren) {
            delete node.data.hasCompactChildren;
            const allChildren = (node.children || []).concat(node._children || []);
            allChildren.forEach(child => {
                delete child.data.isCompact;
            });
        }
    });
}

// Employee detail functions
function showEmployeeDetailById(employeeId) {
    if (!employeeId) return;
    const employee = employeeById.get(employeeId);
    if (employee) {
        showEmployeeDetail(employee);
    }
}

function initializeAvatarFallbacks(container) {
    if (!container) return;
    container.querySelectorAll('[data-role="avatar-image"]').forEach(img => {
        const fallback = img.nextElementSibling && img.nextElementSibling.matches('[data-role="avatar-fallback"]')
            ? img.nextElementSibling
            : null;
        if (!fallback) return;

        const showFallback = () => {
            fallback.hidden = false;
            img.style.display = 'none';
        };

        img.addEventListener('error', showFallback, { once: true });
        if (img.complete && img.naturalWidth === 0) {
            showFallback();
        }
    });
}

function getEmployeeRecord(dataOrNode) {
    if (!dataOrNode) return null;
    if (dataOrNode.data) {
        return dataOrNode.data;
    }
    return dataOrNode;
}

function openTitleEditModal(employeeInput, { focusField = 'title' } = {}) {
    const modal = document.getElementById('titleEditModal');
    const backdrop = document.getElementById('titleEditModalBackdrop');
    if (!modal || !backdrop) {
        return;
    }

    const employee = getEmployeeRecord(employeeInput);
    if (!employee || !employee.id) {
        return;
    }

    if (document.activeElement && typeof document.activeElement.focus === 'function') {
        lastFocusBeforeTitleModal = document.activeElement;
    } else {
        lastFocusBeforeTitleModal = null;
    }

    currentTitleEditEmployeeId = employee.id;
    currentOverrideFocusField = focusField === 'department' ? 'department' : 'title';

    const nameElement = modal.querySelector('[data-role="title-edit-name"]');
    const originalTitleElement = modal.querySelector('[data-role="title-edit-original"]');
    const originalDepartmentElement = modal.querySelector('[data-role="department-edit-original"]');
    const titleInputElement = modal.querySelector('#titleEditInput');
    const departmentInputElement = modal.querySelector('#departmentEditInput');
    const helperElement = modal.querySelector('[data-role="title-edit-helper"]');

    const displayName = getVisibleNameText(employee, { includeFallback: true });
    if (nameElement) {
        nameElement.textContent = displayName;
    }

    const originalTitle = getVisibleJobTitleText(employee, { includeFallback: true, useOverrides: false });
    if (originalTitleElement) {
        originalTitleElement.textContent = originalTitle;
    }

    const originalDepartment = getVisibleDepartmentText(employee, { includeFallback: true, useOverrides: false });
    if (originalDepartmentElement) {
        originalDepartmentElement.textContent = originalDepartment;
    }

    const titleOverrideValue = getTitleOverride(employee.id);
    if (titleInputElement) {
        titleInputElement.value = titleOverrideValue != null ? titleOverrideValue : (employee.title || '');
    }

    const departmentOverrideValue = getDepartmentOverride(employee.id);
    if (departmentInputElement) {
        departmentInputElement.value = departmentOverrideValue != null ? departmentOverrideValue : (employee.department || '');
    }

    if (helperElement) {
        helperElement.textContent = t('index.titleEdit.instructions');
    }

    modal.classList.remove('is-hidden');
    modal.setAttribute('aria-hidden', 'false');
    backdrop.classList.remove('is-hidden');
    backdrop.setAttribute('aria-hidden', 'false');
    modal.focus({ preventScroll: true });
    document.body.classList.add('title-edit-open');

    const focusTarget = currentOverrideFocusField === 'department' ? departmentInputElement : titleInputElement;
    if (focusTarget) {
        setTimeout(() => {
            focusTarget.focus();
            if (typeof focusTarget.select === 'function') {
                focusTarget.select();
            }
        }, 20);
    }
}

function closeTitleEditModal() {
    const modal = document.getElementById('titleEditModal');
    const backdrop = document.getElementById('titleEditModalBackdrop');
    if (!modal || !backdrop) {
        return;
    }
    modal.classList.add('is-hidden');
    modal.setAttribute('aria-hidden', 'true');
    backdrop.classList.add('is-hidden');
    backdrop.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('title-edit-open');
    currentTitleEditEmployeeId = null;
    if (lastFocusBeforeTitleModal && typeof lastFocusBeforeTitleModal.focus === 'function') {
        setTimeout(() => {
            lastFocusBeforeTitleModal?.focus?.();
        }, 0);
    }
    lastFocusBeforeTitleModal = null;
}

function isTitleEditModalOpen() {
    const modal = document.getElementById('titleEditModal');
    return modal ? !modal.classList.contains('is-hidden') : false;
}

function handleTitleEditSave() {
    const modal = document.getElementById('titleEditModal');
    if (!modal || !currentTitleEditEmployeeId) {
        closeTitleEditModal();
        return;
    }
    const titleInputElement = modal.querySelector('#titleEditInput');
    const departmentInputElement = modal.querySelector('#departmentEditInput');
    const newTitleValue = titleInputElement ? titleInputElement.value.trim() : '';
    const newDepartmentValue = departmentInputElement ? departmentInputElement.value.trim() : '';
    setTitleOverride(currentTitleEditEmployeeId, newTitleValue);
    setDepartmentOverride(currentTitleEditEmployeeId, newDepartmentValue);
    closeTitleEditModal();
    refreshAfterOverrideChange();
}

function handleTitleEditClear() {
    if (!currentTitleEditEmployeeId) {
        closeTitleEditModal();
        return;
    }
    setTitleOverride(currentTitleEditEmployeeId, '');
    closeTitleEditModal();
    refreshAfterOverrideChange();
}

function handleDepartmentEditClear() {
    if (!currentTitleEditEmployeeId) {
        closeTitleEditModal();
        return;
    }
    setDepartmentOverride(currentTitleEditEmployeeId, '');
    closeTitleEditModal();
    refreshAfterOverrideChange();
}

function initializeTitleEditModal() {
    const modal = document.getElementById('titleEditModal');
    const backdrop = document.getElementById('titleEditModalBackdrop');
    if (!modal || !backdrop) {
        return;
    }

    const form = modal.querySelector('form');
    const saveBtn = modal.querySelector('[data-action="save-title"]');
    const titleClearBtn = modal.querySelector('[data-action="clear-title"]');
    const departmentClearBtn = modal.querySelector('[data-action="clear-department"]');
    const cancelBtn = modal.querySelector('[data-action="cancel-title"]');
    const closeBtn = modal.querySelector('[data-action="close-title"]');
    const titleInputElement = modal.querySelector('#titleEditInput');
    const departmentInputElement = modal.querySelector('#departmentEditInput');

    if (saveBtn) {
        saveBtn.addEventListener('click', event => {
            event.preventDefault();
            handleTitleEditSave();
        });
    }

    if (titleClearBtn) {
        titleClearBtn.addEventListener('click', event => {
            event.preventDefault();
            handleTitleEditClear();
        });
    }

    if (departmentClearBtn) {
        departmentClearBtn.addEventListener('click', event => {
            event.preventDefault();
            handleDepartmentEditClear();
        });
    }

    if (cancelBtn) {
        cancelBtn.addEventListener('click', event => {
            event.preventDefault();
            closeTitleEditModal();
        });
    }

    if (closeBtn) {
        closeBtn.addEventListener('click', event => {
            event.preventDefault();
            closeTitleEditModal();
        });
    }

    if (titleInputElement) {
        titleInputElement.addEventListener('keydown', event => {
            if (event.key === 'Enter') {
                event.preventDefault();
                handleTitleEditSave();
            }
        });
    }

    if (departmentInputElement) {
        departmentInputElement.addEventListener('keydown', event => {
            if (event.key === 'Enter') {
                event.preventDefault();
                handleTitleEditSave();
            }
        });
    }

    if (form) {
        form.addEventListener('submit', event => {
            event.preventDefault();
            handleTitleEditSave();
        });
    }

    backdrop.addEventListener('click', () => {
        closeTitleEditModal();
    });

    modal.addEventListener('click', event => {
        if (event.target === modal) {
            closeTitleEditModal();
        }
    });
}

function showEmployeeDetail(employee) {
    if (!employee) return;

    const detailEmployee = getEmployeeRecord(employee);
    if (!detailEmployee) return;

    currentDetailEmployeeId = detailEmployee.id || null;

    const detailPanel = document.getElementById('employeeDetail');
    const headerContent = document.getElementById('employeeDetailContent');
    const infoContent = document.getElementById('employeeInfo');

    const translateWithFallback = (key, fallback) => {
        const value = t(key);
        return value === key ? fallback : value;
    };

    const departmentLabel = translateWithFallback('index.employee.detail.department', 'Department');
    const departmentUnknown = translateWithFallback('index.employee.detail.departmentUnknown', 'Unknown department');
    const emailLabel = translateWithFallback('index.employee.detail.email', 'Email');
    const emailUnknown = translateWithFallback('index.employee.detail.emailUnknown', 'No email provided');
    const phoneLabel = translateWithFallback('index.employee.detail.phone', 'Phone');
    const phoneUnknown = translateWithFallback('index.employee.detail.phoneUnknown', 'No phone provided');
    const businessPhoneLabel = translateWithFallback('index.employee.detail.businessPhone', 'Business Phone');
    const businessPhoneUnknown = translateWithFallback('index.employee.detail.businessPhoneUnknown', 'No business phone provided');
    const hireDateLabel = t('index.employee.detail.hireDate');
    const officeLabel = t('index.employee.detail.office');
    const locationLabel = t('index.employee.detail.location');
    const managerHeading = t('index.employee.detail.manager');
    const jobTitleDisplay = getVisibleJobTitleText(detailEmployee, { includeFallback: true });
    const baseTitle = getVisibleJobTitleText(detailEmployee, { includeFallback: true, useOverrides: false });
    const titleOverrideActive = isTitleOverridden(detailEmployee.id);
    const departmentDisplay = getDepartmentDisplayText(detailEmployee, { includeFallback: true, fallback: departmentUnknown });
    const baseDepartment = getVisibleDepartmentText(detailEmployee, { includeFallback: true, useOverrides: false, fallback: departmentUnknown });
    const departmentOverrideActive = isDepartmentOverridden(detailEmployee.id);
    const namesVisible = isNameVisible();
    const displayName = getVisibleNameText(detailEmployee, { includeFallback: true });
    const avatarAlt = namesVisible ? (detailEmployee.name || '') : displayName;
    const initials = namesVisible
        ? (detailEmployee.name || '')
            .split(' ')
            .map(n => n[0])
            .join('')
            .substring(0, 2)
            .toUpperCase()
        : '';

    const employeeAvatar = renderAvatar({
        imageUrl: detailEmployee.photoUrl && detailEmployee.photoUrl.includes('/api/photo/') ? detailEmployee.photoUrl : '',
        name: avatarAlt,
        initials,
        imageClass: 'employee-avatar-image',
        fallbackClass: 'employee-avatar-fallback'
    });

    const titleBadgeText = t('index.employee.titleEditedBadge');
    const departmentBadgeText = t('index.employee.departmentEditedBadge');
    const editButtonLabel = t('index.employee.editTitleButton');
    const editDepartmentLabel = t('index.employee.editDepartmentButton');
    const titleBadgeMarkup = titleOverrideActive
        ? `<span class="title-override-badge">${escapeHtml(titleBadgeText)}</span>`
        : '';
    const titleValueMarkup = jobTitleDisplay
        ? `<div class="employee-title${titleOverrideActive ? ' employee-title--edited' : ''}">${escapeHtml(jobTitleDisplay)}${titleBadgeMarkup}</div>`
        : '';
    const titleButtonMarkup = detailEmployee.id
        ? `<button type="button" class="title-edit-btn" data-action="edit-title" data-employee-id="${escapeHtml(detailEmployee.id)}">${escapeHtml(editButtonLabel)}</button>`
        : '';
    const titleMarkup = `
        <div class="employee-title-row">
            ${titleValueMarkup || `<div class="employee-title-placeholder">${escapeHtml(t('index.employee.noTitle'))}</div>`}
            ${titleButtonMarkup}
        </div>
    `;

    const departmentBadgeMarkup = departmentOverrideActive
        ? `<span class="department-override-badge">${escapeHtml(departmentBadgeText)}</span>`
        : '';
    const hasDepartmentDisplay = Boolean((departmentDisplay || '').trim());
    const departmentValueMarkup = hasDepartmentDisplay
        ? `<div class="employee-department${departmentOverrideActive ? ' employee-department--edited' : ''}">${escapeHtml(departmentDisplay)}${departmentBadgeMarkup}</div>`
        : `<div class="employee-department-placeholder">${escapeHtml(departmentUnknown)}</div>`;
    const departmentButtonMarkup = detailEmployee.id
        ? `<button type="button" class="department-edit-btn" data-action="edit-department" data-employee-id="${escapeHtml(detailEmployee.id)}" data-focus-field="department">${escapeHtml(editDepartmentLabel)}</button>`
        : '';
    const departmentMarkup = `
        <div class="employee-department-row">
            ${departmentValueMarkup}
            ${departmentButtonMarkup}
        </div>
    `;

    headerContent.innerHTML = `
        <div class="employee-avatar-container">
            ${employeeAvatar}
        </div>
        <div class="employee-name">
            ${displayName ? `<h2>${escapeHtml(displayName)}</h2>` : ''}
        </div>
        ${titleMarkup}
        ${departmentMarkup}
    `;

    let infoHTML = '';

    if (namesVisible) {
        infoHTML += `
        <div class="info-item">
            <div class="info-label">${emailLabel}</div>
            <div class="info-value">
                ${detailEmployee.email ? `<a href="mailto:${escapeHtml(detailEmployee.email)}">${escapeHtml(detailEmployee.email)}</a>` : emailUnknown}
            </div>
        </div>
        <div class="info-item">
            <div class="info-label">${phoneLabel}</div>
            <div class="info-value">${escapeHtml(detailEmployee.phone || phoneUnknown)}</div>
        </div>
        <div class="info-item">
            <div class="info-label">${businessPhoneLabel}</div>
            <div class="info-value">${escapeHtml(detailEmployee.businessPhone || businessPhoneUnknown)}</div>
        </div>
        `;
    }

    infoHTML += `
        ${detailEmployee.hireDate ? `
        <div class="info-item">
            <div class="info-label">${hireDateLabel}</div>
            <div class="info-value">${escapeHtml(formatHireDate(detailEmployee.hireDate))}</div>
        </div>
        ` : ''}
        ${detailEmployee.location ? `
        <div class="info-item">
            <div class="info-label">${officeLabel}</div>
            <div class="info-value">${escapeHtml(detailEmployee.location)}</div>
        </div>
        ` : ''}
        ${detailEmployee.city || detailEmployee.state || detailEmployee.country ? `
        <div class="info-item">
            <div class="info-label">${locationLabel}</div>
            <div class="info-value">${[detailEmployee.city, detailEmployee.state, detailEmployee.country].filter(Boolean).map(escapeHtml).join(', ')}</div>
        </div>
        ` : ''}
    `;

    if (titleOverrideActive) {
        const originalLabel = t('index.titleEdit.originalLabel');
        infoHTML = `
        <div class="info-item info-item--title-original">
            <div class="info-label">${originalLabel}</div>
            <div class="info-value">
                ${escapeHtml(baseTitle)}
            </div>
        </div>
        ` + infoHTML;
    }

    if (departmentOverrideActive) {
        const originalDepartmentLabel = t('index.titleEdit.originalDepartmentLabel');
        infoHTML = `
        <div class="info-item info-item--department-original">
            <div class="info-label">${originalDepartmentLabel}</div>
            <div class="info-value">
                ${escapeHtml(baseDepartment)}
            </div>
        </div>
        ` + infoHTML;
    }

    if (detailEmployee.managerId && window.currentOrgData) {
        const manager = findManagerById(window.currentOrgData, detailEmployee.managerId);
        if (manager) {
            const managerDisplayName = getVisibleNameText(manager, { includeFallback: true });
            const managerAvatarAlt = namesVisible ? (manager.name || '') : managerDisplayName;
            const managerInitials = namesVisible
                ? (manager.name || '')
                    .split(' ')
                    .map(n => n[0])
                    .join('')
                    .substring(0, 2)
                    .toUpperCase()
                : '';

            const managerAvatar = renderAvatar({
                imageUrl: manager.photoUrl && manager.photoUrl.includes('/api/photo/') ? manager.photoUrl : '',
                name: managerAvatarAlt,
                initials: managerInitials,
                imageClass: 'manager-avatar-image',
                fallbackClass: 'manager-avatar-fallback'
            });

            const managerTitleDisplay = getVisibleJobTitleText(manager, { includeFallback: true });
            const managerTitleMarkup = managerTitleDisplay
                ? `<div class="manager-title">${escapeHtml(managerTitleDisplay)}</div>`
                : '';

            infoHTML += `
                <div class="manager-section">
                    <h3>${managerHeading}</h3>
                    <div class="manager-item" data-employee-id="${escapeHtml(manager.id)}">
                        <div class="manager-avatar-container">
                            ${managerAvatar}
                        </div>
                        <div class="manager-details">
                            <div class="manager-name">${escapeHtml(managerDisplayName)}</div>
                            ${managerTitleMarkup}
                        </div>
                    </div>
                </div>
            `;
        }
    }

    const directReports = detailEmployee.children || [];
    if (directReports.length > 0) {
        const directReportsLabel = t('index.employee.detail.directReportsWithCount', { count: directReports.length });
        infoHTML += `
            <div class="direct-reports">
                <h3>${directReportsLabel}</h3>
                ${directReports.map(report => {
                    const reportDisplayName = getVisibleNameText(report, { includeFallback: true });
                    const reportAvatarAlt = namesVisible ? (report.name || '') : reportDisplayName;
                    const reportInitials = namesVisible
                        ? (report.name || '')
                            .split(' ')
                            .map(n => n[0])
                            .join('')
                            .substring(0, 2)
                            .toUpperCase()
                        : '';

                    const reportAvatar = renderAvatar({
                        imageUrl: report.photoUrl && report.photoUrl.includes('/api/photo/') ? report.photoUrl : '',
                        name: reportAvatarAlt,
                        initials: reportInitials,
                        imageClass: 'report-avatar-image',
                        fallbackClass: 'report-avatar-fallback'
                    });

                    const reportTitleDisplay = getVisibleJobTitleText(report, { includeFallback: true });
                    const reportTitleMarkup = reportTitleDisplay
                        ? `<div class="report-title">${escapeHtml(reportTitleDisplay)}</div>`
                        : '';

                    return `
                        <div class="report-item" data-employee-id="${escapeHtml(report.id)}">
                            <div class="report-avatar-container">
                                ${reportAvatar}
                            </div>
                            <div class="report-details">
                                <div class="report-name">${escapeHtml(reportDisplayName)}</div>
                                ${reportTitleMarkup}
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;
    }

    infoContent.innerHTML = infoHTML;
    initializeAvatarFallbacks(detailPanel);
    detailPanel.classList.add('active');
}

function refreshEmployeeDetailPanel() {
    const detailPanel = document.getElementById('employeeDetail');
    if (!detailPanel || !detailPanel.classList.contains('active')) {
        return;
    }
    if (!currentDetailEmployeeId) {
        return;
    }
    const employee = employeeById.get(currentDetailEmployeeId);
    if (!employee) {
        currentDetailEmployeeId = null;
        return;
    }
    showEmployeeDetail(employee);
}

function closeEmployeeDetail() {
    document.getElementById('employeeDetail').classList.remove('active');
    currentDetailEmployeeId = null;
}

function findNodeById(node, targetId) {
    if (node.data.id === targetId) {
        return node;
    }
    if (node.children || node._children) {
        const children = node.children || node._children;
        for (let child of children) {
            const result = findNodeById(child, targetId);
            if (result) return result;
        }
    }
    return null;
}

function findManagerById(rootData, managerId) {
    function searchNode(node) {
        if (node.id === managerId) {
            return node;
        }
        if (node.children) {
            for (let child of node.children) {
                const result = searchNode(child);
                if (result) return result;
            }
        }
        return null;
    }
    return searchNode(rootData);
}

function highlightNode(nodeId, highlight = true) {
    if (appSettings.searchHighlight !== false) {
        g.selectAll('.node-rect').each(function(d) {
            if (d.data.id === nodeId) {
                d3.select(this).classed('search-highlight', highlight);
            }
        });
    }
}

function clearHighlights() {
    g.selectAll('.node-rect').classed('search-highlight', false);
}

let searchTimeout;
const searchInput = document.getElementById('searchInput');
const searchResults = document.getElementById('searchResults');

searchInput.addEventListener('input', function(e) {
    clearTimeout(searchTimeout);
    const query = e.target.value.trim();
    
    clearHighlights();
    
    if (query.length < 2) {
        searchResults.classList.remove('active');
        return;
    }
    
    searchTimeout = setTimeout(() => {
        performSearch(query);
    }, 300);
});

searchInput.addEventListener('focus', function(e) {
    if (e.target.value.length >= 2) {
        performSearch(e.target.value);
    }
});

document.addEventListener('click', function(e) {
    if (!e.target.closest('.search-wrapper')) {
        searchResults.classList.remove('active');
    }
});

async function performSearch(query) {
    try {
        const response = await fetch(`${API_BASE_URL}/api/search?q=${encodeURIComponent(query)}`);
        const results = await response.json();
        
        if (results.length > 0) {
            displaySearchResults(results);
        } else {
            searchResults.innerHTML = `<div class="search-result-item">${t('index.search.noResults')}</div>`;
            searchResults.classList.add('active');
        }
    } catch (error) {
        console.error('Search error:', error);
    }
}

function displaySearchResults(results) {
    searchResults.innerHTML = '';

    results.forEach(emp => {
        const item = document.createElement('div');
        item.className = 'search-result-item';
        item.dataset.employeeId = emp.id;
        item.dataset.name = emp.name || '';
        item.dataset.title = emp.title || '';
        item.dataset.department = emp.department || '';
        item.dataset.location = emp.location || emp.officeLocation || '';

        const name = document.createElement('div');
        name.className = 'search-result-name';
        const nameText = getVisibleNameText(emp, { includeFallback: true });
        name.textContent = nameText;
        name.hidden = !nameText;

        const title = document.createElement('div');
        title.className = 'search-result-title';
        populateResultMeta(title, emp);

        item.appendChild(name);
        item.appendChild(title);
        searchResults.appendChild(item);
    });
    searchResults.classList.add('active');
}

function refreshSearchResultsPresentation() {
    if (searchResults) {
        searchResults.querySelectorAll('.search-result-item').forEach(item => {
            const meta = item.querySelector('.search-result-title');
            const nameElement = item.querySelector('.search-result-name');
            if (!meta) return;
            const employeeId = item.dataset.employeeId;
            const metaSource = employeeId && employeeById.has(employeeId)
                ? employeeById.get(employeeId)
                : {
                    id: employeeId,
                    title: item.dataset.title || '',
                    department: item.dataset.department || '',
                    location: item.dataset.location || ''
                };
            populateResultMeta(meta, metaSource);
            if (nameElement) {
                const nameSource = employeeId && employeeById.has(employeeId)
                    ? employeeById.get(employeeId)
                    : { name: item.dataset.name || '' };
                const nameText = getVisibleNameText(nameSource, { includeFallback: true });
                nameElement.textContent = nameText;
                nameElement.hidden = !nameText;
            }
        });
    }

    const topUserResults = document.getElementById('topUserResults');
    if (topUserResults) {
        topUserResults.querySelectorAll('.search-result-item').forEach(item => {
            const meta = item.querySelector('.search-result-title');
            const nameElement = item.querySelector('.search-result-name');
            if (!meta) return;
            const employeeId = item.dataset.employeeId;
            const metaSource = employeeId && employeeById.has(employeeId)
                ? employeeById.get(employeeId)
                : {
                    id: employeeId,
                    title: item.dataset.title || '',
                    department: item.dataset.department || '',
                    location: item.dataset.location || ''
                };
            populateResultMeta(meta, metaSource);
            if (nameElement) {
                const nameSource = employeeId && employeeById.has(employeeId)
                    ? employeeById.get(employeeId)
                    : { name: item.dataset.name || '' };
                const nameText = getVisibleNameText(nameSource, { includeFallback: true });
                nameElement.textContent = nameText;
                nameElement.hidden = !nameText;
            }
        });
    }
}

function selectSearchResult(employeeId) {
    const employee = employeeById.get(employeeId);
    if (employee) {
        showEmployeeDetail(employee);
        searchResults.classList.remove('active');
        searchInput.value = '';
        
        expandToEmployee(employeeId);
    }
}

function expandToEmployee(employeeId) {
    if (appSettings.searchAutoExpand === false) {
        const targetNode = findNodeById(root, employeeId);
        if (targetNode) {
            highlightNode(employeeId);
            showEmployeeDetail(targetNode.data);
        }
        return;
    }
    
    const path = [];
    
    function findPath(node, targetId, currentPath) {
        currentPath.push(node);
        
        if (node.data.id === targetId) {
            path.push(...currentPath);
            return true;
        }
        
        if (node.children || node._children) {
            const children = node.children || node._children;
            for (let child of children) {
                if (findPath(child, targetId, [...currentPath])) {
                    return true;
                }
            }
        }
        
        return false;
    }
    
    findPath(root, employeeId, []);
    
    path.forEach(node => {
        if (node._children) {
            node.children = node._children;
            node._children = null;
        }
    });
    
    update(root);
    
    const targetNode = path[path.length - 1];
    if (targetNode) {
        setTimeout(() => {
            highlightNode(employeeId);
        }, 600);
        
        const container = document.getElementById('orgChart');
        const width = container.clientWidth;
        const height = container.clientHeight;
        
        svg.transition()
            .duration(750)
            .call(zoom.transform, 
                d3.zoomIdentity
                    .translate(width/2, height/2)
                    .scale(1)
                    .translate(-targetNode.x, -targetNode.y)
            );
    }
}

// Export to XLSX function
async function exportToXLSX() {
    await waitForTranslations();
    try {
        const response = await fetch(`${API_BASE_URL}/api/export-xlsx`);
        
        if (response.ok) {
            // Get the blob
            const blob = await response.blob();
            
            // Create download link
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            
            // Get filename from Content-Disposition header or use default
            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = `org-chart-${new Date().toISOString().split('T')[0]}.xlsx`;
            if (contentDisposition && contentDisposition.includes('filename=')) {
                filename = contentDisposition.split('filename=')[1].replace(/"/g, '');
            }
            
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            // Clean up
            window.URL.revokeObjectURL(url);
        } else {
            const errorData = await response.json();
            alert(t('index.alerts.xlsxExportFailed', {
                message: errorData.error || t('index.alerts.unknownError')
            }));
        }
    } catch (error) {
        console.error('Export error:', error);
        alert(t('index.alerts.xlsxExportError'));
    }
}

// Data sync functionality
let _syncPollingInterval = null;

function setSyncButtonState(syncing, statusText = null) {
    const syncBtn = document.getElementById('syncBtn');
    if (!syncBtn) return;
    
    syncBtn.classList.toggle('is-syncing', syncing);
    
    if (syncing) {
        syncBtn.innerHTML = `<span class="sync-spinner"></span>${statusText || t('buttons.syncing', { defaultValue: 'Syncing...' })}`;
    } else if (statusText) {
        syncBtn.textContent = statusText;
        // Reset to default after 3 seconds
        setTimeout(() => {
            syncBtn.textContent = t('buttons.sync', { defaultValue: 'Sync' });
        }, 3000);
    } else {
        syncBtn.textContent = t('buttons.sync', { defaultValue: 'Sync' });
    }
}

function stopSyncPolling() {
    if (_syncPollingInterval) {
        clearInterval(_syncPollingInterval);
        _syncPollingInterval = null;
    }
}

async function pollSyncStatus() {
    try {
        const response = await fetch('/api/settings');
        if (!response.ok) return;
        
        const settings = response.headers.get('content-type')?.includes('application/json')
            ? await response.json()
            : null;
        
        if (!settings) return;
        
        const status = settings.dataUpdateStatus;
        if (!status || status.state !== 'running') {
            stopSyncPolling();
            
            if (status?.state === 'completed') {
                setSyncButtonState(false, t('buttons.syncComplete', { defaultValue: 'Sync complete' }));
                // Update appSettings with new last updated time and refresh header
                appSettings.dataLastUpdatedAt = settings.dataLastUpdatedAt;
                updateHeaderSubtitle();
                // Reload the org chart data
                refreshOrgChart();
            } else if (status?.state === 'failed') {
                setSyncButtonState(false, t('buttons.syncFailed', { defaultValue: 'Sync failed' }));
                // Restore header subtitle after failed sync
                appSettings.dataLastUpdatedAt = settings.dataLastUpdatedAt;
                updateHeaderSubtitle();
            } else {
                setSyncButtonState(false);
                // Restore header subtitle when sync finished (idle state)
                appSettings.dataLastUpdatedAt = settings.dataLastUpdatedAt;
                updateHeaderSubtitle();
            }
        }
    } catch (error) {
        console.warn('Sync status poll error:', error);
    }
}

function startSyncPolling() {
    stopSyncPolling();
    _syncPollingInterval = setInterval(pollSyncStatus, 3000);
}

async function triggerDataSync() {
    const syncBtn = document.getElementById('syncBtn');
    if (!syncBtn || syncBtn.classList.contains('is-syncing')) return;
    
    setSyncButtonState(true);
    updateHeaderSubtitle(true); // Show syncing in header
    
    try {
        const response = await fetch('/api/update-now', { method: 'POST' });
        
        if (response.ok) {
            startSyncPolling();
        } else if (response.status === 409) {
            // Update already in progress
            setSyncButtonState(true, t('buttons.syncing', { defaultValue: 'Syncing...' }));
            startSyncPolling();
        } else {
            setSyncButtonState(false, t('buttons.syncFailed', { defaultValue: 'Sync failed' }));
            updateHeaderSubtitle(); // Restore normal header
        }
    } catch (error) {
        console.error('Sync trigger error:', error);
        setSyncButtonState(false, t('buttons.syncFailed', { defaultValue: 'Sync failed' }));
        updateHeaderSubtitle(); // Restore normal header
    }
}

// Logout function
async function logout() {
    try {
        const response = await fetch('/logout', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            window.location.href = '/';
        }
    } catch (error) {
        console.error('Logout error:', error);
        // Force redirect even if request fails
        window.location.href = '/';
    }
}

function registerEventHandlers() {
    setupStaticEventListeners();
}

window.addEventListener('resize', () => {
    updateSvgSize();
    if (userAdjustedZoom) return;
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
        fitToScreen({ duration: 300, resetUser: true });
    }, RESIZE_DEBOUNCE_MS);
});

document.addEventListener('DOMContentLoaded', () => {
    initializeTitleEditModal();
    registerEventHandlers();
    updateOverrideResetButtons();
    init();
});

document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        if (isTitleEditModalOpen()) {
            closeTitleEditModal();
            return;
        }
        closeEmployeeDetail();
        clearHighlights();
    }
});