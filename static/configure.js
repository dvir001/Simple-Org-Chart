const API_BASE_URL = window.location.origin;
let currentSettings = {};

const EXPORT_COLUMN_KEYS = [
    { key: 'name', selectId: 'exportColumnName' },
    { key: 'title', selectId: 'exportColumnTitle' },
    { key: 'department', selectId: 'exportColumnDepartment' },
    { key: 'email', selectId: 'exportColumnEmail' },
    { key: 'phone', selectId: 'exportColumnPhone' },
    { key: 'businessPhone', selectId: 'exportColumnBusinessPhone' },
    { key: 'hireDate', selectId: 'exportColumnHireDate' },
    { key: 'country', selectId: 'exportColumnCountry' },
    { key: 'state', selectId: 'exportColumnState' },
    { key: 'city', selectId: 'exportColumnCity' },
    { key: 'office', selectId: 'exportColumnOffice' },
    { key: 'manager', selectId: 'exportColumnManager' }
];

const EXPORT_COLUMN_DEFAULTS = {
    name: 'show',
    title: 'show',
    department: 'show',
    email: 'show',
    phone: 'show',
    businessPhone: 'show',
    hireDate: 'admin',
    country: 'show',
    state: 'show',
    city: 'show',
    office: 'show',
    manager: 'show'
};

const tagPickers = {};
let topLevelUserPicker = null;
let filterMetadata = { jobTitles: [], departments: [], employees: [] };

let hasUnsavedChanges = false;
let pendingLogoReset = false;
let pendingFaviconReset = false;
let isInitializing = true;
let beforeUnloadBound = false;
const unsavedReasons = new Set();
const DEFAULT_UPDATE_TIME = '20:00';
const DEFAULT_UPDATE_TIMEZONE = 'UTC';
const DEFAULT_HEADER_COLOR = '#0078D4';
const NODE_COLOR_DEFAULTS = {
    level0: '#90EE90',
    level1: '#FFFFE0',
    level2: '#E0F2FF',
    level3: '#FFE4E1',
    level4: '#E8DFF5',
    level5: '#FFEAA7',
    level6: '#FAD7FF',
    level7: '#D7F8FF'
};

function resolveTranslation(key, fallback) {
    return getTranslation(key, fallback);
}

function updateLastUpdatedDisplay(value) {
    const lastUpdatedEl = document.getElementById('dataLastUpdated');
    if (!lastUpdatedEl) {
        return;
    }
    if (value) {
        const text = typeof value === 'string' ? value.trim() : `${value}`;
        if (text) {
            lastUpdatedEl.textContent = text;
            return;
        }
    }
    lastUpdatedEl.textContent = resolveTranslation('configure.data.manualUpdate.lastUpdatedPending', 'Pending');
}

function setUpdateStatus(key, fallback) {
    const statusEl = document.getElementById('updateStatus');
    if (!statusEl) {
        return;
    }
    statusEl.title = '';
    if (!key) {
        statusEl.textContent = fallback || '';
        return;
    }
    statusEl.textContent = resolveTranslation(key, fallback);
}

function getTranslation(key, fallbackText) {
    try {
        if (window.i18n && typeof window.i18n.t === 'function') {
            const translated = window.i18n.t(key);
            if (translated && translated !== key) {
                return translated;
            }
        }
    } catch (error) {
        console.warn('Translation lookup failed', key, error);
    }
    return fallbackText;
}

function handleBeforeUnload(event) {
    if (!hasUnsavedChanges) {
        return;
    }
    const message = getTranslation(
        'configure.unsavedChanges.confirmLeave',
        'You have unsaved changes. Leave without saving?'
    );
    event.preventDefault();
    event.returnValue = message;
    return message;
}

function updateUnsavedBanner() {
    const banner = document.getElementById('unsavedChangesBar');
    if (!banner) {
        return;
    }
    if (hasUnsavedChanges) {
        banner.classList.add('unsaved-changes--visible');
        if (!beforeUnloadBound) {
            window.addEventListener('beforeunload', handleBeforeUnload);
            beforeUnloadBound = true;
        }
    } else {
        banner.classList.remove('unsaved-changes--visible');
        if (beforeUnloadBound) {
            window.removeEventListener('beforeunload', handleBeforeUnload);
            beforeUnloadBound = false;
        }
    }
}

function markUnsavedChange(reason = 'general') {
    if (isInitializing) {
        return;
    }
    unsavedReasons.add(reason);
    hasUnsavedChanges = unsavedReasons.size > 0;
    updateUnsavedBanner();
}

function clearUnsavedChangeState() {
    hasUnsavedChanges = false;
    pendingLogoReset = false;
    pendingFaviconReset = false;
    unsavedReasons.clear();
    updateUnsavedBanner();
}

function clearUnsavedReason(reason) {
    if (!unsavedReasons.has(reason)) {
        return;
    }
    unsavedReasons.delete(reason);
    hasUnsavedChanges = unsavedReasons.size > 0;
    updateUnsavedBanner();
}

function discardUnsavedChanges() {
    if (!hasUnsavedChanges) {
        return;
    }

    const message = getTranslation(
        'configure.unsavedChanges.confirmDiscard',
        'Discard all unsaved changes?'
    );

    if (!window.confirm(message)) {
        return;
    }

    const previousInitializingState = isInitializing;
    isInitializing = true;
    try {
        applySettings(currentSettings);
    } finally {
        isInitializing = previousInitializingState;
    }

    clearUnsavedChangeState();
    showStatus('Unsaved changes discarded.', 'info');
}

function attachUnsavedListeners() {
    const inputs = document.querySelectorAll('input:not([type="file"]), select, textarea');
    inputs.forEach((field) => {
        const markChange = () => markUnsavedChange();
        field.addEventListener('input', markChange);
        field.addEventListener('change', markChange);
    });
}

function confirmUnsavedNavigation() {
    if (!hasUnsavedChanges) {
        return true;
    }
    const message = getTranslation(
        'configure.unsavedChanges.confirmLeave',
        'You have unsaved changes. Leave without saving?'
    );
    return window.confirm(message);
}

function registerNavigationGuards() {
    const guardSelectors = [
        '.nav-link',
        '[data-config-action="logout"]'
    ];

    guardSelectors.forEach(selector => {
        document.querySelectorAll(selector).forEach(element => {
            element.addEventListener('click', event => {
                if (confirmUnsavedNavigation()) {
                    return;
                }
                event.preventDefault();
                event.stopImmediatePropagation();
            });
        });
    });
}

function parseListString(value) {
    if (!value) {
        return [];
    }
    if (Array.isArray(value)) {
        return value
            .map(item => (typeof item === 'string' ? item.trim() : `${item}`.trim()))
            .filter(item => item.length > 0);
    }
    const trimmed = String(value).trim();
    if (!trimmed) {
        return [];
    }
    if (trimmed.startsWith('[')) {
        try {
            const parsed = JSON.parse(trimmed);
            if (Array.isArray(parsed)) {
                return parsed
                    .map(item => (typeof item === 'string' ? item.trim() : `${item}`.trim()))
                    .filter(item => item.length > 0);
            }
        } catch (error) {
            console.warn('Failed to parse list JSON', error);
        }
    }
    return trimmed
        .split(/\s*[;,]+\s*/)
        .map(item => item.trim())
        .filter(item => item.length > 0);
}

function normalizeHexColor(value, fallback = '#000000') {
    const coerce = (candidate) => {
        if (typeof candidate !== 'string') {
            return null;
        }
        let text = candidate.trim();
        if (!text) {
            return null;
        }
        if (text.startsWith('#')) {
            text = text.slice(1);
        }
        if (/^[0-9a-fA-F]{6}$/.test(text)) {
            return `#${text.toUpperCase()}`;
        }
        return null;
    };

    const normalizedValue = coerce(value);
    if (normalizedValue) {
        return normalizedValue;
    }

    const normalizedFallback = coerce(fallback);
    if (normalizedFallback) {
        return normalizedFallback;
    }

    return '#000000';
}

function applyColorToInputs(colorInput, hexInput, rawValue, fallbackValue = '#000000') {
    const fallback = normalizeHexColor(
        (colorInput && colorInput.value) || (hexInput && hexInput.value) || fallbackValue,
        fallbackValue
    );
    const normalized = normalizeHexColor(rawValue, fallback);
    if (colorInput) {
        colorInput.value = normalized.toLowerCase();
    }
    if (hexInput) {
        hexInput.value = normalized;
    }
    return normalized;
}

class TagPicker {
    constructor({ pickerId, hiddenInputId, options = [], placeholder = '' }) {
        this.root = document.getElementById(pickerId);
        this.hiddenInput = document.getElementById(hiddenInputId);
        if (!this.root || !this.hiddenInput) {
            this.enabled = false;
            return;
        }

        this.enabled = true;
        this.options = Array.isArray(options) ? options.slice() : [];
        this.options = this.options.filter(Boolean);
        this.options.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

        this.tagContainer = this.root.querySelector('[data-role="tag-container"]');
        this.dropdown = this.root.querySelector('[data-role="dropdown"]');
        this.input = this.root.querySelector('.tag-picker__input');
        if (placeholder && this.input) {
            this.input.placeholder = placeholder;
        }

        this.selected = [];
        this.selectedSet = new Set();
        this.filteredOptions = [];

        this.handleDocumentClick = this.handleDocumentClick.bind(this);
        this.handleDropdownClick = this.handleDropdownClick.bind(this);
        this.handleTagClick = this.handleTagClick.bind(this);
        this.handleInput = this.handleInput.bind(this);
    this.handleKeyDown = this.handleKeyDown.bind(this);
    this.focusInputSoon = this.focusInputSoon.bind(this);

        if (this.tagContainer) {
            this.tagContainer.addEventListener('click', this.handleTagClick);
        }
        if (this.dropdown) {
            this.dropdown.addEventListener('click', this.handleDropdownClick);
        }
        if (this.input) {
            this.input.addEventListener('input', this.handleInput);
            this.input.addEventListener('focus', () => this.openDropdown());
            this.input.addEventListener('keydown', this.handleKeyDown);
        }

        document.addEventListener('click', this.handleDocumentClick);
        this.renderTags();
        this.closeDropdown();
        this.updateHiddenInput();
    }

    destroy() {
        if (!this.enabled) {
            return;
        }
        document.removeEventListener('click', this.handleDocumentClick);
        this.enabled = false;
    }

    setOptions(options) {
        if (!this.enabled) {
            return;
        }
        this.options = Array.isArray(options) ? options.slice() : [];
        this.options = this.options.filter(Boolean);
        this.options.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
        this.refreshDropdown();
    }

    setValue(values) {
        if (!this.enabled) {
            return;
        }
        this.selected = [];
        this.selectedSet = new Set();
        (values || []).forEach(value => {
            const normalized = (value || '').trim();
            if (!normalized) {
                return;
            }
            const key = normalized.toLowerCase();
            if (!this.selectedSet.has(key)) {
                this.selected.push(normalized);
                this.selectedSet.add(key);
            }
        });
        this.renderTags();
        this.updateHiddenInput();
        if (this.input) {
            this.input.value = '';
        }
        this.closeDropdown();
        markUnsavedChange();
    }

    getValue() {
        if (!this.enabled) {
            return [];
        }
        return this.selected.slice();
    }

    clear() {
        this.setValue([]);
    }

    handleInput() {
        this.refreshDropdown();
        this.openDropdown();
    }

    handleKeyDown(event) {
        if (event.key === 'Backspace' && this.input && !this.input.value && this.selected.length > 0) {
            const last = this.selected[this.selected.length - 1];
            this.removeValue(last);
            event.preventDefault();
        } else if ((event.key === 'Enter' || event.key === 'Tab') && this.input) {
            const query = this.input.value.trim();
            if (!query) {
                return;
            }
            if (this.filteredOptions.length > 0) {
                this.addValue(this.filteredOptions[0]);
            } else {
                this.addValue(query);
            }
            event.preventDefault();
        }
    }

    handleDropdownClick(event) {
        const option = event.target.closest('[data-value]');
        if (!option) {
            return;
        }
        event.preventDefault();
        event.stopPropagation();
        const value = option.getAttribute('data-value') || '';
        this.addValue(value);
    }

    handleTagClick(event) {
        const removeBtn = event.target.closest('.tag-picker__remove');
        if (!removeBtn) {
            return;
        }
        const value = removeBtn.getAttribute('data-value') || '';
        this.removeValue(value);
    }

    handleDocumentClick(event) {
        if (!this.root) {
            return;
        }
        if (!this.root.contains(event.target)) {
            this.closeDropdown();
        }
    }

    addValue(rawValue) {
        if (!this.enabled) {
            return;
        }
        const normalized = (rawValue || '').trim();
        if (!normalized) {
            return;
        }
        const key = normalized.toLowerCase();
        if (this.selectedSet.has(key)) {
            if (this.input) {
                this.input.value = '';
            }
            this.closeDropdown();
            return;
        }
        this.selected.push(normalized);
        this.selectedSet.add(key);
        this.renderTags();
        this.updateHiddenInput();
        if (this.input) {
            this.input.value = '';
        }
        this.refreshDropdown();
        this.openDropdown();
        this.focusInputSoon();
        markUnsavedChange();
    }

    removeValue(rawValue) {
        if (!this.enabled) {
            return;
        }
        const normalized = (rawValue || '').trim();
        if (!normalized) {
            return;
        }
        const key = normalized.toLowerCase();
        if (!this.selectedSet.has(key)) {
            return;
        }
        this.selected = this.selected.filter(item => item.toLowerCase() !== key);
        this.selectedSet.delete(key);
        this.renderTags();
        this.updateHiddenInput();
        this.refreshDropdown();
        this.openDropdown();
        this.focusInputSoon();
        markUnsavedChange();
    }

    renderTags() {
        if (!this.tagContainer) {
            return;
        }
        this.tagContainer.innerHTML = '';
        if (this.selected.length === 0) {
            return;
        }
        this.selected.forEach(value => {
            const tag = document.createElement('span');
            tag.className = 'tag-picker__tag';

            const label = document.createElement('span');
            label.textContent = value;
            tag.appendChild(label);

            const removeBtn = document.createElement('button');
            removeBtn.type = 'button';
            removeBtn.className = 'tag-picker__remove';
            removeBtn.setAttribute('aria-label', `Remove ${value}`);
            removeBtn.dataset.value = value;
            removeBtn.innerHTML = '&times;';
            tag.appendChild(removeBtn);

            this.tagContainer.appendChild(tag);
        });
    }

    refreshDropdown() {
        if (!this.dropdown) {
            return;
        }
        const query = this.input ? this.input.value.trim().toLowerCase() : '';
        const available = this.options.filter(option => !this.selectedSet.has(option.toLowerCase()));
        let filtered = available;
        if (query) {
            filtered = available.filter(option => option.toLowerCase().includes(query));
        }
        this.filteredOptions = filtered.slice(0, 60);

        this.dropdown.innerHTML = '';
        if (this.filteredOptions.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'tag-picker__option tag-picker__option--empty';
            empty.textContent = query ? 'No matches found' : 'No options available';
            this.dropdown.appendChild(empty);
            return;
        }

        this.filteredOptions.forEach(option => {
            const optionElement = document.createElement('div');
            optionElement.className = 'tag-picker__option';
            optionElement.dataset.value = option;

            const title = document.createElement('span');
            title.className = 'tag-picker__option-title';
            title.textContent = option;
            optionElement.appendChild(title);

            this.dropdown.appendChild(optionElement);
        });
    }

    openDropdown() {
        if (!this.dropdown) {
            return;
        }
        this.dropdown.hidden = false;
    }

    closeDropdown() {
        if (!this.dropdown) {
            return;
        }
        this.dropdown.hidden = true;
    }

    updateHiddenInput() {
        if (!this.hiddenInput) {
            return;
        }
        try {
            this.hiddenInput.value = JSON.stringify(this.selected);
        } catch (error) {
            console.warn('Failed to serialize picker values', error);
            this.hiddenInput.value = this.selected.join(', ');
        }
    }

    focusInputSoon() {
        if (!this.input) {
            return;
        }
        requestAnimationFrame(() => {
            this.input.focus();
        });
    }
}

/**
 * SingleUserPicker - a picker for selecting a single employee (e.g. top-level user)
 * Similar to TagPicker but only allows one selection at a time.
 */
class SingleUserPicker {
    constructor({ pickerId, hiddenInputId, hiddenUserIdInputId, options = [], placeholder = '' }) {
        this.root = document.getElementById(pickerId);
        this.hiddenInput = document.getElementById(hiddenInputId);
        this.hiddenUserIdInput = hiddenUserIdInputId ? document.getElementById(hiddenUserIdInputId) : null;
        if (!this.root || !this.hiddenInput) {
            this.enabled = false;
            return;
        }

        this.enabled = true;
        // Options can be strings or objects with { label, email, id }
        this.options = [];
        this.optionsMap = new Map(); // email/label -> { label, email, id }
        this.setOptions(options);

        this.selectedContainer = this.root.querySelector('[data-role="selected-user"]');
        this.dropdown = this.root.querySelector('[data-role="dropdown"]');
        this.input = this.root.querySelector('.single-user-picker__input');
        if (placeholder && this.input) {
            this.input.placeholder = placeholder;
        }

        this.selectedValue = null; // { label, email, id }
        this.filteredOptions = [];

        this.handleDocumentClick = this.handleDocumentClick.bind(this);
        this.handleDropdownClick = this.handleDropdownClick.bind(this);
        this.handleRemoveClick = this.handleRemoveClick.bind(this);
        this.handleInput = this.handleInput.bind(this);
        this.handleKeyDown = this.handleKeyDown.bind(this);

        if (this.selectedContainer) {
            this.selectedContainer.addEventListener('click', this.handleRemoveClick);
        }
        if (this.dropdown) {
            this.dropdown.addEventListener('click', this.handleDropdownClick);
        }
        if (this.input) {
            this.input.addEventListener('input', this.handleInput);
            this.input.addEventListener('focus', () => this.openDropdown());
            this.input.addEventListener('keydown', this.handleKeyDown);
        }

        document.addEventListener('click', this.handleDocumentClick);
        this.renderSelected();
        this.closeDropdown();
        this.updateHiddenInputs();
    }

    destroy() {
        if (!this.enabled) return;
        document.removeEventListener('click', this.handleDocumentClick);
        this.enabled = false;
    }

    setOptions(options) {
        if (!this.enabled && this.root) return;
        this.options = [];
        this.optionsMap = new Map();

        (Array.isArray(options) ? options : []).forEach(opt => {
            if (!opt) return;
            let entry;
            if (typeof opt === 'object' && opt.label) {
                entry = { label: opt.label, email: opt.email || '', id: opt.id || '' };
            } else if (typeof opt === 'string') {
                // Parse "Name <email>" format
                const match = opt.match(/^(.+?)\s*<([^>]+)>$/);
                if (match) {
                    entry = { label: opt, email: match[2].trim(), id: '' };
                } else {
                    entry = { label: opt, email: opt, id: '' };
                }
            } else {
                return;
            }
            const key = (entry.email || entry.label).toLowerCase();
            if (!this.optionsMap.has(key)) {
                this.options.push(entry);
                this.optionsMap.set(key, entry);
            }
        });

        this.options.sort((a, b) => a.label.localeCompare(b.label, undefined, { sensitivity: 'base' }));
        this.refreshDropdown();
    }

    setValue(email, userId = '') {
        if (!this.enabled) return;
        const normalizedEmail = (email || '').trim().toLowerCase();
        if (!normalizedEmail) {
            this.selectedValue = null;
        } else {
            // Try to find in options map
            const found = this.optionsMap.get(normalizedEmail);
            if (found) {
                this.selectedValue = { ...found, id: userId || found.id };
            } else {
                // Create a simple entry
                this.selectedValue = { label: email, email: email, id: userId };
            }
        }
        this.renderSelected();
        this.updateHiddenInputs();
        if (this.input) {
            this.input.value = '';
        }
        this.closeDropdown();
    }

    getValue() {
        if (!this.enabled) return { email: '', id: '' };
        return this.selectedValue
            ? { email: this.selectedValue.email, id: this.selectedValue.id }
            : { email: '', id: '' };
    }

    clear() {
        this.setValue('');
        markUnsavedChange();
    }

    handleInput() {
        this.refreshDropdown();
        this.openDropdown();
    }

    handleKeyDown(event) {
        if ((event.key === 'Enter' || event.key === 'Tab') && this.input) {
            const query = this.input.value.trim();
            if (!query) return;
            if (this.filteredOptions.length > 0) {
                this.selectOption(this.filteredOptions[0]);
            }
            event.preventDefault();
        }
    }

    handleDropdownClick(event) {
        const option = event.target.closest('[data-email]');
        if (!option) return;
        event.preventDefault();
        event.stopPropagation();
        const email = option.getAttribute('data-email') || '';
        const id = option.getAttribute('data-id') || '';
        const label = option.getAttribute('data-label') || email;
        this.selectOption({ label, email, id });
    }

    handleRemoveClick(event) {
        const removeBtn = event.target.closest('.single-user-picker__remove');
        if (!removeBtn) return;
        this.clear();
    }

    handleDocumentClick(event) {
        if (!this.root) return;
        if (!this.root.contains(event.target)) {
            this.closeDropdown();
        }
    }

    selectOption(entry) {
        if (!this.enabled || !entry) return;
        this.selectedValue = { label: entry.label, email: entry.email, id: entry.id || '' };
        this.renderSelected();
        this.updateHiddenInputs();
        if (this.input) {
            this.input.value = '';
        }
        this.closeDropdown();
        markUnsavedChange();
    }

    renderSelected() {
        if (!this.selectedContainer) return;
        this.selectedContainer.innerHTML = '';
        if (!this.selectedValue) {
            if (this.input) this.input.style.display = '';
            return;
        }
        // Hide input when user is selected
        if (this.input) this.input.style.display = 'none';

        const tag = document.createElement('span');
        tag.className = 'single-user-picker__tag';

        const label = document.createElement('span');
        label.textContent = this.selectedValue.label;
        tag.appendChild(label);

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'single-user-picker__remove';
        removeBtn.setAttribute('aria-label', 'Remove selection');
        removeBtn.innerHTML = '&times;';
        tag.appendChild(removeBtn);

        this.selectedContainer.appendChild(tag);
    }

    refreshDropdown() {
        if (!this.dropdown) return;
        const query = this.input ? this.input.value.trim().toLowerCase() : '';

        let filtered = this.options;
        if (query) {
            filtered = this.options.filter(opt =>
                opt.label.toLowerCase().includes(query) ||
                opt.email.toLowerCase().includes(query)
            );
        }
        this.filteredOptions = filtered.slice(0, 60);

        this.dropdown.innerHTML = '';
        if (this.filteredOptions.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'single-user-picker__option single-user-picker__option--empty';
            empty.textContent = query ? 'No matches found' : 'No options available';
            this.dropdown.appendChild(empty);
            return;
        }

        this.filteredOptions.forEach(opt => {
            const optionElement = document.createElement('div');
            optionElement.className = 'single-user-picker__option';
            optionElement.dataset.email = opt.email;
            optionElement.dataset.id = opt.id || '';
            optionElement.dataset.label = opt.label;

            const title = document.createElement('span');
            title.className = 'single-user-picker__option-title';
            title.textContent = opt.label;
            optionElement.appendChild(title);

            this.dropdown.appendChild(optionElement);
        });
    }

    openDropdown() {
        if (!this.dropdown) return;
        this.dropdown.hidden = false;
    }

    closeDropdown() {
        if (!this.dropdown) return;
        this.dropdown.hidden = true;
    }

    updateHiddenInputs() {
        if (this.hiddenInput) {
            this.hiddenInput.value = this.selectedValue ? this.selectedValue.email : '';
        }
        if (this.hiddenUserIdInput) {
            this.hiddenUserIdInput.value = this.selectedValue ? (this.selectedValue.id || '') : '';
        }
    }
}

async function loadFilterMetadata() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/metadata/options`, { credentials: 'same-origin' });
        if (!response.ok) {
            return { jobTitles: [], departments: [] };
        }
        const data = await response.json();
        return {
            jobTitles: Array.isArray(data.jobTitles) ? data.jobTitles : [],
            departments: Array.isArray(data.departments) ? data.departments : [],
            employees: Array.isArray(data.employees) ? data.employees : []
        };
    } catch (error) {
        console.error('Failed to load filter metadata', error);
        return { jobTitles: [], departments: [] };
    }
}

function initializeTagPickers(metadata) {
    tagPickers.ignoredTitles = new TagPicker({
        pickerId: 'ignoredTitlesPicker',
        hiddenInputId: 'ignoredTitlesInput',
        options: metadata.jobTitles,
        placeholder: document.getElementById('ignoredTitlesSearch')?.getAttribute('placeholder') || ''
    });

    tagPickers.ignoredDepartments = new TagPicker({
        pickerId: 'ignoredDepartmentsPicker',
        hiddenInputId: 'ignoredDepartmentsInput',
        options: metadata.departments,
        placeholder: document.getElementById('ignoredDepartmentsSearch')?.getAttribute('placeholder') || ''
    });

    tagPickers.ignoredEmployees = new TagPicker({
        pickerId: 'ignoredEmployeesPicker',
        hiddenInputId: 'ignoredEmployeesInput',
        options: metadata.employees,
        placeholder: document.getElementById('ignoredEmployeesSearch')?.getAttribute('placeholder') || ''
    });

    // Initialize single user picker for top-level user
    topLevelUserPicker = new SingleUserPicker({
        pickerId: 'topLevelUserPicker',
        hiddenInputId: 'topLevelUserInput',
        hiddenUserIdInputId: 'topLevelUserIdInput',
        options: metadata.employees,
        placeholder: document.getElementById('topLevelUserSearch')?.getAttribute('placeholder') || ''
    });
}

function getIgnoredTitlesValue() {
    const hidden = document.getElementById('ignoredTitlesInput');
    if (!hidden) {
        return '[]';
    }
    const value = hidden.value;
    if (value && value.trim()) {
        return value.trim();
    }
    return '[]';
}

function getIgnoredDepartmentsValue() {
    const hidden = document.getElementById('ignoredDepartmentsInput');
    if (!hidden) {
        return '[]';
    }
    const value = hidden.value;
    if (value && value.trim()) {
        return value.trim();
    }
    return '[]';
}

function getIgnoredEmployeesValue() {
    const hidden = document.getElementById('ignoredEmployeesInput');
    if (!hidden) {
        return '[]';
    }
    const value = hidden.value;
    if (value && value.trim()) {
        return value.trim();
    }
    return '[]';
}

async function loadSettings() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/settings`);
        if (response.ok) {
            currentSettings = await response.json();
            applySettings(currentSettings);
        }
    } catch (error) {
        console.error('Error loading settings:', error);
    }
}

function initTimePicker() {
    const hourSel = document.getElementById('updateHour');
    const minSel = document.getElementById('updateMinute');
    if (!hourSel || !minSel) return;
    if (!hourSel.options.length) {
        for (let h = 0; h < 24; h++) {
            const opt = document.createElement('option');
            opt.value = opt.textContent = h.toString().padStart(2, '0');
            hourSel.appendChild(opt);
        }
    }
    if (!minSel.options.length) {
        for (let m = 0; m < 60; m++) {
            const opt = document.createElement('option');
            opt.value = opt.textContent = m.toString().padStart(2, '0');
            minSel.appendChild(opt);
        }
    }
    function syncHidden() {
        const hidden = document.getElementById('updateTime');
        if (hidden) hidden.value = `${hourSel.value}:${minSel.value}`;
    }
    hourSel.addEventListener('change', syncHidden);
    minSel.addEventListener('change', syncHidden);
    const [defaultHour, defaultMinute] = DEFAULT_UPDATE_TIME.split(':');
    if (!hourSel.value) hourSel.value = defaultHour;
    if (!minSel.value) minSel.value = defaultMinute;
    syncHidden();
}

function initTimezonePicker() {
    const tzSelect = document.getElementById('updateTimezone');
    if (!tzSelect) {
        return;
    }

    let zones = [];
    if (typeof Intl !== 'undefined' && typeof Intl.supportedValuesOf === 'function') {
        try {
            zones = Intl.supportedValuesOf('timeZone');
        } catch (error) {
            console.warn('Failed to enumerate time zones via Intl.supportedValuesOf', error);
        }
    }

    if (!zones || zones.length === 0) {
        zones = [
            'UTC',
            'America/New_York',
            'America/Chicago',
            'America/Denver',
            'America/Los_Angeles',
            'Europe/London',
            'Europe/Berlin',
            'Asia/Tokyo',
            'Australia/Sydney'
        ];
    }

    const fragment = document.createDocumentFragment();
    zones.forEach(zone => {
        const option = document.createElement('option');
        option.value = zone;
        option.textContent = zone;
        fragment.appendChild(option);
    });

    tzSelect.innerHTML = '';
    tzSelect.appendChild(fragment);

    const desired = DEFAULT_UPDATE_TIMEZONE;
    if (zones.includes(desired)) {
        tzSelect.value = desired;
    } else if (zones.length) {
        tzSelect.value = zones[0];
    }
}

function applySettings(settings) {
    document.title = `Configuration - ${settings.chartTitle || 'SimpleOrgChart'}`;
    if (settings.chartTitle) {
        document.getElementById('chartTitle').value = settings.chartTitle;
    }

    const headerColorInput = document.getElementById('headerColor');
    const headerColorHexInput = document.getElementById('headerColorHex');
    const normalizedHeaderColor = applyColorToInputs(
        headerColorInput,
        headerColorHexInput,
        settings.headerColor,
        DEFAULT_HEADER_COLOR
    );
    updateHeaderPreview(normalizedHeaderColor);

    const logoPath = settings.logoPath || '/static/icon.png';
    document.getElementById('currentLogo').src = `${logoPath}?t=${Date.now()}`;
    const logoStatus = document.getElementById('logoStatus');
    if (logoStatus) {
        logoStatus.textContent = logoPath.includes('icon_custom_') ? 'Custom uploaded' : 'Using default';
    }

    const faviconPath = settings.faviconPath || '/favicon.ico';
    document.getElementById('currentFavicon').src = `${faviconPath}?t=${Date.now()}`;
    const favStatus = document.getElementById('faviconStatus');
    if (favStatus) {
        favStatus.textContent = faviconPath.includes('favicon_custom_') ? 'Custom uploaded' : 'Using default';
    }

    Object.keys(NODE_COLOR_DEFAULTS).forEach(level => {
        const colorInput = document.getElementById(`${level}Color`);
        const hexInput = document.getElementById(`${level}ColorHex`);
        const fallback = NODE_COLOR_DEFAULTS[level];
        const rawValue = (settings.nodeColors && settings.nodeColors[level]) || fallback;
        applyColorToInputs(colorInput, hexInput, rawValue, fallback);
    });

    if (settings.autoUpdateEnabled !== undefined) {
        document.getElementById('autoUpdateEnabled').checked = settings.autoUpdateEnabled;
    }

    if (settings.updateTime) {
        const [h, m] = settings.updateTime.split(':');
        const hourSel = document.getElementById('updateHour');
        const minSel = document.getElementById('updateMinute');
        if (hourSel && minSel) {
            hourSel.value = h.padStart(2, '0');
            minSel.value = m.padStart(2, '0');
        }
        const hidden = document.getElementById('updateTime');
        if (hidden) hidden.value = `${h.padStart(2, '0')}:${m.padStart(2, '0')}`;
    }

    const timezoneSelect = document.getElementById('updateTimezone');
    if (timezoneSelect) {
        const desiredTimezone = settings.updateTimezone || DEFAULT_UPDATE_TIMEZONE;
        if (!Array.from(timezoneSelect.options).some(option => option.value === desiredTimezone)) {
            const option = document.createElement('option');
            option.value = desiredTimezone;
            option.textContent = desiredTimezone;
            timezoneSelect.appendChild(option);
        }
        timezoneSelect.value = desiredTimezone;
    }

    if (settings.collapseLevel) {
        document.getElementById('collapseLevel').value = settings.collapseLevel;
    }

    if (settings.searchAutoExpand !== undefined) {
        document.getElementById('searchAutoExpand').checked = settings.searchAutoExpand;
    }

    if (settings.searchHighlight !== undefined) {
        document.getElementById('searchHighlight').checked = settings.searchHighlight;
    }

    if (settings.newEmployeeMonths !== undefined) {
        document.getElementById('newEmployeeMonths').value = settings.newEmployeeMonths;
    }

    if (settings.hideDisabledUsers !== undefined) {
        document.getElementById('hideDisabledUsers').checked = settings.hideDisabledUsers;
    }

    if (settings.hideGuestUsers !== undefined) {
        document.getElementById('hideGuestUsers').checked = settings.hideGuestUsers;
    }

    if (settings.hideNoTitle !== undefined) {
        document.getElementById('hideNoTitle').checked = settings.hideNoTitle;
    }

    if (settings.ignoredDepartments !== undefined) {
        const values = parseListString(settings.ignoredDepartments);
        if (tagPickers.ignoredDepartments && tagPickers.ignoredDepartments.setValue) {
            tagPickers.ignoredDepartments.setOptions(filterMetadata.departments || []);
            tagPickers.ignoredDepartments.setValue(values);
        } else {
            const el = document.getElementById('ignoredDepartmentsInput');
            if (el) el.value = values.join(', ');
        }
    }
    if (settings.ignoredTitles !== undefined) {
        const values = parseListString(settings.ignoredTitles);
        if (tagPickers.ignoredTitles && tagPickers.ignoredTitles.setValue) {
            tagPickers.ignoredTitles.setOptions(filterMetadata.jobTitles || []);
            tagPickers.ignoredTitles.setValue(values);
        } else {
            const el = document.getElementById('ignoredTitlesInput');
            if (el) el.value = values.join(', ');
        }
    }
    if (settings.ignoredEmployees !== undefined) {
        const values = parseListString(settings.ignoredEmployees);
        if (tagPickers.ignoredEmployees && tagPickers.ignoredEmployees.setValue) {
            tagPickers.ignoredEmployees.setOptions(filterMetadata.employees || []);
            tagPickers.ignoredEmployees.setValue(values);
        } else {
            const el = document.getElementById('ignoredEmployeesInput');
            if (el) el.value = values.join(', ');
        }
    }

    // Apply top-level user setting
    if (topLevelUserPicker && topLevelUserPicker.enabled) {
        topLevelUserPicker.setOptions(filterMetadata.employees || []);
        const email = settings.topLevelUserEmail || '';
        const userId = settings.topLevelUserId || '';
        topLevelUserPicker.setValue(email, userId);
    }

    if (settings.printOrientation) {
        document.getElementById('printOrientation').value = settings.printOrientation;
    }

    if (settings.printSize) {
        document.getElementById('printSize').value = settings.printSize;
    }

    if (settings.multiLineChildrenThreshold !== undefined) {
        const el2 = document.getElementById('multiLineChildrenThreshold');
        if (el2) el2.value = settings.multiLineChildrenThreshold;
    }

    const customContactsField = document.getElementById('customDirectoryContacts');
    if (customContactsField) {
        const rawValue = settings.customDirectoryContacts;
        if (typeof rawValue === 'string') {
            customContactsField.value = rawValue;
        } else if (Array.isArray(rawValue)) {
            customContactsField.value = rawValue.join('\n');
        } else {
            customContactsField.value = '';
        }
    }

    // Directory feed settings
    const jsonEnabledEl = document.getElementById('directoryJsonEnabled');
    if (jsonEnabledEl) {
        jsonEnabledEl.checked = settings.directoryJsonEnabled === true;
    }
    const jsonFilenameEl = document.getElementById('directoryJsonFilename');
    if (jsonFilenameEl) {
        jsonFilenameEl.value = settings.directoryJsonFilename || 'contacts';
    }
    const xmlEnabledEl = document.getElementById('directoryXmlEnabled');
    if (xmlEnabledEl) {
        xmlEnabledEl.checked = settings.directoryXmlEnabled === true;
    }
    const xmlFilenameEl = document.getElementById('directoryXmlFilename');
    if (xmlFilenameEl) {
        xmlFilenameEl.value = settings.directoryXmlFilename || 'contacts';
    }

    const updateMeta = settings.dataUpdateStatus || {};
    if (updateMeta.state === 'running') {
        updateLastUpdatedDisplay(resolveTranslation('configure.data.manualUpdate.lastUpdatedUpdating', 'Updating...'));
        setUpdateStatus('configure.data.manualUpdate.statusUpdating', 'Updating...');
        startUpdatePolling();
    } else {
        updateLastUpdatedDisplay(settings.dataLastUpdatedAt);
        if (updateMeta.success === false) {
            setUpdateStatus('configure.data.manualUpdate.statusFailed', '✗ Update failed');
        } else {
            setUpdateStatus(null, '');
        }
    }
    const statusEl = document.getElementById('updateStatus');
    if (statusEl) {
        statusEl.title = updateMeta.error || '';
    }

    applyExportColumnSettings(settings);
}

function applyExportColumnSettings(settings) {
    const modeFor = (key) => {
        const value = (settings.exportXlsxColumns || {})[key];
        if (value === 'hide' || value === 'admin') {
            return value;
        }
        return 'show';
    };

    EXPORT_COLUMN_KEYS.forEach(({ key, selectId }) => {
        const select = document.getElementById(selectId);
        if (select) {
            select.value = modeFor(key);
        }
    });
}

function getExportColumnSettings() {
    const result = {};
    EXPORT_COLUMN_KEYS.forEach(({ key, selectId }) => {
        const select = document.getElementById(selectId);
        if (select) {
            const value = select.value;
            result[key] = (value === 'hide' || value === 'admin') ? value : 'show';
        }
    });
    return result;
}

function resetExportColumns() {
    EXPORT_COLUMN_KEYS.forEach(({ key, selectId }) => {
        const select = document.getElementById(selectId);
        if (select) {
            select.value = EXPORT_COLUMN_DEFAULTS[key] || 'show';
        }
    });
    markUnsavedChange();
}

const headerColorInput = document.getElementById('headerColor');
const headerColorHexInput = document.getElementById('headerColorHex');
if (headerColorInput) {
    headerColorInput.addEventListener('input', event => {
        const normalized = applyColorToInputs(
            headerColorInput,
            headerColorHexInput,
            event.target.value,
            DEFAULT_HEADER_COLOR
        );
        updateHeaderPreview(normalized);
    });
}

if (headerColorHexInput) {
    headerColorHexInput.addEventListener('input', event => {
        if (/^#[0-9A-Fa-f]{6}$/.test(event.target.value)) {
            const normalized = applyColorToInputs(
                headerColorInput,
                headerColorHexInput,
                event.target.value,
                DEFAULT_HEADER_COLOR
            );
            headerColorHexInput.value = normalized;
            updateHeaderPreview(normalized);
        }
    });
}

Object.keys(NODE_COLOR_DEFAULTS).forEach(level => {
    const colorInput = document.getElementById(`${level}Color`);
    const hexInput = document.getElementById(`${level}ColorHex`);

    if (colorInput) {
        colorInput.addEventListener('input', event => {
            applyColorToInputs(colorInput, hexInput, event.target.value, NODE_COLOR_DEFAULTS[level]);
        });
    }

    if (hexInput) {
        hexInput.addEventListener('input', event => {
            if (/^#[0-9A-Fa-f]{6}$/.test(event.target.value)) {
                const normalized = applyColorToInputs(colorInput, hexInput, event.target.value, NODE_COLOR_DEFAULTS[level]);
                hexInput.value = normalized;
            }
        });
    }
});

function updateHeaderPreview(color) {
    const normalized = normalizeHexColor(color, DEFAULT_HEADER_COLOR);
    const darker = adjustColor(normalized, -30);
    const stylesheet = Array.from(document.styleSheets).find(sheet => {
        try {
            return sheet.href && sheet.href.includes('configure.css');
        } catch (error) {
            return false;
        }
    });

    if (!stylesheet) {
        return;
    }

    let rootRule = Array.from(stylesheet.cssRules).find(rule => rule.selectorText === ':root');
    if (!rootRule) {
        const index = stylesheet.cssRules.length;
        stylesheet.insertRule(':root {}', index);
        rootRule = stylesheet.cssRules[index];
    }

    rootRule.style.setProperty('--header-preview-primary', normalized);
    rootRule.style.setProperty('--header-preview-secondary', darker);
}

function adjustColor(color, amount) {
    const num = parseInt(color.replace('#', ''), 16);
    const r = Math.max(0, Math.min(255, (num >> 16) + amount));
    const g = Math.max(0, Math.min(255, ((num >> 8) & 0x00ff) + amount));
    const b = Math.max(0, Math.min(255, (num & 0x0000ff) + amount));
    return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, '0').toUpperCase()}`;
}

async function uploadLogoFile(file) {
    if (!file) return;
    const maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) {
        showStatus('File too large. Maximum size is 5MB.', 'error');
        return;
    }
    const allowedTypes = ['image/png', 'image/jpeg', 'image/jpg'];
    if (!allowedTypes.includes(file.type)) {
        showStatus('Invalid file type. Please use PNG, JPG, or JPEG.', 'error');
        return;
    }
    showStatus('Uploading logo...', 'info');
    try {
        const authCheck = await fetch(`${API_BASE_URL}/api/settings`);
        if (!authCheck.ok && authCheck.status === 401) {
            showStatus('Session expired. Please refresh the page and log in again.', 'error');
            return;
        }
    } catch {
        // ignore auth check errors
    }
    const formData = new FormData();
    formData.append('logo', file);
    try {
        const response = await fetch(`${API_BASE_URL}/api/upload-logo`, { method: 'POST', body: formData });
        if (response.ok) {
            const result = await response.json();
            const logoImg = document.getElementById('currentLogo');
            const newLogoPath = `${result.path}?t=${Date.now()}`;
            logoImg.onload = () => {
                showStatus('Logo uploaded successfully!', 'success');
                const status = document.getElementById('logoStatus');
                if (status) status.textContent = 'Custom uploaded';
            };
            logoImg.onerror = () => {
                console.error('Logo failed to load from path:', newLogoPath);
                showStatus('Logo uploaded but failed to display. Try refreshing the page.', 'warning');
            };
            logoImg.src = newLogoPath;
            pendingLogoReset = false;
            clearUnsavedReason('logoReset');
        } else {
            let errorMessage = 'Unknown error';
            try {
                const errorData = await response.json();
                errorMessage = errorData.error || errorData.message || 'Server error';
            } catch {
                const errorText = await response.text();
                errorMessage = errorText || `HTTP ${response.status} error`;
            }
            if (response.status === 401) showStatus('Authentication required. Please log in again.', 'error');
            else if (response.status === 413) showStatus('File too large. Please use a smaller image.', 'error');
            else if (response.status === 429) showStatus('Too many upload attempts. Please wait a moment.', 'error');
            else showStatus(errorMessage, 'error');
        }
    } catch (error) {
        console.error('Upload error:', error);
        showStatus('Network error uploading logo. Please try again.', 'error');
    }
}

const logoUploadInput = document.getElementById('logoUpload');
if (logoUploadInput) {
    logoUploadInput.addEventListener('change', async event => {
        const file = event.target.files[0];
        await uploadLogoFile(file);
        event.target.value = '';
    });
}

async function uploadFaviconFile(file) {
    if (!file) return;
    const maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) {
        showStatus('File too large. Maximum size is 5MB.', 'error');
        return;
    }
    const allowedTypes = ['image/x-icon', 'image/png', 'image/jpeg', 'image/jpg'];
    if (!allowedTypes.includes(file.type) && !file.name.toLowerCase().endsWith('.ico')) {
        showStatus('Invalid file type. Please use ICO, PNG, JPG, or JPEG.', 'error');
        return;
    }
    showStatus('Uploading favicon...', 'info');
    try {
        const formData = new FormData();
        formData.append('favicon', file);
        const response = await fetch(`${API_BASE_URL}/api/upload-favicon`, { method: 'POST', body: formData });
        if (response.ok) {
            const result = await response.json();
            const faviconImg = document.getElementById('currentFavicon');
            const newFaviconPath = `${result.path}?t=${Date.now()}`;
            faviconImg.onload = () => {
                showStatus('Favicon uploaded successfully!', 'success');
                updatePageFavicon(newFaviconPath);
                const status = document.getElementById('faviconStatus');
                if (status) status.textContent = 'Custom uploaded';
            };
            faviconImg.onerror = () => {
                showStatus('Favicon uploaded but failed to display.', 'warning');
            };
            faviconImg.src = newFaviconPath;
            pendingFaviconReset = false;
            clearUnsavedReason('faviconReset');
        } else {
            let errorMessage = 'Unknown error';
            try {
                const errorData = await response.json();
                errorMessage = errorData.error || errorData.message || 'Server error';
            } catch {
                const errorText = await response.text();
                errorMessage = errorText || `HTTP ${response.status} error`;
            }
            showStatus(errorMessage, 'error');
        }
    } catch (error) {
        console.error('Favicon upload error:', error);
        showStatus('Network error uploading favicon. Please try again.', 'error');
    }
}

const faviconUploadInput = document.getElementById('faviconUpload');
if (faviconUploadInput) {
    faviconUploadInput.addEventListener('change', async event => {
        const file = event.target.files[0];
        await uploadFaviconFile(file);
        event.target.value = '';
    });
}

function updatePageFavicon(faviconPath) {
    let favicon = document.querySelector('link[rel="icon"]') || document.querySelector('link[rel="shortcut icon"]');
    if (!favicon) {
        favicon = document.createElement('link');
        favicon.rel = 'icon';
        document.head.appendChild(favicon);
    }
    favicon.href = faviconPath;
}

function resetFavicon() {
    if (!confirm('Are you sure you want to reset the favicon to default?')) {
        return;
    }

    const defaultPath = '/favicon.ico';
    const previewPath = `${defaultPath}?t=${Date.now()}`;
    document.getElementById('currentFavicon').src = previewPath;
    updatePageFavicon(defaultPath);
    const status = document.getElementById('faviconStatus');
    if (status) {
        status.textContent = 'Reverting to default after save';
    }
    pendingFaviconReset = true;
    markUnsavedChange('faviconReset');
}

function resetChartTitle() {
    document.getElementById('chartTitle').value = 'Organization Chart';
    markUnsavedChange();
}

function resetHeaderColor() {
    const colorInput = document.getElementById('headerColor');
    const hexInput = document.getElementById('headerColorHex');
    const normalized = applyColorToInputs(colorInput, hexInput, DEFAULT_HEADER_COLOR, DEFAULT_HEADER_COLOR);
    updateHeaderPreview(normalized);
    markUnsavedChange();
}

function resetLogo() {
    const defaultPath = '/static/icon.png';
    document.getElementById('currentLogo').src = `${defaultPath}?t=${Date.now()}`;
    const status = document.getElementById('logoStatus');
    if (status) {
        status.textContent = 'Reverting to default after save';
    }
    pendingLogoReset = true;
    markUnsavedChange('logoReset');
}

function resetNodeColors() {
    Object.entries(NODE_COLOR_DEFAULTS).forEach(([level, value]) => {
        applyColorToInputs(
            document.getElementById(`${level}Color`),
            document.getElementById(`${level}ColorHex`),
            value,
            value
        );
    });
    markUnsavedChange();
}

function resetUpdateTime() {
    const hourSel = document.getElementById('updateHour');
    const minSel = document.getElementById('updateMinute');
    const [defaultHour, defaultMinute] = DEFAULT_UPDATE_TIME.split(':');
    if (hourSel) hourSel.value = defaultHour;
    if (minSel) minSel.value = defaultMinute;
    const hidden = document.getElementById('updateTime');
    if (hidden) hidden.value = DEFAULT_UPDATE_TIME;
    const tzSelect = document.getElementById('updateTimezone');
    if (tzSelect) {
        tzSelect.value = DEFAULT_UPDATE_TIMEZONE;
    }
    document.getElementById('autoUpdateEnabled').checked = true;
    markUnsavedChange();
}

function resetCollapseLevel() {
    document.getElementById('collapseLevel').value = '2';
    markUnsavedChange();
}

function resetMultiLineSettings() {
    const thresholdEl = document.getElementById('multiLineChildrenThreshold');
    if (thresholdEl) thresholdEl.value = 20;
    markUnsavedChange();
}

function resetIgnoredDepartments() {
    if (tagPickers.ignoredDepartments && typeof tagPickers.ignoredDepartments.clear === 'function') {
        tagPickers.ignoredDepartments.clear();
    } else {
        const el = document.getElementById('ignoredDepartmentsInput');
        if (el) el.value = '';
    }
    markUnsavedChange();
}

function resetIgnoredTitles() {
    if (tagPickers.ignoredTitles && typeof tagPickers.ignoredTitles.clear === 'function') {
        tagPickers.ignoredTitles.clear();
    } else {
        const el = document.getElementById('ignoredTitlesInput');
        if (el) el.value = '';
    }
    markUnsavedChange();
}

function resetIgnoredEmployees() {
    if (tagPickers.ignoredEmployees && typeof tagPickers.ignoredEmployees.clear === 'function') {
        tagPickers.ignoredEmployees.clear();
    } else {
        const el = document.getElementById('ignoredEmployeesInput');
        if (el) el.value = '';
    }
    markUnsavedChange();
}

function resetTopLevelUser() {
    if (topLevelUserPicker && typeof topLevelUserPicker.clear === 'function') {
        topLevelUserPicker.clear();
    } else {
        const emailEl = document.getElementById('topLevelUserInput');
        const idEl = document.getElementById('topLevelUserIdInput');
        if (emailEl) emailEl.value = '';
        if (idEl) idEl.value = '';
    }
    markUnsavedChange();
}

function resetCustomDirectoryContacts() {
    const textarea = document.getElementById('customDirectoryContacts');
    if (!textarea) {
        return;
    }
    textarea.value = '';
    markUnsavedChange();
}

function resetAllSettings() {
    resetChartTitle();
    resetHeaderColor();
    resetLogo();
    resetNodeColors();
    resetUpdateTime();
    resetCollapseLevel();
    resetMultiLineSettings();
    resetIgnoredDepartments();
    resetIgnoredTitles();
    resetIgnoredEmployees();
    resetExportColumns();
    resetCustomDirectoryContacts();

    document.getElementById('searchAutoExpand').checked = true;
    document.getElementById('searchHighlight').checked = true;
    document.getElementById('newEmployeeMonths').value = '3';
    document.getElementById('hideDisabledUsers').checked = true;
    document.getElementById('hideGuestUsers').checked = true;
    document.getElementById('hideNoTitle').checked = true;
    document.getElementById('printOrientation').value = 'landscape';
    document.getElementById('printSize').value = 'a4';

    markUnsavedChange('resetAll');
    showStatus('Default settings applied. Save to confirm.', 'info');
}

async function logout() {
    try {
        const response = await fetch(`${API_BASE_URL}/logout`, {
            method: 'POST',
            credentials: 'include'
        });

        if (response.ok) {
            window.location.href = '/';
        } else {
            sessionStorage.clear();
            localStorage.clear();
            window.location.href = '/';
        }
    } catch (error) {
        console.error('Logout error:', error);
        window.location.href = '/';
    }
}

async function saveAllSettings() {
    const logoResetRequested = pendingLogoReset;
    const faviconResetRequested = pendingFaviconReset;

    const resolveColorValue = (colorInputId, hexInputId, fallback) => {
        const hexInput = document.getElementById(hexInputId);
        if (hexInput && /^#[0-9A-Fa-f]{6}$/.test(hexInput.value)) {
            return normalizeHexColor(hexInput.value, fallback);
        }
        const colorInput = document.getElementById(colorInputId);
        return normalizeHexColor(colorInput ? colorInput.value : '', fallback);
    };

    const settings = {
        chartTitle: document.getElementById('chartTitle').value || 'Organization Chart',
        headerColor: resolveColorValue('headerColor', 'headerColorHex', DEFAULT_HEADER_COLOR),
        nodeColors: Object.keys(NODE_COLOR_DEFAULTS).reduce((accumulator, level) => {
            accumulator[level] = resolveColorValue(
                `${level}Color`,
                `${level}ColorHex`,
                NODE_COLOR_DEFAULTS[level]
            );
            return accumulator;
        }, {}),
        autoUpdateEnabled: document.getElementById('autoUpdateEnabled').checked,
        updateTime: document.getElementById('updateTime').value,
    updateTimezone: document.getElementById('updateTimezone').value,
        collapseLevel: document.getElementById('collapseLevel').value,
        searchAutoExpand: document.getElementById('searchAutoExpand').checked,
        searchHighlight: document.getElementById('searchHighlight').checked,
        newEmployeeMonths: parseInt(document.getElementById('newEmployeeMonths').value, 10),
        hideDisabledUsers: document.getElementById('hideDisabledUsers').checked,
        hideGuestUsers: document.getElementById('hideGuestUsers').checked,
        hideNoTitle: document.getElementById('hideNoTitle').checked,
        ignoredDepartments: getIgnoredDepartmentsValue(),
        ignoredTitles: getIgnoredTitlesValue(),
    ignoredEmployees: getIgnoredEmployeesValue(),
        printOrientation: document.getElementById('printOrientation').value,
        printSize: document.getElementById('printSize').value,
        multiLineChildrenThreshold: parseInt(document.getElementById('multiLineChildrenThreshold')?.value || '20', 10),
        topLevelUserEmail: document.getElementById('topLevelUserInput')?.value || '',
        topLevelUserId: document.getElementById('topLevelUserIdInput')?.value || '',
        exportXlsxColumns: getExportColumnSettings(),
        customDirectoryContacts: (document.getElementById('customDirectoryContacts')?.value || '').trim(),
        directoryJsonEnabled: document.getElementById('directoryJsonEnabled')?.checked || false,
        directoryJsonFilename: (document.getElementById('directoryJsonFilename')?.value || 'contacts').trim().replace(/[^a-zA-Z0-9_-]/g, '') || 'contacts',
        directoryXmlEnabled: document.getElementById('directoryXmlEnabled')?.checked || false,
        directoryXmlFilename: (document.getElementById('directoryXmlFilename')?.value || 'contacts').trim().replace(/[^a-zA-Z0-9_-]/g, '') || 'contacts'
    };

    try {
        if (logoResetRequested) {
            const response = await fetch(`${API_BASE_URL}/api/reset-logo`, { method: 'POST' });
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                const message = errorData.error || `Error resetting logo (Status: ${response.status})`;
                showStatus(message, 'error');
                return;
            }
            pendingLogoReset = false;
        }

        if (faviconResetRequested) {
            const response = await fetch(`${API_BASE_URL}/api/reset-favicon`, { method: 'POST' });
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                const message = errorData.error || `Error resetting favicon (Status: ${response.status})`;
                showStatus(message, 'error');
                return;
            }
            pendingFaviconReset = false;
        }

        const response = await fetch(`${API_BASE_URL}/api/settings`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(settings)
        });

        if (response.ok) {
            showStatus('Settings saved successfully!', 'success');
            isInitializing = true;
            await loadSettings();
            isInitializing = false;
            clearUnsavedChangeState();
            if (logoResetRequested) {
                const status = document.getElementById('logoStatus');
                if (status) {
                    status.textContent = 'Using default';
                }
            }
            if (faviconResetRequested) {
                const status = document.getElementById('faviconStatus');
                if (status) {
                    status.textContent = 'Using default';
                }
            }
        } else {
            const errorData = await response.json().catch(() => ({}));
            const errorMsg = errorData.error || `Error saving settings (Status: ${response.status})`;
            console.error('Settings save error:', errorData);
            showStatus(errorMsg, 'error');
        }
    } catch (error) {
        console.error('Network error saving settings:', error);
        showStatus(`Error saving settings: ${error.message}`, 'error');
    }
}

let _updatePollingInterval = null;

function stopUpdatePolling() {
    if (_updatePollingInterval) {
        clearInterval(_updatePollingInterval);
        _updatePollingInterval = null;
    }
}

async function pollUpdateStatus() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/settings`);
        if (!response.ok) {
            return;
        }
        const settings = await response.json();
        const status = settings.dataUpdateStatus || {};

        if (status.state !== 'running') {
            stopUpdatePolling();

            if (status.success === true) {
                setUpdateStatus('configure.data.manualUpdate.statusComplete', '✔ Update complete');
                setTimeout(() => {
                    setUpdateStatus(null, '');
                }, 5000);
            } else if (status.success === false) {
                setUpdateStatus('configure.data.manualUpdate.statusFailed', '✗ Update failed');
                const statusEl = document.getElementById('updateStatus');
                if (statusEl) {
                    statusEl.title = status.error || '';
                }
            } else {
                setUpdateStatus(null, '');
            }

            updateLastUpdatedDisplay(settings.dataLastUpdatedAt);
            currentSettings = settings;
        }
    } catch (error) {
        console.warn('Failed to poll update status:', error);
    }
}

function startUpdatePolling() {
    stopUpdatePolling();
    _updatePollingInterval = setInterval(pollUpdateStatus, 3000);
}

async function triggerUpdate() {
    const previousTimestamp = currentSettings?.dataLastUpdatedAt || null;
    setUpdateStatus('configure.data.manualUpdate.statusUpdating', 'Updating...');
    updateLastUpdatedDisplay(resolveTranslation('configure.data.manualUpdate.lastUpdatedUpdating', 'Updating...'));
    if (!currentSettings || typeof currentSettings !== 'object') {
        currentSettings = {};
    }
    currentSettings.dataUpdateStatus = {
        state: 'running',
        startedAt: new Date().toISOString(),
    };

    try {
        const response = await fetch(`${API_BASE_URL}/api/update-now`, { method: 'POST' });

        if (response.ok) {
            setUpdateStatus('configure.data.manualUpdate.statusUpdating', 'Updating...');
            startUpdatePolling();
        } else if (response.status === 409) {
            setUpdateStatus('configure.data.manualUpdate.statusInProgress', 'Update already in progress');
            updateLastUpdatedDisplay(resolveTranslation('configure.data.manualUpdate.lastUpdatedUpdating', 'Updating...'));
            const statusEl = document.getElementById('updateStatus');
            if (statusEl) {
                statusEl.title = resolveTranslation('configure.data.manualUpdate.statusInProgress', 'Update already in progress');
            }
            startUpdatePolling();
        } else {
            setUpdateStatus('configure.data.manualUpdate.statusFailed', '✗ Update failed');
            const errorData = await response.json().catch(() => ({}));
            const detail = errorData.error || '';
            if (detail) {
                const statusEl = document.getElementById('updateStatus');
                if (statusEl) {
                    statusEl.title = detail;
                }
            }
            updateLastUpdatedDisplay(previousTimestamp);
        }
    } catch (error) {
        setUpdateStatus('configure.data.manualUpdate.statusFailed', '✗ Update failed');
        const statusEl = document.getElementById('updateStatus');
        if (statusEl) {
            statusEl.title = error?.message || '';
        }
        updateLastUpdatedDisplay(previousTimestamp);
    }
}

function showStatus(message, type) {
    const statusEl = document.getElementById('statusMessage');
    statusEl.textContent = message;
    statusEl.className = `status-message ${type}`;
    setTimeout(() => {
        statusEl.className = 'status-message';
    }, 3000);
}

function registerConfigActions() {
    const actionHandlers = {
        'reset-chart-title': resetChartTitle,
        'reset-header-color': resetHeaderColor,
        'reset-logo': resetLogo,
        'reset-favicon': resetFavicon,
        'reset-node-colors': resetNodeColors,
        'reset-update-time': resetUpdateTime,
        'reset-collapse-level': resetCollapseLevel,
        'reset-ignored-titles': resetIgnoredTitles,
        'reset-ignored-departments': resetIgnoredDepartments,
    'reset-ignored-employees': resetIgnoredEmployees,
        'reset-top-level-user': resetTopLevelUser,
        'reset-custom-directory-contacts': resetCustomDirectoryContacts,
        'reset-multiline-settings': resetMultiLineSettings,
        'reset-export-columns': resetExportColumns,
        'trigger-update': triggerUpdate,
        'discard-unsaved': discardUnsavedChanges,
        'save-all': saveAllSettings,
        'reset-all': resetAllSettings,
        'logout': logout
    };

    document.querySelectorAll('[data-config-action]').forEach(button => {
        const handler = actionHandlers[button.dataset.configAction];
        if (typeof handler === 'function') {
            button.addEventListener('click', handler);
        }
    });
}

document.addEventListener('DOMContentLoaded', async () => {
    initTimePicker();
    initTimezonePicker();
    registerConfigActions();
    attachUnsavedListeners();
    registerNavigationGuards();

    const metadata = await loadFilterMetadata();
    filterMetadata = metadata || { jobTitles: [], departments: [], employees: [] };
    initializeTagPickers(filterMetadata);
    await loadSettings();

    isInitializing = false;
    clearUnsavedChangeState();

    requestAnimationFrame(() => {
        document.body.classList.remove('is-loading');
    });

    const logoDropZone = document.getElementById('logoDropZone');
    if (logoDropZone) {
        ['dragenter', 'dragover'].forEach(evt => logoDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            logoDropZone.classList.add('dragover');
        }));
        ['dragleave', 'drop'].forEach(evt => logoDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            logoDropZone.classList.remove('dragover');
        }));
        logoDropZone.addEventListener('drop', event => {
            const file = event.dataTransfer.files && event.dataTransfer.files[0];
            if (file) uploadLogoFile(file);
        });
    }

    const faviconDropZone = document.getElementById('faviconDropZone');
    if (faviconDropZone) {
        ['dragenter', 'dragover'].forEach(evt => faviconDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            faviconDropZone.classList.add('dragover');
        }));
        ['dragleave', 'drop'].forEach(evt => faviconDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            faviconDropZone.classList.remove('dragover');
        }));
        faviconDropZone.addEventListener('drop', event => {
            const file = event.dataTransfer.files && event.dataTransfer.files[0];
            if (file) uploadFaviconFile(file);
        });
    }
});
