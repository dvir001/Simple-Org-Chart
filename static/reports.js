const API_BASE_URL = window.location.origin;
let currentReportKey = 'missing-manager';
let latestRecords = [];

const FILTER_REASON_I18N_KEYS = {
    filter_disabled: 'reports.types.filteredLicensed.reason.disabled',
    filter_guest: 'reports.types.filteredLicensed.reason.guest',
    filter_no_title: 'reports.types.filteredLicensed.reason.noTitle',
    filter_ignored_title: 'reports.types.filteredLicensed.reason.ignoredTitle',
    filter_ignored_department: 'reports.types.filteredLicensed.reason.ignoredDepartment',
    filter_ignored_employee: 'reports.types.filteredLicensed.reason.ignoredEmployee',
};

let selectedScanUser = null;
let userScannerEnabled = false;
let scannerSiteFilter = new Set();          // selected site names (empty = all)
let scannerKnownSites = [];                 // populated after first scan
let siteFilterPicker = null;                // TagPicker-style instance (single tab)
let categoryFilterPicker = null;            // TagPicker-style instance for categories (single tab)
let fullScanSiteFilterPicker = null;        // TagPicker-style instance (all tab)
let fullScanCategoryFilterPicker = null;    // TagPicker-style instance for categories (all tab)
let scannerLoudSites = new Set();           // loud sites set, populated after fetch

/** State for the full-scan user-level filters (mirrors last-logins filter toggles). */
const fullScanFilterState = {
    includeUserMailboxes: true,
    includeSharedMailboxes: false,
    includeRoomEquipmentMailboxes: false,
    includeEnabled: true,
    includeDisabled: false,
    includeLicensed: true,
    includeUnlicensed: false,
    includeMembers: true,
    includeGuests: false,
};

/** Filter definitions for the full-scan user filter panel. */
const FULL_SCAN_FILTER_GROUPS = [
    {
        labelKey: 'reports.filters.groups.mailboxTypes',
        filters: [
            { key: 'includeUserMailboxes', labelKey: 'reports.filters.includeUserMailboxes.label' },
            { key: 'includeSharedMailboxes', labelKey: 'reports.filters.includeSharedMailboxes.label' },
            { key: 'includeRoomEquipmentMailboxes', labelKey: 'reports.filters.includeRoomEquipmentMailboxes.label' },
        ],
    },
    {
        labelKey: 'reports.filters.groups.accountStatus',
        filters: [
            { key: 'includeEnabled', labelKey: 'reports.filters.includeEnabled.label' },
            { key: 'includeDisabled', labelKey: 'reports.filters.includeDisabled.label' },
        ],
    },
    {
        labelKey: 'reports.filters.groups.licenseStatus',
        filters: [
            { key: 'includeLicensed', labelKey: 'reports.filters.includeLicensed.label' },
            { key: 'includeUnlicensed', labelKey: 'reports.filters.includeUnlicensed.label' },
        ],
    },
    {
        labelKey: 'reports.filters.groups.userScope',
        filters: [
            { key: 'includeMembers', labelKey: 'reports.filters.includeMembers.label' },
            { key: 'includeGuests', labelKey: 'reports.filters.includeGuests.label' },
        ],
    },
];

const REPORT_CONFIGS = {
    'missing-manager': {
        dataPath: '/api/reports/missing-manager',
        exportPath: '/api/reports/missing-manager/export',
        summaryLabelKey: 'reports.summary.totalLabel',
        tableTitleKey: 'reports.table.title',
        emptyKey: 'reports.table.empty',
        countSummaryKey: 'reports.table.countSummary',
        buildStatusParams: (records) => ({ count: records.length }),
        filters: [
            {
                type: 'toggle',
                key: 'includeUserMailboxes',
                labelKey: 'reports.filters.includeUserMailboxes.label',
                queryParam: 'includeUserMailboxes',
                default: true,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeSharedMailboxes',
                labelKey: 'reports.filters.includeSharedMailboxes.label',
                queryParam: 'includeSharedMailboxes',
                default: false,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeRoomEquipmentMailboxes',
                labelKey: 'reports.filters.includeRoomEquipmentMailboxes.label',
                queryParam: 'includeRoomEquipmentMailboxes',
                default: false,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeEnabled',
                labelKey: 'reports.filters.includeEnabled.label',
                queryParam: 'includeEnabled',
                default: true,
                groupId: 'accountStatus',
                groupLabelKey: 'reports.filters.groups.accountStatus',
            },
            {
                type: 'toggle',
                key: 'includeDisabled',
                labelKey: 'reports.filters.includeDisabled.label',
                queryParam: 'includeDisabled',
                default: false,
                groupId: 'accountStatus',
                groupLabelKey: 'reports.filters.groups.accountStatus',
            },
            {
                type: 'toggle',
                key: 'includeLicensed',
                labelKey: 'reports.filters.includeLicensed.label',
                queryParam: 'includeLicensed',
                default: true,
                groupId: 'licenseStatus',
                groupLabelKey: 'reports.filters.groups.licenseStatus',
            },
            {
                type: 'toggle',
                key: 'includeUnlicensed',
                labelKey: 'reports.filters.includeUnlicensed.label',
                queryParam: 'includeUnlicensed',
                default: true,
                groupId: 'licenseStatus',
                groupLabelKey: 'reports.filters.groups.licenseStatus',
            },
            {
                type: 'toggle',
                key: 'includeMembers',
                labelKey: 'reports.filters.includeMembers.label',
                queryParam: 'includeMembers',
                default: true,
                groupId: 'userScope',
                groupLabelKey: 'reports.filters.groups.userScope',
            },
            {
                type: 'toggle',
                key: 'includeGuests',
                labelKey: 'reports.filters.includeGuests.label',
                queryParam: 'includeGuests',
                default: false,
                groupId: 'userScope',
                groupLabelKey: 'reports.filters.groups.userScope',
            },
        ],
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            { key: 'managerName', labelKey: 'reports.table.columns.manager' },
            {
                key: 'reason',
                labelKey: 'reports.table.columns.reason',
                render: (record, t) => createReasonBadge(record.reason, t),
            },
        ],
    },
    'last-logins': {
        dataPath: '/api/reports/last-logins',
        exportPath: '/api/reports/last-logins/export',
        summaryLabelKey: 'reports.types.lastLogins.summaryLabel',
        tableTitleKey: 'reports.types.lastLogins.tableTitle',
        emptyKey: 'reports.types.lastLogins.empty',
        countSummaryKey: 'reports.types.lastLogins.countSummary',
        showLicenseSummary: true,
        licenseSummaryLabelKey: 'reports.summary.licensesLabel',
        filters: [
            {
                type: 'toggle',
                key: 'includeUserMailboxes',
                labelKey: 'reports.filters.includeUserMailboxes.label',
                queryParam: 'includeUserMailboxes',
                default: true,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeSharedMailboxes',
                labelKey: 'reports.filters.includeSharedMailboxes.label',
                queryParam: 'includeSharedMailboxes',
                default: true,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeRoomEquipmentMailboxes',
                labelKey: 'reports.filters.includeRoomEquipmentMailboxes.label',
                queryParam: 'includeRoomEquipmentMailboxes',
                default: true,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeEnabled',
                labelKey: 'reports.filters.includeEnabled.label',
                queryParam: 'includeEnabled',
                default: true,
                groupId: 'accountStatus',
                groupLabelKey: 'reports.filters.groups.accountStatus',
            },
            {
                type: 'toggle',
                key: 'includeDisabled',
                labelKey: 'reports.filters.includeDisabled.label',
                queryParam: 'includeDisabled',
                default: true,
                groupId: 'accountStatus',
                groupLabelKey: 'reports.filters.groups.accountStatus',
            },
            {
                type: 'toggle',
                key: 'includeLicensed',
                labelKey: 'reports.filters.includeLicensed.label',
                queryParam: 'includeLicensed',
                default: true,
                groupId: 'licenseStatus',
                groupLabelKey: 'reports.filters.groups.licenseStatus',
            },
            {
                type: 'toggle',
                key: 'includeUnlicensed',
                labelKey: 'reports.filters.includeUnlicensed.label',
                queryParam: 'includeUnlicensed',
                default: true,
                groupId: 'licenseStatus',
                groupLabelKey: 'reports.filters.groups.licenseStatus',
            },
            {
                type: 'toggle',
                key: 'includeMembers',
                labelKey: 'reports.filters.includeMembers.label',
                queryParam: 'includeMembers',
                default: true,
                groupId: 'userScope',
                groupLabelKey: 'reports.filters.groups.userScope',
            },
            {
                type: 'toggle',
                key: 'includeGuests',
                labelKey: 'reports.filters.includeGuests.label',
                queryParam: 'includeGuests',
                default: true,
                groupId: 'userScope',
                groupLabelKey: 'reports.filters.groups.userScope',
            },
            {
                type: 'segmented',
                key: 'inactiveDays',
                labelKey: 'reports.filters.inactiveDays.label',
                queryParam: 'inactiveDays',
                default: null,
                options: [
                    { value: null, labelKey: 'reports.filters.inactiveDays.options.all' },
                    { value: 30, labelKey: 'reports.filters.inactiveDays.options.thirty' },
                    { value: 60, labelKey: 'reports.filters.inactiveDays.options.sixty' },
                    { value: 90, labelKey: 'reports.filters.inactiveDays.options.ninety' },
                    { value: 180, labelKey: 'reports.filters.inactiveDays.options.oneEighty' },
                    { value: 365, labelKey: 'reports.filters.inactiveDays.options.year' },
                    { value: 'never', labelKey: 'reports.filters.inactiveDays.options.never' },
                ],
            },
            {
                type: 'segmented',
                key: 'inactiveDaysMax',
                labelKey: 'reports.filters.inactiveDaysMax.label',
                queryParam: 'inactiveDaysMax',
                default: null,
                options: [
                    { value: null, labelKey: 'reports.filters.inactiveDaysMax.options.noLimit' },
                    { value: 30, labelKey: 'reports.filters.inactiveDaysMax.options.thirty' },
                    { value: 60, labelKey: 'reports.filters.inactiveDaysMax.options.sixty' },
                    { value: 90, labelKey: 'reports.filters.inactiveDaysMax.options.ninety' },
                    { value: 180, labelKey: 'reports.filters.inactiveDaysMax.options.oneEighty' },
                    { value: 365, labelKey: 'reports.filters.inactiveDaysMax.options.year' },
                    { value: 730, labelKey: 'reports.filters.inactiveDaysMax.options.twoYears' },
                    { value: 1095, labelKey: 'reports.filters.inactiveDaysMax.options.threeYears' },
                ],
            },
        ],
        buildStatusParams: (records) => ({
            count: records.length,
            licenses: records.reduce((total, item) => total + (item.licenseCount || 0), 0),
        }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            {
                key: 'lastActivityDate',
                labelKey: 'reports.table.columns.lastActivityDate',
                render: renderDateTimeCell('lastActivityDate'),
            },
            {
                key: 'daysSinceLastActivity',
                labelKey: 'reports.table.columns.daysSinceMostRecentSignIn',
            },
            {
                key: 'lastInteractiveSignIn',
                labelKey: 'reports.table.columns.lastInteractiveSignIn',
                render: renderDateTimeCell('lastInteractiveSignIn'),
            },
            {
                key: 'daysSinceInteractiveSignIn',
                labelKey: 'reports.table.columns.daysSinceInteractiveSignIn',
            },
            {
                key: 'lastNonInteractiveSignIn',
                labelKey: 'reports.table.columns.lastNonInteractiveSignIn',
                render: renderDateTimeCell('lastNonInteractiveSignIn'),
            },
            {
                key: 'daysSinceNonInteractiveSignIn',
                labelKey: 'reports.table.columns.daysSinceNonInteractiveSignIn',
            },
            {
                key: 'neverSignedIn',
                labelKey: 'reports.table.columns.neverSignedIn',
                render: renderNeverSignedInCell,
            },
            { key: 'licenseCount', labelKey: 'reports.table.columns.licenseCount' },
            {
                key: 'licenseSkus',
                labelKey: 'reports.table.columns.licenses',
                render: (record) => (record.licenseSkus || []).join(', '),
            },
        ],
    },
    'hired-this-year': {
        dataPath: '/api/reports/hired-this-year',
        exportPath: '/api/reports/hired-this-year/export',
        summaryLabelKey: 'reports.types.hiredThisYear.summaryLabel',
        tableTitleKey: 'reports.types.hiredThisYear.tableTitle',
        emptyKey: 'reports.types.hiredThisYear.empty',
        countSummaryKey: 'reports.types.hiredThisYear.countSummary',
        buildStatusParams: (records) => ({ count: records.length }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            {
                key: 'hireDate',
                labelKey: 'reports.table.columns.hireDate',
                render: (record) => formatDisplayDate(record.hireDate),
            },
            { key: 'daysSinceHire', labelKey: 'reports.table.columns.daysSinceHire' },
            { key: 'managerName', labelKey: 'reports.table.columns.manager' },
        ],
    },
    'filtered-users': {
        dataPath: '/api/reports/filtered-users',
        exportPath: '/api/reports/filtered-users/export',
        summaryLabelKey: 'reports.types.filteredUsers.summaryLabel',
        tableTitleKey: 'reports.types.filteredUsers.tableTitle',
        emptyKey: 'reports.types.filteredUsers.empty',
        countSummaryKey: 'reports.types.filteredUsers.countSummary',
        showLicenseSummary: true,
        licenseSummaryLabelKey: 'reports.summary.licensesLabel',
        filters: [
            {
                type: 'toggle',
                key: 'includeUserMailboxes',
                labelKey: 'reports.filters.includeUserMailboxes.label',
                queryParam: 'includeUserMailboxes',
                default: true,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeSharedMailboxes',
                labelKey: 'reports.filters.includeSharedMailboxes.label',
                queryParam: 'includeSharedMailboxes',
                default: true,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeRoomEquipmentMailboxes',
                labelKey: 'reports.filters.includeRoomEquipmentMailboxes.label',
                queryParam: 'includeRoomEquipmentMailboxes',
                default: true,
                groupId: 'mailboxTypes',
                groupLabelKey: 'reports.filters.groups.mailboxTypes',
            },
            {
                type: 'toggle',
                key: 'includeEnabled',
                labelKey: 'reports.filters.includeEnabled.label',
                queryParam: 'includeEnabled',
                default: true,
                groupId: 'accountStatus',
                groupLabelKey: 'reports.filters.groups.accountStatus',
            },
            {
                type: 'toggle',
                key: 'includeDisabled',
                labelKey: 'reports.filters.includeDisabled.label',
                queryParam: 'includeDisabled',
                default: true,
                groupId: 'accountStatus',
                groupLabelKey: 'reports.filters.groups.accountStatus',
            },
            {
                type: 'toggle',
                key: 'includeLicensed',
                labelKey: 'reports.filters.includeLicensed.label',
                queryParam: 'includeLicensed',
                default: true,
                groupId: 'licenseStatus',
                groupLabelKey: 'reports.filters.groups.licenseStatus',
            },
            {
                type: 'toggle',
                key: 'includeUnlicensed',
                labelKey: 'reports.filters.includeUnlicensed.label',
                queryParam: 'includeUnlicensed',
                default: true,
                groupId: 'licenseStatus',
                groupLabelKey: 'reports.filters.groups.licenseStatus',
            },
            {
                type: 'toggle',
                key: 'includeMembers',
                labelKey: 'reports.filters.includeMembers.label',
                queryParam: 'includeMembers',
                default: true,
                groupId: 'userScope',
                groupLabelKey: 'reports.filters.groups.userScope',
            },
            {
                type: 'toggle',
                key: 'includeGuests',
                labelKey: 'reports.filters.includeGuests.label',
                queryParam: 'includeGuests',
                default: true,
                groupId: 'userScope',
                groupLabelKey: 'reports.filters.groups.userScope',
            },
        ],
        buildStatusParams: (records) => ({
            count: records.length,
            licenses: records.reduce((total, item) => total + (item.licenseCount || 0), 0),
        }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            { key: 'licenseCount', labelKey: 'reports.table.columns.licenseCount' },
            {
                key: 'licenseSkus',
                labelKey: 'reports.table.columns.licenses',
                render: (record) => (record.licenseSkus || []).join(', '),
            },
            {
                key: 'filterReasons',
                labelKey: 'reports.types.filteredLicensed.columns.filterReasons',
                render: renderFilterReasonsCell,
            },
        ],
    },
    'user-scanner': {
        dataPath: '/api/user-scanner/full-scan/results',
        exportPath: null,
        summaryLabelKey: 'reports.types.userScanner.summaryLabel',
        tableTitleKey: 'reports.types.userScanner.tableTitle',
        emptyKey: 'reports.types.userScanner.empty',
        countSummaryKey: 'reports.types.userScanner.countSummary',
        isCustom: true,
        buildStatusParams: (records) => ({ count: records.length }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            { key: 'totalChecked', labelKey: 'reports.types.userScanner.columns.totalChecked' },
            { key: 'registeredCount', labelKey: 'reports.types.userScanner.columns.registeredCount' },
        ],
    },
};

const reportFiltersState = {};

function resolveFilterDefault(filter) {
    if (filter.type === 'toggle') {
        return Boolean(filter.default);
    }
    if (filter.type === 'segmented') {
        return filter.default ?? null;
    }
    return filter.default;
}

function buildDefaultFilters(config) {
    const defaults = {};
    (config.filters || []).forEach((filter) => {
        defaults[filter.key] = resolveFilterDefault(filter);
    });
    return defaults;
}

function ensureFilterState(reportKey) {
    const config = REPORT_CONFIGS[reportKey];
    const existing = reportFiltersState[reportKey];

    if (!config) {
        return existing || {};
    }

    if (!existing) {
        const defaults = buildDefaultFilters(config);
        reportFiltersState[reportKey] = defaults;
        return defaults;
    }

    (config.filters || []).forEach((filter) => {
        if (!Object.prototype.hasOwnProperty.call(existing, filter.key)) {
            existing[filter.key] = resolveFilterDefault(filter);
        }
    });

    return existing;
}

function normalizeFilterValue(filter, value) {
    if (filter.type === 'toggle') {
        if (typeof value === 'string') {
            const lowered = value.toLowerCase();
            if (lowered === 'true') {
                return true;
            }
            if (lowered === 'false') {
                return false;
            }
        }
        return Boolean(value);
    }
    if (filter.type === 'segmented') {
        if (value === null || value === undefined || value === '') {
            return null;
        }
        const parsed = Number(value);
        return Number.isNaN(parsed) ? value : parsed;
    }
    return value;
}

function applyServerFilterState(reportKey, config, serverFilters) {
    if (!config.filters || !serverFilters) {
        return;
    }
    const state = ensureFilterState(reportKey);
    config.filters.forEach((filter) => {
        const hasKey = Object.prototype.hasOwnProperty.call(serverFilters, filter.key);
        const rawValue = hasKey
            ? serverFilters[filter.key]
            : serverFilters[filter.queryParam || filter.key];
        if (rawValue === undefined) {
            return;
        }
        state[filter.key] = normalizeFilterValue(filter, rawValue);
    });
}

function updateFilterValue(reportKey, filter, value) {
    const state = ensureFilterState(reportKey);
    const normalizedValue = filter.type === 'toggle' ? Boolean(value) : value;
    const previous = state[filter.key];
    if (filter.type === 'toggle') {
        if (previous === normalizedValue) {
            return;
        }
    } else if (filter.type === 'segmented') {
        const normalizedPrev = previous === undefined ? null : previous;
        const normalizedNext = normalizedValue === undefined ? null : normalizedValue;
        if (normalizedPrev === normalizedNext) {
            return;
        }
    }

    state[filter.key] = normalizedValue;

    if (reportKey === currentReportKey) {
        const config = REPORT_CONFIGS[reportKey];
        renderFilters(config, reportKey);
        loadReport().catch((error) => {
            console.error('Failed to load report with updated filters:', error);
        });
    }
}

function renderFilters(config, reportKey) {
    const container = qs('reportFilters');
    if (!container) {
        return;
    }

    const filters = config.filters || [];
    if (!filters.length) {
        container.classList.add('is-hidden');
        container.innerHTML = '';
        return;
    }

    const t = getTranslator();
    container.classList.remove('is-hidden');
    container.innerHTML = '';

    const title = document.createElement('span');
    title.className = 'filter-toolbar__title';
    title.textContent = t('reports.filters.title');
    container.appendChild(title);

    const state = ensureFilterState(reportKey);

    const groups = [];
    filters.forEach((filter) => {
        const groupId = filter.groupId || filter.key;
        let group = groups.find((entry) => entry.id === groupId);
        if (!group) {
            group = {
                id: groupId,
                labelKey: filter.groupLabelKey || null,
                filters: [],
            };
            groups.push(group);
        } else if (!group.labelKey && filter.groupLabelKey) {
            group.labelKey = filter.groupLabelKey;
        }
        group.filters.push(filter);
    });

    groups.forEach((group) => {
        if (group.filters.length === 1 && group.filters[0].type === 'segmented') {
            const filter = group.filters[0];
            const groupElement = document.createElement('div');
            groupElement.className = 'filter-group filter-group--segmented';

            const label = document.createElement('span');
            label.className = 'filter-group__label';
            label.textContent = t(filter.labelKey);
            groupElement.appendChild(label);

            const currentValue = Object.prototype.hasOwnProperty.call(state, filter.key)
                ? state[filter.key]
                : resolveFilterDefault(filter);
            (filter.options || []).forEach((option) => {
                const optionValue = option.value ?? null;
                const isSelected = currentValue === optionValue;
                const button = document.createElement('button');
                button.type = 'button';
                button.className = `filter-chip${isSelected ? ' filter-chip--active' : ''}`;
                button.textContent = t(option.labelKey);
                button.setAttribute('aria-pressed', String(isSelected));
                button.addEventListener('click', () => {
                    updateFilterValue(reportKey, filter, optionValue);
                });
                groupElement.appendChild(button);
            });

            container.appendChild(groupElement);
            return;
        }

        const groupElement = document.createElement('div');
        groupElement.className = 'filter-group';

        if (group.labelKey) {
            const label = document.createElement('span');
            label.className = 'filter-group__label';
            label.textContent = t(group.labelKey);
            groupElement.appendChild(label);
        }

        group.filters.forEach((filter) => {
            if (filter.type !== 'toggle') {
                return;
            }

            const rawValue = Object.prototype.hasOwnProperty.call(state, filter.key)
                ? state[filter.key]
                : resolveFilterDefault(filter);
            const isActive = Boolean(rawValue);
            const button = document.createElement('button');
            button.type = 'button';
            button.className = `filter-chip${isActive ? ' filter-chip--active' : ''}`;
            button.textContent = t(filter.labelKey);
            button.setAttribute('aria-pressed', String(isActive));
            button.addEventListener('click', () => {
                updateFilterValue(reportKey, filter, !isActive);
            });
            groupElement.appendChild(button);
        });

        container.appendChild(groupElement);
    });
}

function applyFiltersToUrl(url, config, reportKey) {
    const filters = config.filters || [];
    if (!filters.length) {
        return;
    }

    const state = ensureFilterState(reportKey);

    filters.forEach((filter) => {
        const paramName = filter.queryParam || filter.key;
        const value = Object.prototype.hasOwnProperty.call(state, filter.key)
            ? state[filter.key]
            : resolveFilterDefault(filter);

        if (filter.type === 'toggle') {
            if (value) {
                url.searchParams.set(paramName, 'true');
            } else {
                url.searchParams.set(paramName, 'false');
            }
        } else if (filter.type === 'segmented') {
            if (value === null || value === undefined || value === '') {
                url.searchParams.delete(paramName);
            } else {
                url.searchParams.set(paramName, value);
            }
        } else if (value !== undefined && value !== null && value !== '') {
            url.searchParams.set(paramName, value);
        }
    });
}

function renderFilterReasonsCell(record, t) {
    const reasons = record.filterReasons || [];
    if (!reasons.length) {
        return defaultCellValue([]);
    }

    const container = document.createElement('div');
    container.className = 'reason-badges';

    reasons.forEach((reasonKey) => {
        const badge = document.createElement('span');
        badge.className = 'badge badge--neutral';
        badge.textContent = t(FILTER_REASON_I18N_KEYS[reasonKey] || reasonKey);
        container.appendChild(badge);
    });

    return container;
}

function qs(id) {
    return document.getElementById(id);
}

function getTranslator() {
    return window.i18n?.t || ((key) => key);
}

function formatDate(value) {
    if (!value) {
        return null;
    }
    try {
        const date = new Date(value);
        if (Number.isNaN(date.getTime())) {
            return value;
        }
        const pad = (num) => String(num).padStart(2, '0');
        const year = date.getFullYear();
        const month = pad(date.getMonth() + 1);
        const day = pad(date.getDate());
        const hours = pad(date.getHours());
        const minutes = pad(date.getMinutes());
        const seconds = pad(date.getSeconds());

        const offsetMinutes = -date.getTimezoneOffset();
        const offsetSign = offsetMinutes >= 0 ? '+' : '-';
        const absOffset = Math.abs(offsetMinutes);
        const offsetHours = pad(Math.floor(absOffset / 60));
        const offsetMins = pad(absOffset % 60);

        return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${offsetSign}${offsetHours}:${offsetMins}`;
    } catch (error) {
        return value;
    }
}

function formatDisplayDate(value) {
    if (!value) {
        return '—';
    }
    try {
        const date = new Date(value);
        if (Number.isNaN(date.getTime())) {
            return value;
        }
        return date.toLocaleDateString();
    } catch (error) {
        return value;
    }
}

function formatDisplayDateTime(value) {
    if (!value) {
        return '—';
    }
    try {
        const date = new Date(value);
        if (Number.isNaN(date.getTime())) {
            return value;
        }
        return date.toLocaleString();
    } catch (error) {
        return value;
    }
}

function renderDateTimeCell(field) {
    return (record) => {
        const value = record[field];
        if (!value) {
            return '—';
        }
        return formatDisplayDateTime(value);
    };
}

function renderNeverSignedInCell(record, t) {
    if (record.neverSignedIn) {
        return t('reports.table.neverSignedIn');
    }
    return '—';
}

function applyReportContext(config) {
    const t = getTranslator();
    const labelEl = qs('primarySummaryLabel');
    if (labelEl) {
        labelEl.textContent = t(config.summaryLabelKey);
    }
    const titleEl = qs('tableTitle');
    if (titleEl) {
        titleEl.textContent = t(config.tableTitleKey);
    }
}

function toggleLoading(isLoading, config, records = []) {
    const refreshBtn = qs('refreshReportBtn');
    const exportBtn = qs('exportReportBtn');
    const exportPdfBtn = qs('exportPdfBtn');
    const statusEl = qs('tableStatus');
    const t = getTranslator();

    if (refreshBtn) {
        refreshBtn.disabled = isLoading;
        refreshBtn.setAttribute('aria-busy', String(isLoading));
    }
    if (exportBtn) {
        exportBtn.disabled = isLoading || records.length === 0;
    }
    if (exportPdfBtn) {
        exportPdfBtn.disabled = isLoading || records.length === 0;
    }
    if (statusEl) {
        statusEl.textContent = isLoading
            ? t('reports.table.loading')
            : t(config.countSummaryKey, config.buildStatusParams(records));
    }
}

function showError(messageKey, detail) {
    const banner = qs('errorBanner');
    const t = getTranslator();
    if (!banner) {
        return;
    }
    const message = t(messageKey);
    banner.textContent = detail ? `${message} ${detail}` : message;
    banner.classList.remove('is-hidden');
}

function clearError() {
    const banner = qs('errorBanner');
    if (banner) {
        banner.classList.add('is-hidden');
        banner.textContent = '';
    }
}

function renderSummary(records, generatedAt, config) {
    applyReportContext(config);
    const countEl = qs('primarySummaryValue');
    const generatedEl = qs('generatedAt');
    const licenseCard = qs('licenseSummaryCard');
    const licenseLabel = qs('licenseSummaryLabel');
    const licenseValue = qs('licenseSummaryValue');
    const t = getTranslator();
    const summaryMetrics = config.buildStatusParams ? config.buildStatusParams(records) : { count: records.length };

    if (countEl) {
        countEl.textContent = records.length.toLocaleString();
    }
    if (generatedEl) {
        if (generatedAt) {
            generatedEl.textContent = formatDate(generatedAt);
        } else {
            generatedEl.textContent = t('reports.summary.generatedPending');
        }
    }

    if (licenseCard && licenseLabel && licenseValue) {
        if (config.showLicenseSummary) {
            const labelKey = config.licenseSummaryLabelKey || 'reports.summary.licensesLabel';
            licenseLabel.textContent = t(labelKey);
            const totalLicenses = summaryMetrics.licenses ?? 0;
            licenseValue.textContent = Number.isFinite(totalLicenses)
                ? totalLicenses.toLocaleString()
                : '—';
            licenseCard.classList.remove('is-hidden');
        } else {
            licenseCard.classList.add('is-hidden');
        }
    }
}

function defaultCellValue(value) {
    if (Array.isArray(value)) {
        return value.length ? value.join(', ') : '—';
    }
    if (value === null || value === undefined || value === '') {
        return '—';
    }
    return value;
}

function reasonBadgeClass(reason) {
    switch (reason) {
        case 'manager_not_found':
            return 'badge badge--danger';
        case 'detached':
            return 'badge badge--info';
        case 'filtered':
            return 'badge badge--neutral';
        default:
            return 'badge badge--warning';
    }
}

function createReasonBadge(reason, t) {
    const badge = document.createElement('span');
    badge.className = reasonBadgeClass(reason);
    const labels = {
        no_manager: 'reports.table.reasonLabels.no_manager',
        manager_not_found: 'reports.table.reasonLabels.manager_not_found',
        detached: 'reports.table.reasonLabels.detached',
        filtered: 'reports.table.reasonLabels.filtered',
    };
    const labelKey = labels[reason] || 'reports.table.reasonLabels.unknown';
    badge.textContent = t(labelKey);
    return badge;
}

function renderTable(records, config) {
    const thead = qs('reportTableHead');
    const tbody = qs('reportTableBody');
    const t = getTranslator();

    if (!thead || !tbody) {
        return;
    }

    thead.innerHTML = '';
    const headerRow = document.createElement('tr');
    config.columns.forEach((column) => {
        const th = document.createElement('th');
        th.textContent = t(column.labelKey);
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    tbody.innerHTML = '';

    if (!records.length) {
        const emptyRow = document.createElement('tr');
        emptyRow.className = 'empty-row';
        const cell = document.createElement('td');
        cell.colSpan = config.columns.length;
        cell.textContent = t(config.emptyKey);
        emptyRow.appendChild(cell);
        tbody.appendChild(emptyRow);
        return;
    }

    records.forEach((record) => {
        const row = document.createElement('tr');
        config.columns.forEach((column) => {
            const cell = document.createElement('td');
            let value;
            if (column.render) {
                value = column.render(record, t);
            } else {
                value = defaultCellValue(record[column.key]);
            }

            if (value instanceof HTMLElement) {
                cell.appendChild(value);
            } else {
                cell.textContent = value || '—';
            }

            row.appendChild(cell);
        });
        tbody.appendChild(row);
    });
}

async function loadReport({ refresh = false } = {}) {
    const config = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
    renderFilters(config, currentReportKey);
    toggleLoading(true, config);
    clearError();

    try {
        const url = new URL(config.dataPath, API_BASE_URL);
        if (refresh) {
            url.searchParams.set('refresh', 'true');
        }
        applyFiltersToUrl(url, config, currentReportKey);
        const response = await fetch(url, { credentials: 'include' });
        if (!response.ok) {
            throw new Error(`${response.status}`);
        }
        const payload = await response.json();
        latestRecords = Array.isArray(payload.records) ? payload.records : [];
        if (payload.appliedFilters) {
            applyServerFilterState(currentReportKey, config, payload.appliedFilters);
        }
        renderFilters(config, currentReportKey);
        renderSummary(latestRecords, payload.generatedAt, config);
        renderTable(latestRecords, config);
        toggleLoading(false, config, latestRecords);
    } catch (error) {
        toggleLoading(false, config, []);
        showError('reports.errors.loadFailed', error.message);
        renderSummary([], null, config);
        renderTable([], config);
    }
}

async function exportReport() {
    const config = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
    clearError();

    try {
        const url = new URL(config.exportPath, API_BASE_URL);
        applyFiltersToUrl(url, config, currentReportKey);
        const response = await fetch(url, { credentials: 'include' });
        if (!response.ok) {
            throw new Error(`${response.status}`);
        }

        const blob = await response.blob();
        let filename = `report-${new Date().toISOString().slice(0, 10)}.xlsx`;
        const disposition = response.headers.get('Content-Disposition') || response.headers.get('content-disposition');
        if (disposition) {
            const match = disposition.match(/filename="?([^";]+)"?/i);
            if (match && match[1]) {
                filename = match[1];
            }
        }

        const blobUrl = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = blobUrl;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.URL.revokeObjectURL(blobUrl);
    } catch (error) {
        showError('reports.errors.exportFailed', error.message);
    }
}

async function exportReportToPDF() {
    console.log('exportReportToPDF called');
    const config = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
    const t = getTranslator();
    clearError();

    try {
        console.log('Checking jspdf availability:', typeof window.jspdf);
        if (typeof window.jspdf === 'undefined') {
            console.error('jsPDF library not loaded');
            showError('reports.errors.pdfLibraryMissing');
            return;
        }

        console.log('latestRecords length:', latestRecords.length);
        if (latestRecords.length === 0) {
            showError('reports.errors.noDataToExport');
            return;
        }

        const { jsPDF } = window.jspdf;
        console.log('Creating PDF...');
        const pdf = new jsPDF('l', 'mm', 'a4'); // Landscape for tables

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 15;
        let yPos = margin;

        // Add report title
        const reportTitle = t(config.tableTitleKey) || 'Report';
        console.log('Adding title:', reportTitle);
        pdf.setFontSize(16);
        pdf.setTextColor(0, 0, 0);
        pdf.text(reportTitle, margin, yPos);
        yPos += 10;

        // Add record count
        pdf.setFontSize(10);
        pdf.setTextColor(100, 100, 100);
        pdf.text('Total records: ' + latestRecords.length, margin, yPos);
        yPos += 8;

        // Add active filters
        const filterState = reportFiltersState[currentReportKey] || {};
        const activeFilters = [];
        
        if (config && config.filters) {
            config.filters.forEach(filter => {
                const value = filterState[filter.key];
                if (value !== undefined && value !== null) {
                    const filterLabel = t(filter.labelKey);
                    if (filter.type === 'toggle') {
                        activeFilters.push(filterLabel + ': ' + (value ? 'Yes' : 'No'));
                    } else if (filter.type === 'segmented' && filter.options) {
                        const selectedOption = filter.options.find(opt => (opt.value ?? null) === value);
                        if (selectedOption) {
                            activeFilters.push(filterLabel + ': ' + t(selectedOption.labelKey));
                        }
                    }
                }
            });
        }

        if (activeFilters.length > 0) {
            pdf.setFontSize(8);
            pdf.setTextColor(80, 80, 80);
            pdf.text('Filters: ' + activeFilters.join(' | '), margin, yPos);
            yPos += 6;
        }

        yPos += 2; // Add spacing before table

        // Get table headers and data from the DOM
        const thead = document.getElementById('reportTableHead');
        const tbody = document.getElementById('reportTableBody');
        console.log('Table head element:', thead);
        console.log('Table body element:', tbody);
        if (!thead || !tbody) {
            showError('reports.errors.noTableToExport');
            return;
        }

        const headers = [];
        const headerCells = thead.querySelectorAll('th');
        console.log('Header cells found:', headerCells.length);
        headerCells.forEach(th => {
            headers.push(th.textContent.trim());
        });

        const rows = [];
        const bodyRows = tbody.querySelectorAll('tr:not(.empty-row)');
        console.log('Body rows found:', bodyRows.length);
        bodyRows.forEach(tr => {
            const row = [];
            tr.querySelectorAll('td').forEach(td => {
                row.push(td.textContent.trim());
            });
            if (row.length > 0) {
                rows.push(row);
            }
        });

        console.log('Headers:', headers);
        console.log('Rows count:', rows.length);

        // Calculate column widths to fit within page
        const availableWidth = pageWidth - (2 * margin);
        const colCount = headers.length || 1;
        
        // Equal width columns that fit the page
        const colWidth = availableWidth / colCount;
        const maxCharsPerCol = Math.floor(colWidth / 2.2); // ~2.2mm per character at 6pt

        const rowHeight = 5;
        const headerHeight = 6;

        // Draw header background
        pdf.setFillColor(66, 66, 66);
        pdf.rect(margin, yPos, availableWidth, headerHeight, 'F');
        
        // Draw header text
        pdf.setFontSize(6);
        pdf.setTextColor(255, 255, 255);
        for (let i = 0; i < headers.length; i++) {
            const x = margin + (i * colWidth) + 1;
            const headerText = headers[i].substring(0, maxCharsPerCol);
            pdf.text(headerText, x, yPos + 4);
        }
        yPos += headerHeight;

        // Draw rows
        pdf.setFontSize(5.5);
        const maxY = pageHeight - margin - 15;

        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            if (yPos + rowHeight > maxY) {
                pdf.addPage();
                yPos = margin;
                
                // Redraw header on new page
                pdf.setFillColor(66, 66, 66);
                pdf.rect(margin, yPos, availableWidth, headerHeight, 'F');
                pdf.setFontSize(6);
                pdf.setTextColor(255, 255, 255);
                for (let i = 0; i < headers.length; i++) {
                    const x = margin + (i * colWidth) + 1;
                    pdf.text(headers[i].substring(0, maxCharsPerCol), x, yPos + 4);
                }
                yPos += headerHeight;
                pdf.setFontSize(5.5);
            }

            // Alternate row background
            if (rowIndex % 2 === 0) {
                pdf.setFillColor(245, 245, 245);
                pdf.rect(margin, yPos, availableWidth, rowHeight, 'F');
            }

            pdf.setTextColor(30, 30, 30);
            const row = rows[rowIndex];
            for (let i = 0; i < row.length; i++) {
                const x = margin + (i * colWidth) + 1;
                const cellText = (row[i] || '').substring(0, maxCharsPerCol);
                pdf.text(cellText, x, yPos + 3.5);
            }
            yPos += rowHeight;
        }

        console.log('Adding timestamp...');
        
        // Add timestamp on every page
        const totalPages = pdf.internal.getNumberOfPages();
        const now = new Date();
        const year = now.getUTCFullYear();
        const month = String(now.getUTCMonth() + 1).padStart(2, '0');
        const day = String(now.getUTCDate()).padStart(2, '0');
        const hours = String(now.getUTCHours()).padStart(2, '0');
        const minutes = String(now.getUTCMinutes()).padStart(2, '0');
        const timestamp = 'Generated: ' + year + '-' + month + '-' + day + ' ' + hours + ':' + minutes + ' UTC';

        for (let i = 1; i <= totalPages; i++) {
            pdf.setPage(i);
            pdf.setFontSize(8);
            pdf.setTextColor(128, 128, 128);
            pdf.text(timestamp, pageWidth - margin - 60, pageHeight - 5);
            pdf.text('Page ' + i + ' of ' + totalPages, margin, pageHeight - 5);
        }

        console.log('Saving PDF...');
        const fileName = 'report-' + currentReportKey + '-' + now.toISOString().split('T')[0] + '.pdf';
        pdf.save(fileName);
        console.log('PDF exported successfully:', fileName);

    } catch (error) {
        console.error('Error exporting PDF:', error);
        console.error('Error stack:', error.stack);
        showError('reports.errors.pdfExportFailed', error.message);
    }
}

// ---------------------------------------------------------------------------
// User Scanner panel logic
// ---------------------------------------------------------------------------

function toggleUserScannerPanel(show) {
    const panel = qs('userScannerPanel');
    const tablePanel = document.querySelector('.table-panel');
    const summaryPanel = document.querySelector('.summary-panel');
    const headerActions = document.querySelector('.header-actions');
    const filterPanel = document.querySelector('.filter-panel');
    const reportFilters = qs('reportFilters');
    if (panel) panel.classList.toggle('is-hidden', !show);
    if (tablePanel) tablePanel.classList.toggle('is-hidden', show);
    if (summaryPanel) summaryPanel.classList.toggle('is-hidden', show);
    if (headerActions) headerActions.classList.toggle('is-hidden', show);
    if (filterPanel) filterPanel.classList.toggle('is-hidden', show);
    // Hide report-level filter chips (Mailbox Type, Account Status, etc.)
    if (reportFilters) { reportFilters.classList.add('is-hidden'); reportFilters.innerHTML = ''; }
}

async function checkUserScannerEnabled() {
    try {
        const resp = await fetch(`${API_BASE_URL}/api/user-scanner/status`, { credentials: 'include' });
        if (resp.ok) {
            const data = await resp.json();
            // Auto-install when enabled but not yet installed
            if (data.enabled && !data.installed) {
                try {
                    const installResp = await fetch(`${API_BASE_URL}/api/user-scanner/install`, {
                        method: 'POST',
                        credentials: 'include',
                    });
                    if (installResp.ok) {
                        const installData = await installResp.json();
                        data.installed = installData.success;
                        data.version = installData.version || null;
                    }
                } catch (installErr) {
                    console.error('Auto-install of user-scanner failed:', installErr);
                }
            }
            userScannerEnabled = data.enabled && data.installed;
            return data;
        }
    } catch (e) { /* ignore */ }
    return { enabled: false, installed: false };
}

function renderScanResultsTable(container, results, isEmail) {
    const t = getTranslator();
    container.innerHTML = '';
    if (!results || !results.length) {
        container.textContent = t('reports.types.userScanner.noResults');
        return;
    }

    const table = document.createElement('table');
    table.className = 'user-scanner-results-table';
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['Platform', 'Category', 'Status', 'Reason'].forEach(label => {
        const th = document.createElement('th');
        th.textContent = label;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    results.forEach(r => {
        const tr = document.createElement('tr');
        const statusLabel = r.status || '';
        const isRegistered = statusLabel === 'Registered' || statusLabel === 'Found';

        [r.site_name, r.category, statusLabel, r.reason || ''].forEach((val, i) => {
            const td = document.createElement('td');
            if (i === 2) {
                const badge = document.createElement('span');
                badge.className = isRegistered ? 'badge badge--warning' : 'badge badge--neutral';
                badge.textContent = val;
                td.appendChild(badge);
            } else {
                td.textContent = val || '—';
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    container.appendChild(table);
}

function _looksLikeEmail(str) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(str);
}

// Abort controller for the in-flight individual scan; allows tab-switch to cancel it.
let _singleScanAbort = null;
// Generation counter incremented on every scanner-tab switch.
// The individual scan captures this at start and bails if it changes.
let _scannerTabGen = 0;

async function runSingleUserScan() {
    // If no employee was selected from the dropdown, check if the input
    // contains a raw email and use it directly.
    if (!selectedScanUser) {
        const inputVal = (qs('userScannerInput')?.value || '').trim();
        if (_looksLikeEmail(inputVal)) {
            selectedScanUser = { name: '', email: inputVal };
        } else {
            return;
        }
    }

    // Abort any previous in-flight individual scan
    if (_singleScanAbort) { _singleScanAbort.abort(); }
    _singleScanAbort = new AbortController();
    const signal = _singleScanAbort.signal;
    const myGen = _scannerTabGen;   // capture generation

    const t = getTranslator();
    const scanBtn = qs('runUserScanBtn');
    const tablePanel = document.querySelector('.table-panel');
    if (scanBtn) { scanBtn.disabled = true; scanBtn.textContent = t('reports.types.userScanner.scanning'); }
    if (tablePanel) tablePanel.classList.remove('is-hidden');

    const thead = qs('reportTableHead');
    const tbody = qs('reportTableBody');
    const titleEl = qs('tableTitle');
    const statusEl = qs('tableStatus');
    if (titleEl) titleEl.textContent = t('reports.types.userScanner.singleResultTitle') + ' — ' + (selectedScanUser.name || selectedScanUser.email);
    if (statusEl) statusEl.textContent = t('reports.types.userScanner.scanning');
    if (thead) thead.innerHTML = '';
    if (tbody) tbody.innerHTML = '<tr><td colspan="4">' + t('reports.types.userScanner.scanning') + '</td></tr>';

    try {
        const resp = await fetch(`${API_BASE_URL}/api/user-scanner/scan`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            credentials: 'include',
            signal,
            body: JSON.stringify({
                email: selectedScanUser.email || null,
                ..._getScanOptions(),
            }),
        });
        // If user switched tabs while fetch was in-flight, discard results
        if (myGen !== _scannerTabGen) return;
        const data = await resp.json();
        if (myGen !== _scannerTabGen) return;
        if (!resp.ok) {
            if (tbody) tbody.innerHTML = '<tr class="empty-row"><td colspan="4">' + (data.error || 'Scan failed') + '</td></tr>';
            return;
        }

        // Final gate before rendering – bail if tab changed
        if (myGen !== _scannerTabGen) return;

        // Email results only – filtered server-side when sites are selected
        const rawResults = data.email_results || [];
        _collectKnownSites(rawResults);
        const nonErrorResults = rawResults.filter(r => (r.status || '').toLowerCase() !== 'error');
        const errorResults = rawResults.filter(r => (r.status || '').toLowerCase() === 'error');
        const allResults = nonErrorResults;
        latestRecords = allResults;

        // Update status text with individual scan count
        if (statusEl) {
            const found = allResults.filter(r => r.status === 'Registered' || r.status === 'Found').length;
            statusEl.textContent = allResults.length + ' sites checked · ' + found + ' registered';
        }

        if (thead) {
            thead.innerHTML = '';
            const hr = document.createElement('tr');
            ['Platform', 'Category', 'Status', 'Reason'].forEach(h => {
                const th = document.createElement('th'); th.textContent = h; hr.appendChild(th);
            });
            thead.appendChild(hr);
        }

        if (tbody) {
            tbody.innerHTML = '';
            if (!allResults.length) {
                tbody.innerHTML = '<tr class="empty-row"><td colspan="4">' + t('reports.types.userScanner.noResults') + '</td></tr>';
            } else {
                allResults.forEach(r => {
                    const tr = document.createElement('tr');
                    const statusLabel = r.status || '';
                    const isRegistered = statusLabel === 'Registered' || statusLabel === 'Found';
                    [r.site_name, r.category, statusLabel, r.reason || ''].forEach((val, i) => {
                        const td = document.createElement('td');
                        if (i === 2) {
                            const badge = document.createElement('span');
                            badge.className = isRegistered ? 'badge badge--warning' : 'badge badge--neutral';
                            badge.textContent = val;
                            td.appendChild(badge);
                        } else {
                            td.textContent = val || '—';
                        }
                        tr.appendChild(td);
                    });
                    tbody.appendChild(tr);
                });

                // Collapsible errors section
                if (errorResults.length) {
                    const errorRow = document.createElement('tr');
                    const errorTd = document.createElement('td');
                    errorTd.colSpan = 4;
                    errorTd.style.padding = '0';
                    const details = document.createElement('details');
                    details.style.padding = '10px 14px';
                    const summary = document.createElement('summary');
                    summary.style.cursor = 'pointer';
                    summary.style.color = 'var(--text-muted)';
                    summary.style.fontSize = '0.88rem';
                    summary.textContent = errorResults.length + ' site(s) returned errors (click to expand)';
                    details.appendChild(summary);
                    const errorTable = document.createElement('table');
                    errorTable.className = 'user-scanner-results-table';
                    errorTable.style.marginTop = '8px';
                    errorResults.forEach(r => {
                        const tr2 = document.createElement('tr');
                        [r.site_name, r.category, 'Error', r.reason || ''].forEach(val => {
                            const td2 = document.createElement('td');
                            td2.textContent = val || '—';
                            td2.style.color = 'var(--text-muted)';
                            tr2.appendChild(td2);
                        });
                        errorTable.appendChild(tr2);
                    });
                    details.appendChild(errorTable);
                    errorTd.appendChild(details);
                    errorRow.appendChild(errorTd);
                    tbody.appendChild(errorRow);
                }
            }
        }
    } catch (err) {
        if (err.name === 'AbortError') return;   // scan was cancelled by tab switch
        if (myGen === _scannerTabGen && tbody) tbody.innerHTML = '<tr class="empty-row"><td colspan="4">Scan failed: ' + err.message + '</td></tr>';
    } finally {
        if (scanBtn) { scanBtn.disabled = false; scanBtn.textContent = t('reports.userScanner.scanButton'); }
    }
}

// ---------------------------------------------------------------------------
// Site-filter helpers  (TagPicker-style component)
// ---------------------------------------------------------------------------

/**
 * Lightweight TagPicker clone for filtering scan results by site name.
 * Mirrors the TagPicker class from configure.js but fires an onChange
 * callback instead of calling markUnsavedChange().
 */
class SiteFilterPicker {
    constructor({ pickerId, hiddenInputId, options = [], placeholder = '', onChange }) {
        this.root = document.getElementById(pickerId);
        this.hiddenInput = document.getElementById(hiddenInputId);
        this.onChange = typeof onChange === 'function' ? onChange : () => {};
        if (!this.root || !this.hiddenInput) { this.enabled = false; return; }

        this.enabled = true;
        this.options = (Array.isArray(options) ? options.slice() : []).filter(Boolean);
        this.options.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

        this.tagContainer = this.root.querySelector('[data-role="tag-container"]');
        this.dropdown    = this.root.querySelector('[data-role="dropdown"]');
        this.input       = this.root.querySelector('.tag-picker__input');
        if (placeholder && this.input) this.input.placeholder = placeholder;

        this.selected    = [];
        this.selectedSet = new Set();
        this.filteredOptions = [];

        this._onDocClick    = this._onDocClick.bind(this);
        this._onDropClick   = this._onDropClick.bind(this);
        this._onTagClick    = this._onTagClick.bind(this);
        this._onInput       = this._onInput.bind(this);
        this._onKeyDown     = this._onKeyDown.bind(this);

        if (this.tagContainer) this.tagContainer.addEventListener('click', this._onTagClick);
        if (this.dropdown) this.dropdown.addEventListener('click', this._onDropClick);
        if (this.input) {
            this.input.addEventListener('input', this._onInput);
            this.input.addEventListener('focus', () => this._openDropdown());
            this.input.addEventListener('keydown', this._onKeyDown);
        }
        document.addEventListener('click', this._onDocClick);
        this._renderTags();
        this._closeDropdown();
        this._syncHidden();
    }

    /* --- public API --- */
    setOptions(opts) {
        if (!this.enabled) return;
        this.options = (Array.isArray(opts) ? opts.slice() : []).filter(Boolean);
        this.options.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
        this._refreshDropdown();
    }
    setValue(values) {
        if (!this.enabled) return;
        this.selected = []; this.selectedSet = new Set();
        (values || []).forEach(v => {
            const n = (v || '').trim(); if (!n) return;
            const k = n.toLowerCase();
            if (!this.selectedSet.has(k)) { this.selected.push(n); this.selectedSet.add(k); }
        });
        this._renderTags(); this._syncHidden();
        if (this.input) this.input.value = '';
        this._closeDropdown();
    }
    getValue() { return this.enabled ? this.selected.slice() : []; }
    clear()    { this.setValue([]); this.onChange(); }

    /* --- internal: add / remove --- */
    _addValue(raw) {
        if (!this.enabled) return;
        const n = (raw || '').trim(); if (!n) return;
        const k = n.toLowerCase();
        if (this.selectedSet.has(k)) { if (this.input) this.input.value = ''; this._closeDropdown(); return; }
        this.selected.push(n); this.selectedSet.add(k);
        this._renderTags(); this._syncHidden();
        if (this.input) this.input.value = '';
        this._refreshDropdown(); this._openDropdown(); this._focusSoon();
        this.onChange();
    }
    _removeValue(raw) {
        if (!this.enabled) return;
        const n = (raw || '').trim(); if (!n) return;
        const k = n.toLowerCase();
        if (!this.selectedSet.has(k)) return;
        this.selected = this.selected.filter(i => i.toLowerCase() !== k);
        this.selectedSet.delete(k);
        this._renderTags(); this._syncHidden();
        this._refreshDropdown(); this._openDropdown(); this._focusSoon();
        this.onChange();
    }

    /* --- internal: render --- */
    _renderTags() {
        if (!this.tagContainer) return;
        this.tagContainer.innerHTML = '';
        this.selected.forEach(value => {
            const tag = document.createElement('span');
            tag.className = 'tag-picker__tag';
            const label = document.createElement('span');
            label.textContent = value;
            tag.appendChild(label);
            const btn = document.createElement('button');
            btn.type = 'button'; btn.className = 'tag-picker__remove';
            btn.setAttribute('aria-label', `Remove ${value}`);
            btn.dataset.value = value; btn.innerHTML = '&times;';
            tag.appendChild(btn);
            this.tagContainer.appendChild(tag);
        });
    }
    _refreshDropdown() {
        if (!this.dropdown) return;
        const q = this.input ? this.input.value.trim().toLowerCase() : '';
        const avail = this.options.filter(o => !this.selectedSet.has(o.toLowerCase()));
        let filtered = avail;
        if (q) filtered = avail.filter(o => o.toLowerCase().includes(q));
        this.filteredOptions = filtered.slice(0, 60);
        this.dropdown.innerHTML = '';
        if (!this.filteredOptions.length) {
            const empty = document.createElement('div');
            empty.className = 'tag-picker__option tag-picker__option--empty';
            empty.textContent = q ? 'No matches found' : 'No options available';
            this.dropdown.appendChild(empty); return;
        }
        this.filteredOptions.forEach(opt => {
            const el = document.createElement('div');
            el.className = 'tag-picker__option'; el.dataset.value = opt;
            const title = document.createElement('span');
            title.className = 'tag-picker__option-title'; title.textContent = opt;
            el.appendChild(title);
            if (scannerLoudSites.has(opt)) {
                const badge = document.createElement('span');
                badge.className = 'tag-picker__loud-badge';
                badge.textContent = 'loud';
                el.appendChild(badge);
            }
            this.dropdown.appendChild(el);
        });
    }
    _openDropdown()  { if (this.dropdown) this.dropdown.hidden = false; }
    _closeDropdown() { if (this.dropdown) this.dropdown.hidden = true; }
    _syncHidden() {
        if (!this.hiddenInput) return;
        try { this.hiddenInput.value = JSON.stringify(this.selected); }
        catch { this.hiddenInput.value = this.selected.join(', '); }
    }
    _focusSoon() { if (this.input) requestAnimationFrame(() => this.input.focus()); }

    /* --- event handlers --- */
    _onDocClick(e)  { if (this.root && !this.root.contains(e.target)) this._closeDropdown(); }
    _onDropClick(e) {
        const opt = e.target.closest('[data-value]'); if (!opt) return;
        e.preventDefault(); e.stopPropagation();
        this._addValue(opt.getAttribute('data-value') || '');
    }
    _onTagClick(e)  {
        const btn = e.target.closest('.tag-picker__remove'); if (!btn) return;
        this._removeValue(btn.getAttribute('data-value') || '');
    }
    _onInput() { this._refreshDropdown(); this._openDropdown(); }
    _onKeyDown(e) {
        if (e.key === 'Backspace' && this.input && !this.input.value && this.selected.length) {
            this._removeValue(this.selected[this.selected.length - 1]); e.preventDefault();
        } else if ((e.key === 'Enter' || e.key === 'Tab') && this.input) {
            const q = this.input.value.trim(); if (!q) return;
            this._addValue(this.filteredOptions.length ? this.filteredOptions[0] : q);
            e.preventDefault();
        }
    }
}

function _renderLoudSitesList(containerId, loudList) {
    const el = document.getElementById(containerId);
    if (!el || !loudList.length) return;
    el.textContent = loudList.join(', ');
}

function _collectKnownSites(results) {
    const names = new Set(scannerKnownSites);
    (results || []).forEach(r => { if (r.site_name) names.add(r.site_name); });
    scannerKnownSites = Array.from(names).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
    // Keep picker options in sync (supplement API-sourced list with any new names)
    if (siteFilterPicker) siteFilterPicker.setOptions(scannerKnownSites);
}

/** Return the list of selected site names (empty array = all). */
function _getSelectedSiteNames() {
    return siteFilterPicker ? siteFilterPicker.getValue() : [];
}

/** Return the list of selected category names (empty array = all). */
function _getSelectedCategories() {
    return categoryFilterPicker ? categoryFilterPicker.getValue() : [];
}

/** Read toggle states and return the common scan options object (single tab). */
function _getScanOptions() {
    const sites = _getSelectedSiteNames();
    const categories = _getSelectedCategories();
    const allowLoudEl = document.getElementById('scanOptAllowLoud');
    const onlyFoundEl = document.getElementById('scanOptOnlyFound');
    return {
        sites: sites.length ? sites : null,
        categories: categories.length ? categories : null,
        allowLoud: allowLoudEl ? allowLoudEl.checked : false,
        onlyFound: onlyFoundEl ? onlyFoundEl.checked : false,
    };
}

/** Read toggle states and return scan options for the All tab (full scan). */
function _getFullScanOptions() {
    const sites = fullScanSiteFilterPicker ? fullScanSiteFilterPicker.getValue() : [];
    const categories = fullScanCategoryFilterPicker ? fullScanCategoryFilterPicker.getValue() : [];
    const allowLoudEl = document.getElementById('fullScanOptAllowLoud');
    const onlyFoundEl = document.getElementById('fullScanOptOnlyFound');
    return {
        sites: sites.length ? sites : null,
        categories: categories.length ? categories : null,
        allowLoud: allowLoudEl ? allowLoudEl.checked : false,
        onlyFound: onlyFoundEl ? onlyFoundEl.checked : false,
    };
}

/** Read the user-level filter state for the All tab. */
function _getFullScanUserFilters() {
    return { ...fullScanFilterState };
}

/**
 * Render the full-scan user filters using the same filter-group / filter-chip
 * pattern as the report filter toolbar.
 */
function renderFullScanFilters() {
    const container = document.getElementById('fullScanFilters');
    if (!container) return;

    const t = getTranslator();
    container.innerHTML = '';

    const title = document.createElement('span');
    title.className = 'filter-toolbar__title';
    title.textContent = t('reports.userScanner.fullScan.filtersTitle');
    container.appendChild(title);

    FULL_SCAN_FILTER_GROUPS.forEach(group => {
        const groupEl = document.createElement('div');
        groupEl.className = 'filter-group';

        const label = document.createElement('span');
        label.className = 'filter-group__label';
        label.textContent = t(group.labelKey);
        groupEl.appendChild(label);

        group.filters.forEach(filter => {
            const isActive = !!fullScanFilterState[filter.key];
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = `filter-chip${isActive ? ' filter-chip--active' : ''}`;
            btn.textContent = t(filter.labelKey);
            btn.setAttribute('aria-pressed', String(isActive));
            btn.addEventListener('click', () => {
                fullScanFilterState[filter.key] = !fullScanFilterState[filter.key];
                renderFullScanFilters();
            });
            groupEl.appendChild(btn);
        });

        container.appendChild(groupEl);
    });
}

function _syncSiteFilterFromPicker() {
    scannerSiteFilter = new Set(siteFilterPicker ? siteFilterPicker.getValue() : []);
}

async function initSiteFilter() {
    // --- Single-scan pickers ---
    siteFilterPicker = new SiteFilterPicker({
        pickerId: 'siteFilterPicker',
        hiddenInputId: 'siteFilterHiddenInput',
        options: scannerKnownSites,
        onChange: _syncSiteFilterFromPicker,
    });

    const resetBtn = qs('siteFilterResetBtn');
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            if (siteFilterPicker) siteFilterPicker.clear();
            scannerSiteFilter.clear();
        });
    }

    // Pre-populate picker with all available sites from the backend
    let allSites = [];
    try {
        const resp = await fetch(`${API_BASE_URL}/api/user-scanner/sites`, { credentials: 'include' });
        if (resp.ok) {
            allSites = await resp.json();
            if (Array.isArray(allSites) && allSites.length) {
                scannerKnownSites = allSites;
                siteFilterPicker.setOptions(scannerKnownSites);
            }
        }
    } catch (e) {
        // Non-critical – picker will start empty and get populated after first scan
    }

    // Fetch loud sites and render their names next to the toggle
    try {
        const resp = await fetch(`${API_BASE_URL}/api/user-scanner/loud-sites`, { credentials: 'include' });
        if (resp.ok) {
            const loudList = await resp.json();
            if (Array.isArray(loudList)) {
                scannerLoudSites = new Set(loudList);
                _renderLoudSitesList('loudSitesList', loudList);
                _renderLoudSitesList('fullScanLoudSitesList', loudList);
            }
        }
    } catch (_) { /* non-critical */ }

    // Category filter (single)
    categoryFilterPicker = new SiteFilterPicker({
        pickerId: 'categoryFilterPicker',
        hiddenInputId: 'categoryFilterHiddenInput',
        options: [],
        onChange: () => {},
    });

    const catResetBtn = qs('categoryFilterResetBtn');
    if (catResetBtn) {
        catResetBtn.addEventListener('click', () => {
            if (categoryFilterPicker) categoryFilterPicker.clear();
        });
    }

    let allCats = [];
    try {
        const resp = await fetch(`${API_BASE_URL}/api/user-scanner/categories`, { credentials: 'include' });
        if (resp.ok) {
            allCats = await resp.json();
            if (Array.isArray(allCats) && allCats.length) {
                categoryFilterPicker.setOptions(allCats);
            }
        }
    } catch (e) {
        // Non-critical
    }

    // --- Full-scan pickers ---
    fullScanSiteFilterPicker = new SiteFilterPicker({
        pickerId: 'fullScanSiteFilterPicker',
        hiddenInputId: 'fullScanSiteFilterHiddenInput',
        options: scannerKnownSites.length ? scannerKnownSites : allSites,
        onChange: () => {},
    });

    const fullSiteResetBtn = document.getElementById('fullScanSiteFilterResetBtn');
    if (fullSiteResetBtn) {
        fullSiteResetBtn.addEventListener('click', () => {
            if (fullScanSiteFilterPicker) fullScanSiteFilterPicker.clear();
        });
    }

    fullScanCategoryFilterPicker = new SiteFilterPicker({
        pickerId: 'fullScanCategoryFilterPicker',
        hiddenInputId: 'fullScanCategoryFilterHiddenInput',
        options: allCats,
        onChange: () => {},
    });

    const fullCatResetBtn = document.getElementById('fullScanCategoryFilterResetBtn');
    if (fullCatResetBtn) {
        fullCatResetBtn.addEventListener('click', () => {
            if (fullScanCategoryFilterPicker) fullScanCategoryFilterPicker.clear();
        });
    }
}

let _fullScanPollTimer = null;
let _fullScanGeneration = 0;
let _lastLogLen = 0;                 // track how many log lines we've rendered

function _stopPolling() {
    if (_fullScanPollTimer) {
        clearInterval(_fullScanPollTimer);
        _fullScanPollTimer = null;
    }
}

/* ---------- progress bar + terminal helpers ---------- */

function _showScanUI(show) {
    const wrap = qs('fullScanProgressWrap');
    const term = qs('fullScanTerminal');
    if (wrap) wrap.classList.toggle('is-hidden', !show);
    if (term) term.classList.toggle('is-hidden', !show);
    if (show) _lastLogLen = 0;
}

function _updateProgress(progress) {
    const fill = qs('fullScanProgressFill');
    const label = qs('fullScanProgressLabel');
    if (!progress) return;
    const pct = progress.total ? Math.round((progress.current / progress.total) * 100) : 0;
    if (fill) fill.style.width = pct + '%';
    if (label) label.textContent = `${progress.current} / ${progress.total}  (${pct}%)`;
}

function _appendLog(lines) {
    const body = qs('fullScanTerminalBody');
    if (!body || !lines) return;

    // Only append new lines since last render
    const newLines = lines.slice(_lastLogLen);
    if (newLines.length === 0) return;
    _lastLogLen = lines.length;

    for (const line of newLines) {
        body.textContent += (body.textContent ? '\n' : '') + line;
    }
    body.scrollTop = body.scrollHeight;
}

function _finishScanUI() {
    // Keep terminal visible so user can review, but hide progress bar
    const wrap = qs('fullScanProgressWrap');
    if (wrap) wrap.classList.add('is-hidden');
}

/* ---------- polling ---------- */

function _startPolling(generation) {
    _stopPolling();
    const t = getTranslator();
    const statusEl = qs('fullScanStatus');
    const btn = qs('runFullScanBtn');
    const stopBtn = qs('stopFullScanBtn');

    async function _poll() {
        try {
            const resp = await fetch(`${API_BASE_URL}/api/user-scanner/full-scan/status`, { credentials: 'include', cache: 'no-store' });
            if (!resp.ok) return;
            const state = await resp.json();

            // Ignore results from a different generation
            if (generation && state.generation !== generation) {
                _stopPolling();
                return;
            }

            // Always push progress & log updates while running
            if (state.progress) _updateProgress(state.progress);
            if (state.log) _appendLog(state.log);

            if (!state.running && state.cancelled && state.scanId) {
                _stopPolling();
                _finishScanUI();
                if (btn) btn.disabled = false;
                if (stopBtn) stopBtn.classList.add('is-hidden');
                if (statusEl) statusEl.textContent = t('reports.types.userScanner.fullScan.stopped');

                loadFullScanHistory();
            } else if (!state.running && state.scanId) {
                _stopPolling();
                _finishScanUI();
                if (btn) btn.disabled = false;
                if (stopBtn) stopBtn.classList.add('is-hidden');
                if (statusEl) statusEl.textContent = t('reports.types.userScanner.fullScan.complete');

                loadFullScanHistory();
            } else if (!state.running && state.error) {
                _stopPolling();
                _finishScanUI();
                if (btn) btn.disabled = false;
                if (stopBtn) stopBtn.classList.add('is-hidden');
                if (statusEl) statusEl.textContent = t('reports.types.userScanner.fullScan.failed').replace('{error}', state.error);
            } else if (state.running && state.progress) {
                // Update status text with live count
                if (statusEl) {
                    const p = state.progress;
                    statusEl.textContent = t('reports.types.userScanner.fullScan.progressStatus')
                        .replace('{current}', p.current)
                        .replace('{total}', p.total);
                }
            }
        } catch (_) { /* ignore transient errors */ }
    }

    // Fire immediately, then every 2 seconds
    _poll();
    _fullScanPollTimer = setInterval(_poll, 2000);
}

async function runFullScan() {
    const t = getTranslator();
    const emailInput = qs('fullScanEmail');
    const statusEl = qs('fullScanStatus');
    const btn = qs('runFullScanBtn');
    const stopBtn = qs('stopFullScanBtn');
    const notifyEmail = emailInput ? emailInput.value.trim() : '';

    if (btn) btn.disabled = true;
    if (statusEl) {
        statusEl.classList.remove('is-hidden');
        statusEl.textContent = t('reports.types.userScanner.fullScan.starting');
    }

    // Reset terminal + progress
    const termBody = qs('fullScanTerminalBody');
    if (termBody) termBody.textContent = '';

    try {
        const resp = await fetch(`${API_BASE_URL}/api/user-scanner/full-scan`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            credentials: 'include',
            body: JSON.stringify({
                notifyEmail: notifyEmail || null,
                ..._getFullScanOptions(),
                ..._getFullScanUserFilters(),
            }),
        });
        const data = await resp.json();
        if (!resp.ok) {
            if (statusEl) statusEl.textContent = data.error || 'Failed to start scan';
            if (btn) btn.disabled = false;
            return;
        }
        const msg = t('reports.types.userScanner.fullScan.started')
            .replace('{count}', data.employeeCount || 0);
        if (statusEl) statusEl.textContent = msg + (notifyEmail ? ' ' + t('reports.types.userScanner.fullScan.willNotify').replace('{email}', notifyEmail) : '');

        // Show stop button, progress bar, terminal
        if (stopBtn) {
            stopBtn.classList.remove('is-hidden');
            stopBtn.disabled = false;
        }
        _showScanUI(true);

        // Start polling for completion
        _fullScanGeneration = data.generation || 0;
        _startPolling(_fullScanGeneration);
    } catch (err) {
        if (statusEl) statusEl.textContent = 'Error: ' + err.message;
        if (btn) btn.disabled = false;
    }
}

async function stopFullScan() {
    const t = getTranslator();
    const statusEl = qs('fullScanStatus');
    const stopBtn = qs('stopFullScanBtn');

    if (stopBtn) stopBtn.disabled = true;
    if (statusEl) statusEl.textContent = t('reports.types.userScanner.fullScan.stopping');

    try {
        await fetch(`${API_BASE_URL}/api/user-scanner/full-scan/stop`, {
            method: 'POST',
            credentials: 'include',
        });
    } catch (_) { /* polling will handle the final state */ }
}

async function loadFullScanHistory() {
    const t = getTranslator();
    const container = qs('fullScanHistory');
    if (!container) return;

    try {
        const resp = await fetch(`${API_BASE_URL}/api/user-scanner/full-scan/history`, { credentials: 'include' });
        if (!resp.ok) return;
        const history = await resp.json();

        if (!history || history.length === 0) {
            container.innerHTML = '';
            return;
        }

        const heading = t('reports.types.userScanner.fullScan.historyTitle');
        const clearLabel = t('reports.types.userScanner.fullScan.clearHistory');
        let html = `<div class="full-scan-history-header"><h4>${heading}</h4><button type="button" class="btn btn-secondary btn-sm full-scan-history-clear" id="clearScanHistoryBtn">${clearLabel}</button></div><ul class="full-scan-history-list">`;
        for (const entry of history) {
            const date = entry.scannedAt ? new Date(entry.scannedAt).toLocaleString() : '—';
            const dlUrl = `${API_BASE_URL}/api/user-scanner/full-scan/download/${encodeURIComponent(entry.scanId)}`;
            html += `<li class="full-scan-history-item">
                <div class="full-scan-history-info">
                    <span class="full-scan-history-date">${date}</span>
                    <span class="full-scan-history-meta">${entry.totalEmployees} ${t('reports.types.userScanner.fullScan.historyEmployees')} · ${entry.registeredTotal} ${t('reports.types.userScanner.fullScan.historyRegistered')}</span>
                </div>
                <a href="${dlUrl}" class="btn btn-secondary btn-sm full-scan-history-dl" download>${t('reports.types.userScanner.fullScan.download')}</a>
            </li>`;
        }
        html += '</ul>';
        container.innerHTML = html;

        // Wire up clear button
        const clearBtn = document.getElementById('clearScanHistoryBtn');
        if (clearBtn) {
            clearBtn.addEventListener('click', async () => {
                try {
                    const r = await fetch(`${API_BASE_URL}/api/user-scanner/full-scan/history`, { method: 'DELETE', credentials: 'include' });
                    if (r.ok) loadFullScanHistory();
                } catch (_) { /* ignore */ }
            });
        }
    } catch (_) { /* ignore */ }
}

function initUserScannerPanel() {
    const input = qs('userScannerInput');
    const searchResults = qs('userScannerSearchResults');
    const scanBtn = qs('runUserScanBtn');
    const fullScanBtn = qs('runFullScanBtn');

    // --- Tab switching ---
    const tabs = document.querySelectorAll('.scanner-tab');
    const tabContents = document.querySelectorAll('.scanner-tab-content');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const target = tab.dataset.tab;
            tabs.forEach(t => { t.classList.toggle('is-active', t.dataset.tab === target); t.setAttribute('aria-selected', t.dataset.tab === target ? 'true' : 'false'); });
            tabContents.forEach(c => { c.classList.toggle('is-active', c.dataset.tabContent === target); });

            // Bump generation so any in-flight individual scan discards results
            _scannerTabGen++;
            // Abort any in-flight individual scan when leaving the tab
            if (_singleScanAbort) { _singleScanAbort.abort(); _singleScanAbort = null; }

            const tablePanel = document.querySelector('.table-panel');
            const thead = qs('reportTableHead');
            const tbody = qs('reportTableBody');
            const titleEl = qs('tableTitle');
            const statusEl = qs('tableStatus');
            // Always clear the shared table on tab switch
            if (thead) thead.innerHTML = '';
            if (tbody) tbody.innerHTML = '';
            if (titleEl) titleEl.textContent = '';
            if (statusEl) statusEl.textContent = '';

            if (target === 'all') {
                // Organization tab: hide results table, only show terminal + downloads
                if (tablePanel) tablePanel.classList.add('is-hidden');
                loadFullScanHistory();
            } else {
                // Individual tab: show the results table for single-user scans
                if (tablePanel) tablePanel.classList.remove('is-hidden');
            }
        });
    });

    if (scanBtn) {
        scanBtn.disabled = true;
        scanBtn.addEventListener('click', () => runSingleUserScan());
    }

    if (fullScanBtn) {
        fullScanBtn.addEventListener('click', () => runFullScan());
    }

    const stopScanBtn = qs('stopFullScanBtn');
    if (stopScanBtn) {
        stopScanBtn.addEventListener('click', () => stopFullScan());
    }

    // Load history and check if a scan is currently running
    loadFullScanHistory();
    (async () => {
        try {
            const resp = await fetch(`${API_BASE_URL}/api/user-scanner/full-scan/status`, { credentials: 'include', cache: 'no-store' });
            if (resp.ok) {
                const state = await resp.json();
                if (state.running) {
                    const statusEl = qs('fullScanStatus');
                    if (statusEl) {
                        statusEl.classList.remove('is-hidden');
                        statusEl.textContent = getTranslator()('reports.types.userScanner.fullScan.running');
                    }
                    if (fullScanBtn) fullScanBtn.disabled = true;
                    if (stopScanBtn) stopScanBtn.classList.remove('is-hidden');
                    _showScanUI(true);
                    if (state.progress) _updateProgress(state.progress);
                    if (state.log) _appendLog(state.log);
                    _fullScanGeneration = state.generation || 0;
                    _startPolling(_fullScanGeneration);
                }
            }
        } catch (_) { /* ignore */ }
    })();

    if (!input || !searchResults) return;

    let debounceTimer = null;
    input.addEventListener('input', () => {
        clearTimeout(debounceTimer);
        const query = input.value.trim();
        if (query.length < 2) {
            searchResults.classList.add('is-hidden');
            searchResults.innerHTML = '';
            selectedScanUser = null;
            if (scanBtn) scanBtn.disabled = true;
            return;
        }

        // Enable the scan button when the input looks like a valid email,
        // even without selecting from the dropdown.
        if (_looksLikeEmail(query)) {
            selectedScanUser = null;  // will be resolved in runSingleUserScan
            if (scanBtn) scanBtn.disabled = false;
        } else {
            if (!selectedScanUser && scanBtn) scanBtn.disabled = true;
        }

        debounceTimer = setTimeout(async () => {
            try {
                // Use scanner-specific endpoint that includes all non-guest users
                const resp = await fetch(`${API_BASE_URL}/api/user-scanner/users?q=${encodeURIComponent(query)}`, { credentials: 'include' });
                const employees = await resp.json();
                searchResults.innerHTML = '';
                if (!employees.length) {
                    searchResults.classList.add('is-hidden');
                    return;
                }
                searchResults.classList.remove('is-hidden');
                employees.forEach(emp => {
                    const item = document.createElement('div');
                    item.className = 'user-scanner-search__item';

                    const nameEl = document.createElement('strong');
                    nameEl.textContent = emp.name || '—';

                    const emailEl = document.createElement('span');
                    emailEl.textContent = emp.email || emp.mail || '';

                    const titleEl = document.createElement('span');
                    titleEl.className = 'badge badge--neutral';
                    titleEl.textContent = emp.title || '';

                    item.appendChild(nameEl);
                    item.appendChild(document.createTextNode(' '));
                    item.appendChild(emailEl);
                    item.appendChild(document.createTextNode(' '));
                    item.appendChild(titleEl);
                    item.addEventListener('click', () => {
                        selectedScanUser = {
                            name: emp.name || '',
                            email: emp.email || emp.mail || '',
                        };
                        input.value = emp.name + (emp.email ? ' (' + emp.email + ')' : '');
                        searchResults.classList.add('is-hidden');
                        if (scanBtn) scanBtn.disabled = false;
                    });
                    searchResults.appendChild(item);
                });
            } catch (e) {
                searchResults.classList.add('is-hidden');
            }
        }, 300);
    });

    initSiteFilter();
    renderFullScanFilters();
}

async function initializeReportsPage() {
    const htmlElement = document.documentElement;
    const i18nReadyPromise = window.i18n?.ready;

    try {
        if (i18nReadyPromise && typeof i18nReadyPromise.then === 'function') {
            try {
                await i18nReadyPromise;
            } catch (error) {
                console.error('Failed to load translations for reports page:', error);
            }
        }

        const reportSelect = qs('reportTypeSelect');
        if (reportSelect) {
            // Check if user-scanner is enabled; hide option if not
            const scannerStatus = await checkUserScannerEnabled();
            const scannerOption = reportSelect.querySelector('option[value="user-scanner"]');
            if (scannerOption && !scannerStatus.enabled) {
                scannerOption.style.display = 'none';
            }

            reportSelect.value = currentReportKey;
            // Show scanner panel if it's the initial selection
            if (currentReportKey === 'user-scanner' && scannerStatus.enabled) {
                toggleUserScannerPanel(true);
            }
            reportSelect.addEventListener('change', () => {
                currentReportKey = reportSelect.value;
                const isScanner = currentReportKey === 'user-scanner';
                toggleUserScannerPanel(isScanner);
                if (isScanner) return;
                const config = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
                ensureFilterState(currentReportKey);
                renderSummary([], null, config);
                renderTable([], config);
                renderFilters(config, currentReportKey);
                loadReport().catch((error) => {
                    console.error('Failed to load report:', error);
                });
            });
        }

        initUserScannerPanel();

        const refreshBtn = qs('refreshReportBtn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => loadReport({ refresh: true }));
            const t = getTranslator();
            refreshBtn.title = t('reports.buttons.refreshTooltip');
        }

        const exportBtn = qs('exportReportBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', exportReport);
            const t = getTranslator();
            exportBtn.title = t('reports.buttons.exportTooltip');
            exportBtn.disabled = true;
        }

        const exportPdfBtn = qs('exportPdfBtn');
        console.log('exportPdfBtn element:', exportPdfBtn);
        if (exportPdfBtn) {
            exportPdfBtn.addEventListener('click', () => {
                console.log('Export PDF button clicked');
                exportReportToPDF();
            });
            const t = getTranslator();
            exportPdfBtn.title = t('reports.buttons.exportPdfTooltip');
            exportPdfBtn.disabled = true;
        }

        const initialConfig = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
        ensureFilterState(currentReportKey);
        renderSummary([], null, initialConfig);
        renderTable([], initialConfig);
        renderFilters(initialConfig, currentReportKey);

        await loadReport();
    } finally {
        htmlElement.classList.remove('i18n-loading');
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initializeReportsPage().catch((error) => {
        console.error('Failed to initialize reports page:', error);
        showError('reports.errors.initializationFailed');
    });
});
