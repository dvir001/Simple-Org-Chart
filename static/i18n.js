(function () {
    const DEFAULT_LOCALE = 'en-US';
    const htmlLang = (document.documentElement.getAttribute('lang') || '').trim();
    const normalizedLowerLang = htmlLang ? htmlLang.toLowerCase() : DEFAULT_LOCALE.toLowerCase();

    const localeCandidates = [];
    const addCandidate = (candidate) => {
        if (candidate && !localeCandidates.includes(candidate)) {
            localeCandidates.push(candidate);
        }
    };

    // Build locale candidates from most-specific to least-specific.
    // Only add candidates that follow the xx-YY pattern used by our locale
    // files to avoid 404s for bare language codes like "en" or invented
    // region codes like "en-EN".
    if (htmlLang && htmlLang.includes('-')) {
        // e.g. "en-US" or "en-us" → normalise to "en-US"
        const [langPart, regionPart] = htmlLang.split('-', 2);
        addCandidate(`${langPart.toLowerCase()}-${(regionPart || '').toUpperCase()}`);
    }

    addCandidate(DEFAULT_LOCALE);

    const defaultLocalePath = `/static/locales/${DEFAULT_LOCALE}.json`;

    let translations = {};
    let fallbackTranslations = {};

    async function fetchJson(path) {
        const response = await fetch(path, { cache: 'no-cache' });
        if (!response.ok) {
            throw new Error(`Failed to load locale file: ${path}`);
        }
        return response.json();
    }

    function getNestedValue(obj, key) {
        return key.split('.').reduce((acc, part) => (acc && acc[part] !== undefined ? acc[part] : undefined), obj);
    }

    function formatString(template, params) {
        if (!params) {
            return template;
        }
        return template.replace(/\{([^}]+)\}/g, (_, match) => {
            const value = params[match.trim()];
            return value === undefined ? `{${match}}` : value;
        });
    }

    function camelToKebab(value) {
        return value
            .replace(/([a-z0-9])([A-Z])/g, '$1-$2')
            .replace(/([A-Z])([A-Z][a-z])/g, '$1-$2')
            .toLowerCase();
    }

    function parseParams(value) {
        if (!value) {
            return undefined;
        }
        try {
            return JSON.parse(value);
        } catch (error) {
            console.warn('[i18n] Failed to parse translation params', value, error);
            return undefined;
        }
    }

    function translate(key, params) {
        const direct = getNestedValue(translations, key);
        const fallback = getNestedValue(fallbackTranslations, key);
        const template = (direct !== undefined ? direct : fallback);
        if (template === undefined) {
            return key;
        }
        return formatString(String(template), params);
    }

    function applyTranslations(root = document) {
        const elements = root.querySelectorAll('[data-i18n], [data-i18n-html], [data-i18n-placeholder], [data-i18n-title], [data-i18n-ariaLabel]');
        elements.forEach((el) => {
            const params = parseParams(el.dataset.i18nParams);
            if (el.dataset.i18n) {
                el.textContent = translate(el.dataset.i18n, params);
            }
            if (el.dataset.i18nHtml) {
                el.innerHTML = translate(el.dataset.i18nHtml, params);
            }

            Object.keys(el.dataset)
                .filter((key) => key.startsWith('i18n') && key !== 'i18n' && key !== 'i18nHtml')
                .forEach((dataKey) => {
                    const attrNamePart = dataKey.substring('i18n'.length);
                    if (!attrNamePart) {
                        return;
                    }
                    const attrName = camelToKebab(attrNamePart);
                    el.setAttribute(attrName, translate(el.dataset[dataKey], params));
                });
        });
    }

    const readyPromise = (async () => {
        try {
            fallbackTranslations = await fetchJson(defaultLocalePath);
        } catch (error) {
            console.error('[i18n] Failed to load default locale', error);
            fallbackTranslations = {};
        }

        for (const candidate of localeCandidates) {
            const path = `/static/locales/${candidate}.json`;
            try {
                translations = await fetchJson(path);
                return;
            } catch (error) {
                if (candidate === DEFAULT_LOCALE) {
                    console.warn('[i18n] Default locale missing or failed to load, continuing with fallback object.');
                }
            }
        }

        translations = fallbackTranslations;
    })();

    window.i18n = {
        t: translate,
        ready: readyPromise,
        applyTranslations,
    };

    readyPromise.then(() => {
        applyTranslations(document);
    }).catch((error) => {
        console.error('[i18n] Initialization error', error);
    });
})();
