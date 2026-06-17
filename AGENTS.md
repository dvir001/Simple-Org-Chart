# AGENTS.md

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

## 5. Project Architecture

**Flask + vanilla JS + D3. No ORM, no front-end framework.**

```
simple_org_chart/   ← Python package (Flask app)
  app_main.py       ← All Flask routes, filter parsers, API endpoints
  msgraph.py        ← Graph API helpers: fetch_all_employees, probe_graph_capabilities, etc.
  reports.py        ← Pure filter functions (apply_tagpicker_filters, apply_last_login_filters, …)
  data_update.py    ← Sync orchestration: token → probe caps → fetch employees → write caches
  config.py         ← Path constants (DATA_DIR, *_FILE paths)
  settings.py       ← load_settings / save_settings helpers
  auth.py           ← @require_auth decorator
static/
  app.js            ← D3 org chart, node click/drag, employee detail panel
  reports.js        ← Report filter UI, tagpicker/toggle renderers, REPORT_CONFIGS
  reports.css       ← Report filter layout styles
  locales/en-US.json← All user-facing strings (i18n keys)
data/               ← JSON caches written at sync time (git-ignored)
```

## 6. Adding a Report Filter

Every filter touches all of these layers — miss one and it silently does nothing:

1. **`msgraph.py`** — add the field to `$select` in `fetch_all_employees` / `collect_last_login_records`; populate it on each record dict.
2. **`reports.py`** — add the parameter to every `apply_*_filters` function signature; implement the filtering logic.
3. **`app_main.py`** — parse the query-string param in `_parse_standard_toggle_args` or `_parse_tagpicker_args`; forward it to the filter function in every relevant route (GET + export).
4. **`static/reports.js`** — add the filter object to `_standardToggleFilters()` or `TAGPICKER_FILTERS`; add `requiredCapability` if the filter needs a Graph permission beyond `User.Read.All`.
5. **`static/locales/en-US.json`** — add label and (for tagpickers) placeholder/mode keys.

## 7. Graph Permission & Capability Gating

Filters that depend on Graph permissions beyond the base `User.Read.All` must be gated:

| Filter group | Required permission | Capability key |
|---|---|---|
| Mailbox type (User / Shared / Room) | `MailboxSettings.Read` | `mailbox_settings_read` |
| GAL visibility (Hidden / Visible) | `MailboxSettings.Read` | `mailbox_settings_read` |
| Inactivity day-range | `AuditLog.Read.All` + Entra P1/P2 | `audit_log_read_all` |

**How it works:**
- `msgraph.probe_graph_capabilities(token)` decodes the JWT access token's `roles` claim (no extra API calls) at the start of every sync and writes `data/graph_capabilities.json`.
- `GET /api/graph-capabilities` serves that file to the front end.
- `reports.js` fetches capabilities before first render; `_isFilterCapable(filter)` checks `filter.requiredCapability` against the loaded flags.
- Incapable filter buttons get `aria-disabled` + class `filter-chip--unavailable` (greyed out, tooltip explains missing permission).

When adding a new filter that needs a Graph permission:
1. Add a probe call in `probe_graph_capabilities()` if the capability isn't already detected.
2. Set `requiredCapability: '<key>'` on the filter object in `reports.js`.

## 8. Data Flow: Sync → Cache → API → UI

```
data_update.run_data_update()
  ├─ probe_graph_capabilities()  → data/graph_capabilities.json
  ├─ fetch_all_employees()       → data/employee_data.json, missing_manager, filtered_users, …
  ├─ collect_last_login_records()
  │    └─ enrich with managerId from employee list
  │                               → data/last_login_records.json
  └─ collect_disabled_users()    → data/disabled_user_records.json

GET /api/reports/<type>  →  load cache  →  apply_*_filters()  →  JSON response
GET /api/graph-capabilities  →  data/graph_capabilities.json  →  JSON response
```

## 9. i18n Rules

- Every user-visible string must have a key in `static/locales/en-US.json`.
- The translator `t(key)` is available in `reports.js` via `getTranslator()`.
- Never hardcode English strings in JS templates.
- When adding filters: add `labelKey`, `placeholderKey`, `resetLabelKey`, `modeIncludeLabelKey`, `modeExcludeLabelKey` references and the matching JSON entries.
