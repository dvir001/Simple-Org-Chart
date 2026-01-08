# SimpleOrgChart

> **Note:** This repository is a maintained fork of [jaffster595/DB-Auto-Org-Chart](https://github.com/jaffster595/DB-Auto-Org-Chart).

SimpleOrgChart is a Flask application backed by Azure Active Directory (Entra ID) data that renders a fully interactive, client-side organisation chart. A static JavaScript front end (vanilla JS + D3) consumes cached Graph API data, offers rich filtering, and exposes an admin dashboard with compliance-friendly exports.

<img width="1640" height="527" alt="SimpleOrgChart preview (names hidden toggle)" src="https://github.com/user-attachments/assets/cde7e312-f7e6-466c-9750-df84026b3990" />

## At a Glance

- Hardened security defaults: strict Content Security Policy, sanitized redirects, login isolation, and placeholder-secret protection.
- Modular front end: no inline scripts or styles; shared CSS variables power `configure`, `reports`, and org chart experiences.
- Daily automation: background scheduler refreshes Azure AD data (20:00 local time) and persists JSON caches under `data/`.
- Admin reporting: missing managers, filtered users, and last-login inactivity insights—each with one-click XLSX export.
- Export tooling: SVG/PNG/PDF org chart capture and server-backed XLSX generation for the current chart tree.
- Deployment ready: ships with Docker Compose and a Gunicorn configuration (`deploy/gunicorn.conf.py`) for containerized hosting.

## How It Works

| Layer | Description |
| --- | --- |
| Flask backend (`simple_org_chart/` package) | Serves templates/static assets, authenticates admin endpoints, manages schedulers, and stores cached Graph API responses.
| Static front end (`static/*.js`) | Renders the D3 org chart, configuration UI, and reports dashboard using cached JSON.
| Data cache (`data/*.json`) | Holds employee hierarchies, report snapshots, and last login activity to reduce Graph API calls.
| Scheduler | Nightly job (20:00) refreshes employee data; manual refresh endpoints and CLI helpers are available.

## Prerequisites

1. Docker Desktop (or Docker Engine with the Compose plugin) for container-based deployment.
2. An Azure AD tenant with privileges to create app registrations and grant Graph application permissions.

## Azure AD Setup

1. **Create an App Registration**
   - Azure Portal ➜ Azure Active Directory ➜ App registrations ➜ **New registration**.
   - Choose a name (for example, `SimpleOrgChart`) and leave Redirect URI empty.

2. **Assign Microsoft Graph Application Permissions**
   - `User.Read.All`
   - `LicenseAssignment.Read.All` *(required for licensing insights and admin reports)*
   - `AuditLog.Read.All` *(required for last sign-in metrics and disabled-user audit timestamps)*
   - `MailboxSettings.Read` *(enables mailbox-type metadata used by last sign-in filters; without it, all mailboxes are treated as standard users)*
   - Grant admin consent for the tenant.

3. **Create a Client Secret**
   - Certificates & secrets ➜ **New client secret**.

4. **Capture Identifiers**
   - Application (client) ID → `AZURE_CLIENT_ID`
   - Directory (tenant) ID → `AZURE_TENANT_ID`

## Configure Environment Variables

Copy the template and fill in your secrets.

```bash
cp .env.template .env
# edit .env with tenant/client IDs, secret, admin password, etc.
```

**Required values**

- `AZURE_TENANT_ID` – Directory (tenant) ID.
- `AZURE_CLIENT_ID` – Application (client) ID.
- `AZURE_CLIENT_SECRET` – Client secret value.
- `TOP_LEVEL_USER_EMAIL` – Email for the org chart root user.
- `ADMIN_PASSWORD` – Protects `/configure` and `/reports`.
- `SECRET_KEY` – 64+ character random string for Flask sessions.

Generate a strong secret key:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

**Optional values**

- `TOP_LEVEL_USER_ID` – Explicit Graph object ID for the root user.
- `CORS_ALLOWED_ORIGINS` – Comma-separated list of allowed cross-origin hosts.
- `RUN_INITIAL_UPDATE` – Set to `false` to skip automatic data refresh at startup.
- `APP_PORT` – Port the application listens on (defaults to `5000`).

## Running the Application

### Docker (recommended)

```bash
docker compose pull
docker compose up -d
```

- Default port: `APP_PORT` (defaults to `5000`). Override it in `.env` to change container and host bindings.
- Persistent data resides in the `orgchart_data` volume. Remove it to rebuild caches from scratch.
- Local execution outside Docker is not supported; use the provided container workflow for development and production.

## Key Features

- **Interactive D3 Org Chart**: Pan, zoom, and expand/collapse hierarchies with persistent hidden subtrees.
- **Search & Discovery**: Real-time directory search, quick navigation helpers, and configurable filters for guests/disabled users.
- **Configuration UI** (`/configure`): Adjust styling, filtering, export columns, and scheduling without editing files.
- **Admin Reports** (`/reports`):
  - Missing managers
   - Users by last sign-in activity
   - Employees hired in the last 365 days
   - Users hidden by filters
- **Export Options**: SVG/PNG/PDF snapshots and XLSX exports for reports and chart data.
- **MicroSIP Directory Feed**: Serve a MicroSIP contacts JSON at /contacts.json using cached employee data.
- **Caching & Scheduling**: JSON caches regenerate nightly; manual refresh endpoints keep data current on demand.

## MicroSIP Directory

SimpleOrgChart publishes a MicroSIP-compatible directory export at /contacts.json.

- The feed reuses the cached employee list; trigger a manual refresh if the response is empty.
- Response headers disable caching so MicroSIP always retrieves the latest contacts.
- Fields include number, name, firstname, lastname, phone, mobile, email, address, city, state, comment, presence, starred, and info.
- Contacts without a desk or mobile number are skipped to keep the directory free of unreachable entries.
- Add extra entries from the Configure → Custom MicroSIP Contacts textarea (one Name,Number pair per line) when you need off-chart contacts in the feed.

Example payload:

```json
{
   "refresh": 1736073600,
   "items": [
      {
         "number": "5551234567",
         "name": "Ada Lovelace",
         "firstname": "Ada",
         "lastname": "Lovelace",
         "phone": "555-123-4567",
         "mobile": "555-987-6543",
         "email": "ada@example.com",
         "city": "London",
         "state": "",
         "comment": "Engineering - Research",
         "presence": 0,
         "starred": 0
      }
   ]
}
```

## Reporting Caches

- `data/employee_data.json` – Full org hierarchy.
- `data/missing_manager_records.json` – Missing manager snapshot.
- `data/disabled_user_records.json` – Disabled users enriched with license and sign-in metadata.
- `data/last_login_records.json` – Active users with last sign-in timestamps.
- Additional files exist for filtered/disabled-with-license/hiring reports.

If a cache is missing or stale, hit **Refresh Data** on the reports page or start the app with `RUN_INITIAL_UPDATE=true`.

## Security Guidance

- Store secrets in Azure Key Vault or your host’s secret manager; never commit `.env` files.
- Restrict `/configure` and `/reports` behind reverse-proxy auth if deployed on the public internet.
- Monitor `AuditLog.Read.All` usage—limit consent scope to required admins.
- Rotate `AZURE_CLIENT_SECRET` regularly and update the environment accordingly.

## Troubleshooting

- **Graph permission errors**: Ensure admin consent is granted; check logs for 403 responses when fetching `signInActivity`.
- **Stale data**: Run `curl -X POST http://<host>/api/update-now` (with admin auth) or remove the `data/*.json` caches and restart.
- **Export failures**: Confirm `openpyxl` is installed (bundled via `requirements.txt`). The API returns a 500 with JSON error details if export dependencies are missing.
- **Missing logos**: Upload custom branding via `/configure`; static assets persist in `data/`.

## Contributing

Issues and pull requests are welcome. Please document new locale strings in `static/locales/en-US.json`, update report caches when introducing routes, and include manual validation steps if automated tests are not available.

---

SimpleOrgChart keeps your organisation chart and admin insights in sync with Azure AD while staying lightweight, portable, and secure.
