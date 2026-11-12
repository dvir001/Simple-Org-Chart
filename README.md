# Simple Org Chart

> **Note:** This repository is a maintained fork of [jaffster595/DB-Auto-Org-Chart](https://github.com/jaffster595/DB-Auto-Org-Chart).

SimpleOrgChart is a Flask application backed by Azure Active Directory (Entra ID) data that renders a fully interactive, client-side organisation chart. A static JavaScript front end (vanilla JS + D3) consumes cached Graph API data, offers rich filtering, and exposes an admin dashboard with compliance-friendly exports.

<img width="1640" height="527" alt="SimpleOrgChart preview (names hidden toggle)" src="https://github.com/user-attachments/assets/0f8eebbe-b071-42f8-adc4-674051022a9d" />

## At a Glance

- Hardened security defaults: strict Content Security Policy, sanitized redirects, login isolation, and placeholder-secret protection.
- Modular front end: no inline scripts or styles; shared CSS variables power `configure`, `reports`, and org chart experiences.
- Daily automation: background scheduler refreshes Azure AD data once a day (20:00 UTC by default, configurable under `/configure`) and persists JSON caches under `data/`.
- Admin reporting: missing managers, filtered users, last-login inactivity, missing profile pictures, data quality issues, and recently hired users—each with one-click XLSX export.
- Export tooling: SVG/PNG/PDF org chart capture and server-backed XLSX generation for the current chart tree.
- Deployment ready: ships with Docker Compose and a Gunicorn configuration (`deploy/gunicorn.conf.py`) for containerized hosting.

## How It Works

| Layer | Description |
| --- | --- |
| Flask backend (`simple_org_chart/` package) | Serves templates/static assets, authenticates admin endpoints, manages schedulers, and stores cached Graph API responses.
| Static front end (`static/*.js`) | Renders the D3 org chart, configuration UI, and reports dashboard using cached JSON.
| Data cache (`data/*.json`) | Holds employee hierarchies, report snapshots, and last login activity to reduce Graph API calls.
| Scheduler | Daily job (20:00 UTC by default) refreshes employee data; the **Sync** button triggers a refresh on demand.

## Quick Start

You need Docker and an Entra ID (Azure AD) app registration. The app runs from the published image `ghcr.io/dvir001/db-auto-org-chart`; there is nothing to clone or build.

1. **Register the app in Entra ID** and note the tenant ID, client ID, and client secret — see [Azure AD Setup](#azure-ad-setup).
2. **Download the two files you need** into an empty folder:
   ```bash
   curl -O https://raw.githubusercontent.com/dvir001/Simple-Org-Chart/main/docker-compose.yml
   curl -o .env https://raw.githubusercontent.com/dvir001/Simple-Org-Chart/main/.env.template
   ```
3. **Fill in the five required values** in `.env` (details in [Configure Environment Variables](#configure-environment-variables)):
   ```dotenv
   AZURE_TENANT_ID=<directory (tenant) ID>
   AZURE_CLIENT_ID=<application (client) ID>
   AZURE_CLIENT_SECRET=<client secret value>
   ADMIN_PASSWORD=<a strong password of your choice>
   SECRET_KEY=<output of the command below>
   ```
   ```bash
   python -c "import secrets; print(secrets.token_hex(32))"
   ```
   Also remove or adjust the sample `CORS_ALLOWED_ORIGINS` line if you do not need cross-origin access.
4. **Start the container** (this pulls the published image):
   ```bash
   docker compose up -d
   ```
5. **Open** `http://localhost:5000`. The first start downloads your directory from Microsoft Graph, which can take a few minutes for large tenants; follow progress with `docker compose logs -f`.
6. **Sign in as administrator** at `http://localhost:5000/login` with `ADMIN_PASSWORD` to reach **Configure** and **Reports**.

If something does not work, see [Troubleshooting](#troubleshooting).

## Prerequisites

1. Docker Desktop (or Docker Engine with the Compose plugin) for container-based deployment.
2. An Azure AD tenant with privileges to create app registrations and grant Graph application permissions.

## Azure AD Setup

1. **Create an App Registration**
   - Azure Portal ➜ Azure Active Directory ➜ App registrations ➜ **New registration**.
   - Choose a name (for example, `SimpleOrgChart`) and leave Redirect URI empty.

2. **Assign Microsoft Graph Application Permissions**
   - `User.Read.All`
   - `Presence.Read.All` *(enables live Teams presence status indicators on org chart cards)*
   - `LicenseAssignment.Read.All` *(required for licensing insights and admin reports)*
   - `AuditLog.Read.All` *(required for last sign-in metrics and disabled-user audit timestamps; also enables the inactivity day-range filters on the Last Logins report — without it those filters are greyed out)*
   - `MailboxSettings.Read` *(enables mailbox-type metadata used by mailbox-type filters; without it the User/Shared/Room mailbox filters are greyed out)*
   - Grant admin consent for the tenant.

3. **Create a Client Secret**
   - Certificates & secrets ➜ **New client secret**.
   - Copy the secret **Value** (not the Secret ID) immediately; it is shown only once → `AZURE_CLIENT_SECRET`

4. **Capture Identifiers**
   - Application (client) ID → `AZURE_CLIENT_ID`
   - Directory (tenant) ID → `AZURE_TENANT_ID`

Only `User.Read.All` is strictly required. The other permissions each unlock the feature noted beside them and can be added later; run a sync after granting a new permission.

## Configure Environment Variables

Create `.env` from the template (the Quick Start download already does this) and fill in your secrets.

```bash
curl -o .env https://raw.githubusercontent.com/dvir001/Simple-Org-Chart/main/.env.template
# edit .env with tenant/client IDs, secret, admin password, etc.
```

**Required values**

- `AZURE_TENANT_ID` – Directory (tenant) ID.
- `AZURE_CLIENT_ID` – Application (client) ID.
- `AZURE_CLIENT_SECRET` – Client secret value.
- `ADMIN_PASSWORD` – Administrator password for `/configure` and `/reports`. Required when `AUTH_TYPE=simple` (the default); the app refuses to start if it is missing or left at the template placeholder.
- `SECRET_KEY` – 64+ character random string for Flask sessions. If omitted, a random key is generated at every start and all users are signed out on each restart.

Generate a strong secret key:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

### Optional Microsoft Entra Sign-In

Set `AUTH_TYPE=oidc` to require Microsoft sign-in for the entire app. The default, `AUTH_TYPE=simple`, keeps the existing public chart and password-based administrator login. OIDC uses a second App Registration, separate from the application-permission registration used to sync org data.

Configure the sign-in App Registration as a **Web** application with redirect URI `https://<your-host>/auth/callback` and grant these delegated Microsoft Graph permissions with tenant admin consent:

- `User.Read`
- `GroupMember.Read.All`

In **Enterprise applications** (not App registrations), open the corresponding application and configure:

1. Under **Properties**, set **Assignment required?** to **Yes**.
2. Under **Users and groups**, assign the reader, privileged, and admin security groups used below.
3. Do not rely on tenant membership alone; the application also verifies membership in one of these three groups during every new login.

Microsoft Entra rejects an unassigned user before redirecting back to the org chart. Users assigned to the Enterprise Application but absent from all three access groups are denied by the application with a 403 error.

Then configure:

| Variable | Description |
| --- | --- |
| `AUTH_TYPE` | Authentication mode: `simple` or `oidc`. Defaults to `simple`. |
| `OIDC_TENANT_ID` | Tenant ID for the sign-in App Registration. |
| `OIDC_CLIENT_ID` | Client ID for the sign-in App Registration. |
| `OIDC_CLIENT_SECRET` | Client secret for the sign-in App Registration. |
| `OIDC_READER_GROUP_ID` | Security group object ID allowed to read the org chart. |
| `OIDC_PRIVILEGED_GROUP_ID` | Security group object ID allowed to sync, use admin XLSX columns, and access reports. |
| `OIDC_ADMIN_GROUP_ID` | Security group object ID granted full configuration access in addition to privileged access. |
| `OIDC_REDIRECT_URI` | Optional callback override when the externally visible URL cannot be inferred. |

Users outside all three configured groups are denied by the application even if Microsoft completes authentication. Membership is resolved through Microsoft Graph's transitive membership check, so nested group membership is honored. Privileged membership takes precedence over reader membership, and admin membership takes precedence over both.

The configuration page displays the three OIDC group IDs as read-only values. Full admin group members always have every application capability. From the same page, an admin can independently grant these capabilities to the privileged group:

- Reports, report exports, and user-scanner tools
- Manual data synchronization
- XLSX columns marked `Restricted`

Existing XLSX column settings stored as `admin` are retained for compatibility but are displayed as `Restricted`; visibility is determined by the restricted-XLSX capability rather than the user's role name.

With `AUTH_TYPE=simple`, the existing public chart and `ADMIN_PASSWORD` administrator flow remain unchanged. `ADMIN_PASSWORD` is not required with `AUTH_TYPE=oidc`.

**Optional values**

| Variable | Default | Description |
| --- | --- | --- |
| `APP_PORT` | `5000` | Port the application listens on. |
| `GUNICORN_TIMEOUT` | `600` | Worker timeout in seconds. Increase if PNG generation times out on very large charts. |
| `CORS_ALLOWED_ORIGINS` | *(none)* | Comma-separated list of allowed cross-origin hosts. |
| `RUN_INITIAL_UPDATE` | `auto` | Set to `true` to force data refresh at startup, `false` to skip. |
| `SESSION_TYPE` | `filesystem` | Flask session backend type. |
| `MAX_FILE_SIZE_MB` | `5` | Maximum upload size (MB) for logos and favicons. |
| `ALLOWED_LOGO_EXTENSIONS` | `png,jpg,jpeg` | Comma-separated list of allowed logo image formats. |
| `ALLOWED_FAVICON_EXTENSIONS` | `ico,png,jpg,jpeg` | Comma-separated list of allowed favicon formats. |
| `PHOTO_CACHE_SECONDS` | `3600` | Browser cache duration (seconds) for profile photos. |
| `PHOTO_CACHE_FILE_SECONDS` | `86400` | File cache duration (seconds) for profile photos on disk. |
| `RATE_LIMIT_DEFAULT` | `200 per day,50 per hour` | Default rate limits for all endpoints. |
| `RATE_LIMIT_LOGIN` | `5 per minute` | Rate limit for the login endpoint. |
| `RATE_LIMIT_PHOTO` | `500 per hour` | Rate limit for the photo endpoint. |
| `RATE_LIMIT_SETTINGS` | `20 per minute` | Rate limit for settings endpoints. |
| `RATE_LIMIT_UPLOAD` | `5 per minute` | Rate limit for file upload endpoints. |
| `RATE_LIMIT_REFRESH` | `1 per minute` | Rate limit for data refresh endpoints. |
| `RATE_LIMIT_PRESENCE` | `60 per minute` | Rate limit for the Teams presence endpoint. |
| `PRESENCE_REFRESH_SECONDS` | `120` | How often the client refreshes Teams presence status (seconds). |
| `PRESENCE_BATCH_SIZE` | `650` | Maximum user IDs per Graph presence batch request (Graph API max is 650). |
| `SECURITY_HEADER_CONTENT_TYPE_OPTIONS` | `nosniff` | `X-Content-Type-Options` header value. |
| `SECURITY_HEADER_FRAME_OPTIONS` | `DENY` | `X-Frame-Options` header value. |
| `SECURITY_HEADER_XSS_PROTECTION` | `1; mode=block` | `X-XSS-Protection` header value. |
| `SECURITY_HEADER_HSTS` | `max-age=31536000; includeSubDomains` | `Strict-Transport-Security` header value. |
| `SECURITY_HEADER_CSP` | `default-src 'self'; ...` | `Content-Security-Policy` header value. |
| `GRAPH_API_ENDPOINT` | `https://graph.microsoft.com/v1.0` | Microsoft Graph API v1.0 endpoint. |
| `GRAPH_API_BETA_ENDPOINT` | `https://graph.microsoft.com/beta` | Microsoft Graph API beta endpoint. |

**SMTP Email Configuration (optional, for automated reports)**

| Variable | Default | Description |
| --- | --- | --- |
| `SMTP_SERVER` | *(none)* | SMTP server hostname (e.g., `smtp.gmail.com`, `smtp.office365.com`). |
| `SMTP_PORT` | `587` | SMTP server port (587 for STARTTLS, 465 for SSL/TLS, 25 for plain). |
| `SMTP_USERNAME` | *(none)* | SMTP username (often your email address). |
| `SMTP_PASSWORD` | *(none)* | SMTP password or app-specific password. |
| `SMTP_FROM_ADDRESS` | *(none)* | From address for sent emails (must be authorized by SMTP server). |
| `SMTP_ENCRYPTION` | `TLS` | Encryption protocol: `TLS` (STARTTLS for port 587), `SSL` (SSL/TLS for port 465), or `None`. |
| `APP_BASE_URL` | `http://localhost:5000` | URL the app uses to reach **itself** when rendering PNG email attachments. Must use the port the app listens on (`APP_PORT`), not an external reverse-proxy address. |

## Running the Application

### Docker (recommended)

```bash
docker compose pull
docker compose up -d
```

- Default port: `APP_PORT` (defaults to `5000`). Override it in `.env` to change container and host bindings.
- Persistent data resides in the `./data` and `./config` bind-mounted directories. Remove their contents to rebuild caches from scratch.
- Local execution outside Docker is not supported; use the provided container workflow for development and production.

Common operations:

| Task | Command |
| --- | --- |
| Follow logs | `docker compose logs -f` |
| Apply `.env` changes | `docker compose up -d` (recreates the container) |
| Update to the latest image | `docker compose pull && docker compose up -d` |
| Stop | `docker compose down` |

### Building from source (contributors only)

Not needed to run the app. If you are changing the code, clone the repository; `docker-compose-dev.yml` builds the image from your working copy instead of pulling it, and publishes it on host port `5002`:

```bash
docker compose -f docker-compose-dev.yml up -d --build
```

Run the test suite with:

```bash
pip install -r requirements.txt pytest
python -m pytest .tests/
```

## Key Features

- **Interactive D3 Org Chart**: Pan, zoom, and expand/collapse hierarchies with persistent hidden subtrees.
- **Teams Presence Indicators**: Live Microsoft Teams status icons on each card (Available, Busy, Do Not Disturb, Away, Out of Office, Offline) with automatic refresh.
- **Search & Discovery**: Real-time directory search, quick navigation helpers, and configurable filters for guests/disabled users.
- **Configuration UI** (`/configure`): Adjust styling, filtering, export columns, scheduling, and email reports without editing files.
- **Admin Reports** (`/reports`):
  - Missing managers
  - Users by last sign-in activity
  - Employees hired in the last 365 days
  - Users hidden by filters
  - Users without profile picture
  - Users without a hiring date
  - Users with data quality issues (whitespace, uppercase emails)
  - User Scanner (OSINT) — individual and organization-wide email presence scans
- **Rich Report Filters**: Toggle groups (mailbox type, account status, license status, user type, GAL visibility, mailbox presence, manager presence) plus tagpicker filters for title, department, country, and state/province. Filters that require Graph permissions not currently granted are automatically greyed out with a tooltip explaining what is missing.
- **Permission-Gated Filters**: On each data sync the app probes which Graph API capabilities are available and persists the result. Filters that require `AuditLog.Read.All` (inactivity ranges), `MailboxSettings.Read` (mailbox type), or Exchange-backed `showInAddressList` (GAL visibility) are disabled in the UI until the corresponding permission is granted and a sync is run.
- **Automated Email Reports**: Schedule daily, weekly, or monthly reports sent via SMTP after data synchronization.
- **Export Options**: SVG/PNG/PDF snapshots and XLSX exports for reports and chart data.
- **Caching & Scheduling**: JSON caches regenerate nightly; manual refresh endpoints keep data current on demand.

## Automated Email Reports

SimpleOrgChart can send scheduled email reports with organization chart data as attachments after each successful data synchronization. **Disabled by default.**

### Prerequisites

1. Configure SMTP environment variables in `.env` (see SMTP Email Configuration section above)
2. Ensure all required SMTP settings are provided: server, port, username, password, and from address
3. Set `APP_BASE_URL` in `.env` to the URL where the app can reach itself: `http://localhost:<APP_PORT>`, which is `http://localhost:5000` unless you changed `APP_PORT`. Inside Docker this is the port the app listens on in the container, regardless of how the port is published on the host.
4. **(Docker users)** PNG screenshots are automatically supported - Playwright is included in the Docker image
5. **(Non-Docker users)** For PNG chart screenshots, install Playwright:
   ```bash
   pip install playwright
   playwright install --with-deps chromium
   ```

### Configuration

1. Navigate to `/configure` and locate the **📧 Email Reports** section
2. Enable **Automated Email Reports** toggle
3. Configure the following options:
   - **Recipient Email**: Email address(es) to receive reports (comma-separated for multiple recipients)
   - **Report Frequency**: Choose daily, weekly, or monthly
   - **Day of Week**: For weekly schedules, select which day to send
   - **Day of Month**: For monthly schedules, select first or last day
   - **Attachment Types**: Select which file formats to include:
     - **XLSX (Excel)**: Employee data spreadsheet (always available)
     - **PNG (Chart Image)**: Visual org chart diagram (requires Playwright)
4. Click **Send Test Email** to verify SMTP configuration
5. Save settings

### How It Works

- Email reports are checked after each **scheduled** data synchronization (manual syncs never send them)
- The check is calendar-based, and at most one email is sent per period:
- For **daily** reports: one email per day
- For **weekly** reports: one email per week, on the configured day of the week
- For **monthly** reports: one email per month, on the first or last day of the month

Email reports support two attachment types:

1. **XLSX (Excel)** - Always available
   - Complete employee directory with name, title, department, email, phone, manager, location details
   - Server-side generation, no additional dependencies

2. **PNG (Chart Image)** - Included in Docker, optional for manual installs
   - Visual organization chart diagram
   - Full chart view with all employees
   - Generated via headless browser screenshot
   - **Docker**: Automatically available (Playwright included in image)
   - **Manual installs**: Requires `playwright install --with-deps chromium`

> **Note**: SVG and PDF exports require client-side rendering and are not available for automated email reports. These formats can still be generated manually from the web interface.

### Test Email

Use the **Send Test Email** button to verify your SMTP configuration before enabling automated reports. Use **Manual Send Now** to immediately send a report with XLSX and/or PNG attachments based on your configuration.

### SMTP Status

The email reports section displays the current SMTP configuration status:
- ✓ **SMTP configured**: Shows the server, port, and from address
- ⚠ **SMTP not configured**: Indicates missing SMTP environment variables

If SMTP is not configured, ensure all required variables are set in your `.env` file and restart the application.

## Teams Presence

SimpleOrgChart displays live Microsoft Teams presence status on each org chart card. **Disabled by default.**

Presence icons follow the standard Teams visual language:

| Icon | Status | Color |
| --- | --- | --- |
| ✓ checkmark | Available | Green |
| solid circle | Busy / In a Call / In a Meeting / Presenting | Red |
| — bar | Do Not Disturb | Red |
| clock | Away / Be Right Back | Yellow |
| → arrow | Out of Office | Purple |
| ✕ cross | Offline / Off Work | Gray |

### Enabling

1. Ensure `Presence.Read.All` (Application) is granted and admin-consented in your Entra ID app registration
2. Navigate to `/configure` and locate the **🟢 Teams Presence** section
3. Toggle **Enable Teams Presence**
4. Save settings

Once enabled, the org chart fetches presence for all visible employees on load and refreshes automatically (every two minutes by default; configurable via the `PRESENCE_REFRESH_SECONDS` environment variable). Each card shows a small Teams-style icon at the top-left corner of the card; hovering displays the detailed status (e.g., "In a Meeting" rather than just "Busy").

The backend calls the Graph `POST /communications/getPresencesByUserId` endpoint in batches of up to `PRESENCE_BATCH_SIZE` user IDs per request (default 650, the Graph API maximum) and relies on Microsoft Graph's built-in rate limiting (it may receive 429 responses if limits are exceeded under heavy load).

## User Scanner (OSINT)

SimpleOrgChart includes a built-in OSINT scanner that checks employee email addresses against hundreds of online platforms to identify where accounts are registered. **Disabled by default.**

The scanner uses the open-source [user-scanner](https://pypi.org/project/user-scanner/) package, which is downloaded automatically from PyPI when the feature is enabled and kept up to date.

### Enabling

1. Navigate to `/configure` and locate the **🔍 User Scanner** section
2. Toggle **Enable User Scanner**
3. Save settings — the package is installed on first use

### Individual Scan

Scan a single employee or any arbitrary email address from the **Individual** tab on the reports page.

1. Go to `/reports` and select **User Scanner (OSINT)** from the report dropdown
2. On the **Individual** tab, search for an employee or type a raw email address
3. Optionally filter by specific sites or categories, toggle loud-site inclusion, or limit to found-only results
4. Click **Run Scan** — results appear inline in a per-platform table showing registration status

### Organization Scan

Scan all employees (or a filtered subset) from the **Organization** tab.

1. Switch to the **Organization** tab
2. Optionally narrow the scan with user filters (mailbox type, account status, license status, user type)
3. Use site and category pickers to target specific platforms
4. Toggle "Include loud sites" to include platforms that may notify the target (e.g., Instagram, Flirtbate)
5. Optionally enter an email address to receive the report when the scan completes
6. Click **Run Full Scan**

During the scan:

- A progress bar and live terminal display scanning status in real time
- The scan can be stopped at any point with **Stop Scan**; partial results are preserved
- Rate limiting (1.5 s between employees) prevents HTTP 429 errors from target platforms

When the scan completes (or is stopped):

- An XLSX workbook is generated with a **Summary** sheet plus one sheet per scanned site
- The last five scan runs are listed under **Recent Scan Results** with one-click **Download** buttons
- If an email address was provided, the XLSX is sent as an attachment
- Use the **Clear** button to remove scan history

### Multi-Worker Support

Organization scans use file-based state (`data/full_scan_state.json`) and a cancel sentinel (`data/full_scan_cancel.flag`) so progress and stop signals work correctly across multiple Gunicorn workers.

## Reporting Caches

- `data/employee_data.json` – Full org hierarchy.
- `data/missing_manager_records.json` – Missing manager snapshot.
- `data/filtered_user_records.json` – Users hidden by org chart filters.
- `data/disabled_user_records.json` – Disabled users enriched with license and sign-in metadata.
- `data/last_login_records.json` – Active users with last sign-in timestamps.
- `data/recently_hired_employees.json` – Users hired in the last 365 days.
- `data/missing_photo_records.json` – Users without a profile picture.
- `data/missing_hire_date_records.json` – Users without a hiring date (`employeeHireDate`).
- `data/dirty_data_records.json` – Users with data quality issues (whitespace, uppercase emails).
- `data/graph_capabilities.json` – Graph API capability flags written on each sync; controls which report filters are enabled in the UI.
- `data/user_scanner_results.json` – Cached results from the most recent organization-wide OSINT scan.
- `data/user_scanner_history.json` – Last five scan run metadata entries.
- `data/user_scanner_exports/` – Downloaded XLSX workbooks for completed scans.

If a cache is missing or stale, click **Sync** in the page header or start the app with `RUN_INITIAL_UPDATE=true`.

## Troubleshooting

| Symptom | Likely cause and fix |
| --- | --- |
| Container exits immediately with `ADMIN_PASSWORD ... must be set` or `must not use the default placeholder value` | `ADMIN_PASSWORD` is missing from `.env` or still set to the template placeholder. Set a real password and run `docker compose up -d`. |
| Chart stays empty; logs show `Missing Azure AD credentials` or `Access token retrieval failed` | One of `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` is wrong, the secret expired, or the secret ID was pasted instead of its value. |
| Sync runs but returns no users or fails with 403 | Admin consent was not granted for the Graph **application** permissions (delegated permissions are not enough). |
| Report filters are greyed out | The permission named in the tooltip is not granted. Grant it, then click **Sync**. |
| Everyone is signed out after each restart | `SECRET_KEY` is not set. |
| `.env` changes have no effect | `docker compose restart` does not re-read `.env`; use `docker compose up -d`. |
| PNG attachment missing from email reports | `APP_BASE_URL` does not point at `http://localhost:<APP_PORT>`, or the chart takes longer than `GUNICORN_TIMEOUT` to render. |
| Cannot write caches / permission errors on Linux | The container fixes ownership of `./data` and `./config` on start; make sure both are bind-mounted as in `docker-compose.yml`. |

## Contributing

Issues and pull requests are welcome. Please document new locale strings in `static/locales/en-US.json`, update report caches when introducing routes, run `python -m pytest .tests/`, and include manual validation steps for anything the automated tests do not cover.
