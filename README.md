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
- `ADMIN_PASSWORD` – Protects `/configure` and `/reports`.
- `SECRET_KEY` – 64+ character random string for Flask sessions.

Generate a strong secret key:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

**Optional values**

| Variable | Default | Description |
| --- | --- | --- |
| `APP_PORT` | `5000` | Port the application listens on. |
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
| `APP_BASE_URL` | `http://localhost:5000` | Base URL for generating PNG screenshots (required for PNG email attachments). |

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
- **Configuration UI** (`/configure`): Adjust styling, filtering, export columns, scheduling, and email reports without editing files.
- **Admin Reports** (`/reports`):
  - Missing managers
  - Users by last sign-in activity
  - Employees hired in the last 365 days
  - Users hidden by filters
- **Automated Email Reports**: Schedule daily, weekly, or monthly reports sent via SMTP after data synchronization.
- **Export Options**: SVG/PNG/PDF snapshots and XLSX exports for reports and chart data.
- **MicroSIP Directory Feed**: Serve a MicroSIP contacts JSON at /contacts.json using cached employee data.
- **Desk Phone Directory**: Yealink-compatible XML phonebook at /contacts.xml for T31P, T33G, T46U, and similar models.
- **Caching & Scheduling**: JSON caches regenerate nightly; manual refresh endpoints keep data current on demand.

## Automated Email Reports

SimpleOrgChart can send scheduled email reports with organization chart data as attachments after each successful data synchronization. **Disabled by default.**

### Prerequisites

1. Configure SMTP environment variables in `.env` (see SMTP Email Configuration section above)
2. Ensure all required SMTP settings are provided: server, port, username, password, and from address
3. Set `APP_BASE_URL` in `.env` to the **internal** URL where the app listens:
   - **Docker**: `http://localhost` (port 80 inside container, regardless of external port mapping)
   - **Non-Docker**: `http://localhost:5000` (or whatever port you configured)
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

- Email reports are triggered automatically after successful data synchronization
- The scheduler checks if an email should be sent based on the configured frequency
- For **daily** reports: Emails are sent on the specified day of the week if at least 24 hours have passed since the last email
- For **weekly** reports: Emails are sent on the specified day of the week if at least 7 days have passed
- For **monthly** reports: Emails are sent on the first or last day of the month if at least 28 days have passed

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

## MicroSIP Directory

SimpleOrgChart can publish a MicroSIP-compatible directory export. **Disabled by default.**

1. Navigate to `/configure` and enable **MicroSIP Directory (JSON)**
2. Optionally change the filename (default: `contacts` → `/contacts.json`)
3. Save settings

Once enabled:

- The feed reuses the cached employee list; trigger a manual refresh if the response is empty.
- Response headers disable caching so MicroSIP always retrieves the latest contacts.
- Fields include number, name, firstname, lastname, phone, mobile, email, address, city, state, comment, presence, starred, and info.
- Contacts without a desk or mobile number are skipped to keep the directory free of unreachable entries.
- Add extra entries from the Configure → Custom Directory Contacts textarea (one Name,Number pair per line) when you need off-chart contacts in the feed.

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

## Desk Phone Directory (Yealink XML)

SimpleOrgChart can provide a Yealink-compatible remote phonebook. **Disabled by default.**

1. Navigate to `/configure` and enable **Desk Phone Directory (XML)**
2. Optionally change the filename (default: `contacts` → `/contacts.xml`)
3. Save settings

Once enabled:

- Compatible with Yealink T31P, T33G, T46U, and other models supporting remote XML phonebooks.
- Reuses the same cached employee list and custom contacts as the MicroSIP directory.
- Contacts without a desk or mobile number are omitted.
- Employees with both office and mobile phones will have both numbers listed.
- The phonebook title is derived from your configured Chart Title.

### Yealink Phone Configuration

1. Access your phone's web interface (typically `http://<phone-ip>`)
2. Navigate to **Directory** → **Remote Phone Book**
3. Add a new remote phonebook entry:
   - **Remote URL**: `http://<your-server>:5000/<filename>.xml` (use your configured filename)
   - **Display Name**: `Company Directory`
4. Save and reboot the phone if required

Example XML structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<YealinkIPPhoneDirectory>
  <Title>Organization Directory</Title>
  <DirectoryEntry>
    <Name>Ada Lovelace</Name>
    <Telephone>555-123-4567</Telephone>
    <Telephone>555-987-6543</Telephone>
  </DirectoryEntry>
  <DirectoryEntry>
    <Name>Charles Babbage</Name>
    <Telephone>555-111-2222</Telephone>
  </DirectoryEntry>
</YealinkIPPhoneDirectory>
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
