# Employee Request Form — Live SharePoint Excel Viewer

Static frontend + Azure Functions API, deployed together as an **Azure Static
Web App** with CI/CD from GitHub. The Functions app authenticates to
Microsoft Graph with app-only (client credentials) auth and reads your
Excel table on SharePoint; the client secret lives only in the Static Web
App's configuration, never in the repo or the browser.

```
/public                    → static frontend (index.html)
/api                        → Azure Functions app (Node.js v4 model)
  /src/graphClient.js        → Graph auth + table-read logic
  /src/functions/getData.js  → GET /api/data
  /src/functions/getConfig.js→ GET /api/config
staticwebapp.config.json    → SWA routing config
.github/workflows/          → CI/CD workflow
```

---

## 1. App registration (Entra ID / Azure AD)

1. **Entra ID → App registrations → New registration**
   - Name: e.g. `SharePoint Excel Viewer`
   - Account type: *Accounts in this organizational directory only*
   - No redirect URI needed — this is a daemon/service app.
2. Copy the **Application (client) ID** and **Directory (tenant) ID** from
   the Overview page.
3. **Certificates & secrets → New client secret** — copy the secret **value**
   immediately (it's hidden after you navigate away).
4. **API permissions → Add a permission → Microsoft Graph → Application
   permissions → `Sites.Read.All`** (or `Sites.Selected` if you want to scope
   it to just this one site — ask if you want that tighter setup instead).
5. **Grant admin consent** for the tenant.

## 2. Find your SharePoint path values

- `SITE_HOSTNAME` — e.g. `powercen.sharepoint.com`
- `SITE_PATH` — e.g. `/sites/EmployeeRequests`
- `FILE_PATH` — path to the xlsx inside the doc library, e.g.
  `/Shared Documents/EmployeeRequestForm.xlsx`
- `TABLE_NAME` — open the file in Excel, click the data range, check the
  **Table Design** ribbon tab for the table's name (default is often `Table1`).

## 3. Create the Azure Static Web App and link GitHub

1. Push this repo to GitHub first.
2. In the Azure Portal: **Create a resource → Static Web App**.
3. Under **Deployment details**, sign in to GitHub and pick this repo/branch.
   Azure will offer to auto-generate the GitHub Actions workflow and add the
   `AZURE_STATIC_WEB_APPS_API_TOKEN` secret to the repo for you — let it.
4. Build details:
   - **Build presets:** Custom
   - **App location:** `/public`
   - **Api location:** `api`
   - **Output location:** *(leave blank)*

   If Azure's auto-generated workflow doesn't match, compare it against
   `.github/workflows/azure-static-web-apps.yml` here and adjust the
   `app_location` / `api_location` / `output_location` values to match.
5. Once created, every push to the linked branch redeploys automatically —
   that's your CI/CD.

## 4. Set the secrets (never in the repo)

In the Static Web App resource: **Configuration → Application settings**,
add:

| Name | Value |
|---|---|
| `TENANT_ID` | from step 1 |
| `CLIENT_ID` | from step 1 |
| `CLIENT_SECRET` | from step 1 |
| `SITE_HOSTNAME` | from step 2 |
| `SITE_PATH` | from step 2 |
| `FILE_PATH` | from step 2 |
| `TABLE_NAME` | from step 2 |
| `POLL_INTERVAL_SECONDS` | e.g. `20` |

These are injected into the Functions app's environment at runtime — the
`api/local.settings.json.example` file is only a template for local testing
and is gitignored once you copy it to `local.settings.json`.

## 5. Local development (optional, before pushing)

```bash
cd api
cp local.settings.json.example local.settings.json
# fill in local.settings.json with your real values
npm install
npm start          # requires Azure Functions Core Tools installed
```

For a full local SWA experience (frontend + API together, matching
production routing), use the SWA CLI:

```bash
npm install -g @azure/static-web-apps-cli
swa start public --api-location api
```

## Notes

- `staticwebapp.config.json` sets `/api/*` as anonymously callable (SWA's
  platform still fronts it — this just governs SWA's own role-based routing,
  it doesn't change how Graph auth works).
- If the SharePoint file ever moves/renames, the next `/api/data` call fails
  once, then self-heals (cached site/item IDs are cleared and re-resolved).
- Azure Functions on the SWA-managed plan may cold-start after idle periods —
  the first poll after a while can take a couple seconds longer.
