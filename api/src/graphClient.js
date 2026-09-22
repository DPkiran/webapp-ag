const axios = require('axios');

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  SITE_HOSTNAME,
  SITE_PATH,
  LIBRARY_NAME,   // optional — set this if the file isn't in the default "Documents" library
  FILE_PATH,
  TABLE_NAME,
} = process.env;

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

// In-memory caches. Note: Azure Functions consumption plan instances can be
// recycled/cold-started at any time, so these are a best-effort optimization,
// not guaranteed persistence — that's fine, they just get rebuilt on the next call.
let cachedToken = null;   // { value, expiresAt }
let cachedSiteId = null;
let cachedDriveId = null;
let cachedItemId = null;

function assertConfig() {
  // LIBRARY_NAME is intentionally optional — omit it to use the site's default library
  const required = { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_HOSTNAME, SITE_PATH, FILE_PATH, TABLE_NAME };
  const missing = Object.entries(required).filter(([, v]) => !v).map(([k]) => k);
  if (missing.length) {
    throw new Error(`Missing required application settings: ${missing.join(', ')}`);
  }
}

async function getAccessToken() {
  if (cachedToken && Date.now() < cachedToken.expiresAt - 60_000) {
    return cachedToken.value;
  }
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const { data } = await axios.post(tokenUrl, params, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });

  cachedToken = {
    value: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };
  return cachedToken.value;
}

async function graphGet(url, token) {
  try {
    const { data } = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    return data;
  } catch (err) {
    // TEMP: attach exactly what URL was requested, to debug "not found" errors
    err.requestedUrl = url;
    throw err;
  }
}

async function resolveSiteId(token) {
  if (cachedSiteId) return cachedSiteId;
  // SITE_PATH must be just the site's own path, e.g. "/sites/Bltest" —
  // no https://, no hostname, and no document-library name folded in.
  const url = `${GRAPH_ROOT}/sites/${SITE_HOSTNAME}:${SITE_PATH}`;
  const data = await graphGet(url, token);
  cachedSiteId = data.id;
  return cachedSiteId;
}

async function resolveDriveId(token, siteId) {
  if (cachedDriveId) return cachedDriveId;

  if (!LIBRARY_NAME) {
    // No library specified — use the site's default document library.
    const url = `${GRAPH_ROOT}/sites/${siteId}/drive`;
    const data = await graphGet(url, token);
    cachedDriveId = data.id;
    return cachedDriveId;
  }

  // A specific (non-default) library was named — find it among the site's drives.
  const url = `${GRAPH_ROOT}/sites/${siteId}/drives`;
  const data = await graphGet(url, token);
  const match = data.value.find(
    (d) => d.name && d.name.toLowerCase() === LIBRARY_NAME.toLowerCase()
  );
  if (!match) {
    const available = data.value.map((d) => d.name).join(', ');
    throw new Error(`No document library named "${LIBRARY_NAME}" found. Available libraries: ${available}`);
  }
  cachedDriveId = match.id;
  return cachedDriveId;
}

async function resolveFileItemId(token, driveId) {
  if (cachedItemId) return cachedItemId;
  // FILE_PATH is the path *within that library*, e.g. "/Test WEBapps.xlsx"
  // — don't include the library name itself here, that's LIBRARY_NAME's job.
  const encodedPath = FILE_PATH.split('/').map(encodeURIComponent).join('/');
  const url = `${GRAPH_ROOT}/drives/${driveId}/root:${encodedPath}`;
  const data = await graphGet(url, token);
  cachedItemId = data.id;
  return cachedItemId;
}

async function fetchTableRows() {
  assertConfig();
  const token = await getAccessToken();
  const siteId = await resolveSiteId(token);
  const driveId = await resolveDriveId(token, siteId);
  const itemId = await resolveFileItemId(token, driveId);

  const columnsUrl = `${GRAPH_ROOT}/drives/${driveId}/items/${itemId}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/columns`;
  const rowsUrl = `${GRAPH_ROOT}/drives/${driveId}/items/${itemId}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows`;

  const [columnsData, rowsData] = await Promise.all([
    graphGet(columnsUrl, token),
    graphGet(rowsUrl, token),
  ]);

  const headers = columnsData.value
    .sort((a, b) => a.index - b.index)
    .map((c) => c.name);

  const rows = rowsData.value.map((r) => r.values[0]);

  return { headers, rows };
}

function clearCaches() {
  cachedSiteId = null;
  cachedDriveId = null;
  cachedItemId = null;
}

module.exports = { fetchTableRows, clearCaches };
