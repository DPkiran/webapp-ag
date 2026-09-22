const axios = require('axios');

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  SITE_HOSTNAME,
  SITE_PATH,
  FILE_PATH,
  TABLE_NAME,
} = process.env;

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

// In-memory caches. Note: Azure Functions consumption plan instances can be
// recycled/cold-started at any time, so these are a best-effort optimization,
// not guaranteed persistence — that's fine, they just get rebuilt on the next call.
let cachedToken = null;   // { value, expiresAt }
let cachedSiteId = null;
let cachedItemId = null;

function assertConfig() {
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
  const { data } = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  return data;
}

async function resolveSiteId(token) {
  if (cachedSiteId) return cachedSiteId;
  const url = `${GRAPH_ROOT}/sites/${SITE_HOSTNAME}:${SITE_PATH}`;
  const data = await graphGet(url, token);
  cachedSiteId = data.id;
  return cachedSiteId;
}

async function resolveFileItemId(token, siteId) {
  if (cachedItemId) return cachedItemId;
  const encodedPath = FILE_PATH.split('/').map(encodeURIComponent).join('/');
  const url = `${GRAPH_ROOT}/sites/${siteId}/drive/root:${encodedPath}`;
  const data = await graphGet(url, token);
  cachedItemId = data.id;
  return cachedItemId;
}

async function fetchTableRows() {
  assertConfig();
  const token = await getAccessToken();
  const siteId = await resolveSiteId(token);
  const itemId = await resolveFileItemId(token, siteId);

  const columnsUrl = `${GRAPH_ROOT}/sites/${siteId}/drive/items/${itemId}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/columns`;
  const rowsUrl = `${GRAPH_ROOT}/sites/${siteId}/drive/items/${itemId}/workbook/tables/${encodeURIComponent(TABLE_NAME)}/rows`;

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
  cachedItemId = null;
}

module.exports = { fetchTableRows, clearCaches };
