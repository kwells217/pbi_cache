const { app } = require('@azure/functions');

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const STORAGE_ACCOUNT = process.env.STORAGE_ACCOUNT;
const STORAGE_SAS = process.env.STORAGE_SAS; // e.g. ?sv=2021-06-08&ss=b&...
const CACHE_CONTAINER = 'pbi-cache';
const CACHE_BLOB = 'catalog.json';
const CACHE_TTL_HOURS = 24;

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': process.env.ALLOWED_ORIGIN || '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json'
};

function blobUrl(blob) {
  return `https://${STORAGE_ACCOUNT}.blob.core.windows.net/${CACHE_CONTAINER}/${blob}${STORAGE_SAS}`;
}

function containerUrl() {
  return `https://${STORAGE_ACCOUNT}.blob.core.windows.net/${CACHE_CONTAINER}?restype=container&${STORAGE_SAS.replace(/^\?/, '')}`;
}

async function readBlobCache() {
  try {
    const res = await fetch(blobUrl(CACHE_BLOB));
    if (!res.ok) return null;
    const text = await res.text();
    const cached = JSON.parse(text);
    const ageHours = (Date.now() - cached.timestamp) / (1000 * 60 * 60);
    if (ageHours < CACHE_TTL_HOURS) return cached.data;
    return null;
  } catch (e) { return null; }
}

async function writeBlobCache(data) {
  try {
    // Create container if it doesn't exist
    const cRes = await fetch(containerUrl(), {
      method: 'PUT',
      headers: { 'x-ms-version': '2020-10-02', 'Content-Length': '0' }
    });
    if (cRes.status !== 201 && cRes.status !== 409) {
      console.error('Container create failed:', cRes.status, await cRes.text());
    }

    // Write blob
    const content = JSON.stringify({ timestamp: Date.now(), data });
    const bRes = await fetch(blobUrl(CACHE_BLOB), {
      method: 'PUT',
      headers: {
        'x-ms-blob-type': 'BlockBlob',
        'x-ms-version': '2020-10-02',
        'Content-Type': 'application/json',
        'Content-Length': String(Buffer.byteLength(content, 'utf8'))
      },
      body: content
    });
    if (!bRes.ok) {
      console.error('Blob write failed:', bRes.status, await bRes.text());
    } else {
      console.log('Cache written successfully');
    }
  } catch (e) { console.error('Cache write exception:', e.message); }
}

async function getToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: 'https://analysis.windows.net/powerbi/api/.default'
      })
    }
  );
  const data = await res.json();
  if (data.error) throw new Error(data.error_description || data.error);
  return data.access_token;
}

async function pbiGet(token, path) {
  const res = await fetch(`https://api.powerbi.com/v1.0/myorg${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`Power BI API error ${res.status}: ${path}`);
  return res.json();
}

async function fetchGateways(token) {
  try {
    const data = await pbiGet(token, '/gateways');
    const result = [];
    for (const gw of (data.value || [])) {
      try {
        const ds = await pbiGet(token, `/gateways/${gw.id}/datasources`);
        (ds.value || []).forEach(src => result.push({
          id: src.id,
          displayName: src.displayName || src.datasourceName || src.name || null,
          connectionDetails: src.connectionDetails,
          credentialType: src.credentialType || null,
          username: src.credentialDetails?.username || src.username || null
        }));
      } catch (e) { }
    }
    return result;
  } catch (e) { return []; }
}

async function fetchConnections(token) {
  try {
    const data = await pbiGet(token, '/connections');
    return (data.value || []).map(c => ({
      id: c.id,
      displayName: c.displayName || c.name || null,
      connectionDetails: c.connectionDetails,
      username: c.credentialDetails?.username || c.username || null
    }));
  } catch (e) { return []; }
}

async function buildFullCatalog(token) {
  const [wsData, gatewayData, fabricData] = await Promise.all([
    pbiGet(token, '/groups?$top=1000'),
    fetchGateways(token),
    fetchConnections(token)
  ]);

  const workspaces = wsData.value || [];
  const connections = [...gatewayData, ...fabricData];
  const reports = [], datasets = [], permissions = [];

  for (const ws of workspaces) {
    try {
      const [rData, dData, wsUsers] = await Promise.all([
        pbiGet(token, `/groups/${ws.id}/reports`),
        pbiGet(token, `/groups/${ws.id}/datasets`),
        pbiGet(token, `/groups/${ws.id}/users`).catch(() => ({ value: [] }))
      ]);
      (rData.value || []).forEach(r => { r._workspace = ws.name; r._workspaceId = ws.id; reports.push(r); });
      (dData.value || []).forEach(d => { d._workspace = ws.name; d._workspaceId = ws.id; datasets.push(d); });
      (wsUsers.value || []).forEach(u => permissions.push({
        itemType: 'Workspace', itemName: ws.name, itemId: ws.id, workspace: ws.name,
        identifier: u.emailAddress || u.displayName || u.identifier,
        displayName: u.displayName || u.emailAddress || u.identifier,
        principalType: u.principalType || 'User',
        role: u.groupUserAccessRight || '—',
        accessType: u.principalType === 'Group' ? 'Group' : 'Direct'
      }));
    } catch (e) { }
  }

  for (const ds of datasets) {
    try {
      const [srcData, refreshData, scheduleData, dsUsers] = await Promise.all([
        pbiGet(token, `/groups/${ds._workspaceId}/datasets/${ds.id}/datasources`),
        pbiGet(token, `/groups/${ds._workspaceId}/datasets/${ds.id}/refreshes?$top=1`).catch(() => ({ value: [] })),
        pbiGet(token, `/groups/${ds._workspaceId}/datasets/${ds.id}/refreshSchedule`).catch(() => ({ enabled: false })),
        pbiGet(token, `/groups/${ds._workspaceId}/datasets/${ds.id}/users`).catch(() => ({ value: [] }))
      ]);
      ds._sources = srcData.value || [];
      ds._lastRefresh = (refreshData.value || [])[0] || null;
      ds._schedule = scheduleData || null;
      (dsUsers.value || []).forEach(u => permissions.push({
        itemType: 'Dataset', itemName: ds.name, itemId: ds.id, workspace: ds._workspace,
        identifier: u.emailAddress || u.displayName || u.identifier,
        displayName: u.displayName || u.emailAddress || u.identifier,
        principalType: u.principalType || 'User',
        role: u.datasetUserAccessRight || '—',
        accessType: u.principalType === 'Group' ? 'Group' : 'Direct'
      }));
    } catch (e) { ds._sources = []; ds._lastRefresh = null; ds._schedule = null; }
  }

  for (const r of reports) {
    try {
      const rUsers = await pbiGet(token, `/groups/${r._workspaceId}/reports/${r.id}/users`);
      (rUsers.value || []).forEach(u => permissions.push({
        itemType: 'Report', itemName: r.name, itemId: r.id, workspace: r._workspace,
        identifier: u.emailAddress || u.displayName || u.identifier,
        displayName: u.displayName || u.emailAddress || u.identifier,
        principalType: u.principalType || 'User',
        role: u.reportUserAccessRight || '—',
        accessType: u.principalType === 'Group' ? 'Group' : 'Direct'
      }));
    } catch (e) { }
  }

  return { workspaces, reports, datasets, connections, permissions };
}

app.http('pbi-proxy', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'function',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: CORS_HEADERS };
    }

    const url = new URL(request.url);
    const action = url.searchParams.get('action');

    try {
      if (action === 'catalog') {
        const cached = await readBlobCache();
        if (cached) {
          context.log('Serving from cache');
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ source: 'cache', ...cached }) };
        }
        context.log('Cache miss — fetching from Power BI');
        const token = await getToken();
        const data = await buildFullCatalog(token);
        await writeBlobCache(data);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ source: 'live', ...data }) };
      }

      if (action === 'refresh') {
        context.log('Force refresh');
        const token = await getToken();
        const data = await buildFullCatalog(token);
        await writeBlobCache(data);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ source: 'live', ...data }) };
      }

      return { status: 400, headers: CORS_HEADERS, body: JSON.stringify({ error: 'Unknown action' }) };

    } catch (e) {
      context.error('Proxy error:', e.message);
      return { status: 500, headers: CORS_HEADERS, body: JSON.stringify({ error: e.message }) };
    }
  }
});
