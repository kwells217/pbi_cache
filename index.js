const { app } = require('@azure/functions');
const { BlobServiceClient } = require('@azure/storage-blob');

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const STORAGE_CONNECTION = process.env.STORAGE_CONNECTION;
const CACHE_CONTAINER = 'pbi-cache';
const CACHE_BLOB = 'catalog.json';
const CACHE_TTL_HOURS = 24;

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': process.env.ALLOWED_ORIGIN || '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json'
};

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

async function readCache() {
  try {
    const blobClient = BlobServiceClient.fromConnectionString(STORAGE_CONNECTION)
      .getContainerClient(CACHE_CONTAINER)
      .getBlobClient(CACHE_BLOB);
    const download = await blobClient.download();
    const text = await streamToText(download.readableStreamBody);
    const cached = JSON.parse(text);
    const age = (Date.now() - cached.timestamp) / (1000 * 60 * 60);
    if (age < CACHE_TTL_HOURS) return cached.data;
    return null;
  } catch (e) { return null; }
}

async function writeCache(data) {
  try {
    const containerClient = BlobServiceClient.fromConnectionString(STORAGE_CONNECTION)
      .getContainerClient(CACHE_CONTAINER);
    await containerClient.createIfNotExists();
    const blobClient = containerClient.getBlockBlobClient(CACHE_BLOB);
    const content = JSON.stringify({ timestamp: Date.now(), data });
    await blobClient.upload(content, Buffer.byteLength(content), {
      blobHTTPHeaders: { blobContentType: 'application/json' }
    });
  } catch (e) { console.error('Cache write failed:', e.message); }
}

async function streamToText(stream) {
  const chunks = [];
  for await (const chunk of stream) chunks.push(chunk);
  return Buffer.concat(chunks).toString('utf8');
}

async function buildFullCatalog(token) {
  const [wsData, gatewayData, fabricData] = await Promise.all([
    pbiGet(token, '/groups?$top=1000'),
    fetchGateways(token),
    fetchConnections(token)
  ]);

  const workspaces = wsData.value || [];
  const connections = [...gatewayData, ...fabricData];
  const reports = [], datasets = [];

  for (const ws of workspaces) {
    try {
      const [rData, dData] = await Promise.all([
        pbiGet(token, `/groups/${ws.id}/reports`),
        pbiGet(token, `/groups/${ws.id}/datasets`)
      ]);
      (rData.value || []).forEach(r => { r._workspace = ws.name; r._workspaceId = ws.id; reports.push(r); });
      (dData.value || []).forEach(d => { d._workspace = ws.name; d._workspaceId = ws.id; datasets.push(d); });
    } catch (e) { /* skip inaccessible */ }
  }

  for (const ds of datasets) {
    try {
      const [srcData, refreshData, scheduleData] = await Promise.all([
        pbiGet(token, `/groups/${ds._workspaceId}/datasets/${ds.id}/datasources`),
        pbiGet(token, `/groups/${ds._workspaceId}/datasets/${ds.id}/refreshes?$top=1`).catch(() => ({ value: [] })),
        pbiGet(token, `/groups/${ds._workspaceId}/datasets/${ds.id}/refreshSchedule`).catch(() => ({ enabled: false }))
      ]);
      ds._sources = srcData.value || [];
      ds._lastRefresh = (refreshData.value || [])[0] || null;
      ds._schedule = scheduleData || null;
    } catch (e) { ds._sources = []; ds._lastRefresh = null; ds._schedule = null; }
  }

  return { workspaces, reports, datasets, connections };
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
          connectionDetails: src.connectionDetails
        }));
      } catch (e) { /* skip */ }
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
      connectionDetails: c.connectionDetails
    }));
  } catch (e) { return []; }
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
      // Serve full catalog from cache
      if (action === 'catalog') {
        const cached = await readCache();
        if (cached) {
          context.log('Serving from cache');
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ source: 'cache', ...cached }) };
        }
        // Cache miss — build fresh
        context.log('Cache miss — fetching from Power BI');
        const token = await getToken();
        const data = await buildFullCatalog(token);
        await writeCache(data);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ source: 'live', ...data }) };
      }

      // Force refresh cache
      if (action === 'refresh') {
        context.log('Force refreshing cache');
        const token = await getToken();
        const data = await buildFullCatalog(token);
        await writeCache(data);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ source: 'live', ...data }) };
      }

      return {
        status: 400,
        headers: CORS_HEADERS,
        body: JSON.stringify({ error: 'Unknown action. Use: catalog, refresh' })
      };

    } catch (e) {
      context.error('Proxy error:', e.message);
      return {
        status: 500,
        headers: CORS_HEADERS,
        body: JSON.stringify({ error: e.message })
      };
    }
  }
});


const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': process.env.ALLOWED_ORIGIN || '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json'
};

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
      const token = await getToken();

      if (action === 'workspaces') {
        const data = await pbiGet(token, '/groups?$top=1000');
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      if (action === 'reports') {
        const wsId = url.searchParams.get('wsId');
        const data = await pbiGet(token, `/groups/${wsId}/reports`);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      if (action === 'datasets') {
        const wsId = url.searchParams.get('wsId');
        const data = await pbiGet(token, `/groups/${wsId}/datasets`);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      if (action === 'datasources') {
        const wsId = url.searchParams.get('wsId');
        const dsId = url.searchParams.get('dsId');
        const data = await pbiGet(token, `/groups/${wsId}/datasets/${dsId}/datasources`);
        return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
      }

      // New: Get dataset refresh history (for last refreshed date)
      if (action === 'refreshHistory') {
        const wsId = url.searchParams.get('wsId');
        const dsId = url.searchParams.get('dsId');
        try {
          const data = await pbiGet(token, `/groups/${wsId}/datasets/${dsId}/refreshes?$top=1`);
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
        } catch (e) {
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ value: [] }) };
        }
      }

      // New: Get dataset refresh schedule
      if (action === 'refreshSchedule') {
        const wsId = url.searchParams.get('wsId');
        const dsId = url.searchParams.get('dsId');
        try {
          const data = await pbiGet(token, `/groups/${wsId}/datasets/${dsId}/refreshSchedule`);
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
        } catch (e) {
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ enabled: false }) };
        }
      }

      // Get all gateways and their datasources for display names
      if (action === 'gateways') {
        try {
          const data = await pbiGet(token, '/gateways');
          const gateways = data.value || [];
          // Fetch datasources for each gateway
          const result = [];
          for (const gw of gateways) {
            try {
              const ds = await pbiGet(token, `/gateways/${gw.id}/datasources`);
              (ds.value || []).forEach(src => result.push({
                id: src.id,
                gatewayId: gw.id,
                displayName: src.displayName || src.datasourceName || src.name || null,
                connectionDetails: src.connectionDetails
              }));
            } catch (e) { /* skip inaccessible */ }
          }
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ value: result }) };
        } catch (e) {
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ value: [] }) };
        }
      }

      // Get all Fabric managed connections for display names
      if (action === 'connections') {
        try {
          const data = await pbiGet(token, '/connections');
          const result = (data.value || []).map(c => ({
            id: c.id,
            displayName: c.displayName || c.name || null,
            connectionDetails: c.connectionDetails
          }));
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ value: result }) };
        } catch (e) {
          return { status: 200, headers: CORS_HEADERS, body: JSON.stringify({ value: [] }) };
        }
      }

      return {
        status: 400,
        headers: CORS_HEADERS,
        body: JSON.stringify({ error: 'Unknown action' })
      };

    } catch (e) {
      context.error('Proxy error:', e.message);
      return {
        status: 500,
        headers: CORS_HEADERS,
        body: JSON.stringify({ error: e.message })
      };
    }
  }
});
