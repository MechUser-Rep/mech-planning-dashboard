// graph.js
// Requires auth.js loaded before this script.

let _siteId = null;

async function graphFetch(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${url}\n${text}`);
  }
  return res.json();
}

async function getSiteId(token) {
  if (_siteId) return _siteId;
  const url = `https://graph.microsoft.com/v1.0/sites/${CONFIG.sharePointHostname}:${CONFIG.sitePath}`;
  const data = await graphFetch(url, token);
  _siteId = data.id;
  return _siteId;
}

// Returns the used range as a 2D array of values: rows[rowIndex][colIndex]
async function getExcelUsedRange(token, filePath, sheetName) {
  const siteId = await getSiteId(token);
  const encodedPath = filePath.split('/').map(p => encodeURIComponent(p)).join('/');
  const encodedSheet = encodeURIComponent(sheetName);
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/workbook/worksheets/${encodedSheet}/usedRange`;
  const data = await graphFetch(url, token);
  return data.values; // 2D array
}

// Returns array of worksheet name strings for a workbook
async function getWorksheetNames(token, filePath) {
  const siteId = await getSiteId(token);
  const encodedPath = filePath.split('/').map(p => encodeURIComponent(p)).join('/');
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/workbook/worksheets`;
  const data = await graphFetch(url, token);
  return data.value.map(ws => ws.name);
}
