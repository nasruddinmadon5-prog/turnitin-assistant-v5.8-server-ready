/**
 * Turnitin Assistant v5.8-server-ready — Remote Dashboard JS
 * Connects to the FastAPI backend running on your VPS.
 */

function getBaseUrl() {
  return document.getElementById('apiUrl').value.replace(/\/+$/, '');
}

function getHeaders() {
  const token = document.getElementById('apiToken').value;
  const headers = { 'Content-Type': 'application/json' };
  if (token) {
    headers['Authorization'] = `Bearer ${token}`;
  }
  return headers;
}

function appendLog(msg, level) {
  const el = document.getElementById('logOutput');
  const ts = new Date().toLocaleTimeString();
  el.textContent += `[${ts}] [${level}] ${msg}\n`;
  el.scrollTop = el.scrollHeight;
}

async function apiFetch(endpoint, options = {}) {
  const url = `${getBaseUrl()}${endpoint}`;
  const headers = getHeaders();
  if (options.body) {
    headers['Content-Type'] = 'application/json';
  }
  try {
    const res = await fetch(url, { ...options, headers });
    const data = await res.json();
    return { ok: res.ok, status: res.status, data };
  } catch (err) {
    return { ok: false, status: 0, data: { error: err.message } };
  }
}

async function checkHealth() {
  const el = document.getElementById('healthResult');
  el.innerHTML = 'Connecting...';
  const { ok, data } = await apiFetch('/health');
  if (ok) {
    el.innerHTML = `<span style="color:#16a34a;">✅ Connected — ${data.version} (${data.status})</span>`;
    appendLog(`Health check OK: ${data.version}`, 'INFO');
    await fetchStatus();
  } else {
    el.innerHTML = `<span style="color:#dc2626;">❌ Connection failed: ${data.error || 'Unknown error'}</span>`;
    appendLog(`Health check FAILED: ${data.error || 'Unknown'}`, 'ERROR');
  }
}

async function fetchStatus() {
  const card = document.getElementById('statusCard');
  const content = document.getElementById('statusContent');
  card.style.display = 'block';
  content.innerHTML = 'Loading...';

  const { ok, data } = await apiFetch('/api/status');
  if (!ok) {
    content.innerHTML = `<span style="color:#dc2626;">Error: ${data.error || 'Failed to fetch status'}</span>`;
    return;
  }

  // Workflow status grid
  let wfHtml = '<div class="wf-grid">';
  const wfLabels = {
    ojs_download: 'OJS Download',
    turnitin_upload: 'Turnitin Upload',
    template_screening: 'Template Screening',
    ojs_report_upload: 'OJS Report Upload'
  };
  for (const [key, wf] of Object.entries(data.workflows || {})) {
    const label = wfLabels[key] || key;
    const statusClass = `status-${wf.status}`;
    wfHtml += `
      <div class="wf-card">
        <div class="wf-name">${label}</div>
        <div class="wf-status"><span class="status-badge ${statusClass}">${wf.status}</span></div>
        <div class="wf-msg">${wf.message || ''}</div>
      </div>
    `;
  }
  wfHtml += '</div>';

  const meExists = data.master_excel_exists
    ? '<span style="color:#16a34a;">✅ Exists</span>'
    : '<span style="color:#dc2626;">❌ Not found</span>';

  content.innerHTML = `
    <p><strong>Version:</strong> ${data.version} | <strong>Mode:</strong> ${data.mode}</p>
    <p><strong>Master Excel:</strong> ${meExists}</p>
    <p><strong>Storage:</strong> OJS: ${data.storage?.ojs_downloads}<br/>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       Turnitin: ${data.storage?.turnitin_reports}<br/>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       Screening: ${data.storage?.screening_reports}</p>
    <h3 style="margin-top:0.75rem; font-size:0.9rem; color:#38bdf8;">Workflow Status</h3>
    ${wfHtml}
  `;

  appendLog('Status fetched successfully', 'INFO');
}

async function startWorkflow(name) {
  const endpoint = `/api/${name}/start`;
  appendLog(`Starting workflow: ${name}...`, 'INFO');

  const { ok, data } = await apiFetch(endpoint, { method: 'POST' });
  if (ok) {
    appendLog(`✅ ${name} — ${data.status || 'started'}`, 'INFO');
  } else if (data.status === 'running') {
    appendLog(`⚠️ ${name} — already running`, 'WARN');
  } else {
    appendLog(`❌ ${name} — ${data.error || 'failed'}`, 'ERROR');
  }

  // Refresh status after a short delay
  setTimeout(fetchStatus, 1000);
}

async function fetchLogs() {
  const el = document.getElementById('logOutput');
  el.textContent = 'Loading logs...';

  const { ok, data } = await apiFetch('/api/logs');
  if (!ok) {
    el.textContent = `Error: ${data.error || 'Failed to fetch logs'}`;
    return;
  }

  const logs = data.logs || [];
  if (logs.length === 0) {
    el.textContent = 'No logs available.';
    return;
  }

  el.textContent = logs.map(e => `[${e.timestamp}] ${e.message}`).join('\n');
  el.scrollTop = el.scrollHeight;
}