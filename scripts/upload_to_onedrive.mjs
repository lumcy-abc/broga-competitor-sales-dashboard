// Uploads the weekly competitor workbook to a SharePoint/OneDrive library folder via Microsoft Graph.
// Needs these GitHub Actions secrets: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, ONEDRIVE_DRIVE_ID, ONEDRIVE_FOLDER_PATH
// (ONEDRIVE_FOLDER_PATH = folder inside the library, e.g. "Intern/03_Outputs/Dashboards/Broga Competitor Sales (TEDUH)").
import fs from 'node:fs';

const FILES = ['Broga_Weekly_Tracking.xlsx'];
const required = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'ONEDRIVE_DRIVE_ID', 'ONEDRIVE_FOLDER_PATH'];
const missing = required.filter(k => !process.env[k]);
if (missing.length) throw new Error(`Missing GitHub Actions secrets: ${missing.join(', ')}`);
// Check before uploading anything: the workbook exists and is not empty.
for (const f of FILES) {
  if (!fs.existsSync(f) || fs.statSync(f).size === 0) throw new Error(`Workbook missing or empty: ${f}`);
}

const form = new URLSearchParams({
  client_id: process.env.MS_CLIENT_ID,
  client_secret: process.env.MS_CLIENT_SECRET,
  scope: 'https://graph.microsoft.com/.default',
  grant_type: 'client_credentials',
});
const tokenRes = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`, {
  method: 'POST',
  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  body: form,
});
const tokenText = await tokenRes.text();
if (!tokenRes.ok) throw new Error(`Microsoft Graph token request failed (${tokenRes.status}): ${tokenText.slice(0, 500)}`);
const token = JSON.parse(tokenText).access_token;

const folder = process.env.ONEDRIVE_FOLDER_PATH.split('/').filter(Boolean).map(encodeURIComponent).join('/');
const failures = [];
for (const f of FILES) {
  const body = fs.readFileSync(f);
  const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(process.env.ONEDRIVE_DRIVE_ID)}/root:/${folder}/${encodeURIComponent(f)}:/content`;
  try {
    const res = await fetch(url, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body,
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`HTTP ${res.status}: ${text.slice(0, 300)}`);
    const result = JSON.parse(text);
    // Reconciliation: the size SharePoint reports must equal the size we sent.
    if (result.size !== body.length) throw new Error(`size mismatch (sent ${body.length}, library reports ${result.size})`);
    console.log(`Uploaded ${f} (${body.length} bytes) -> ${result.webUrl ?? 'no web URL returned'}`);
  } catch (err) {
    failures.push(`${f}: ${err.message}`);
    console.error(`UPLOAD FAILED ${f}: ${err.message}`);
  }
  await new Promise(r => setTimeout(r, 1000));
}
if (failures.length) throw new Error(`${failures.length} of ${FILES.length} uploads failed:\n${failures.join('\n')}`);
console.log(`Weekly workbook uploaded.`);
