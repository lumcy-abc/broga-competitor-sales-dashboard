import fs from 'node:fs';

const required = ['MS_TENANT_ID','MS_CLIENT_ID','MS_CLIENT_SECRET','ONEDRIVE_DRIVE_ID','ONEDRIVE_ITEM_PATH'];
const missing = required.filter(k => !process.env[k]);
if (missing.length) throw new Error(`Missing GitHub Actions secrets: ${missing.join(', ')}`);
const filePath = 'Broga_Competitor_Sales_Analysis.xlsx';
if (!fs.existsSync(filePath)) throw new Error(`Workbook not found: ${filePath}`);

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
if (!tokenRes.ok) throw new Error(`Microsoft Graph token request failed (${tokenRes.status}): ${tokenText.slice(0,500)}`);
const token = JSON.parse(tokenText).access_token;
const encodedPath = process.env.ONEDRIVE_ITEM_PATH.split('/').filter(Boolean).map(encodeURIComponent).join('/');
const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(process.env.ONEDRIVE_DRIVE_ID)}/root:/${encodedPath}:/content`;
const uploadRes = await fetch(url, {
  method: 'PUT',
  headers: {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  body: fs.readFileSync(filePath),
});
const uploadText = await uploadRes.text();
if (!uploadRes.ok) throw new Error(`OneDrive upload failed (${uploadRes.status}): ${uploadText.slice(0,700)}`);
const result = JSON.parse(uploadText);
console.log(`OneDrive workbook updated: ${result.name ?? process.env.ONEDRIVE_ITEM_PATH}`);
console.log(`Item ID: ${result.id ?? 'not returned'}; web URL: ${result.webUrl ?? 'not returned'}`);
