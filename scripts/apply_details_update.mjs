// Applies an "Update details" GitHub issue to data/property-details.json.
// Run by .github/workflows/update-details.yml with the issue body in ISSUE_BODY.
// Each survey is kept as a dated entry; the dashboard shows the latest one and the
// weekly Excel Master View uses each entry from its survey date onwards.
// Prints a short result for the issue comment; exits 1 (changing nothing) if the request is invalid.
import fs from 'node:fs/promises';

const FILE = 'data/property-details.json';
const FIELDS = ['land area', 'bua', 'spa price', 'nett price', 'rebate', 'remarks'];
const body = process.env.ISSUE_BODY ?? '';
const issue = Number(process.env.ISSUE_NUMBER) || null;

function fail(message) {
  console.log(`Not applied: ${message}`);
  process.exit(1);
}

// The form sits inside a ``` block: "project:" and "survey date:" lines, then [field] sections.
const block = body.match(/```[^\n]*\n([\s\S]*?)```/)?.[1] ?? body;
const lines = block.replace(/\r/g, '').split('\n');
const header = name => lines.find(l => l.toLowerCase().startsWith(name + ':'))?.slice(name.length + 1).trim() ?? '';
const project = header('project');
const surveyDate = header('survey date');
const details = {};
let current = null;
for (const line of lines) {
  const m = line.match(/^\[(.+)\]\s*$/);
  if (m) { current = m[1].trim().toLowerCase(); if (current !== 'end' && !FIELDS.includes(current)) fail(`unknown section [${m[1]}].`); continue; }
  if (current && current !== 'end') details[current] = details[current] === undefined ? line : `${details[current]}\n${line}`;
}
for (const k of Object.keys(details)) details[k] = details[k].trim();
for (const k of Object.keys(details)) if (!details[k]) delete details[k];

const html = await fs.readFile('index.html', 'utf8');
const start = html.indexOf('const DATA=') + 'const DATA='.length;
const data = JSON.parse(html.slice(start, start + html.slice(start).search(/;\r?\nconst \$/)));
if (!project) fail('the issue has no "project:" line.');
if (!data.projects.includes(project)) fail(`"${project}" is not a dashboard project (${data.projects.join(', ')}).`);
if (!/^\d{4}-\d{2}-\d{2}$/.test(surveyDate) || Number.isNaN(Date.parse(surveyDate))) fail(`survey date "${surveyDate}" must be YYYY-MM-DD.`);
const todayMY = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Kuala_Lumpur' }).format(new Date());
if (surveyDate > todayMY) fail(`survey date ${surveyDate} is in the future.`);
if (!Object.keys(details).length) fail('no detail values were filled in.');

const store = await fs.readFile(FILE, 'utf8').then(JSON.parse).catch(err => {
  if (err.code !== 'ENOENT') throw err;
  return { schemaVersion: 1, about: 'Survey updates to land area, BUA, SPA, nett price, rebate and remarks, saved from the dashboard through "Update details" GitHub issues. Base values (13/10/2026) live in index.html.', projects: {} };
});
const list = store.projects[project] ??= [];
const replaced = list.findIndex(e => e.surveyDate === surveyDate);
const entry = { surveyDate, details, issue, savedOn: todayMY };
if (replaced >= 0) list[replaced] = entry; else list.push(entry);
list.sort((a, b) => a.surveyDate.localeCompare(b.surveyDate));
await fs.writeFile(FILE, JSON.stringify(store, null, 2) + '\n');

const shown = Object.entries(details).map(([k, v]) => `- **${k}**: ${v.replace(/\n/g, ' / ')}`).join('\n');
console.log(`Applied: ${project} details for survey date ${surveyDate}${replaced >= 0 ? ' (replaced the earlier entry for that date)' : ''}.\n\n${shown}\n\nThe dashboard shows them on its next refresh; the weekly Excel Master View uses them from ${surveyDate} onwards after its next run.`);
