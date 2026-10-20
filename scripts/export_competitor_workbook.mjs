import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

const ROOT = process.cwd();
const HTML_PATH = path.join(ROOT, 'index.html');
const HISTORY_PATH = path.join(ROOT, 'data', 'teduh-history.json');
const OUT_PATH = path.join(ROOT, 'Broga_Competitor_Sales_Analysis.xlsx');

function loadEmbeddedData(html) {
  const marker = 'const DATA=';
  const startAt = html.indexOf(marker);
  if (startAt < 0) throw new Error('Could not find the embedded DATA object in index.html.');
  const start = startAt + marker.length;
  const end = html.indexOf(';\nconst $', start);
  if (end < 0) throw new Error('Could not locate the end of the embedded DATA object.');
  return JSON.parse(html.slice(start, end));
}
function isSigned(c) {
  const label = String(c?.label ?? '').toLowerCase();
  const raw = String(c?.raw ?? '').toLowerCase();
  return label.includes('signed') || raw.includes('signed');
}
function isBooked(c) {
  const label = String(c?.label ?? '').toLowerCase();
  const raw = String(c?.raw ?? '').toLowerCase();
  return label.includes('booked') || raw.includes('booked');
}
function componentsFor(r) {
  if (Array.isArray(r.components)) return r.components;
  if (Number.isFinite(Number(r.signed)) && Number.isFinite(Number(r.total)) && Number(r.total) > 0) {
    return [{ label: 'Signed', units: Number(r.signed), total: Number(r.total), raw: r.raw ?? '' }];
  }
  return [];
}
function groupLabel(data, project, component, index, components) {
  // Match the revised source workbook's unit-group headers. Ignore the legacy Group 1 column.
  // Historical 189-unit observations and later 308-unit observations both belong to Group B.
  if (project === 'Springvale Sempurna Heights') {
    const groupText = [component.groupLabel, component.unitGroup, component.group, component.layoutLabel, component.raw, component.label].filter(Boolean).join(' ').toLowerCase();
    if (/group\s*a|front[- ]kitchen/.test(groupText) || Number(component.total) === 27) return 'Group A · front-kitchen layout';
    if (/group\s*b|rear[- ]kitchen/.test(groupText) || Number(component.total) === 189 || Number(component.total) === 308) return 'Group B · rear-kitchen layout';
  }

  const ruleGroups = data.scopeRules?.[project]?.groups ?? [];
  const total = Number(component.total);
  const stableId = component.unitGroupId ?? component.groupId ?? component.apiGroupId ?? component.teduhGroupId ?? component.group_id ?? component.unitTypeId ?? component.id;
  if (stableId !== undefined && stableId !== null && String(stableId).trim()) {
    const byId = ruleGroups.filter(g => [g.apiGroupId, g.unitGroupId, g.groupId, g.teduhGroupId, g.id]
      .some(id => id !== undefined && id !== null && String(id) === String(stableId)));
    if (byId.length === 1) return byId[0].label;
  }

  const explicit = component.groupLabel || component.unitGroup || component.group || component.layoutLabel;
  if (explicit && !/^(signed|booked|yet to sign|yet to launch|all configured comparison units)$/i.test(String(explicit).trim())) {
    return String(explicit).trim();
  }

  // A unique denominator can safely map to a configured group. Never use array order
  // to decide between multiple groups with the same denominator.
  const sameTotal = ruleGroups.filter(g => Number(g.total) === total);
  if (sameTotal.length === 1) return sameTotal[0].label;

  const landLines = String(data.static?.[project]?.['land area'] ?? '').split(/\n+/).map(x => x.trim()).filter(Boolean);
  if (ruleGroups.length === 0 && components.length === 1) {
    return data.scopeRules?.[project]?.label || landLines[0] || 'All comparison units';
  }
  if (components.length === 1) return data.scopeRules?.[project]?.label || 'All comparison units';

  const idSuffix = stableId !== undefined && stableId !== null && String(stableId).trim()
    ? ` [TEDUH ID: ${String(stableId).trim()}]`
    : '';
  console.warn(`UNMAPPED UNIT GROUP: ${project}; component ${index + 1} has no unique stable mapping${idSuffix}. Add a verified TEDUH ID/scopeRules mapping before interpreting group-level trends.`);
  return `Unmapped unit group ${index + 1} — verify TEDUH mapping${idSuffix}`;
}
function normalizeRecord(data, r, source) {
  const project = String(r.project || r.development || r.name || '').trim();
  let date = String(r.date || r.recordDate || '').slice(0, 10);
  if (!project || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;
  // User-verified correction: these two M Legasi snapshots belong to 27 Feb 2026,
  // as confirmed by the latest original dated source workbook.
  if ((project === 'M Legasi Phase 1' || project === 'M Legasi Phase 2') && date === '2025-02-27') date = '2026-02-27';
  const components = componentsFor(r).map((c, index, arr) => ({
    ...c,
    _groupLabel: groupLabel(data, project, c, index, arr),
  }));
  return { ...r, project, date, raw: r.raw ?? '', components, source };
}
function mergeRecords(data, sharedRecords) {
  const byKey = new Map();
  for (const r of data.history ?? []) {
    const normalized = normalizeRecord(data, r, 'Embedded historical record');
    if (normalized) byKey.set(`${normalized.project}|${normalized.date}`, normalized);
  }
  for (const r of sharedRecords ?? []) {
    const normalized = normalizeRecord(data, r, 'Saved TEDUH record');
    if (!normalized) continue;
    const key = `${normalized.project}|${normalized.date}`;
    const existing = byKey.get(key);
    if (!existing || normalized.components.length) byKey.set(key, normalized);
  }
  return [...byKey.values()].sort((a, b) => a.project.localeCompare(b.project) || a.date.localeCompare(b.date));
}
function signedMap(record) {
  const result = new Map();
  if (!record) return result;
  for (const c of record.components ?? []) {
    if (!isSigned(c)) continue;
    result.set(c._groupLabel, c);
  }
  return result;
}
function percent(c) {
  // Recalculate from the recorded counts; source-reported percentages have known historical errors.
  if (!c) return null;
  const total = Number(c.total);
  const units = Number(c.units);
  return Number.isFinite(units) && Number.isFinite(total) && total > 0 ? units / total * 100 : null;
}
function formatDate(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}
function weeklyBaseline(signedRecords, latest) {
  if (!latest) return null;
  const latestTime = Date.parse(`${latest.date}T00:00:00Z`);
  const min = latestTime - 14 * 86400000;
  const max = latestTime - 7 * 86400000;
  return [...signedRecords].reverse().find(r => {
    const t = Date.parse(`${r.date}T00:00:00Z`);
    return t >= min && t <= max;
  }) ?? null;
}

const html = fs.readFileSync(HTML_PATH, 'utf8');
const data = loadEmbeddedData(html);
const shared = JSON.parse(fs.readFileSync(HISTORY_PATH, 'utf8'));
if (!Array.isArray(shared.records)) throw new Error('data/teduh-history.json does not contain a records array.');
const history = mergeRecords(data, shared.records);
if (!history.length) throw new Error('No valid dated records were found.');

const byProject = new Map();
for (const r of history) {
  if (!byProject.has(r.project)) byProject.set(r.project, []);
  byProject.get(r.project).push(r);
}
const projects = [];
for (const [project, records] of byProject) {
  records.sort((a, b) => a.date.localeCompare(b.date));
  const signedRecords = records.filter(r => signedMap(r).size > 0);
  const latest = signedRecords.at(-1) ?? records.at(-1);
  const previous = signedRecords.length ? weeklyBaseline(signedRecords, latest) : null;
  const latestMap = signedMap(latest);
  const previousMap = signedMap(previous);
  let labels = [...new Set([...latestMap.keys(), ...previousMap.keys()])];
  const configuredOrder = (data.scopeRules?.[project]?.groups ?? []).map(g => g.label);
  if (configuredOrder.some(label => labels.includes(label))) {
    labels.sort((a, b) => {
      const ai = configuredOrder.indexOf(a), bi = configuredOrder.indexOf(b);
      if (ai >= 0 && bi >= 0) return ai - bi;
      if (ai >= 0) return -1;
      if (bi >= 0) return 1;
      return 0;
    });
  }
  if (!labels.length) labels = ['Overall status'];
  projects.push({ project, records, latest, previous, latestMap, previousMap, labels });
}
projects.sort((a, b) => a.project.localeCompare(b.project));

const wb = new ExcelJS.Workbook();
wb.creator = 'Broga Competitor Sales Dashboard';
wb.created = new Date();
wb.modified = new Date();
const ws = wb.addWorksheet('Weekly Comparison', { views: [{ state: 'frozen', xSplit: 1, ySplit: 5, topLeftCell: 'B6' }] });
ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 5, topLeftCell: 'B6', zoomScale: 80 }];
ws.getColumn(1).width = 25;

const totalCols = 1 + projects.reduce((n, p) => n + p.labels.length, 0);
ws.mergeCells(1, 1, 1, totalCols);
ws.getCell(1, 1).value = 'BROGA COMPETITOR SALES ANALYSIS';
ws.getCell(1, 1).font = { name: 'Aptos Display', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
ws.getCell(1, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172554' } };
ws.getCell(1, 1).alignment = { vertical: 'middle' };
ws.getRow(1).height = 34;
ws.mergeCells(2, 1, 2, totalCols);
ws.getCell(2, 1).value = `Generated from embedded historical records + saved TEDUH records · ${formatDate(new Date().toISOString().slice(0,10))}`;
ws.getCell(2, 1).font = { size: 10, bold: true, color: { argb: 'FF475467' } };
ws.getCell(2, 1).alignment = { vertical: 'middle', wrapText: true };
ws.getRow(2).height = 24;
ws.getCell(4, 1).value = 'RECORD DETAIL';
ws.getCell(5, 1).value = 'Unit group / layout';

let col = 2;
for (const p of projects) {
  p.startCol = col;
  p.endCol = col + p.labels.length - 1;
  if (p.endCol > p.startCol) ws.mergeCells(4, p.startCol, 4, p.endCol);
  ws.getCell(4, p.startCol).value = p.project;
  p.labels.forEach((label, i) => {
    ws.getCell(5, col + i).value = label;
    ws.getColumn(col + i).width = 21;
  });
  col = p.endCol + 1;
}
const navy = 'FF172554', lightBlue = 'FFEAF2FF', lightGrey = 'FFF3F5F8', border = 'FFD7DDE7';
for (let c = 1; c <= totalCols; c++) {
  for (const r of [4, 5]) {
    const cell = ws.getCell(r, c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: r === 4 ? navy : lightBlue } };
    cell.font = { bold: true, color: { argb: r === 4 ? 'FFFFFFFF' : navy } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: border } };
  }
}
ws.getRow(4).height = 34; ws.getRow(5).height = 48;
const rowSpecs = [
  ['Latest record', 'dateLatest'],
  ['Sales Unit · latest', 'unitsLatest'],
  ['Sales Status · latest', 'pctLatest'],
  ['Weekly baseline', 'datePrevious'],
  ['Sales Unit · baseline', 'unitsPrevious'],
  ['Sales Status · baseline', 'pctPrevious'],
  ['Weekly Change', 'change'],
];
for (let i = 0; i < rowSpecs.length; i++) {
  const rowNo = 6 + i;
  const [label, kind] = rowSpecs[i];
  ws.getCell(rowNo, 1).value = label;
  ws.getCell(rowNo, 1).font = { bold: true, color: { argb: 'FF5D6878' } };
  ws.getCell(rowNo, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: lightGrey } };
  ws.getCell(rowNo, 1).alignment = { vertical: 'middle', wrapText: true };
  for (const p of projects) {
    p.labels.forEach((g, idx) => {
      const cell = ws.getCell(rowNo, p.startCol + idx);
      const cur = p.latestMap.get(g); const prev = p.previousMap.get(g);
      const latestRaw = String(p.latest.raw ?? '').toLowerCase();
      const prevRaw = String(p.previous?.raw ?? '').toLowerCase();
      if (kind === 'dateLatest') cell.value = formatDate(p.latest.date);
      else if (kind === 'datePrevious') cell.value = p.previous ? formatDate(p.previous.date) : 'No record 7–14 days earlier';
      else if (kind === 'unitsLatest') cell.value = cur ? `${cur.units} / ${cur.total}` : '—';
      else if (kind === 'pctLatest') cell.value = cur ? `${Math.round(percent(cur))}%` : (latestRaw.includes('yet to launch') ? 'Yet to launch' : latestRaw.includes('yet to sign') ? 'Yet to sign' : '—');
      else if (kind === 'unitsPrevious') cell.value = prev ? `${prev.units} / ${prev.total}` : '—';
      else if (kind === 'pctPrevious') cell.value = prev ? `${Math.round(percent(prev))}%` : (prevRaw.includes('yet to launch') ? 'Yet to launch' : prevRaw.includes('yet to sign') ? 'Yet to sign' : '—');
      else if (kind === 'change') {
        let value;
        if (!cur || !prev) value = (!cur && /yet to (sign|launch)/i.test(p.latest.raw ?? '')) ? 'No signed status yet' : 'Insufficient weekly history';
        else if (Number(cur.total) !== Number(prev.total)) value = 'Total changed — review';
        else {
          const delta = Number(cur.units) - Number(prev.units);
          value = delta === 0 ? 'No change' : `${delta > 0 ? '+' : ''}${delta} signed`;
        }
        cell.value = value;
        if (value === 'No change') { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F2F5' } }; cell.font = { bold: true, color: { argb: 'FF5D6878' } }; }
        else if (/review|insufficient/i.test(value)) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF1D6' } }; cell.font = { bold: true, color: { argb: 'FFB54708' } }; }
        else if (/signed/i.test(value)) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F7EF' } }; cell.font = { bold: true, color: { argb: 'FF067647' } }; }
      }
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { bottom: { style: 'thin', color: border } };
      if (rowNo >= 6 && rowNo <= 8) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FBFF' } };
    });
  }
  ws.getRow(rowNo).height = kind === 'change' || kind.startsWith('date') ? 38 : 32;
}
ws.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5, column: totalCols } };
ws.pageSetup = { orientation: 'landscape', paperSize: 8, fitToPage: true, fitToWidth: 2, fitToHeight: 0, horizontalDpi: 300, verticalDpi: 300 };
ws.pageMargins = { left: 0.25, right: 0.25, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 };
ws.printTitlesRow = '1:5';

const hs = wb.addWorksheet('Historical Records', { views: [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2' }] });
hs.views = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', zoomScale: 90 }];
const histHeaders = ['Record Date','Project','Unit Group / Layout','Status Type','Signed Units','Total Units','Signed %','Raw Recorded Status','Source'];
hs.addRow(histHeaders);
hs.getRow(1).height = 30;
hs.getRow(1).eachCell((cell) => {
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: navy } };
  cell.alignment = { vertical: 'middle', wrapText: true };
});
for (const r of history) {
  const comps = r.components ?? [];
  if (!comps.length) {
    hs.addRow([formatDate(r.date),r.project,'No group details','Unclassified','','','','No parseable group-level values: '+String(r.raw ?? ''),r.source]);
    continue;
  }
  comps.forEach((c, idx) => {
    const type = isSigned(c) ? 'Signed' : isBooked(c) ? 'Booked' : String(c.label ?? 'Unclassified');
    const total = Number(c.total) || '';
    const signedUnits = type === 'Signed' ? Number(c.units) : '';
    const pctValue = type === 'Signed' && total ? Number(c.units) / Number(total) : null;
    const row = hs.addRow([formatDate(r.date),r.project,c._groupLabel ?? groupLabel(data,r.project,c,idx,comps),type,signedUnits,total,pctValue,c.raw ?? r.raw ?? '',r.source]);
    row.getCell(7).numFmt = '0%';
  });
}
[15,32,32,17,14,14,12,46,24].forEach((width, i) => hs.getColumn(i + 1).width = width);
hs.autoFilter = { from: { row: 1, column: 1 }, to: { row: hs.rowCount, column: histHeaders.length } };
for (let r = 2; r <= hs.rowCount; r++) {
  hs.getRow(r).eachCell(cell => { cell.border = { bottom: { style: 'hair', color: 'FFE5E7EB' } }; cell.alignment = { vertical: 'top', wrapText: [3,8].includes(cell.col) }; });
}

const readme = wb.addWorksheet('Read Me');
readme.addRows([
  ['Purpose','Horizontal competitor comparison by development and unit group; historical observations are preserved in Historical Records.'],
  ['Weekly baseline','The comparison uses a signed record dated 7–14 days before the latest signed record. If no such record exists, the workbook shows Insufficient weekly history rather than comparing against a much older date.'],
  ['Sales Unit','Signed units / total units, shown separately for every unit group. Groups are not combined into one total.'],
  ['Sales Status','Signed units divided by that group’s confirmed total, rounded to a whole percentage for display. Booking-only observations are not treated as zero signed sales.'],
  ['Weekly Change','The signed-unit change versus the weekly baseline. If the denominator changed, the workbook flags review instead of presenting a misleading sales change.'],
  ['Scope','Birchpark remains four terrace groups and excludes rumah bandar. Sapphire uses the terrace-only comparison scope. Springvale keeps the 27-unit front-kitchen and 308-unit rear-kitchen groups separate.'],
  ['OneDrive destination','Use a separate file named Broga_Competitor_Sales_Analysis.xlsx. Do not overwrite the original source workbook.'],
]);
readme.getColumn(1).width = 26; readme.getColumn(2).width = 115;
for (let r = 1; r <= readme.rowCount; r++) {
  readme.getCell(r,1).font = { bold: true, color: { argb: navy } };
  readme.getCell(r,1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: lightBlue } };
  readme.getCell(r,1).alignment = { vertical: 'top', wrapText: true };
  readme.getCell(r,2).alignment = { vertical: 'top', wrapText: true };
  readme.getRow(r).height = 46;
}

await wb.xlsx.writeFile(OUT_PATH);
console.log(`Created ${OUT_PATH} from ${history.length} merged dated project records.`);
