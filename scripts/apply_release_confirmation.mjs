// Applies a "Confirm release" GitHub issue to data/scope-config.json.
// Run by .github/workflows/confirm-release.yml with the issue body in ISSUE_BODY.
// Prints a short result for the issue comment; exits 1 (changing nothing) if the request is invalid.
import fs from 'node:fs/promises';

const CONFIG = 'data/scope-config.json';
const body = process.env.ISSUE_BODY ?? '';
const issue = process.env.ISSUE_NUMBER ?? '';

function fail(message) {
  console.log(`Not applied: ${message}`);
  process.exit(1);
}
const field = name => body.match(new RegExp(`^\\s*${name}\\s*:\\s*(.+?)\\s*$`, 'mi'))?.[1] ?? null;

const project = field('project');
const range = field('range');
const units = Number(field('units'));
const group = field('group');
if (!project || !range || !group) fail('the issue must contain "project:", "range:" and "group:" lines.');

const config = JSON.parse(await fs.readFile(CONFIG, 'utf8'));
const scope = config.projects?.[project];
if (!scope) fail(`"${project}" has no range-based scope in ${CONFIG}.`);

const prefix = scope.unitPrefix ?? 'L';
const m = range.match(new RegExp(`^${prefix}\\s*(\\d+)\\s*[-–]\\s*${prefix}\\s*(\\d+)$`, 'i'));
if (!m) fail(`range "${range}" is not in the form ${prefix}336-${prefix}373.`);
const lo = Number(m[1]), hi = Number(m[2]);
if (hi < lo) fail(`range "${range}" ends before it starts.`);

const max = Math.max(...scope.groups.flatMap(g => g.ranges.map(r => r[1])));
const overlaps = scope.groups.some(g => g.ranges.some(([a, b]) => lo <= b && hi >= a));
if (overlaps) fail(`${range} overlaps units already in the comparison scope.`);
if (lo !== max + 1) fail(`${range} must start right after the current boundary ${prefix}${max}.`);

const target = scope.groups.find(g => g.label === group);
if (!target) fail(`group "${group}" is not one of: ${scope.groups.map(g => g.label).join(', ')}.`);

const before = target.ranges.reduce((n, [a, b]) => n + b - a + 1, 0);
const last = target.ranges.at(-1);
if (last && last[1] + 1 === lo) last[1] = hi; else target.ranges.push([lo, hi]);
const after = before + hi - lo + 1;
const size = hi - lo + 1;

scope.confirmedReleases = scope.confirmedReleases ?? [];
scope.confirmedReleases.push({
  range: `${prefix}${lo}–${prefix}${hi}`, units: size, group,
  confirmedOn: new Date().toISOString().slice(0, 10), issue: issue ? Number(issue) : null
});
await fs.writeFile(CONFIG, JSON.stringify(config, null, 2) + '\n');

const total = scope.groups.reduce((n, g) => n + g.ranges.reduce((k, [a, b]) => k + b - a + 1, 0), 0);
console.log(`Applied: ${prefix}${lo}–${prefix}${hi} (${size} units) added to ${project} ${group}: ${before} → ${after} units; project comparison total now ${total}.`);
if (Number.isFinite(units) && units !== size) console.log(`Note: the dashboard saw ${units} units in this range, but the range covers ${size} unit numbers. The daily collector will flag the difference until TEDUH lists every unit.`);
console.log('The dashboard uses the new scope on its next refresh; the daily collection and Excel reports use it from their next run, and the first new observation is flagged as a scope change.');
