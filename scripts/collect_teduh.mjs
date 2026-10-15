// Daily TEDUH collector for the Broga Competitor Sales Intelligence Dashboard.
// Runs on GitHub Actions; it does not write to the user's local computer.
import fs from 'node:fs/promises';

const OUT = 'data/teduh-history.json';
const ENDPOINTS = [
  { project: 'Balau Homes 2A', id: '20094-4', total: 113 },
  { project: 'Balau Homes 2C', id: '20094-5', total: 48 },
  { project: 'Springvale Sempurna Heights', id: '19321-4', total: 335, type: 'springvale' },
  { project: 'Zinnia Residences', id: '19959-3', total: 149 },
  { project: 'M Legasi Phase 1', id: '30897-1', total: 330, type: 'mlegasi1' },
  { project: 'M Legasi Phase 2', id: '30897-2', total: 256, type: 'mlegasi2' },
  { project: 'M Legasi Phase 3', id: '30897-3', total: 231, type: 'mlegasi3' },
  { project: 'Sapphire Residence @Diamond City', id: '13170-7', total: 136, type: 'terrace' },
  { project: 'Birchpark @Eco Forest', id: '19150-17', total: 152, type: 'birchpark' },
  { project: 'Chorus @Setia Ecohill 2', id: '14421-21', total: 99 },
  { project: 'Banyan @Setia Ecohill 2', id: '14421-22', total: 113 },
  { project: 'Amani Gallery Homes', id: '6868-20', total: 81 }
];
const BASE = 'https://teduh.kpkt.gov.my/api/unit-projek-swasta/';
const MALAYSIA_DATE = new Intl.DateTimeFormat('en-CA', {
  timeZone: 'Asia/Kuala_Lumpur', year: 'numeric', month: '2-digit', day: '2-digit'
}).format(new Date());

function groupsFrom(data) {
  const groups = data?.unitGroups ?? data?.data?.unitGroups;
  if (!Array.isArray(groups)) throw new Error('API response has no unitGroups array');
  return groups;
}
function flatten(groups) {
  const out = [];
  for (const g of groups) {
    if (!g || typeof g !== 'object') continue;
    const list = g.units ?? g.unit ?? [];
    if (!Array.isArray(list)) continue;
    for (const u of list) if (u && typeof u === 'object') out.push({
      ...u,
      __groupLabel: String(g.jenis ?? g.namaKumpulan ?? g.namaJenisRumah ?? g.jenisRumah ?? g.namaJenis ?? g.name ?? ''),
      __groupType: String(g.jenis ?? g.jenisRumah ?? g.type ?? g.kategori ?? ''),
      __groupId: g.pembangunan_id ?? g.id ?? null
    });
  }
  if (!out.length) throw new Error('API returned no unit records');
  return out;
}
function label(u) {
  for (const k of ['no','noUnit','noUnitRumah','unitNo','noRumah','noLot','noUnitPemajuan','unitNumber','unit_number','unitNoRumah','nomborUnit','nomborRumah','noPintu','noBangunan','kodUnit','unit']) {
    if (u?.[k] !== undefined && u?.[k] !== null && String(u[k]).trim()) return String(u[k]).trim();
  }
  return '';
}
function sold(u) {
  const a = String(u.status ?? '').trim().toLowerCase();
  const b = String(u.statusJualan ?? '').trim().toLowerCase();
  return a === 'sold' || b === 'telah dijual' || b === 'sold';
}
function inRange(value, lo, hi, prefix) {
  const s = String(value || '').trim();
  let m;
  if (prefix === 'Z') m = s.match(/\bZ\s*-\s*0*(\d+)\b/i);
  else m = s.match(/\bL\s*0*(\d+)\b/i) || s.match(/^(?:LOT|UNIT|NO\.?\s*)?\s*0*(\d+)\s*$/i);
  if (!m) return false;
  const n = Number(m[1]);
  return n >= lo && n <= hi;
}
function exactCount(name, units, expected, groupName) {
  if (units.length !== expected) throw new Error(`${name}: ${groupName} matched ${units.length} units; expected ${expected}. Refusing to save a possibly incorrect snapshot.`);
}
function makeComponents(endpoint, units) {
  const project = endpoint.project;
  if (endpoint.type === 'springvale') {
    const a = units.filter(u => inRange(label(u), 1, 27, 'L'));
    const b = units.filter(u => inRange(label(u), 28, 335, 'L'));
    exactCount(project, a, 27, 'Group A L1-L27');
    exactCount(project, b, 308, 'Group B L28-L335');
    return [[a,27,'Group A · front-kitchen layout'],[b,308,'Group B · rear-kitchen layout']].map(([list,total,group]) => component(list,total,group));
  }
  if (endpoint.type === 'mlegasi1') {
    const g1 = units.filter(u => inRange(label(u),1,102,'Z') || inRange(label(u),552,560,'Z'));
    const g2 = units.filter(u => [[103,162],[189,197],[223,243],[367,390],[417,440]].some(([lo,hi]) => inRange(label(u),lo,hi,'Z')));
    const g3 = units.filter(u => Number(u.__groupId ?? u.pembangunan_id ?? u.id) === 57721);
    exactCount(project,g1,111,'Group 1 Z ranges');
    exactCount(project,g2,138,'Group 2 Z ranges');
    exactCount(project,g3,81,'Group 3 API group 57721');
    return [component(g1,111,'Group 1'),component(g2,138,'Group 2'),component(g3,81,'Group 3')];
  }
  if (endpoint.type === 'mlegasi2' || endpoint.type === 'mlegasi3') {
    const sizes = endpoint.type === 'mlegasi2' ? [26,126,104] : [44,85,102];
    if (units.length !== sizes.reduce((a,b)=>a+b,0)) throw new Error(`${project}: returned ${units.length} units; expected ${sizes.reduce((a,b)=>a+b,0)}`);
    let offset=0;
    return sizes.map((size,i) => {
      const groupUnits=units.slice(offset,offset+size); offset+=size;
      return component(groupUnits,size,`Group ${i+1}`);
    });
  }
  if (endpoint.type === 'terrace' || endpoint.type === 'birchpark') {
    const byGroup = new Map();
    for (const u of units) {
      const type = `${u.__groupLabel} ${u.__groupType} ${u.groupName ?? ''}`.toLowerCase();
      const isTerrace = /rumah\s*teres|terrace/.test(type) && !/rumah\s*bandar|bandar|sesebuah|detached/.test(type);
      if (isTerrace) {
        const key = endpoint.type === 'birchpark' ? String(u.__groupId ?? u.pembangunan_id ?? '') : 'terrace';
        if (!byGroup.has(key)) byGroup.set(key,[]);
        byGroup.get(key).push(u);
      }
    }
    const selected=[...byGroup.values()].flat();
    exactCount(project,selected,endpoint.total,'terrace-only scope');
    if (endpoint.type === 'birchpark') {
      const expected = new Map([['57513',46],['57514',56],['57515',25],['57516',25]]);
      const components=[];
      for (const [id,total] of expected) {
        const list=byGroup.get(id)||[];
        exactCount(project,list,total,`TEDUH group ${id}`);
        components.push(component(list,total,`Group ${id}`));
      }
      return components;
    }
    return [component(selected,136,'Rumah teres only')];
  }
  exactCount(project, units, endpoint.total, 'configured comparison scope');
  return [component(units, endpoint.total, 'All configured comparison units')];
}
function component(units,total,label) {
  const count=units.filter(sold).length;
  return { label:'Signed', units:count, total, reportedPct:Math.round(count/total*100), raw:`${count}/${total} signed`, group:label };
}
async function collect(endpoint) {
  const response=await fetch(BASE+endpoint.id,{headers:{Accept:'application/json'},signal:AbortSignal.timeout(25000)});
  if(!response.ok) throw new Error(`${endpoint.project} (${endpoint.id}): HTTP ${response.status}`);
  const data=await response.json();
  const units=flatten(groupsFrom(data));
  const components=makeComponents(endpoint,units);
  const signed=components.reduce((n,c)=>n+c.units,0);
  const total=components.reduce((n,c)=>n+c.total,0);
  if(total!==endpoint.total) throw new Error(`${endpoint.project}: scoped total ${total} differs from expected ${endpoint.total}`);
  return {
    project:endpoint.project,date:MALAYSIA_DATE,
    raw:components.map(c=>`${c.reportedPct}% Signed (${c.units}/${c.total} units)`).join('\n'),
    components,signed,booked:0,total,signedPct:total?signed/total*100:null,
    mappedId:endpoint.id,static:{},source:'TEDUH API · scheduled snapshot'
  };
}

const file = await fs.readFile(OUT,'utf8').catch(err => err.code==='ENOENT' ? '{"schemaVersion":1,"lastRun":null,"records":[]}' : Promise.reject(err));
const store=JSON.parse(file);
if (!Array.isArray(store.records)) store.records=[];
const collected=[];
for (const endpoint of ENDPOINTS) {
  const snapshot=await collect(endpoint);
  collected.push(snapshot);
  console.log(`${snapshot.project}: ${snapshot.signed}/${snapshot.total} signed (${Math.round(snapshot.signedPct)}%)`);
}
// Save only after every configured endpoint passed validation; no partial daily snapshots.
const byKey=new Map(store.records.map(r=>[`${r.project}|${r.date}`,r]));
for (const snapshot of collected) byKey.set(`${snapshot.project}|${snapshot.date}`,snapshot);
store.schemaVersion=1;
store.lastRun=new Date().toISOString();
store.records=[...byKey.values()].sort((a,b)=>a.date.localeCompare(b.date)||a.project.localeCompare(b.project));
await fs.writeFile(OUT,JSON.stringify(store,null,2)+'\n');
console.log(`Saved ${collected.length} observations for ${MALAYSIA_DATE}; ${store.records.length} dated records total.`);
