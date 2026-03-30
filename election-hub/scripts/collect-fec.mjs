/**
 * FEC Bulk Data Collector
 * Downloads the FEC candidate summary CSV for 2026 and outputs per-state JSON files.
 * Source: https://www.fec.gov/files/bulk-downloads/2026/candidate_summary_2026.csv
 */
import { writeFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = join(__dirname, '..', 'public', 'data');
const FEC_URL = 'https://www.fec.gov/files/bulk-downloads/2026/candidate_summary_2026.csv';

// Simple CSV parser that handles quoted fields
function parseCSV(text) {
  const lines = text.split('\n');
  const headers = parseCSVLine(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const vals = parseCSVLine(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => obj[h.trim()] = (vals[idx] || '').trim());
    rows.push(obj);
  }
  return rows;
}

function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}

function normalizeOffice(code, state, district) {
  if (code === 'P') return 'President';
  if (code === 'S') return `U.S. Senate (${state})`;
  if (code === 'H') return `U.S. House ${state}-${district || 'AL'}`;
  return code;
}

function normalizeName(raw) {
  if (!raw) return { first: '', last: '', full: raw || '' };
  // FEC format: "LAST, FIRST MIDDLE SUFFIX" or "LAST, FIRST"
  const parts = raw.split(',').map(s => s.trim());
  const last = titleCase(parts[0] || '');
  const firstParts = (parts[1] || '').replace(/\b(MR|MRS|MS|DR|JR|SR|II|III|IV)\.?\b/gi, '').trim();
  const first = titleCase(firstParts.split(/\s+/)[0] || '');
  return { first, last, full: `${first} ${last}`.trim() };
}

function titleCase(s) {
  return s.toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

async function main() {
  console.log('Fetching FEC candidate summary...');
  const resp = await fetch(FEC_URL, { redirect: 'follow' });
  if (!resp.ok) throw new Error(`FEC download failed: ${resp.status}`);
  const text = await resp.text();

  console.log('Parsing CSV...');
  const rows = parseCSV(text);
  console.log(`Parsed ${rows.length} candidates`);

  // Group by state
  const byState = {};
  for (const row of rows) {
    const st = row.Cand_Office_St;
    if (!st || st.length !== 2) continue;

    const name = normalizeName(row.Cand_Name);
    const candidate = {
      name: name.full,
      firstName: name.first,
      lastName: name.last,
      office: normalizeOffice(row.Cand_Office, st, row.Cand_Office_Dist),
      officeLevel: 'federal',
      party: row.Cand_Party_Affiliation || '',
      status: row.Cand_Incumbent_Challenger_Open_Seat || '',
      fecId: row.Cand_Id || '',
      address: [row.Cand_Street_1, row.Cand_Street_2].filter(Boolean).join(', '),
      city: row.Cand_City || '',
      state: row.Cand_State || st,
      zip: row.Cand_Zip || '',
      totalReceipts: parseFloat(row.Total_Receipt) || 0,
      cashOnHand: parseFloat(row.Cash_On_Hand_COP) || 0,
      source: 'fec',
    };

    if (!byState[st]) byState[st] = [];
    byState[st].push(candidate);
  }

  // Write per-state files
  mkdirSync(OUT_DIR, { recursive: true });
  let totalWritten = 0;
  for (const [st, candidates] of Object.entries(byState)) {
    candidates.sort((a, b) => a.office.localeCompare(b.office) || a.lastName.localeCompare(b.lastName));
    const outPath = join(OUT_DIR, `${st.toLowerCase()}-federal.json`);
    writeFileSync(outPath, JSON.stringify(candidates, null, 2));
    totalWritten += candidates.length;
    console.log(`  ${st}: ${candidates.length} candidates`);
  }

  console.log(`\nDone! Wrote ${totalWritten} candidates across ${Object.keys(byState).length} states to ${OUT_DIR}`);
}

main().catch(err => { console.error(err); process.exit(1); });
