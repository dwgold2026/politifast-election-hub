/**
 * NC NCSBE Candidate Data Collector
 * Downloads the NC State Board of Elections candidate CSV (all levels, includes contact info).
 * Source: https://s3.amazonaws.com/dl.ncsbe.gov/Elections/2026/Candidate%20Filing/Candidate_Listing_2026.csv
 */
import { writeFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = join(__dirname, '..', 'public', 'data');
const NC_URL = 'https://s3.amazonaws.com/dl.ncsbe.gov/Elections/2026/Candidate%20Filing/Candidate_Listing_2026.csv';

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

function parseCSV(text) {
  const lines = text.split('\n');
  const headers = parseCSVLine(lines[0]).map(h => h.trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const vals = parseCSVLine(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => obj[h] = (vals[idx] || '').trim());
    rows.push(obj);
  }
  return rows;
}

function titleCase(s) {
  if (!s) return '';
  return s.toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

function classifyOffice(contestName) {
  const c = contestName.toUpperCase();
  if (c.includes('US SENATE') || c.includes('US HOUSE') || c.includes('PRESIDENT')) return 'federal';
  if (c.includes('GOVERNOR') || c.includes('LT. GOVERNOR') || c.includes('ATTORNEY GENERAL') ||
      c.includes('STATE SENATE') || c.includes('STATE HOUSE') || c.includes('NC SENATE') ||
      c.includes('NC HOUSE') || c.includes('SECRETARY OF STATE') || c.includes('STATE AUDITOR') ||
      c.includes('STATE TREASURER') || c.includes('SUPT OF PUBLIC') || c.includes('COMMISSIONER OF')) return 'state';
  if (c.includes('BOARD OF EDUCATION') || c.includes('SCHOOL')) return 'school';
  if (c.includes('COUNTY') || c.includes('SHERIFF') || c.includes('REGISTER OF DEEDS') ||
      c.includes('CLERK OF') || c.includes('DISTRICT COURT') || c.includes('SUPERIOR COURT')) return 'county';
  if (c.includes('CITY') || c.includes('TOWN') || c.includes('MUNICIPAL') || c.includes('MAYOR') ||
      c.includes('ALDERMAN')) return 'municipal';
  if (c.includes('SOIL') || c.includes('WATER') || c.includes('SANITARY') || c.includes('FIRE')) return 'special';
  return 'other';
}

async function main() {
  console.log('Fetching NC NCSBE candidate data...');
  const resp = await fetch(NC_URL, { redirect: 'follow' });
  if (!resp.ok) throw new Error(`NC download failed: ${resp.status}`);
  const text = await resp.text();

  console.log('Parsing CSV...');
  const rows = parseCSV(text);
  console.log(`Parsed ${rows.length} candidate rows`);

  // Deduplicate by name+contest (candidates appear once per county for statewide races)
  const seen = new Set();
  const candidates = [];
  for (const row of rows) {
    const key = `${row.first_name}|${row.last_name}|${row.contest_name}`;
    if (seen.has(key)) continue;
    seen.add(key);

    const phone = row.phone || row.office_phone || row.business_phone || '';
    candidates.push({
      name: row.name_on_ballot || `${titleCase(row.first_name)} ${titleCase(row.last_name)}`,
      firstName: titleCase(row.first_name),
      lastName: titleCase(row.last_name),
      office: row.contest_name || '',
      officeLevel: classifyOffice(row.contest_name || ''),
      party: row.party_candidate || '',
      county: row.county_name || '',
      email: (row.email || '').toLowerCase(),
      phone: phone,
      address: row.street_address || '',
      city: row.city || '',
      state: 'NC',
      zip: row.zip_code || '',
      filingDate: row.candidacy_dt || '',
      electionDate: row.election_dt || '',
      source: 'ncsbe',
    });
  }

  // Sort by office level then name
  const levelOrder = { federal: 0, state: 1, county: 2, municipal: 3, school: 4, special: 5, other: 6 };
  candidates.sort((a, b) =>
    (levelOrder[a.officeLevel] ?? 9) - (levelOrder[b.officeLevel] ?? 9) ||
    a.office.localeCompare(b.office) ||
    a.lastName.localeCompare(b.lastName)
  );

  // Stats
  const withEmail = candidates.filter(c => c.email).length;
  const withPhone = candidates.filter(c => c.phone).length;
  const byLevel = {};
  for (const c of candidates) byLevel[c.officeLevel] = (byLevel[c.officeLevel] || 0) + 1;

  mkdirSync(OUT_DIR, { recursive: true });
  const outPath = join(OUT_DIR, 'nc-all.json');
  writeFileSync(outPath, JSON.stringify(candidates, null, 2));

  console.log(`\nNC Candidate Summary:`);
  console.log(`  Total unique candidates: ${candidates.length}`);
  console.log(`  With email: ${withEmail} (${Math.round(withEmail/candidates.length*100)}%)`);
  console.log(`  With phone: ${withPhone} (${Math.round(withPhone/candidates.length*100)}%)`);
  console.log(`  By level:`, byLevel);
  console.log(`\nWrote to ${outPath}`);
}

main().catch(err => { console.error(err); process.exit(1); });
