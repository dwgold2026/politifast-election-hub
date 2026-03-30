/**
 * Vercel Serverless Function: Fetch and parse NC NCSBE candidate data
 * GET /api/collect/nc
 */
export const config = { maxDuration: 30 };
import { fetchWithCache } from '../lib/cache.js';

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

function titleCase(s) {
  if (!s) return '';
  return s.toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

function classifyOffice(c) {
  const u = c.toUpperCase();
  if (u.includes('US SENATE') || u.includes('US HOUSE') || u.includes('PRESIDENT')) return 'federal';
  if (u.includes('GOVERNOR') || u.includes('LT. GOVERNOR') || u.includes('ATTORNEY GENERAL') ||
      u.includes('STATE SENATE') || u.includes('STATE HOUSE') || u.includes('NC SENATE') ||
      u.includes('NC HOUSE') || u.includes('SECRETARY OF STATE') || u.includes('STATE AUDITOR') ||
      u.includes('STATE TREASURER') || u.includes('SUPT OF PUBLIC') || u.includes('COMMISSIONER OF')) return 'state';
  if (u.includes('BOARD OF EDUCATION') || u.includes('SCHOOL')) return 'school';
  if (u.includes('COUNTY') || u.includes('SHERIFF') || u.includes('REGISTER OF DEEDS') ||
      u.includes('CLERK OF') || u.includes('DISTRICT COURT') || u.includes('SUPERIOR COURT')) return 'county';
  if (u.includes('CITY') || u.includes('TOWN') || u.includes('MUNICIPAL') || u.includes('MAYOR') ||
      u.includes('ALDERMAN')) return 'municipal';
  if (u.includes('SOIL') || u.includes('WATER') || u.includes('SANITARY') || u.includes('FIRE')) return 'special';
  return 'other';
}

export default async function handler(req, res) {
  try {
    const result = await fetchWithCache('ncsbe', 'NC', async () => {
    const resp = await fetch('https://s3.amazonaws.com/dl.ncsbe.gov/Elections/2026/Candidate%20Filing/Candidate_Listing_2026.csv', { redirect: 'follow' });
    if (!resp.ok) throw new Error(`NC SBE returned ${resp.status}`);
    const text = await resp.text();

    const lines = text.split('\n');
    const headers = parseCSVLine(lines[0]).map(h => h.trim());

    const seen = new Set();
    const candidates = [];
    for (let i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      const vals = parseCSVLine(lines[i]);
      const row = {};
      headers.forEach((h, idx) => row[h] = (vals[idx] || '').trim());

      const key = `${row.first_name}|${row.last_name}|${row.contest_name}`;
      if (seen.has(key)) continue;
      seen.add(key);

      candidates.push({
        name: row.name_on_ballot || `${titleCase(row.first_name)} ${titleCase(row.last_name)}`,
        firstName: titleCase(row.first_name),
        lastName: titleCase(row.last_name),
        office: row.contest_name || '',
        officeLevel: classifyOffice(row.contest_name || ''),
        party: row.party_candidate || '',
        county: row.county_name || '',
        email: (row.email || '').toLowerCase(),
        phone: row.phone || row.office_phone || row.business_phone || '',
        address: row.street_address || '',
        city: row.city || '',
        state: 'NC',
        zip: row.zip_code || '',
        filingDate: row.candidacy_dt || '',
        electionDate: row.election_dt || '',
        source: 'ncsbe',
      });
    }

    const levelOrder = { federal: 0, state: 1, county: 2, municipal: 3, school: 4, special: 5, other: 6 };
    candidates.sort((a, b) =>
      (levelOrder[a.officeLevel] ?? 9) - (levelOrder[b.officeLevel] ?? 9) ||
      a.office.localeCompare(b.office) ||
      a.lastName.localeCompare(b.lastName)
    );

    return {
      source: 'ncsbe',
      totalCandidates: candidates.length,
      withEmail: candidates.filter(c => c.email).length,
      withPhone: candidates.filter(c => c.phone).length,
      data: candidates,
    };
    });

    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json(result);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
