/**
 * Vercel Serverless Function: Fetch and parse FEC bulk candidate data
 * GET /api/collect/fec?state=NC  (optional state filter)
 */
export const config = { maxDuration: 30 };

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
  return (s || '').toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

function normalizeName(raw) {
  if (!raw) return { first: '', last: '', full: '' };
  const parts = raw.split(',').map(s => s.trim());
  const last = titleCase(parts[0] || '');
  const firstRaw = (parts[1] || '').replace(/\b(MR|MRS|MS|DR|JR|SR|II|III|IV)\.?\b/gi, '').trim();
  const first = titleCase(firstRaw.split(/\s+/)[0] || '');
  return { first, last, full: `${first} ${last}`.trim() };
}

export default async function handler(req, res) {
  const stateFilter = (req.query.state || '').toUpperCase();

  try {
    const resp = await fetch('https://www.fec.gov/files/bulk-downloads/2026/candidate_summary_2026.csv', { redirect: 'follow' });
    if (!resp.ok) return res.status(502).json({ error: `FEC returned ${resp.status}` });
    const text = await resp.text();

    const lines = text.split('\n');
    const headers = parseCSVLine(lines[0]).map(h => h.trim());

    const byState = {};
    for (let i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      const vals = parseCSVLine(lines[i]);
      const row = {};
      headers.forEach((h, idx) => row[h] = (vals[idx] || '').trim());

      const st = row.Cand_Office_St;
      if (!st || st.length !== 2) continue;
      if (stateFilter && st !== stateFilter) continue;

      const name = normalizeName(row.Cand_Name);
      const officeCode = row.Cand_Office;
      let office = officeCode;
      if (officeCode === 'P') office = 'President';
      else if (officeCode === 'S') office = `U.S. Senate (${st})`;
      else if (officeCode === 'H') office = `U.S. House ${st}-${row.Cand_Office_Dist || 'AL'}`;

      const candidate = {
        name: name.full,
        firstName: name.first,
        lastName: name.last,
        office,
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

    // Sort each state's candidates
    for (const st of Object.keys(byState)) {
      byState[st].sort((a, b) => a.office.localeCompare(b.office) || a.lastName.localeCompare(b.lastName));
    }

    const totalCandidates = Object.values(byState).reduce((sum, arr) => sum + arr.length, 0);

    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json({
      source: 'fec',
      updated: new Date().toISOString(),
      totalCandidates,
      states: Object.keys(byState).length,
      data: stateFilter ? (byState[stateFilter] || []) : byState,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
