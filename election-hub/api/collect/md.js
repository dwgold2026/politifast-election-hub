/**
 * Vercel Serverless Function: Fetch MD State Board of Elections candidate data
 * MD provides CSV downloads of candidate lists
 * GET /api/collect/md
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

function classifyOffice(title) {
  const u = title.toUpperCase();
  if (u.includes('U.S.') || u.includes('PRESIDENT') || u.includes('CONGRESS')) return 'federal';
  if (u.includes('GOVERNOR') || u.includes('COMPTROLLER') || u.includes('ATTORNEY GENERAL') ||
      u.includes('STATE SENATE') || u.includes('DELEGATE') || u.includes('HOUSE OF')) return 'state';
  if (u.includes('COUNTY') || u.includes('JUDGE') || u.includes('SHERIFF') ||
      u.includes("STATE'S ATTORNEY")) return 'county';
  if (u.includes('BOARD OF EDUCATION') || u.includes('SCHOOL')) return 'school';
  if (u.includes('CITY') || u.includes('MAYOR') || u.includes('COUNCIL')) return 'municipal';
  return 'other';
}

export default async function handler(req, res) {
  try {
    // MD publishes candidate lists as CSV
    const csvUrl = 'https://elections.maryland.gov/elections/2026/primary_candidates/2026_GP_all_counties_candidatelist.html';
    const resp = await fetch(csvUrl, {
      headers: { 'User-Agent': 'Mozilla/5.0 PolitiFast Election Hub' },
    });

    if (!resp.ok) {
      return res.json({
        source: 'md_sbe',
        updated: new Date().toISOString(),
        totalCandidates: 0,
        note: 'MD SBE candidate list not accessible via API. Use FEC data via /api/collect/fec?state=MD',
        data: [],
      });
    }

    const html = await resp.text();
    const candidates = [];

    // Parse HTML table rows
    const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    const cellRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    let rowMatch;

    while ((rowMatch = rowRegex.exec(html)) !== null) {
      const cells = [];
      let cellMatch;
      while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
        cells.push(cellMatch[1].replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim());
      }
      if (cells.length >= 3 && cells[0] && !cells[0].toUpperCase().includes('OFFICE')) {
        candidates.push({
          name: titleCase(cells[1] || cells[0]),
          firstName: '',
          lastName: '',
          office: cells[0] || '',
          officeLevel: classifyOffice(cells[0] || ''),
          party: cells[2] || '',
          state: 'MD',
          source: 'md_sbe',
        });
      }
    }

    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json({
      source: 'md_sbe',
      updated: new Date().toISOString(),
      totalCandidates: candidates.length,
      data: candidates,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
