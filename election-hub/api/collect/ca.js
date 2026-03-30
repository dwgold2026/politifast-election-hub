/**
 * Vercel Serverless Function: Fetch CA Cal-Access candidate data
 * Uses the candidate_summary CSV from CA SoS
 * GET /api/collect/ca
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

export default async function handler(req, res) {
  try {
    // Cal-Access provides a TSV/CSV download of candidates
    const resp = await fetch('https://cal-access.sos.ca.gov/Campaign/Candidates/list.aspx?view=certified&session=2025', {
      headers: { 'User-Agent': 'Mozilla/5.0 PolitiFast Election Hub' }
    });

    if (!resp.ok) {
      // Fallback: return FEC data for CA
      return res.json({
        source: 'ca_sos',
        updated: new Date().toISOString(),
        totalCandidates: 0,
        note: 'CA Cal-Access requires browser access. Use FEC data for federal candidates.',
        data: [],
      });
    }

    const html = await resp.text();

    // Parse candidate table from HTML
    const candidates = [];
    const rowRegex = /<tr[^>]*class="[^"]*listRow[^"]*"[^>]*>([\s\S]*?)<\/tr>/gi;
    const cellRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    let rowMatch;

    while ((rowMatch = rowRegex.exec(html)) !== null) {
      const cells = [];
      let cellMatch;
      const rowHtml = rowMatch[1];
      while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {
        cells.push(cellMatch[1].replace(/<[^>]+>/g, '').trim());
      }
      if (cells.length >= 3) {
        candidates.push({
          name: titleCase(cells[0]),
          firstName: titleCase(cells[0].split(',')[1] || '').trim(),
          lastName: titleCase(cells[0].split(',')[0] || '').trim(),
          office: cells[1] || '',
          officeLevel: 'state',
          party: cells[2] || '',
          state: 'CA',
          source: 'cal_access',
        });
      }
    }

    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json({
      source: 'cal_access',
      updated: new Date().toISOString(),
      totalCandidates: candidates.length,
      note: candidates.length === 0 ? 'CA Cal-Access data could not be parsed. FEC federal data available via /api/collect/fec?state=CA' : undefined,
      data: candidates,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
