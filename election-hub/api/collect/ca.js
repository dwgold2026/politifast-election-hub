/**
 * Vercel Serverless Function: Fetch CA candidate data
 * Cal-Access is behind WAF and JS-rendered — not scrapable server-side.
 * Uses OpenFEC API for federal candidates + broader search.
 * GET /api/collect/ca
 */
export const config = { maxDuration: 30 };
import { getCached, setCache } from '../lib/cache.js';

function titleCase(s) {
  return (s || '').toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

const FEC_API_KEY = 'DEMO_KEY';

async function fetchFECCandidates(state, office) {
  const candidates = [];
  let page = 1;
  let hasMore = true;
  while (hasMore && page <= 5) {
    const url = `https://api.open.fec.gov/v1/candidates/?api_key=${FEC_API_KEY}&election_year=2026&state=${state}&office=${office}&per_page=100&page=${page}&sort=name`;
    const resp = await fetch(url);
    if (!resp.ok) break;
    const data = await resp.json();
    if (!data.results || data.results.length === 0) break;
    for (const c of data.results) {
      const nameParts = (c.name || '').split(',').map(s => s.trim());
      const last = titleCase(nameParts[0] || '');
      const first = titleCase((nameParts[1] || '').split(/\s+/)[0] || '');
      candidates.push({
        name: `${first} ${last}`.trim(),
        firstName: first,
        lastName: last,
        office: c.office_full ? `${c.office_full}${c.district && c.district !== '00' ? ` District ${c.district}` : ''}` : '',
        officeLevel: 'federal',
        party: c.party || '',
        fecId: c.candidate_id || '',
        status: c.incumbent_challenge_full || '',
        state: 'CA',
        source: 'openfec',
      });
    }
    hasMore = data.pagination && page < data.pagination.pages;
    page++;
  }
  return candidates;
}

export default async function handler(req, res) {
  try {
    // Fetch Senate and House candidates from OpenFEC
    const [senate, house] = await Promise.all([
      fetchFECCandidates('CA', 'S'),
      fetchFECCandidates('CA', 'H'),
    ]);

    const candidates = [...senate, ...house];
    candidates.sort((a, b) => a.office.localeCompare(b.office) || a.lastName.localeCompare(b.lastName));

    const payload = {
      source: 'openfec',
      updated: new Date().toISOString(),
      totalCandidates: candidates.length,
      note: 'CA Cal-Access is behind WAF. Using OpenFEC API for federal candidates. State-level candidates require Cal-Access browser access.',
      data: candidates,
    };

    if (candidates.length > 0) await setCache('ca_sos', 'CA', payload).catch(() => {});
    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json(payload);
  } catch (err) {
    const cached = await getCached('ca_sos', 'CA').catch(() => null);
    if (cached) return res.json({ ...cached, fromCache: true, note: 'Source unavailable. Showing cached data.' });
    res.status(500).json({ error: err.message });
  }
}
