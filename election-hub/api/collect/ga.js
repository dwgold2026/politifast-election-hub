/**
 * Vercel Serverless Function: Fetch GA qualified candidates
 * GA SoS provides downloadable candidate lists
 * GET /api/collect/ga
 */
export const config = { maxDuration: 30 };
import { getCached, setCache } from '../lib/cache.js';

function titleCase(s) {
  return (s || '').toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

function classifyOffice(title) {
  const u = title.toUpperCase();
  if (u.includes('U.S.') || u.includes('UNITED STATES') || u.includes('PRESIDENT')) return 'federal';
  if (u.includes('GOVERNOR') || u.includes('STATE') || u.includes('GENERAL ASSEMBLY') ||
      u.includes('SENATE DISTRICT') || u.includes('HOUSE DISTRICT') || u.includes('SECRETARY') ||
      u.includes('ATTORNEY GENERAL') || u.includes('LIEUTENANT') || u.includes('COMMISSIONER OF')) return 'state';
  if (u.includes('COUNTY') || u.includes('SHERIFF') || u.includes('PROBATE') ||
      u.includes('CLERK') || u.includes('TAX COMMISSIONER') || u.includes('CORONER') ||
      u.includes('SUPERIOR COURT') || u.includes('MAGISTRATE')) return 'county';
  if (u.includes('CITY') || u.includes('MAYOR') || u.includes('COUNCIL') || u.includes('MUNICIPAL')) return 'municipal';
  if (u.includes('SCHOOL') || u.includes('BOARD OF EDUCATION')) return 'school';
  return 'other';
}

export default async function handler(req, res) {
  try {
    // Try the GA qualified candidates CSV
    const csvUrl = 'https://sos.ga.gov/sites/default/files/2026-qualified-candidates.csv';
    const resp = await fetch(csvUrl, {
      headers: { 'User-Agent': 'Mozilla/5.0 PolitiFast Election Hub' },
      redirect: 'follow',
    });

    if (!resp.ok) {
      
      return res.json({
        source: 'ga_sos',
        updated: new Date().toISOString(),
        totalCandidates: 0,
        note: 'GA SoS candidate list not available via API. Use FEC data for federal candidates via /api/collect/fec?state=GA',
        data: [],
      });
    }

    const text = await resp.text();
    // Parse if we got data
    const lines = text.split('\n');
    const candidates = [];

    for (let i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      const fields = lines[i].split(',').map(f => f.replace(/^"|"$/g, '').trim());
      if (fields.length >= 3) {
        candidates.push({
          name: titleCase(fields[0]),
          firstName: '',
          lastName: '',
          office: fields[1] || '',
          officeLevel: classifyOffice(fields[1] || ''),
          party: fields[2] || '',
          state: 'GA',
          source: 'ga_sos',
        });
      }
    }

    const payload = { source: 'ga_sos', updated: new Date().toISOString(), totalCandidates: candidates.length, data: candidates };
    if (candidates.length > 0) await setCache('ga_sos', 'GA', payload).catch(() => {});
    res.json(payload);
  } catch (err) {
    const cached = await getCached('ga_sos', 'GA').catch(() => null);
    if (cached) return res.json({ ...cached, fromCache: true, note: 'Source unavailable. Showing cached data.' });
    res.status(500).json({ error: err.message });
  }
}
