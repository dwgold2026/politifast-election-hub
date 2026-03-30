/**
 * Vercel Serverless Function: Fetch OR ORESTAR candidate/committee data
 * GET /api/collect/or
 */
export const config = { maxDuration: 30 };

function titleCase(s) {
  return (s || '').toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

export default async function handler(req, res) {
  try {
    // ORESTAR provides a search API for committees
    const apiUrl = 'https://secure.sos.state.or.us/orestar/GNR_ActiveCandComm.do?commandName=search&electionYear=2026&format=json';
    const resp = await fetch(apiUrl, {
      headers: { 'User-Agent': 'Mozilla/5.0 PolitiFast Election Hub' },
    });

    if (!resp.ok) {
      return res.json({
        source: 'orestar',
        updated: new Date().toISOString(),
        totalCandidates: 0,
        note: 'OR ORESTAR API not accessible. Use FEC data via /api/collect/fec?state=OR',
        data: [],
      });
    }

    let data;
    const text = await resp.text();
    try { data = JSON.parse(text); } catch { data = null; }

    const candidates = [];
    if (data && Array.isArray(data.items || data.results || data)) {
      const items = data.items || data.results || data;
      for (const item of items) {
        candidates.push({
          name: titleCase(item.candidateName || item.name || item.committeeName || ''),
          firstName: titleCase(item.firstName || ''),
          lastName: titleCase(item.lastName || ''),
          office: item.officeSought || item.office || '',
          officeLevel: 'state',
          party: item.party || item.partyAffiliation || '',
          email: item.email || '',
          phone: item.phone || '',
          website: item.webAddress || item.website || '',
          state: 'OR',
          source: 'orestar',
        });
      }
    }

    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json({
      source: 'orestar',
      updated: new Date().toISOString(),
      totalCandidates: candidates.length,
      note: candidates.length === 0 ? 'ORESTAR data format may have changed. Use FEC data for federal candidates.' : undefined,
      data: candidates,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
