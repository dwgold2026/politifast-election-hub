/**
 * API: Browse data pull history and past snapshots
 * GET /api/history?source=fec&state=NC         — list all pull dates
 * GET /api/history?source=fec&state=NC&date=2026-03-29  — get specific snapshot
 */
import { getHistory, getSnapshot, getCached } from './lib/cache.js';

export default async function handler(req, res) {
  const { source, state, date } = req.query;

  if (!source) {
    return res.status(400).json({ error: 'Missing ?source= parameter (fec, ncsbe, mn_sos, ca_sos, ga_sos, md_sbe, orestar)' });
  }

  // If date specified, return that specific snapshot
  if (date) {
    const snapshot = await getSnapshot(source, state, date);
    if (!snapshot) return res.status(404).json({ error: `No snapshot found for ${source}/${state || 'all'}/${date}` });
    return res.json({ source, state, date, ...snapshot });
  }

  // Otherwise return history + latest summary
  const history = await getHistory(source, state);
  const latest = await getCached(source, state);

  res.json({
    source,
    state: state || 'all',
    totalSnapshots: history.length,
    dates: history,
    latest: latest ? {
      date: latest.snapshotDate,
      candidateCount: latest.candidateCount || (Array.isArray(latest.data) ? latest.data.length : 0),
      pulledAt: latest.pulledAt || latest.updated,
    } : null,
  });
}
