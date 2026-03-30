/**
 * Cron job: runs all collectors weekly to keep cached data fresh.
 * Triggered by Vercel Cron: every Monday at 6am UTC
 *
 * Each collector is called internally and results are saved to Redis.
 * GET /api/cron/collect-all
 */
export const config = { maxDuration: 60 };

const COLLECTORS = [
  { name: 'FEC (all states)', url: '/api/collect/fec' },
  { name: 'NC (all levels)', url: '/api/collect/nc' },
  { name: 'CA (certified list)', url: '/api/collect/ca' },
  { name: 'GA (qualified candidates)', url: '/api/collect/ga' },
  { name: 'MD (SBE candidates)', url: '/api/collect/md' },
  { name: 'MN (SoS filings)', url: '/api/collect/mn' },
  { name: 'OR (ORESTAR)', url: '/api/collect/or' },
];

export default async function handler(req, res) {
  // Verify cron secret to prevent unauthorized triggers
  if (req.headers.authorization !== `Bearer ${process.env.CRON_SECRET}`) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const baseUrl = `https://${req.headers.host}`;
  const results = [];

  for (const collector of COLLECTORS) {
    const start = Date.now();
    try {
      const resp = await fetch(`${baseUrl}${collector.url}`, {
        headers: { 'User-Agent': 'Election-Hub-Cron/1.0' },
      });
      const data = await resp.json();
      results.push({
        name: collector.name,
        status: resp.ok ? 'success' : 'error',
        candidates: data.totalCandidates || 0,
        fromCache: data.fromCache || false,
        cacheStatus: data.cacheStatus || '',
        duration: `${Date.now() - start}ms`,
      });
    } catch (err) {
      results.push({
        name: collector.name,
        status: 'failed',
        error: err.message,
        duration: `${Date.now() - start}ms`,
      });
    }
  }

  const totalCandidates = results.reduce((sum, r) => sum + (r.candidates || 0), 0);
  const successful = results.filter(r => r.status === 'success').length;

  res.json({
    ran: new Date().toISOString(),
    summary: `${successful}/${COLLECTORS.length} collectors succeeded, ${totalCandidates} total candidates`,
    results,
  });
}
