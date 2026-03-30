/**
 * Simple password authentication
 * POST /api/auth { password: "..." }
 * Set APP_PASSWORD env var in Vercel
 */
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST only' });

  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const { password } = JSON.parse(Buffer.concat(chunks).toString('utf8'));

  if (!process.env.APP_PASSWORD) {
    return res.status(500).json({ error: 'APP_PASSWORD not configured' });
  }

  if (password === process.env.APP_PASSWORD) {
    return res.json({ success: true });
  }

  res.status(401).json({ error: 'Invalid password' });
}
