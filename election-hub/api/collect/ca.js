/**
 * Vercel Serverless Function: Fetch CA certified candidate list + contact list
 * Parses the official SoS PDF certified list and contact list
 * GET /api/collect/ca
 */
export const config = { maxDuration: 30 };
import pdf from 'pdf-parse';
import { fetchWithCache } from '../lib/cache.js';

function titleCase(s) {
  return (s || '').toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

function classifyOffice(title) {
  const u = title.toUpperCase();
  if (u.includes('REPRESENTATIVE') || u.includes('UNITED STATES') || u.includes('U.S.')) return 'federal';
  if (u.includes('GOVERNOR') || u.includes('LIEUTENANT') || u.includes('SECRETARY OF STATE') ||
      u.includes('CONTROLLER') || u.includes('TREASURER') || u.includes('ATTORNEY GENERAL') ||
      u.includes('INSURANCE') || u.includes('SUPERINTENDENT') || u.includes('EQUALIZATION') ||
      u.includes('STATE SENATE') || u.includes('STATE ASSEMBLY') || u.includes('SENATE DISTRICT') ||
      u.includes('ASSEMBLY DISTRICT')) return 'state';
  return 'other';
}

export default async function handler(req, res) {
  try {
    // Fetch the certified list PDF
    const pdfUrl = 'https://elections.cdn.sos.ca.gov/statewide-elections/2026-primary/cert-list-candidates.pdf';
    const resp = await fetch(pdfUrl, { redirect: 'follow' });

    if (!resp.ok) {
      return res.json({
        source: 'ca_sos',
        updated: new Date().toISOString(),
        totalCandidates: 0,
        note: 'CA certified list PDF not available. Use FEC data for federal candidates via /api/collect/fec?state=CA',
        data: [],
      });
    }

    const buffer = Buffer.from(await resp.arrayBuffer());
    const pdfData = await pdf(buffer);
    const text = pdfData.text;

    // Parse candidates from PDF text
    // Format is typically: office header, then candidate lines with party and name
    const candidates = [];
    let currentOffice = '';
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

    // Patterns for office headers (all caps, specific keywords)
    const officePattern = /^(GOVERNOR|LIEUTENANT GOVERNOR|SECRETARY OF STATE|CONTROLLER|TREASURER|ATTORNEY GENERAL|INSURANCE COMMISSIONER|SUPERINTENDENT|MEMBER.*EQUALIZATION|UNITED STATES REPRESENTATIVE|STATE SENATOR|MEMBER.*ASSEMBLY|U\.S\. REPRESENTATIVE)/i;
    const partyAbbrevs = ['DEM', 'REP', 'AI', 'GRN', 'LIB', 'PF', 'NPP', 'NP'];

    for (const line of lines) {
      // Check if this is an office header
      const officeMatch = line.match(officePattern);
      if (officeMatch) {
        currentOffice = line;
        continue;
      }

      if (!currentOffice) continue;

      // Try to extract candidate: look for party abbreviation followed by name
      // Or name followed by party
      let party = '';
      let name = line;

      for (const p of partyAbbrevs) {
        const partyRegex = new RegExp(`\\b${p}\\b`, 'i');
        if (partyRegex.test(line)) {
          party = p.toUpperCase();
          name = line.replace(partyRegex, '').trim();
          break;
        }
      }

      // Clean up name - remove district numbers, asterisks, etc
      name = name.replace(/[*#†]/g, '').replace(/\s+/g, ' ').trim();

      // Skip if it looks like a header or metadata line
      if (!name || name.length < 3 || name.length > 60) continue;
      if (/^\d+$/.test(name) || /^(Page|CERTIFIED|JUNE|PRIMARY|ELECTION|OFFICE|DISTRICT|COUNTY)/i.test(name)) continue;

      if (name && currentOffice) {
        const nameParts = name.split(/\s+/);
        candidates.push({
          name: titleCase(name),
          firstName: titleCase(nameParts[0] || ''),
          lastName: titleCase(nameParts[nameParts.length - 1] || ''),
          office: currentOffice,
          officeLevel: classifyOffice(currentOffice),
          party: party || '',
          state: 'CA',
          source: 'ca_sos_certified',
        });
      }
    }

    // Also try to fetch and parse contact list for email/phone
    let contactMap = {};
    try {
      const contactResp = await fetch('https://elections.cdn.sos.ca.gov/statewide-elections/2026-primary/contact-list.pdf', { redirect: 'follow' });
      if (contactResp.ok) {
        const contactBuf = Buffer.from(await contactResp.arrayBuffer());
        const contactPdf = await pdf(contactBuf);
        const contactText = contactPdf.text;

        // Try to extract email addresses and phone numbers associated with candidate names
        const emailRegex = /[\w.-]+@[\w.-]+\.\w+/g;
        const phoneRegex = /\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/g;
        const contactLines = contactText.split('\n');

        let currentName = '';
        for (const cl of contactLines) {
          const trimmed = cl.trim();
          if (!trimmed) continue;

          const emails = trimmed.match(emailRegex) || [];
          const phones = trimmed.match(phoneRegex) || [];

          // If line has a name-like pattern (capitalized words, no numbers at start)
          if (/^[A-Z][a-zA-Z]/.test(trimmed) && !emails.length && !phones.length) {
            currentName = trimmed.toLowerCase();
          }

          if (currentName && (emails.length || phones.length)) {
            contactMap[currentName] = {
              email: emails[0] || '',
              phone: phones[0] || '',
            };
          }
        }

        // Enrich candidates with contact info
        for (const c of candidates) {
          const key = c.name.toLowerCase();
          if (contactMap[key]) {
            c.email = contactMap[key].email;
            c.phone = contactMap[key].phone;
          }
        }
      }
    } catch (e) {
      // Contact enrichment is optional
    }

    const payload = {
      source: 'ca_sos_certified',
      updated: new Date().toISOString(),
      totalCandidates: candidates.length,
      withEmail: candidates.filter(c => c.email).length,
      withPhone: candidates.filter(c => c.phone).length,
      data: candidates,
    };
    // Cache successful results
    if (candidates.length > 0) {
      const { setCache } = await import('../lib/cache.js');
      await setCache('ca_sos', 'CA', payload).catch(() => {});
    }
    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json(payload);
  } catch (err) {
    // Try cache on failure
    const { getCached } = await import('../lib/cache.js');
    const cached = await getCached('ca_sos', 'CA').catch(() => null);
    if (cached) return res.json({ ...cached, fromCache: true, note: `Source unavailable. Showing cached data.` });
    res.status(500).json({ error: err.message });
  }
}
