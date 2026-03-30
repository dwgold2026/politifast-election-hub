/**
 * Vercel Serverless Function: Fetch and parse MN SoS candidate data
 * MN provides semicolon-delimited text files with campaign email, phone, website
 * Filing opens Jun 2, 2026 — returns empty until files are available
 * GET /api/collect/mn
 */
export const config = { maxDuration: 30 };

function titleCase(s) {
  if (!s) return '';
  return s.toLowerCase().replace(/(?:^|\s|-)\S/g, c => c.toUpperCase());
}

function classifyOffice(title) {
  const u = title.toUpperCase();
  if (u.includes('U.S.') || u.includes('UNITED STATES')) return 'federal';
  if (u.includes('STATE') || u.includes('GOVERNOR') || u.includes('ATTORNEY GENERAL') ||
      u.includes('SECRETARY OF STATE') || u.includes('AUDITOR')) return 'state';
  if (u.includes('COUNTY') || u.includes('SHERIFF') || u.includes('JUDGE')) return 'county';
  if (u.includes('CITY') || u.includes('MAYOR') || u.includes('COUNCIL')) return 'municipal';
  if (u.includes('SCHOOL') || u.includes('ISD')) return 'school';
  if (u.includes('TOWNSHIP') || u.includes('HOSPITAL')) return 'special';
  return 'other';
}

export default async function handler(req, res) {
  const urls = [
    'https://candidates.sos.mn.gov/FedStateCountyCandidates.txt',
    'https://candidates.sos.mn.gov/LocalCandidates.txt',
  ];

  try {
    const results = await Promise.all(urls.map(u =>
      fetch(u).then(r => r.ok ? r.text() : '').catch(() => '')
    ));

    const candidates = [];
    for (const text of results) {
      if (!text.trim()) continue;
      const lines = text.split('\n');
      for (const line of lines) {
        if (!line.trim()) continue;
        const fields = line.split(';').map(f => f.trim());
        // Layout: OfficeID;CandName;OfficeID2;OfficeTitle;CountyID;Party;
        //         ResAddr;ResCity;ResState;ResZip;CampAddr;CampCity;CampState;CampZip;
        //         CampPhone;CampWebsite;CampEmail;RunMateWebsite;RunMateEmail;RunMatePhone
        if (fields.length < 17) continue;

        const name = titleCase(fields[1] || '');
        const nameParts = name.split(/\s+/);
        candidates.push({
          name,
          firstName: nameParts[0] || '',
          lastName: nameParts[nameParts.length - 1] || '',
          office: fields[3] || '',
          officeLevel: classifyOffice(fields[3] || ''),
          party: fields[5] || '',
          email: (fields[16] || '').toLowerCase(),
          phone: fields[14] || '',
          website: fields[15] || '',
          address: fields[10] || fields[6] || '',
          city: fields[11] || fields[7] || '',
          state: 'MN',
          zip: fields[13] || fields[9] || '',
          source: 'mn_sos',
        });
      }
    }

    const levelOrder = { federal: 0, state: 1, county: 2, municipal: 3, school: 4, special: 5, other: 6 };
    candidates.sort((a, b) =>
      (levelOrder[a.officeLevel] ?? 9) - (levelOrder[b.officeLevel] ?? 9) ||
      a.office.localeCompare(b.office) ||
      a.lastName.localeCompare(b.lastName)
    );

    res.setHeader('Cache-Control', 's-maxage=3600, stale-while-revalidate=86400');
    res.json({
      source: 'mn_sos',
      updated: new Date().toISOString(),
      totalCandidates: candidates.length,
      withEmail: candidates.filter(c => c.email).length,
      withPhone: candidates.filter(c => c.phone).length,
      withWebsite: candidates.filter(c => c.website).length,
      note: candidates.length === 0 ? 'MN filing opens Jun 2, 2026 — data not yet available' : undefined,
      data: candidates,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
