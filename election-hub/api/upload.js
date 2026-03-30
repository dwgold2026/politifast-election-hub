/**
 * API: Upload candidate data manually via CSV
 * POST /api/upload
 * Body: multipart/form-data with fields:
 *   - file: CSV file
 *   - state: two-letter state code (e.g. "CA")
 *   - source: optional label (defaults to "manual_upload")
 *
 * CSV must have a header row. Recognized columns (case-insensitive, flexible matching):
 *   name, first_name, last_name, office, party, email, phone, website,
 *   address, city, state, zip, county, level/office_level
 *
 * Deduplication: matches on name + office. Existing records are enriched
 * (new fields added) but never overwritten with empty values.
 */
export const config = { maxDuration: 30 };
import { getCached, setCache } from './lib/cache.js';

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

// Map flexible column names to our standard fields
function mapHeader(raw) {
  const h = raw.toLowerCase().replace(/[^a-z0-9]/g, '');
  const mappings = {
    name: 'name', fullname: 'name', candidatename: 'name', candidate: 'name', nameonballot: 'name',
    firstname: 'firstName', first: 'firstName', fname: 'firstName',
    lastname: 'lastName', last: 'lastName', lname: 'lastName',
    office: 'office', contest: 'office', contestname: 'office', position: 'office', race: 'office', officesought: 'office',
    party: 'party', partyaffiliation: 'party', partycandidate: 'party',
    email: 'email', emailaddress: 'email', campaignemail: 'email',
    phone: 'phone', phonenumber: 'phone', telephone: 'phone', campaignphone: 'phone', officephone: 'phone', businessphone: 'phone',
    website: 'website', url: 'website', web: 'website', campaignwebsite: 'website', webaddress: 'website',
    address: 'address', streetaddress: 'address', mailingaddress: 'address', street: 'address', candstreet1: 'address',
    city: 'city', candcity: 'city',
    state: 'state', st: 'state', candstate: 'state',
    zip: 'zip', zipcode: 'zip', postalcode: 'zip', candzip: 'zip',
    county: 'county', countyname: 'county',
    level: 'officeLevel', officelevel: 'officeLevel',
    district: 'district',
    status: 'status', incumbentchallenger: 'status',
    filingdate: 'filingDate', candidacydt: 'filingDate',
  };
  return mappings[h] || null;
}

function classifyOffice(title) {
  if (!title) return 'other';
  const u = title.toUpperCase();
  if (u.includes('U.S.') || u.includes('UNITED STATES') || u.includes('CONGRESS') || u.includes('SENATE') && !u.includes('STATE')) return 'federal';
  if (u.includes('GOVERNOR') || u.includes('STATE') || u.includes('ASSEMBLY') || u.includes('LEGISLATURE') ||
      u.includes('ATTORNEY GENERAL') || u.includes('SECRETARY') || u.includes('TREASURER') ||
      u.includes('COMPTROLLER') || u.includes('AUDITOR')) return 'state';
  if (u.includes('COUNTY') || u.includes('SHERIFF') || u.includes('CLERK') || u.includes('JUDGE') ||
      u.includes('COMMISSIONER') || u.includes('CORONER')) return 'county';
  if (u.includes('SCHOOL') || u.includes('BOARD OF EDUCATION') || u.includes('ISD')) return 'school';
  if (u.includes('CITY') || u.includes('MAYOR') || u.includes('COUNCIL') || u.includes('ALDERMAN') ||
      u.includes('MUNICIPAL') || u.includes('TOWN')) return 'municipal';
  return 'other';
}

function parseCSVData(text) {
  const lines = text.split('\n');
  if (lines.length < 2) return [];

  const rawHeaders = parseCSVLine(lines[0]);
  const headerMap = rawHeaders.map(h => mapHeader(h.trim()));

  const candidates = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const vals = parseCSVLine(lines[i]);
    const row = {};
    headerMap.forEach((field, idx) => {
      if (field && vals[idx]) row[field] = vals[idx].trim();
    });

    // Build candidate — require at least a name or first+last
    let name = row.name || '';
    if (!name && row.firstName && row.lastName) {
      name = `${titleCase(row.firstName)} ${titleCase(row.lastName)}`;
    } else if (!name && row.firstName) {
      name = titleCase(row.firstName);
    } else if (!name && row.lastName) {
      name = titleCase(row.lastName);
    }
    if (!name) continue;

    candidates.push({
      name: name.includes(',') ? titleCase(name.split(',').reverse().join(' ').trim()) : titleCase(name),
      firstName: titleCase(row.firstName || ''),
      lastName: titleCase(row.lastName || ''),
      office: row.office || '',
      officeLevel: row.officeLevel || classifyOffice(row.office || ''),
      party: (row.party || '').toUpperCase(),
      email: (row.email || '').toLowerCase(),
      phone: row.phone || '',
      website: row.website || '',
      address: row.address || '',
      city: row.city || '',
      state: row.state || '',
      zip: row.zip || '',
      county: row.county || '',
      district: row.district || '',
      status: row.status || '',
      filingDate: row.filingDate || '',
      source: 'manual_upload',
    });
  }
  return candidates;
}

// Merge new candidates into existing, deduplicating by name+office
function mergeCandidate(existing, incoming) {
  const merged = { ...existing };
  for (const [key, val] of Object.entries(incoming)) {
    if (!val) continue; // Don't overwrite with empty
    if (key === 'source') continue; // Keep original source unless empty
    if (!merged[key]) merged[key] = val; // Only fill gaps
  }
  // Always mark that manual data was merged
  if (!merged.source) merged.source = incoming.source;
  return merged;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'POST only' });
  }

  try {
    // Read raw body
    const chunks = [];
    for await (const chunk of req) chunks.push(chunk);
    const body = Buffer.concat(chunks).toString('utf8');

    // Parse multipart or JSON
    let state, source, csvText;

    if (req.headers['content-type']?.includes('application/json')) {
      const json = JSON.parse(body);
      state = json.state;
      source = json.source || 'manual_upload';
      csvText = json.csv;
    } else {
      // Simple multipart parser — extract fields
      const boundary = req.headers['content-type']?.split('boundary=')[1];
      if (!boundary) return res.status(400).json({ error: 'Missing boundary in multipart form' });

      const parts = body.split(`--${boundary}`);
      for (const part of parts) {
        if (part.includes('name="state"')) {
          state = part.split('\r\n\r\n')[1]?.trim().replace(/\r\n--$/, '');
        } else if (part.includes('name="source"')) {
          source = part.split('\r\n\r\n')[1]?.trim().replace(/\r\n--$/, '');
        } else if (part.includes('name="file"') || part.includes('name="csv"')) {
          csvText = part.split('\r\n\r\n').slice(1).join('\r\n\r\n').replace(/\r\n--$/, '').trim();
        }
      }
      source = source || 'manual_upload';
    }

    if (!state || !csvText) {
      return res.status(400).json({ error: 'Missing required fields: state (2-letter code) and file/csv data' });
    }

    state = state.toUpperCase();
    if (state.length !== 2) {
      return res.status(400).json({ error: 'State must be a 2-letter code (e.g. CA, NC)' });
    }

    // Parse uploaded CSV
    const uploaded = parseCSVData(csvText);
    if (uploaded.length === 0) {
      return res.status(400).json({ error: 'No valid candidates found in CSV. Ensure it has a header row with recognizable column names (name, office, party, email, phone, etc.)' });
    }

    // Set state on all records
    for (const c of uploaded) c.state = state;

    // Get existing cached data for this state
    const existing = await getCached(`manual_${state.toLowerCase()}`, state);
    const existingData = existing && Array.isArray(existing.data) ? existing.data : [];

    // Build dedup map: key = lowercase(name|office)
    const byKey = new Map();
    for (const c of existingData) {
      byKey.set(`${c.name?.toLowerCase()}|${c.office?.toLowerCase()}`, c);
    }

    // Merge uploaded candidates
    let newCount = 0;
    let enrichedCount = 0;
    for (const c of uploaded) {
      const key = `${c.name?.toLowerCase()}|${c.office?.toLowerCase()}`;
      if (byKey.has(key)) {
        // Existing candidate — enrich with new fields only
        const merged = mergeCandidate(byKey.get(key), c);
        byKey.set(key, merged);
        enrichedCount++;
      } else {
        // New candidate — add
        byKey.set(key, c);
        newCount++;
      }
    }

    const merged = [...byKey.values()];
    merged.sort((a, b) => (a.office || '').localeCompare(b.office || '') || (a.lastName || '').localeCompare(b.lastName || ''));

    // Save as new snapshot
    const payload = {
      source: `manual_${state.toLowerCase()}`,
      totalCandidates: merged.length,
      withEmail: merged.filter(c => c.email).length,
      withPhone: merged.filter(c => c.phone).length,
      data: merged,
    };
    await setCache(`manual_${state.toLowerCase()}`, state, payload);

    res.json({
      success: true,
      state,
      uploaded: uploaded.length,
      newCandidates: newCount,
      enrichedExisting: enrichedCount,
      totalAfterMerge: merged.length,
      withEmail: payload.withEmail,
      withPhone: payload.withPhone,
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
