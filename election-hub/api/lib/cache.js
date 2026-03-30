/**
 * Redis cache layer for candidate data — append-only versioned snapshots.
 *
 * Every pull is stored as a dated snapshot. Data is never overwritten or deleted.
 * The API always serves the latest snapshot.
 *
 * Key structure:
 *   candidates:{source}:{state}:{YYYY-MM-DD}  — dated snapshot
 *   candidates:{source}:{state}:latest         — pointer to most recent date
 *   candidates:{source}:{state}:history        — list of all pull dates
 */
import { Redis } from '@upstash/redis';

const redis = new Redis({
  url: process.env.KV_REST_API_URL,
  token: process.env.KV_REST_API_TOKEN,
});

function today() {
  return new Date().toISOString().split('T')[0]; // YYYY-MM-DD
}

function baseKey(source, state) {
  return state ? `candidates:${source}:${state}` : `candidates:${source}`;
}

/**
 * Get the latest snapshot for a source+state
 */
export async function getCached(source, state) {
  try {
    const base = baseKey(source, state);
    // Get the latest date pointer
    const latestDate = await redis.get(`${base}:latest`);
    if (!latestDate) return null;
    // Get that snapshot
    const cached = await redis.get(`${base}:${latestDate}`);
    if (cached && typeof cached === 'object') return { ...cached, snapshotDate: latestDate };
    if (cached && typeof cached === 'string') return { ...JSON.parse(cached), snapshotDate: latestDate };
    return null;
  } catch {
    return null;
  }
}

/**
 * Store a new dated snapshot — never overwrites previous snapshots.
 * Returns { saved: boolean, reason: string, date: string }
 */
export async function setCache(source, state, data) {
  try {
    const base = baseKey(source, state);
    const date = today();
    const freshCount = Array.isArray(data.data) ? data.data.length : 0;

    if (freshCount === 0) {
      return { saved: false, reason: 'Empty data — not saved', date };
    }

    const payload = {
      pulledAt: new Date().toISOString(),
      candidateCount: freshCount,
      ...data,
    };

    // Save the dated snapshot
    await redis.set(`${base}:${date}`, JSON.stringify(payload));

    // Update the latest pointer
    await redis.set(`${base}:latest`, date);

    // Append to history list
    const history = await redis.get(`${base}:history`);
    const dates = history ? (typeof history === 'string' ? JSON.parse(history) : history) : [];
    if (!dates.includes(date)) {
      dates.push(date);
      await redis.set(`${base}:history`, JSON.stringify(dates));
    }

    return { saved: true, reason: `Saved snapshot ${date} with ${freshCount} candidates`, date };
  } catch {
    return { saved: false, reason: 'Redis write failed', date: today() };
  }
}

/**
 * Get a specific dated snapshot
 */
export async function getSnapshot(source, state, date) {
  try {
    const key = `${baseKey(source, state)}:${date}`;
    const data = await redis.get(key);
    if (data && typeof data === 'object') return data;
    if (data && typeof data === 'string') return JSON.parse(data);
    return null;
  } catch {
    return null;
  }
}

/**
 * Get list of all snapshot dates for a source+state
 */
export async function getHistory(source, state) {
  try {
    const key = `${baseKey(source, state)}:history`;
    const history = await redis.get(key);
    if (!history) return [];
    return typeof history === 'string' ? JSON.parse(history) : history;
  } catch {
    return [];
  }
}

/**
 * Wrapper: fetch fresh data, save as new snapshot, fall back to latest snapshot on failure.
 */
export async function fetchWithCache(source, state, fetchFn) {
  try {
    const freshData = await fetchFn();
    const freshCount = freshData && Array.isArray(freshData.data) ? freshData.data.length : 0;

    if (freshCount > 0) {
      const cacheResult = await setCache(source, state, freshData);
      return { ...freshData, fromCache: false, cacheStatus: cacheResult.reason };
    }

    // Fresh fetch returned empty — serve latest snapshot
    const cached = await getCached(source, state);
    if (cached) return { ...cached, fromCache: true, note: 'Source returned empty. Showing latest cached snapshot.' };
    return freshData;
  } catch (err) {
    // Fetch failed — serve latest snapshot
    const cached = await getCached(source, state);
    if (cached) return { ...cached, fromCache: true, note: `Source unavailable (${err.message}). Showing latest cached snapshot.` };
    throw err;
  }
}
