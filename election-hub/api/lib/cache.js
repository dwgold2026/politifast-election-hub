/**
 * Redis cache layer for candidate data.
 * Stores collected data so we don't lose it if a source goes down.
 *
 * Safety rules:
 * - Fresh data only overwrites cache if it has MORE or EQUAL candidates
 * - If fresh data has significantly fewer candidates (>30% drop), cache is kept
 * - Empty fetches never overwrite existing cache
 * - All cache entries are indefinite (no TTL)
 *
 * Keys: "candidates:{source}:{state}" or "candidates:{source}"
 */
import { Redis } from '@upstash/redis';

const redis = new Redis({
  url: process.env.KV_REST_API_URL,
  token: process.env.KV_REST_API_TOKEN,
});

/**
 * Get cached candidate data for a source+state combo
 */
export async function getCached(source, state) {
  try {
    const key = state ? `candidates:${source}:${state}` : `candidates:${source}`;
    const cached = await redis.get(key);
    if (cached && typeof cached === 'object') return cached;
    if (cached && typeof cached === 'string') return JSON.parse(cached);
    return null;
  } catch {
    return null;
  }
}

/**
 * Store candidate data in cache — only if it passes safety checks
 * Returns { saved: boolean, reason: string }
 */
export async function setCache(source, state, data) {
  try {
    const key = state ? `candidates:${source}:${state}` : `candidates:${source}`;
    const freshCount = Array.isArray(data.data) ? data.data.length : 0;

    // Never cache empty data
    if (freshCount === 0) return { saved: false, reason: 'Fresh data is empty — cache preserved' };

    // Check existing cache
    const existing = await getCached(source, state);
    const existingCount = existing && Array.isArray(existing.data) ? existing.data.length : 0;

    // If existing cache has data, only overwrite if fresh data isn't a major regression
    if (existingCount > 0) {
      const dropPercent = ((existingCount - freshCount) / existingCount) * 100;

      if (dropPercent > 30) {
        // Fresh data lost more than 30% of candidates — suspicious, keep cache
        return {
          saved: false,
          reason: `Fresh data has ${freshCount} candidates vs ${existingCount} cached (${Math.round(dropPercent)}% drop). Cache preserved.`,
        };
      }
    }

    // Safe to save — fresh data is equal, larger, or within 30% of existing
    const payload = {
      updated: new Date().toISOString(),
      cachedAt: new Date().toISOString(),
      previousCount: existingCount || undefined,
      ...data,
    };
    await redis.set(key, JSON.stringify(payload));
    return { saved: true, reason: `Cached ${freshCount} candidates (was ${existingCount})` };
  } catch {
    return { saved: false, reason: 'Redis write failed' };
  }
}

/**
 * Wrapper: try to fetch fresh data, fall back to cache if fetch fails.
 * Fresh data only replaces cache if it passes safety checks.
 */
export async function fetchWithCache(source, state, fetchFn) {
  try {
    const freshData = await fetchFn();
    const freshCount = freshData && Array.isArray(freshData.data) ? freshData.data.length : 0;

    if (freshCount > 0) {
      // Attempt to cache — setCache will reject if data looks like a regression
      const cacheResult = await setCache(source, state, freshData);
      return { ...freshData, fromCache: false, cacheStatus: cacheResult.reason };
    }

    // Fresh fetch returned empty — serve cache instead
    const cached = await getCached(source, state);
    if (cached) return { ...cached, fromCache: true, note: 'Source returned empty. Showing cached data.' };
    return freshData;
  } catch (err) {
    // Fetch failed entirely — serve cache
    const cached = await getCached(source, state);
    if (cached) return { ...cached, fromCache: true, note: `Source unavailable (${err.message}). Showing cached data.` };
    throw err;
  }
}
