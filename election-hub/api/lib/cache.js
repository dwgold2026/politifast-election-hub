/**
 * Redis cache layer for candidate data.
 * Stores collected data so we don't lose it if a source goes down.
 *
 * Keys: "candidates:{source}:{state}" or "candidates:{source}"
 * Values: JSON stringified { updated, data, meta }
 * TTL: 7 days (data refreshed on each successful collection)
 */
import { Redis } from '@upstash/redis';

const redis = new Redis({
  url: process.env.KV_REST_API_URL,
  token: process.env.KV_REST_API_TOKEN,
});

const CACHE_TTL = 7 * 24 * 60 * 60; // 7 days in seconds

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
 * Store candidate data in cache
 */
export async function setCache(source, state, data) {
  try {
    const key = state ? `candidates:${source}:${state}` : `candidates:${source}`;
    const payload = {
      updated: new Date().toISOString(),
      cachedAt: new Date().toISOString(),
      ...data,
    };
    await redis.set(key, JSON.stringify(payload), { ex: CACHE_TTL });
    return true;
  } catch {
    return false;
  }
}

/**
 * Wrapper: try to fetch fresh data, fall back to cache if fetch fails
 */
export async function fetchWithCache(source, state, fetchFn) {
  try {
    // Try fresh fetch
    const freshData = await fetchFn();
    if (freshData && freshData.data && freshData.data.length > 0) {
      // Cache the successful result
      await setCache(source, state, freshData);
      return { ...freshData, fromCache: false };
    }
    // Fresh fetch returned empty — try cache
    const cached = await getCached(source, state);
    if (cached) return { ...cached, fromCache: true, note: 'Source returned empty. Showing cached data.' };
    return freshData; // Return the empty result
  } catch (err) {
    // Fetch failed — try cache
    const cached = await getCached(source, state);
    if (cached) return { ...cached, fromCache: true, note: `Source unavailable (${err.message}). Showing cached data.` };
    throw err; // No cache available either
  }
}
