import { Site, Entry } from '../../types';
import { STORAGE_CONFIG, sanitizeUrl } from './constants';

export async function loadSitesFromLocal(): Promise<Site[]> {
  try {
    const sitesData = localStorage.getItem(STORAGE_CONFIG.sitesKey);
    if (!sitesData) return [];
    
    const sites = JSON.parse(sitesData);
    return sites.map((site: any) => ({
      ...site,
      lastUpdated: new Date(site.lastUpdated)
    }));
  } catch (error) {
    console.error('Failed to load sites from local storage:', error);
    return [];
  }
}

export async function loadEntriesFromLocal(siteUrl: string, limit?: number, offset?: number): Promise<Entry[]> {
  try {
    const key = STORAGE_CONFIG.entriesPrefix + sanitizeUrl(siteUrl);
    const entriesData = localStorage.getItem(key);
    if (!entriesData) return [];
    
    let entries = JSON.parse(entriesData);
    
    // Sort by published date (newest first)
    entries.sort((a: Entry, b: Entry) => 
      new Date(b.publishedDate).getTime() - new Date(a.publishedDate).getTime()
    );
    
    // Apply pagination if specified
    if (offset || limit) {
      const start = offset || 0;
      const end = limit ? start + limit : undefined;
      entries = entries.slice(start, end);
    }
    
    return entries;
  } catch (error) {
    console.error('Failed to load entries from local storage:', error);
    return [];
  }
}

export async function loadAllEntriesFromLocal(sites: Site[]): Promise<Entry[]> {
  const allEntries: Entry[] = [];

  for (const site of sites) {
    const siteEntries = await loadEntriesFromLocal(site.url);
    allEntries.push(...siteEntries);
  }

  return allEntries;
}

export async function loadEntriesByIdsFromLocal(siteUrl: string, ids: string[]): Promise<Entry[]> {
  if (ids.length === 0) {
    return [];
  }

  const entries = await loadEntriesFromLocal(siteUrl);
  const idSet = new Set(ids);
  return entries.filter(entry => idSet.has(entry.id));
}

export async function loadUnseenEntriesFromLocal(siteUrl: string, limit: number, offset: number, recentOnly: boolean = false): Promise<Entry[]> {
  try {
    const key = STORAGE_CONFIG.entriesPrefix + sanitizeUrl(siteUrl);
    const entriesData = localStorage.getItem(key);
    if (!entriesData) return [];
    
    let entries = JSON.parse(entriesData);
    
    // Filter for unseen entries first
    let unseenEntries = entries.filter((entry: Entry) => !entry.seen);
    
    // If recentOnly is true, only include entries from the last 24 hours
    if (recentOnly) {
      const oneDayAgo = new Date();
      oneDayAgo.setDate(oneDayAgo.getDate() - 1);
      
      unseenEntries = unseenEntries.filter((entry: Entry) => {
        // Check when entry was stored in database (created_at), fall back to published_date only if no created_at
        const entryDate = entry.metadata?.created_at ? 
          new Date(entry.metadata.created_at) : 
          (entry.publishedDate ? new Date(entry.publishedDate) : new Date(0));
        return entryDate >= oneDayAgo;
      });
    }
    
    const recentText = recentOnly ? ' (recent only)' : '';
    console.log(`Local storage: Found ${unseenEntries.length} unseen entries${recentText} out of ${entries.length} total`);
    // Sort by published date (newest first)
    unseenEntries.sort((a: Entry, b: Entry) => 
      new Date(b.publishedDate).getTime() - new Date(a.publishedDate).getTime()
    );
    
    // Apply pagination
    const start = offset;
    const end = start + limit;
    return unseenEntries.slice(start, end);
  } catch (error) {
    console.error('Failed to load unseen entries from local storage:', error);
    return [];
  }
}

export async function saveEntryToLocal(siteUrl: string, entry: Entry): Promise<void> {
  try {
    const entries = await loadEntriesFromLocal(siteUrl);
    const existingIndex = entries.findIndex(e => e.id === entry.id);
    
    if (existingIndex >= 0) {
      entries[existingIndex] = entry;
    } else {
      entries.push(entry);
    }
    
    const key = STORAGE_CONFIG.entriesPrefix + sanitizeUrl(siteUrl);
    localStorage.setItem(key, JSON.stringify(entries));
    console.log(`Entry saved to local storage: ${entry.title}`);
  } catch (error) {
    if (error.name === 'QuotaExceededError') {
      console.warn('Local storage quota exceeded. Clearing old entries and retrying...');
      
      try {
        // Clear some old entries to make space
        await cleanupLocalStorage();
        
        // Retry saving
        const entries = await loadEntriesFromLocal(siteUrl);
        const existingIndex = entries.findIndex(e => e.id === entry.id);
        
        if (existingIndex >= 0) {
          entries[existingIndex] = entry;
        } else {
          entries.push(entry);
        }
        
        const key = STORAGE_CONFIG.entriesPrefix + sanitizeUrl(siteUrl);
        localStorage.setItem(key, JSON.stringify(entries));
        console.log(`Entry saved to local storage after cleanup: ${entry.title}`);
      } catch (retryError) {
        console.error('Failed to save entry even after cleanup. Local storage is full.');
        throw new Error(`Local storage quota exceeded. Entry "${entry.title}" could not be saved to local storage.`);
      }
    } else {
      console.error('Failed to save entry to local storage:', error);
      throw error;
    }
  }
}

export async function loadSitemapFromLocal(siteUrl: string): Promise<string[]> {
  try {
    const key = STORAGE_CONFIG.sitemapPrefix + sanitizeUrl(siteUrl);
    const sitemapData = localStorage.getItem(key);
    if (!sitemapData) return [];
    
    return JSON.parse(sitemapData);
  } catch (error) {
    console.error('Failed to load sitemap from local storage:', error);
    return [];
  }
}

export function getStorageUsage(): { used: number; total: number; percentage: number } {
  try {
    let totalSize = 0;
    for (let key in localStorage) {
      if (localStorage.hasOwnProperty(key)) {
        totalSize += localStorage[key].length + key.length;
      }
    }
    
    // Estimate total available storage (5MB is typical for localStorage)
    const estimatedTotal = 5 * 1024 * 1024; // 5MB in bytes
    const percentage = (totalSize / estimatedTotal) * 100;
    
    return {
      used: totalSize,
      total: estimatedTotal,
      percentage: Math.min(percentage, 100)
    };
  } catch (error) {
    console.error('Failed to calculate storage usage:', error);
    return { used: 0, total: 0, percentage: 0 };
  }
}

/**
 * Clean up local storage by removing old entries
 */
export async function cleanupLocalStorage(): Promise<void> {
  try {
    console.log('Cleaning up local storage...');
    
    // Get all entry keys from localStorage
    const entryKeys = Object.keys(localStorage).filter(key => 
      key.startsWith(STORAGE_CONFIG.entriesPrefix)
    );
    
    let totalCleaned = 0;
    
    for (const key of entryKeys) {
      try {
        const entries = JSON.parse(localStorage.getItem(key) || '[]');
        
        if (entries.length > 50) {
          // Keep only the 50 most recent entries (by publishedDate)
          entries.sort((a: Entry, b: Entry) => 
            new Date(b.publishedDate).getTime() - new Date(a.publishedDate).getTime()
          );
          const keptEntries = entries.slice(0, 50);
          const removedCount = entries.length - keptEntries.length;
          
          localStorage.setItem(key, JSON.stringify(keptEntries));
          totalCleaned += removedCount;
          console.log(`Cleaned ${removedCount} old entries from ${key}`);
        }
      } catch (error) {
        console.warn(`Failed to cleanup entries for key ${key}:`, error);
      }
    }
    
    console.log(`Local storage cleanup complete. Removed ${totalCleaned} old entries total.`);
  } catch (error) {
    console.error('Failed to cleanup local storage:', error);
  }
}