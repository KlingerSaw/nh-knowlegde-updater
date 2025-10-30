import { Site, Entry } from '../../types';
import type { StorageUsage } from './types';
import { sanitizeUrl } from './constants';
import {
  getStorageType,
  isAvailable,
  testDatabaseConnection,
  loadSitesFromDatabase,
  saveSitesToDatabase,
  getSiteByUrlFromDatabase,
  deleteSiteFromDatabase,
  getActualEntryCountFromDatabase,
  getUnseenEntryCountFromDatabase,
  getSitemapCountFromDatabase,
  loadAllEntriesFromDatabase,
  loadEntriesFromDatabase,
  saveEntryToDatabase,
  saveEntriesBatchToDatabase,
  deleteEntryFromDatabase,
  deleteAllEntriesFromDatabase,
  loadSitemapFromDatabase,
  saveSitemapToDatabase,
  getExistingEntryIdsFromDatabase,
  loadUnseenEntriesFromDatabase,
  loadEntriesByIdsFromDatabase
} from './database';
import {
  loadSitesFromLocal,
  loadEntriesFromLocal,
  loadAllEntriesFromLocal,
  saveEntryToLocal,
  loadSitemapFromLocal,
  loadUnseenEntriesFromLocal,
  loadEntriesByIdsFromLocal,
  getStorageUsage,
  cleanupLocalStorage
} from './localStorage';

export type { StorageUsage } from './types';

export class StorageService {
  static async testDatabaseConnection(): Promise<void> {
    return testDatabaseConnection();
  }

  static getStorageType(): 'database' | 'local' {
    return getStorageType();
  }

  static async loadSites(): Promise<Site[]> {
    const storageType = getStorageType();
    console.log(`Loading sites from ${storageType}`);

    try {
      if (isAvailable()) {
        return await loadSitesFromDatabase();
      } else {
        return await loadSitesFromLocal();
      }
    } catch (error) {
      console.warn(`Failed to load sites from ${storageType}, falling back to local storage:`, error);
      return await loadSitesFromLocal();
    }
  }

  static async saveSites(sites: Site[]): Promise<void> {
    const storageType = getStorageType();

    // Always save to database if available
    if (isAvailable()) {
      return await saveSitesToDatabase(sites);
    }

    // Only use local storage if database is not available
    throw new Error('Database not available - sites must be saved to database');
  }

  static async getSiteByUrl(url: string): Promise<Site | null> {
    const storageType = getStorageType();
    console.log(`Checking for existing site with URL ${url} in ${storageType}`);

    try {
      if (isAvailable()) {
        return await getSiteByUrlFromDatabase(url);
      } else {
        // Check local storage
        const sites = await loadSitesFromLocal();
        return sites.find(site => site.url === url) || null;
      }
    } catch (error) {
      console.warn(`Failed to check for existing site in ${storageType}, falling back to local storage:`, error);
      const sites = await loadSitesFromLocal();
      return sites.find(site => site.url === url) || null;
    }
  }

  static async getExistingEntryIdsForSite(siteUrl: string, guids: string[]): Promise<string[]> {
    const storageType = getStorageType();

    if (guids.length === 0) {
      return [];
    }

    try {
      if (isAvailable()) {
        return await getExistingEntryIdsFromDatabase(siteUrl, guids);
      } else {
        // Fallback to local storage
        const entries = await loadEntriesFromLocal(siteUrl);
        const existingIds = new Set(entries.map(entry => entry.id));
        return guids.filter(guid => existingIds.has(guid));
      }
    } catch (error) {
      if (error.message && error.message.includes('timeout')) {
        console.warn(`Database timeout getting existing entry IDs, trying local storage fallback:`, error);
        try {
          const entries = await loadEntriesFromLocal(siteUrl);
          const existingIds = new Set(entries.map(entry => entry.id));
          const result = guids.filter(guid => existingIds.has(guid));
          console.log(`Fallback to local storage returned ${result.length} existing IDs`);
          return result;
        } catch (localError) {
          console.warn('Local storage fallback also failed:', localError);
          return [];
        }
      } else {
        console.warn(`Failed to get existing entry IDs from ${storageType}:`, error);
        return [];
      }
    }
  }

  static async deleteSite(siteId: string): Promise<void> {
    const storageType = getStorageType();
    console.log(`Deleting site ${siteId} from ${storageType}`);

    try {
      const sites = await this.loadSites();
      const site = sites.find(s => s.id === siteId);
      
      if (site) {
        // Delete entries and sitemap for this site
        await this.deleteAllEntries(site.url);
        
        // Delete from database if available
        if (isAvailable()) {
          await deleteSiteFromDatabase(siteId);
        }
        
        // Delete from local storage
        localStorage.removeItem('knowledge-updater-sitemap-' + sanitizeUrl(site.url));
      }
    } catch (error) {
      console.error('Failed to delete site:', error);
      throw new Error('Failed to delete site from storage');
    }
  }

  static async loadEntries(siteUrl: string): Promise<Entry[]> {
    // For the main entry list, load all entries without limit
    return this.loadAllEntriesForSite(siteUrl);
  }

  static async loadAllEntriesForSite(siteUrl: string): Promise<Entry[]> {
    const storageType = getStorageType();
    console.log(`Loading ALL entries for ${siteUrl} from ${storageType}`);

    const addLog = (window as any).addLog;

    try {
      if (isAvailable()) {
        return await loadEntriesFromDatabase(siteUrl);
      } else {
        return await loadEntriesFromLocal(siteUrl);
      }
    } catch (error) {
      if (error.message && error.message.includes('timeout')) {
        console.warn(`Database timeout loading entries from ${storageType}, falling back to local storage:`, error);
        if (addLog) {
          addLog(`⚠️ Database timeout loading entries for ${new URL(siteUrl).hostname}, using local storage`, 'warning');
        }
        try {
          const localEntries = await loadEntriesFromLocal(siteUrl);
          console.log(`Local storage fallback returned ${localEntries.length} entries`);
          return localEntries;
        } catch (localError) {
          console.warn('Local storage fallback failed:', localError);
          return [];
        }
      } else {
        console.warn(`Failed to load all entries from ${storageType} (${error.message}), falling back to local storage:`, error);
        return await loadEntriesFromLocal(siteUrl);
      }
    }
  }

  static async loadEntriesWithLimit(siteUrl: string, limit?: number, offset?: number): Promise<Entry[]> {
    const storageType = getStorageType();
    console.log(`Loading entries for ${siteUrl} from ${storageType} (limit: ${limit || 'none'}, offset: ${offset || 0})`);

    try {
      if (isAvailable()) {
        return await loadEntriesFromDatabase(siteUrl, limit, offset);
      } else {
        return await loadEntriesFromLocal(siteUrl, limit, offset);
      }
    } catch (error) {
      console.warn(`Failed to load entries from ${storageType} (${error.message}), falling back to local storage:`, error);
      return await loadEntriesFromLocal(siteUrl, limit, offset);
    }
  }

  static async loadEntriesByIds(siteUrl: string, ids: string[]): Promise<Entry[]> {
    const storageType = getStorageType();
    console.log(`Loading ${ids.length} entries by ID for ${siteUrl} from ${storageType}`);

    if (ids.length === 0) {
      return [];
    }

    try {
      if (isAvailable()) {
        return await loadEntriesByIdsFromDatabase(siteUrl, ids);
      }

      return await loadEntriesByIdsFromLocal(siteUrl, ids);
    } catch (error) {
      console.warn(`Failed to load entries by IDs from ${storageType}:`, error);

      if (storageType === 'database') {
        try {
          return await loadEntriesByIdsFromLocal(siteUrl, ids);
        } catch (localError) {
          console.warn('Local storage fallback failed for loadEntriesByIds:', localError);
          return [];
        }
      }

      return [];
    }
  }

  static async loadUnseenEntriesWithLimit(siteUrl: string, limit: number, offset: number): Promise<Entry[]> {
    const storageType = getStorageType();
    console.log(`Loading ${limit} unseen entries for ${siteUrl} from ${storageType} (offset: ${offset})`);

    try {
      if (isAvailable()) {
        return await loadUnseenEntriesFromDatabase(siteUrl, limit, offset);
      } else {
        return await loadUnseenEntriesFromLocal(siteUrl, limit, offset, true); // Pass true for recent only
      }
    } catch (error) {
      console.warn(`Failed to load unseen entries from ${storageType} (${error.message}), falling back to local storage:`, error);
      return await loadUnseenEntriesFromLocal(siteUrl, limit, offset, true); // Pass true for recent only
    }
  }

  static async loadAllEntries(): Promise<Entry[]> {
    const storageType = getStorageType();
    console.log(`Loading all entries from ${storageType}`);

    try {
      if (isAvailable()) {
        return await loadAllEntriesFromDatabase();
      } else {
        // Load from all local storage entries
        const sites = await this.loadSites();
        return await loadAllEntriesFromLocal(sites);
      }
    } catch (error) {
      console.warn(`Failed to load all entries from ${storageType}:`, error);
      return [];
    }
  }

  static async saveEntry(siteUrl: string, entry: Entry): Promise<void> {
    // Always save to database if available
    if (isAvailable()) {
      return await saveEntryToDatabase(entry);
    }

    // Only use local storage if database is not available
    throw new Error('Database not available - entries must be saved to database');
  }

  static async saveEntriesBatch(siteUrl: string, entries: Entry[]): Promise<void> {
    const storageType = getStorageType();
    console.log(`Batch saving ${entries.length} entries to ${storageType}`);

    // Always save to database if available
    if (isAvailable()) {
      return await saveEntriesBatchToDatabase(entries);
    }

    // Only use local storage if database is not available
    throw new Error('Database not available - entries must be saved to database');
  }

  static async markEntryAsSeen(entry: Entry): Promise<void> {
    const storageType = getStorageType();
    console.log(`Marking entry "${entry.title}" as seen in ${storageType}`);

    try {
      if (!entry.siteUrl) {
        throw new Error('Entry must have a siteUrl to mark as seen');
      }
      await this.saveEntry(entry.siteUrl, entry);
    } catch (error) {
      console.error('Failed to mark entry as seen:', error);
      throw new Error('Failed to mark entry as seen');
    }
  }

  static async deleteEntry(entry: Entry): Promise<void> {
    const storageType = getStorageType();
    console.log(`Deleting entry "${entry.title}" from ${storageType}`);

    // Always delete from database if available
    if (isAvailable()) {
      return await deleteEntryFromDatabase(entry.id);
    }

    // Only use local storage if database is not available
    throw new Error('Database not available - entries must be deleted from database');
  }

  static async deleteAllEntries(siteUrl: string): Promise<void> {
    const storageType = getStorageType();
    console.log(`Deleting all entries for ${siteUrl} from ${storageType}`);

    // Always delete from database if available
    if (isAvailable()) {
      return await deleteAllEntriesFromDatabase(siteUrl);
    }

    // Only use local storage if database is not available
    throw new Error('Database not available - entries must be deleted from database');
  }

  static async getActualEntryCount(siteUrl: string): Promise<number> {
    const storageType = getStorageType();
    console.log(`📊 Getting actual entry count for ${new URL(siteUrl).hostname} from ${storageType}`);

    const addLog = (window as any).addLog;

    try {
      if (isAvailable()) {
        return await getActualEntryCountFromDatabase(siteUrl);
      } else {
        const entries = await loadEntriesFromLocal(siteUrl);
        const count = entries.length;
        console.log(`📊 Entry count from local storage: ${count} for ${new URL(siteUrl).hostname}`);
        return count;
      }
    } catch (error) {
      console.warn(`📊 Failed to get entry count from ${storageType} for ${new URL(siteUrl).hostname}:`, error);
      
      // Try fallback to local storage if database fails
      if (storageType === 'database') {
        console.log(`📊 Trying local storage fallback for entry count...`);
        try {
          const entries = await loadEntriesFromLocal(siteUrl);
          const count = entries.length;
          console.log(`📊 Fallback: Entry count from local storage: ${count} for ${new URL(siteUrl).hostname}`);
          if (addLog) {
            addLog(`⚠️ Used local storage fallback for entry count: ${count} for ${new URL(siteUrl).hostname}`, 'warning');
          }
          return count;
        } catch (fallbackError) {
          console.warn(`📊 Local storage fallback also failed:`, fallbackError);
        }
      }
      
      if (addLog) {
        if (error.message && error.message.includes('timeout')) {
          addLog(`⏱️ Database timeout getting entry count for ${new URL(siteUrl).hostname}`, 'warning');
        } else {
          addLog(`❌ Failed to get entry count for ${new URL(siteUrl).hostname}: ${error.message}`, 'error');
        }
      }
      return 0;
    }
  }

  static async getUnseenEntryCount(siteUrl: string): Promise<number> {
    const storageType = getStorageType();
    console.log(`📊 Getting unseen entry count for ${new URL(siteUrl).hostname} from ${storageType}`);

    const addLog = (window as any).addLog;

    try {
      if (isAvailable()) {
        return await getUnseenEntryCountFromDatabase(siteUrl);
      } else {
        const entries = await loadEntriesFromLocal(siteUrl);
        
        // Filter for unseen entries from the last 24 hours only
        const oneDayAgo = new Date();
        oneDayAgo.setDate(oneDayAgo.getDate() - 1);
        
        const recentUnseenEntries = entries.filter(entry => {
          if (entry.seen) return false;
          
          // Check when entry was stored in database (created_at), fall back to published_date only if no created_at
          const entryDate = entry.metadata?.created_at ? 
            new Date(entry.metadata.created_at) : 
            (entry.publishedDate ? new Date(entry.publishedDate) : new Date(0));
          return entryDate >= oneDayAgo;
        });
        
        const count = recentUnseenEntries.length;
        console.log(`📊 Unseen entry count from local storage: ${count} for ${new URL(siteUrl).hostname}`);
        return count;
      }
    } catch (error) {
      console.warn(`📊 Failed to get unseen entry count from ${storageType} for ${new URL(siteUrl).hostname}:`, error);
      
      // Try fallback to local storage if database fails
      if (storageType === 'database') {
        console.log(`📊 Trying local storage fallback for unseen entry count...`);
        try {
          const entries = await loadEntriesFromLocal(siteUrl);
          
          // Filter for unseen entries from the last 24 hours only
          const oneDayAgo = new Date();
          oneDayAgo.setDate(oneDayAgo.getDate() - 1);
          
          const recentUnseenEntries = entries.filter(entry => {
            if (entry.seen) return false;
            
            // Check when entry was stored in database (created_at), fall back to published_date only if no created_at
            const entryDate = entry.metadata?.created_at ? 
              new Date(entry.metadata.created_at) : 
              (entry.publishedDate ? new Date(entry.publishedDate) : new Date(0));
            return entryDate >= oneDayAgo;
          });
          
          const count = recentUnseenEntries.length;
          console.log(`📊 Fallback: Unseen entry count from local storage: ${count} for ${new URL(siteUrl).hostname}`);
          if (addLog) {
            addLog(`⚠️ Used local storage fallback for unseen count: ${count} for ${new URL(siteUrl).hostname}`, 'warning');
          }
          return count;
        } catch (fallbackError) {
          console.warn(`📊 Local storage fallback also failed:`, fallbackError);
        }
      }
      
      if (addLog) {
        if (error.message && error.message.includes('timeout')) {
          addLog(`⏱️ Database timeout getting unseen entry count for ${new URL(siteUrl).hostname}`, 'warning');
        } else {
          addLog(`❌ Failed to get unseen entry count for ${new URL(siteUrl).hostname}: ${error.message}`, 'error');
        }
      }
      return 0;
    }
  }

  static async getSitemapCount(siteUrl: string): Promise<number> {
    const storageType = getStorageType();
    console.log(`📊 Getting sitemap count for ${new URL(siteUrl).hostname} from ${storageType}`);

    const addLog = (window as any).addLog;

    try {
      if (isAvailable()) {
        return await getSitemapCountFromDatabase(siteUrl);
      } else {
        const sitemap = await loadSitemapFromLocal(siteUrl);
        
        if (sitemap.length === 0) {
          console.warn(`📊 ⚠️ SITEMAP COUNT IS 0: Local storage has 0 entries for ${new URL(siteUrl).hostname}`);
          if (addLog) {
            addLog(`⚠️ No sitemap entries found in local storage for ${new URL(siteUrl).hostname}`, 'warning');
          }
        }
        
        console.log(`📊 Sitemap count from local storage: ${sitemap.length} for ${new URL(siteUrl).hostname}`);
        return sitemap.length;
      }
    } catch (error) {
      console.warn(`📊 Failed to get sitemap count from ${storageType}:`, error);
      console.warn(`📊 ⚠️ SITEMAP COUNT SET TO 0: Exception in getSitemapCount for ${new URL(siteUrl).hostname}: ${error.message}`);
      if (addLog) {
        if (error.message && error.message.includes('timeout')) {
          addLog(`⚠️ Database timeout getting sitemap count for ${new URL(siteUrl).hostname}`, 'warning');
        } else {
          addLog(`❌ Failed to get sitemap count for ${new URL(siteUrl).hostname}: ${error.message}`, 'error');
        }
      }
      return 0;
    }
  }

  static async loadSitemap(siteUrl: string): Promise<string[]> {
    const storageType = getStorageType();
    console.log(`Loading sitemap for ${siteUrl} from ${storageType}`);

    try {
      if (isAvailable()) {
        return await loadSitemapFromDatabase(siteUrl);
      } else {
        return await loadSitemapFromLocal(siteUrl);
      }
    } catch (error) {
      console.warn(`Failed to load sitemap from ${storageType}, falling back to local storage:`, error);
      return await loadSitemapFromLocal(siteUrl);
    }
  }

  static async saveSitemap(siteUrl: string, guids: string[]): Promise<void> {
    const storageType = getStorageType();
    console.log(`Saving sitemap for ${new URL(siteUrl).hostname} (${guids.length} entries) to ${storageType}`);

    // Always save to database if available
    if (isAvailable()) {
      return await saveSitemapToDatabase(siteUrl, guids);
    }

    // Only use local storage if database is not available
    throw new Error('Database not available - sitemaps must be saved to database');
  }

  static getStorageUsage(): StorageUsage {
    return getStorageUsage();
  }
}