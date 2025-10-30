import { Site, Entry } from '../../types';
import { DatabaseService } from '../DatabaseService';

const db = DatabaseService.getInstance();

export function getStorageType(): 'database' | 'local' {
  return db.getStorageType();
}

export function isAvailable(): boolean {
  return db.isAvailable();
}

export async function testDatabaseConnection(): Promise<void> {
  const storageType = getStorageType();
  
  if (storageType === 'database') {
    // Test database connection by attempting a simple query
    try {
      await db.testConnection();
    } catch (error) {
      throw new Error(`Database connection test failed: ${error.message}`);
    }
  } else {
    throw new Error('Database not available - using local storage');
  }
}

export async function loadSitesFromDatabase(): Promise<Site[]> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    const sites = await db.loadSites();
    console.log(`Loaded ${sites.length} sites from database:`, sites.map(s => `${s.name}: ${s.entryCount} entries, ${s.sitemapEntryCount} sitemap`));
    return sites;
  } catch (error) {
    console.error('Failed to load sites from database:', error);
    throw error;
  }
}

export async function saveSitesToDatabase(sites: Site[]): Promise<void> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  console.log(`📊 Saving ${sites.length} sites to database:`, sites.map(s => `${s.name}: ${s.entryCount} entries, ${s.sitemapEntryCount} sitemap`));

  try {
    for (const site of sites) {
      await db.saveSite(site);
    }
    console.log(`📊 Sites saved to database successfully`);
  } catch (error) {
    console.error('Failed to save sites to database:', error);
    throw new Error(`Failed to save sites to database: ${error.message}`);
  }
}

export async function getSiteByUrlFromDatabase(url: string): Promise<Site | null> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    return await db.getSiteByUrl(url);
  } catch (error) {
    console.error('Failed to get site by URL from database:', error);
    throw error;
  }
}

export async function deleteSiteFromDatabase(siteId: string): Promise<void> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    await db.deleteSite(siteId);
  } catch (error) {
    console.error('Failed to delete site from database:', error);
    throw error;
  }
}

export async function getActualEntryCountFromDatabase(siteUrl: string): Promise<number> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    const count = await db.getEntryCount(siteUrl);
    console.log(`📊 Entry count from database: ${count} for ${new URL(siteUrl).hostname}`);
    return count;
  } catch (error) {
    // Handle network errors gracefully
    if (error instanceof TypeError && error.message === 'Failed to fetch') {
      console.warn(`Network error getting entry count for ${new URL(siteUrl).hostname}, returning 0`);
      return 0;
    }
    console.error('Failed to get entry count from database:', error);
    throw error;
  }
}

export async function getUnseenEntryCountFromDatabase(siteUrl: string): Promise<number> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    const count = await db.getUnseenEntryCount(siteUrl, true); // Pass true for recent only
    console.log(`📊 Unseen entry count from database: ${count} for ${new URL(siteUrl).hostname}`);
    return count;
  } catch (error) {
    // Handle network errors gracefully
    if (error instanceof TypeError && error.message === 'Failed to fetch') {
      console.warn(`Network error getting unseen entry count for ${new URL(siteUrl).hostname}, returning 0`);
      return 0;
    }
    console.error('Failed to get unseen entry count from database:', error);
    throw error;
  }
}

export async function getSitemapCountFromDatabase(siteUrl: string): Promise<number> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    const count = await db.getSitemapCount(siteUrl);
    console.log(`📊 Sitemap count from database: ${count} for ${new URL(siteUrl).hostname}`);
    
    if (count === 0) {
      console.warn(`📊 ⚠️ SITEMAP COUNT IS 0: Database returning 0 for ${new URL(siteUrl).hostname}`);
      const addLog = (window as any).addLog;
      if (addLog) {
        addLog(`⚠️ No sitemap entries found in database for ${new URL(siteUrl).hostname}`, 'warning');
      }
    }
    
    return count;
  } catch (error) {
    // Handle network errors gracefully
    if (error instanceof TypeError && error.message === 'Failed to fetch') {
      console.warn(`Network error getting sitemap count for ${new URL(siteUrl).hostname}, returning 0`);
      return 0;
    }
    console.error('Failed to get sitemap count from database:', error);
    throw error;
  }
}

export async function loadAllEntriesFromDatabase(): Promise<Entry[]> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    return await db.loadAllEntries();
  } catch (error) {
    console.error('Failed to load all entries from database:', error);
    throw error;
  }
}

export async function loadEntriesFromDatabase(siteUrl: string, limit?: number, offset?: number): Promise<Entry[]> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    const result = await db.loadEntriesWithLimit(siteUrl, limit, offset);
    
    if (result.length > 0) {
      console.log(`Database query completed: loaded ${result.length} entries (limit: ${limit || 'none'}, offset: ${offset || 0})`);
    }
    return result;
  } catch (error) {
    console.error('Failed to load entries from database:', error);
    throw error;
  }
}

export async function loadUnseenEntriesFromDatabase(siteUrl: string, limit: number, offset: number): Promise<Entry[]> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    return await db.loadUnseenEntriesWithLimit(siteUrl, limit, offset);
  } catch (error) {
    console.error('Failed to load unseen entries from database:', error);
    throw error;
  }
}

export async function loadEntriesByIdsFromDatabase(siteUrl: string, ids: string[]): Promise<Entry[]> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    return await db.loadEntriesByIds(siteUrl, ids);
  } catch (error) {
    console.error('Failed to load entries by IDs from database:', error);
    throw error;
  }
}

export async function saveEntryToDatabase(entry: Entry): Promise<void> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    await db.saveEntry(entry);
    console.log(`Entry saved to database: ${entry.title}`);
  } catch (error) {
    console.error('Failed to save entry to database:', error);
    throw new Error(`Failed to save entry to database: ${error.message}`);
  }
}

export async function saveEntriesBatchToDatabase(entries: Entry[]): Promise<void> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  console.log(`Batch saving ${entries.length} entries to database`);

  const addLog = (window as any).addLog;

  try {
    await db.saveEntriesBatch(entries);
    console.log(`Batch saved ${entries.length} entries to database`);
  } catch (error) {
    console.error('Failed to batch save entries to database:', error);
    if (addLog) {
      if (error.message && error.message.includes('timeout')) {
        addLog(`⚠️ Database timeout saving ${entries.length} entries`, 'warning');
      } else {
        addLog(`❌ Failed to save ${entries.length} entries to database: ${error.message}`, 'error');
      }
    }
    throw new Error(`Failed to batch save entries to database: ${error.message}`);
  }
}

export async function deleteEntryFromDatabase(entryId: string): Promise<void> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    await db.deleteEntry(entryId);
    console.log(`Entry deleted from database: ${entryId}`);
  } catch (error) {
    console.error('Failed to delete entry from database:', error);
    throw new Error(`Failed to delete entry from database: ${error.message}`);
  }
}

export async function deleteAllEntriesFromDatabase(siteUrl: string): Promise<void> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    await db.deleteAllEntries(siteUrl);
    console.log(`All entries deleted from database for ${siteUrl}`);
  } catch (error) {
    console.error('Failed to delete entries from database:', error);
    throw new Error(`Failed to delete entries from database: ${error.message}`);
  }
}

export async function loadSitemapFromDatabase(siteUrl: string): Promise<string[]> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  try {
    return await db.loadSitemap(siteUrl);
  } catch (error) {
    console.error('Failed to load sitemap from database:', error);
    throw error;
  }
}

export async function saveSitemapToDatabase(siteUrl: string, guids: string[]): Promise<void> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  console.log(`Saving sitemap for ${new URL(siteUrl).hostname} (${guids.length} entries) to database`);

  try {
    // First verify the site exists
    const site = await db.getSiteByUrl(siteUrl);
    if (!site) {
      throw new Error(`Site not found for URL: ${siteUrl}`);
    }
    console.log(`Found site for sitemap save: ${site.name} (ID: ${site.id})`);
    
    await db.saveSitemap(siteUrl, guids);
    console.log(`Sitemap saved to database successfully: ${guids.length} entries`);
  } catch (error) {
    console.error(`Failed to save sitemap to database for ${new URL(siteUrl).hostname}:`, error);
    throw new Error(`Failed to save sitemap to database: ${error.message}`);
  }
}

export async function getExistingEntryIdsFromDatabase(siteUrl: string, guids: string[]): Promise<string[]> {
  if (!db.isAvailable()) {
    throw new Error('Database not available');
  }

  console.log(`Getting existing entry IDs for ${siteUrl} from database (checking ${guids.length} GUIDs)`);

  const addLog = (window as any).addLog;

  if (guids.length === 0) {
    return [];
  }

  try {
    // Get the site to get its ID
    const site = await db.getSiteByUrl(siteUrl);
    if (!site) {
      return [];
    }
    
    return await db.getExistingEntryIds(site.id, guids);
  } catch (error) {
    if (error.message && error.message.includes('timeout')) {
      console.warn(`Database timeout getting existing entry IDs, falling back to local storage:`, error);
      if (addLog) {
        addLog(`⚠️ Database timeout checking existing entries for ${new URL(siteUrl).hostname}`, 'warning');
      }
      throw error; // Let the caller handle the fallback
    } else {
      console.warn(`Failed to get existing entry IDs from database:`, error);
      return [];
    }
  }
}