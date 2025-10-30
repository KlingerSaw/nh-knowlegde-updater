import { Site, Entry } from '../../types';
import { DatabaseConnection } from './connection';
import { SitesRepository } from './sites';
import { EntriesRepository } from './entries';
import { EntriesBatchRepository } from './entriesBatch';
import { SitemapsRepository } from './sitemaps';

export class DatabaseService {
  private static instance: DatabaseService;
  private connection: DatabaseConnection;
  private sitesRepo: SitesRepository | null = null;
  private entriesRepo: EntriesRepository | null = null;
  private entriesBatchRepo: EntriesBatchRepository | null = null;
  private sitemapsRepo: SitemapsRepository | null = null;

  private constructor() {
    this.connection = new DatabaseConnection();
    this.initializeRepositories();
  }

  static getInstance(): DatabaseService {
    if (!DatabaseService.instance) {
      DatabaseService.instance = new DatabaseService();
    }
    return DatabaseService.instance;
  }

  private initializeRepositories() {
    const supabase = this.connection.getSupabaseClient();
    if (supabase) {
      this.sitesRepo = new SitesRepository(supabase);
      this.entriesRepo = new EntriesRepository(supabase);
      this.entriesBatchRepo = new EntriesBatchRepository(supabase);
      this.sitemapsRepo = new SitemapsRepository(supabase);
    }
  }

  isAvailable(): boolean {
    return this.connection.isAvailable();
  }

  getStorageType(): 'database' | 'local' {
    return this.connection.getStorageType();
  }

  getSupabaseClient() {
    return this.connection.getSupabaseClient();
  }

  async testConnection(): Promise<void> {
    return this.connection.testConnection();
  }

  // Sites methods
  async loadSites(): Promise<Site[]> {
    if (!this.sitesRepo) throw new Error('Sites repository not available');
    return this.sitesRepo.loadSites();
  }

  async saveSite(site: Site): Promise<void> {
    if (!this.sitesRepo) throw new Error('Sites repository not available');
    return this.sitesRepo.saveSite(site);
  }

  async getSiteByUrl(url: string): Promise<Site | null> {
    if (!this.sitesRepo) throw new Error('Sites repository not available');
    return this.sitesRepo.getSiteByUrl(url);
  }

  async deleteSite(siteId: string): Promise<void> {
    if (!this.sitesRepo) throw new Error('Sites repository not available');
    return this.sitesRepo.deleteSite(siteId);
  }

  // Entries methods
  async getEntryCount(siteUrl: string): Promise<number> {
    if (!this.entriesRepo) throw new Error('Entries repository not available');
    
    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      console.log(`📊 No site found for URL ${siteUrl}, returning 0 entry count`);
      return 0;
    }

    return this.entriesRepo.getEntryCount(site.id, site.name);
  }

  async getUnseenEntryCount(siteUrl: string, recentOnly: boolean = false): Promise<number> {
    if (!this.entriesRepo) throw new Error('Entries repository not available');
    
    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      console.log(`📊 No site found for URL ${siteUrl}, returning 0 unseen entry count`);
      return 0;
    }

    return this.entriesRepo.getUnseenEntryCount(site.id, site.name, recentOnly);
  }

  async getExistingEntryIds(siteId: string, guids: string[]): Promise<string[]> {
    if (!this.entriesRepo) throw new Error('Entries repository not available');
    return this.entriesRepo.getExistingEntryIds(siteId, guids);
  }

  async loadEntriesWithLimit(siteUrl: string, limit?: number, offset?: number): Promise<Entry[]> {
    if (!this.entriesRepo) throw new Error('Entries repository not available');

    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      return [];
    }

    const entries = await this.entriesRepo.loadEntriesWithLimit(site.id, site.name, limit, offset);
    // Set the siteUrl for each entry
    return entries.map(entry => ({ ...entry, siteUrl }));
  }

  async loadEntriesByIds(siteUrl: string, ids: string[]): Promise<Entry[]> {
    if (!this.entriesRepo) throw new Error('Entries repository not available');

    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      return [];
    }

    const entries = await this.entriesRepo.loadEntriesByIds(site.id, site.name, ids);
    return entries.map(entry => ({ ...entry, siteUrl }));
  }

  async loadUnseenEntriesWithLimit(siteUrl: string, limit: number, offset: number): Promise<Entry[]> {
    if (!this.entriesRepo) throw new Error('Entries repository not available');

    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      return [];
    }

    const entries = await this.entriesRepo.loadUnseenEntriesWithLimit(site.id, site.name, limit, offset);
    // Set the siteUrl for each entry
    return entries.map(entry => ({ ...entry, siteUrl }));
  }

  async loadAllEntries(): Promise<Entry[]> {
    if (!this.entriesRepo || !this.sitesRepo) throw new Error('Repositories not available');
    
    const sites = await this.sitesRepo.loadSites();
    const allEntries: Entry[] = [];
    
    for (const site of sites) {
      const siteEntries = await this.entriesRepo.loadEntriesWithLimit(site.id, site.name);
      const entriesWithSiteUrl = siteEntries.map(entry => ({ ...entry, siteUrl: site.url }));
      allEntries.push(...entriesWithSiteUrl);
    }
    
    return allEntries;
  }

  // Entry batch operations
  async saveEntry(entry: Entry): Promise<void> {
    if (!this.entriesBatchRepo) throw new Error('Entries batch repository not available');
    
    const site = await this.getSiteByUrl(entry.siteUrl!);
    if (!site) {
      throw new Error(`Site not found for URL: ${entry.siteUrl}`);
    }

    return this.entriesBatchRepo.saveEntry(entry, site.id);
  }

  async saveEntriesBatch(entries: Entry[]): Promise<void> {
    if (!this.entriesBatchRepo) throw new Error('Entries batch repository not available');
    if (!this.entriesRepo) throw new Error('Entries repository not available');

    if (entries.length === 0) return;

    const site = await this.getSiteByUrl(entries[0].siteUrl!);
    if (!site) {
      throw new Error(`Site not found for URL: ${entries[0].siteUrl}`);
    }

    // Only persist entries that aren't already stored
    const guids = entries.map(e => e.id);
    const existingIds = new Set(await this.entriesRepo.getExistingEntryIds(site.id, guids));
    const newEntries = entries.filter(e => !existingIds.has(e.id));

    if (newEntries.length === 0) {
      console.log(`No new entries to save for site ${site.name}`);
      return;
    }

    return this.entriesBatchRepo.saveEntriesBatch(newEntries, site.id);
  }

  async deleteEntry(entryId: string): Promise<void> {
    if (!this.entriesBatchRepo) throw new Error('Entries batch repository not available');
    return this.entriesBatchRepo.deleteEntry(entryId);
  }

  async deleteAllEntries(siteUrl: string): Promise<void> {
    if (!this.entriesBatchRepo) throw new Error('Entries batch repository not available');
    
    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      return; // No site found, nothing to delete
    }

    return this.entriesBatchRepo.deleteAllEntries(site.id);
  }

  // Sitemaps methods
  async getSitemapCount(siteUrl: string): Promise<number> {
    if (!this.sitemapsRepo) throw new Error('Sitemaps repository not available');
    
    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      console.log(`No site found for URL ${siteUrl}, returning 0 sitemap count`);
      console.warn(`⚠️ SITEMAP COUNT SET TO 0: No site found for URL ${siteUrl}`);
      return 0;
    }

    return this.sitemapsRepo.getSitemapCount(site.id, site.name);
  }

  async loadSitemap(siteUrl: string): Promise<string[]> {
    if (!this.sitemapsRepo) throw new Error('Sitemaps repository not available');
    
    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      return [];
    }

    return this.sitemapsRepo.loadSitemap(site.id);
  }

  async saveSitemap(siteUrl: string, guids: string[]): Promise<void> {
    if (!this.sitemapsRepo) throw new Error('Sitemaps repository not available');
    
    const site = await this.getSiteByUrl(siteUrl);
    if (!site) {
      throw new Error(`Site not found for URL: ${siteUrl}`);
    }

    return this.sitemapsRepo.saveSitemap(site.id, site.name, guids);
  }
}