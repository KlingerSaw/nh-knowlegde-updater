import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { Entry, Site } from '../types';
import { StorageService } from './StorageService';
import { ApiService } from './ApiService';

export type ExcelImportStage =
  | 'parsing'
  | 'matching'
  | 'fetching'
  | 'saving'
  | 'zipping'
  | 'complete';

export interface ExcelImportProgress {
  stage: ExcelImportStage;
  message: string;
  current?: number;
  total?: number;
}

export interface ExcelImportSiteSummary {
  siteId: string;
  siteName: string;
  totalLinks: number;
  uniqueEntries: number;
  newEntriesSaved: number;
  existingEntries: number;
  failedEntries: number;
}

export interface ExcelImportSummary {
  totalLinks: number;
  validLinks: number;
  processedEntries: number;
  newEntriesSaved: number;
  failedLinks: string[];
  unmatchedLinks: string[];
  siteSummaries: ExcelImportSiteSummary[];
}

export interface ExcelImportResult {
  zipBlob: Blob;
  zipFileName: string;
  summary: ExcelImportSummary;
}

interface ParsedLink {
  originalUrl: string;
  site?: Site | null;
  guid?: string | null;
}

interface SiteProcessingData {
  site: Site;
  guids: string[];
  urlByGuid: Map<string, string>;
  totalLinks: number;
}

export class ExcelImportService {
  static async processFile(
    file: File,
    sites: Site[],
    onProgress?: (progress: ExcelImportProgress) => void
  ): Promise<ExcelImportResult> {
    if (!file) {
      throw new Error('No file provided');
    }

    if (!sites || sites.length === 0) {
      throw new Error('No sites configured. Please add sites before importing an Excel file.');
    }

    onProgress?.({ stage: 'parsing', message: `Reading ${file.name}...` });

    const workbook = await this.readWorkbook(file);
    const rawUrls = this.extractUrlsFromWorkbook(workbook);

    if (rawUrls.length === 0) {
      throw new Error('The uploaded Excel file did not contain any valid URLs.');
    }

    onProgress?.({
      stage: 'matching',
      message: `Processing ${rawUrls.length} links from workbook...`
    });

    const { uniqueLinks, unmatchedLinks, siteData, validLinkCount } = this.prepareLinkData(rawUrls, sites);

    if (uniqueLinks.length === 0) {
      throw new Error('No links in the Excel file matched known sites or contained recognizable GUIDs.');
    }

    const failedLinks = new Set<string>();
    unmatchedLinks.forEach(url => failedLinks.add(url));

    const entryMap = new Map<string, Entry>();
    const siteSummaries: ExcelImportSiteSummary[] = [];
    let totalNewEntriesSaved = 0;

    for (const [siteId, info] of siteData.entries()) {
      onProgress?.({
        stage: 'fetching',
        message: `Checking existing entries for ${info.site.name}...`
      });

      const existingIds = await StorageService.getExistingEntryIdsForSite(info.site.url, info.guids);
      const existingIdSet = new Set(existingIds);

      const missingGuids = info.guids.filter(guid => !existingIdSet.has(guid));
      let fetchedEntries: Entry[] = [];
      let savedEntriesForSite = 0;
      let failedForSite = 0;

      if (missingGuids.length > 0) {
        onProgress?.({
          stage: 'fetching',
          message: `Fetching ${missingGuids.length} new entries for ${info.site.name}...`,
          current: 0,
          total: missingGuids.length
        });

        const { success: fetchedData, failed: failedGuids } = await ApiService.fetchPublicationsBatch(
          info.site.url,
          missingGuids,
          undefined,
          (completed, total) => {
            onProgress?.({
              stage: 'fetching',
              message: `Fetching entries for ${info.site.name}: ${completed}/${total}`,
              current: completed,
              total
            });
          }
        );

        fetchedEntries = fetchedData.map(data => this.mapToEntry(data, info.site.url));

        if (fetchedEntries.length > 0) {
          onProgress?.({
            stage: 'saving',
            message: `Saving ${fetchedEntries.length} entries for ${info.site.name}...`
          });

          await StorageService.saveEntriesBatch(info.site.url, fetchedEntries);
          savedEntriesForSite = fetchedEntries.length;
          totalNewEntriesSaved += fetchedEntries.length;
        }

        if (failedGuids.length > 0) {
          failedForSite += failedGuids.length;
          failedGuids.forEach(guid => {
            const originalUrl = info.urlByGuid.get(guid);
            if (originalUrl) {
              failedLinks.add(originalUrl);
            }
          });
        }
      }

      const existingEntries = await StorageService.loadEntriesByIds(
        info.site.url,
        info.guids.filter(guid => existingIdSet.has(guid))
      );

      // Combine new and existing entries into the map for zip generation
      [...existingEntries, ...fetchedEntries].forEach(entry => {
        const key = `${siteId}|${entry.id}`;
        entryMap.set(key, { ...entry, siteUrl: info.site.url });
      });

      siteSummaries.push({
        siteId,
        siteName: info.site.name,
        totalLinks: info.totalLinks,
        uniqueEntries: info.guids.length,
        newEntriesSaved: savedEntriesForSite,
        existingEntries: info.guids.length - missingGuids.length,
        failedEntries: failedForSite
      });
    }

    const orderedEntries: Entry[] = [];
    const seenKeys = new Set<string>();

    for (const link of uniqueLinks) {
      if (!link.site || !link.guid) {
        continue;
      }

      const key = `${link.site.id}|${link.guid}`;
      if (seenKeys.has(key)) {
        continue;
      }

      const entry = entryMap.get(key);
      if (entry) {
        orderedEntries.push(entry);
        seenKeys.add(key);
      } else {
        failedLinks.add(link.originalUrl);
      }
    }

    if (orderedEntries.length === 0) {
      throw new Error('No entries could be resolved for download. Please verify the links in the Excel file.');
    }

    onProgress?.({ stage: 'zipping', message: 'Generating ZIP file...' });

    const zipBlob = await this.createZip(orderedEntries);
    const zipFileName = this.buildZipFileName();

    onProgress?.({ stage: 'complete', message: 'Excel import completed' });

    const summary: ExcelImportSummary = {
      totalLinks: rawUrls.length,
      validLinks: validLinkCount,
      processedEntries: orderedEntries.length,
      newEntriesSaved: totalNewEntriesSaved,
      failedLinks: Array.from(failedLinks),
      unmatchedLinks,
      siteSummaries
    };

    return { zipBlob, zipFileName, summary };
  }

  private static async readWorkbook(file: File): Promise<XLSX.WorkBook> {
    const data = await file.arrayBuffer();
    return XLSX.read(data, { type: 'array' });
  }

  private static extractUrlsFromWorkbook(workbook: XLSX.WorkBook): string[] {
    const urls: string[] = [];

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        return;
      }

      const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
      if (!range) {
        return;
      }

      const headerRow = range.s.r;
      const linkColumns: number[] = [];

      for (let col = range.s.c; col <= range.e.c; col++) {
        const headerCellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
        const headerCell = sheet[headerCellAddress] as XLSX.CellObject | undefined;
        const headerValue = this.getCellStringValue(headerCell);

        if (headerValue && headerValue.trim().toLowerCase() === 'link') {
          linkColumns.push(col);
        }
      }

      const columnsToScan = linkColumns.length > 0 ? linkColumns : [...Array(range.e.c - range.s.c + 1).keys()].map(
        index => range.s.c + index
      );

      for (let row = headerRow + (linkColumns.length > 0 ? 1 : 0); row <= range.e.r; row++) {
        for (const col of columnsToScan) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = sheet[cellAddress] as XLSX.CellObject | undefined;
          const url = this.getUrlFromCell(cell);

          if (url) {
            urls.push(url);
          }
        }
      }
    });

    return urls;
  }

  private static getUrlFromCell(cell: XLSX.CellObject | undefined): string | null {
    if (!cell) {
      return null;
    }

    if (cell.l?.Target && this.isValidUrl(cell.l.Target)) {
      return cell.l.Target.trim();
    }

    const value = this.getCellStringValue(cell);
    if (value && this.isValidUrl(value)) {
      return value;
    }

    return null;
  }

  private static getCellStringValue(cell: XLSX.CellObject | undefined): string | null {
    if (!cell) {
      return null;
    }

    if (typeof cell.v === 'string') {
      const trimmed = cell.v.trim();
      return trimmed || null;
    }

    return null;
  }

  private static prepareLinkData(rawUrls: string[], sites: Site[]) {
    const unmatchedSet = new Set<string>();
    const uniqueLinkMap = new Map<string, ParsedLink>();
    const siteData = new Map<string, SiteProcessingData>();
    let validLinkCount = 0;

    for (const urlString of rawUrls) {
      const parsedLink = this.parseLink(urlString, sites);

      if (parsedLink.site && parsedLink.guid) {
        validLinkCount += 1;
        const siteId = parsedLink.site.id;
        const key = `${siteId}|${parsedLink.guid}`;

        if (!uniqueLinkMap.has(key)) {
          uniqueLinkMap.set(key, parsedLink);
        }

        let info = siteData.get(siteId);
        if (!info) {
          info = {
            site: parsedLink.site,
            guids: [],
            urlByGuid: new Map(),
            totalLinks: 0
          };
          siteData.set(siteId, info);
        }

        info.totalLinks += 1;
        if (!info.urlByGuid.has(parsedLink.guid)) {
          info.guids.push(parsedLink.guid);
          info.urlByGuid.set(parsedLink.guid, parsedLink.originalUrl);
        }
      } else {
        unmatchedSet.add(urlString);
      }
    }

    const uniqueLinks = Array.from(uniqueLinkMap.values());

    return {
      uniqueLinks,
      unmatchedLinks: Array.from(unmatchedSet),
      siteData,
      validLinkCount
    };
  }

  private static parseLink(urlString: string, sites: Site[]): ParsedLink {
    try {
      const url = new URL(urlString);
      const site = this.matchSite(url, sites);
      const guid = this.extractGuid(url);

      return { originalUrl: urlString, site, guid };
    } catch {
      return { originalUrl: urlString, site: null, guid: null };
    }
  }

  private static matchSite(url: URL, sites: Site[]): Site | null {
    const hostname = this.normalizeHostname(url.hostname);

    for (const site of sites) {
      try {
        const siteHost = this.normalizeHostname(new URL(site.url).hostname);
        if (hostname === siteHost) {
          return site;
        }
      } catch {
        // Ignore invalid site URLs
      }
    }

    return null;
  }

  private static normalizeHostname(hostname: string): string {
    return hostname.replace(/^www\./i, '').toLowerCase();
  }

  private static extractGuid(url: URL): string | null {
    const guidParam = url.searchParams.get('guid');
    if (guidParam) {
      const trimmed = guidParam.trim();
      if (trimmed) {
        return trimmed;
      }
    }

    const pathSegments = url.pathname.split('/').filter(Boolean).reverse();
    for (const segment of pathSegments) {
      const cleaned = segment.trim();
      if (!cleaned) {
        continue;
      }

      if (this.isGuid(cleaned)) {
        return cleaned;
      }

      const match = cleaned.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
      if (match) {
        return match[0];
      }
    }

    return null;
  }

  private static isGuid(value: string): boolean {
    return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value);
  }

  private static isValidUrl(value: string): boolean {
    try {
      new URL(value);
      return true;
    } catch {
      return false;
    }
  }

  private static mapToEntry(data: any, siteUrl: string): Entry {
    const id = data.id || data.guid;
    if (!id) {
      throw new Error('Fetched entry did not contain an ID or GUID.');
    }

    return {
      id,
      title: data.title || '',
      abstract: data.abstract || '',
      body: data.body || '',
      publishedDate: data.published_date || data.date || '',
      type: data.type || 'publication',
      seen: false,
      siteUrl,
      metadata: data
    };
  }

  private static async createZip(entries: Entry[]): Promise<Blob> {
    const zip = new JSZip();

    entries.forEach(entry => {
      const fileName = `${this.sanitizeFileName(this.getSiteIdentifier(entry.siteUrl))}_${this.sanitizeFileName(entry.id)}.txt`;
      zip.file(fileName, this.formatEntryContent(entry));
    });

    return zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 1 }
    });
  }

  private static getSiteIdentifier(siteUrl?: string): string {
    if (!siteUrl) {
      return 'unknown_site';
    }

    try {
      const url = new URL(siteUrl);
      return url.hostname || 'unknown_site';
    } catch {
      return 'unknown_site';
    }
  }

  private static formatEntryContent(entry: Entry): string {
    const metadata = entry.metadata || {};

    return [
      `site_url: ${entry.siteUrl || ''}`,
      `id: ${entry.id || ''}`,
      `type: ${entry.type || ''}`,
      `jnr: ${metadata.jnr || ''}`,
      `title: ${entry.title || ''}`,
      `published_date: ${entry.publishedDate || ''}`,
      `date: ${metadata.date || entry.publishedDate || ''}`,
      `is_board_ruling: ${metadata.is_board_ruling || ''}`,
      `is_brought_to_court: ${metadata.is_brought_to_court || ''}`,
      `authority: ${metadata.authority || ''}`,
      `categories: ${Array.isArray(metadata.categories) ? metadata.categories.join(', ') : metadata.categories || ''}`,
      `original_url: ${metadata.url || ''}`,
      `abstract: ${entry.abstract || ''}`,
      `body: ${entry.body || ''}`
    ].join('\n');
  }

  private static sanitizeFileName(fileName: string): string {
    return fileName
      .replace(/[<>:"/\\|?*]/g, '_')
      .replace(/\s+/g, '_')
      .substring(0, 120);
  }

  private static buildZipFileName(): string {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    return `excel_import_${timestamp}.zip`;
  }
}
