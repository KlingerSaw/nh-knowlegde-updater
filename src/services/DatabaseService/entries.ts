import { SupabaseClient } from '@supabase/supabase-js';
import { Entry } from '../../types';
import { getTableName, DEFAULT_QUERY_OPTIONS, BATCH_CONFIG } from './constants';

export class EntriesRepository {
  constructor(private supabase: SupabaseClient) {}

  async getEntryCount(siteId: string, siteName: string): Promise<number> {
    try {
      const tableName = getTableName('entries');
      
      console.log(`📊 Querying entry count for site ${siteName} (ID: ${siteId}) from table ${tableName}`);
      
      const timeoutPromise = new Promise<never>((_, reject) => {
        setTimeout(() => reject(new Error('Entry count query timeout')), 15000); // Shorter timeout for count queries
      });
      
      const queryPromise = this.supabase
        .from(tableName)
        .select('id', { count: 'exact', head: true })
        .eq('site_id', siteId);
      
      const { count, error } = await Promise.race([queryPromise, timeoutPromise]);

      if (error) {
        // Handle "Failed to fetch" errors as warnings since they're network-related
        if (error instanceof TypeError && error.message === 'Failed to fetch') {
          console.warn(`📊 Network error getting entry count for ${siteName}:`, error);
          
          const addLog = (window as any).addLog;
          if (addLog) {
            addLog(`⚠️ Network error getting entry count for ${siteName}: ${error.message}`, 'warning');
          }
        } else {
          console.error(`📊 Error getting entry count for ${siteName}:`, error);
          
          // Log the specific error details for debugging
          const addLog = (window as any).addLog;
          if (addLog) {
            addLog(`❌ Database error getting entry count for ${siteName}: ${error.message}`, 'error');
          }
        }
        
        // Return 0 instead of throwing to prevent app crashes
        console.warn(`📊 Returning 0 entries due to database error for ${siteName}`);
        return 0;
      }

      console.log(`📊 Entry count query result for ${siteName}: ${count || 0}`);
      return count || 0;
    } catch (error) {
      // Handle "Failed to fetch" errors as warnings
      if (error instanceof TypeError && error.message === 'Failed to fetch') {
        console.warn('📊 Network error getting entry count:', error);
        
        const addLog = (window as any).addLog;
        if (addLog) {
          addLog(`⚠️ Network error getting entry count for ${siteName}: ${error.message}`, 'warning');
        }
        
        // Return 0 instead of throwing to prevent app crashes
        console.warn(`📊 Returning 0 entries due to network error for ${siteName}`);
        return 0;
      }
      
      if (error.message === 'Entry count query timeout') {
        console.error('📊 Entry count query timed out after 8 seconds');
        
        const addLog = (window as any).addLog;
        if (addLog) {
          addLog(`⏱️ Database timeout getting entry count for ${siteName}`, 'warning');
        }
        
        // Return 0 instead of throwing to prevent app crashes
        console.warn(`📊 Returning 0 entries due to timeout for ${siteName}`);
        return 0;
      }
      console.error('Failed to get entry count from database:', error);
      
      const addLog = (window as any).addLog;
      if (addLog) {
        addLog(`❌ Failed to get entry count for ${siteName}: ${error.message}`, 'error');
      }
      
      // Return 0 instead of throwing to prevent app crashes
      console.warn(`📊 Returning 0 entries due to error for ${siteName}`);
      return 0;
    }
  }

  async getUnseenEntryCount(siteId: string, siteName: string, recentOnly: boolean = false): Promise<number> {
    try {
      const tableName = getTableName('entries');
      
      console.log(`📊 Querying unseen entry count for site ${siteName} (ID: ${siteId}) from table ${tableName}`);
      
      const timeoutPromise = new Promise<never>((_, reject) => {
        setTimeout(() => reject(new Error('Unseen entry count query timeout')), 15000); // Shorter timeout for count queries
      });
      
      let query = this.supabase
        .from(tableName)
        .select('*', { count: 'exact', head: true })
        .eq('site_id', siteId)
        .is('seen', false); // Use 'is' instead of 'eq' for boolean false
      
      // If recentOnly is true, only count entries stored in database within the last 24 hours
      if (recentOnly) {
        const oneDayAgo = new Date();
        oneDayAgo.setDate(oneDayAgo.getDate() - 1);
        query = query.gte('created_at', oneDayAgo.toISOString());
      }
      
      const { count, error } = await Promise.race([query, timeoutPromise]);

      if (error) {
        // Handle "Failed to fetch" errors as warnings since they're network-related
        if (error instanceof TypeError && error.message === 'Failed to fetch') {
          console.warn(`📊 Network error getting unseen entry count for ${siteName}:`, error);
          
          const addLog = (window as any).addLog;
          if (addLog) {
            addLog(`⚠️ Network error getting unseen count for ${siteName}: ${error.message}`, 'warning');
          }
        } else {
          console.error(`📊 Error getting unseen entry count for ${siteName}:`, error);
          
          // Log the specific error details for debugging
          const addLog = (window as any).addLog;
          if (addLog) {
            addLog(`❌ Database error getting unseen count for ${siteName}: ${error.message}`, 'error');
          }
        }
        
        // Return 0 instead of throwing to prevent app crashes
        console.warn(`📊 Returning 0 unseen entries due to database error for ${siteName}`);
        return 0;
      }

      const resultText = recentOnly ? `${count || 0} (recent only)` : `${count || 0}`;
      console.log(`📊 Unseen entry count query result for ${siteName}: ${resultText}`);
      return count || 0;
    } catch (error) {
      // Handle "Failed to fetch" errors as warnings
      if (error instanceof TypeError && error.message === 'Failed to fetch') {
        console.warn('📊 Network error getting unseen entry count:', error);
        
        const addLog = (window as any).addLog;
        if (addLog) {
          addLog(`⚠️ Network error getting unseen count for ${siteName}: ${error.message}`, 'warning');
        }
        
        // Return 0 instead of throwing to prevent app crashes
        console.warn(`📊 Returning 0 unseen entries due to network error for ${siteName}`);
        return 0;
      }
      
      if (error.message === 'Unseen entry count query timeout') {
        console.error('📊 Unseen entry count query timed out after 8 seconds');
        
        const addLog = (window as any).addLog;
        if (addLog) {
          addLog(`⏱️ Database timeout getting unseen count for ${siteName}`, 'warning');
        }
        
        // Return 0 instead of throwing to prevent app crashes
        console.warn(`📊 Returning 0 unseen entries due to timeout for ${siteName}`);
        return 0;
      }
      console.error('Failed to get unseen entry count from database:', error);
      
      const addLog = (window as any).addLog;
      if (addLog) {
        addLog(`❌ Failed to get unseen count for ${siteName}: ${error.message}`, 'error');
      }
      
      // Return 0 instead of throwing to prevent app crashes
      console.warn(`📊 Returning 0 unseen entries due to error for ${siteName}`);
      return 0;
    }
  }

  async getExistingEntryIds(siteId: string, guids: string[]): Promise<string[]> {
    if (guids.length === 0) {
      return [];
    }

    try {
      const tableName = getTableName('entries');
      
      // Process in smaller batches to avoid query size limits and timeouts
      const batchSize = 500;
      const existingIds: string[] = [];
      
      for (let i = 0; i < guids.length; i += batchSize) {
        const batch = guids.slice(i, i + batchSize);
        
        // Add timeout to each batch query
        const timeoutPromise = new Promise<never>((_, reject) => {
          setTimeout(() => reject(new Error('Batch query timeout')), 30000);
        });
        
        const queryPromise = this.supabase
          .from(tableName)
          .select('id')
          .eq('site_id', siteId)
          .in('id', batch);
        
        const { data, error } = await Promise.race([queryPromise, timeoutPromise]);

        if (error) {
          throw error;
        }

        if (data) {
          existingIds.push(...data.map(row => row.id));
        }
        
        // Small delay between batches to avoid overwhelming the database
        if (i + batchSize < guids.length) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      }
      
      return existingIds;
    } catch (error) {
      if (error.message === 'Batch query timeout') {
        console.error('Batch query timed out - returning partial results');
        throw new Error('Database query timeout - operation may be incomplete');
      }
      console.error('Failed to get existing entry IDs from database:', error);
      throw error;
    }
  }

  async loadEntriesWithLimit(siteId: string, siteName: string, limit?: number, offset?: number): Promise<Entry[]> {
    try {
      const tableName = getTableName('entries');

      // Always use batching for reliability, even with limits
      if (!limit) {
        const batchSize = 50; // Very small batch size for maximum reliability
        let allEntries: any[] = [];
        let currentOffset = offset || 0;
        
        while (true) {
          // Add timeout to each batch
          const timeoutPromise = new Promise<never>((_, reject) => {
            setTimeout(() => reject(new Error('Load entries timeout')), 20000); // 20 second timeout per batch
          });
          
          const queryPromise = this.supabase
            .from(tableName)
            .select('*')
            .eq('site_id', siteId)
            .order('published_date', { ascending: false })
            .range(currentOffset, currentOffset + batchSize - 1);
          
          const { data, error } = await Promise.race([queryPromise, timeoutPromise]);

          if (error) {
            throw error;
          }

          if (!data || data.length === 0) {
            break;
          }

          allEntries.push(...data);
          console.log(`Loaded batch: ${data.length} entries (total so far: ${allEntries.length})`);
          
          // If we got less than the batch size, we've reached the end
          if (data.length < batchSize) {
            break;
          }
          
          currentOffset += batchSize;
          
          // Even longer delay between batches for stability
          await new Promise(resolve => setTimeout(resolve, 1000)); // 1 second delay
        }

        return this.mapEntriesToDomain(allEntries);
      } else {
        // For limited queries, also use chunking if the limit is large
        if (limit > 100) { // Lower threshold for chunking
          // Use chunking for large limits
          const chunkSize = 50; // Smaller chunks
          let allEntries: any[] = [];
          let currentOffset = offset || 0;
          let remaining = limit;
          
          while (remaining > 0) {
            const currentLimit = Math.min(chunkSize, remaining);
            
            const timeoutPromise = new Promise<never>((_, reject) => {
              setTimeout(() => reject(new Error('Chunked query timeout')), 20000); // 20 second timeout
            });
            
            const queryPromise = this.supabase
              .from(tableName)
              .select('*')
              .eq('site_id', siteId)
              .order('published_date', { ascending: false })
              .range(currentOffset, currentOffset + currentLimit - 1);
            
            const { data, error } = await Promise.race([queryPromise, timeoutPromise]);

            if (error) {
              throw error;
            }

            if (!data || data.length === 0) {
              break;
            }

            allEntries.push(...data);
            console.log(`Loaded chunk: ${data.length} entries (${allEntries.length}/${limit} total)`);
            
            currentOffset += data.length;
            remaining -= data.length;
            
            // If we got fewer entries than requested, we've reached the end
            if (data.length < currentLimit) {
              break;
            }
            
            // Delay between chunks
            await new Promise(resolve => setTimeout(resolve, 800)); // Longer delay
          }
          
          return this.mapEntriesToDomain(allEntries);
        } else {
          // For small limits, use direct query
        const timeoutPromise = new Promise<never>((_, reject) => {
            setTimeout(() => reject(new Error('Small query timeout')), 15000); // 15 second timeout for small queries
        });
        
        let query = this.supabase
          .from(tableName)
          .select('*')
          .eq('site_id', siteId)
          .order('published_date', { ascending: false });

        if (offset) {
          query = query.range(offset, offset + limit - 1);
        } else {
          query = query.limit(limit);
        }

        const { data, error } = await Promise.race([query, timeoutPromise]);

        if (error) {
          throw error;
        }

          console.log(`Small query completed: ${(data || []).length} entries (limit: ${limit}, offset: ${offset || 0})`);
        return this.mapEntriesToDomain(data || []);
        }
      }

    } catch (error) {
      if (error.message && error.message.includes('timeout')) {
        console.error('Database query timed out:', error);
        throw new Error(`Database query timeout - dataset too large for single query (limit: ${limit}, offset: ${offset})`);
      }
      console.error('Failed to load entries from database:', error);
      throw error;
    }
  }

  async loadUnseenEntriesWithLimit(siteId: string, siteName: string, limit: number, offset: number, recentOnly: boolean = false): Promise<Entry[]> {
    try {
      const tableName = getTableName('entries');
      
      console.log(`📊 Loading ${limit} unseen entries for ${siteName} (offset: ${offset})`);
      
      // First, let's check how many unseen entries exist
      console.log(`📊 Checking unseen entries count before loading...`);
      const timeoutPromise = new Promise<never>((_, reject) => {
        setTimeout(() => reject(new Error('Unseen entries query timeout')), 30000);
      });
      
      let query = this.supabase
        .from(tableName)
        .select('*')
        .eq('site_id', siteId)
        .is('seen', false) // Use 'is' instead of 'eq' for boolean false
        .order('published_date', { ascending: false })
        .range(offset, offset + limit - 1);
      
      // If recentOnly is true, only load entries stored in database within the last 24 hours
      if (recentOnly) {
        const oneDayAgo = new Date();
        oneDayAgo.setDate(oneDayAgo.getDate() - 1);
        query = query.gte('created_at', oneDayAgo.toISOString());
      }
      
      const { data, error } = await Promise.race([query, timeoutPromise]);

      if (error) {
        throw error;
      }
      
      // Also get a count to debug
      try {
        let countQuery = this.supabase
          .from(tableName)
          .select('*', { count: 'exact', head: true })
          .eq('site_id', siteId)
          .is('seen', false);
        
        if (recentOnly) {
          const oneDayAgo = new Date();
          oneDayAgo.setDate(oneDayAgo.getDate() - 1);
          countQuery = countQuery.gte('created_at', oneDayAgo.toISOString());
        }
        
        const { count } = await countQuery;
        const countText = recentOnly ? `${count || 0} (recent only)` : `${count || 0}`;
        console.log(`📊 Total unseen entries available: ${countText}`);
      } catch (countError) {
        console.warn('Failed to get unseen count for debugging:', countError);
      }

      console.log(`📊 Loaded ${(data || []).length} unseen entries for ${siteName}`);
      return this.mapEntriesToDomain(data || []);
    } catch (error) {
      if (error.message && error.message.includes('timeout')) {
        console.error('Database query timed out:', error);
        throw new Error('Database query timeout - please try again');
      }
      console.error('Failed to load unseen entries from database:', error);
      throw error;
    }
  }

  async loadEntriesByIds(siteId: string, siteName: string, ids: string[]): Promise<Entry[]> {
    if (ids.length === 0) {
      return [];
    }

    try {
      const tableName = getTableName('entries');
      const uniqueIds = Array.from(new Set(ids));
      const batchSize = 50;
      const allEntries: any[] = [];

      for (let i = 0; i < uniqueIds.length; i += batchSize) {
        const chunk = uniqueIds.slice(i, i + batchSize);

        const timeoutPromise = new Promise<never>((_, reject) => {
          setTimeout(() => reject(new Error('Load entries by IDs timeout')), 20000);
        });

        const queryPromise = this.supabase
          .from(tableName)
          .select('*')
          .eq('site_id', siteId)
          .in('id', chunk);

        const { data, error } = await Promise.race([queryPromise, timeoutPromise]);

        if (error) {
          throw error;
        }

        if (data && data.length > 0) {
          allEntries.push(...data);
        }

        if (i + batchSize < uniqueIds.length) {
          await new Promise(resolve => setTimeout(resolve, 200));
        }
      }

      return this.mapEntriesToDomain(allEntries);
    } catch (error) {
      if (error.message && error.message.includes('timeout')) {
        console.error(`Database query timed out while loading entries for ${siteName}:`, error);
        throw new Error('Database query timeout while loading entries by IDs');
      }
      console.error('Failed to load entries by IDs from database:', error);
      throw error;
    }
  }

  private mapEntriesToDomain(data: any[]): Entry[] {
    return data.map(entry => ({
      id: entry.id,
      title: entry.title || '',
      abstract: entry.abstract || '',
      body: entry.body || '',
      publishedDate: entry.published_date || '',
      type: entry.type || 'publication',
      seen: entry.seen || false,
      metadata: entry.metadata || {},
      siteUrl: '' // Will be set by the calling service
    }));
  }
}