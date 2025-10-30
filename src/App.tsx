import React, { useState, useEffect, useRef } from 'react';
import { v4 as uuidv4 } from 'uuid';
import { Database, HardDrive, Wifi, WifiOff, X } from 'lucide-react';
import { saveAs } from 'file-saver';
import { Site, Entry, LogEntry } from './types';
import { StorageService } from './services/StorageService';
import { ExportService } from './services/ExportService';
import { SyncService } from './services/SyncService';
import SiteManager from './components/SiteManager';
import EntryList from './components/EntryList';
import PreviewPanel from './components/PreviewPanel';
import LogPanel from './components/LogPanel';
import SyncSplashScreen from './components/SyncSplashScreen';
import ExportSplashScreen from './components/ExportSplashScreen';
import { useSiteOperations } from './hooks/useSiteOperations';
import { useEntryOperations } from './hooks/useEntryOperations';
import { useAutoLoad } from './hooks/useAutoLoad';
import { usePersistentOperations } from './hooks/usePersistentOperations';
import BackgroundExportModal from './components/BackgroundExportModal';
import ExcelImportPanel from './components/ExcelImportPanel';
import { ExcelImportService, ExcelImportProgress, ExcelImportSummary } from './services/ExcelImportService';

function App() {
  const [isInitializing, setIsInitializing] = useState(true);
  const [sites, setSites] = useState<Site[]>([]);
  const [selectedSite, setSelectedSite] = useState<Site | null>(null);
  const [entries, setEntries] = useState<Entry[]>([]);
  const [selectedEntry, setSelectedEntry] = useState<Entry | null>(null);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [maxEntries, setMaxEntries] = useState(0);
  const [newEntriesCount, setNewEntriesCount] = useState<{ [siteId: string]: number }>({});
  const [storageUsage, setStorageUsage] = useState(StorageService.getStorageUsage());
  const [showSites, setShowSites] = useState(false);
  const [syncProgress, setSyncProgress] = useState<{
    isVisible: boolean;
    step: string;
    currentSite: string;
    sitesProcessed: number;
    totalSites: number;
    entriesProcessed: number;
    totalEntries: number;
    isComplete: boolean;
  }>({
    isVisible: false,
    step: '',
    currentSite: '',
    sitesProcessed: 0,
    totalSites: 0,
    entriesProcessed: 0,
    totalEntries: 0,
    isComplete: false
  });
  const [exportProgress, setExportProgress] = useState<{
    isVisible: boolean;
    step: string;
    currentSite: string;
    sitesProcessed: number;
    totalSites: number;
    entriesProcessed: number;
    totalEntries: number;
    isComplete: boolean;
  }>({
    isVisible: false,
    step: '',
    currentSite: '',
    sitesProcessed: 0,
    totalSites: 0,
    entriesProcessed: 0,
    totalEntries: 0,
    isComplete: false
  });
  const [showBackgroundExportModal, setShowBackgroundExportModal] = useState(false);
  const abortControllerRef = useRef<AbortController | null>(null);

  const [isExcelProcessing, setIsExcelProcessing] = useState(false);
  const [excelProgress, setExcelProgress] = useState<ExcelImportProgress | null>(null);
  const [excelSummary, setExcelSummary] = useState<ExcelImportSummary | null>(null);

  // Add state for persistent fetch progress display
  const [persistentFetchProgress, setPersistentFetchProgress] = useState<{
    isVisible: boolean;
    siteName: string;
    step: string;
    current: number;
    total: number;
    message: string;
  }>({
    isVisible: false,
    siteName: '',
    step: '',
    current: 0,
    total: 0,
    message: ''
  });

  const addLog = (message: string, type: 'info' | 'success' | 'error' | 'warning' = 'info') => {
    const log: LogEntry = {
      id: uuidv4(),
      message,
      type,
      timestamp: new Date()
    };
    setLogs(prev => [...prev, log]);
  };

  const stopCurrentOperation = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      abortControllerRef.current = null;
      setIsLoading(false);
      addLog('Operation stopped by user', 'warning');
    }
  };

  const loadSites = async () => {
    try {
      const loadedSites = await StorageService.loadSites();
      setSites(loadedSites);
      addLog(`Loaded ${loadedSites.length} sites from ${StorageService.getStorageType()}`, 'success');
      return loadedSites;
    } catch (error) {
      addLog(`Failed to load sites: ${error.message}`, 'error');
      return [];
    }
  };

  const loadEntries = async (siteUrl: string) => {
    try {
      const loadedEntries = await StorageService.loadEntries(siteUrl);
      setEntries(loadedEntries);
      setSelectedEntry(null);
      addLog(`Loaded ${loadedEntries.length} entries for display`, 'info');
    } catch (error) {
      addLog(`Failed to load entries: ${error.message}`, 'error');
    }
  };

  // Initialize hooks
  const { autoLoadStatus, startAutoLoad } = useAutoLoad(addLog, setShowSites);
  
  const { activeOperations, startPersistentFetch, stopOperation, cancelOperation, verifyingSiteIds } = usePersistentOperations({
    sites,
    setSites,
    addLog,
    maxEntries,
    setPersistentFetchProgress
  });
  
  const siteOperations = useSiteOperations({
    sites,
    setSites,
    selectedSite,
    setSelectedSite,
    addLog,
    isLoading,
    setIsLoading,
    abortControllerRef,
    newEntriesCount,
    setNewEntriesCount,
    loadEntries,
    maxEntries
  });

  const entryOperations = useEntryOperations({
    entries,
    setEntries,
    selectedEntry,
    setSelectedEntry,
    selectedSite,
    sites,
    setSites,
    addLog
  });

  useEffect(() => {
    const initializeApp = async () => {
      setIsInitializing(true);
      
      // Make addLog available globally for services
      (window as any).addLog = addLog;
      
      // Load sites first
      const loadedSites = await loadSites();
      
      // Show sites immediately
      setShowSites(true);
      
      // End initialization phase quickly
      setIsInitializing(false);
      
      // Start background auto-load to update site counts
      startAutoLoad(loadedSites, setSelectedSite, setEntries, setSites);
    };
    
    initializeApp();
    
    const interval = setInterval(() => {
      setStorageUsage(StorageService.getStorageUsage());
    }, 30000);

    return () => clearInterval(interval);
  }, []);

  // Cleanup global addLog on unmount
  useEffect(() => {
    return () => {
      delete (window as any).addLog;
    };
  }, []);

  useEffect(() => {
    if (selectedSite) {
      // Don't auto-load entries - let EntryList handle loading based on active tab
    }
  }, [selectedSite]);

  // Filter new entries to only include those from the last 24 hours
  const getRecentNewEntries = (entries: Entry[]) => {
    const oneDayAgo = new Date();
    oneDayAgo.setDate(oneDayAgo.getDate() - 1);
    
    return entries.filter(entry => {
      if (entry.seen) return false;
      
      // Check when entry was stored in database (created_at), fall back to published_date only if no created_at
      const entryDate = entry.metadata?.created_at ? 
        new Date(entry.metadata.created_at) : 
        (entry.publishedDate ? new Date(entry.publishedDate) : new Date(0));
      
      return entryDate >= oneDayAgo;
    });
  };

  const recentNewEntries = getRecentNewEntries(entries);
  const handleExportEntries = async (site: Site) => {
    setShowBackgroundExportModal(true);
  };

  const handleExportAllSites = async () => {
    setShowBackgroundExportModal(true);
  };

  const handleSyncToFolder = async () => {
    setShowBackgroundExportModal(true);
  };

  const handleStartBackgroundExport = async (type: 'single_site' | 'all_sites', siteId?: string) => {
    // Export is handled directly in the modal component
  };

  const handleExcelImport = async (file: File) => {
    if (sites.length === 0) {
      addLog('Cannot import Excel file without configured sites', 'warning');
      return;
    }

    setIsExcelProcessing(true);
    setExcelProgress(null);
    setExcelSummary(null);
    addLog(`Starting Excel import for ${file.name}`, 'info');

    try {
      const result = await ExcelImportService.processFile(file, sites, progress => {
        setExcelProgress(progress);
      });

      saveAs(result.zipBlob, result.zipFileName);
      setExcelSummary(result.summary);
      addLog(
        `Excel import complete: ${result.summary.processedEntries} entries bundled, ${result.summary.newEntriesSaved} new saved`,
        'success'
      );

      const processedSiteIds = new Set(result.summary.siteSummaries.map(site => site.siteId));

      if (processedSiteIds.size > 0) {
        const updatedSites = await Promise.all(sites.map(async site => {
          if (!processedSiteIds.has(site.id)) {
            return site;
          }

          try {
            const entryCount = await StorageService.getActualEntryCount(site.url);
            return {
              ...site,
              entryCount,
              lastUpdated: new Date()
            };
          } catch (error) {
            addLog(`Failed to refresh entry count for ${site.name}: ${error.message}`, 'warning');
            return site;
          }
        }));

        setSites(updatedSites);

        try {
          await StorageService.saveSites(updatedSites);
        } catch (error) {
          addLog(`Unable to persist site updates after Excel import: ${error.message}`, 'warning');
        }

        setNewEntriesCount(prev => {
          const next = { ...prev };
          processedSiteIds.forEach(id => {
            next[id] = 0;
          });
          return next;
        });

        if (selectedSite && processedSiteIds.has(selectedSite.id)) {
          await loadEntries(selectedSite.url);
        }
      }
    } catch (error) {
      addLog(`Excel import failed: ${error.message}`, 'error');
    } finally {
      setIsExcelProcessing(false);
      setExcelProgress(null);
    }
  };

  const storageType = StorageService.getStorageType();

  // Show loading screen during initialization
  if (isInitializing) {
    return (
      <div className="min-h-screen bg-gray-100 flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-16 w-16 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <h2 className="text-xl font-semibold text-gray-900 mb-2">Loading Knowledge Updater</h2>
          <p className="text-gray-600">Loading sites and checking for active operations...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col">
      {/* Header */}
      <header className="bg-white shadow-sm border-b border-gray-200">
        <div className="px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <Database className="w-8 h-8 text-blue-600" />
              <div>
                <h1 className="text-2xl font-bold text-gray-900">Knowledge Updater</h1>
                <p className="text-sm text-gray-600">Publication tracking and management system</p>
              </div>
            </div>
            
            <div className="flex items-center gap-4">
              <div className="flex items-center gap-2 px-3 py-2 bg-gray-50 rounded-lg">
                {storageType === 'database' ? (
                  <>
                    <Wifi className="w-4 h-4 text-green-600" />
                    <span className="text-sm font-medium text-green-700">Database</span>
                  </>
                ) : (
                  <>
                    <WifiOff className="w-4 h-4 text-orange-600" />
                    <span className="text-sm font-medium text-orange-700">Local Storage</span>
                  </>
                )}
              </div>

              {storageType === 'local' && (
                <div className="flex items-center gap-2 px-3 py-2 bg-gray-50 rounded-lg">
                  <HardDrive className="w-4 h-4 text-gray-600" />
                  <span className="text-sm text-gray-600">
                    {storageUsage.percentage.toFixed(1)}% used
                  </span>
                </div>
              )}

              <div className="text-right">
                <div className="text-sm font-medium text-gray-900">{sites.length} Sites</div>
                <div className="text-xs text-gray-500">{recentNewEntries.length} New Entries (24h)</div>
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* Main content */}
      <div className="flex-1 flex overflow-hidden">
        <div className="w-80 bg-white border-r border-gray-200 flex flex-col">
          <SiteManager
            sites={sites}
            selectedSite={selectedSite}
            onSelectSite={setSelectedSite}
            onAddSite={siteOperations.handleAddSite}
            onEditSite={siteOperations.handleEditSite}
            onRemoveSite={siteOperations.handleRemoveSite}
            onUpdateSitemap={siteOperations.handleUpdateSitemap}
            onFetchEntries={siteOperations.handleFetchEntries}
            onDeleteAllEntries={siteOperations.handleDeleteAllEntries}
            isLoading={isLoading}
            maxEntries={maxEntries}
            onMaxEntriesChange={setMaxEntries}
            onExportEntries={handleExportEntries}
            onUpdateAllSitemaps={siteOperations.handleUpdateAllSitemaps}
            onFetchAllEntries={siteOperations.handleFetchAllEntries}
            onExportAllSites={handleExportAllSites}
            onDeleteAllEntriesAllSites={siteOperations.handleDeleteAllEntriesAllSites}
            onSyncToFolder={handleSyncToFolder}
            onStopOperation={stopCurrentOperation}
            newEntriesCount={newEntriesCount}
            autoLoadStatus={autoLoadStatus}
          fetchStatus={siteOperations.fetchStatus}
          showSites={showSites}
          activeOperations={activeOperations}
          verifyingSiteIds={verifyingSiteIds}
          onStartPersistentFetch={startPersistentFetch}
          onStopPersistentOperation={stopOperation}
          onCancelPersistentOperation={cancelOperation}
        />
          
          <EntryList
            entries={entries}
            newEntries={recentNewEntries}
            selectedEntry={selectedEntry}
            selectedSite={selectedSite}
            onSelectEntry={setSelectedEntry}
            onMarkAsSeen={entryOperations.handleMarkAsSeen}
            onDeleteEntry={entryOperations.handleDeleteEntry}
          />
        </div>

        <div className="flex-1 flex flex-col">
          <ExcelImportPanel
            isProcessing={isExcelProcessing}
            onUpload={handleExcelImport}
            progress={excelProgress}
            summary={excelSummary}
            hasSites={sites.length > 0}
          />
          <PreviewPanel
            entry={selectedEntry}
            onMarkAsSeen={entryOperations.handleMarkAsSeen}
            onDeleteEntry={entryOperations.handleDeleteEntry}
          />
          
          <LogPanel logs={logs} />
        </div>
      </div>
      
      {/* Sync Splash Screen */}
      <SyncSplashScreen
        isVisible={syncProgress.isVisible}
        progress={syncProgress}
      />
      
      {/* Export Splash Screen */}
      <ExportSplashScreen
        isVisible={exportProgress.isVisible}
        progress={exportProgress}
      />
      
      {/* Background Export Modal */}
      <BackgroundExportModal
        isOpen={showBackgroundExportModal}
        onClose={() => setShowBackgroundExportModal(false)}
        onStartExport={handleStartBackgroundExport}
        sites={sites}
        selectedSite={selectedSite}
      />
      
      {/* Persistent Fetch Progress Display */}
      {persistentFetchProgress.isVisible && (
        <div className="fixed bottom-4 right-4 bg-white rounded-lg shadow-lg border border-gray-200 p-4 max-w-sm z-40">
          <div className="flex items-center justify-between mb-2">
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 bg-blue-600 rounded-full animate-pulse"></div>
              <span className="text-sm font-medium text-gray-900">Persistent Fetch</span>
            </div>
            <button
              onClick={() => setPersistentFetchProgress(prev => ({ ...prev, isVisible: false }))}
              className="text-gray-400 hover:text-gray-600"
            >
              <X size={16} />
            </button>
          </div>
          
          <div className="text-sm text-gray-700 mb-2">
            <strong>{persistentFetchProgress.siteName}</strong>
          </div>
          
          <div className="text-sm text-gray-600 mb-2">
            {persistentFetchProgress.message}
          </div>
          
          {persistentFetchProgress.total > 0 && (
            <div className="space-y-1">
              <div className="flex justify-between text-xs text-gray-600">
                <span>{persistentFetchProgress.current} / {persistentFetchProgress.total}</span>
                <span>{Math.round((persistentFetchProgress.current / persistentFetchProgress.total) * 100)}%</span>
              </div>
              <div className="w-full bg-gray-200 rounded-full h-2">
                <div 
                  className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                  style={{ 
                    width: `${Math.min((persistentFetchProgress.current / persistentFetchProgress.total) * 100, 100)}%` 
                  }}
                />
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export default App;