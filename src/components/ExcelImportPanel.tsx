import React, { useRef } from 'react';
import { Upload, Loader2, FileWarning } from 'lucide-react';
import { ExcelImportProgress, ExcelImportSummary } from '../services/ExcelImportService';

interface ExcelImportPanelProps {
  isProcessing: boolean;
  onUpload: (file: File) => void;
  progress: ExcelImportProgress | null;
  summary: ExcelImportSummary | null;
  hasSites: boolean;
}

const ExcelImportPanel: React.FC<ExcelImportPanelProps> = ({
  isProcessing,
  onUpload,
  progress,
  summary,
  hasSites
}) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    onUpload(file);

    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const disabled = isProcessing || !hasSites;

  return (
    <div className="bg-white border-b border-gray-200">
      <div className="p-4">
        <div className="flex items-center justify-between mb-3">
          <div>
            <h2 className="text-lg font-semibold text-gray-900">Excel Import</h2>
            <p className="text-sm text-gray-600">
              Upload an Excel file with links to decisions and news. The entries will be fetched, saved and bundled in a ZIP file.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <label className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg border ${disabled ? 'border-gray-200 text-gray-400 cursor-not-allowed' : 'border-blue-500 text-blue-600 hover:bg-blue-50 cursor-pointer'} transition-colors`}>
            <Upload size={18} />
            <span className="text-sm font-medium">Choose Excel File</span>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={handleFileChange}
              disabled={disabled}
            />
          </label>

          {!hasSites && (
            <div className="flex items-center gap-2 text-sm text-orange-600">
              <FileWarning size={16} />
              <span>Add at least one site before importing.</span>
            </div>
          )}

          {isProcessing && (
            <div className="flex items-center gap-2 text-sm text-blue-600">
              <Loader2 size={16} className="animate-spin" />
              <span>Processing Excel file...</span>
            </div>
          )}
        </div>

        {progress && (
          <div className="mt-3 text-sm text-gray-700 bg-blue-50 border border-blue-100 rounded-md px-3 py-2">
            <div className="font-medium text-blue-900">{progress.message}</div>
            {typeof progress.current === 'number' && typeof progress.total === 'number' && progress.total > 0 && (
              <div className="text-xs text-blue-700 mt-1">
                {progress.current}/{progress.total} completed
              </div>
            )}
          </div>
        )}

        {summary && (
          <div className="mt-4 space-y-4">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <div className="p-3 bg-gray-50 rounded-lg">
                <div className="text-xs uppercase text-gray-500 tracking-wide">Total Links</div>
                <div className="text-lg font-semibold text-gray-900">{summary.totalLinks}</div>
              </div>
              <div className="p-3 bg-gray-50 rounded-lg">
                <div className="text-xs uppercase text-gray-500 tracking-wide">Valid Links</div>
                <div className="text-lg font-semibold text-gray-900">{summary.validLinks}</div>
              </div>
              <div className="p-3 bg-gray-50 rounded-lg">
                <div className="text-xs uppercase text-gray-500 tracking-wide">Entries in ZIP</div>
                <div className="text-lg font-semibold text-gray-900">{summary.processedEntries}</div>
              </div>
              <div className="p-3 bg-gray-50 rounded-lg">
                <div className="text-xs uppercase text-gray-500 tracking-wide">New Entries Saved</div>
                <div className="text-lg font-semibold text-gray-900">{summary.newEntriesSaved}</div>
              </div>
            </div>

            {summary.siteSummaries.length > 0 && (
              <div>
                <h3 className="text-sm font-semibold text-gray-800 mb-2">Per Site Summary</h3>
                <div className="space-y-2">
                  {summary.siteSummaries.map(site => (
                    <div key={site.siteId} className="border border-gray-200 rounded-lg p-3">
                      <div className="flex flex-wrap justify-between gap-2">
                        <div>
                          <div className="text-sm font-semibold text-gray-900">{site.siteName}</div>
                          <div className="text-xs text-gray-500">{site.uniqueEntries} unique entries from {site.totalLinks} links</div>
                        </div>
                        <div className="flex flex-wrap gap-3 text-sm">
                          <span className="text-green-700 font-medium">+{site.newEntriesSaved} saved</span>
                          <span className="text-gray-600">{site.existingEntries} existing</span>
                          {site.failedEntries > 0 && (
                            <span className="text-red-600">{site.failedEntries} failed</span>
                          )}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {(summary.failedLinks.length > 0 || summary.unmatchedLinks.length > 0) && (
              <div className="bg-red-50 border border-red-100 rounded-lg p-3">
                <div className="flex items-center gap-2 text-sm font-semibold text-red-700 mb-1">
                  <FileWarning size={16} />
                  Issues Detected
                </div>
                {summary.unmatchedLinks.length > 0 && (
                  <div className="text-sm text-red-700 mb-2">
                    <span className="font-medium">Unmatched Links:</span> {summary.unmatchedLinks.slice(0, 5).join(', ')}
                    {summary.unmatchedLinks.length > 5 && ' ...'}
                  </div>
                )}
                {summary.failedLinks.length > 0 && (
                  <div className="text-sm text-red-700">
                    <span className="font-medium">Failed to Fetch:</span> {summary.failedLinks.slice(0, 5).join(', ')}
                    {summary.failedLinks.length > 5 && ' ...'}
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default ExcelImportPanel;
