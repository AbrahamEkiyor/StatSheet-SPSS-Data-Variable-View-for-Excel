import { CheckCircle2, Download, FileCode2, ShieldCheck, X } from 'lucide-react';
import { useState } from 'react';
import { generateOfficeManifestXml } from '../engine/spssEngine';
import type { HostEnvironment } from '../types/spss';

interface ExcelAddinModalProps {
  host: HostEnvironment;
  lastBackupTimestamp: string | null;
  onClose: () => void;
}

type PlatformGuide = 'windows' | 'mac' | 'web';

export function ExcelAddinModal({
  host,
  lastBackupTimestamp,
  onClose,
}: ExcelAddinModalProps) {
  const defaultOrigin =
    typeof window !== 'undefined'
      ? window.location.origin
      : 'https://localhost:3000';

  const [manifestUrl, setManifestUrl] = useState(defaultOrigin);
  const [platform, setPlatform] = useState<PlatformGuide>('windows');

  const handleDownloadManifest = () => {
    const clean = manifestUrl.trim() || defaultOrigin;
    const xml = generateOfficeManifestXml(clean);
    const blob = new Blob([xml], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'statsheet-excel-manifest.xml';
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="addin-modal-title"
    >
      <div className="bg-white border border-slate-300 rounded-lg shadow-xl w-full max-w-2xl flex flex-col max-h-[90vh]">
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2
              id="addin-modal-title"
              className="font-display text-sm font-semibold text-slate-900"
            >
              Excel Desktop &amp; Web Add-in Setup
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Works on Excel Desktop (Windows &amp; macOS) as well as Excel on the Web.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close modal"
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-5 overflow-y-auto space-y-4 text-xs text-slate-700">
          {/* Host Status */}
          <div className="p-3.5 border border-slate-200 rounded-md bg-slate-50 flex items-start gap-3">
            <CheckCircle2
              className={`w-4 h-4 mt-0.5 shrink-0 ${
                host === 'excel' ? 'text-[#107C41]' : 'text-blue-700'
              }`}
            />
            <div>
              <div className="font-semibold text-slate-900">
                {host === 'excel'
                  ? 'Connected to Microsoft Excel (Office.js Active)'
                  : 'Running in Browser Simulator Mode'}
              </div>
              <p className="text-slate-600 mt-0.5 leading-relaxed">
                {host === 'excel'
                  ? 'StatSheet is running inside Excel. Use Sync to Excel to write Value Labels and Data Validation rules while preserving raw numeric codes in hidden backup sheets.'
                  : 'You are previewing the interactive simulator. Download the manifest below to install StatSheet as a native Task Pane add-in inside your desktop Excel.'}
              </p>
            </div>
          </div>

          {/* Reversible Sync Explanation */}
          <div className="p-3.5 border border-slate-200 rounded-md space-y-1.5">
            <div className="flex items-center gap-2 font-semibold text-slate-900">
              <ShieldCheck className="w-4 h-4 text-[#107C41]" />
              <span>Reversible Value Labels Sync in Excel</span>
            </div>
            <p className="text-slate-600 leading-relaxed">
              When syncing Value Labels to your active Excel sheet, StatSheet creates two hidden worksheets inside your workbook:
            </p>
            <ul className="list-disc pl-5 space-y-1 text-slate-600">
              <li>
                <code className="font-mono text-slate-800">_SPSS_RawBackup</code> — Stores exact raw numeric codes so toggling Value Labels off restores numbers without breaking formulas.
              </li>
              <li>
                <code className="font-mono text-slate-800">_SPSS_Metadata</code> — Saves your 11-column Variable View definitions, missing-value rules, and validation bounds inside the `.xlsx` file.
              </li>
            </ul>
            {lastBackupTimestamp && (
              <p className="text-[11px] text-slate-500 pt-1">
                Last snapshot: {new Date(lastBackupTimestamp).toLocaleString()}
              </p>
            )}
          </div>

          {/* Sideloading Manifest & Desktop Instructions */}
          <div className="p-3.5 border border-slate-200 rounded-md space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2 font-semibold text-slate-900">
                <FileCode2 className="w-4 h-4 text-blue-700" />
                <span>1. Configure &amp; Download Manifest</span>
              </div>
              <button
                type="button"
                onClick={handleDownloadManifest}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#107C41] hover:bg-[#0D6635] rounded-md cursor-pointer shrink-0"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Download manifest.xml</span>
              </button>
            </div>

            <div>
              <label
                htmlFor="manifest-host-url"
                className="block text-[11px] font-medium text-slate-600 mb-1"
              >
                Hosted Add-in HTTPS URL (use your deployed/shared public URL or local HTTPS server)
              </label>
              <input
                id="manifest-host-url"
                type="url"
                value={manifestUrl}
                onChange={(e) => setManifestUrl(e.target.value)}
                className="w-full px-2.5 py-1.5 text-xs font-mono bg-white border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
            </div>

            {/* Platform Tabs */}
            <div className="pt-1">
              <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-md w-fit mb-3">
                {(
                  [
                    { id: 'windows', label: 'Windows Excel Desktop' },
                    { id: 'mac', label: 'macOS Excel Desktop' },
                    { id: 'web', label: 'Excel on the Web' },
                  ] as const
                ).map((tab) => (
                  <button
                    key={tab.id}
                    type="button"
                    onClick={() => setPlatform(tab.id)}
                    className={`px-2.5 py-1 text-[11px] font-medium rounded transition-colors cursor-pointer whitespace-nowrap ${
                      platform === tab.id
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>

              {platform === 'windows' && (
                <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 leading-relaxed">
                  <li>
                    Put <code className="font-mono text-slate-800">statsheet-excel-manifest.xml</code> in a folder (for example <code className="font-mono text-slate-800">C:\ExcelAddins\StatSheet</code>).
                  </li>
                  <li>
                    Right-click the folder → <strong>Properties → Sharing → Share…</strong> to share it with yourself and note its network path (e.g. <code className="font-mono text-slate-800">\\YOUR-PC\StatSheet</code>).
                  </li>
                  <li>
                    In Desktop Excel, open <strong>File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs</strong>.
                  </li>
                  <li>
                    Paste the network path into <strong>Catalog Url</strong>, click <strong>Add catalog</strong>, check <strong>Show in Menu</strong>, click <strong>OK</strong>, and restart Excel.
                  </li>
                  <li>
                    Go to <strong>Home → Add-ins → More Add-ins → Shared Folder</strong> and select <strong>StatSheet</strong>.
                  </li>
                </ol>
              )}

              {platform === 'mac' && (
                <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 leading-relaxed">
                  <li>
                    In Finder, press <code className="font-mono text-slate-800">Cmd + Shift + G</code> and open{' '}
                    <code className="font-mono text-slate-800">
                      ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
                    </code>{' '}
                    (create the <code className="font-mono text-slate-800">wef</code> folder if missing).
                  </li>
                  <li>
                    Copy <code className="font-mono text-slate-800">statsheet-excel-manifest.xml</code> into that folder and restart Excel.
                  </li>
                  <li>
                    Open <strong>Home → Add-ins → My Add-ins</strong> (or <strong>Insert → Add-ins → My Add-ins</strong>) and choose <strong>StatSheet</strong> under Developer Add-ins.
                  </li>
                </ol>
              )}

              {platform === 'web' && (
                <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 leading-relaxed">
                  <li>
                    Open Excel on the web, go to <strong>Home → Add-ins → More Add-ins → My Add-ins</strong>.
                  </li>
                  <li>
                    Click <strong>Upload My Add-in</strong> and choose{' '}
                    <code className="font-mono text-slate-800">statsheet-excel-manifest.xml</code>.
                  </li>
                  <li>
                    Click <strong>Data &amp; Variable View</strong> on the Home ribbon to open the StatSheet Task Pane.
                  </li>
                </ol>
              )}
            </div>
          </div>
        </div>

        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-md cursor-pointer"
          >
            Done
          </button>
        </div>
      </div>
    </div>
  );
}
