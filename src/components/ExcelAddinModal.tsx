import {
  Check,
  CheckCircle2,
  Copy,
  Download,
  FileCode2,
  ShieldCheck,
  Terminal,
  X,
} from 'lucide-react';
import { useState } from 'react';
import { generateOfficeManifestXml } from '../engine/spssEngine';
import {
  generateCustomUiRibbonXml,
  generateStatSheetVbaModule,
  generateWindowsXlamBuilderScript,
} from '../engine/vbaAddinGenerator';
import type { HostEnvironment } from '../types/spss';

interface ExcelAddinModalProps {
  host: HostEnvironment;
  lastBackupTimestamp: string | null;
  onClose: () => void;
}

type AddinModeTab = 'vba' | 'officejs';
type VbaOsTab = 'windows' | 'mac';
type OfficeJsOsTab = 'windows' | 'mac' | 'web';

function triggerTextDownload(filename: string, content: string, mime: string) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function ExcelAddinModal({
  host,
  lastBackupTimestamp,
  onClose,
}: ExcelAddinModalProps) {
  const defaultOrigin =
    typeof window !== 'undefined'
      ? window.location.origin
      : 'https://localhost:3000';

  const [activeTab, setActiveTab] = useState<AddinModeTab>('vba');
  const [vbaOs, setVbaOs] = useState<VbaOsTab>('windows');
  const [showVbaSource, setShowVbaSource] = useState(false);
  const [copiedVba, setCopiedVba] = useState(false);

  const [manifestUrl, setManifestUrl] = useState(defaultOrigin);
  const [officeJsPlatform, setOfficeJsPlatform] =
    useState<OfficeJsOsTab>('windows');

  const vbaCode = generateStatSheetVbaModule();

  const handleCopyVba = async () => {
    try {
      await navigator.clipboard.writeText(vbaCode);
      setCopiedVba(true);
      window.setTimeout(() => setCopiedVba(false), 2000);
    } catch {
      setShowVbaSource(true);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="addin-modal-title"
    >
      <div className="bg-white border border-slate-300 rounded-lg shadow-xl w-full max-w-2xl flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between gap-4">
          <div>
            <h2
              id="addin-modal-title"
              className="font-display text-sm font-semibold text-slate-900"
            >
              Desktop Excel Add-in Installer
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Choose the zero-server offline VBA add-in (.xlam) or the live Office.js Task Pane manifest.
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

        {/* Mode Selector */}
        <div className="px-5 pt-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
          <button
            type="button"
            onClick={() => setActiveTab('vba')}
            className={`px-3.5 py-2 text-xs font-semibold border-b-2 transition-colors cursor-pointer whitespace-nowrap ${
              activeTab === 'vba'
                ? 'border-[#107C41] text-[#107C41]'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            Offline VBA Add-in (.xlam)
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('officejs')}
            className={`px-3.5 py-2 text-xs font-semibold border-b-2 transition-colors cursor-pointer whitespace-nowrap ${
              activeTab === 'officejs'
                ? 'border-blue-700 text-blue-700'
                : 'border-transparent text-slate-600 hover:text-slate-900'
            }`}
          >
            Office.js Task Pane (manifest.xml)
          </button>
        </div>

        {/* Body */}
        <div className="p-5 overflow-y-auto space-y-4 text-xs text-slate-700">
          {activeTab === 'vba' ? (
            <>
              <div className="p-3.5 border border-slate-200 rounded-md bg-slate-50 space-y-1.5">
                <div className="flex items-center gap-2 font-semibold text-slate-900">
                  <ShieldCheck className="w-4 h-4 text-[#107C41]" />
                  <span>True Zero-Server Offline Excel Add-in</span>
                </div>
                <p className="text-slate-600 leading-relaxed">
                  Pure cross-platform VBA (<code className="font-mono text-slate-800">StatSheet_SPSS.bas</code>) for Windows and macOS Desktop Excel. Requires no internet connection, no web server, and no external libraries.
                </p>
                <div className="pt-1 text-[11px] text-slate-600">
                  Includes: 11-column <code className="font-mono">Variable_View</code> sync · Reversible <code className="font-mono">1↔A</code> Value Labels toggle (<code className="font-mono">Ctrl+Shift+L</code>) with hidden <code className="font-mono">_SPSS_RawBackup</code> · View switcher (<code className="font-mono">Ctrl+Shift+V</code>) · Missing &amp; validation highlighter · Descriptives report · JSON Codebook export.
                </div>
              </div>

              {/* Downloads */}
              <div className="p-3.5 border border-slate-200 rounded-md space-y-3">
                <div className="font-semibold text-slate-900">
                  1. Download Offline Add-in Files
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() =>
                      triggerTextDownload(
                        'StatSheet_SPSS.bas',
                        vbaCode,
                        'text/plain;charset=utf-8'
                      )
                    }
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#107C41] hover:bg-[#0D6635] rounded-md cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Download StatSheet_SPSS.bas</span>
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      triggerTextDownload(
                        'Build_StatSheet_xlam.ps1',
                        generateWindowsXlamBuilderScript(),
                        'text/plain;charset=utf-8'
                      )
                    }
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-800 bg-white border border-slate-300 hover:bg-slate-50 rounded-md cursor-pointer"
                  >
                    <Terminal className="w-3.5 h-3.5 text-slate-600" />
                    <span>Windows Auto-Builder (.ps1)</span>
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      triggerTextDownload(
                        'customUI14.xml',
                        generateCustomUiRibbonXml(),
                        'application/xml'
                      )
                    }
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md cursor-pointer"
                  >
                    <FileCode2 className="w-3.5 h-3.5 text-slate-500" />
                    <span>Ribbon XML (opt.)</span>
                  </button>

                  <button
                    type="button"
                    onClick={handleCopyVba}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md cursor-pointer"
                  >
                    {copiedVba ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-[#16A34A]" />
                        <span>Copied</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3.5 h-3.5" />
                        <span>Copy VBA</span>
                      </>
                    )}
                  </button>
                </div>

                <button
                  type="button"
                  onClick={() => setShowVbaSource((s) => !s)}
                  className="text-[11px] font-medium text-blue-700 hover:underline cursor-pointer"
                >
                  {showVbaSource ? 'Hide VBA source code' : 'Preview VBA source code'}
                </button>

                {showVbaSource && (
                  <pre className="max-h-52 overflow-auto p-3 bg-slate-900 text-slate-100 rounded-md font-mono text-[11px] leading-relaxed">
                    {vbaCode}
                  </pre>
                )}
              </div>

              {/* Install Steps */}
              <div className="p-3.5 border border-slate-200 rounded-md space-y-3">
                <div className="flex items-center justify-between gap-3">
                  <span className="font-semibold text-slate-900">
                    2. Install as a Permanent .xlam Add-in
                  </span>
                  <div className="flex items-center gap-1 p-0.5 bg-slate-100 rounded-md">
                    {(
                      [
                        { id: 'windows', label: 'Windows' },
                        { id: 'mac', label: 'macOS' },
                      ] as const
                    ).map((t) => (
                      <button
                        key={t.id}
                        type="button"
                        onClick={() => setVbaOs(t.id)}
                        className={`px-2.5 py-1 text-[11px] font-medium rounded cursor-pointer ${
                          vbaOs === t.id
                            ? 'bg-white text-slate-900 shadow-xs'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        {t.label}
                      </button>
                    ))}
                  </div>
                </div>

                {vbaOs === 'windows' ? (
                  <div className="space-y-2.5 text-slate-600 leading-relaxed">
                    <div>
                      <strong className="text-slate-800">Option A — One-Click Windows Builder:</strong>{' '}
                      Download both <code className="font-mono text-slate-800">StatSheet_SPSS.bas</code> and{' '}
                      <code className="font-mono text-slate-800">Build_StatSheet_xlam.ps1</code> into the same folder, right-click the <code className="font-mono">.ps1</code> file → <strong>Run with PowerShell</strong>. It compiles and enables <code className="font-mono">StatSheet_SPSS.xlam</code> automatically.
                    </div>
                    <div>
                      <strong className="text-slate-800">Option B — Manual 30-Second Import (no Trust Center changes):</strong>
                      <ol className="list-decimal pl-5 mt-1 space-y-1">
                        <li>Open a blank workbook in Excel and press <code className="font-mono text-slate-800">Alt + F11</code> to open the VBA Editor.</li>
                        <li>Click <strong>File → Import File…</strong> and choose <code className="font-mono text-slate-800">StatSheet_SPSS.bas</code>.</li>
                        <li>Close the VBA Editor, then press <code className="font-mono text-slate-800">F12</code> (Save As) and set Save as type to <strong>Excel Add-in (*.xlam)</strong> (Excel automatically selects <code className="font-mono">%APPDATA%\Microsoft\AddIns</code>).</li>
                        <li>Enable it via <strong>File → Options → Add-ins → Manage: Excel Add-ins → Go…</strong> and check <strong>StatSheet_SPSS</strong>.</li>
                      </ol>
                    </div>
                  </div>
                ) : (
                  <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 leading-relaxed">
                    <li>Open a blank workbook in Excel for Mac and press <code className="font-mono text-slate-800">Option + F11</code> (or <strong>Tools → Macro → Visual Basic Editor</strong>).</li>
                    <li>Click <strong>File → Import File…</strong> and select <code className="font-mono text-slate-800">StatSheet_SPSS.bas</code>.</li>
                    <li>Return to Excel, choose <strong>File → Save As…</strong>, and change File Format to <strong>Excel Add-in (.xlam)</strong>.</li>
                    <li>Enable it under <strong>Tools → Excel Add-ins…</strong>, then use <strong>Add-ins → StatSheet SPSS</strong> or <code className="font-mono">Ctrl+Shift+L</code> / <code className="font-mono">Ctrl+Shift+V</code>.</li>
                  </ol>
                )}
              </div>
            </>
          ) : (
            <>
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
                      : 'You are previewing the interactive simulator. Office.js Task Pane add-ins require an HTTPS host URL; use the Offline VBA Add-in tab if you need zero-server offline operation.'}
                  </p>
                  {lastBackupTimestamp && (
                    <p className="text-[11px] text-slate-500 mt-1">
                      Last backup snapshot: {new Date(lastBackupTimestamp).toLocaleString()}
                    </p>
                  )}
                </div>
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
                    onClick={() => {
                      const clean = manifestUrl.trim() || defaultOrigin;
                      triggerTextDownload(
                        'statsheet-excel-manifest.xml',
                        generateOfficeManifestXml(clean),
                        'application/xml'
                      );
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-700 hover:bg-blue-800 rounded-md cursor-pointer shrink-0"
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

                <div className="pt-1">
                  <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-md w-fit mb-3">
                    {(
                      [
                        { id: 'windows', label: 'Windows Desktop' },
                        { id: 'mac', label: 'macOS Desktop' },
                        { id: 'web', label: 'Excel on the Web' },
                      ] as const
                    ).map((tab) => (
                      <button
                        key={tab.id}
                        type="button"
                        onClick={() => setOfficeJsPlatform(tab.id)}
                        className={`px-2.5 py-1 text-[11px] font-medium rounded transition-colors cursor-pointer whitespace-nowrap ${
                          officeJsPlatform === tab.id
                            ? 'bg-white text-slate-900 shadow-xs'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        {tab.label}
                      </button>
                    ))}
                  </div>

                  {officeJsPlatform === 'windows' && (
                    <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 leading-relaxed">
                      <li>
                        Put <code className="font-mono text-slate-800">statsheet-excel-manifest.xml</code> in a shared folder (e.g. <code className="font-mono text-slate-800">\\YOUR-PC\StatSheet</code>).
                      </li>
                      <li>
                        In Excel, go to <strong>File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs</strong>, add the network path, check <strong>Show in Menu</strong>, and restart Excel.
                      </li>
                      <li>
                        Open <strong>Home → Add-ins → More Add-ins → Shared Folder</strong> and select <strong>StatSheet</strong>.
                      </li>
                    </ol>
                  )}

                  {officeJsPlatform === 'mac' && (
                    <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 leading-relaxed">
                      <li>
                        Copy <code className="font-mono text-slate-800">statsheet-excel-manifest.xml</code> into{' '}
                        <code className="font-mono text-slate-800">
                          ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
                        </code>{' '}
                        and restart Excel.
                      </li>
                      <li>
                        Open <strong>Home → Add-ins → My Add-ins</strong> and select <strong>StatSheet</strong>.
                      </li>
                    </ol>
                  )}

                  {officeJsPlatform === 'web' && (
                    <ol className="list-decimal pl-5 space-y-1.5 text-slate-600 leading-relaxed">
                      <li>
                        Open Excel on the web → <strong>Home → Add-ins → More Add-ins → My Add-ins</strong>.
                      </li>
                      <li>
                        Click <strong>Upload My Add-in</strong> and select <code className="font-mono text-slate-800">statsheet-excel-manifest.xml</code>.
                      </li>
                    </ol>
                  )}
                </div>
              </div>
            </>
          )}
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
