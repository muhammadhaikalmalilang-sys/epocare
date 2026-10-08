import React, { useState } from 'react';
import { User } from 'firebase/auth';
import { 
  X, 
  FileSpreadsheet, 
  ArrowDownToLine, 
  ArrowUpFromLine, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCw, 
  Copy, 
  Check, 
  Zap,
  ExternalLink,
  Link2,
  Download,
  HelpCircle
} from 'lucide-react';
import { PatientRecord, SpreadsheetConfig } from '../types/dialysis';
import { APPS_SCRIPT_SAMPLE_CODE, DEFAULT_APPS_SCRIPT_URL, OFFICIAL_SPREADSHEET_URL, OFFICIAL_SPREADSHEET_ID, sanitizeAppsScriptUrl } from '../services/googleSheets';

interface GoogleSheetsSyncModalProps {
  isOpen: boolean;
  onClose: () => void;
  user?: User | null;
  spreadsheetConfig: SpreadsheetConfig | null;
  onSaveSpreadsheetConfig: (config: SpreadsheetConfig) => void;
  patients: PatientRecord[];
  onPullFromSheet?: (spreadsheetId: string) => Promise<void>;
  onPushToSheet?: (spreadsheetId: string) => Promise<void>;
  onPullViaAppsScript: (url: string) => Promise<void>;
  onPushViaAppsScript: (url: string) => Promise<void>;
  onImportCsv?: (csvText: string) => void;
  onSignIn?: () => void;
  onOpenExportExcelModal?: () => void;
  selectedMonth: string;
  autoSyncEnabled?: boolean;
  onToggleAutoSync?: () => void;
}

export const GoogleSheetsSyncModal: React.FC<GoogleSheetsSyncModalProps> = ({
  isOpen,
  onClose,
  spreadsheetConfig,
  onSaveSpreadsheetConfig,
  patients,
  onPullViaAppsScript,
  onPushViaAppsScript,
  onOpenExportExcelModal,
  autoSyncEnabled = true,
  onToggleAutoSync,
}) => {
  // Input states - Default to target official Web App URL
  const [appsScriptUrl, setAppsScriptUrl] = useState(
    spreadsheetConfig?.appsScriptUrl || DEFAULT_APPS_SCRIPT_URL
  );
  const [spreadsheetUrlInput, setSpreadsheetUrlInput] = useState(
    spreadsheetConfig?.spreadsheetUrl || OFFICIAL_SPREADSHEET_URL
  );

  // Operation states
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);
  const [copiedCode, setCopiedCode] = useState(false);
  const [showGuide, setShowGuide] = useState(false);

  // Perbarui nilai input sesuai konfigurasi aktif saat modal dibuka
  React.useEffect(() => {
    if (isOpen) {
      setAppsScriptUrl(spreadsheetConfig?.appsScriptUrl || DEFAULT_APPS_SCRIPT_URL);
      setSpreadsheetUrlInput(spreadsheetConfig?.spreadsheetUrl || OFFICIAL_SPREADSHEET_URL);
    }
  }, [isOpen, spreadsheetConfig]);

  if (!isOpen) return null;

  // Salin Kode Apps Script
  const handleCopyCode = () => {
    navigator.clipboard.writeText(APPS_SCRIPT_SAMPLE_CODE);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 3000);
  };

  // Simpan URL Apps Script
  const handleSaveAppsScriptUrl = () => {
    const raw = appsScriptUrl.trim() || DEFAULT_APPS_SCRIPT_URL;
    const targetUrl = sanitizeAppsScriptUrl(raw) || DEFAULT_APPS_SCRIPT_URL;
    setAppsScriptUrl(targetUrl);
    if (!targetUrl.startsWith('http')) {
      setStatusMessage({ 
        type: 'error', 
        text: 'Masukkan Web App URL Apps Script yang valid (diawali https://script.google.com/...)' 
      });
      return;
    }
    const updated: SpreadsheetConfig = {
      spreadsheetId: spreadsheetConfig?.spreadsheetId || 'appsscript-connected',
      sheetName: 'Senin-Kamis, Selasa-Jumat, Rabu-Sabtu',
      appsScriptUrl: targetUrl,
      syncMode: 'appsscript',
      lastSyncedAt: new Date().toISOString(),
    };
    onSaveSpreadsheetConfig(updated);
    setStatusMessage({ type: 'success', text: 'Web App URL Apps Script berhasil disimpan!' });
  };

  // Simpan URL Spreadsheet
  const handleSaveSpreadsheetUrl = () => {
    const cleanUrl = spreadsheetUrlInput.trim();
    let extractedId = spreadsheetConfig?.spreadsheetId || 'appsscript-connected';
    if (cleanUrl) {
      const match = cleanUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) {
        extractedId = match[1];
      }
    }
    const updated: SpreadsheetConfig = {
      spreadsheetId: extractedId,
      sheetName: spreadsheetConfig?.sheetName || 'Senin-Kamis, Selasa-Jumat, Rabu-Sabtu',
      spreadsheetUrl: cleanUrl,
      appsScriptUrl: appsScriptUrl.trim() || spreadsheetConfig?.appsScriptUrl || DEFAULT_APPS_SCRIPT_URL,
      syncMode: 'appsscript',
      lastSyncedAt: new Date().toISOString(),
    };
    onSaveSpreadsheetConfig(updated);
    setStatusMessage({ type: 'success', text: 'URL Google Sheets berhasil disimpan!' });
  };

  // Tarik via Apps Script (Tanpa Login)
  const handlePullAppsScript = async () => {
    const url = appsScriptUrl.trim() || spreadsheetConfig?.appsScriptUrl || DEFAULT_APPS_SCRIPT_URL;
    if (!url) {
      setStatusMessage({ type: 'error', text: 'Masukkan Web App URL Apps Script terlebih dahulu.' });
      return;
    }
    try {
      setIsProcessing(true);
      setStatusMessage({ type: 'info', text: 'Menghubungi Apps Script dan membaca data...' });
      await onPullViaAppsScript(url);
      handleSaveAppsScriptUrl();
      setStatusMessage({ type: 'success', text: 'Berhasil membaca data alokasi dari Google Sheet tanpa login!' });
    } catch (err: any) {
      console.error(err);
      setStatusMessage({ type: 'error', text: err.message || 'Gagal terhubung dengan Apps Script.' });
    } finally {
      setIsProcessing(false);
    }
  };

  // Kirim via Apps Script (Tanpa Login)
  const handlePushAppsScript = async () => {
    const url = appsScriptUrl.trim() || spreadsheetConfig?.appsScriptUrl || DEFAULT_APPS_SCRIPT_URL;
    if (!url) {
      setStatusMessage({ type: 'error', text: 'Masukkan Web App URL Apps Script terlebih dahulu.' });
      return;
    }
    try {
      setIsProcessing(true);
      setStatusMessage({ type: 'info', text: 'Mengirimkan data alokasi pasien ke Google Sheet...' });
      await onPushViaAppsScript(url);
      handleSaveAppsScriptUrl();
      setStatusMessage({ type: 'success', text: `Sukses menyimpan ${patients.length} data pasien ke Google Sheet via Apps Script!` });
    } catch (err: any) {
      console.error(err);
      setStatusMessage({ type: 'error', text: err.message || 'Gagal mengirim ke Apps Script.' });
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-3 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-xl max-w-md sm:max-w-lg w-full shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden flex flex-col max-h-[85vh]">
        
        {/* Header (Pinned) */}
        <div className="px-4 py-2.5 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between bg-emerald-50/50 dark:bg-emerald-950/30 shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 rounded-lg bg-emerald-600 text-white flex items-center justify-center shadow-xs shrink-0">
              <FileSpreadsheet className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900 dark:text-white leading-tight">
                Sinkronisasi Google Sheets
              </h3>
              <p className="text-[11px] text-slate-500 dark:text-slate-400">
                Alokasi 2 arah via Apps Script (Bebas Login Akun)
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition cursor-pointer"
            aria-label="Tutup"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body Content (Compact) */}
        <div className="p-3 space-y-2.5 text-xs overflow-y-auto flex-1">
          
          {/* Notification Message */}
          {statusMessage && (
            <div className={`p-2 rounded-lg flex items-start gap-1.5 ${
              statusMessage.type === 'success' 
                ? 'bg-emerald-50 text-emerald-800 border border-emerald-200 dark:bg-emerald-950/40 dark:text-emerald-300 dark:border-emerald-800'
                : statusMessage.type === 'error'
                ? 'bg-rose-50 text-rose-800 border border-rose-200 dark:bg-rose-950/40 dark:text-rose-300 dark:border-rose-800'
                : 'bg-blue-50 text-blue-800 border border-blue-200 dark:bg-blue-950/40 dark:text-blue-300 dark:border-blue-800'
            }`}>
              {statusMessage.type === 'success' && <CheckCircle2 className="w-3.5 h-3.5 mt-0.5 shrink-0" />}
              {statusMessage.type === 'error' && <AlertCircle className="w-3.5 h-3.5 mt-0.5 shrink-0" />}
              {statusMessage.type === 'info' && <RefreshCw className="w-3.5 h-3.5 mt-0.5 shrink-0 animate-spin" />}
              <span className="text-[10px] sm:text-[11px] font-medium">{statusMessage.text}</span>
            </div>
          )}

          {/* Panduan Pemecahan Masalah jika Error Apps Script */}
          {statusMessage?.type === 'error' && (
            <div className="p-2.5 rounded-lg border border-amber-300 dark:border-amber-800 bg-amber-50/90 dark:bg-amber-950/40 text-amber-900 dark:text-amber-200 text-[10.5px] space-y-1.5 animate-in fade-in duration-200">
              <div className="font-bold flex items-center gap-1.5 text-xs text-amber-950 dark:text-amber-100">
                <HelpCircle className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                <span>Solusi Mengatasi "Gagal menghubungi Apps Script":</span>
              </div>
              <ol className="list-decimal pl-4 space-y-1 text-slate-700 dark:text-slate-300">
                <li>
                  <strong>Akses Deployment Belum "Anyone":</strong> Di Google Apps Script, klik tombol biru <strong>Deploy</strong> &gt; <strong>Kelola deployment (Manage deployments)</strong>. Pastikan pada opsi <em>"Who has access" (Siapa yang memiliki akses)</em> dipilih <strong>Anyone (Siapa saja)</strong>, bukan "Only myself" (Hanya saya).
                </li>
                <li>
                  <strong>Akhiran URL:</strong> Pastikan Web App URL berakhiran <code>/exec</code> (bukan <code>/edit</code>).
                </li>
                <li>
                  <strong>Koneksi Langsung:</strong> Anda juga dapat mengklik tombol <button type="button" onClick={() => { setAppsScriptUrl(DEFAULT_APPS_SCRIPT_URL); handleSaveAppsScriptUrl(); }} className="text-emerald-700 dark:text-emerald-400 font-bold underline cursor-pointer hover:text-emerald-900">"Gunakan URL Resmi"</button> di bawah untuk menggunakan URL Web App bawaan yang sudah terverifikasi.
                </li>
              </ol>
            </div>
          )}

          {/* Auto-Sync Real-Time Banner */}
          {onToggleAutoSync && (
            <div className={`p-2.5 rounded-lg border transition ${
              autoSyncEnabled
                ? 'bg-emerald-50/90 dark:bg-emerald-950/40 border-emerald-300 dark:border-emerald-800'
                : 'bg-slate-50 dark:bg-slate-850 border-slate-200 dark:border-slate-800'
            }`}>
              <div className="flex items-center justify-between gap-2">
                <div className="flex items-center gap-2 min-w-0">
                  <div className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 ${
                    autoSyncEnabled ? 'bg-emerald-600 text-white shadow-2xs' : 'bg-slate-300 dark:bg-slate-700 text-slate-600 dark:text-slate-300'
                  }`}>
                    <RefreshCw className={`w-3.5 h-3.5 ${autoSyncEnabled ? 'animate-spin' : ''}`} />
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 dark:text-white text-xs leading-tight flex items-center gap-1.5">
                      <span>Auto-Sync Real-Time ke Google Sheets</span>
                      <span className={`text-[9px] px-1.5 py-0.2 rounded font-extrabold ${
                        autoSyncEnabled ? 'bg-emerald-200 text-emerald-900 dark:bg-emerald-900 dark:text-emerald-100' : 'bg-slate-200 text-slate-700 dark:bg-slate-700 dark:text-slate-300'
                      }`}>
                        {autoSyncEnabled ? 'AKTIF' : 'NONAKTIF'}
                      </span>
                    </h4>
                    <p className="text-[10px] text-slate-600 dark:text-slate-400 leading-tight">
                      {autoSyncEnabled
                        ? 'Setiap kali Anda mengedit pasien, input Hb, atau ubah jadwal, sistem langsung mengedit data di Google Sheets secara otomatis tanpa perlu kirim manual.'
                        : 'Auto-sync dinonaktifkan. Pengiriman data ke Google Sheets harus ditekan manual melalui tombol Kirim.'}
                    </p>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={onToggleAutoSync}
                  className={`h-7 px-3 rounded-md text-xs font-bold transition cursor-pointer shrink-0 shadow-2xs ${
                    autoSyncEnabled
                      ? 'bg-rose-700 hover:bg-rose-800 text-white'
                      : 'bg-emerald-700 hover:bg-emerald-800 text-white'
                  }`}
                >
                  {autoSyncEnabled ? 'Matikan' : 'Aktifkan'}
                </button>
              </div>
            </div>
          )}

          {/* Web App URL Panel */}
          <div className="p-2 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50/80 dark:bg-slate-850/60 space-y-1.5">
            <div className="flex items-center justify-between">
              <label className="font-bold text-slate-800 dark:text-slate-200 text-[10px] sm:text-[11px] flex items-center gap-1">
                <Link2 className="w-3 h-3 text-emerald-600" />
                <span>Web App URL Google Apps Script:</span>
              </label>
              <button
                type="button"
                onClick={() => {
                  setAppsScriptUrl(DEFAULT_APPS_SCRIPT_URL);
                  handleSaveAppsScriptUrl();
                }}
                className="text-[9px] sm:text-[10px] text-emerald-700 dark:text-emerald-400 hover:underline font-bold cursor-pointer"
                title="Gunakan Web App URL Resmi"
              >
                Gunakan URL Resmi
              </button>
            </div>

            <div className="flex gap-1.5">
              <input
                type="text"
                placeholder="https://script.google.com/macros/s/.../exec"
                value={appsScriptUrl}
                onChange={(e) => setAppsScriptUrl(e.target.value)}
                className="flex-1 px-2 py-1 rounded-md border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white font-mono text-[10px] sm:text-[11px] focus:ring-1 focus:ring-emerald-500 focus:outline-hidden"
              />
              <button
                type="button"
                onClick={handleSaveAppsScriptUrl}
                className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-md font-bold text-[11px] transition cursor-pointer shrink-0"
              >
                Simpan
              </button>
            </div>

            <div className="flex items-center gap-1 text-[9px] text-slate-500 dark:text-slate-400">
              <CheckCircle2 className="w-2.5 h-2.5 text-emerald-600 shrink-0" />
              <span className="truncate">Tersambung otomatis ke 3 sheet: <em>Senin-Kamis, Selasa-Jumat, Rabu-Sabtu</em></span>
            </div>
          </div>

          {/* Direct Spreadsheet URL Panel */}
          <div className="p-2 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50/80 dark:bg-slate-850/60 space-y-1.5">
            <div className="flex items-center justify-between">
              <label className="font-bold text-slate-800 dark:text-slate-200 text-[10px] sm:text-[11px] flex items-center gap-1">
                <FileSpreadsheet className="w-3 h-3 text-emerald-600" />
                <span>Tautan Dokumen Google Sheets:</span>
              </label>
            </div>

            <div className="flex gap-1.5 flex-wrap sm:flex-nowrap">
              <input
                type="text"
                placeholder="https://docs.google.com/spreadsheets/d/.../edit"
                value={spreadsheetUrlInput}
                onChange={(e) => setSpreadsheetUrlInput(e.target.value)}
                className="flex-1 min-w-[180px] px-2 py-1 rounded-md border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white font-mono text-[10px] sm:text-[11px] focus:ring-1 focus:ring-emerald-500 focus:outline-hidden"
              />
              <button
                type="button"
                onClick={handleSaveSpreadsheetUrl}
                className="px-2.5 py-1 bg-slate-700 hover:bg-slate-800 text-white rounded-md font-bold text-[11px] transition cursor-pointer shrink-0"
              >
                Simpan
              </button>
              {(spreadsheetUrlInput || spreadsheetConfig?.spreadsheetUrl || OFFICIAL_SPREADSHEET_URL) && (
                <a
                  href={spreadsheetUrlInput || spreadsheetConfig?.spreadsheetUrl || OFFICIAL_SPREADSHEET_URL}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 active:bg-emerald-700 text-white rounded-md font-bold text-[11px] transition cursor-pointer shrink-0 flex items-center gap-1 shadow-xs"
                  title="Buka dokumen Google Sheets langsung di tab browser baru"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  <span>Buka Sheet</span>
                  <ExternalLink className="w-2.5 h-2.5" />
                </a>
              )}
            </div>
          </div>

          {/* Action Buttons: Tarik & Kirim */}
          <div className="grid grid-cols-2 gap-2">
            <button
              type="button"
              onClick={handlePullAppsScript}
              disabled={isProcessing || !appsScriptUrl.trim()}
              className="h-8.5 px-3 rounded-lg font-semibold text-xs bg-blue-50 text-blue-700 hover:bg-blue-100 dark:bg-blue-950/60 dark:text-blue-300 dark:hover:bg-blue-900 border border-blue-200 dark:border-blue-800 disabled:opacity-50 transition flex items-center justify-center gap-1.5 cursor-pointer shadow-2xs"
            >
              {isProcessing ? (
                <RefreshCw className="w-3.5 h-3.5 shrink-0 animate-spin" />
              ) : (
                <ArrowDownToLine className="w-3.5 h-3.5 shrink-0" />
              )}
              <span>{isProcessing ? 'Menarik...' : 'Tarik Data (Pull)'}</span>
            </button>
            <button
              type="button"
              onClick={handlePushAppsScript}
              disabled={isProcessing || !appsScriptUrl.trim()}
              className="h-8.5 px-3 rounded-lg font-semibold text-xs bg-emerald-600 text-white hover:bg-emerald-700 active:bg-emerald-800 disabled:opacity-50 transition shadow-xs flex items-center justify-center gap-1.5 cursor-pointer"
            >
              {isProcessing ? (
                <RefreshCw className="w-3.5 h-3.5 shrink-0 animate-spin" />
              ) : (
                <ArrowUpFromLine className="w-3.5 h-3.5 shrink-0" />
              )}
              <span>{isProcessing ? 'Mengirim...' : 'Kirim ke Sheet (Push)'}</span>
            </button>
          </div>

          {/* Opsi Ekspor File Excel Offline (Bebas Login) */}
          {onOpenExportExcelModal && (
            <div className="p-2.5 rounded-lg border border-emerald-300 dark:border-emerald-800 bg-emerald-50/80 dark:bg-emerald-950/40 flex items-center justify-between gap-2">
              <div className="space-y-0.5">
                <span className="font-bold text-emerald-900 dark:text-emerald-200 text-xs flex items-center gap-1">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Ekspor Menjadi File Excel (.xlsx)</span>
                </span>
                <p className="text-[10px] text-emerald-800 dark:text-emerald-300">
                  Unduh langsung file Excel ke komputer/HP tanpa perlu login akun Google.
                </p>
              </div>
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onOpenExportExcelModal();
                }}
                className="h-7.5 px-3 bg-emerald-700 hover:bg-emerald-800 text-white rounded-md font-bold text-[11px] transition cursor-pointer flex items-center gap-1.5 shrink-0 shadow-xs"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Unduh Excel</span>
              </button>
            </div>
          )}

          {/* Penjelasan Arsitektur Terintegrasi 1 Tahun & Kecepatan Data */}
          <div className="p-2.5 rounded-lg border border-emerald-300 dark:border-emerald-800 bg-emerald-50/70 dark:bg-emerald-950/30 text-[10px] sm:text-[11px] text-slate-700 dark:text-slate-300 space-y-1.5">
            <div className="font-bold text-emerald-900 dark:text-emerald-200 flex items-center gap-1.5 text-xs">
              <Zap className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
              <span>Arsitektur 1 Tahun Terintegrasi (5 Lembar Tab Resmi)</span>
            </div>
            <ul className="list-disc pl-4 space-y-1 text-slate-600 dark:text-slate-400">
              <li>
                <strong className="text-slate-800 dark:text-slate-200">1. Tab Master <code>REKAP HB TAHUNAN</code>:</strong> Menyimpan 12 bulan (Jan–Des). <strong>Nilai Hb hanya ada pada sheet ini</strong>. Nilai Hb bulan terbaru otomatis menjadi acuan untuk menentukan pasien wajib <strong>⭐ Cek Hb Pilihan (&lt; 9.0)</strong> di bulan berikutnya (contoh: Nilai September menjadi acuan untuk sesi Oktober).
              </li>
              <li>
                <strong className="text-slate-800 dark:text-slate-200">2. Tab Matriks <code>MATRIK CEK HB</code>:</strong> Matriks kalender 6 hari sesi HD pertama awal bulan untuk penjadwalan kehadiran sampling pasien. Lembar <code>JADWAL CEK HB</code> lama telah dihapus sepenuhnya.
              </li>
              <li>
                <strong className="text-slate-800 dark:text-slate-200">3, 4, 5. 3 Tab Jadwal Harian (Tanpa Nilai HB):</strong> <code>Senin-Kamis</code>, <code>Selasa-Jumat</code>, <code>Rabu-Sabtu</code> dengan matriks tgl 1–31, pemisahan Shift Pagi/Siang, dan status tindakan (✅ / ❌ / 2000 / PRC) murni tanpa kolom nilai Hb.
              </li>
            </ul>
          </div>

          {/* Collapsible Script Guide & Code */}
          <div className="pt-0.5">
            <button
              type="button"
              onClick={() => setShowGuide(!showGuide)}
              className="w-full h-8 px-2.5 rounded-lg border border-emerald-300 dark:border-emerald-700 bg-emerald-50/70 dark:bg-emerald-950/40 hover:bg-emerald-100/80 dark:hover:bg-emerald-900/40 flex items-center justify-between text-xs font-semibold text-emerald-800 dark:text-emerald-300 transition cursor-pointer shadow-2xs"
            >
              <span className="flex items-center gap-1.5">
                <Zap className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                <span>{showGuide ? 'Sembunyikan Kode Skrip' : '⚡ Salin Kode Skrip Versi Cepat (Turbo Speed - 10x Lebih Cepat)'}</span>
              </span>
              <span className="text-emerald-600 dark:text-emerald-400 text-[10px]">{showGuide ? '▲' : '▼'}</span>
            </button>

            {showGuide && (
              <div className="mt-2 space-y-2 p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-850 text-xs">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-1.5">
                    <Zap className="w-3.5 h-3.5 text-amber-500 shrink-0" />
                    <span className="font-bold text-slate-800 dark:text-slate-200 text-xs">
                      Kode Apps Script Turbo Speed (1 Tahun Terintegrasi):
                    </span>
                  </div>
                  <button
                    onClick={handleCopyCode}
                    className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-md text-[11px] font-bold transition cursor-pointer shadow-xs"
                  >
                    {copiedCode ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                    <span>{copiedCode ? 'Tersalin!' : 'Salin Seluruh Kode'}</span>
                  </button>
                </div>
                <div className="text-[10.5px] text-emerald-800 dark:text-emerald-300 bg-emerald-50/80 dark:bg-emerald-950/40 p-2 rounded border border-emerald-200 dark:border-emerald-800">
                  ⚡ <strong>Peningkatan Kecepatan:</strong> Kode ini menghilangkan proses formatting berulang yang lambat di server Google Sheets dan beralih ke <em>2D batch array</em>. Pengiriman data kini berkurang dari ~15-20 detik menjadi <strong>hanya ~1-2 detik</strong>!
                </div>
                <pre className="p-2.5 bg-slate-900 text-emerald-300 rounded-md text-[10px] font-mono overflow-x-auto max-h-36 leading-relaxed select-all">
                  {APPS_SCRIPT_SAMPLE_CODE}
                </pre>
                <div className="text-[10px] text-slate-600 dark:text-slate-400 space-y-1 bg-white dark:bg-slate-900 p-2 rounded border border-slate-200 dark:border-slate-800">
                  <p className="font-semibold text-slate-800 dark:text-slate-200">Cara Memperbarui Skrip di Google Sheet:</p>
                  <ol className="list-decimal pl-4 space-y-0.5">
                    <li>Buka Google Spreadsheet Anda &gt; menu <strong>Ekstensi</strong> &gt; <strong>Apps Script</strong>.</li>
                    <li>Ganti seluruh isi file <code>Code.gs</code> dengan kode turbo yang disalin di atas.</li>
                    <li>Klik <strong>Terapkan (Deploy)</strong> &gt; <strong>Kelola Penerapan</strong> &gt; Edit (ikon pensil) &gt; Versi Baru &gt; <strong>Terapkan</strong>.</li>
                    <li>Pastikan izin akses disetel ke <em>"Siapa saja" (Anyone)</em>.</li>
                  </ol>
                </div>
              </div>
            )}
          </div>

        </div>

        {/* Modal Footer (Pinned) */}
        <div className="px-4 py-2.5 bg-slate-50 dark:bg-slate-850 border-t border-slate-200 dark:border-slate-800 flex items-center justify-between shrink-0">
          <span className="text-[11px] text-slate-500 truncate">
            {spreadsheetConfig?.lastSyncedAt 
              ? `Sinkron: ${new Date(spreadsheetConfig.lastSyncedAt).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}` 
              : 'Siap sinkronisasi'}
          </span>
          <button
            onClick={onClose}
            className="h-8.5 px-3.5 rounded-lg font-medium text-xs text-slate-700 dark:text-slate-300 hover:bg-slate-200/60 dark:hover:bg-slate-800 transition cursor-pointer"
          >
            Tutup
          </button>
        </div>

      </div>
    </div>
  );
};
