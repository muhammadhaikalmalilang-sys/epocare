/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { User } from 'firebase/auth';
import { Header } from './components/Header';
import { DashboardStats } from './components/DashboardStats';
import { PatientTable } from './components/PatientTable';
import { PatientModal } from './components/PatientModal';
import { GoogleSheetsSyncModal } from './components/GoogleSheetsSyncModal';
import { ClinicalCalculatorModal } from './components/ClinicalCalculatorModal';
import { PrintReportModal } from './components/PrintReportModal';
import { BulkImportModal } from './components/BulkImportModal';
import { MonthlyHbInputModal } from './components/MonthlyHbInputModal';
import { LabScheduleModal } from './components/LabScheduleModal';
import { ExportExcelModal } from './components/ExportExcelModal';
import { WorkflowGuideModal } from './components/WorkflowGuideModal';
import { EpoSchedulePrintModal } from './components/EpoSchedulePrintModal';
import { PatientRecord, SpreadsheetConfig, DoseStatus, HDDaySchedule, LabSchedule, LabScheduleScope, isSelectiveHbCandidate } from './types/dialysis';
import { getInitialDemoPatients, calculateClinicalRecommendation, generateDefaultWeeks, isDemoPatient } from './services/clinicalRules';
import { 
  initAuth, 
  googleSignIn, 
  logout, 
  getAccessToken 
} from './services/firebaseAuth';
import { 
  readPatientsFromSheet, 
  pushPatientsToSheet,
  pullViaAppsScript,
  pushViaAppsScript,
  importFromCSV,
  getPairedHDDate,
  getDateDoseDisplayInfo,
  getMonthDaysInfo,
  getHDSessionWeek,
  getFirstHDDateOfMonth,
  getSecondHDDateOfMonth,
  getNextYearMonth,
  getPatientMonthHDSessions,
  DEFAULT_APPS_SCRIPT_URL,
  OFFICIAL_SPREADSHEET_ID,
  OFFICIAL_SPREADSHEET_URL,
  isSummaryOrHeaderRow
} from './services/googleSheets';
import { 
  FileSpreadsheet, 
  Sparkles, 
  CheckCircle2, 
  AlertCircle, 
  Info,
  ArrowRight
} from 'lucide-react';

const isRemovedDemoNote = (note?: string): boolean => {
  if (!note) return false;
  const lower = note.toLowerCase().trim();
  return (
    lower.includes('penjadwalan epo hari awal: pemberian epo tgl 7 september 2026 ditunda') ||
    lower.includes('pasien mengeluh lemas, konjungtiva anemis berat') ||
    lower.includes('hasil lab hb awal bulan belum diinputkan') ||
    lower.includes('transfusi 1 bag saat hd running') ||
    lower.includes('rutin epo 4x sebulan di hari awal (senin)') ||
    lower.includes('rutin epo 4x sebulan, periksa saturasi transferin') ||
    lower.includes('target hb tercapai stabil, maintenance 1 ampul di m1') ||
    lower.includes('hb di atas target 12 g/dl, tunda injeksi epo bulan ini') ||
    lower.includes('hb diatas 12.00 mg/dl')
  );
};

/**
 * Memastikan seluruh pasien memiliki ID unik tanpa duplikasi (mencegah error React key)
 */
export const ensureUniquePatientIds = (patientList: PatientRecord[]): PatientRecord[] => {
  const seenIds = new Set<string>();
  return patientList.map((p, index) => {
    let patientId = (p.id || '').trim();
    if (!patientId || seenIds.has(patientId)) {
      const cleanRm = (p.noRm || `pat${index}`).replace(/[^a-zA-Z0-9_-]/g, '');
      patientId = `${patientId || 'pat'}-${cleanRm}-${index}-${Math.random().toString(36).substring(2, 7)}`;
    }
    seenIds.add(patientId);
    return {
      ...p,
      id: patientId,
    };
  });
};

export default function App() {
  // Current Month: format YYYY-MM
  const now = new Date();
  const defaultMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  const [selectedMonth, setSelectedMonth] = useState<string>(defaultMonth);

  // Auth State
  const [user, setUser] = useState<User | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  // Patient Records State (Dikosongkan dari data bawaan sistem, baris rekapitulasi, serta ID terduplikasi)
  const [patients, setPatients] = useState<PatientRecord[]>(() => {
    try {
      const saved = localStorage.getItem('dialysis_patients_data');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          // Bersihkan secara otomatis pasien demo bawaan sistem serta baris total/rekapitulasi & rapikan ID
          const cleaned = ensureUniquePatientIds(
            parsed
              .filter((p: PatientRecord) => !isSummaryOrHeaderRow(p.name, p.noRm) && !isDemoPatient(p))
              .map((p: PatientRecord) => ({
                ...p,
                clinicalNotes: isRemovedDemoNote(p.clinicalNotes) ? '' : p.clinicalNotes,
              }))
          );
          localStorage.setItem('dialysis_patients_data', JSON.stringify(cleaned));
          return cleaned;
        }
      }
    } catch (e) {
      console.error('Error loading saved patients:', e);
    }
    return [];
  });

  // Bersihkan data pasien dari baris total/rekapitulasi, data demo bawaan sistem, serta jamin ID unik
  useEffect(() => {
    setPatients((prev) => {
      let changed = false;
      const seenIds = new Set<string>();
      const filtered = prev
        .filter((p) => {
          if (isSummaryOrHeaderRow(p.name, p.noRm) || isDemoPatient(p)) {
            changed = true;
            return false;
          }
          return true;
        })
        .map((p, idx) => {
          let currentId = (p.id || '').trim();
          let note = p.clinicalNotes;
          if (isRemovedDemoNote(note)) {
            changed = true;
            note = '';
          }
          if (!currentId || seenIds.has(currentId)) {
            changed = true;
            const cleanRm = (p.noRm || `pat${idx}`).replace(/[^a-zA-Z0-9_-]/g, '');
            currentId = `${currentId || 'pat'}-${cleanRm}-${idx}-${Math.random().toString(36).substring(2, 7)}`;
          }
          seenIds.add(currentId);

          // Sinkronisasi status Cek Hb Pilihan: HANYA nilai Hb <= 8.9 mg/dl yang masuk Cek Hb Pilihan
          const refHb = (typeof p.prevHbValue === 'number' && p.prevHbValue > 0)
            ? Number(p.prevHbValue.toFixed(1))
            : (p.hbValue > 0 ? Number(p.hbValue.toFixed(1)) : undefined);
          
          let isSelective = p.isSelectiveHb;
          let labSchedule = p.labSchedule;

          if (refHb !== undefined) {
            if (refHb >= 9.0) {
              if (isSelective || labSchedule?.isSelectiveHb || labSchedule?.testType?.includes('Pilihan')) {
                changed = true;
                isSelective = false;
                if (labSchedule) {
                  labSchedule = {
                    ...labSchedule,
                    testType: 'Rutin Hb (Evaluasi EPO)',
                    isSelectiveHb: false,
                    selectiveReason: undefined,
                  };
                }
              }
            } else if (refHb <= 8.9) {
              if (labSchedule?.testType && (labSchedule.testType === 'Cek Hb Pilihan (Hb < 9.0)' || (labSchedule.testType as any) === 'Cek Hb Pilihan (Hb ≤ 9.0)')) {
                changed = true;
                labSchedule = {
                  ...labSchedule,
                  testType: 'Cek Hb Pilihan (Hb ≤ 8.9)',
                };
              }
            }
          }

          // Sinkronisasi alokasi Terapi EPO: Penjadwalan otomatis pada Pertemuan HD Kedua
          let weeks = p.weeks;
          let dailyRecords = p.dailyRecords;
          const monthInfo = getMonthDaysInfo(selectedMonth);
          const sessions = getPatientMonthHDSessions(p, monthInfo);
          const secondHdDate = sessions.length >= 2 ? sessions[1] : (sessions[0] || null);

          // Bersihkan status 'Belum' yang tidak sengaja tercatat di pertemuan pertama (sesi Cek Hb)
          // Kecuali untuk pasien HD 1x/2 minggu yang kategori rekomendasi EPO_4X_2000 (sesi 1 terjadwal EPO 2000)
          if (sessions.length >= 2 && dailyRecords && !(p.hdFrequency === '1 kali / 2 minggu' && p.recommendation?.category === 'EPO_4X_2000')) {
            const firstSessionDate = sessions[0].dateNumber;
            if (dailyRecords[firstSessionDate]?.status === 'Belum' && !dailyRecords[firstSessionDate]?.postponedFromDate) {
              changed = true;
              const newDaily = { ...dailyRecords };
              delete newDaily[firstSessionDate];
              dailyRecords = newDaily;
            }
          }

          // Sinkronisasi rekomendasi pasien HD 1x/2 minggu
          let recommendation = p.recommendation;
          if (p.hdFrequency === '1 kali / 2 minggu' && recommendation?.category === 'EPO_4X_2000' && recommendation.totalEpoVials !== 2) {
            changed = true;
            recommendation = calculateClinicalRecommendation(p.hbValue, p.hdFrequency);
          }

          if (p.recommendation?.category === 'EPO_1X_2000') {
            if (weeks) {
              const w2Status = weeks.week2?.status;
              const w3Status = weeks.week3?.status;
              const w4Status = weeks.week4?.status;
              if (
                (w2Status === 'Belum' || (weeks.week2?.doseIU && weeks.week2.doseIU > 0)) ||
                (w3Status === 'Belum' || (weeks.week3?.doseIU && weeks.week3.doseIU > 0)) ||
                (w4Status === 'Belum' || (weeks.week4?.doseIU && weeks.week4.doseIU > 0)) ||
                (secondHdDate && weeks.week1?.plannedDate !== secondHdDate.dateString)
              ) {
                changed = true;
                weeks = {
                  ...weeks,
                  week1: { 
                    weekNumber: 1, 
                    status: weeks.week1?.status || 'Belum', 
                    doseIU: 2000, 
                    plannedDate: secondHdDate ? secondHdDate.dateString : weeks.week1?.plannedDate,
                    notes: 'Otomatis dijadwalkan pada pertemuan HD kedua di bulan terkait',
                  },
                  week2: w2Status === 'Diberikan' ? weeks.week2 : { weekNumber: 2, status: 'Tidak Ada Jadwal', doseIU: 0 },
                  week3: w3Status === 'Diberikan' ? weeks.week3 : { weekNumber: 3, status: 'Tidak Ada Jadwal', doseIU: 0 },
                  week4: w4Status === 'Diberikan' ? weeks.week4 : { weekNumber: 4, status: 'Tidak Ada Jadwal', doseIU: 0 },
                };
              }
            }

            // Bersihkan sisa 'Belum' duplikat di dailyRecords
            if (dailyRecords) {
              const hasGiven = Object.values(dailyRecords).some((r) => r.status === 'Diberikan');
              const belumKeys = Object.entries(dailyRecords)
                .filter(([_, r]) => r.status === 'Belum' && !r.postponedFromDate)
                .map(([k]) => parseInt(k, 10));

              if (hasGiven && belumKeys.length > 0) {
                changed = true;
                const newDaily = { ...dailyRecords };
                belumKeys.forEach((k) => delete newDaily[k]);
                dailyRecords = newDaily;
              } else if (belumKeys.length > 1) {
                changed = true;
                const newDaily = { ...dailyRecords };
                belumKeys.slice(1).forEach((k) => delete newDaily[k]);
                dailyRecords = newDaily;
              }
            }
          } else if (p.recommendation?.category === 'EPO_4X_2000') {
            if (p.hdFrequency === '1 kali / 2 minggu') {
              const firstSession = sessions[0] || null;
              const secondSession = sessions.length >= 2 ? sessions[1] : null;
              if (weeks && firstSession && secondSession) {
                const firstWeekNum = firstSession.dateNumber <= 7 ? 1 : 2;
                const secondWeekNum = secondSession.dateNumber <= 21 ? 3 : 4;
                const w1Key = `week${firstWeekNum}` as 'week1' | 'week2';
                const w2Key = `week${secondWeekNum}` as 'week3' | 'week4';
                if (weeks[w1Key]?.plannedDate !== firstSession.dateString || weeks[w2Key]?.plannedDate !== secondSession.dateString) {
                  changed = true;
                  const newWeeks: PatientRecord['weeks'] = {
                    week1: { weekNumber: 1, status: 'Tidak Ada Jadwal', doseIU: 0 },
                    week2: { weekNumber: 2, status: 'Tidak Ada Jadwal', doseIU: 0 },
                    week3: { weekNumber: 3, status: 'Tidak Ada Jadwal', doseIU: 0 },
                    week4: { weekNumber: 4, status: 'Tidak Ada Jadwal', doseIU: 0 },
                  };
                  newWeeks[w1Key] = {
                    weekNumber: firstWeekNum as 1 | 2,
                    plannedDate: firstSession.dateString,
                    status: weeks[w1Key]?.status || 'Belum',
                    doseIU: 2000,
                    notes: 'Terjadwal EPO Sesi HD Pertama (1x/2 Minggu)',
                  };
                  newWeeks[w2Key] = {
                    weekNumber: secondWeekNum as 3 | 4,
                    plannedDate: secondSession.dateString,
                    status: weeks[w2Key]?.status || 'Belum',
                    doseIU: 2000,
                    notes: 'Terjadwal EPO Sesi HD Kedua (1x/2 Minggu)',
                  };
                  weeks = newWeeks;
                }
              }
            } else if (weeks && secondHdDate && weeks.week1?.plannedDate !== secondHdDate.dateString && weeks.week1?.status !== 'Diberikan') {
              changed = true;
              weeks = {
                ...weeks,
                week1: {
                  ...weeks.week1,
                  plannedDate: secondHdDate.dateString,
                  notes: 'Otomatis dijadwalkan pada pertemuan HD kedua di bulan terkait',
                }
              };
            }
          }

          if (
            currentId !== p.id ||
            note !== p.clinicalNotes ||
            isSelective !== p.isSelectiveHb ||
            labSchedule !== p.labSchedule ||
            recommendation !== p.recommendation ||
            weeks !== p.weeks ||
            dailyRecords !== p.dailyRecords
          ) {
            changed = true;
            return { 
              ...p, 
              id: currentId, 
              clinicalNotes: note,
              isSelectiveHb: isSelective,
              labSchedule,
              recommendation,
              weeks,
              dailyRecords,
            };
          }
          return p;
        });
      if (changed || filtered.length !== prev.length) {
        try {
          localStorage.setItem('dialysis_patients_data', JSON.stringify(filtered));
        } catch (e) {
          // ignore
        }
        return filtered;
      }
      return prev;
    });
  }, []);

  // Spreadsheet Config State - Tetapkan URL resmi Google Sheet & Apps Script
  const [spreadsheetConfig, setSpreadsheetConfig] = useState<SpreadsheetConfig>(() => {
    const DEFAULT_CONFIG: SpreadsheetConfig = {
      spreadsheetId: OFFICIAL_SPREADSHEET_ID,
      spreadsheetUrl: OFFICIAL_SPREADSHEET_URL,
      sheetName: 'Senin-Kamis, Selasa-Jumat, Rabu-Sabtu',
      appsScriptUrl: DEFAULT_APPS_SCRIPT_URL,
      syncMode: 'appsscript',
      lastSyncedAt: new Date().toISOString(),
    };
    try {
      const saved = localStorage.getItem('dialysis_spreadsheet_config');
      if (saved) {
        const parsed = JSON.parse(saved);
        const merged: SpreadsheetConfig = {
          ...DEFAULT_CONFIG,
          ...parsed,
          spreadsheetId: OFFICIAL_SPREADSHEET_ID,
          spreadsheetUrl: OFFICIAL_SPREADSHEET_URL,
          appsScriptUrl: DEFAULT_APPS_SCRIPT_URL, // Tetapkan URL resmi Google Apps Script
        };
        localStorage.setItem('dialysis_spreadsheet_config', JSON.stringify(merged));
        return merged;
      }
    } catch (e) {
      console.error('Error loading saved spreadsheet config:', e);
    }
    return DEFAULT_CONFIG;
  });

  // Pastikan URL Apps Script resmi selalu tersimpan aktif
  useEffect(() => {
    try {
      const saved = localStorage.getItem('dialysis_spreadsheet_config');
      const parsed = saved ? JSON.parse(saved) : {};
      if (parsed.appsScriptUrl !== DEFAULT_APPS_SCRIPT_URL) {
        const updated = {
          ...parsed,
          spreadsheetId: OFFICIAL_SPREADSHEET_ID,
          spreadsheetUrl: OFFICIAL_SPREADSHEET_URL,
          appsScriptUrl: DEFAULT_APPS_SCRIPT_URL,
          syncMode: 'appsscript',
        };
        localStorage.setItem('dialysis_spreadsheet_config', JSON.stringify(updated));
      }
    } catch (e) {
      // ignore
    }
  }, []);

  // Modals state
  const [isPatientModalOpen, setIsPatientModalOpen] = useState(false);
  const [editingPatient, setEditingPatient] = useState<PatientRecord | null>(null);
  const [isBulkImportModalOpen, setIsBulkImportModalOpen] = useState(false);
  const [isMonthlyHbModalOpen, setIsMonthlyHbModalOpen] = useState(false);
  const [isSyncModalOpen, setIsSyncModalOpen] = useState(false);
  const [isCalculatorModalOpen, setIsCalculatorModalOpen] = useState(false);
  const [isPrintModalOpen, setIsPrintModalOpen] = useState(false);
  const [isEpoSchedulePrintModalOpen, setIsEpoSchedulePrintModalOpen] = useState(false);
  const [isLabScheduleModalOpen, setIsLabScheduleModalOpen] = useState(false);
  const [isExportExcelModalOpen, setIsExportExcelModalOpen] = useState(false);
  const [isWorkflowGuideOpen, setIsWorkflowGuideOpen] = useState(false);
  const [exportExcelTarget, setExportExcelTarget] = useState<HDDaySchedule | 'ALL'>('ALL');

  // Syncing status & notifications
  const [isSyncing, setIsSyncing] = useState(false);
  const [toastNotification, setToastNotification] = useState<{
    type: 'success' | 'error' | 'info';
    message: string;
  } | null>(null);

  // Auto-hide toast notification
  useEffect(() => {
    if (toastNotification) {
      const timer = setTimeout(() => {
        setToastNotification(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [toastNotification]);

  // Save patients to localStorage on change
  useEffect(() => {
    try {
      localStorage.setItem('dialysis_patients_data', JSON.stringify(patients));
      if (selectedMonth && patients.length > 0) {
        localStorage.setItem('dialysis_patients_data_' + selectedMonth, JSON.stringify(patients));
      }
    } catch (e) {
      console.error('Failed to save patients to localStorage', e);
    }
  }, [patients, selectedMonth]);

  // Initialize Firebase Auth listener
  useEffect(() => {
    const unsubscribe = initAuth(
      (currentUser) => {
        setUser(currentUser);
      },
      () => {
        setUser(null);
      }
    );
    return () => unsubscribe();
  }, []);

  // Google Sign In Handler
  const handleSignIn = async () => {
    try {
      setIsLoggingIn(true);
      const res = await googleSignIn();
      if (res) {
        setUser(res.user);
        setToastNotification({
          type: 'success',
          message: `Berhasil masuk dengan akun Google: ${res.user.displayName || res.user.email}`,
        });
      }
    } catch (err: any) {
      console.error(err);
      setToastNotification({
        type: 'error',
        message: err.message || 'Gagal masuk dengan akun Google. Periksa koneksi atau izin popup.',
      });
    } finally {
      setIsLoggingIn(false);
    }
  };

  // Google Sign Out Handler
  const handleSignOut = async () => {
    await logout();
    setUser(null);
    setToastNotification({
      type: 'info',
      message: 'Anda telah keluar dari akun Google.',
    });
  };

  // Save Spreadsheet Config
  const handleSaveSpreadsheetConfig = (config: SpreadsheetConfig) => {
    setSpreadsheetConfig(config);
    try {
      localStorage.setItem('dialysis_spreadsheet_config', JSON.stringify(config));
    } catch (e) {
      console.error('Failed to save spreadsheet config', e);
    }
  };

  // Pull data from Google Sheet
  const handlePullFromSheet = async (spreadsheetId: string) => {
    const token = await getAccessToken();
    if (!token) {
      throw new Error('Silakan masuk dengan akun Google terlebih dahulu untuk menarik data.');
    }

    setIsSyncing(true);
    try {
      const importedPatients = await readPatientsFromSheet(spreadsheetId, token, selectedMonth);
      if (importedPatients.length === 0) {
        setToastNotification({
          type: 'info',
          message: 'Sheet terhubung, namun belum ada baris pasien yang ditemukan pada spreadsheet.',
        });
      } else {
        const uniqueImported = ensureUniquePatientIds(importedPatients);
        setPatients(uniqueImported);
        setToastNotification({
          type: 'success',
          message: `Berhasil mengimpor ${uniqueImported.length} pasien dari Google Sheet!`,
        });
      }

      handleSaveSpreadsheetConfig({
        spreadsheetId,
        sheetName: 'Alokasi_EPO_HD',
        lastSyncedAt: new Date().toISOString(),
      });
    } finally {
      setIsSyncing(false);
    }
  };

  // Push data to Google Sheet
  const handlePushToSheet = async (spreadsheetId: string) => {
    const token = await getAccessToken();
    if (!token) {
      throw new Error('Silakan masuk dengan akun Google terlebih dahulu untuk menyimpan ke Sheet.');
    }

    setIsSyncing(true);
    try {
      await pushPatientsToSheet(spreadsheetId, patients, token, selectedMonth);
      handleSaveSpreadsheetConfig({
        spreadsheetId,
        sheetName: 'Senin-Kamis, Selasa-Jumat, Rabu-Sabtu',
        lastSyncedAt: new Date().toISOString(),
      });
      setToastNotification({
        type: 'success',
        message: `Berhasil menyimpan & mensinkronisasikan ${patients.length} pasien ke 3 sheet jadwal di Google Sheet!`,
      });
    } finally {
      setIsSyncing(false);
    }
  };

  // Pull via Apps Script (Tanpa Login Google)
  const handlePullViaAppsScript = async (url: string) => {
    setIsSyncing(true);
    try {
      const imported = await pullViaAppsScript(url, selectedMonth);
      if (imported.length === 0) {
        if (patients.length > 0) {
          setToastNotification({
            type: 'info',
            message: `Terhubung ke Apps Script! Spreadsheet Google Sheets Anda saat ini masih kosong. Klik tombol 'Kirim ke Sheet (Push)' untuk mengirimkan ${patients.length} data pasien dari aplikasi ke spreadsheet Anda.`,
          });
        } else {
          setToastNotification({
            type: 'info',
            message: 'Terhubung ke Apps Script! Spreadsheet Google Sheets Anda saat ini belum memiliki baris data pasien.',
          });
        }
      } else {
        const uniqueImported = ensureUniquePatientIds(imported);
        setPatients(uniqueImported);
        setToastNotification({
          type: 'success',
          message: `Berhasil membaca ${uniqueImported.length} data pasien via Apps Script tanpa login!`,
        });
      }
    } finally {
      setIsSyncing(false);
    }
  };

  // Push via Apps Script (Tanpa Login Google)
  const handlePushViaAppsScript = async (url: string, scheduleScope?: LabScheduleScope) => {
    setIsSyncing(true);
    try {
      await pushViaAppsScript(url, patients, selectedMonth, undefined, scheduleScope);
      setToastNotification({
        type: 'success',
        message: `Berhasil mengirim ${patients.length} data pasien ke 3 sheet jadwal dan matriks Cek HB via Apps Script!`,
      });
    } finally {
      setIsSyncing(false);
    }
  };

  // Quick Push to Google Sheets (Per Sheet / Semua Sheet) langsung dari tombol tabel
  const handleQuickPushToSheet = async (targetSchedule?: HDDaySchedule, scheduleScope?: LabScheduleScope) => {
    const url = spreadsheetConfig?.appsScriptUrl || DEFAULT_APPS_SCRIPT_URL;
    setIsSyncing(true);
    try {
      if (url) {
        await pushViaAppsScript(url, patients, selectedMonth, targetSchedule, scheduleScope);
        setToastNotification({
          type: 'success',
          message: targetSchedule
            ? `Berhasil mengirim data yang diinputkan (${targetSchedule}) ke sheet di Google Sheets!`
            : `Berhasil mensinkronisasikan seluruh jadwal dan matriks Cek HB ke Google Sheets!`,
        });
      } else if (spreadsheetConfig?.spreadsheetId) {
        const token = await getAccessToken();
        if (token) {
          await pushPatientsToSheet(spreadsheetConfig.spreadsheetId, patients, token, selectedMonth, targetSchedule, scheduleScope);
          setToastNotification({
            type: 'success',
            message: targetSchedule
              ? `Berhasil mengirim data ${targetSchedule} ke tab spreadsheet!`
              : `Berhasil mensinkronisasikan seluruh data dan matriks Cek HB ke spreadsheet!`,
          });
        } else {
          setIsSyncModalOpen(true);
        }
      } else {
        setIsSyncModalOpen(true);
      }
    } catch (err: any) {
      console.error(err);
      setToastNotification({
        type: 'error',
        message: 'Gagal mengirim data ke Google Sheet. Periksa koneksi internet atau URL Apps Script.',
      });
    } finally {
      setIsSyncing(false);
    }
  };

  // Import from CSV (Tanpa Login Google)
  const handleImportCsv = (csvText: string) => {
    const imported = importFromCSV(csvText, selectedMonth);
    if (imported.length > 0) {
      const uniqueImported = ensureUniquePatientIds(imported);
      setPatients(uniqueImported);
      setToastNotification({
        type: 'success',
        message: `Berhasil memuat ${uniqueImported.length} data pasien dari file CSV!`,
      });
    } else {
      setToastNotification({
        type: 'error',
        message: 'File CSV tidak berisi data pasien yang valid.',
      });
    }
  };

  // Add or Edit Patient
  const handleSavePatient = (patientData: Partial<PatientRecord>) => {
    if (editingPatient) {
      // Update existing
      setPatients((prev) => {
        const updated = prev.map((p) => {
          if (p.id !== editingPatient.id) return p;
          const merged: PatientRecord = {
            ...p,
            ...patientData,
            id: p.id,
            updatedAt: new Date().toISOString(),
          };
          return merged;
        });

        try {
          localStorage.setItem('dialysis_patients_data_' + selectedMonth, JSON.stringify(updated));
          localStorage.setItem('dialysis_patients_data', JSON.stringify(updated));
        } catch (e) {
          console.error('Failed to save updated patient:', e);
        }

        return updated;
      });

      setToastNotification({
        type: 'success',
        message: `Data pasien ${patientData.name || editingPatient.name} berhasil diperbarui.`,
      });
    } else {
      // Create new
      const newPatient: PatientRecord = {
        id: `pat-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
        noRm: patientData.noRm || `RM-${Math.floor(10000 + Math.random() * 90000)}`,
        name: patientData.name || 'Pasien Baru',
        age: patientData.age,
        gender: patientData.gender || 'L',
        hdFrequency: patientData.hdFrequency || '2 kali dalam satu minggu',
        singleDay: patientData.singleDay,
        lastHdDate: patientData.lastHdDate,
        scheduleDay: patientData.scheduleDay || 'Senin - Kamis',
        scheduleShift: patientData.scheduleShift || 'Shift 1 (Pagi)',
        hbValue: typeof patientData.hbValue === 'number' ? patientData.hbValue : 0,
        hbDate: patientData.hbDate || getFirstHDDateOfMonth(selectedMonth, patientData.scheduleDay || 'Senin - Kamis', patientData.singleDay, patientData.hdFrequency, patientData.lastHdDate).dateString,
        monthPeriod: selectedMonth,
        recommendation: patientData.recommendation!,
        weeks: patientData.weeks!,
        overallStatus: patientData.overallStatus || 'Berjalan',
        clinicalNotes: patientData.clinicalNotes,
        doctorInCharge: patientData.doctorInCharge || 'dr. Sp.PD-KGH',
        updatedAt: new Date().toISOString(),
      };
      setPatients((prev) => {
        const updated = [newPatient, ...prev];
        try {
          localStorage.setItem('dialysis_patients_data_' + selectedMonth, JSON.stringify(updated));
          localStorage.setItem('dialysis_patients_data', JSON.stringify(updated));
        } catch (e) {}
        return updated;
      });
      setToastNotification({
        type: 'success',
        message: `Pasien baru ${newPatient.name} ditambahkan ke alokasi bulan ini.`,
      });
    }
    setEditingPatient(null);
  };

  // Navigasi & Sinkronisasi Perpindahan Bulan:
  // Sesuai Permintaan Inti:
  // 1. Pada sesi CEK HB SELURUH PASIEN:
  //    Contoh: input CEK HB SELURUH PASIEN untuk BULAN FEBRUARI 2026 di bulan JANUARI 2026,
  //    maka seluruh pasien Tn A, Tn B, Tn C dan Tn D pada bulan Februari 2026 nilai HB nya berstatus 0 (menunggu hasil lab).
  // 2. Pada sesi CEK HB PILIHAN:
  //    Pasien yang tidak terjadwal otomatis terisi nilai HB nya sesuai dengan nilai bulan sebelumnya karena tidak melakukan CEK HB.
  //    Contoh: input CEK HB PILIHAN untuk BULAN FEBRUARI 2026 di bulan JANUARI 2026,
  //    maka pada bulan Februari 2026 Tn A otomatis terisi nilai HB nya (12.0),
  //    sedangkan Tn B, Tn C, dan Tn D terisi nilai 0 (Menunggu Hasil Lab).
  const handleMonthChange = (newMonth: string) => {
    // 1. Simpan data pasien bulan aktif saat ini sebelum berpindah
    try {
      localStorage.setItem('dialysis_patients_data_' + selectedMonth, JSON.stringify(patients));
      localStorage.setItem('dialysis_selected_month', newMonth);
    } catch (e) {}

    setSelectedMonth(newMonth);

    // 2. Cek apakah sudah ada data tersimpan khusus untuk newMonth
    let savedForNewMonth: PatientRecord[] | null = null;
    try {
      const raw = localStorage.getItem('dialysis_patients_data_' + newMonth);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed) && parsed.length > 0) {
          savedForNewMonth = ensureUniquePatientIds(parsed);
        }
      }
    } catch (e) {}

    if (savedForNewMonth && savedForNewMonth.length > 0) {
      setPatients(savedForNewMonth);
      try {
        localStorage.setItem('dialysis_patients_data', JSON.stringify(savedForNewMonth));
      } catch (e) {}
      return;
    }

    // 3. Jika belum ada data tersimpan untuk newMonth, generate dari data pasien saat ini
    // Sesuai mode sesi yang terkonfigurasi untuk newMonth ('SELURUH' atau 'PILIHAN'):
    let scope: 'SELURUH' | 'PILIHAN' = 'PILIHAN';
    try {
      const stored = localStorage.getItem('epocare_lab_schedule_mode_' + newMonth) || localStorage.getItem('epocare_lab_schedule_mode');
      if (stored === 'SELURUH' || stored === 'PILIHAN') {
        scope = stored;
      }
    } catch (e) {}

    setPatients((prev) => {
      const generated: PatientRecord[] = prev.map((p) => {
        const targetFirstDate = getFirstHDDateOfMonth(
          newMonth,
          p.scheduleDay,
          p.singleDay,
          p.hdFrequency,
          p.lastHdDate
        ).dateString;

        const priorHb = (typeof p.hbValue === 'number' && p.hbValue > 0)
          ? p.hbValue
          : (typeof p.prevHbValue === 'number' && p.prevHbValue > 0 ? p.prevHbValue : 0);

        const isCandidate = (priorHb > 0 && priorHb < 9.0) || isSelectiveHbCandidate(p);

        if (scope === 'SELURUH') {
          // ATURAN 1 (CEK HB SELURUH PASIEN):
          // Seluruh pasien (Tn A, Tn B, Tn C, Tn D) berstatus 0 (menunggu hasil lab)
          const reco = calculateClinicalRecommendation(0, p.hdFrequency);
          const weeks = generateDefaultWeeks('MENUNGGU_HASIL_LAB', newMonth, p.hdFrequency, p.scheduleDay, p.singleDay, p.lastHdDate);
          return {
            ...p,
            hbValue: 0,
            prevHbValue: priorHb > 0 ? priorHb : p.prevHbValue,
            hbDate: targetFirstDate,
            monthPeriod: newMonth,
            recommendation: reco,
            overallStatus: 'Menunggu',
            weeks,
            dailyRecords: {},
            labSchedule: {
              scheduledDate: targetFirstDate,
              testType: 'Rutin Hb (Evaluasi EPO)',
              status: 'Terjadwal',
              notes: 'Evaluasi rutin awal bulan (Cek Hb Seluruh Pasien)',
              isSelectiveHb: false,
            },
            updatedAt: new Date().toISOString(),
          };
        } else {
          // ATURAN 2 (CEK HB PILIHAN):
          if (priorHb >= 9.0 && !isCandidate) {
            // Pasien tidak terjadwal (Tn A) otomatis terisi nilai bulan sebelumnya (12.0)
            const reco = calculateClinicalRecommendation(priorHb, p.hdFrequency);
            const weeks = generateDefaultWeeks(reco.category, newMonth, p.hdFrequency, p.scheduleDay, p.singleDay, p.lastHdDate);
            return {
              ...p,
              hbValue: priorHb,
              prevHbValue: priorHb,
              hbDate: targetFirstDate,
              monthPeriod: newMonth,
              recommendation: reco,
              overallStatus: reco.category === 'TRANSFUSI_2_RAWAT_INAP' ? 'Perlu Perhatian' : 'Berjalan',
              weeks,
              dailyRecords: {},
              labSchedule: {
                scheduledDate: '',
                testType: 'Rutin Hb (Evaluasi EPO)',
                status: 'Tidak Ada Jadwal',
                notes: `Tidak dijadwalkan Cek Hb bulan depan (Hb acuan ${priorHb.toFixed(1)} g/dL)`,
                isSelectiveHb: false,
              },
              updatedAt: new Date().toISOString(),
            };
          } else {
            // Pasien terjadwal (Tn B, Tn C, Tn D) berstatus 0 (menunggu hasil lab)
            const reco = calculateClinicalRecommendation(0, p.hdFrequency);
            const weeks = generateDefaultWeeks('MENUNGGU_HASIL_LAB', newMonth, p.hdFrequency, p.scheduleDay, p.singleDay, p.lastHdDate);
            return {
              ...p,
              hbValue: 0,
              prevHbValue: priorHb > 0 ? priorHb : p.prevHbValue,
              hbDate: targetFirstDate,
              monthPeriod: newMonth,
              recommendation: reco,
              overallStatus: 'Menunggu',
              weeks,
              dailyRecords: {},
              labSchedule: {
                scheduledDate: targetFirstDate,
                testType: 'Cek Hb Pilihan (Hb ≤ 8.9)',
                status: 'Terjadwal',
                notes: `⭐ Cek Hb Pilihan: Evaluasi anemia (Hb acuan ${priorHb > 0 ? priorHb.toFixed(1) : '< 9.0'} g/dL)`,
                isSelectiveHb: true,
                selectiveReason: `Nilai Hb acuan < 9.0 mg/dL (${priorHb > 0 ? priorHb.toFixed(1) : '< 9.0'})`,
              },
              updatedAt: new Date().toISOString(),
            };
          }
        }
      });

      try {
        localStorage.setItem('dialysis_patients_data_' + newMonth, JSON.stringify(generated));
        localStorage.setItem('dialysis_patients_data', JSON.stringify(generated));
      } catch (e) {}

      return generated;
    });
  };

  // Delete Patient
  const handleDeletePatient = (patientId: string) => {
    const target = patients.find((p) => p.id === patientId);
    const confirmed = window.confirm(
      `Apakah Anda yakin ingin menghapus data pasien ${target?.name || 'ini'} dari jadwal alokasi?`
    );
    if (!confirmed) return;

    setPatients((prev) => prev.filter((p) => p.id !== patientId));
    setToastNotification({
      type: 'info',
      message: `Pasien ${target?.name || ''} telah dihapus dari daftar.`,
    });
  };

  // Update Week Status (M1, M2, M3, M4)
  const handleUpdateWeekStatus = (
    patientId: string,
    weekKey: 'week1' | 'week2' | 'week3' | 'week4',
    newStatus: DoseStatus,
    nurseName?: string
  ) => {
    setPatients((prev) =>
      prev.map((patient) => {
        if (patient.id !== patientId) return patient;

        const currentWeek = patient.weeks[weekKey];
        const updatedWeek = {
          ...currentWeek,
          status: newStatus,
          administeredAt: newStatus === 'Diberikan' 
            ? new Date().toLocaleString('id-ID', { dateStyle: 'short', timeStyle: 'short' })
            : undefined,
          administeredBy: newStatus === 'Diberikan' ? (nurseName || 'Perawat HD') : undefined,
        };

        const updatedWeeks = {
          ...patient.weeks,
          [weekKey]: updatedWeek,
        };

        // Recalculate overall status
        const allActiveWeeks = ['week1', 'week2', 'week3', 'week4'].filter(
          (k) => (updatedWeeks as any)[k].status !== 'Tidak Ada Jadwal'
        );
        const allDone = allActiveWeeks.every(
          (k) => (updatedWeeks as any)[k].status === 'Diberikan'
        );

        let overallStatus = patient.overallStatus;
        if (patient.recommendation.category === 'TRANSFUSI_2_RAWAT_INAP') {
          overallStatus = 'Perlu Perhatian';
        } else if (allDone && allActiveWeeks.length > 0) {
          overallStatus = 'Selesai';
        } else {
          overallStatus = 'Berjalan';
        }

        return {
          ...patient,
          weeks: updatedWeeks,
          overallStatus,
          updatedAt: new Date().toISOString(),
        };
      })
    );
  };

  // Update Date-level action (Injeksi Diberikan, Tunda ke Hari Kedua, Reset, dll)
  const handleUpdateDateAction = (
    patientId: string,
    dateNumber: number,
    newStatus: DoseStatus,
    nurseName?: string,
    notes?: string
  ) => {
    setPatients((prev) =>
      prev.map((patient) => {
        if (patient.id !== patientId) return patient;

        const monthInfo = getMonthDaysInfo(selectedMonth);
        const dayInfo = monthInfo.days[dateNumber - 1];
        if (!dayInfo) return patient;

        const isOnceWeekly = patient.hdFrequency === '1 kali dalam satu minggu';
        const { isPrimary, pairedDate } = getPairedHDDate(
          dateNumber,
          dayInfo.dayOfWeek,
          patient.scheduleDay,
          monthInfo.daysInMonth,
          patient.hdFrequency,
          patient.singleDay
        );

        const sessions = getPatientMonthHDSessions(patient, monthInfo);
        const currIdx = sessions.findIndex((s) => s.dateNumber === dateNumber);
        const nextSession = currIdx >= 0 && currIdx < sessions.length - 1 ? sessions[currIdx + 1] : null;
        const targetNextDate = nextSession ? nextSession.dateNumber : pairedDate;

        const currentRecords = { ...(patient.dailyRecords || {}) };

        if (newStatus === 'Tunda') {
          if (targetNextDate) {
            const isBiweekly = patient.hdFrequency === '1 kali / 2 minggu';
            const defaultNote = isBiweekly
              ? `Ditunda tgl ${dateNumber}, dialihkan ke sesi HD 14 hari berikutnya tgl ${targetNextDate}`
              : `Ditunda tgl ${dateNumber}, dialihkan ke sesi HD berikutnya tgl ${targetNextDate}`;
            currentRecords[dateNumber] = {
              status: 'Tunda',
              postponedToDate: targetNextDate,
              notes: notes || defaultNote,
            };
            currentRecords[targetNextDate] = {
              status: 'Belum',
              postponedFromDate: dateNumber,
              notes: `Jadwal Pengganti dari penundaan tgl ${dateNumber}`,
            };
            setToastNotification({
              type: 'info',
              message: isBiweekly
                ? `Pemberian EPO tgl ${dateNumber} Ditunda (❌) dan dialihkan ke sesi HD 14 hari berikutnya tgl ${targetNextDate}.`
                : `Pemberian EPO tgl ${dateNumber} Ditunda (❌) dan otomatis dialihkan ke sesi HD berikutnya tgl ${targetNextDate}.`,
            });
          } else {
            currentRecords[dateNumber] = {
              status: 'Tunda',
              notes: notes || 'Pemberian ditunda',
            };
            setToastNotification({
              type: 'info',
              message: `Pemberian EPO tgl ${dateNumber} Ditunda (❌).`,
            });
          }
        } else if (newStatus === 'Diberikan') {
          const nowStr = new Date().toLocaleString('id-ID', { dateStyle: 'short', timeStyle: 'short' });
          currentRecords[dateNumber] = {
            ...(currentRecords[dateNumber] || {}),
            status: 'Diberikan',
            administeredAt: nowStr,
            administeredBy: nurseName || 'Ns. Maya',
            notes: notes || undefined,
          };
          setToastNotification({
            type: 'success',
            message: `EPO Diberikan (✅) 2000 IU pada tanggal ${dateNumber}.`,
          });
        } else if (newStatus === 'Batal') {
          currentRecords[dateNumber] = {
            ...(currentRecords[dateNumber] || {}),
            status: 'Batal',
            notes: notes || 'Dibatalkan',
          };
          setToastNotification({
            type: 'info',
            message: `Pemberian EPO tgl ${dateNumber} dibatalkan.`,
          });
        } else if (newStatus === 'Belum') {
          // Reset status tanggal
          const existing = currentRecords[dateNumber];
          if (existing?.postponedToDate) {
            delete currentRecords[existing.postponedToDate];
          }
          if (existing?.postponedFromDate) {
            delete currentRecords[existing.postponedFromDate];
          }
          delete currentRecords[dateNumber];
          setToastNotification({
            type: 'info',
            message: `Status tanggal ${dateNumber} di-reset ke 2000 (Terjadwal EPO).`,
          });
        }

        // Hitung total dosis diberikan dan sinkronkan dengan status overall
        let totalGiven = 0;
        for (let d = 1; d <= monthInfo.daysInMonth; d++) {
          const info = getDateDoseDisplayInfo({ ...patient, dailyRecords: currentRecords }, d, monthInfo);
          if (info.cellCode === '✅' || (info.cellCode as any) === 'P') totalGiven++;
        }

        let overallStatus = patient.overallStatus;
        if (patient.recommendation.category === 'TRANSFUSI_2_RAWAT_INAP') {
          overallStatus = 'Perlu Perhatian';
        } else if (totalGiven >= patient.recommendation.totalEpoVials && patient.recommendation.totalEpoVials > 0) {
          overallStatus = 'Selesai';
        } else if (totalGiven > 0) {
          overallStatus = 'Berjalan';
        }

        // Sinkronkan ke alokasi minggu (weeks) secara akurat
        const updatedWeeks = { ...patient.weeks };
        const originalSourceDate = currentRecords[dateNumber]?.postponedFromDate || dateNumber;
        const weekNum = getHDSessionWeek(originalSourceDate, monthInfo.daysInMonth, patient.scheduleDay, currentRecords[dateNumber]?.postponedFromDate);
        const weekKey = `week${weekNum}` as 'week1' | 'week2' | 'week3' | 'week4';
        if (updatedWeeks[weekKey]) {
          if (newStatus === 'Diberikan') {
            updatedWeeks[weekKey] = {
              ...updatedWeeks[weekKey],
              status: 'Diberikan',
              administeredAt: currentRecords[dateNumber]?.administeredAt,
              administeredBy: currentRecords[dateNumber]?.administeredBy,
              notes: currentRecords[dateNumber]?.postponedFromDate 
                ? `Diberikan pada Hari Kedua (tgl ${dateNumber}) pengalihan dari penundaan tgl ${currentRecords[dateNumber].postponedFromDate}`
                : undefined,
            };
          } else if (newStatus === 'Tunda') {
            updatedWeeks[weekKey] = {
              ...updatedWeeks[weekKey],
              status: 'Tunda',
              notes: pairedDate ? `Ditunda di Hari Awal (tgl ${dateNumber}), dialihkan ke Hari Kedua (tgl ${pairedDate})` : 'Ditunda',
            };
          } else if (newStatus === 'Belum') {
            updatedWeeks[weekKey] = {
              ...updatedWeeks[weekKey],
              status: 'Belum',
              notes: undefined,
              administeredAt: undefined,
              administeredBy: undefined,
            };
          }
        }

        return {
          ...patient,
          dailyRecords: currentRecords,
          weeks: updatedWeeks,
          overallStatus,
          updatedAt: new Date().toISOString(),
        };
      })
    );
  };

  // Bulk Import Patients Handler
  const handleImportPatients = (newPatients: PatientRecord[], mode: 'append' | 'replace') => {
    const sanitizedNew = ensureUniquePatientIds(newPatients);
    if (mode === 'replace') {
      setPatients(sanitizedNew);
      setToastNotification({
        type: 'success',
        message: `Berhasil mengganti seluruh data dengan ${sanitizedNew.length} pasien baru.`,
      });
    } else {
      // Append mode: update if same RM exists, or add new
      const incomingRmMap = new Map(sanitizedNew.map((p) => [p.noRm.toLowerCase(), p]));
      const updatedExisting = patients.map((p) => {
        const match = incomingRmMap.get(p.noRm.toLowerCase());
        if (match) {
          incomingRmMap.delete(p.noRm.toLowerCase());
          return match;
        }
        return p;
      });
      const remainingNew = Array.from(incomingRmMap.values());
      const combined = ensureUniquePatientIds([...remainingNew, ...updatedExisting]);
      setPatients(combined);
      setToastNotification({
        type: 'success',
        message: `Berhasil memasukkan ${sanitizedNew.length} pasien (${remainingNew.length} pasien baru, ${sanitizedNew.length - remainingNew.length} data diperbarui).`,
      });
    }
  };

  // Batch Update Hb Awal Bulan (Runtutan 2, 3, 4, 5)
  const handleSaveBatchHb = (
    updatedHbList: { id: string; hbValue: number; hbDate: string }[],
    targetMonth?: string,
    scope?: 'SELURUH' | 'PILIHAN'
  ) => {
    const effectiveMonth = targetMonth || selectedMonth;
    const effectiveScope = scope || 'PILIHAN';

    const isSingle = updatedHbList.length === 1;
    const targetPat = isSingle ? patients.find((p) => p.id === updatedHbList[0].id) : null;

    if (!isSingle) {
      try {
        localStorage.setItem('epocare_lab_schedule_mode_' + effectiveMonth, effectiveScope);
        localStorage.setItem('epocare_lab_schedule_mode', effectiveScope);
      } catch (e) {}
    }

    const updateMap = new Map(updatedHbList.map((item) => [item.id, item]));

    const updatePatientItem = (patient: PatientRecord): PatientRecord => {
      const update = updateMap.get(patient.id);
      if (!update) return patient;

      const validHb = update.hbValue;
      const reco = calculateClinicalRecommendation(validHb, patient.hdFrequency);
      const isCategoryChanged = reco.category !== patient.recommendation.category;
      const weeks = isCategoryChanged
        ? generateDefaultWeeks(
            reco.category,
            effectiveMonth,
            patient.hdFrequency,
            patient.scheduleDay,
            patient.singleDay,
            patient.lastHdDate
          )
        : patient.weeks;
      const dailyRecords = isCategoryChanged ? undefined : patient.dailyRecords;

      // PENTING (Runtutan 3 & 4):
      // Catat nilai Hb bulan sebelumnya sebagai riwayat prevHbValue jika nilai Hb diperbarui
      const prevHb = (patient.hbValue > 0 && patient.hbValue !== validHb)
        ? patient.hbValue 
        : (patient.prevHbValue !== undefined && patient.prevHbValue > 0 ? patient.prevHbValue : (patient.hbValue > 0 ? patient.hbValue : 0));

      const isSelective = (prevHb > 0 && prevHb < 9.0) || isSelectiveHbCandidate(patient);

      let updatedLabSchedule = patient.labSchedule;
      if (validHb > 0) {
        updatedLabSchedule = {
          ...(patient.labSchedule || {
            scheduledDate: update.hbDate || patient.hbDate,
            testType: isSelective ? 'Cek Hb Pilihan (Hb ≤ 8.9)' : 'Rutin Hb (Evaluasi EPO)',
            notes: '',
          }),
          status: 'Selesai',
          resultHb: validHb,
          completedAt: new Date().toISOString(),
        };
      } else if (effectiveScope === 'SELURUH') {
        updatedLabSchedule = {
          ...(patient.labSchedule || {
            scheduledDate: update.hbDate || patient.hbDate,
            testType: 'Rutin Hb (Evaluasi EPO)',
            notes: '',
          }),
          status: 'Terjadwal',
          resultHb: undefined,
        };
      } else if (effectiveScope === 'PILIHAN') {
        if (prevHb >= 9.0 && !isSelective) {
          updatedLabSchedule = {
            ...(patient.labSchedule || {
              scheduledDate: '',
              testType: 'Rutin Hb (Evaluasi EPO)',
              notes: '',
            }),
            status: 'Tidak Ada Jadwal',
            resultHb: undefined,
          };
        } else {
          updatedLabSchedule = {
            ...(patient.labSchedule || {
              scheduledDate: update.hbDate || patient.hbDate,
              testType: 'Cek Hb Pilihan (Hb ≤ 8.9)',
              notes: '',
            }),
            status: 'Terjadwal',
            resultHb: undefined,
          };
        }
      }

      return {
        ...patient,
        hbValue: validHb,
        prevHbValue: prevHb,
        hbDate: update.hbDate || patient.hbDate,
        monthPeriod: effectiveMonth,
        recommendation: reco,
        weeks,
        dailyRecords,
        labSchedule: updatedLabSchedule,
        overallStatus: validHb === 0 
          ? 'Menunggu' 
          : reco.category === 'TRANSFUSI_2_RAWAT_INAP' 
          ? 'Perlu Perhatian' 
          : 'Berjalan',
        updatedAt: new Date().toISOString(),
      };
    };

    if (effectiveMonth === selectedMonth) {
      setPatients((prev) => {
        const updated = prev.map(updatePatientItem);
        try {
          localStorage.setItem('dialysis_patients_data_' + effectiveMonth, JSON.stringify(updated));
          localStorage.setItem('dialysis_patients_data', JSON.stringify(updated));
        } catch (e) {}
        return updated;
      });
    } else {
      try {
        const raw = localStorage.getItem('dialysis_patients_data_' + effectiveMonth);
        const existing = raw ? JSON.parse(raw) : patients;
        const updated = (existing as PatientRecord[]).map(updatePatientItem);
        localStorage.setItem('dialysis_patients_data_' + effectiveMonth, JSON.stringify(updated));
      } catch (e) {}
    }

    if (isSingle && targetPat) {
      setToastNotification({
        type: 'success',
        message: `Nilai Hb ${updatedHbList[0].hbValue.toFixed(1)} g/dL untuk ${targetPat.name} berhasil disimpan tanpa merubah nilai pasien lain.`,
      });
    } else {
      const isSelectiveMode = effectiveScope === 'PILIHAN';
      setToastNotification({
        type: 'success',
        message: isSelectiveMode
          ? `Berhasil memperbarui data Cek HB Pilihan untuk bulan ${effectiveMonth}! Pasien stabil (≥ 9.0) terisi otomatis nilai bulan lalu, pasien terjadwal anemia (< 9.0) berstatus 0 (Menunggu Hasil Lab).`
          : `Berhasil memperbarui data Cek HB Seluruh Pasien untuk bulan ${effectiveMonth}! Seluruh pasien tanpa hasil lab berstatus 0 (Menunggu Hasil Lab).`,
      });
    }
  };

  // Simpan Jadwal Cek Lab Manual & Cek Hb Pilihan
  const handleSaveLabSchedules = (
    updatedSchedules: { id: string; labSchedule: LabSchedule; hbDate?: string; prevHbValue?: number; isSelectiveHb?: boolean }[],
    scope?: LabScheduleScope,
    targetMonth?: string
  ) => {
    const effectiveScope = scope || 'PILIHAN';
    const effectiveTargetMonth = targetMonth || getNextYearMonth(selectedMonth);

    try {
      localStorage.setItem('epocare_lab_schedule_mode_' + effectiveTargetMonth, effectiveScope);
      localStorage.setItem('epocare_lab_schedule_mode', effectiveScope);
    } catch (e) {}

    const scheduleMap = new Map(updatedSchedules.map((item) => [item.id, item]));

    // 1. Update status labSchedule di bulan berjalan
    setPatients((prev) => {
      const updated = prev.map((patient) => {
        const item = scheduleMap.get(patient.id);
        if (!item) return patient;

        return {
          ...patient,
          labSchedule: item.labSchedule,
          prevHbValue: item.prevHbValue !== undefined ? item.prevHbValue : patient.prevHbValue,
          isSelectiveHb: item.isSelectiveHb !== undefined ? item.isSelectiveHb : patient.isSelectiveHb,
          updatedAt: new Date().toISOString(),
        };
      });

      try {
        localStorage.setItem('dialysis_patients_data', JSON.stringify(updated));
        localStorage.setItem('dialysis_patients_data_' + selectedMonth, JSON.stringify(updated));
      } catch (e) {
        console.error('Error saving lab schedules to storage:', e);
      }

      return updated;
    });

    // 2. OTOMATIS GENERATE & SIMPAN DATA UNTUK BULAN SELANJUTNYA (effectiveTargetMonth):
    // Sesuai Permintaan Inti:
    // 1. Pada sesi CEK HB SELURUH PASIEN: Seluruh pasien Tn A, Tn B, Tn C dan Tn D pada bulan target nilai HB nya berstatus 0 (menunggu hasil lab)
    // 2. Pada sesi CEK HB PILIHAN: Pasien yang tidak terjadwal otomatis terisi nilai HB nya sesuai nilai bulan sebelumnya karena tidak melakukan CEK HB.
    //    Contoh: Tn A otomatis terisi nilai HB nya (12.0). Sedangkan Tn B, Tn C, dan Tn D terisi nilai 0 (Menunggu Hasil Lab)
    const targetMonthPatients: PatientRecord[] = patients.map((p) => {
      const targetFirstDate = getFirstHDDateOfMonth(
        effectiveTargetMonth,
        p.scheduleDay,
        p.singleDay,
        p.hdFrequency,
        p.lastHdDate
      ).dateString;

      const priorHb = (typeof p.hbValue === 'number' && p.hbValue > 0)
        ? p.hbValue
        : (typeof p.prevHbValue === 'number' && p.prevHbValue > 0 ? p.prevHbValue : 0);

      const isCandidate = (priorHb > 0 && priorHb < 9.0) || isSelectiveHbCandidate(p);

      if (effectiveScope === 'SELURUH') {
        // ATURAN 1 (CEK HB SELURUH PASIEN):
        // Seluruh pasien (Tn A, Tn B, Tn C, Tn D) berstatus 0 (menunggu hasil lab)
        const reco = calculateClinicalRecommendation(0, p.hdFrequency);
        const weeks = generateDefaultWeeks('MENUNGGU_HASIL_LAB', effectiveTargetMonth, p.hdFrequency, p.scheduleDay, p.singleDay, p.lastHdDate);
        return {
          ...p,
          hbValue: 0,
          prevHbValue: priorHb > 0 ? priorHb : p.prevHbValue,
          hbDate: targetFirstDate,
          monthPeriod: effectiveTargetMonth,
          recommendation: reco,
          overallStatus: 'Menunggu',
          weeks,
          dailyRecords: {},
          labSchedule: {
            scheduledDate: targetFirstDate,
            testType: 'Rutin Hb (Evaluasi EPO)',
            status: 'Terjadwal',
            notes: 'Evaluasi rutin awal bulan (Cek Hb Seluruh Pasien)',
            isSelectiveHb: false,
          },
          updatedAt: new Date().toISOString(),
        };
      } else {
        // ATURAN 2 (CEK HB PILIHAN):
        if (priorHb >= 9.0 && !isCandidate) {
          // Pasien tidak terjadwal (Tn A) otomatis terisi nilai bulan sebelumnya (12.0)
          const reco = calculateClinicalRecommendation(priorHb, p.hdFrequency);
          const weeks = generateDefaultWeeks(reco.category, effectiveTargetMonth, p.hdFrequency, p.scheduleDay, p.singleDay, p.lastHdDate);
          return {
            ...p,
            hbValue: priorHb,
            prevHbValue: priorHb,
            hbDate: targetFirstDate,
            monthPeriod: effectiveTargetMonth,
            recommendation: reco,
            overallStatus: reco.category === 'TRANSFUSI_2_RAWAT_INAP' ? 'Perlu Perhatian' : 'Berjalan',
            weeks,
            dailyRecords: {},
            labSchedule: {
              scheduledDate: '',
              testType: 'Rutin Hb (Evaluasi EPO)',
              status: 'Tidak Ada Jadwal',
              notes: `Tidak dijadwalkan Cek Hb bulan depan (Hb acuan ${priorHb.toFixed(1)} g/dL)`,
              isSelectiveHb: false,
            },
            updatedAt: new Date().toISOString(),
          };
        } else {
          // Pasien terjadwal (Tn B, Tn C, Tn D) berstatus 0 (menunggu hasil lab)
          const reco = calculateClinicalRecommendation(0, p.hdFrequency);
          const weeks = generateDefaultWeeks('MENUNGGU_HASIL_LAB', effectiveTargetMonth, p.hdFrequency, p.scheduleDay, p.singleDay, p.lastHdDate);
          return {
            ...p,
            hbValue: 0,
            prevHbValue: priorHb > 0 ? priorHb : p.prevHbValue,
            hbDate: targetFirstDate,
            monthPeriod: effectiveTargetMonth,
            recommendation: reco,
            overallStatus: 'Menunggu',
            weeks,
            dailyRecords: {},
            labSchedule: {
              scheduledDate: targetFirstDate,
              testType: 'Cek Hb Pilihan (Hb ≤ 8.9)',
              status: 'Terjadwal',
              notes: `⭐ Cek Hb Pilihan: Evaluasi anemia (Hb acuan ${priorHb > 0 ? priorHb.toFixed(1) : '< 9.0'} g/dL)`,
              isSelectiveHb: true,
              selectiveReason: `Nilai Hb acuan < 9.0 mg/dL (${priorHb > 0 ? priorHb.toFixed(1) : '< 9.0'})`,
            },
            updatedAt: new Date().toISOString(),
          };
        }
      }
    });

    try {
      localStorage.setItem('dialysis_patients_data_' + effectiveTargetMonth, JSON.stringify(targetMonthPatients));
    } catch (e) {
      console.error('Error saving target month data:', e);
    }

    if (selectedMonth === effectiveTargetMonth) {
      setPatients(targetMonthPatients);
    }

    const isScopePilihan = effectiveScope === 'PILIHAN';
    const scheduledCount = updatedSchedules.filter((s) => s.labSchedule.status === 'Terjadwal').length;
    setToastNotification({
      type: 'success',
      message: isScopePilihan
        ? `Jadwal Cek HB Pilihan tersimpan untuk ${effectiveTargetMonth}! Pasien stabil (≥ 9.0) otomatis terisi nilai bulan lalu, pasien anemia (< 9.0) berstatus 0 (Menunggu Hasil Lab).`
        : `Jadwal Cek HB Seluruh Pasien tersimpan untuk ${effectiveTargetMonth}! Seluruh (${scheduledCount}) pasien berstatus 0 (Menunggu Hasil Lab).`,
    });
  };

  const pendingHbCount = patients.filter((p) => p.hbValue <= 0).length;
  const scheduledLabCount = patients.filter((p) => p.labSchedule?.status === 'Terjadwal' || (!p.labSchedule && p.hbValue <= 0)).length;

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 flex flex-col font-sans selection:bg-rose-600 selection:text-white">
      
      {/* Toast Notification Banner */}
      {toastNotification && (
        <div className="print:hidden fixed top-20 right-4 z-50 max-w-md animate-in slide-in-from-top-4 duration-200">
          <div className={`p-4 rounded-xl shadow-xl border flex items-start gap-3 text-xs ${
            toastNotification.type === 'success'
              ? 'bg-emerald-50 text-emerald-900 border-emerald-300 dark:bg-emerald-950 dark:text-emerald-100 dark:border-emerald-700'
              : toastNotification.type === 'error'
              ? 'bg-rose-50 text-rose-900 border-rose-300 dark:bg-rose-950 dark:text-rose-100 dark:border-rose-700'
              : 'bg-indigo-50 text-indigo-900 border-indigo-300 dark:bg-indigo-950 dark:text-indigo-100 dark:border-indigo-700'
          }`}>
            {toastNotification.type === 'success' && <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />}
            {toastNotification.type === 'error' && <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />}
            {toastNotification.type === 'info' && <Info className="w-5 h-5 text-indigo-600 shrink-0 mt-0.5" />}
            <div className="flex-1 font-medium">{toastNotification.message}</div>
            <button
              onClick={() => setToastNotification(null)}
              className="text-slate-400 hover:text-slate-600 ml-2"
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* Main App Header */}
      <div className="print:hidden">
        <Header
          user={user}
          spreadsheetConfig={spreadsheetConfig}
          selectedMonth={selectedMonth}
          onMonthChange={handleMonthChange}
          onOpenSyncModal={() => setIsSyncModalOpen(true)}
          onOpenMonthlyHbModal={() => setIsMonthlyHbModalOpen(true)}
          onOpenLabScheduleModal={() => setIsLabScheduleModalOpen(true)}
          onOpenCalculatorModal={() => setIsCalculatorModalOpen(true)}
          onOpenPrintModal={() => setIsPrintModalOpen(true)}
          onOpenEpoSchedulePrintModal={() => setIsEpoSchedulePrintModalOpen(true)}
          onOpenExportExcelModal={() => {
            setExportExcelTarget('ALL');
            setIsExportExcelModalOpen(true);
          }}
          onOpenAddPatientModal={() => {
            setEditingPatient(null);
            setIsPatientModalOpen(true);
          }}
          onOpenBulkImportModal={() => setIsBulkImportModalOpen(true)}
          onOpenWorkflowGuide={() => setIsWorkflowGuideOpen(true)}
          onSignIn={handleSignIn}
          onSignOut={handleSignOut}
          isSyncing={isSyncing}
          isLoggingIn={isLoggingIn}
          pendingHbCount={pendingHbCount}
          scheduledLabCount={scheduledLabCount}
        />
      </div>

      {/* Main Content Area */}
      <main className="print:hidden flex-1 max-w-[1920px] 2xl:max-w-[2400px] w-full mx-auto px-2 sm:px-4 lg:px-6 py-4 space-y-3.5">
        
        {/* Google Sheets Sync Banner Indicator if Not Connected */}
        {!spreadsheetConfig?.spreadsheetId && (
          <div className="bg-gradient-to-r from-emerald-600 via-teal-600 to-indigo-700 text-white rounded-xl p-3.5 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-white/20 backdrop-blur-xs shrink-0">
                <FileSpreadsheet className="w-5 h-5" />
              </div>
              <div>
                <h3 className="font-bold text-xs sm:text-sm">Sinkronisasikan dengan Google Sheets (2 Arah)</h3>
                <p className="text-[11px] text-emerald-100">
                  Hubungkan dengan spreadsheet tim HD atau buat file spreadsheet baru otomatis di Google Drive Anda.
                </p>
              </div>
            </div>
            <button
              onClick={() => setIsSyncModalOpen(true)}
              className="h-8.5 inline-flex items-center justify-center gap-1.5 px-3.5 text-xs font-semibold rounded-lg bg-white text-emerald-800 hover:bg-emerald-50 active:bg-emerald-100 shadow-xs transition cursor-pointer shrink-0"
            >
              <span>Setup Google Sheets</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* Clinical Protocol & Key Statistics */}
        <DashboardStats patients={patients} />

        {/* Patient Table & Schedule Management */}
        <PatientTable
          patients={patients}
          selectedMonth={selectedMonth}
          onMonthChange={handleMonthChange}
          onEditPatient={(patient) => {
            setEditingPatient(patient);
            setIsPatientModalOpen(true);
          }}
          onDeletePatient={handleDeletePatient}
          onUpdateWeekStatus={handleUpdateWeekStatus}
          onUpdateDateAction={handleUpdateDateAction}
          onOpenBulkImportModal={() => setIsBulkImportModalOpen(true)}
          onOpenAddPatientModal={() => {
            setEditingPatient(null);
            setIsPatientModalOpen(true);
          }}
          onOpenMonthlyHbModal={() => setIsMonthlyHbModalOpen(true)}
          onOpenLabScheduleModal={() => setIsLabScheduleModalOpen(true)}
          onOpenExportExcelModal={(target) => {
            setExportExcelTarget(target || 'ALL');
            setIsExportExcelModalOpen(true);
          }}
          onOpenEpoSchedulePrintModal={() => setIsEpoSchedulePrintModalOpen(true)}
          onPushToSheet={handleQuickPushToSheet}
          isSyncing={isSyncing}
        />

      </main>

      {/* Footer */}
      <footer className="print:hidden bg-white dark:bg-slate-900 border-t border-slate-200 dark:border-slate-800 py-3 text-[11px] text-slate-500 dark:text-slate-400 text-center">
        <p className="flex items-center justify-center gap-1.5 flex-wrap">
          <strong className="text-rose-700 dark:text-rose-400 font-bold">EPOCARE</strong>
          <span>•</span>
          <span>Sistem Alokasi Terapi Eritropoietin & Transfusi Pasien Hemodialisa RS Happy Land Medical Centre</span>
        </p>
      </footer>

      {/* Modals */}
      <PatientModal
        isOpen={isPatientModalOpen}
        onClose={() => {
          setIsPatientModalOpen(false);
          setEditingPatient(null);
        }}
        onSave={handleSavePatient}
        initialPatient={editingPatient}
        currentMonth={selectedMonth}
      />

      <GoogleSheetsSyncModal
        isOpen={isSyncModalOpen}
        onClose={() => setIsSyncModalOpen(false)}
        user={user}
        spreadsheetConfig={spreadsheetConfig}
        onSaveSpreadsheetConfig={handleSaveSpreadsheetConfig}
        patients={patients}
        onPullFromSheet={handlePullFromSheet}
        onPushToSheet={handlePushToSheet}
        onPullViaAppsScript={handlePullViaAppsScript}
        onPushViaAppsScript={handlePushViaAppsScript}
        onImportCsv={handleImportCsv}
        onSignIn={handleSignIn}
        onOpenExportExcelModal={() => {
          setExportExcelTarget('ALL');
          setIsExportExcelModalOpen(true);
        }}
        selectedMonth={selectedMonth}
      />

      <ClinicalCalculatorModal
        isOpen={isCalculatorModalOpen}
        onClose={() => setIsCalculatorModalOpen(false)}
        onApplyToNewPatient={(hb) => {
          setEditingPatient(null);
          setIsPatientModalOpen(true);
        }}
      />

      <PrintReportModal
        isOpen={isPrintModalOpen}
        onClose={() => setIsPrintModalOpen(false)}
        patients={patients}
        selectedMonth={selectedMonth}
        onSwitchToEpoSchedule={() => {
          setIsPrintModalOpen(false);
          setIsEpoSchedulePrintModalOpen(true);
        }}
      />

      <EpoSchedulePrintModal
        isOpen={isEpoSchedulePrintModalOpen}
        onClose={() => setIsEpoSchedulePrintModalOpen(false)}
        patients={patients}
        selectedMonth={selectedMonth}
      />

      <BulkImportModal
        isOpen={isBulkImportModalOpen}
        onClose={() => setIsBulkImportModalOpen(false)}
        onImportPatients={handleImportPatients}
        selectedMonth={selectedMonth}
      />

      <MonthlyHbInputModal
        isOpen={isMonthlyHbModalOpen}
        onClose={() => setIsMonthlyHbModalOpen(false)}
        patients={patients}
        selectedMonth={selectedMonth}
        onMonthChange={handleMonthChange}
        onSaveBatchHb={handleSaveBatchHb}
        onOpenLabScheduleModal={() => {
          setIsMonthlyHbModalOpen(false);
          setIsLabScheduleModalOpen(true);
        }}
        onOpenWorkflowGuide={() => setIsWorkflowGuideOpen(true)}
      />

      <LabScheduleModal
        isOpen={isLabScheduleModalOpen}
        onClose={() => setIsLabScheduleModalOpen(false)}
        patients={patients}
        selectedMonth={selectedMonth}
        onMonthChange={handleMonthChange}
        onSaveLabSchedules={handleSaveLabSchedules}
        onSyncToSheets={(scope) => handleQuickPushToSheet(undefined, scope)}
        isSyncing={isSyncing}
      />

      <ExportExcelModal
        isOpen={isExportExcelModalOpen}
        onClose={() => setIsExportExcelModalOpen(false)}
        patients={patients}
        selectedMonth={selectedMonth}
        onMonthChange={handleMonthChange}
        initialSchedule={exportExcelTarget}
      />

      <WorkflowGuideModal
        isOpen={isWorkflowGuideOpen}
        onClose={() => setIsWorkflowGuideOpen(false)}
        onOpenMonthlyHbModal={() => setIsMonthlyHbModalOpen(true)}
        onOpenLabScheduleModal={() => setIsLabScheduleModalOpen(true)}
      />

    </div>
  );
}
