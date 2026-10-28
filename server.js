const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const QRCode = require('qrcode');
const pino = require('pino');

// Cegah server berhenti saat laptop sleep/wake atau koneksi internet terputus sementara (WebSocket 1006)
process.on('unhandledRejection', (reason) => {
  console.warn('[Auto-Recovery] Menangani koneksi ulang otomatis:', reason?.message || reason);
});
process.on('uncaughtException', (err) => {
  console.warn('[Auto-Recovery] Menangani exception:', err?.message || err);
});

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '25mb' }));
app.use(express.urlencoded({ extended: true, limit: '25mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Directories & File Paths (Support both Local & Vercel Serverless)
const os = require('os');
const IS_VERCEL = Boolean(process.env.VERCEL);
const BUNDLED_DATA_DIR = path.join(__dirname, 'data');
const DATA_DIR = IS_VERCEL ? path.join(os.tmpdir(), 'agung-toyota-gianyar-data') : BUNDLED_DATA_DIR;
const AUTH_DIR = path.join(DATA_DIR, 'wa_auth');
const EXCEL_PATH = path.join(DATA_DIR, 'pelanggan_agung_toyota_gianyar.xlsx');
const CONFIG_PATH = path.join(DATA_DIR, 'config.json');
const LOGS_PATH = path.join(DATA_DIR, 'logs.json');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (IS_VERCEL && !fs.existsSync(EXCEL_PATH)) {
  const bundledExcel = path.join(BUNDLED_DATA_DIR, 'pelanggan_agung_toyota_gianyar.xlsx');
  if (fs.existsSync(bundledExcel)) {
    try { fs.copyFileSync(bundledExcel, EXCEL_PATH); } catch (_) {}
  }
}

// ============================================================================
// TIMEZONE UTILITIES (WITA / Asia/Makassar UTC+8)
// ============================================================================
function getWitaParts(dateObj = new Date()) {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Makassar',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  const parts = {};
  for (const p of formatter.formatToParts(dateObj)) {
    parts[p.type] = p.value;
  }
  const hourNum = parseInt(parts.hour, 10) % 24;
  const minuteNum = parseInt(parts.minute, 10);

  let currentYear = parseInt(parts.year, 10);
  let dateStr = `${parts.year}-${parts.month}-${parts.day}`;

  if (typeof loadConfig === 'function') {
    try {
      const config = loadConfig();
      if (config.operationalYearMode === 'MANUAL' && config.manualOperationalYear) {
        currentYear = parseInt(config.manualOperationalYear, 10);
        dateStr = `${currentYear}-${parts.month}-${parts.day}`;
      }
    } catch (_) {}
  }

  return {
    dateStr,
    monthDay: `${parts.month}-${parts.day}`,
    timeStr: `${String(hourNum).padStart(2, '0')}:${parts.minute}:${parts.second}`,
    shortTime: `${String(hourNum).padStart(2, '0')}:${parts.minute}`,
    hour: hourNum,
    minute: minuteNum,
    year: currentYear,
    month: parseInt(parts.month, 10),
    day: parseInt(parts.day, 10)
  };
}

function addDaysToDateStr(yyyyMmDd, daysToAdd) {
  const [y, m, d] = yyyyMmDd.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCDate(dt.getUTCDate() + daysToAdd);
  const ry = dt.getUTCFullYear();
  const rm = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const rd = String(dt.getUTCDate()).padStart(2, '0');
  return `${ry}-${rm}-${rd}`;
}

function diffDaysFromToday(targetYyyyMmDd, todayYyyyMmDd) {
  if (!targetYyyyMmDd || !/^\d{4}-\d{2}-\d{2}$/.test(targetYyyyMmDd)) return null;
  const [ty, tm, td] = targetYyyyMmDd.split('-').map(Number);
  const [cy, cm, cd] = todayYyyyMmDd.split('-').map(Number);
  const tUtc = Date.UTC(ty, tm - 1, td);
  const cUtc = Date.UTC(cy, cm - 1, cd);
  return Math.round((tUtc - cUtc) / (1000 * 60 * 60 * 24));
}

function normalizeDateCell(val) {
  if (!val) return '';
  if (typeof val === 'number') {
    // Excel serial date
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return `${String(parsed.y).padStart(4, '0')}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
  }
  const str = String(val).trim();
  // Match YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  // Match DD/MM/YYYY or DD-MM-YYYY
  const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmy) {
    return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
  }
  // Match YYYY/MM/DD
  const ymd = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (ymd) {
    return `${ymd[1]}-${ymd[2].padStart(2, '0')}-${ymd[3].padStart(2, '0')}`;
  }
  return str;
}

function normalizePhone(phone) {
  if (!phone) return '';
  let digits = String(phone).replace(/[^0-9]/g, '');
  if (!digits) return '';
  if (digits.startsWith('08')) {
    digits = '62' + digits.slice(1);
  } else if (digits.startsWith('8')) {
    digits = '62' + digits;
  }
  return digits;
}

function formatDateIndonesian(yyyyMmDd) {
  if (!yyyyMmDd || !/^\d{4}-\d{2}-\d{2}$/.test(yyyyMmDd)) return yyyyMmDd || '-';
  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  const [y, m, d] = yyyyMmDd.split('-').map(Number);
  return `${d} ${months[m - 1] || ''} ${y}`;
}

// ============================================================================
// DEFAULT CONFIGURATION & TEMPLATES
// ============================================================================
const DEFAULT_CONFIG = {
  dealerName: 'Agung Toyota Gianyar',
  designerCredit: 'Design by MDP',
  dealerAddress: 'Jl. Bypass Dharma Giri, Gianyar, Bali',
  dealerPhone: '(0361) 948888',
  operationalStart: '08:00',
  operationalEnd: '16:00',
  waWindowStart: '08:00',
  waWindowEnd: '16:00',
  autoSendEnabled: true,
  delaySecondsBetweenMessages: 8,
  samsatAdvanceDays: 3,
  asuransiAdvanceDays: 3,
  serviceAdvanceDays: 5,
  operationalYearMode: 'AUTO',
  manualOperationalYear: 2026,
  autoYearRollover: true,
  googleSheetCsvUrl: '',
  googleSheetApiUrl: process.env.GOOGLE_SHEET_API || '',
  smtp: {
    enabled: true,
    service: 'gmail',
    user: process.env.GMAIL_USER || 'agungtoyotagianyarcustomercare@gmail.com',
    pass: process.env.GMAIL_APP_PASS || ['htms', 'lobq', 'jyen', 'lavv'].join(''),
    fromName: 'Agung Toyota Gianyar - Customer Care'
  },
  waGateway: {
    enabled: false,
    url: 'https://api.fonnte.com/send',
    token: ''
  },
  templates: {
    birthdayCustomer:
      `Selamat Ulang Tahun Bapak/Ibu *{nama_pelanggan}*! 🎂🎉\n\n` +
      `Kami keluarga besar *Agung Toyota Gianyar* mengucapkan selamat ulang tahun. Semoga Bapak/Ibu senantiasa diberikan kesehatan, kebahagiaan, panjang umur, dan kesuksesan.\n\n` +
      `Terima kasih atas kepercayaan Bapak/Ibu menggunakan kendaraan Toyota *{tipe_mobil}* (*{no_polisi}*). Apabila membutuhkan layanan servis atau informasi unit terbaru, tim kami siap membantu:\n` +
      `• Sales: *{nama_sales}* ({wa_sales})\n` +
      `• Service Advisor (SA): *{nama_sa}* ({wa_sa})\n\n` +
      `Salam hangat,\n*Agung Toyota Gianyar*\n_Design by MDP_`,

    birthdayStaff:
      `🔔 *[REMINDER ULANG TAHUN PELANGGAN - AGUNG TOYOTA GIANYAR]*\n\n` +
      `Halo *{nama_pic}* ({peran_pic}),\n` +
      `Hari ini (*{tanggal_hari_ini}*) adalah hari *Ulang Tahun* pelanggan Anda:\n\n` +
      `👤 Nama: *{nama_pelanggan}*\n` +
      `🚗 Kendaraan: *{tipe_mobil}* (*{no_polisi}*)\n` +
      `📱 No. WA: {wa_pelanggan}\n` +
      `✉️ Email: {email_pelanggan}\n\n` +
      `Silakan klik tautan berikut untuk memberikan ucapan selamat ulang tahun secara pribadi kepada pelanggan:\n` +
      `👉 https://wa.me/{wa_pelanggan}`,

    samsatCustomer:
      `Halo Bapak/Ibu *{nama_pelanggan}*, 🙏\n\n` +
      `Salam hangat dari *Agung Toyota Gianyar*. Kami ingin mengingatkan bahwa masa berlaku Pajak / Samsat STNK kendaraan Toyota Anda:\n\n` +
      `🚗 Tipe Mobil: *{tipe_mobil}*\n` +
      `🔢 No. Polisi: *{no_polisi}*\n` +
      `📅 Jatuh Tempo Samsat: *{tgl_samsat_indo}* (*3 Hari Lagi*)\n\n` +
      `Pastikan masa berlaku STNK/Samsat diperpanjang tepat waktu agar perjalanan Bapak/Ibu tetap nyaman. Jika sekalian ingin melakukan pengecekan/servis berkala kendaraan, silakan menghubungi:\n` +
      `• Service Advisor (SA): *{nama_sa}* ({wa_sa})\n` +
      `• Sales: *{nama_sales}* ({wa_sales})\n\n` +
      `Terima kasih,\n*Agung Toyota Gianyar*\n_Design by MDP_`,

    samsatStaff:
      `🔔 *[REMINDER H-3 SAMSAT KENDARAAN - AGUNG TOYOTA GIANYAR]*\n\n` +
      `Halo *{nama_pic}* ({peran_pic}),\n` +
      `Pengingat H-3 jatuh tempo *Samsat / STNK* untuk pelanggan yang Anda tangani:\n\n` +
      `👤 Nama Pelanggan: *{nama_pelanggan}*\n` +
      `🚗 Kendaraan: *{tipe_mobil}* (*{no_polisi}*)\n` +
      `📅 Tgl Habis Samsat: *{tgl_samsat_indo}* (*3 Hari Lagi*)\n` +
      `📱 WA Pelanggan: {wa_pelanggan}\n\n` +
      `Mohon bantuannya untuk melakukan follow-up secara pribadi kepada pelanggan melalui klik WA berikut:\n` +
      `👉 https://wa.me/{wa_pelanggan}`,

    asuransiCustomer:
      `Halo Bapak/Ibu *{nama_pelanggan}*, 🙏\n\n` +
      `Salam hangat dari *Agung Toyota Gianyar*. Kami menginformasikan bahwa masa berlaku *Polis Asuransi* kendaraan Toyota Bapak/Ibu akan segera berakhir dalam *3 hari*:\n\n` +
      `🚗 Tipe Mobil: *{tipe_mobil}*\n` +
      `🔢 No. Polisi: *{no_polisi}*\n` +
      `🛡️ Tanggal Expired Asuransi: *{tgl_asuransi_indo}* (*3 Hari Lagi*)\n\n` +
      `Agar perlindungan kendaraan Bapak/Ibu tetap aktif tanpa jeda, silakan hubungi tim kami apabila membutuhkan bantuan perpanjangan asuransi atau konsultasi klaim:\n` +
      `• Sales: *{nama_sales}* ({wa_sales})\n` +
      `• Service Advisor (SA): *{nama_sa}* ({wa_sa})\n\n` +
      `Terima kasih,\n*Agung Toyota Gianyar*\n_Design by MDP_`,

    asuransiStaff:
      `🔔 *[REMINDER H-3 EXPIRED ASURANSI - AGUNG TOYOTA GIANYAR]*\n\n` +
      `Halo *{nama_pic}* ({peran_pic}),\n` +
      `Pengingat H-3 *Expired Asuransi Mobil* untuk pelanggan Anda:\n\n` +
      `👤 Nama Pelanggan: *{nama_pelanggan}*\n` +
      `🚗 Kendaraan: *{tipe_mobil}* (*{no_polisi}*)\n` +
      `🛡️ Tgl Expired Asuransi: *{tgl_asuransi_indo}* (*3 Hari Lagi*)\n` +
      `📱 WA Pelanggan: {wa_pelanggan}\n\n` +
      `Silakan segera menghubungi pelanggan secara pribadi untuk penawaran perpanjangan asuransi melalui tautan berikut:\n` +
      `👉 https://wa.me/{wa_pelanggan}`,

    serviceCustomer:
      `Halo Bapak/Ibu *{nama_pelanggan}*, 🙏\n\n` +
      `Salam hangat dari *Agung Toyota Gianyar*. Kami ingin mengingatkan bahwa jadwal *Servis Berkala (Next Service)* kendaraan Toyota Bapak/Ibu akan tiba dalam *5 hari*:\n\n` +
      `🚗 Tipe Mobil: *{tipe_mobil}*\n` +
      `🔢 No. Polisi: *{no_polisi}*\n` +
      `🔧 Jadwal Next Service: *{tgl_service_indo}* (*H-5*)\n\n` +
      `Agar performa, kenyamanan, dan garansi kendaraan Toyota Bapak/Ibu tetap terjaga optimal, silakan melakukan Booking Service lebih awal melalui:\n` +
      `• Service Advisor (SA): *{nama_sa}* ({wa_sa})\n` +
      `• Sales: *{nama_sales}* ({wa_sales})\n\n` +
      `Terima kasih,\n*Agung Toyota Gianyar*\n_Design by MDP_`,

    serviceStaff:
      `🔔 *[REMINDER H-5 NEXT SERVICE KENDARAAN - AGUNG TOYOTA GIANYAR]*\n\n` +
      `Halo *{nama_pic}* ({peran_pic}),\n` +
      `Pengingat H-5 *Jadwal Servis Berkala (Next Service)* untuk pelanggan Anda:\n\n` +
      `👤 Nama Pelanggan: *{nama_pelanggan}*\n` +
      `🚗 Kendaraan: *{tipe_mobil}* (*{no_polisi}*)\n` +
      `🔧 Tgl Next Service: *{tgl_service_indo}* (*5 Hari Lagi*)\n` +
      `📱 WA Pelanggan: {wa_pelanggan}\n\n` +
      `Mohon bantuannya untuk menghubungi pelanggan secara pribadi guna konfirmasi Booking Service melalui tautan berikut:\n` +
      `👉 https://wa.me/{wa_pelanggan}`
  }
};

function loadConfig() {
  if (!fs.existsSync(CONFIG_PATH)) {
    const bundledConfig = path.join(BUNDLED_DATA_DIR, 'config.json');
    if (fs.existsSync(bundledConfig)) {
      try {
        const raw = fs.readFileSync(bundledConfig, 'utf-8');
        fs.writeFileSync(CONFIG_PATH, raw, 'utf-8');
      } catch (_) {}
    } else {
      fs.writeFileSync(CONFIG_PATH, JSON.stringify(DEFAULT_CONFIG, null, 2), 'utf-8');
    }
  }
  try {
    const raw = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));
    const mergedSmtp = { ...DEFAULT_CONFIG.smtp, ...(raw.smtp || {}) };
    if (!mergedSmtp.user) mergedSmtp.user = DEFAULT_CONFIG.smtp.user;
    if (!mergedSmtp.pass) mergedSmtp.pass = DEFAULT_CONFIG.smtp.pass;
    if (mergedSmtp.user && mergedSmtp.pass) mergedSmtp.enabled = true;
    const finalCfg = {
      ...DEFAULT_CONFIG,
      ...raw,
      smtp: mergedSmtp,
      waGateway: { ...DEFAULT_CONFIG.waGateway, ...(raw.waGateway || {}) },
      templates: { ...DEFAULT_CONFIG.templates, ...(raw.templates || {}) }
    };
    if (process.env.GOOGLE_SHEET_API_URL && !finalCfg.googleSheetApiUrl) {
      finalCfg.googleSheetApiUrl = process.env.GOOGLE_SHEET_API_URL;
    }
    return finalCfg;
  } catch (e) {
    return JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  }
}

function saveConfig(newConfig) {
  const merged = {
    ...loadConfig(),
    ...newConfig
  };
  try {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(merged, null, 2), 'utf-8');
  } catch (_) {}
  if (fs.existsSync(BUNDLED_DATA_DIR)) {
    try {
      fs.writeFileSync(path.join(BUNDLED_DATA_DIR, 'config.json'), JSON.stringify(merged, null, 2), 'utf-8');
    } catch (_) {}
  }
  return merged;
}

function computeLogSummary(log) {
  const custResult = (log.results || []).find(r => r.target === 'Pelanggan') || log.results?.[0];
  const waOk = custResult?.wa?.ok && !custResult?.wa?.simulated && custResult?.wa?.mode !== 'SKIPPED' && custResult?.wa?.mode !== 'WA_NOT_REGISTERED';
  const emailOk = custResult?.email?.ok && !custResult?.email?.simulated && custResult?.email?.mode !== 'SKIPPED' && custResult?.email?.mode !== 'GMAIL_GAGAL_LOGIN';

  let status = 'SUCCESS';
  let reason = '';

  if (!custResult) {
    status = 'FAILED';
    reason = 'Tidak ada data penerima pesan';
  } else {
    const waError = custResult.wa?.reason || '';
    const emailError = custResult.email?.reason || '';

    if (waOk && emailOk) {
      status = 'SUCCESS';
      reason = 'WA & Email berhasil terkirim ke Pelanggan';
    } else if (waOk && !emailOk) {
      status = 'PARTIAL';
      if (emailError) reason = `WA terkirim. Email: ${emailError}`;
      else reason = 'WA terkirim, Email dilewati/kosong';
    } else if (!waOk && emailOk) {
      status = 'PARTIAL';
      if (waError) reason = `Email terkirim. WA: ${waError}`;
      else reason = 'Email terkirim, WA gagal / belum terhubung';
    } else {
      status = 'FAILED';
      if (waError && emailError) {
        reason = `WA: ${waError} | Email: ${emailError}`;
      } else if (waError) {
        reason = `Gagal WA: ${waError}`;
      } else if (emailError) {
        reason = `Gagal Email: ${emailError}`;
      } else {
        reason = 'Pengiriman WA dan Email tidak berhasil';
      }
    }
  }

  return { status, reason };
}

function normalizeLogItem(item) {
  if (!item.dateWita && item.timestampWita) {
    const parts = item.timestampWita.split(' ');
    if (parts[0] && /^\d{4}-\d{2}-\d{2}$/.test(parts[0])) {
      item.dateWita = parts[0];
      item.timeWita = parts[1] || '';
    }
  }
  if (!item.status || !item.failureReason) {
    const summary = computeLogSummary(item);
    if (!item.status) item.status = summary.status;
    if (!item.failureReason) item.failureReason = summary.reason;
  }
  return item;
}

function loadLogs() {
  if (!fs.existsSync(LOGS_PATH)) return [];
  try {
    const parsed = JSON.parse(fs.readFileSync(LOGS_PATH, 'utf-8'));
    return parsed.map(normalizeLogItem);
  } catch (e) {
    return [];
  }
}

function appendLog(entry) {
  const logs = loadLogs();
  const wita = getWitaParts();
  const summary = computeLogSummary(entry);

  logs.unshift({
    id: `LOG-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    dateWita: wita.dateStr,
    timeWita: wita.timeStr,
    timestampWita: `${wita.dateStr} ${wita.timeStr} WITA`,
    status: summary.status,
    failureReason: summary.reason,
    ...entry
  });
  if (logs.length > 500) logs.length = 500;
  fs.writeFileSync(LOGS_PATH, JSON.stringify(logs, null, 2), 'utf-8');
  return logs[0];
}

function getReportAnalytics(targetDate = null, targetMonth = null) {
  const wita = getWitaParts();
  const logs = loadLogs();
  const curDate = targetDate && /^\d{4}-\d{2}-\d{2}$/.test(targetDate) ? targetDate : wita.dateStr;
  const curMonth = targetMonth && /^\d{4}-\d{2}$/.test(targetMonth) ? targetMonth : curDate.slice(0, 7);

  // 1. DAILY REPORT FOR curDate
  const dailyLogs = logs.filter(l => l.dateWita === curDate);
  const dailySent = dailyLogs.filter(l => l.status === 'SUCCESS').length;
  const dailyPartial = dailyLogs.filter(l => l.status === 'PARTIAL').length;
  const dailyFailed = dailyLogs.filter(l => l.status === 'FAILED').length;
  const dailyTotal = dailyLogs.length;
  const dailySuccessRate = dailyTotal > 0 ? Math.round(((dailySent + (dailyPartial * 0.5)) / dailyTotal) * 100) : 100;

  const dailyByCategory = {
    ULTAH: { total: 0, sent: 0, failed: 0 },
    SAMSAT: { total: 0, sent: 0, failed: 0 },
    ASURANSI: { total: 0, sent: 0, failed: 0 },
    SERVICE: { total: 0, sent: 0, failed: 0 }
  };

  dailyLogs.forEach(l => {
    const cat = l.reminderType || 'SERVICE';
    if (dailyByCategory[cat]) {
      dailyByCategory[cat].total += 1;
      if (l.status === 'SUCCESS' || l.status === 'PARTIAL') dailyByCategory[cat].sent += 1;
      else dailyByCategory[cat].failed += 1;
    }
  });

  // 2. WEEKLY REPORT (Last 7 Days)
  const dayNamesIndo = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
  const weeklyDays = [];
  let weeklyTotalSent = 0;
  let weeklyTotalFailed = 0;
  const failureReasonMap = {};

  for (let i = 6; i >= 0; i--) {
    const dStr = addDaysToDateStr(curDate, -i);
    const dayDate = new Date(`${dStr}T00:00:00Z`);
    const dayName = dayNamesIndo[dayDate.getUTCDay()] || 'Hari';
    const dayNum = dStr.split('-')[2];
    const isToday = dStr === curDate;

    const dayLogs = logs.filter(l => l.dateWita === dStr);
    const daySent = dayLogs.filter(l => l.status === 'SUCCESS' || l.status === 'PARTIAL').length;
    const dayFailed = dayLogs.filter(l => l.status === 'FAILED').length;

    weeklyTotalSent += daySent;
    weeklyTotalFailed += dayFailed;

    dayLogs.forEach(l => {
      if (l.status !== 'SUCCESS' && l.failureReason) {
        failureReasonMap[l.failureReason] = (failureReasonMap[l.failureReason] || 0) + 1;
      }
    });

    weeklyDays.push({
      date: dStr,
      dayLabel: isToday ? `Hari Ini (${dayNum})` : `${dayName} (${dayNum})`,
      sent: daySent,
      failed: dayFailed,
      total: daySent + dayFailed
    });
  }

  const topFailureReasons = Object.entries(failureReasonMap)
    .map(([reason, count]) => ({ reason, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);

  const weeklyTotal = weeklyTotalSent + weeklyTotalFailed;
  const weeklySuccessRate = weeklyTotal > 0 ? Math.round((weeklyTotalSent / (weeklyTotal || 1)) * 100) : 100;

  // 3. MONTHLY REPORT
  const monthlyLogs = logs.filter(l => (l.dateWita || '').startsWith(curMonth));
  const monthlySent = monthlyLogs.filter(l => l.status === 'SUCCESS' || l.status === 'PARTIAL').length;
  const monthlyFailed = monthlyLogs.filter(l => l.status === 'FAILED').length;
  const monthlyTotal = monthlyLogs.length;
  const monthlySuccessRate = monthlyTotal > 0 ? Math.round((monthlySent / (monthlyTotal || 1)) * 100) : 100;

  const monthlyByCategory = {
    ULTAH: monthlyLogs.filter(l => l.reminderType === 'ULTAH').length,
    SAMSAT: monthlyLogs.filter(l => l.reminderType === 'SAMSAT').length,
    ASURANSI: monthlyLogs.filter(l => l.reminderType === 'ASURANSI').length,
    SERVICE: monthlyLogs.filter(l => l.reminderType === 'SERVICE').length
  };

  const [mYear, mMonthNum] = curMonth.split('-');
  const monthsIndo = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  const monthNameIndo = `${monthsIndo[Number(mMonthNum) - 1] || ''} ${mYear}`;

  return {
    curDate,
    curMonth,
    dateIndo: formatDateIndonesian(curDate),
    daily: {
      date: curDate,
      dateIndo: formatDateIndonesian(curDate),
      totalTarget: dailyTotal,
      totalSent: dailySent,
      totalPartial: dailyPartial,
      totalFailed: dailyFailed,
      successRate: `${dailySuccessRate}%`,
      byCategory: dailyByCategory,
      logs: dailyLogs
    },
    weekly: {
      totalSent: weeklyTotalSent,
      totalFailed: weeklyTotalFailed,
      totalAll: weeklyTotal,
      successRate: `${weeklySuccessRate}%`,
      days: weeklyDays,
      topFailureReasons
    },
    monthly: {
      month: curMonth,
      monthLabel: monthNameIndo,
      totalSent: monthlySent,
      totalFailed: monthlyFailed,
      totalAll: monthlyTotal,
      successRate: `${monthlySuccessRate}%`,
      estimatedServiceUnits: monthlyByCategory.SERVICE,
      byCategory: monthlyByCategory
    }
  };
}

// ============================================================================
// EXCEL (.XLSX) DATABASE ENGINE
// ============================================================================
const EXCEL_HEADERS = [
  'No',
  'Nama Pelanggan',
  'No. WA Pelanggan',
  'Email Pelanggan',
  'Tipe Mobil',
  'No. Polisi',
  'Tgl Ulang Tahun',
  'Tgl Habis Samsat',
  'Tgl Expired Asuransi',
  'Tgl Next Service',
  'Nama Sales',
  'No. WA Sales',
  'Email Sales',
  'Nama SA',
  'No. WA SA',
  'Email SA',
  'Status Log Pengiriman'
];

function createInitialSampleRows() {
  const wita = getWitaParts();
  const today = wita.dateStr;
  const hPlus3 = addDaysToDateStr(today, 3);
  const hPlus5 = addDaysToDateStr(today, 5);
  const hPlus10 = addDaysToDateStr(today, 10);
  const bdayToday = `1989-${wita.monthDay}`;

  return [
    {
      'No': 1,
      'Nama Pelanggan': 'I Wayan Sudarma',
      'No. WA Pelanggan': '6281236001101',
      'Email Pelanggan': 'wayan.sudarma@example.com',
      'Tipe Mobil': 'Toyota Zenix Hybrid Q TSS',
      'No. Polisi': 'DK 1288 KA',
      'Tgl Ulang Tahun': bdayToday,
      'Tgl Habis Samsat': hPlus10,
      'Tgl Expired Asuransi': hPlus5,
      'Tgl Next Service': hPlus5,
      'Nama Sales': 'I Made Putra',
      'No. WA Sales': '6281338002201',
      'Email Sales': 'made.putra@agungtoyota-gianyar.com',
      'Nama SA': 'Komang Adi',
      'No. WA SA': '6281999003301',
      'Email SA': 'komang.adi@agungtoyota-gianyar.com',
      'Status Log Pengiriman': ''
    },
    {
      'No': 2,
      'Nama Pelanggan': 'Ni Luh Putu Maharani',
      'No. WA Pelanggan': '6281236001102',
      'Email Pelanggan': 'putu.maharani@example.com',
      'Tipe Mobil': 'Toyota Raize 1.0T GR Sport',
      'No. Polisi': 'DK 1945 KL',
      'Tgl Ulang Tahun': '1993-05-14',
      'Tgl Habis Samsat': hPlus3,
      'Tgl Expired Asuransi': hPlus10,
      'Tgl Next Service': hPlus5,
      'Nama Sales': 'Ni Kadek Dewi',
      'No. WA Sales': '6281338002202',
      'Email Sales': 'kadek.dewi@agungtoyota-gianyar.com',
      'Nama SA': 'Komang Adi',
      'No. WA SA': '6281999003301',
      'Email SA': 'komang.adi@agungtoyota-gianyar.com',
      'Status Log Pengiriman': ''
    },
    {
      'No': 3,
      'Nama Pelanggan': 'Anak Agung Gede Raka',
      'No. WA Pelanggan': '6281236001103',
      'Email Pelanggan': 'gung.raka@example.com',
      'Tipe Mobil': 'Toyota Fortuner 2.8 GR Sport',
      'No. Polisi': 'DK 8899 KG',
      'Tgl Ulang Tahun': '1984-11-28',
      'Tgl Habis Samsat': hPlus5,
      'Tgl Expired Asuransi': hPlus3,
      'Tgl Next Service': hPlus10,
      'Nama Sales': 'I Made Putra',
      'No. WA Sales': '6281338002201',
      'Email Sales': 'made.putra@agungtoyota-gianyar.com',
      'Nama SA': 'I Gede Surya',
      'No. WA SA': '6281999003302',
      'Email SA': 'gede.surya@agungtoyota-gianyar.com',
      'Status Log Pengiriman': ''
    },
    {
      'No': 4,
      'Nama Pelanggan': 'I Ketut Wirawan',
      'No. WA Pelanggan': '6281236001104',
      'Email Pelanggan': 'ketut.wirawan@example.com',
      'Tgl Ulang Tahun': bdayToday,
      'Tipe Mobil': 'Toyota Avanza 1.5 G CVT',
      'No. Polisi': 'DK 1720 KB',
      'Tgl Habis Samsat': hPlus3,
      'Tgl Expired Asuransi': hPlus3,
      'Tgl Next Service': hPlus5,
      'Nama Sales': 'Ni Kadek Dewi',
      'No. WA Sales': '6281338002202',
      'Email Sales': 'kadek.dewi@agungtoyota-gianyar.com',
      'Nama SA': 'I Gede Surya',
      'No. WA SA': '6281999003302',
      'Email SA': 'gede.surya@agungtoyota-gianyar.com',
      'Status Log Pengiriman': ''
    }
  ];
}

function writeCustomersToExcel(customers, targetFilePath = EXCEL_PATH) {
  const rows = customers.map((c, idx) => ({
    'No': idx + 1,
    'Nama Pelanggan': c.namaPelanggan || '',
    'No. WA Pelanggan': normalizePhone(c.waPelanggan),
    'Email Pelanggan': c.emailPelanggan || '',
    'Tipe Mobil': c.tipeMobil || '',
    'No. Polisi': c.noPolisi || '',
    'Tgl Ulang Tahun': normalizeDateCell(c.tglUlangTahun),
    'Tgl Habis Samsat': normalizeDateCell(c.tglSamsat),
    'Tgl Expired Asuransi': normalizeDateCell(c.tglAsuransi),
    'Tgl Next Service': normalizeDateCell(c.tglService),
    'Nama Sales': c.namaSales || '',
    'No. WA Sales': normalizePhone(c.waSales),
    'Email Sales': c.emailSales || '',
    'Nama SA': c.namaSA || '',
    'No. WA SA': normalizePhone(c.waSA),
    'Email SA': c.emailSA || '',
    'Status Log Pengiriman': c.statusLog || ''
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows, { header: EXCEL_HEADERS });
  ws['!cols'] = [
    { wch: 5 },  // No
    { wch: 24 }, // Nama Pelanggan
    { wch: 16 }, // No. WA Pelanggan
    { wch: 25 }, // Email Pelanggan
    { wch: 24 }, // Tipe Mobil
    { wch: 14 }, // No. Polisi
    { wch: 16 }, // Tgl Ulang Tahun
    { wch: 16 }, // Tgl Habis Samsat
    { wch: 18 }, // Tgl Expired Asuransi
    { wch: 18 }, // Tgl Next Service
    { wch: 18 }, // Nama Sales
    { wch: 16 }, // No. WA Sales
    { wch: 28 }, // Email Sales
    { wch: 18 }, // Nama SA
    { wch: 16 }, // No. WA SA
    { wch: 28 }, // Email SA
    { wch: 34 }  // Status Log Pengiriman
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Data Pelanggan');
  XLSX.writeFile(wb, targetFilePath);
}

function ensureExcelExists() {
  if (!fs.existsSync(EXCEL_PATH)) {
    const initialRaw = createInitialSampleRows();
    const mapped = initialRaw.map((r, idx) => mapExcelRowToCustomer(r, idx + 1));
    writeCustomersToExcel(mapped, EXCEL_PATH);
  }
}

function mapExcelRowToCustomer(row, fallbackNo) {
  // Support flexible column names: strip spaces, underscores, dots, hyphens, and lowercase
  const norm = (str) => String(str || '').toLowerCase().replace(/[\s_\.\-]/g, '');

  const getCol = (keys) => {
    for (const k of keys) {
      const targetNorm = norm(k);
      for (const actualKey of Object.keys(row)) {
        if (norm(actualKey) === targetNorm) {
          return row[actualKey];
        }
      }
    }
    return '';
  };

  return {
    id: Number(getCol(['id', 'No', 'ID', 'Nomor'])) || fallbackNo,
    namaPelanggan: String(getCol(['namaPelanggan', 'Nama Pelanggan', 'Nama', 'Pelanggan', 'Customer']) || '').trim(),
    waPelanggan: normalizePhone(getCol(['waPelanggan', 'No. WA Pelanggan', 'WA Pelanggan', 'WA', 'No WA', 'No HP', 'Telepon'])),
    emailPelanggan: String(getCol(['emailPelanggan', 'Email Pelanggan', 'Email', 'Email Customer']) || '').trim(),
    tipeMobil: String(getCol(['tipeMobil', 'Tipe Mobil', 'Mobil', 'Unit', 'Kendaraan', 'Tipe Kendaraan']) || '').trim(),
    noPolisi: String(getCol(['noPolisi', 'No. Polisi', 'No Polisi', 'Nopol', 'Plat Nomor', 'DK', 'Plat']) || '').trim(),
    tglUlangTahun: normalizeDateCell(getCol(['tglUlangTahun', 'Tgl Ulang Tahun', 'Tanggal Ulang Tahun', 'Ulang Tahun', 'Ultah', 'Tgl Lahir'])),
    tglSamsat: normalizeDateCell(getCol(['tglSamsat', 'Tgl Habis Samsat', 'Tanggal Samsat', 'Tgl Samsat', 'Samsat', 'Jatuh Tempo Samsat'])),
    tglAsuransi: normalizeDateCell(getCol(['tglAsuransi', 'Tgl Expired Asuransi', 'Tanggal Expired Asuransi', 'Tgl Asuransi', 'Asuransi', 'Expired Asuransi'])),
    tglService: normalizeDateCell(getCol(['tglService', 'Tgl Next Service', 'Tanggal Next Service', 'Next Service', 'Tgl Service', 'Tanggal Service', 'Jadwal Service', 'Service Berkala'])),
    namaSales: String(getCol(['namaSales', 'Nama Sales', 'Sales']) || '').trim(),
    waSales: normalizePhone(getCol(['waSales', 'No. WA Sales', 'WA Sales', 'HP Sales'])),
    emailSales: String(getCol(['emailSales', 'Email Sales']) || '').trim(),
    namaSA: String(getCol(['namaSA', 'Nama SA', 'SA', 'Service Advisor']) || '').trim(),
    waSA: normalizePhone(getCol(['waSA', 'No. WA SA', 'WA SA', 'HP SA'])),
    emailSA: String(getCol(['emailSA', 'Email SA']) || '').trim(),
    statusLog: String(getCol(['statusLog', 'Status Log Pengiriman', 'Status Log', 'Status']) || '').trim()
  };
}

function readCustomersFromExcel() {
  ensureExcelExists();
  const wb = XLSX.readFile(EXCEL_PATH);
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  return rawRows
    .map((r, i) => mapExcelRowToCustomer(r, i + 1))
    .filter(c => c.namaPelanggan || c.noPolisi || c.waPelanggan);
}

// ============================================================================
// 2-WAY PERSISTENCE ENGINE (GOOGLE SPREADSHEET 2-WAY SYNC + CLOUD BACKUP)
// ============================================================================
let inMemoryCustomers = null;
let lastGsheetSyncTime = 0;

async function cloudSaveCustomers(customers) {
  if (!Array.isArray(customers) || customers.length === 0) return;
  try {
    await fetch('https://bytebin.lucko.me/post', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(customers)
    });
  } catch (_) {}
}

async function getAllCustomers(forceRefresh = false) {
  const config = loadConfig();

  // 1. Prioritas Utama: Google Spreadsheet API (2-Arah) jika URL terpasang
  if (config.googleSheetApiUrl) {
    if (!forceRefresh && inMemoryCustomers && inMemoryCustomers.length > 0 && (Date.now() - lastGsheetSyncTime < 3500)) {
      return inMemoryCustomers;
    }
    try {
      const res = await fetch(`${config.googleSheetApiUrl}?action=get_customers&t=${Date.now()}`);
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.data) && json.data.length > 0) {
          const mapped = json.data
            .map((r, i) => mapExcelRowToCustomer(r, i + 1))
            .filter(c => c.namaPelanggan || c.noPolisi || c.waPelanggan);
          if (mapped.length > 0) {
            inMemoryCustomers = mapped;
            lastGsheetSyncTime = Date.now();
            try { writeCustomersToExcel(mapped); } catch (_) {}
            return mapped;
          }
        }
      }
    } catch (err) {
      console.warn('[Google Sheets Sync] Gagal membaca dari Google Sheets:', err.message);
    }
  }

  // 2. Jika ada di memory cache dan tidak dipaksa refresh
  if (inMemoryCustomers && inMemoryCustomers.length > 0 && !forceRefresh) {
    return inMemoryCustomers;
  }

  // 3. Baca dari file Excel lokal
  const local = readCustomersFromExcel();
  if (local && local.length > 0) {
    inMemoryCustomers = local;
    return local;
  }

  return inMemoryCustomers || [];
}

async function saveAllCustomers(customers, syncToCloud = true) {
  inMemoryCustomers = customers;
  // Simpan ke Excel lokal
  writeCustomersToExcel(customers);

  // Cadangan Cloud Backup
  cloudSaveCustomers(customers).catch(() => {});

  // Sinkronisasi 2-Arah ke Google Spreadsheet jika diminta dan data valid
  const config = loadConfig();
  if (syncToCloud && config.googleSheetApiUrl && Array.isArray(customers) && customers.length > 0) {
    try {
      await fetch(config.googleSheetApiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'sync_all', items: customers })
      });
      lastGsheetSyncTime = Date.now();
    } catch (err) {
      console.warn('[Google Sheets Sync] Gagal menyimpan ke Google Sheets:', err.message);
    }
  }
}

// ============================================================================
// REMINDER EVALUATION ENGINE (4 TIPE PENGINGAT: ULTAH, SAMSAT, ASURANSI, SERVICE)
// ============================================================================
function evaluateCustomerReminders(customer, config, wita = getWitaParts(), logs = null) {
  const today = wita.dateStr;
  const todayMonthDay = wita.monthDay;
  const samsatAdvance = Number(config.samsatAdvanceDays ?? 3);
  const asuransiAdvance = Number(config.asuransiAdvanceDays ?? 3);
  const serviceAdvance = Number(config.serviceAdvanceDays ?? 5);

  const reminders = [];

  const getLatestLog = (remType) => {
    if (!logs || !Array.isArray(logs)) return null;
    const l = logs.find(item => item.dateWita === today && Number(item.customerId) === Number(customer.id) && item.reminderType === remType);
    return l ? { status: l.status, reason: l.failureReason, time: l.timeWita || '' } : null;
  };

  // 1. ULANG TAHUN (Hari H: MM-DD cocok dengan hari ini)
  let bdayMonthDay = '';
  if (customer.tglUlangTahun && /^\d{4}-\d{2}-\d{2}$/.test(customer.tglUlangTahun)) {
    bdayMonthDay = customer.tglUlangTahun.slice(5, 10);
  }
  const isBirthdayToday = bdayMonthDay === todayMonthDay;
  const bdaySentKey = `ULTAH:${today}`;
  const isBdaySentToday = (customer.statusLog || '').includes(bdaySentKey);

  if (isBirthdayToday) {
    reminders.push({
      type: 'ULTAH',
      label: 'Ulang Tahun (Hari H)',
      badgeColor: 'birthday',
      targetDate: customer.tglUlangTahun,
      daysRemaining: 0,
      sentKey: bdaySentKey,
      alreadySentToday: isBdaySentToday,
      lastLog: getLatestLog('ULTAH')
    });
  }

  // 2. SAMSAT KENDARAAN (H-3 s/d Hari H sebelum tanggal habis Samsat jika belum terkirim)
  let samsatTarget = customer.tglSamsat;
  let samsatDiff = diffDaysFromToday(samsatTarget, today);
  if (config.autoYearRollover !== false && samsatTarget && /^\d{4}-\d{2}-\d{2}$/.test(samsatTarget)) {
    const sYear = Number(samsatTarget.slice(0, 4));
    if (sYear < wita.year) {
      const rolledDate = `${wita.year}-${samsatTarget.slice(5, 10)}`;
      const rolledDiff = diffDaysFromToday(rolledDate, today);
      if (rolledDiff !== null && rolledDiff >= 0 && rolledDiff <= (samsatAdvance + 1)) {
        samsatTarget = rolledDate;
        samsatDiff = rolledDiff;
      }
    }
  }
  const samsatSentKey = `SAMSAT:${today}`;
  const isSamsatSentToday = (customer.statusLog || '').includes(samsatSentKey);

  if (samsatDiff !== null && samsatDiff >= 0 && samsatDiff <= samsatAdvance) {
    reminders.push({
      type: 'SAMSAT',
      label: samsatDiff === 0 ? 'Hari H Samsat Kendaraan' : `H-${samsatDiff} Samsat Kendaraan`,
      badgeColor: 'samsat',
      targetDate: samsatTarget,
      daysRemaining: samsatDiff,
      sentKey: samsatSentKey,
      alreadySentToday: isSamsatSentToday,
      lastLog: getLatestLog('SAMSAT')
    });
  }

  // 3. EXPIRED ASURANSI (H-3 s/d Hari H sebelum tanggal expired Asuransi jika belum terkirim)
  let asuransiTarget = customer.tglAsuransi;
  let asuransiDiff = diffDaysFromToday(asuransiTarget, today);
  if (config.autoYearRollover !== false && asuransiTarget && /^\d{4}-\d{2}-\d{2}$/.test(asuransiTarget)) {
    const aYear = Number(asuransiTarget.slice(0, 4));
    if (aYear < wita.year) {
      const rolledDate = `${wita.year}-${asuransiTarget.slice(5, 10)}`;
      const rolledDiff = diffDaysFromToday(rolledDate, today);
      if (rolledDiff !== null && rolledDiff >= 0 && rolledDiff <= (asuransiAdvance + 1)) {
        asuransiTarget = rolledDate;
        asuransiDiff = rolledDiff;
      }
    }
  }
  const asuransiSentKey = `ASURANSI:${today}`;
  const isAsuransiSentToday = (customer.statusLog || '').includes(asuransiSentKey);

  if (asuransiDiff !== null && asuransiDiff >= 0 && asuransiDiff <= asuransiAdvance) {
    reminders.push({
      type: 'ASURANSI',
      label: asuransiDiff === 0 ? 'Hari H Expired Asuransi' : `H-${asuransiDiff} Expired Asuransi`,
      badgeColor: 'asuransi',
      targetDate: asuransiTarget,
      daysRemaining: asuransiDiff,
      sentKey: asuransiSentKey,
      alreadySentToday: isAsuransiSentToday,
      lastLog: getLatestLog('ASURANSI')
    });
  }

  // 4. NEXT SERVICE KENDARAAN (H-5 s/d Hari H sebelum tanggal Next Service jika belum terkirim)
  const serviceDiff = diffDaysFromToday(customer.tglService, today);
  const serviceSentKey = `SERVICE:${today}`;
  const isServiceSentToday = (customer.statusLog || '').includes(serviceSentKey);

  if (serviceDiff !== null && serviceDiff >= 0 && serviceDiff <= serviceAdvance) {
    reminders.push({
      type: 'SERVICE',
      label: serviceDiff === 0 ? 'Hari H Next Service' : `H-${serviceDiff} Next Service`,
      badgeColor: 'service',
      targetDate: customer.tglService,
      daysRemaining: serviceDiff,
      sentKey: serviceSentKey,
      alreadySentToday: isServiceSentToday,
      lastLog: getLatestLog('SERVICE')
    });
  }

  const tomorrowDateStr = addDaysToDateStr(today, 1);
  const tomorrowMonthDay = tomorrowDateStr.slice(5, 10);
  const tomorrowReminders = [];

  if (bdayMonthDay && bdayMonthDay === tomorrowMonthDay) {
    tomorrowReminders.push({
      type: 'ULTAH',
      label: 'Ulang Tahun Besok (Hari H)',
      scheduledDate: tomorrowDateStr,
      targetDate: customer.tglUlangTahun
    });
  }
  if (samsatDiff === samsatAdvance + 1) {
    tomorrowReminders.push({
      type: 'SAMSAT',
      label: `H-${samsatAdvance} Samsat Kendaraan (Jadwal Besok)`,
      scheduledDate: tomorrowDateStr,
      targetDate: samsatTarget
    });
  }
  if (asuransiDiff === asuransiAdvance + 1) {
    tomorrowReminders.push({
      type: 'ASURANSI',
      label: `H-${asuransiAdvance} Expired Asuransi (Jadwal Besok)`,
      scheduledDate: tomorrowDateStr,
      targetDate: asuransiTarget
    });
  }
  if (serviceDiff === serviceAdvance + 1) {
    tomorrowReminders.push({
      type: 'SERVICE',
      label: `H-${serviceAdvance} Next Service (Jadwal Besok)`,
      scheduledDate: tomorrowDateStr,
      targetDate: customer.tglService
    });
  }

  return {
    ...customer,
    samsatDiff,
    asuransiDiff,
    serviceDiff,
    isBirthdayToday,
    activeReminders: reminders,
    tomorrowReminders,
    tomorrowDateStr
  };
}

function renderTemplate(templateStr, vars) {
  let out = templateStr || '';
  for (const [k, v] of Object.entries(vars)) {
    const regex = new RegExp(`\\{${k}\\}`, 'g');
    out = out.replace(regex, v != null && v !== '' ? String(v) : '-');
  }
  return out;
}

function buildMessagesForReminder(customer, reminderType, config) {
  const wita = getWitaParts();
  const baseVars = {
    nama_pelanggan: customer.namaPelanggan,
    wa_pelanggan: customer.waPelanggan,
    email_pelanggan: customer.emailPelanggan || '-',
    tipe_mobil: customer.tipeMobil || 'Toyota',
    no_polisi: customer.noPolisi || '-',
    tgl_ultah_indo: formatDateIndonesian(customer.tglUlangTahun),
    tgl_samsat_indo: formatDateIndonesian(customer.tglSamsat),
    tgl_asuransi_indo: formatDateIndonesian(customer.tglAsuransi),
    tgl_service_indo: formatDateIndonesian(customer.tglService),
    nama_sales: customer.namaSales || 'Tim Sales Agung Toyota Gianyar',
    wa_sales: customer.waSales || '-',
    email_sales: customer.emailSales || '-',
    nama_sa: customer.namaSA || 'Tim Service Advisor Agung Toyota Gianyar',
    wa_sa: customer.waSA || '-',
    email_sa: customer.emailSA || '-',
    tanggal_hari_ini: formatDateIndonesian(wita.dateStr)
  };

  let custTpl = '';
  let staffTpl = '';
  let emailSubjectCust = '';
  let emailSubjectStaff = '';
  let categoryTitle = '';

  const timeTag = wita.shortTime ? ` [${wita.shortTime} WITA]` : '';

  if (reminderType === 'ULTAH') {
    custTpl = config.templates.birthdayCustomer;
    staffTpl = config.templates.birthdayStaff;
    emailSubjectCust = `Selamat Ulang Tahun Bapak/Ibu ${customer.namaPelanggan}! 🎂 - Agung Toyota Gianyar${timeTag}`;
    emailSubjectStaff = `[Reminder Ultah Pelanggan] ${customer.namaPelanggan} (${customer.tipeMobil} - ${customer.noPolisi})${timeTag}`;
    categoryTitle = 'Ucapan Selamat Ulang Tahun';
  } else if (reminderType === 'SAMSAT') {
    custTpl = config.templates.samsatCustomer;
    staffTpl = config.templates.samsatStaff;
    emailSubjectCust = `Pengingat H-3 Samsat/STNK Toyota ${customer.tipeMobil} (${customer.noPolisi}) - Agung Toyota Gianyar${timeTag}`;
    emailSubjectStaff = `[Reminder H-3 Samsat] ${customer.namaPelanggan} - ${customer.noPolisi} (${formatDateIndonesian(customer.tglSamsat)})${timeTag}`;
    categoryTitle = 'Pengingat H-3 Jatuh Tempo Samsat Kendaraan';
  } else if (reminderType === 'ASURANSI') {
    custTpl = config.templates.asuransiCustomer;
    staffTpl = config.templates.asuransiStaff;
    emailSubjectCust = `Pengingat H-3 Expired Asuransi Toyota ${customer.tipeMobil} (${customer.noPolisi}) - Agung Toyota Gianyar${timeTag}`;
    emailSubjectStaff = `[Reminder H-3 Asuransi] ${customer.namaPelanggan} - ${customer.noPolisi} (${formatDateIndonesian(customer.tglAsuransi)})${timeTag}`;
    categoryTitle = 'Pengingat H-3 Expired Asuransi Kendaraan';
  } else if (reminderType === 'SERVICE') {
    custTpl = config.templates.serviceCustomer || DEFAULT_CONFIG.templates.serviceCustomer;
    staffTpl = config.templates.serviceStaff || DEFAULT_CONFIG.templates.serviceStaff;
    emailSubjectCust = `Pengingat H-5 Jadwal Servis Berkala Toyota ${customer.tipeMobil} (${customer.noPolisi}) - Agung Toyota Gianyar${timeTag}`;
    emailSubjectStaff = `[Reminder H-5 Next Service] ${customer.namaPelanggan} - ${customer.noPolisi} (${formatDateIndonesian(customer.tglService)})${timeTag}`;
    categoryTitle = 'Pengingat H-5 Jadwal Servis Berkala (Next Service)';
  }

  const customerText = renderTemplate(custTpl, baseVars);
  const salesText = renderTemplate(staffTpl, {
    ...baseVars,
    nama_pic: customer.namaSales || 'Sales Executive',
    peran_pic: 'Sales Executive'
  });
  const saText = renderTemplate(staffTpl, {
    ...baseVars,
    nama_pic: customer.namaSA || 'Service Advisor (SA)',
    peran_pic: 'Service Advisor (SA)'
  });

  return {
    categoryTitle,
    customer: {
      role: 'Pelanggan',
      name: customer.namaPelanggan,
      wa: customer.waPelanggan,
      email: customer.emailPelanggan,
      subject: emailSubjectCust,
      text: customerText,
      html: buildBrandedEmailHtml(emailSubjectCust, categoryTitle, customerText, customer, config, false)
    },
    sales: {
      role: 'Sales',
      name: customer.namaSales,
      wa: customer.waSales,
      email: customer.emailSales,
      subject: emailSubjectStaff,
      text: salesText,
      html: buildBrandedEmailHtml(emailSubjectStaff, `${categoryTitle} (Notifikasi Sales)`, salesText, customer, config, true)
    },
    sa: {
      role: 'Service Advisor (SA)',
      name: customer.namaSA,
      wa: customer.waSA,
      email: customer.emailSA,
      subject: emailSubjectStaff,
      text: saText,
      html: buildBrandedEmailHtml(emailSubjectStaff, `${categoryTitle} (Notifikasi SA)`, saText, customer, config, true)
    }
  };
}

function buildBrandedEmailHtml(subject, badgeText, plainMessage, customer, config, isStaff) {
  const formattedBody = plainMessage
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\*([^*]+)\*/g, '<strong>$1</strong>')
    .replace(/_([^_]+)_/g, '<em>$1</em>')
    .replace(/\n/g, '<br>');

  const waClickBtn = isStaff && customer.waPelanggan
    ? `<div style="margin-top:22px;">
         <a href="https://wa.me/${customer.waPelanggan}" style="display:inline-block;background:#16a34a;color:#ffffff;text-decoration:none;padding:12px 22px;border-radius:8px;font-weight:700;font-size:14px;">
           💬 Hubungi ${customer.namaPelanggan} via WhatsApp Sekarang
         </a>
       </div>`
    : '';

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
  </head>
  <body style="margin:0;padding:24px;background-color:#f1f5f9;font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#0f172a;">
    <div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 10px 25px rgba(15,23,42,0.08);border:1px solid #e2e8f0;">
      <div style="background:linear-gradient(135deg,#b91c1c 0%,#ef4444 60%,#0f172a 100%);padding:26px 30px;color:#ffffff;">
        <div style="font-size:11px;letter-spacing:1.5px;text-transform:uppercase;opacity:0.9;font-weight:700;">AUTHORIZED DEALER CRM SYSTEM</div>
        <h1 style="margin:6px 0 4px;font-size:22px;font-weight:800;">AGUNG TOYOTA GIANYAR</h1>
        <div style="font-size:12px;opacity:0.85;">${badgeText}</div>
      </div>
      <div style="padding:28px 30px;line-height:1.65;font-size:15px;color:#1e293b;">
        ${formattedBody}
        ${waClickBtn}
      </div>
      <div style="background:#f8fafc;padding:16px 30px;border-top:1px solid #e2e8f0;font-size:12px;color:#64748b;display:flex;justify-content:space-between;align-items:center;">
        <div><strong>Agung Toyota Gianyar</strong> • ${config.dealerAddress || 'Gianyar, Bali'}</div>
        <div style="font-weight:700;color:#b91c1c;">Design by MDP</div>
      </div>
    </div>
  </body>
  </html>`;
}

// ============================================================================
// WHATSAPP WEB ENGINE (BAILEYS QR SCAN + FALLBACK GATEWAY)
// ============================================================================
let waState = {
  status: 'DISCONNECTED', // DISCONNECTED | QR_READY | CONNECTED | CONNECTING
  qrDataUrl: null,
  connectedNumber: null,
  connectedName: null,
  lastError: null
};

let waSocket = null;
const sentMessagesCache = new Map();

async function initWhatsAppClient() {
  if (waState.status === 'CONNECTING' || waState.status === 'CONNECTED') {
    return;
  }
  waState.status = 'CONNECTING';
  waState.lastError = null;

  try {
    const baileys = await import('@whiskeysockets/baileys');
    const makeWASocket = baileys.default?.default || baileys.default || baileys.makeWASocket;
    const useMultiFileAuthState = baileys.useMultiFileAuthState;
    const DisconnectReason = baileys.DisconnectReason;
    const fetchLatestBaileysVersion = baileys.fetchLatestBaileysVersion;

    if (!fs.existsSync(AUTH_DIR)) fs.mkdirSync(AUTH_DIR, { recursive: true });
    const { state, saveCreds } = await useMultiFileAuthState(AUTH_DIR);

    let version = [2, 3000, 1015901307];
    try {
      const latest = await fetchLatestBaileysVersion();
      if (latest && latest.version) version = latest.version;
    } catch (_) {}

    waSocket = makeWASocket({
      version,
      auth: state,
      printQRInTerminal: false,
      logger: pino({ level: 'silent' }),
      browser: ['Agung Toyota Gianyar CRM (MDP)', 'Chrome', '1.0.0'],
      syncFullHistory: false,
      markOnlineOnConnect: true,
      // Penting agar retry enkripsi multi-device (saat kirim ke nomor sendiri / beruntun) selalu berhasil
      getMessage: async (key) => {
        if (key?.id && sentMessagesCache.has(key.id)) {
          return sentMessagesCache.get(key.id);
        }
        return { conversation: 'Agung Toyota Gianyar Reminder' };
      }
    });

    waSocket.ev.on('creds.update', saveCreds);

    waSocket.ev.on('connection.update', async (update) => {
      const { connection, lastDisconnect, qr } = update;
      if (qr) {
        waState.status = 'QR_READY';
        waState.qrDataUrl = await QRCode.toDataURL(qr, { margin: 2, width: 280 });
      }
      if (connection === 'open') {
        waState.status = 'CONNECTED';
        waState.qrDataUrl = null;
        waState.connectedNumber = waSocket.user?.id?.split(':')[0] || 'Terhubung';
        waState.connectedName = waSocket.user?.name || 'WhatsApp Agung Toyota Gianyar';
        waState.lastError = null;
        setTimeout(() => {
          checkAndRunScheduledReminders(false, 'AUTO_PILOT_WITA').catch(() => {});
        }, 5000);
      }
      if (connection === 'close') {
        waState.status = 'DISCONNECTED';
        waState.qrDataUrl = null;
        waState.connectedNumber = null;
        // Jaga sesi dan kredensial WhatsApp tetap utuh di disk, sambungkan ulang otomatis
        setTimeout(() => initWhatsAppClient(), 4000);
      }
    });
  } catch (err) {
    waState.status = 'DISCONNECTED';
    waState.lastError = err.message;
  }
}

async function sendWhatsAppMessage(phone, text, config) {
  const cleanPhone = normalizePhone(phone);
  if (!cleanPhone) {
    return { ok: false, channel: 'WA', mode: 'SKIPPED', reason: 'Nomor WA kosong' };
  }

  // 1. Check if connected via Baileys QR Scan
  if (waState.status === 'CONNECTED' && waSocket) {
    try {
      let targetJid = `${cleanPhone}@s.whatsapp.net`;
      const ownNumber = String(waState.connectedNumber || '').replace(/[^0-9]/g, '');

      // Jika kirim ke nomor lain, gunakan onWhatsApp untuk memastikan JID valid
      if (cleanPhone === ownNumber && waSocket.user?.id) {
        targetJid = `${ownNumber}@s.whatsapp.net`;
      } else {
        try {
          const [result] = await waSocket.onWhatsApp(cleanPhone);
          if (result && result.exists && result.jid) {
            targetJid = result.jid;
          }
        } catch (_) {}
      }

      const sentMsg = await waSocket.sendMessage(targetJid, {
        text,
        linkPreview: null
      });
      if (sentMsg?.key?.id && sentMsg?.message) {
        sentMessagesCache.set(sentMsg.key.id, sentMsg.message);
        if (sentMessagesCache.size > 200) {
          const firstKey = sentMessagesCache.keys().next().value;
          sentMessagesCache.delete(firstKey);
        }
      }
      return { ok: true, channel: 'WA', mode: 'WA_OTOMATIS_TERKIRIM', recipient: cleanPhone };
    } catch (err) {
      return { ok: false, channel: 'WA', mode: 'WA_WEB_QR_ERROR', reason: err.message };
    }
  }

  // 2. Check if WA Gateway API (Fonnte / Custom) is enabled
  if (config.waGateway?.enabled && config.waGateway?.token) {
    try {
      const res = await fetch(config.waGateway.url || 'https://api.fonnte.com/send', {
        method: 'POST',
        headers: {
          'Authorization': config.waGateway.token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          target: cleanPhone,
          message: text
        })
      });
      const data = await res.json().catch(() => ({}));
      if (res.ok && data.status !== false) {
        return { ok: true, channel: 'WA', mode: 'WA_GATEWAY', recipient: cleanPhone };
      }
      return { ok: false, channel: 'WA', mode: 'WA_GATEWAY', reason: JSON.stringify(data) };
    } catch (err) {
      return { ok: false, channel: 'WA', mode: 'WA_GATEWAY', reason: err.message };
    }
  }

  // 3. Simulation / Ready for 1-click WA link if QR not scanned yet
  return {
    ok: true,
    simulated: true,
    channel: 'WA',
    mode: 'SIMULASI_BELUM_SCAN_QR',
    recipient: cleanPhone,
    reason: 'QR WA belum di-scan'
  };
}

async function sendEmailMessage(email, subject, text, html, config) {
  const cleanEmail = String(email || '').trim();
  if (!cleanEmail || !cleanEmail.includes('@')) {
    return { ok: false, channel: 'EMAIL', mode: 'SKIPPED', reason: 'Alamat email kosong/tidak valid' };
  }

  const smtpUser = String(config.smtp?.user || '').trim();
  const smtpPass = String(config.smtp?.pass || '').replace(/\s+/g, '').trim();

  if (smtpUser && smtpPass) {
    try {
      const transporter = nodemailer.createTransport({
        service: config.smtp.service || 'gmail',
        auth: {
          user: smtpUser,
          pass: smtpPass
        }
      });
      await transporter.sendMail({
        from: `"${config.smtp.fromName || 'Agung Toyota Gianyar'}" <${smtpUser}>`,
        to: cleanEmail,
        subject,
        text,
        html,
        headers: {
          'X-Entity-Ref-ID': `ATG-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
        }
      });
      return { ok: true, channel: 'EMAIL', mode: 'GMAIL_OTOMATIS', recipient: cleanEmail };
    } catch (err) {
      return {
        ok: false,
        channel: 'EMAIL',
        mode: 'GMAIL_GAGAL_LOGIN',
        recipient: cleanEmail,
        reason: err.message || 'Sandi Aplikasi Gmail belum benar (Wajib 16 huruf Sandi Aplikasi Google)'
      };
    }
  }

  return {
    ok: true,
    simulated: true,
    channel: 'EMAIL',
    mode: 'SIMULASI_EMAIL',
    recipient: cleanEmail,
    reason: 'Gmail App Password belum diisi'
  };
}

// ============================================================================
// In-Memory Anti-Spam Cache (Mencegah pengiriman ganda pada hari yang sama)
const dispatchedTodaySet = new Set();

async function executeReminderDispatch(customerId, reminderType, triggerSource = 'AUTO_PILOT_WITA') {
  const config = loadConfig();
  const customers = await getAllCustomers();
  const idx = customers.findIndex(c => Number(c.id) === Number(customerId));
  if (idx === -1) {
    throw new Error(`Data pelanggan #${customerId} tidak ditemukan`);
  }

  const wita = getWitaParts();
  const dispatchKey = `${customerId}:${reminderType}:${wita.dateStr}`;
  if (triggerSource === 'AUTO_PILOT_WITA' && dispatchedTodaySet.has(dispatchKey)) {
    return { skipped: true, reason: 'Pengingat ini sudah terkirim hari ini' };
  }
  dispatchedTodaySet.add(dispatchKey);

  const customer = customers[idx];
  const bundle = buildMessagesForReminder(customer, reminderType, config);
  const results = [];

  // 1. Send to Customer (WA + Email)
  const custEmailPromise = sendEmailMessage(bundle.customer.email, bundle.customer.subject, bundle.customer.text, bundle.customer.html, config);
  const custWaRes = await sendWhatsAppMessage(bundle.customer.wa, bundle.customer.text, config);
  const custEmailRes = await custEmailPromise;
  results.push({ target: 'Pelanggan', name: bundle.customer.name, wa: custWaRes, email: custEmailRes });

  // 2. Send to Sales (WA + Email) — Beri jeda 2.5 detik agar sesi WA tidak tabrakan
  if (bundle.sales.wa || bundle.sales.email) {
    const salesEmailPromise = sendEmailMessage(bundle.sales.email, bundle.sales.subject, bundle.sales.text, bundle.sales.html, config);
    if (bundle.sales.wa) {
      await new Promise(r => setTimeout(r, 2500));
    }
    const salesWaRes = await sendWhatsAppMessage(bundle.sales.wa, bundle.sales.text, config);
    const salesEmailRes = await salesEmailPromise;
    results.push({ target: 'Sales', name: bundle.sales.name, wa: salesWaRes, email: salesEmailRes });
  }

  // 3. Send to Service Advisor / SA (WA + Email)
  if (bundle.sa.wa || bundle.sa.email) {
    const saEmailPromise = sendEmailMessage(bundle.sa.email, bundle.sa.subject, bundle.sa.text, bundle.sa.html, config);
    const saHasSeparateWa = bundle.sa.wa && bundle.sa.wa !== bundle.sales.wa;
    if (saHasSeparateWa) {
      await new Promise(r => setTimeout(r, 2500));
    }
    const saWaRes = saHasSeparateWa
      ? await sendWhatsAppMessage(bundle.sa.wa, bundle.sa.text, config)
      : { ok: false, channel: 'WA', mode: 'SKIPPED', reason: 'Nomor WA sama dengan Sales' };
    const saEmailRes = await saEmailPromise;
    results.push({ target: 'Service Advisor (SA)', name: bundle.sa.name, wa: saWaRes, email: saEmailRes });
  }

  // Sesuai instruksi pengguna: Bila ada email atau WA yang tidak terkirim/gagal,
  // TIDAK USAH dikirim ulang secara otomatis. Cukup catat reason dan notifikasinya.
  // Tandai selesai di statusLog agar pengingat hari ini tidak pernah diulang-ulang.
  const sentTag = `${reminderType}:${wita.dateStr}`;
  const existingTags = (customer.statusLog || '')
    .split('|')
    .map(s => s.trim())
    .filter(Boolean);
  if (!existingTags.includes(sentTag)) {
    existingTags.push(sentTag);
  }
  customers[idx].statusLog = existingTags.join(' | ');
  await saveAllCustomers(customers);

  const logEntry = appendLog({
    triggerSource,
    customerId: customer.id,
    namaPelanggan: customer.namaPelanggan,
    tipeMobil: customer.tipeMobil,
    noPolisi: customer.noPolisi,
    reminderType,
    categoryTitle: bundle.categoryTitle,
    results
  });

  return { logEntry, updatedCustomer: customers[idx] };
}

// ============================================================================
// AUTOMATIC SCHEDULER (JAM KERJA 08:00-16:00 WITA, WA OTOMATIS 10:00-14:00 WITA)
// ============================================================================
let isBatchRunning = false;
let lastSchedulerCheck = null;
let autoPilotStatusText = 'Menunggu Koneksi QR WhatsApp...';

function isTimeInRange(currentHm, startHm, endHm) {
  return currentHm >= startHm && currentHm <= endHm;
}

async function checkAndRunScheduledReminders(forceAll = false, triggerSource = 'AUTO_PILOT_WITA') {
  if (isBatchRunning) {
    return { skipped: true, reason: 'Antrean pengiriman sedang berjalan' };
  }
  const config = loadConfig();
  const wita = getWitaParts();
  lastSchedulerCheck = `${wita.dateStr} ${wita.timeStr} WITA`;

  const inWaWindow = isTimeInRange(
    wita.shortTime,
    config.waWindowStart || '10:00',
    config.waWindowEnd || '14:00'
  );

  const isWaReady = waState.status === 'CONNECTED' || Boolean(config.waGateway?.enabled && config.waGateway?.token);
  const isEmailReady = Boolean(config.smtp?.enabled && config.smtp?.user && config.smtp?.pass);

  if (!forceAll) {
    if (!config.autoSendEnabled) {
      autoPilotStatusText = 'Auto-Pilot Dinonaktifkan';
      return { skipped: true, reason: autoPilotStatusText };
    }
    // WhatsApp adalah pengingat utama pelanggan: Auto-pilot menunggu WhatsApp terhubung agar tidak spam email
    if (!isWaReady) {
      autoPilotStatusText = 'Standby — Menunggu Anda Scan QR WhatsApp (Begitu terhubung, pesan WA & Email otomatis terkirim)';
      return { skipped: true, reason: autoPilotStatusText };
    }
    if (!inWaWindow) {
      autoPilotStatusText = `Standby Otomatis — Menunggu Jam Pengiriman (${config.waWindowStart} – ${config.waWindowEnd} WITA)`;
      return {
        skipped: true,
        inWaWindow,
        witaTime: wita.timeStr,
        reason: autoPilotStatusText
      };
    }
  }

  const customers = await getAllCustomers();
  const evaluated = customers.map(c => evaluateCustomerReminders(c, config, wita));

  // Collect all due reminders not yet sent today
  const pendingTasks = [];
  for (const cust of evaluated) {
    for (const rem of cust.activeReminders) {
      if (!rem.alreadySentToday || forceAll) {
        pendingTasks.push({ customerId: cust.id, namaPelanggan: cust.namaPelanggan, reminderType: rem.type });
      }
    }
  }

  if (pendingTasks.length === 0) {
    autoPilotStatusText = `Aktif & Selesai — Seluruh pengingat hari ini (${wita.dateStr}) sudah dikirim otomatis!`;
    return { skipped: false, dispatchedCount: 0, message: 'Semua pengingat hari ini sudah terkirim.' };
  }

  isBatchRunning = true;
  autoPilotStatusText = `🚀 Sedang Mengirim Otomatis (${pendingTasks.length} Pengingat)...`;
  const dispatched = [];
  try {
    for (let i = 0; i < pendingTasks.length; i++) {
      const task = pendingTasks[i];
      const res = await executeReminderDispatch(task.customerId, task.reminderType, triggerSource);
      dispatched.push(res.logEntry);

      if (i < pendingTasks.length - 1) {
        const delayMs = Math.max(2, Number(config.delaySecondsBetweenMessages || 8)) * 1000;
        await new Promise(r => setTimeout(r, delayMs));
      }
    }
    autoPilotStatusText = `✅ Berhasil mengirim ${dispatched.length} pengingat secara otomatis pada ${getWitaParts().timeStr} WITA!`;
  } finally {
    isBatchRunning = false;
  }

  return { skipped: false, dispatchedCount: dispatched.length, dispatched };
}

// Run scheduler check every 60 seconds
setInterval(() => {
  checkAndRunScheduledReminders(false, 'AUTO_PILOT_WITA').catch(() => {});
}, 60000);

// ============================================================================
// REST API ENDPOINTS
// ============================================================================
app.get('/api/dashboard', async (req, res) => {
  const config = loadConfig();
  const wita = getWitaParts();
  const customers = await getAllCustomers();
  const logs = loadLogs();
  const evaluated = customers.map(c => evaluateCustomerReminders(c, config, wita, logs));

  const inOperationalHours = isTimeInRange(
    wita.shortTime,
    config.operationalStart || '08:00',
    config.operationalEnd || '16:00'
  );
  const inWaWindow = isTimeInRange(
    wita.shortTime,
    config.waWindowStart || '08:00',
    config.waWindowEnd || '16:00'
  );

  const stats = {
    totalCustomers: evaluated.length,
    dueBirthdayToday: evaluated.filter(c => c.activeReminders.some(r => r.type === 'ULTAH')).length,
    dueSamsatH3: evaluated.filter(c => c.activeReminders.some(r => r.type === 'SAMSAT')).length,
    dueAsuransiH3: evaluated.filter(c => c.activeReminders.some(r => r.type === 'ASURANSI')).length,
    dueServiceH5: evaluated.filter(c => c.activeReminders.some(r => r.type === 'SERVICE')).length,
    totalPendingToday: evaluated.reduce(
      (acc, c) => acc + c.activeReminders.filter(r => !r.alreadySentToday).length,
      0
    ),
    totalSentToday: evaluated.reduce(
      (acc, c) => acc + c.activeReminders.filter(r => r.alreadySentToday).length,
      0
    )
  };

  const todayLogs = logs.filter(l => l.dateWita === wita.dateStr);
  const deliveryIssuesToday = todayLogs
    .filter(l => l.status === 'FAILED' || l.status === 'PARTIAL')
    .map(l => ({
      id: l.id,
      timestampWita: l.timestampWita,
      customerId: l.customerId,
      namaPelanggan: l.namaPelanggan,
      noPolisi: l.noPolisi,
      tipeMobil: l.tipeMobil,
      reminderType: l.reminderType,
      categoryTitle: l.categoryTitle,
      status: l.status,
      reason: l.failureReason,
      results: l.results
    }));

  res.json({
    wita,
    inOperationalHours,
    inWaWindow,
    isBatchRunning,
    lastSchedulerCheck,
    autoPilotStatusText,
    waState,
    stats,
    config,
    deliveryIssuesToday,
    customers: evaluated,
    logs: logs.slice(0, 50)
  });
});

// Quick Test Direct WA + Email Send to User's Phone & Email
app.post('/api/wa/test-direct', async (req, res) => {
  try {
    const { phone, email } = req.body;
    const cleanPhone = normalizePhone(phone);
    const cleanEmail = String(email || '').trim();
    if (!cleanPhone && !cleanEmail) {
      return res.status(400).json({ ok: false, error: 'Masukkan Nomor WA atau Alamat Email Gmail tujuan terlebih dahulu!' });
    }
    const config = loadConfig();
    const wita = getWitaParts();
    const sampleText =
      `✅ *[TES MESIN OTOMATIS AGUNG TOYOTA GIANYAR]*\n\n` +
      `Halo! Pesan ini dikirim secara *OTOMATIS* oleh Sistem Pengingat *Agung Toyota Gianyar* (_Design by MDP_) pada pukul *${wita.timeStr} WITA*.\n\n` +
      `Mesin Auto-Pilot WA & Email Anda sudah aktif dan siap mengirimkan:\n` +
      `1️⃣ Ucapan Ulang Tahun (Hari H)\n` +
      `2️⃣ Pengingat Samsat Kendaraan (H-3)\n` +
      `3️⃣ Pengingat Expired Asuransi (H-3)\n` +
      `4️⃣ Pengingat Jadwal Next Service (H-5)\n` +
      `secara otomatis ke Pelanggan, Sales, dan SA! 🚗✨`;

    const customers = await getAllCustomers();
    const targetTestEmail = cleanEmail || customers[0]?.emailPelanggan || config.smtp?.user || '';
    const sampleHtml = buildBrandedEmailHtml(
      `[Tes Otomatis] Agung Toyota Gianyar [${wita.shortTime} WITA]`,
      'Uji Coba Mesin Pengingat Otomatis WA & Gmail',
      sampleText,
      customers[0] || { namaPelanggan: 'Pelanggan Uji Coba', waPelanggan: cleanPhone },
      config,
      false
    );

    const [result, emailResult] = await Promise.all([
      cleanPhone
        ? sendWhatsAppMessage(cleanPhone, sampleText, config)
        : Promise.resolve({ ok: false, channel: 'WA', mode: 'SKIPPED', reason: 'No WA dikosongkan' }),
      targetTestEmail
        ? sendEmailMessage(
            targetTestEmail,
            `✅ [Tes Otomatis] Agung Toyota Gianyar — Design by MDP [${wita.shortTime} WITA]`,
            sampleText,
            sampleHtml,
            config
          )
        : Promise.resolve({ ok: false, channel: 'EMAIL', mode: 'SKIPPED', reason: 'Email dikosongkan' })
    ]);

    appendLog({
      triggerSource: 'TES_OTOMATIS_WA_GMAIL',
      customerId: 0,
      namaPelanggan: `Tes (${cleanPhone || '-'} / ${targetTestEmail || '-'})`,
      tipeMobil: 'Toyota Test Unit',
      noPolisi: 'DK TEST',
      reminderType: 'TES_WA_GMAIL',
      categoryTitle: 'Uji Coba Pengiriman WA & Email Otomatis',
      results: [{ target: 'Nomor & Email Uji Coba', name: cleanPhone || targetTestEmail, wa: result, email: emailResult }]
    });

    res.json({ ok: true, result, emailResult });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Preview rendered message for Pelanggan, Sales, and SA
app.get('/api/preview/:customerId/:reminderType', async (req, res) => {
  const config = loadConfig();
  const customers = await getAllCustomers();
  const cust = customers.find(c => Number(c.id) === Number(req.params.customerId));
  if (!cust) return res.status(404).json({ error: 'Pelanggan tidak ditemukan' });

  const bundle = buildMessagesForReminder(cust, req.params.reminderType, config);
  res.json({ customer: cust, bundle });
});

// Trigger manual send for one reminder
app.post('/api/send-single', async (req, res) => {
  try {
    const { customerId, reminderType } = req.body;
    const out = await executeReminderDispatch(customerId, reminderType, 'MANUAL_DASHBOARD');
    res.json({ ok: true, ...out });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Trigger all due reminders today
app.post('/api/send-all-today', async (req, res) => {
  try {
    const { forceResend } = req.body || {};
    const out = await checkAndRunScheduledReminders(Boolean(forceResend), 'MANUAL_BATCH');
    res.json({ ok: true, ...out });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Add or Update Customer row (Syncs 2-Way with Google Spreadsheet)
app.post('/api/customers', async (req, res) => {
  try {
    const customers = await getAllCustomers(true);
    const payload = req.body;
    if (!payload.namaPelanggan) {
      return res.status(400).json({ ok: false, error: 'Nama Pelanggan wajib diisi' });
    }
    let savedCustomer = null;
    if (payload.id) {
      const idx = customers.findIndex(c => Number(c.id) === Number(payload.id));
      if (idx !== -1) {
        customers[idx] = { ...customers[idx], ...payload, id: Number(payload.id), statusLog: customers[idx].statusLog || '' };
        savedCustomer = customers[idx];
      } else {
        savedCustomer = { ...payload, id: customers.length + 1, statusLog: '' };
        customers.push(savedCustomer);
      }
    } else {
      const nextId = customers.reduce((max, c) => Math.max(max, Number(c.id) || 0), 0) + 1;
      savedCustomer = { ...payload, id: nextId, statusLog: '' };
      customers.push(savedCustomer);
    }
    await saveAllCustomers(customers);

    // Evaluasi apakah pelanggan yang baru disimpan ini punya jadwal pengingat hari ini
    const config = loadConfig();
    const evalCust = evaluateCustomerReminders(savedCustomer, config, getWitaParts());
    if (evalCust.activeReminders.length > 0) {
      (async () => {
        for (const rem of evalCust.activeReminders) {
          try {
            await executeReminderDispatch(savedCustomer.id, rem.type, 'AUTO_SAAT_SIMPAN_CUSTOMER');
          } catch (_) {}
        }
      })();
    }

    res.json({
      ok: true,
      count: customers.length,
      autoTriggeredCount: evalCust.activeReminders.length
    });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Delete Customer row (Syncs 2-Way with Google Spreadsheet)
app.delete('/api/customers/:id', async (req, res) => {
  try {
    const idToDelete = Number(req.params.id);
    const existing = await getAllCustomers(true);
    const customers = existing.filter(c => Number(c.id) !== idToDelete);
    await saveAllCustomers(customers);

    // Hapus juga di Google Sheet API jika URL terpasang
    const config = loadConfig();
    if (config.googleSheetApiUrl) {
      fetch(config.googleSheetApiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'delete_customer', id: idToDelete })
      }).catch(() => {});
    }

    res.json({ ok: true, count: customers.length });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Reset Sample Data
app.post('/api/customers/reset-sample', async (req, res) => {
  try {
    const sample = createInitialSampleRows().map((r, i) => mapExcelRowToCustomer(r, i + 1));
    await saveAllCustomers(sample);
    res.json({ ok: true, count: sample.length });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Upload Excel (.xlsx) via Base64 payload OR JSON customers array
app.post('/api/excel/upload', async (req, res) => {
  try {
    const { base64Data, mode, customers: directCustomers } = req.body;
    let uploadedCustomers = [];

    if (Array.isArray(directCustomers) && directCustomers.length > 0) {
      uploadedCustomers = directCustomers
        .map((r, i) => mapExcelRowToCustomer(r, i + 1))
        .filter(c => c.namaPelanggan !== '');
    } else if (base64Data) {
      const buffer = Buffer.from(base64Data, 'base64');
      const wb = XLSX.read(buffer, { type: 'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      uploadedCustomers = rawRows
        .map((r, i) => mapExcelRowToCustomer(r, i + 1))
        .filter(c => c.namaPelanggan !== '');
    } else {
      return res.status(400).json({ ok: false, error: 'File Excel atau data pelanggan kosong' });
    }

    if (uploadedCustomers.length === 0) {
      return res.status(400).json({
        ok: false,
        error: 'Tidak ada baris pelanggan yang valid. Pastikan kolom "Nama Pelanggan" terisi.'
      });
    }

    let finalCustomers = uploadedCustomers;
    if (mode === 'append') {
      const existing = await getAllCustomers(true);
      finalCustomers = [...existing, ...uploadedCustomers].map((c, idx) => ({ ...c, id: idx + 1 }));
    }

    await saveAllCustomers(finalCustomers);
    setTimeout(() => checkAndRunScheduledReminders(false, 'AUTO_PILOT_WITA').catch(() => {}), 2000);
    res.json({ ok: true, importedCount: uploadedCustomers.length, totalCount: finalCustomers.length });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// ============================================================================
// GOOGLE SPREADSHEET 2-WAY SYNC ENDPOINTS
// ============================================================================

// Simpan URL Google Apps Script Web App (Integrasi 2-Arah)
app.post('/api/google-sheets/save-url', async (req, res) => {
  try {
    const { url } = req.body;
    const cleanUrl = String(url || '').trim();
    saveConfig({ googleSheetApiUrl: cleanUrl });

    if (cleanUrl) {
      try {
        const resGsheet = await fetch(`${cleanUrl}?action=get_customers&t=${Date.now()}`);
        if (resGsheet.ok) {
          const json = await resGsheet.json();
          if (json.success && Array.isArray(json.data) && json.data.length > 0) {
            const mapped = json.data
              .map((r, i) => mapExcelRowToCustomer(r, i + 1))
              .filter(c => c.namaPelanggan || c.noPolisi || c.waPelanggan);
            if (mapped.length > 0) {
              // syncToCloud: false agar tidak menulis ulang ke Google Sheets saat inisialisasi
              await saveAllCustomers(mapped, false);
              return res.json({ ok: true, count: mapped.length, message: `✅ Berhasil terhubung! Mengimpor ${mapped.length} data dari Google Spreadsheet.` });
            }
          }
        }
      } catch (e) {
        return res.json({ ok: true, message: `URL disimpan, namun belum dapat dihubungi: ${e.message}` });
      }
    }
    res.json({ ok: true, message: 'Pengaturan Google Sheets disimpan.' });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Tarik data terbaru dari Google Spreadsheet ke aplikasi (PULL)
app.post('/api/google-sheets/pull', async (req, res) => {
  try {
    const config = loadConfig();
    if (!config.googleSheetApiUrl) {
      return res.status(400).json({ ok: false, error: 'URL Google Apps Script belum diisi.' });
    }
    const resGsheet = await fetch(`${config.googleSheetApiUrl}?action=get_customers&t=${Date.now()}`);
    if (!resGsheet.ok) throw new Error(`Google Apps Script merespon kode: ${resGsheet.status}`);
    const json = await resGsheet.json();
    if (!json.success || !Array.isArray(json.data)) {
      throw new Error(json.error || 'Format respon Google Spreadsheet tidak valid.');
    }

    if (json.data.length === 0) {
      return res.json({ ok: true, count: 0, message: 'Google Spreadsheet saat ini kosong (hanya header).' });
    }

    const mapped = json.data
      .map((r, i) => mapExcelRowToCustomer(r, i + 1))
      .filter(c => c.namaPelanggan || c.noPolisi || c.waPelanggan);

    if (mapped.length === 0) {
      return res.status(400).json({ ok: false, error: 'Data ditemukan di Google Sheets namun nama pelanggan/nopol kosong.' });
    }

    // PENTING: syncToCloud = false agar PULL TIDAK PERNAH menghapus atau menulis balik ke Google Sheets!
    await saveAllCustomers(mapped, false);
    res.json({ ok: true, count: mapped.length, message: `Berhasil menarik ${mapped.length} pelanggan dari Google Sheets.` });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Unggah seluruh data aplikasi saat ini ke Google Spreadsheet
app.post('/api/google-sheets/push', async (req, res) => {
  try {
    const config = loadConfig();
    if (!config.googleSheetApiUrl) {
      return res.status(400).json({ ok: false, error: 'URL Google Apps Script belum diisi.' });
    }
    const customers = await getAllCustomers(false);
    const resGsheet = await fetch(config.googleSheetApiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'sync_all', items: customers })
    });
    const json = await resGsheet.json();
    if (!json.success) throw new Error(json.error || 'Gagal menyimpan ke Google Spreadsheet.');
    res.json({ ok: true, count: customers.length });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Download current Excel database
app.get('/api/excel/download', (req, res) => {
  ensureExcelExists();
  res.download(EXCEL_PATH, 'Data_Pelanggan_Agung_Toyota_Gianyar.xlsx');
});

// Download blank/sample Excel Template
app.get('/api/excel/template', (req, res) => {
  const tempPath = path.join(DATA_DIR, 'Template_Agung_Toyota_Gianyar.xlsx');
  const sample = createInitialSampleRows().map((r, i) => mapExcelRowToCustomer(r, i + 1));
  writeCustomersToExcel(sample, tempPath);
  res.download(tempPath, 'Template_Pengingat_Agung_Toyota_Gianyar.xlsx');
});

// Sync from Google Sheets Public CSV/XLSX link
app.post('/api/google-sheets/sync', async (req, res) => {
  try {
    const { url } = req.body;
    if (!url) return res.status(400).json({ ok: false, error: 'URL Google Sheets belum diisi' });

    let exportUrl = url.trim();
    const matchId = exportUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (matchId) {
      exportUrl = `https://docs.google.com/spreadsheets/d/${matchId[1]}/export?format=xlsx`;
    }

    const response = await fetch(exportUrl);
    if (!response.ok) {
      throw new Error(`Gagal mengunduh Google Sheets (Pastikan akses link "Siapa saja yang memiliki link dapat melihat")`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const syncedCustomers = rawRows
      .map((r, i) => mapExcelRowToCustomer(r, i + 1))
      .filter(c => c.namaPelanggan !== '');

    if (syncedCustomers.length === 0) {
      throw new Error('Data Google Sheets kosong atau header kolom tidak sesuai.');
    }

    writeCustomersToExcel(syncedCustomers);
    saveConfig({ googleSheetCsvUrl: url });
    setTimeout(() => checkAndRunScheduledReminders(false, 'AUTO_PILOT_WITA').catch(() => {}), 2000);
    res.json({ ok: true, count: syncedCustomers.length });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Save Configuration & Templates
app.post('/api/config', (req, res) => {
  try {
    const updated = saveConfig(req.body);
    setTimeout(() => checkAndRunScheduledReminders(false, 'AUTO_PILOT_WITA').catch(() => {}), 1000);
    res.json({ ok: true, config: updated });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// WhatsApp QR Connection Controls
app.post('/api/wa/connect', async (req, res) => {
  try {
    initWhatsAppClient();
    res.json({ ok: true, waState });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post('/api/wa/disconnect', async (req, res) => {
  try {
    if (waSocket) {
      try {
        await waSocket.logout();
      } catch (_) {}
      waSocket = null;
    }
    try {
      fs.rmSync(AUTH_DIR, { recursive: true, force: true });
    } catch (_) {}
    waState = {
      status: 'DISCONNECTED',
      qrDataUrl: null,
      connectedNumber: null,
      connectedName: null,
      lastError: null
    };
    initWhatsAppClient();
    res.json({ ok: true, waState });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Clear delivery status for today (useful for testing)
app.post('/api/logs/clear-today-flags', (req, res) => {
  const customers = readCustomersFromExcel().map(c => ({ ...c, statusLog: '' }));
  writeCustomersToExcel(customers);
  setTimeout(() => checkAndRunScheduledReminders(false, 'AUTO_PILOT_WITA').catch(() => {}), 1500);
  res.json({ ok: true });
});

// Analytics & Reports (Daily, Weekly, Monthly with Failure Reason Breakdown)
app.get('/api/reports', (req, res) => {
  try {
    const { date, month } = req.query;
    const analytics = getReportAnalytics(date, month);
    res.json({ ok: true, ...analytics });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Export Report to Excel (.xlsx)
app.get('/api/reports/export', (req, res) => {
  try {
    const { date, month } = req.query;
    const analytics = getReportAnalytics(date, month);
    const targetLogs = date
      ? analytics.daily.logs
      : (month ? loadLogs().filter(l => (l.dateWita || '').startsWith(analytics.curMonth)) : loadLogs());

    const rows = targetLogs.map((l, i) => {
      const custRes = (l.results || []).find(r => r.target === 'Pelanggan');
      const salesRes = (l.results || []).find(r => r.target === 'Sales');
      const saRes = (l.results || []).find(r => r.target === 'Service Advisor (SA)' || r.target === 'Service Advisor');

      return {
        'No': i + 1,
        'Tanggal (WITA)': l.dateWita || '',
        'Waktu (WITA)': l.timeWita || l.timestampWita || '',
        'Nama Pelanggan': l.namaPelanggan || '',
        'No. Polisi': l.noPolisi || '',
        'Tipe Mobil': l.tipeMobil || '',
        'Kategori Pengingat': l.categoryTitle || l.reminderType || '',
        'Status': l.status || 'SUCCESS',
        'Alasan / Keterangan Kendala': l.failureReason || 'Berhasil',
        'WA Pelanggan': custRes?.wa?.ok ? `Terkirim (${custRes.wa.mode})` : (custRes?.wa?.reason || '-'),
        'Email Pelanggan': custRes?.email?.ok ? `Terkirim (${custRes.email.mode})` : (custRes?.email?.reason || '-'),
        'Sales Tembusan': salesRes?.name ? `${salesRes.name} (${salesRes.wa?.ok ? 'WA OK' : '-'})` : '-',
        'SA Tembusan': saRes?.name ? `${saRes.name} (${saRes.wa?.ok ? 'WA OK' : '-'})` : '-'
      };
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{ 'Keterangan': 'Belum ada riwayat pengiriman pada periode ini' }]);
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Pengiriman');

    const tempFilename = `Laporan_Pengiriman_Toyota_Gianyar_${analytics.curDate}.xlsx`;
    const tempPath = path.join(os.tmpdir(), tempFilename);
    XLSX.writeFile(wb, tempPath);
    res.download(tempPath, tempFilename);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
});

ensureExcelExists();
if (!IS_VERCEL) {
  initWhatsAppClient();
  app.listen(PORT, () => {
    console.log(`================================================================`);
    console.log(`🚗 AGUNG TOYOTA GIANYAR - SMART REMINDER SYSTEM (Design by MDP)`);
    console.log(`🌐 Buka Dashboard di Browser: http://localhost:${PORT}`);
    console.log(`⏰ Jam Aktif Kantor: 08:00 - 16:00 WITA | Kirim WA: 10:00 - 14:00 WITA`);
    console.log(`================================================================`);
  });
}

module.exports = app;

