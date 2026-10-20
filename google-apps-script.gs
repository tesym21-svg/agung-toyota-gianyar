/**
 * ============================================================================
 * AGUNG TOYOTA GIANYAR — GOOGLE SHEETS + GMAIL & WA AUTO REMINDER
 * Design by MDP
 * ============================================================================
 * Cara Pakai di Google Sheets (Opsional jika ingin jalan langsung di Cloud Google):
 * 1. Buka Google Sheets Anda -> Menu "Ekstensi" -> "Apps Script"
 * 2. Tempel kode ini, lalu simpan.
 * 3. Buka menu "Pemicu" (Triggers / Ikon Jam di kiri):
 *    - Pilih fungsi: jalankanPengingatHarianWITA
 *    - Sumber acara: Dipicu waktu (Time-driven) -> Timer harian -> 10.00 hingga 11.00 WITA
 */

const CONFIG_ATG = {
  DEALER_NAME: 'Agung Toyota Gianyar',
  DESIGN_CREDIT: 'Design by MDP',
  TIMEZONE: 'Asia/Makassar', // WITA (UTC+8)
  WA_START_HOUR: 10,         // Mulai kirim WA jam 10:00 WITA
  WA_END_HOUR: 14,           // Batas akhir jam 14:00 WITA
  WA_GATEWAY_TOKEN: ''       // Opsional: Isi jika menggunakan Fonnte di Cloud
};

function jalankanPengingatHarianWITA() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return;

  const now = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG_ATG.TIMEZONE, 'yyyy-MM-dd');
  const todayMonthDay = Utilities.formatDate(now, CONFIG_ATG.TIMEZONE, 'MM-dd');
  const currentHour = Number(Utilities.formatDate(now, CONFIG_ATG.TIMEZONE, 'HH'));

  // Pastikan berjalan di rentang 10:00 - 14:00 WITA
  if (currentHour < CONFIG_ATG.WA_START_HOUR || currentHour > CONFIG_ATG.WA_END_HOUR) {
    Logger.log('Di luar jam pengiriman 10:00 - 14:00 WITA.');
    return;
  }

  const hPlus3Date = new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000);
  const hPlus3Str = Utilities.formatDate(hPlus3Date, CONFIG_ATG.TIMEZONE, 'yyyy-MM-dd');

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const namaPelanggan = String(row[1] || '').trim();
    if (!namaPelanggan) continue;

    const waPelanggan = bersihkanNomorWA(row[2]);
    const emailPelanggan = String(row[3] || '').trim();
    const tipeMobil = String(row[4] || 'Toyota').trim();
    const noPolisi = String(row[5] || '-').trim();
    const tglUltah = formatTgl(row[6]);
    const tglSamsat = formatTgl(row[7]);
    const tglAsuransi = formatTgl(row[8]);
    const namaSales = String(row[9] || 'Sales Agung Toyota Gianyar').trim();
    const waSales = bersihkanNomorWA(row[10]);
    const emailSales = String(row[11] || '').trim();
    const namaSA = String(row[12] || 'Service Advisor Agung Toyota Gianyar').trim();
    const waSA = bersihkanNomorWA(row[13]);
    const emailSA = String(row[14] || '').trim();
    let statusLog = String(row[15] || '').trim();

    const data = {
      namaPelanggan, waPelanggan, emailPelanggan, tipeMobil, noPolisi,
      tglUltah, tglSamsat, tglAsuransi, namaSales, waSales, emailSales, namaSA, waSA, emailSA
    };

    // 1. Cek Ulang Tahun (Hari H)
    if (tglUltah && tglUltah.slice(5, 10) === todayMonthDay && !statusLog.includes('ULTAH:' + todayStr)) {
      kirimPaketNotifikasi('ULTAH', data);
      statusLog += (statusLog ? ' | ' : '') + 'ULTAH:' + todayStr;
    }

    // 2. Cek Samsat Kendaraan (H-3)
    if (tglSamsat === hPlus3Str && !statusLog.includes('SAMSAT:' + todayStr)) {
      kirimPaketNotifikasi('SAMSAT', data);
      statusLog += (statusLog ? ' | ' : '') + 'SAMSAT:' + todayStr;
    }

    // 3. Cek Expired Asuransi (H-3)
    if (tglAsuransi === hPlus3Str && !statusLog.includes('ASURANSI:' + todayStr)) {
      kirimPaketNotifikasi('ASURANSI', data);
      statusLog += (statusLog ? ' | ' : '') + 'ASURANSI:' + todayStr;
    }

    sheet.getRange(i + 1, 16).setValue(statusLog);
  }
}

function kirimPaketNotifikasi(tipe, d) {
  let subjekCust = '';
  let pesanCust = '';
  let subjekStaff = '';
  let pesanStaff = '';

  if (tipe === 'ULTAH') {
    subjekCust = `Selamat Ulang Tahun Bapak/Ibu ${d.namaPelanggan}! - Agung Toyota Gianyar`;
    pesanCust = `Selamat Ulang Tahun Bapak/Ibu *${d.namaPelanggan}*! 🎂🎉\n\nKami keluarga besar *Agung Toyota Gianyar* mengucapkan selamat ulang tahun. Terima kasih telah mempercayakan kendaraan *${d.tipeMobil}* (*${d.noPolisi}*) kepada kami.\n• Sales: ${d.namaSales} (${d.waSales})\n• SA: ${d.namaSA} (${d.waSA})\n\n*Agung Toyota Gianyar* — _Design by MDP_`;
    subjekStaff = `[Reminder Ulang Tahun Pelanggan] ${d.namaPelanggan} (${d.noPolisi})`;
    pesanStaff = `🔔 *[REMINDER ULANG TAHUN PELANGGAN - AGUNG TOYOTA GIANYAR]*\nPelanggan *${d.namaPelanggan}* (${d.tipeMobil} - ${d.noPolisi}) berulang tahun hari ini.\nKlik untuk ucapkan selamat: https://wa.me/${d.waPelanggan}`;
  } else if (tipe === 'SAMSAT') {
    subjekCust = `Pengingat H-3 Samsat/STNK ${d.tipeMobil} (${d.noPolisi}) - Agung Toyota Gianyar`;
    pesanCust = `Halo Bapak/Ibu *${d.namaPelanggan}*,\nMasa berlaku Samsat/STNK kendaraan *${d.tipeMobil}* (*${d.noPolisi}*) akan jatuh tempo 3 hari lagi pada *${d.tglSamsat}*.\nHubungi SA (${d.namaSA} - ${d.waSA}) atau Sales (${d.namaSales} - ${d.waSales}) bila memerlukan bantuan.\n\n*Agung Toyota Gianyar* — _Design by MDP_`;
    subjekStaff = `[Reminder H-3 Samsat] ${d.namaPelanggan} (${d.noPolisi})`;
    pesanStaff = `🔔 *[REMINDER H-3 SAMSAT - AGUNG TOYOTA GIANYAR]*\nSamsat kendaraan *${d.namaPelanggan}* (${d.tipeMobil} - ${d.noPolisi}) jatuh tempo pada ${d.tglSamsat}.\nFollow-up WA: https://wa.me/${d.waPelanggan}`;
  } else if (tipe === 'ASURANSI') {
    subjekCust = `Pengingat H-3 Expired Asuransi ${d.tipeMobil} (${d.noPolisi}) - Agung Toyota Gianyar`;
    pesanCust = `Halo Bapak/Ibu *${d.namaPelanggan}*,\nMasa berlaku Asuransi kendaraan *${d.tipeMobil}* (*${d.noPolisi}*) akan berakhir 3 hari lagi pada *${d.tglAsuransi}*.\nHubungi Sales (${d.namaSales} - ${d.waSales}) atau SA (${d.namaSA} - ${d.waSA}) untuk bantuan perpanjangan.\n\n*Agung Toyota Gianyar* — _Design by MDP_`;
    subjekStaff = `[Reminder H-3 Asuransi] ${d.namaPelanggan} (${d.noPolisi})`;
    pesanStaff = `🔔 *[REMINDER H-3 ASURANSI - AGUNG TOYOTA GIANYAR]*\nAsuransi mobil *${d.namaPelanggan}* (${d.tipeMobil} - ${d.noPolisi}) berakhir pada ${d.tglAsuransi}.\nFollow-up WA: https://wa.me/${d.waPelanggan}`;
  }

  // Kirim Email ke Pelanggan, Sales, dan SA
  if (d.emailPelanggan) MailApp.sendEmail(d.emailPelanggan, subjekCust, pesanCust);
  if (d.emailSales) MailApp.sendEmail(d.emailSales, subjekStaff, pesanStaff);
  if (d.emailSA && d.emailSA !== d.emailSales) MailApp.sendEmail(d.emailSA, subjekStaff, pesanStaff);
}

function formatTgl(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG_ATG.TIMEZONE, 'yyyy-MM-dd');
  }
  return String(val).trim();
}

function bersihkanNomorWA(hp) {
  let d = String(hp || '').replace(/[^0-9]/g, '');
  if (d.startsWith('08')) d = '62' + d.slice(1);
  return d;
}
