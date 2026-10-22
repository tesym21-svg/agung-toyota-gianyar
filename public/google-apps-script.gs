/**
 * ============================================================================
 * AGUNG TOYOTA GIANYAR — GOOGLE APPS SCRIPT 2-WAY DATABASE SYNC (Design by MDP)
 * ============================================================================
 * Pasang skrip ini di Google Sheets Anda:
 * 1. Buka Google Spreadsheet baru di Google Drive Anda.
 * 2. Klik menu: Ekstensi (Extensions) -> Apps Script.
 * 3. Hapus semua kode yang ada, lalu tempel SELURUH KODE ini.
 * 4. Klik "Simpan" (ikon disket).
 * 5. Klik tombol biru "Terapkan" (Deploy) di kanan atas:
 *    - Pilih "Penerapan baru" (New deployment).
 *    - Pilih jenis: "Aplikasi Web" (Web app).
 *    - Deskripsi: "Agung Toyota Gianyar API"
 *    - Jalankan sebagai (Execute as): "Saya" (Me / email Anda).
 *    - Yang memiliki akses (Who has access): "Siapa saja" (Anyone).
 * 6. Klik "Terapkan" (Deploy) & Izinkan Akses (Authorize access).
 * 7. Salin URL Aplikasi Web (berakhiran /exec) dan tempelkan ke aplikasi Agung Toyota Gianyar.
 * ============================================================================
 */

const HEADERS = [
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

function getOrCreateCustomerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Data Pelanggan');
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName('Data Pelanggan');
  }
  
  // Pastikan header ada di baris 1
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground('#dc2626').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizeDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Makassar', 'yyyy-MM-dd');
  }
  const str = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmy) {
    return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
  }
  return str;
}

function cleanPhoneVal(val) {
  if (!val) return '';
  let digits = String(val).replace(/[^0-9]/g, '');
  if (digits.startsWith('08')) digits = '62' + digits.slice(1);
  else if (digits.startsWith('8')) digits = '62' + digits;
  return digits;
}

function readAllCustomersFromSheet() {
  const sheet = getOrCreateCustomerSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  if (lastRow <= 1) return [];

  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRow = allValues[0];

  const norm = (s) => String(s || '').toLowerCase().replace(/[\s_\.\-]/g, '');

  const colMap = {};
  for (let c = 0; c < headerRow.length; c++) {
    const h = norm(headerRow[c]);
    if (h) colMap[h] = c;
  }

  const getCell = (row, candidates, fallbackIdx) => {
    for (let i = 0; i < candidates.length; i++) {
      const key = norm(candidates[i]);
      if (colMap[key] !== undefined && colMap[key] < row.length) {
        return row[colMap[key]];
      }
    }
    if (fallbackIdx !== undefined && fallbackIdx < row.length) {
      return row[fallbackIdx];
    }
    return '';
  };

  const customers = [];

  for (let i = 1; i < allValues.length; i++) {
    const r = allValues[i];
    const nama = String(getCell(r, ['Nama Pelanggan', 'Nama', 'Customer', 'namaPelanggan'], 1) || '').trim();
    const nopol = String(getCell(r, ['No. Polisi', 'No Polisi', 'Plat Nomor', 'Nopol', 'Plat', 'noPolisi'], 5) || '').trim();
    const wa = cleanPhoneVal(getCell(r, ['No. WA Pelanggan', 'WA Pelanggan', 'No WA', 'WA', 'Telepon', 'waPelanggan'], 2));

    // Selama ada Nama ATAU Plat Nomor ATAU No WA, anggap baris valid
    if (!nama && !nopol && !wa) continue;

    customers.push({
      id: Number(getCell(r, ['No', 'ID', 'Nomor'], 0)) || i,
      namaPelanggan: nama || 'Pelanggan Tanpa Nama',
      waPelanggan: wa,
      emailPelanggan: String(getCell(r, ['Email Pelanggan', 'Email', 'emailPelanggan'], 3) || '').trim(),
      tipeMobil: String(getCell(r, ['Tipe Mobil', 'Mobil', 'Unit', 'tipeMobil'], 4) || '').trim(),
      noPolisi: nopol,
      tglUlangTahun: normalizeDateVal(getCell(r, ['Tgl Ulang Tahun', 'Tanggal Ulang Tahun', 'Ulang Tahun', 'Ultah', 'tglUlangTahun'], 6)),
      tglSamsat: normalizeDateVal(getCell(r, ['Tgl Habis Samsat', 'Tanggal Samsat', 'Tgl Samsat', 'Samsat', 'tglSamsat'], 7)),
      tglAsuransi: normalizeDateVal(getCell(r, ['Tgl Expired Asuransi', 'Tgl Asuransi', 'Asuransi', 'tglAsuransi'], 8)),
      tglService: normalizeDateVal(getCell(r, ['Tgl Next Service', 'Next Service', 'Service', 'Tgl Service', 'tglService'], 9)),
      namaSales: String(getCell(r, ['Nama Sales', 'Sales', 'namaSales'], 10) || '').trim(),
      waSales: cleanPhoneVal(getCell(r, ['No. WA Sales', 'WA Sales', 'HP Sales', 'waSales'], 11)),
      emailSales: String(getCell(r, ['Email Sales', 'emailSales'], 12) || '').trim(),
      namaSA: String(getCell(r, ['Nama SA', 'SA', 'Service Advisor', 'namaSA'], 13) || '').trim(),
      waSA: cleanPhoneVal(getCell(r, ['No. WA SA', 'WA SA', 'HP SA', 'waSA'], 14)),
      emailSA: String(getCell(r, ['Email SA', 'emailSA'], 15) || '').trim(),
      statusLog: String(getCell(r, ['Status Log Pengiriman', 'Status Log', 'Status', 'statusLog'], 16) || '').trim()
    });
  }
  return customers;
}

function syncAllCustomersToSheet(customers) {
  const sheet = getOrCreateCustomerSheet();

  // PROTEKSI KEAMANAN TINGKAT TINGGI:
  // Jangan pernah menghapus baris spreadsheet jika array data kosong atau tidak valid!
  if (!Array.isArray(customers) || customers.length === 0) {
    return { success: false, error: 'Daftar pelanggan kosong. Penghapusan baris dicegah demi keselamatan data.' };
  }

  // Hapus semua baris data lama (pertahankan baris 1 header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  const rows = customers.map((c, idx) => [
    Number(c.id) || (idx + 1),
    String(c.namaPelanggan || ''),
    String(c.waPelanggan || ''),
    String(c.emailPelanggan || ''),
    String(c.tipeMobil || ''),
    String(c.noPolisi || ''),
    String(c.tglUlangTahun || ''),
    String(c.tglSamsat || ''),
    String(c.tglAsuransi || ''),
    String(c.tglService || ''),
    String(c.namaSales || ''),
    String(c.waSales || ''),
    String(c.emailSales || ''),
    String(c.namaSA || ''),
    String(c.waSA || ''),
    String(c.emailSA || ''),
    String(c.statusLog || '')
  ]);

  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  return { success: true, count: rows.length };
}

function upsertCustomerToSheet(cust) {
  const sheet = getOrCreateCustomerSheet();
  const customers = readAllCustomersFromSheet();
  const targetId = Number(cust.id);
  const targetNopol = String(cust.noPolisi || '').toUpperCase().trim();

  let existingIndex = -1;
  if (targetId) {
    existingIndex = customers.findIndex(c => Number(c.id) === targetId);
  }
  if (existingIndex === -1 && targetNopol) {
    existingIndex = customers.findIndex(c => String(c.noPolisi || '').toUpperCase().trim() === targetNopol);
  }

  if (existingIndex !== -1) {
    customers[existingIndex] = { ...customers[existingIndex], ...cust };
  } else {
    const nextId = customers.reduce((max, c) => Math.max(max, Number(c.id) || 0), 0) + 1;
    customers.push({ ...cust, id: cust.id ? Number(cust.id) : nextId });
  }

  syncAllCustomersToSheet(customers);
  return { success: true, count: customers.length };
}

function deleteCustomerFromSheet(id, nopol) {
  const sheet = getOrCreateCustomerSheet();
  let customers = readAllCustomersFromSheet();
  const numId = Number(id);
  const cleanNop = String(nopol || '').toUpperCase().trim();

  customers = customers.filter(c => {
    if (numId && Number(c.id) === numId) return false;
    if (cleanNop && String(c.noPolisi || '').toUpperCase().trim() === cleanNop) return false;
    return true;
  });

  syncAllCustomersToSheet(customers);
  return { success: true, count: customers.length };
}

// ============================================================================
// WEB APP API (doGet & doPost)
// ============================================================================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'get_customers';
  
  if (action === 'ping') {
    return createJsonResponse({ success: true, message: 'Agung Toyota Gianyar Sheet API Ready', timestamp: new Date().toISOString() });
  }

  const customers = readAllCustomersFromSheet();
  return createJsonResponse({
    success: true,
    action: 'get_customers',
    count: customers.length,
    data: customers
  });
}

function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }

    const action = payload.action || 'sync_all';

    if (action === 'sync_all' || action === 'save_all') {
      const items = payload.items || payload.customers || [];
      const res = syncAllCustomersToSheet(items);
      return createJsonResponse(res);
    }

    if (action === 'upsert_customer' || action === 'add_customer' || action === 'save_customer') {
      const item = payload.customer || payload.item || payload;
      const res = upsertCustomerToSheet(item);
      return createJsonResponse(res);
    }

    if (action === 'delete_customer' || action === 'remove_customer') {
      const id = payload.id;
      const nopol = payload.noPolisi || payload.nopol;
      const res = deleteCustomerFromSheet(id, nopol);
      return createJsonResponse(res);
    }

    if (action === 'get_customers') {
      const customers = readAllCustomersFromSheet();
      return createJsonResponse({ success: true, count: customers.length, data: customers });
    }

    return createJsonResponse({ success: false, error: 'Aksi tidak dikenali: ' + action });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
