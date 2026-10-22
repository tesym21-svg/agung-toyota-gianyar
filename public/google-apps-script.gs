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
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const customers = [];

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const nama = String(r[1] || '').trim();
    if (!nama) continue;

    customers.push({
      id: Number(r[0]) || (i + 1),
      namaPelanggan: nama,
      waPelanggan: cleanPhoneVal(r[2]),
      emailPelanggan: String(r[3] || '').trim(),
      tipeMobil: String(r[4] || '').trim(),
      noPolisi: String(r[5] || '').trim(),
      tglUlangTahun: normalizeDateVal(r[6]),
      tglSamsat: normalizeDateVal(r[7]),
      tglAsuransi: normalizeDateVal(r[8]),
      tglService: normalizeDateVal(r[9]),
      namaSales: String(r[10] || '').trim(),
      waSales: cleanPhoneVal(r[11]),
      emailSales: String(r[12] || '').trim(),
      namaSA: String(r[13] || '').trim(),
      waSA: cleanPhoneVal(r[14]),
      emailSA: String(r[15] || '').trim(),
      statusLog: String(r[16] || '').trim()
    });
  }
  return customers;
}

function syncAllCustomersToSheet(customers) {
  const sheet = getOrCreateCustomerSheet();
  
  // Hapus semua baris data lama (pertahankan baris 1 header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  if (!Array.isArray(customers) || customers.length === 0) {
    return { success: true, count: 0 };
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
