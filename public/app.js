// ============================================================================
// AGUNG TOYOTA GIANYAR — SMART REMINDER CLIENT (Design by MDP)
// ============================================================================

let appState = {
  customers: [],
  config: {},
  stats: {},
  waState: {},
  logs: [],
  selectedTypeFilter: 'ALL',
  searchQuery: '',
  activePreviewTarget: null, // { customerId, reminderType }
  pagination: {
    page: 1,
    perPage: 15
  },
  activeReportPeriod: 'daily',
  selectedReportDate: '',
  selectedReportMonth: '',
  reportData: null
};

// 1. Mandatory Light-Dismiss Fallback for <dialog closedby="any">
function setupDialogLightDismissFallbacks() {
  const dialogs = document.querySelectorAll('dialog.app-dialog');
  dialogs.forEach((dialog) => {
    if (!('closedBy' in HTMLDialogElement.prototype)) {
      dialog.addEventListener('click', (event) => {
        if (event.target !== dialog) return;
        const rect = dialog.getBoundingClientRect();
        const isDialogContent = (
          rect.top <= event.clientY &&
          event.clientY <= rect.top + rect.height &&
          rect.left <= event.clientX &&
          event.clientX <= rect.left + rect.width
        );
        if (isDialogContent) return;
        dialog.close();
      });
    }
  });

  document.querySelectorAll('[data-close-dialog]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-close-dialog');
      const dlg = document.getElementById(id);
      if (dlg) dlg.close();
    });
  });
}

function showToast(message) {
  const container = document.getElementById('toastContainer');
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = message;
  container.appendChild(el);
  setTimeout(() => {
    el.remove();
  }, 4200);
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDateIndoShort(yyyyMmDd) {
  if (!yyyyMmDd || !/^\d{4}-\d{2}-\d{2}$/.test(yyyyMmDd)) return yyyyMmDd || '-';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const [y, m, d] = yyyyMmDd.split('-').map(Number);
  return `${d} ${months[m - 1]} ${y}`;
}

// Live WITA Clock (Asia/Makassar UTC+8)
function updateLiveWitaClock() {
  const now = new Date();
  const timeFmt = new Intl.DateTimeFormat('id-ID', {
    timeZone: 'Asia/Makassar',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  }).format(now).replace(/\./g, ':');

  const dateFmt = new Intl.DateTimeFormat('id-ID', {
    timeZone: 'Asia/Makassar',
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  }).format(now);

  const clockEl = document.getElementById('liveWitaClock');
  const dateEl = document.getElementById('liveWitaDate');
  if (clockEl) clockEl.textContent = `${timeFmt} WITA`;
  if (dateEl) dateEl.textContent = dateFmt;
}

// Fetch Dashboard Data
async function fetchDashboard(populateForms = false) {
  try {
    const res = await fetch('/api/dashboard');
    const data = await res.json();
    appState.customers = data.customers || [];
    appState.config = data.config || {};
    appState.stats = data.stats || {};
    appState.waState = data.waState || {};
    appState.logs = data.logs || [];
    appState.deliveryIssuesToday = data.deliveryIssuesToday || [];
    appState.autoPilotStatusText = data.autoPilotStatusText || '';

    // LocalStorage Caching & Recovery
    const localUrl = localStorage.getItem('toyota_gsheet_url');
    if (localUrl && !appState.config.googleSheetApiUrl) {
      appState.config.googleSheetApiUrl = localUrl;
      fetch('/api/google-sheets/save-url', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url: localUrl })
      }).catch(() => {});
    }

    if (appState.customers.length > 0) {
      try {
        localStorage.setItem('toyota_customers_cache', JSON.stringify(appState.customers));
      } catch (_) {}
    } else {
      // If server returned 0 customers (e.g. cold start on Vercel), recover from cache
      try {
        const cached = localStorage.getItem('toyota_customers_cache');
        if (cached) {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed) && parsed.length > 0) {
            appState.customers = parsed;
            if (data.stats) data.stats.totalCustomers = parsed.length;
            fetch('/api/excel/upload', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ mode: 'replace', customers: parsed })
            }).catch(() => {});
          }
        }
      } catch (_) {}
    }

    renderHeaderAndStats(data);
    renderTodayReminders();
    renderCustomerTable();
    renderWaAndSmtpStatus();
    renderLogs();
    fetchReportAnalytics().catch(() => {});

    if (populateForms) {
      populateConfigForms(appState.config);
      const urlInput = document.getElementById('inputGoogleSheetApiUrl');
      if (urlInput && !urlInput.value && localUrl) {
        urlInput.value = localUrl;
      }
    }
  } catch (err) {
    console.error('Gagal memuat dashboard:', err);
  }
}


function renderHeaderAndStats(data) {
  const safeSetText = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.textContent = text ?? '0';
  };

  safeSetText('statTotalCustomers', data.stats?.totalCustomers);
  safeSetText('statBirthday', data.stats?.dueBirthdayToday);
  safeSetText('statSamsat', data.stats?.dueSamsatH3);
  safeSetText('statAsuransi', data.stats?.dueAsuransiH3);
  safeSetText('statService', data.stats?.dueServiceH5);
  safeSetText('badgePendingToday', data.stats?.totalPendingToday);
  safeSetText('summaryCountCustomers', data.stats?.totalCustomers ?? (appState.customers?.length || 0));

  // Sync Sidebar Drawer Badges
  safeSetText('sidebarBadgeCustomers', data.stats?.totalCustomers ?? (appState.customers?.length || 0));
  safeSetText('sidebarBadgePending', data.stats?.totalPendingToday);

  const sbWa = document.getElementById('sidebarBadgeWa');
  if (sbWa) {
    const isConn = data.waState?.status === 'CONNECTED';
    sbWa.textContent = isConn ? 'ONLINE' : 'SCAN QR';
    sbWa.className = isConn ? 'nav-badge badge-emerald' : 'nav-badge badge-amber';
  }

  const isGsheetConnected = Boolean(data.config?.googleSheetApiUrl || localStorage.getItem('toyota_gsheet_url'));
  const sbGsheet = document.getElementById('sidebarBadgeGsheet');
  if (sbGsheet) {
    sbGsheet.textContent = isGsheetConnected ? '2-Arah ✓' : 'Excel';
    sbGsheet.className = isGsheetConnected ? 'nav-badge badge-sky' : 'nav-badge badge-amber';
  }

  const gsheetPill = document.getElementById('gsheetStatusPill');
  if (gsheetPill) {
    if (isGsheetConnected) {
      gsheetPill.className = 'gsheet-pill connected';
      gsheetPill.textContent = '🟢 Terhubung ke Google Spreadsheet (2-Arah)';
    } else {
      gsheetPill.className = 'gsheet-pill disconnected';
      gsheetPill.textContent = '⚪ Belum Terhubung (Penyimpanan Lokal / Excel)';
    }
  }

  const toolbarBadge = document.getElementById('toolbarGSheetBadge');
  if (toolbarBadge) {
    toolbarBadge.className = isGsheetConnected ? 'toolbar-mini-badge connected' : 'toolbar-mini-badge disconnected';
    toolbarBadge.textContent = isGsheetConnected ? '🟢 Google Sheets' : '⚪ Excel';
  }
}

function renderTodayReminders() {
  const container = document.getElementById('todayRemindersContainer');
  const issuesContainer = document.getElementById('alertDeliveryIssuesContainer');

  // Render Notifikasi & Reason Kendala Pengiriman Hari Ini
  const issues = appState.deliveryIssuesToday || [];
  if (issuesContainer) {
    if (issues.length === 0) {
      issuesContainer.classList.add('hidden');
      issuesContainer.innerHTML = '';
    } else {
      issuesContainer.classList.remove('hidden');
      issuesContainer.innerHTML = `
        <div style="background: #fffbeb; border: 1px solid #fde68a; border-radius: 12px; padding: 14px 18px;">
          <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <div style="display: flex; align-items: center; gap: 10px;">
              <span style="font-size: 26px;">⚠️</span>
              <div>
                <h4 style="margin: 0; color: #92400e; font-size: 14.5px; font-weight: 800;">
                  Notifikasi Pengiriman Hari Ini: ${issues.length} Pesan Tidak Terkirim / Terkendala
                </h4>
                <p style="margin: 3px 0 0; font-size: 12px; color: #b45309;">
                  Sesuai instruksi Anda: Pesan yang gagal/terkendala <strong>TIDAK dikirim ulang secara otomatis</strong> agar tidak terjadi pengulangan pesan ke inbox/pelanggan.
                </p>
              </div>
            </div>
          </div>
          <div style="margin-top: 12px; display: flex; flex-direction: column; gap: 8px;">
            ${issues.map(iss => `
              <div style="background: #ffffff; border: 1px solid #fed7aa; border-radius: 8px; padding: 10px 14px; font-size: 12.5px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
                <div>
                  <strong>${escapeHtml(iss.namaPelanggan)}</strong> (${escapeHtml(iss.tipeMobil)} • <span style="font-weight:700;">${escapeHtml(iss.noPolisi)}</span>) — <span style="font-weight: 700; color: #c2410c;">${escapeHtml(iss.categoryTitle || iss.reminderType)}</span>
                  <div style="color: #b91c1c; font-size: 12px; margin-top: 3px; font-weight: 600;">
                    ⚠️ Alasan (Reason): ${escapeHtml(iss.reason || 'Koneksi WA / Email tidak tersedia')}
                  </div>
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                  <span style="font-size: 11.5px; color: var(--slate-500); font-weight: 600;">
                    🕒 ${escapeHtml(iss.timestampWita)} • <em>Tidak diulang</em>
                  </span>
                  <button type="button" class="btn btn-outline btn-sm" data-action="send-single" data-id="${iss.customerId}" data-type="${iss.reminderType}" style="padding: 3px 8px; font-size: 11px;">
                    🚀 Kirim Manual
                  </button>
                </div>
              </div>
            `).join('')}
          </div>
        </div>
      `;
    }
  }

  const items = [];
  const tomorrowItems = [];

  for (const cust of appState.customers) {
    for (const rem of cust.activeReminders || []) {
      if (appState.selectedTypeFilter === 'ALL' || appState.selectedTypeFilter === rem.type) {
        items.push({ cust, rem, isTomorrow: false });
      }
    }
    for (const rem of cust.tomorrowReminders || []) {
      if (appState.selectedTypeFilter === 'ALL' || appState.selectedTypeFilter === rem.type) {
        tomorrowItems.push({ cust, rem, isTomorrow: true });
      }
    }
  }

  if (items.length === 0 && tomorrowItems.length === 0) {
    container.innerHTML = `
      <div class="settings-card" style="grid-column: 1 / -1; text-align: center; padding: 36px;">
        <div style="font-size: 38px; margin-bottom: 8px;">✅</div>
        <h3 style="margin: 0 0 6px;">Tidak Ada Pengingat pada Filter Ini untuk Hari Ini</h3>
        <p style="margin: 0; color: var(--slate-600); font-size: 13.5px;">
          Semua data pelanggan dari Excel telah diperiksa. Anda dapat melihat seluruh jadwal di tab <strong>Data Pelanggan &amp; Excel</strong>.
        </p>
      </div>
    `;
    return;
  }

  const renderCard = ({ cust, rem, isTomorrow }) => {
    const icon = rem.type === 'ULTAH'
      ? '🎂'
      : rem.type === 'SAMSAT'
        ? '🚗'
        : rem.type === 'ASURANSI'
          ? '🛡️'
          : '🔧';
    const dateLabel = rem.type === 'ULTAH'
      ? `Tgl Lahir: ${formatDateIndoShort(cust.tglUlangTahun)}`
      : rem.type === 'SAMSAT'
        ? `Jatuh Tempo Samsat: ${formatDateIndoShort(cust.tglSamsat)} (H-3)`
        : rem.type === 'ASURANSI'
          ? `Expired Asuransi: ${formatDateIndoShort(cust.tglAsuransi)} (H-3)`
          : `Jadwal Next Service: ${formatDateIndoShort(cust.tglService)} (H-5)`;

    let statusBadgeHtml = '';
    if (isTomorrow) {
      statusBadgeHtml = `<span class="sent-pill pending" style="background:#dbeafe;color:#1e40af;border:1px solid #93c5fd;">
           🗓️ Terjadwal Otomatis Besok (${formatDateIndoShort(rem.scheduledDate)} • 10:00–14:00 WITA)
         </span>`;
    } else if (rem.alreadySentToday) {
      const last = rem.lastLog;
      if (!last || last.status === 'SUCCESS') {
        statusBadgeHtml = `<span class="sent-pill done" style="background:#ecfdf5;color:#047857;border:1px solid #a7f3d0;">
             ✓ Terkirim Sukses (Selesai)
           </span>`;
      } else if (last.status === 'PARTIAL') {
        statusBadgeHtml = `
          <div style="display:flex;flex-direction:column;align-items:flex-end;gap:3px;">
            <span class="sent-pill" style="background:#fffbeb;color:#b45309;border:1px solid #fde68a;">
              ⚠️ Sebagian Terkendala
            </span>
            <span style="font-size:11px;color:#c2410c;font-weight:600;max-width:240px;text-align:right;">
              Alasan: ${escapeHtml(last.reason)} (Tidak diulang)
            </span>
          </div>`;
      } else {
        statusBadgeHtml = `
          <div style="display:flex;flex-direction:column;align-items:flex-end;gap:3px;">
            <span class="sent-pill" style="background:#fef2f2;color:#b91c1c;border:1px solid #fecaca;">
              ❌ Tidak Terkirim
            </span>
            <span style="font-size:11px;color:#b91c1c;font-weight:600;max-width:240px;text-align:right;">
              Alasan: ${escapeHtml(last.reason)} (Tidak diulang)
            </span>
          </div>`;
      }
    } else {
      statusBadgeHtml = `<span class="sent-pill pending">
           ⏳ Menunggu Kirim (10:00–14:00 WITA)
         </span>`;
    }

    return `
      <article class="rem-card type-${escapeHtml(rem.type)}">
        <div>
          <div class="rem-card-top">
            <span class="rem-badge ${escapeHtml(rem.type)}">${icon} ${escapeHtml(rem.label)}</span>
            ${statusBadgeHtml}
          </div>
          <h3 class="rem-cust-name">${escapeHtml(cust.namaPelanggan)}</h3>
          <div class="rem-vehicle">🚘 ${escapeHtml(cust.tipeMobil)} • <span class="plate-tag">${escapeHtml(cust.noPolisi)}</span></div>
          <div style="font-size: 12.5px; color: var(--toyota-red-dark); font-weight: 700; margin-top: 6px;">
            📅 ${escapeHtml(dateLabel)}
          </div>

          <div class="rem-recipients-box mt-12">
            <div class="recip-row">
              <span class="recip-label">👤 Pelanggan (WA &amp; Email):</span>
              <span class="recip-val">${escapeHtml(cust.waPelanggan || '-')} • ${escapeHtml(cust.emailPelanggan || '-')}</span>
            </div>
            <div class="recip-row">
              <span class="recip-label">👨‍💼 Sales PIC (${escapeHtml(cust.namaSales || '-')}):</span>
              <span class="recip-val">${escapeHtml(cust.waSales || '-')} • ${escapeHtml(cust.emailSales || '-')}</span>
            </div>
            <div class="recip-row">
              <span class="recip-label">🔧 SA PIC (${escapeHtml(cust.namaSA || '-')}):</span>
              <span class="recip-val">${escapeHtml(cust.waSA || '-')} • ${escapeHtml(cust.emailSA || '-')}</span>
            </div>
          </div>
        </div>

        <div class="rem-card-actions">
          <button type="button" class="btn btn-sm btn-outline" data-action="preview" data-id="${cust.id}" data-type="${rem.type}">
            👁️ Preview 3 Pesan (Pelanggan, Sales, SA)
          </button>
          <button type="button" class="btn btn-sm btn-primary" data-action="send-single" data-id="${cust.id}" data-type="${rem.type}">
            🚀 ${isTomorrow ? 'Tes Kirim Sekarang' : (rem.alreadySentToday ? 'Kirim Ulang WA & Gmail' : 'Kirim Sekarang')}
          </button>
        </div>
      </article>
    `;
  };

  let htmlOut = '';
  if (items.length > 0) {
    htmlOut += items.map(renderCard).join('');
  }
  if (tomorrowItems.length > 0) {
    const tmrDateFormatted = formatDateIndoShort(tomorrowItems[0].rem.scheduledDate);
    htmlOut += `
      <div style="grid-column: 1 / -1; background: #eff6ff; border: 1.5px solid #93c5fd; border-radius: 12px; padding: 14px 18px; margin-top: 4px;">
        <div style="font-weight: 800; color: #1e3a8a; font-size: 14.5px;">
          🗓️ Daftar Pengingat yang Sudah Terjadwal Otomatis untuk Besok (${escapeHtml(tmrDateFormatted)} — Pukul ${escapeHtml(appState.config.waWindowStart || '10:00')}–${escapeHtml(appState.config.waWindowEnd || '14:00')} WITA)
        </div>
        <div style="font-size: 12.5px; color: #1e40af; margin-top: 3px;">
          ✅ Setting tanggal Anda untuk besok sudah terbaca oleh sistem! Besok saat memasuki jam <strong>${escapeHtml(appState.config.waWindowStart || '10:00')} WITA</strong>, ke-${tomorrowItems.length} pengingat di bawah ini akan otomatis dikirim ke WhatsApp &amp; Gmail tanpa perlu klik apa pun.
        </div>
      </div>
    ` + tomorrowItems.map(renderCard).join('');
  }

  container.innerHTML = htmlOut;
}

function renderCustomerTable() {
  const tbody = document.getElementById('customerTableBody');
  if (!tbody) return;
  const q = appState.searchQuery.trim().toLowerCase();

  const filtered = appState.customers.filter((c) => {
    if (!q) return true;
    return [
      c.namaPelanggan,
      c.waPelanggan,
      c.emailPelanggan,
      c.tipeMobil,
      c.noPolisi,
      c.namaSales,
      c.namaSA
    ].some(v => String(v || '').toLowerCase().includes(q));
  });

  const totalCount = filtered.length;
  const perPage = appState.pagination.perPage === 'ALL' ? totalCount : Number(appState.pagination.perPage || 15);
  const totalPages = Math.max(1, Math.ceil(totalCount / (perPage || 1)));

  if (appState.pagination.page > totalPages) {
    appState.pagination.page = totalPages;
  }
  const currentPage = appState.pagination.page;
  const startIndex = (currentPage - 1) * perPage;
  const endIndex = Math.min(startIndex + perPage, totalCount);
  const pageRows = (appState.pagination.perPage === 'ALL') ? filtered : filtered.slice(startIndex, endIndex);

  // Update pagination info & controls
  const infoEl = document.getElementById('paginationInfo');
  if (infoEl) {
    infoEl.innerHTML = totalCount === 0
      ? 'Tidak ada data pelanggan'
      : `Menampilkan <strong>${startIndex + 1}–${endIndex}</strong> dari <strong>${totalCount}</strong> pelanggan`;
  }

  const btnPrev = document.getElementById('btnPrevPage');
  const btnNext = document.getElementById('btnNextPage');
  if (btnPrev) btnPrev.disabled = currentPage <= 1;
  if (btnNext) btnNext.disabled = currentPage >= totalPages;

  const pagesContainer = document.getElementById('paginationPages');
  if (pagesContainer) {
    let pagesHtml = '';
    const maxVisible = 5;
    let startPage = Math.max(1, currentPage - 2);
    let endPage = Math.min(totalPages, startPage + maxVisible - 1);
    if (endPage - startPage < maxVisible - 1) {
      startPage = Math.max(1, endPage - maxVisible + 1);
    }
    for (let p = startPage; p <= endPage; p++) {
      pagesHtml += `<button type="button" class="page-btn ${p === currentPage ? 'active' : ''}" data-goto-page="${p}">${p}</button>`;
    }
    pagesContainer.innerHTML = pagesHtml;
  }

  if (pageRows.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="10" style="text-align: center; padding: 28px; color: var(--slate-600);">
          Data pelanggan tidak ditemukan. Silakan tambah pelanggan atau tarik dari Google Sheets.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = pageRows.map((c, idx) => {
    const globalIdx = startIndex + idx + 1;

    const bdayBadge = c.isBirthdayToday
      ? `<span class="mini-tag bday">🎂 HARI INI</span>`
      : '';

    const samsatBadge = c.samsatDiff === 3
      ? `<span class="mini-tag samsat">🚗 H-3</span>`
      : (c.samsatDiff !== null && c.samsatDiff >= 0 && c.samsatDiff <= 14)
        ? `<span class="mini-tag samsat">H-${c.samsatDiff}</span>`
        : '';

    const asuransiBadge = c.asuransiDiff === 3
      ? `<span class="mini-tag asuransi">🛡️ H-3</span>`
      : (c.asuransiDiff !== null && c.asuransiDiff >= 0 && c.asuransiDiff <= 14)
        ? `<span class="mini-tag asuransi">H-${c.asuransiDiff}</span>`
        : '';

    const serviceAdv = Number(appState.config.serviceAdvanceDays ?? 5);
    const serviceBadge = c.serviceDiff === serviceAdv
      ? `<span class="mini-tag service">🔧 H-${serviceAdv}</span>`
      : (c.serviceDiff !== null && c.serviceDiff >= 0 && c.serviceDiff <= 14)
        ? `<span class="mini-tag service">H-${c.serviceDiff}</span>`
        : '';

    const waLink = c.waPelanggan ? `https://wa.me/${c.waPelanggan}` : '#';

    return `
      <tr>
        <td style="text-align: center; color: var(--slate-500); font-weight: 700;">${globalIdx}</td>
        <td>
          <div class="compact-cust-cell">
            <span class="cust-name">${escapeHtml(c.namaPelanggan)}</span>
            ${c.waPelanggan ? `<a href="${waLink}" target="_blank" rel="noopener" class="mini-wa-chip" title="Chat WhatsApp Pelanggan">📱 ${escapeHtml(c.waPelanggan)}</a>` : ''}
          </div>
          ${c.emailPelanggan ? `<div style="font-size:11px;color:var(--slate-500);margin-top:1px;">✉️ ${escapeHtml(c.emailPelanggan)}</div>` : ''}
        </td>
        <td>
          <span class="compact-plate">${escapeHtml(c.noPolisi || '-')}</span>
          <span class="compact-car-model">${escapeHtml(c.tipeMobil || '-')}</span>
        </td>
        <td style="white-space: nowrap;">
          <span>${formatDateIndoShort(c.tglUlangTahun)}</span>
          ${bdayBadge}
        </td>
        <td style="white-space: nowrap;">
          <span>${formatDateIndoShort(c.tglSamsat)}</span>
          ${samsatBadge}
        </td>
        <td style="white-space: nowrap;">
          <span>${formatDateIndoShort(c.tglAsuransi)}</span>
          ${asuransiBadge}
        </td>
        <td style="white-space: nowrap;">
          <span>${formatDateIndoShort(c.tglService)}</span>
          ${serviceBadge}
        </td>
        <td>
          <div style="font-size:11.5px;font-weight:600;color:var(--slate-700);">
            ${escapeHtml(c.namaSales || '-')}
            <span style="color:var(--slate-400);font-weight:400;">(Sales)</span>
          </div>
          <div style="font-size:11px;color:var(--slate-500);">
            ${escapeHtml(c.namaSA || '-')}
            <span style="color:var(--slate-400);">(SA)</span>
          </div>
        </td>
        <td style="white-space: nowrap;">
          ${(() => {
            const activeLog = (c.activeReminders || []).map(r => r.lastLog).find(Boolean);
            if (activeLog) {
              if (activeLog.status === 'SUCCESS') {
                return `<span class="status-dot-ok" style="color:#10b981;font-weight:700;" title="Berhasil dikirim pada ${activeLog.time || ''} WITA">🟢 ✓ Terkirim</span>`;
              } else if (activeLog.status === 'PARTIAL') {
                return `<span class="status-dot-warn" style="color:#f59e0b;font-weight:700;" title="${escapeHtml(activeLog.reason || 'Terkirim sebagian')}">🟡 ⚠️ Sebagian</span>`;
              } else if (activeLog.status === 'FAILED') {
                return `<span class="status-dot-fail" style="color:#ef4444;font-weight:700;" title="${escapeHtml(activeLog.reason || 'Gagal terkirim')}">🔴 ✕ Gagal</span>`;
              }
            }
            if (c.statusLog) {
              const isGagal = c.statusLog.includes('GAGAL') || c.statusLog.includes('FAIL');
              const isParsial = c.statusLog.includes('PARSIAL') || c.statusLog.includes('PARTIAL') || c.statusLog.includes('WARN');
              const isSukses = c.statusLog.includes('OK') || c.statusLog.includes('SUCCESS');
              if (isGagal) {
                return `<span class="status-dot-fail" style="color:#ef4444;font-weight:700;" title="${escapeHtml(c.statusLog)}">🔴 ✕ Gagal</span>`;
              } else if (isParsial) {
                return `<span class="status-dot-warn" style="color:#f59e0b;font-weight:700;" title="${escapeHtml(c.statusLog)}">🟡 ⚠️ Sebagian</span>`;
              } else if (isSukses) {
                return `<span class="status-dot-ok" style="color:#10b981;font-weight:700;" title="${escapeHtml(c.statusLog)}">🟢 ✓ Terkirim</span>`;
              }
            }
            if (c.activeReminders && c.activeReminders.length > 0) {
              return '<span style="color:#d97706;font-size:11.5px;font-weight:700;">⏳ Belum Dikirim</span>';
            }
            return '<span style="color:#94a3b8;font-size:11px;">-</span>';
          })()}
        </td>
        <td style="text-align: right; white-space: nowrap;">
          <div style="display:inline-flex;gap:4px;">
            <button type="button" class="btn btn-sm btn-primary" data-action="send-single" data-id="${c.id}" data-type="${c.activeReminders?.[0]?.type || 'SERVICE'}" style="padding: 4px 8px; font-size: 11.5px;" title="Langsung Kirim WA & Email">🚀</button>
            <button type="button" class="btn btn-sm btn-outline" data-action="edit-customer" data-id="${c.id}" style="padding: 4px 8px; font-size: 11.5px;" title="Edit Data Pelanggan">✏️</button>
            <button type="button" class="btn btn-sm btn-outline" data-action="preview" data-id="${c.id}" data-type="${c.activeReminders?.[0]?.type || 'SERVICE'}" style="padding: 4px 8px; font-size: 11.5px;" title="Preview Pesan">👁️</button>
            <button type="button" class="btn btn-sm btn-danger" data-action="delete-customer" data-id="${c.id}" style="padding: 4px 8px; font-size: 11.5px;" title="Hapus">🗑️</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function renderWaAndSmtpStatus() {
  const waPill = document.getElementById('waConnectionPill');
  const waDesc = document.getElementById('waStatusDescription');
  const qrImg = document.getElementById('waQrImage');

  // Front-page Auto-Pilot & QR Elements
  const mainWaChip = document.getElementById('mainWaStatusChip');
  const mainGmailChip = document.getElementById('mainGmailStatusChip');
  const mainAutoStatus = document.getElementById('mainAutoPilotStatusText');
  const mainQrImg = document.getElementById('mainWaQrImage');
  const mainQrLoading = document.getElementById('mainWaQrLoading');
  const mainNotConn = document.getElementById('mainQrNotConnected');
  const mainConnBox = document.getElementById('mainQrConnectedBox');
  const mainConnNum = document.getElementById('mainWaConnectedNumber');
  const mainGmailAcc = document.getElementById('mainGmailConnectedAccount');

  if (mainAutoStatus && appState.autoPilotStatusText) {
    mainAutoStatus.innerHTML = `<strong>Status Mesin Otomatis:</strong> ${escapeHtml(appState.autoPilotStatusText)}`;
  }

  if (appState.waState.status === 'CONNECTED') {
    if (waPill) {
      waPill.textContent = `✓ Terhubung (${appState.waState.connectedNumber || 'WhatsApp Aktif'})`;
      waPill.className = 'status-pill connected';
    }
    if (waDesc) {
      waDesc.innerHTML = `✅ <strong>WhatsApp Kantor Terhubung!</strong> (${escapeHtml(appState.waState.connectedName || '')}). Siap mengirim pesan otomatis pada pukul <strong>${escapeHtml(appState.config.waWindowStart || '10:00')} – ${escapeHtml(appState.config.waWindowEnd || '14:00')} WITA</strong>.`;
    }
    if (qrImg) qrImg.classList.add('hidden');

    if (mainWaChip) {
      mainWaChip.textContent = `📱 WA Terhubung: ${appState.waState.connectedNumber || 'Aktif'}`;
      mainWaChip.style.background = '#16a34a';
    }
    if (mainNotConn) mainNotConn.classList.add('hidden');
    if (mainConnBox) mainConnBox.classList.remove('hidden');
    if (mainConnNum) mainConnNum.textContent = `📱 WA: ${appState.waState.connectedName || ''} (${appState.waState.connectedNumber || ''})`;
  } else if (appState.waState.status === 'QR_READY' && appState.waState.qrDataUrl) {
    if (waPill) {
      waPill.textContent = '📲 Menunggu Scan QR Code';
      waPill.className = 'status-pill';
    }
    if (waDesc) {
      waDesc.innerHTML = 'Silakan buka WhatsApp di HP Kantor &rarr; <strong>Perangkat Tertaut</strong> &rarr; <strong>Tautkan Perangkat</strong> lalu scan QR di bawah ini:';
    }
    if (qrImg) {
      qrImg.src = appState.waState.qrDataUrl;
      qrImg.classList.remove('hidden');
    }

    if (mainWaChip) {
      mainWaChip.textContent = '📲 WA Belum Tertaut — Silakan Scan QR di Tab Auto-Pilot';
      mainWaChip.style.background = '#dc2626';
    }
    if (mainNotConn) mainNotConn.classList.remove('hidden');
    if (mainConnBox) mainConnBox.classList.add('hidden');
    if (mainQrImg) {
      mainQrImg.src = appState.waState.qrDataUrl;
      mainQrImg.classList.remove('hidden');
    }
    if (mainQrLoading) mainQrLoading.classList.add('hidden');
  } else if (appState.waState.status === 'CONNECTING') {
    if (waPill) {
      waPill.textContent = '⏳ Menyiapkan QR Code...';
      waPill.className = 'status-pill';
    }
    if (waDesc) {
      waDesc.textContent = 'Sedang menghubungi server WhatsApp Web, mohon tunggu beberapa detik...';
    }
    if (qrImg) qrImg.classList.add('hidden');

    if (mainWaChip) mainWaChip.textContent = '⏳ Menyiapkan QR WhatsApp...';
    if (mainNotConn) mainNotConn.classList.remove('hidden');
    if (mainConnBox) mainConnBox.classList.add('hidden');
    if (mainQrImg) mainQrImg.classList.add('hidden');
    if (mainQrLoading) mainQrLoading.classList.remove('hidden');
  } else {
    if (waPill) {
      waPill.textContent = 'Siap Scan QR';
      waPill.className = 'status-pill';
    }
    if (waDesc) {
      waDesc.innerHTML = 'Klik tombol <strong>"Tampilkan / Hubungkan QR WhatsApp"</strong> di bawah untuk menautkan WhatsApp kantor secara gratis.';
    }
    if (qrImg) qrImg.classList.add('hidden');

    if (mainWaChip) mainWaChip.textContent = '📲 Menunggu Koneksi QR';
  }

  const smtpPill = document.getElementById('smtpConnectionPill');
  if (appState.config?.smtp?.enabled && appState.config?.smtp?.user && appState.config?.smtp?.pass) {
    if (smtpPill) {
      smtpPill.textContent = `✓ Gmail Aktif (${appState.config.smtp.user})`;
      smtpPill.className = 'status-pill connected';
    }
    if (mainGmailChip) {
      mainGmailChip.textContent = `✉️ Gmail Terhubung: ${appState.config.smtp.user}`;
      mainGmailChip.style.background = '#16a34a';
    }
    if (mainGmailAcc) {
      mainGmailAcc.textContent = `✉️ Gmail: ${appState.config.smtp.user}`;
    }
  } else {
    if (smtpPill) {
      smtpPill.textContent = 'Mode Simulasi Email';
      smtpPill.className = 'status-pill';
    }
    if (mainGmailChip) {
      mainGmailChip.textContent = '✉️ Gmail Belum Diatur';
      mainGmailChip.style.background = '#64748b';
    }
    if (mainGmailAcc) {
      mainGmailAcc.textContent = '';
    }
  }
}

function setReportPeriod(period) {
  appState.activeReportPeriod = period;
  
  const secDaily = document.getElementById('reportSectionDaily');
  const secWeekly = document.getElementById('reportSectionWeekly');
  const secMonthly = document.getElementById('reportSectionMonthly');

  const btnDaily = document.getElementById('btnReportPeriodDaily');
  const btnWeekly = document.getElementById('btnReportPeriodWeekly');
  const btnMonthly = document.getElementById('btnReportPeriodMonthly');

  if (secDaily) secDaily.classList.toggle('hidden', period !== 'daily');
  if (secWeekly) secWeekly.classList.toggle('hidden', period !== 'weekly');
  if (secMonthly) secMonthly.classList.toggle('hidden', period !== 'monthly');

  if (btnDaily) btnDaily.classList.toggle('active', period === 'daily');
  if (btnWeekly) btnWeekly.classList.toggle('active', period === 'weekly');
  if (btnMonthly) btnMonthly.classList.toggle('active', period === 'monthly');
}
window.setReportPeriod = setReportPeriod;

async function fetchReportAnalytics(targetDate, targetMonth) {
  try {
    const qDate = targetDate !== undefined ? targetDate : appState.selectedReportDate;
    const qMonth = targetMonth !== undefined ? targetMonth : appState.selectedReportMonth;

    const params = new URLSearchParams();
    if (qDate) params.set('date', qDate);
    if (qMonth) params.set('month', qMonth);

    const res = await fetch(`/api/reports?${params.toString()}`);
    if (!res.ok) return;
    const data = await res.json();
    if (!data.ok) return;

    appState.reportData = data;
    if (data.curDate && !appState.selectedReportDate) {
      appState.selectedReportDate = data.curDate;
    }
    if (data.curMonth && !appState.selectedReportMonth) {
      appState.selectedReportMonth = data.curMonth;
    }

    renderReportsDashboard();
  } catch (err) {
    console.error('Gagal mengambil data laporan analitik:', err);
  }
}

function renderReportsDashboard() {
  const d = appState.reportData;
  if (!d) return;

  const safeSet = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val ?? '';
  };

  // ==================== 1. DAILY REPORT ====================
  const daily = d.daily || {};
  const dateInput = document.getElementById('inputReportDailyDate');
  if (dateInput && !dateInput.value) {
    dateInput.value = daily.date || d.curDate || '';
  }
  const dateLabel = document.getElementById('dailyReportDateLabel');
  if (dateLabel) {
    dateLabel.textContent = (daily.date === d.curDate)
      ? `🟢 Hari Ini — ${daily.dateIndo || ''}`
      : `📅 ${daily.dateIndo || daily.date}`;
  }

  safeSet('kpiDailyTotal', daily.totalTarget || 0);
  safeSet('kpiDailySent', daily.totalSent || 0);
  safeSet('kpiDailySuccessRate', daily.successRate || '100%');
  safeSet('kpiDailyFailed', daily.totalFailed || 0);
  const failRate = daily.totalTarget > 0 ? `${Math.round((daily.totalFailed / daily.totalTarget) * 100)}%` : '0%';
  safeSet('kpiDailyFailedRate', failRate);
  safeSet('kpiDailyStaff', (daily.totalSent + (daily.totalPartial || 0)) * 2);

  // Daily Category Progress Bars
  const catBarsEl = document.getElementById('dailyCategoryBars');
  if (catBarsEl) {
    const cats = [
      { key: 'ULTAH', label: '🎂 Ulang Tahun', fillClass: 'fill-bday', color: '#db2777' },
      { key: 'SAMSAT', label: '🚗 Pajak / Samsat H-3', fillClass: 'fill-samsat', color: '#d97706' },
      { key: 'ASURANSI', label: '🛡️ Expired Asuransi H-3', fillClass: 'fill-asuransi', color: '#2563eb' },
      { key: 'SERVICE', label: '🔧 Next Service H-5', fillClass: 'fill-service', color: '#0d9488' }
    ];

    catBarsEl.innerHTML = cats.map(c => {
      const data = daily.byCategory?.[c.key] || { total: 0, sent: 0, failed: 0 };
      const pct = data.total > 0 ? Math.round((data.sent / data.total) * 100) : 0;
      const failPct = 100 - pct;

      return `
        <div class="category-progress-item">
          <div class="category-label-row">
            <span style="color: ${c.color}; font-weight: 700;">${c.label}</span>
            <span style="color: var(--slate-600); font-weight: 600;">${data.sent} / ${data.total} (${pct}%)</span>
          </div>
          <div class="progress-bar-track">
            ${data.total === 0
              ? `<div class="progress-bar-fill" style="width: 100%; background: #e2e8f0;"></div>`
              : `
                <div class="progress-bar-fill ${c.fillClass}" style="width: ${pct}%;"></div>
                ${data.failed > 0 ? `<div class="progress-bar-fill fill-failed" style="width: ${failPct}%;"></div>` : ''}
              `
            }
          </div>
        </div>
      `;
    }).join('');
  }

  safeSet('dailyChartFooterSent', `${daily.totalSent || 0} Pesan`);
  safeSet('dailyChartFooterFailed', `${daily.totalFailed || 0} Pesan`);

  // Daily Reason & Message Details Table
  const tableBody = document.getElementById('dailyReportTableBody');
  const attnBadge = document.getElementById('dailyAttentionBadge');
  if (attnBadge) {
    if (daily.totalFailed > 0) {
      attnBadge.textContent = `⚠️ ${daily.totalFailed} Butuh Perhatian`;
      attnBadge.classList.remove('hidden');
    } else {
      attnBadge.classList.add('hidden');
    }
  }

  if (tableBody) {
    const logs = daily.logs || [];
    if (logs.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="5" style="text-align:center; padding: 28px; color: var(--slate-500);">
            Belum ada riwayat pengiriman pesan yang tercatat pada tanggal <strong>${escapeHtml(daily.dateIndo || daily.date)}</strong>.
          </td>
        </tr>
      `;
    } else {
      tableBody.innerHTML = logs.map(l => {
        let statusBadge = '<span class="status-pill-ok">✓ Terkirim</span>';
        let reasonClass = 'text-green font-semibold';
        if (l.status === 'FAILED') {
          statusBadge = '<span class="status-pill-fail">✕ Gagal</span>';
          reasonClass = 'reason-text-highlight';
        } else if (l.status === 'PARTIAL') {
          statusBadge = '<span class="status-pill-partial">⚠️ Parsial</span>';
          reasonClass = 'reason-text-partial';
        }

        const catIcons = { ULTAH: '🎂 Ultah', SAMSAT: '🚗 Samsat', ASURANSI: '🛡️ Asuransi', SERVICE: '🔧 Service' };
        const catLabel = catIcons[l.reminderType] || l.categoryTitle || 'Pengingat';

        return `
          <tr>
            <td style="white-space: nowrap; color: var(--slate-500); font-weight: 600;">
              ${escapeHtml(l.timeWita || l.timestampWita?.split(' ')[1] || '-')} WITA
            </td>
            <td>
              <strong style="color: var(--obsidian); display: block;">${escapeHtml(l.namaPelanggan)}</strong>
              <span style="font-size: 11px; color: var(--slate-500);">${escapeHtml(l.noPolisi || '-')} • ${escapeHtml(l.tipeMobil || '-')}</span>
            </td>
            <td style="white-space: nowrap; font-weight: 700;">
              ${escapeHtml(catLabel)}
            </td>
            <td style="text-align: center; white-space: nowrap;">
              ${statusBadge}
            </td>
            <td>
              <div class="${reasonClass}">${escapeHtml(l.failureReason || 'Berhasil terkirim')}</div>
            </td>
          </tr>
        `;
      }).join('');
    }
  }

  // ==================== 2. WEEKLY REPORT ====================
  const weekly = d.weekly || {};
  safeSet('weeklyTotalSummaryBadge', `Total Minggu Ini: ${weekly.totalAll || 0} Pesan`);
  safeSet('weeklyLegendSent', weekly.totalSent || 0);
  safeSet('weeklyLegendFailed', weekly.totalFailed || 0);

  const weeklyContainer = document.getElementById('weeklyBarsContainer');
  if (weeklyContainer) {
    const days = weekly.days || [];
    const maxVal = Math.max(...days.map(day => day.total || 0), 1);

    weeklyContainer.innerHTML = days.map(day => {
      const heightPct = Math.max(12, Math.round((day.total / maxVal) * 90));
      const failPct = day.total > 0 ? Math.round((day.failed / day.total) * 100) : 0;
      const successPct = 100 - failPct;

      return `
        <div class="weekly-day-col">
          <span style="font-size: 11px; font-weight: 800; color: ${day.failed > 0 ? '#b91c1c' : '#475569'};">
            ${day.failed > 0 ? `${day.sent}/${day.failed}!` : day.sent}
          </span>
          <div class="weekly-bar-stacked" style="height: ${heightPct}%;">
            ${day.failed > 0 ? `<div class="weekly-bar-segment-fail" style="height: ${failPct}%;"></div>` : ''}
            <div class="weekly-bar-segment-success" style="height: ${successPct}%;"></div>
          </div>
          <span class="weekly-day-label ${day.date === d.curDate ? 'today' : ''}">
            ${escapeHtml(day.dayLabel)}
          </span>
        </div>
      `;
    }).join('');
  }

  const reasonsContainer = document.getElementById('weeklyTopReasonsContainer');
  if (reasonsContainer) {
    const reasons = weekly.topFailureReasons || [];
    if (reasons.length === 0) {
      reasonsContainer.innerHTML = `
        <div style="grid-column: 1 / -1; text-align: center; padding: 20px; color: var(--slate-500); font-size: 13px;">
          🎉 <strong>Luar Biasa!</strong> Tidak ada kegagalan pengiriman yang terdeteksi dalam 7 hari terakhir.
        </div>
      `;
    } else {
      reasonsContainer.innerHTML = reasons.map((r, idx) => `
        <div class="reason-card">
          <div class="reason-card-title">
            <span>${idx + 1}. ${escapeHtml(r.reason)}</span>
            <span style="background: #fee2e2; padding: 2px 7px; border-radius: 6px; font-size: 11px;">${r.count} Kasus</span>
          </div>
          <p class="reason-card-desc">Periksa kelengkapan nomor WhatsApp dan email pelanggan pada data spreadsheet/Excel.</p>
        </div>
      `).join('');
    }
  }

  // ==================== 3. MONTHLY REPORT ====================
  const monthly = d.monthly || {};
  safeSet('monthlyReportTitle', `📈 Laporan Akumulasi Bulanan (${monthly.monthLabel || ''})`);
  safeSet('kpiMonthlyTotal', monthly.totalAll || 0);
  safeSet('kpiMonthlySent', monthly.totalSent || 0);
  safeSet('kpiMonthlySentRate', `Tingkat keberhasilan: ${monthly.successRate || '100%'}`);
  safeSet('kpiMonthlyFailed', monthly.totalFailed || 0);
  const mFailRate = monthly.totalAll > 0 ? `${Math.round((monthly.totalFailed / monthly.totalAll) * 100)}%` : '0%';
  safeSet('kpiMonthlyFailedRate', `${mFailRate} dari total data`);
  safeSet('kpiMonthlyServiceUnits', monthly.estimatedServiceUnits || 0);

  // Calendar Month Picker (Perpetual Multi-Year Support via Calendar)
  const monthPicker = document.getElementById('inputReportMonthPicker');
  const currentYearNum = (d.curDate ? parseInt(d.curDate.slice(0, 4), 10) : new Date().getFullYear()) || 2026;
  const activeMonthStr = monthly.month || d.curMonth || `${currentYearNum}-10`;

  if (monthPicker) {
    monthPicker.value = activeMonthStr;
  }

  const monthlyCatBars = document.getElementById('monthlyCategoryBars');
  if (monthlyCatBars) {
    const mTotal = monthly.totalAll || 1;
    const cats = [
      { key: 'ULTAH', label: '🎂 Ulang Tahun (Hari H)', fillClass: 'fill-bday', color: '#db2777' },
      { key: 'SAMSAT', label: '🚗 Pajak / Samsat Kendaraan (H-3)', fillClass: 'fill-samsat', color: '#d97706' },
      { key: 'ASURANSI', label: '🛡️ Polis Asuransi Mobil (H-3)', fillClass: 'fill-asuransi', color: '#2563eb' },
      { key: 'SERVICE', label: '🔧 Next Service Berkala (H-5)', fillClass: 'fill-service', color: '#0d9488' }
    ];

    monthlyCatBars.innerHTML = cats.map(c => {
      const count = monthly.byCategory?.[c.key] || 0;
      const pct = monthly.totalAll > 0 ? Math.round((count / mTotal) * 100) : 0;

      return `
        <div class="category-progress-item">
          <div class="category-label-row">
            <span style="color: ${c.color}; font-weight: 700;">${c.label}</span>
            <span style="color: var(--slate-600); font-weight: 600;">${count} Pesan (${pct}%)</span>
          </div>
          <div class="progress-bar-track">
            <div class="progress-bar-fill ${c.fillClass}" style="width: ${pct}%;"></div>
          </div>
        </div>
      `;
    }).join('');
  }
}

function renderLogs() {
  const container = document.getElementById('logsContainer');
  if (!container) return;
  if (!appState.logs || appState.logs.length === 0) {
    container.innerHTML = `
      <div class="settings-card" style="text-align:center;color:var(--slate-600);">
        Belum ada riwayat pengiriman pesan.
      </div>
    `;
    return;
  }

  container.innerHTML = appState.logs.map((log) => {
    const details = (log.results || []).map((r) => {
      const waBadge = r.wa?.ok ? `WA: ✓ (${r.wa.mode})` : `WA: -`;
      const emBadge = r.email?.ok ? `Email: ✓ (${r.email.mode})` : `Email: -`;
      return `<span class="status-pill">${escapeHtml(r.target)} (${escapeHtml(r.name || '-')}): ${escapeHtml(waBadge)} | ${escapeHtml(emBadge)}</span>`;
    }).join(' ');

    return `
      <div class="log-item">
        <div>
          <div style="font-size:12px;color:var(--slate-600);font-weight:700;">
            🕒 ${escapeHtml(log.timestampWita)} • Sumber: ${escapeHtml(log.triggerSource)}
          </div>
          <div style="font-size:15px;font-weight:800;margin:3px 0;">
            ${escapeHtml(log.namaPelanggan)} — ${escapeHtml(log.tipeMobil)} (${escapeHtml(log.noPolisi)})
          </div>
          <div style="font-size:13px;color:var(--toyota-red-dark);font-weight:700;">
            📌 ${escapeHtml(log.categoryTitle)}
          </div>
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;">
          ${details}
        </div>
      </div>
    `;
  }).join('');
}

function populateConfigForms(cfg) {
  if (!cfg) return;
  const safeSetVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val ?? '';
  };
  const safeSetChecked = (id, bool) => {
    const el = document.getElementById(id);
    if (el) el.checked = Boolean(bool);
  };

  const quickPhoneInput = document.getElementById('inputQuickTestPhone');
  const quickEmailInput = document.getElementById('inputQuickTestEmail');
  if (quickPhoneInput && !quickPhoneInput.value && appState.customers[0]?.waPelanggan) {
    quickPhoneInput.value = appState.customers[0].waPelanggan;
  }
  if (quickEmailInput && !quickEmailInput.value) {
    quickEmailInput.value = appState.customers[0]?.emailPelanggan || cfg.smtp?.user || '';
  }

  safeSetVal('inputGoogleSheetApiUrl', cfg.googleSheetApiUrl || cfg.googleSheetCsvUrl || '');
  const gsheetPill = document.getElementById('gsheetStatusPill');
  if (gsheetPill) {
    if (cfg.googleSheetApiUrl) {
      gsheetPill.className = 'gsheet-pill connected';
      gsheetPill.textContent = '🟢 Terhubung ke Google Spreadsheet (2-Arah)';
    } else {
      gsheetPill.className = 'gsheet-pill disconnected';
      gsheetPill.textContent = '⚪ Belum Terhubung (Penyimpanan Lokal / Excel)';
    }
  }

  // WA Gateway
  safeSetChecked('chkWaGatewayEnabled', cfg.waGateway?.enabled);
  safeSetVal('inputWaGatewayUrl', cfg.waGateway?.url || 'https://api.fonnte.com/send');
  safeSetVal('inputWaGatewayToken', cfg.waGateway?.token || '');

  // SMTP
  safeSetChecked('chkSmtpEnabled', cfg.smtp?.enabled);
  safeSetVal('inputSmtpFromName', cfg.smtp?.fromName || 'Agung Toyota Gianyar - Customer Care');
  safeSetVal('inputSmtpUser', cfg.smtp?.user || '');
  safeSetVal('inputSmtpPass', cfg.smtp?.pass || '');

  // Schedule & Templates
  safeSetVal('inputOpStart', cfg.operationalStart || '08:00');
  safeSetVal('inputOpEnd', cfg.operationalEnd || '16:00');
  safeSetVal('inputWaStart', cfg.waWindowStart || '10:00');
  safeSetVal('inputWaEnd', cfg.waWindowEnd || '14:00');
  safeSetVal('inputSamsatDays', cfg.samsatAdvanceDays ?? 3);
  safeSetVal('inputAsuransiDays', cfg.asuransiAdvanceDays ?? 3);
  safeSetVal('inputServiceDays', cfg.serviceAdvanceDays ?? 5);
  safeSetVal('inputDelaySec', cfg.delaySecondsBetweenMessages ?? 8);
  safeSetChecked('chkAutoSendEnabled', cfg.autoSendEnabled !== false);

  // Multi-Year & Operational Year Configuration
  safeSetVal('selectOperationalYearMode', cfg.operationalYearMode || 'AUTO');
  safeSetVal('inputManualOperationalYear', cfg.manualOperationalYear || (new Date()).getFullYear());
  safeSetChecked('chkAutoYearRollover', cfg.autoYearRollover !== false);
  const grpManualYear = document.getElementById('groupManualYear');
  if (grpManualYear) {
    grpManualYear.style.display = (cfg.operationalYearMode === 'MANUAL') ? 'block' : 'none';
  }

  if (cfg.templates) {
    safeSetVal('tplBdayCustomer', cfg.templates.birthdayCustomer || '');
    safeSetVal('tplBdayStaff', cfg.templates.birthdayStaff || '');
    safeSetVal('tplSamsatCustomer', cfg.templates.samsatCustomer || '');
    safeSetVal('tplSamsatStaff', cfg.templates.samsatStaff || '');
    safeSetVal('tplAsuransiCustomer', cfg.templates.asuransiCustomer || '');
    safeSetVal('tplAsuransiStaff', cfg.templates.asuransiStaff || '');
    safeSetVal('tplServiceCustomer', cfg.templates.serviceCustomer || '');
    safeSetVal('tplServiceStaff', cfg.templates.serviceStaff || '');
  }
}

// Open Preview Modal for Pelanggan, Sales, and SA
async function openPreviewModal(customerId, reminderType) {
  try {
    const res = await fetch(`/api/preview/${customerId}/${reminderType}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Gagal memuat preview');

    appState.activePreviewTarget = { customerId, reminderType };
    const { customer, bundle } = data;

    const modalBody = document.getElementById('previewModalBody');
    const makeWaLink = (phone, text) =>
      phone ? `https://wa.me/${phone}?text=${encodeURIComponent(text)}` : '#';

    modalBody.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;margin-bottom:16px;">
        <div>
          <strong>Pelanggan:</strong> ${escapeHtml(customer.namaPelanggan)} (${escapeHtml(customer.tipeMobil)} • ${escapeHtml(customer.noPolisi)})
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <button type="button" class="btn btn-sm ${reminderType === 'ULTAH' ? 'btn-primary' : 'btn-outline'}" data-switch-preview="ULTAH">🎂 Ulang Tahun</button>
          <button type="button" class="btn btn-sm ${reminderType === 'SAMSAT' ? 'btn-primary' : 'btn-outline'}" data-switch-preview="SAMSAT">🚗 Samsat (H-3)</button>
          <button type="button" class="btn btn-sm ${reminderType === 'ASURANSI' ? 'btn-primary' : 'btn-outline'}" data-switch-preview="ASURANSI">🛡️ Asuransi (H-3)</button>
          <button type="button" class="btn btn-sm ${reminderType === 'SERVICE' ? 'btn-primary' : 'btn-outline'}" data-switch-preview="SERVICE">🔧 Next Service (H-5)</button>
        </div>
      </div>

      <div class="preview-grid">
        <!-- 1. PELANGGAN -->
        <div class="preview-box">
          <div>
            <h4>1️⃣ Pesan ke Pelanggan (${escapeHtml(bundle.customer.name)})</h4>
            <div style="font-size:12px;color:var(--slate-600);margin-bottom:8px;">
              📱 WA: <strong>${escapeHtml(bundle.customer.wa || '-')}</strong><br>
              ✉️ Email: <strong>${escapeHtml(bundle.customer.email || '-')}</strong>
            </div>
            <div class="wa-bubble">${escapeHtml(bundle.customer.text)}</div>
          </div>
          <a href="${makeWaLink(bundle.customer.wa, bundle.customer.text)}" target="_blank" rel="noopener" class="btn btn-sm btn-wa">
            💬 Kirim WA Manual (wa.me)
          </a>
        </div>

        <!-- 2. SALES -->
        <div class="preview-box">
          <div>
            <h4>2️⃣ Notifikasi ke Sales (${escapeHtml(bundle.sales.name || '-')})</h4>
            <div style="font-size:12px;color:var(--slate-600);margin-bottom:8px;">
              📱 WA Sales: <strong>${escapeHtml(bundle.sales.wa || '-')}</strong><br>
              ✉️ Email Sales: <strong>${escapeHtml(bundle.sales.email || '-')}</strong>
            </div>
            <div class="wa-bubble">${escapeHtml(bundle.sales.text)}</div>
          </div>
          <a href="${makeWaLink(bundle.sales.wa, bundle.sales.text)}" target="_blank" rel="noopener" class="btn btn-sm btn-wa">
            💬 Kirim WA Manual (wa.me)
          </a>
        </div>

        <!-- 3. SERVICE ADVISOR (SA) -->
        <div class="preview-box">
          <div>
            <h4>3️⃣ Notifikasi ke SA (${escapeHtml(bundle.sa.name || '-')})</h4>
            <div style="font-size:12px;color:var(--slate-600);margin-bottom:8px;">
              📱 WA SA: <strong>${escapeHtml(bundle.sa.wa || '-')}</strong><br>
              ✉️ Email SA: <strong>${escapeHtml(bundle.sa.email || '-')}</strong>
            </div>
            <div class="wa-bubble">${escapeHtml(bundle.sa.text)}</div>
          </div>
          <a href="${makeWaLink(bundle.sa.wa, bundle.sa.text)}" target="_blank" rel="noopener" class="btn btn-sm btn-wa">
            💬 Kirim WA Manual (wa.me)
          </a>
        </div>
      </div>
    `;

    modalBody.querySelectorAll('[data-switch-preview]').forEach((btn) => {
      btn.addEventListener('click', () => {
        openPreviewModal(customerId, btn.getAttribute('data-switch-preview'));
      });
    });

    modalBody.querySelectorAll('a.btn-wa').forEach((link) => {
      link.addEventListener('click', () => {
        triggerSingleSend(customerId, reminderType);
      });
    });

    const dlg = document.getElementById('dialogPreviewMessage');
    if (!dlg.open) dlg.showModal();
  } catch (err) {
    showToast(`Gagal menampilkan preview: ${err.message}`);
  }
}

async function triggerSingleSend(customerId, reminderType) {
  try {
    showToast('⏳ Mengirim WA & Gmail ke Pelanggan, Sales, dan SA...');
    const res = await fetch('/api/send-single', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ customerId, reminderType })
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.error || 'Gagal mengirim pesan');
    const waOk = (data.logEntry?.results || []).some(r => r.wa?.mode === 'WA_OTOMATIS_TERKIRIM');
    const emailOk = (data.logEntry?.results || []).some(r => r.email?.mode === 'GMAIL_OTOMATIS');
    const statusParts = [
      waOk ? '✅ WA Terkirim' : '⚠️ WA Simulasi',
      emailOk ? '✅ Gmail Terkirim' : '⚠️ Gmail Belum Terkirim'
    ].join(' & ');
    showToast(`${statusParts} — Pengingat ${reminderType} untuk ${data.updatedCustomer.namaPelanggan}!`);
    await fetchDashboard(false);
  } catch (err) {
    showToast(`❌ Gagal: ${err.message}`);
  }
}

async function triggerSendAllToday() {
  try {
    showToast('🚀 Memproses & mengirim seluruh pengingat hari ini (WA & Email ke Pelanggan, Sales & SA)...');
    const res = await fetch('/api/send-all-today', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ forceResend: true })
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.error || 'Gagal memproses antrean');
    if (!data.dispatchedCount) {
      showToast('ℹ️ Tidak ada jadwal pengingat aktif untuk hari ini.');
    } else {
      showToast(`✅ WA Terkirim & ✅ Gmail Terkirim (${data.dispatchedCount} pengingat ke Pelanggan, Sales & SA)!`);
    }
    await fetchDashboard(false);
  } catch (err) {
    showToast(`❌ Gagal: ${err.message}`);
  }
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach((b) => {
    b.classList.toggle('active', b.getAttribute('data-tab') === tabId);
  });
  document.querySelectorAll('.sidebar-nav-item[data-tab-nav]').forEach((b) => {
    b.classList.toggle('active', b.getAttribute('data-tab-nav') === tabId);
  });
  document.querySelectorAll('.tab-panel').forEach((p) => {
    p.classList.toggle('active', p.id === tabId);
  });
}

function toggleSidebar(open) {
  const sidebar = document.getElementById('appSidebar');
  const backdrop = document.getElementById('sidebarBackdrop');
  if (!sidebar) return;
  const shouldOpen = typeof open === 'boolean' ? open : !sidebar.classList.contains('open');
  sidebar.classList.toggle('open', shouldOpen);
  if (backdrop) backdrop.classList.toggle('open', shouldOpen);
  document.body.classList.toggle('sidebar-locked', shouldOpen);
}
window.toggleSidebar = toggleSidebar;
window.switchTab = switchTab;

// ============================================================================
// INITIALIZE EVENT LISTENERS
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
  setupDialogLightDismissFallbacks();
  updateLiveWitaClock();
  setInterval(updateLiveWitaClock, 1000);

  fetchDashboard(true);
  // Poll every 3.5 seconds to update QR status & automatic sends
  setInterval(() => fetchDashboard(false), 3500);

  // AccuCare-Style Sidebar Drawer Navigation
  document.querySelectorAll('.sidebar-nav-item[data-tab-nav]').forEach((item) => {
    item.addEventListener('click', () => {
      const targetTab = item.getAttribute('data-tab-nav');
      if (targetTab) {
        switchTab(targetTab);
        toggleSidebar(false);
      }
    });
  });

  const sidebarBtnDb = document.getElementById('sidebarMenuBtnDatabase');
  if (sidebarBtnDb) {
    sidebarBtnDb.addEventListener('click', () => {
      toggleSidebar(false);
      document.getElementById('dialogDatabaseMenu')?.showModal();
    });
  }

  const btnOpenSidebar = document.getElementById('btnOpenSidebar');
  if (btnOpenSidebar) {
    btnOpenSidebar.addEventListener('click', () => toggleSidebar(true));
  }

  const btnCloseSidebar = document.getElementById('btnCloseSidebar');
  if (btnCloseSidebar) {
    btnCloseSidebar.addEventListener('click', () => toggleSidebar(false));
  }

  const sidebarBackdrop = document.getElementById('sidebarBackdrop');
  if (sidebarBackdrop) {
    sidebarBackdrop.addEventListener('click', () => toggleSidebar(false));
  }

  // Pagination Event Listeners
  const btnPrevPage = document.getElementById('btnPrevPage');
  if (btnPrevPage) {
    btnPrevPage.addEventListener('click', () => {
      if (appState.pagination.page > 1) {
        appState.pagination.page -= 1;
        renderCustomerTable();
      }
    });
  }

  const btnNextPage = document.getElementById('btnNextPage');
  if (btnNextPage) {
    btnNextPage.addEventListener('click', () => {
      appState.pagination.page += 1;
      renderCustomerTable();
    });
  }

  const selectPageSize = document.getElementById('selectPageSize');
  if (selectPageSize) {
    selectPageSize.addEventListener('change', (e) => {
      appState.pagination.perPage = e.target.value === 'ALL' ? 'ALL' : Number(e.target.value);
      appState.pagination.page = 1;
      renderCustomerTable();
    });
  }

  const paginationPages = document.getElementById('paginationPages');
  if (paginationPages) {
    paginationPages.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-goto-page]');
      if (!btn) return;
      const targetPage = Number(btn.getAttribute('data-goto-page'));
      if (targetPage && targetPage !== appState.pagination.page) {
        appState.pagination.page = targetPage;
        renderCustomerTable();
      }
    });
  }

  // Tab Switching
  document.querySelectorAll('.tab-btn').forEach((btn) => {
    btn.addEventListener('click', () => switchTab(btn.getAttribute('data-tab')));
  });

  // Quick Bar Buttons
  const btnQuickUp = document.getElementById('btnQuickUpload');
  if (btnQuickUp) {
    btnQuickUp.addEventListener('click', () => {
      document.getElementById('dialogUploadExcel').showModal();
    });
  }
  const btnSendAll = document.getElementById('btnSendAllPending');
  if (btnSendAll) btnSendAll.addEventListener('click', triggerSendAllToday);

  // Quick Test Direct WA + Email Send Form (Both in Tab Autopilot & in Dedicated Menu Modal)
  async function handleQuickTestSubmit(phone, email, submitBtn) {
    if (!phone && !email) {
      showToast('⚠️ Masukkan Nomor WA atau Alamat Email Gmail tujuan terlebih dahulu!');
      return;
    }
    showToast(`📲✉️ Mengirim tes otomatis ke WA (${phone || '-'}) & Gmail (${email || '-'})...`);
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.textContent = '⏳ Mengirim...';
    }
    try {
      const res = await fetch('/api/wa/test-direct', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ phone, email })
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        showToast(`⚠️ ${data.error || 'Gagal mengirim tes.'}`);
        return;
      }
      const parts = [];
      if (data.result?.mode === 'WA_OTOMATIS_TERKIRIM') {
        parts.push(`✅ WA Terkirim (${data.result.recipient})`);
      } else if (data.result?.mode === 'SIMULASI_BELUM_SCAN_QR') {
        parts.push(`ℹ️ WA Simulasi (QR belum di-scan)`);
      } else if (data.result?.reason) {
        parts.push(`⚠️ WA: ${data.result.reason}`);
      }

      if (data.emailResult?.mode === 'GMAIL_OTOMATIS') {
        parts.push(`✅ Gmail Terkirim (${data.emailResult.recipient})`);
      } else if (data.emailResult?.mode === 'SIMULASI_GMAIL') {
        parts.push(`ℹ️ Gmail Simulasi`);
      } else if (data.emailResult?.reason) {
        parts.push(`⚠️ Email: ${data.emailResult.reason}`);
      }
      showToast(parts.length ? parts.join(' | ') : 'ℹ️ Pengiriman selesai.');

      // Close modal if open
      const dlg = document.getElementById('dialogTestSend');
      if (dlg && typeof dlg.close === 'function') {
        try { dlg.close(); } catch (_) {}
      }

      fetchDashboard(false);
    } catch (err) {
      showToast(`❌ Gagal: ${err.message}`);
    } finally {
      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = '⚡ Tes Kirim WA & Gmail';
      }
    }
  }

  const formQuickTestWa = document.getElementById('formQuickTestWa');
  if (formQuickTestWa) {
    formQuickTestWa.addEventListener('submit', (e) => {
      e.preventDefault();
      const phone = (document.getElementById('inputQuickTestPhone')?.value || '').trim();
      const email = (document.getElementById('inputQuickTestEmail')?.value || '').trim();
      const btn = formQuickTestWa.querySelector('button[type="submit"]');
      handleQuickTestSubmit(phone, email, btn);
    });
  }

  const formModalQuickTestWa = document.getElementById('formModalQuickTestWa');
  if (formModalQuickTestWa) {
    formModalQuickTestWa.addEventListener('submit', (e) => {
      e.preventDefault();
      const phone = (document.getElementById('modalInputQuickTestPhone')?.value || '').trim();
      const email = (document.getElementById('modalInputQuickTestEmail')?.value || '').trim();
      const btn = formModalQuickTestWa.querySelector('button[type="submit"]');
      handleQuickTestSubmit(phone, email, btn);
    });
  }

  // Extend Auto-Pilot Window to 08:00 - 18:00 WITA so user can test automatic scheduler right now in the afternoon
  const btnExtend = document.getElementById('btnExtendWindowNow');
  if (btnExtend) {
    btnExtend.addEventListener('click', async () => {
      await fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          waWindowStart: '08:00',
          waWindowEnd: '18:00',
          autoSendEnabled: true
        })
      });
      showToast('🔓 Jam Auto-Pilot diperluas ke 08:00–18:00 WITA! Jika WA sudah terhubung, mesin otomatis langsung jalan.');
      fetchDashboard(true);
    });
  }

  // Reset Today's Sent Flags
  const btnResetFlags = document.getElementById('btnResetTodayFlags');
  if (btnResetFlags) {
    btnResetFlags.addEventListener('click', async () => {
      await fetch('/api/logs/clear-today-flags', { method: 'POST' });
      showToast('🔄 Status pengiriman hari ini di-reset. Siap diuji kirim ulang!');
      fetchDashboard(false);
    });
  }

  // Filter Chips for 4 Reminder Types
  document.querySelectorAll('[data-type-filter]').forEach((chip) => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('[data-type-filter]').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      appState.selectedTypeFilter = chip.getAttribute('data-type-filter');
      renderTodayReminders();
    });
  });

  // Delegated clicks on Today Reminders & Customer Table
  document.body.addEventListener('click', async (e) => {
    const btn = e.target.closest('[data-action]');
    if (!btn) return;
    const action = btn.getAttribute('data-action');
    const id = Number(btn.getAttribute('data-id'));
    const type = btn.getAttribute('data-type');

    if (action === 'preview') {
      openPreviewModal(id, type || 'SERVICE');
    } else if (action === 'send-single') {
      triggerSingleSend(id, type);
    } else if (action === 'edit-customer') {
      const cust = appState.customers.find(c => Number(c.id) === id);
      if (!cust) return;
      const titleEl = document.getElementById('titleCustomerModal');
      if (titleEl) titleEl.textContent = `✏️ Edit Pelanggan: ${cust.namaPelanggan}`;
      const safeVal = (elemId, val) => {
        const el = document.getElementById(elemId);
        if (el) el.value = val ?? '';
      };
      safeVal('custId', cust.id);
      safeVal('custNama', cust.namaPelanggan);
      safeVal('custWa', cust.waPelanggan);
      safeVal('custEmail', cust.emailPelanggan);
      safeVal('custMobil', cust.tipeMobil);
      safeVal('custNopol', cust.noPolisi);
      safeVal('custUltah', cust.tglUlangTahun);
      safeVal('custSamsat', cust.tglSamsat);
      safeVal('custAsuransi', cust.tglAsuransi);
      safeVal('custService', cust.tglService);
      safeVal('custNamaSales', cust.namaSales);
      safeVal('custWaSales', cust.waSales);
      safeVal('custEmailSales', cust.emailSales);
      safeVal('custNamaSA', cust.namaSA);
      safeVal('custWaSA', cust.waSA);
      safeVal('custEmailSA', cust.emailSA);
      document.getElementById('dialogCustomerForm')?.showModal();
    } else if (action === 'delete-customer') {
      if (!confirm('Hapus baris pelanggan ini dari file Excel & Google Spreadsheet?')) return;
      await fetch(`/api/customers/${id}`, { method: 'DELETE' });
      showToast('🗑️ Data pelanggan dihapus dari Excel & Google Spreadsheet.');
      fetchDashboard(false);
    }
  });

  // Confirm Send inside Preview Modal
  const btnConfirmSend = document.getElementById('btnConfirmSendFromPreview');
  if (btnConfirmSend) {
    btnConfirmSend.addEventListener('click', () => {
      if (!appState.activePreviewTarget) return;
      document.getElementById('dialogPreviewMessage')?.close();
      triggerSingleSend(appState.activePreviewTarget.customerId, appState.activePreviewTarget.reminderType);
    });
  }

  // Search Input
  const searchInput = document.getElementById('searchCustomerInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      appState.searchQuery = e.target.value;
      renderCustomerTable();
    });
  }

  // Open Upload Excel Modal
  const btnOpenExcel = document.getElementById('btnOpenExcelModal');
  if (btnOpenExcel) {
    btnOpenExcel.addEventListener('click', () => {
      document.getElementById('dialogUploadExcel')?.showModal();
    });
  }

  // File Input Change label update
  const fileInput = document.getElementById('inputExcelFile');
  if (fileInput) {
    fileInput.addEventListener('change', () => {
      const file = fileInput.files[0];
      const lbl = document.getElementById('selectedFileName');
      if (lbl) {
        lbl.textContent = file
          ? `✅ File dipilih: ${file.name}`
          : 'Klik di sini untuk memilih file Excel (.xlsx)';
      }
    });
  }

  // Handle Upload Excel Submit
  const formUpload = document.getElementById('formUploadExcel');
  if (formUpload && fileInput) {
    formUpload.addEventListener('submit', async (e) => {
      e.preventDefault();
      const file = fileInput.files[0];
      if (!file) {
        showToast('Pilih file Excel (.xlsx) terlebih dahulu');
        return;
      }
      const mode = document.getElementById('selectUploadMode')?.value || 'replace';
      const reader = new FileReader();
      reader.onload = async () => {
        const base64Data = reader.result.split(',')[1];
        const res = await fetch('/api/excel/upload', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ base64Data, mode })
        });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          showToast(`❌ Gagal upload Excel: ${data.error}`);
          return;
        }
        document.getElementById('dialogUploadExcel')?.close();
        showToast(`✅ Berhasil mengimpor ${data.importedCount} pelanggan dari Excel!`);
        fetchDashboard(false);
      };
      reader.readAsDataURL(file);
    });
  }

  // Open Add Customer Modal
  const btnOpenAdd = document.getElementById('btnOpenAddCustomer');
  if (btnOpenAdd) {
    btnOpenAdd.addEventListener('click', () => {
      document.getElementById('formCustomerData')?.reset();
      const custId = document.getElementById('custId');
      if (custId) custId.value = '';
      const titleModal = document.getElementById('titleCustomerModal');
      if (titleModal) titleModal.textContent = '➕ Tambah Data Pelanggan Baru ke Excel';
      document.getElementById('dialogCustomerForm')?.showModal();
    });
  }

  // Save Customer Form
  const formCust = document.getElementById('formCustomerData');
  if (formCust) {
    formCust.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const payload = Object.fromEntries(fd.entries());
      const res = await fetch('/api/customers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        showToast(`❌ Gagal menyimpan: ${data.error}`);
        return;
      }
      document.getElementById('dialogCustomerForm')?.close();
      if (data.autoTriggeredCount > 0) {
        showToast(`✅ Data disimpan & sedang mengirim ${data.autoTriggeredCount} pengingat otomatis ke WA & Email!`);
      } else {
        showToast('✅ Data disimpan! Klik tombol "🚀 Kirim" pada baris pelanggan jika ingin langsung tes kirim WA/Email.');
      }
      fetchDashboard(false);
      setTimeout(() => fetchDashboard(false), 3500);
    });
  }

  // Reset Sample Data
  const btnResetSample = document.getElementById('btnResetSampleData');
  if (btnResetSample) {
    btnResetSample.addEventListener('click', async () => {
      await fetch('/api/customers/reset-sample', { method: 'POST' });
      showToast('✨ Contoh data uji coba Agung Toyota Gianyar berhasil dimuat!');
      fetchDashboard(false);
    });
  }

  // ==========================================================================
  // GOOGLE SPREADSHEET 2-WAY DATABASE INTEGRATION (ACCU CARE PATTERN)
  // ==========================================================================
  
  // Unified Pull Function with Direct Browser Fetch Fallback
  async function executeGoogleSheetsPull() {
    let sheetUrl = (appState.config?.googleSheetApiUrl || localStorage.getItem('toyota_gsheet_url') || '').trim();
    if (!sheetUrl) {
      showToast('⚠️ URL Google Spreadsheet belum dihubungkan. Silakan tempel URL Web App (/exec) di menu ini.');
      const dlg = document.getElementById('dialogDatabaseMenu');
      if (dlg) {
        if (typeof dlg.showModal === 'function') dlg.showModal();
        else dlg.style.display = 'block';
      }
      return;
    }

    localStorage.setItem('toyota_gsheet_url', sheetUrl);
    showToast('🔄 Menarik data pelanggan terbaru dari Google Spreadsheet...');

    // 1. Coba Tarik melalui Backend API
    try {
      const res = await fetch('/api/google-sheets/pull', { method: 'POST' });
      const data = await res.json();
      if (res.ok && data.ok && data.count > 0) {
        showToast(`✅ Sukses! ${data.count} pelanggan dimuat dari Google Spreadsheet.`);
        await fetchDashboard(false);
        return;
      }
    } catch (backendErr) {
      console.warn('Backend pull error, falling back to direct browser fetch:', backendErr);
    }

    // 2. Direct Browser Fetch Fallback (Persis pola AccuCare)
    try {
      showToast('🌐 Mengambil data langsung dari Google Apps Script...');
      const directUrl = `${sheetUrl}${sheetUrl.includes('?') ? '&' : '?'}action=get_customers&t=${Date.now()}`;
      const directRes = await fetch(directUrl);
      const directJson = await directRes.json();

      if (directJson && directJson.success && Array.isArray(directJson.data) && directJson.data.length > 0) {
        // Simpan ke cache lokal browser
        localStorage.setItem('toyota_customers_cache', JSON.stringify(directJson.data));
        
        // Simpan ke backend server
        await fetch('/api/excel/upload', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode: 'replace', customers: directJson.data })
        });

        showToast(`✅ Sukses! ${directJson.data.length} pelanggan dimuat langsung dari Google Spreadsheet.`);
        await fetchDashboard(false);
        return;
      } else if (directJson && directJson.data && directJson.data.length === 0) {
        showToast('ℹ️ Google Spreadsheet Anda saat ini masih kosong (hanya header).');
        return;
      }
      throw new Error(directJson?.error || 'Respon Google Spreadsheet tidak valid');
    } catch (directErr) {
      showToast(`❌ Gagal tarik data: ${directErr.message}. Pastikan izin deployment Web App = "Anyone"`);
    }
  }

  // 1. Simpan & Hubungkan URL Google Apps Script Web App
  const formSaveGSheet = document.getElementById('formSaveGoogleSheetUrl');
  if (formSaveGSheet) {
    formSaveGSheet.addEventListener('submit', async (e) => {
      e.preventDefault();
      const url = (document.getElementById('inputGoogleSheetApiUrl')?.value || '').trim();
      if (!url) {
        showToast('⚠️ Masukkan URL Web App Google Apps Script (/exec) terlebih dahulu');
        return;
      }
      localStorage.setItem('toyota_gsheet_url', url);
      showToast('⏳ Menghubungkan dan memverifikasi koneksi Google Spreadsheet...');
      try {
        const res = await fetch('/api/google-sheets/save-url', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ url })
        });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          showToast(`❌ Gagal terhubung: ${data.error || 'Periksa izin publikasi Web App'}`);
          return;
        }
        showToast(`✅ ${data.message || 'Berhasil terhubung ke Google Spreadsheet!'}`);
        await executeGoogleSheetsPull();
      } catch (err) {
        showToast(`❌ Terjadi kesalahan: ${err.message}`);
      }
    });
  }

  // 2. Tarik Data Terbaru dari Google Spreadsheet (Pull)
  const btnPullGSheet = document.getElementById('btnPullFromGSheet');
  if (btnPullGSheet) {
    btnPullGSheet.addEventListener('click', executeGoogleSheetsPull);
  }

  // 3. Quick Pull langsung dari Toolbar Utama
  const btnQuickPull = document.getElementById('btnQuickPullFromGSheet');
  if (btnQuickPull) {
    btnQuickPull.addEventListener('click', executeGoogleSheetsPull);
  }

  // 4. Unggah Semua Data ke Google Spreadsheet (Push Backup)
  const btnPushGSheet = document.getElementById('btnPushToGSheet');
  if (btnPushGSheet) {
    btnPushGSheet.addEventListener('click', async () => {
      if (!confirm('Apakah Anda ingin mengunggah semua data pelanggan saat ini untuk menggantikan isi Google Spreadsheet Anda?')) return;
      showToast('☁️ Mengunggah data ke Google Spreadsheet...');
      try {
        const res = await fetch('/api/google-sheets/push', { method: 'POST' });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          showToast(`❌ Gagal upload: ${data.error}`);
          return;
        }
        showToast(`✅ Berhasil! ${data.count} pelanggan tersimpan aman di Google Spreadsheet.`);
      } catch (err) {
        showToast(`❌ Terjadi kesalahan: ${err.message}`);
      }
    });
  }

  // 5. Modal Panduan & Salin Script
  const btnOpenGScriptModal = document.getElementById('btnOpenGoogleScriptModal');
  const txtAreaGScript = document.getElementById('googleAppsScriptCodeArea');
  const dlgGScript = document.getElementById('dialogGoogleScript');

  async function loadAndShowScriptModal() {
    if (txtAreaGScript && !txtAreaGScript.value) {
      try {
        const res = await fetch('/google-apps-script.gs');
        if (res.ok) {
          txtAreaGScript.value = await res.text();
        }
      } catch (err) {
        console.error('Gagal mengambil script file:', err);
      }
    }
    if (dlgGScript) dlgGScript.showModal();
  }

  if (btnOpenGScriptModal) {
    btnOpenGScriptModal.addEventListener('click', loadAndShowScriptModal);
  }

  function copyGScriptText() {
    if (!txtAreaGScript) return;
    navigator.clipboard.writeText(txtAreaGScript.value).then(() => {
      showToast('📋 Kode Google Apps Script berhasil disalin ke clipboard!');
    }).catch(() => {
      txtAreaGScript.select();
      document.execCommand('copy');
      showToast('📋 Kode Google Apps Script disalin!');
    });
  }

  const btnCopyGScript1 = document.getElementById('btnCopyGoogleScript');
  if (btnCopyGScript1) btnCopyGScript1.addEventListener('click', copyGScriptText);

  const btnCopyGScript2 = document.getElementById('btnCopyGoogleScript2');
  if (btnCopyGScript2) {
    btnCopyGScript2.addEventListener('click', () => {
      copyGScriptText();
      if (dlgGScript) dlgGScript.close();
    });
  }

  // 6. Tombol Garis 3 (☰) Buka Menu Database & File Excel
  const btnToggleDbMenu = document.getElementById('btnToggleDatabaseMenu');
  const dlgDbMenu = document.getElementById('dialogDatabaseMenu');
  if (btnToggleDbMenu && dlgDbMenu) {
    btnToggleDbMenu.addEventListener('click', () => {
      dlgDbMenu.showModal();
    });
  }

  // WhatsApp QR Connect / Disconnect
  const btnConnect = document.getElementById('btnConnectWa');
  if (btnConnect) {
    btnConnect.addEventListener('click', async () => {
      showToast('📲 Menyiapkan QR Code WhatsApp...');
      await fetch('/api/wa/connect', { method: 'POST' });
      setTimeout(() => fetchDashboard(false), 1500);
    });
  }

  const btnDisconnect = document.getElementById('btnDisconnectWa');
  if (btnDisconnect) {
    btnDisconnect.addEventListener('click', async () => {
      await fetch('/api/wa/disconnect', { method: 'POST' });
      showToast('🔌 Sesi WhatsApp telah diputuskan.');
      fetchDashboard(false);
    });
  }

  // Password Toggle for Gmail App Password
  const btnTogglePass = document.getElementById('btnToggleSmtpPass');
  const inputPass = document.getElementById('inputSmtpPass');
  if (btnTogglePass && inputPass) {
    btnTogglePass.addEventListener('click', () => {
      const isHidden = inputPass.type === 'password';
      inputPass.type = isHidden ? 'text' : 'password';
      btnTogglePass.setAttribute('aria-pressed', String(isHidden));
      btnTogglePass.textContent = isHidden ? '🙈 Sembunyikan' : '👁️ Lihat';
    });
  }

  // Save WA Gateway Form
  const formWaGw = document.getElementById('formWaGateway');
  if (formWaGw) {
    formWaGw.addEventListener('submit', async (e) => {
      e.preventDefault();
      const payload = {
        waGateway: {
          enabled: Boolean(document.getElementById('chkWaGatewayEnabled')?.checked),
          url: (document.getElementById('inputWaGatewayUrl')?.value || '').trim(),
          token: (document.getElementById('inputWaGatewayToken')?.value || '').trim()
        }
      };
      await fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      showToast('💾 Pengaturan WA Gateway berhasil disimpan.');
      fetchDashboard(false);
    });
  }

  // Save SMTP Form
  const formSmtp = document.getElementById('formSmtpConfig');
  if (formSmtp) {
    formSmtp.addEventListener('submit', async (e) => {
      e.preventDefault();
      const userVal = (document.getElementById('inputSmtpUser')?.value || '').trim();
      const passVal = (document.getElementById('inputSmtpPass')?.value || '').trim();
      const chkSmtp = document.getElementById('chkSmtpEnabled');
      const autoEnable = Boolean(userVal && passVal) || (chkSmtp ? chkSmtp.checked : false);
      if (chkSmtp) chkSmtp.checked = autoEnable;

      const payload = {
        smtp: {
          enabled: autoEnable,
          service: 'gmail',
          fromName: (document.getElementById('inputSmtpFromName')?.value || '').trim(),
          user: userVal,
          pass: passVal
        }
      };
      await fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      showToast('💾 Pengaturan Gmail berhasil disimpan & diaktifkan!');
      fetchDashboard(false);
    });
  }

  // Save Schedule & Templates Form
  const formSched = document.getElementById('formScheduleAndTemplates');
  if (formSched) {
    formSched.addEventListener('submit', async (e) => {
      e.preventDefault();
      const payload = {
        operationalStart: document.getElementById('inputOpStart')?.value || '08:00',
        operationalEnd: document.getElementById('inputOpEnd')?.value || '16:00',
        waWindowStart: document.getElementById('inputWaStart')?.value || '10:00',
        waWindowEnd: document.getElementById('inputWaEnd')?.value || '14:00',
        samsatAdvanceDays: Number(document.getElementById('inputSamsatDays')?.value ?? 3),
        asuransiAdvanceDays: Number(document.getElementById('inputAsuransiDays')?.value ?? 3),
        serviceAdvanceDays: Number(document.getElementById('inputServiceDays')?.value ?? 5),
        delaySecondsBetweenMessages: Number(document.getElementById('inputDelaySec')?.value ?? 8),
        autoSendEnabled: document.getElementById('chkAutoSendEnabled') ? document.getElementById('chkAutoSendEnabled').checked : true,
        operationalYearMode: document.getElementById('selectOperationalYearMode')?.value || 'AUTO',
        manualOperationalYear: Number(document.getElementById('inputManualOperationalYear')?.value || 2026),
        autoYearRollover: document.getElementById('chkAutoYearRollover') ? document.getElementById('chkAutoYearRollover').checked : true,
        templates: {
          birthdayCustomer: document.getElementById('tplBdayCustomer')?.value || '',
          birthdayStaff: document.getElementById('tplBdayStaff')?.value || '',
          samsatCustomer: document.getElementById('tplSamsatCustomer')?.value || '',
          samsatStaff: document.getElementById('tplSamsatStaff')?.value || '',
          asuransiCustomer: document.getElementById('tplAsuransiCustomer')?.value || '',
          asuransiStaff: document.getElementById('tplAsuransiStaff')?.value || '',
          serviceCustomer: document.getElementById('tplServiceCustomer')?.value || '',
          serviceStaff: document.getElementById('tplServiceStaff')?.value || ''
        }
      };
      await fetch('/api/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      showToast('✅ Pengaturan Jadwal WITA, Tahun Operasional & Template Pesan berhasil disimpan!');
      fetchDashboard(false);
    });
  }

  // Toggle Manual Year input visibility on change
  const selectOpYearMode = document.getElementById('selectOperationalYearMode');
  if (selectOpYearMode) {
    selectOpYearMode.addEventListener('change', (e) => {
      const grp = document.getElementById('groupManualYear');
      if (grp) grp.style.display = (e.target.value === 'MANUAL') ? 'block' : 'none';
    });
  }

  // ==================== REPORT & ANALYTICS EVENT LISTENERS ====================
  const reportDateInput = document.getElementById('inputReportDailyDate');
  if (reportDateInput) {
    reportDateInput.addEventListener('change', (e) => {
      appState.selectedReportDate = e.target.value;
      fetchReportAnalytics(e.target.value, null);
    });
  }

  const monthPicker = document.getElementById('inputReportMonthPicker');
  if (monthPicker) {
    monthPicker.addEventListener('change', (e) => {
      const val = e.target.value;
      if (!val) return;
      appState.selectedReportMonth = val;
      fetchReportAnalytics(null, val);
    });
  }

  const btnReportThisMonth = document.getElementById('btnReportThisMonth');
  if (btnReportThisMonth) {
    btnReportThisMonth.addEventListener('click', () => {
      const now = new Date();
      const curYear = appState.reportData?.curDate ? appState.reportData.curDate.slice(0, 4) : String(now.getFullYear());
      const curM = appState.reportData?.curDate ? appState.reportData.curDate.slice(5, 7) : String(now.getMonth() + 1).padStart(2, '0');
      const combinedMonth = `${curYear}-${curM}`;
      const picker = document.getElementById('inputReportMonthPicker');
      if (picker) picker.value = combinedMonth;
      appState.selectedReportMonth = combinedMonth;
      fetchReportAnalytics(null, combinedMonth);
    });
  }

  const btnDownloadReport = document.getElementById('btnDownloadReportExcel');
  if (btnDownloadReport) {
    btnDownloadReport.addEventListener('click', () => {
      const q = new URLSearchParams();
      if (appState.activeReportPeriod === 'daily' && appState.selectedReportDate) {
        q.set('date', appState.selectedReportDate);
      } else if (appState.activeReportPeriod === 'monthly' && appState.selectedReportMonth) {
        q.set('month', appState.selectedReportMonth);
      }
      showToast('📥 Mengunduh Laporan Pengiriman Excel...');
      window.open(`/api/reports/export?${q.toString()}`, '_blank');
    });
  }

  const btnRefreshReport = document.getElementById('btnRefreshReport');
  if (btnRefreshReport) {
    btnRefreshReport.addEventListener('click', () => {
      showToast('🔄 Memperbarui data laporan pengiriman...');
      fetchReportAnalytics();
    });
  }
});
