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
  activePreviewTarget: null // { customerId, reminderType }
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
    appState.autoPilotStatusText = data.autoPilotStatusText || '';

    renderHeaderAndStats(data);
    renderTodayReminders();
    renderCustomerTable();
    renderWaAndSmtpStatus();
    renderLogs();

    if (populateForms) {
      populateConfigForms(appState.config);
    }
  } catch (err) {
    console.error('Gagal memuat dashboard:', err);
  }
}

function renderHeaderAndStats(data) {
  document.getElementById('statTotalCustomers').textContent = data.stats.totalCustomers ?? 0;
  document.getElementById('statBirthday').textContent = data.stats.dueBirthdayToday ?? 0;
  document.getElementById('statSamsat').textContent = data.stats.dueSamsatH3 ?? 0;
  document.getElementById('statAsuransi').textContent = data.stats.dueAsuransiH3 ?? 0;
  const statServiceEl = document.getElementById('statService');
  if (statServiceEl) statServiceEl.textContent = data.stats.dueServiceH5 ?? 0;
  document.getElementById('badgePendingToday').textContent = data.stats.totalPendingToday ?? 0;

  const opStart = data.config.operationalStart || '08:00';
  const opEnd = data.config.operationalEnd || '16:00';
  const waStart = data.config.waWindowStart || '10:00';
  const waEnd = data.config.waWindowEnd || '14:00';

  document.getElementById('officeHoursLabel').textContent = `${opStart} – ${opEnd} WITA`;
  document.getElementById('waWindowLabel').textContent = `${waStart} – ${waEnd} WITA`;

  const officeBadge = document.getElementById('officeStatusBadge');
  const waBadge = document.getElementById('waWindowBadge');
  officeBadge.classList.toggle('active-now', Boolean(data.inOperationalHours));
  waBadge.classList.toggle('active-now', Boolean(data.inWaWindow));

  const gsheetPill = document.getElementById('gsheetStatusPill');
  const toolbarBadge = document.getElementById('toolbarGSheetBadge');
  const isGsheetConnected = Boolean(data.config?.googleSheetApiUrl);

  if (gsheetPill) {
    if (isGsheetConnected) {
      gsheetPill.className = 'gsheet-pill connected';
      gsheetPill.textContent = '🟢 Terhubung ke Google Spreadsheet (2-Arah)';
    } else {
      gsheetPill.className = 'gsheet-pill disconnected';
      gsheetPill.textContent = '⚪ Belum Terhubung (Penyimpanan Lokal / Excel)';
    }
  }

  if (toolbarBadge) {
    if (isGsheetConnected) {
      toolbarBadge.className = 'toolbar-mini-badge connected';
      toolbarBadge.textContent = '🟢 Google Sheets';
    } else {
      toolbarBadge.className = 'toolbar-mini-badge disconnected';
      toolbarBadge.textContent = '⚪ Excel';
    }
  }

  const summaryCountEl = document.getElementById('summaryCountCustomers');
  if (summaryCountEl) {
    summaryCountEl.textContent = data.stats.totalCustomers ?? (appState.customers?.length || 0);
  }
}

function renderTodayReminders() {
  const container = document.getElementById('todayRemindersContainer');
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

    const statusBadgeHtml = isTomorrow
      ? `<span class="sent-pill pending" style="background:#dbeafe;color:#1e40af;border:1px solid #93c5fd;">
           🗓️ Terjadwal Otomatis Besok (${formatDateIndoShort(rem.scheduledDate)} • 10:00–14:00 WITA)
         </span>`
      : `<span class="sent-pill ${rem.alreadySentToday ? 'done' : 'pending'}">
           ${rem.alreadySentToday ? '✓ WA & Gmail Terkirim' : '⏳ Menunggu Kirim (10:00–14:00 WITA)'}
         </span>`;

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

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="11" style="text-align: center; padding: 28px; color: var(--slate-600);">
          Data pelanggan tidak ditemukan. Silakan tambah pelanggan atau upload file Excel.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = filtered.map((c, idx) => {
    const bdayBadge = c.isBirthdayToday
      ? `<span class="rem-badge ULTAH" style="margin-top:4px;">🎂 Hari Ini Ultah!</span>`
      : '';

    const samsatHint = c.samsatDiff === 3
      ? `<span class="rem-badge SAMSAT" style="margin-top:4px;">🚗 H-3 Hari Ini!</span>`
      : c.samsatDiff !== null && c.samsatDiff >= 0 && c.samsatDiff <= 14
        ? `<div style="font-size:11px;color:#b45309;font-weight:700;">H-${c.samsatDiff} hari lagi</div>`
        : '';

    const asuransiHint = c.asuransiDiff === 3
      ? `<span class="rem-badge ASURANSI" style="margin-top:4px;">🛡️ H-3 Hari Ini!</span>`
      : c.asuransiDiff !== null && c.asuransiDiff >= 0 && c.asuransiDiff <= 14
        ? `<div style="font-size:11px;color:#1d4ed8;font-weight:700;">H-${c.asuransiDiff} hari lagi</div>`
        : '';

    const serviceAdv = Number(appState.config.serviceAdvanceDays ?? 5);
    const serviceHint = c.serviceDiff === serviceAdv
      ? `<span class="rem-badge SERVICE" style="margin-top:4px;">🔧 H-${serviceAdv} Hari Ini!</span>`
      : c.serviceDiff !== null && c.serviceDiff >= 0 && c.serviceDiff <= 14
        ? `<div style="font-size:11px;color:#0f766e;font-weight:700;">H-${c.serviceDiff} hari lagi</div>`
        : '';

    return `
      <tr>
        <td><strong>${idx + 1}</strong></td>
        <td>
          <strong>${escapeHtml(c.namaPelanggan)}</strong>
          <div style="font-size:12px;color:var(--slate-600);">📱 ${escapeHtml(c.waPelanggan || '-')}</div>
          <div style="font-size:11.5px;color:var(--slate-600);">✉️ ${escapeHtml(c.emailPelanggan || '-')}</div>
        </td>
        <td>
          <strong>${escapeHtml(c.tipeMobil)}</strong><br>
          <span class="plate-tag">${escapeHtml(c.noPolisi)}</span>
        </td>
        <td>
          <div>${formatDateIndoShort(c.tglUlangTahun)}</div>
          ${bdayBadge}
        </td>
        <td>
          <div>${formatDateIndoShort(c.tglSamsat)}</div>
          ${samsatHint}
        </td>
        <td>
          <div>${formatDateIndoShort(c.tglAsuransi)}</div>
          ${asuransiHint}
        </td>
        <td>
          <div>${formatDateIndoShort(c.tglService)}</div>
          ${serviceHint}
        </td>
        <td>
          <strong>${escapeHtml(c.namaSales || '-')}</strong>
          <div style="font-size:11.5px;color:var(--slate-600);">📱 ${escapeHtml(c.waSales || '-')}</div>
          <div style="font-size:11px;color:var(--slate-600);">${escapeHtml(c.emailSales || '')}</div>
        </td>
        <td>
          <strong>${escapeHtml(c.namaSA || '-')}</strong>
          <div style="font-size:11.5px;color:var(--slate-600);">📱 ${escapeHtml(c.waSA || '-')}</div>
          <div style="font-size:11px;color:var(--slate-600);">${escapeHtml(c.emailSA || '')}</div>
        </td>
        <td>
          <div style="font-size:11.5px;color:var(--slate-600);max-width:180px;">
            ${c.statusLog
              ? `<span class="sent-pill done">✓ WA &amp; Gmail Terkirim (${escapeHtml(c.statusLog)})</span>`
              : '<span style="color:#94a3b8;">Belum ada pengiriman</span>'}
          </div>
        </td>
        <td>
          <div style="display:flex;gap:6px;flex-wrap:wrap;">
            <button type="button" class="btn btn-sm btn-primary" data-action="send-single" data-id="${c.id}" data-type="${c.activeReminders?.[0]?.type || 'SERVICE'}" title="Langsung Kirim WA & Email ke Pelanggan, Sales, dan SA">🚀 Kirim</button>
            <button type="button" class="btn btn-sm btn-outline" data-action="edit-customer" data-id="${c.id}">✏️ Edit</button>
            <button type="button" class="btn btn-sm btn-outline" data-action="preview" data-id="${c.id}" data-type="${c.activeReminders?.[0]?.type || 'SERVICE'}" title="Preview & Pilih Tipe Pesan">👁️ Pesan</button>
            <button type="button" class="btn btn-sm btn-danger" data-action="delete-customer" data-id="${c.id}">🗑️</button>
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
    waPill.textContent = `✓ Terhubung (${appState.waState.connectedNumber || 'WhatsApp Aktif'})`;
    waPill.className = 'status-pill connected';
    waDesc.innerHTML = `✅ <strong>WhatsApp Kantor Terhubung!</strong> (${escapeHtml(appState.waState.connectedName || '')}). Siap mengirim pesan otomatis pada pukul <strong>${escapeHtml(appState.config.waWindowStart || '10:00')} – ${escapeHtml(appState.config.waWindowEnd || '14:00')} WITA</strong>.`;
    qrImg.classList.add('hidden');

    if (mainWaChip) {
      mainWaChip.textContent = `📱 WA Terhubung: ${appState.waState.connectedNumber || 'Aktif'}`;
      mainWaChip.style.background = '#16a34a';
    }
    if (mainNotConn) mainNotConn.classList.add('hidden');
    if (mainConnBox) mainConnBox.classList.remove('hidden');
    if (mainConnNum) mainConnNum.textContent = `📱 WA: ${appState.waState.connectedName || ''} (${appState.waState.connectedNumber || ''})`;
  } else if (appState.waState.status === 'QR_READY' && appState.waState.qrDataUrl) {
    waPill.textContent = '📲 Menunggu Scan QR Code';
    waPill.className = 'status-pill';
    waDesc.innerHTML = 'Silakan buka WhatsApp di HP Kantor &rarr; <strong>Perangkat Tertaut</strong> &rarr; <strong>Tautkan Perangkat</strong> lalu scan QR di bawah ini:';
    qrImg.src = appState.waState.qrDataUrl;
    qrImg.classList.remove('hidden');

    if (mainWaChip) {
      mainWaChip.textContent = '📲 WA Belum Tertaut — Silakan Scan QR di Kanan';
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
    waPill.textContent = '⏳ Menyiapkan QR Code...';
    waPill.className = 'status-pill';
    waDesc.textContent = 'Sedang menghubungi server WhatsApp Web, mohon tunggu beberapa detik...';
    qrImg.classList.add('hidden');

    if (mainWaChip) mainWaChip.textContent = '⏳ Menyiapkan QR WhatsApp...';
    if (mainNotConn) mainNotConn.classList.remove('hidden');
    if (mainConnBox) mainConnBox.classList.add('hidden');
    if (mainQrImg) mainQrImg.classList.add('hidden');
    if (mainQrLoading) mainQrLoading.classList.remove('hidden');
  } else {
    waPill.textContent = 'Siap Scan QR';
    waPill.className = 'status-pill';
    waDesc.innerHTML = 'Klik tombol <strong>"Tampilkan / Hubungkan QR WhatsApp"</strong> di bawah untuk menautkan WhatsApp kantor secara gratis.';
    qrImg.classList.add('hidden');

    if (mainWaChip) mainWaChip.textContent = '📲 Menunggu Koneksi QR';
  }

  const smtpPill = document.getElementById('smtpConnectionPill');
  if (appState.config.smtp?.enabled && appState.config.smtp?.user && appState.config.smtp?.pass) {
    smtpPill.textContent = `✓ Gmail Aktif (${appState.config.smtp.user})`;
    smtpPill.className = 'status-pill connected';
    if (mainGmailChip) {
      mainGmailChip.textContent = `✉️ Gmail Terhubung: ${appState.config.smtp.user}`;
      mainGmailChip.style.background = '#16a34a';
    }
    if (mainGmailAcc) {
      mainGmailAcc.textContent = `✉️ Gmail: ${appState.config.smtp.user}`;
    }
  } else {
    smtpPill.textContent = 'Mode Simulasi Email';
    smtpPill.className = 'status-pill';
    if (mainGmailChip) {
      mainGmailChip.textContent = '✉️ Gmail Belum Diatur';
      mainGmailChip.style.background = '#64748b';
    }
    if (mainGmailAcc) {
      mainGmailAcc.textContent = '';
    }
  }
}

function renderLogs() {
  const container = document.getElementById('logsContainer');
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
  const quickPhoneInput = document.getElementById('inputQuickTestPhone');
  const quickEmailInput = document.getElementById('inputQuickTestEmail');
  if (quickPhoneInput && !quickPhoneInput.value && appState.customers[0]?.waPelanggan) {
    quickPhoneInput.value = appState.customers[0].waPelanggan;
  }
  if (quickEmailInput && !quickEmailInput.value) {
    quickEmailInput.value = appState.customers[0]?.emailPelanggan || cfg.smtp?.user || '';
  }

  const gsheetInput = document.getElementById('inputGoogleSheetApiUrl');
  if (gsheetInput) {
    gsheetInput.value = cfg.googleSheetApiUrl || cfg.googleSheetCsvUrl || '';
  }
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
  document.getElementById('chkWaGatewayEnabled').checked = Boolean(cfg.waGateway?.enabled);
  document.getElementById('inputWaGatewayUrl').value = cfg.waGateway?.url || 'https://api.fonnte.com/send';
  document.getElementById('inputWaGatewayToken').value = cfg.waGateway?.token || '';

  // SMTP
  document.getElementById('chkSmtpEnabled').checked = Boolean(cfg.smtp?.enabled);
  document.getElementById('inputSmtpFromName').value = cfg.smtp?.fromName || 'Agung Toyota Gianyar - Customer Care';
  document.getElementById('inputSmtpUser').value = cfg.smtp?.user || '';
  document.getElementById('inputSmtpPass').value = cfg.smtp?.pass || '';

  // Schedule & Templates
  document.getElementById('inputOpStart').value = cfg.operationalStart || '08:00';
  document.getElementById('inputOpEnd').value = cfg.operationalEnd || '16:00';
  document.getElementById('inputWaStart').value = cfg.waWindowStart || '10:00';
  document.getElementById('inputWaEnd').value = cfg.waWindowEnd || '14:00';
  document.getElementById('inputSamsatDays').value = cfg.samsatAdvanceDays ?? 3;
  document.getElementById('inputAsuransiDays').value = cfg.asuransiAdvanceDays ?? 3;
  const srvDaysEl = document.getElementById('inputServiceDays');
  if (srvDaysEl) srvDaysEl.value = cfg.serviceAdvanceDays ?? 5;
  document.getElementById('inputDelaySec').value = cfg.delaySecondsBetweenMessages ?? 8;
  document.getElementById('chkAutoSendEnabled').checked = cfg.autoSendEnabled !== false;

  if (cfg.templates) {
    document.getElementById('tplBdayCustomer').value = cfg.templates.birthdayCustomer || '';
    document.getElementById('tplBdayStaff').value = cfg.templates.birthdayStaff || '';
    document.getElementById('tplSamsatCustomer').value = cfg.templates.samsatCustomer || '';
    document.getElementById('tplSamsatStaff').value = cfg.templates.samsatStaff || '';
    document.getElementById('tplAsuransiCustomer').value = cfg.templates.asuransiCustomer || '';
    document.getElementById('tplAsuransiStaff').value = cfg.templates.asuransiStaff || '';
    const tplSrvCust = document.getElementById('tplServiceCustomer');
    const tplSrvStaff = document.getElementById('tplServiceStaff');
    if (tplSrvCust) tplSrvCust.value = cfg.templates.serviceCustomer || '';
    if (tplSrvStaff) tplSrvStaff.value = cfg.templates.serviceStaff || '';
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
  document.querySelectorAll('.tab-panel').forEach((p) => {
    p.classList.toggle('active', p.id === tabId);
  });
}

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
  document.getElementById('btnSendAllPending').addEventListener('click', triggerSendAllToday);

  // Quick Test Direct WA + Email Send Form on Front Page
  const formQuickTestWa = document.getElementById('formQuickTestWa');
  if (formQuickTestWa) {
    formQuickTestWa.addEventListener('submit', async (e) => {
      e.preventDefault();
      const phone = document.getElementById('inputQuickTestPhone').value.trim();
      const email = (document.getElementById('inputQuickTestEmail')?.value || '').trim();
      showToast(`📲✉️ Mengirim tes otomatis ke WA (${phone || '-'}) & Gmail (${email || '-'})...`);
      const res = await fetch('/api/wa/test-direct', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ phone, email })
      });
      const data = await res.json();
      if (!res.ok || !data.ok) {
        showToast(`⚠️ ${data.error}`);
        return;
      }
      const parts = [];
      if (data.result?.mode === 'WA_OTOMATIS_TERKIRIM') {
        parts.push(`✅ WA Terkirim (${data.result.recipient})`);
      }
      if (data.emailResult?.mode === 'GMAIL_OTOMATIS') {
        parts.push(`✅ Gmail Terkirim (${data.emailResult.recipient})`);
      }
      showToast(parts.length ? parts.join(' & ') : 'ℹ️ Pengiriman selesai.');
      fetchDashboard(false);
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
  document.getElementById('btnResetTodayFlags').addEventListener('click', async () => {
    await fetch('/api/logs/clear-today-flags', { method: 'POST' });
    showToast('🔄 Status pengiriman hari ini di-reset. Siap diuji kirim ulang!');
    fetchDashboard(false);
  });

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
      document.getElementById('titleCustomerModal').textContent = `✏️ Edit Pelanggan: ${cust.namaPelanggan}`;
      document.getElementById('custId').value = cust.id;
      document.getElementById('custNama').value = cust.namaPelanggan || '';
      document.getElementById('custWa').value = cust.waPelanggan || '';
      document.getElementById('custEmail').value = cust.emailPelanggan || '';
      document.getElementById('custMobil').value = cust.tipeMobil || '';
      document.getElementById('custNopol').value = cust.noPolisi || '';
      document.getElementById('custUltah').value = cust.tglUlangTahun || '';
      document.getElementById('custSamsat').value = cust.tglSamsat || '';
      document.getElementById('custAsuransi').value = cust.tglAsuransi || '';
      const srvInput = document.getElementById('custService');
      if (srvInput) srvInput.value = cust.tglService || '';
      document.getElementById('custNamaSales').value = cust.namaSales || '';
      document.getElementById('custWaSales').value = cust.waSales || '';
      document.getElementById('custEmailSales').value = cust.emailSales || '';
      document.getElementById('custNamaSA').value = cust.namaSA || '';
      document.getElementById('custWaSA').value = cust.waSA || '';
      document.getElementById('custEmailSA').value = cust.emailSA || '';
      document.getElementById('dialogCustomerForm').showModal();
    } else if (action === 'delete-customer') {
      if (!confirm('Hapus baris pelanggan ini dari file Excel?')) return;
      await fetch(`/api/customers/${id}`, { method: 'DELETE' });
      showToast('🗑️ Data pelanggan dihapus dari Excel.');
      fetchDashboard(false);
    }
  });

  // Confirm Send inside Preview Modal
  document.getElementById('btnConfirmSendFromPreview').addEventListener('click', () => {
    if (!appState.activePreviewTarget) return;
    document.getElementById('dialogPreviewMessage').close();
    triggerSingleSend(appState.activePreviewTarget.customerId, appState.activePreviewTarget.reminderType);
  });

  // Search Input
  document.getElementById('searchCustomerInput').addEventListener('input', (e) => {
    appState.searchQuery = e.target.value;
    renderCustomerTable();
  });

  // Open Upload Excel Modal
  document.getElementById('btnOpenExcelModal').addEventListener('click', () => {
    document.getElementById('dialogUploadExcel').showModal();
  });

  // File Input Change label update
  const fileInput = document.getElementById('inputExcelFile');
  fileInput.addEventListener('change', () => {
    const file = fileInput.files[0];
    document.getElementById('selectedFileName').textContent = file
      ? `✅ File dipilih: ${file.name}`
      : 'Klik di sini untuk memilih file Excel (.xlsx)';
  });

  // Handle Upload Excel Submit
  document.getElementById('formUploadExcel').addEventListener('submit', async (e) => {
    e.preventDefault();
    const file = fileInput.files[0];
    if (!file) {
      showToast('Pilih file Excel (.xlsx) terlebih dahulu');
      return;
    }
    const mode = document.getElementById('selectUploadMode').value;
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
      document.getElementById('dialogUploadExcel').close();
      showToast(`✅ Berhasil mengimpor ${data.importedCount} pelanggan dari Excel!`);
      fetchDashboard(false);
    };
    reader.readAsDataURL(file);
  });

  // Open Add Customer Modal
  document.getElementById('btnOpenAddCustomer').addEventListener('click', () => {
    document.getElementById('formCustomerData').reset();
    document.getElementById('custId').value = '';
    document.getElementById('titleCustomerModal').textContent = '➕ Tambah Data Pelanggan Baru ke Excel';
    document.getElementById('dialogCustomerForm').showModal();
  });

  // Save Customer Form
  document.getElementById('formCustomerData').addEventListener('submit', async (e) => {
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
    document.getElementById('dialogCustomerForm').close();
    if (data.autoTriggeredCount > 0) {
      showToast(`✅ Data disimpan & sedang mengirim ${data.autoTriggeredCount} pengingat otomatis ke WA & Email!`);
    } else {
      showToast('✅ Data disimpan! Klik tombol "🚀 Kirim" pada baris pelanggan jika ingin langsung tes kirim WA/Email.');
    }
    fetchDashboard(false);
    setTimeout(() => fetchDashboard(false), 3500);
  });

  // Reset Sample Data
  document.getElementById('btnResetSampleData').addEventListener('click', async () => {
    await fetch('/api/customers/reset-sample', { method: 'POST' });
    showToast('✨ Contoh data uji coba Agung Toyota Gianyar berhasil dimuat!');
    fetchDashboard(false);
  });

  // ==========================================================================
  // GOOGLE SPREADSHEET 2-WAY DATABASE INTEGRATION (ACCU CARE PATTERN)
  // ==========================================================================
  
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
        fetchDashboard(false);
      } catch (err) {
        showToast(`❌ Terjadi kesalahan: ${err.message}`);
      }
    });
  }

  // 2. Tarik Data Terbaru dari Google Spreadsheet (Pull)
  const btnPullGSheet = document.getElementById('btnPullFromGSheet');
  if (btnPullGSheet) {
    btnPullGSheet.addEventListener('click', async () => {
      showToast('🔄 Menarik data pelanggan terbaru dari Google Spreadsheet...');
      try {
        const res = await fetch('/api/google-sheets/pull', { method: 'POST' });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          showToast(`❌ Gagal tarik data: ${data.error}`);
          return;
        }
        showToast(`✅ Sukses! ${data.count} pelanggan dimuat dari Google Spreadsheet.`);
        fetchDashboard(false);
      } catch (err) {
        showToast(`❌ Terjadi kesalahan: ${err.message}`);
      }
    });
  }

  // 3. Unggah Semua Data ke Google Spreadsheet (Push Backup)
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

  // 4. Modal Panduan & Salin Script
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

  // 5. Tombol Garis 3 (☰) Buka Menu Database & File Excel
  const btnToggleDbMenu = document.getElementById('btnToggleDatabaseMenu');
  const dlgDbMenu = document.getElementById('dialogDatabaseMenu');
  if (btnToggleDbMenu && dlgDbMenu) {
    btnToggleDbMenu.addEventListener('click', () => {
      dlgDbMenu.showModal();
    });
  }

  // 6. Quick Pull langsung dari Toolbar Utama
  const btnQuickPull = document.getElementById('btnQuickPullFromGSheet');
  if (btnQuickPull) {
    btnQuickPull.addEventListener('click', async () => {
      showToast('🔄 Menarik data pelanggan terbaru dari Google Spreadsheet...');
      try {
        const res = await fetch('/api/google-sheets/pull', { method: 'POST' });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          showToast(`❌ Gagal tarik data: ${data.error || 'Periksa URL Google Sheets'}`);
          return;
        }
        showToast(`✅ Sukses! ${data.count} pelanggan dimuat dari Google Spreadsheet.`);
        fetchDashboard(false);
      } catch (err) {
        showToast(`❌ Terjadi kesalahan: ${err.message}`);
      }
    });
  }

  // WhatsApp QR Connect / Disconnect
  document.getElementById('btnConnectWa').addEventListener('click', async () => {
    showToast('📲 Menyiapkan QR Code WhatsApp...');
    await fetch('/api/wa/connect', { method: 'POST' });
    setTimeout(() => fetchDashboard(false), 1500);
  });

  document.getElementById('btnDisconnectWa').addEventListener('click', async () => {
    await fetch('/api/wa/disconnect', { method: 'POST' });
    showToast('🔌 Sesi WhatsApp telah diputuskan.');
    fetchDashboard(false);
  });

  // Password Toggle for Gmail App Password
  const btnTogglePass = document.getElementById('btnToggleSmtpPass');
  const inputPass = document.getElementById('inputSmtpPass');
  btnTogglePass.addEventListener('click', () => {
    const isHidden = inputPass.type === 'password';
    inputPass.type = isHidden ? 'text' : 'password';
    btnTogglePass.setAttribute('aria-pressed', String(isHidden));
    btnTogglePass.textContent = isHidden ? '🙈 Sembunyikan' : '👁️ Lihat';
  });

  // Save WA Gateway Form
  document.getElementById('formWaGateway').addEventListener('submit', async (e) => {
    e.preventDefault();
    const payload = {
      waGateway: {
        enabled: document.getElementById('chkWaGatewayEnabled').checked,
        url: document.getElementById('inputWaGatewayUrl').value.trim(),
        token: document.getElementById('inputWaGatewayToken').value.trim()
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

  // Save SMTP Form
  document.getElementById('formSmtpConfig').addEventListener('submit', async (e) => {
    e.preventDefault();
    const userVal = document.getElementById('inputSmtpUser').value.trim();
    const passVal = document.getElementById('inputSmtpPass').value.trim();
    const autoEnable = Boolean(userVal && passVal) || document.getElementById('chkSmtpEnabled').checked;
    document.getElementById('chkSmtpEnabled').checked = autoEnable;

    const payload = {
      smtp: {
        enabled: autoEnable,
        service: 'gmail',
        fromName: document.getElementById('inputSmtpFromName').value.trim(),
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

  // Save Schedule & Templates Form
  document.getElementById('formScheduleAndTemplates').addEventListener('submit', async (e) => {
    e.preventDefault();
    const payload = {
      operationalStart: document.getElementById('inputOpStart').value,
      operationalEnd: document.getElementById('inputOpEnd').value,
      waWindowStart: document.getElementById('inputWaStart').value,
      waWindowEnd: document.getElementById('inputWaEnd').value,
      samsatAdvanceDays: Number(document.getElementById('inputSamsatDays').value),
      asuransiAdvanceDays: Number(document.getElementById('inputAsuransiDays').value),
      serviceAdvanceDays: Number(document.getElementById('inputServiceDays')?.value || 5),
      delaySecondsBetweenMessages: Number(document.getElementById('inputDelaySec').value),
      autoSendEnabled: document.getElementById('chkAutoSendEnabled').checked,
      templates: {
        birthdayCustomer: document.getElementById('tplBdayCustomer').value,
        birthdayStaff: document.getElementById('tplBdayStaff').value,
        samsatCustomer: document.getElementById('tplSamsatCustomer').value,
        samsatStaff: document.getElementById('tplSamsatStaff').value,
        asuransiCustomer: document.getElementById('tplAsuransiCustomer').value,
        asuransiStaff: document.getElementById('tplAsuransiStaff').value,
        serviceCustomer: document.getElementById('tplServiceCustomer')?.value || '',
        serviceStaff: document.getElementById('tplServiceStaff')?.value || ''
      }
    };
    await fetch('/api/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    showToast('✅ Jadwal WITA (10:00–14:00 WITA) & Template Pesan berhasil disimpan!');
    fetchDashboard(false);
  });
});
