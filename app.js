
// ============================================
// 🔧 CONFIGURATION
// ============================================
const API_URL = 'https://script.google.com/macros/s/AKfycbzEOewMA5fh7azh3LR9iqNY6hqQGeaFjt1BRaNRWvmnwNuQIleWzOyKqFkOZTU8Z4ywEg/exec';

/** ค่าที่เปลี่ยนได้ต่อระบบ — โค้ดโหลดใช้รูปแบบเดียวกัน */
const APP_CONFIG = {
  apiUrl: API_URL,
  loadingMsg: 'กำลังดำเนินการ...',
  loadingTimeoutMs: 15000,
  pollSeatMs: 4000
};

// ============================================
// 🚀 POLYFILL: google.script.run for External Hosting
// ============================================
(function() {
  if (window.google && window.google.script) return;
  window.google = window.google || {};
  window.google.script = window.google.script || {};

  class ExternalRunner {
    constructor() {
      this.successCallback = () => {};
      this.failureCallback = () => {};
    }
    withSuccessHandler(callback) { this.successCallback = callback; return this; }
    withFailureHandler(callback) { this.failureCallback = callback; return this; }

    async callServer(functionName, args) {
      try {
        const response = await fetch(APP_CONFIG.apiUrl, {
          method: 'POST',
          mode: 'cors',
          cache: 'no-cache',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({ action: functionName, args: args }),
        });
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const data = await response.json();
        if (data && data.status === 'success') {
          this.successCallback(data.result);
        } else if (data && data.status === 'error') {
          this.failureCallback(data.message || 'Unknown server error');
        } else {
          this.failureCallback('Invalid server response');
        }
      } catch (error) {
        console.error('API Call Error:', error);
        this.failureCallback(error?.message || 'Unknown network error');
      }
    }
  }

  const runnerProxy = {
    get: function(runner, prop) {
      if (prop === 'withSuccessHandler' || prop === 'withFailureHandler') {
        return function(cb) {
          runner[prop](cb);
          return new Proxy(runner, runnerProxy);
        };
      }
      if (prop in runner) return runner[prop].bind(runner);
      return function(...args) {
        runner.callServer(prop, args);
        return runner;
      };
    }
  };

  window.google.script.run = new Proxy({}, {
    get: function(_, prop) {
      // Each top-level property access creates a fresh ExternalRunner
      const runner = new ExternalRunner();
      if (prop === 'withSuccessHandler' || prop === 'withFailureHandler') {
        return function(cb) {
          runner[prop](cb);
          return new Proxy(runner, runnerProxy);
        };
      }
      return function(...args) {
        runner.callServer(prop, args);
        return runner;
      };
    }
  });
})();

// --- Global Variables ---
let u = null, students = [], subjects = [];
let hwData = [], moneyData = [], leaveData = [], loanData = [];
let lastTheme = localStorage.getItem('theme') || 'light';
let currentHwCount = 0, currentTrCount = 0, currentLvCount = 0;
let currentTrPayCounter = -1;
let currentSeatVersion = -1;
let seatSelectedIds = [];
let currentZoomLevel = 1.0;
let isAutoFitEnabled = true;
let autoRefreshTimer = null;
let loadingCount = 0;
let loadingTimeout = null;
let html5QrCode = null;
let html5QrCodeLogin = null;

const SUBJECTS_FALLBACK = ['ไทยหลัก','ไทยเสริม','คณิตหลัก','คณิตเสริม','วิทย์หลัก','วิทย์เสริม',
'อังกฤษหลัก','อังกฤษเสริม Joshua','อังกฤษเสริม จิรารัตน์','IS','ประวัติ','สังคม',
'ป้องกันการทุจริต','วิทยาการคำนวณ','มัลติมีเดีย','แนะแนว','นาฏศิลป์','ทัศนศิลป์',
'การงาน','สุขศึกษา','พลศึกษา','อื่นๆ'];

document.addEventListener('DOMContentLoaded', () => { 
  initForms(); 
  checkSession(); 
  applyTheme(); 
  window.addEventListener('resize', () => {
    if (seatSnap?.layout && document.getElementById('panelSeats')?.classList.contains('active')) {
      seatRenderAll();
    }
  });
});

function safeAddListener(id, event, handler) {
  const el = document.getElementById(id);
  if (el) el.addEventListener(event, handler);
  else console.warn(`safeAddListener: '${id}' not found`);
}

function getLuminousColor(hex, percent = 40) {
  if (!hex || hex.length < 4) return '#FFFFFF';
  hex = hex.replace(/^#/, '');
  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);

  // ปรับให้สว่างขึ้น (Lighten)
  r = Math.round(r + (255 - r) * (percent / 100));
  g = Math.round(g + (255 - g) * (percent / 100));
  b = Math.round(b + (255 - b) * (percent / 100));

  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

function getDarkColor(hex, amount) {
  if (!hex || hex.length < 4) return '#000000';
  hex = hex.replace(/^#/, '');
  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);
  r = Math.max(0, r + amount);
  g = Math.max(0, g + amount);
  b = Math.max(0, b + amount);
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

function checkChanges() {
  google.script.run
    .withSuccessHandler(r => {
      if (r && r.success) {
        if (r.hwCount !== currentHwCount || r.trCount !== currentTrCount || r.lvCount !== currentLvCount || r.trPayCounter !== currentTrPayCounter) {
          manualRefresh();
        }
        if (r.seatVersion !== currentSeatVersion) {
          currentSeatVersion = r.seatVersion;
          seatsReloadSilent();
        }
      }
    })
    .getCounts();
}

function initForms() {
  safeAddListener('loginForm', 'submit', doLogin);
  safeAddListener('registerForm', 'submit', doRegister);
  safeAddListener('formHw', 'submit', doAddHw);
  safeAddListener('formMoney', 'submit', doAddMoney);
  safeAddListener('formLeave', 'submit', doAddLeave);
  safeAddListener('formPw', 'submit', doChangePw);
  safeAddListener('formLoan', 'submit', doAddLoan);
  
  const today = new Date().toISOString().split('T')[0];
  const hwDate = document.getElementById('hwDate');
  if (hwDate) hwDate.value = today;
  const lDate = document.getElementById('lDate');
  if (lDate) lDate.value = today;
  
  updateCodeForm();
}

// เพิ่มฟังก์ชันนี้ และเปลี่ยน onclick ใน HTML เป็น toggleSidebar()
function toggleSidebar() {
  const desktopSidebar = document.getElementById('sidebar');
  const mobileSidebar = document.getElementById('mobileSidebar');
  const wrapper = document.querySelector('.main-wrapper');
  const overlay = document.getElementById('mobileOverlay');
  const btn = document.getElementById('hamburgerBtn');

  if (window.innerWidth <= 1041) {
    const willOpen = !mobileSidebar?.classList.contains('show');
    if (mobileSidebar) mobileSidebar.classList.toggle('show', willOpen);
    if (overlay) overlay.classList.toggle('show', willOpen);
    if (btn) btn.classList.toggle('is-open', willOpen);
  } else {
    if (desktopSidebar) desktopSidebar.classList.toggle('collapsed');
    if (wrapper) wrapper.classList.toggle('expanded');
  }
}

function closeMobileSidebar() {
  const mobileSidebar = document.getElementById('mobileSidebar');
  const overlay = document.getElementById('mobileOverlay');
  const btn = document.getElementById('hamburgerBtn');
  if (mobileSidebar) mobileSidebar.classList.remove('show');
  if (overlay) overlay.classList.remove('show');
  if (btn) btn.classList.remove('is-open');
}

// --- UI Helpers ---
function openModal(id) { 
  if (id === 'modalHw') populateSubjects(); 
  if (id === 'modalMoney') populateMoneyStudents();
  if (id === 'modalLoan') populateLoanStudents();
  document.getElementById(id).classList.add('show'); 
}
function closeModal(id) { document.getElementById(id).classList.remove('show'); }

function toggleTheme() {
  const newTheme = lastTheme === 'light' ? 'dark' : 'light';
  setTheme(newTheme);
}
function setTheme(mode) {
  if (mode !== 'light' && mode !== 'dark') return;
  lastTheme = mode;
  localStorage.setItem('theme', lastTheme);
  applyTheme();
  updateThemeIcons();
}
function applyTheme() { 
  document.documentElement.setAttribute('data-theme', lastTheme);
}
function updateThemeIcons() {
  const isDark = lastTheme === 'dark';
  const icon = isDark ? 'fa-sun' : 'fa-moon';
  const label = isDark ? 'โหมดสว่าง' : 'โหมดมืด';
  
  const sideBtn = document.getElementById('sidebarThemeBtn');
  if (sideBtn) {
    sideBtn.innerHTML = `<i class="fas ${icon}"></i>`;
    sideBtn.title = label;
  }
  const mobBtn = document.getElementById('mobileThemeBtn');
  if (mobBtn) {
    mobBtn.innerHTML = `<i class="fas ${icon}"></i>`;
    mobBtn.title = label;
  }
}

let globalLoadingToast = null;

function showLoading(msg = APP_CONFIG.loadingMsg) {
  loadingCount = Math.max(0, loadingCount + 1);
  if (loadingCount === 1) {
    const container = document.getElementById('toastContainer');
    globalLoadingToast = document.createElement('div');
    // ใช้คลาส info เพื่อให้เป็นสีฟ้าและมีไอคอนหมุน
    globalLoadingToast.className = 'toast info';
    globalLoadingToast.innerHTML = `
      <i class="fas fa-spinner fa-spin"></i>
      <span class="toast-message">${msg}</span>
    `;
    container.appendChild(globalLoadingToast);
  } else if (globalLoadingToast) {
    globalLoadingToast.querySelector('.toast-message').textContent = msg;
  }
  
  if (loadingTimeout) clearTimeout(loadingTimeout);
  loadingTimeout = setTimeout(() => { console.warn('loading timeout'); resetLoadingOverlay(); }, APP_CONFIG.loadingTimeoutMs);
}

function hideLoading() {
  loadingCount = Math.max(0, loadingCount - 1);
  if (loadingCount <= 0) resetLoadingOverlay();
}

function resetLoadingOverlay() {
  loadingCount = 0;
  if (globalLoadingToast && globalLoadingToast.parentElement) {
    globalLoadingToast.remove();
    globalLoadingToast = null;
  }
  if (loadingTimeout) { clearTimeout(loadingTimeout); loadingTimeout = null; }
}

// Set a button to loading state, returns restore function
function btnLoading(btn) {
  if (!btn) return () => {};
  const orig = btn.innerHTML;
  btn.classList.add('btn-loading');
  btn.innerHTML = orig + '<span class="btn-spin"></span>';
  return () => { btn.classList.remove('btn-loading'); btn.innerHTML = orig; };
}

function showError(id, m) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = m;
  el.style.display = 'block';
}

// Toast — bottom-right, supports 'broadcast' type for cross-user notifications
function showToast(m, type = 'info', duration = 4000) {
  const c = document.getElementById('toastContainer');
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  const icon = type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : type === 'broadcast' ? 'bell' : 'info-circle';
  t.innerHTML = `
    <i class="fas fa-${icon}"></i>
    <span class="toast-message">${m}</span>
    <button class="toast-close" aria-label="ปิด">&times;</button>
  `;
  c.appendChild(t);
  const removeToast = () => { if (t.parentElement) t.remove(); };
  t.querySelector('.toast-close').addEventListener('click', removeToast);
  if (duration > 0) {
    let timer = setTimeout(removeToast, duration);
    t.addEventListener('mouseenter', () => clearTimeout(timer));
    t.addEventListener('mouseleave', () => { timer = setTimeout(removeToast, 2000); });
  }
}

// Show skeleton placeholder while loading data
function showSkeleton(containerId, count = 3) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = Array(count).fill(`<div class="skeleton skeleton-card"></div>`).join('');
}

function formatDate(d) {
  if (!d) return '-';
  try { return new Date(d).toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' }); }
  catch(e) { return d; }
}

// --- Tab Navigation ---
// FIX: setTab now syncs both desktop and mobile sidebar active states
function setTab(id) {
  document.querySelectorAll('.sidebar-item').forEach(item => {
    item.classList.toggle('active', item.dataset.tab === id);
  });
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('panel' + id.charAt(0).toUpperCase() + id.slice(1)).classList.add('active');
  if (window.innerWidth <= 1041) closeMobileSidebar();
  if (id !== 'redeem' && html5QrCode) { 
    try { html5QrCode.stop(); } catch(e) {} 
  }
  if (id === 'schedule') timetableRefresh();
  if (id === 'seats') seatsInit();
  if (id === 'loan') loadLoan();
  if (id === 'receipt') loadReceipts();
}

function escapeHtml(str) {
  return String(str ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function toDatetimeLocal(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
}

// ============================================
// 📷 TIMETABLE (Google Sheet + Drive รูป)
// ============================================
function timetableRefresh() {
  const wrap = document.getElementById('timetablePublicWrap');
  if (!wrap) return;
  google.script.run
    .withSuccessHandler(r => {
      if (!r || !r.success) {
        wrap.innerHTML = `<p style="color:var(--danger);">${escapeHtml(r?.message || 'โหลดไม่สำเร็จ')}</p>`;
        return;
      }
      timetableRenderPublic(r);
      const linkEl = document.getElementById('ttLinkUrl');
      if (linkEl) linkEl.value = r.linkUrl || '';
    })
    .withFailureHandler(() => {
      wrap.innerHTML = '<p style="color:var(--danger);">เชื่อมต่อเซิร์ฟเวอร์ไม่ได้</p>';
    })
    .getTimetable();
}

function timetableRenderPublic(data) {
  const wrap = document.getElementById('timetablePublicWrap');
  if (!wrap) return;
  const link = data.linkUrl
    ? `<p style="margin-top:4px;"><a href="${escapeHtml(data.linkUrl)}" target="_blank" rel="noopener noreferrer" class="btn btn-primary btn-sm" style="display:inline-flex;"><i class="fas fa-external-link-alt"></i> เปิดลิงก์ตารางเรียน</a></p>`
    : '<p style="color:var(--text-muted); margin-top:8px;">ยังไม่มีลิงก์ตารางเรียน (เจ้าของเว็บสามารถตั้งลิงก์ได้จากเมนูตั้งค่า)</p>';
  const img = data.imageUrl
    ? `<img src="${escapeHtml(data.imageUrl)}" style="max-width:100%; border-radius:10px; box-shadow:var(--shadow); margin-top:15px;" alt="ตารางเรียน">`
    : '<p style="color:var(--text-muted); margin-top:10px;">ยังไม่ได้อัปโหลดรูปตารางเรียน</p>';
  wrap.innerHTML = `<div>${link}${img}</div>`;
}

function timetableUploadImage() {
  const inp = document.getElementById('ttImageFile');
  const file = inp?.files?.[0];
  if (!file) { showToast('เลือกไฟล์รูปก่อน', 'error'); return; }
  showLoading('กำลังอัปโหลด...');
  const reader = new FileReader();
  reader.onload = () => {
    google.script.run
      .withSuccessHandler(r => {
        hideLoading();
        inp.value = '';
        if (r.success) {
          showToast('บันทึกรูปแล้ว', 'success');
          timetableRefresh();
        } else showToast(r.message || 'อัปโหลดไม่สำเร็จ', 'error');
      })
      .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
      .setTimetable(u.id, reader.result);
  };
  reader.readAsDataURL(file);
}

function timetableSaveLink() {
  const linkUrl = document.getElementById('ttLinkUrl')?.value?.trim() || '';
  showLoading();
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        showToast('บันทึกลิงก์แล้ว', 'success');
        timetableRefresh();
      } else showToast(r.message || 'บันทึกไม่สำเร็จ', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .setTimetable(u.id, null, linkUrl);
}

// ============================================
// 🪑 SEAT MAP + BOOKING
// ============================================
const SEAT_GUEST_TOKEN_KEY = 'seat_guest_edit_token';
let seatSnap = null;
let seatSelectedId = null;
let seatDragState = null;
let seatScrollState = null;
let seatSuppressClickUntil = 0;

function seatGuestToken() {
  try { return sessionStorage.getItem(SEAT_GUEST_TOKEN_KEY) || ''; } catch { return ''; }
}

function seatCanAdmin() {
  return u && (u.roleKey === 'OWNER' || u.roleKey === 'TEACHER');
}

function seatsInit() {
  seatsLoadSnapshot(true);
  if (seatCanAdmin()) seatRefreshGuestLists();
}

function seatsLoadSnapshot(showToastErr) {
  google.script.run
    .withSuccessHandler(r => {
      if (!r || !r.success) {
        if (showToastErr) showToast(r?.message || 'โหลดแผนผังไม่สำเร็จ', 'error');
        return;
      }
      seatSnap = r;
      currentSeatVersion = r.version;
      
      // Clear selection states
      seatSelectedIds = [];
      seatSelectedId = null;
      
      const fb = document.getElementById('seatFrontBand');
      if (fb && r.layout) fb.value = r.layout.frontBand ?? 2;
      const ws = document.getElementById('seatWinStart');
      const we = document.getElementById('seatWinEnd');
      if (ws) ws.value = toDatetimeLocal(r.bookingStart);
      if (we) we.value = toDatetimeLocal(r.bookingEnd);
      seatRenderAll();
      seatUpdateStatusLine();
    })
    .withFailureHandler(() => { if (showToastErr) showToast('เชื่อมต่อเซิร์ฟเวอร์ไม่ได้', 'error'); })
    .seatGetSnapshot(u.id, seatGuestToken());
}

function seatsReloadSilent() {
  google.script.run
    .withSuccessHandler(r => {
      if (r && r.success) {
        seatSnap = r;
        currentSeatVersion = r.version;
        seatSelectedIds = [];
        seatSelectedId = null;
        seatRenderAll();
        seatUpdateStatusLine();
      }
    })
    .seatGetSnapshot(u.id, seatGuestToken());
}

function seatUpdateStatusLine() {
  const el = document.getElementById('seatStatusLine');
  const actions = document.getElementById('seatHeaderActions');
  if (!el || !seatSnap) return;
  const open = seatSnap.bookingOpen;
  el.innerHTML = open
    ? '<span style="color:var(--success);"><i class="fas fa-circle" style="font-size:0.5rem;"></i> เปิดรับจอง</span>'
    : '<span style="color:var(--warning);"><i class="fas fa-circle" style="font-size:0.5rem;"></i> ปิดจอง / นอกช่วงเวลา</span>';
  if (actions) {
    actions.innerHTML = seatSnap.canManageSettings && seatSnap.version != null
      ? `<span style="font-size:0.85rem;color:var(--text-muted);">แผนผัง v${seatSnap.version}</span>`
      : '';
  }
}

function bookingBySeat(seatId) {
  return (seatSnap.bookings || []).find(b => String(b.seatId) === String(seatId));
}

function seatGetRenderCell() {
  const g = seatSnap?.layout?.grid || { cols: 22, cell: 24 };
  return Number(g.cell) || 24;
}

function seatUpdateSelectedSummary() {
  const box = document.getElementById('seatSelectedSummary');
  if (!box) return;

  const bookBtn = document.getElementById('seatBtnBook');
  const cancelBtn = document.getElementById('seatBtnCancel');
  const codeInput = document.getElementById('seatBookStudentCode');
  const booking = seatSelectedId ? bookingBySeat(seatSelectedId) : null;
  const seat = seatSelectedId ? (seatSnap?.layout?.seats || []).find(s => s.id === seatSelectedId) : null;
  const canBook = !!seat && !booking && !!seatSnap?.bookingOpen;
  const canCancel = !!booking;

  if (!seat) {
    box.innerHTML = `
      <div class="seat-selected-seat">กรุณาเลือกโต๊ะจากแผนผัง</div>
      <div class="seat-selected-meta">แตะหนึ่งครั้งเพื่อเลือก หรือลากเพื่อเลื่อนดูแผนผังห้องเรียน</div>
    `;
  } else if (booking) {
    const name = booking.studentName || 'จองแล้ว';
    const number = booking.studentNo || '-';
    box.innerHTML = `
      <div class="seat-selected-seat">โต๊ะ ${escapeHtml(seat.label || seat.id)}</div>
      <div class="seat-selected-meta">จองโดย ${escapeHtml(name)} | เลขที่ ${escapeHtml(String(number))}</div>
    `;
  } else {
    box.innerHTML = `
      <div class="seat-selected-seat">โต๊ะ ${escapeHtml(seat.label || seat.id)}</div>
      <div class="seat-selected-meta">ว่างสำหรับการจอง กรุณากรอกรหัสนักเรียน 5 หลักเพื่อยืนยัน</div>
    `;
  }

  if (bookBtn) bookBtn.disabled = !canBook;
  if (cancelBtn) cancelBtn.disabled = !canCancel;
  if (codeInput) codeInput.disabled = !!booking;
}

function seatSetupGrabScroll() {
  const scroll = document.getElementById('seatMapScroll');
  if (!scroll || scroll.dataset.grabReady === '1') return;
  scroll.dataset.grabReady = '1';

  const finish = ev => {
    if (!seatScrollState || seatScrollState.pointerId !== ev.pointerId) return;
    if (seatScrollState.moved) seatSuppressClickUntil = Date.now() + 180;
    scroll.classList.remove('drag-scroll');
    try { scroll.releasePointerCapture(ev.pointerId); } catch (_) {}
    seatScrollState = null;
  };

  scroll.addEventListener('pointerdown', ev => {
    if (ev.pointerType === 'mouse' && ev.button !== 0) return;
    if (seatSnap?.canEditLayout && ev.target.closest('.seat-desk')) return;

    seatScrollState = {
      pointerId: ev.pointerId,
      startY: ev.clientY,
      startTop: scroll.scrollTop,
      moved: false
    };
    scroll.classList.add('drag-scroll');
    try { scroll.setPointerCapture(ev.pointerId); } catch (_) {}
  });

  scroll.addEventListener('pointermove', ev => {
    if (!seatScrollState || seatScrollState.pointerId !== ev.pointerId) return;
    const deltaY = ev.clientY - seatScrollState.startY;
    if (Math.abs(deltaY) > 4) seatScrollState.moved = true;
    scroll.scrollTop = seatScrollState.startTop - deltaY;
    if (seatScrollState.moved) ev.preventDefault();
  });

  scroll.addEventListener('pointerup', finish);
  scroll.addEventListener('pointercancel', finish);
  scroll.addEventListener('pointerleave', ev => {
    if (seatScrollState && seatScrollState.pointerId === ev.pointerId && ev.pointerType === 'mouse') finish(ev);
  });
  scroll.addEventListener('click', ev => {
    if (Date.now() < seatSuppressClickUntil) {
      ev.preventDefault();
      ev.stopPropagation();
    }
  }, true);
}

function seatRenderAll() {
  const inner = document.getElementById('seatMapInner');
  if (!inner || !seatSnap || !seatSnap.layout) return;
  const g = seatSnap.layout.grid || { cols: 22, rows: 16, cell: 24 };
  const cell = seatGetRenderCell();
  const seats = seatSnap.layout.seats || [];
  
  inner.style.setProperty('--seat-cell', String(cell));
  inner.style.width = `${g.cols * cell}px`;
  inner.style.height = `${g.rows * cell}px`;
  inner.style.minWidth = '0';
  inner.style.minHeight = `${Math.max(g.rows * cell, 250)}px`;
  inner.style.backgroundSize = `${cell}px ${cell}px`;

  const fb = Number(seatSnap.layout.frontBand) || 0;
  let html = `
    <div class="seat-stage-arc"></div>
    <div class="seat-front-band" style="height:${fb * cell}px;"></div>
  `;

  const canEdit = seatSnap.canEditLayout;
  if (seats.length === 0) {
    html = '<div style="padding:40px; text-align:center; color:var(--text-muted); width:100%;">ไม่มีที่นั่งในแผนผังนี้</div>';
  } else {
    seats.forEach(st => {
      const bk = bookingBySeat(st.id);
      const bookedClass = bk ? ' booked' : '';
      const lockedClass = st.lock ? ' locked' : '';
      
      const isSelectedForEdit = canEdit && seatSelectedIds.includes(st.id);
      const isSelectedForBook = !canEdit && seatSelectedId === st.id;
      const selectedClass = (isSelectedForEdit || isSelectedForBook) ? ' selected' : '';
      
      let label = `<div class="seat-desk-label">${escapeHtml(st.label || 'Seat')}</div>`;
      if (bk) {
        const nick = bk.studentName ? bk.studentName.split(' ')[0] : '-';
        label = `<div class="seat-desk-name">${escapeHtml(nick)}</div>
                 <div class="seat-desk-no">No. ${escapeHtml(String(bk.studentNo || '-'))}</div>`;
      }

      html += `<div class="seat-desk${bookedClass}${lockedClass}${selectedClass}" data-seat="${escapeHtml(st.id)}"
        style="left:${st.gx * cell}px;top:${st.gy * cell}px;width:${(st.gw || 3) * cell}px;height:${(st.gh || 2) * cell}px;"
        title="${bk ? escapeHtml(bk.studentName || '') : ''}">${label}</div>`;
    });
  }

  inner.innerHTML = html;
  seatSetupGrabScroll();
  seatUpdateSelectedSummary();

  inner.querySelectorAll('.seat-desk').forEach(el => {
    el.addEventListener('click', () => seatOnSelect(el.dataset.seat));
    if (canEdit) seatBindDrag(el, cell, fb);
  });
  
  // Apply zoom/fit scaling
  seatApplyZoomAndFit();
  
  // Update bulk selection toolbar details
  if (canEdit) {
    const selCount = seatSelectedIds.length;
    const countText = document.getElementById('seatSelCountText');
    if (countText) countText.textContent = `เลือกแล้ว ${selCount} โต๊ะ`;
    
    const btnAlignTop = document.getElementById('btnAlignTop');
    const btnAlignLeft = document.getElementById('btnAlignLeft');
    const btnDeleteSelected = document.getElementById('btnDeleteSelected');
    
    if (btnAlignTop) btnAlignTop.disabled = selCount < 2;
    if (btnAlignLeft) btnAlignLeft.disabled = selCount < 2;
    if (btnDeleteSelected) btnDeleteSelected.disabled = selCount === 0;
  }
}

function seatOnSelect(id) {
  if (Date.now() < seatSuppressClickUntil) return;
  
  if (seatSnap?.canEditLayout) {
    const idx = seatSelectedIds.indexOf(id);
    if (idx > -1) {
      seatSelectedIds.splice(idx, 1);
    } else {
      seatSelectedIds.push(id);
    }
  } else {
    seatSelectedId = id;
  }
  seatRenderAll();
}

function seatBindDrag(el, cell, fb) {
  el.addEventListener('pointerdown', ev => {
    const id = el.dataset.seat;
    const st = seatSnap.layout.seats.find(s => s.id === id);
    if (!st || st.lock || st.gy < fb) return;
    ev.preventDefault();
    el.classList.add('dragging');
    
    // If the dragged desk is NOT part of the selected desks, make it the ONLY selected desk
    if (!seatSelectedIds.includes(id)) {
      seatSelectedIds = [id];
      seatRenderAll();
    }
    
    // Store original position of ALL selected desks
    const draggedSeats = seatSnap.layout.seats.filter(s => seatSelectedIds.includes(s.id) && !s.lock && s.gy >= fb);
    const seatPositions = draggedSeats.map(s => ({
      id: s.id,
      ogx: s.gx,
      ogy: s.gy,
      gw: s.gw || 3,
      gh: s.gh || 2,
      el: document.querySelector(`[data-seat="${s.id}"]`)
    }));
    
    seatDragState = {
      draggedId: id,
      startX: ev.clientX,
      startY: ev.clientY,
      cell,
      fb,
      positions: seatPositions
    };
    
    el.setPointerCapture(ev.pointerId);
  });

  el.addEventListener('pointermove', ev => {
    if (!seatDragState || seatDragState.draggedId !== el.dataset.seat) return;
    const dx = ev.clientX - seatDragState.startX;
    const dy = ev.clientY - seatDragState.startY;
    const cell = seatDragState.cell;
    
    // Calculate grid step delta
    const dgx = Math.round(dx / cell);
    const dgy = Math.round(dy / cell);
    
    const g = seatSnap.layout.grid;
    const cols = g.cols || 22;
    const rows = g.rows || 16;
    
    // Move ALL desks in the selection by the delta
    seatDragState.positions.forEach(pos => {
      const st = seatSnap.layout.seats.find(s => s.id === pos.id);
      if (!st) return;
      
      st.gx = Math.max(0, Math.min(cols - pos.gw, pos.ogx + dgx));
      st.gy = Math.max(seatDragState.fb, Math.min(rows - pos.gh, pos.ogy + dgy));
      
      // Update DOM style elements directly for smooth dragging
      if (pos.el) {
        pos.el.style.left = `${st.gx * cell}px`;
        pos.el.style.top = `${st.gy * cell}px`;
      }
    });
  });

  el.addEventListener('pointerup', ev => {
    if (seatDragState && seatDragState.draggedId === el.dataset.seat) {
      el.classList.remove('dragging');
      seatDragState = null;
      try { el.releasePointerCapture(ev.pointerId); } catch (_) {}
      seatRenderAll(); // final render to snap positions perfectly
    }
  });
}

// --- Zoom & Fit Utilities ---
function seatApplyZoomAndFit() {
  const scroll = document.getElementById('seatMapScroll');
  const outer = document.getElementById('seatScaleOuter');
  const wrapper = document.getElementById('seatMapScaleWrapper');
  const inner = document.getElementById('seatMapInner');
  
  if (!scroll || !outer || !wrapper || !inner || !seatSnap || !seatSnap.layout) return;
  
  const g = seatSnap.layout.grid || { cols: 22, rows: 16, cell: 24 };
  const cell = Number(g.cell) || 24;
  const W = g.cols * cell;
  const H = g.rows * cell;
  
  let S = currentZoomLevel;
  
  // Update checkbox state
  const fitCheck = document.getElementById('seatAutoFitCheck');
  if (fitCheck) {
    isAutoFitEnabled = fitCheck.checked;
  }
  
  if (isAutoFitEnabled) {
    const padX = 32; // padding of seatMapScroll
    const availW = scroll.clientWidth - padX;
    S = Math.min(1.2, availW / W);
    // Update zoom slider and percentage text
    const slider = document.getElementById('seatZoomSlider');
    if (slider) slider.value = S.toFixed(2);
  }
  
  const pctText = document.getElementById('seatZoomPercent');
  if (pctText) pctText.textContent = `${Math.round(S * 100)}%`;
  
  // Set physical sizes of outer container to match the scaled dimensions
  outer.style.width = `${W * S}px`;
  outer.style.height = `${Math.max(H * S, 250 * S)}px`;
  
  // Set wrapper constraints
  wrapper.style.transform = `scale(${S})`;
  wrapper.style.width = `${W}px`;
  wrapper.style.height = `${Math.max(H, 250)}px`;
}

function seatSetZoomFromSlider(val) {
  const fitCheck = document.getElementById('seatAutoFitCheck');
  if (fitCheck) fitCheck.checked = false; // Turn off Auto-fit when user manually interacts with slider
  isAutoFitEnabled = false;
  currentZoomLevel = Number(val);
  seatApplyZoomAndFit();
}

function seatAdjustZoom(delta) {
  const fitCheck = document.getElementById('seatAutoFitCheck');
  if (fitCheck) fitCheck.checked = false;
  isAutoFitEnabled = false;
  
  currentZoomLevel = Math.max(0.4, Math.min(1.5, currentZoomLevel + delta));
  const slider = document.getElementById('seatZoomSlider');
  if (slider) slider.value = currentZoomLevel.toFixed(2);
  
  seatApplyZoomAndFit();
}

function seatToggleAutoFit(checked) {
  isAutoFitEnabled = checked;
  seatApplyZoomAndFit();
}

// --- Bulk Desk Utilities ---
function seatSelectAll() {
  if (!seatSnap?.layout?.seats) return;
  seatSelectedIds = seatSnap.layout.seats.map(s => s.id);
  seatRenderAll();
}

function seatClearSelection() {
  seatSelectedIds = [];
  seatRenderAll();
}

function seatAlignSelected(direction) {
  if (seatSelectedIds.length < 2) return;
  const seats = seatSnap.layout.seats.filter(s => seatSelectedIds.includes(s.id));
  if (direction === 'top') {
    // Find top-most Y (minimum gy)
    let minY = Infinity;
    seats.forEach(s => minY = Math.min(minY, s.gy));
    seats.forEach(s => s.gy = minY);
    showToast('จัดระดับแนวตั้งตรงกันแล้ว — อย่าลืมกดบันทึกแผนผัง', 'info');
  } else if (direction === 'left') {
    // Find left-most X (minimum gx)
    let minX = Infinity;
    seats.forEach(s => minX = Math.min(minX, s.gx));
    seats.forEach(s => s.gx = minX);
    showToast('จัดระดับแนวนอนตรงกันแล้ว — อย่าลืมกดบันทึกแผนผัง', 'info');
  }
  seatRenderAll();
}

function seatDeleteSelected() {
  if (seatSelectedIds.length === 0) return;
  if (!confirm(`ต้องการลบโต๊ะที่เลือกทั้งหมด ${seatSelectedIds.length} โต๊ะหรือไม่?`)) return;
  
  // Remove from layout
  seatSnap.layout.seats = seatSnap.layout.seats.filter(s => !seatSelectedIds.includes(s.id));
  
  // Clear selection
  seatSelectedIds = [];
  seatRenderAll();
  showToast('ลบโต๊ะที่เลือกเรียบร้อยแล้ว — อย่าลืมกดบันทึกแผนผัง', 'warning');
}

function seatApplyTemplate(templateId) {
  if (!seatSnap?.canEditLayout) return;
  const val = templateId;
  if (!val) { showToast('', 'error'); return; }

  // :  2 cells,  2 cells
  // : 1 cell, : 2 cells
  // gw=2, gh=2  
  const positions = [];
  let cols = 22, rows = 16;

  if (val === 'temp1') {
    //  1: 4   2 , 4  = 32 
    //  x-start: 1, 6, 11, 16  ( 2   2+1+2=5,  3)
    const groupXs = [1, 6, 12, 17]; // x 
    const colsInGroup = [0, 3];     // offset  (2   3)
    const rowYs = [1, 4, 7, 10];    // 4 
    cols = 22; rows = 14;
    rowYs.forEach(y => {
      groupXs.forEach(gx => {
        colsInGroup.forEach(dx => {
          positions.push({ gx: gx + dx, gy: y, gw: 2, gh: 2 });
        });
      });
    });

  } else if (val === 'temp2') {
    //  2:  1  + 3  2 , 5  = 35 
    const rowYs = [1, 4, 7, 10, 13];
    cols = 20; rows = 17;
    rowYs.forEach(y => {
      //  1: 1  (x=1)
      positions.push({ gx: 1, gy: y, gw: 2, gh: 2 });
      //  2-4: 2  (x=5,8 / 12,15 / 19,22  )
      [[5,8],[11,14],[17,20]].forEach(([x1,x2]) => {
        positions.push({ gx: x1, gy: y, gw: 2, gh: 2 });
        positions.push({ gx: x2, gy: y, gw: 2, gh: 2 });
      });
    });

  } else if (val === 'temp3') {
    //  3: 5   2 , 4  = 40 
    const groupXs = [1, 6, 12, 17, 23];
    const colsInGroup = [0, 3];
    const rowYs = [1, 4, 7, 10];
    cols = 28; rows = 14;
    rowYs.forEach(y => {
      groupXs.forEach(gx => {
        colsInGroup.forEach(dx => {
          positions.push({ gx: gx + dx, gy: y, gw: 2, gh: 2 });
        });
      });
    });
  }

  const confirmMsg = `การจัดที่นั่งตามแม่แบบจะปรับขนาดตารางเป็น ${cols}x${rows} และจัดวางตำแหน่งโต๊ะใหม่ทั้งหมด ${positions.length} โต๊ะ\n\nต้องการดำเนินการต่อหรือไม่?`;
  if (!confirm(confirmMsg)) return;

  // Set grid dimensions
  seatSnap.layout.grid.cols = cols;
  seatSnap.layout.grid.rows = rows;
  
  const existingSeats = [...(seatSnap.layout.seats || [])];
  const newSeats = [];
  
  for (let i = 0; i < positions.length; i++) {
    const pos = positions[i];
    if (i < existingSeats.length) {
      // Reuse existing seat
      const seat = existingSeats[i];
      seat.gx = pos.gx;
      seat.gy = pos.gy;
      seat.gw = pos.gw;
      seat.gh = pos.gh;
      newSeats.push(seat);
    } else {
      // Create new seat
      newSeats.push({
        id: 'S_' + Math.random().toString(36).slice(2, 10),
        gx: pos.gx,
        gy: pos.gy,
        gw: pos.gw,
        gh: pos.gh,
        label: String(i + 1),
        lock: false
      });
    }
  }
  
  // If there are more existing seats than template capacity, keep booked seats at the bottom
  if (existingSeats.length > positions.length) {
    const extraSeats = existingSeats.slice(positions.length);
    let overflowY = rows;
    extraSeats.forEach(seat => {
      const bk = bookingBySeat(seat.id);
      if (bk) {
        seat.gx = 0;
        seat.gy = overflowY;
        seat.gw = 3;
        seat.gh = 2;
        newSeats.push(seat);
        overflowY += 3;
        if (overflowY >= seatSnap.layout.grid.rows) {
          seatSnap.layout.grid.rows = overflowY + 2;
        }
      }
    });
  }

  seatSnap.layout.seats = newSeats;
  seatSelectedIds = [];
  seatRenderAll();
  showToast('จัดวางโต๊ะตามแม่แบบเรียบร้อยแล้ว — อย่าลืมกดบันทึกแผนผัง', 'success', 5000);
}

function seatSaveLayoutNow() {
  if (!seatSnap?.canEditLayout) { showToast('ไม่มีสิทธิ์บันทึกแผนผัง', 'error'); return; }
  showLoading('กำลังบันทึกแผนผัง...');
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        showToast('บันทึกแผนผังแล้ว', 'success');
        seatsLoadSnapshot(false);
      } else showToast(r.message || 'บันทึกไม่สำเร็จ', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .seatSaveLayout(u.id, JSON.stringify(seatSnap.layout), seatGuestToken());
}

function seatSaveBookingWindow() {
  if (!seatCanAdmin()) return;
  const ws = document.getElementById('seatWinStart')?.value;
  const we = document.getElementById('seatWinEnd')?.value;
  showLoading();
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        showToast('บันทึกช่วงเวลาแล้ว', 'success');
        seatsLoadSnapshot(false);
      } else showToast(r.message || 'บันทึกไม่สำเร็จ', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .seatSetBookingWindow(u.id, ws ? new Date(ws).toISOString() : '', we ? new Date(we).toISOString() : '');
}

function seatAddDesk() {
  if (!seatSnap?.canEditLayout) return;
  const g = seatSnap.layout.grid;
  const cols = g.cols || 22;
  const rows = g.rows || 16;
  const gw = 3;
  const gh = 2;
  const gx = Math.max(0, Math.floor(cols / 2 - gw / 2));
  const gy = Math.max(seatSnap.layout.frontBand || 2, Math.floor(rows / 2));
  seatSnap.layout.seats.push({
    id: 'S_' + Math.random().toString(36).slice(2, 10),
    gx,
    gy,
    gw,
    gh,
    label: String((seatSnap.layout.seats || []).length + 1),
    lock: false
  });
  seatRenderAll();
}

function seatCenterDesks() {
  if (!seatSnap?.canEditLayout) return;
  const seats = seatSnap.layout.seats || [];
  if (!seats.length) return;
  const g = seatSnap.layout.grid;
  const cols = g.cols || 22;
  const rows = g.rows || 16;
  let minGx = Infinity, maxGx = -Infinity, minGy = Infinity, maxGy = -Infinity;
  seats.forEach(s => {
    const gw = s.gw || 3;
    const gh = s.gh || 2;
    minGx = Math.min(minGx, s.gx);
    maxGx = Math.max(maxGx, s.gx + gw);
    minGy = Math.min(minGy, s.gy);
    maxGy = Math.max(maxGy, s.gy + gh);
  });
  const bw = maxGx - minGx;
  const bh = maxGy - minGy;
  const dx = Math.floor((cols - bw) / 2) - minGx;
  const dy = Math.floor((rows - bh) / 2) - minGy;
  seats.forEach(s => {
    s.gx = Math.max(0, Math.min(cols - (s.gw || 3), s.gx + dx));
    s.gy = Math.max(seatSnap.layout.frontBand || 0, Math.min(rows - (s.gh || 2), s.gy + dy));
  });
  seatRenderAll();
  showToast('จัดกึ่งกลางแล้ว (ยังไม่บันทึกจนกดบันทึกแผนผัง)', 'info', 3500);
}

function seatSaveLayoutNow() {
  if (!seatSnap || !seatSnap.layout) return;
  const fb = Number(document.getElementById('seatFrontBand')?.value) || seatSnap.layout.frontBand || 2;
  seatSnap.layout.frontBand = fb;
  
  showLoading('กำลังบันทึกแผนผัง...');
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        showToast('บันทึกแผนผังสำเร็จ!', 'success');
        seatsLoadSnapshot(false);
      } else showToast(r.message || 'บันทึกไม่สำเร็จ', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .seatSaveLayout(u.id, seatSnap.layout, seatGuestToken());
}

function seatToggleLockFrontRow() {
  if (!seatSnap?.canEditLayout) return;
  const fb = Number(document.getElementById('seatFrontBand')?.value) || seatSnap.layout.frontBand || 2;
  seatSnap.layout.frontBand = fb;
  const seats = seatSnap.layout.seats || [];
  const lockNext = !seats.some(s => s.gy < fb && s.lock);
  seats.forEach(s => {
    if (s.gy < fb) s.lock = lockNext;
  });
  seatRenderAll();
  showToast(lockNext ? 'ล็อกโซนหน้าห้องแล้ว — อย่าลืมกดบันทึกแผนผัง' : 'ปลดล็อกโซนหน้าห้องแล้ว — อย่าลืมกดบันทึกแผนผัง', 'info', 4000);
}

function seatDoBook() {
  const code = document.getElementById('seatBookStudentCode')?.value?.trim();
  if (!seatSelectedId) { showToast('เลือกที่นั่งบนแผนผังก่อน', 'error'); return; }
  if (!/^\d{5}$/.test(code)) { showToast('กรุณากรอกรหัสนักเรียน 5 หลัก', 'error'); return; }
  showLoading('กำลังจอง...');
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        showToast(r.message || 'จองสำเร็จ', 'success');
        seatsLoadSnapshot(false);
      } else showToast(r.message || 'จองไม่ได้', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .seatBook(u.id, seatSelectedId, code);
}

function seatDoCancel() {
  if (!seatSelectedId) { showToast('เลือกที่นั่งที่จองแล้วก่อน', 'error'); return; }
  showLoading();
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        showToast(r.message || 'ยกเลิกแล้ว', 'success');
        seatsLoadSnapshot(false);
      } else showToast(r.message || 'ยกเลิกไม่ได้', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .seatCancelBooking(u.id, seatSelectedId);
}

function seatCreateGuestCode() {
  let plain = document.getElementById('seatGuestPlainCode')?.value?.trim();
  const mins = Number(document.getElementById('seatGuestMins')?.value) || 60;
  const label = document.getElementById('seatGuestLabel')?.value?.trim() || '';
  
  if (!plain) {
    plain = Math.random().toString(36).substring(2, 8).toUpperCase();
    showToast('กำลังสร้างโค้ดอัตโนมัติ...', 'info');
  }
  
  if (plain.length < 4) { showToast('ตั้งโค้ดอย่างน้อย 4 ตัว', 'error'); return; }
  showLoading();
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        showToast('สร้างโค้ดแล้ว — ส่งโค้ดให้ผู้ช่วยได้ (ไม่แสดงซ้ำเพื่อความปลอดภัย)', 'success', 5000);
        document.getElementById('seatGuestPlainCode').value = '';
        seatRefreshGuestLists();
      } else showToast(r.message || 'สร้างไม่สำเร็จ', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .seatCreateEditCode(u.id, plain, mins, label);
}

function seatApplyGuestCode() {
  const c = document.getElementById('seatGuestInputCode')?.value?.trim();
  if (!c) return;
  showLoading();
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success && r.token) {
        sessionStorage.setItem(SEAT_GUEST_TOKEN_KEY, r.token);
        showToast('เปิดโหมดแก้แผนผังแล้ว', 'success');
        seatsLoadSnapshot(false);
      } else showToast(r.message || 'โค้ดไม่ถูกต้อง', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อไม่ได้', 'error'); })
    .seatValidateEditCode(c);
}

function seatClearGuestToken() {
  sessionStorage.removeItem(SEAT_GUEST_TOKEN_KEY);
  showToast('ออกจากโหมดแก้ไขแล้ว', 'info');
  seatsLoadSnapshot(false);
}

function seatRefreshGuestLists() {
  google.script.run
    .withSuccessHandler(r => {
      const box = document.getElementById('seatGuestCodesList');
      if (!box || !r.success) return;
      box.innerHTML = '<b>โค้ดที่สร้าง</b><ul style="margin:6px 0 0 18px;">' +
        (r.codes || []).map(c => `<li>${escapeHtml(c.label || '')} — ${c.active ? 'ใช้งานได้ถึง' : 'หมดอายุ/ยกเลิก'} ${escapeHtml(c.expiresAt || '')}
          ${c.active ? `<button type="button" class="btn btn-danger btn-sm" onclick="seatRevokeCode('${escapeHtml(c.id)}')">ยกเลิกโค้ด</button>` : ''}</li>`).join('') +
        '</ul>';
    })
    .seatListEditCodes(u.id);
  google.script.run
    .withSuccessHandler(r => {
      const box = document.getElementById('seatGuestSessionsList');
      if (!box || !r.success) return;
      box.innerHTML = '<b>เซสชันที่กำลังแก้ไข</b><ul style="margin:6px 0 0 18px;">' +
        (r.sessions || []).map(s => `<li>หมดอายุ ${escapeHtml(s.expiresAt)}
          <button type="button" class="btn btn-danger btn-sm" onclick="seatKickSession('${escapeHtml(s.token)}')">ตัดการเชื่อมต่อ</button></li>`).join('') +
        '</ul>';
    })
    .seatListActiveSessions(u.id);
}

function seatRevokeCode(codeId) {
  google.script.run
    .withSuccessHandler(r => {
      if (r.success) { showToast('ยกเลิกโค้ดแล้ว', 'success'); seatRefreshGuestLists(); }
      else showToast(r.message || 'ล้มเหลว', 'error');
    })
    .seatRevokeEditCode(u.id, codeId);
}

function seatKickSession(token) {
  google.script.run
    .withSuccessHandler(r => {
      if (r.success) { showToast('ตัดการเชื่อมต่อแล้ว', 'success'); seatRefreshGuestLists(); }
      else showToast(r.message || 'ล้มเหลว', 'error');
    })
    .seatRevokeSession(u.id, token);
}

// --- Auth ---
function doLogin(e) {
  e.preventDefault(); 
  const id = document.getElementById('loginId').value.trim(); 
  const pw = document.getElementById('loginPw').value; 
  const loginCode = document.getElementById('loginCode')?.value?.trim().toUpperCase();
  showLoading();
  
  google.script.run
    .withSuccessHandler(r => { 
      hideLoading(); 
      if (r.success) { 
        u = r.user; 
        localStorage.setItem('user', JSON.stringify(u)); 
        showApp();
        // FIX: Pass u.id when redeeming code after login
        if (loginCode) {
          showLoading();
          google.script.run
            .withSuccessHandler(rr => {
              hideLoading();
              showToast(rr.success ? rr.message : rr.message, rr.success ? 'success' : 'error');
              if (rr.success) manualRefresh();
            })
            .withFailureHandler(() => { hideLoading(); showToast('Redeem error', 'error'); })
            .redeemCode(loginCode, u.id);
        }
      } else {
        showError('loginError', r.message); 
      }
    })
    .withFailureHandler(err => { 
      hideLoading(); 
      showError('loginError', 'เชื่อมต่อ Server ไม่ได้'); 
    })
    .loginUser(id, pw);
}

function doRegister(e) {
  e.preventDefault(); 
  const n = document.getElementById('regName').value.trim(); 
  const em = document.getElementById('regEmail').value.trim(); 
  const c = document.getElementById('regCode').value.trim(); 
  const p = document.getElementById('regPw').value; 
  const cp = document.getElementById('regCPw').value; 
  const h = document.getElementById('regHint').value; 
  showLoading();
  
  google.script.run
    .withSuccessHandler(r => { 
      hideLoading(); 
      if (r.success) { 
        showToast(r.message, 'success'); 
        toggleForm('login'); 
      } else {
        showError('regError', r.message); 
      }
    })
    .withFailureHandler(err => { hideLoading(); showError('regError', 'Server error (register)'); })
    .registerUser(n, em, p, cp, c, h);
}

function checkSession() { 
  resetLoadingOverlay();
  const saved = localStorage.getItem('user'); 
  if (saved) { 
    try {
      u = JSON.parse(saved);
      // Guest account: verify not expired before restoring session
      if (u.isGuest) {
        const expiresAt = localStorage.getItem('guestExpiresAt');
        if (!expiresAt || new Date(expiresAt).getTime() <= Date.now()) {
          // Expired — delete on server then clear
          if (u.id) {
            google.script.run
              .withSuccessHandler(() => {})
              .withFailureHandler(() => {})
              .deleteGuestAccount(u.id);
          }
          localStorage.clear();
          return;
        }
      }
      showApp();
    } catch(e) {
      localStorage.clear();
    }
  } 
}

function logout() { 
  if (autoRefreshTimer) clearInterval(autoRefreshTimer);
  if (guestCountdownInterval) clearInterval(guestCountdownInterval);
  // If guest, delete account on server before clearing
  if (u && u.isGuest && u.id) {
    google.script.run
      .withSuccessHandler(() => {})
      .withFailureHandler(() => {})
      .deleteGuestAccount(u.id);
  }
  localStorage.clear(); 
  location.reload(); 
}

function showApp() {
  document.getElementById('loginPage').style.display = 'none'; 
  document.getElementById('app').style.display = 'block';
  
  document.getElementById('sidebarUserName').textContent = u.displayName;
  document.getElementById('sidebarUserRole').textContent = u.role;
  document.getElementById('mobileSidebarUserName').textContent = u.displayName;
  document.getElementById('mobileSidebarUserRole').textContent = u.role;
  document.getElementById('mobileSidebarCredits').textContent = u.hwCredits || 0;
  
  if (u.canManageHomework || u.hwCredits > 0) document.getElementById('btnAddHw').style.display = 'inline-flex';
  if (u.canManageTreasury) document.getElementById('btnAddMoney').style.display = 'inline-flex';
  if (u.canManageCodes) {
    document.getElementById('createCodeCard').style.display = 'block';
    updateCodeForm();
  }

  if (u.roleKey === 'OWNER') {
    const ttCard = document.getElementById('timetableOwnerCard');
    if (ttCard) ttCard.style.display = 'block';
  }

  const canSeatAdmin = u.roleKey === 'OWNER' || u.roleKey === 'TEACHER';
  const seatAd = document.getElementById('seatAdminCard');
  const seatGuest = document.getElementById('seatGuestLoginCard');
  if (seatAd) seatAd.style.display = canSeatAdmin ? 'block' : 'none';
  if (seatGuest) seatGuest.style.display = canSeatAdmin ? 'none' : 'block';

  // --- Guest-specific UI ---
  if (u.isGuest) {
    // Show timer cards
    document.getElementById('guestTimerCard').style.display = 'block';
    document.getElementById('sidebarTimer').classList.add('show');
    document.getElementById('mobileSidebarTimer').classList.add('show');
    // Hide password change (guest has no password)
    document.getElementById('changePwCard').style.display = 'none';
    document.getElementById('hwCreditsRow').style.display = 'none';
    // Restore countdown if refreshed
    restoreGuestCountdownIfNeeded();
  }
  
  renderMe();
  loadMasters();
  startAutoRefresh();
  resetLoadingOverlay();
  updateNotifBadges();
  updateNotifBtnState();

  timetableRefresh();
  seatsLoadSnapshot(false);
}

function togglePushNotifications() {
  if (!("Notification" in window)) {
    showToast('เบราว์เซอร์นี้ไม่รองรับการแจ้งเตือน', 'error');
    return;
  }
  
  if (Notification.permission === "granted") {
    showToast('คุณได้เปิดการแจ้งเตือนเรียบร้อยแล้ว', 'success');
  } else if (Notification.permission !== "denied") {
    Notification.requestPermission().then(permission => {
      if (permission === "granted") {
        showToast('เปิดการแจ้งเตือนสำเร็จ!', 'success');
        updateNotifBtnState();
      }
    });
  } else {
    showToast('คุณได้ปิดการแจ้งเตือนไว้ กรุณาเปิดในตั้งค่าเบราว์เซอร์', 'warning');
  }
}

function updateNotifBtnState() {
  const btn = document.getElementById('btnPushNotif');
  if (!btn) return;
  if (!("Notification" in window)) {
    btn.style.display = 'none';
    return;
  }
  if (Notification.permission === "granted") {
    btn.innerHTML = '<i class="fas fa-check-circle"></i> การแจ้งเตือนเปิดอยู่';
    btn.classList.replace('btn-secondary', 'btn-success');
    btn.disabled = true;
  } else if (Notification.permission === "denied") {
    btn.innerHTML = '<i class="fas fa-exclamation-triangle"></i> การแจ้งเตือนถูกบล็อก';
    btn.classList.replace('btn-secondary', 'btn-danger');
  }
}

// --- Load Masters ---
function loadMasters() {
  google.script.run
    .withSuccessHandler(r => {
      students = r.students || [];
      subjects = r.subjects || SUBJECTS_FALLBACK;
      populateSubjects();
      loadAll();
    })
    .withFailureHandler(err => {
      console.error('loadMasters failed:', err);
      subjects = SUBJECTS_FALLBACK;
      populateSubjects();
      loadAll();
    })
    .getStudents();
}

function populateMoneyStudents() {
  const container = document.getElementById('mStudentSelection');
  if (!container) return;
  container.innerHTML = '';
  
  students.forEach(s => {
    const div = document.createElement('div');
    div.className = 'template-item'; // reuse template-item style for selection
    div.style.padding = '5px';
    div.style.fontSize = '0.75rem';
    div.dataset.no = s.no;
    div.textContent = `${s.no}. ${s.name.split(' ')[0]}`;
    div.classList.add('active'); // Default to selected
    div.onclick = () => {
      div.classList.toggle('active');
      updateMTargetList();
    };
    container.appendChild(div);
  });
  updateMTargetList();
}

function updateMTargetList() {
  const container = document.getElementById('mStudentSelection');
  const input = document.getElementById('mTargetList');
  if (!container || !input) return;
  const selected = Array.from(container.querySelectorAll('.template-item.active'))
                        .map(el => el.dataset.no);
  input.value = selected.join(',');
}

function mSelectAll() {
  document.querySelectorAll('#mStudentSelection .template-item').forEach(el => el.classList.add('active'));
  updateMTargetList();
}

function mClearAll() {
  document.querySelectorAll('#mStudentSelection .template-item').forEach(el => el.classList.remove('active'));
  updateMTargetList();
}

function populateSubjects() {
  const sel = document.getElementById('hwSub'); 
  if (!sel) return;
  sel.innerHTML = '<option value="">-- เลือกวิชา --</option>';
  (subjects && subjects.length ? subjects : SUBJECTS_FALLBACK).forEach(s => {
    sel.innerHTML += `<option value="${s}">${s}</option>`;
  });
}

// --- Main Loaders ---
// Silent load — no overlay, just skeleton then render
function loadAll() {
  const elHw = document.getElementById('listHw');
  const elTr = document.getElementById('listMoney');
  const elLv = document.getElementById('listLeave');
  
  if (!hwData.length) showSkeleton('listHw', 2);
  if (!moneyData.length) showSkeleton('listMoney', 2);
  if (!leaveData.length) showSkeleton('listLeave', 2);

  google.script.run
    .withSuccessHandler(r => {
      if (r && r.success) {
        hwData = r.homework || [];
        moneyData = r.treasury || [];
        leaveData = r.leaveRequests || [];
        
        renderHw();
        renderMoney();
        renderLeave();
        
        currentHwCount = hwData.length;
        currentTrCount = moneyData.length;
        currentLvCount = leaveData.length;

        updateNotifBadges();
      }
      timetableRefresh();
    })
    .withFailureHandler(err => {
      console.error('getDashboardData failed:', err);
    })
    .getDashboardData();
}

function updateNotifBadges() {
  const pendingHw = hwData.filter(h => !h.noDueDate && new Date(h.dueDate) >= new Date()).length;
  const pendingMoney = moneyData.filter(m => u.studentNo && m.payments[u.studentNo] && !m.payments[u.studentNo].isComplete).length;
  const pendingLeave = u.canApproveLeave ? leaveData.filter(l => l.status === 'PENDING').length : 0;

  const bHw = document.getElementById('notifHw');
  const bMoney = document.getElementById('notifMoney');
  const bLeave = document.getElementById('notifLeave');

  if (bHw) { bHw.textContent = pendingHw; bHw.style.display = pendingHw > 0 ? 'flex' : 'none'; }
  if (bMoney) { bMoney.textContent = pendingMoney; bMoney.style.display = pendingMoney > 0 ? 'flex' : 'none'; }
  if (bLeave) { bLeave.textContent = pendingLeave; bLeave.style.display = pendingLeave > 0 ? 'flex' : 'none'; }
}

function manualRefresh() { loadAll(); showToast('รีเฟรชข้อมูลแล้ว', 'info', 2000); }

function startAutoRefresh() {
  if (autoRefreshTimer) clearInterval(autoRefreshTimer);
  autoRefreshTimer = setInterval(checkChanges, 3000); // ลดเหลือ 3 วินาที
}

// Broadcast notifications — tracks what already shown to avoid duplicate toasts
let lastBroadcastHwCount = -1;
let lastBroadcastTrCount = -1;
let lastBroadcastLvCount = -1;

function checkChanges() {
  google.script.run
    .withSuccessHandler(r => {
      if (!r || !r.success) return;

      // Homework changes
      if (currentHwCount >= 0 && r.hwCount !== currentHwCount) {
        loadHw();
        if (lastBroadcastHwCount !== r.hwCount) {
          lastBroadcastHwCount = r.hwCount;
          if (r.hwCount > currentHwCount) showToast('📚 มีการบ้านใหม่!', 'broadcast', 5000);
        }
      }

      // Treasury item changes
      if (currentTrCount >= 0 && r.trCount !== currentTrCount) {
        loadMoney();
        if (lastBroadcastTrCount !== r.trCount) {
          lastBroadcastTrCount = r.trCount;
          if (r.trCount > currentTrCount) showToast('💰 มีรายการเงินใหม่!', 'broadcast', 5000);
        }
      }

      // Payment changes (silent refresh, no toast needed)
      if (r.trPayCounter !== currentTrPayCounter && currentTrPayCounter !== -1) {
        loadMoney();
      }
      currentTrPayCounter = r.trPayCounter;

      if (typeof r.seatVersion === 'number' && currentSeatVersion >= 0 && r.seatVersion !== currentSeatVersion) {
        const activePanel = document.querySelector('.tab-panel.active');
        if (activePanel && activePanel.id === 'panelSeats') seatsReloadSilent();
      }
      if (typeof r.seatVersion === 'number') currentSeatVersion = r.seatVersion;

      // Leave/activity changes
      if (r.pendingLeaveCount !== currentLvCount) {
        if (r.pendingLeaveCount > currentLvCount) {
          if (u.canApproveLeave && lastBroadcastLvCount !== r.pendingLeaveCount) {
            lastBroadcastLvCount = r.pendingLeaveCount;
            showToast(`🚨 มี${r.lastPendingType || 'คำขอ'}ใหม่รออนุมัติ!`, 'error', 0);
          }
          if (!u.canApproveLeave && lastBroadcastLvCount !== r.pendingLeaveCount) {
            lastBroadcastLvCount = r.pendingLeaveCount;
          }
        }
        loadLeave();
      }
    })
    .getLastUpdate();
}

// --- Homework ---
function loadHw() {
  const el = document.getElementById('listHw');
  if (!hwData.length) showSkeleton('listHw', 2);

  google.script.run
    .withSuccessHandler(r => {
      if (r && r.success) hwData = Array.isArray(r.homework) ? r.homework : [];
      else if (r && Array.isArray(r.data)) hwData = r.data;
      else if (Array.isArray(hwData)) {}
      else hwData = [];
      renderHw();
      currentHwCount = hwData.length;
      updateNotifBadges();
    })    .withFailureHandler(err => {
      console.error('loadHw failed', err);
      if (!hwData.length && el) el.innerHTML = '<p style="text-align:center;color:var(--danger);padding:20px;">เชื่อมต่อไม่ได้</p>';
    })
    .getHomework();
}

function saveHomework() {
  const btn = document.getElementById('btnSaveHw');
  const originalText = btn.innerHTML;
  
  btn.disabled = true;
  btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> กำลังบันทึก...';
  
  const data = { /* ข้อมูลการบ้าน */ };
  
  google.script.run.withSuccessHandler(res => {
    btn.disabled = false;
    btn.innerHTML = originalText;
    showToast('บันทึกสำเร็จ!', 'success');
    closeModal('modalHw');
    loadAll();
  }).addHomework(data);
}

function renderHw() {
  const el = document.getElementById('listHw');
  if (!hwData.length) { 
    el.innerHTML = '<p style="text-align:center; color:var(--text-muted); padding: 40px;">ไม่มีการบ้าน</p>';
    return; 
  }
  
  el.innerHTML = hwData.map(h => {
    const isDone = u.studentNo && h.statuses[u.studentNo]?.status === 'completed';
    const imagePath = h.statuses[u.studentNo]?.imagePath;
    const hasStudentStatus = u.studentNo && h.statuses[u.studentNo];
    const isDarkMode = document.documentElement.getAttribute('data-theme') === 'dark';
    const baseColor = h.color || '#E0E0E0';
    const subjectColor = isDarkMode ? getLuminousColor(baseColor, 50) : baseColor;
    const textColor = isDarkMode ? getLuminousColor(baseColor, 85) : getDarkColor(baseColor, -55);
    const assignedDate = h.assignedDate ? new Date(h.assignedDate) : null;
    const dueDate = h.noDueDate ? null : h.dueDate ? new Date(h.dueDate) : null;
    const now = new Date();
    const daysSinceAssigned = assignedDate ? Math.max(0, Math.floor((now - assignedDate) / 86400000)) : 0;
    const isOverdue = dueDate && dueDate < new Date(new Date().setHours(0,0,0,0)) && !isDone;
    const ageClass = isOverdue ? 'overdue' : (daysSinceAssigned >= 3 ? 'warning' : '');
    const ageLabel = isOverdue ? 'เลยกำหนดแล้ว' : `สั่งมา ${daysSinceAssigned} วัน`;
    return `
<div class="card">
  <div class="subject-color-bar" style="background-color: ${subjectColor};"></div>
  <div class="card-head">
    <div>
      <span class="subject-title" style="background-color: ${subjectColor}40; color: ${textColor};">${escapeHtml(h.subject)}</span><br>
      <small style="color:var(--text-secondary)">${escapeHtml(h.description)}</small>
    </div>
    ${u.canManageHomework ? `<button class="btn btn-danger btn-icon" onclick="delHw('${h.id}')"><i class="fas fa-trash"></i></button>` : ''}
  </div>
  <div class="card-body">
    <p style="margin-bottom:8px;">
      <i class="fas fa-calendar-alt"></i> สั่ง: ${formatDate(h.assignedDate)} | 
      <i class="fas fa-calendar-check"></i> ส่ง: ${h.noDueDate ? 'ไม่มีกำหนด' : formatDate(h.dueDate)}
    </p>
    <div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:8px;">
      <span class="hw-age-badge ${ageClass}"><i class="fas fa-clock"></i> ${ageLabel}</span>
      ${isOverdue ? '<span class="hw-age-badge overdue"><i class="fas fa-exclamation-triangle"></i> ต้องส่งทันที</span>' : ''}
    </div>
    ${hasStudentStatus ? `
    <div style="background:var(--bg-tertiary); padding:10px; border-radius:8px; margin-top:10px;">
      <div style="margin-bottom:8px; display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
        <span class="badge ${isDone ? 'badge-approved' : 'badge-pending'}">${isDone ? 'ส่งแล้ว' : 'ยังไม่ส่ง'}</span>
        ${imagePath ? '<span style="font-size:0.8rem; color:var(--success);"><i class="fas fa-image"></i> มีรูป</span>' : ''}
      </div>
      ${imagePath ? `<img src="${imagePath}" style="max-width:100px; border-radius:8px; margin-bottom:10px; border: 2px solid var(--border);">` : ''}
      <div style="display:flex; gap:8px; flex-wrap:wrap;">
        <button class="btn btn-secondary btn-sm" onclick="hwUploadImg('${h.id}')">
          <i class="fas fa-camera"></i> แนบรูปภาพ
        </button>
        <button class="btn ${isDone ? 'btn-danger' : 'btn-success'} btn-sm" onclick="hwToggleDone('${h.id}', ${isDone}, '${imagePath || ''}')">
          <i class="fas ${isDone ? 'fa-undo' : 'fa-check'}"></i> ${isDone ? 'ย้อนสถานะ' : 'ทำเสร็จแล้ว'}
        </button>
      </div>
    </div>
    ` : ''}
  </div>
</div>`;
  }).join('');
}

function previewImage(input, previewId) {
  const file = input.files[0];
  const preview = document.getElementById(previewId);
  if (file) {
    const reader = new FileReader();
    reader.onload = function(e) {
      preview.src = e.target.result;
      preview.style.display = 'block'; // แสดงรูปเมื่อโหลดเสร็จ
    }
    reader.readAsDataURL(file);
  } else {
    preview.style.display = 'none'; // ซ่อนรูปถ้าไม่ได้เลือกไฟล์
  }
}

// ตัวอย่าง HTML ส่วนเลือกรูป
// <input type="file" onchange="previewImage(this, 'imgPreview')">
// <img id="imgPreview" style="display:none; width:100%; border-radius:8px;">

function hwUploadImg(hid) {
  const input = document.createElement('input'); 
  input.type = 'file'; 
  input.accept = 'image/*';
  input.onchange = e => {
    const file = e.target.files[0]; 
    if (!file) return;
    // Show uploading toast immediately
    showToast('⬆️ กำลังอัปโหลดรูป...', 'info', 0);
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      google.script.run
        .withSuccessHandler(r => { 
          // Clear uploading toast and show result
          document.querySelectorAll('.toast').forEach(t => { if (t.querySelector('.toast-message')?.textContent.includes('อัปโหลด')) t.remove(); });
          showToast('📎 แนบรูปสำเร็จ', 'success'); 
          loadHw(); 
        })
        .withFailureHandler(() => { 
          document.querySelectorAll('.toast').forEach(t => { if (t.querySelector('.toast-message')?.textContent.includes('อัปโหลด')) t.remove(); });
          showToast('ไม่สามารถอัปโหลดรูปได้', 'error'); 
        })
        .updateHomeworkStatus(hid, u.studentNo, null, reader.result);
    };
  };
  input.click();
}

function hwToggleDone(hid, isCurrentlyDone, imagePath) {
  if (!isCurrentlyDone && !imagePath) {
    if (!confirm("คุณจะไม่แนบรูปภาพจริงๆใช่ไหม?")) return;
  }
  const newStatus = isCurrentlyDone ? 'pending' : 'completed';
  // Optimistic update immediately
  hwData = hwData.map(h => {
    if (h.id === hid && u.studentNo && h.statuses[u.studentNo]) {
      return { ...h, statuses: { ...h.statuses, [u.studentNo]: { ...h.statuses[u.studentNo], status: newStatus } } };
    }
    return h;
  });
  renderHw();

  google.script.run
    .withSuccessHandler(r => { 
      showToast(newStatus === 'completed' ? '✅ ทำเสร็จแล้ว!' : '↩️ ย้อนสถานะแล้ว', 'success'); 
      loadHw(); // sync from server
    })
    .withFailureHandler(() => { 
      showToast('ไม่สามารถอัปเดตสถานะได้', 'error'); 
      loadHw(); // revert
    })
    .updateHomeworkStatus(hid, u.studentNo, newStatus, null);
}

function doAddHw(e) {
  e.preventDefault();
  const sub = document.getElementById('hwSub').value;
  const desc = document.getElementById('hwDesc').value.trim();
  const ad = document.getElementById('hwDate').value;
  const dd = document.getElementById('hwDue').value;
  const nd = document.getElementById('hwNoDue').checked;
  const by = u.displayName;

  if (!sub || !desc) {
    showToast('กรุณากรอกข้อมูลให้ครบถ้วน', 'error');
    return;
  }

  const btn = e.submitter || e.target.querySelector('button[type="submit"]');
  const restoreBtn = btnLoading(btn);

  google.script.run
    .withSuccessHandler(r => {
      restoreBtn();
      if (r.success) {
        showToast('บันทึกสำเร็จ!', 'success');
        closeModal('modalHw');
        loadHw();
      } else {
        showToast(r.message || 'บันทึกไม่สำเร็จ', 'error');
      }
    })
    .withFailureHandler(() => {
      restoreBtn();
      showToast('เกิดข้อผิดพลาดในการเชื่อมต่อ', 'error');
    })
    .addHomework(sub, desc, ad, dd, nd, by);
}

function delHw(id) { 
  if (confirm('ยืนยันลบการบ้าน?')) { 
    // Optimistic remove
    const backup = [...hwData];
    hwData = hwData.filter(h => h.id !== id);
    renderHw();
    google.script.run
      .withSuccessHandler(() => { showToast('ลบการบ้านแล้ว', 'success'); loadHw(); })
      .withFailureHandler(() => { showToast('ไม่สามารถลบการบ้านได้', 'error'); hwData = backup; renderHw(); })
      .deleteHomework(id); 
  } 
}

// --- Money ---
function loadMoney() {
  const el = document.getElementById('listMoney');
  if (!moneyData.length) showSkeleton('listMoney', 2);
  google.script.run
    .withSuccessHandler(r => {
      if (r && r.success) moneyData = Array.isArray(r.treasury) ? r.treasury : [];
      else if (r && Array.isArray(r.data)) moneyData = r.data;
      else if (Array.isArray(moneyData)) {}
      else moneyData = [];
      currentTrCount = moneyData.length;
      renderMoney();
      updateNotifBadges();
    })
    .withFailureHandler(() => {
      if (!moneyData.length && el) el.innerHTML = '<p style="text-align:center;color:var(--danger);padding:20px;">เชื่อมต่อไม่ได้</p>';
    })
    .getTreasuryItems();
}

function renderMoney() {
  const el = document.getElementById('listMoney');
  if (!moneyData.length) { 
    el.innerHTML = '<p style="text-align:center; color:var(--text-muted); padding: 40px;">ไม่มีรายการ</p>';
    return; 
  }
  
  el.innerHTML = moneyData.map(m => {
    const itemColor = m.color || '#E0E0E0';
    const textColor = getDarkColor(itemColor, -55);
    const payments = m.payments || {};
    const paidCount = Object.values(payments).filter(p => p.amountPaid >= m.amountPerPerson).length;
    const totalCount = Object.keys(payments).length;
    const userPayment = u.studentNo ? payments[u.studentNo] : null;
    const isUserPaid = userPayment && userPayment.amountPaid >= m.amountPerPerson;
    const receiptButton = isUserPaid
      ? `<button class="btn btn-secondary btn-sm" onclick='showReceipt("money", {item: ${JSON.stringify(m).replace(/'/g,"&#39;")}, studentNo: "${u.studentNo}"})'>
           <i class="fas fa-receipt"></i> ใบเสร็จ</button>`
      : '';

    return `
<div class="card">
  <div class="subject-color-bar" style="background-color: ${itemColor};"></div>
  <div class="card-head">
    <div>
      <span class="subject-title" style="background-color: ${itemColor}40; color: ${textColor};">${escapeHtml(m.title)}</span><br>
      <small style="color:var(--text-secondary)">จำนวนเงิน/คน ${parseFloat(m.amountPerPerson).toLocaleString('th-TH')} บาท</small>
    </div>
    ${u.canManageTreasury ? `<button class="btn btn-danger btn-icon" onclick="delMoney('${m.id}')"><i class="fas fa-trash"></i></button>` : ''}
  </div>
  <div class="card-body">
    <div class="check-grid">
      ${Object.entries(payments).map(([no, p]) => { 
        const std = students.find(s => String(s.no) === String(no)); 
        const isComplete = p.amountPaid >= m.amountPerPerson; 
        const canEdit = u.canManageTreasury;
        return `
<label class="check-card ${isComplete ? 'checked' : ''}" ${canEdit ? `onclick="handleCheck('${m.id}', ${no}, ${p.amountPaid}, ${m.amountPerPerson}, this)"` : ''} style="${!canEdit ? 'cursor:default;' : ''}">
  <input type="checkbox" ${isComplete ? 'checked' : ''} ${!canEdit ? 'disabled' : ''}>
  <span class="check-mark"><i class="fas fa-check"></i></span>
  <div class="check-card-content">
    <div class="check-no"> ${escapeHtml(String(no))}</div>
    <div class="check-name">${std ? escapeHtml(std.name.split(' ')[0]) : '-'}</div>
    <div class="check-amt">${parseFloat(p.amountPaid).toLocaleString('th-TH')}/${parseFloat(m.amountPerPerson).toLocaleString('th-TH')}</div>
  </div>
</label>`;
      }).join('')}
    </div>
    <div style="margin-top:15px; padding-top:10px; border-top:1px solid var(--border); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
      <span style="font-size:0.9rem; color:var(--text-secondary);">
        ชำระแล้ว <b style="color:var(--success)">${paidCount}</b> / ${totalCount} คน
      </span>
      ${receiptButton}
    </div>
    ${userPayment ? `<div style="margin-top:10px; font-size:0.88rem; color:${isUserPaid ? '#16a34a' : '#d97706'}; font-weight:600;">สถานะของคุณ: ${isUserPaid ? 'ชำระเรียบร้อยแล้ว' : 'ค้างชำระ'}</div>` : ''}
  </div>
</div>
`;
  }).join('');
}

// FIX: handleCheck now calls loadMoney on failure to revert optimistic UI
function handleCheck(tid, sno, currentPaid, required, element) {
  if (!u.canManageTreasury) return;
  const willPay = (currentPaid < required) ? required : 0;
  // Optimistic UI update
  const checkbox = element.querySelector('input');
  checkbox.checked = (willPay > 0);
  element.classList.toggle('checked', willPay > 0);
  const amtDiv = element.querySelector('.check-amt');
  if (amtDiv) amtDiv.textContent = willPay > 0 ? `${required}/${required}` : `0/${required}`;

  google.script.run
    .withSuccessHandler(r => { if (!r || !r.success) loadMoney(); })
    .withFailureHandler(() => { showToast('บันทึกไม่สำเร็จ', 'error'); loadMoney(); })
    .updatePayment(tid, sno, willPay);
}

function doAddMoney(e) {
  e.preventDefault(); 
  const t = document.getElementById('mTitle').value.trim(); 
  const a = document.getElementById('mAmt').value;
  const targetStr = document.getElementById('mTargetList')?.value?.trim() || '';
  if (!t || !a) return;

  let targetList = null;
  if (targetStr) {
    targetList = [];
    targetStr.split(',').forEach(p => {
      p = p.trim();
      if (p.includes('-')) {
        const [start, end] = p.split('-').map(Number);
        if (!isNaN(start) && !isNaN(end)) {
          for (let i = Math.min(start, end); i <= Math.max(start, end); i++) targetList.push(i);
        }
      } else {
        const n = Number(p);
        if (!isNaN(n)) targetList.push(n);
      }
    });
    targetList = [...new Set(targetList)];
  }

  const btn = e.submitter;
  const restore = btnLoading(btn);
  closeModal('modalMoney');
  document.getElementById('formMoney').reset();
  showToast('⏳ กำลังบันทึก...', 'info', 0);
  google.script.run
    .withSuccessHandler(r => { 
      restore();
      document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('กำลังบันทึก')) el.remove(); });
      if (r.success) { showToast('💰 ตั้งเก็บเงินสำเร็จ!', 'success'); loadMoney(); }
      else showToast(r.message || 'ไม่สำเร็จ', 'error');
    })
    .withFailureHandler(() => { 
      restore();
      document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('กำลังบันทึก')) el.remove(); });
      showToast('เกิดข้อผิดพลาด', 'error'); 
    })
    .addTreasuryItem(t, parseFloat(a), u.displayName, targetList);
}

function delMoney(id) { 
  if (confirm('?')) { 
    const backup = [...moneyData];
    moneyData = moneyData.filter(m => m.id !== id);
    renderMoney();
    google.script.run
      .withSuccessHandler(() => { showToast('', 'success'); loadMoney(); })
      .withFailureHandler(() => { showToast('', 'error'); moneyData = backup; renderMoney(); })
      .deleteTreasuryItem(id); 
  } 
}

// ============================================
//  RECEIPT ()
// ============================================
function showReceipt(type, data) {
  // type: 'money' | 'loan'
  const el = document.getElementById('receiptContent');
  if (!el) return;

  const now = new Date();
  const fallbackReceiptNo = 'RC-' + now.getFullYear() + String(now.getMonth()+1).padStart(2,'0') + String(now.getDate()).padStart(2,'0') + '-' + String(Math.floor(Math.random()*9000)+1000);
  const dateStr = now.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
  const timeStr = now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });

  if (type === 'money') {
    const { item, studentNo } = data;
    const std = students.find(s => String(s.no) === String(studentNo));
    const payment = item.payments?.[studentNo] || { amountPaid: 0 };
    const receiptNo = payment.receiptNo || ('RC-' + now.getFullYear() + String(now.getMonth()+1).padStart(2,'0') + String(now.getDate()).padStart(2,'0') + '-' + String(Math.floor(Math.random()*9000)+1000));
    const isPaid = payment.amountPaid >= item.amountPerPerson;
    const paidAtRaw = payment.paidAt || payment.updatedAt || payment.timestamp || null;
    const paidDate = paidAtRaw ? new Date(paidAtRaw) : null;
    const paidDateStr = paidDate ? paidDate.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' }) : dateStr;
    const paidTimeStr = paidDate ? paidDate.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }) : timeStr;

    el.innerHTML = `
      <div class="receipt-header">
        <div>
          <div class="receipt-logo"><i class="fas fa-school"></i></div>
        </div>
        <div>
          <h2>ใบเสร็จรับเงิน</h2>
          <p class="receipt-subtitle">รายการเงินห้องและค่าธรรมเนียม</p>
        </div>
        <div class="receipt-company">
          <strong>MyClass Web</strong>
          <span>ระบบจัดการห้องเรียน</span>
          <span>โทร. 02-123-4567</span>
        </div>
      </div>
      <div class="receipt-title-bar"></div>
      <div class="receipt-meta">
        <div>
          <div class="receipt-meta-label">เลขที่ใบเสร็จ</div>
          <div class="receipt-meta-value">${receiptNo}</div>
        </div>
        <div>
          <div class="receipt-meta-label">วันที่ / เวลา</div>
          <div class="receipt-meta-value">${paidDateStr} ${paidTimeStr}</div>
        </div>
        <div>
          <div class="receipt-meta-label">ผู้ชำระ</div>
          <div class="receipt-meta-value">${escapeHtml(std ? std.name : '-')}</div>
        </div>
        <div>
          <div class="receipt-meta-label">เลขที่นักเรียน</div>
          <div class="receipt-meta-value">${escapeHtml(studentNo)}</div>
        </div>
      </div>
      <table class="receipt-table">
        <thead>
          <tr>
            <th>รายการ</th>
            <th style="text-align:right;">จำนวนเงิน</th>
            <th style="text-align:right;">ชำระจริง</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>${escapeHtml(item.title)}</td>
            <td style="text-align:right;">${parseFloat(item.amountPerPerson).toLocaleString('th-TH')} บาท</td>
            <td style="text-align:right;">${parseFloat(payment.amountPaid).toLocaleString('th-TH')} บาท</td>
          </tr>
        </tbody>
        <tfoot>
          <tr class="receipt-total-row">
            <td><b>รวมทั้งสิ้น</b></td>
            <td></td>
            <td style="text-align:right;"><b>${parseFloat(payment.amountPaid).toLocaleString('th-TH')} บาท</b></td>
          </tr>
        </tfoot>
      </table>
      <div style="margin-top:12px; font-size:0.9rem; color:${isPaid ? '#16a34a' : '#dc2626'}; font-weight:600;">
        สถานะการชำระ: ${isPaid ? 'ชำระครบแล้ว' : 'ค้างชำระ'}
      </div>
      <div class="receipt-footer">
         <div>ผู้รับเงิน: MyClass Web</div>
         <div>${escapeHtml(u.displayName || 'ผู้ใช้งาน')}</div>
         <div>${dateStr}</div>
      </div>`;

  } else if (type === 'loan') {
    const loan = data;
    const borrower = students.find(s => String(s.no) === String(loan.borrowerNo));
    const lender = students.find(s => String(s.no) === String(loan.lenderNo));
    const isReturned = loan.status === 'returned';
    const loanDate = loan.createdAt ? formatDate(loan.createdAt) : dateStr;

    el.innerHTML = `
      <div class="receipt-header">
        <div>
          <div class="receipt-logo"><i class="fas fa-hand-holding-usd"></i></div>
        </div>
        <div>
          <h2>ใบเสร็จเงินยืม</h2>
          <p class="receipt-subtitle">ธุรกรรมเงินยืมภายในกลุ่ม</p>
        </div>
        <div class="receipt-company">
          <strong>MyClass Web</strong>
          <span>ระบบจัดการห้องเรียน</span>
          <span>โทร. 02-123-4567</span>
        </div>
      </div>
      <div class="receipt-title-bar receipt-loan"></div>
      <div class="receipt-meta">
        <div>
          <div class="receipt-meta-label">เลขที่</div>
          <div class="receipt-meta-value">${escapeHtml(loan.id || fallbackReceiptNo)}</div>
        </div>
        <div>
          <div class="receipt-meta-label">วันที่</div>
          <div class="receipt-meta-value">${loanDate}</div>
        </div>
        <div>
          <div class="receipt-meta-label">ผู้ยืม</div>
          <div class="receipt-meta-value">${escapeHtml(borrower ? borrower.name : loan.borrowerNo)}</div>
        </div>
        <div>
          <div class="receipt-meta-label">ผู้ให้ยืม</div>
          <div class="receipt-meta-value">${escapeHtml(lender ? lender.name : loan.lenderNo)}</div>
        </div>
      </div>
      <table class="receipt-table">
        <thead>
          <tr>
            <th>รายการ</th>
            <th style="text-align:right;">จำนวนเงิน</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>${loan.note ? escapeHtml(loan.note) : 'เงินยืมภายในกลุ่ม'}</td>
            <td style="text-align:right;">${parseFloat(loan.amount).toLocaleString('th-TH')} บาท</td>
          </tr>
        </tbody>
        <tfoot>
          <tr class="receipt-total-row">
            <td><b>รวมทั้งสิ้น</b></td>
            <td style="text-align:right;"><b>${parseFloat(loan.amount).toLocaleString('th-TH')} บาท</b></td>
          </tr>
        </tfoot>
      </table>
      ${loan.proofUrl ? `<div class="receipt-qr-area"><p style="font-size:0.8rem;color:#64748b;margin-bottom:4px;">หลักฐานการยืม</p><img src="${escapeHtml(loan.proofUrl)}" class="receipt-proof-img" alt="หลักฐาน"></div>` : ''}
      <div style="margin-top:12px; font-size:0.9rem; color:${isReturned ? '#16a34a' : '#2563eb'}; font-weight:600;">
        สถานะ: ${isReturned ? 'คืนเงินแล้ว' : 'ยังไม่คืน'}</div>
      <div class="receipt-footer">
         <div>ผู้รับรอง: MyClass Web</div>
         <div>${escapeHtml(u.displayName || 'ผู้ใช้งาน')}</div>
         <div>${dateStr}</div>
      </div>`;
  }

  openModal('modalReceipt');
}

function printReceipt() {
  window.print();
}

function showLoanReceipt(id) {
  const loan = loanData.find(l => l.id === id);
  if (!loan) return showToast('', 'error');
  showReceipt('loan', loan);
}

// --- RECEIPT PANEL / LIST ---
function loadReceipts() {
  // Ensure latest data exists (moneyData and loanData are kept up-to-date by loadAll/loadMoney/loadLoan)
  renderReceipts();
}

function renderReceipts() {
  const wrap = document.getElementById('listReceipt');
  if (!wrap) return;
  const receipts = [];

  // Money receipts (only for current user where amountPaid > 0)
  (moneyData || []).forEach(item => {
    const payments = item.payments || {};
    Object.entries(payments).forEach(([no, p]) => {
      // Show only if this receipt belongs to current user
      if (String(no) === String(u.studentNo)) {
        const amt = parseFloat(p.amountPaid || 0);
        if (amt > 0) {
          receipts.push({ type: 'money', item, studentNo: no, payment: p });
        }
      }
    });
  });

  // Loan receipts (only where current user is borrower or lender)
  (loanData || []).forEach(l => {
    if (String(l.borrowerNo) === String(u.studentNo) || String(l.lenderNo) === String(u.studentNo)) {
      receipts.push({ type: 'loan', loan: l });
    }
  });

  if (!receipts.length) {
    wrap.innerHTML = '<p style="text-align:center; color:var(--text-muted); padding:30px;">ไม่มีใบเสร็จ</p>';
    return;
  }

  wrap.innerHTML = receipts.map(r => {
    if (r.type === 'money') {
      const m = r.item;
      const std = students.find(s => String(s.no) === String(r.studentNo));
      const paidAt = r.payment?.paidAt || r.payment?.updatedAt || '';
      const dateStr = paidAt ? new Date(paidAt).toLocaleString('th-TH') : '';
      return `
      <div class="card">
        <div class="card-head">
          <div>
            <span class="subject-title">${escapeHtml(m.title)}</span>
            <div style="font-size:0.85rem; color:var(--text-secondary);">ผู้ชำระ: ${escapeHtml(std ? std.name : r.studentNo)}</div>
          </div>
          <div>
            <button class="btn btn-secondary btn-sm" onclick='showReceipt("money", {item: ${JSON.stringify(m).replace(/'/g,"&#39;")}, studentNo: "${escapeHtml(r.studentNo)}"})'><i class="fas fa-receipt"></i> ดู</button>
          </div>
        </div>
        <div class="card-body">
          <div>ยอดชำระ: <b>${parseFloat(r.payment.amountPaid).toLocaleString('th-TH')} บาท</b></div>
          <div style="color:var(--text-muted); margin-top:6px;">วันที่ชำระ: ${escapeHtml(dateStr)}</div>
        </div>
      </div>`;
    } else {
      const loan = r.loan;
      const borrower = students.find(s => String(s.no) === String(loan.borrowerNo));
      const lender = students.find(s => String(s.no) === String(loan.lenderNo));
      return `
      <div class="card loan-card">
        <div class="card-head">
          <div>
            <span class="loan-amount">${parseFloat(loan.amount).toLocaleString('th-TH')} บาท</span>
            <div style="font-size:0.85rem; color:var(--text-secondary);">ผู้ยืม: ${escapeHtml(borrower ? borrower.name : loan.borrowerNo)}</div>
          </div>
          <div>
            <button class="btn btn-secondary btn-sm" onclick='showReceipt("loan", ${JSON.stringify(loan).replace(/'/g,"&#39;")})'><i class="fas fa-receipt"></i> ดู</button>
          </div>
        </div>
        <div class="card-body">
          <div>สถานะ: ${loan.status === 'returned' ? 'คืนแล้ว' : 'ยังไม่คืน'}</div>
          <div style="color:var(--text-muted); margin-top:6px;">วันที่: ${escapeHtml(formatDate(loan.createdAt))}</div>
        </div>
      </div>`;
    }
  }).join('');
}

// ============================================
//  LOAN ()
// ============================================
function populateLoanStudents() {
  const borrowerSel = document.getElementById('loanBorrowerNo');
  const lenderSel = document.getElementById('loanLenderNo');
  if (!borrowerSel || !lenderSel) return;

  const opts = students.map(s => `<option value="${s.no}">${s.no} - ${s.name}</option>`).join('');
  borrowerSel.innerHTML = '<option value="">--  --</option>' + opts;
  lenderSel.innerHTML = '<option value="">--  --</option>' + opts;
}

function loadLoan() {
  const el = document.getElementById('listLoan');
  if (!loanData.length && el) showSkeleton('listLoan', 2);
  google.script.run
    .withSuccessHandler(r => {
      loanData = (r && r.success && Array.isArray(r.loans)) ? r.loans : (loanData || []);
      renderLoan();
    })
    .withFailureHandler(() => {
      if (!loanData.length && el) el.innerHTML = '<p style="text-align:center;color:var(--danger);padding:20px;"></p>';
    })
    .getLoans();
}

function renderLoan() {
  const el = document.getElementById('listLoan');
  if (!el) return;

  if (!loanData.length) {
    el.innerHTML = `
      <div class="card">
        <div class="card-body" style="text-align:center; color:var(--text-muted); padding:40px;">
          <i class="fas fa-hand-holding-usd" style="font-size:2rem; margin-bottom:10px; display:block;"></i>
          ไม่มีรายการเงินยืมในขณะนี้
        </div>
      </div>`;
    return;
  }

  el.innerHTML = loanData.map(loan => {
    const borrower = students.find(s => String(s.no) === String(loan.borrowerNo));
    const lender = students.find(s => String(s.no) === String(loan.lenderNo));
    const isReturned = loan.status === 'returned';
    const badgeClass = isReturned ? 'returned' : 'overdue';
    const badgeText = isReturned ? 'คืนแล้ว' : 'ค้างคืน';

    const canManage = u.canManageTreasury || String(loan.borrowerNo) === String(u.studentNo) || String(loan.lenderNo) === String(u.studentNo);

    return `
<div class="card loan-card">
  <div class="card-head">
    <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
      <span class="loan-amount">${parseFloat(loan.amount).toLocaleString('th-TH')} บาท</span>
      <span class="loan-badge ${badgeClass}">${badgeText}</span>
    </div>
    <div style="display:flex; gap:6px; flex-wrap:wrap;">
      <button class="btn btn-secondary btn-sm" onclick="showLoanReceipt('${escapeHtml(loan.id)}')">
        <i class="fas fa-receipt"></i> ใบเสร็จ
      </button>
      ${u.canManageTreasury && !isReturned ? `<button class="btn btn-success btn-sm" onclick="markLoanReturned('${loan.id}')"><i class="fas fa-check"></i> คืนแล้ว</button>` : ''}
      ${u.canManageTreasury ? `<button class="btn btn-danger btn-icon" onclick="deleteLoan('${loan.id}')"><i class="fas fa-trash"></i></button>` : ''}
    </div>
  </div>
  <div class="card-body">
    <div class="loan-meta">
      <span><i class="fas fa-user-graduate"></i> ผู้ยืม: <b>${escapeHtml(borrower ? borrower.name : loan.borrowerNo)}</b></span>
      <span><i class="fas fa-user"></i> ผู้ให้ยืม: <b>${escapeHtml(lender ? lender.name : loan.lenderNo)}</b></span>
      <span><i class="fas fa-hashtag"></i> เลขที่ผู้ยืม: <b>${escapeHtml(String(loan.borrowerNo))}</b></span>
      <span><i class="fas fa-hashtag"></i> เลขที่ผู้ให้ยืม: <b>${escapeHtml(String(loan.lenderNo))}</b></span>
      <span><i class="fas fa-calendar-alt"></i> วันที่: ${formatDate(loan.createdAt)}</span>
      ${loan.note ? `<span><i class="fas fa-sticky-note"></i> บันทึก: ${escapeHtml(loan.note)}</span>` : ''}
    </div>
    ${loan.proofUrl ? `
      <details class="leave-proof-details" style="margin-top:8px;">
        <summary><i class="fas fa-image"></i> รูปหลักฐาน</summary>
        <img src="${escapeHtml(loan.proofUrl)}" alt="หลักฐานเงินยืม" style="max-width:100%; margin-top:8px; border-radius:8px;">
      </details>` : ''}
  </div>
</div>`;
  }).join('');
}

function doAddLoan(e) {
  e.preventDefault();
  const borrowerNo = document.getElementById('loanBorrowerNo').value;
  const lenderNo = document.getElementById('loanLenderNo').value;
  const amount = document.getElementById('loanAmount').value;
  const note = document.getElementById('loanNote').value.trim();
  const imageFile = document.getElementById('loanImage').files?.[0];

  if (!borrowerNo || !lenderNo || !amount) {
    showToast('กรุณาเลือกผู้ยืม ผู้ให้ยืม และจำนวนเงินให้ครบ', 'error');
    return;
  }
  if (borrowerNo === lenderNo) {
    showToast('ผู้ยืมและผู้ให้ยืมต้องไม่ใช่คนเดียวกัน', 'error');
    return;
  }

  const btn = e.submitter || document.querySelector('#formLoan button[type="submit"]');
  const restore = btnLoading(btn);
  closeModal('modalLoan');
  document.getElementById('formLoan').reset();
  document.getElementById('loanPreview').style.display = 'none';
  showToast('กำลังบันทึกข้อมูลเงินยืม...', 'info', 0);

  const save = (proofUrl) => {
    google.script.run
      .withSuccessHandler(r => {
        restore();
        document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('กำลังบันทึก')) el.remove(); });
        if (r && r.success) {
          showToast('บันทึกเงินยืมสำเร็จ', 'success');
          loadLoan();
        } else {
          showToast(r?.message || 'ไม่สามารถบันทึกเงินยืมได้', 'error');
        }
      })
      .withFailureHandler(() => {
        restore();
        document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('กำลังบันทึก')) el.remove(); });
        showToast('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้', 'error');
      })
      .addLoan(parseInt(borrowerNo), parseInt(lenderNo), parseFloat(amount), note, proofUrl || '', u.displayName);
  };

  if (imageFile) {
    const reader = new FileReader();
    reader.onload = () => save(reader.result);
    reader.readAsDataURL(imageFile);
  } else {
    save('');
  }
}

function markLoanReturned(id) {
  if (!confirm('ยืนยันว่าเงินยืมรายการนี้ได้คืนแล้ว?')) return;
  showLoading('กำลังอัปเดตสถานะ...');
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r && r.success) { showToast('บันทึกสถานะคืนเงินแล้ว', 'success'); loadLoan(); }
      else showToast(r?.message || 'ไม่สามารถอัปเดตสถานะได้', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เกิดข้อผิดพลาดในการเชื่อมต่อ', 'error'); })
    .updateLoanStatus(id, 'returned');
}

function deleteLoan(id) {
  if (!confirm('ยืนยันลบรายการเงินยืมนี้?')) return;
  showLoading('กำลังลบรายการ...');
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r && r.success) { showToast('ลบรายการเงินยืมแล้ว', 'success'); loadLoan(); }
      else showToast(r?.message || 'ไม่สามารถลบรายการได้', 'error');
    })
    .withFailureHandler(() => { hideLoading(); showToast('เกิดข้อผิดพลาดในการเชื่อมต่อ', 'error'); })
    .deleteLoan(id);
}

// --- Leave ---
function loadLeave() {
  if (!leaveData.length) showSkeleton('listLeave', 2);
  google.script.run
    .withSuccessHandler(r => {
      if (Array.isArray(r)) leaveData = r;
      else if (r && Array.isArray(r.data)) leaveData = r.data;
      else if (Array.isArray(leaveData)) {} // keep existing
      else leaveData = [];
      renderLeave();
      currentLvCount = leaveData.filter(l => l.status === 'PENDING').length;
      updateNotifBadges();
    })
    .withFailureHandler(() => {
      if (!leaveData.length) document.getElementById('listLeave').innerHTML = '<p style="text-align:center;color:var(--danger);padding:20px;">เชื่อมต่อไม่ได้</p>';
    })
    .getLeaveRequests();
}

function renderLeave() {
  const el = document.getElementById('listLeave');
  if (!el) return;
  
  let list = Array.isArray(leaveData) ? leaveData : [];
  if (!u.canApproveLeave && u.studentNo) {
    list = list.filter(l => String(l.studentNo) === String(u.studentNo));
  }
  
  if (!list.length) {
    el.innerHTML = `
      <div class="card">
        <div class="card-body" style="text-align:center; color:var(--text-muted); padding:30px;">
          <i class="fas fa-calendar-times" style="font-size:1.6rem; margin-bottom:10px; display:block;"></i>
          ${u.canApproveLeave ? 'ไม่มีคำขอลา / กิจกรรม' : 'ไม่มีคำขอของคุณ'}
        </div>
      </div>`;
    return;
  }
  
  el.innerHTML = list.map(l => {
    const isPending = l.status === 'PENDING';
    const badgeClass = isPending ? 'badge-pending' : (l.status === 'APPROVED' ? 'badge-approved' : 'badge-rejected');
    const badgeText = isPending ? 'รออนุมัติ' : (l.status === 'APPROVED' ? 'อนุมัติแล้ว' : 'ไม่อนุมัติ');
    const isOwner = String(l.studentNo) === String(u.studentNo);
    
    return `
<div class="card">
  <div class="card-head">
    <div>
      <b style="font-size: 1.1rem;">${escapeHtml(l.studentName)}</b><br>
      <small style="color:var(--text-secondary)">เลขที่ ${escapeHtml(String(l.studentNo))} | ${escapeHtml(l.type)}</small>
    </div>
    <span class="badge ${badgeClass}">${badgeText}</span>
  </div>
  <div class="card-body">
    <p style="margin-bottom:10px;">
      <i class="fas fa-calendar-alt"></i> วันที่: ${formatDate(l.date)}<br>
      <i class="fas fa-comment-dots"></i> เหตุผล: ${escapeHtml(l.reason)}
    </p>
    ${l.proofImage ? `
    <details class="leave-proof-details">
      <summary><i class="fas fa-image"></i> หลักฐานการลา (แตะเพื่อเปิด/ปิด)</summary>
      <img src="${l.proofImage}" alt="หลักฐานการลา" loading="lazy">
    </details>` : ''}
    ${isPending ? `
    <div style="margin-top:12px; padding:10px; background:var(--bg-tertiary); border-radius:8px;">
      <div style="display:flex; gap:10px; flex-wrap:wrap;">
        ${u.canApproveLeave ? `
        <button class="btn btn-success btn-sm" onclick="updLeave('${l.id}', 'APPROVED')">
          <i class="fas fa-check"></i> อนุมัติ
        </button>
        <button class="btn btn-danger btn-sm" onclick="updLeave('${l.id}', 'REJECTED')">
          <i class="fas fa-times"></i> ไม่อนุมัติ
        </button>` : ''}
        ${isOwner && !l.proofImage ? `
        <button class="btn btn-secondary btn-sm" onclick="confirmLeave('${l.id}')">
          <i class="fas fa-camera"></i> แนบ/อัพเดทรูป
        </button>` : ''}
        ${isOwner && l.proofImage ? `
        <span style="color:var(--success); font-size:0.9rem; line-height:32px;">
          <i class="fas fa-check-circle"></i> แนบรูปแล้ว
        </span>` : ''}
      </div>
    </div>` : ''}
  </div>
</div>`;
  }).join('');
}

function doAddLeave(e) {
  e.preventDefault();
  const type = document.getElementById('lType').value;
  const date = document.getElementById('lDate').value;
  const reason = document.getElementById('lReason').value;
  const file = document.getElementById('lImage').files[0];
  
  if (!type || !date || !reason) { showToast('กรุณากรอกข้อมูลให้ครบ', 'error'); return; }
  if (!file && !confirm("คุณจะส่งคำขอโดยไม่มีรูปภาพจริง ๆ ใช่ไหม?")) return;
  
  closeModal('modalLeave');
  document.getElementById('formLeave').reset();
  showToast('⏳ กำลังส่งคำขอ...', 'info', 0);

  const doSubmit = (imageData) => {
    google.script.run
      .withSuccessHandler(r => {
        document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('กำลังส่งคำขอ')) el.remove(); });
        showToast(imageData ? '📋 ส่งคำขอลาสำเร็จ!' : '📋 ส่งคำขอลาสำเร็จ (ไม่มีรูป)', 'success');
        loadLeave();
      })
      .withFailureHandler(() => { 
        document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('กำลังส่งคำขอ')) el.remove(); });
        showToast('ส่งคำขอไม่สำเร็จ', 'error'); 
      })
      .submitLeaveRequest(u.studentNo, u.displayName, type, date, reason, imageData || null);
  };
  
  if (file) {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => doSubmit(reader.result);
  } else {
    doSubmit(null);
  }
}

function updLeave(id, stat) {
  // Optimistic update
  leaveData = leaveData.map(l => l.id === id ? { ...l, status: stat } : l);
  renderLeave();
  google.script.run
    .withSuccessHandler(() => { 
      showToast(stat === 'APPROVED' ? '✅ อนุมัติแล้ว' : '❌ ไม่อนุมัติ', 'success'); 
      loadLeave(); 
    })
    .withFailureHandler(() => { showToast('ไม่สามารถอัปเดตสถานะได้', 'error'); loadLeave(); })
    .updateLeaveStatus(id, stat);
}

function confirmLeave(id) {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    showToast('⬆️ กำลังอัปโหลดรูป...', 'info', 0);
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      google.script.run
        .withSuccessHandler(() => { 
          document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('อัปโหลด')) el.remove(); });
          showToast('📎 อัปโหลดรูปยืนยันสำเร็จ', 'success'); 
          loadLeave(); 
        })
        .withFailureHandler(() => { 
          document.querySelectorAll('.toast').forEach(el => { if (el.querySelector('.toast-message')?.textContent.includes('อัปโหลด')) el.remove(); });
          showToast('ไม่สามารถอัปโหลดรูปได้', 'error'); 
        })
        .confirmLeaveRequest(id, reader.result);
    };
  };
  input.click();
}

function toggleAllLeaveEvidence() {
  const details = document.querySelectorAll('.leave-proof-details');
  if (!details.length) return;
  const anyOpen = Array.from(details).some(d => d.open);
  details.forEach(d => d.open = !anyOpen);
  showToast(anyOpen ? 'ปิดหลักฐานทั้งหมดแล้ว' : 'เปิดหลักฐานทั้งหมดแล้ว', 'info', 2000);
}

function doChangePw(e) { 
  e.preventDefault(); 
  const o = document.getElementById('pwOld').value; 
  const n = document.getElementById('pwNew').value; 
  const c = document.getElementById('pwConf').value;
  const btn = e.submitter;
  const restore = btnLoading(btn);
  google.script.run
    .withSuccessHandler(r => { restore(); showToast(r.message, r.success ? 'success' : 'error'); })
    .withFailureHandler(() => { restore(); showToast('ไม่สามารถเปลี่ยนรหัสผ่านได้', 'error'); })
    .changePassword(u.id, o, n, c); 
}

// ฟังก์ชันปรับสีให้เข้มขึ้นสำหรับตัวอักษร
function getDarkColor(hex, percent = -55) {
  if (!hex || hex.length < 4) return '#333333';
  hex = hex.replace(/^#/, '');
  if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
  let r = parseInt(hex.substring(0, 2), 16);
  let g = parseInt(hex.substring(2, 4), 16);
  let b = parseInt(hex.substring(4, 6), 16);

  r = Math.max(0, Math.min(255, r + Math.round(255 * (percent / 100))));
  g = Math.max(0, Math.min(255, g + Math.round(255 * (percent / 100))));
  b = Math.max(0, Math.min(255, b + Math.round(255 * (percent / 100))));

  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).padStart(6, '0');
}

// --- Utilities ---
function toggleForm(t) { 
  document.getElementById('loginBox').style.display = t === 'reg' ? 'none' : 'block'; 
  document.getElementById('regForm').style.display = t === 'reg' ? 'block' : 'none'; 
  document.getElementById('loginError').style.display = 'none'; 
  document.getElementById('regError').style.display = 'none'; 
}

function getHint() { 
  const name = document.getElementById('forgotName').value.trim(); 
  if (!name) return; 
  const res = document.getElementById('forgotResult');
  res.style.display = 'block';
  res.style.background = 'var(--bg-tertiary)';
  res.style.color = 'var(--text-secondary)';
  res.innerHTML = '<i class="fas fa-spinner fa-spin"></i> กำลังค้นหา...';
  google.script.run
    .withSuccessHandler(r => { 
      res.innerHTML = r.success ? `<b>คำใบ้:</b> ${r.hint}` : r.message;
      res.style.color = r.success ? 'var(--text-primary)' : 'var(--danger)';
    })
    .withFailureHandler(() => {
      res.innerHTML = 'เกิดข้อผิดพลาดในการเชื่อมต่อ';
      res.style.color = 'var(--danger)';
    })
    .getPasswordHint(name); 
}

// --- Code Form Dynamic ---
function updateCodeForm() {
  const typeEl = document.getElementById('codeType');
  if (!typeEl) return;
  const type = typeEl.value;
  const el = document.getElementById('codeFormDynamic');
  if (!el) return;
  
  if (type === 'TEMP_ACCESS') {
    el.innerHTML = `
      <div class="form-group">
        <label>ระยะเวลาที่ได้รับสิทธิ์</label>
        <div style="display:flex; gap:10px;">
          <input type="number" id="tempQty" value="1" min="1" style="flex:2;">
          <select id="tempUnit" style="flex:3;">
            <option value="days">วัน</option>
            <option value="weeks">อาทิตย์</option>
            <option value="months">เดือน</option>
            <option value="years">ปี</option>
          </select>
        </div>
      </div>`;
  } else if (type === 'ADD_HW') {
    el.innerHTML = `
      <div class="form-group">
        <label>จำนวนเครดิตตั้งการบ้าน</label>
        <input type="number" id="hwCredits" min="1" value="1" placeholder="เช่น 1">
      </div>`;
  } else if (type === 'ADD_MONEY') {
    el.innerHTML = `
      <div class="form-group">
        <label>ชื่อรายการ (เช่น ค่าเสื้อห้อง)</label>
        <input type="text" id="codeTitle" placeholder="ระบุชื่อรายการ">
      </div>
      <div class="form-group">
        <label>จำนวนเงินต่อคน (บาท)</label>
        <input type="number" id="codeAmt" placeholder="0" min="1">
      </div>`;
  }
}

function createCode() {
  const type = document.getElementById('codeType').value;
  const note = document.getElementById('codeNote').value;
  let val = '';
  const details = { note };

  if (type === 'ADD_MONEY') {
    details.title = document.getElementById('codeTitle')?.value?.trim();
    details.amount = document.getElementById('codeAmt')?.value;
    if (!details.title || !details.amount) {
      showToast('กรุณากรอกชื่อรายการและจำนวนเงิน', 'error');
      return;
    }
    val = '1';
  } else if (type === 'TEMP_ACCESS') {
    const qty = Number(document.getElementById('tempQty')?.value);
    const unit = document.getElementById('tempUnit')?.value || 'days';
    if (!qty || qty < 1) { showToast('กรุณากรอกจำนวนที่จะให้สิทธิ์', 'error'); return; }
    // FIX: Store value as "qty_unit" so server can parse both qty and unit correctly
    val = `${qty}_${unit.toUpperCase()}`;
    details.unit = unit;
  } else if (type === 'ADD_HW') {
    const hw = Number(document.getElementById('hwCredits')?.value);
    if (!hw || hw < 1) { showToast('กรุณากรอกจำนวนเครดิตให้ถูกต้อง', 'error'); return; }
    val = String(hw);
  }

  const createBtn = document.querySelector('#createCodeCard .btn-success');
  const restore = btnLoading(createBtn);
  google.script.run
    .withSuccessHandler(r => {
      restore();
      if (r.success) {
        document.getElementById('generatedCodeBox').style.display = 'block';
        document.getElementById('generatedCodeText').textContent = r.code;
        const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${r.code}&ecc=H`;
        document.getElementById('qrImage').src = qrUrl;
        showToast('🎟️ สร้างโค้ดสำเร็จ!', 'success');
        document.getElementById('generatedCodeBox').scrollIntoView({ behavior: 'smooth' });
      } else {
        showToast(r.message || 'ไม่สามารถสร้างโค้ดได้', 'error');
      }
    })
    .withFailureHandler(() => { restore(); showToast('เกิดข้อผิดพลาดสร้างโค้ด', 'error'); })
    .generateCode(type, val, details);
}

function copyCode() {
  const text = document.getElementById('generatedCodeText').textContent;
  navigator.clipboard.writeText(text).then(() => showToast('Copied!', 'success'));
}

// doRedeem — no overlay, use inline button spinner
function doRedeem() {
  const code = document.getElementById('redeemInput').value.trim().toUpperCase();
  if (!code) { showToast('กรุณากรอกโค้ด', 'error'); return; }
  const btn = document.querySelector('#panelRedeem .btn-primary');
  const restore = btnLoading(btn);
  google.script.run
    .withSuccessHandler(r => {
      restore();
      if (r.success) {
        showToast(r.message, 'success');
        document.getElementById('redeemInput').value = '';
        manualRefresh();
        if (r.action === 'ADD_MONEY') setTab('money');
        if (r.action === 'ADD_HW') {
          u.hwCredits = (u.hwCredits || 0) + parseInt(r.addedCredits || 0);
          localStorage.setItem('user', JSON.stringify(u));
          document.getElementById('meHwCredits').textContent = u.hwCredits;
          document.getElementById('mobileSidebarCredits').textContent = u.hwCredits;
          document.getElementById('btnAddHw').style.display = 'inline-flex';
        }
      } else {
        showToast(r.message, 'error');
      }
    })
    .withFailureHandler(() => { restore(); showToast('Server error redeem', 'error'); })
    .redeemCode(code, u.id);
}

// ============================================
// 🔑 GUEST ACCOUNT SYSTEM (TEMP_ACCESS)
// ใส่โค้ด → สร้างบัญชี GUEST อัตโนมัติ → login เลย → countdown → kick + ลบบัญชี
// ============================================

let guestCountdownInterval = null;

function doLoginRedeem() {
  const code = document.getElementById('loginCode')?.value?.trim().toUpperCase();
  if (!code) { showToast('กรุณากรอกโค้ดก่อน', 'error'); return; }

  // Keep showLoading here — this transitions to a new page (login → app)
  showLoading();
  google.script.run
    .withSuccessHandler(r => {
      hideLoading();
      if (r.success) {
        u = r.user;
        localStorage.setItem('user', JSON.stringify(u));
        localStorage.setItem('guestExpiresAt', r.expiresAt);
        showApp();
        showToast(`🎉 เข้าสู่ระบบในฐานะ "${u.displayName}" สำเร็จ! (${r.durationText})`, 'success');
        startGuestCountdown(r.expiresAt);
      } else {
        showToast(r.message, 'error');
      }
    })
    .withFailureHandler(() => { hideLoading(); showToast('เชื่อมต่อ Server ไม่ได้', 'error'); })
    .createGuestAccount(code);
}

// Start countdown timer — runs every second, kicks out when expired
function startGuestCountdown(expiresAtISO) {
  if (guestCountdownInterval) clearInterval(guestCountdownInterval);

  // Show the timer elements
  document.getElementById('sidebarTimer')?.classList.add('show');
  document.getElementById('mobileSidebarTimer')?.classList.add('show');

  const expiresAt = new Date(expiresAtISO).getTime();

  function tick() {
    const now = Date.now();
    const remaining = expiresAt - now;

    if (remaining <= 0) {
      // TIME'S UP — kick out and delete the account
      clearInterval(guestCountdownInterval);
      guestCountdownInterval = null;
      expireGuestAccount();
      return;
    }

    const totalSeconds = Math.floor(remaining / 1000);
    const d = Math.floor(totalSeconds / 86400);
    const h = Math.floor((totalSeconds % 86400) / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;

    let display;
    if (d > 0) {
      display = `${d}วัน ${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
    } else {
      display = `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
    }

    // Update all timer displays
    const isUrgent = remaining < 5 * 60 * 1000; // last 5 minutes = urgent
    ['sidebarTimerDigits','mobileSidebarTimerDigits','meTimerDigits'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = display;
    });
    ['sidebarTimer','mobileSidebarTimer','guestTimerCard'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.classList.toggle('urgent', isUrgent);
    });

    // Warn at 5 minutes left (once)
    if (remaining < 5 * 60 * 1000 && remaining > (5 * 60 * 1000 - 2000)) {
      showToast('⚠️ เวลาใช้งานเหลือน้อยกว่า 5 นาที!', 'error', 0);
    }
    // Warn at 1 minute left (once)
    if (remaining < 60 * 1000 && remaining > (60 * 1000 - 2000)) {
      showToast('🚨 เวลาใช้งานเหลือน้อยกว่า 1 นาที!', 'error', 0);
    }
  }

  tick(); // run immediately
  guestCountdownInterval = setInterval(tick, 1000);
}

// Called when countdown hits zero
function expireGuestAccount() {
  const guestId = u?.id;
  const guestName = u?.displayName || 'บัญชีชั่วคราว';

  // Clear local session immediately
  if (autoRefreshTimer) clearInterval(autoRefreshTimer);
  localStorage.clear();
  u = null;

  // Tell server to delete the guest account
  if (guestId) {
    google.script.run
      .withSuccessHandler(() => {})
      .withFailureHandler(() => {})
      .deleteGuestAccount(guestId);
  }

  // Show expired overlay instead of just reloading
  document.getElementById('app').style.display = 'none';
  document.getElementById('loginPage').style.display = 'flex';

  // Show big toast
  showToast(`⏰ เวลาใช้งาน "${guestName}" หมดแล้ว บัญชีถูกลบเรียบร้อย`, 'error', 0);

  // Clear code input
  const codeInput = document.getElementById('loginCode');
  if (codeInput) codeInput.value = '';
}

// Restore countdown if user refreshes page while still within guest session
function restoreGuestCountdownIfNeeded() {
  if (!u || !u.isGuest) return;
  const expiresAt = localStorage.getItem('guestExpiresAt');
  if (!expiresAt) return;
  if (new Date(expiresAt).getTime() <= Date.now()) {
    // Already expired — clean up
    expireGuestAccount();
    return;
  }
  startGuestCountdown(expiresAt);
}

// --- QR Scanner ---
function startScanner() {
  document.getElementById('scanStatus').textContent = "กำลังขอสิทธิ์กล้อง...";
  document.getElementById('scanStatus').style.color = 'var(--text-secondary)';
  document.getElementById('stopCamBtn').style.display = 'none';
  document.getElementById('startCamBtn').style.display = 'none';
  
  html5QrCode = new Html5Qrcode("reader");
  const readerEl = document.getElementById('reader');
  const maxSize = Math.min(360, readerEl.clientWidth || 360);
  const qrSize = Math.max(200, Math.min(300, maxSize - 20));
  
  html5QrCode.start({ facingMode: "environment" }, { fps: 10, qrbox: { width: qrSize, height: qrSize } }, onScanSuccess, () => {})
    .then(() => {
      document.getElementById('scanStatus').textContent = "สแกน QR Code ที่นี่";
      document.getElementById('stopCamBtn').style.display = 'inline-flex';
    })
    .catch((err) => {
      document.getElementById('startCamBtn').style.display = 'inline-flex';
      let errMsg = "เกิดข้อผิดพลาด";
      if (String(err).includes("NotAllowedError") || String(err).includes("Permission denied")) {
        errMsg = "🚫 ถูกปฏิเสธการเข้าถึงกล้อง กรุณาอนุญาตสิทธิ์กล้องในเบราว์เซอร์";
      } else if (String(err).includes("NotFoundError")) {
        errMsg = "❌ ไม่พบกล้องในอุปกรณ์นี้";
      }
      document.getElementById('scanStatus').innerHTML = `<span style="color:var(--danger)">${errMsg}</span>`;
    });
}

function stopScanner() {
  if (html5QrCode) {
    html5QrCode.stop().catch(() => {}).finally(() => {
      document.getElementById('stopCamBtn').style.display = 'none';
      document.getElementById('startCamBtn').style.display = 'inline-flex';
      document.getElementById('scanStatus').textContent = "กดปุ่มด้านล่างเพื่อเปิดกล้อง";
    });
  }
}

function onScanSuccess(decodedText) {
  stopScanner();
  document.getElementById('redeemInput').value = decodedText.toUpperCase();
  doRedeem();
}

function startLoginScanner() {
  const loginReader = document.getElementById('loginReader');
  if (!loginReader) return;
  loginReader.innerHTML = '<div id="loginReaderInner" style="width:100%; min-height:220px;"></div>';
  if (html5QrCodeLogin) { try { html5QrCodeLogin.stop(); } catch(e) {} }
  html5QrCodeLogin = new Html5Qrcode('loginReaderInner');
  const maxSize = Math.min(320, loginReader.clientWidth || 320);
  const qrSize = Math.max(200, Math.min(300, maxSize - 20));
  html5QrCodeLogin.start({ facingMode: 'environment' }, { fps: 10, qrbox: { width: qrSize, height: qrSize } }, loginDecodeSuccess, () => {})
    .then(() => { document.getElementById('stopLoginCamBtn').style.display = 'inline-flex'; })
    .catch(() => { showToast('ไม่สามารถเปิดกล้องได้', 'error'); });
}

function stopLoginScanner() {
  if (html5QrCodeLogin) {
    html5QrCodeLogin.stop().catch(() => {}).finally(() => {
      document.getElementById('stopLoginCamBtn').style.display = 'none';
    });
  }
}

function loginDecodeSuccess(decodedText) {
  stopLoginScanner();
  const loginCodeEl = document.getElementById('loginCode');
  if (loginCodeEl) {
    loginCodeEl.value = decodedText.toUpperCase();
    showToast('สแกนโค้ดสำเร็จ', 'success');
    // Auto-open the confirm modal immediately after scan
    doLoginRedeem();
  }
}

function renderMe() {
  document.getElementById('meName').textContent = u.displayName;
  document.getElementById('meEmail').textContent = u.isGuest ? '(บัญชีชั่วคราว)' : (u.email || '-');
  document.getElementById('meRole').textContent = u.role;
  document.getElementById('meNo').textContent = u.studentNo || '-';
  document.getElementById('meHwCredits').textContent = u.hwCredits || 0;
}

