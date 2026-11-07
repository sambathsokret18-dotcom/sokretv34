/* ==========================================
    FULL SCRIPT.JS - TIME STUDY PRO (UPGRADED)
    ✔ CT / Average time in MM:SS
    ✔ Per-filter Refresh buttons on Dashboard
    ✔ NEW 1: Abnormal Lap detection + Min/Max/Avg/StdDev
    ✔ NEW 2: Line Balancing (Takt, Bottleneck, Balance Rate) + Yamazumi chart
    ✔ NEW 3: KPI Cards on Dashboard
    ✔ NEW 4: Export Excel (CT per lap) + Export PDF report (print)
    ✔ NEW 5: Keyboard shortcuts, Undo Lap, Beep/Vibrate, Auto-save Draft, Wake Lock
========================================== */

// ---------- IndexedDB storage bootstrap for large Excel data ----------
// Large upload records must not fall back to localStorage when IndexedDB is still opening.
(function initTspLargeStorage() {
    if (!('indexedDB' in window)) {
        window.tspLargeStorageReady = Promise.resolve(false);
        return;
    }
    window.tspLargeStorageCache = window.tspLargeStorageCache || {};
    window.tspLargeStorageReady = new Promise(resolve => {
        let request;
        try { request = indexedDB.open('TimeStudyProLargeStorage', 1); }
        catch (e) { console.error('IndexedDB open failed:', e); resolve(false); return; }
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains('records')) db.createObjectStore('records');
        };
        request.onerror = () => { console.error('IndexedDB unavailable:', request.error); resolve(false); };
        request.onsuccess = () => {
            const db = request.result;
            window.tspLargeStorageDB = db;
            const keys = ['uploadedExcelFiles', 'uploadedExcelTrash'];
            let pending = keys.length;
            const finish = () => {
                if (--pending === 0) {
                    try { window.dispatchEvent(new CustomEvent('tsp:large-storage-ready')); } catch (_) {}
                    resolve(true);
                }
            };
            keys.forEach(key => {
                const tx = db.transaction('records', 'readonly');
                const store = tx.objectStore('records');
                const manifestReq = store.get(key + ':manifest');
                manifestReq.onsuccess = () => {
                    const recordKeys = manifestReq.result;
                    if (Array.isArray(recordKeys)) {
                        const out = new Array(recordKeys.length);
                        if (!recordKeys.length) { window.tspLargeStorageCache[key] = out; finish(); return; }
                        let left = recordKeys.length;
                        recordKeys.forEach((rk, i) => {
                            const getReq = store.get(rk);
                            getReq.onsuccess = () => { out[i] = getReq.result; if (--left === 0) { window.tspLargeStorageCache[key] = out.filter(v => v !== undefined); finish(); } };
                            getReq.onerror = () => { if (--left === 0) { window.tspLargeStorageCache[key] = out.filter(v => v !== undefined); finish(); } };
                        });
                    } else {
                        // Migrate any existing localStorage array without deleting it until the DB write succeeds.
                        let legacy = [];
                        try { legacy = JSON.parse(localStorage.getItem(key) || '[]'); } catch (_) {}
                        if (!Array.isArray(legacy)) legacy = [];
                        const write = db.transaction('records', 'readwrite');
                        const ws = write.objectStore('records');
                        const prefix = key + ':';
                        const recordKeysNew = legacy.map((item, i) => prefix + tspRecordKey(item, i));
                        legacy.forEach((item, i) => ws.put(item, recordKeysNew[i]));
                        ws.put(recordKeysNew, key + ':manifest');
                        write.oncomplete = () => {
                            window.tspLargeStorageCache[key] = legacy;
                            try { localStorage.removeItem(key); } catch (_) {}
                            finish();
                        };
                        write.onerror = write.onabort = () => {
                            console.error('IndexedDB migration failed:', write.error);
                            window.tspLargeStorageCache[key] = legacy;
                            finish();
                        };
                    }
                };
                manifestReq.onerror = () => { window.tspLargeStorageCache[key] = []; finish(); };
            });
        };
    });
})();

// ---------- Global helper functions ----------
// Data ownership guard: viewing/using shared data is allowed; modifying/deleting requires owner or Admin.
function tspCurrentPermissionUser() {
    try {
        const users = JSON.parse(localStorage.getItem('tsp_user_management_users_v1') || '[]');
        const sid = localStorage.getItem('tsp_user_management_session_v1');
        return Array.isArray(users) ? (users.find(u => u && u.id === sid) || null) : null;
    } catch (_) { return null; }
}
function tspCanManageRecord(record) {
    const user = tspCurrentPermissionUser();
    if (!user) return false;
    if (user.role === 'Admin') return true;
    if (!record || !user.id) return false;
    const owner = record.ownerId || record.uploaderId || record.createdById || record.userId;
    return !!owner && String(owner) === String(user.id);
}
function tspPermissionDenied(record) {
    const ownerName = record && (record.uploaderName || record.createdBy || record.user || record.sewerName) ? (record.uploaderName || record.createdBy || record.user || record.sewerName) : 'User ផ្សេង';
    if (typeof showCustomAlert === 'function') showCustomAlert('⛔ មិនអាចលុប ឬកែប្រែបានទេ', 'ទិន្នន័យនេះមិនមែនជារបស់អ្នកទេ។ ម្ចាស់ទិន្នន័យ៖ ' + ownerName + '. អ្នកអាចមើល និងប្រើបាន ប៉ុន្តែមិនអាចលុប ឬកែប្រែបានទេ។ Admin មានសិទ្ធិគ្រប់យ៉ាង។', 'fa-solid fa-shield-halved', '#dc2626');
    else if (typeof toast === 'function') toast('គ្មានសិទ្ធិកែប្រែ ឬលុបទិន្នន័យនេះទេ។', 'error');
}

function escapeHtml(v) {
    return String(v === undefined || v === null ? '' : v)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function safeLoad(key) {
    try {
        if ((key === 'uploadedExcelFiles' || key === 'uploadedExcelTrash') && window.tspLargeStorageCache && Object.prototype.hasOwnProperty.call(window.tspLargeStorageCache, key)) {
            const cached = window.tspLargeStorageCache[key];
            // Return a detached copy. Callers frequently edit arrays before safeSave();
            // sharing the IndexedDB cache here destroys the old snapshot, so deleted
            // records cannot be detected and removed from IndexedDB on persistence.
            return Array.isArray(cached) ? JSON.parse(JSON.stringify(cached)) : [];
        }
        const v = JSON.parse(localStorage.getItem(key));
        return Array.isArray(v) ? v : [];
    } catch (e) {
        return [];
    }
}

function safeSave(key, value) {
    try {
        if (key === 'uploadedExcelFiles' || key === 'uploadedExcelTrash') {
            // Do not replace the cache before persistence: IndexedDB needs the old
            // cached records to detect which entries actually changed.
            tspPersistLargeStorage(key, value).then(ok => {
                if (ok && window.tspLargeStorageCache) window.tspLargeStorageCache[key] = value;
                if (!ok && typeof showCustomAlert === 'function') showCustomAlert('កំហុសរក្សាទុក', 'Browser មិនអាចរក្សាទុកទិន្នន័យបានទេ។ សូមពិនិត្យទំហំផ្ទុកទំនេរ។', 'fa-solid fa-triangle-exclamation', '#dc2626');
            });
            return true;
        }
        localStorage.setItem(key, JSON.stringify(value));
        return true;
    } catch (e) {
        console.error('localStorage error:', e);
        if (typeof showCustomAlert === 'function') {
            showCustomAlert('កំហុសឆ្គង', 'មិនអាចរក្សាទុកទិន្នន័យបានទេ (Storage ពេញ)!', 'fa-solid fa-triangle-exclamation', '#dc2626');
        }
        return false;
    }
}

// Store uploaded-file records individually in IndexedDB. Rewriting one giant array on
// every upload can temporarily require space for both the old and new copies.
function tspRecordKey(item, index) {
    const stable = item && (item.fileHash || item.id || (item.name && item.uploadedAt ? item.name + '|' + item.uploadedAt : ''));
    return 'item:' + encodeURIComponent(String(stable || ('record-' + index)));
}
async function tspPersistLargeStorage(key, value) {
    if (['uploadedExcelFiles', 'uploadedExcelTrash'].includes(key) && !window.tspLargeStorageDB && window.tspLargeStorageReady) {
        await window.tspLargeStorageReady;
    }
    if (!['uploadedExcelFiles', 'uploadedExcelTrash'].includes(key) || !window.tspLargeStorageDB) {
        try { localStorage.setItem(key, JSON.stringify(value)); return Promise.resolve(true); }
        catch (e) { console.error('Storage write failed:', e); return Promise.resolve(false); }
    }
    const db = window.tspLargeStorageDB;
    const previous = (window.tspLargeStorageCache && window.tspLargeStorageCache[key]) || [];
    const prefix = key + ':';
    const oldMap = new Map(previous.map((item, i) => [prefix + tspRecordKey(item, i), item]));
    const nextMap = new Map(value.map((item, i) => [prefix + tspRecordKey(item, i), item]));
    return new Promise(resolve => {
        let tx;
        try {
            tx = db.transaction('records', 'readwrite');
            const store = tx.objectStore('records');
            oldMap.forEach((_, recordKey) => { if (!nextMap.has(recordKey)) store.delete(recordKey); });
            nextMap.forEach((item, recordKey) => {
                const old = oldMap.get(recordKey);
                // Write only records that are new or actually changed.
                if (!old || JSON.stringify(old) !== JSON.stringify(item)) store.put(item, recordKey);
            });
            store.put(Array.from(nextMap.keys()), key + ':manifest');
            tx.oncomplete = () => {
                if (window.tspLargeStorageCache) window.tspLargeStorageCache[key] = value;
                try { localStorage.removeItem(key); } catch (_) {}
                resolve(true);
            };
            tx.onerror = tx.onabort = () => {
                console.error('IndexedDB record write failed:', tx.error);
                resolve(false);
            };
        } catch (e) { console.error('IndexedDB write failed:', e); resolve(false); }
    });
}

function loadObject(key) {
    try {
        const v = JSON.parse(localStorage.getItem(key));
        return (v && typeof v === 'object' && !Array.isArray(v)) ? v : {};
    } catch (e) {
        return {};
    }
}

const normStr = (x) => String(x === undefined || x === null ? '' : x).trim().toLowerCase();
const validSewerName = (s) => (s && s !== 'undefined' && String(s).trim() !== '') ? String(s).trim() : 'No Sewer';

// isMs = true  → input in Milliseconds
// isMs = false → input in Seconds or String ("00:00.00", "45s", 59.6)
function formatToMinSec(val, isMs = false) {
    const ZERO = '00:00';
    if (val === undefined || val === null || val === '') return ZERO;

    let totalSeconds = 0;
    if (typeof val === 'number') {
        if (!isFinite(val) || val <= 0) return ZERO;
        totalSeconds = isMs ? val / 1000 : val;
    } else if (typeof val === 'string') {
        const cleanVal = val.trim().toLowerCase().replace('s', '');
        if (cleanVal.includes(':')) {
            const parts = cleanVal.split(':');
            totalSeconds = ((parseFloat(parts[0]) || 0) * 60) + (parseFloat(parts[1]) || 0);
        } else {
            totalSeconds = parseFloat(cleanVal) || 0;
        }
    } else {
        return ZERO;
    }

    totalSeconds = Math.round(totalSeconds);
    if (!isFinite(totalSeconds) || totalSeconds <= 0) return ZERO;

    const m = Math.floor(totalSeconds / 60);
    const s = totalSeconds % 60;
    return String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0');
}

function parseTimeToSeconds(str) {
    const parts = String(str === undefined || str === null ? '0' : str).replace('s', '').trim().split(':');
    return parts.length === 2
        ? (parseFloat(parts[0]) || 0) * 60 + (parseFloat(parts[1]) || 0)
        : (parseFloat(parts[0]) || 0);
}

// Does a saved study belong to one exact process row?
// A Sewer/ID Card can be shared by several processes in the same group, so
// PrimaryKey alone must NEVER mark every process as completed or overwrite them.
function studyMatchesItem(s, item, ctx) {
    if (!s || !item) return false;
    ctx = ctx || {};

    // Compare group context when both sides have a meaningful value.
    for (const key of ['title', 'section', 'styleName', 'rampupDay']) {
        const a = normStr(s[key]);
        const b = normStr(item[key] || ctx[key]);
        if (a && b && a !== 'n/a' && b !== 'n/a' && a !== b) return false;
    }

    // Process and Prep/Assembly are essential row-level identity fields.
    // They are checked BEFORE PrimaryKey because one Sewer ID can cover many processes.
    const sProcess = normStr(s.processRegion || s.processName);
    const iProcess = normStr(item.processRegion || item.processName || ctx.processRegion || ctx.processName);
    if (!sProcess || !iProcess || sProcess !== iProcess) return false;

    const sPrep = normStr(s.prepAssembly || s.prepName || s.prep);
    const iPrep = normStr(item.prepAssembly || item.prepName || item.prep || ctx.prepAssembly || ctx.prepName || ctx.prep);
    if (sPrep && iPrep && sPrep !== iPrep) return false;

    // If both rows have a stable row ID, it must also agree; a shared ID Card
    // is not sufficient to make two different processes the same study.
    const sPk = normStr(s.primaryKey);
    const iPk = normStr(item.primaryKey || ctx.primaryKey);
    if (sPk && sPk !== 'n/a' && iPk && iPk !== 'n/a' && sPk !== iPk) return false;

    // Sewer name is an attribute, not a process identifier. If both are known,
    // mismatches mean the records must remain separate.
    const sSewer = normStr(validSewerName(s.sewerName));
    const iSewer = normStr(validSewerName(item.sewerName || ctx.sewerName));
    if (sSewer && iSewer && sSewer !== iSewer) return false;

    return true;
}

function pfmColors(pfmVal) {
    if (pfmVal < 85) return { bg: '#fee2e2', color: '#dc2626' };
    if (pfmVal < 100) return { bg: '#fef3c7', color: '#d97706' };
    return { bg: '#e0f2fe', color: '#0284c7' };
}

// Lap statistics: use only valid positive finite durations.
// With 1-2 laps, calculate the available mean but label it as a small sample.
function calcLapStats(laps) {
    const source = Array.isArray(laps) ? laps : [];
    const durations = source
        .map(l => Number(l && l.durationMs))
        .filter(d => Number.isFinite(d) && d > 0);
    const n = durations.length;
    if (n === 0) return { n: 0, avg: 0, min: 0, max: 0, std: 0, outliers: [], deviations: [], sampleStatus: 'none' };
    const sum = durations.reduce((a, b) => a + b, 0);
    const avg = sum / n;
    const min = Math.min(...durations);
    const max = Math.max(...durations);
    const variance = durations.reduce((a, d) => a + Math.pow(d - avg, 2), 0) / n;
    const std = Math.sqrt(variance);
    const deviations = durations.map(d => avg > 0 ? (d - avg) / avg : 0);
    // Do not flag laps as abnormal based on a 25% deviation threshold.
    const outliers = [];
    const sampleStatus = n < 3 ? 'small' : n < 5 ? 'limited' : 'ok';
    return { n, avg, min, max, std, outliers, deviations, sampleStatus };
}

// ✨ NEW 5: sound / vibration / wake lock
let audioCtx = null;
function beep(freq = 880, duration = 70) {
    try {
        const AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) return;
        if (!audioCtx) audioCtx = new AC();
        const osc = audioCtx.createOscillator();
        const gain = audioCtx.createGain();
        osc.type = 'sine';
        osc.frequency.value = freq;
        gain.gain.value = 0.08;
        osc.connect(gain);
        gain.connect(audioCtx.destination);
        osc.start();
        osc.stop(audioCtx.currentTime + duration / 1000);
    } catch (e) { /* ignore */ }
}
function vibrate(ms = 40) {
    try { if (navigator.vibrate) navigator.vibrate(ms); } catch (e) { /* ignore */ }
}

let wakeLockSentinel = null;
async function requestWakeLock() {
    try {
        if ('wakeLock' in navigator && !wakeLockSentinel) {
            wakeLockSentinel = await navigator.wakeLock.request('screen');
            wakeLockSentinel.addEventListener('release', () => { wakeLockSentinel = null; });
        }
    } catch (e) { wakeLockSentinel = null; }
}
function releaseWakeLock() {
    try {
        if (wakeLockSentinel) { wakeLockSentinel.release(); wakeLockSentinel = null; }
    } catch (e) { wakeLockSentinel = null; }
}

// Confirm dialog (SweetAlert2 if available, otherwise custom confirm)
function confirmDialog(opts) {
    return new Promise((resolve) => {
        if (typeof Swal !== 'undefined') {
            Swal.fire({
                title: '<span style="font-size: 20px; font-weight: bold;">' + opts.title + '</span>',
                html: '<span style="font-size: 13px; color: #64748b;">' + opts.html + '</span>',
                icon: opts.icon || 'warning',
                width: '280px',
                padding: '0.5rem',
                showCancelButton: true,
                confirmButtonColor: opts.confirmColor || '#d33',
                cancelButtonColor: '#3085d6',
                confirmButtonText: '<span style="font-size: 13px;">' + (opts.confirmText || 'យល់ព្រម') + '</span>',
                cancelButtonText: '<span style="font-size: 13px;">បោះបង់</span>',
                customClass: { popup: 'custom-swal-popup', icon: 'custom-swal-icon' }
            }).then((r) => resolve(!!r.isConfirmed));
        } else {
            showCustomConfirm(
                opts.title.replace(/<[^>]*>/g, ''),
                opts.html.replace(/<[^>]*>/g, ''),
                () => resolve(true),
                opts.confirmText || 'យល់ព្រម',
                () => resolve(false)
            );
        }
    });
}

function successDialog(title, html) {
    if (typeof Swal !== 'undefined') {
        Swal.fire({
            title: '<span style="font-size: 15px;">' + title + '</span>',
            html: '<span style="font-size: 12px;">' + html + '</span>',
            icon: 'success',
            width: '280px',
            timer: 1500,
            showConfirmButton: false,
            customClass: { popup: 'custom-swal-popup' }
        });
    } else {
        showToast(title, html, 'fa-solid fa-circle-check', '#10b981');
    }
}

document.addEventListener('DOMContentLoaded', () => {

    // Mobile-only sidebar drawer; desktop navigation remains unchanged.
    const mobileMenuToggle = document.getElementById('mobile-menu-toggle');
    const mobileSidebarBackdrop = document.getElementById('mobile-sidebar-backdrop');
    const appSidebar = document.getElementById('app-sidebar');
    const closeMobileSidebar = () => {
        if (appSidebar) appSidebar.classList.remove('mobile-sidebar-open');
        if (mobileSidebarBackdrop) mobileSidebarBackdrop.hidden = true;
        if (mobileMenuToggle) {
            mobileMenuToggle.setAttribute('aria-expanded', 'false');
            mobileMenuToggle.innerHTML = '<i class="fa-solid fa-bars"></i>';
            mobileMenuToggle.setAttribute('aria-label', 'បើកម៉ឺនុយ');
        }
    };
    if (mobileMenuToggle && appSidebar && mobileSidebarBackdrop) {
        mobileMenuToggle.addEventListener('click', () => {
            const open = !appSidebar.classList.contains('mobile-sidebar-open');
            appSidebar.classList.toggle('mobile-sidebar-open', open);
            mobileSidebarBackdrop.hidden = !open;
            mobileMenuToggle.setAttribute('aria-expanded', String(open));
            mobileMenuToggle.innerHTML = open ? '<i class="fa-solid fa-xmark"></i>' : '<i class="fa-solid fa-bars"></i>';
            mobileMenuToggle.setAttribute('aria-label', open ? 'បិទម៉ឺនុយ' : 'បើកម៉ឺនុយ');
        });
        mobileSidebarBackdrop.addEventListener('click', closeMobileSidebar);
        document.addEventListener('keydown', e => { if (e.key === 'Escape') closeMobileSidebar(); });
    }

    // --- ១. ប្ដូរទំព័រពេលចុចលើ Sidebar Menu ---
    const menuLinks = document.querySelectorAll('.sidebar ul li a');
    const pages = document.querySelectorAll('.page-content');

    menuLinks.forEach(link => {
        link.addEventListener('click', function (e) {
            e.preventDefault();
            closeMobileSidebar();

            menuLinks.forEach(item => item.classList.remove('active'));
            this.classList.add('active');

            const targetPageId = this.getAttribute('data-page');
            const currentPageLabel = document.getElementById('tsp-current-page-name');
            if (currentPageLabel) {
                const pageLabel = (this.textContent || '').replace(/\s+/g, ' ').trim();
                currentPageLabel.textContent = pageLabel || (targetPageId || '').replace(/-page$/, '').replace(/-/g, ' ');
            }

            pages.forEach(page => {
                if (page.id === targetPageId) {
                    page.style.display = 'block';
                    if (targetPageId === 'time-page') {
                        ['section-page', 'style-page', 'rampup-page', 'prep-page', 'process-page', 'timestudy-page'].forEach(pId => {
                            const p = document.getElementById(pId);
                            if (p) p.style.display = 'none';
                        });
                        renderMasterSectList();
                    } else if (targetPageId === 'dashboard-page') {
                        renderDashboardStats();
                    }
                } else {
                    page.style.display = 'none';
                }
            });
        });
    });



    // Mobile Glass navigation: handle page changes directly so taps remain reliable on phones.
    const mobileBottomLinks = document.querySelectorAll('.mobile-bottom-nav a[data-page]');
    const activateMobilePage = (targetPageId) => {
        if (!targetPageId) return;
        const targetLink = document.querySelector('.sidebar ul li a[data-page="' + targetPageId + '"]');
        const targetPage = document.getElementById(targetPageId);
        if (!targetPage) return;
        // Dismiss the Welcome overlay before showing the selected destination.
        const welcomeOverlay = document.getElementById('tsp-home-welcome');
        if (welcomeOverlay) welcomeOverlay.hidden = true;
        // Keep the original sidebar's active state and page label in sync.
        menuLinks.forEach(item => item.classList.toggle('active', item.getAttribute('data-page') === targetPageId));
        const currentPageLabel = document.getElementById('tsp-current-page-name');
        if (currentPageLabel) currentPageLabel.textContent = targetLink ? (targetLink.textContent || '').replace(/\s+/g, ' ').trim() : targetPageId.replace(/-page$/, '').replace(/-/g, ' ');
        pages.forEach(page => { page.style.display = page.id === targetPageId ? 'block' : 'none'; });
        if (targetPageId === 'time-page') {
            ['section-page', 'style-page', 'rampup-page', 'prep-page', 'process-page', 'timestudy-page'].forEach(pId => {
                const p = document.getElementById(pId);
                if (p) p.style.display = 'none';
            });
            if (typeof renderMasterSectList === 'function') renderMasterSectList();
        } else if (targetPageId === 'dashboard-page' && typeof renderDashboardStats === 'function') {
            renderDashboardStats();
        }
        mobileBottomLinks.forEach(item => item.classList.toggle('active', item.getAttribute('data-page') === targetPageId));
    };
    const mobileMoreToggle = document.getElementById('mobile-more-toggle');
    const mobileMoreSheet = document.getElementById('mobile-more-sheet');
    const mobileMoreBackdrop = document.getElementById('mobile-more-backdrop');
    const closeMobileMore = () => {
        if (mobileMoreSheet) { mobileMoreSheet.classList.remove('open'); mobileMoreSheet.setAttribute('aria-hidden', 'true'); }
        if (mobileMoreBackdrop) mobileMoreBackdrop.hidden = true;
        if (mobileMoreToggle) mobileMoreToggle.setAttribute('aria-expanded', 'false');
    };
    const openMobileMore = () => {
        if (mobileMoreSheet) { mobileMoreSheet.classList.add('open'); mobileMoreSheet.setAttribute('aria-hidden', 'false'); }
        if (mobileMoreBackdrop) mobileMoreBackdrop.hidden = false;
        if (mobileMoreToggle) mobileMoreToggle.setAttribute('aria-expanded', 'true');
    };
    mobileBottomLinks.forEach(mobileLink => {
        mobileLink.addEventListener('click', function (e) {
            e.preventDefault();
            const targetPageId = this.getAttribute('data-page');
            activateMobilePage(targetPageId);
            closeMobileMore();
        });
    });
    if (mobileMoreToggle) mobileMoreToggle.addEventListener('click', () => mobileMoreSheet && mobileMoreSheet.classList.contains('open') ? closeMobileMore() : openMobileMore());
    if (mobileMoreBackdrop) mobileMoreBackdrop.addEventListener('click', closeMobileMore);
    const mobileMoreClose = document.getElementById('mobile-more-close');
    if (mobileMoreClose) mobileMoreClose.addEventListener('click', closeMobileMore);
    document.addEventListener('keydown', e => { if (e.key === 'Escape') closeMobileMore(); });
    if (mobileMoreSheet) mobileMoreSheet.querySelectorAll('[data-mobile-action]').forEach(btn => btn.addEventListener('click', () => {
        const action = btn.dataset.mobileAction;
        if (action === 'page') {
            activateMobilePage(btn.dataset.page);
        } else if (action === 'profile') {
            const profile = document.getElementById('profile-btn');
            if (profile) {
                profile.click();
                const profileMenu = document.getElementById('tsp-profile-menu');
                const profileAction = profileMenu && profileMenu.querySelector('[data-action="profile"], [data-action="login"]');
                if (profileAction) profileAction.click();
            }
        } else if (action === 'settings') {
            const profile = document.getElementById('profile-btn');
            if (profile) profile.click();
            const settingsAction = document.querySelector('#tsp-profile-menu [data-action="workspace-settings"]');
            if (settingsAction) settingsAction.click();
            else {
                const workspaceSettings = document.querySelector('#tsp-profile-menu [data-action="smart-tools"]');
                if (workspaceSettings) workspaceSettings.click();
                else alert('សូមបើក Profile Menu ដើម្បីចូល Settings។');
            }
        } else if (action === 'search') {
            const uploadLink = document.querySelector('.sidebar ul li a[data-page="upload-page"]');
            if (uploadLink) uploadLink.click();
            const uploadSearch = document.getElementById('uploaded-file-search');
            if (uploadSearch) { window.setTimeout(() => { uploadSearch.focus(); uploadSearch.scrollIntoView({behavior:'smooth',block:'center'}); }, 80); }
        } else if (action === 'filters') {
            const dashboardLink = document.querySelector('.sidebar ul li a[data-page="dashboard-page"]');
            if (dashboardLink) dashboardLink.click();
            const dashboard = document.getElementById('dashboard-page');
            if (dashboard) { dashboard.scrollIntoView({behavior:'smooth',block:'start'}); const firstSelect = dashboard.querySelector('select'); if (firstSelect) firstSelect.focus({preventScroll:true}); }
        }
        closeMobileMore();
    }));

    // --- ២. Upload និង Excel Processing ---
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('excel-file');
    const fileInfo = document.getElementById('file-info');
    const fileNameSpan = document.getElementById('file-name');
    const removeFileBtn = document.getElementById('remove-file');
    const uploadForm = document.getElementById('upload-form');

    let uploadedFilesList = safeLoad('uploadedExcelFiles');
    let selectedFile = null;
    const selectedUploadIndexes = new Set();
    let uploadViewMode = localStorage.getItem('tspUploadViewMode') === 'grid' ? 'grid' : 'list';
    let uploadedFilesTrash = safeLoad('uploadedExcelTrash');
    window.addEventListener('tsp:large-storage-ready', () => {
        uploadedFilesList = safeLoad('uploadedExcelFiles');
        uploadedFilesTrash = safeLoad('uploadedExcelTrash');
        if (typeof renderTable === 'function') renderTable();
        if (typeof updateTotalFilesCard === 'function') updateTotalFilesCard();
    });
    let uploadMonthFilter = ''; // YYYY-MM; empty means all months
    let uploadDateFilter = ''; // YYYY-MM-DD; when set, show files from the exact selected day
    let uploadActivityLog = safeLoad('tspUploadActivityLog');
    function logUploadActivity(action, item, detail = '') {
        const user = getUploadUser();
        uploadActivityLog.unshift({ action, name: item && item.name ? item.name : '—', user: user.fullName || user.username, username: user.username, avatar: user.avatar || '', detail, at: new Date().toISOString() });
        uploadActivityLog = uploadActivityLog.slice(0, 300);
        safeSave('tspUploadActivityLog', uploadActivityLog);
        try { window.dispatchEvent(new CustomEvent('tsp:activity-updated', { detail: uploadActivityLog[0] })); } catch (_) {}
    }
    function saveUploadTrash() { return safeSave('uploadedExcelTrash', uploadedFilesTrash); }
    function canManageUploadFiles(item) {
        const user = getUploadUser();
        return user.role === 'Admin' || (!!user.permissions?.deleteOwn && !!user.id && !!item && !!item.uploaderId && user.id === item.uploaderId);
    }
    function safeUploadHtml(value) { return escapeHtml(value); }
    function fileSnapshot(item) {
        const copy = JSON.parse(JSON.stringify(item));
        delete copy.versionHistory;
        delete copy.lastVersionSnapshotAt;
        return copy;
    }
    const uploadProgressWrap = document.getElementById('upload-progress-wrap');
    const uploadProgress = document.getElementById('upload-progress');
    const uploadProgressText = document.getElementById('upload-progress-text');
    const uploadProgressPercent = document.getElementById('upload-progress-percent');
    function setUploadProgress(value, label) {
        const pct = Math.max(0, Math.min(100, Number(value) || 0));
        if (uploadProgressWrap) uploadProgressWrap.hidden = false;
        if (uploadProgress) uploadProgress.value = pct;
        if (uploadProgressPercent) uploadProgressPercent.textContent = Math.round(pct) + '%';
        if (uploadProgressText && label) uploadProgressText.textContent = label;
    }
    function hideUploadProgress() { if (uploadProgressWrap) uploadProgressWrap.hidden = true; }
    const TSP_USERS_KEY = 'tsp_user_management_users_v1';
    const TSP_SESSION_KEY = 'tsp_user_management_session_v1';
    function getUploadUser() {
        try {
            const users = JSON.parse(localStorage.getItem(TSP_USERS_KEY) || '[]');
            const sid = localStorage.getItem(TSP_SESSION_KEY);
            const user = Array.isArray(users) ? users.find(u => u.id === sid) : null;
            if (user) return { id: user.id, username: user.username || user.fullName || 'User', fullName: user.fullName || user.username || 'User', avatar: user.avatar || '', role: user.role || 'User', status: user.status || 'active', permissions: user.permissions || { upload: true, download: true, deleteOwn: true } };
        } catch (_) {}
        return { id: '', username: 'SOKRET', fullName: 'SOKRET', avatar: '', role: 'User' };
    }
    function normalizeFileName(name) { return String(name || '').trim().normalize('NFC').toLocaleLowerCase(); }
    function findDuplicateFile(name) {
        const wanted = normalizeFileName(name);
        return uploadedFilesList.find(item => normalizeFileName(item.name) === wanted);
    }
    function formatFileSize(bytes) {
        const n = Number(bytes) || 0;
        if (n < 1024) return n + ' B';
        if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
        return (n / (1024 * 1024)).toFixed(2) + ' MB';
    }

    function resetFileInput() {
        selectedFile = null;
        if (fileInput) fileInput.value = '';
        if (fileInfo) fileInfo.style.display = 'none';
        if (dropZone) dropZone.style.display = 'flex';
        hideUploadProgress();
    }

    function validateAndProcessFile(file) {
        if (!file) return;
        const fileName = file.name.toLowerCase();
        const duplicate = findDuplicateFile(file.name);
        if (duplicate) {
            showCustomAlert('ឯកសារមានរួចហើយ', 'មិនអាច Upload បានទេ ព្រោះមាន File ឈ្មោះ [' + duplicate.name + '] រួចហើយ។ សូមប្តូរឈ្មោះ File មុនសិន។', 'fa-solid fa-copy', '#dc2626');
            resetFileInput();
            return;
        }

        if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
            showCustomAlert('សូមអភ័យទោស!', 'ប្រព័ន្ធអនុញ្ញាតឱ្យ Upload តែឯកសារ Excel (.xlsx ឬ .xls) ប៉ុណ្ណោះ។', 'fa-solid fa-triangle-exclamation', '#f59e0b');
            resetFileInput();
            return;
        }

        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (jsonData.length === 0) {
                    showCustomAlert('ទម្រង់មិនត្រឹមត្រូវ', 'ហ្វាល Excel នេះទទេ គ្មានទិន្នន័យទេ!', 'fa-solid fa-triangle-exclamation', '#dc2626');
                    resetFileInput();
                    return;
                }

                const fileHeaders = jsonData[0].map(h => String(h).trim().toLowerCase());
                const requiredHeaders = ['title', 'primarykey', 'smv'];
                const missingHeaders = requiredHeaders.filter(req => !fileHeaders.some(fh => fh.includes(req)));

                if (missingHeaders.length > 0) {
                    showCustomAlert('ខុសទម្រង់ក្បាលតារាង', 'ហ្វាល Excel ខ្វះក្បាលតារាងចាំបាច់ (Title, PrimaryKey, SMV)', 'fa-solid fa-triangle-exclamation', '#dc2626');
                    resetFileInput();
                } else {
                    selectedFile = file;
                    hideUploadProgress();
                    fileInfo.style.display = 'flex';
                    dropZone.style.display = 'none';
                    fileNameSpan.textContent = file.name;
                }
            } catch (error) {
                console.error('Excel Error:', error);
                showCustomAlert('កំហុសឆ្គង', 'មិនអាចអានហ្វាល Excel នេះបានទេ: ' + error.message, 'fa-solid fa-triangle-exclamation', '#dc2626');
                resetFileInput();
            }
        };
        reader.readAsArrayBuffer(file);
    }

    if (dropZone && fileInput) {
        dropZone.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) validateAndProcessFile(e.target.files[0]);
        });

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#2563eb';
            dropZone.style.backgroundColor = '#eff6ff';
        });

        dropZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#cbd5e1';
            dropZone.style.backgroundColor = '#f8fafc';
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#cbd5e1';
            dropZone.style.backgroundColor = '#f8fafc';
            if (e.dataTransfer.files.length > 1) { showCustomAlert('ជ្រើសរើស File មួយ', 'សូមជ្រើសរើស Excel មួយក្នុងមួយដង ដើម្បីពិនិត្យទិន្នន័យឱ្យបានត្រឹមត្រូវ។', 'fa-solid fa-circle-info', '#2563eb'); return; }
            if (e.dataTransfer.files.length > 0) validateAndProcessFile(e.dataTransfer.files[0]);
        });
    }

    if (removeFileBtn) removeFileBtn.addEventListener('click', () => resetFileInput());

    if (uploadForm) {
        uploadForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const uploadUser = getUploadUser();
            if (uploadUser.status === 'disabled' || (uploadUser.role !== 'Admin' && uploadUser.permissions?.upload === false)) { showCustomAlert('គ្មានសិទ្ធិ', 'គណនីនេះមិនមានសិទ្ធិ Upload ឯកសារទេ។', 'fa-solid fa-lock', '#dc2626'); return; }
            const file = selectedFile || (fileInput && fileInput.files[0]);
            if (!file) {
                showCustomAlert('ជូនដំណឹង', 'សូមជ្រើសរើសឯកសារ Excel ជាមុនសិន!', 'fa-solid fa-circle-info', '#3b82f6');
                return;
            }

            const duplicate = findDuplicateFile(file.name);
            if (duplicate) {
                showCustomAlert('ឯកសារមានរួចហើយ', 'មិនអាច Upload បានទេ ព្រោះមាន File ឈ្មោះ [' + duplicate.name + '] រួចហើយ។', 'fa-solid fa-copy', '#dc2626');
                resetFileInput();
                return;
            }

            setUploadProgress(5, 'កំពុងអាន និងពិនិត្យ File...');
            const reader = new FileReader();
            reader.onprogress = function (ev) { if (ev.lengthComputable) setUploadProgress(5 + (ev.loaded / ev.total) * 20, 'កំពុងអាន File...'); };
            reader.onload = async function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    setUploadProgress(30, 'កំពុងពិនិត្យមាតិកា Excel...');
                    let fileHash = '';
                    if (window.crypto && crypto.subtle && e.target.result) {
                        const digest = await crypto.subtle.digest('SHA-256', e.target.result);
                        fileHash = Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
                        const sameContent = uploadedFilesList.find(item => item.fileHash && item.fileHash === fileHash);
                        if (sameContent) { showCustomAlert('មាតិកា File ស្ទួន', 'File នេះមានមាតិកាដូចគ្នានឹង [' + sameContent.name + '] ដែលមានក្នុងប្រព័ន្ធរួចហើយ។', 'fa-solid fa-copy', '#dc2626'); resetFileInput(); return; }
                    }
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

                    const rawArrayData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    const rawHeaders = rawArrayData.length > 0 ? rawArrayData[0].map(h => String(h).trim()) : [];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);

                    const titlesSet = new Set();
                    const sectionsMap = {};
                    const stylesMap = {};
                    const rampupMap = {};
                    const prepMap = {};
                    const stamp = Date.now();

                    jsonData.forEach((row, index) => {
                        const normalizeHeader = (value) => String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
                        const getRowVal = (keys) => {
                            const normalizedAliases = keys.map(normalizeHeader);
                            const rowKeys = Object.keys(row);
                            // Prefer exact header matches. For general fields, permit descriptive prefixes,
                            // but never let identity/card columns be treated as a Sewer field.
                            let foundKey = rowKeys.find(k => normalizedAliases.includes(normalizeHeader(k)));
                            if (!foundKey) {
                                foundKey = rowKeys.find(k => {
                                    const normalizedKey = normalizeHeader(k);
                                    const isIdentityColumn = /(idcard|cardid|employeeid|operatorid|staffid|badgeid|primarykey|identitycard)/.test(normalizedKey);
                                    if (isIdentityColumn) return false;
                                    return normalizedAliases.some(alias => alias && normalizedKey.startsWith(alias));
                                });
                            }
                            return foundKey && row[foundKey] !== undefined && row[foundKey] !== null
                                ? String(row[foundKey]).trim() : '';
                        };

                        const titleVal = getRowVal(['title', 'factory']) || 'General Factory';
                        const sectionVal = getRowVal(['section']) || 'General Section';
                        const styleVal = getRowVal(['style', 'stylename']) || 'General Style';
                        const rampupVal = getRowVal(['rampupd', 'rampupday', 'ramp_up_day', 'rampup', 'day']) || 'Day 1';
                        // Read Group only from the Excel Group column; keep the original value (Prep, Ass, Ass1, Ass2, etc.).
                        const groupVal = getRowVal(['group']);
                        const primaryKeyVal = getRowVal(['primarykey', 'primary_key', 'idcard', 'cardid', 'employeeid']) || 'N/A';
                        // Sewer is read only from a column explicitly named Sewer / Sewer Name.
                        // Never use an ID Card / employee ID column as a fallback for the operator's name.
                        const sewerHeader = Object.keys(row).find(k => {
                            const h = normalizeHeader(k);
                            const identityColumn = /(idcard|cardid|employeeid|operatorid|staffid|badgeid|primarykey|identitycard)/.test(h);
                            return !identityColumn && (h === 'sewer' || h === 'sewername' || h.startsWith('sewername') || h.startsWith('seweroperator'));
                        });
                        const sewerNameVal = (sewerHeader && row[sewerHeader] != null && String(row[sewerHeader]).trim())
                            ? String(row[sewerHeader]).trim()
                            : (getRowVal(['sewername', 'sewer']) || 'No Sewer');
                        const prepAssemblyVal = getRowVal(['prep_assembly', 'prepassembly', 'prep', 'assembly', 'prep/assembly']) || 'Detail Item';

                        const processKhmerVal = getRowVal(['process(region)', 'processregion', 'process_region', 'process_khmer', 'khmer']) || '';
                        const processThaiVal = getRowVal(['process', 'processname', 'operation', 'operationname', 'thai']) || '';

                        const finalProcessRegion = processKhmerVal ? processKhmerVal : (processThaiVal || 'General Process');
                        const finalProcessName = processThaiVal ? processThaiVal : processKhmerVal;

                        const noVal = getRowVal(['no', 'number']) || String(index + 1);
                        const seqVal = getRowVal(['seq', 'sequence', 'opseq']) || '99';
                        const mcVal = getRowVal(['mc', 'machine', 'machinename', 'm/c']) || '-';
                        const weightVal = getRowVal(['weight', 'wt']) || '-';
                        const dateTimeStartVal = getRowVal(['datetime-start', 'datetimestart', 'start_time']) || '';

                        const rawSmv = parseFloat(getRowVal(['smv', 'ngiesmv'])) || 0;
                        const formattedSmv = !isNaN(rawSmv) ? rawSmv.toFixed(2) : '0.00';

                        titlesSet.add(titleVal);

                        if (!sectionsMap[titleVal]) sectionsMap[titleVal] = new Set();
                        sectionsMap[titleVal].add(sectionVal);

                        const styleCompositeKey = `${titleVal}_${sectionVal}`;
                        if (!stylesMap[styleCompositeKey]) stylesMap[styleCompositeKey] = new Set();
                        stylesMap[styleCompositeKey].add(styleVal);

                        const rampupCompositeKey = `${titleVal}_${sectionVal}_${styleVal}`;
                        if (!rampupMap[rampupCompositeKey]) rampupMap[rampupCompositeKey] = new Set();
                        rampupMap[rampupCompositeKey].add(rampupVal);

                        const prepCompositeKey = `${titleVal}_${sectionVal}_${styleVal}_${rampupVal}`;
                        if (!prepMap[prepCompositeKey]) prepMap[prepCompositeKey] = [];

                        prepMap[prepCompositeKey].push({
                            id: 'EXCEL_ROW_' + stamp + '_' + index,
                            title: titleVal,
                            factory: titleVal,
                            section: sectionVal,
                            styleName: styleVal,
                            rampupDay: rampupVal,
                            group: groupVal,
                            primaryKey: primaryKeyVal,
                            sewerName: sewerNameVal,
                            prepAssembly: prepAssemblyVal,
                            processName: finalProcessName,
                            processRegion: finalProcessRegion,
                            smv: formattedSmv,
                            no: noVal,
                            seq: seqVal,
                            operatorSeq: parseInt(seqVal) || 99,
                            mc: mcVal,
                            weight: weightVal,
                            dateTimeStart: dateTimeStartVal,
                            rawRowData: row
                        });
                    });

                    const uniqueTitles = Array.from(titlesSet);
                    const formattedSectionsMap = {};
                    for (const t in sectionsMap) formattedSectionsMap[t] = Array.from(sectionsMap[t]);
                    const formattedStylesMap = {};
                    for (const k in stylesMap) formattedStylesMap[k] = Array.from(stylesMap[k]);
                    const formattedRampupMap = {};
                    for (const k in rampupMap) formattedRampupMap[k] = Array.from(rampupMap[k]);

                    const now = new Date();
                    const dateString = now.toLocaleDateString('km-KH') + ' ' + now.toLocaleTimeString('km-KH', { hour: '2-digit', minute: '2-digit' });

                    const uploadUser = getUploadUser();
                    uploadedFilesList.unshift({
                        user: uploadUser.username,
                        uploaderId: uploadUser.id,
                        uploaderName: uploadUser.fullName,
                        uploaderAvatar: uploadUser.avatar,
                        uploaderRole: uploadUser.role,
                        name: file.name,
                        fileHash: fileHash,
                        fileSize: file.size || 0,
                        fileType: file.name.split('.').pop().toUpperCase(),
                        rowCount: jsonData.length,
                        uploadedAt: now.toISOString(),
                        date: dateString,
                        excelHeaders: rawHeaders,
                        titles: uniqueTitles,
                        sectionsMap: formattedSectionsMap,
                        stylesMap: formattedStylesMap,
                        rampupMap: formattedRampupMap,
                        prepMap: prepMap
                    });

                    const uploadSaved = await tspPersistLargeStorage('uploadedExcelFiles', uploadedFilesList);
                    if (!uploadSaved) {
                        uploadedFilesList.shift();
                        hideUploadProgress();
                        showCustomAlert('រក្សាទុកមិនបានសម្រេច', 'Browser មិនអាចរក្សាទុក File នេះបានទេ។ សូមពិនិត្យទំហំផ្ទុកទំនេរ ឬសាកល្បងបើក TimeStudyPro ក្នុង Browser ដដែល។', 'fa-solid fa-triangle-exclamation', '#dc2626');
                        return;
                    }
                    logUploadActivity('Uploaded', uploadedFilesList[0], 'New Excel file');
                    setUploadProgress(100, 'Upload ជោគជ័យ');
                    renderTable();
                    updateTotalFilesCard();
                    showCustomAlert('ជោគជ័យ', 'ឯកសារ Excel ["' + file.name + '"] ត្រូវបាន Upload ជោគជ័យ!', 'fa-solid fa-circle-check', '#10b981');
                    resetFileInput();
                } catch (err) {
                    console.error('Upload Error:', err);
                    showCustomAlert('កំហុសឆ្គង', 'មិនអាចអានទិន្នន័យពីហ្វាល Excel នេះបានទេ: ' + err.message, 'fa-solid fa-triangle-exclamation', '#dc2626');
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    function renderTable() {
        Array.from(selectedUploadIndexes).forEach(i => { if (i < 0 || i >= uploadedFilesList.length) selectedUploadIndexes.delete(i); });
        const tableBody = document.getElementById('files-table-body');
        const gridContainer = document.getElementById('uploaded-files-grid');
        const fileCountSpan = document.getElementById('file-count');
        const searchInput = document.getElementById('uploaded-file-search');
        const query = searchInput ? searchInput.value.trim().toLocaleLowerCase() : '';
        const typeFilter = document.getElementById('uploaded-file-type-filter')?.value || 'all';
        const sortMode = document.getElementById('uploaded-file-sort')?.value || 'newest';
        const visibleFiles = uploadedFilesList.map((item, index) => ({ item, index })).filter(({ item }) => {
            const uploadedDate = item.uploadedAt ? new Date(item.uploadedAt) : null;
            const monthKey = uploadedDate && !Number.isNaN(uploadedDate.getTime())
                ? `${uploadedDate.getFullYear()}-${String(uploadedDate.getMonth() + 1).padStart(2, '0')}`
                : '';
            const dayKey = uploadedDate && !Number.isNaN(uploadedDate.getTime())
                ? `${uploadedDate.getFullYear()}-${String(uploadedDate.getMonth() + 1).padStart(2, '0')}-${String(uploadedDate.getDate()).padStart(2, '0')}`
                : '';
            return (!query || [item.name, item.user, item.uploaderName, item.fileType].some(v => String(v || '').toLocaleLowerCase().includes(query))) &&
                (typeFilter === 'all' || String(item.name || '').toLowerCase().endsWith('.' + typeFilter)) &&
                (!uploadMonthFilter || monthKey === uploadMonthFilter) &&
                (!uploadDateFilter || dayKey === uploadDateFilter);
        });
        visibleFiles.sort((a, b) => {
            if (sortMode === 'oldest') return new Date(a.item.uploadedAt || 0) - new Date(b.item.uploadedAt || 0);
            if (sortMode === 'name-asc') return String(a.item.name).localeCompare(String(b.item.name));
            if (sortMode === 'name-desc') return String(b.item.name).localeCompare(String(a.item.name));
            if (sortMode === 'size-desc') return (Number(b.item.fileSize) || 0) - (Number(a.item.fileSize) || 0);
            if (sortMode === 'size-asc') return (Number(a.item.fileSize) || 0) - (Number(b.item.fileSize) || 0);
            return new Date(b.item.uploadedAt || 0) - new Date(a.item.uploadedAt || 0);
        });
        if (fileCountSpan) fileCountSpan.textContent = (query || typeFilter !== 'all' || uploadMonthFilter || uploadDateFilter) ? `${visibleFiles.length} / ${uploadedFilesList.length}` : uploadedFilesList.length;
        if (!tableBody) return;
        const viewListBtn = document.getElementById('upload-view-list');
        const viewGridBtn = document.getElementById('upload-view-grid');
        const tableWrap = document.querySelector('#upload-page .table-responsive');
        if (tableWrap) { tableWrap.hidden = uploadViewMode !== 'list'; tableWrap.style.display = uploadViewMode === 'list' ? '' : 'none'; }
        if (gridContainer) gridContainer.hidden = uploadViewMode !== 'grid';
        if (viewListBtn) { viewListBtn.classList.toggle('active', uploadViewMode === 'list'); viewListBtn.setAttribute('aria-pressed', String(uploadViewMode === 'list')); }
        if (viewGridBtn) { viewGridBtn.classList.toggle('active', uploadViewMode === 'grid'); viewGridBtn.setAttribute('aria-pressed', String(uploadViewMode === 'grid')); }
        if (visibleFiles.length === 0) {
            const message = uploadedFilesList.length === 0 ? 'មិនទាន់មានឯកសារត្រូវបាន Upload នៅឡើយទេ' : 'រកមិនឃើញ File ដែលត្រូវនឹងពាក្យស្វែងរកទេ';
            tableBody.innerHTML = `<tr id="empty-row"><td colspan="8" style="text-align:center;color:#9ca3af;padding:24px;">${escapeHtml(message)}</td></tr>`;
            if (gridContainer) gridContainer.innerHTML = `<div class="upload-grid-empty"><i class="fa-solid fa-folder-open"></i><p>${escapeHtml(message)}</p></div>`;
            updateBulkToolbar();
            return;
        }

        const currentUser = getUploadUser();
        let html = '';
        visibleFiles.forEach(({ item, index }, rowIndex) => {
            const canDelete = canManageUploadFiles(item);
            const ownershipBadge = currentUser.role === 'Admin' ? '<span class="upload-owner-badge upload-owner-admin">Admin · គ្រប់គ្រងបាន</span>' : (currentUser.id && item.uploaderId && currentUser.id === item.uploaderId) ? '<span class="upload-owner-badge upload-owner-mine">របស់ខ្ញុំ</span>' : '<span class="upload-owner-badge upload-owner-other">របស់អ្នកដទៃ</span>';
            const avatar = item.uploaderAvatar
                ? `<img class="upload-user-avatar" src="${escapeHtml(item.uploaderAvatar)}" alt="">`
                : `<span class="upload-user-avatar upload-user-avatar-fallback"><i class="fa-solid fa-user"></i></span>`;
            html += `
                <tr>
                    <td><input type="checkbox" class="upload-file-select" data-index="${index}" aria-label="ជ្រើសរើស ${escapeHtml(item.name)}" ${selectedUploadIndexes.has(index) ? 'checked' : ''}></td>
                    <td>${rowIndex + 1}</td>
                    <td><div class="upload-user-cell">${avatar}<span><b>${escapeHtml(item.uploaderName || item.user || 'Unknown')}</b><small>${escapeHtml(item.user || '')}</small>${ownershipBadge}</span></div></td>
                    <td><div class="upload-file-cell"><span class="upload-file-icon"><i class="fa-solid fa-file-excel"></i></span><span><b>${escapeHtml(item.name)}</b><small>${escapeHtml(item.fileType || (item.name.split('.').pop() || 'FILE').toUpperCase())} · ${formatFileSize(item.fileSize)}${item.rowCount ? ' · ' + item.rowCount + ' rows' : ''}</small></span></div></td>
                    <td>${escapeHtml(item.date || '—')}</td>
                    <td><div class="upload-file-action-group"><button class="action-btn upload-details-btn" data-index="${index}" title="File Details & Version History"><i class="fa-solid fa-circle-info"></i> File Details</button><button class="action-btn preview-file-btn" data-index="${index}"><i class="fa-solid fa-eye"></i> Preview</button></div></td>
                    <td><button class="action-btn download-btn-action" data-index="${index}"><i class="fa-solid fa-download"></i> Download</button></td>
                    <td><button type="button" class="action-btn delete-btn-action ${canDelete ? '' : 'upload-delete-denied'}" data-index="${index}" title="${canDelete ? 'Delete file' : 'File របស់អ្នកដទៃ — ចុចដើម្បីមើលសារ​សិទ្ធិ'}"><i class="fa-solid fa-trash"></i> ${canDelete ? 'Delete' : 'គ្មានសិទ្ធិ'}</button></td>
                </tr>
            `;
        });
        tableBody.innerHTML = html;
        if (gridContainer) {
            gridContainer.innerHTML = visibleFiles.map(({item,index}) => {
                const avatar = item.uploaderAvatar ? `<img class="upload-user-avatar" src="${escapeHtml(item.uploaderAvatar)}" alt="">` : `<span class="upload-user-avatar upload-user-avatar-fallback"><i class="fa-solid fa-user"></i></span>`;
                const canDelete = canManageUploadFiles(item);
                const ownershipBadge = currentUser.role === 'Admin' ? '<span class="upload-owner-badge upload-owner-admin">Admin · គ្រប់គ្រងបាន</span>' : (currentUser.id && item.uploaderId && currentUser.id === item.uploaderId) ? '<span class="upload-owner-badge upload-owner-mine">របស់ខ្ញុំ</span>' : '<span class="upload-owner-badge upload-owner-other">របស់អ្នកដទៃ</span>';
                return `<article class="upload-file-card"><div class="upload-file-card-top"><input type="checkbox" class="upload-file-select" data-index="${index}" aria-label="ជ្រើសរើស ${escapeHtml(item.name)}" ${selectedUploadIndexes.has(index)?'checked':''}><span class="upload-file-icon"><i class="fa-solid fa-file-excel"></i></span><button type="button" class="upload-card-details" data-index="${index}" title="File Details"><i class="fa-solid fa-ellipsis"></i></button></div><h4 title="${escapeHtml(item.name)}">${escapeHtml(item.name)}</h4><p>${escapeHtml(item.fileType || 'EXCEL')} · ${formatFileSize(item.fileSize)}</p><div class="upload-file-card-user">${avatar}<span>${escapeHtml(item.uploaderName || item.user || 'Unknown')} ${ownershipBadge}</span></div><small class="upload-file-card-date"><i class="fa-regular fa-calendar"></i> ${escapeHtml(item.date || '—')}</small><div class="upload-file-card-actions"><button type="button" class="action-btn upload-details-btn" data-index="${index}" title="File Details & Version History"><i class="fa-solid fa-circle-info"></i> File Details</button><button type="button" class="action-btn preview-file-btn" data-index="${index}"><i class="fa-solid fa-eye"></i> Preview</button><button type="button" class="action-btn download-btn-action" data-index="${index}" title="Download"><i class="fa-solid fa-download"></i></button><button type="button" class="action-btn delete-btn-action ${canDelete?'':'upload-delete-denied'}" data-index="${index}" title="${canDelete?'Delete file':'File របស់អ្នកដទៃ — ចុចដើម្បីមើលសារសិទ្ធិ'}"><i class="fa-solid fa-trash"></i></button></div></article>`;
            }).join('');
        }

        document.querySelectorAll('#files-table-body .upload-file-select, #uploaded-files-grid .upload-file-select').forEach(box => box.addEventListener('change', function () {
            const idx = parseInt(this.getAttribute('data-index'), 10);
            if (this.checked) selectedUploadIndexes.add(idx); else selectedUploadIndexes.delete(idx);
            updateBulkToolbar();
        }));
        document.querySelectorAll('#files-table-body .delete-btn-action, #uploaded-files-grid .delete-btn-action').forEach(btn => btn.addEventListener('click', function () {
            window.deleteFile(parseInt(this.getAttribute('data-index'), 10));
        }));
        document.querySelectorAll('#files-table-body .download-btn-action, #uploaded-files-grid .download-btn-action').forEach(btn => btn.addEventListener('click', function () {
            if (getUploadUser().role !== 'Admin' && getUploadUser().permissions?.download === false) { showCustomAlert('គ្មានសិទ្ធិ', 'គណនីនេះមិនមានសិទ្ធិ Download ឯកសារទេ។', 'fa-solid fa-lock', '#dc2626'); return; }
            downloadFile(parseInt(this.getAttribute('data-index'), 10));
        }));
        document.querySelectorAll('#files-table-body .preview-file-btn, #uploaded-files-grid .preview-file-btn').forEach(btn => btn.addEventListener('click', function () {
            previewUploadedFile(parseInt(this.getAttribute('data-index'), 10));
        }));
        document.querySelectorAll('#files-table-body .upload-details-btn, #uploaded-files-grid .upload-details-btn, #uploaded-files-grid .upload-card-details').forEach(btn => btn.addEventListener('click', function () { showUploadFileDetails(parseInt(this.getAttribute('data-index'), 10)); }));
        updateBulkToolbar();
    }

    function updateBulkToolbar() {
        const count = document.getElementById('upload-selected-count');
        if (count) count.textContent = 'បានជ្រើស ' + selectedUploadIndexes.size + ' File';
        const all = document.getElementById('upload-select-all');
        const tableAll = document.getElementById('upload-table-select-all');
        const selectedVisible = Array.from(document.querySelectorAll('.upload-file-select'));
        const everyVisible = selectedVisible.length > 0 && selectedVisible.every(box => box.checked);
        if (all) all.checked = everyVisible;
        if (tableAll) tableAll.checked = everyVisible;
        const bulkDelete = document.getElementById('upload-bulk-delete');
        const bulkDownload = document.getElementById('upload-bulk-download');
        if (bulkDelete) bulkDelete.disabled = selectedUploadIndexes.size === 0;
        if (bulkDownload) bulkDownload.disabled = selectedUploadIndexes.size === 0;
        const storageCount = document.getElementById('upload-storage-count');
        const storageSize = document.getElementById('upload-storage-size');
        if (storageCount) storageCount.textContent = uploadedFilesList.length;
        let bytes = 0;
        try { bytes = new Blob([JSON.stringify(safeLoad('uploadedExcelFiles'))]).size; } catch (_) { bytes = uploadedFilesList.reduce((sum, item) => sum + (Number(item.fileSize) || 0), 0); }
        if (storageSize) storageSize.textContent = formatFileSize(bytes);
    }

    function setVisibleUploadSelection(checked) {
        const visibleIndexes = new Set();
        document.querySelectorAll('.upload-file-select').forEach(box => { visibleIndexes.add(parseInt(box.getAttribute('data-index'), 10)); box.checked = checked; });
        visibleIndexes.forEach(idx => { if (checked) selectedUploadIndexes.add(idx); else selectedUploadIndexes.delete(idx); });
        updateBulkToolbar();
    }
    ['upload-select-all', 'upload-table-select-all'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('change', () => setVisibleUploadSelection(el.checked));
    });
    ['uploaded-file-type-filter', 'uploaded-file-sort'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('change', renderTable);
    });
    const uploadMonthFilterButton = document.getElementById('uploaded-file-month-filter');
    const uploadMonthFilterLabel = document.getElementById('uploaded-file-month-filter-label');
    function updateUploadMonthFilterButton() {
        if (!uploadMonthFilterLabel) return;
        if (uploadDateFilter) {
            const [year, month, day] = uploadDateFilter.split('-').map(Number);
            uploadMonthFilterLabel.textContent = new Date(year, month - 1, day).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        } else if (!uploadMonthFilter) uploadMonthFilterLabel.textContent = 'ខែ / ឆ្នាំ';
        else {
            const [year, month] = uploadMonthFilter.split('-');
            const date = new Date(Number(year), Number(month) - 1, 1);
            uploadMonthFilterLabel.textContent = date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
        }
        if (uploadMonthFilterButton) uploadMonthFilterButton.classList.toggle('is-filtered', !!(uploadMonthFilter || uploadDateFilter));
    }
    function buildUploadMonthCalendar(monthValue) {
        const [yearText, monthText] = monthValue.split('-');
        const year = Number(yearText), month = Number(monthText);
        const firstDay = new Date(year, month - 1, 1).getDay();
        const daysInMonth = new Date(year, month, 0).getDate();
        const today = new Date();
        const isCurrentMonth = year === today.getFullYear() && month === today.getMonth() + 1;
        const weekdays = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        let days = '';
        for (let i = 0; i < firstDay; i++) days += '<span class="upload-calendar-empty" aria-hidden="true"></span>';
        for (let day = 1; day <= daysInMonth; day++) {
            const isToday = isCurrentMonth && day === today.getDate();
            const dateKey = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            const isSelected = dateKey === uploadDateFilter;
            const dayOfWeek = new Date(year, month - 1, day).getDay();
            const weekendClass = dayOfWeek === 6 ? ' is-saturday' : dayOfWeek === 0 ? ' is-sunday' : '';
            const weekendLabel = dayOfWeek === 6 ? ' Saturday' : dayOfWeek === 0 ? ' Sunday' : '';
            days += `<button type="button" class="upload-calendar-day${weekendClass}${isToday ? ' is-today' : ''}${isSelected ? ' is-selected' : ''}" data-day="${day}" aria-label="${dateKey}${weekendLabel}">${day}</button>`;
        }
        return `<div class="upload-calendar-top"><button type="button" class="upload-calendar-nav" id="upload-calendar-prev" aria-label="ខែមុន"><i class="fa-solid fa-chevron-left"></i></button><div class="upload-calendar-month-title">${new Date(year, month - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}</div><button type="button" class="upload-calendar-nav" id="upload-calendar-next" aria-label="ខែបន្ទាប់"><i class="fa-solid fa-chevron-right"></i></button></div><div class="upload-calendar-weekdays">${weekdays.map(day => `<span>${day}</span>`).join('')}</div><div class="upload-calendar-days">${days}</div>`;
    }
    if (uploadMonthFilterButton) uploadMonthFilterButton.addEventListener('click', async () => {
        const now = new Date();
        let selectedMonth = uploadMonthFilter || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        let displayedMonth = selectedMonth;
        if (window.Swal && typeof window.Swal.fire === 'function') {
            const result = await window.Swal.fire({
                title: '<span class="upload-month-modal-heading"><span class="upload-month-modal-icon"><i class="fa-solid fa-calendar-days"></i></span><span><strong>ស្វែងរក File តាមខែ / ឆ្នាំ</strong></span></span>',
                html: `<div id="upload-month-calendar">${buildUploadMonthCalendar(displayedMonth)}</div>`,
                showCancelButton: false,
                showCloseButton: true,
                showDenyButton: true,
                confirmButtonText: '<i class="fa-solid fa-filter"></i> បង្ហាញ File',
                denyButtonText: '<i class="fa-solid fa-layer-group"></i> បង្ហាញទាំងអស់',
                reverseButtons: true,
                buttonsStyling: false,
                focusConfirm: false,
                customClass: {
                    popup: 'upload-month-swal-popup',
                    title: 'upload-month-swal-title',
                    htmlContainer: 'upload-month-swal-content',
                    actions: 'upload-month-swal-actions',
                    confirmButton: 'upload-month-swal-confirm',
                    denyButton: 'upload-month-swal-all',
                    cancelButton: 'upload-month-swal-cancel',
                    closeButton: 'upload-month-swal-close'
                },
                didOpen: () => {
                    const calendar = document.getElementById('upload-month-calendar');
                    const draw = () => { if (calendar) calendar.innerHTML = buildUploadMonthCalendar(displayedMonth); };
                    if (!calendar) return;
                    calendar.addEventListener('click', event => {
                        const prev = event.target.closest('#upload-calendar-prev');
                        const next = event.target.closest('#upload-calendar-next');
                        const dayButton = event.target.closest('.upload-calendar-day');
                        if (prev || next) {
                            const [y, m] = displayedMonth.split('-').map(Number);
                            const date = new Date(y, m - 1 + (prev ? -1 : 1), 1);
                            displayedMonth = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
                            selectedMonth = displayedMonth;
                            // Month navigation applies the month filter immediately.
                            uploadMonthFilter = displayedMonth;
                            uploadDateFilter = '';
                            updateUploadMonthFilterButton();
                            renderTable();
                            draw();
                        } else if (dayButton) {
                            selectedMonth = displayedMonth;
                            const day = Number(dayButton.dataset.day);
                            const [year, month] = displayedMonth.split('-').map(Number);
                            uploadMonthFilter = displayedMonth;
                            uploadDateFilter = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                            updateUploadMonthFilterButton();
                            renderTable();
                            // Keep the calendar open so another date can be selected immediately.
                            // Redraw to update the selected-day highlight without closing the modal.
                            draw();
                        }
                    });
                },
                preConfirm: () => selectedMonth
            });
            if (result.isDenied) { uploadMonthFilter = ''; uploadDateFilter = ''; }
            else if (result.isConfirmed && /^\d{4}-(0[1-9]|1[0-2])$/.test(result.value || '')) { uploadMonthFilter = result.value; uploadDateFilter = ''; }
            else return;
        } else {
            const value = window.prompt('បញ្ចូលខែ/ឆ្នាំ ជាទម្រង់ YYYY-MM (ទទេ = ទាំងអស់)', uploadMonthFilter || '');
            if (value === null) return;
            if (value.trim() && !/^\d{4}-(0[1-9]|1[0-2])$/.test(value.trim())) {
                showCustomAlert('ទម្រង់មិនត្រឹមត្រូវ', 'សូមបញ្ចូលទម្រង់ YYYY-MM ឧ. 2026-11', 'fa-solid fa-calendar-days', '#dc2626');
                return;
            }
            uploadMonthFilter = value.trim();
            uploadDateFilter = '';
        }
        updateUploadMonthFilterButton();
        renderTable();
    });
    updateUploadMonthFilterButton();
    const bulkDownloadBtn = document.getElementById('upload-bulk-download');
    if (bulkDownloadBtn) bulkDownloadBtn.addEventListener('click', () => {
        const indexes = Array.from(selectedUploadIndexes).filter(i => uploadedFilesList[i]);
        if (!indexes.length) return;
        indexes.forEach(index => downloadFile(index));
    });
    const bulkDeleteBtn = document.getElementById('upload-bulk-delete');
    if (bulkDeleteBtn) bulkDeleteBtn.addEventListener('click', () => {
        const currentUser = getUploadUser();
        const selectedIndexes = Array.from(selectedUploadIndexes).filter(i => uploadedFilesList[i]);
        const forbidden = currentUser.role === 'Admin' ? [] : selectedIndexes.filter(i => !canManageUploadFiles(uploadedFilesList[i]));
        if (forbidden.length) { const names = forbidden.slice(0, 3).map(i => uploadedFilesList[i].name).join(', '); showCustomAlert('⛔ មិនអាចលុប File របស់អ្នកដទៃបានទេ', 'មាន File ' + forbidden.length + ' ដែលមិនមែនជារបស់អ្នក (' + names + (forbidden.length > 3 ? '…' : '') + ')។ សូមដក File ទាំងនោះចេញពីជម្រើសជាមុន។ គ្មាន File ណាមួយត្រូវបានលុបទេ។', 'fa-solid fa-shield-halved', '#dc2626'); return; }
        const indexes = selectedIndexes;
        if (!indexes.length) { showCustomAlert('គ្មានសិទ្ធិ', 'សូមជ្រើសរើស File ដែលត្រូវលុបជាមុន។', 'fa-solid fa-lock', '#dc2626'); return; }
        showCustomConfirm('បញ្ជាក់ការលុប', 'តើអ្នកពិតជាចង់លុប File ចំនួន ' + indexes.length + ' មែនទេ?', () => {
            indexes.sort((a,b) => b-a).forEach(i => { const item = uploadedFilesList[i]; if (item) { uploadedFilesTrash.unshift({ ...item, deletedAt: new Date().toISOString(), deletedBy: currentUser.username }); logUploadActivity('Moved to Recycle Bin', item); uploadedFilesList.splice(i, 1); } });
            saveUploadTrash();
            selectedUploadIndexes.clear();
            safeSave('uploadedExcelFiles', uploadedFilesList);
            renderTable(); updateTotalFilesCard(); updateBulkToolbar();
            showToast('លុបជោគជ័យ', 'បានលុប File ដែលបានជ្រើសរើស។', 'fa-solid fa-trash', '#dc2626');
        });
    });

    function previewUploadedFile(index) {
        const item = uploadedFilesList[index];
        if (!item) return;
        const avatar = item.uploaderAvatar
            ? `<img src="${escapeHtml(item.uploaderAvatar)}" alt="" class="upload-preview-avatar">`
            : `<span class="upload-preview-avatar upload-user-avatar-fallback"><i class="fa-solid fa-user"></i></span>`;
        const details = `<div class="upload-preview-modal"><div class="upload-preview-heading"><span class="upload-preview-file-icon"><i class="fa-solid fa-file-excel"></i></span><div><h3>${escapeHtml(item.name)}</h3><p>${escapeHtml(item.fileType || 'EXCEL')} · ${formatFileSize(item.fileSize)}</p></div></div><div class="upload-preview-uploader">${avatar}<div><b>${escapeHtml(item.uploaderName || item.user || 'Unknown')}</b><small>${escapeHtml(item.user || '')}</small></div></div><div class="upload-preview-grid"><div><small>Upload Date</small><b>${escapeHtml(item.date || '—')}</b></div><div><small>Rows</small><b>${Number(item.rowCount) || Object.values(item.prepMap || {}).reduce((n, rows) => n + (Array.isArray(rows) ? rows.length : 0), 0)}</b></div><div><small>Titles</small><b>${(item.titles || []).length}</b></div><div><small>File Size</small><b>${formatFileSize(item.fileSize)}</b></div></div><p class="upload-preview-note">Preview បង្ហាញព័ត៌មាន និងស្ថិតិរបស់ Excel ដែលបានរក្សាទុកក្នុង TimeStudyPro។</p></div>`;
        if (window.Swal && typeof window.Swal.fire === 'function') {
            window.Swal.fire({
                title: 'File Preview',
                html: details,
                width: 340,
                padding: '14px',
                showCloseButton: true,
                showConfirmButton: false,
                focusConfirm: false,
                customClass: { popup: 'upload-preview-swal-popup' }
            });
        } else {
            showCustomAlert('File Preview', item.name + '\nUploader: ' + (item.uploaderName || item.user || 'Unknown') + '\nDate: ' + (item.date || '—') + '\nSize: ' + formatFileSize(item.fileSize) + '\nRows: ' + (Number(item.rowCount) || 0), 'fa-solid fa-file-excel', '#2563eb');
        }
    }

    function showUploadFileDetails(index) {
        const item = uploadedFilesList[index]; if (!item) return;
        item.versionHistory = Array.isArray(item.versionHistory) ? item.versionHistory : [];
        const historyRows = item.versionHistory.length ? item.versionHistory.map((v,i) => `<div class="upload-history-row"><div><b>Version ${item.versionHistory.length-i}</b><small>${escapeHtml(v.savedAt || '—')} · ${escapeHtml(v.savedBy || 'Unknown')}</small><small>${formatFileSize(v.fileSize)} · ${Number(v.rowCount)||0} rows</small></div><button type="button" class="action-btn upload-restore-version" data-version="${i}"><i class="fa-solid fa-rotate-left"></i> Restore</button></div>`).join('') : '<p class="upload-muted">មិនទាន់មាន Version Snapshot ទេ។</p>';
        const html = `<div class="upload-details-panel"><div class="upload-details-head"><span class="upload-file-icon"><i class="fa-solid fa-file-excel"></i></span><div><h3>${escapeHtml(item.name)}</h3><p>${escapeHtml(item.fileType||'EXCEL')} · ${formatFileSize(item.fileSize)}</p></div></div><div class="upload-details-grid"><div><small>Uploader</small><b>${escapeHtml(item.uploaderName||item.user||'Unknown')}</b></div><div><small>Upload Date</small><b>${escapeHtml(item.date||'—')}</b></div><div><small>Rows</small><b>${Number(item.rowCount)||0}</b></div><div><small>Titles</small><b>${(item.titles||[]).length}</b></div><div><small>Permission</small><b>${canManageUploadFiles(item)?'Can manage':'Read / Download'}</b></div><div><small>Saved versions</small><b>${item.versionHistory.length}</b></div></div><div class="upload-detail-actions"><button type="button" id="upload-save-version" class="action-btn"><i class="fa-solid fa-code-branch"></i> Save Version Snapshot</button></div><h4>Version History</h4><div class="upload-version-history">${historyRows}</div></div>`;
        if (!window.Swal || typeof window.Swal.fire !== 'function') { previewUploadedFile(index); return; }
        window.Swal.fire({ title:'File Details', html, width:520, showConfirmButton:false, showCloseButton:true, customClass:{popup:'upload-preview-swal-popup upload-details-swal'} , didOpen: () => {
            const saveBtn = document.getElementById('upload-save-version');
            if (saveBtn) { saveBtn.disabled = !canManageUploadFiles(item); saveBtn.title = saveBtn.disabled ? 'មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះអាចកែប្រែ Version បាន' : ''; saveBtn.addEventListener('click', () => {
                const latest = uploadedFilesList[index]; if (!latest) return; if (!canManageUploadFiles(latest)) { showCustomAlert('គ្មានសិទ្ធិ', 'មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះអាចរក្សាទុក Version បាន។', 'fa-solid fa-lock', '#dc2626'); return; }
                latest.versionHistory = Array.isArray(latest.versionHistory) ? latest.versionHistory : [];
                latest.versionHistory.unshift({ ...fileSnapshot(latest), savedAt:new Date().toLocaleString(), savedBy:getUploadUser().username, snapshotId:'v-'+Date.now() });
                latest.versionHistory = latest.versionHistory.slice(0,10);
                if (safeSave('uploadedExcelFiles', uploadedFilesList)) { logUploadActivity('Saved version snapshot', latest); window.Swal.close(); showUploadFileDetails(index); }
            }); }
            document.querySelectorAll('.upload-restore-version').forEach(btn => { btn.disabled = !canManageUploadFiles(item); btn.title = btn.disabled ? 'មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះអាច Restore បាន' : ''; btn.addEventListener('click', () => {
                const latest = uploadedFilesList[index]; const version = latest && latest.versionHistory && latest.versionHistory[Number(btn.dataset.version)]; if (!version) return; if (!canManageUploadFiles(latest)) { showCustomAlert('គ្មានសិទ្ធិ', 'មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះអាច Restore Version បាន។', 'fa-solid fa-lock', '#dc2626'); return; }
                const currentSnapshot = { ...fileSnapshot(latest), savedAt:new Date().toLocaleString(), savedBy:getUploadUser().username, snapshotId:'v-'+Date.now() };
                const keepHistory = latest.versionHistory.slice();
                Object.keys(latest).forEach(k => { if (k !== 'versionHistory') delete latest[k]; });
                Object.assign(latest, fileSnapshot(version), { versionHistory:keepHistory });
                latest.versionHistory.unshift(currentSnapshot);
                if (safeSave('uploadedExcelFiles', uploadedFilesList)) { logUploadActivity('Restored version', latest); window.Swal.close(); renderTable(); updateTotalFilesCard(); showUploadFileDetails(index); }
            }); });
        }});
    }
    function openUploadRecycleBin() {
        const html = uploadedFilesTrash.length ? `<div class="upload-trash-list">${uploadedFilesTrash.map((item,i)=>`<div class="upload-trash-row"><div><b>${escapeHtml(item.name)}</b><small>Deleted: ${escapeHtml(item.deletedAt||'—')} · ${formatFileSize(item.fileSize)}</small></div><button type="button" class="action-btn upload-trash-restore" data-index="${i}" ${canManageUploadFiles(item) ? '' : 'disabled title="មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះ"'}><i class="fa-solid fa-rotate-left"></i> Restore</button><button type="button" class="action-btn delete-btn-action upload-trash-permanent" data-index="${i}" title="Delete permanently" ${canManageUploadFiles(item) ? '' : 'disabled'}><i class="fa-solid fa-trash-can"></i></button></div>`).join('')}</div>` : '<div class="upload-grid-empty"><i class="fa-solid fa-trash-can-arrow-up"></i><p>Recycle Bin ទទេ</p></div>';
        if (!window.Swal || typeof window.Swal.fire !== 'function') { showCustomAlert('Recycle Bin', 'File ក្នុង Recycle Bin: '+uploadedFilesTrash.length, 'fa-solid fa-trash-arrow-up', '#2563eb'); return; }
        window.Swal.fire({title:'<span class="upload-feature-modal-title"><i class="fa-solid fa-trash-arrow-up"></i><span>Recycle Bin</span></span>', html, width:540, showConfirmButton:false, showCloseButton:true, customClass:{popup:'upload-preview-swal-popup upload-details-swal',title:'upload-feature-modal-title-wrap',closeButton:'upload-feature-modal-close'}, didOpen:()=>{
            document.querySelectorAll('.upload-trash-restore').forEach(btn=>btn.addEventListener('click',()=>{ const i=Number(btn.dataset.index), item=uploadedFilesTrash[i]; if(!item)return; if(!canManageUploadFiles(item)){showCustomAlert('គ្មានសិទ្ធិ','មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះអាច Restore បាន។','fa-solid fa-lock','#dc2626');return;} if(findDuplicateFile(item.name)){showCustomAlert('ឈ្មោះស្ទួន','មាន File ឈ្មោះនេះក្នុងបញ្ជីបច្ចុប្បន្នរួចហើយ។ សូមប្ដូរឈ្មោះ ឬលុប File ស្ទួនជាមុន។','fa-solid fa-copy','#dc2626');return;} const restored={...item}; delete restored.deletedAt; delete restored.deletedBy; uploadedFilesList.unshift(restored); uploadedFilesTrash.splice(i,1); if(safeSave('uploadedExcelFiles',uploadedFilesList)&&saveUploadTrash()){logUploadActivity('Restored from Recycle Bin',restored);window.Swal.close();renderTable();updateTotalFilesCard();openUploadRecycleBin();} }));
            document.querySelectorAll('.upload-trash-permanent').forEach(btn=>btn.addEventListener('click',()=>{const i=Number(btn.dataset.index),item=uploadedFilesTrash[i];if(!item)return;if(!canManageUploadFiles(item)){showCustomAlert('គ្មានសិទ្ធិ','មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះអាចលុបជាអចិន្ត្រៃយ៍បាន។','fa-solid fa-lock','#dc2626');return;}showCustomConfirm('លុបជាអចិន្ត្រៃយ៍','តើអ្នកចង់លុប '+item.name+' ជាអចិន្ត្រៃយ៍មែនទេ? មិនអាច Restore បានទេ។',()=>{uploadedFilesTrash.splice(i,1);saveUploadTrash();logUploadActivity('Permanently deleted',item);window.Swal.close();openUploadRecycleBin();});}));
        }});
    }
    function openUploadStorageDashboard() {
        let serializedBytes=0; try { serializedBytes=new Blob([JSON.stringify(safeLoad('uploadedExcelFiles'))]).size; } catch(_) {}
        const totalBytes=uploadedFilesList.reduce((sum,f)=>sum+(Number(f.fileSize)||0),0), xls=uploadedFilesList.filter(f=>String(f.name).toLowerCase().endsWith('.xls')&&!String(f.name).toLowerCase().endsWith('.xlsx')), xlsx=uploadedFilesList.filter(f=>String(f.name).toLowerCase().endsWith('.xlsx')), other=uploadedFilesList.filter(f=>!(/\.xlsx?$/i).test(String(f.name)));
        const xlsxBytes=xlsx.reduce((n,f)=>n+(Number(f.fileSize)||0),0), xlsBytes=xls.reduce((n,f)=>n+(Number(f.fileSize)||0),0), otherBytes=other.reduce((n,f)=>n+(Number(f.fileSize)||0),0);
        const largest=uploadedFilesList.slice().sort((a,b)=>(Number(b.fileSize)||0)-(Number(a.fileSize)||0))[0];
        const pct=(n)=>totalBytes?Math.round(n/totalBytes*100):0, p1=pct(xlsxBytes), p2=pct(xlsBytes), p3=pct(otherBytes);
        const ring=totalBytes?`conic-gradient(#2563eb 0% ${p1}%, #14b8a6 ${p1}% ${p1+p2}%, #f59e0b ${p1+p2}% 100%)`:'#e2e8f0';
        const html=`<div class="upload-storage-dashboard"><div class="upload-storage-kpis"><div><small>Active Files</small><b>${uploadedFilesList.length}</b></div><div><small>Original File Size</small><b>${formatFileSize(totalBytes)}</b></div><div><small>Recycle Bin</small><b>${uploadedFilesTrash.length}</b></div><div><small>localStorage JSON</small><b>${formatFileSize(serializedBytes)}</b></div></div><h4>Storage by File Type</h4><div class="upload-storage-chart"><div class="upload-storage-donut" style="background:${ring}"><div class="upload-storage-donut-center"><b>${formatFileSize(totalBytes)}</b><small>Total Used</small></div></div><div class="upload-storage-legend"><div><i class="upload-legend-xlsx"></i><span><b>.xlsx</b><small>${xlsx.length} files · ${formatFileSize(xlsxBytes)}</small></span><strong>${p1}%</strong></div><div><i class="upload-legend-xls"></i><span><b>.xls</b><small>${xls.length} files · ${formatFileSize(xlsBytes)}</small></span><strong>${p2}%</strong></div><div><i class="upload-legend-other"></i><span><b>Other</b><small>${other.length} files · ${formatFileSize(otherBytes)}</small></span><strong>${p3}%</strong></div></div></div><h4>Largest File</h4><p class="upload-storage-largest"><i class="fa-solid fa-file-excel"></i> ${largest?escapeHtml(largest.name)+' · '+formatFileSize(largest.fileSize):'មិនទាន់មាន File'}</p><p class="upload-muted">ទំហំ localStorage គឺជាការប៉ាន់ស្មានរបស់ JSON ក្នុង Browser មិនមែនទំហំ Disk/Server ទេ។</p></div>`;
        if(window.Swal) window.Swal.fire({title:'<span class="upload-feature-modal-title"><i class="fa-solid fa-chart-pie"></i><span>Storage</span></span>',html,width:500,showConfirmButton:false,showCloseButton:true,customClass:{popup:'upload-preview-swal-popup upload-details-swal',title:'upload-feature-modal-title-wrap',closeButton:'upload-feature-modal-close'}}); else showCustomAlert('Storage Dashboard',`Files: ${uploadedFilesList.length}; Size: ${formatFileSize(totalBytes)}`,'fa-solid fa-database','#2563eb');
    }
    function openUploadActivityLog() {
        const html=uploadActivityLog.length?`<div class="upload-activity-list">${uploadActivityLog.slice(0,100).map(a=>{const avatar=a.avatar?`<img class="upload-activity-avatar" src="${escapeHtml(a.avatar)}" alt="">`:`<span class="upload-activity-avatar upload-user-avatar-fallback"><i class="fa-solid fa-user"></i></span>`;return `<div class="upload-activity-row">${avatar}<div class="upload-activity-main"><b>${escapeHtml(a.action)}</b><small>${escapeHtml(a.name)} · ${escapeHtml(a.user||'Unknown')}</small></div><time>${escapeHtml(a.at?new Date(a.at).toLocaleString():'—')}</time></div>`;}).join('')}</div>`:'<p class="upload-muted">មិនទាន់មានសកម្មភាពទេ។</p>';
        if(window.Swal) window.Swal.fire({title:'<span class="upload-feature-modal-title"><i class="fa-solid fa-clock-rotate-left"></i><span>Activity Log</span></span>',html,width:520,showConfirmButton:false,showCloseButton:true,customClass:{popup:'upload-preview-swal-popup upload-details-swal',title:'upload-feature-modal-title-wrap',closeButton:'upload-feature-modal-close'}}); else showCustomAlert('Activity Log','Activities: '+uploadActivityLog.length,'fa-solid fa-clock-rotate-left','#2563eb');
    }
    function openUploadPermissions() {
        const user=getUploadUser(); const isAdmin=user.role==='Admin';
        const html=`<div class="upload-permissions-panel"><div class="upload-permission-current"><i class="fa-solid fa-user-shield"></i><div><b>${escapeHtml(user.fullName||user.username)}</b><small>Current role: ${escapeHtml(user.role)}</small></div></div><div class="upload-permission-row"><span>View File Details / Preview</span><b>All signed-in users</b></div><div class="upload-permission-row"><span>Download File</span><b>All signed-in users</b></div><div class="upload-permission-row"><span>Delete / Move to Recycle Bin</span><b>Admin or original uploader</b></div><div class="upload-permission-row"><span>Restore / Permanently Delete Trash</span><b>${isAdmin?'Admin recommended':'Current UI allows local restore'}</b></div><p class="upload-muted">សិទ្ធិទាំងនេះត្រូវបានអនុវត្តនៅក្នុង Browser/localStorage។ សម្រាប់ការការពារពិតប្រាកដរវាងអ្នកប្រើច្រើន ត្រូវមាន Server-side authentication និង database rules។</p></div>`;
        if(window.Swal) window.Swal.fire({title:'<span class="upload-feature-modal-title"><i class="fa-solid fa-user-shield"></i><span>Permissions</span></span>',html,width:500,showConfirmButton:false,showCloseButton:true,customClass:{popup:'upload-preview-swal-popup upload-details-swal',title:'upload-feature-modal-title-wrap',closeButton:'upload-feature-modal-close'}}); else showCustomAlert('File Permissions',`Role: ${user.role}`,'fa-solid fa-user-shield','#2563eb');
    }
    const viewListButton=document.getElementById('upload-view-list'), viewGridButton=document.getElementById('upload-view-grid');
    if(viewListButton) viewListButton.addEventListener('click',()=>{uploadViewMode='list';localStorage.setItem('tspUploadViewMode',uploadViewMode);renderTable();});
    if(viewGridButton) viewGridButton.addEventListener('click',()=>{uploadViewMode='grid';localStorage.setItem('tspUploadViewMode',uploadViewMode);renderTable();});
    const trashButton=document.getElementById('upload-open-trash'); if(trashButton) trashButton.addEventListener('click',openUploadRecycleBin);
    const storageButton=document.getElementById('upload-storage-dashboard'); if(storageButton) storageButton.addEventListener('click',openUploadStorageDashboard);
    const activityButton=document.getElementById('upload-activity-log'); if(activityButton) activityButton.addEventListener('click',openUploadActivityLog);
    const permissionsButton=document.getElementById('upload-permissions'); if(permissionsButton) permissionsButton.addEventListener('click',openUploadPermissions);

    function updateTotalFilesCard() {
        const totalFileCountCard = document.getElementById('total-file-count');
        if (totalFileCountCard) totalFileCountCard.textContent = uploadedFilesList.length;
        const dashCount = document.getElementById('dash-total-file-count');
        if (dashCount) dashCount.textContent = uploadedFilesList.length;
    }

    const uploadedFileSearch = document.getElementById('uploaded-file-search');
    if (uploadedFileSearch) uploadedFileSearch.addEventListener('input', renderTable);
    renderTable();
    updateTotalFilesCard();

    window.deleteFile = function (index) {
        const target = uploadedFilesList[index];
        if (!target) return;
        const fileName = target.name;
        const currentUser = getUploadUser();
        const canDelete = canManageUploadFiles(target);
        if (!canDelete) {
            showCustomAlert('⛔ មិនអាចលុប File នេះបានទេ', 'File នេះមិនមែនជារបស់អ្នកទេ។ ម្ចាស់ File៖ ' + (target.uploaderName || target.user || 'User ផ្សេង') + '. អ្នកអាចមើល និង Download បាន ប៉ុន្តែមិនអាចលុបបានទេ។ Admin មានសិទ្ធិគ្រប់យ៉ាង។', 'fa-solid fa-shield-halved', '#dc2626');
            return;
        }
        showCustomConfirm('បញ្ជាក់ការលុប', 'តើអ្នកពិតជាចង់លុបឯកសារ [' + fileName + '] នេះមែនទេ?', () => {
            const removedItem = uploadedFilesList[index];
            if (removedItem) { uploadedFilesTrash.unshift({ ...removedItem, deletedAt: new Date().toISOString(), deletedBy: currentUser.username }); saveUploadTrash(); logUploadActivity('Moved to Recycle Bin', removedItem); }
            uploadedFilesList.splice(index, 1);
            selectedUploadIndexes.clear();
            safeSave('uploadedExcelFiles', uploadedFilesList);
            renderTable();
            updateTotalFilesCard();
            showToast('លុបជោគជ័យ', 'ឯកសារ [' + fileName + '] ត្រូវបានលុបចេញពីប្រព័ន្ធ', 'fa-solid fa-trash', '#dc2626');
        });
    };

    function downloadFile(index) {
        const fileObj = uploadedFilesList[index];
        if (!fileObj) return;
        if (typeof XLSX === 'undefined') {
            showCustomAlert('កំហុសឆ្គង', 'មិនអាចទាញយកបានទេ (XLSX library មិនទាន់ Load)', 'fa-solid fa-triangle-exclamation', '#dc2626');
            return;
        }
        try {
            const rows = [];
            Object.values(fileObj.prepMap || {}).forEach(arr => arr.forEach(i => rows.push(i.rawRowData || {})));
            if (rows.length === 0) {
                showCustomAlert('ជូនដំណឹង', 'គ្មានទិន្នន័យសម្រាប់ទាញយកទេ!', 'fa-solid fa-circle-info', '#3b82f6');
                return;
            }
            const ws = XLSX.utils.json_to_sheet(rows, fileObj.excelHeaders && fileObj.excelHeaders.length ? { header: fileObj.excelHeaders } : undefined);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
            showToast('កំពុងទាញយក...', fileObj.name, 'fa-solid fa-circle-down', '#2563eb');
            logUploadActivity('Downloaded', fileObj);
            XLSX.writeFile(wb, fileObj.name);
        } catch (err) {
            console.error('Download error:', err);
            showCustomAlert('កំហុសឆ្គង', 'មិនអាចទាញយកឯកសារបានទេ: ' + err.message, 'fa-solid fa-triangle-exclamation', '#dc2626');
        }
    }

    // --- ៣. Master Sect, Section, Style, Rampup, Prep, Process ---
    function bindItemClicks(container, handler) {
        container.querySelectorAll('.observe-item').forEach(el => {
            el.addEventListener('click', () => handler(el.dataset));
        });
    }

    function observeTrashButton(level, attrs) {
        const data = Object.entries(attrs || {}).map(([k, v]) => `data-${k}="${escapeHtml(v)}"`).join(' ');
        // Match the Process delete button: large red trash button on the LEFT of each row.
        return `<button type="button" class="observe-trash-btn" data-level="${level}" ${data} title="លុប ${level}" style="background:#fee2e2;color:#dc2626;border:none;padding:10px 14px;border-radius:8px;font-size:16px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;transition:all .2s;flex-shrink:0;"><i class="fa-solid fa-trash-can" style="font-size:16px;"></i></button>`;
    }

    function observeGroupExistsInFile(file, level, d) {
        if (!file) return false;
        const n = normStr;
        const title = d.title || '', section = d.section || '', styleName = d.style || '', rampupDay = d.rampup || '', prepName = d.prep || '';
        const k2 = `${title}_${section}`, k3 = `${title}_${section}_${styleName}`, k4 = `${title}_${section}_${styleName}_${rampupDay}`;
        const hasKey = (obj, prefix) => !!obj && Object.keys(obj).some(k => k === prefix || k.startsWith(prefix + '_'));
        if (level === 'title') return (file.titles || []).some(x => n(x) === n(title)) || hasKey(file.sectionsMap, title) || hasKey(file.stylesMap, title) || hasKey(file.rampupMap, title) || hasKey(file.prepMap, title);
        if (level === 'section') return !!((file.sectionsMap && Array.isArray(file.sectionsMap[title]) && file.sectionsMap[title].some(x => n(x) === n(section))) || hasKey(file.stylesMap, k2) || hasKey(file.rampupMap, k2) || hasKey(file.prepMap, k2));
        if (level === 'style') return !!((file.stylesMap && Array.isArray(file.stylesMap[k2]) && file.stylesMap[k2].some(x => n(x) === n(styleName))) || hasKey(file.rampupMap, k3) || hasKey(file.prepMap, k3));
        if (level === 'rampup') return !!((file.rampupMap && Array.isArray(file.rampupMap[k3]) && file.rampupMap[k3].some(x => n(x) === n(rampupDay))) || hasKey(file.prepMap, k4));
        if (level === 'prep') return !!(file.prepMap && Array.isArray(file.prepMap[k4]) && file.prepMap[k4].some(x => n(x.prepAssembly) === n(prepName)));
        return false;
    }

    async function deleteObserveGroup(level, d) {
        const title = d.title || '';
        const section = d.section || '';
        const styleName = d.style || '';
        const rampupDay = d.rampup || '';
        const prepName = d.prep || '';
        const label = ({title, section, style: styleName, rampup: rampupDay, prep: prepName})[level] || '';
        const levelLabel = ({title:'Master Sect / Title', section:'Section', style:'Style', rampup:'Rampup Day', prep:'Prep / Assembly'})[level] || level;
        const files = safeLoad('uploadedExcelFiles');
        const matchingFiles = files.filter(f => observeGroupExistsInFile(f, level, d));
        const ownedMatchingFiles = matchingFiles.filter(f => tspCanManageRecord(f));
        if (!ownedMatchingFiles.length) {
            tspPermissionDenied(matchingFiles[0]);
            return;
        }
        const ok = await confirmDialog({
            title: 'តើអ្នកពិតជាចង់លុបប្រអប់នេះទេ?',
            html: `<b>[${escapeHtml(label)}]</b><br>រួមទាំងទិន្នន័យ ${escapeHtml(levelLabel)} និង Time Study ដែលពាក់ព័ន្ធនឹងត្រូវលុបចោលទាំងស្រុង។`,
            icon: 'error',
            confirmColor: '#dc2626',
            confirmText: 'លុបចោល'
        });
        if (!ok) return;

        const studies = safeLoad('savedTimeStudies');
        const n = normStr;
        const k2 = `${title}_${section}`;
        const k3 = `${title}_${section}_${styleName}`;
        const k4 = `${title}_${section}_${styleName}_${rampupDay}`;
        files.forEach(f => {
            if (!tspCanManageRecord(f)) return;
            if (level === 'title') {
                f.titles = (f.titles || []).filter(x => n(x) !== n(title));
                if (f.sectionsMap) delete f.sectionsMap[title];
                ['stylesMap','rampupMap','prepMap'].forEach(mapName => {
                    if (f[mapName]) Object.keys(f[mapName]).forEach(k => { if (k === title || k.startsWith(title + '_')) delete f[mapName][k]; });
                });
            } else if (level === 'section') {
                if (f.sectionsMap && Array.isArray(f.sectionsMap[title])) f.sectionsMap[title] = f.sectionsMap[title].filter(x => n(x) !== n(section));
                if (f.stylesMap) delete f.stylesMap[k2];
                ['rampupMap','prepMap'].forEach(mapName => { if (f[mapName]) Object.keys(f[mapName]).forEach(k => { if (k.startsWith(k2 + '_')) delete f[mapName][k]; }); });
            } else if (level === 'style') {
                if (f.stylesMap && Array.isArray(f.stylesMap[k2])) f.stylesMap[k2] = f.stylesMap[k2].filter(x => n(x) !== n(styleName));
                if (f.rampupMap) delete f.rampupMap[k3];
                if (f.prepMap) Object.keys(f.prepMap).forEach(k => { if (k.startsWith(k3 + '_')) delete f.prepMap[k]; });
            } else if (level === 'rampup') {
                if (f.rampupMap && Array.isArray(f.rampupMap[k3])) f.rampupMap[k3] = f.rampupMap[k3].filter(x => n(x) !== n(rampupDay));
                if (f.prepMap) delete f.prepMap[k4];
            } else if (level === 'prep') {
                if (f.prepMap && Array.isArray(f.prepMap[k4])) f.prepMap[k4] = f.prepMap[k4].filter(x => n(x.prepAssembly) !== n(prepName));
                if (f.prepMap && f.prepMap[k4] && !f.prepMap[k4].length) delete f.prepMap[k4];
            }
        });
        // Wait for IndexedDB persistence before re-rendering Time Observe.
        // Otherwise safeLoad() can read the old cache and the deleted row remains visible until reload.
        const filesSaved = await tspPersistLargeStorage('uploadedExcelFiles', files);
        if (!filesSaved) {
            showCustomAlert('កំហុសរក្សាទុក', 'មិនអាចរក្សាទុកការលុបទិន្នន័យបានទេ។ សូមព្យាយាមម្ដងទៀត។', 'fa-solid fa-triangle-exclamation', '#dc2626');
            return;
        }
        const belongs = s => {
            if (title && n(s.title || s.factory) !== n(title)) return false;
            if (['section','style','rampup','prep'].includes(level) && n(s.section) !== n(section)) return false;
            if (['style','rampup','prep'].includes(level) && n(s.styleName) !== n(styleName)) return false;
            if (['rampup','prep'].includes(level) && n(s.rampupDay) !== n(rampupDay)) return false;
            if (level === 'prep' && n(s.prepAssembly || s.prepName) !== n(prepName)) return false;
            return true;
        };
        safeSave('savedTimeStudies', studies.filter(s => !tspCanManageRecord(s) || !belongs(s)));
        if (level === 'title') window.renderMasterSectList(document.getElementById('section-search-input')?.value || '');
        else if (level === 'section') window.renderSectionList(title, document.getElementById('section-item-search-input')?.value || '');
        else if (level === 'style') window.renderStyleList(title, section, document.getElementById('style-item-search-input')?.value || '');
        else if (level === 'rampup') window.renderRampupList(title, section, styleName, document.getElementById('rampup-item-search-input')?.value || '');
        else if (level === 'prep') window.renderPrepAssemblyList(title, section, styleName, rampupDay, document.getElementById('prep-item-search-input')?.value || '');
        successDialog('បានលុបជោគជ័យ!', `ប្រអប់ ${escapeHtml(levelLabel)} [${escapeHtml(label)}] ត្រូវបានលុបចេញរួចរាល់។`);
    }

    function bindObserveTrash(container) {
        container.querySelectorAll('.observe-trash-btn').forEach(btn => btn.addEventListener('click', e => {
            e.preventDefault(); e.stopPropagation();
            const d = {...btn.dataset};
            deleteObserveGroup(d.level, d);
        }));
    }

    window.renderMasterSectList = function (filterText = '') {
        const container = document.getElementById('section-list-container');
        if (!container) return;
        const uploadedFiles = safeLoad('uploadedExcelFiles');

        if (uploadedFiles.length === 0) {
            container.innerHTML = `<div class="empty-observe-state"><p><i class="fa-solid fa-triangle-exclamation" style="color: #f59e0b;"></i> សូម Upload ហ្វាល Excel ជាមុនសិន!</p></div>`;
            return;
        }

        const allTitlesSet = new Set();
        uploadedFiles.forEach(f => f.titles && f.titles.forEach(t => allTitlesSet.add(t)));
        let titles = Array.from(allTitlesSet);
        if (filterText) titles = titles.filter(t => t.toLowerCase().includes(filterText.toLowerCase()));

        if (titles.length === 0) {
            container.innerHTML = `<div class="empty-observe-state"><p>រកមិនឃើញ Master Sect ដែលត្រូវគ្នាទេ។</p></div>`;
            return;
        }

        let html = '';
        titles.forEach(title => {
            html += `
                <div class="observe-item" data-title="${escapeHtml(title)}" style="cursor:pointer;">
                    <div class="observe-item-left" style="display:flex;align-items:center;gap:12px;">${observeTrashButton('title', {title})}<i class="fa-solid fa-folder-tree" style="color: #2563eb; margin-right: 8px;"></i><span class="observe-title-text">${escapeHtml(title)}</span></div>
                    <div class="observe-item-right"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html;
        bindObserveTrash(container);
        bindItemClicks(container, d => window.openTimeObserveDetail(d.title));
    };

    const searchInput = document.getElementById('section-search-input');
    if (searchInput) searchInput.addEventListener('input', (e) => window.renderMasterSectList(e.target.value));

    window.openTimeObserveDetail = function (title) {
        document.getElementById('time-page').style.display = 'none';
        document.getElementById('section-page').style.display = 'block';
        document.getElementById('section-page-title').innerHTML = `<i class="fa-solid fa-layer-group"></i> SECTIONS (${escapeHtml(title)})`;
        window.currentSelectedTitle = title;
        window.renderSectionList(title);
    };

    const backBtn = document.getElementById('back-to-observe-btn');
    if (backBtn) backBtn.addEventListener('click', () => {
        document.getElementById('section-page').style.display = 'none';
        document.getElementById('time-page').style.display = 'block';
    });

    window.renderSectionList = function (title, filterText = '') {
        const container = document.getElementById('section-item-list-container');
        if (!container) return;
        const uploadedFiles = safeLoad('uploadedExcelFiles');
        const allSectionsSet = new Set();
        uploadedFiles.forEach(f => f.sectionsMap && f.sectionsMap[title] && f.sectionsMap[title].forEach(s => allSectionsSet.add(s)));

        let sections = Array.from(allSectionsSet);
        sections.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        if (filterText) sections = sections.filter(s => s.toLowerCase().includes(filterText.toLowerCase()));

        let html = '';
        sections.forEach(section => {
            html += `
                <div class="observe-item" data-title="${escapeHtml(title)}" data-section="${escapeHtml(section)}" style="cursor:pointer;">
                    <div class="observe-item-left" style="display:flex;align-items:center;gap:12px;">${observeTrashButton('section', {title, section})}<i class="fa-solid fa-folder-open" style="color: #2563eb;"></i><span class="observe-title-text">${escapeHtml(section)}</span></div>
                    <div class="observe-item-right"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញ Section ក្រោម Title នេះទេ។</p></div>`;
        bindObserveTrash(container);
        bindItemClicks(container, d => window.openStyleObservePage(d.title, d.section));
    };

    const sectionItemSearchInput = document.getElementById('section-item-search-input');
    if (sectionItemSearchInput) sectionItemSearchInput.addEventListener('input', (e) => {
        if (window.currentSelectedTitle) window.renderSectionList(window.currentSelectedTitle, e.target.value);
    });

    window.openStyleObservePage = function (title, section) {
        document.getElementById('section-page').style.display = 'none';
        document.getElementById('style-page').style.display = 'block';
        document.getElementById('style-page-title').innerHTML = `<i class="fa-solid fa-shirt"></i> STYLES (${escapeHtml(title)} / ${escapeHtml(section)})`;
        window.currentSelectedTitleForStyle = title;
        window.currentSelectedSection = section;
        window.renderStyleList(title, section);
    };

    const backToSectionBtn = document.getElementById('back-to-section-btn');
    if (backToSectionBtn) backToSectionBtn.addEventListener('click', () => {
        document.getElementById('style-page').style.display = 'none';
        document.getElementById('section-page').style.display = 'block';
    });

    window.renderStyleList = function (title, section, filterText = '') {
        const container = document.getElementById('style-item-list-container');
        if (!container) return;
        const uploadedFiles = safeLoad('uploadedExcelFiles');
        const compositeKey = `${title}_${section}`;
        const allStylesSet = new Set();
        uploadedFiles.forEach(f => f.stylesMap && f.stylesMap[compositeKey] && f.stylesMap[compositeKey].forEach(st => allStylesSet.add(st)));

        let styles = Array.from(allStylesSet);
        styles.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        if (filterText) styles = styles.filter(s => s.toLowerCase().includes(filterText.toLowerCase()));

        let html = '';
        styles.forEach(styleName => {
            html += `
                <div class="observe-item" data-title="${escapeHtml(title)}" data-section="${escapeHtml(section)}" data-style="${escapeHtml(styleName)}" style="cursor:pointer;">
                    <div class="observe-item-left" style="display:flex;align-items:center;gap:12px;">${observeTrashButton('style', {title, section, style: styleName})}<i class="fa-solid fa-tag" style="color: #2563eb;"></i><span class="observe-title-text">${escapeHtml(styleName)}</span></div>
                    <div class="observe-item-right"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញ Style ក្រោម Section នេះទេ។</p></div>`;
        bindObserveTrash(container);
        bindItemClicks(container, d => window.openRampupPage(d.title, d.section, d.style));
    };

    const styleItemSearchInput = document.getElementById('style-item-search-input');
    if (styleItemSearchInput) styleItemSearchInput.addEventListener('input', (e) => {
        if (window.currentSelectedTitleForStyle && window.currentSelectedSection) {
            window.renderStyleList(window.currentSelectedTitleForStyle, window.currentSelectedSection, e.target.value);
        }
    });

    window.openRampupPage = function (title, section, styleName) {
        document.getElementById('style-page').style.display = 'none';
        document.getElementById('rampup-page').style.display = 'block';
        document.getElementById('rampup-page-title').innerHTML = `<i class="fa-solid fa-calendar-day"></i> RAMPUP DAYS (${escapeHtml(title)} / ${escapeHtml(section)} / ${escapeHtml(styleName)})`;
        window.currentSelectedTitleForRampup = title;
        window.currentSelectedSectionForRampup = section;
        window.currentSelectedStyle = styleName;
        window.renderRampupList(title, section, styleName);
    };

    const backToStyleBtn = document.getElementById('back-to-style-btn');
    if (backToStyleBtn) backToStyleBtn.addEventListener('click', () => {
        document.getElementById('rampup-page').style.display = 'none';
        document.getElementById('style-page').style.display = 'block';
    });

    window.renderRampupList = function (title, section, styleName, filterText = '') {
        const container = document.getElementById('rampup-item-list-container');
        if (!container) return;
        const uploadedFiles = safeLoad('uploadedExcelFiles');
        const compositeKey = `${title}_${section}_${styleName}`;
        const allRampupSet = new Set();
        uploadedFiles.forEach(f => f.rampupMap && f.rampupMap[compositeKey] && f.rampupMap[compositeKey].forEach(rd => allRampupSet.add(rd)));

        let rampups = Array.from(allRampupSet).map(String);
        rampups.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        if (filterText) rampups = rampups.filter(r => r.toLowerCase().includes(filterText.toLowerCase()));

        let html = '';
        rampups.forEach(rampupDay => {
            html += `
                <div class="observe-item" data-title="${escapeHtml(title)}" data-section="${escapeHtml(section)}" data-style="${escapeHtml(styleName)}" data-rampup="${escapeHtml(rampupDay)}" style="cursor:pointer;">
                    <div class="observe-item-left" style="display:flex;align-items:center;gap:12px;">${observeTrashButton('rampup', {title, section, style: styleName, rampup: rampupDay})}<i class="fa-solid fa-calendar-check" style="color: #2563eb;"></i><span class="observe-title-text">${escapeHtml(rampupDay)}</span></div>
                    <div class="observe-item-right"><i class="fa-solid fa-chevron-right"></i></div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញ RampupDay ក្រោម Style នេះទេ។</p></div>`;
        bindObserveTrash(container);
        bindItemClicks(container, d => window.openPrepAssemblyPage(d.title, d.section, d.style, d.rampup));
    };

    const rampupItemSearchInput = document.getElementById('rampup-item-search-input');
    if (rampupItemSearchInput) rampupItemSearchInput.addEventListener('input', (e) => {
        if (window.currentSelectedTitleForRampup && window.currentSelectedSectionForRampup && window.currentSelectedStyle) {
            window.renderRampupList(window.currentSelectedTitleForRampup, window.currentSelectedSectionForRampup, window.currentSelectedStyle, e.target.value);
        }
    });

    window.openPrepAssemblyPage = function (title, section, styleName, rampupDay) {
        document.getElementById('rampup-page').style.display = 'none';
        document.getElementById('prep-page').style.display = 'block';
        document.getElementById('prep-page-title').innerHTML = `<i class="fa-solid fa-list-check"></i> ${escapeHtml(rampupDay)}`;
        window.prepContext = { title, section, styleName, rampupDay };
        window.renderPrepAssemblyList(title, section, styleName, rampupDay);
    };

    const backToRampupBtn = document.getElementById('back-to-rampup-btn');
    if (backToRampupBtn) backToRampupBtn.addEventListener('click', () => {
        document.getElementById('prep-page').style.display = 'none';
        document.getElementById('rampup-page').style.display = 'block';
    });

    window.renderPrepAssemblyList = function (title, section, styleName, rampupDay, filterText = '') {
        const container = document.getElementById('prep-item-list-container');
        if (!container) return;
        const uploadedFiles = safeLoad('uploadedExcelFiles');
        const savedStudies = safeLoad('savedTimeStudies');
        const compositeKey = `${title}_${section}_${styleName}_${rampupDay}`;
        let rawItems = [];
        uploadedFiles.forEach(f => f.prepMap && f.prepMap[compositeKey] && (rawItems = rawItems.concat(f.prepMap[compositeKey])));

        const ctx = { title, section, styleName, rampupDay };
        const groupedMap = {};
        rawItems.forEach(item => {
            const prepName = item.prepAssembly;
            if (!groupedMap[prepName]) {
                groupedMap[prepName] = {
                    prepAssembly: prepName,
                    primaryKeys: new Set(),
                    smvsList: [],
                    totalProcesses: 0,
                    completedCount: 0
                };
            }
            const g = groupedMap[prepName];
            if (item.primaryKey) g.primaryKeys.add(item.primaryKey);
            g.totalProcesses++;

            if (savedStudies.some(s => studyMatchesItem(s, item, ctx))) g.completedCount++;

            const numSmv = parseFloat(item.smv);
            if (!isNaN(numSmv)) g.smvsList.push(numSmv);
        });

        let prepItems = Object.values(groupedMap).map(g => {
            const totalSmv = g.smvsList.reduce((acc, curr) => acc + curr, 0);
            return {
                prepAssembly: g.prepAssembly,
                primaryKey: Array.from(g.primaryKeys).join(', '),
                smv: totalSmv > 0 ? totalSmv.toFixed(2) : '0.00',
                totalProcesses: g.totalProcesses,
                completedCount: g.completedCount,
                isCompleted: g.totalProcesses > 0 && g.completedCount >= g.totalProcesses
            };
        });

        if (filterText) {
            const kw = filterText.trim().toLowerCase();
            prepItems = prepItems.filter(it =>
                (it.prepAssembly || '').toLowerCase().includes(kw) ||
                (it.primaryKey || '').toLowerCase().includes(kw)
            );
        }

        let html = '';
        prepItems.forEach((item, index) => {
            let statusBadgeHtml = '';
            if (item.isCompleted) {
                statusBadgeHtml = `
                    <span class="prep-process-status prep-process-status-complete" style="background: #d1fae5; color: #047857; padding: 4px 12px; border-radius: 20px; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-circle-check" aria-hidden="true"></i><span class="prep-status-message"> សូមអបអរសាទរ👏🎉 អ្នកបានចាប់នាទីរួចរាល់</span><span class="prep-status-count">(${item.completedCount}/${item.totalProcesses})</span>
                    </span>`;
            } else {
                statusBadgeHtml = `
                    <span class="prep-process-status prep-process-status-incomplete" style="background: #fee2e2; color: #dc2626; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: bold; display: inline-flex; align-items: center; gap: 6px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-triangle-exclamation" aria-hidden="true"></i><span class="prep-status-message"> បញ្ជាក់៖ អ្នកមិនទាន់ចាប់នាទីគ្រប់ចំនួនProcessនោះទេ😁😁!!</span><span class="prep-status-count">(${item.completedCount}/${item.totalProcesses})</span>
                    </span>`;
            }

            html += `
                <div class="observe-item">
                    <div style="display: flex; align-items: center; gap: 12px; width: 100%;">
                        <button type="button" class="delete-prep-btn" data-index="${index}" title="លុបប្រអប់នេះចោលទាំងស្រុង" style="background: #fee2e2; color: #dc2626; border: none; padding: 10px 14px; border-radius: 8px; font-size: 16px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; transition: all 0.2s; flex-shrink: 0;">
                            <i class="fa-solid fa-trash-can" style="font-size: 16px;"></i>
                        </button>
                        <div class="observe-item-left prep-open-btn" data-prep="${escapeHtml(item.prepAssembly)}" style="flex-direction: column; align-items: flex-start; gap: 8px; width: 100%; cursor: pointer;">
                            <div class="prep-title-row" style="display: flex; align-items: center; width: 100%; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                                <span class="observe-title-text prep-title-text" style="font-size: 16px; font-weight: bold;">${escapeHtml(item.prepAssembly)}</span>
                                <div class="prep-status-wrap">${statusBadgeHtml}</div>
                            </div>
                            <div style="display: flex; gap: 15px; font-size: 12px; color: #94a3b8; align-items: center; width: 100%; justify-content: space-between; flex-wrap: wrap;">
                                <span>${escapeHtml(item.primaryKey)}</span>
                                <span style="font-weight: bold; color: #0284c7; background: #e0f2fe; padding: 3px 10px; border-radius: 6px; font-size: 13px;">
                                    Total SMV: ${item.smv}
                                </span>
                            </div>
                        </div>
                    </div>
                    <div class="observe-item-right prep-open-btn" data-prep="${escapeHtml(item.prepAssembly)}" style="cursor:pointer;">
                        <i class="fa-solid fa-chevron-right"></i>
                    </div>
                </div>
            `;
        });
        container.innerHTML = html || `<div class="empty-observe-state"><p>រកមិនឃើញទិន្នន័យ Prep / Assembly ក្រោម RampupDay នេះទេ។</p></div>`;

        container.querySelectorAll('.prep-open-btn').forEach(el => {
            el.addEventListener('click', () => window.openProcessPage(title, section, styleName, rampupDay, el.dataset.prep));
        });

        container.querySelectorAll('.delete-prep-btn').forEach(trashBtn => {
            trashBtn.addEventListener('click', async function (e) {
                e.stopPropagation();
                const targetPrep = prepItems[parseInt(this.getAttribute('data-index'))];
                if (!targetPrep) return;
                const targetName = (targetPrep.prepAssembly || '').trim();
                const filesData = safeLoad('uploadedExcelFiles');
                const matchingPrepFiles = filesData.filter(file => file.prepMap && Array.isArray(file.prepMap[compositeKey]) && file.prepMap[compositeKey].some(i => normStr(i.prepAssembly) === normStr(targetName)));
                if (!matchingPrepFiles.some(file => tspCanManageRecord(file))) { tspPermissionDenied(matchingPrepFiles[0]); return; }

                const ok = await confirmDialog({
                    title: 'តើអ្នកពិតជាចង់លុបប្រអប់នេះទេ?',
                    html: '[' + escapeHtml(targetName) + '] រួមទាំង Process និង Time Study ទាំងអស់នឹងត្រូវលុបចោលទាំងស្រុង។',
                    icon: 'error',
                    confirmColor: '#dc2626',
                    confirmText: 'លុបចោល'
                });
                if (!ok) return;

                if (!filesData.some(file => tspCanManageRecord(file) && file.prepMap && file.prepMap[compositeKey] && file.prepMap[compositeKey].some(i => normStr(i.prepAssembly) === normStr(targetName)))) { tspPermissionDenied(); return; }
                filesData.forEach(file => {
                    if (!tspCanManageRecord(file)) return;
                    if (file.prepMap && file.prepMap[compositeKey]) {
                        file.prepMap[compositeKey] = file.prepMap[compositeKey].filter(
                            i => (i.prepAssembly || '').trim() !== targetName
                        );
                    }
                });

                let rampupRemoved = false;
                const stillHasItems = filesData.some(file =>
                    tspCanManageRecord(file) && file.prepMap && file.prepMap[compositeKey] && file.prepMap[compositeKey].length > 0
                );
                if (!stillHasItems) {
                    const rampupKey = `${title}_${section}_${styleName}`;
                    filesData.forEach(file => {
                        if (!tspCanManageRecord(file)) return;
                        if (file.prepMap) delete file.prepMap[compositeKey];
                        if (file.rampupMap && Array.isArray(file.rampupMap[rampupKey])) {
                            file.rampupMap[rampupKey] = file.rampupMap[rampupKey].filter(r => String(r) !== String(rampupDay));
                        }
                    });
                    rampupRemoved = true;
                }
                // Finish IndexedDB save before refreshing the visible Time Observe list.
                const filesSaved = await tspPersistLargeStorage('uploadedExcelFiles', filesData);
                if (!filesSaved) {
                    showCustomAlert('កំហុសរក្សាទុក', 'មិនអាចរក្សាទុកការលុប Prep / Assembly បានទេ។ សូមព្យាយាមម្ដងទៀត។', 'fa-solid fa-triangle-exclamation', '#dc2626');
                    return;
                }
                uploadedFilesList = filesData;

                let studiesData = safeLoad('savedTimeStudies');
                studiesData = studiesData.filter(s => !tspCanManageRecord(s) || !(
                    normStr(s.styleName) === normStr(styleName) &&
                    normStr(s.rampupDay) === normStr(rampupDay) &&
                    normStr(s.prepAssembly || s.prepName) === normStr(targetName) &&
                    (!s.title || s.title === 'N/A' || normStr(s.title) === normStr(title)) &&
                    (!s.section || s.section === 'N/A' || normStr(s.section) === normStr(section))
                ));
                safeSave('savedTimeStudies', studiesData);

                if (rampupRemoved) {
                    document.getElementById('prep-page').style.display = 'none';
                    document.getElementById('rampup-page').style.display = 'block';
                    window.renderRampupList(title, section, styleName);
                } else {
                    window.renderPrepAssemblyList(title, section, styleName, rampupDay, filterText);
                }
                successDialog('បានលុបជោគជ័យ!', 'ប្រអប់ Prep / Assembly ត្រូវបានលុបចេញរួចរាល់។');
            });
        });
    };

    const prepItemSearchInput = document.getElementById('prep-item-search-input');
    if (prepItemSearchInput) prepItemSearchInput.addEventListener('input', (e) => {
        if (window.prepContext) {
            const c = window.prepContext;
            window.renderPrepAssemblyList(c.title, c.section, c.styleName, c.rampupDay, e.target.value);
        }
    });

    window.openProcessPage = function (title, section, styleName, rampupDay, prepName) {
        document.getElementById('prep-page').style.display = 'none';
        document.getElementById('process-page').style.display = 'block';
        document.getElementById('process-page-title').innerHTML = `<i class="fa-solid fa-list-check"></i> ${escapeHtml(prepName)}`;
        window.processContext = { title, section, styleName, rampupDay, prepName };
        window.renderProcessList(title, section, styleName, rampupDay, prepName);
    };

    const backToPrepBtn = document.getElementById('back-to-prep-btn');
    if (backToPrepBtn) backToPrepBtn.addEventListener('click', () => {
        document.getElementById('process-page').style.display = 'none';
        document.getElementById('prep-page').style.display = 'block';
        if (window.prepContext) {
            const pc = window.prepContext;
            window.renderPrepAssemblyList(pc.title, pc.section, pc.styleName, pc.rampupDay);
        }
    });

    window.renderProcessList = function (title, section, styleName, rampupDay, prepName, filterText = '') {
        const container = document.getElementById('process-item-list-container');
        if (!container) return;
        const uploadedFiles = safeLoad('uploadedExcelFiles');
        const savedStudies = safeLoad('savedTimeStudies');
        const compositeKey = `${title}_${section}_${styleName}_${rampupDay}`;
        let rawItems = [];
        uploadedFiles.forEach(f => f.prepMap && f.prepMap[compositeKey] && (rawItems = rawItems.concat(f.prepMap[compositeKey])));

        const ctx = { title, section, styleName, rampupDay, prepName };
        let processItems = rawItems.filter(item => item.prepAssembly === prepName);
        if (filterText) {
            const kw = filterText.toLowerCase();
            processItems = processItems.filter(item =>
                (item.processRegion && item.processRegion.toLowerCase().includes(kw)) ||
                (item.primaryKey && item.primaryKey.toLowerCase().includes(kw)));
        }

        if (processItems.length === 0) {
            container.innerHTML = `<div class="empty-observe-state"><p>រកមិនឃើញ Process ក្រោម Prep នេះទេ។</p></div>`;
            return;
        }

        const drafts = loadObject('tsDrafts');

        let html = '';
        processItems.forEach((item, index) => {
            const khmerText = item.processRegion || 'អត់មានឈ្មោះខ្មែរ';
            const thaiText = item.processName || '';
            const sewerNameText = validSewerName(item.sewerName);
            const primaryKeyText = item.primaryKey || 'N/A';
            const processNo = item.no || (index + 1);

            const numSmv = parseFloat(item.smv);
            const smvText = !isNaN(numSmv) ? numSmv.toFixed(2) : (item.smv || '0.00');

            const matchedStudy = savedStudies.find(s => studyMatchesItem(s, item, ctx));
            const hasDraft = !matchedStudy && item.id && drafts[item.id] && ((drafts[item.id].laps || []).length > 0);

            let checkmarkBadgeHtml = '';
            let avgCtBadgeHtml = '';
            let pfmBadgeHtml = '';
            let deleteBtnHtml = '';

            if (matchedStudy) {
                checkmarkBadgeHtml = `
                    <span class="observe-complete-badge" title="រួចរាល់" aria-label="រួចរាល់" style="background: #d1fae5; color: #047857; padding: 3px 8px; border-radius: 20px; font-size: 12px; font-weight: 600; display: inline-flex; align-items: center; justify-content:center; gap: 0;">
                        <i class="fa-solid fa-circle-check"></i>
                    </span>`;

                let ctSecondsTotal = parseFloat(matchedStudy.avgSeconds);
                if (isNaN(ctSecondsTotal)) ctSecondsTotal = parseTimeToSeconds(matchedStudy.averageCT);
                avgCtBadgeHtml = `
                    <span style="background: #e0f2fe; color: #0369a1; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 5px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-stopwatch"></i> CT: ${Math.round(ctSecondsTotal)}
                    </span>`;

                const pfmVal = parseFloat(matchedStudy.pfm) || 0;
                const pc = pfmColors(pfmVal);
                pfmBadgeHtml = `
                    <span style="background: ${pc.bg}; color: ${pc.color}; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                        <i class="fa-solid fa-chart-pie"></i> PFM: ${Math.round(pfmVal)}%
                    </span>`;

                deleteBtnHtml = `
                    <button type="button" class="delete-study-btn observe-refresh-icon" data-study-id="${escapeHtml(matchedStudy.id)}" title="Reset / ចាប់ផ្ដើមការចាប់នាទីសាថ្មី" aria-label="Refresh" style="background: #fee2e2; color: #dc2626; border: none; padding: 6px; border-radius: 8px; font-size: 15px; cursor: pointer; display: inline-flex; align-items: center; justify-content:center; gap: 0; transition: all 0.2s;">
                        <i class="fa-solid fa-arrows-rotate"></i>
                    </button>`;
            } else if (hasDraft) {
                checkmarkBadgeHtml = `
                    <span style="background: #fef3c7; color: #b45309; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; display: inline-flex; align-items: center; gap: 4px;">
                        <i class="fa-solid fa-floppy-disk"></i> Draft (${drafts[item.id].laps.length} Lap)
                    </span>`;
            }

            const leftTrashBtnHtml = `
                <button type="button" class="delete-entire-study-btn" data-index="${index}" title="លុបប្រអប់នេះចោលទាំងស្រុង" style="background: #fee2e2; color: #dc2626; border: none; padding: 10px 14px; border-radius: 8px; font-size: 16px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; transition: all 0.2s; flex-shrink: 0;">
                    <i class="fa-solid fa-trash-can" style="font-size: 16px;"></i>
                </button>`;

            html += `
                <div class="observe-item mobile-process-observe-item" style="display: flex; align-items: center; justify-content: space-between; padding: 14px 18px; margin-bottom: 12px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; transition: all 0.2s ease;">
                    <div style="display: flex; align-items: center; gap: 12px; width: 100%;">
                        ${leftTrashBtnHtml}

                        <div class="observe-item-left process-item-btn" data-index="${index}" style="flex-direction: column; align-items: flex-start; gap: 8px; width: 100%; cursor: pointer;">
                            <div class="mobile-process-title-row" style="font-size: 15px; font-weight: 600; color: #1e293b; display: flex; align-items: center; width: 100%; justify-content: space-between;">
                                <div>
                                    <span style="color: #2563eb; margin-right: 5px;">${escapeHtml(processNo)}.</span> ${escapeHtml(khmerText)}
                                </div>
                                <div class="mobile-process-title-status" style="display: flex; align-items: center; gap: 8px;">
                                    ${checkmarkBadgeHtml}
                                </div>
                            </div>

                            <span style="font-size: 13px; color: #64748b;">${escapeHtml(thaiText)}</span>

                            <div style="display: flex; gap: 15px; font-size: 13px; color: #475569; background: #f8fafc; padding: 8px 12px; border-radius: 8px; width: 100%; flex-wrap: wrap; align-items: center; border: 1px solid #f1f5f9;">
                                <span><b>Seq:</b> ${escapeHtml(item.seq || 99)}</span>
                                <span><b>Sewer:</b> ${escapeHtml(sewerNameText)}</span>
                                <span><b>MC:</b> ${escapeHtml(item.mc || '-')}</span>
                            </div>

                            <div class="mobile-process-primary-metrics" style="display: flex; width: 100%; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-top: 2px;">
                                <span style="font-size: 12px; color: #64748b;"><b>PrimaryKey:</b> ${escapeHtml(primaryKeyText)}</span>
                                <div class="mobile-process-metrics-row" style="display: flex; align-items: center; gap: 8px;">
                                    ${pfmBadgeHtml}
                                    ${avgCtBadgeHtml}
                                    <span class="observe-smv-badge" style="font-weight: bold; color: #584ff4; background: #6661cb5d; padding: 4px 10px; border-radius: 6px; border: 1px solid #fefefe; font-size: 13px;">
                                        SMV: ${smvText}
                                    </span>
                                    ${deleteBtnHtml}
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="observe-item-right process-item-btn" data-index="${index}" style="margin-left: 15px; color: #94a3b8; cursor: pointer;">
                        <i class="fa-solid fa-chevron-right"></i>
                    </div>
                </div>
            `;
        });
        container.innerHTML = html;

        container.querySelectorAll('.process-item-btn').forEach(btn => {
            btn.addEventListener('click', function () {
                const selectedItem = processItems[parseInt(this.getAttribute('data-index'))];
                if (selectedItem) window.openTimeStudyForm(selectedItem);
            });
        });

        container.querySelectorAll('.delete-study-btn').forEach(delBtn => {
            delBtn.addEventListener('click', async function (e) {
                e.stopPropagation();
                const studyId = this.getAttribute('data-study-id');
                const existingStudies = safeLoad('savedTimeStudies');
                const existingStudy = existingStudies.find(s => s.id === studyId);
                if (!existingStudy) return;
                if (!tspCanManageRecord(existingStudy)) { tspPermissionDenied(existingStudy); return; }
                const ok = await confirmDialog({
                    title: 'តើអ្នកពិតជាចង់ Reset?',
                    html: 'ឬចាប់ផ្ដើមការចាប់នាទី Process នេះសាថ្មីមែនទេ?',
                    icon: 'warning',
                    confirmText: 'យល់ព្រម'
                });
                if (!ok) return;

                let studies = safeLoad('savedTimeStudies');
                const targetStudy = studies.find(s => s.id === studyId); if (!targetStudy) return; if (!tspCanManageRecord(targetStudy)) { tspPermissionDenied(targetStudy); return; }
                studies = studies.filter(s => s.id !== studyId);
                safeSave('savedTimeStudies', studies);

                window.renderProcessList(title, section, styleName, rampupDay, prepName, filterText);
                window.renderPrepAssemblyList(title, section, styleName, rampupDay);
                successDialog('បានជោគជ័យ!', 'ទិន្នន័យត្រូវបាន Reset សាថ្មីរួចរាល់។');
            });
        });

        container.querySelectorAll('.delete-entire-study-btn').forEach(trashBtn => {
            trashBtn.addEventListener('click', async function (e) {
                e.stopPropagation();
                const targetItem = processItems[parseInt(this.getAttribute('data-index'))];
                if (!targetItem) return;
                const permissionFiles = safeLoad('uploadedExcelFiles');
                const permissionMatches = permissionFiles.filter(file => file.prepMap && Array.isArray(file.prepMap[compositeKey]) && file.prepMap[compositeKey].some(i => (i.id && i.id === targetItem.id) || (String(i.primaryKey || '') === String(targetItem.primaryKey || '') && normStr(i.processRegion || i.processName) === normStr(targetItem.processRegion || targetItem.processName) && normStr(i.sewerName) === normStr(targetItem.sewerName) && normStr(i.prepAssembly) === normStr(targetItem.prepAssembly))));
                if (!permissionMatches.some(file => tspCanManageRecord(file))) { tspPermissionDenied(permissionMatches[0]); return; }

                const ok = await confirmDialog({
                    title: 'តើអ្នកពិតជាចង់លុបប្រអប់នេះទេ?',
                    html: 'ប្រអប់ Process នេះនឹងត្រូវលុបចោលទាំងស្រុង។',
                    icon: 'error',
                    confirmColor: '#dc2626',
                    confirmText: 'លុបចោល'
                });
                if (!ok) return;

                const filesData = safeLoad('uploadedExcelFiles');
                if (!filesData.some(file => tspCanManageRecord(file) && file.prepMap && Array.isArray(file.prepMap[compositeKey]) && file.prepMap[compositeKey].some(i => (i.id && i.id === targetItem.id) || (String(i.primaryKey || '') === String(targetItem.primaryKey || '') && normStr(i.processRegion || i.processName) === normStr(targetItem.processRegion || targetItem.processName) && normStr(i.sewerName) === normStr(targetItem.sewerName) && normStr(i.prepAssembly) === normStr(targetItem.prepAssembly))))) { tspPermissionDenied(targetItem); return; }
                filesData.forEach(file => {
                    if (!tspCanManageRecord(file)) return;
                    if (file.prepMap && Array.isArray(file.prepMap[compositeKey])) {
                        const arr = file.prepMap[compositeKey];
                        let realIdx = arr.findIndex(i => i.id && i.id === targetItem.id);
                        if (realIdx === -1) {
                            realIdx = arr.findIndex(i =>
                                String(i.primaryKey || '') === String(targetItem.primaryKey || '') &&
                                normStr(i.processRegion || i.processName) === normStr(targetItem.processRegion || targetItem.processName) &&
                                normStr(i.sewerName) === normStr(targetItem.sewerName) &&
                                normStr(i.prepAssembly) === normStr(targetItem.prepAssembly));
                        }
                        if (realIdx !== -1) arr.splice(realIdx, 1);

                        // ប្រសិនបើ Prep/Assembly លែងមាន Process ទៀត ត្រូវលុប Group នោះចេញ
                        const prepNameStillExists = arr.some(i => normStr(i.prepAssembly) === normStr(targetItem.prepAssembly));
                        if (!prepNameStillExists) {
                            file.prepMap[compositeKey] = arr.filter(i => normStr(i.prepAssembly) !== normStr(targetItem.prepAssembly));
                        }
                        if (file.prepMap[compositeKey].length === 0) delete file.prepMap[compositeKey];
                    }
                });

                // សម្អាត Rampup Day ដែលគ្មាន Prep/Assembly/Process នៅសល់
                const rampupKey = `${title}_${section}_${styleName}`;
                const dayStillHasPrep = filesData.some(file => tspCanManageRecord(file) && file.prepMap && Array.isArray(file.prepMap[compositeKey]) && file.prepMap[compositeKey].length > 0);
                if (!dayStillHasPrep) {
                    filesData.forEach(file => {
                        if (!tspCanManageRecord(file)) return;
                        if (file.rampupMap && Array.isArray(file.rampupMap[rampupKey])) {
                            file.rampupMap[rampupKey] = file.rampupMap[rampupKey].filter(day => String(day) !== String(rampupDay));
                            if (file.rampupMap[rampupKey].length === 0) delete file.rampupMap[rampupKey];
                        }
                    });
                }
                // Wait for IndexedDB to commit deletion before updating the visible list.
                const filesSaved = await tspPersistLargeStorage('uploadedExcelFiles', filesData);
                if (!filesSaved) {
                    showCustomAlert('កំហុសរក្សាទុក', 'មិនអាចរក្សាទុកការលុប Process បានទេ។ សូមព្យាយាមម្ដងទៀត។', 'fa-solid fa-triangle-exclamation', '#dc2626');
                    return;
                }
                uploadedFilesList = filesData;

                let studies = safeLoad('savedTimeStudies');
                studies = studies.filter(s => !tspCanManageRecord(s) || !studyMatchesItem(s, targetItem, ctx));
                safeSave('savedTimeStudies', studies);

                // លុប Draft ដែលជាប់ទាក់ទង
                const d = loadObject('tsDrafts');
                if (targetItem.id && d[targetItem.id]) {
                    delete d[targetItem.id];
                    try { localStorage.setItem('tsDrafts', JSON.stringify(d)); } catch (err) { /* ignore */ }
                }

                const remainingInCurrentPrep = safeLoad('uploadedExcelFiles').some(file =>
                    file.prepMap && Array.isArray(file.prepMap[compositeKey]) &&
                    file.prepMap[compositeKey].some(i => normStr(i.prepAssembly) === normStr(prepName))
                );
                if (!remainingInCurrentPrep) {
                    const processPage = document.getElementById('process-page');
                    const prepPage = document.getElementById('prep-page');
                    if (processPage) processPage.style.display = 'none';
                    if (prepPage) prepPage.style.display = 'block';
                    window.renderPrepAssemblyList(title, section, styleName, rampupDay);
                } else {
                    window.renderProcessList(title, section, styleName, rampupDay, prepName, filterText);
                    window.renderPrepAssemblyList(title, section, styleName, rampupDay);
                }
                successDialog('បានលុបជោគជ័យ!', 'ប្រអប់ Process ត្រូវបានលុបចេញរួចរាល់។');
            });
        });
    };

    const processItemSearchInput = document.getElementById('process-item-search-input');
    if (processItemSearchInput) processItemSearchInput.addEventListener('input', (e) => {
        if (window.processContext) {
            const c = window.processContext;
            window.renderProcessList(c.title, c.section, c.styleName, c.rampupDay, c.prepName, e.target.value);
        }
    });

    // --- ៤. Time Study Form, Stopwatch, Lap Analytics, Draft & Shortcuts ---
    let stopwatchInterval = null;
    let stopwatchTime = 0;
    let stopwatchState = 'STOPPED';
    let lapRecords = [];
    let lastLapTotalTime = 0;
    let currentStartTimeStamp = '';
    let lastDraftSave = 0;

    function formatStopwatch(ms) {
        const minutes = Math.floor(ms / 60000);
        const seconds = Math.floor((ms % 60000) / 1000);
        const hundredths = Math.floor((ms % 1000) / 10);
        return String(minutes).padStart(2, '0') + ':' +
            String(seconds).padStart(2, '0') + '.' +
            String(hundredths).padStart(2, '0');
    }

    const mainActionBtn = document.getElementById('ts-main-action-btn');
    const secondaryActionBtn = document.getElementById('ts-lap-btn');
    const display = document.getElementById('stopwatch-display');

    // ---- Draft (Auto-save) ----
    function draftKey(ctx) {
        if (!ctx) return '';
        return ctx.id || [ctx.primaryKey, ctx.processRegion, ctx.sewerName, ctx.prepAssembly].join('|');
    }

    function saveDraft() {
        const ctx = window.timeStudyContext;
        const key = draftKey(ctx);
        if (!key) return;
        if (stopwatchTime === 0 && lapRecords.length === 0) return;
        const drafts = loadObject('tsDrafts');
        drafts[key] = {
            time: stopwatchTime,
            laps: lapRecords.map(l => ({ ...l })),
            lastLap: lastLapTotalTime,
            start: currentStartTimeStamp,
            savedAt: Date.now()
        };
        try { localStorage.setItem('tsDrafts', JSON.stringify(drafts)); } catch (e) { /* ignore */ }
        lastDraftSave = Date.now();
    }

    function clearDraft() {
        const key = draftKey(window.timeStudyContext);
        if (!key) return;
        const drafts = loadObject('tsDrafts');
        if (drafts[key]) {
            delete drafts[key];
            try { localStorage.setItem('tsDrafts', JSON.stringify(drafts)); } catch (e) { /* ignore */ }
        }
    }

    function tryRestoreDraft() {
        const key = draftKey(window.timeStudyContext);
        if (!key) return;
        const d = loadObject('tsDrafts')[key];
        if (!d || (!(d.laps && d.laps.length) && !(d.time > 0))) return;

        stopwatchTime = d.time || 0;
        lapRecords = Array.isArray(d.laps) ? d.laps : [];
        lastLapTotalTime = d.lastLap || 0;
        currentStartTimeStamp = d.start || '';
        pauseStopwatch();
        updateDisplay();
        renderLapTable();
        showToast('ស្ដារ Draft', 'បានស្ដារទិន្នន័យចាប់នាទីដែលមិនទាន់រក្សាទុក (' + lapRecords.length + ' Lap)', 'fa-solid fa-floppy-disk', '#f59e0b');
    }

    window.addEventListener('beforeunload', saveDraft);
    document.addEventListener('visibilitychange', () => {
        if (document.hidden) saveDraft();
        else if (stopwatchState === 'RUNNING') requestWakeLock();
    });

    // ---- Extra controls (Undo button + Stats panel) ----
    function ensureExtraControls() {
        if (secondaryActionBtn && !document.getElementById('ts-undo-lap-btn')) {
            const b = document.createElement('button');
            b.type = 'button';
            b.id = 'ts-undo-lap-btn';
            b.className = secondaryActionBtn.className;
            b.innerHTML = '<i class="fa-solid fa-delete-left"></i> Undo';
            b.title = 'លុប Lap ចុងក្រោយ (U)';
            b.style.cssText = 'background:#fef3c7;color:#b45309;border:none;cursor:pointer;margin-left:8px;';
            b.setAttribute('disabled', 'true');
            secondaryActionBtn.insertAdjacentElement('afterend', b);
            b.addEventListener('click', undoLastLap);
        }
        if (!document.getElementById('lap-stats-panel')) {
            const tbody = document.getElementById('lap-table-body');
            const table = tbody ? tbody.closest('table') : null;
            if (table && table.parentNode) {
                const panel = document.createElement('div');
                panel.id = 'lap-stats-panel';
                panel.style.cssText = 'margin:10px 0;';
                table.parentNode.insertBefore(panel, table);
            }
        }
    }

    function undoLastLap() {
        if (lapRecords.length === 0) return;
        lapRecords.pop();
        const last = lapRecords[lapRecords.length - 1];
        lastLapTotalTime = last && last.endMs !== undefined ? last.endMs : 0;
        renderLapTable();
        saveDraft();
        vibrate(30);
        showToast('Undo', 'បានលុប Lap ចុងក្រោយ', 'fa-solid fa-delete-left', '#f59e0b');
    }

    window.openTimeStudyForm = function (itemData) {
        document.getElementById('process-page').style.display = 'none';
        document.getElementById('timestudy-page').style.display = 'block';

        const validSewer = validSewerName(itemData.sewerName);
        const numSmv = parseFloat(itemData.smv);
        const formattedSmv = !isNaN(numSmv) ? numSmv.toFixed(2) : (itemData.smv || '0.00');

        document.getElementById('ts-process-khmer').textContent = itemData.processRegion || itemData.processName || 'N/A';
        document.getElementById('ts-process-thai').textContent = itemData.processName || '';
        document.getElementById('ts-primary-key').textContent = itemData.primaryKey || 'N/A';
        document.getElementById('ts-smv').textContent = formattedSmv;

        const sewerBadge = document.getElementById('ts-sewer-badge');
        if (sewerBadge) sewerBadge.textContent = validSewer;

        window.timeStudyContext = { ...itemData, smv: formattedSmv, sewerName: validSewer };
        ensureExtraControls();
        resetStopwatch(false);
        tryRestoreDraft();
    };

    const backToProcessBtn = document.getElementById('back-to-process-btn');
    if (backToProcessBtn) backToProcessBtn.addEventListener('click', () => {
        saveDraft(); // រក្សា Draft ទុក មិនបាត់ Lap
        resetStopwatch(false);
        document.getElementById('timestudy-page').style.display = 'none';
        document.getElementById('process-page').style.display = 'block';
        if (window.processContext) {
            const c = window.processContext;
            window.renderProcessList(c.title, c.section, c.styleName, c.rampupDay, c.prepName);
        }
    });

    function togglePlayPause() {
        if (stopwatchState === 'STOPPED' || stopwatchState === 'PAUSED') startStopwatch();
        else if (stopwatchState === 'RUNNING') pauseStopwatch();
    }

    if (mainActionBtn) mainActionBtn.addEventListener('click', togglePlayPause);

    if (secondaryActionBtn) {
        secondaryActionBtn.addEventListener('click', () => {
            if (stopwatchState === 'RUNNING') recordLap();
            else if (stopwatchState === 'PAUSED') resetStopwatch(true);
        });
    }

    function startStopwatch() {
        if (!mainActionBtn || !secondaryActionBtn) return;
        if (!currentStartTimeStamp || stopwatchState === 'STOPPED') {
            const now = new Date();
            currentStartTimeStamp = now.toLocaleDateString('km-KH') + ' ' + now.toLocaleTimeString('km-KH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        }

        stopwatchState = 'RUNNING';
        mainActionBtn.innerHTML = `<i class="fa-solid fa-pause"></i> Pause`;
        mainActionBtn.style.backgroundColor = '#f59e0b';
        secondaryActionBtn.innerHTML = `<i class="fa-solid fa-flag"></i> កត់ត្រា (Lap)`;
        secondaryActionBtn.style.backgroundColor = '#e2e8f0';
        secondaryActionBtn.style.color = '#1e293b';
        secondaryActionBtn.removeAttribute('disabled');

        beep(660, 60);
        requestWakeLock();

        clearInterval(stopwatchInterval);
        const startTime = Date.now() - stopwatchTime;
        stopwatchInterval = setInterval(() => {
            stopwatchTime = Date.now() - startTime;
            updateDisplay();
            if (Date.now() - lastDraftSave > 2000) saveDraft();
        }, 10);
    }

    function pauseStopwatch() {
        if (!mainActionBtn || !secondaryActionBtn) return;
        stopwatchState = 'PAUSED';
        mainActionBtn.innerHTML = `<i class="fa-solid fa-play"></i> Play`;
        mainActionBtn.style.backgroundColor = '#10b981';
        secondaryActionBtn.innerHTML = `<i class="fa-solid fa-rotate-left"></i> Restart`;
        secondaryActionBtn.style.backgroundColor = '#fee2e2';
        secondaryActionBtn.style.color = '#dc2626';
        secondaryActionBtn.removeAttribute('disabled');
        clearInterval(stopwatchInterval);
        releaseWakeLock();
        saveDraft();
    }

    // clearDraftFlag === false → រក្សា Draft ទុក; default → លុប Draft
    function resetStopwatch(clearDraftFlag) {
        clearInterval(stopwatchInterval);
        releaseWakeLock();
        if (clearDraftFlag !== false) clearDraft();
        stopwatchState = 'STOPPED';
        stopwatchTime = 0;
        lastLapTotalTime = 0;
        currentStartTimeStamp = '';
        lapRecords = [];
        updateDisplay();
        renderLapTable();

        if (mainActionBtn) {
            mainActionBtn.innerHTML = `<i class="fa-solid fa-play"></i> Play`;
            mainActionBtn.style.backgroundColor = '#10b981';
        }
        if (secondaryActionBtn) {
            secondaryActionBtn.innerHTML = `<i class="fa-solid fa-flag"></i> កត់ត្រា (Lap)`;
            secondaryActionBtn.style.backgroundColor = '#e2e8f0';
            secondaryActionBtn.style.color = '#94a3b8';
            secondaryActionBtn.setAttribute('disabled', 'true');
        }
    }

    function updateDisplay() {
        if (display) display.textContent = formatStopwatch(stopwatchTime);
    }

    function recordLap() {
        const currentTotalTime = stopwatchTime;
        const lapDuration = currentTotalTime - lastLapTotalTime;
        lastLapTotalTime = currentTotalTime;

        lapRecords.push({
            round: lapRecords.length + 1,
            time: formatStopwatch(lapDuration),
            durationMs: lapDuration,
            endMs: currentTotalTime
        });
        beep(880, 70);
        vibrate(40);
        renderLapTable();
        saveDraft();
    }

    function renderLapStats(stats) {
        const panel = document.getElementById('lap-stats-panel');
        if (!panel) return;

        const hint = `<div style="font-size:11px;color:#94a3b8;margin-top:6px;">⌨️ Space = Start / Lap &nbsp;•&nbsp; P = Pause / Play &nbsp;•&nbsp; U = Undo Lap &nbsp;•&nbsp; R = Restart (ពេល Pause)</div>`;

        if (stats.n === 0) {
            panel.innerHTML = hint;
            return;
        }

        const chip = (label, value, bg, color) =>
            `<span style="background:${bg};color:${color};padding:4px 10px;border-radius:6px;font-size:12px;font-weight:bold;">${label}: ${value}</span>`;


        panel.innerHTML = `
            <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;">
                ${chip('Avg', formatStopwatch(stats.avg), '#e0f2fe', '#0369a1')}
                ${chip('Min', formatStopwatch(stats.min), '#d1fae5', '#047857')}
                ${chip('Max', formatStopwatch(stats.max), '#fee2e2', '#dc2626')}
                ${chip('Std Dev', (stats.std / 1000).toFixed(2) + 's', '#ede9fe', '#6d28d9')}
                ${chip('Laps', stats.n, '#f1f5f9', '#475569')}
            </div>
            ${stats.n < 3 ? `<div style="font-size:12px;color:#b45309;margin-top:7px;">ⓘ Average CT បច្ចុប្បន្នគណនាពី ${stats.n} Lap ប៉ុណ្ណោះ — លទ្ធផលបណ្តោះអាសន្ន។ ចាប់ Lap បន្ថែម 3–5 ដង ដើម្បីឱ្យតម្លៃមានភាពជឿជាក់ជាងមុន។</div>` : (stats.n < 5 ? `<div style="font-size:12px;color:#64748b;margin-top:7px;">Average CT គណនាពី ${stats.n} Lap។ បន្ថែម Lap ដល់ 5 ដង ប្រសិនបើអាច ដើម្បីឱ្យតម្លៃមានស្ថិរភាពជាងមុន។</div>` : '')}
            ${hint}
        `;
    }

    function renderLapTable() {
        ensureExtraControls();
        const tbody = document.getElementById('lap-table-body');
        const undoBtn = document.getElementById('ts-undo-lap-btn');
        if (undoBtn) {
            if (lapRecords.length > 0) undoBtn.removeAttribute('disabled');
            else undoBtn.setAttribute('disabled', 'true');
        }

        const stats = calcLapStats(lapRecords);
        renderLapStats(stats);

        if (!tbody) return;

        if (lapRecords.length === 0) {
            tbody.innerHTML = `<tr><td colspan="3" style="text-align: center; color: #94a3b8;">មិនទាន់មានទិន្នន័យចាប់ម៉ោងនៅឡើយទេ</td></tr>`;
            return;
        }

        let html = '';
        lapRecords.forEach((lap, index) => {
            html += `
                <tr>
                    <td>Round ${lap.round}</td>
                    <td style="font-weight: 600; color: #2563eb;">${lap.time}</td>
                    <td>
                        <button type="button" onclick="deleteLap(${index})" style="background: none; border: none; color: #dc2626; cursor: pointer;"><i class="fa-solid fa-trash"></i></button>
                    </td>
                </tr>
            `;
        });
        tbody.innerHTML = html;
    }

    window.deleteLap = function (index) {
        lapRecords.splice(index, 1);
        lapRecords.forEach((lap, i) => lap.round = i + 1);
        renderLapTable();
        saveDraft();
    };

    // ---- Keyboard shortcuts ----
    document.addEventListener('keydown', (e) => {
        const page = document.getElementById('timestudy-page');
        if (!page || page.style.display === 'none' || page.style.display === '') return;
        if (e.ctrlKey || e.metaKey || e.altKey) return;

        const tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : '';
        if (tag === 'input' || tag === 'textarea' || tag === 'select' || (e.target && e.target.isContentEditable)) return;

        const modalOpen = ['custom-modal', 'timeObserveEditModal', 'editStudyModal', 'custom-confirm-modal'].some(id => {
            const m = document.getElementById(id);
            return m && m.style.display === 'flex';
        });
        if (modalOpen) return;

        const key = e.key.toLowerCase();
        if (e.code === 'Space' || key === ' ') {
            e.preventDefault();
            if (stopwatchState === 'RUNNING') recordLap();
            else startStopwatch();
        } else if (key === 'p') {
            e.preventDefault();
            togglePlayPause();
        } else if (key === 'u') {
            e.preventDefault();
            undoLastLap();
        } else if (key === 'r' && stopwatchState === 'PAUSED') {
            e.preventDefault();
            resetStopwatch(true);
        }
    });

    // --- ៥. Save Time Study ---
    const saveBtn = document.getElementById('ts-save-btn');
    if (saveBtn) {
        saveBtn.addEventListener('click', async () => {
            // ត្រូវមាន CT Lap យ៉ាងហោចណាស់ 3 ដង ទើបអាច Save បាន
            if (lapRecords.length < 3) {
                showCustomAlert('មិនអាចរក្សាទុកបានទេ!', 'សូមចាប់ CT Lap យ៉ាងហោចណាស់ 3 ដងជាមុនសិន។ បច្ចុប្បន្នមាន ' + lapRecords.length + ' Lap ប៉ុណ្ណោះ។', 'fa-solid fa-triangle-exclamation', '#f59e0b');
                return;
            }

            const stats = calcLapStats(lapRecords);

            if (stats.n < 1 || stats.avg <= 0) {
                showCustomAlert('មិនអាចរក្សាទុកបានទេ!', 'មិនមាន Lap ដែលមានពេលវេលាត្រឹមត្រូវទេ។ សូមចាប់ Lap ម្តងទៀត។', 'fa-solid fa-triangle-exclamation', '#f59e0b');
                return;
            }
            // Use the same validated durations as the live Average CT panel.
            const avgMs = stats.avg;
            const totalSeconds = avgMs / 1000;
            const avgFormattedTime = formatToMinSec(avgMs, true);

            const context = window.timeStudyContext || {};
            const smvVal = parseFloat(context.smv) || 0;

            let pfmPercent = 0;
            if (totalSeconds > 0 && smvVal > 0) pfmPercent = (smvVal / totalSeconds) * 100;

            let outputPcs = 0;
            if (totalSeconds > 0) outputPcs = Math.round(3600 / totalSeconds);

            const now = new Date();
            const dateTimeEndVal = now.toLocaleDateString('km-KH') + ' ' + now.toLocaleTimeString('km-KH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
            const startTimeVal = currentStartTimeStamp || dateTimeEndVal;

            const studyRecord = {
                id: 'TS_' + Date.now(),
                date: dateTimeEndVal,
                dateTimeStart: startTimeVal,
                dateTimeEnd: dateTimeEndVal,
                title: context.title || 'N/A',
                factory: context.title || context.factory || 'N/A',
                section: context.section || 'N/A',
                styleName: context.styleName || 'N/A',
                rampupDay: context.rampupDay || 'N/A',
                group: context.group || '',
                prepName: context.prepName || context.prepAssembly || 'N/A',
                prepAssembly: context.prepAssembly || context.prepName || 'N/A',
                sewerName: context.sewerName || 'No Sewer',
                processRegion: context.processRegion || 'N/A',
                processName: context.processName || 'N/A',
                primaryKey: context.primaryKey || 'N/A',
                smv: context.smv || '0.00',
                no: context.no || '-',
                seq: parseInt(context.seq) || 99,
                operatorSeq: parseInt(context.seq) || 99,
                mc: context.mc || '-',
                weight: context.weight || '-',
                rawRowData: context.rawRowData || {},
                laps: lapRecords.map(l => ({ ...l })),
                averageCT: avgFormattedTime,
                avgSeconds: totalSeconds,
                minSeconds: stats.min / 1000,
                maxSeconds: stats.max / 1000,
                stdSeconds: stats.std / 1000,
                outlierCount: stats.outliers.length,
                output: outputPcs,
                pfm: pfmPercent.toFixed(1)
            };

            const savedStudies = safeLoad('savedTimeStudies');
            studyRecord.ownerId = (tspCurrentPermissionUser() || {}).id || '';
            const existingIndex = savedStudies.findIndex(s => studyMatchesItem(s, studyRecord, studyRecord) && tspCanManageRecord(s));
            if (existingIndex !== -1) {
                studyRecord.id = savedStudies[existingIndex].id || studyRecord.id;
                savedStudies[existingIndex] = studyRecord;
            } else {
                savedStudies.unshift(studyRecord);
            }

            if (!safeSave('savedTimeStudies', savedStudies)) return;

            showToast('ជោគជ័យ!', 'ទិន្នន័យ Time Study ត្រូវបានរក្សាទុក', 'fa-solid fa-circle-check', '#10b981');
            beep(1040, 120);
            resetStopwatch(true); // លុប Draft ព្រោះបាន Save រួច

            if (window.prepContext) {
                const p = window.prepContext;
                window.renderPrepAssemblyList(p.title, p.section, p.styleName, p.rampupDay);
            }
            if (window.processContext) {
                const c = window.processContext;
                window.renderProcessList(c.title, c.section, c.styleName, c.rampupDay, c.prepName);
            }

            document.getElementById('timestudy-page').style.display = 'none';
            document.getElementById('process-page').style.display = 'block';
        });
    }

    // --- Modal កែសម្រួលទិន្នន័យ (Time Observe) ---
    window.openTimeObserveEditModal = function () {
        let studyData = window.timeStudyContext;

        if (!studyData || !studyData.primaryKey) {
            const pkElement = document.getElementById('ts-primary-key');
            const pkText = pkElement ? pkElement.textContent.trim() : '';
            studyData = safeLoad('savedTimeStudies').find(item => item.primaryKey === pkText);
        }

        if (!studyData) {
            alert('សូមជ្រើសរើស Process ណាមួយជាមុនសិន!');
            return;
        }

        let studyCopy = JSON.parse(JSON.stringify(studyData));

        const savedStudies = safeLoad('savedTimeStudies');
        let foundStudy = studyCopy.id ? savedStudies.find(item => item.id === studyCopy.id) : null;
        if (!foundStudy) foundStudy = savedStudies.find(s => studyMatchesItem(s, studyCopy, studyCopy));
        if (foundStudy) studyCopy = JSON.parse(JSON.stringify(foundStudy));

        window.timeStudyContext = studyCopy;
        window.currentEditMode = 'observe';

        const uploadedFiles = safeLoad('uploadedExcelFiles');
        const fileObj = uploadedFiles[0];
        const baseHeaders = fileObj && fileObj.excelHeaders && fileObj.excelHeaders.length
            ? fileObj.excelHeaders
            : ['Style', 'Process', 'Process(Region)', 'StdCode', 'NGIETYPE', 'NGIEDescription', 'NGIESMV', 'MC', 'SMV', 'Weight', 'NGIEaudit'];

        const nonCtHeaders = baseHeaders.filter(h => {
            const cleanH = String(h).trim().toLowerCase().replace(/[\s_]+/g, '');
            return !/^ct\d*$/.test(cleanH);
        });
        // Always expose Sewer Name in Edit Data, even when the source workbook has no Sewer column.
        const hasSewerHeader = nonCtHeaders.some(h => String(h).trim().toLowerCase().replace(/[^a-z0-9]/g, '').includes('sewer'));
        if (!hasSewerHeader) nonCtHeaders.push('Sewer Name');

        const currentLaps = (studyCopy.laps && Array.isArray(studyCopy.laps)) ? studyCopy.laps : [];
        const dynamicCtHeaders = currentLaps.map((_, i) => 'CT' + (i + 1));

        const excelHeaders = [...nonCtHeaders, ...dynamicCtHeaders];
        const rawRow = studyCopy.rawRowData || {};

        const container = document.getElementById('dynamic-edit-inputs-container');
        if (container) {
            let html = '';
            let ctIndex = 0;

            excelHeaders.forEach(header => {
                let val = '';
                const cleanHeader = String(header).trim().toLowerCase().replace(/[\s_]+/g, '');
                const isCt = /^ct\d*$/.test(cleanHeader);

                if (isCt) {
                    const lapItem = currentLaps[ctIndex];
                    if (lapItem !== undefined) {
                        const duration = lapItem.durationMs !== undefined ? lapItem.durationMs : (lapItem.duration || 0);
                        val = formatToMinSec(duration, true);
                    }
                    ctIndex++;
                    val = formatToMinSec(val);
                } else {
                    const matchedKey = Object.keys(rawRow).find(k => k.trim().toLowerCase().replace(/[\s_]+/g, '') === cleanHeader);

                    if (matchedKey && rawRow[matchedKey] !== undefined && rawRow[matchedKey] !== null && String(rawRow[matchedKey]).trim() !== '') {
                        val = rawRow[matchedKey];
                    } else if (cleanHeader === 'no') val = studyCopy.no || '';
                    else if (cleanHeader === 'seq' || cleanHeader === 'sequence') val = studyCopy.seq || studyCopy.operatorSeq || '';
                    else if (cleanHeader === 'mc' || cleanHeader === 'machine') val = studyCopy.mc || '';
                    else if (cleanHeader === 'weight' || cleanHeader === 'wt') val = studyCopy.weight || '';
                    else if (cleanHeader.includes('datetime') || cleanHeader.includes('start')) val = studyCopy.dateTimeStart || '';
                    else if (cleanHeader.includes('title') || cleanHeader.includes('factory')) val = studyCopy.title || studyCopy.factory || '';
                    else if (cleanHeader.includes('section')) val = studyCopy.section || '';
                    else if (cleanHeader.includes('style')) val = studyCopy.styleName || '';
                    else if (cleanHeader.includes('rampup')) val = studyCopy.rampupDay || '';
                    else if (cleanHeader.includes('prep') || cleanHeader.includes('assembly')) val = studyCopy.prepAssembly || studyCopy.prepName || '';
                    else if (cleanHeader.includes('sewer')) {
                        const rawSewerKey = Object.keys(rawRow).find(k => String(k).trim().toLowerCase().replace(/[^a-z0-9]/g, '').includes('sewer'));
                        val = studyCopy.sewerName || studyCopy.sewer || (rawSewerKey ? rawRow[rawSewerKey] : '') || '';
                    }
                    else if (cleanHeader.includes('primary')) val = studyCopy.primaryKey || '';
                    else if (cleanHeader === 'smv') val = studyCopy.smv || '';
                    else if (cleanHeader.includes('process')) val = studyCopy.processRegion || studyCopy.processName || '';
                }

                html += `
                    <div style="display: flex; flex-direction: column; gap: 5px;">
                        <label style="font-size: 12px; color: #334155; font-weight: bold;">${escapeHtml(header)}</label>
                        <input type="text" class="dynamic-obs-input" data-header="${escapeHtml(header)}" value="${escapeHtml(val)}" style="padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 13px; outline: none; width: 100%;">
                    </div>
                `;
            });
            container.innerHTML = html;
        }

        const modal = document.getElementById('timeObserveEditModal') || document.getElementById('editStudyModal');
        if (modal) modal.style.display = 'flex';
    };

    window.closeEditModal = window.closeTimeObserveEditModal = function () {
        ['timeObserveEditModal', 'editStudyModal'].forEach(mid => {
            const m = document.getElementById(mid);
            if (m) m.style.display = 'none';
        });
    };

    window.saveTimeObserveEdit = function () {
        const inputs = document.querySelectorAll('.dynamic-obs-input');
        const studyData = window.timeStudyContext;
        if (!studyData) return;
        const sourceFiles = safeLoad('uploadedExcelFiles');
        const hasOwnedSource = sourceFiles.some(file => tspCanManageRecord(file) && Object.values(file.prepMap || {}).some(rows => Array.isArray(rows) && rows.some(item => (studyData.id && item.id === studyData.id) || (String(item.primaryKey || '') === String(studyData.primaryKey || '') && normStr(item.processRegion || item.processName) === normStr(studyData.processRegion || studyData.processName) && normStr(item.sewerName) === normStr(studyData.sewerName) && normStr(item.prepAssembly) === normStr(studyData.prepAssembly || studyData.prepName)))));
        if (!hasOwnedSource && (tspCurrentPermissionUser() || {}).role !== 'Admin') { tspPermissionDenied(); return; }

        if (!studyData.rawRowData) studyData.rawRowData = {};
        const updatedLaps = [];
        let hasCtFieldInModal = false;

        inputs.forEach(input => {
            const header = input.getAttribute('data-header');
            const val = input.value;
            studyData.rawRowData[header] = val;

            const cleanH = header.toLowerCase().replace(/[\s_]+/g, '');

            if (/^ct\d*$/.test(cleanH)) {
                hasCtFieldInModal = true;
                if (val && val.trim() !== '' && val.trim() !== '00:00') {
                    let totalMs = 0;
                    if (val.includes(':')) {
                        const parts = val.split(':');
                        totalMs = ((parseFloat(parts[0]) || 0) * 60 + (parseFloat(parts[1]) || 0)) * 1000;
                    } else {
                        totalMs = (parseFloat(val) || 0) * 1000;
                    }
                    if (totalMs > 0) updatedLaps.push({ durationMs: Math.round(totalMs) });
                }
            }
            else if (cleanH === 'no') studyData.no = val;
            else if (cleanH === 'seq' || cleanH === 'sequence') { studyData.seq = val; studyData.operatorSeq = parseInt(val) || 99; }
            else if (cleanH === 'mc' || cleanH === 'machine') studyData.mc = val;
            else if (cleanH === 'weight' || cleanH === 'wt') studyData.weight = val;
            else if (cleanH.includes('datetime') || cleanH.includes('start')) studyData.dateTimeStart = val;
            else if (cleanH.includes('title') || cleanH.includes('factory')) { studyData.title = val; studyData.factory = val; }
            else if (cleanH.includes('section')) studyData.section = val;
            else if (cleanH.includes('style')) studyData.styleName = val;
            else if (cleanH.includes('rampup')) studyData.rampupDay = val;
            else if (cleanH.includes('prep') || cleanH.includes('assembly')) { studyData.prepAssembly = val; studyData.prepName = val; }
            else if (cleanH === 'sewer' || cleanH === 'sewername' || cleanH.startsWith('sewername') || cleanH.startsWith('seweroperator')) {
                // Only explicitly named Sewer columns may update the Sewer Name.
                // ID Card / ID columns must never overwrite this field.
                studyData.sewerName = validSewerName(val);
                studyData.sewer = studyData.sewerName;
                studyData.rawRowData[header] = studyData.sewerName;
            }
            else if (cleanH.includes('primarykey') || cleanH === 'primary_key' || cleanH === 'id') studyData.primaryKey = val;
            else if (cleanH === 'smv' || cleanH.includes('ngiesmv')) studyData.smv = val;
            else if (cleanH.includes('process') && cleanH.includes('region')) studyData.processRegion = val;
            else if (cleanH === 'process' || cleanH === 'processname') studyData.processName = val;
            else if (cleanH.includes('process')) studyData.processRegion = val;
        });

        // កុំអនុញ្ញាតឱ្យ Save ការកែ Time Observe ប្រសិនបើ CT Lap តិចជាង 3
        if (hasCtFieldInModal && updatedLaps.length < 3) {
            showCustomAlert('មិនអាចរក្សាទុកបានទេ!', 'សូមបញ្ចូល CT Lap ដែលមានពេលវេលាត្រឹមត្រូវយ៉ាងហោចណាស់ 3 ដង។ បច្ចុប្បន្នមាន ' + updatedLaps.length + ' Lap ប៉ុណ្ណោះ។', 'fa-solid fa-triangle-exclamation', '#f59e0b');
            return;
        }

        if (hasCtFieldInModal) {
            studyData.laps = updatedLaps.map((l, i) => ({ round: i + 1, durationMs: l.durationMs }));

            if (updatedLaps.length > 0) {
                const st = calcLapStats(studyData.laps);
                const totalSeconds = st.avg / 1000;
                const smvNum = parseFloat(studyData.smv) || 0;

                studyData.avgSeconds = totalSeconds;
                studyData.averageCT = formatToMinSec(st.avg, true);
                studyData.minSeconds = st.min / 1000;
                studyData.maxSeconds = st.max / 1000;
                studyData.stdSeconds = st.std / 1000;
                studyData.outlierCount = st.outliers.length;
                studyData.pfm = (totalSeconds > 0 && smvNum > 0) ? ((smvNum / totalSeconds) * 100).toFixed(1) : '0.0';
                studyData.output = totalSeconds > 0 ? Math.round(3600 / totalSeconds) : 0;
            } else {
                studyData.avgSeconds = 0;
                studyData.averageCT = '00:00';
                studyData.minSeconds = 0;
                studyData.maxSeconds = 0;
                studyData.stdSeconds = 0;
                studyData.outlierCount = 0;
                studyData.pfm = '0.0';
                studyData.output = 0;
            }
        }

        const savedStudies = safeLoad('savedTimeStudies');
        let index = studyData.id ? savedStudies.findIndex(item => item.id === studyData.id) : -1;
        if (index !== -1 && !tspCanManageRecord(savedStudies[index])) { tspPermissionDenied(); return; }
        if (index === -1) index = savedStudies.findIndex(s => studyMatchesItem(s, studyData, studyData) && tspCanManageRecord(s));

        if (index !== -1) {
            studyData.id = savedStudies[index].id;
            studyData.ownerId = savedStudies[index].ownerId || (tspCurrentPermissionUser() || {}).id || ''; 
            savedStudies[index] = studyData;
        } else {
            studyData.id = 'TS_' + Date.now();
            studyData.ownerId = (tspCurrentPermissionUser() || {}).id || '';
            savedStudies.push(studyData);
        }
        if (!safeSave('savedTimeStudies', savedStudies)) {
            showCustomAlert('កំហុសរក្សាទុក', 'មិនអាចរក្សាទុកការកែទិន្នន័យបានទេ។ សូមព្យាយាមម្ដងទៀត។', 'fa-solid fa-triangle-exclamation', '#dc2626');
            return;
        }
        if (typeof window.tspRecordRecentActivity === 'function') {
            const processLabel = studyData.processRegion || studyData.processName || 'Process';
            window.tspRecordRecentActivity('Time Observe data edited', studyData.title || studyData.factory || 'Time Study', `${processLabel}${studyData.sewerName ? ' · Sewer: ' + studyData.sewerName : ''}`);
        }

        window.closeEditModal();

        setTimeout(() => {
            const dash = document.getElementById('dashboard-page');
            if (dash && dash.style.display === 'block') window.renderDashboardStats();
            if (window.processContext) {
                const c = window.processContext;
                window.renderProcessList(c.title, c.section, c.styleName, c.rampupDay, c.prepName);
            }
            if (window.prepContext) {
                const pc = window.prepContext;
                window.renderPrepAssemblyList(pc.title, pc.section, pc.styleName, pc.rampupDay);
            }
        }, 150);

        showToast('ជោគជ័យ', 'ទិន្នន័យត្រូវបានរក្សាទុកដោយសុវត្ថិភាព!', 'fa-solid fa-check', '#10b981');
    };

    // --- ៦. Dashboard, KPI, Line Balancing, Yamazumi & Export ---
    const studyFactory = (s) => s.title || s.factory || '';

    const studyGroup = (s) => {
        if (!s) return '';
        const direct = s.group || s.groupName || s.groupLabel || s.workGroup || s.team;
        if (direct != null && String(direct).trim() !== '') return String(direct).trim();
        const raw = s.rawRowData;
        if (raw && typeof raw === 'object') {
            const key = Object.keys(raw).find(k => {
                const h = String(k).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
                return ['group','groupname','workgroup','team','sectiongroup','sectiongroupname'].includes(h) || (h.includes('group') && !h.includes('grouping'));
            });
            if (key && raw[key] != null && String(raw[key]).trim() !== '') return String(raw[key]).trim();
        }
        // Compatibility with earlier workbook versions: Prep, Ass, Ass1, Center etc.
        // may be stored in the Prep/Assembly field rather than the explicit Group field.
        const prep = s.prepAssembly || s.prepName;
        return prep != null ? String(prep).trim() : '';
    };

    const FILTER_DEFS = {
        factory: { id: 'filter-factory', label: 'Factory', get: studyFactory },
        section: { id: 'filter-section', label: 'Section', get: (s) => s.section || '' },
        style:   { id: 'filter-style',   label: 'Style',   get: (s) => s.styleName || '' },
        rampup:  { id: 'filter-rampup',  label: 'Day (D)', get: (s) => s.rampupDay || '' },
        group:   { id: 'filter-group',   label: 'Group',   get: studyGroup }
    };

    function injectDashboardStyles() {
        if (document.getElementById('dashboard-dynamic-style')) return;
        const st = document.createElement('style');
        st.id = 'dashboard-dynamic-style';
        st.textContent = `
            @keyframes dashSpin { to { transform: rotate(360deg); } }
            .filter-refresh-btn { background: #e0f2fe; color: #0284c7; border: none; width: 24px; height: 24px; border-radius: 6px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; font-size: 12px; transition: all .2s; flex-shrink: 0; }
            .filter-refresh-btn:hover { background: #bae6fd; }
            .filter-refresh-btn.spinning i { animation: dashSpin .6s linear infinite; }
            .dash-export-btn { border: none; padding: 8px 14px; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; transition: all .2s; }
            .dash-export-btn:hover { filter: brightness(0.95); transform: translateY(-1px); }
            .takt-input { width: 90px; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 13px; outline: none; }
            /* Modern dashboard KPI cards — visual-only enhancement */
            #dash-kpi-container { align-items: stretch; }
            .dash-kpi-card { position: relative; isolation: isolate; display: flex; align-items: center; gap: 13px; min-width: 0; min-height: 104px; padding: 17px 16px 17px 19px; overflow: hidden; border: 1px solid #e2e8f0; border-radius: 15px; background: linear-gradient(135deg,#fff 0%,#fbfdff 100%); box-shadow: 0 3px 10px rgba(15,23,42,.045); transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease; }
            .dash-kpi-card::before { content: ''; position: absolute; z-index: -1; inset: 0 auto 0 0; width: 4px; border-radius: 15px 0 0 15px; background: var(--kpi-accent,#2563eb); }
            .dash-kpi-card::after { content: ''; position: absolute; z-index: -1; width: 100px; height: 100px; right: -40px; top: -42px; border-radius: 50%; background: var(--kpi-accent,#2563eb); opacity: .045; }
            .dash-kpi-card:hover { transform: translateY(-3px); border-color: color-mix(in srgb, var(--kpi-accent,#2563eb) 30%, #e2e8f0); box-shadow: 0 10px 24px rgba(15,23,42,.09); }
            .dash-kpi-icon { display: flex; flex: 0 0 46px; align-items: center; justify-content: center; width: 46px; height: 46px; border: 1px solid color-mix(in srgb, var(--kpi-accent,#2563eb) 12%, #fff); border-radius: 13px; background: color-mix(in srgb, var(--kpi-accent,#2563eb) 10%, #fff); color: var(--kpi-accent,#2563eb); font-size: 19px; }
            .dash-kpi-copy { min-width: 0; flex: 1; }
            .dash-kpi-title { margin-bottom: 4px; color: #64748b; font-size: 11px; line-height: 1.4; font-weight: 700; letter-spacing: .015em; }
            .dash-kpi-value { color: #0f172a; font-size: clamp(17px,1.5vw,21px); line-height: 1.25; font-weight: 800; font-variant-numeric: tabular-nums; overflow-wrap: anywhere; }
            .dash-kpi-sub { margin-top: 5px; color: #94a3b8; font-size: 10.5px; line-height: 1.4; overflow-wrap: anywhere; }
            .dash-kpi-card.kpi-pfm { background: linear-gradient(135deg,#fff 0%,#f0fdf8 100%); }
            .dash-kpi-card.kpi-pfm .dash-kpi-value { color: #047857; }
            .dash-kpi-card.kpi-slowest { background: linear-gradient(135deg,#fff 0%,#fff7f7 100%); }
            .dash-kpi-card.kpi-slowest .dash-kpi-value { color: #b91c1c; }
            .oa-open-btn { background: linear-gradient(135deg,#2563eb,#4f46e5); color:#fff; box-shadow:0 3px 10px rgba(37,99,235,.18); }
            .oa-overlay { position:fixed; inset:0; z-index:99999; display:flex; align-items:center; justify-content:center; padding:18px; background:rgba(15,23,42,.62); backdrop-filter:blur(5px); }
            .oa-dialog { width:min(1080px,100%); max-height:92vh; display:flex; flex-direction:column; overflow:hidden; border:1px solid #e2e8f0; border-radius:20px; background:#f8fafc; box-shadow:0 25px 80px rgba(0,0,0,.28); }
            .oa-topbar { display:flex; align-items:flex-start; justify-content:space-between; gap:16px; padding:22px 24px 18px; color:#fff; background:linear-gradient(120deg,#1d4ed8,#4338ca); }
            .oa-eyebrow { font-size:10px; font-weight:800; letter-spacing:.15em; opacity:.8; }
            .oa-topbar h2 { margin:5px 0 4px; font-size:clamp(18px,2.5vw,24px); color:#fff; }
            .oa-topbar p { margin:0; color:#dbeafe; font-size:13px; }
            .oa-close { flex:0 0 36px; width:36px; height:36px; border:1px solid rgba(255,255,255,.25); border-radius:10px; background:rgba(255,255,255,.12); color:#fff; font-size:19px; cursor:pointer; }
            .oa-summary { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px; padding:14px 22px 0; }
            .oa-summary>div { display:flex; flex-direction:column; gap:3px; padding:12px 14px; border:1px solid #e2e8f0; border-radius:12px; background:#fff; }
            .oa-summary span { color:#64748b; font-size:11px; font-weight:700; }
            .oa-summary strong { color:#0f172a; font-size:20px; }
            .oa-toolbar { display:flex; align-items:center; gap:12px; padding:12px 22px 0; min-width:0; }
            .oa-summary-refresh { display:flex; justify-content:flex-start; flex:0 0 auto; padding:0; }
            .oa-summary-refresh button { display:inline-flex; align-items:center; gap:7px; border:1px solid #bfdbfe; border-radius:8px; padding:10px 14px; white-space:nowrap; background:#eff6ff; color:#1d4ed8; font-weight:700; font-size:12px; cursor:pointer; }
            .oa-summary-refresh button:hover { background:#dbeafe; }
            .oa-search-wrap { display:flex; align-items:center; gap:10px; flex:1 1 auto; min-width:0; margin:0; padding:0 13px; border:1px solid #cbd5e1; border-radius:11px; background:#fff; color:#64748b; }
            .oa-search-wrap input { width:100%; min-width:0; padding:10px 0; border:0; outline:0; background:transparent; color:#0f172a; font-size:13px; }
            .oa-content { min-height:0; overflow:auto; padding:0 22px 18px; }
            .oa-worker-card { margin-bottom:14px; overflow:hidden; border:1px solid #e2e8f0; border-radius:14px; background:#fff; box-shadow:0 2px 5px rgba(15,23,42,.025); }
            .oa-worker-heading { display:flex; align-items:center; gap:11px; padding:14px 16px; border-bottom:1px solid #eef2f7; background:linear-gradient(90deg,#eff6ff,#fff); }
            .oa-avatar { display:flex; align-items:center; justify-content:center; width:38px; height:38px; border-radius:12px; background:#dbeafe; color:#1d4ed8; }
            .oa-worker-heading h3 { margin:0; color:#0f172a; font-size:15px; }
            .oa-worker-heading p { margin:3px 0 0; color:#64748b; font-size:11px; }
            .oa-table { width:100%; border-collapse:collapse; text-align:left; }
            .oa-table th { padding:10px 15px; background:#f8fafc; color:#64748b; font-size:10px; text-transform:uppercase; letter-spacing:.04em; }
            .oa-table td { padding:12px 15px; border-top:1px solid #f1f5f9; color:#334155; font-size:13px; }
            .oa-table td:nth-child(2) { color:#047857; }
            .oa-number { white-space:nowrap; font-weight:800; font-variant-numeric:tabular-nums; }
            .oa-number small { color:#94a3b8; font-size:10px; font-weight:600; }
            .oa-meta { margin-top:4px; color:#94a3b8; font-size:10px; font-weight:400; }
            .oa-empty { padding:36px 12px; text-align:center; color:#64748b; }
            .oa-empty i { font-size:25px; color:#2563eb; }
            .oa-footer { display:flex; align-items:center; justify-content:space-between; gap:12px; padding:12px 22px; border-top:1px solid #e2e8f0; background:#fff; color:#64748b; font-size:11px; }
            @media(max-width:600px) { .oa-overlay { padding:7px; } .oa-dialog { max-height:96vh; border-radius:14px; } .oa-topbar { padding:17px 15px 14px; } .oa-summary { gap:6px; padding:10px 12px 0; } .oa-summary>div { padding:9px; } .oa-summary strong { font-size:17px; } .oa-search-wrap { margin:10px 12px; } .oa-content { padding:0 12px 12px; } .oa-table th,.oa-table td { padding:10px 8px; font-size:11px; } .oa-footer { padding:10px 12px; font-size:10px; } }
            @media (max-width: 600px) {
                .oa-toolbar { gap:8px; padding:10px 12px 0; }
                .oa-summary-refresh button { padding:9px 10px; font-size:11px; }
                .oa-search-wrap { gap:7px; }
                .oa-search-wrap input { font-size:12px; } .dash-kpi-card { min-height: 90px; gap: 10px; padding: 13px 11px 13px 15px; border-radius: 12px; } .dash-kpi-icon { flex-basis: 38px; width: 38px; height: 38px; border-radius: 11px; font-size: 16px; } .dash-kpi-title { font-size: 10.5px; } .dash-kpi-value { font-size: 17px; } .dash-kpi-sub { font-size: 10px; } }
            @media (prefers-reduced-motion: reduce) { .dash-kpi-card { transition: none; } .dash-kpi-card:hover { transform: none; } }

            /* Mobile-only Dashboard layout. Desktop rules and Dashboard logic remain unchanged. */
            @media (max-width: 600px) {
                #dashboard-page { min-width: 0; max-width: 100%; overflow-x: clip; }
                #dashboard-page .observe-container { min-width: 0; width: 100%; max-width: 100%; padding: 8px 10px calc(92px + env(safe-area-inset-bottom, 0px)); box-sizing: border-box; }
                #dashboard-page .observe-header { align-items: stretch !important; gap: 12px !important; margin-bottom: 12px; }
                #dashboard-page .observe-header h2 { width: 100%; margin: 0; font-size: 18px; line-height: 1.35; }
                #dashboard-page .observe-header > div { display: grid !important; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 7px !important; width: 100%; }
                #dashboard-page .observe-header .dash-export-btn { justify-content: center; min-width: 0; padding: 10px 7px; font-size: 11px; white-space: normal; }
                #dashboard-page .cards-grid { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; gap: 8px !important; margin-top: 12px !important; }
                #dashboard-page .cards-grid > .card { min-width: 0; padding: 12px 10px; border-radius: 13px; gap: 8px; }
                #dashboard-page .cards-grid .card-info h3 { font-size: 11px; line-height: 1.35; }
                #dashboard-page .cards-grid .card-info p { font-size: 21px; line-height: 1.2; }
                #dashboard-page .dash-filter-grid { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; gap: 8px !important; margin-top: 14px !important; padding: 10px !important; border-radius: 14px !important; }
                #dashboard-page .dash-filter-grid > * { min-width: 0; }
                #dashboard-page #dash-kpi-container { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; gap: 8px !important; margin-top: 14px !important; }
                #dashboard-page .dash-kpi-card { min-width: 0; min-height: 84px; gap: 8px; padding: 11px 9px 11px 11px; }
                #dashboard-page .dash-kpi-icon { flex: 0 0 32px; width: 32px; height: 32px; font-size: 14px; }
                #dashboard-page .dash-kpi-title { font-size: 10px; }
                #dashboard-page .dash-kpi-value { font-size: 16px; overflow-wrap: anywhere; }
                #dashboard-page .dash-kpi-sub { font-size: 9px; }
                #dashboard-page .uploaded-files-section { min-width: 0; max-width: 100%; margin-top: 14px !important; padding: 12px !important; border-radius: 14px !important; box-sizing: border-box; }
                #dashboard-page .uploaded-files-section h3 { font-size: 14px; line-height: 1.45; }
                #dashboard-page .uploaded-files-section > div[style*="height: 340px"] { height: 250px !important; max-width: 100%; }
                #dashboard-page .uploaded-files-section > div[style*="height: 340px"] canvas { max-width: 100% !important; }
                #dashboard-page #yamazumi-kpi-summary { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; gap: 7px !important; }
                #dashboard-page .takt-input { width: 100%; max-width: 100%; box-sizing: border-box; font-size: 16px; }
                #dashboard-page .table-responsive { display: block; width: 100%; max-width: 100%; overflow-x: auto !important; overflow-y: hidden; -webkit-overflow-scrolling: touch; overscroll-behavior-x: contain; touch-action: pan-x pan-y; border-radius: 10px; }
                #dashboard-page #dashboard-main-table { min-width: 1050px; width: max-content; max-width: none; }
                #dashboard-page #dashboard-main-table th, #dashboard-page #dashboard-main-table td { white-space: nowrap; }
                #dashboard-page .table-responsive::-webkit-scrollbar { height: 5px; }
                #dashboard-page .table-responsive::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 10px; }
                #dashboard-page .dash-export-btn, #dashboard-page button { touch-action: manipulation; }
            }
            @media (max-width: 380px) {
                #dashboard-page .observe-container { padding-left: 8px; padding-right: 8px; }
                #dashboard-page .cards-grid > .card { padding: 10px 8px; }
                #dashboard-page .dash-kpi-card { padding: 9px 7px 9px 9px; gap: 6px; }
                #dashboard-page .dash-kpi-icon { flex-basis: 28px; width: 28px; height: 28px; font-size: 12px; }
                #dashboard-page .dash-kpi-value { font-size: 14px; }
            }
        `;
        document.head.appendChild(st);
    }

    // Collect Group values from the uploaded Excel rows (prepMap), not only saved studies.
    // This makes Group available as soon as an Excel file is uploaded and keeps it linked to Style.
    function getExcelGroupRows() {
        const rows = [];
        safeLoad('uploadedExcelFiles').forEach(file => {
            Object.values(file.prepMap || {}).forEach(arr => {
                if (Array.isArray(arr)) arr.forEach(item => rows.push(item));
            });
        });
        return rows;
    }

    const FILTER_ORDER = ['factory', 'section', 'style', 'rampup', 'group'];

    function recordMatchesFilters(record, filters, keys) {
        return keys.every(key => {
            const wanted = filters && filters[key];
            if (!wanted) return true;
            const getter = FILTER_DEFS[key] && FILTER_DEFS[key].get;
            return String(getter ? getter(record) : '') === String(wanted);
        });
    }

    function getGroupFilterValues(filtersArg) {
        const f = filtersArg || getCurrentFilterValues();
        const priorKeys = FILTER_ORDER.slice(0, FILTER_ORDER.indexOf('group'));
        const rows = getExcelGroupRows();
        const fromExcel = rows.filter(item => {
            const normalized = {
                factory: item.title || item.factory || '',
                section: item.section || '',
                style: item.styleName || '',
                rampup: item.rampupDay || ''
            };
            return priorKeys.every(key => !f[key] || String(normalized[key] || '') === String(f[key]));
        }).map(item => studyGroup(item) || String(item.prepAssembly || item.prepName || '').trim()).filter(Boolean);
        const saved = safeLoad('savedTimeStudies');
        const fromSaved = saved.filter(item => recordMatchesFilters(item, f, priorKeys))
            .map(item => studyGroup(item) || String(item.prepAssembly || item.prepName || '').trim()).filter(Boolean);
        return [...new Set([...fromExcel, ...fromSaved].map(v => String(v).trim()).filter(Boolean))]
            .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
    }

    function buildFilterOptions(key, studies, selected, selectedFilters = {}) {
        const def = FILTER_DEFS[key];
        const priorKeys = FILTER_ORDER.slice(0, FILTER_ORDER.indexOf(key));
        const eligible = studies.filter(item => recordMatchesFilters(item, selectedFilters, priorKeys));
        const values = key === 'group'
            ? getGroupFilterValues(selectedFilters)
            : [...new Set(eligible.map(def.get).filter(Boolean).map(String))]
                .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        return `<option value="">All</option>` +
            values.map(v => `<option value="${escapeHtml(v)}"${v === selected ? ' selected' : ''}>${escapeHtml(v)}</option>`).join('');
    }

    function refreshCascadingFilterOptions() {
        const filters = getCurrentFilterValues();
        const studies = safeLoad('savedTimeStudies');
        FILTER_ORDER.forEach((key, index) => {
            if (index === 0) return; // Factory is the root list; it is not constrained by another filter.
            const select = document.getElementById(FILTER_DEFS[key].id);
            if (!select) return;
            const previous = select.value;
            const priorKeys = FILTER_ORDER.slice(0, index);
            const eligible = studies.filter(item => recordMatchesFilters(item, filters, priorKeys));
            const values = key === 'group'
                ? getGroupFilterValues(filters)
                : [...new Set(eligible.map(FILTER_DEFS[key].get).filter(Boolean).map(String))]
                    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
            select.innerHTML = `<option value="">All</option>` + values.map(v =>
                `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`
            ).join('');
            select.value = values.includes(previous) ? previous : '';
            filters[key] = select.value;
        });
    }

    function refreshGroupFilterOptions() {
        refreshCascadingFilterOptions();
    }

    function buildFilterBox(key, studies, selected) {
        const def = FILTER_DEFS[key];
        return `
            <div style="background: #fff; padding: 8px; border-radius: 6px; border: 1px solid #7dd3fc;">
                <div style="display: flex; align-items: center; justify-content: space-between; gap: 6px;">
                    <label for="${def.id}" style="font-size: 11px; color: #334155; font-weight: bold; display: block;">${def.label}</label>
                    <button type="button" class="filter-refresh-btn" data-filter="${key}" title="Refresh ${def.label}" onclick="refreshDashboardFilter('${key}', this)">
                        <i class="fa-solid fa-rotate"></i>
                    </button>
                </div>
                <select id="${def.id}" class="form-control" style="width: 100%; border: none; background: transparent; font-size: 13px; outline: none;" onchange="filterDashboardTable()">
                    ${buildFilterOptions(key, studies, selected, window.__dashboardFilterSeed || {})}
                </select>
            </div>`;
    }

    function getCurrentFilterValues() {
        const vals = {};
        Object.keys(FILTER_DEFS).forEach(k => {
            const el = document.getElementById(FILTER_DEFS[k].id);
            vals[k] = el ? el.value : '';
        });
        return vals;
    }

    function getFilteredStudies() {
        const savedStudies = safeLoad('savedTimeStudies');
        const f = getCurrentFilterValues();
        return savedStudies.filter(s => Object.keys(FILTER_DEFS).every(k => !f[k] || String(FILTER_DEFS[k].get(s)) === f[k]));
    }

    // Output Average popup: groups the currently filtered saved studies by worker.
    function closeOutputAverageModal() {
        const modal = document.getElementById('output-average-modal');
        if (modal) modal.remove();
        document.removeEventListener('keydown', outputAverageEscapeHandler);
    }

    function outputAverageEscapeHandler(e) {
        if (e.key === 'Escape') closeOutputAverageModal();
    }

    window.showOutputAverage = function () {
        closeOutputAverageModal();
        const studies = getFilteredStudies();
        const groups = new Map();
        studies.forEach((s) => {
            const worker = String(s.sewerName || '').trim() || 'No Sewer';
            if (!groups.has(worker)) groups.set(worker, []);
            groups.get(worker).push(s);
        });

        const workerNames = [...groups.keys()].sort((a, b) => a.localeCompare(b, undefined, {numeric:true, sensitivity:'base'}));
        const body = workerNames.map((worker) => {
            const items = groups.get(worker).slice().sort((a, b) =>
                String(a.processRegion || a.processName || '').localeCompare(String(b.processRegion || b.processName || ''), undefined, {numeric:true, sensitivity:'base'}));
            const totalCtSeconds = items.reduce((sum, s) => sum + studySeconds(s), 0);
            const combinedOutput = totalCtSeconds > 0 ? Math.round(3600 / totalCtSeconds) : 0;
            const rows = items.map((s) => {
                const sec = studySeconds(s);
                // CT Average is shown as seconds (e.g. 110), not MM:SS.
                const ct = sec > 0 ? String(Math.round(sec)) : '—';
                // Keep the individual process output visible for each process.
                const savedOutput = Number(s.output);
                const output = Number.isFinite(savedOutput) && savedOutput > 0
                    ? Math.round(savedOutput)
                    : (sec > 0 ? Math.round(3600 / sec) : 0);
                const process = s.processRegion || s.processName || 'N/A';
                const meta = [s.title || s.factory, s.section, s.styleName, s.rampupDay].filter(v => v && v !== 'N/A').join(' · ');
                return `<tr class="oa-process-row" data-search="${escapeHtml([worker, process, s.processName, s.title, s.factory, s.section, s.styleName, s.rampupDay].join(' ').toLowerCase())}">
                    <td><strong>${escapeHtml(process)}</strong>${meta ? `<div class="oa-meta">${escapeHtml(meta)}</div>` : ''}</td>
                    <td class="oa-number">${output > 0 ? output.toLocaleString() : '—'} <small>pcs/hr</small></td>
                    <td class="oa-number">${escapeHtml(ct)} <small>sec</small></td>
                </tr>`;
            }).join('') + `<tr class="oa-total-row" style="font-weight:700;background:rgba(59,130,246,.10);border-top:2px solid var(--border-color,#cbd5e1)">
                <td><strong>សរុប ${items.length} Process</strong><div class="oa-meta">Output គិតតាម CT សរុបរបស់មនុស្សម្នាក់</div></td>
                <td class="oa-number">${combinedOutput > 0 ? combinedOutput.toLocaleString() : '—'} <small>pcs/hr</small></td>
                <td class="oa-number">${totalCtSeconds > 0 ? String(Math.round(totalCtSeconds)) : '—'} <small>sec</small></td>
            </tr>`;
            return `<section class="oa-worker-card" data-worker="${escapeHtml(worker.toLowerCase())}" data-worker-name="${escapeHtml(worker)}" data-process-count="${items.length}" data-output="${combinedOutput}">
                <button type="button" class="oa-worker-heading" aria-label="ជ្រើសរើស ${escapeHtml(worker)}" title="ចុចដើម្បីបង្ហាញសេចក្ដីសង្ខេបរបស់បុគ្គលិកនេះ" style="display:flex;width:100%;text-align:left;align-items:center;gap:12px;border:0;background:transparent;padding:12px 14px;cursor:pointer;border-radius:10px;">
                    <span class="oa-avatar"><i class="fa-solid fa-user"></i></span><div><h3>${escapeHtml(worker)}</h3><p>${items.length} Process${items.length === 1 ? '' : 'es'}</p></div><span style="margin-left:auto;color:var(--ts-primary,#2563eb);font-size:12px;font-weight:700;">ចុចដើម្បីជ្រើសរើស</span>
                </button>
                <div class="table-responsive"><table class="oa-table"><thead><tr><th>Process</th><th>Output Average</th><th>CT Average</th></tr></thead><tbody>${rows}</tbody></table></div>
            </section>`;
        }).join('');

        const modal = document.createElement('div');
        modal.id = 'output-average-modal';
        modal.className = 'oa-overlay';
        modal.innerHTML = `<div class="oa-dialog" role="dialog" aria-modal="true" aria-labelledby="oa-title">
            <div class="oa-topbar"><div><span class="oa-eyebrow">TIME STUDY PRO</span><h2 id="oa-title"><i class="fa-solid fa-chart-simple"></i> Output Average & CT Average</h2><p>បង្ហាញ Process ទាំងអស់តាមមនុស្សម្នាក់</p></div><button type="button" class="oa-close" aria-label="Close" onclick="closeOutputAverageModal()"><i class="fa-solid fa-xmark"></i></button></div>
            <div class="oa-summary"><div><span>Workers</span><strong id="oa-summary-workers">${workerNames.length}</strong></div><div><span>Processes</span><strong id="oa-summary-processes">${studies.length}</strong></div><div><span>Output unit</span><strong id="oa-summary-output">pcs/hour</strong></div></div>
            <div class="oa-toolbar"><div class="oa-summary-refresh"><button type="button" id="oa-summary-refresh-btn" title="ដោះការជ្រើសរើសបុគ្គលិក និងរាប់ Workers និង Processes ទាំងអស់ឡើងវិញ"><i class="fa-solid fa-arrows-rotate"></i> Refresh</button></div>
            <label class="oa-search-wrap"><i class="fa-solid fa-magnifying-glass"></i><input id="oa-search" type="search" placeholder="ស្វែងរកឈ្មោះមនុស្ស ឬ Process..." autocomplete="off"></label></div>
            <div class="oa-content" id="oa-content">${body || '<div class="oa-empty"><i class="fa-solid fa-circle-info"></i><p>មិនមានទិន្នន័យ Time Study សម្រាប់ Filter បច្ចុប្បន្នទេ។</p></div>'}</div>
            <div class="oa-footer"><span>Output Average គិតជា pcs/hour · CT Average គិតជា seconds</span><button type="button" class="dash-export-btn" onclick="closeOutputAverageModal()">បិទ</button></div>
        </div>`;
        document.body.appendChild(modal);
        modal.addEventListener('click', (e) => { if (e.target === modal) closeOutputAverageModal(); });
        // Clicking a worker updates the three summary values at the top of the modal.
        const summaryWorkers = document.getElementById('oa-summary-workers');
        const summaryProcesses = document.getElementById('oa-summary-processes');
        const summaryOutput = document.getElementById('oa-summary-output');
        const summaryRefresh = document.getElementById('oa-summary-refresh-btn');
        if (summaryRefresh) summaryRefresh.addEventListener('click', () => {
            // Clear any worker selection and restore totals for all workers/processes
            // currently included by the dashboard's Factory/Section/Style/Day filters.
            modal.querySelectorAll('.oa-worker-card').forEach(card => {
                card.style.outline = 'none';
                card.style.borderRadius = '12px';
            });
            if (summaryWorkers) summaryWorkers.textContent = String(workerNames.length);
            if (summaryProcesses) summaryProcesses.textContent = String(studies.length);
            if (summaryOutput) summaryOutput.textContent = 'pcs/hour';
            summaryRefresh.classList.add('spinning');
            window.setTimeout(() => summaryRefresh.classList.remove('spinning'), 600);
        });
        modal.querySelectorAll('.oa-worker-card').forEach(card => {
            const heading = card.querySelector('.oa-worker-heading');
            if (!heading) return;
            heading.addEventListener('click', () => {
                modal.querySelectorAll('.oa-worker-card').forEach(other => {
                    other.style.outline = 'none';
                    other.style.borderRadius = '12px';
                });
                card.style.outline = '2px solid var(--ts-primary,#2563eb)';
                card.style.borderRadius = '12px';
                if (summaryWorkers) summaryWorkers.textContent = '1';
                if (summaryProcesses) summaryProcesses.textContent = card.getAttribute('data-process-count') || '0';
                const outputValue = Number(card.getAttribute('data-output')) || 0;
                if (summaryOutput) summaryOutput.textContent = outputValue > 0 ? outputValue.toLocaleString() + ' pcs/hour' : '— pcs/hour';
            });
        });
        const search = document.getElementById('oa-search');
        if (search) search.addEventListener('input', () => {
            const q = search.value.trim().toLowerCase();
            modal.querySelectorAll('.oa-worker-card').forEach(card => {
                let visible = false;
                card.querySelectorAll('.oa-process-row').forEach(row => {
                    const match = !q || (row.getAttribute('data-search') || '').includes(q);
                    row.style.display = match ? '' : 'none';
                    if (match) visible = true;
                });
                card.style.display = visible ? '' : 'none';
            });
        });
        document.addEventListener('keydown', outputAverageEscapeHandler);
    };

    window.closeOutputAverageModal = closeOutputAverageModal;

    function getTaktSeconds() {
        return parseFloat(localStorage.getItem('dashTaktSec')) || 0;
    }

    function studySeconds(item) {
        // Prefer stored average seconds, but recover older records from CT text or lap data.
        let sec = parseFloat(item && item.avgSeconds);
        if (!Number.isFinite(sec) || sec <= 0) {
            sec = item && item.averageCT ? parseTimeToSeconds(item.averageCT) : 0;
        }
        if ((!Number.isFinite(sec) || sec <= 0) && item && Array.isArray(item.laps) && item.laps.length) {
            const lapSeconds = item.laps.map(lap => {
                if (lap && Number.isFinite(Number(lap.durationMs)) && Number(lap.durationMs) > 0) return Number(lap.durationMs) / 1000;
                if (lap && Number.isFinite(Number(lap.duration)) && Number(lap.duration) > 0) return Number(lap.duration);
                if (lap && lap.time != null) return parseTimeToSeconds(lap.time);
                return 0;
            }).filter(value => Number.isFinite(value) && value > 0);
            if (lapSeconds.length) sec = lapSeconds.reduce((sum, value) => sum + value, 0) / lapSeconds.length;
        }
        return Number.isFinite(sec) && sec > 0 ? sec : 0;
    }

    // ✨ NEW 2: Line Balancing calculation (per Sewer / Operator)
    function computeLineBalance(studies) {
        const totals = {};
        const seqMap = {};
        studies.forEach(s => {
            const name = validSewerName(s.sewerName);
            totals[name] = (totals[name] || 0) + studySeconds(s);
            seqMap[name] = parseInt(s.seq) || parseInt(s.operatorSeq) || 99;
        });
        const names = Object.keys(totals).sort((a, b) => (seqMap[a] || 99) - (seqMap[b] || 99));
        const values = names.map(n => totals[n]);
        const n = names.length;
        const sum = values.reduce((a, b) => a + b, 0);
        const max = n ? Math.max(...values) : 0;
        const bottleneck = n ? names[values.indexOf(max)] : '';
        const takt = getTaktSeconds();
        const balanceRate = (n > 0 && max > 0) ? (sum / (n * max)) * 100 : 0;
        const lineEfficiency = (n > 0 && takt > 0) ? (sum / (n * takt)) * 100 : 0;
        const overTakt = takt > 0 ? names.filter((nm, i) => values[i] > takt) : [];
        return { names, totals, sum, max, bottleneck, n, takt, balanceRate, lineEfficiency, overTakt, hourlyOutput: max > 0 ? 3600 / max : 0 };
    }

    // ✨ NEW 3: KPI Cards
    function kpiCard(icon, color, title, value, sub) {
        // Styling-only classes: KPI calculations and existing dashboard behavior remain unchanged.
        const kpiVariant = icon.includes('fa-percent') ? 'kpi-pfm'
            : icon.includes('fa-turtle') ? 'kpi-slowest' : 'kpi-standard';
        return `
            <div class="dash-kpi-card ${kpiVariant}" style="--kpi-accent:${color};">
                <div class="dash-kpi-icon"><i class="${icon}"></i></div>
                <div class="dash-kpi-copy">
                    <div class="dash-kpi-title">${title}</div>
                    <div class="dash-kpi-value">${value}</div>
                    <div class="dash-kpi-sub">${sub}</div>
                </div>
            </div>`;
    }

    function renderDashboardKPIs(studies) {
        const container = document.getElementById('dash-kpi-container');
        if (!container) return;

        const f = getCurrentFilterValues();
        const uploadedFiles = safeLoad('uploadedExcelFiles');
        let totalRows = 0;
        uploadedFiles.forEach(file => {
            Object.values(file.prepMap || {}).forEach(arr => arr.forEach(it => {
                if (f.factory && String(it.title || it.factory || '') !== f.factory) return;
                if (f.section && String(it.section || '') !== f.section) return;
                if (f.style && String(it.styleName || '') !== f.style) return;
                if (f.rampup && String(it.rampupDay || '') !== f.rampup) return;
                totalRows++;
            }));
        });

        const pfms = studies.map(s => parseFloat(s.pfm) || 0).filter(v => v > 0);
        const avgPfm = pfms.length ? pfms.reduce((a, b) => a + b, 0) / pfms.length : 0;
        const pc = pfmColors(avgPfm);

        let slowest = null;
        studies.forEach(s => { if (!slowest || studySeconds(s) > studySeconds(slowest)) slowest = s; });

        const totalSmv = studies.reduce((a, s) => a + (parseFloat(s.smv) || 0), 0);
        const totalCt = studies.reduce((a, s) => a + studySeconds(s), 0);

        const progressPct = totalRows > 0 ? Math.min(100, Math.round((studies.length / totalRows) * 100)) : 0;

        const lb = computeLineBalance(studies);
        // Count unique operators represented by the currently filtered chart/dashboard data.
        const peopleInGraph = new Set((studies || []).map(s => validSewerName(s.sewerName))).size;
        let balanceSub = 'Bottleneck: ' + (lb.bottleneck ? escapeHtml(lb.bottleneck) + ' (' + formatToMinSec(lb.max) + ')' : '-');
        let balanceColor = '#10b981';
        if (lb.balanceRate > 0 && lb.balanceRate < 70) balanceColor = '#dc2626';
        else if (lb.balanceRate < 85) balanceColor = '#d97706';

        let effCard;
        if (lb.takt > 0) {
            const effColor = lb.lineEfficiency >= 85 ? '#10b981' : (lb.lineEfficiency >= 70 ? '#d97706' : '#dc2626');
            const overTxt = lb.overTakt.length ? lb.overTakt.length + ' Operator លើស Takt' : 'គ្រប់ Operator ក្នុង Takt';
            effCard = kpiCard('fa-solid fa-gauge-high', effColor, 'Line Efficiency (vs Takt ' + lb.takt + 's)', lb.lineEfficiency.toFixed(1) + '%', overTxt);
        } else {
            effCard = kpiCard('fa-solid fa-gauge-high', '#94a3b8', 'Line Efficiency', '—', 'បញ្ចូល Takt Time ដើម្បីគណនា');
        }

        container.innerHTML = `
            ${kpiCard('fa-solid fa-percent', pc.color, 'PFM មធ្យម', avgPfm > 0 ? Math.round(avgPfm) + '%' : '—', pfms.length + ' Process')}
            ${kpiCard('fa-solid fa-turtle', '#dc2626', 'Process យឺតបំផុត', slowest ? formatToMinSec(studySeconds(slowest)) : '—', slowest ? escapeHtml(slowest.processRegion || slowest.processName || '') : 'គ្មានទិន្នន័យ')}
            ${kpiCard('fa-solid fa-scale-balanced', '#2563eb', 'Total SMV vs Total CT', totalSmv.toFixed(1) + 's / ' + formatToMinSec(totalCt), 'SMV ធៀបនឹង CT ជាក់ស្ដែង')}
            ${kpiCard('fa-solid fa-list-check', '#8b5cf6', 'វឌ្ឍនភាពចាប់នាទី', studies.length + ' / ' + totalRows, progressPct + '% រួចរាល់')}
            ${kpiCard('fa-solid fa-diagram-project', balanceColor, 'Balance Rate', lb.balanceRate > 0 ? lb.balanceRate.toFixed(1) + '%' : '—', balanceSub)}
            ${kpiCard('fa-solid fa-users', '#0891b2', 'People in Graph', String(peopleInGraph), 'ចំនួនមនុស្ស/Operator មិនស្ទួនក្នុងទិន្នន័យដែលបាន Filter')}
            ${effCard}
            ${kpiCard('fa-solid fa-boxes-stacked', '#0ea5e9', 'Output អតិបរមា/ម៉ោង', lb.hourlyOutput > 0 ? Math.floor(lb.hourlyOutput) + ' Pcs' : '—', 'កំណត់ដោយ Bottleneck')}
        `;
    }

    // resetKey: បើបញ្ជាក់ → កំណត់ Filter នោះទៅ "All"
    window.renderDashboardStats = function (resetKey) {
        const dashboardPage = document.getElementById('dashboard-page');
        if (!dashboardPage) return;
        injectDashboardStyles();

        const savedStudies = safeLoad('savedTimeStudies');
        const uploadedFiles = safeLoad('uploadedExcelFiles');

        const prev = getCurrentFilterValues();
        if (resetKey && Object.prototype.hasOwnProperty.call(prev, resetKey)) prev[resetKey] = '';
        window.__dashboardFilterSeed = prev;

        dashboardPage.innerHTML = `
            <div class="observe-container">
                <div class="observe-header" style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;">
                    <h2><i class="fa-solid fa-chart-line"></i> DASHBOARD ANALYSIS & TIME STUDIES</h2>
                    <div style="display:flex;gap:8px;flex-wrap:wrap;">
                        <button type="button" class="dash-export-btn oa-open-btn" onclick="showOutputAverage()"><i class="fa-solid fa-chart-simple"></i> Output Average</button>
                        <button type="button" class="dash-export-btn" style="background:#d1fae5;color:#047857;" onclick="exportDashboardExcel()"><i class="fa-solid fa-file-excel"></i> Export Excel</button>
                        <button type="button" class="dash-export-btn" style="background:#fee2e2;color:#dc2626;" onclick="exportDashboardPDF()"><i class="fa-solid fa-file-pdf"></i> Export PDF</button>
                    </div>
                </div>

                <div class="cards-grid" style="margin-top: 20px;">
                    <div class="card">
                        <div class="card-icon card-1 icon"><i class="fa-solid fa-file-excel"></i></div>
                        <div class="card-info">
                            <h3>Total Excel Files</h3>
                            <p id="dash-total-file-count">${uploadedFiles.length}</p>
                        </div>
                    </div>
                    <div class="card">
                        <div class="card-icon card-2 icon"><i class="fa-solid fa-stopwatch"></i></div>
                        <div class="card-info">
                            <h3>Total Studies Saved</h3>
                            <p>${savedStudies.length}</p>
                        </div>
                    </div>
                </div>

                <!-- Filter Selection Bar (Refresh ក្នុងប្រអប់នីមួយៗ) -->
                <div class="dash-filter-grid" style="background: #e0f2fe; padding: 15px; border-radius: 10px; display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 10px; margin-top: 25px; border: 1px solid #bae6fd;">
                    ${buildFilterBox('factory', savedStudies, prev.factory)}
                    ${buildFilterBox('section', savedStudies, prev.section)}
                    ${buildFilterBox('style', savedStudies, prev.style)}
                    ${buildFilterBox('rampup', savedStudies, prev.rampup)}
                    ${buildFilterBox('group', savedStudies, prev.group)}
                </div>

                <!-- KPI Cards -->
                <div id="dash-kpi-container" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-top:20px;"></div>

                <!-- Yamazumi / Line Balancing Chart -->
                <div class="uploaded-files-section" style="margin-top: 25px; background: #fff; padding: 20px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:15px;">
                        <h3 style="margin:0;"><i class="fa-solid fa-chart-bar"></i> Yamazumi Chart: Cycle Time (CT Average) តាម Seq និង Process</h3>
                        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;font-size:12px;color:#475569;">
                            <label for="takt-input" style="font-weight:bold;">Takt Time (s)</label>
                            <input type="number" id="takt-input" class="takt-input" min="0" step="1" placeholder="ឧ. 60">
                            <label for="target-input" style="font-weight:bold;">ឬ Target/ម៉ោង</label>
                            <input type="number" id="target-input" class="takt-input" min="0" step="1" placeholder="ឧ. 60">
                        </div>
                    </div>
                    <div id="yamazumi-kpi-summary" class="yamazumi-kpi-grid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:10px;margin:0 0 16px;"></div>
                    <div style="display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap;margin-bottom:10px;">
                        <button type="button" class="dash-export-btn" onclick="exportYamazumiPNG()"><i class="fa-solid fa-image"></i> Export PNG</button>
                    </div>
                    <div style="position: relative; height: 340px; width: 100%;">
                        <canvas id="ctProcessChart"></canvas>
                    </div>
                </div>

                <!-- Main Dashboard Table -->
                <div class="uploaded-files-section" style="margin-top: 25px;">
                    <h3><i class="fa-solid fa-list-check"></i> បញ្ជីទិន្នន័យ Time Study និងការវិភាគ (CT & PFM)</h3>
                    <div class="table-responsive" style="margin-top: 15px;">
                        <table id="dashboard-main-table" class="files-table" style="width: 100%; text-align: left;">
                            <thead>
                                <tr>
                                    <th>No</th>
                                    <th>Seq</th>
                                    <th>Start Time</th>
                                    <th>Factory / Title</th>
                                    <th>Section</th>
                                    <th>Style</th>
                                    <th>Day</th>
                                    <th>Sewer</th>
                                    <th>MC</th>
                                    <th>Process (Khmer)</th>
                                    <th>PrimaryKey</th>
                                    <th>SMV</th>
                                    <th>CT (Avg)</th>
                                    <th>Output</th>
                                    <th>PFM (%)</th>
                                    <th>Action</th>
                                </tr>
                            </thead>
                            <tbody id="dashboard-table-body">
                                ${savedStudies.length === 0 ? `
                                    <tr><td colspan="16" style="text-align: center; color: #9ca3af; padding: 20px;">មិនទាន់មានទិន្នន័យ Time Study ត្រូវបាន Save នៅឡើយទេ</td></tr>
                                ` : savedStudies.map((item, index) => {
                                    const pfmVal = parseFloat(item.pfm) || 0;
                                    const outputVal = item.output || 0;
                                    const opSeq = item.seq !== undefined ? item.seq : (item.operatorSeq !== undefined ? item.operatorSeq : 99);
                                    const itemNo = item.no || (index + 1);
                                    const pc = pfmColors(pfmVal);
                                    const formattedCtAvg = String(Math.round((item.averageCT !== undefined && item.averageCT !== null && String(item.averageCT).trim() !== '') ? parseTimeToSeconds(item.averageCT) : (parseFloat(item.avgSeconds) || 0)));

                                    return `
                                        <tr data-factory="${escapeHtml(studyFactory(item))}"
                                            data-section="${escapeHtml(item.section || '')}"
                                            data-style="${escapeHtml(item.styleName || '')}"
                                            data-rampup="${escapeHtml(item.rampupDay || '')}"
                                            data-group="${escapeHtml(studyGroup(item))}">
                                            <td><strong>${escapeHtml(itemNo)}</strong></td>
                                            <td><span style="background: #f3f4f6; padding: 2px 6px; border-radius: 4px; font-weight: bold;">${escapeHtml(opSeq)}</span></td>
                                            <td style="font-size: 12px; color: #64748b;">${escapeHtml(item.dateTimeStart || item.date)}</td>
                                            <td>${escapeHtml(studyFactory(item) || 'N/A')}</td>
                                            <td>${escapeHtml(item.section || 'N/A')}</td>
                                            <td>${escapeHtml(item.styleName || 'N/A')}</td>
                                            <td>${escapeHtml(item.rampupDay || 'N/A')}</td>
                                            <td><strong>${escapeHtml(item.sewerName || 'No Sewer')}</strong></td>
                                            <td><span style="background: #f1f5f9; padding: 2px 6px; border-radius: 4px;">${escapeHtml(item.mc || '-')}</span></td>
                                            <td><strong>${escapeHtml(item.processRegion || 'N/A')}</strong></td>
                                            <td>${escapeHtml(item.primaryKey || 'N/A')}</td>
                                            <td>${escapeHtml(item.smv || '0.00')}</td>
                                            <td><span style="font-weight: bold; color: #2563eb; background: #eff6ff; padding: 3px 6px; border-radius: 4px;">${formattedCtAvg}</span></td>
                                            <td><span style="font-weight: bold; color: #10b981; background: #ecfdf5; padding: 3px 6px; border-radius: 4px;">${escapeHtml(outputVal)} Pcs</span></td>
                                            <td>
                                                <span style="font-weight: bold; background: ${pc.bg}; color: ${pc.color}; padding: 4px 8px; border-radius: 6px;">
                                                    ${Math.round(pfmVal)}%
                                                </span>
                                            </td>
                                            <td>
                                                <button type="button" onclick="editStudyRecord('${escapeHtml(item.id)}')" style="background: none; border: none; color: #2563eb; cursor: pointer; margin-right: 5px;" title="Edit Study"><i class="fa-solid fa-pen-to-square"></i></button>
                                                <button type="button" onclick="deleteDashboardStudy('${escapeHtml(item.id)}')" style="background: none; border: none; color: #dc2626; cursor: pointer;" title="Delete"><i class="fa-solid fa-trash"></i></button>
                                            </td>
                                        </tr>
                                    `;
                                }).join('')}
                                <tr id="dashboard-no-match-row" style="display:none;"><td colspan="16" style="text-align: center; color: #9ca3af; padding: 20px;">រកមិនឃើញទិន្នន័យដែលត្រូវនឹង Filter ទេ</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;

        // Takt inputs
        const taktEl = document.getElementById('takt-input');
        const tgtEl = document.getElementById('target-input');
        const takt = getTaktSeconds();
        if (taktEl) taktEl.value = takt || '';
        if (tgtEl) tgtEl.value = takt ? Math.round(3600 / takt) : '';
        if (taktEl) taktEl.addEventListener('change', () => {
            const v = Math.max(0, parseFloat(taktEl.value) || 0);
            localStorage.setItem('dashTaktSec', String(v));
            if (tgtEl) tgtEl.value = v ? Math.round(3600 / v) : '';
            window.filterDashboardTable();
        });
        if (tgtEl) tgtEl.addEventListener('change', () => {
            const t = Math.max(0, parseFloat(tgtEl.value) || 0);
            const v = t ? +(3600 / t).toFixed(1) : 0;
            localStorage.setItem('dashTaktSec', String(v));
            if (taktEl) taktEl.value = v || '';
            window.filterDashboardTable();
        });

        delete window.__dashboardFilterSeed;
        installTableResizerStyles();
        makeTableResizable(document.getElementById('dashboard-main-table'));
        window.filterDashboardTable();
    };

    // 🔄 Refresh Filter ប្រអប់នីមួយៗ
    window.refreshDashboardFilter = function (key, btn) {
        if (btn) btn.classList.add('spinning');
        setTimeout(() => {
            window.renderDashboardStats(key);
            const label = FILTER_DEFS[key] ? FILTER_DEFS[key].label : key;
            showToast('Refresh ជោគជ័យ', 'ទិន្នន័យ [' + label + '] ត្រូវបានធ្វើបច្ចុប្បន្នភាពសាថ្មី', 'fa-solid fa-rotate', '#2563eb');
        }, 450);
    };

    function renderYamazumiKPIs(studies) {
        const el = document.getElementById('yamazumi-kpi-summary');
        if (!el) return;
        const vals = (studies || []).map(s => ({ s, ct: studySeconds(s) })).filter(x => Number.isFinite(x.ct) && x.ct > 0);
        const avg = vals.length ? vals.reduce((a, x) => a + x.ct, 0) / vals.length : 0;
        const high = vals.length ? vals.reduce((a, x) => x.ct > a.ct ? x : a, vals[0]) : null;
        const low = vals.length ? vals.reduce((a, x) => x.ct < a.ct ? x : a, vals[0]) : null;
        const proc = s => String(s.processRegion || s.processName || 'General Process');
        const card = (label, value, note, color) => `<div style="background:#f8fafc;border:1px solid #e2e8f0;border-left:4px solid ${color};border-radius:10px;padding:12px 14px;min-width:0;"><div style="font-size:11px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.04em;">${label}</div><div style="font-size:22px;font-weight:800;color:#0f172a;margin-top:5px;overflow-wrap:anywhere;">${value}</div><div style="font-size:11px;color:#64748b;margin-top:3px;overflow-wrap:anywhere;">${note}</div></div>`;
        el.innerHTML = [
            card('Average CT', avg ? formatToMinSec(avg) : '—', vals.length + ' records', '#2563eb'),
            card('Highest CT', high ? formatToMinSec(high.ct) : '—', high ? escapeHtml(proc(high.s)) : 'No data', '#dc2626'),
            card('Lowest CT', low ? formatToMinSec(low.ct) : '—', low ? escapeHtml(proc(low.s)) : 'No data', '#16a34a'),
            card('Bottleneck Process', high ? escapeHtml(proc(high.s)) : '—', high ? 'Seq ' + escapeHtml(String(high.s.seq || high.s.operatorSeq || '—')) : 'No data', '#ea580c')
        ].join('');
    }

    window.exportYamazumiPNG = function () {
        if (!window.myCTChart || typeof window.myCTChart.toBase64Image !== 'function') {
            showCustomAlert('ជូនដំណឹង', 'មិនមាន Chart សម្រាប់ Export ទេ។', 'fa-solid fa-circle-info', '#2563eb');
            return;
        }
        const a = document.createElement('a');
        a.href = window.myCTChart.toBase64Image('image/png', 1);
        a.download = 'TimeStudyPro_Yamazumi_Chart.png';
        document.body.appendChild(a); a.click(); a.remove();
    };

    function uniqueProcessTotal(processMap) {
        return Object.values(processMap || {}).reduce((sum, value) => sum + (Number(value) || 0), 0);
    }

    function renderCTChart(studies) {
        renderYamazumiKPIs(studies);
        const ctx = document.getElementById('ctProcessChart');
        if (!ctx) return;
        if (typeof Chart === 'undefined') {
            console.warn('Chart.js មិនទាន់ Load');
            return;
        }

        const sewerDataMap = {};
        const sewerSeqMap = {};
        const allProcessesSet = new Set();

        studies.forEach(item => {
            const sewerName = validSewerName(item.sewerName);
            const opSeq = parseInt(item.seq) || parseInt(item.operatorSeq) || 99;
            sewerSeqMap[sewerName] = opSeq;

            const processName = item.processRegion || item.processName || 'General Process';
            allProcessesSet.add(processName);

            if (!sewerDataMap[sewerName]) sewerDataMap[sewerName] = {};
            sewerDataMap[sewerName][processName] = (sewerDataMap[sewerName][processName] || 0) + studySeconds(item);
        });

        const labels = Object.keys(sewerDataMap).sort((a, b) => (sewerSeqMap[a] || 99) - (sewerSeqMap[b] || 99));
        const uniqueProcesses = Array.from(allProcessesSet);

        const lb = computeLineBalance(studies);
        const takt = lb.takt;
        const bottleneckIndex = lb.bottleneck ? labels.indexOf(lb.bottleneck) : -1;

        // Calculate dynamic Yamazumi limits from the currently filtered chart totals.
        // BPT is the mean CT per displayed Seq; UCL/LCL use mean +/- 3 population SD.
        const seqTotals = labels.map(sewerName => uniqueProcessTotal(sewerDataMap[sewerName]));
        const validSeqTotals = seqTotals.filter(v => Number.isFinite(v) && v >= 0);
        const bpt = validSeqTotals.length
            ? validSeqTotals.reduce((sum, value) => sum + value, 0) / validSeqTotals.length
            : 0;
        const variance = validSeqTotals.length
            ? validSeqTotals.reduce((sum, value) => sum + Math.pow(value - bpt, 2), 0) / validSeqTotals.length
            : 0;
        const stdDev = Math.sqrt(variance);
        const ucl = bpt + (3 * stdDev);
        const lcl = Math.max(0, bpt - (3 * stdDev));

        // Auto-scale the Yamazumi Y-axis to a clean, readable ceiling based on the
        // largest visible Seq total / UCL / Takt value. Include stacked totals and
        // reference lines so no meaningful value is clipped.
        // Common scales: 40 -> 50/10, 80 -> 100/20, 120 -> 150/30,
        // 180 -> 200/50, 260 -> 300/50, 450 -> 500/100, 850 -> 1000/200.
        const yScalePeak = Math.max(1, ...seqTotals, ucl || 0, takt || 0);
        const standardYScales = [
            { ceiling: 50, step: 10 },
            { ceiling: 100, step: 20 },
            { ceiling: 150, step: 30 },
            { ceiling: 200, step: 50 },
            { ceiling: 300, step: 50 },
            { ceiling: 500, step: 100 },
            { ceiling: 1000, step: 200 }
        ];
        const matchedYScale = standardYScales.find(scale => yScalePeak <= scale.ceiling);
        let yAxisStep;
        let yAxisMax;
        if (matchedYScale) {
            yAxisStep = matchedYScale.step;
            yAxisMax = matchedYScale.ceiling;
        } else {
            // Beyond the common ranges, continue with a sensible 1/2/5/10 tick step.
            const rawYStep = yScalePeak / 5;
            const yMagnitude = Math.pow(10, Math.floor(Math.log10(Math.max(rawYStep, 1))));
            const normalizedStep = rawYStep / yMagnitude;
            yAxisStep = (normalizedStep <= 1 ? 1 : normalizedStep <= 2 ? 2 : normalizedStep <= 5 ? 5 : 10) * yMagnitude;
            yAxisMax = Math.ceil(yScalePeak / yAxisStep) * yAxisStep;
            if (yAxisMax <= yScalePeak) yAxisMax += yAxisStep;
        }

        const colorPalette = [
            'rgba(16, 185, 129, 0.85)',
            'rgba(37, 99, 235, 0.85)',
            'rgba(245, 158, 11, 0.85)',
            'rgba(239, 68, 68, 0.85)',
            'rgba(139, 92, 246, 0.85)',
            'rgba(236, 72, 153, 0.85)'
        ];

        const datasets = uniqueProcesses.map((processName, idx) => ({
            label: processName,
            data: labels.map(sewerName => {
                const val = sewerDataMap[sewerName][processName];
                return val !== undefined ? Math.round(val) : 0;
            }),
            backgroundColor: colorPalette[idx % colorPalette.length],
            borderRadius: 4
        }));

        if (window.myCTChart && typeof window.myCTChart.destroy === 'function') {
            window.myCTChart.destroy();
            window.myCTChart = null;
        }

        const limitLinesAndValuesPlugin = {
            id: 'limitLinesAndValuesPlugin',
            afterDraw(chart) {
                const { ctx, chartArea, scales } = chart;
                if (!chartArea || !scales || !scales.y) return;
                const { left, right } = chartArea;
                const y = scales.y;
                const limits = [];
                if (validSeqTotals.length) {
                    limits.push(
                        { value: ucl, label: 'UCL (' + Math.round(ucl) + 's)', color: '#dc2626', dash: [6, 6] },
                        { value: bpt, label: 'BPT (' + Math.round(bpt) + 's)', color: '#2563eb', dash: [6, 6] },
                        { value: lcl, label: 'LCL (' + Math.round(lcl) + 's)', color: '#d97706', dash: [6, 6] }
                    );
                }
                if (takt > 0) limits.push({ value: takt, label: 'TAKT (' + takt + 's)', color: '#7c3aed', dash: [] });

                ctx.save();
                limits.forEach(limit => {
                    const yPos = y.getPixelForValue(limit.value);
                    if (yPos >= chartArea.top && yPos <= chartArea.bottom) {
                        ctx.beginPath();
                        ctx.moveTo(left, yPos);
                        ctx.lineTo(right, yPos);
                        ctx.lineWidth = limit.dash.length ? 2 : 3;
                        ctx.strokeStyle = limit.color;
                        ctx.setLineDash(limit.dash);
                        ctx.stroke();

                        ctx.setLineDash([]);
                        ctx.fillStyle = limit.color;
                        ctx.font = 'bold 11px sans-serif';
                        ctx.textAlign = limit.dash.length ? 'right' : 'left';
                        ctx.textBaseline = 'bottom';
                        ctx.fillText(limit.label, limit.dash.length ? right - 5 : left + 5, yPos - 3);
                    }
                });
                ctx.restore();
            },
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                if (!chart.data.datasets.length) return;
                const meta0 = chart.getDatasetMeta(0);
                if (!meta0 || !meta0.data) return;

                meta0.data.forEach((bar, index) => {
                    let totalSum = 0;
                    chart.data.datasets.forEach(dataset => { totalSum += (dataset.data[index] || 0); });

                    if (totalSum > 0) {
                        let topY = null, topX = null;
                        for (let d = chart.data.datasets.length - 1; d >= 0; d--) {
                            if ((chart.data.datasets[d].data[index] || 0) > 0) {
                                const m = chart.getDatasetMeta(d);
                                if (m && m.data[index]) {
                                    const el = m.data[index];
                                    topX = el.tooltipPosition().x;
                                    topY = Math.min(el.y, el.base !== undefined ? el.base : el.y);
                                    break;
                                }
                            }
                        }
                        if (topY === null) return;

                        const isBottleneck = index === bottleneckIndex;
                        const overTakt = takt > 0 && totalSum > takt;
                        ctx.save();
                        ctx.font = 'bold 11px sans-serif';
                        ctx.fillStyle = (isBottleneck || overTakt) ? '#dc2626' : '#1e293b';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'bottom';
                        ctx.fillText((isBottleneck ? '🔥 ' : '') + String(Math.round(totalSum)), topX, topY - 6);
                        ctx.restore();
                    }
                });
            }
        };

        window.myCTChart = new Chart(ctx, {
            type: 'bar',
            data: { labels: labels, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { top: 30 } },
                scales: {
                    x: { stacked: true, title: { display: true, text: 'Sewer Name (Sorted by Seq)' } },
                    y: {
                        stacked: true,
                        beginAtZero: true,
                        title: { display: true, text: 'Total Time (Seconds)' },
                        // Explicit max + step keep the Y axis close to the visible data.
                        // Unlike suggestedMax, max prevents Chart.js from expanding the scale unnecessarily.
                        min: 0,
                        max: yAxisMax,
                        ticks: {
                            stepSize: yAxisStep,
                            maxTicksLimit: 7,
                            callback: (value) => String(Math.round(Number(value) || 0))
                        }
                    }
                },
                plugins: {
                    legend: { display: true, position: 'top' },
                    tooltip: {
                        callbacks: {
                            label: function (context) {
                                return context.dataset.label + ': ' + String(Math.round(Number(context.raw) || 0)) + ' s';
                            }
                        }
                    }
                }
            },
            plugins: [limitLinesAndValuesPlugin]
        });
    }

    // Filter ទាំងតារាង KPI និងក្រាហ្វ
    window.filterDashboardTable = function () {
        // Cascade filters in order: Factory -> Section -> Style -> Day -> Group.
        // Each downstream list only shows values related to the selections above it.
        refreshCascadingFilterOptions();
        const f = getCurrentFilterValues();

        let visible = 0;
        document.querySelectorAll('#dashboard-table-body tr[data-factory]').forEach(row => {
            let match = true;
            if (f.factory && row.getAttribute('data-factory') !== f.factory) match = false;
            if (f.section && row.getAttribute('data-section') !== f.section) match = false;
            if (f.style && row.getAttribute('data-style') !== f.style) match = false;
            if (f.rampup && row.getAttribute('data-rampup') !== f.rampup) match = false;
            if (f.group && row.getAttribute('data-group') !== f.group) match = false;
            row.style.display = match ? '' : 'none';
            if (match) visible++;
        });

        const noMatchRow = document.getElementById('dashboard-no-match-row');
        if (noMatchRow) {
            const hasRows = document.querySelectorAll('#dashboard-table-body tr[data-factory]').length > 0;
            noMatchRow.style.display = (hasRows && visible === 0) ? '' : 'none';
        }

        const filtered = getFilteredStudies();
        renderDashboardKPIs(filtered);
        renderCTChart(filtered);
    };

    // ✨ NEW 4a: Export Excel (CT ក្នុងគ្រប់ Lap)
    window.exportDashboardExcel = function () {
        if (typeof XLSX === 'undefined') {
            showCustomAlert('កំហុសឆ្គង', 'មិនអាច Export បានទេ (XLSX library មិនទាន់ Load)', 'fa-solid fa-triangle-exclamation', '#dc2626');
            return;
        }
        const studies = getFilteredStudies();
        if (studies.length === 0) {
            showCustomAlert('ជូនដំណឹង', 'គ្មានទិន្នន័យសម្រាប់ Export ទេ!', 'fa-solid fa-circle-info', '#3b82f6');
            return;
        }

        try {
            const maxLaps = Math.max(...studies.map(s => (s.laps || []).length), 0);
            const sorted = [...studies].sort((a, b) =>
                (parseInt(a.seq) || parseInt(a.operatorSeq) || 99) - (parseInt(b.seq) || parseInt(b.operatorSeq) || 99));

            const rows = sorted.map((s, i) => {
                const row = {
                    'No': s.no || (i + 1),
                    'Seq': s.seq !== undefined ? s.seq : (s.operatorSeq || ''),
                    'DateTime-Start': s.dateTimeStart || s.date || '',
                    'DateTime-End': s.dateTimeEnd || '',
                    'Title': studyFactory(s),
                    'Section': s.section || '',
                    'Style': s.styleName || '',
                    'RampupDay': s.rampupDay || '',
                    'Prep_Assembly': s.prepAssembly || s.prepName || '',
                    'SewerName': s.sewerName || 'No Sewer',
                    'MC': s.mc || '',
                    'PrimaryKey': s.primaryKey || '',
                    'Process(Region)': s.processRegion || '',
                    'Process': s.processName || '',
                    'SMV': parseFloat(s.smv) || 0
                };
                for (let k = 0; k < maxLaps; k++) {
                    const lap = (s.laps || [])[k];
                    row['CT' + (k + 1)] = lap ? formatToMinSec(lap.durationMs || 0, true) : '';
                }
                row['CT_Avg'] = formatToMinSec(s.averageCT || s.avgSeconds || 0);
                row['CT_Avg(s)'] = Math.round(studySeconds(s));
                row['Min(s)'] = s.minSeconds !== undefined ? +parseFloat(s.minSeconds).toFixed(1) : '';
                row['Max(s)'] = s.maxSeconds !== undefined ? +parseFloat(s.maxSeconds).toFixed(1) : '';
                row['StdDev(s)'] = s.stdSeconds !== undefined ? +parseFloat(s.stdSeconds).toFixed(2) : '';
                row['Outlier Laps'] = s.outlierCount || 0;
                row['Output(Pcs/h)'] = s.output || 0;
                row['PFM(%)'] = parseFloat(s.pfm) || 0;
                return row;
            });

            const lb = computeLineBalance(studies);
            const pfms = studies.map(s => parseFloat(s.pfm) || 0).filter(v => v > 0);
            const summary = [
                { 'KPI': 'Total Studies', 'Value': studies.length },
                { 'KPI': 'Average PFM (%)', 'Value': pfms.length ? +(pfms.reduce((a, b) => a + b, 0) / pfms.length).toFixed(1) : 0 },
                { 'KPI': 'Total SMV (s)', 'Value': +studies.reduce((a, s) => a + (parseFloat(s.smv) || 0), 0).toFixed(2) },
                { 'KPI': 'Total CT (MM:SS)', 'Value': formatToMinSec(studies.reduce((a, s) => a + studySeconds(s), 0)) },
                { 'KPI': 'Operators', 'Value': lb.n },
                { 'KPI': 'Bottleneck', 'Value': lb.bottleneck ? lb.bottleneck + ' (' + formatToMinSec(lb.max) + ')' : '-' },
                { 'KPI': 'Balance Rate (%)', 'Value': +lb.balanceRate.toFixed(1) },
                { 'KPI': 'Takt Time (s)', 'Value': lb.takt || '-' },
                { 'KPI': 'Line Efficiency vs Takt (%)', 'Value': lb.takt ? +lb.lineEfficiency.toFixed(1) : '-' },
                { 'KPI': 'Max Output/hour (Pcs)', 'Value': lb.hourlyOutput ? Math.floor(lb.hourlyOutput) : 0 }
            ];

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'TimeStudy');
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'Summary');

            const d = new Date();
            const stamp = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
            XLSX.writeFile(wb, 'TimeStudy_Report_' + stamp + '.xlsx');
            showToast('Export ជោគជ័យ', 'ឯកសារ Excel ត្រូវបានទាញយក', 'fa-solid fa-file-excel', '#10b981');
        } catch (err) {
            console.error('Export Excel error:', err);
            showCustomAlert('កំហុសឆ្គង', 'មិនអាច Export បានទេ: ' + err.message, 'fa-solid fa-triangle-exclamation', '#dc2626');
        }
    };

    // ✨ NEW 4b: Export PDF (បើកទំព័ររបាយការណ៍ → Print / Save as PDF)
    window.exportDashboardPDF = function () {
        const studies = getFilteredStudies();
        if (studies.length === 0) {
            showCustomAlert('ជូនដំណឹង', 'គ្មានទិន្នន័យសម្រាប់ Export ទេ!', 'fa-solid fa-circle-info', '#3b82f6');
            return;
        }

        const f = getCurrentFilterValues();
        const lb = computeLineBalance(studies);
        const pfms = studies.map(s => parseFloat(s.pfm) || 0).filter(v => v > 0);
        const avgPfm = pfms.length ? (pfms.reduce((a, b) => a + b, 0) / pfms.length).toFixed(1) : '0';
        let chartImg = '';
        try { if (window.myCTChart) chartImg = window.myCTChart.toBase64Image(); } catch (e) { chartImg = ''; }

        const sorted = [...studies].sort((a, b) =>
            (parseInt(a.seq) || parseInt(a.operatorSeq) || 99) - (parseInt(b.seq) || parseInt(b.operatorSeq) || 99));

        const rowsHtml = sorted.map((s, i) => {
            const pfm = parseFloat(s.pfm) || 0;
            const pc = pfmColors(pfm);
            return `<tr>
                <td>${escapeHtml(s.no || (i + 1))}</td>
                <td>${escapeHtml(s.seq !== undefined ? s.seq : (s.operatorSeq || ''))}</td>
                <td>${escapeHtml(s.sewerName || 'No Sewer')}</td>
                <td>${escapeHtml(s.processRegion || '')}</td>
                <td>${escapeHtml(s.styleName || '')}</td>
                <td>${escapeHtml(s.rampupDay || '')}</td>
                <td>${escapeHtml(s.smv || '0.00')}</td>
                <td><b>${formatToMinSec(s.averageCT || s.avgSeconds || 0)}</b></td>
                <td>${escapeHtml(s.output || 0)}</td>
                <td style="background:${pc.bg};color:${pc.color};font-weight:bold;">${pfm}%</td>
            </tr>`;
        }).join('');

        const filterText = Object.keys(FILTER_DEFS)
            .map(k => FILTER_DEFS[k].label + ': ' + (f[k] || 'All')).join(' | ');

        const html = `<!DOCTYPE html>
<html lang="km"><head><meta charset="UTF-8"><title>Time Study Report</title>
<style>
    body{font-family:'Khmer OS Battambang','Noto Sans Khmer',Arial,sans-serif;margin:24px;color:#1e293b;}
    h1{font-size:20px;margin:0 0 4px;} .sub{color:#64748b;font-size:12px;margin-bottom:14px;}
    .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0;}
    .kpi{border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;} .kpi small{color:#64748b;display:block;font-size:11px;} .kpi b{font-size:16px;}
    img{max-width:100%;margin:10px 0;border:1px solid #e2e8f0;border-radius:8px;}
    table{width:100%;border-collapse:collapse;font-size:11px;margin-top:10px;}
    th,td{border:1px solid #cbd5e1;padding:5px 6px;text-align:left;} th{background:#f1f5f9;}
    @media print{body{margin:10mm;} .noprint{display:none;}}
</style></head><body>
<div class="noprint" style="margin-bottom:10px;"><button onclick="window.print()" style="padding:8px 16px;font-size:14px;cursor:pointer;">🖨️ Print / Save as PDF</button></div>
<h1>Time Study Report</h1>
<div class="sub">${escapeHtml(filterText)} &nbsp;•&nbsp; បង្កើតនៅ: ${escapeHtml(new Date().toLocaleString('km-KH'))}</div>
<div class="kpis">
    <div class="kpi"><small>Total Studies</small><b>${studies.length}</b></div>
    <div class="kpi"><small>PFM មធ្យម</small><b>${avgPfm}%</b></div>
    <div class="kpi"><small>Balance Rate</small><b>${lb.balanceRate ? lb.balanceRate.toFixed(1) + '%' : '—'}</b></div>
    <div class="kpi"><small>Bottleneck</small><b>${escapeHtml(lb.bottleneck || '—')}${lb.max ? ' (' + formatToMinSec(lb.max) + ')' : ''}</b></div>
    <div class="kpi"><small>Takt Time</small><b>${lb.takt ? lb.takt + 's' : '—'}</b></div>
    <div class="kpi"><small>Line Efficiency</small><b>${lb.takt ? lb.lineEfficiency.toFixed(1) + '%' : '—'}</b></div>
    <div class="kpi"><small>Total CT</small><b>${formatToMinSec(lb.sum)}</b></div>
    <div class="kpi"><small>Output អតិបរមា/ម៉ោង</small><b>${lb.hourlyOutput ? Math.floor(lb.hourlyOutput) + ' Pcs' : '—'}</b></div>
</div>
${chartImg ? `<img src="${chartImg}" alt="Yamazumi Chart">` : ''}
<table>
    <thead><tr><th>No</th><th>Seq</th><th>Sewer</th><th>Process</th><th>Style</th><th>Day</th><th>SMV</th><th>CT (Avg)</th><th>Output</th><th>PFM</th></tr></thead>
    <tbody>${rowsHtml}</tbody>
</table>
</body></html>`;

        const w = window.open('', '_blank');
        if (!w) {
            showCustomAlert('ជូនដំណឹង', 'Browser បានទប់ស្កាត់ Popup សូមអនុញ្ញាត Popup សម្រាប់គេហទំព័រនេះ រួចព្យាយាមម្ដងទៀត។', 'fa-solid fa-circle-info', '#3b82f6');
            return;
        }
        w.document.open();
        w.document.write(html);
        w.document.close();
        setTimeout(() => { try { w.focus(); w.print(); } catch (e) { /* user can print manually */ } }, 600);
    };

    window.deleteDashboardStudy = function (id) {
        const existing = safeLoad('savedTimeStudies').find(item => item.id === id); if (!existing) return; if (!tspCanManageRecord(existing)) { tspPermissionDenied(); return; }
        showCustomConfirm('បញ្ជាក់ការលុប', 'តើអ្នកពិតជាចង់លុបទិន្នន័យ Time Study នេះមែនទេ?', () => {
            let savedStudies = safeLoad('savedTimeStudies');
            savedStudies = savedStudies.filter(item => item.id !== id);
            safeSave('savedTimeStudies', savedStudies);
            window.renderDashboardStats();
            showToast('លុបជោគជ័យ', 'ទិន្នន័យត្រូវបានលុបចេញពីប្រព័ន្ធ', 'fa-solid fa-trash', '#dc2626');
        });
    };

    window.editStudyRecord = function (id) {
        const savedStudies = safeLoad('savedTimeStudies');
        const study = savedStudies.find(item => item.id === id);
        if (!study) return;
        if (!tspCanManageRecord(study)) { tspPermissionDenied(); return; }

        window.currentEditMode = 'dashboard';

        const setVal = (elId, v) => { const el = document.getElementById(elId); if (el) el.value = v; };
        setVal('edit-study-id', study.id);
        setVal('edit-factory', study.title || study.factory || '');
        setVal('edit-section', study.section || '');
        setVal('edit-prep', study.prepAssembly || study.prepName || study.prep || study.prep_assembly || '');
        setVal('edit-style', study.styleName || '');
        setVal('edit-rampup', study.rampupDay || '');
        setVal('edit-sewer', study.sewerName || '');
        setVal('edit-seq', study.operatorSeq !== undefined ? study.operatorSeq : (study.seq || ''));
        setVal('edit-process', study.processRegion || '');
        setVal('edit-smv', study.smv || 0);

        const lapsContainer = document.getElementById('edit-laps-container');
        if (lapsContainer) {
            if (study.laps && study.laps.length > 0) {
                lapsContainer.innerHTML = study.laps.map((lap, index) => {
                    const secVal = formatToMinSec(lap.durationMs || 0, true);
                    return `
                        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; gap: 8px;">
                            <span style="font-size: 12px; font-weight: bold; color: #475569;">Round ${lap.round || (index + 1)}:</span>
                            <input type="text" class="edit-lap-input" data-index="${index}" value="${secVal}" style="flex: 1; padding: 4px 8px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 13px;" placeholder="MM:SS">
                            <span style="font-size: 12px; color: #64748b;">(MM:SS)</span>
                        </div>
                    `;
                }).join('');
            } else {
                lapsContainer.innerHTML = '<div style="font-size: 12px; color: #94a3b8; text-align: center;">គ្មានទិន្នន័យ Lap ទេ</div>';
            }
        }

        const editModal = document.getElementById('editStudyModal');
        if (editModal) editModal.style.display = 'flex';
    };

    window.saveEditedStudy = function () {
        if (window.currentEditMode === 'observe') {
            return window.saveTimeObserveEdit();
        }

        const editIdEl = document.getElementById('edit-study-id');
        if (!editIdEl) return;
        const id = editIdEl.value;
        const savedStudies = safeLoad('savedTimeStudies');

        const index = savedStudies.findIndex(item => item.id === id);
        if (index === -1) return;
        if (!tspCanManageRecord(savedStudies[index])) { tspPermissionDenied(); return; }

        const rec = savedStudies[index];
        const getVal = (elId) => { const el = document.getElementById(elId); return el ? el.value : ''; };

        rec.title = getVal('edit-factory');
        rec.factory = rec.title;
        rec.section = getVal('edit-section');

        const newPrepAssembly = getVal('edit-prep');
        rec.prepAssembly = newPrepAssembly;
        rec.prepName = newPrepAssembly;

        rec.styleName = getVal('edit-style');
        rec.rampupDay = getVal('edit-rampup');
        rec.sewerName = getVal('edit-sewer');
        rec.operatorSeq = parseInt(getVal('edit-seq')) || 99;
        rec.seq = rec.operatorSeq;
        rec.processRegion = getVal('edit-process');

        const newSmv = parseFloat(getVal('edit-smv')) || 0;
        rec.smv = newSmv.toFixed(2);
        // មិនប៉ះពាល់ primaryKey ទេ (ត្រូវការសម្រាប់ផ្គូផ្គងជាមួយ Excel)

        const lapInputs = document.querySelectorAll('.edit-lap-input');
        if (lapInputs.length > 0) {
            const newLaps = [];

            lapInputs.forEach((input, idx) => {
                const valStr = input.value.trim();
                let durationMs = 0;

                if (valStr.includes(':')) {
                    const parts = valStr.split(':');
                    durationMs = Math.round(((parseFloat(parts[0]) || 0) * 60 + (parseFloat(parts[1]) || 0)) * 1000);
                } else {
                    durationMs = Math.round((parseFloat(valStr) || 0) * 1000);
                }
                newLaps.push({ round: idx + 1, durationMs: durationMs });
            });

            rec.laps = newLaps;

            const st = calcLapStats(newLaps);
            const totalSeconds = st.avg / 1000;

            rec.avgSeconds = totalSeconds;
            rec.averageCT = formatToMinSec(st.avg, true);
            rec.minSeconds = st.min / 1000;
            rec.maxSeconds = st.max / 1000;
            rec.stdSeconds = st.std / 1000;
            rec.outlierCount = st.outliers.length;

            if (totalSeconds > 0) {
                rec.pfm = ((newSmv / totalSeconds) * 100).toFixed(1);
                rec.output = Math.round(3600 / totalSeconds);
            } else {
                rec.pfm = '0.0';
                rec.output = 0;
            }
        } else if (parseFloat(rec.avgSeconds) > 0) {
            rec.pfm = ((newSmv / parseFloat(rec.avgSeconds)) * 100).toFixed(1);
        }

        if (!safeSave('savedTimeStudies', savedStudies)) {
            showCustomAlert('កំហុសរក្សាទុក', 'មិនអាចរក្សាទុកការកែទិន្នន័យបានទេ។ សូមព្យាយាមម្ដងទៀត។', 'fa-solid fa-triangle-exclamation', '#dc2626');
            return;
        }
        if (typeof window.tspRecordRecentActivity === 'function') {
            const processLabel = studyData.processRegion || studyData.processName || 'Process';
            window.tspRecordRecentActivity('Time Observe data edited', studyData.title || studyData.factory || 'Time Study', `${processLabel}${studyData.sewerName ? ' · Sewer: ' + studyData.sewerName : ''}`);
        }

        window.closeEditModal();
        window.renderDashboardStats();

        showToast('ជោគជ័យ', 'ទិន្នន័យត្រូវបានកែប្រែ និងរក្សាទុករួចរាល់!', 'fa-solid fa-check', '#10b981');
    };

});

// --- ៧. Global Utility Functions ---
function showCustomAlert(title, message, iconClass = 'fa-solid fa-circle-check', iconColor = '#10b981') {
    const modal = document.getElementById('custom-modal');
    if (!modal) {
        alert(title + '\n' + message);
        return;
    }
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-message').innerText = message;
    document.getElementById('modal-icon').innerHTML = `<i class="${iconClass}" style="color: ${iconColor};"></i>`;
    const btnCancel = document.getElementById('modal-btn-cancel');
    if (btnCancel) btnCancel.style.display = 'none';
    const btnOk = document.getElementById('modal-btn-ok');
    modal.style.display = 'flex';
    if (btnOk) {
        btnOk.textContent = 'យល់ព្រម';
        btnOk.className = 'modal-btn modal-btn-primary';
        const newBtnOk = btnOk.cloneNode(true);
        btnOk.parentNode.replaceChild(newBtnOk, btnOk);
        newBtnOk.addEventListener('click', () => modal.style.display = 'none');
    }
}

function showCustomConfirm(title, message, onConfirm, yesLabel = 'លុប', onCancel) {
    let modal = document.getElementById('custom-confirm-modal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'custom-confirm-modal';
        modal.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); display:flex; align-items:center; justify-content:center; z-index:9999;';
        modal.innerHTML = `
            <div style="background:#fff; padding:25px; border-radius:12px; width:350px; text-align:center; box-shadow:0 4px 6px rgba(0,0,0,0.1);">
                <h3 id="confirm-title" style="margin-bottom:10px; color:#1e293b;"></h3>
                <p id="confirm-msg" style="margin-bottom:20px; color:#64748b; font-size:14px;"></p>
                <div style="display:flex; gap:10px; justify-content:center;">
                    <button id="confirm-yes-btn" style="background:#dc2626; color:#fff; border:none; padding:8px 20px; border-radius:6px; cursor:pointer; font-weight:bold;">លុប</button>
                    <button id="confirm-no-btn" style="background:#e2e8f0; color:#334155; border:none; padding:8px 20px; border-radius:6px; cursor:pointer; font-weight:bold;">បោះបង់</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-msg').textContent = message;
    document.getElementById('confirm-yes-btn').textContent = yesLabel;
    modal.style.display = 'flex';

    document.getElementById('confirm-yes-btn').onclick = function () {
        modal.style.display = 'none';
        if (onConfirm) onConfirm();
    };
    document.getElementById('confirm-no-btn').onclick = function () {
        modal.style.display = 'none';
        if (onCancel) onCancel();
    };
}

function showToast(title, message, iconClass, iconColor) {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = 'toast success';
    toast.innerHTML = `<i class="${escapeHtml(iconClass)}" style="color: ${escapeHtml(iconColor)};"></i><div class="toast-content"><h4>${escapeHtml(title)}</h4><p>${escapeHtml(message)}</p></div>`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

/* Table Resizer: pointer-based dragging, minimum widths, and table expands as columns grow. */
function makeTableResizable(table) {
    if (!table || table.dataset.resizerReady === 'true') return;
    const headerRow = table.tHead && table.tHead.rows[0];
    if (!headerRow || !headerRow.cells.length) return;

    table.dataset.resizerReady = 'true';
    table.classList.add('resizable-table');
    table.style.setProperty('table-layout', 'fixed', 'important');
    table.style.setProperty('width', 'max-content', 'important');
    table.style.setProperty('min-width', '100%', 'important');

    // A colgroup gives the browser a stable width target for each column.
    let colgroup = table.querySelector(':scope > colgroup');
    if (!colgroup) {
        colgroup = document.createElement('colgroup');
        table.insertBefore(colgroup, table.firstChild);
    }
    const headers = Array.from(headerRow.cells);
    while (colgroup.children.length < headers.length) colgroup.appendChild(document.createElement('col'));
    while (colgroup.children.length > headers.length) colgroup.lastElementChild.remove();

    headers.forEach((th, index) => {
        th.style.setProperty('position', 'relative', 'important');
        th.style.setProperty('overflow', 'visible', 'important');
        th.style.setProperty('white-space', 'nowrap', 'important');
        const col = colgroup.children[index];
        const initialWidth = Math.max(70, Math.round(th.getBoundingClientRect().width || th.offsetWidth || 100));
        col.style.width = initialWidth + 'px';
        th.style.width = initialWidth + 'px';
        th.style.minWidth = '70px';

        let handle = th.querySelector('.resizer');
        if (!handle) {
            handle = document.createElement('div');
            handle.className = 'resizer';
            handle.setAttribute('role', 'separator');
            handle.setAttribute('aria-orientation', 'vertical');
            handle.setAttribute('aria-label', 'Resize column ' + (th.textContent.trim() || (index + 1)));
            handle.tabIndex = 0;
            th.appendChild(handle);
        }
        if (handle.dataset.bound === 'true') return;
        handle.dataset.bound = 'true';

        let startX = 0;
        let startWidth = 0;
        let activePointerId = null;

        const stopResize = () => {
            if (activePointerId === null) return;
            handle.classList.remove('dragging');
            th.classList.remove('resizing');
            document.body.classList.remove('table-column-resizing');
            try { if (handle.hasPointerCapture(activePointerId)) handle.releasePointerCapture(activePointerId); } catch (_) {}
            activePointerId = null;
        };

        handle.addEventListener('pointerdown', (event) => {
            if (event.button !== undefined && event.button !== 0) return;
            event.preventDefault();
            event.stopPropagation();
            activePointerId = event.pointerId;
            startX = event.clientX;
            startWidth = col.getBoundingClientRect().width || th.getBoundingClientRect().width;
            handle.classList.add('dragging');
            th.classList.add('resizing');
            document.body.classList.add('table-column-resizing');
            try { handle.setPointerCapture(event.pointerId); } catch (_) {}
        });

        handle.addEventListener('pointermove', (event) => {
            if (activePointerId === null || event.pointerId !== activePointerId) return;
            const minWidth = 70;
            const maxWidth = 1200;
            const nextWidth = Math.max(minWidth, Math.min(maxWidth, Math.round(startWidth + event.clientX - startX)));
            col.style.width = nextWidth + 'px';
            th.style.width = nextWidth + 'px';
            th.style.minWidth = minWidth + 'px';
            th.style.maxWidth = 'none';
            // Keep the table at least as wide as the sum of the columns so widths do not snap back.
            const totalWidth = Array.from(colgroup.children).reduce((sum, item) => sum + (parseFloat(item.style.width) || item.getBoundingClientRect().width || 70), 0);
            table.style.setProperty('width', Math.max(totalWidth, table.parentElement?.clientWidth || 0) + 'px', 'important');
        });

        handle.addEventListener('pointerup', stopResize);
        handle.addEventListener('pointercancel', stopResize);
        handle.addEventListener('lostpointercapture', stopResize);
        handle.addEventListener('keydown', (event) => {
            if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
            event.preventDefault();
            const current = parseFloat(col.style.width) || th.getBoundingClientRect().width;
            const next = Math.max(70, Math.min(1200, current + (event.key === 'ArrowRight' ? 10 : -10)));
            col.style.width = next + 'px';
            th.style.width = next + 'px';
            const totalWidth = Array.from(colgroup.children).reduce((sum, item) => sum + (parseFloat(item.style.width) || item.getBoundingClientRect().width || 70), 0);
            table.style.setProperty('width', Math.max(totalWidth, table.parentElement?.clientWidth || 0) + 'px', 'important');
        });
    });
}

function installTableResizerStyles() {
    if (document.getElementById('table-resizer-improved-styles')) return;
    const style = document.createElement('style');
    style.id = 'table-resizer-improved-styles';
    style.textContent = `
        .table-responsive { position: relative; width: 100%; max-width: 100%; overflow-x: auto !important; overflow-y: hidden; -webkit-overflow-scrolling: touch; }
        .files-table.resizable-table { border-collapse: separate !important; border-spacing: 0; }
        .files-table.resizable-table th { position: relative !important; user-select: none; white-space: nowrap !important; overflow: visible !important; }
        .files-table.resizable-table td { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .files-table .resizer { position: absolute; top: 0; right: -4px; width: 9px; height: 100%; min-height: 100%; cursor: col-resize; z-index: 20; background: transparent; touch-action: none; user-select: none; }
        .files-table .resizer::after { content: ''; position: absolute; top: 0; bottom: 0; left: 4px; width: 1px; background: #cbd5e1; }
        .files-table .resizer:hover, .files-table th.resizing .resizer, .files-table .resizer.dragging { background: rgba(59,130,246,.16); }
        .files-table th.resizing .resizer::after, .files-table .resizer:hover::after, .files-table .resizer.dragging::after { width: 2px; background: #2563eb; }
        body.table-column-resizing { cursor: col-resize !important; user-select: none !important; }
        body.table-column-resizing * { cursor: col-resize !important; user-select: none !important; }
    `;
    document.head.appendChild(style);
}

window.addEventListener('pageshow', function () {
    const activeRampupDay = localStorage.getItem('lastSelectedRampupDay');
    const activeTitle = localStorage.getItem('lastSelectedTitle');
    const activeSection = localStorage.getItem('lastSelectedSection');
    const activeStyle = localStorage.getItem('lastSelectedStyle');

    if (activeRampupDay && activeTitle && activeSection && activeStyle && typeof window.renderPrepAssemblyList === 'function') {
        window.renderPrepAssemblyList(activeTitle, activeSection, activeStyle, activeRampupDay);
    }
});

/* Workspace zoom controls now live inside Settings; no floating zoom panel is created. */


/* Home Welcome Screen: always shown on browser refresh; navigation continues to use existing pages. */
(function installHomeWelcomeScreen() {
    function setupHomeWelcomeScreen() {
        const workspace = document.querySelector('.main-content');
        if (!workspace || document.getElementById('tsp-home-welcome')) return;

        const navLinks = Array.from(document.querySelectorAll('.sidebar ul li a[data-page]'))
            .filter(link => link.getAttribute('data-page') && (link.textContent || '').trim());
        const dashboardLink = navLinks.find(link => link.getAttribute('data-page') === 'dashboard-page');
        const quickLinks = navLinks.filter(link => link !== dashboardLink && link.getAttribute('data-page') !== 'home-page').slice(0, 6);

        const style = document.createElement('style');
        style.id = 'tsp-home-welcome-styles';
        style.textContent = `
            .main-content { position: relative; }
            #tsp-home-welcome { position: absolute; inset: 0; z-index: 5000; overflow: auto; padding: clamp(22px, 3.2vw, 48px); isolation: isolate; background: radial-gradient(ellipse at 14% 7%, rgba(59,130,246,.29), transparent 34%), radial-gradient(ellipse at 90% 82%, rgba(139,92,246,.25), transparent 36%), radial-gradient(ellipse at 48% 110%, rgba(14,165,233,.10), transparent 42%), linear-gradient(135deg,#060d19 0%,#0b172a 48%,#090f1d 100%); color: #f8fafc; border-radius: 14px; animation: tspHomeFade .65s ease both; }
            /* Leave the fixed mobile navigation outside the Welcome overlay's hit area. */
            @media(max-width:600px) { #tsp-home-welcome { bottom: calc(82px + env(safe-area-inset-bottom)); } }
            #tsp-home-welcome::before, #tsp-home-welcome::after { content:''; position:absolute; z-index:-1; width: min(38vw, 420px); height: min(38vw, 420px); border-radius:50%; filter:blur(3px); pointer-events:none; opacity:.4; animation:tspHomeOrb 13s ease-in-out infinite alternate; }
            #tsp-home-welcome::marker { content:''; }
            #tsp-home-welcome .tsp-home-inner::before { content:''; position:absolute; inset:0; z-index:-1; pointer-events:none; opacity:.13; background-image:linear-gradient(rgba(148,163,184,.16) 1px,transparent 1px),linear-gradient(90deg,rgba(148,163,184,.16) 1px,transparent 1px); background-size:34px 34px; mask-image:linear-gradient(to bottom,black,transparent 82%); }
            #tsp-home-welcome::before { top:-130px; right:4%; background:radial-gradient(circle,rgba(59,130,246,.55),rgba(59,130,246,0) 70%); }
            #tsp-home-welcome::after { bottom:-160px; left:3%; background:radial-gradient(circle,rgba(139,92,246,.45),rgba(139,92,246,0) 70%); animation-delay:-5s; }
            @keyframes tspHomeFade { from { opacity:0; } to { opacity:1; } }
            @keyframes tspHomeOrb { from { transform:translate3d(-10px,0,0) scale(.94); } to { transform:translate3d(26px,18px,0) scale(1.08); } }
            @keyframes tspHomeRise { from { opacity:0; transform:translateY(16px); } to { opacity:1; transform:translateY(0); } }
            #tsp-home-welcome[hidden] { display: none !important; }
            #tsp-home-welcome .tsp-home-inner { position:relative; width: min(100%, 1120px); margin: 0 auto; }
            #tsp-home-welcome .tsp-home-brand { display:flex; align-items:center; gap:14px; margin-bottom: clamp(26px,4.5vh,46px); animation:tspHomeRise .55s .04s ease both; }
            #tsp-home-welcome .tsp-home-logo { width:56px;height:56px;display:grid;place-items:center;border-radius:17px;background:linear-gradient(135deg,#38bdf8 0%,#3b82f6 48%,#8b5cf6 100%);color:#fff;font-size:26px;box-shadow:0 0 0 1px rgba(255,255,255,.2),0 14px 34px rgba(59,130,246,.34),inset 0 1px 0 rgba(255,255,255,.22); }
            #tsp-home-welcome .tsp-home-brand-name { font-size:clamp(19px,2.3vw,25px);font-weight:800;letter-spacing:-.5px; }
            #tsp-home-welcome .tsp-home-subtitle { color:#a8b8d0;font-size:13px; }
            #tsp-home-welcome .tsp-home-hero { position:relative; overflow:hidden; padding:clamp(26px,4.2vw,48px);border:1px solid rgba(191,219,254,.22);border-radius:26px;background:linear-gradient(118deg,rgba(255,255,255,.115),rgba(255,255,255,.055) 58%,rgba(59,130,246,.075));backdrop-filter:blur(22px);-webkit-backdrop-filter:blur(22px);box-shadow:0 28px 76px rgba(0,0,0,.30),inset 0 1px 0 rgba(255,255,255,.14),inset 0 -1px 0 rgba(255,255,255,.035);animation:tspHomeRise .65s .12s ease both; }
            #tsp-home-welcome .tsp-home-hero::after { content:''; position:absolute; width:260px;height:260px;right:-112px;top:-142px;border-radius:50%;background:radial-gradient(circle,rgba(96,165,250,.24),rgba(99,102,241,.06) 52%,transparent 72%);filter:blur(8px);pointer-events:none; }
            #tsp-home-welcome .tsp-home-eyebrow { display:inline-flex;align-items:center;gap:8px;color:#bfdbfe;font-size:11px;font-weight:800;letter-spacing:1.8px;text-transform:uppercase;margin-bottom:15px;padding:7px 11px;border:1px solid rgba(147,197,253,.19);border-radius:999px;background:rgba(59,130,246,.09); }
            #tsp-home-welcome .tsp-home-eyebrow::before { content:'';width:6px;height:6px;border-radius:50%;background:#60a5fa;box-shadow:0 0 12px #60a5fa; }
            #tsp-home-welcome h1 { font-size:clamp(32px,4.8vw,54px);line-height:1.08;letter-spacing:-1.65px;margin:0 0 17px;font-weight:850;background:linear-gradient(95deg,#fff 0%,#e0eaff 55%,#a5b4fc 100%);-webkit-background-clip:text;background-clip:text;color:transparent; }
            #tsp-home-welcome .tsp-home-description { max-width:690px;color:#bac8dc;font-size:clamp(14px,1.5vw,16px);line-height:1.85;margin-bottom:27px; }
            #tsp-home-welcome .tsp-home-actions { display:flex;gap:10px;flex-wrap:wrap; }
            #tsp-home-welcome button { border:0;cursor:pointer;border-radius:11px;padding:12px 17px;font-weight:700;transition:transform .22s ease,box-shadow .22s ease,background .22s ease,border-color .22s ease; }
            #tsp-home-welcome button:hover { transform:translateY(-1px); }
            #tsp-home-welcome .tsp-home-primary { color:#fff;background:linear-gradient(110deg,#2563eb,#4f46e5 70%,#6366f1);box-shadow:0 10px 28px rgba(59,130,246,.32),inset 0 1px 0 rgba(255,255,255,.2); }
            #tsp-home-welcome .tsp-home-primary:hover { box-shadow:0 12px 30px rgba(59,130,246,.45); }
            #tsp-home-welcome .tsp-home-secondary { color:#e2e8f0;background:rgba(15,23,42,.24);border:1px solid rgba(203,213,225,.24);box-shadow:inset 0 1px 0 rgba(255,255,255,.06); }
            #tsp-home-welcome .tsp-home-section-title { display:flex;align-items:center;gap:10px;margin:31px 0 15px;font-size:15px;font-weight:800;letter-spacing:.15px;color:#e5edf8;animation:tspHomeRise .55s .2s ease both; }
            #tsp-home-welcome .tsp-home-section-title::after { content:'';height:1px;flex:1;background:linear-gradient(90deg,rgba(148,163,184,.3),transparent); }
            #tsp-home-welcome .tsp-home-cards { display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,235px),1fr));gap:14px; }
            #tsp-home-welcome .tsp-home-card { position:relative;display:flex;align-items:center;gap:14px;text-align:left;width:100%;min-height:86px;padding:17px;background:linear-gradient(135deg,rgba(255,255,255,.105),rgba(255,255,255,.04) 75%);border:1px solid rgba(148,163,184,.23) !important;border-radius:17px;backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);box-shadow:0 14px 32px rgba(0,0,0,.18),inset 0 1px 0 rgba(255,255,255,.08);animation:tspHomeRise .55s ease both; }
            #tsp-home-welcome .tsp-home-card:nth-child(1) { animation-delay:.22s; } #tsp-home-welcome .tsp-home-card:nth-child(2) { animation-delay:.28s; } #tsp-home-welcome .tsp-home-card:nth-child(3) { animation-delay:.34s; } #tsp-home-welcome .tsp-home-card:nth-child(4) { animation-delay:.4s; } #tsp-home-welcome .tsp-home-card:nth-child(5) { animation-delay:.46s; } #tsp-home-welcome .tsp-home-card:nth-child(6) { animation-delay:.52s; }
            #tsp-home-welcome .tsp-home-card:hover { transform:translateY(-5px);border-color:rgba(96,165,250,.62) !important;box-shadow:0 22px 42px rgba(0,0,0,.25),0 0 0 1px rgba(96,165,250,.10),0 0 28px rgba(59,130,246,.08);background:linear-gradient(135deg,rgba(59,130,246,.19),rgba(255,255,255,.065)); }
            #tsp-home-welcome .tsp-home-card-icon { width:44px;height:44px;flex:0 0 44px;display:grid;place-items:center;border-radius:13px;background:linear-gradient(135deg,rgba(59,130,246,.24),rgba(139,92,246,.22));border:1px solid rgba(147,197,253,.22);color:#a5cfff;font-size:18px;box-shadow:inset 0 1px 0 rgba(255,255,255,.07);transition:transform .22s ease,background .22s ease; }
            #tsp-home-welcome .tsp-home-card:hover .tsp-home-card-icon { transform:scale(1.07) rotate(-3deg);background:linear-gradient(135deg,rgba(59,130,246,.32),rgba(139,92,246,.3)); }
            #tsp-home-welcome .tsp-home-card-text { min-width:0; }
            #tsp-home-welcome .tsp-home-card-title { display:block;font-size:13px;font-weight:800;color:#f1f5f9;letter-spacing:.1px; }
            #tsp-home-welcome .tsp-home-card-hint { display:block;margin-top:5px;font-size:11px;font-weight:400;line-height:1.45;color:#a8b8d0; }
            #tsp-home-welcome .tsp-home-footer { margin-top:24px;text-align:center;color:#8495af;font-size:11px; }
            @media(prefers-reduced-motion:reduce) { #tsp-home-welcome, #tsp-home-welcome::before, #tsp-home-welcome::after, #tsp-home-welcome .tsp-home-brand, #tsp-home-welcome .tsp-home-hero, #tsp-home-welcome .tsp-home-section-title, #tsp-home-welcome .tsp-home-card { animation:none !important; transition:none !important; } }
            @media(max-width:600px) { #tsp-home-welcome { padding:15px; } #tsp-home-welcome .tsp-home-brand { margin-bottom:22px; } #tsp-home-welcome .tsp-home-hero { border-radius:17px; } #tsp-home-welcome h1 { letter-spacing:-.8px; } #tsp-home-welcome .tsp-home-card { min-height:74px;padding:13px; } }
        `;
        document.head.appendChild(style);

        const home = document.createElement('section');
        home.id = 'tsp-home-welcome';
        home.setAttribute('aria-label', 'TimeStudyPro Home Welcome');
        const navCardMarkup = quickLinks.map((link, index) => {
            const title = (link.textContent || '').trim().replace(/\s+/g, ' ');
            const icons = ['fa-list-check','fa-stopwatch','fa-file-arrow-up','fa-chart-column','fa-folder-open','fa-sliders'];
            const hints = ['Open this workspace section','Go to this feature','Open and manage your data','Review your work'];
            return `<button type="button" class="tsp-home-card" data-home-page="${escapeHtml(link.getAttribute('data-page'))}"><span class="tsp-home-card-icon"><i class="fa-solid ${icons[index % icons.length]}"></i></span><span class="tsp-home-card-text"><span class="tsp-home-card-title">${escapeHtml(title)}</span><span class="tsp-home-card-hint">${hints[index % hints.length]}</span></span></button>`;
        }).join('');
        home.innerHTML = `
            <div class="tsp-home-inner">
                <div class="tsp-home-brand"><div class="tsp-home-logo"><i class="fa-solid fa-stopwatch"></i></div><div><div class="tsp-home-brand-name">TimeStudyPro</div><div class="tsp-home-subtitle">Time Study &amp; Process Analysis</div></div></div>
                <div class="tsp-home-hero">
                    <div class="tsp-home-eyebrow">Welcome to your workspace</div>
                    <h1>Work smarter.<br>Improve every process.</h1>
                    <p class="tsp-home-description">Your workspace for time study, process analysis, line balancing and production improvement. Choose a section below to get started.</p>
                    <div class="tsp-home-actions">
                        ${dashboardLink ? '<button type="button" class="tsp-home-primary" data-home-page="dashboard-page"><i class="fa-solid fa-chart-line"></i>&nbsp; Open Dashboard &nbsp; <i class="fa-solid fa-arrow-right"></i></button>' : ''}
                        <button type="button" class="tsp-home-secondary" data-home-dismiss><i class="fa-solid fa-house"></i>&nbsp; Continue to workspace</button>
                    </div>
                </div>
                ${navCardMarkup ? `<div class="tsp-home-section-title">Quick Access</div><div class="tsp-home-cards">${navCardMarkup}</div>` : ''}
                <div class="tsp-home-footer">TimeStudyPro · Time Study &amp; Process Analysis</div>
            </div>`;
        workspace.appendChild(home);

        function hideHome() { home.hidden = true; }
        function openPage(pageId) {
            const link = navLinks.find(item => item.getAttribute('data-page') === pageId);
            if (link) link.click();
            hideHome();
        }
        home.addEventListener('click', event => {
            const pageButton = event.target.closest('[data-home-page]');
            if (pageButton) { openPage(pageButton.getAttribute('data-home-page')); return; }
            if (event.target.closest('[data-home-dismiss]')) {
                hideHome();
                const visiblePage = Array.from(document.querySelectorAll('.page-content')).find(page => getComputedStyle(page).display !== 'none');
                if (!visiblePage && navLinks[0]) navLinks[0].click();
            }
        });
        // Every refresh starts at Home. Any normal sidebar navigation dismisses the welcome screen.
        navLinks.forEach(link => link.addEventListener('click', hideHome, true));
        document.querySelectorAll('.page-content').forEach(page => { page.style.display = 'none'; });
        menuLinksHomeReset();
        function menuLinksHomeReset() {
            navLinks.forEach(link => link.classList.remove('active'));
        }
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', setupHomeWelcomeScreen, { once: true });
    else setupHomeWelcomeScreen();
})();


/* ===== Integrated Login / Profile / User Management (isolated add-on) ===== */
/* TimeStudyPro User Management add-on. Browser-local demo only; use a backend for secure authentication. */
(function () {
  if (window.__TSP_USER_MANAGEMENT_V1__) return;
  window.__TSP_USER_MANAGEMENT_V1__ = true;
  const USERS_KEY = 'tsp_user_management_users_v1';
  const SESSION_KEY = 'tsp_user_management_session_v1';
  const defaultUsers = [{ id: 'admin-default', username: 'SOKRET11', fullName: 'SOKRET11', password: 'Admin123!', role: 'Admin', createdAt: new Date().toISOString() }];
  const readUsers = () => {
    try { const v = JSON.parse(localStorage.getItem(USERS_KEY) || 'null'); return Array.isArray(v) && v.length ? v : defaultUsers.slice(); }
    catch (_) { return defaultUsers.slice(); }
  };
  let users = readUsers();
  let current = null;
  try { const sid = localStorage.getItem(SESSION_KEY); current = users.find(u => u.id === sid && u.status !== 'disabled') || null; if (!current) localStorage.removeItem(SESSION_KEY); } catch (_) {}
  const saveUsers = () => { try { localStorage.setItem(USERS_KEY, JSON.stringify(users)); } catch (_) {} };
  const AUDIT_KEY = 'tspSecurityAuditLog_v1';
  function auditEvent(action, detail = '', username = (current && current.username) || '') {
    try {
      const rows = JSON.parse(localStorage.getItem(AUDIT_KEY) || '[]');
      rows.unshift({ action, detail: String(detail || ''), username: username || 'Unknown', userId: (current && current.id) || '', at: new Date().toISOString() });
      localStorage.setItem(AUDIT_KEY, JSON.stringify(rows.slice(0, 500)));
    } catch (_) {}
  }
  // Keep the user's current profile photo consistent across all saved UI/history records.
  function syncUserAvatarReferences(user) {
    if (!user || !user.id) return;
    const username = String(user.username || '').toLowerCase();
    const avatar = user.avatar || '';
    const isThisUser = record => !!record && (
      String(record.uploaderId || record.userId || record.id || '') === String(user.id) ||
      String(record.username || '').toLowerCase() === username ||
      String(record.user || '').toLowerCase() === username ||
      String(record.uploaderName || '').toLowerCase() === String(user.fullName || '').toLowerCase()
    );
    const updateList = (key, updater) => {
      try {
        const value = safeLoad(key);
        if (!Array.isArray(value)) return;
        updater(value);
        safeSave(key, value);
      } catch (_) {}
    };
    const updateNestedSnapshots = record => {
      if (!record || typeof record !== 'object') return;
      if (isThisUser(record)) {
        if ('avatar' in record) record.avatar = avatar;
        if ('uploaderAvatar' in record) record.uploaderAvatar = avatar;
        if ('uploaderName' in record) record.uploaderName = user.fullName || user.username;
      }
      if (Array.isArray(record)) record.forEach(updateNestedSnapshots);
      else Object.keys(record).forEach(key => {
        const child = record[key];
        if (child && typeof child === 'object') updateNestedSnapshots(child);
      });
    };
    ['uploadedExcelFiles', 'uploadedExcelTrash'].forEach(key => updateList(key, list => {
      list.forEach(item => {
        if (String(item && item.uploaderId || '') === String(user.id) || String(item && item.user || '').toLowerCase() === username) {
          item.uploaderAvatar = avatar;
          item.uploaderName = user.fullName || user.username;
          item.user = user.username;
        }
        updateNestedSnapshots(item);
      });
    }));
    ['tspUploadActivityLog', 'tsp_user_activity_log_v1'].forEach(key => updateList(key, list => {
      list.forEach(item => {
        if (isThisUser(item)) {
          if ('avatar' in item) item.avatar = avatar;
          if ('uploaderAvatar' in item) item.uploaderAvatar = avatar;
          if ('fullName' in item) item.fullName = user.fullName || user.username;
          if ('user' in item && String(item.user).toLowerCase() === username) item.user = user.fullName || user.username;
        }
      });
    }));
    try { window.dispatchEvent(new CustomEvent('tsp:profile-updated', { detail: { userId: user.id, avatar } })); } catch (_) {}
  }
  saveUsers();
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const $ = id => document.getElementById(id);
  function toast(msg, kind) {
    const box = $('toast-container');
    if (box) { const item = document.createElement('div'); item.className = 'tsp-user-toast ' + (kind || ''); item.textContent = msg; box.appendChild(item); setTimeout(() => item.remove(), 3000); }
    else alert(msg);
  }
  const overlay = document.createElement('div'); overlay.id = 'tsp-user-overlay'; overlay.className = 'tsp-user-overlay'; overlay.style.display = 'none'; document.body.appendChild(overlay);
  function showLogin() {
    overlay.innerHTML = `<div class="tsp-user-card tsp-login-card"><div class="tsp-user-logo"><i class="fa-solid fa-user-shield"></i></div><h2>TimeStudyPro Login</h2><p class="tsp-user-muted">ចូលប្រើគណនីរបស់អ្នក</p><form id="tsp-login-form"><label>Username</label><input id="tsp-login-name" autocomplete="username" required autofocus><label>Password</label><input id="tsp-login-pass" type="password" autocomplete="current-password" required><p id="tsp-login-error" class="tsp-user-error"></p><button class="tsp-user-primary" type="submit"><i class="fa-solid fa-right-to-bracket"></i> Login</button></form><p class="tsp-user-hint">គណនីដំបូង: <b>SOKRET11</b> / <b>Admin123!</b></p><p class="tsp-user-security">នេះជាកំណែ Browser-local សម្រាប់សាកល្បង។ មិនទាន់មាន Server Authentication ទេ។</p></div>`;
    overlay.style.display = 'flex';
    $('tsp-login-form').addEventListener('submit', e => { e.preventDefault(); const name = $('tsp-login-name').value.trim().toLowerCase(); const pass = $('tsp-login-pass').value; const u = users.find(x => x.username.toLowerCase() === name && x.password === pass && x.status !== 'disabled'); if (!u) { auditEvent('Login failed', 'Invalid credentials or disabled account', name); $('tsp-login-error').textContent = 'Username/Password មិនត្រឹមត្រូវ ឬគណនីត្រូវបានបិទ។'; return; } current = u; try { localStorage.setItem(SESSION_KEY, u.id); } catch (_) {} auditEvent('Login successful', 'Browser-local session', u.username); overlay.style.display = 'none'; updateProfile(); toast('Login ជោគជ័យ'); });
  }
  function updateProfile() {
    const profile = $('profile-btn'); if (!profile) return;
    const img = profile.querySelector('img'); const span = profile.querySelector('span');
    if (span) span.innerHTML = `${esc(current ? current.username : 'Login')} <i class="fa-solid fa-chevron-down"></i>`;
    if (img) img.src = (current && current.avatar) || 'SOKRET.png.jpg';
    profile.title = current ? `${current.fullName || current.username} (${current.role})` : 'Login';
    let menu = $('tsp-profile-menu'); if (!menu) { menu = document.createElement('div'); menu.id = 'tsp-profile-menu'; menu.className = 'tsp-profile-menu'; profile.parentElement.style.position = 'relative'; profile.parentElement.appendChild(menu); }
    menu.innerHTML = current ? `<div class="tsp-profile-summary"><div class="tsp-avatar">${current.avatar ? `<img src="${current.avatar}" alt="Profile photo">` : '<i class="fa-solid fa-user"></i>'}</div><div><b>${esc(current.fullName || current.username)}</b><small>${esc(current.username)} · ${esc(current.role)}</small></div></div><button data-action="profile"><i class="fa-regular fa-id-card"></i> My Profile</button>${current.role === 'Admin' ? '<button data-action="users"><i class="fa-solid fa-users-gear"></i> User Management</button>' : ''}<button data-action="logout" class="tsp-logout"><i class="fa-solid fa-right-from-bracket"></i> Logout</button>` : `<button data-action="login"><i class="fa-solid fa-right-to-bracket"></i> Login</button>`;
    menu.querySelectorAll('button').forEach(b => b.addEventListener('click', e => { e.stopPropagation(); menu.classList.remove('open'); if (b.dataset.action === 'login') showLogin(); if (b.dataset.action === 'logout') { auditEvent('Logout', 'User logged out'); current = null; try { localStorage.removeItem(SESSION_KEY); } catch (_) {} refreshPresence(); updateProfile(); showLogin(); } if (b.dataset.action === 'users') showUsers(); if (b.dataset.action === 'profile') showProfile(); if (b.dataset.action === 'active-users') showActiveUsers(); if (b.dataset.action === 'smart-tools') showSmartTools(); if (b.dataset.action === 'workspace-settings') showWorkspaceSettings(); }));
    addActiveUserMenuItem();
    initAdminNotification();
    refreshPresence();
  }
  function showModal(title, body, footer) {
    overlay.innerHTML = `<div class="tsp-user-card tsp-user-dialog"><div class="tsp-user-dialog-head"><h2>${title}</h2><button type="button" id="tsp-user-close" aria-label="Close"><i class="fa-solid fa-xmark"></i></button></div><div class="tsp-user-dialog-body">${body}</div><div class="tsp-user-dialog-foot">${footer || '<button type="button" class="tsp-user-secondary" id="tsp-user-close-2">Close</button>'}</div></div>`;
    overlay.style.display = 'flex';
    overlay.querySelectorAll('#tsp-user-close, #tsp-user-close-2').forEach(b => b.addEventListener('click', () => { overlay.style.display = 'none'; }));
    overlay.onclick = e => { if (e.target === overlay && current) overlay.style.display = 'none'; };
  }
  function showProfile() {
    if (!current) return showLogin();
    showModal('My Profile', `<div class="tsp-profile-edit"><div class="tsp-profile-photo-area"><div class="tsp-avatar tsp-avatar-large" id="tsp-profile-avatar">${current.avatar ? `<img src="${current.avatar}" alt="Profile photo">` : '<i class="fa-solid fa-user"></i>'}</div><label class="tsp-user-primary tsp-photo-upload" for="tsp-profile-photo"><i class="fa-solid fa-camera"></i> Upload Photo</label><input id="tsp-profile-photo" type="file" accept="image/*" hidden><small class="tsp-photo-hint">រូបភាព JPG, PNG ឬ WebP (អតិបរមា 2 MB)</small><button type="button" class="tsp-user-secondary tsp-photo-remove" id="tsp-profile-photo-remove" ${current.avatar ? '' : 'hidden'}>Remove Photo</button></div><label>Full name</label><input id="tsp-profile-fullname" value="${esc(current.fullName || '')}" required><label>Username</label><input id="tsp-profile-username" value="${esc(current.username)}" required><label>New password <small>(ទុកទទេ ប្រសិនបើមិនប្តូរ)</small></label><input id="tsp-profile-password" type="password" autocomplete="new-password"></div>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Cancel</button><button type="button" class="tsp-user-primary" id="tsp-profile-save">Save Profile</button>');
    overlay.querySelector('.tsp-user-dialog').classList.add('tsp-profile-dialog');
    let selectedAvatar = current.avatar || '';
    $('tsp-profile-photo').addEventListener('change', e => { const file = e.target.files && e.target.files[0]; if (!file) return; if (!file.type.startsWith('image/')) { toast('សូមជ្រើសរើសឯកសាររូបភាព។','error'); e.target.value = ''; return; } if (file.size > 2 * 1024 * 1024) { toast('រូបភាពធំពេក។ សូមជ្រើសរើសរូបភាពក្រោម 2 MB។','error'); e.target.value = ''; return; } const reader = new FileReader(); reader.onload = () => { selectedAvatar = String(reader.result || ''); $('tsp-profile-avatar').innerHTML = `<img src="${selectedAvatar}" alt="Profile photo">`; $('tsp-profile-photo-remove').hidden = false; }; reader.readAsDataURL(file); });
    $('tsp-profile-photo-remove').addEventListener('click', () => { selectedAvatar = ''; $('tsp-profile-avatar').innerHTML = '<i class="fa-solid fa-user"></i>'; $('tsp-profile-photo-remove').hidden = true; $('tsp-profile-photo').value = ''; });
    $('tsp-profile-save').addEventListener('click', () => { const fullName = $('tsp-profile-fullname').value.trim(), username = $('tsp-profile-username').value.trim(), password = $('tsp-profile-password').value; if (!fullName || !username) return toast('សូមបំពេញព័ត៌មានចាំបាច់។','error'); if (users.some(u => u.id !== current.id && u.username.toLowerCase() === username.toLowerCase())) return toast('Username នេះមានរួចហើយ។','error'); current.fullName = fullName; current.username = username; current.avatar = selectedAvatar; if (password) current.password = password; saveUsers(); auditEvent(password ? 'Profile and password updated' : 'Profile updated', 'Profile details/photo saved'); syncUserAvatarReferences(current); updateProfile(); overlay.style.display = 'none'; toast('បានរក្សាទុក Profile និងធ្វើបច្ចុប្បន្នភាពរូបភាពគ្រប់កន្លែង'); });
  }
  function showUsers() {
    if (!current || current.role !== 'Admin') return toast('មានតែ Admin ប៉ុណ្ណោះដែលអាចគ្រប់គ្រង User បាន។','error');
    const activeIds = getActiveUserIds();
    const rows = users.map((u,i) => `<tr><td>${i+1}</td><td><div class="tsp-user-cell"><span class="tsp-avatar tsp-table-avatar">${u.avatar ? `<img src="${esc(u.avatar)}" alt="">` : '<i class="fa-solid fa-user"></i>'}</span><span><b>${esc(u.fullName || u.username)}</b><small>${esc(u.username)}</small></span></div></td><td><span class="tsp-role ${u.role === 'Admin' ? 'admin' : ''}">${esc(u.role)}</span></td><td><span class="tsp-presence ${activeIds.has(u.id) ? 'is-active' : 'is-inactive'}"><i></i>${activeIds.has(u.id) ? 'Active' : 'Offline'}</span><small class="tsp-account-state ${u.status === 'disabled' ? 'disabled' : ''}">${u.status === 'disabled' ? 'Disabled' : 'Enabled'}</small></td><td>${u.id === current.id ? '<span class="tsp-current-tag">You</span>' : ''}</td><td class="tsp-row-actions"><button data-edit="${esc(u.id)}" title="Edit / Reset Password"><i class="fa-solid fa-pen"></i></button><button data-toggle-status="${esc(u.id)}" title="${u.status === 'disabled' ? 'Enable account' : 'Disable account'}" ${u.id === current.id || (u.role === 'Admin' && users.filter(x => x.role === 'Admin' && x.status !== 'disabled').length <= 1) ? 'disabled' : ''}><i class="fa-solid ${u.status === 'disabled' ? 'fa-user-check' : 'fa-user-lock'}"></i></button><button data-delete="${esc(u.id)}" title="Delete" ${u.id === current.id ? 'disabled' : ''}><i class="fa-solid fa-trash"></i></button></td></tr>`).join('');
    showModal('User Management', `<div class="tsp-users-toolbar"><p>គណនី: <b>${users.length}</b> · <span class="tsp-presence is-active"><i></i>${activeIds.size} Active</span></p><input id="tsp-user-search" type="search" placeholder="ស្វែងរក User…" aria-label="Search users"><button type="button" id="tsp-add-user" class="tsp-user-primary"><i class="fa-solid fa-user-plus"></i> Add User</button></div><div class="tsp-users-table-wrap"><table class="tsp-users-table"><thead><tr><th>#</th><th>User</th><th>Role</th><th>Status</th><th></th><th>Actions</th></tr></thead><tbody>${rows || '<tr><td colspan="6">No users</td></tr>'}</tbody></table></div><p class="tsp-user-security">Account permissions in this browser are local UI controls; production-grade security requires server-side authentication.</p>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Done</button>');
    $('tsp-add-user').addEventListener('click', () => editUser());
    $('tsp-user-search').addEventListener('input', e => { const q=e.target.value.trim().toLowerCase(); overlay.querySelectorAll('.tsp-users-table tbody tr').forEach(tr=>tr.hidden=!tr.textContent.toLowerCase().includes(q)); });
    overlay.querySelectorAll('[data-edit]').forEach(b => b.addEventListener('click', () => editUser(b.dataset.edit)));
    overlay.querySelectorAll('[data-delete]').forEach(b => b.addEventListener('click', () => deleteUser(b.dataset.delete)));
    overlay.querySelectorAll('[data-toggle-status]').forEach(b => b.addEventListener('click', () => { const target=users.find(x=>x.id===b.dataset.toggleStatus); if(!target || target.id===current.id)return; if(target.role==='Admin' && target.status!=='disabled' && users.filter(x=>x.role==='Admin' && x.status!=='disabled').length<=1)return toast('ត្រូវរក្សា Admin សកម្មយ៉ាងតិចម្នាក់។','error'); target.status=target.status==='disabled'?'active':'disabled'; saveUsers(); auditEvent(target.status==='disabled'?'Account disabled':'Account enabled', target.username); showUsers(); toast(target.status==='disabled'?'បានបិទគណនី User':'បានបើកគណនី User'); }));
  }
  function editUser(id) {
    const u = users.find(x => x.id === id) || { id: '', username: '', fullName: '', password: '', role: 'User', status: 'active', permissions: { upload: true, download: true, deleteOwn: true } };
    showModal(id ? 'Edit User' : 'Add User', `<form id="tsp-user-form"><label>Full name</label><input id="tsp-user-fullname" value="${esc(u.fullName)}" required><label>Username</label><input id="tsp-user-username" value="${esc(u.username)}" required><label>${id ? 'Reset Password (ទុកទទេបើមិនប្តូរ)' : 'Password'}</label><input id="tsp-user-password" type="password" ${id ? '' : 'required'} autocomplete="new-password"><small class="tsp-password-note">សម្រាប់សុវត្ថិភាព Admin អាចកំណត់/Reset Password ថ្មីបាន ប៉ុន្តែមិនបង្ហាញ Password ដែល User កំពុងប្រើទេ។</small><label>Role</label><select id="tsp-user-role"><option value="User" ${u.role==='User'?'selected':''}>User</option><option value="Admin" ${u.role==='Admin'?'selected':''}>Admin</option></select><label>សិទ្ធិ User</label><div class="tsp-permission-options"><label><input type="checkbox" id="tsp-perm-upload" ${(u.permissions?.upload !== false)?'checked':''}> Upload files</label><label><input type="checkbox" id="tsp-perm-download" ${(u.permissions?.download !== false)?'checked':''}> Download files</label><label><input type="checkbox" id="tsp-perm-delete-own" ${(u.permissions?.deleteOwn !== false)?'checked':''}> Delete own files</label></div></form>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Cancel</button><button type="button" class="tsp-user-primary" id="tsp-user-save"><i class="fa-solid fa-floppy-disk"></i> Save User</button>');
    $('tsp-user-save').addEventListener('click', () => { const fullName = $('tsp-user-fullname').value.trim(), username = $('tsp-user-username').value.trim(), password = $('tsp-user-password').value, role = $('tsp-user-role').value, permissions = { upload: $('tsp-perm-upload').checked, download: $('tsp-perm-download').checked, deleteOwn: $('tsp-perm-delete-own').checked }; if (!fullName || !username || (!id && !password)) return toast('សូមបំពេញព័ត៌មានចាំបាច់។','error'); if (users.some(x => x.id !== id && x.username.toLowerCase() === username.toLowerCase())) return toast('Username នេះមានរួចហើយ។','error'); if (id) { const target = users.find(x => x.id === id); if (!target) return; if (target.role === 'Admin' && role !== 'Admin' && users.filter(x => x.role === 'Admin').length === 1) return toast('ត្រូវរក្សា Admin យ៉ាងតិចម្នាក់។','error'); target.fullName = fullName; target.username = username; target.role = role; target.permissions = permissions; if (!target.status) target.status = 'active'; if (password) target.password = password; } else users.push({ id: 'u-' + Date.now() + '-' + Math.random().toString(36).slice(2,8), fullName, username, password, role, status: 'active', permissions, createdAt: new Date().toISOString() }); saveUsers(); auditEvent(id ? 'User updated' : 'User created', username); current = users.find(x => x.id === current.id) || current; updateProfile(); showUsers(); toast(id ? 'បានកែប្រែ User' : 'បានបង្កើត User'); });
  }
  function deleteUser(id) {
    const u = users.find(x => x.id === id); if (!u || u.id === current.id) return;
    showModal('Delete User?', `<p>តើអ្នកពិតជាចង់លុបគណនី <b>${esc(u.username)}</b> មែនទេ? សកម្មភាពនេះមិនអាចត្រឡប់វិញបានទេ។</p>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Cancel</button><button type="button" class="tsp-user-danger" id="tsp-user-delete-confirm">Delete User</button>');
    $('tsp-user-delete-confirm').addEventListener('click', () => { if (users.filter(x => x.role === 'Admin').length === 1 && u.role === 'Admin') return toast('ត្រូវរក្សា Admin យ៉ាងតិចម្នាក់។','error'); users = users.filter(x => x.id !== id); saveUsers(); auditEvent('User deleted', u.username); showUsers(); toast('បានលុប User'); });
  }
  const ACTIVE_KEY = 'tsp_user_presence_v1';
  const ACTIVITY_KEY = 'tspUploadActivityLog';
  // Shared writer for confirmed Clone Day and Time Observe save operations.
  window.tspRecordRecentActivity = function(action, name, detail = '') {
    try {
      const rows = JSON.parse(localStorage.getItem(ACTIVITY_KEY) || '[]');
      const actor = current || {};
      rows.unshift({ action: String(action || 'Activity'), name: String(name || '—'), detail: String(detail || ''), user: actor.fullName || actor.username || 'Unknown', fullName: actor.fullName || actor.username || 'Unknown', username: actor.username || 'Unknown', at: new Date().toISOString() });
      localStorage.setItem(ACTIVITY_KEY, JSON.stringify(rows.slice(0, 300)));
      window.dispatchEvent(new CustomEvent('tsp:activity-updated', {detail: rows[0]}));
      return true;
    } catch (_) { return false; }
  };
  const PRESENCE_TTL = 65000;
  let presenceTimer = null;
  let activeRefreshTimer = null;
  let activityPanel = null;

  function getActiveSessions() {
    try {
      const now = Date.now();
      const list = JSON.parse(localStorage.getItem(ACTIVE_KEY) || '[]');
      const fresh = Array.isArray(list) ? list.filter(x => x && x.userId && Number(x.lastSeen) && now - Number(x.lastSeen) < PRESENCE_TTL) : [];
      if (JSON.stringify(fresh) !== JSON.stringify(list)) localStorage.setItem(ACTIVE_KEY, JSON.stringify(fresh));
      return fresh;
    } catch (_) { return []; }
  }
  function getActiveUserIds() {
    return new Set(getActiveSessions().map(x => x.userId));
  }
  function getTabId() {
    try {
      let id = sessionStorage.getItem('tsp_presence_tab_id');
      if (!id) { id = 'tab-' + Date.now() + '-' + Math.random().toString(36).slice(2, 9); sessionStorage.setItem('tsp_presence_tab_id', id); }
      return id;
    } catch (_) { return 'tab-' + Math.random().toString(36).slice(2, 10); }
  }
  const tabId = getTabId();
  function refreshPresence() {
    try {
      const sid = localStorage.getItem(SESSION_KEY);
      const activeUser = users.find(u => u.id === sid);
      let sessions = getActiveSessions().filter(x => x.tabId !== tabId);
      if (activeUser) {
        sessions.push({ tabId, userId: activeUser.id, username: activeUser.username, fullName: activeUser.fullName || activeUser.username, role: activeUser.role, lastSeen: Date.now() });
      }
      localStorage.setItem(ACTIVE_KEY, JSON.stringify(sessions));
      window.dispatchEvent(new CustomEvent('tsp:presence-updated'));
    } catch (_) {}
    refreshActiveLabels();
  }
  function refreshActiveLabels() {
    const ids = getActiveUserIds();
    const rows = overlay.querySelectorAll('[data-presence-id]');
    rows.forEach(el => {
      const on = ids.has(el.getAttribute('data-presence-id'));
      el.classList.toggle('is-active', on); el.classList.toggle('is-inactive', !on);
      const label = el.querySelector('span'); if (label) label.textContent = on ? 'Active' : 'Offline';
    });
    const count = overlay.querySelector('#tsp-active-count'); if (count) count.textContent = String(ids.size);
  }
  function todaysActivities() {
    try {
      const list = JSON.parse(localStorage.getItem(ACTIVITY_KEY) || '[]');
      const today = new Date().toDateString();
      return (Array.isArray(list) ? list : []).filter(a => a && a.at && new Date(a.at).toDateString() === today);
    } catch (_) { return []; }
  }
  function allActivities() {
    try { const list = JSON.parse(localStorage.getItem(ACTIVITY_KEY) || '[]'); return (Array.isArray(list) ? list : []).filter(a => a && a.at).sort((a,b) => new Date(b.at) - new Date(a.at)); } catch (_) { return []; }
  }
  function activitySummary() {
    const list = todaysActivities();
    const usersFor = predicate => new Set(list.filter(predicate).map(a => a.username || a.user || 'Unknown'));
    const uploads = usersFor(a => /upload/i.test(a.action || ''));
    const deletes = usersFor(a => /delete|recycle bin/i.test(a.action || ''));
    const downloads = usersFor(a => /download/i.test(a.action || ''));
    return { list, uploads, deletes, downloads };
  }
  const ACTIVITY_READ_KEY = 'tspAdminActivityReadAt_v1';
  function getActivityReadAt() {
    try { return Number(localStorage.getItem(ACTIVITY_READ_KEY) || 0); } catch (_) { return 0; }
  }
  function markActivitiesRead() {
    try { localStorage.setItem(ACTIVITY_READ_KEY, String(Date.now())); } catch (_) {}
    refreshNotificationBadge();
  }
  function refreshNotificationBadge() {
    const badge = document.getElementById('tsp-notification-badge');
    if (!badge) return;
    const isAdmin = current && current.role === 'Admin';
    const readAt = getActivityReadAt();
    const count = isAdmin ? todaysActivities().filter(a => new Date(a.at).getTime() > readAt).length : 0;
    badge.hidden = !isAdmin || count === 0;
    badge.textContent = count > 99 ? '99+' : String(count);
    const bell = $('notification-btn');
    if (bell) bell.title = isAdmin ? (count ? `${count} unread activity notifications` : 'Admin notifications — all caught up') : 'Notifications';
  }
  function showActivityPanel() {
    if (!current || current.role !== 'Admin') return toast('ការជូនដំណឹងសកម្មភាពអាចមើលបានតែ Admin ប៉ុណ្ណោះ។', 'error');
    if (activityPanel) activityPanel.remove();
    markActivitiesRead();
    const summary = activitySummary();
    const activityRecords = allActivities();
    activityPanel = document.createElement('div');
    activityPanel.className = 'tsp-admin-activity-panel';
    let activeFilter = 'all';
    const filterActivities = () => { const query = (activityPanel?.querySelector('#tsp-activity-search')?.value || '').trim().toLowerCase(); const dateValue = activityPanel?.querySelector('#tsp-activity-date')?.value || ''; return activityRecords.filter(a => {
      const identity = `${a.fullName || ''} ${a.username || ''} ${a.user || ''} ${a.name || ''}`.toLowerCase();
      if (query && !identity.includes(query)) return false;
      if (dateValue && new Date(a.at).toISOString().slice(0,10) !== dateValue) return false;
      if (activeFilter === 'upload') return /upload/i.test(a.action || '');
      if (activeFilter === 'delete') return /delete|recycle bin/i.test(a.action || '');
      if (activeFilter === 'download') return /download/i.test(a.action || '');
      return true;
    }); };
    const renderRows = () => {
      const list = filterActivities().slice(0, 5);
      const target = activityPanel && activityPanel.querySelector('.tsp-admin-activity-list');
      if (!target) return;
      target.innerHTML = list.length ? list.map(a => `<div class="tsp-admin-activity-item"><span class="tsp-activity-avatar">${a.avatar ? `<img src="${esc(a.avatar)}" alt="">` : '<i class="fa-solid fa-user"></i>'}</span><div class="tsp-admin-activity-main"><b>${esc(a.fullName || a.username || a.user || 'Unknown')}</b><small>${esc(a.action || 'Activity')} · ${esc(a.name || '—')}</small><small class="tsp-activity-username">${esc(a.username || '')}${a.detail ? ' · ' + esc(a.detail) : ''}</small></div><time>${new Date(a.at).toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'})}</time></div>`).join('') : '<p class="tsp-admin-activity-empty">មិនមានសកម្មភាពក្នុងប្រភេទនេះទេ។</p>';
      activityPanel.querySelectorAll('.tsp-admin-activity-stat').forEach(b => b.classList.toggle('is-selected', b.dataset.filter === activeFilter));
    };
    activityPanel.innerHTML = `<div class="tsp-admin-activity-head"><div><b><i class="fa-regular fa-bell"></i> សកម្មភាពថ្ងៃនេះ</b><small>${new Date().toLocaleDateString()}</small></div><button type="button" aria-label="Close"><i class="fa-solid fa-xmark"></i></button></div><div class="tsp-admin-activity-stats"><button type="button" class="tsp-admin-activity-stat" data-filter="upload" aria-label="View users who uploaded files today"><i class="fa-solid fa-upload"></i><b>${summary.uploads.size}</b><small>User Upload</small></button><button type="button" class="tsp-admin-activity-stat" data-filter="delete" aria-label="View users who deleted files today"><i class="fa-solid fa-trash"></i><b>${summary.deletes.size}</b><small>User Delete</small></button><button type="button" class="tsp-admin-activity-stat" data-filter="download" aria-label="View users who downloaded files today"><i class="fa-solid fa-download"></i><b>${summary.downloads.size}</b><small>User Download</small></button></div><div class="tsp-activity-filters"><input id="tsp-activity-search" type="search" placeholder="ស្វែងរក User ឬ File…"><input id="tsp-activity-date" type="date" aria-label="Filter activity date"></div><div class="tsp-admin-activity-caption">ចុចស្ថិតិដើម្បីច្រោះប្រភេទសកម្មភាព · បង្ហាញ 5 សកម្មភាពថ្មីៗបំផុត</div><div class="tsp-admin-activity-list"></div>`;
    const header = document.querySelector('header');
    if (header) { header.appendChild(activityPanel); } else document.body.appendChild(activityPanel);
    activityPanel.querySelector('.tsp-admin-activity-head button').addEventListener('click', () => { activityPanel.remove(); activityPanel = null; });
    activityPanel.querySelectorAll('.tsp-admin-activity-stat').forEach(button => button.addEventListener('click', () => { activeFilter = button.dataset.filter; renderRows(); }));
    activityPanel.querySelector('#tsp-activity-search').addEventListener('input', renderRows);
    activityPanel.querySelector('#tsp-activity-date').addEventListener('change', renderRows);
    renderRows();
    document.addEventListener('click', function closeActivityOnce(e) {
      if (activityPanel && !activityPanel.contains(e.target) && e.target.id !== 'notification-btn' && !e.target.closest('#notification-btn')) {
        activityPanel.remove(); activityPanel = null; document.removeEventListener('click', closeActivityOnce);
      }
    });
  }
  function initAdminNotification() {
    const bell = $('notification-btn');
    if (!bell || bell.dataset.tspAdminBound) return;
    bell.dataset.tspAdminBound = '1';
    // Keep the unread counter anchored to the bell itself, not the whole user-menu.
    let bellWrap = bell.parentElement && bell.parentElement.classList.contains('tsp-notification-bell-wrap') ? bell.parentElement : null;
    if (!bellWrap) {
      bellWrap = document.createElement('span');
      bellWrap.className = 'tsp-notification-bell-wrap';
      bell.parentNode.insertBefore(bellWrap, bell);
      bellWrap.appendChild(bell);
    }
    bell.classList.add('tsp-notification-bell');
    let badge = document.getElementById('tsp-notification-badge');
    if (!badge) {
      badge = document.createElement('span');
      badge.id = 'tsp-notification-badge';
      badge.className = 'tsp-notification-badge';
      badge.hidden = true;
      bellWrap.appendChild(badge);
    }
    bell.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); if (activityPanel) { activityPanel.remove(); activityPanel = null; } else showActivityPanel(); });
    window.addEventListener('tsp:activity-updated', refreshNotificationBadge);
    window.addEventListener('storage', e => { if (e.key === ACTIVITY_KEY || e.key === ACTIVE_KEY || e.key === ACTIVITY_READ_KEY) refreshNotificationBadge(); });
    setInterval(refreshNotificationBadge, 15000);
    refreshNotificationBadge();
  }
  function addActiveUserMenuItem() {
    const menu = $('tsp-profile-menu');
    if (!menu || !current) return;
    if (current.role === 'Admin') {
      const items = [
        ['active-users', '<i class="fa-solid fa-user-group"></i> Active Users'],
        ['smart-tools', '<i class="fa-solid fa-wand-magic-sparkles"></i> Smart Admin Tools']
      ];
      items.forEach(([action, html]) => {
        if (menu.querySelector(`[data-action="${action}"]`)) return;
        const btn = document.createElement('button'); btn.dataset.action = action; btn.innerHTML = html;
        const logout = menu.querySelector('[data-action="logout"]'); menu.insertBefore(btn, logout || null);
        btn.addEventListener('click', e => { e.stopPropagation(); menu.classList.remove('open'); if (action === 'active-users') showActiveUsers(); else showSmartTools(); });
      });
    }
    if (!menu.querySelector('[data-action="workspace-settings"]')) {
      const btn = document.createElement('button'); btn.dataset.action = 'workspace-settings'; btn.innerHTML = '<i class="fa-solid fa-sliders"></i> Workspace Settings';
      const logout = menu.querySelector('[data-action="logout"]'); menu.insertBefore(btn, logout || null);
      btn.addEventListener('click', e => { e.stopPropagation(); menu.classList.remove('open'); showWorkspaceSettings(); });
    }
  }

  function readAuditLog() { try { const a = JSON.parse(localStorage.getItem(AUDIT_KEY) || '[]'); return Array.isArray(a) ? a : []; } catch (_) { return []; } }
  function localBackupKeys() {
    const keys = [];
    try {
      for (let i=0;i<localStorage.length;i++) {
        const k = localStorage.key(i); if (!k) continue;
        if (k === SESSION_KEY || k === ACTIVE_KEY || k.includes('Presence') || k.includes('presence')) continue;
        if (/^(tsp|uploadedExcel|uploadedExcelTrash|uploadedExcelFiles|tspUpload|tsp_user|tspAdmin|tspWorkspace|tspSecurity|upload)/i.test(k)) keys.push(k);
      }
    } catch (_) {}
    return keys.sort();
  }
  function downloadBackup() {
    try {
      const data = {}; localBackupKeys().forEach(k => { try { data[k] = JSON.parse(localStorage.getItem(k)); } catch (_) { data[k] = localStorage.getItem(k); } });
      if (window.tspLargeStorageCache) ['uploadedExcelFiles','uploadedExcelTrash'].forEach(k => { data[k] = window.tspLargeStorageCache[k] || []; });
      const payload = { app: 'TimeStudyPro', formatVersion: 1, createdAt: new Date().toISOString(), note: 'Browser-local backup. Protect this file because it may contain account and business data.', data };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `TimeStudyPro_Backup_${new Date().toISOString().slice(0,10)}.json`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
      auditEvent('Backup exported', `${Object.keys(data).length} localStorage keys`); toast('បានបង្កើត Backup ឯកសារ JSON');
    } catch (e) { toast('មិនអាចបង្កើត Backup បានទេ។ សូមពិនិត្យទំហំទិន្នន័យ Browser។','error'); }
  }
  function restoreBackupFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const payload = JSON.parse(String(reader.result || '{}'));
        if (payload.app !== 'TimeStudyPro' || !payload.data || typeof payload.data !== 'object') throw new Error('Invalid backup');
        if (!window.confirm('Restore នឹងជំនួសទិន្នន័យ localStorage ដែលមានក្នុង Backup។ សូមប្រាកដថាអ្នកបាន Backup ទិន្នន័យបច្ចុប្បន្នរួច។ បន្តឬទេ?')) return;
        Object.entries(payload.data).forEach(([k,v]) => {
          if (k === SESSION_KEY || k === ACTIVE_KEY || !/^(tsp|uploadedExcel|uploadedExcelTrash|uploadedExcelFiles|tspUpload|tsp_user|tspAdmin|tspWorkspace|tspSecurity|upload)/i.test(k)) return;
          if (k === 'uploadedExcelFiles' || k === 'uploadedExcelTrash') { if (!safeSave(k, Array.isArray(v) ? v : [])) throw new Error('Could not restore large data'); }
          else localStorage.setItem(k, typeof v === 'string' ? v : JSON.stringify(v));
        });
        auditEvent('Backup restored', file.name); toast('Restore ជោគជ័យ។ សូម Refresh ទំព័រ។');
        if (window.confirm('Refresh ទំព័រឥឡូវនេះឬទេ?')) location.reload();
      } catch (_) { toast('File Backup មិនត្រឹមត្រូវ ឬមិនអាចអានបាន។','error'); }
    };
    reader.onerror = () => toast('មិនអាចអាន Backup File បានទេ។','error'); reader.readAsText(file);
  }
  function showSecurityHistory() {
    const rows = readAuditLog();
    const html = `<div class="tsp-smart-list">${rows.length ? rows.slice(0,100).map(r=>`<div class="tsp-smart-row"><span class="tsp-smart-icon"><i class="fa-solid fa-shield-halved"></i></span><div><b>${esc(r.action)}</b><small>${esc(r.username)}${r.detail ? ' · '+esc(r.detail):''}</small></div><time>${esc(new Date(r.at).toLocaleString())}</time></div>`).join('') : '<p class="tsp-user-muted">មិនទាន់មានប្រវត្តិសុវត្ថិភាពទេ។</p>'}</div><p class="tsp-smart-note">កំណត់ត្រានេះរក្សាទុកក្នុង Browser បច្ចុប្បន្ន។ វាមិនអាចបញ្ជាក់អត្តសញ្ញាណ ឬការពារការកែប្រែដោយអ្នកមានសិទ្ធិចូល Browser បានទេ។</p>`;
    showModal('Security & Login History', html, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button>');
  }
  function showAnalytics() {
    const activities = allActivities(); const today = new Date().toISOString().slice(0,10);
    const todayRows = activities.filter(a => a.at && String(a.at).slice(0,10) === today);
    const count = re => activities.filter(a => re.test(a.action || '')).length;
    const active = getActiveUserIds().size; let files=safeLoad('uploadedExcelFiles');
    const bars = [['Upload',count(/upload/i)],['Download',count(/download/i)],['Delete',count(/delete|recycle bin/i)],['Other',Math.max(0,activities.length-count(/upload|download|delete|recycle bin/i))]];
    const max = Math.max(1,...bars.map(x=>x[1]));
    const html = `<div class="tsp-analytics-kpis"><div><small>Total Users</small><b>${users.length}</b></div><div><small>Active Now</small><b>${active}</b></div><div><small>Uploaded Files</small><b>${files.length}</b></div><div><small>Today Activity</small><b>${todayRows.length}</b></div></div><h3 class="tsp-smart-subtitle">Activity Overview</h3><div class="tsp-analytics-bars">${bars.map(([label,n])=>`<div class="tsp-analytics-bar-row"><span>${label}</span><div><i style="width:${Math.round(n/max*100)}%"></i></div><b>${n}</b></div>`).join('')}</div><p class="tsp-smart-note">ស្ថិតិគណនាពីប្រវត្តិដែលមានក្នុង Browser នេះប៉ុណ្ណោះ។</p>`;
    showModal('User Analytics Dashboard', html, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button>');
  }
  function showSmartFiles() {
    let files=safeLoad('uploadedExcelFiles');
    const render = () => {
      const q = (overlay.querySelector('#tsp-smart-file-search')?.value || '').toLowerCase();
      const body = overlay.querySelector('#tsp-smart-file-rows'); if(!body)return;
      const list = files.map((f,i)=>({...f,__index:i})).filter(f=>String(f.name||'').toLowerCase().includes(q)||String(f.user||f.uploaderName||'').toLowerCase().includes(q));
      body.innerHTML = list.length ? list.map(f=>`<div class="tsp-smart-file-row"><button type="button" class="tsp-favorite-btn ${f.favorite?'is-favorite':''}" data-fav="${f.__index}" aria-label="Toggle favorite"><i class="fa-${f.favorite?'solid':'regular'} fa-star"></i></button><div><b>${esc(f.name||'Unnamed file')}</b><small>${esc(f.uploaderName||f.user||'Unknown User')} · ${esc(f.uploadDate||f.uploadedAt||'')}</small></div><span>${f.favorite?'Favorite':''}</span></div>`).join('') : '<p class="tsp-user-muted">មិនមាន File ដែលត្រូវនឹងការស្វែងរកទេ។</p>';
      body.querySelectorAll('[data-fav]').forEach(btn=>btn.addEventListener('click',()=>{const i=Number(btn.dataset.fav); if(!files[i])return; files[i].favorite=!files[i].favorite; if(safeSave('uploadedExcelFiles',files)){auditEvent(files[i].favorite?'File favorited':'Favorite removed',files[i].name);render();}else{toast('មិនអាចរក្សាទុក Favorite បានទេ។','error');}}));
    };
    showModal('Smart File Management', `<div class="tsp-smart-file-toolbar"><input id="tsp-smart-file-search" type="search" placeholder="ស្វែងរក File ឬ User…"><span>${files.length} Files</span></div><div id="tsp-smart-file-rows" class="tsp-smart-file-list"></div><p class="tsp-smart-note">File Version History និង Recycle Bin ដែលមានស្រាប់នៅ Upload File ត្រូវបានរក្សាទុកដដែល។</p>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button>');
    $('tsp-smart-file-search').addEventListener('input',render); render();
  }
  function showNotifications() {
    const records = allActivities(); const readAt = getActivityReadAt();
    const html = `<div class="tsp-notification-summary"><b>${records.filter(a=>new Date(a.at).getTime()>readAt).length}</b><span>Unread activity records</span><button type="button" id="tsp-notifications-mark-read" class="tsp-user-secondary">Mark all read</button></div><div class="tsp-smart-list">${records.length ? records.slice(0,100).map(a=>`<div class="tsp-smart-row"><span class="tsp-smart-icon"><i class="fa-regular fa-bell"></i></span><div><b>${esc(a.fullName||a.username||a.user||'Unknown')} · ${esc(a.action||'Activity')}</b><small>${esc(a.name||'—')} ${a.detail?'· '+esc(a.detail):''}</small></div><time>${esc(new Date(a.at).toLocaleString())}</time></div>`).join('') : '<p class="tsp-user-muted">មិនទាន់មាន Notification ទេ។</p>'}</div>`;
    showModal('Smart Notification Center', html, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button>');
    $('tsp-notifications-mark-read')?.addEventListener('click',()=>{markActivitiesRead(); auditEvent('Notifications marked read','All activity notifications'); showNotifications();});
  }
  function showWorkspaceSettings() {
    // These preferences are intentionally session-only: a reload starts at defaults.
    const prefs = { zoom: 100, textSize: 'default', theme: 'system' };
    showModal('TimeStudyPro Settings', `<div class="tsp-workspace-settings"><label class="tsp-mobile-hide-zoom" for="tsp-pref-zoom">Zoom Content <b id="tsp-pref-zoom-value">100%</b></label><input class="tsp-mobile-hide-zoom" id="tsp-pref-zoom" type="range" min="80" max="150" step="5" value="100"><p class="tsp-smart-note tsp-mobile-hide-zoom">លើទូរស័ព្ទ អាចប្រើម្រាមដៃពីរដើម្បី Zoom តាម Browser បាន។</p><label for="tsp-pref-theme">Theme</label><select id="tsp-pref-theme"><option value="system" selected>System</option><option value="light">Light</option><option value="dark">Dark</option></select><label for="tsp-pref-text-size">Text Size</label><select id="tsp-pref-text-size"><option value="small">Small</option><option value="default" selected>Default</option><option value="large">Large</option></select><button type="button" class="tsp-user-secondary tsp-mobile-hide-zoom" id="tsp-pref-reset-zoom">Reset Zoom to 100%</button><button type="button" class="tsp-user-secondary" id="tsp-pref-fullscreen"><i class="fa-solid fa-expand"></i> Full Screen</button><p class="tsp-smart-note">ការកំណត់នឹងត្រឡប់ទៅ Default ពេល Reload ឬបើកកម្មវិធីឡើងវិញ។</p></div>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button><button type="button" class="tsp-user-primary" id="tsp-workspace-save">Apply Settings</button>');
    const zoom=$('tsp-pref-zoom');
    zoom.addEventListener('input',()=>{$('tsp-pref-zoom-value').textContent=zoom.value+'%';});
    $('tsp-pref-reset-zoom').addEventListener('click',()=>{zoom.value='100';$('tsp-pref-zoom-value').textContent='100%';applyWorkspacePrefs({zoom:100,theme:$('tsp-pref-theme').value,textSize:$('tsp-pref-text-size').value});});
    $('tsp-pref-fullscreen').addEventListener('click',async()=>{try{if(!document.fullscreenElement) await document.documentElement.requestFullscreen();else await document.exitFullscreen();}catch(_){toast('Full Screen មិនអាចប្រើបានលើ Browser នេះទេ។','error');}});
    $('tsp-workspace-save').addEventListener('click',()=>{const next={zoom:window.matchMedia('(max-width: 600px)').matches?100:Number(zoom.value),theme:$('tsp-pref-theme').value,textSize:$('tsp-pref-text-size').value};applyWorkspacePrefs(next);auditEvent('Workspace preferences applied',JSON.stringify(next));overlay.style.display='none';toast('បានអនុវត្ត Settings');});
  }
  function applyWorkspacePrefs(prefs) {
    try {
      const textSizes = { small: '12px', default: '13px', large: '15px' };
      document.documentElement.style.setProperty('--tsp-user-font-size', textSizes[prefs.textSize] || '13px');
      const workspace = document.querySelector('.main-content');
      if (workspace) workspace.style.zoom = String(Math.max(80, Math.min(150, Number(prefs.zoom) || 100))/100);
      const dark = prefs.theme === 'dark' || (prefs.theme === 'system' && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
      document.body.classList.toggle('tsp-user-dark-theme', dark);
    } catch(_){}
  }
  function showSmartTools() {
    if (!current || current.role !== 'Admin') return;
    showModal('TimeStudyPro Smart Admin Tools', `<p class="tsp-smart-intro">គ្រប់គ្រងប្រវត្តិ សុវត្ថិភាព ទិន្នន័យ និងការកំណត់ Workspace នៅកន្លែងតែមួយ។</p><div class="tsp-smart-tools-grid"><button type="button" data-smart="security"><i class="fa-solid fa-shield-halved"></i><b>Security & Login History</b><small>ប្រវត្តិ Login និងសកម្មភាព Admin</small></button><button type="button" data-smart="backup"><i class="fa-solid fa-database"></i><b>Backup & Restore</b><small>នាំចេញ ឬ Restore JSON Backup</small></button><button type="button" data-smart="analytics"><i class="fa-solid fa-chart-column"></i><b>User Analytics</b><small>ស្ថិតិ User និងសកម្មភាព</small></button><button type="button" data-smart="files"><i class="fa-solid fa-folder-open"></i><b>Smart File Management</b><small>ស្វែងរក និង Favorite Files</small></button><button type="button" data-smart="notifications"><i class="fa-regular fa-bell"></i><b>Notification Center</b><small>មើល និងសម្គាល់សារថាបានអាន</small></button><button type="button" data-smart="workspace"><i class="fa-solid fa-sliders"></i><b>Personalized Workspace</b><small>Zoom, Font និង Theme</small></button></div><p class="tsp-smart-note">Browser-local edition: សុវត្ថិភាព និងការសមកាលកម្មរវាងឧបករណ៍ត្រូវការ Backend/Database ដើម្បីការពារពិតប្រាកដ។</p>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button>');
    overlay.querySelectorAll('[data-smart]').forEach(btn=>btn.addEventListener('click',()=>{const key=btn.dataset.smart;if(key==='security')showSecurityHistory();if(key==='backup')showBackupRestore();if(key==='analytics')showAnalytics();if(key==='files')showSmartFiles();if(key==='notifications')showNotifications();if(key==='workspace')showWorkspaceSettings();}));
  }
  function showBackupRestore() {
    const keys=localBackupKeys();
    showModal('Backup & Restore', `<div class="tsp-backup-panel"><div class="tsp-backup-stat"><i class="fa-solid fa-database"></i><div><b>${keys.length} data items</b><small>Browser-local storage keys</small></div></div><button type="button" class="tsp-user-primary" id="tsp-backup-export"><i class="fa-solid fa-download"></i> Download Backup JSON</button><label class="tsp-backup-import"><i class="fa-solid fa-upload"></i> Choose Backup JSON to Restore<input id="tsp-backup-file" type="file" accept=".json,application/json" hidden></label><p class="tsp-smart-note">⚠️ Backup អាចមានព័ត៌មានគណនី និងទិន្នន័យផ្ទាល់ខ្លួន។ រក្សាឯកសារនេះនៅកន្លែងមានសុវត្ថិភាព។ Restore អាចជំនួសទិន្នន័យបច្ចុប្បន្ន។</p></div>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button>');
    $('tsp-backup-export').addEventListener('click',downloadBackup); $('tsp-backup-file').addEventListener('change',e=>restoreBackupFile(e.target.files && e.target.files[0]));
  }

  function showActiveUsers() {
    if (!current || current.role !== 'Admin') return;
    const ids = getActiveUserIds();
    const rows = users.map((u, i) => {
      const on = ids.has(u.id); const session = getActiveSessions().find(x => x.userId === u.id);
      return `<tr><td>${i+1}</td><td><div class="tsp-user-cell"><span class="tsp-avatar tsp-table-avatar">${u.avatar ? `<img src="${esc(u.avatar)}" alt="">` : '<i class="fa-solid fa-user"></i>'}</span><span><b>${esc(u.fullName || u.username)}</b><small>${esc(u.username)}</small></span></div></td><td><span class="tsp-role ${u.role === 'Admin' ? 'admin' : ''}">${esc(u.role)}</span></td><td><span class="tsp-presence ${on ? 'is-active' : 'is-inactive'}" data-presence-id="${esc(u.id)}"><i></i><span>${on ? 'Active' : 'Offline'}</span></span><small class="tsp-last-seen">${on ? 'Now active' : (session?.lastSeen ? 'Last seen ' + new Date(session.lastSeen).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'}) : 'No recent session')}</small></td></tr>`;
    }).join('');
    showModal('Live Active Users', `<div class="tsp-users-toolbar"><p>កំពុង Active: <b id="tsp-active-count">${ids.size}</b> User</p><span class="tsp-live-label"><i></i> Live status</span></div><div class="tsp-users-table-wrap"><table class="tsp-users-table"><thead><tr><th>#</th><th>User</th><th>Role</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></div>`, '<button type="button" class="tsp-user-secondary" id="tsp-user-close">Close</button>');
    // Give Live Active Users its own compact layout; the header X is the only close control.
    overlay.querySelector('.tsp-user-dialog')?.classList.add('tsp-live-users-dialog');
    if (activeRefreshTimer) clearInterval(activeRefreshTimer);
    activeRefreshTimer = setInterval(refreshActiveLabels, 5000);
    $('tsp-user-close').addEventListener('click', () => { if (activeRefreshTimer) clearInterval(activeRefreshTimer); activeRefreshTimer = null; overlay.style.display = 'none'; });
  }

  function init() {
    const profile = $('profile-btn'); if (profile) { profile.addEventListener('click', e => { e.stopPropagation(); const menu = $('tsp-profile-menu'); if (menu) menu.classList.toggle('open'); }); document.addEventListener('click', e => { const menu = $('tsp-profile-menu'); if (menu && !menu.contains(e.target) && !profile.contains(e.target)) menu.classList.remove('open'); }); }
    updateProfile();
    try { applyWorkspacePrefs({zoom:100,textSize:'default',theme:'system'}); } catch (_) {}
    initAdminNotification();
    refreshPresence();
    if (presenceTimer) clearInterval(presenceTimer);
    presenceTimer = setInterval(refreshPresence, 15000);
    window.addEventListener('storage', e => {
      if (e.key === SESSION_KEY || e.key === USERS_KEY) {
        users = readUsers();
        const sid = localStorage.getItem(SESSION_KEY);
        current = users.find(u => u.id === sid && u.status !== 'disabled') || null;
        if (!current) { try { localStorage.removeItem(SESSION_KEY); } catch (_) {} }
        updateProfile();
        refreshPresence();
      }
      if (e.key === ACTIVE_KEY) refreshActiveLabels();
    });
    window.addEventListener('beforeunload', () => {
      try { localStorage.setItem(ACTIVE_KEY, JSON.stringify(getActiveSessions().filter(x => x.tabId !== tabId))); } catch (_) {}
    });
    if (!current) showLogin();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true }); else init();
})();


/* Clone Day Smart Merge — appended module; existing features remain untouched. */
(function initCloneDaySmartMerge(){
  const start = () => {
    const page = document.getElementById('clone-page');
    const fileSelect = document.getElementById('clone-source-file');
    const daySelect = document.getElementById('clone-source-day');
    const targetsInput = document.getElementById('clone-target-days');
    const previewBtn = document.getElementById('clone-preview-btn');
    const resetBtn = document.getElementById('clone-reset-btn');
    const previewPanel = document.getElementById('clone-preview-panel');
    const reportPanel = document.getElementById('clone-report-panel');
    const sourcePanelPrimary = document.querySelector('.clone-source-panel-primary');
    const sourceBrowser = document.getElementById('clone-source-browser');
    const sourceList = document.getElementById('clone-source-day-list');
    const sourceSearch = document.getElementById('clone-source-search');
    const sourceListSecondary = document.getElementById('clone-source-day-list-secondary');
    const sourceSearchSecondary = document.getElementById('clone-source-search-secondary');
    const sourcePanelSecondary = document.getElementById('clone-source-panel-secondary');
    const sourceFirstHint = document.getElementById('clone-source-first-hint');
    const sourceSecondaryHint = document.getElementById('clone-source-secondary-hint');
    if (!page || !fileSelect || fileSelect.dataset.cloneBound) return;
    fileSelect.dataset.cloneBound = '1';
    let previewPlan = null;
    const sourceEdits = Object.create(null);
    const loadFiles = () => safeLoad('uploadedExcelFiles');
    const userNow = () => { try { const us=JSON.parse(localStorage.getItem('tsp_user_management_users_v1')||'[]'); const sid=localStorage.getItem('tsp_user_management_session_v1'); return (Array.isArray(us)?us:[]).find(u=>u&&u.id===sid)||null; } catch(_){return null;} };
    const canEdit = f => { const u=userNow(); return !!u && (u.role==='Admin' || (!!f && !!u.id && String(f.uploaderId||'')===String(u.id) && (u.permissions?.deleteOwn!==false))); };
    const esc = v => String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
    const norm = v => String(v==null?'':v).trim().toLocaleLowerCase();
    const getRows = file => { const rows=[]; Object.keys(file?.prepMap||{}).forEach(k=>{ const arr=file.prepMap[k]; if(Array.isArray(arr)) arr.forEach(row=>{ if(row) rows.push(row); }); }); return rows; };
    const identity = r => [r.title,r.section,r.styleName,r.rampupDay,r.group,r.prepAssembly,r.primaryKey,r.processRegion,r.processName,r.seq,r.mc,r.weight].map(norm).join('¦');
    const processIdentity = r => [r.title,r.section,r.styleName,r.group,r.prepAssembly,r.primaryKey,r.processRegion,r.processName,r.seq,r.mc,r.weight].map(norm).join('¦');
    const dayNames = file => [...new Set(getRows(file).map(r=>String(r.rampupDay||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:'base'}));
    // Show every non-empty Style value in the source picker, including styles made of letters only.
    const fileStyleLabel = file => {
      const styles = [...new Set(getRows(file).map(r=>String(r?.styleName||'').trim()).filter(Boolean))];
      return styles.length ? styles.join(', ') : 'មិនមាន Style';
    };
    const setStep = n => document.querySelectorAll('[data-clone-step]').forEach(el=>{const i=Number(el.dataset.cloneStep);el.classList.toggle('is-active',i===n);el.classList.toggle('is-done',i<n);});
    const refreshFiles = () => { const files=loadFiles(); const old=fileSelect.value; fileSelect.innerHTML='<option value="">— ជ្រើស File ទិន្នន័យ —</option>'+files.map((f,i)=>`<option value="${i}">${esc(fileStyleLabel(f))}</option>`).join(''); if(files[Number(old)]) fileSelect.value=old; refreshDays();  };
    const renderSourceDays = () => {
      if(!sourceBrowser||!sourceList)return;
      const files=loadFiles(), f=fileSelect.value!==''?files[Number(fileSelect.value)]:null;
      sourceBrowser.hidden=!f;
      if(sourceFirstHint)sourceFirstHint.hidden=!!f;
      if(!f){sourceList.innerHTML='';if(sourceListSecondary)sourceListSecondary.innerHTML='';if(sourcePanelSecondary)sourcePanelSecondary.hidden=true;if(sourceSecondaryHint)sourceSecondaryHint.hidden=true;return;}
      const rows=getRows(f), days=dayNames(f), q=norm(sourceSearchSecondary?.value || sourceSearch?.value || '');
      const uniqueValues=(items,field)=>[...new Set(items.map(r=>String(r?.[field]??'').trim()).filter(Boolean))];
      const dayInfo=days.map(day=>{const dr=rows.filter(r=>norm(r.rampupDay)===norm(day));const titles=uniqueValues(dr,'title'),sections=uniqueValues(dr,'section'),styles=uniqueValues(dr,'styleName'),groups=uniqueValues(dr,'group');const combos=[...new Set(dr.map(r=>[r.title||'',r.styleName||'',r.section||'',r.group||''].join(' · ')).filter(x=>x.replace(/[ ·]/g,'').trim()))];return {day,count:dr.length,combos,sections,styles,titles,groups};}).filter(x=>[x.day,...x.combos,...x.titles,...x.sections,...x.styles,...x.groups].some(v=>norm(v).includes(q)));
      if(!dayInfo.length){sourceList.innerHTML='<div class="clone-empty">មិនមាន Day ត្រូវនឹងពាក្យស្វែងរកទេ។</div>';if(sourceListSecondary)sourceListSecondary.innerHTML=sourceList.innerHTML;return;}
      const cell=(values)=>`${esc(values.slice(0,2).join(', ')||'—')}${values.length>2?` <small>+${values.length-2}</small>`:''}`;
      const primaryHtml=`<div class="clone-source-table-wrap"><table class="clone-source-table"><thead><tr><th>Day</th><th>Factory</th><th>Style</th><th>Section</th><th>Process</th><th>ជ្រើស</th></tr></thead><tbody>${dayInfo.map(x=>{const chosen=norm(daySelect.value)===norm(x.day);return `<tr class="${chosen?'is-selected':''}"><td><strong>${esc(x.day)}</strong>${chosen?'<small class="clone-source-current">បានជ្រើស</small>':''}</td><td>${cell(x.titles)}</td><td>${cell(x.styles)}</td><td>${cell(x.sections)}</td><td><span class="clone-source-count">${x.count}</span></td><td><button type="button" class="clone-source-pick" data-pick-source="${esc(x.day)}">${chosen?'បានជ្រើស ✓':'ជ្រើសថ្ងៃនេះ'}</button></td></tr>`;}).join('')}</tbody></table></div>`;
      const secondaryHtml=`<div class="clone-source-table-wrap"><table class="clone-source-table"><thead><tr><th>Day</th><th>Factory</th><th>Style</th><th>Section</th><th>Process</th></tr></thead><tbody>${dayInfo.map(x=>{const chosen=norm(daySelect.value)===norm(x.day), edits=sourceEdits[norm(x.day)]||{};const editCell=(field,label,values)=>{const original=values.join(', ')||'';return `<td><input class="clone-source-edit-input" aria-label="${label} ${esc(x.day)}" data-source-edit-day="${esc(x.day)}" data-source-edit-field="${field}" value="${esc(edits[field]!==undefined?edits[field]:original)}" placeholder="${label}" ${chosen?'':'disabled'}></td>`;};return `<tr class="${chosen?'is-selected':''}">${editCell('rampupDay','Day',[x.day])}${editCell('title','Factory',x.titles)}${editCell('styleName','Style',x.styles)}${editCell('section','Section',x.sections)}<td><span class="clone-source-count">${x.count}</span></td></tr>`;}).join('')}</tbody></table></div>`;
      sourceList.innerHTML=primaryHtml;
      if(sourceListSecondary)sourceListSecondary.innerHTML=secondaryHtml;
      sourceList.querySelectorAll('[data-pick-source]').forEach(btn=>btn.addEventListener('click',()=>{
        const sameDay=norm(daySelect.value)===norm(btn.dataset.pickSource);
        if(sameDay&&sourcePanelSecondary&&!sourcePanelSecondary.hidden){sourcePanelSecondary.hidden=true;renderSourceDays();return;}
        daySelect.value=btn.dataset.pickSource;daySelect.dispatchEvent(new Event('change',{bubbles:true}));if(sourcePanelSecondary)sourcePanelSecondary.hidden=false;if(sourceSecondaryHint)sourceSecondaryHint.hidden=true;renderSourceDays();
      }));
      if(sourceListSecondary){
        sourceListSecondary.querySelectorAll('[data-source-edit-field]').forEach(input=>input.addEventListener('input',()=>{
          const key=norm(input.dataset.sourceEditDay);if(!sourceEdits[key])sourceEdits[key]={};sourceEdits[key][input.dataset.sourceEditField]=input.value;
        }));
      }
    };
    const savePreviewBtn=document.getElementById('clone-source-save-preview');
    if(savePreviewBtn)savePreviewBtn.addEventListener('click',()=>{previewPlan=buildPlan(true);if(previewPlan)renderPreview(previewPlan);});
    const refreshDays = () => { const files=loadFiles(); const f=fileSelect.value!==''?files[Number(fileSelect.value)]:null; if(sourcePanelPrimary)sourcePanelPrimary.hidden=!f; daySelect.innerHTML='<option value="">— ជ្រើសថ្ងៃដើម —</option>'; daySelect.disabled=!f; (f?dayNames(f):[]).forEach(d=>daySelect.add(new Option(d,d))); targetsInput.disabled=!f; previewBtn.disabled=true; previewPanel.hidden=true; reportPanel.hidden=true; previewPlan=null; if(sourceSearch)sourceSearch.value=''; if(sourceSearchSecondary)sourceSearchSecondary.value=''; if(sourcePanelSecondary)sourcePanelSecondary.hidden=true; if(sourceSecondaryHint)sourceSecondaryHint.hidden=true; renderSourceDays(); setStep(1); };
    const toast = (title,msg) => { if(typeof window.showCustomAlert==='function') window.showCustomAlert(title,msg,'fa-solid fa-circle-info','#2563eb'); else alert(title+'\n'+msg); };
    const buildPlan = (fromEditSave=false) => {
      const files=loadFiles(), idx=Number(fileSelect.value), file=files[idx], source=daySelect.value;
      const edits=sourceEdits[norm(source)]||{};
      let targets=[...new Set(targetsInput.value.split(',').map(s=>s.trim()).filter(Boolean))];
      if(!file||!source){toast('ត្រូវការជ្រើសថ្ងៃដើម','សូមជ្រើស File និង Source Day ជាមុន។');return null;}
      if(!targets.length && fromEditSave){
        const editedDays=edits.rampupDay===undefined?[]:[...new Set(String(edits.rampupDay).split(',').map(s=>s.trim()).filter(Boolean))];
        if(editedDays.length && editedDays.some(t=>norm(t)!==norm(source))) targets=editedDays;
        else {
          const numbered=String(source).match(/^(.*?)(\d+)$/);
          targets=[numbered?`${numbered[1]}${Number(numbered[2])+1}`:`${source}_Copy`];
          if(targetsInput)targetsInput.value=targets.join(', ');
        }
      }
      if(!targets.length){toast('ត្រូវការថ្ងៃគោលដៅ','សូមបញ្ចូល Target Day(s) ឬកែ Day ក្នុងប្រអប់ទី២ ជាមុន។');return null;}
      if(targets.some(t=>norm(t)===norm(source))){toast('ថ្ងៃគោលដៅមិនត្រឹមត្រូវ','Target Day មិនអាចដូច Source Day បានទេ។ សូមកែ Day ក្នុងប្រអប់ទី២ ឬ Target Day(s)។');return null;}
      const all=getRows(file), sourceRows=all.filter(r=>norm(r.rampupDay)===norm(source)).map(r=>({...r,rawRowData:r.rawRowData&&typeof r.rawRowData==='object'?{...r.rawRowData}:r.rawRowData}));
      if(!sourceRows.length){toast('គ្មានទិន្នន័យ','Source Day នេះមិនមាន Process សម្រាប់ Clone ទេ។');return null;}
      const fields=[['title','Factory','Title'],['styleName','Style','Style'],['section','Section','Section'],['rampupDay','Day','RampupDay']];
      fields.forEach(([prop,label,rawKey])=>{
        if(edits[prop]===undefined)return;
        const newValues=String(edits[prop]).split(',').map(v=>v.trim());
        const originalValues=[...new Set(sourceRows.map(r=>String(r[prop]??'').trim()))];
        sourceRows.forEach(row=>{
          const old=String(row[prop]??'').trim();
          const index=originalValues.indexOf(old);
          const next=newValues.length===1?newValues[0]:(index>=0&&newValues[index]!==undefined?newValues[index]:old);
          row[prop]=next;
          if(row.rawRowData&&typeof row.rawRowData==='object')row.rawRowData[rawKey]=next;
        });
      });
      const entries=[];
      targets.forEach(target=>{
        const targetRows=all.filter(r=>norm(r.rampupDay)===norm(target));
        const existing=new Set(targetRows.map(processIdentity));
        sourceRows.forEach((row,si)=>{ const key=processIdentity(row); const duplicate=existing.has(key); entries.push({target,row,duplicate,sourceIndex:si}); if(!duplicate) existing.add(key); });
      });
      return {fileIndex:idx,fileName:file.name,source,targets,sourceRows:sourceRows.map(r=>({...r})),entries,add:entries.filter(e=>!e.duplicate).length,skip:entries.filter(e=>e.duplicate).length,sourceCount:sourceRows.length};
    };
    const renderPreview = plan => {
      const targetSummary=plan.targets.map(t=>{const rows=plan.entries.filter(e=>e.target===t);return `<tr><td>${esc(t)}</td><td>${rows.length}</td><td class="clone-status-new">${rows.filter(e=>!e.duplicate).length}</td><td class="clone-status-skip">${rows.filter(e=>e.duplicate).length}</td></tr>`;}).join('');
      const sample=plan.entries.slice(0,60).map(e=>`<tr><td>${esc(e.target)}</td><td>${esc(e.row.section||'—')}</td><td>${esc(e.row.styleName||'—')}</td><td>${esc(e.row.prepAssembly||'—')}</td><td>${esc(e.row.processName||e.row.processRegion||'—')}</td><td>${e.duplicate?'<span class="clone-status-skip">ស្ទួន · រំលង</span>':'<span class="clone-status-new">ថ្មី · បន្ថែម</span>'}</td></tr>`).join('');
      const denied=!canEdit(loadFiles()[plan.fileIndex]);
      previewPanel.innerHTML=`<div class="clone-day-card-title"><i class="fa-solid fa-code-compare"></i><div><h3>3. Preview & Compare</h3><p>Source: <b>${esc(plan.source)}</b> · ${plan.sourceCount} source Process · ${plan.targets.length} target day(s)</p></div></div><div class="clone-preview-summary"><div class="clone-stat"><small>Process ត្រូវពិនិត្យ</small><strong>${plan.entries.length}</strong></div><div class="clone-stat is-new"><small>Process ថ្មីនឹងបន្ថែម</small><strong>${plan.add}</strong></div><div class="clone-stat is-skip"><small>ស្ទួននឹងរំលង</small><strong>${plan.skip}</strong></div></div><h4>សង្ខេបតាម Target Day</h4><div class="clone-preview-table-wrap"><table class="clone-preview-table"><thead><tr><th>Target Day</th><th>Checked</th><th>Add</th><th>Skip duplicate</th></tr></thead><tbody>${targetSummary}</tbody></table></div><h4 style="margin-top:20px">Process Preview ${plan.entries.length>60?'(បង្ហាញ 60 ជួរដំបូង)':''}</h4><div class="clone-preview-table-wrap"><table class="clone-preview-table"><thead><tr><th>Target Day</th><th>Section</th><th>Style</th><th>Prep/Assembly</th><th>Process</th><th>Result</th></tr></thead><tbody>${sample}</tbody></table></div>${denied?'<div class="clone-permission-hint"><i class="fa-solid fa-lock"></i> អ្នកអាចមើល Preview ប៉ុន្តែមិនមានសិទ្ធិ Clone File នេះទេ។ មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះ។</div>':''}<div class="clone-day-actions"><button type="button" id="clone-confirm-btn" class="clone-btn-primary" ${denied||plan.add===0?'disabled':''}><i class="fa-solid fa-check-double"></i> Confirm Clone (${plan.add})</button><button type="button" id="clone-back-btn" class="clone-btn-light">កែជម្រើស</button></div>`;
      previewPanel.hidden=false;reportPanel.hidden=true;setStep(3);
      document.getElementById('clone-back-btn').onclick=()=>{previewPanel.hidden=true;setStep(2);};
      document.getElementById('clone-confirm-btn').onclick=()=>{if(!previewPlan)return; const fresh=loadFiles()[previewPlan.fileIndex]; if(!canEdit(fresh)){toast('គ្មានសិទ្ធិ','មានតែម្ចាស់ File ឬ Admin ប៉ុណ្ណោះអាច Clone បាន។');return;} if(!window.confirm(`បញ្ជាក់ Clone ${previewPlan.add} Process ទៅ ${previewPlan.targets.length} Target Day(s)?
Process ស្ទួន ${previewPlan.skip} នឹងត្រូវរំលង។ ទិន្នន័យចាស់នឹងរក្សាទុក។`))return; executeClone(previewPlan);};
    };
    const executeClone = async plan => {
      // Work on a detached copy: loadFiles() returns the live IndexedDB cache by reference.
      // Mutating it before tspPersistLargeStorage compares old/new records makes both
      // snapshots identical, so the changed Clone record is never written to IndexedDB.
      const files=loadFiles().map(item => JSON.parse(JSON.stringify(item))), file=files[plan.fileIndex]; if(!file||!canEdit(file)){toast('គ្មានសិទ្ធិ','សូមចូលដោយម្ចាស់ File ឬ Admin។');return;}
      const sourceRows=plan.sourceRows.map(r=>({...r})); let added=0,skipped=0; const newIds=[]; const now=Date.now();
      // Keep the Time Observe hierarchy maps in sync with every confirmed Clone.
      // Add only missing values so existing Factory/Section/Style/Day tables are reused.
      file.titles=Array.isArray(file.titles)?file.titles:[];
      file.sectionsMap=file.sectionsMap||{};
      file.stylesMap=file.stylesMap||{};
      file.rampupMap=file.rampupMap||{};
      file.prepMap=file.prepMap||{};
      const addUnique=(arr,value)=>{if(!Array.isArray(arr))arr=[];if(value!=null&&String(value).trim()&&!arr.some(v=>norm(v)===norm(value)))arr.push(String(value).trim());return arr;};
      const ensureObserveHierarchy=(row,target)=>{
        const title=String(row.title||row.factory||'').trim();
        const section=String(row.section||'').trim();
        const style=String(row.styleName||'').trim();
        if(title)file.titles=addUnique(file.titles,title);
        if(title&&section)file.sectionsMap[title]=addUnique(file.sectionsMap[title],section);
        if(title&&section&&style){const styleKey=`${title}_${section}`;file.stylesMap[styleKey]=addUnique(file.stylesMap[styleKey],style);}
        if(title&&section&&style&&target){const rampupKey=`${title}_${section}_${style}`;file.rampupMap[rampupKey]=addUnique(file.rampupMap[rampupKey],target);}
      };
      plan.targets.forEach((target,ti)=>{
        sourceRows.forEach((row,ri)=>{
          const cloneBase={...row,rampupDay:target};
          ensureObserveHierarchy(cloneBase,target);
          const key=`${cloneBase.title||''}_${cloneBase.section||''}_${cloneBase.styleName||''}_${target}`;
          if(!Array.isArray(file.prepMap[key])) file.prepMap[key]=[];
          const targetRows=file.prepMap[key];
          if(targetRows.some(existing=>processIdentity(existing)===processIdentity(cloneBase))){skipped++;return;}
          const clone={...cloneBase,id:`CLONE_${now}_${ti}_${ri}_${Math.random().toString(36).slice(2,7)}`,clonedFromDay:plan.source,clonedAt:new Date().toISOString()};
          if(clone.rawRowData&&typeof clone.rawRowData==='object') clone.rawRowData={...clone.rawRowData,RampupDay:target,Day:target,Title:clone.title,Factory:clone.title,Section:clone.section,Style:clone.styleName};
          targetRows.push(clone);added++;newIds.push(clone.id);
        });
      });
      file.rowCount=(Number(file.rowCount)||0)+added;
      file.lastCloneAt=new Date().toISOString();file.lastCloneBy=(userNow()?.username)||'User';
      // Wait until IndexedDB/localStorage confirms the write before reporting success.
      // This prevents Refresh from racing ahead of the Clone save operation.
      const saved = await tspPersistLargeStorage('uploadedExcelFiles', files);
      if(!saved){toast('រក្សាទុកមិនបាន','Browser មិនអាចរក្សាទុក Clone បានទេ។ ទិន្នន័យមិនត្រូវបានបញ្ជាក់ថាបាន Clone រួចឡើយ។');return;}
      if(window.tspLargeStorageCache) window.tspLargeStorageCache['uploadedExcelFiles'] = files;
      if(typeof window.tspRecordRecentActivity==='function') window.tspRecordRecentActivity('Clone Day', file.name, `Source: ${plan.source}; Target: ${plan.targets.join(', ')}; Added: ${added}; Skipped: ${skipped}`);
      if(typeof window.dispatchEvent==='function')window.dispatchEvent(new CustomEvent('tsp:clone-day-complete',{detail:{fileName:file.name,source:plan.source,targets:plan.targets,added,skipped}}));
      setStep(5);previewPanel.hidden=true;reportPanel.hidden=false;
      reportPanel.innerHTML=`<div class="clone-report-success"><i class="fa-solid fa-circle-check"></i><div><h3>Clone Day បានបញ្ចប់</h3><p>បានរក្សាទុកការផ្លាស់ប្ដូរទៅក្នុង File: ${esc(file.name)}</p></div></div><div class="clone-preview-summary"><div class="clone-stat is-new"><small>បានបន្ថែម</small><strong>${added}</strong></div><div class="clone-stat is-skip"><small>ស្ទួន · រំលង</small><strong>${skipped}</strong></div><div class="clone-stat"><small>Target Day</small><strong>${plan.targets.length}</strong></div></div><p style="font-size:12px;color:#64748b;line-height:1.7">Source Day: <b>${esc(plan.source)}</b><br>Target Day(s): <b>${plan.targets.map(esc).join(', ')}</b><br>ទិន្នន័យដែលមានស្រាប់ត្រូវបានរក្សាទុក។ ${added===0?'មិនមាន Process ថ្មីត្រូវបន្ថែមទេ។':''}</p><div class="clone-report-actions"><button type="button" id="clone-new-run"><i class="fa-solid fa-plus"></i> Clone ម្តងទៀត</button><button type="button" id="clone-refresh-data"><i class="fa-solid fa-rotate"></i> Refresh File Data</button></div>`;
      document.getElementById('clone-new-run').onclick=()=>{refreshFiles();targetsInput.value='';previewPanel.hidden=true;reportPanel.hidden=true;setStep(1);};
      document.getElementById('clone-refresh-data').onclick=()=>{refreshFiles();toast('Refresh ជោគជ័យ','បានផ្ទុកបញ្ជី File និង Day ឡើងវិញ។');};
    };
    fileSelect.addEventListener('change',()=>{refreshDays();});
    daySelect.addEventListener('change',()=>{previewBtn.disabled=!(fileSelect.value!==''&&daySelect.value&&targetsInput.value.trim());if(daySelect.value)setStep(2);renderSourceDays();});
    targetsInput.addEventListener('input',()=>{previewBtn.disabled=!(fileSelect.value!==''&&daySelect.value&&targetsInput.value.trim());if(targetsInput.value.trim()&&daySelect.value)setStep(2);});
    if(sourceSearch)sourceSearch.addEventListener('input',()=>{if(sourceSearchSecondary)sourceSearchSecondary.value=sourceSearch.value;renderSourceDays();});
    if(sourceSearchSecondary)sourceSearchSecondary.addEventListener('input',()=>{if(sourceSearch)sourceSearch.value=sourceSearchSecondary.value;renderSourceDays();});
    previewBtn.addEventListener('click',()=>{previewPlan=buildPlan();if(previewPlan)renderPreview(previewPlan);});
    resetBtn.addEventListener('click',()=>{fileSelect.value='';targetsInput.value='';refreshDays();});
    document.querySelectorAll('[data-page="clone-page"]').forEach(link=>link.addEventListener('click',()=>{refreshFiles();}));
    refreshFiles();
  };
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();


/* V34: Prevent multi-touch pinch gestures where the browser allows it.
   Single-finger scrolling and ordinary taps are intentionally left enabled. */
(function preventPinchZoomOnly() {
    'use strict';
    document.addEventListener('touchmove', function (event) {
        if (event.touches && event.touches.length > 1 && event.cancelable) {
            event.preventDefault();
        }
    }, { passive: false });
    ['gesturestart', 'gesturechange', 'gestureend'].forEach(function (type) {
        document.addEventListener(type, function (event) {
            if (event.cancelable) event.preventDefault();
        }, { passive: false });
    });
})();
