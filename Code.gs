/**
 * JADWAL TOKO — Google Apps Script (database di Google Sheets + backup + alarm email).
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 * Jalankan setup() sekali, lalu installTriggers() sekali. Salin API_SECRET (lihat Logs) ke env Vercel (GS_SECRET).
 * Atur zona waktu: Project Settings > Time zone = Asia/Jakarta.
 *
 * UPGRADE dari versi lama: tempel kode ini, jalankan setup() + installTriggers() lagi, lalu
 * Deploy > Manage deployments > (pensil) > Version: New version > Deploy. Data lama tetap aman:
 * kolom baru hanya DITAMBAHKAN di sebelah kanan sheet yang sudah ada.
 */
const SS = SpreadsheetApp.getActive();
const TZ = Session.getScriptTimeZone();
const P = PropertiesService.getScriptProperties();
const HARI = ['MINGGU','SENIN','SELASA','RABU','KAMIS','JUMAT','SABTU'];
const DEF_BRANCH = ['BANDUNG 1','BANDUNG 2','DC CIANJUR','KARAWANG','PLUMBON','TEGAL','CILACAP'];
const DEF_SET = { DC_NAME: 'SS08 - WH FRESH PADALARANG [WFPDG]', ALARM_EMAIL: 'Fijarfresh@gmail.com' };
const STALE_DAYS = 4;                       // toko dianggap terlambat jika > 4 hari belum ada SL padahal ada jadwal
const RE_DATE = /^\d{4}-\d{2}-\d{2}$/, RE_PER = /^\d{4}-\d{2}$/;
const HDR = {
  TOKO: ['KODE','INS','NAMA','BRANCH','GROUP','STATUS','TUTUP_MULAI','GO_MULAI','GROUP_LAMA','GROUP_BERLAKU','KODE_LAMA','KET','UPDATED','RIT'],
  GROUP: ['GROUP','BRANCH','H1','H2','H3'],
  BLACKOUT: ['TANGGAL','SCOPE','NILAI','KET','OLEH','MODE','GESER_HARI','JML_TOKO','WAKTU','TGL_AKHIR'],
  PLAN_BASE: ['PERIODE','KODE','PLAN'],
  HISTORY: ['WAKTU','USER','AKSI','KODE','DETAIL'],
  UPLOAD: ['WAKTU','USER','FILE','BARIS','TGL_MIN','TGL_MAX'],
  TARGET: ['PERIODE','SCOPE','NILAI','TARGET','OLEH','WAKTU'],
  SETTING: ['KEY','VALUE'],
  BRANCH: ['BRANCH'],
  SL: ['TANGGAL','KODE','NAMA','BRANCH','QTY_PB','QTY_SHIP','RP_PB','QTY_PICK','RP_PICK','RP_SHIP','SL_QTY','TOT_ITEM']
};
const CHK = {};                             // cek/migrasi header cukup sekali per eksekusi

// ---------- setup ----------
function setup() {
  ['TOKO','GROUP','BLACKOUT','PLAN_BASE','HISTORY','UPLOAD','TARGET','SETTING','BRANCH'].forEach(sh_);
  if (!branchSheetRows_().length) put_(sh_('BRANCH'), DEF_BRANCH.map(b => [b]));
  Object.keys(DEF_SET).forEach(k => { if (settingAll_()[k] === undefined) put_(sh_('SETTING'), [[k, DEF_SET[k]]]); });
  if (!P.getProperty('API_SECRET')) P.setProperty('API_SECRET', Utilities.getUuid() + Utilities.getUuid());
  Logger.log('API_SECRET = ' + P.getProperty('API_SECRET'));
}
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();
  ScriptApp.newTrigger('dailyAlarm').timeBased().everyDays(1).atHour(6).create();
}
function dailyBackup() { doBackup_('harian'); }
function dailyAlarm() { sendAlarm_(false); }

// ---------- web app ----------
function doGet() { return ContentService.createTextOutput('JADWAL TOKO API OK'); }
function doPost(e) {
  let out;
  try {
    const b = JSON.parse(e.postData.contents);
    if (!b.secret || b.secret !== P.getProperty('API_SECRET')) throw new Error('unauthorized');
    out = { ok: true, data: route_(b.action, b.payload || {}, b.user || '?') };
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
const WRITE = ['tokoSave','setStatus','pindahGroup','gantiKode','blackout','groupSave','groupDelete','targetSave','settingSave','slBegin','slAppend','slEnd','imporBegin','imporToko','imporEnd'];
function route_(a, p, u) {
  if (a === 'ver') return { ver: P.getProperty('VER') || '0' };
  if (a === 'state') return state_(p.periode);
  if (a === 'history') return history_(p);
  if (a === 'backups') return backups_();
  if (a === 'backupNow') return { name: doBackup_('manual') };
  if (a === 'alarmEmail') return sendAlarm_(true);
  const fn = ACT[a];
  if (!fn) throw new Error('aksi tidak dikenal: ' + a);
  if (WRITE.indexOf(a) < 0) return fn(p, u);
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  try { const r = fn(p, u); P.setProperty('VER', String(Date.now())); return r; } finally { lock.releaseLock(); }
}

// ---------- helpers ----------
function hdrOf_(n) { return n.indexOf('SL_') === 0 ? HDR.SL : HDR[n]; }
function sh_(n) {
  const h = hdrOf_(n); let s = SS.getSheetByName(n);
  if (!s) {
    s = SS.insertSheet(n);
    s.getRange(1, 1, 1, h.length).setNumberFormat('@').setValues([h]).setFontWeight('bold');
    s.setFrozenRows(1); CHK[n] = 1; return s;
  }
  if (!CHK[n]) {                              // migrasi: tambah kolom header baru di kanan, data lama tidak disentuh
    if (s.getMaxColumns() < h.length) s.insertColumnsAfter(s.getMaxColumns(), h.length - s.getMaxColumns());
    const cur = s.getRange(1, 1, 1, h.length).getValues()[0];
    cur.forEach((v, i) => { if (v === '' || v == null) s.getRange(1, i + 1).setNumberFormat('@').setValue(h[i]).setFontWeight('bold'); });
    CHK[n] = 1;
  }
  return s;
}
function ds_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : (v == null || v === '' ? '' : String(v).slice(0, 10)); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function dnum_(s) { return Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) / 864e5; }
function diff_(a, b) { return Math.round(dnum_(b) - dnum_(a)); }      // b - a (hari)
function put_(s, rows) {            // tulis sebagai teks (anti auto-format tanggal/angka)
  if (!rows.length) return;
  const start = s.getLastRow() + 1, need = start + rows.length - 1;
  if (need > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), need - s.getMaxRows());
  const r = s.getRange(start, 1, rows.length, rows[0].length);
  r.setNumberFormat('@'); r.setValues(rows);
}
function clearData_(s) { const n = s.getLastRow(); if (n > 1) s.getRange(2, 1, n - 1, s.getMaxColumns()).clearContent(); }
function readAll_(name) { const s = sh_(name), n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, hdrOf_(name).length).getValues(); }
function hist_(u, aksi, kode, detail) { put_(sh_('HISTORY'), [[now_(), u, aksi, kode || '', typeof detail === 'string' ? detail : JSON.stringify(detail || {})]]); }
function chkDate_(v, label) { if (!RE_DATE.test(String(v || ''))) throw new Error((label || 'Tanggal') + ' tidak valid (format yyyy-mm-dd)'); return String(v); }
function num_(v) { const n = Number(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? n : 0; }

function branchSheetRows_() { return readAll_('BRANCH').map(r => String(r[0]).trim()).filter(Boolean); }
function settingAll_() { const o = {}; readAll_('SETTING').forEach(r => { if (r[0] !== '') o[String(r[0])] = String(r[1]); }); return o; }
function settingGet_(k) { const v = settingAll_()[k]; return v === undefined || v === '' ? DEF_SET[k] : v; }
function branchAll_(toko, groups) {
  const set = {}; const add = b => { b = String(b || '').trim().toUpperCase(); if (b) set[b] = 1; };
  const fromSheet = branchSheetRows_(); (fromSheet.length ? fromSheet : DEF_BRANCH).forEach(add);
  groups.forEach(g => add(g.branch)); toko.forEach(t => add(t.branch));
  const all = Object.keys(set), ord = b => { const i = DEF_BRANCH.indexOf(b); return i < 0 ? 99 : i; };
  return all.sort((a, b) => ord(a) - ord(b) || a.localeCompare(b));
}

function tokoAll_() {
  return readAll_('TOKO').filter(r => r[0] !== '').map(r => ({
    kode: String(r[0]), ins: String(r[1]), nama: String(r[2]), branch: String(r[3]), group: String(r[4]),
    status: String(r[5] || 'AKTIF'), tutup: ds_(r[6]), go: ds_(r[7]), groupLama: String(r[8] || ''),
    berlaku: ds_(r[9]), kodeLama: String(r[10] || ''), ket: String(r[11] || ''), rit: String(r[13] || '1')
  }));
}
function tokoRow_(t) { return [t.kode, t.ins, t.nama, t.branch, t.group, t.status, t.tutup, t.go, t.groupLama, t.berlaku, t.kodeLama, t.ket, now_(), t.rit || '1']; }
function rowOf_(kode) {
  const s = sh_('TOKO'), n = s.getLastRow(); if (n < 2) return -1;
  const i = s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])).indexOf(String(kode));
  return i < 0 ? -1 : i + 2;
}
function getToko_(kode) {
  const row = rowOf_(kode); if (row < 0) return null;
  const all = tokoAll_(); return all.filter(t => t.kode === String(kode))[0] || null;
}
function saveToko_(t) {
  const row = rowOf_(t.kode), s = sh_('TOKO');
  if (row < 0) put_(s, [tokoRow_(t)]); else { const r = s.getRange(row, 1, 1, HDR.TOKO.length); r.setNumberFormat('@'); r.setValues([tokoRow_(t)]); }
}
function groupAll_() { return readAll_('GROUP').filter(r => r[0] !== '').map(r => ({ group: String(r[0]), branch: String(r[1]), h1: String(r[2]), h2: String(r[3]), h3: String(r[4]) })); }
function blackAll_() {
  return readAll_('BLACKOUT').filter(r => r[0] !== '').map(r => ({
    tgl: ds_(r[0]), scope: String(r[1]), nilai: String(r[2]), ket: String(r[3]), oleh: String(r[4]),
    mode: String(r[5] || 'KOSONG'), geser: Number(r[6]) || 0, jml: Number(r[7]) || 0, waktu: String(r[8] || ''), tglAkhir: ds_(r[9]) || ds_(r[0])
  }));
}
function baseMap_(per) { const m = {}; readAll_('PLAN_BASE').forEach(r => { if (String(r[0]) === per) m[String(r[1])] = String(r[2]); }); return m; }
function targetAll_(per) {
  const o = { total: 0, branch: {} };
  readAll_('TARGET').forEach(r => {
    if (String(r[0]) !== per) return;
    const v = num_(r[3]); if (String(r[1]) === 'ALL') o.total = v; else if (String(r[1]) === 'BRANCH') o.branch[String(r[2])] = v;
  });
  return o;
}

// aturan aktif: GO belum mulai / sudah TUTUP; jika GO > TUTUP berarti RE-GO (tutup sementara lalu buka lagi)
function aktif_(t, d) {
  if (t.tutup && t.go && t.go > t.tutup) return !(d >= t.tutup && d < t.go);
  if (t.go && d < t.go) return false;
  if (t.tutup && d >= t.tutup) return false;
  return true;
}
function bmatch_(b, t, g) { return b.scope === 'ALL' || (b.scope === 'BRANCH' && b.nilai === t.branch) || (b.scope === 'GROUP' && b.nilai === g); }

// ---------- plan & aktual ----------
// Tahap 1: jadwal awal (aturan group / plan dasar file + status tutup/GO). Index 0 = hari terakhir bulan lalu (kolom carry).
function datesOf_(y, m, nd) {
  const dates = [], dows = [];
  for (let i = 0; i <= 31; i++) {
    if (i > nd) { dates.push(null); dows.push(0); continue; }
    const dt = new Date(y, m - 1, i, 12);
    dates.push(Utilities.formatDate(dt, TZ, 'yyyy-MM-dd')); dows.push(dt.getDay());
  }
  return { dates: dates, dows: dows, nd: nd };
}
function planRaw_(dd, toko, gmap, base) {
  const out = {};
  toko.forEach(t => {
    const bs = base[t.kode], a = new Array(32).fill(0);
    for (let i = 0; i <= 31; i++) {
      const ds = dd.dates[i]; if (!ds) continue;
      const g = (t.berlaku && t.groupLama && ds < t.berlaku) ? t.groupLama : t.group;
      let p;
      if (bs && !(t.berlaku && ds >= t.berlaku)) p = bs.charAt(i) === '1';
      else { const G = gmap[g]; p = !!G && [G.h1, G.h2, G.h3].indexOf(HARI[dd.dows[i]]) >= 0; }
      if (p && !aktif_(t, ds)) p = false;
      a[i] = p ? 1 : 0;
    }
    out[t.kode] = a;
  });
  return out;
}
// Tahap 2: terapkan Libur/SO untuk rentang tanggal. KOSONG = plan dikosongkan; GESER = plan tiap tanggal dalam rentang
// dipindah N hari ke depan (sama seperti "jadwal geser" di file JPF). Pergeseran yang melewati akhir bulan dibuang.
function planApply_(dd, raw, toko, bo) {
  const list = [];
  bo.forEach(b => { const idx = []; dd.dates.forEach((d, i) => { if (d && d >= b.tgl && d <= (b.tglAkhir || b.tgl)) idx.push(i); }); if (idx.length) list.push({ b: b, idx: idx }); });
  const out = {}, awal = {};
  toko.forEach(t => {
    const r = raw[t.kode], p = r.slice();
    if (list.length) {
      const gOf = ds => (t.berlaku && t.groupLama && ds < t.berlaku) ? t.groupLama : t.group, shifts = [];
      list.forEach(e => e.idx.forEach(i => {
        if (!r[i] || !bmatch_(e.b, t, gOf(dd.dates[i]))) return;
        p[i] = 0;
        if (e.b.mode === 'GESER' && i >= 1) { const j = i + (e.b.geser || 1); if (j <= dd.nd && aktif_(t, dd.dates[j])) shifts.push(j); }
      }));
      shifts.forEach(j => { p[j] = 1; });
      list.forEach(e => { if (e.b.mode !== 'GESER') e.idx.forEach(i => { if (bmatch_(e.b, t, gOf(dd.dates[i]))) p[i] = 0; }); });   // KOSONG menang atas hasil geser
    }
    out[t.kode] = p.join('');
    if (list.length && out[t.kode] !== r.join('')) awal[t.kode] = r.join('');
  });
  return { plan: out, awal: awal };
}
function buildAct_(per, nd) {
  const s = SS.getSheetByName('SL_' + per), act = {}, sl = {}; let last = '';
  if (!s || s.getLastRow() < 2) return { act: act, sl: sl, last: last };
  sh_('SL_' + per);
  s.getRange(2, 1, s.getLastRow() - 1, HDR.SL.length).getValues().forEach(r => {
    const d = ds_(r[0]); if (d.slice(0, 7) !== per) return;
    const day = +d.slice(8, 10), k = String(r[1]), ship = num_(r[5]), rps = num_(r[9]);
    if (d > last) last = d;
    const x = sl[day] = sl[day] || { toko: 0, pb: 0, ship: 0, rpPb: 0, rpShip: 0 };
    x.pb += num_(r[4]); x.ship += ship; x.rpPb += num_(r[6]); x.rpShip += rps;
    if (ship <= 0 && rps <= 0) return;                 // baris tanpa kirim tidak dihitung sebagai aktual
    if (!act[k]) act[k] = new Array(32).fill('0');
    if (act[k][day] !== '1') { act[k][day] = '1'; x.toko++; }
  });
  Object.keys(act).forEach(k => act[k] = act[k].join(''));
  return { act: act, sl: sl, last: last };
}
function monthData_(per, toko, groups, bo) {
  const y = +per.slice(0, 4), m = +per.slice(5), nd = new Date(y, m, 0).getDate();
  const base = baseMap_(per), gmap = {}; groups.forEach(g => gmap[g.group] = g);
  const dd = datesOf_(y, m, nd), raw = planRaw_(dd, toko, gmap, base), pa = planApply_(dd, raw, toko, bo), A = buildAct_(per, nd);
  return { per: per, y: y, m: m, nd: nd, dd: dd, raw: raw, plan: pa.plan, awal: pa.awal, act: A.act, sl: A.sl, last: A.last, hasBase: Object.keys(base).length > 0 };
}
function state_(per) {
  const cur = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  per = RE_PER.test(per || '') ? per : cur;
  const toko = tokoAll_(), groups = groupAll_(), bo = blackAll_();
  const md = monthData_(per, toko, groups, bo), cmd = per === cur ? md : monthData_(cur, toko, groups, bo);
  return { ver: P.getProperty('VER') || '0', periode: per, nd: md.nd, today: today_(), toko: toko, groups: groups, blackout: bo,
    branches: branchAll_(toko, groups), plan: md.plan, awal: md.awal, act: md.act, sl: md.sl, lastSL: md.last, hasBase: md.hasBase,
    target: targetAll_(per), setting: { DC_NAME: settingGet_('DC_NAME'), ALARM_EMAIL: settingGet_('ALARM_EMAIL') }, alarm: alarms_(toko, cmd) };
}
function history_(p) {
  const s = sh_('HISTORY'), n = s.getLastRow(); if (n < 2) return [];
  const from = Math.max(2, n - 2999);
  let v = s.getRange(from, 1, n - from + 1, 5).getValues().map(r => ({ waktu: String(r[0]), user: r[1], aksi: String(r[2]), kode: String(r[3]), detail: r[4] }));
  if (p.kode) v = v.filter(x => x.kode === String(p.kode));
  if (p.aksi) v = v.filter(x => x.aksi.indexOf(String(p.aksi)) === 0);
  return v.reverse().slice(0, p.limit || 400);
}

// ---------- alarm (dipakai dashboard, menu Tutup/GO, dan email harian) ----------
// level: crit (merah) / warn (kuning) / info (abu). Selalu dihitung untuk BULAN BERJALAN.
function alarms_(toko, md) {
  const T = today_(), cur = md.per, out = { today: T, cut: 0, lastSL: md.last || '', global: [], list: [] };
  const tIdx = T.slice(0, 7) === cur ? +T.slice(8, 10) : 0;
  const cut = md.last && md.last.slice(0, 7) === cur ? Math.min(+md.last.slice(8, 10), tIdx || md.nd) : 0;
  out.cut = cut;
  if (!md.last || diff_(md.last, T) >= 2) out.global.push({ type: 'SL_LAMA', level: md.last && diff_(md.last, T) < 4 ? 'warn' : 'crit', hari: md.last ? diff_(md.last, T) : null,
    msg: md.last ? 'SL terakhir di-upload tanggal ' + md.last + ' (' + diff_(md.last, T) + ' hari lalu)' : 'Belum ada data SL yang di-upload' });
  const add = (t, type, level, hari, msg) => out.list.push({ kode: t.kode, nama: t.nama, branch: t.branch, group: t.group, type: type, level: level, hari: hari, msg: msg });
  toko.forEach(t => {
    const pl = md.plan[t.kode] || '', ac = md.act[t.kode] || '';
    if (t.go) {
      const d = diff_(T, t.go);
      if (d === 0) add(t, 'GO_HARI_INI', 'crit', 0, 'GO HARI INI (' + t.go + ')');
      else if (d > 0 && d <= 3) add(t, 'GO_SEGERA', 'warn', d, 'GO H-' + d + ' (' + t.go + ')');
      else if (d > 3 && d <= 14) add(t, 'GO_AKAN', 'info', d, 'GO ' + d + ' hari lagi (' + t.go + ')');
      else if (d < 0 && d >= -14 && T.slice(0, 7) === cur) {
        let fp = 0; for (let i = 1; i <= tIdx; i++) { if (pl[i] === '1' && md.dd.dates[i] >= t.go) { fp = i; break; } }
        if (fp && cut >= fp) {
          let got = false; for (let i = fp; i <= md.nd; i++) if (ac[i] === '1') { got = true; break; }
          if (!got) add(t, 'GO_BELUM_SL', (cut - fp) >= 3 ? 'crit' : 'warn', cut - fp, 'GO ' + t.go + ' sudah jalan, belum ada SL sejak jadwal pertama tgl ' + fp);
        }
      }
    }
    if (t.tutup) {
      const d = diff_(T, t.tutup);
      if (d === 0) add(t, 'TUTUP_HARI_INI', 'crit', 0, 'TUTUP HARI INI (' + t.tutup + ')');
      else if (d > 0 && d <= 3) add(t, 'TUTUP_SEGERA', 'warn', d, 'Tutup H-' + d + ' (' + t.tutup + ')');
      else if (d > 3 && d <= 14) add(t, 'TUTUP_AKAN', 'info', d, 'Tutup ' + d + ' hari lagi (' + t.tutup + ')');
      else if (d < 0 && !aktif_(t, T) && T.slice(0, 7) === cur) {
        let bad = 0; for (let i = 1; i <= md.nd; i++) if (ac[i] === '1' && md.dd.dates[i] >= t.tutup) bad++;
        if (bad) add(t, 'TUTUP_ADA_SL', 'crit', bad, 'Sudah tutup ' + t.tutup + ' tetapi masih ada SL (' + bad + ' hari) — periksa');
      }
    }
    if (cut > 0 && aktif_(t, md.dd.dates[cut])) {       // toko > 4 hari belum kirim
      let la = 0; for (let i = cut; i >= 1; i--) if (ac[i] === '1') { la = i; break; }
      let missed = false; for (let i = la + 1; i <= cut; i++) if (pl[i] === '1') { missed = true; break; }
      const gap = cut - la;
      if (missed && gap > STALE_DAYS) add(t, 'BELUM_KIRIM', gap > 7 ? 'crit' : 'warn', gap, la ? gap + ' hari sejak SL terakhir (tgl ' + la + ')' : gap + ' hari belum ada SL bulan ini');
    }
  });
  const rank = { crit: 0, warn: 1, info: 2 };
  out.list.sort((a, b) => rank[a.level] - rank[b.level] || a.hari - b.hari || a.kode.localeCompare(b.kode));
  return out;
}
function sendAlarm_(force) {
  const toko = tokoAll_(), groups = groupAll_(), bo = blackAll_();
  const al = alarms_(toko, monthData_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), toko, groups, bo));
  const useful = al.list.filter(x => x.level !== 'info'), to = settingGet_('ALARM_EMAIL');
  if (!force && !useful.length && !al.global.length) return { sent: false, n: 0, to: to };
  const sec = (title, types) => {
    const l = useful.filter(x => types.indexOf(x.type) >= 0); if (!l.length) return '';
    return '<h3>' + title + ' (' + l.length + ')</h3><ul>' + l.slice(0, 60).map(x => '<li><b>' + x.kode + '</b> ' + x.nama + ' — ' + x.branch + ' / ' + x.group + ': ' + x.msg + '</li>').join('') + (l.length > 60 ? '<li>... dan ' + (l.length - 60) + ' lainnya (lihat web)</li>' : '') + '</ul>';
  };
  const html = '<div style="font-family:Arial,sans-serif;font-size:13px"><h2>Alarm Jadwal Toko — ' + al.today + '</h2>' +
    (al.global.length ? '<p style="color:#b45309"><b>' + al.global.map(g => g.msg).join('; ') + '</b></p>' : '') +
    sec('Toko GO / Grand Opening', ['GO_HARI_INI', 'GO_SEGERA']) + sec('GO sudah jalan tapi belum ada SL', ['GO_BELUM_SL']) +
    sec('Toko TUTUP', ['TUTUP_HARI_INI', 'TUTUP_SEGERA']) + sec('Tutup tapi masih ada SL', ['TUTUP_ADA_SL']) +
    sec('Toko > ' + STALE_DAYS + ' hari belum kirim', ['BELUM_KIRIM']) +
    (!useful.length ? '<p>Tidak ada alarm toko GO / tutup hari ini.</p>' : '') + '</div>';
  MailApp.sendEmail({ to: to, subject: 'Alarm Jadwal Toko ' + al.today + ' (' + useful.length + ' toko)', htmlBody: html });
  return { sent: true, n: useful.length, to: to };
}

// ---------- backup ----------
function doBackup_(why) {
  let f; try { f = DriveApp.getFolderById(P.getProperty('BACKUP_FOLDER_ID')); } catch (e) { f = DriveApp.createFolder('JADWAL_TOKO_BACKUP'); P.setProperty('BACKUP_FOLDER_ID', f.getId()); }
  const name = 'BACKUP_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm') + '_' + why;
  DriveApp.getFileById(SS.getId()).makeCopy(name, f);
  const files = []; const it = f.getFiles(); while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()); files.slice(60).forEach(x => x.setTrashed(true));
  P.setProperty('LAST_SNAP', String(Date.now()));
  return name;
}
function snap_(why) { if (Date.now() - Number(P.getProperty('LAST_SNAP') || 0) > 10 * 60 * 1000) doBackup_(why); }   // snapshot otomatis sebelum operasi besar
function backups_() {
  try { const f = DriveApp.getFolderById(P.getProperty('BACKUP_FOLDER_ID')), out = [], it = f.getFiles();
    while (it.hasNext()) { const x = it.next(); out.push({ name: x.getName(), url: x.getUrl(), at: Utilities.formatDate(x.getDateCreated(), TZ, 'yyyy-MM-dd HH:mm') }); }
    return out.sort((a, b) => b.at.localeCompare(a.at)).slice(0, 60);
  } catch (e) { return []; }
}

// ---------- Libur / SO: hitung dampak ----------
function planOf_(per, toko, groups, bo) {
  const y = +per.slice(0, 4), m = +per.slice(5), nd = new Date(y, m, 0).getDate(), gmap = {}; groups.forEach(g => gmap[g.group] = g);
  const dd = datesOf_(y, m, nd);
  return { dd: dd, plan: planApply_(dd, planRaw_(dd, toko, gmap, baseMap_(per)), toko, bo).plan };
}
function changed_(a, b) { let n = 0; Object.keys(a).forEach(k => { if (a[k] !== b[k]) n++; }); return n; }
function blackRows_(list) { return list.map(b => [b.tgl, b.scope, b.nilai, b.ket, b.oleh, b.mode, b.geser, b.jml, b.waktu, b.tglAkhir].map(String)); }
function rewriteBlack_(list) { const s = sh_('BLACKOUT'); clearData_(s); put_(s, blackRows_(list)); }

// ---------- aksi tulis ----------
const ACT = {
  tokoSave(p, u) {                       // tambah toko baru (GO) atau ubah info dasar
    const kode = String(p.kode || '').trim().toUpperCase(); if (!kode) throw new Error('Kode toko wajib');
    if (!/^[A-Z0-9._-]{2,12}$/.test(kode)) throw new Error('Kode toko hanya huruf/angka (2-12 karakter)');
    const old = getToko_(kode);
    const rit = p.rit != null && p.rit !== '' ? String(p.rit) : null;
    if (rit !== null && !/^[1-9]$/.test(rit)) throw new Error('Rit harus angka 1-9');
    if (old) {
      if (p.baru) throw new Error('Kode toko ' + kode + ' sudah terdaftar (group ' + old.group + ')');
      const t = Object.assign({}, old, { ins: p.ins || old.ins, nama: p.nama || old.nama, branch: p.branch || old.branch, ket: p.ket != null ? p.ket : old.ket, rit: rit || old.rit });
      saveToko_(t); hist_(u, 'UBAH_INFO', kode, { dari: old, ke: t }); return { kode: kode };
    }
    const g = groupAll_().filter(x => x.group === String(p.group || ''))[0];
    if (!g) throw new Error('Group "' + (p.group || '') + '" tidak ada di master group');
    const branch = String(p.branch || g.branch || '').trim().toUpperCase(); if (!branch) throw new Error('Branch wajib untuk toko baru');
    if (p.tgl) chkDate_(p.tgl, 'Tanggal GO');
    const t = { kode: kode, ins: p.ins || '', nama: p.nama || '', branch: branch, group: g.group, status: p.tgl ? 'GO' : 'AKTIF',
      tutup: '', go: p.tgl || '', groupLama: '', berlaku: '', kodeLama: '', ket: p.ket || '', rit: rit || '1' };
    saveToko_(t); hist_(u, 'TOKO_BARU', kode, t);
    return { kode: kode, group: g.group, branch: branch, hari: [g.h1, g.h2, g.h3].filter(Boolean), go: t.go };
  },
  setStatus(p, u) {                      // TUTUP / GO / AKTIF(batalkan)
    const t = getToko_(p.kode); if (!t) throw new Error('Toko tidak ditemukan');
    if (!/^(TUTUP|GO|AKTIF)$/.test(p.status)) throw new Error('Status tidak valid');
    if (p.status !== 'AKTIF') chkDate_(p.tgl, 'Tanggal ' + p.status);
    const before = { status: t.status, tutup: t.tutup, go: t.go };
    if (p.status === 'TUTUP') { t.tutup = p.tgl; if (t.go && t.go > p.tgl) t.go = ''; }
    else if (p.status === 'GO') { t.go = p.tgl; if (t.tutup && t.tutup > p.tgl) t.tutup = ''; }
    else { t.tutup = ''; t.go = ''; }
    t.status = p.status; if (p.ket) t.ket = p.ket;
    saveToko_(t); hist_(u, 'STATUS_' + p.status, t.kode, { sebelum: before, tgl: p.tgl, ket: p.ket || '' }); return {};
  },
  pindahGroup(p, u) {
    const t = getToko_(p.kode); if (!t) throw new Error('Toko tidak ditemukan');
    if (!p.groupBaru || !p.berlaku) throw new Error('Group baru dan tanggal berlaku wajib');
    chkDate_(p.berlaku, 'Tanggal berlaku');
    const g = groupAll_().filter(x => x.group === p.groupBaru)[0]; if (!g) throw new Error('Group ' + p.groupBaru + ' tidak ada');
    const lama = (t.berlaku && t.groupLama && p.berlaku < t.berlaku) ? t.groupLama : t.group;
    t.groupLama = lama; t.group = p.groupBaru; t.berlaku = p.berlaku;
    if (g.branch) t.branch = g.branch;
    saveToko_(t); hist_(u, 'PINDAH_GROUP', t.kode, { dari: lama, ke: p.groupBaru, berlaku: p.berlaku }); return {};
  },
  gantiKode(p, u) {                      // toko lama TUTUP (take over), toko baru GO dengan group sama
    const o = getToko_(p.kodeLama); if (!o) throw new Error('Kode lama tidak ditemukan');
    const nk = String(p.kodeBaru || '').trim().toUpperCase(); if (!nk || !p.tgl) throw new Error('Kode baru dan tanggal wajib');
    if (!/^[A-Z0-9._-]{2,12}$/.test(nk)) throw new Error('Kode baru hanya huruf/angka (2-12 karakter)');
    chkDate_(p.tgl, 'Tanggal');
    if (getToko_(nk)) throw new Error('Kode baru sudah ada');
    const n = { kode: nk, ins: p.ins || o.ins, nama: p.nama || o.nama, branch: o.branch, group: o.group, status: 'GO', tutup: '', go: p.tgl,
      groupLama: o.groupLama, berlaku: o.berlaku, kodeLama: o.kode, ket: 'TAKE OVER DARI ' + o.kode, rit: o.rit };
    o.tutup = p.tgl; o.status = 'TUTUP'; o.ket = 'TAKE OVER KE ' + nk;
    saveToko_(o); saveToko_(n); hist_(u, 'GANTI_KODE', o.kode, { baru: nk, tgl: p.tgl }); hist_(u, 'GANTI_KODE', nk, { lama: o.kode, tgl: p.tgl }); return {};
  },
  blackout(p, u) {                       // hari libur / stock opname: ALL | BRANCH | GROUP ; mode KOSONG | GESER
    const toko = tokoAll_(), groups = groupAll_(), all = blackAll_();
    if (p.op === 'del') {
      const hit = all.filter(b => b.tgl === p.tgl && b.scope === p.scope && b.nilai === (p.nilai || ''));
      if (!hit.length) throw new Error('Data libur/SO tidak ditemukan (mungkin sudah dihapus)');
      const rest = all.filter(b => hit.indexOf(b) < 0), per = p.tgl.slice(0, 7);
      const kembali = changed_(planOf_(per, toko, groups, all).plan, planOf_(per, toko, groups, rest).plan);
      rewriteBlack_(rest); hist_(u, 'LIBUR_HAPUS', '', { tgl: p.tgl, tglAkhir: hit[0].tglAkhir, scope: p.scope, nilai: p.nilai || '', mode: hit[0].mode, geser: hit[0].geser, jmlKembali: kembali }); return { kembali: kembali };
    }
    if (p.op === 'reset') {              // kembalikan SEMUA libur/SO pada satu periode ke jadwal awal
      if (!RE_PER.test(p.periode || '')) throw new Error('Periode tidak valid');
      const hit = all.filter(b => b.tgl.slice(0, 7) === p.periode), rest = all.filter(b => b.tgl.slice(0, 7) !== p.periode);
      if (!hit.length) throw new Error('Tidak ada libur/SO pada periode ' + p.periode);
      const kembali = changed_(planOf_(p.periode, toko, groups, all).plan, planOf_(p.periode, toko, groups, rest).plan);
      rewriteBlack_(rest); hist_(u, 'LIBUR_RESET', '', { periode: p.periode, entri: hit.length, jmlKembali: kembali }); return { kembali: kembali, entri: hit.length };
    }
    const tgl = chkDate_(p.tgl), scope = p.scope || 'ALL', nilai = scope === 'ALL' ? '' : String(p.nilai || ''), mode = p.mode === 'GESER' ? 'GESER' : 'KOSONG';
    if (['ALL', 'BRANCH', 'GROUP'].indexOf(scope) < 0) throw new Error('Cakupan tidak valid');
    if (scope !== 'ALL' && !nilai) throw new Error('Pilih ' + (scope === 'BRANCH' ? 'branch' : 'group'));
    const akhir = p.tglAkhir ? chkDate_(p.tglAkhir, 'Tanggal akhir') : tgl;
    if (akhir < tgl) throw new Error('Tanggal akhir tidak boleh sebelum tanggal mulai');
    if (akhir.slice(0, 7) !== tgl.slice(0, 7)) throw new Error('Rentang tanggal harus dalam bulan yang sama');
    let geser = 0;
    if (mode === 'GESER') { geser = p.geser == null || p.geser === '' ? 1 : Math.round(Number(p.geser)); if (!(geser >= 1 && geser <= 14)) throw new Error('Geser harus 1-14 hari'); }
    if (all.some(b => b.tgl === tgl && b.scope === scope && b.nilai === nilai)) throw new Error('Libur/SO untuk tanggal & cakupan ini sudah ada — hapus dulu bila ingin mengubah');
    const entry = { tgl: tgl, scope: scope, nilai: nilai, ket: p.ket || '', oleh: u, mode: mode, geser: geser, jml: 0, waktu: now_(), tglAkhir: akhir };
    const per = tgl.slice(0, 7), before = planOf_(per, toko, groups, all).plan, after = planOf_(per, toko, groups, all.concat([entry])).plan;
    entry.jml = changed_(before, after);
    put_(sh_('BLACKOUT'), blackRows_([entry]));
    hist_(u, 'LIBUR_TAMBAH', '', { tgl: tgl, tglAkhir: akhir, scope: scope, nilai: nilai, mode: mode, geser: geser, ket: entry.ket, jmlToko: entry.jml });
    return { jml: entry.jml };
  },
  groupSave(p, u) {
    const g = String(p.group || '').trim().toUpperCase(); if (!g) throw new Error('Group wajib');
    if (!/^[A-Z0-9._-]{2,12}$/.test(g)) throw new Error('Nama group hanya huruf/angka (2-12 karakter)');
    const branch = String(p.branch || '').trim().toUpperCase(); if (!branch) throw new Error('Branch wajib');
    const hs = [p.h1, p.h2, p.h3].map(x => String(x || '').trim().toUpperCase());
    if (!hs[0]) throw new Error('Hari 1 wajib diisi');
    hs.forEach(h => { if (h && HARI.indexOf(h) < 0) throw new Error('Hari tidak valid: ' + h); });
    const filled = hs.filter(Boolean); if (new Set(filled).size !== filled.length) throw new Error('Hari kirim tidak boleh sama');
    const s = sh_('GROUP'), v = readAll_('GROUP').map(r => r.map(String)), row = [g, branch, hs[0], hs[1], hs[2]];
    const i = v.findIndex(r => r[0] === g); if (i < 0) v.push(row); else v[i] = row;
    clearData_(s); put_(s, v); hist_(u, i < 0 ? 'GROUP_BARU' : 'GROUP_SIMPAN', g, row); return {};
  },
  groupDelete(p, u) {
    const g = String(p.group || '').trim().toUpperCase(); if (!g) throw new Error('Group wajib');
    const pakai = tokoAll_().filter(t => t.group === g || t.groupLama === g).length;
    if (pakai) throw new Error('Group ' + g + ' masih dipakai ' + pakai + ' toko — pindahkan tokonya dulu');
    const s = sh_('GROUP'), v = readAll_('GROUP').map(r => r.map(String));
    if (!v.some(r => r[0] === g)) throw new Error('Group tidak ditemukan');
    clearData_(s); put_(s, v.filter(r => r[0] !== g)); hist_(u, 'GROUP_HAPUS', g, {}); return {};
  },
  targetSave(p, u) {                     // target kiriman per bulan: total (ALL) atau per branch; kosong/0 = hapus
    if (!RE_PER.test(p.periode || '')) throw new Error('Periode tidak valid');
    const scope = p.scope === 'BRANCH' ? 'BRANCH' : 'ALL', nilai = scope === 'ALL' ? '' : String(p.nilai || '').trim().toUpperCase();
    if (scope === 'BRANCH' && !nilai) throw new Error('Branch wajib');
    const raw = String(p.target == null ? '' : p.target).trim(), v = raw === '' ? 0 : Number(raw.replace(/\./g, '').replace(',', '.'));
    if (!isFinite(v) || v < 0) throw new Error('Target harus angka positif');
    const s = sh_('TARGET'), rows = readAll_('TARGET').map(r => r.map(String)).filter(r => !(r[0] === p.periode && r[1] === scope && r[2] === nilai));
    if (v > 0) rows.push([p.periode, scope, nilai, String(Math.round(v)), u, now_()]);
    clearData_(s); put_(s, rows); hist_(u, 'TARGET_SIMPAN', '', { periode: p.periode, scope: scope, nilai: nilai, target: Math.round(v) }); return {};
  },
  settingSave(p, u) {
    if (!(p.key in DEF_SET)) throw new Error('Pengaturan tidak dikenal');
    const val = String(p.value == null ? '' : p.value).trim(); if (!val) throw new Error('Nilai wajib diisi');
    if (p.key === 'ALARM_EMAIL' && !/^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$/.test(val)) throw new Error('Format email tidak valid');
    const s = sh_('SETTING'), rows = readAll_('SETTING').map(r => r.map(String)), i = rows.findIndex(r => r[0] === p.key);
    if (i < 0) rows.push([p.key, val]); else rows[i][1] = val;
    clearData_(s); put_(s, rows); hist_(u, 'SETTING', '', { key: p.key, value: val }); return {};
  },
  slBegin(p, u) {                        // hapus data SL pada tanggal yang akan diupload ulang
    snap_('sebelum_upload_sl');
    const byPer = {}; (p.dates || []).forEach(d => { chkDate_(d); (byPer[d.slice(0, 7)] = byPer[d.slice(0, 7)] || {})[d] = 1; });
    Object.keys(byPer).forEach(per => {
      const s = sh_('SL_' + per), n = s.getLastRow(); if (n < 2) return;
      const keep = s.getRange(2, 1, n - 1, HDR.SL.length).getValues().filter(r => !byPer[per][ds_(r[0])]).map(r => r.map(String));
      clearData_(s); put_(s, keep);
    });
    return { periods: Object.keys(byPer) };
  },
  slAppend(p) {
    const by = {}; (p.rows || []).forEach(r => { const per = String(r[0]).slice(0, 7); if (!RE_PER.test(per)) throw new Error('Tanggal SL tidak valid: ' + r[0]); const row = r.map(String); while (row.length < HDR.SL.length) row.push(''); (by[per] = by[per] || []).push(row.slice(0, HDR.SL.length)); });
    Object.keys(by).forEach(per => put_(sh_('SL_' + per), by[per]));
    return { n: (p.rows || []).length };
  },
  slEnd(p, u) { put_(sh_('UPLOAD'), [[now_(), u, p.file || '', p.count || 0, p.min || '', p.max || '']]); hist_(u, 'UPLOAD_SL', '', { file: p.file, baris: p.count, min: p.min, max: p.max, rpPb: p.rpPb, rpShip: p.rpShip }); return {}; },
  imporBegin(p, u) {                     // impor awal dari file JPF (master toko + group + plan dasar periode)
    snap_('sebelum_impor');
    clearData_(sh_('TOKO')); clearData_(sh_('GROUP'));
    put_(sh_('GROUP'), (p.groups || []).map(g => [g.group, g.branch, g.h1, g.h2, g.h3]));
    const pb = sh_('PLAN_BASE'), keep = readAll_('PLAN_BASE').filter(r => String(r[0]) !== p.periode).map(r => r.map(String));
    clearData_(pb); put_(pb, keep);
    const have = branchSheetRows_(), add = (p.branches || []).map(b => String(b).trim().toUpperCase()).filter(b => b && have.indexOf(b) < 0);
    if (add.length) put_(sh_('BRANCH'), add.map(b => [b]));
    return {};
  },
  imporToko(p) {
    const seen = {}; readAll_('TOKO').forEach(r => seen[String(r[0])] = 1);
    const rows = (p.rows || []).filter(t => t.kode && !seen[t.kode] && (seen[t.kode] = 1));
    put_(sh_('TOKO'), rows.map(t => [t.kode, t.ins, t.nama, t.branch, t.group, t.tutup ? 'TUTUP' : (t.go ? 'GO' : 'AKTIF'), t.tutup || '', t.go || '', '', '', '', t.ket || '', now_(), String(t.rit || '1')]));
    put_(sh_('PLAN_BASE'), rows.map(t => [p.periode, t.kode, t.plan]));
    return { n: rows.length };
  },
  imporEnd(p, u) {
    if (RE_PER.test(p.periode || '') && Number(p.target) > 0 && !targetAll_(p.periode).total) put_(sh_('TARGET'), [[p.periode, 'ALL', '', String(Math.round(Number(p.target))), u, now_()]]);
    hist_(u, 'IMPOR_AWAL', '', { periode: p.periode, toko: p.count, group: p.groupCount, duplikat: p.dups || [], target: p.target || 0 }); return {};
  }
};
