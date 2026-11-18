/**
 * JADWAL TOKO — Google Apps Script (database di Google Sheets + backup).
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 * Jalankan setup() sekali, lalu installTriggers() sekali. Salin API_SECRET (lihat Logs) ke env Vercel (GS_SECRET).
 * Atur zona waktu: Project Settings > Time zone = Asia/Jakarta.
 */
const SS = SpreadsheetApp.getActive();
const TZ = Session.getScriptTimeZone();
const P = PropertiesService.getScriptProperties();
const HARI = ['MINGGU','SENIN','SELASA','RABU','KAMIS','JUMAT','SABTU'];
const HDR = {
  TOKO: ['KODE','INS','NAMA','BRANCH','GROUP','STATUS','TUTUP_MULAI','GO_MULAI','GROUP_LAMA','GROUP_BERLAKU','KODE_LAMA','KET','UPDATED'],
  GROUP: ['GROUP','BRANCH','H1','H2','H3'],
  BLACKOUT: ['TANGGAL','SCOPE','NILAI','KET','OLEH'],
  PLAN_BASE: ['PERIODE','KODE','PLAN'],
  HISTORY: ['WAKTU','USER','AKSI','KODE','DETAIL'],
  UPLOAD: ['WAKTU','USER','FILE','BARIS','TGL_MIN','TGL_MAX'],
  SL: ['TANGGAL','KODE','NAMA','BRANCH','QTY_PB','QTY_SHIP']
};

// ---------- setup ----------
function setup() {
  ['TOKO','GROUP','BLACKOUT','PLAN_BASE','HISTORY','UPLOAD'].forEach(sh_);
  if (!P.getProperty('API_SECRET')) P.setProperty('API_SECRET', Utilities.getUuid() + Utilities.getUuid());
  Logger.log('API_SECRET = ' + P.getProperty('API_SECRET'));
}
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();
}
function dailyBackup() { doBackup_('harian'); }

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
const WRITE = ['tokoSave','setStatus','pindahGroup','gantiKode','blackout','groupSave','slBegin','slAppend','slEnd','imporBegin','imporToko','imporEnd'];
function route_(a, p, u) {
  if (a === 'ver') return { ver: P.getProperty('VER') || '0' };
  if (a === 'state') return state_(p.periode);
  if (a === 'history') return history_(p);
  if (a === 'backups') return backups_();
  if (a === 'backupNow') return { name: doBackup_('manual') };
  const fn = ACT[a];
  if (!fn) throw new Error('aksi tidak dikenal: ' + a);
  if (WRITE.indexOf(a) < 0) return fn(p, u);
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  try { const r = fn(p, u); P.setProperty('VER', String(Date.now())); return r; } finally { lock.releaseLock(); }
}

// ---------- helpers ----------
function sh_(n) {
  let s = SS.getSheetByName(n);
  if (!s) {
    s = SS.insertSheet(n);
    const h = n.indexOf('SL_') === 0 ? HDR.SL : HDR[n];
    s.getRange(1, 1, 1, h.length).setNumberFormat('@').setValues([h]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function ds_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : (v == null || v === '' ? '' : String(v).slice(0, 10)); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function put_(s, rows) {            // tulis sebagai teks (anti auto-format tanggal/angka)
  if (!rows.length) return;
  const start = s.getLastRow() + 1, need = start + rows.length - 1;
  if (need > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), need - s.getMaxRows());
  const r = s.getRange(start, 1, rows.length, rows[0].length);
  r.setNumberFormat('@'); r.setValues(rows);
}
function clearData_(s) { const n = s.getLastRow(); if (n > 1) s.getRange(2, 1, n - 1, s.getMaxColumns()).clearContent(); }
function readAll_(name) { const s = sh_(name), n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, HDR[name].length).getValues(); }
function hist_(u, aksi, kode, detail) { put_(sh_('HISTORY'), [[now_(), u, aksi, kode || '', typeof detail === 'string' ? detail : JSON.stringify(detail || {})]]); }

function tokoAll_() {
  return readAll_('TOKO').filter(r => r[0] !== '').map(r => ({
    kode: String(r[0]), ins: String(r[1]), nama: String(r[2]), branch: String(r[3]), group: String(r[4]),
    status: String(r[5] || 'AKTIF'), tutup: ds_(r[6]), go: ds_(r[7]), groupLama: String(r[8] || ''),
    berlaku: ds_(r[9]), kodeLama: String(r[10] || ''), ket: String(r[11] || '')
  }));
}
function tokoRow_(t) { return [t.kode, t.ins, t.nama, t.branch, t.group, t.status, t.tutup, t.go, t.groupLama, t.berlaku, t.kodeLama, t.ket, now_()]; }
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
  if (row < 0) put_(s, [tokoRow_(t)]); else { const r = s.getRange(row, 1, 1, 13); r.setNumberFormat('@'); r.setValues([tokoRow_(t)]); }
}
function groupAll_() { return readAll_('GROUP').filter(r => r[0] !== '').map(r => ({ group: String(r[0]), branch: String(r[1]), h1: String(r[2]), h2: String(r[3]), h3: String(r[4]) })); }
function blackAll_() { return readAll_('BLACKOUT').filter(r => r[0] !== '').map(r => ({ tgl: ds_(r[0]), scope: String(r[1]), nilai: String(r[2]), ket: String(r[3]), oleh: String(r[4]) })); }
function baseMap_(per) { const m = {}; readAll_('PLAN_BASE').forEach(r => { if (String(r[0]) === per) m[String(r[1])] = String(r[2]); }); return m; }

// aturan aktif: GO belum mulai / sudah TUTUP; jika GO > TUTUP berarti RE-GO (tutup sementara lalu buka lagi)
function aktif_(t, d) {
  if (t.tutup && t.go && t.go > t.tutup) return !(d >= t.tutup && d < t.go);
  if (t.go && d < t.go) return false;
  if (t.tutup && d >= t.tutup) return false;
  return true;
}

// ---------- plan & aktual ----------
function buildPlan_(y, m, nd, toko, groups, bo, base) {
  const gmap = {}; groups.forEach(g => gmap[g.group] = g);
  const dates = [], dows = [];
  for (let i = 0; i <= 31; i++) {
    if (i > nd) { dates.push(null); dows.push(0); continue; }
    const dt = new Date(y, m - 1, i, 12);             // i=0 => hari terakhir bulan lalu (kolom carry)
    dates.push(Utilities.formatDate(dt, TZ, 'yyyy-MM-dd')); dows.push(dt.getDay());
  }
  const boBy = {}; bo.forEach(b => (boBy[b.tgl] = boBy[b.tgl] || []).push(b));
  const out = {};
  toko.forEach(t => {
    const bs = base[t.kode]; let s = '';
    for (let i = 0; i <= 31; i++) {
      const ds = dates[i]; if (!ds) { s += '0'; continue; }
      const g = (t.berlaku && t.groupLama && ds < t.berlaku) ? t.groupLama : t.group;
      let p;
      if (bs && !(t.berlaku && ds >= t.berlaku)) p = bs.charAt(i) === '1';
      else { const G = gmap[g]; p = !!G && [G.h1, G.h2, G.h3].indexOf(HARI[dows[i]]) >= 0; }
      if (p && !aktif_(t, ds)) p = false;
      if (p && boBy[ds] && boBy[ds].some(b => b.scope === 'ALL' || (b.scope === 'BRANCH' && b.nilai === t.branch) || (b.scope === 'GROUP' && b.nilai === g))) p = false;
      s += p ? '1' : '0';
    }
    out[t.kode] = s;
  });
  return out;
}
function buildAct_(per) {
  const s = SS.getSheetByName('SL_' + per), out = {};
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(r => {
    const d = ds_(r[0]); if (d.slice(0, 7) !== per) return;
    const k = String(r[1]); if (!out[k]) out[k] = new Array(32).fill('0');
    out[k][+d.slice(8, 10)] = '1';
  });
  Object.keys(out).forEach(k => out[k] = out[k].join(''));
  return out;
}
function state_(per) {
  per = /^\d{4}-\d{2}$/.test(per || '') ? per : Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const y = +per.slice(0, 4), m = +per.slice(5), nd = new Date(y, m, 0).getDate();
  const toko = tokoAll_(), groups = groupAll_(), bo = blackAll_(), base = baseMap_(per);
  return { ver: P.getProperty('VER') || '0', periode: per, nd: nd, toko: toko, groups: groups, blackout: bo,
    plan: buildPlan_(y, m, nd, toko, groups, bo, base), act: buildAct_(per), hasBase: Object.keys(base).length > 0 };
}
function history_(p) {
  const s = sh_('HISTORY'), n = s.getLastRow(); if (n < 2) return [];
  const from = Math.max(2, n - 2999);
  let v = s.getRange(from, 1, n - from + 1, 5).getValues().map(r => ({ waktu: String(r[0]), user: r[1], aksi: r[2], kode: String(r[3]), detail: r[4] }));
  if (p.kode) v = v.filter(x => x.kode === String(p.kode));
  return v.reverse().slice(0, p.limit || 400);
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

// ---------- aksi tulis ----------
const ACT = {
  tokoSave(p, u) {                       // tambah toko baru (GO) atau ubah info dasar
    const kode = String(p.kode || '').trim().toUpperCase(); if (!kode) throw new Error('Kode toko wajib');
    const old = getToko_(kode);
    if (old) {
      const t = Object.assign({}, old, { ins: p.ins || old.ins, nama: p.nama || old.nama, branch: p.branch || old.branch, ket: p.ket != null ? p.ket : old.ket });
      saveToko_(t); hist_(u, 'UBAH_INFO', kode, { dari: old, ke: t }); return { kode: kode };
    }
    if (!p.group || !p.branch) throw new Error('Branch dan group wajib untuk toko baru');
    const t = { kode: kode, ins: p.ins || '', nama: p.nama || '', branch: p.branch, group: p.group, status: p.tgl ? 'GO' : 'AKTIF',
      tutup: '', go: p.tgl || '', groupLama: '', berlaku: '', kodeLama: '', ket: p.ket || '' };
    saveToko_(t); hist_(u, 'TOKO_BARU', kode, t); return { kode: kode };
  },
  setStatus(p, u) {                      // TUTUP / GO / AKTIF(batalkan)
    const t = getToko_(p.kode); if (!t) throw new Error('Toko tidak ditemukan');
    const before = { status: t.status, tutup: t.tutup, go: t.go };
    if (p.status === 'TUTUP') { t.tutup = p.tgl; if (t.go && t.go > p.tgl) t.go = ''; }
    else if (p.status === 'GO') { t.go = p.tgl; if (t.tutup && t.tutup > p.tgl) t.tutup = ''; }
    else { t.tutup = ''; t.go = ''; }
    if (!/^(TUTUP|GO|AKTIF)$/.test(p.status)) throw new Error('Status tidak valid');
    if (p.status !== 'AKTIF' && !p.tgl) throw new Error('Tanggal wajib');
    t.status = p.status; if (p.ket) t.ket = p.ket;
    saveToko_(t); hist_(u, 'STATUS_' + p.status, t.kode, { sebelum: before, tgl: p.tgl, ket: p.ket || '' }); return {};
  },
  pindahGroup(p, u) {
    const t = getToko_(p.kode); if (!t) throw new Error('Toko tidak ditemukan');
    if (!p.groupBaru || !p.berlaku) throw new Error('Group baru dan tanggal berlaku wajib');
    const lama = (t.berlaku && t.groupLama && p.berlaku < t.berlaku) ? t.groupLama : t.group;
    t.groupLama = lama; t.group = p.groupBaru; t.berlaku = p.berlaku;
    const g = groupAll_().filter(x => x.group === p.groupBaru)[0]; if (g && g.branch) t.branch = g.branch;
    saveToko_(t); hist_(u, 'PINDAH_GROUP', t.kode, { dari: lama, ke: p.groupBaru, berlaku: p.berlaku }); return {};
  },
  gantiKode(p, u) {                      // toko lama TUTUP (take over), toko baru GO dengan group sama
    const o = getToko_(p.kodeLama); if (!o) throw new Error('Kode lama tidak ditemukan');
    const nk = String(p.kodeBaru || '').trim().toUpperCase(); if (!nk || !p.tgl) throw new Error('Kode baru dan tanggal wajib');
    if (getToko_(nk)) throw new Error('Kode baru sudah ada');
    const n = { kode: nk, ins: p.ins || o.ins, nama: p.nama || o.nama, branch: o.branch, group: o.group, status: 'GO', tutup: '', go: p.tgl,
      groupLama: o.groupLama, berlaku: o.berlaku, kodeLama: o.kode, ket: 'TAKE OVER DARI ' + o.kode };
    o.tutup = p.tgl; o.status = 'TUTUP'; o.ket = 'TAKE OVER KE ' + nk;
    saveToko_(o); saveToko_(n); hist_(u, 'GANTI_KODE', o.kode, { baru: nk, tgl: p.tgl }); hist_(u, 'GANTI_KODE', nk, { lama: o.kode, tgl: p.tgl }); return {};
  },
  blackout(p, u) {                       // hari libur / stock opname: ALL | BRANCH | GROUP
    const s = sh_('BLACKOUT');
    if (p.op === 'del') {
      const v = readAll_('BLACKOUT').filter(r => !(ds_(r[0]) === p.tgl && String(r[1]) === p.scope && String(r[2]) === (p.nilai || '')));
      clearData_(s); put_(s, v.map(r => r.map(String)));
      hist_(u, 'LIBUR_HAPUS', '', p); return {};
    }
    if (!p.tgl) throw new Error('Tanggal wajib');
    put_(s, [[p.tgl, p.scope || 'ALL', p.nilai || '', p.ket || '', u]]); hist_(u, 'LIBUR_TAMBAH', '', p); return {};
  },
  groupSave(p, u) {
    const g = String(p.group || '').trim().toUpperCase(); if (!g) throw new Error('Group wajib');
    const s = sh_('GROUP'), v = readAll_('GROUP').map(r => r.map(String)), row = [g, p.branch || '', p.h1 || '', p.h2 || '', p.h3 || ''];
    const i = v.findIndex(r => r[0] === g); if (i < 0) v.push(row); else v[i] = row;
    clearData_(s); put_(s, v); hist_(u, 'GROUP_SIMPAN', g, row); return {};
  },
  slBegin(p, u) {                        // hapus data SL pada tanggal yang akan diupload ulang
    snap_('sebelum_upload_sl');
    const byPer = {}; (p.dates || []).forEach(d => { (byPer[d.slice(0, 7)] = byPer[d.slice(0, 7)] || {})[d] = 1; });
    Object.keys(byPer).forEach(per => {
      const s = sh_('SL_' + per), n = s.getLastRow(); if (n < 2) return;
      const keep = s.getRange(2, 1, n - 1, 6).getValues().filter(r => !byPer[per][ds_(r[0])]).map(r => r.map(String));
      clearData_(s); put_(s, keep);
    });
    return { periods: Object.keys(byPer) };
  },
  slAppend(p) {
    const by = {}; (p.rows || []).forEach(r => { const per = String(r[0]).slice(0, 7); (by[per] = by[per] || []).push(r.map(String)); });
    Object.keys(by).forEach(per => put_(sh_('SL_' + per), by[per]));
    return { n: (p.rows || []).length };
  },
  slEnd(p, u) { put_(sh_('UPLOAD'), [[now_(), u, p.file || '', p.count || 0, p.min || '', p.max || '']]); hist_(u, 'UPLOAD_SL', '', { file: p.file, baris: p.count, min: p.min, max: p.max }); return {}; },
  imporBegin(p, u) {                     // impor awal dari file JPF (master toko + group + plan dasar periode)
    snap_('sebelum_impor');
    clearData_(sh_('TOKO')); clearData_(sh_('GROUP'));
    put_(sh_('GROUP'), (p.groups || []).map(g => [g.group, g.branch, g.h1, g.h2, g.h3]));
    const pb = sh_('PLAN_BASE'), keep = readAll_('PLAN_BASE').filter(r => String(r[0]) !== p.periode).map(r => r.map(String));
    clearData_(pb); put_(pb, keep); return {};
  },
  imporToko(p) {
    const rows = p.rows || [];
    put_(sh_('TOKO'), rows.map(t => [t.kode, t.ins, t.nama, t.branch, t.group, t.tutup ? 'TUTUP' : (t.go ? 'GO' : 'AKTIF'), t.tutup || '', t.go || '', '', '', '', t.ket || '', now_()]));
    put_(sh_('PLAN_BASE'), rows.map(t => [p.periode, t.kode, t.plan]));
    return { n: rows.length };
  },
  imporEnd(p, u) { hist_(u, 'IMPOR_AWAL', '', { periode: p.periode, toko: p.count }); return {}; }
};
