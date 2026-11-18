// Vercel serverless: POST /api  { action, payload }  + header Authorization: Bearer <Firebase ID token>
// Tugas: verifikasi login, cek hak akses, lalu teruskan ke Google Apps Script (yang dilindungi GS_SECRET).
const admin = require('firebase-admin');

if (!admin.apps.length) {
  admin.initializeApp({
    credential: admin.credential.cert({
      projectId: process.env.FB_PROJECT_ID,
      clientEmail: process.env.FB_CLIENT_EMAIL,
      privateKey: (process.env.FB_PRIVATE_KEY || '').replace(/\\n/g, '\n')
    })
  });
}

const DOMAIN = '@outboundpro.app';
const ALL = ['view', 'edit', 'upload', 'download', 'admin'];
const DEFAULT_PERMS = { admin: ALL, editor: ['view', 'edit', 'upload', 'download'], viewer: ['view', 'download'] };

// aksi -> izin yang dibutuhkan
const NEED = {
  me: 'view', ver: 'view', state: 'view', history: 'view',
  tokoSave: 'edit', setStatus: 'edit', pindahGroup: 'edit', gantiKode: 'edit', blackout: 'edit', groupSave: 'edit',
  slBegin: 'upload', slAppend: 'upload', slEnd: 'upload',
  imporBegin: 'admin', imporToko: 'admin', imporEnd: 'admin', backupNow: 'admin', backups: 'admin',
  'users.list': 'admin', 'users.save': 'admin'
};

function who(d) {
  const name = (d.email || '').replace(DOMAIN, '');
  // akun lama dengan kata "admin" di username tetap dianggap admin (sama seperti dashboard lama)
  const role = d.role || (/admin/.test(name) ? 'admin' : 'none');
  const perms = role === 'admin' ? ALL : (Array.isArray(d.perms) ? d.perms : (DEFAULT_PERMS[role] || []));
  return { uid: d.uid, name, role, perms };
}

async function callGS(action, payload, user) {
  const r = await fetch(process.env.GS_URL, {
    method: 'POST', redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ secret: process.env.GS_SECRET, action, payload, user })
  });
  const t = await r.text();
  try { return JSON.parse(t); } catch (e) { throw new Error('Respon GS tidak valid: ' + t.slice(0, 160)); }
}

async function users(action, p, me) {
  const a = admin.auth();
  if (action === 'users.list') {
    const r = await a.listUsers(1000);
    return r.users.map(u => ({
      uid: u.uid, username: (u.email || '').replace(DOMAIN, ''), disabled: u.disabled,
      role: (u.customClaims || {}).role || '-', perms: (u.customClaims || {}).perms || [],
      last: u.metadata.lastSignInTime || ''
    }));
  }
  const username = String(p.username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('Username 3-30 karakter (huruf kecil, angka, . _ -)');
  const role = ['admin', 'editor', 'viewer'].includes(p.role) ? p.role : 'viewer';
  const perms = (p.perms || []).filter(x => ALL.includes(x) && x !== 'admin');
  let u;
  try { u = await a.getUserByEmail(username + DOMAIN); }
  catch (e) {
    if (!p.password || p.password.length < 6) throw new Error('Password minimal 6 karakter');
    u = await a.createUser({ email: username + DOMAIN, password: p.password });
  }
  if (u.uid === me.uid && (role !== 'admin' || p.disabled)) throw new Error('Tidak bisa menurunkan / menonaktifkan akun sendiri');
  const upd = { disabled: !!p.disabled };
  if (p.password) { if (p.password.length < 6) throw new Error('Password minimal 6 karakter'); upd.password = p.password; }
  await a.updateUser(u.uid, upd);
  await a.setCustomUserClaims(u.uid, { role, perms });
  await a.revokeRefreshTokens(u.uid);              // klaim baru berlaku saat login ulang
  return { uid: u.uid };
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'POST only' });
  try {
    const token = (req.headers.authorization || '').replace(/^Bearer /, '');
    if (!token) return res.status(401).json({ ok: false, error: 'Belum login' });
    let dec;
    try { dec = await admin.auth().verifyIdToken(token, true); }
    catch (e) { return res.status(401).json({ ok: false, error: 'Sesi tidak valid, silakan login ulang' }); }

    const me = who(dec);
    const { action, payload } = req.body || {};
    const need = NEED[action];
    if (!need) return res.status(400).json({ ok: false, error: 'Aksi tidak dikenal' });
    if (!me.perms.includes(need)) return res.status(403).json({ ok: false, error: 'Tidak punya izin: ' + need });

    if (action === 'me') return res.json({ ok: true, data: me });
    if (action.startsWith('users.')) return res.json({ ok: true, data: await users(action, payload || {}, me) });
    return res.json(await callGS(action, payload || {}, me.name));
  } catch (e) {
    return res.status(500).json({ ok: false, error: String(e.message || e) });
  }
};
module.exports.config = { api: { bodyParser: { sizeLimit: '4mb' } } };
