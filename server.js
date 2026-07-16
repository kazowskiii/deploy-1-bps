/**
 * SIMONEV BPS — Sistem Monitoring & Evaluasi Kinerja
 * Backend server: Express + penyimpanan file lokal di server (folder /uploads)
 * Database sederhana disimpan sebagai JSON di /data/db.json
 * Autentikasi session, validasi input, backup harian, export JSON/Excel/PDF,
 * dan audit log.
 *
 * Menjalankan:
 *   npm install
 *   npm start
 * Lalu buka http://localhost:3000
 */

require('dotenv').config();
const express = require('express');
const session = require('express-session');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const Joi = require('joi');
const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const onedrive = require('./onedrive');

const app = express();
const PORT = process.env.PORT || 3000;
const SESSION_SECRET = process.env.SESSION_SECRET || 'simonev-session-secret';

const DATA_DIR = path.join(__dirname, 'data');
const BACKUP_DIR = path.join(DATA_DIR, 'backup');
const DB_FILE = path.join(DATA_DIR, 'db.json');
const MAX_FILE_MB = 10;

const USER_ACCOUNTS = [
  { username: 'admin', password: 'admin123', role: 'admin' },
  { username: 'operator_sosial', password: 'opsosial1', role: 'operator', teamName: 'Tim Statistik Sosial' },
  { username: 'operator_produksi', password: 'opprod1', role: 'operator', teamName: 'Tim Statistik Produksi' },
  { username: 'operator_distribusi', password: 'opdist1', role: 'operator', teamName: 'Tim Statistik Distribusi' },
  { username: 'operator_neraca', password: 'opneraca1', role: 'operator', teamName: 'Tim Neraca & Analisis Statistik' },
  { username: 'operator_ipds', password: 'opipds1', role: 'operator', teamName: 'Tim IPDS (Integrasi Pengolahan & Diseminasi Statistik)' },
  { username: 'operator_tu', password: 'optu1', role: 'operator', teamName: 'Tim Tata Usaha' },
];

function ensureDirectory(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}
ensureDirectory(DATA_DIR);
ensureDirectory(BACKUP_DIR);

function sanitizeText(value) {
  if (typeof value !== 'string') return value;
  return value.replace(/<[^>]*>/g, '').trim();
}

function sanitizePayload(value) {
  if (Array.isArray(value)) return value.map(sanitizePayload);
  if (value && typeof value === 'object') {
    return Object.keys(value).reduce((acc, key) => {
      acc[key] = sanitizePayload(value[key]);
      return acc;
    }, {});
  }
  return sanitizeText(value);
}

function seedTeams() {
  return [
    { id: uuidv4(), name: 'Tim Statistik Sosial' },
    { id: uuidv4(), name: 'Tim Statistik Produksi' },
    { id: uuidv4(), name: 'Tim Statistik Distribusi' },
    { id: uuidv4(), name: 'Tim Neraca & Analisis Statistik' },
    { id: uuidv4(), name: 'Tim IPDS (Integrasi Pengolahan & Diseminasi Statistik)' },
    { id: uuidv4(), name: 'Tim Tata Usaha' },
  ];
}

function createInitialDB() {
  return {
    teams: seedTeams(),
    iku: [],
    fra: [],
    kegiatan: [],
    tugas: [],
    auditLogs: [],
  };
}

function backupCorruptedDb() {
  try {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupFile = path.join(BACKUP_DIR, `corrupt-db-${timestamp}.json`);
    fs.copyFileSync(DB_FILE, backupFile);
    console.error('Salinan db rusak disimpan di', backupFile);
  } catch (err) {
    console.error('Gagal membuat salinan backup db rusak:', err.message);
  }
}

function loadDB() {
  if (!fs.existsSync(DB_FILE)) {
    const initial = createInitialDB();
    fs.writeFileSync(DB_FILE, JSON.stringify(initial, null, 2));
    return initial;
  }
  try {
    return JSON.parse(fs.readFileSync(DB_FILE, 'utf-8'));
  } catch (err) {
    console.error('db.json rusak, membuat ulang dari kosong. Error:', err.message);
    backupCorruptedDb();
    const initial = createInitialDB();
    fs.writeFileSync(DB_FILE, JSON.stringify(initial, null, 2));
    return initial;
  }
}

let DB = loadDB();
let writeQueue = Promise.resolve();
function saveDB() {
  writeQueue = writeQueue.then(() => fs.promises.writeFile(DB_FILE, JSON.stringify(DB, null, 2)));
  return writeQueue;
}

async function backupDB() {
  try {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupFile = path.join(BACKUP_DIR, `db-backup-${timestamp}.json`);
    await fs.promises.copyFile(DB_FILE, backupFile);
    console.log('Backup db disimpan di', backupFile);
  } catch (err) {
    console.error('Gagal membuat backup db:', err.message);
  }
}

backupDB();
setInterval(backupDB, 24 * 60 * 60 * 1000);

/**
 * Pengingat email (opsional).
 * Aktif hanya jika variabel environment SMTP_HOST, SMTP_USER, SMTP_PASS diisi.
 * Kalau tidak dikonfigurasi, fitur ini otomatis dilewati (tidak mengganggu server).
 * Setiap tim bisa diberi alamat email tujuan lewat REMINDER_EMAILS (opsional),
 * format: "Tim Statistik Sosial=sosial@bps.go.id,Tim Statistik Produksi=produksi@bps.go.id"
 */
function loadReminderEmailMap() {
  const raw = process.env.REMINDER_EMAILS || '';
  const map = {};
  raw.split(',').filter(Boolean).forEach((pair) => {
    const [team, email] = pair.split('=').map((s) => (s || '').trim());
    if (team && email) map[team] = email;
  });
  return map;
}

async function sendReminderEmails() {
  const { SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM } = process.env;
  if (!SMTP_HOST || !SMTP_USER || !SMTP_PASS) {
    return; // Email reminder tidak dikonfigurasi — dilewati secara diam-diam.
  }
  const { daysLeft, items } = computeReminders({ role: 'admin' });
  if (daysLeft > 7 || !items.length) return; // hanya kirim di H-7 sebelum akhir triwulan

  let nodemailer;
  try {
    nodemailer = require('nodemailer');
  } catch (err) {
    console.warn('Paket nodemailer belum terpasang. Jalankan "npm install nodemailer" untuk mengaktifkan pengingat email.');
    return;
  }

  const emailMap = loadReminderEmailMap();
  const grouped = items.reduce((acc, item) => {
    (acc[item.timName] = acc[item.timName] || []).push(item);
    return acc;
  }, {});

  const transporter = nodemailer.createTransport({
    host: SMTP_HOST,
    port: Number(SMTP_PORT) || 587,
    secure: Number(SMTP_PORT) === 465,
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });

  for (const [timName, tasks] of Object.entries(grouped)) {
    const to = emailMap[timName];
    if (!to) continue; // tidak ada alamat tujuan yang dikonfigurasi untuk tim ini
    const list = tasks.map((t) => `- ${t.nama} (${t.quarterLabel})`).join('\n');
    try {
      await transporter.sendMail({
        from: SMTP_FROM || SMTP_USER,
        to,
        subject: `[SIMONEV] Pengingat: realisasi ${tasks[0].quarterLabel} belum diisi`,
        text: `Halo ${timName},\n\nBerikut tugas yang belum diisi realisasinya untuk triwulan berjalan (H-${daysLeft}):\n\n${list}\n\nMohon segera dilengkapi di SIMONEV.\n\n— Sistem SIMONEV BPS`,
      });
      console.log(`Pengingat email terkirim ke ${to} (${timName})`);
    } catch (err) {
      console.error(`Gagal mengirim pengingat email ke ${to}:`, err.message);
    }
  }
}

// Cek sekali saat start, lalu setiap hari sekali (fitur ini aman walau SMTP tidak diisi).
sendReminderEmails();
setInterval(sendReminderEmails, 24 * 60 * 60 * 1000);

function findTeamIdByName(name) {
  const team = DB.teams.find((t) => t.name === name);
  return team ? team.id : null;
}

function mapUserTeams() {
  USER_ACCOUNTS.forEach((user) => {
    if (user.teamName) {
      user.teamId = findTeamIdByName(user.teamName);
    }
  });
}
mapUserTeams();

function getUserByUsername(username) {
  return USER_ACCOUNTS.find((u) => u.username === username);
}

function getItemStatus(collection, item) {
  if (collection === 'iku') {
    const target = Number(item.target) || 0;
    const capaian = Number(item.capaian) || 0;
    if (!target) return 'Tertinggal';
    const pct = Math.round((capaian / target) * 100);
    if (pct >= 100) return 'Tercapai';
    if (pct >= 60) return 'Perlu Perhatian';
    return 'Tertinggal';
  }
  if (collection === 'tugas') {
    const total = ['q1', 'q2', 'q3', 'q4'].reduce((sum, q) => sum + (Number((item[q] || {}).realisasi) || 0), 0);
    const target = Number(item.target) || 0;
    if (!target) return 'Tertinggal';
    const pct = Math.round((total / target) * 100);
    if (pct >= 100) return 'Tercapai';
    if (pct >= 60) return 'Perlu Perhatian';
    return 'Tertinggal';
  }
  return item.status || '—';
}

function currentQuarterInfo() {
  const now = new Date();
  const q = Math.floor(now.getMonth() / 3); // 0..3
  const quarterEndMonth = q * 3 + 2; // 0-indexed month of last month in quarter
  const quarterEnd = new Date(now.getFullYear(), quarterEndMonth + 1, 0); // last day of quarter
  const daysLeft = Math.ceil((quarterEnd - now) / (24 * 60 * 60 * 1000));
  return { quarterKey: `q${q + 1}`, quarterIndex: q, year: now.getFullYear(), quarterEnd, daysLeft };
}

function computeReminders(user) {
  const { quarterKey, quarterIndex, year, daysLeft } = currentQuarterInfo();
  const scoped = filterItems(DB.tugas, user, 'tugas', {}).filter((t) => Number(t.tahun) === year);
  const reminders = [];
  scoped.forEach((t) => {
    const q = t[quarterKey] || {};
    const realisasi = Number(q.realisasi) || 0;
    if (realisasi <= 0) {
      const team = DB.teams.find((tm) => tm.id === t.timId);
      reminders.push({
        tugasId: t.id,
        nama: t.nama,
        timId: t.timId,
        timName: team ? team.name : '—',
        quarter: quarterKey,
        quarterLabel: `Triwulan ${['I', 'II', 'III', 'IV'][quarterIndex]}`,
        daysLeft,
        urgent: daysLeft <= 7,
      });
    }
  });
  return { year, quarterKey, quarterIndex, daysLeft, items: reminders };
}

function auditLog(user, action, collection, itemId, details) {
  DB.auditLogs = DB.auditLogs || [];
  DB.auditLogs.unshift({
    id: uuidv4(),
    timestamp: new Date().toISOString(),
    user: user?.username || 'unknown',
    role: user?.role || 'anonymous',
    action,
    collection,
    itemId,
    details,
  });
  return saveDB();
}

function filterItems(items, user, collection, query) {
  let result = Array.isArray(items) ? [...items] : [];
  if (user.role !== 'admin') {
    if (collection === 'teams') {
      result = result.filter((item) => item.id === user.teamId);
    } else {
      result = result.filter((item) => item.timId === user.teamId);
    }
  }
  if (query.teamId) {
    result = result.filter((item) => item.timId === query.teamId || item.id === query.teamId);
  }
  if (query.tahun) {
    result = result.filter((item) => Number(item.tahun) === Number(query.tahun));
  }
  if (query.status) {
    result = result.filter((item) => getItemStatus(collection, item).toLowerCase() === String(query.status).toLowerCase());
  }
  if (query.search) {
    const needle = String(query.search).toLowerCase();
    result = result.filter((item) => JSON.stringify(item).toLowerCase().includes(needle));
  }
  return result;
}

const teamSchema = Joi.object({
  name: Joi.string().trim().min(3).max(120).required(),
});
const fraSchema = Joi.object({
  timId: Joi.string().required(),
  periode: Joi.string().trim().min(3).max(120).required(),
  uraian: Joi.string().trim().min(3).max(1000).required(),
  persentase: Joi.number().min(0).max(200).required(),
  tahun: Joi.number().integer().min(2020).max(2100).required(),
  catatan: Joi.string().trim().max(1000).allow('', null),
  hasEvidence: Joi.boolean().optional(),
  evidenceFileName: Joi.string().trim().allow('', null),
  evidenceOriginalName: Joi.string().trim().allow('', null),
  evidenceSize: Joi.number().min(0).allow(null),
});
const kegiatanSchema = Joi.object({
  timId: Joi.string().required(),
  tanggal: Joi.string().isoDate().required(),
  nama: Joi.string().trim().min(3).max(250).required(),
  kendala: Joi.string().trim().min(3).max(1000).required(),
  solusi: Joi.string().trim().min(3).max(1000).required(),
  rtl: Joi.string().trim().min(3).max(1000).required(),
  status: Joi.string().valid('Belum Ditindaklanjuti', 'Dalam Proses', 'Selesai').required(),
});
const tugasSchema = Joi.object({
  timId: Joi.string().required(),
  tahun: Joi.number().integer().min(2020).max(2100).required(),
  nama: Joi.string().trim().min(3).max(250).required(),
  target: Joi.number().min(1).required(),
  satuan: Joi.string().trim().min(1).max(50).required(),
  q1: Joi.object().pattern(/.*/, Joi.any()).default({}),
  q2: Joi.object().pattern(/.*/, Joi.any()).default({}),
  q3: Joi.object().pattern(/.*/, Joi.any()).default({}),
  q4: Joi.object().pattern(/.*/, Joi.any()).default({}),
});
const ikuSchema = Joi.object({
  kode: Joi.string().trim().max(50).allow('', null),
  timId: Joi.string().required(),
  nama: Joi.string().trim().min(3).max(250).required(),
  target: Joi.number().min(1).required(),
  capaian: Joi.number().min(0).required(),
  satuan: Joi.string().trim().min(1).max(50).required(),
  tahun: Joi.number().integer().min(2020).max(2100).required(),
  hasEvidence: Joi.boolean().optional(),
  evidenceFileName: Joi.string().trim().allow('', null),
  evidenceOriginalName: Joi.string().trim().allow('', null),
  evidenceSize: Joi.number().min(0).allow(null),
});

function validateBody(schema) {
  return (req, res, next) => {
    const payload = sanitizePayload(req.body);
    const { error, value } = schema.validate(payload, {
      abortEarly: false,
      stripUnknown: true,
    });
    if (error) {
      return res.status(400).json({ error: 'Validasi gagal', details: error.details.map((d) => d.message) });
    }
    req.body = value;
    next();
  };
}

function isJsonRequest(req, res, next) {
  if (req.is('application/json')) return next();
  return res.status(400).json({ error: 'Request harus JSON' });
}

function requireLogin(req, res, next) {
  if (req.session && req.session.user) return next();
  return res.status(401).json({ error: 'Login diperlukan' });
}

function requireAdmin(req, res, next) {
  if (req.session && req.session.user && req.session.user.role === 'admin') return next();
  return res.status(403).json({ error: 'Akses ditolak' });
}

function authorizePayload(req, res, next) {
  const user = req.session.user;
  if (!user) return res.status(401).json({ error: 'Login diperlukan' });
  if (user.role === 'admin') return next();
  if ((req.body.timId && req.body.timId !== user.teamId) || (req.body.teamId && req.body.teamId !== user.teamId)) {
    return res.status(403).json({ error: 'Tidak dapat mengelola data untuk tim lain' });
  }
  next();
}

app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: '256kb' }));
app.use(express.urlencoded({ extended: false, limit: '64kb' }));
app.set('trust proxy', 1);
app.use(session({
  secret: process.env.SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: { 
    maxAge: undefined,   // tidak ada masa berlaku tetap = "session cookie"
    secure: true,         // wajib HTTPS (aman karena Railway sudah HTTPS)
    httpOnly: true        // cookie tidak bisa diakses lewat JavaScript (lebih aman)
  }
}));
app.use(express.static(path.join(__dirname, 'public')));

const storage = multer.memoryStorage();

function fileFilter(req, file, cb) {
  const okExt = ['.pdf', '.doc', '.docx'];
  const ext = path.extname(file.originalname).toLowerCase();
  if (okExt.includes(ext)) cb(null, true);
  else cb(new Error('Format berkas harus PDF atau Word (.pdf, .doc, .docx)'));
}

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: MAX_FILE_MB * 1024 * 1024 },
});

app.post('/api/login', isJsonRequest, validateBody(Joi.object({ username: Joi.string().trim().required(), password: Joi.string().required() })), (req, res) => {
  const { username, password } = req.body;
  const user = getUserByUsername(username);
  if (!user || user.password !== password) {
    return res.status(401).json({ error: 'Username atau kata sandi tidak cocok' });
  }
  req.session.user = { username: user.username, role: user.role, teamId: user.teamId || null };
  auditLog(req.session.user, 'login', 'auth', null, { ip: req.ip });
  res.json({ username: user.username, role: user.role, teamId: user.teamId || null });
});

app.post('/api/logout', requireLogin, (req, res) => {
  auditLog(req.session.user, 'logout', 'auth', null, { ip: req.ip });
  req.session.destroy(() => {});
  res.json({ ok: true });
});

app.get('/api/me', (req, res) => {
  if (req.session && req.session.user) {
    return res.json(req.session.user);
  }
  res.status(401).json({ error: 'Belum login' });
});

function collectEvidenceFileIds(name, item) {
  if (!item) return [];
  const ids = [];
  if (item.evidenceFileName) ids.push(item.evidenceFileName);
  if (name === 'tugas') {
    ['q1', 'q2', 'q3', 'q4'].forEach((q) => {
      const quarter = item[q];
      if (quarter && quarter.evidenceFileName) ids.push(quarter.evidenceFileName);
    });
  }
  return ids;
}

async function deleteEvidenceFiles(fileIds) {
  if (!fileIds.length) {
    console.log('[OneDrive] Tidak ada berkas bukti terkait untuk dihapus (evidenceFileName kosong).');
    return;
  }
  for (const fileId of fileIds) {
    console.log(`[OneDrive] Mencoba menghapus berkas dengan id: ${fileId}`);
    try {
      await onedrive.deleteFile(fileId);
      console.log(`[OneDrive] Berhasil menghapus berkas dengan id: ${fileId}`);
    } catch (err) {
      // Jangan gagalkan proses hapus data hanya karena berkas di OneDrive
      // bermasalah (mis. sudah dihapus manual sebelumnya) — cukup dicatat.
      console.error(`[OneDrive] Gagal menghapus berkas OneDrive terkait (${fileId}):`, err.message);
    }
  }
}

function getCollectionRoute(name, schema, options = {}) {
  const isAdminOnly = options.adminOnly || false;

  app.get(`/api/${name}`, requireLogin, (req, res) => {
    const user = req.session.user;
    res.json(filterItems(DB[name], user, name, req.query));
  });

  app.post(`/api/${name}`, requireLogin, isAdminOnly ? requireAdmin : authorizePayload, validateBody(schema), async (req, res) => {
    const user = req.session.user;
    if (user.role !== 'admin' && req.body.timId && req.body.timId !== user.teamId) {
      return res.status(403).json({ error: 'Tidak dapat menambahkan data untuk tim lain' });
    }
    const item = { id: uuidv4(), ...req.body };
    DB[name].push(item);
    await auditLog(user, 'create', name, item.id, { item });
    await saveDB();
    res.json(item);
  });

  app.put(`/api/${name}/:id`, requireLogin, isAdminOnly ? requireAdmin : authorizePayload, validateBody(schema), async (req, res) => {
    const user = req.session.user;
    const idx = DB[name].findIndex((x) => x.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Data tidak ditemukan' });
    if (user.role !== 'admin' && DB[name][idx].timId !== user.teamId) {
      return res.status(403).json({ error: 'Tidak dapat mengubah data tim lain' });
    }
    if (user.role !== 'admin' && req.body.timId && req.body.timId !== user.teamId) {
      return res.status(403).json({ error: 'Tidak dapat mengubah data untuk tim lain' });
    }
    const before = DB[name][idx];
    DB[name][idx] = { ...before, ...req.body, id: req.params.id };
    await auditLog(user, 'update', name, req.params.id, { newValue: DB[name][idx] });
    await saveDB();
    res.json(DB[name][idx]);
    // Kalau berkas bukti diganti dengan yang baru, hapus berkas LAMA di
    // OneDrive yang sudah tidak terpakai lagi (dijalankan setelah respons
    // dikirim supaya tidak menghambat UI). Berkas yang masih dipakai (id-nya
    // sama persis di data baru) tidak akan ikut terhapus.
    const oldIds = collectEvidenceFileIds(name, before);
    const newIds = new Set(collectEvidenceFileIds(name, DB[name][idx]));
    const orphanedIds = oldIds.filter((id) => !newIds.has(id));
    deleteEvidenceFiles(orphanedIds);
  });

  app.delete(`/api/${name}/:id`, requireLogin, isAdminOnly ? requireAdmin : authorizePayload, async (req, res) => {
    const user = req.session.user;
    const idx = DB[name].findIndex((x) => x.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Data tidak ditemukan' });
    if (user.role !== 'admin' && DB[name][idx].timId !== user.teamId) {
      return res.status(403).json({ error: 'Tidak dapat menghapus data tim lain' });
    }
    const deleted = DB[name].splice(idx, 1)[0];
    await auditLog(user, 'delete', name, req.params.id, { deleted });
    await saveDB();
    res.json({ deleted: true });
    // Hapus berkas bukti terkait di OneDrive setelah respons dikirim,
    // supaya user tidak menunggu proses ini (tidak menghambat UI).
    deleteEvidenceFiles(collectEvidenceFileIds(name, deleted));
  });
}

getCollectionRoute('teams', teamSchema, { adminOnly: true });
getCollectionRoute('fra', fraSchema);
getCollectionRoute('kegiatan', kegiatanSchema);
getCollectionRoute('tugas', tugasSchema);
getCollectionRoute('iku', ikuSchema);

app.get('/api/audit', requireLogin, requireAdmin, (req, res) => {
  res.json(DB.auditLogs || []);
});

app.get('/api/reminders', requireLogin, (req, res) => {
  res.json(computeReminders(req.session.user));
});

app.get('/api/export/json', requireLogin, (req, res) => {
  const user = req.session.user;
  const collection = req.query.collection;
  const targetCollections = collection ? [collection] : ['teams', 'fra', 'kegiatan', 'tugas', 'iku'];
  const payload = {};
  targetCollections.forEach((name) => {
    if (DB[name]) payload[name] = filterItems(DB[name], user, name, req.query);
  });
  const filename = collection ? `simonev-${collection}.json` : 'simonev-all.json';
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.setHeader('Content-Type', 'application/json');
  res.send(JSON.stringify(payload, null, 2));
});

app.get('/api/export/excel', requireLogin, async (req, res) => {
  const user = req.session.user;
  const collection = req.query.collection;
  const targetCollections = collection ? [collection] : ['teams', 'fra', 'kegiatan', 'tugas', 'iku'];
  const workbook = new ExcelJS.Workbook();
  for (const name of targetCollections) {
    if (!DB[name]) continue;
    const items = filterItems(DB[name], user, name, req.query);
    if (!items.length) continue;
    const sheet = workbook.addWorksheet(name);
    sheet.columns = Object.keys(items[0]).map((key) => ({ header: key, key, width: 20 }));
    items.forEach((item) => sheet.addRow(item));
  }
  const filename = collection ? `simonev-${collection}.xlsx` : 'simonev-all.xlsx';
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  await workbook.xlsx.write(res);
  res.end();
});

app.get('/api/export/pdf', requireLogin, (req, res) => {
  const user = req.session.user;
  const collection = req.query.collection || 'tugas';
  const items = filterItems(DB[collection] || [], user, collection, req.query);
  const doc = new PDFDocument({ size: 'A4', margin: 40 });
  res.setHeader('Content-Disposition', `attachment; filename=simonev-${collection}.pdf`);
  res.setHeader('Content-Type', 'application/pdf');
  doc.pipe(res);
  doc.fontSize(18).text(`Laporan ${collection.toUpperCase()}`, { underline: true });
  doc.moveDown(0.5);
  doc.fontSize(12).text(`Dibuat: ${new Date().toLocaleString('id-ID')}`);
  doc.moveDown(1);
  items.forEach((item, idx) => {
    doc.fontSize(12).text(`${idx + 1}. ${item.nama || item.periode || item.kode || item.name || 'Data'}`);
    doc.fontSize(10).text(JSON.stringify(item, null, 2));
    doc.moveDown(0.5);
  });
  doc.end();
});

app.get('/api/onedrive/status', requireLogin, requireAdmin, (req, res) => {
  res.json({ connected: onedrive.isConnected() });
});

app.get('/api/onedrive/connect', requireLogin, requireAdmin, (req, res) => {
  try {
    res.redirect(onedrive.getAuthUrl());
  } catch (err) {
    res.status(500).send(`Gagal memulai koneksi OneDrive: ${err.message}`);
  }
});

app.get('/auth/redirect', requireLogin, requireAdmin, async (req, res) => {
  const { code, error, error_description: errorDescription } = req.query;
  if (error) {
    return res.status(400).send(`Login OneDrive gagal: ${errorDescription || error}`);
  }
  if (!code) {
    return res.status(400).send('Kode otorisasi tidak ditemukan pada callback OneDrive.');
  }
  try {
    await onedrive.exchangeCodeForTokens(code);
    res.send('<h2>OneDrive berhasil terhubung.</h2><p>Anda bisa menutup tab ini dan kembali ke aplikasi SIMONEV.</p>');
  } catch (err) {
    console.error('Gagal menghubungkan OneDrive:', err.message);
    res.status(500).send(`Gagal menghubungkan OneDrive: ${err.message}`);
  }
});

// Pemetaan modul -> nama folder OneDrive. Kalau kind tidak dikenali/tidak
// dikirim, jatuh ke folder default (ONEDRIVE_FOLDER di .env / 'SIMONEV-Uploads').
const UPLOAD_FOLDER_BY_KIND = {
  fra: 'SIMONEV-FRA',
  iku: 'SIMONEV-IKU',
  tugas: 'SIMONEV-Tugas',
  kegiatan: 'SIMONEV-Kegiatan',
};

app.post('/api/upload', requireLogin, (req, res) => {
  upload.single('file')(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'Tidak ada berkas yang dikirim' });
    try {
      const originalName = sanitizeText(req.file.originalname);
      const ext = path.extname(req.file.originalname).toLowerCase();
      const baseName = path.basename(originalName, ext)
      .replace(/[^a-zA-Z0-9-_ ]/g, '')   // buang karakter yang tidak aman untuk nama file
      .trim()
      .slice(0, 100);                    // batasi panjang biar tidak kepanjangan
      const shortId = uuidv4().split('-')[0]; // ambil 8 karakter pertama UUID saja
      const storedName = `${baseName}-${shortId}${ext}`;
      const kind = sanitizeText(req.body.kind || '');
      const folder = UPLOAD_FOLDER_BY_KIND[kind]; // undefined -> onedrive.js pakai default
      console.log(`[Upload] req.body.kind = "${kind}" -> folder tujuan: ${folder || '(default) ' + (process.env.ONEDRIVE_FOLDER || 'SIMONEV-Uploads')}`);
      const item = await onedrive.uploadFile(req.file.buffer, storedName, folder);
      // "filename" yang dikembalikan sekarang adalah ID item OneDrive
      // (bukan lagi nama file di disk lokal), dipakai untuk lihat/hapus berkas.
      res.json({ filename: item.id, originalName, size: req.file.size, mimetype: req.file.mimetype });
    } catch (uploadErr) {
      console.error('Gagal mengunggah ke OneDrive:', uploadErr.message);
      res.status(502).json({ error: `Gagal mengunggah berkas ke OneDrive: ${uploadErr.message}` });
    }
  });
});

app.get('/api/files/:filename', requireLogin, async (req, res) => {
  try {
    const itemId = req.params.filename;
    const info = await onedrive.getDownloadUrl(itemId);
    if (!info || !info.downloadUrl) return res.status(404).json({ error: 'Berkas tidak ditemukan' });
    // Redirect ke link unduh sementara dari OneDrive (pre-authenticated, berlaku singkat).
    res.redirect(info.downloadUrl);
  } catch (err) {
    console.error('Gagal mengambil berkas dari OneDrive:', err.message);
    res.status(502).json({ error: `Gagal mengambil berkas dari OneDrive: ${err.message}` });
  }
});

app.delete('/api/files/:filename', requireLogin, async (req, res) => {
  try {
    const itemId = req.params.filename;
    await onedrive.deleteFile(itemId);
    res.json({ deleted: true });
  } catch (err) {
    console.error('Gagal menghapus berkas di OneDrive:', err.message);
    res.status(502).json({ error: `Gagal menghapus berkas di OneDrive: ${err.message}` });
  }
});

app.get('*', (req, res) => {
  if (req.path.startsWith('/api/')) return res.status(404).json({ error: 'Endpoint tidak ditemukan' });
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`✓ SIMONEV BPS berjalan di http://localhost:${PORT}`);
  console.log(`  Data JSON   : ${DB_FILE}`);
  console.log(`  Berkas bukti: OneDrive (${process.env.ONEDRIVE_USER || 'ONEDRIVE_USER belum diisi di .env'} / ${process.env.ONEDRIVE_FOLDER || 'SIMONEV-Uploads'})`);
});