/**
 * SIMORA BPS — Sistem Monitoring & Evaluasi Kinerja
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
const bcrypt = require('bcryptjs');
const { tanyaAI } = require('./groqService');

const app = express();
const PORT = process.env.PORT || 3000;
const SESSION_SECRET = process.env.SESSION_SECRET || 'simora-session-secret';

const DATA_DIR = path.join(__dirname, 'data');
const BACKUP_DIR = path.join(DATA_DIR, 'backup');
const DB_FILE = path.join(DATA_DIR, 'db.json');
const MAX_FILE_MB = 10;

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

// Hapus karakter kontrol yang tidak valid di XML 1.0 (penyebab Excel minta "recover").
// Tab, newline (\n), dan carriage return (\r) tetap dibiarkan karena itu valid.
function sanitizeExcelText(value) {
  if (typeof value !== 'string') return value;
  return value.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
}
function sanitizeExcelRow(row) {
  const clean = {};
  Object.keys(row).forEach((key) => {
    clean[key] = sanitizeExcelText(row[key]);
  });
  return clean;
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
    users: [],
    iku: [],
    fra: [],
    kegiatan: [],
    tugas: [],
    auditLogs: [],
    notifications: [],
    ikuTargets: [],
    settings: { capaianGlobal: 120 },
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
DB.users = DB.users || [];
DB.notifications = DB.notifications || [];
DB.ikuTargets = DB.ikuTargets || [];
DB.settings = DB.settings || { capaianGlobal: 120 };
if (DB.settings.capaianGlobal === undefined || DB.settings.capaianGlobal === null) DB.settings.capaianGlobal = 120;
seedDefaultUsersIfEmpty();
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

// BARU: cek berkas yang mungkin sudah dihapus manual di OneDrive, setiap 20 menit.
syncEvidenceFilesWithOneDrive();
setInterval(syncEvidenceFilesWithOneDrive, 20 * 60 * 1000);

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
        subject: `[SIMORA] Pengingat: realisasi ${tasks[0].quarterLabel} belum diisi`,
        text: `Halo ${timName},\n\nBerikut tugas yang belum diisi realisasinya untuk triwulan berjalan (H-${daysLeft}):\n\n${list}\n\nMohon segera dilengkapi di SIMORA.\n\n— Sistem SIMORA BPS`,
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

function findUserByLogin(loginInput) {
  const needle = String(loginInput || '').trim().toLowerCase();
  return (DB.users || []).find(
    (u) => u.username.toLowerCase() === needle || u.email.toLowerCase() === needle
  );
}

function hasAnyUser() {
  return Array.isArray(DB.users) && DB.users.length > 0;
}

function sanitizeUserForClient(u) {
  const { passwordHash, ...rest } = u;
  return rest;
}

async function seedDefaultUsersIfEmpty() {
  if (DB.users && DB.users.length) return;
  const defaults = [
    { username: 'admin', email: 'admin@bps.go.id', password: 'admin123', role: 'admin' },
    { username: 'operator_sosial', email: 'sosial@bps.go.id', password: 'opsosial1', role: 'operator', teamName: 'Tim Statistik Sosial' },
    { username: 'operator_produksi', email: 'produksi@bps.go.id', password: 'opprod1', role: 'operator', teamName: 'Tim Statistik Produksi' },
    { username: 'operator_distribusi', email: 'distribusi@bps.go.id', password: 'opdist1', role: 'operator', teamName: 'Tim Statistik Distribusi' },
    { username: 'operator_neraca', email: 'neraca@bps.go.id', password: 'opneraca1', role: 'operator', teamName: 'Tim Neraca & Analisis Statistik' },
    { username: 'operator_ipds', email: 'ipds@bps.go.id', password: 'opipds1', role: 'operator', teamName: 'Tim IPDS (Integrasi Pengolahan & Diseminasi Statistik)' },
    { username: 'operator_tu', email: 'tu@bps.go.id', password: 'optu1', role: 'operator', teamName: 'Tim Tata Usaha' },
  ];
  DB.users = [];
  for (const d of defaults) {
    const passwordHash = await bcrypt.hash(d.password, 10);
    DB.users.push({
      id: uuidv4(),
      username: d.username,
      email: d.email,
      passwordHash,
      role: d.role,
      teamId: d.teamName ? findTeamIdByName(d.teamName) : null,
    });
  }
  await saveDB();
  console.log('Akun default (admin & 6 operator) dibuat otomatis. Segera ganti password default lewat menu Kelola User.');
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
    const target = Number(item.target) || 0;
    const realisasi = Number(item.realisasi) || 0;
    if (!target) return 'Tertinggal';
    const pct = Math.round((realisasi / target) * 100);
    if (pct >= 100) return 'Tercapai';
    if (pct >= 60) return 'Perlu Perhatian';
    return 'Tertinggal';
  }
  if (collection === 'fra') {
    const pct = Number(item.persentase) || 0;
    if (pct >= 100) return 'Tercapai';
    if (pct >= 60) return 'Perlu Perhatian';
    return 'Tertinggal';
  }
  return item.status || '—';
}
/* ===================== Helper export PDF (cover + tabel) ===================== */
const LOGO_PATH = path.join(__dirname, 'public', 'assets', 'logo-bps.png');
const TRIWULAN_LABELS_SERVER = { q1: 'Triwulan I', q2: 'Triwulan II', q3: 'Triwulan III', q4: 'Triwulan IV' };

const PDF_TITLES = {
  fra: 'Laporan FRA — Capaian Kinerja Tim',
  kegiatan: 'Laporan Analisis Kegiatan',
  tugas: 'Laporan Tugas Tim & Target Tahunan',
  iku: 'Laporan Indikator Kinerja Utama (IKU)',
  teams: 'Laporan Data Tim',
};
const EXPORT_VIEW_META = {
  pencapaian: {
    filename: 'simora-tl-tw-sebelumnya',
    title: 'Laporan TL TW Sebelumnya — Bukti Tindak Lanjut',
  },
};
function resolveExportMeta(collection, view) {
  if (view && EXPORT_VIEW_META[view]) return EXPORT_VIEW_META[view];
  return {
    filename: collection ? `simora-${collection}` : 'simora-all',
    title: PDF_TITLES[collection] || `Laporan ${String(collection || '').toUpperCase()}`,
  };
}

const PDF_COLUMNS = {
  teams: [{ key: 'name', label: 'Nama Tim', width: 480 }],
  fra: [
    { key: 'team', label: 'Tim', width: 90 },
    { key: 'iku', label: 'IKU', width: 45 },
    { key: 'periode', label: 'Triwulan', width: 65 },
    { key: 'target', label: 'Target', width: 55 },
    { key: 'realisasi', label: 'Realisasi', width: 60 },
    { key: 'persen', label: 'Capaian', width: 55 },
    { key: 'status', label: 'Status', width: 65 },
  ],
  kegiatan: [
    { key: 'team', label: 'Tim', width: 75 },
    { key: 'iku', label: 'IKU', width: 35 },
    { key: 'nama', label: 'Kegiatan', width: 100 },
    { key: 'kendala', label: 'Kendala', width: 85 },
    { key: 'solusi', label: 'Solusi', width: 85 },
    { key: 'rtl', label: 'RTL', width: 60 },
    { key: 'status', label: 'Status', width: 60 },
  ],
  tugas: [
    { key: 'team', label: 'Tim', width: 75 },
    { key: 'nama', label: 'Nama Tugas', width: 110 },
    { key: 'triwulan', label: 'Triwulan', width: 55 },
    { key: 'target', label: 'Target', width: 55 },
    { key: 'realisasi', label: 'Realisasi', width: 60 },
    { key: 'persen', label: '%', width: 35 },
    { key: 'status', label: 'Status', width: 70 },
  ],
  iku: [
    { key: 'kode', label: 'Kode', width: 40 },
    { key: 'nama', label: 'Indikator', width: 135 },
    { key: 'team', label: 'Tim', width: 90 },
    { key: 'target', label: 'Target', width: 55 },
    { key: 'capaian', label: 'Capaian', width: 55 },
    { key: 'persen', label: '%', width: 35 },
    { key: 'status', label: 'Status', width: 65 },
  ],
};

function teamNameServer(id) {
  const t = DB.teams.find((x) => x.id === id);
  return t ? t.name : '—';
}
/* ===================== Export sesuai tampilan per-navbar ===================== */
const FRA_IKU_LIST = [
  { no: '1.1.1.1', nama: 'Persentase Publikasi/Laporan Statistik Kependudukan dan Ketenagakerjaan yang Berkualitas' },
  { no: '1.1.3.1', nama: 'Persentase Publikasi/Laporan Statistik Kesejahteraan Rakyat yang Berkualitas' },
  { no: '1.1.5.1', nama: 'Persentase Publikasi/Laporan Statistik Ketahanan Sosial yang Berkualitas' },
  { no: '1.2.1.1', nama: 'Persentase Publikasi/Laporan Statistik Tanaman Pangan, Hortikultura, dan Perkebunan yang Berkualitas' },
  { no: '1.2.2.1', nama: 'Persentase Publikasi/Laporan Statistik Peternakan, Perikanan, dan Kehutanan yang Berkualitas' },
  { no: '1.2.3.1', nama: 'Persentase Publikasi/Laporan Statistik Industri yang Berkualitas' },
  { no: '1.3.1.1', nama: 'Persentase Publikasi/Laporan Statistik Distribusi yang Berkualitas' },
  { no: '1.3.3.1', nama: 'Persentase Publikasi/Laporan Statistik Harga yang Berkualitas' },
  { no: '1.3.5.1', nama: 'Persentase Publikasi/Laporan Statistik Keuangan, Teknologi Informasi, dan Pariwisata yang Berkualitas' },
  { no: '1.4.1.1', nama: 'Persentase Publikasi/Laporan Neraca Produksi yang Berkualitas' },
  { no: '1.4.1.2', nama: 'Persentase Publikasi/Laporan Neraca Pengeluaran yang Berkualitas' },
  { no: '1.4.1.3', nama: 'Persentase Publikasi/Laporan Analisis Statistik dan Neraca Satelit yang Berkualitas' },
  { no: '2.1.4.1', nama: 'Persentase Kumulatif Desa Yang Berpredikat Desa Cinta Statistik' },
  { no: '2.5.1.1', nama: 'Tingkat Penyelenggaraan Pembinaan Statistik Sektoral sesuai standar' },
  { no: '2.7.1.1', nama: 'Indeks Pelayanan Publik - Penilaian mandiri' },
  { no: '3.2.4.1', nama: 'Nilai SAKIP oleh Inspektorat' },
  { no: '3.2.4.2', nama: 'Indeks Implementasi BerAKHLAK' },
];
function fraIkuNamaByNoServer(no) {
  const item = FRA_IKU_LIST.find(x => x.no === no);
  return item ? item.nama : '';
}
function ikuNamaByNoAnyServer(no) {
  const known = fraIkuNamaByNoServer(no);
  if (known) return known;
  const custom = (DB.ikuTargets || []).find(t => t.ikuNomor === no && t.ikuNama);
  return custom ? custom.ikuNama : '';
}
function compareIkuNomorServer(a, b) {
  const pa = String(a || '').split('.').map(n => parseInt(n, 10));
  const pb = String(b || '').split('.').map(n => parseInt(n, 10));
  const len = Math.max(pa.length, pb.length);
  for (let i = 0; i < len; i++) {
    const va = isNaN(pa[i]) ? 0 : pa[i];
    const vb = isNaN(pb[i]) ? 0 : pb[i];
    if (va !== vb) return va - vb;
  }
  return 0;
}

/* ---------- FRA: grouping per Tim + No IKU, sama seperti tabel di halaman FRA ---------- */
function groupFraByTeamIkuServer(items) {
  const map = {};
  items.forEach((f) => {
    const key = f.timId + '|' + f.ikuNomor;
    if (!map[key]) map[key] = { timId: f.timId, ikuNomor: f.ikuNomor, target: Number(f.target) || 0, quarters: { 1: null, 2: null, 3: null, 4: null } };
    const q = periodeToQuarterNum(f.periode);
    if (q) {
      map[key].quarters[q] = f;
      if (Number(f.target)) map[key].target = Number(f.target);
    }
  });
  return Object.values(map);
}
function fraStatusLabelServer(pct) {
  if (pct >= 100) return 'Tercapai';
  if (pct >= 60) return 'Perlu Perhatian';
  return 'Tertinggal';
}
function groupFraStatusServer(g) {
  const filled = [1, 2, 3, 4].map((q) => g.quarters[q]).filter(Boolean);
  if (!filled.length) return '—';
  const avg = Math.round(filled.reduce((s, f) => s + (Number(f.persentase) || 0), 0) / filled.length);
  return fraStatusLabelServer(avg);
}
function buildFraExportRows(user, query) {
  const tahun = query.tahun ? Number(query.tahun) : null;
  let fraItems = filterItems(DB.fra, user, 'fra', query).map(enrichFraWithLiveTarget);
  if (query.teamId) fraItems = fraItems.filter((f) => f.timId === query.teamId);
  let groups = groupFraByTeamIkuServer(fraItems);

  // baris "kosong" untuk Target IKU yang sudah diatur tapi belum ada capaian sama sekali —
  // supaya ikut tampil di export, sama seperti di halaman web (bukan hilang).
  const targetsYear = (DB.ikuTargets || []).filter(
    (t) => (!tahun || Number(t.tahun) === tahun) && (!query.teamId || t.timId === query.teamId)
  );
  targetsYear.forEach((t) => {
    const existing = groups.find((g) => g.timId === t.timId && g.ikuNomor === t.ikuNomor);
    if (!existing) groups.push({ timId: t.timId, ikuNomor: t.ikuNomor, target: Number(t.target) || 0, quarters: { 1: null, 2: null, 3: null, 4: null } });
    else if (!existing.target) existing.target = Number(t.target) || 0;
  });

  if (query.status) {
    groups = groups.filter((g) => groupFraStatusServer(g).toLowerCase() === String(query.status).toLowerCase());
  }
  groups.sort((a, b) => compareIkuNomorServer(a.ikuNomor, b.ikuNomor) || teamNameServer(a.timId).localeCompare(teamNameServer(b.timId)));

  const qLabels = { 1: 'I', 2: 'II', 3: 'III', 4: 'IV' };
  return groups.map((g) => {
    const totalRealisasi = [1, 2, 3, 4].reduce((s, q) => s + (g.quarters[q] ? Number(g.quarters[q].realisasi) || 0 : 0), 0);
    const totalCapaian = g.target ? Math.min(120, Math.round((totalRealisasi / g.target) * 10000) / 100) : 0;
    const row = {
      ikuNomor: g.ikuNomor || '-',
      iku: ikuNamaByNoAnyServer(g.ikuNomor) || '-',
      team: teamNameServer(g.timId),
      targetTahunan: String(g.target || 0),
      realisasiTahunan: String(totalRealisasi),
      capaianTahunan: `${totalCapaian}%`,
      status: groupFraStatusServer(g),
    };
    [1, 2, 3, 4].forEach((q) => {
      const entry = g.quarters[q];
      const record = findIkuTargetRecord(g.ikuNomor, tahun || (entry ? entry.tahun : new Date().getFullYear()), g.timId);
      const qTarget = entry ? Number(entry.target) || 0 : (record ? Number(record['targetTw' + q]) || 0 : 0);
      row['targetTw' + q] = qTarget ? String(qTarget) : '-';
      row['realisasiTw' + q] = entry ? String(Number(entry.realisasi) || 0) : '-';
      row['capaianTw' + q] = entry ? `${Number(entry.persentase) || 0}%` : '-';
    });
    return row;
  });
}
const FRA_EXPORT_COLUMNS = [
  { key: 'ikuNomor', label: 'No IKU', width: 38 },
  { key: 'iku', label: 'IKU', width: 100 },
  { key: 'team', label: 'Tim', width: 50 },
  { key: 'targetTahunan', label: 'Target Th.', width: 36 },
  { key: 'realisasiTahunan', label: 'Realisasi Th.', width: 36 },
  { key: 'capaianTahunan', label: 'Capaian Th.', width: 34 },
  { key: 'targetTw1', label: 'Target TW I', width: 32 },
  { key: 'targetTw2', label: 'Target TW II', width: 32 },
  { key: 'targetTw3', label: 'Target TW III', width: 32 },
  { key: 'targetTw4', label: 'Target TW IV', width: 32 },
  { key: 'realisasiTw1', label: 'Real. TW I', width: 32 },
  { key: 'realisasiTw2', label: 'Real. TW II', width: 32 },
  { key: 'realisasiTw3', label: 'Real. TW III', width: 32 },
  { key: 'realisasiTw4', label: 'Real. TW IV', width: 32 },
  { key: 'capaianTw1', label: 'Cap. TW I', width: 32 },
  { key: 'capaianTw2', label: 'Cap. TW II', width: 32 },
  { key: 'capaianTw3', label: 'Cap. TW III', width: 32 },
  { key: 'capaianTw4', label: 'Cap. TW IV', width: 32 },
  { key: 'status', label: 'Status', width: 32 },
];

/* ---------- Kegiatan & TL TW Sebelumnya: grid No IKU x Tim x Triwulan, sama seperti halaman web ---------- */
const KEGIATAN_PERIODE_LABEL = { 1: 'Triwulan I', 2: 'Triwulan II', 3: 'Triwulan III', 4: 'Triwulan IV' };
function buildKegiatanGridServer(tahun, triwulanFilter) {
  const quarters = triwulanFilter ? [Number(triwulanFilter)] : [1, 2, 3, 4];
  let targets = (DB.ikuTargets || []).filter((t) => Number(t.tahun) === tahun);
  const items = [];
  const usedIds = new Set();
  targets.forEach((t) => {
    quarters.forEach((q) => {
      const periode = KEGIATAN_PERIODE_LABEL[q];
      const match = DB.kegiatan.find(
        (k) => k.timId === t.timId && k.ikuNomor === t.ikuNomor && (!k.tahun || Number(k.tahun) === tahun) && k.periode === periode
      );
      if (match) {
        usedIds.add(match.id);
        items.push(match);
      } else {
        items.push({
          __virtual: true, timId: t.timId, ikuNomor: t.ikuNomor, tahun, periode,
          nama: '', kendala: '', solusi: '', rtl: '', catatan: '', catatanTl: '',
          picTindakLanjut: '', batasWaktuTindakLanjut: '', status: 'Belum Ditindaklanjuti',
          hasEvidence: false, evidenceFiles: [], hasTlEvidence: false, tlEvidenceFiles: [],
        });
      }
    });
  });
  // data lama yang No IKU/Tim-nya belum terdaftar di Kelola Target IKU tetap ikut tampil
  DB.kegiatan.forEach((k) => {
    if (usedIds.has(k.id)) return;
    if (k.tahun && Number(k.tahun) !== tahun) return;
    items.push(k);
  });
  return items;
}
function baseKegiatanExportItems(query) {
  const tahun = query.tahun ? Number(query.tahun) : new Date().getFullYear();
  let items = buildKegiatanGridServer(tahun, query.triwulan);
  if (query.teamId) items = items.filter((k) => k.timId === query.teamId);
  if (query.status) items = items.filter((k) => String(k.status || 'Belum Ditindaklanjuti').toLowerCase() === String(query.status).toLowerCase());
  items.sort(
    (a, b) =>
      compareIkuNomorServer(a.ikuNomor, b.ikuNomor) ||
      teamNameServer(a.timId).localeCompare(teamNameServer(b.timId)) ||
      (periodeToQuarterNum(a.periode) || 0) - (periodeToQuarterNum(b.periode) || 0)
  );
  return items;
}
function buildKegiatanExportRows(user, query) {
  return baseKegiatanExportItems(query).map((k) => {
    const hasEv = !!(k.hasEvidence || (k.evidenceFiles && k.evidenceFiles.length));
    return {
      ikuNomor: k.ikuNomor || '-',
      iku: ikuNamaByNoAnyServer(k.ikuNomor) || '-',
      team: teamNameServer(k.timId),
      nama: k.nama || '-',
      periode: k.periode || '-',
      kendala: k.kendala || '-',
      solusi: k.solusi || '-',
      rtl: k.rtl || '-',
      pic: k.picTindakLanjut || '-',
      batasWaktu: k.batasWaktuTindakLanjut ? formatDateIndo(k.batasWaktuTindakLanjut) : '-',
      buktiDukung: hasEv ? 'Ada' : 'Belum ada',
      status: k.status || 'Belum Ditindaklanjuti',
      catatan: k.catatan || '-',
    };
  });
}
const KEGIATAN_EXPORT_COLUMNS = [
  { key: 'ikuNomor', label: 'No IKU', width: 38 },
  { key: 'iku', label: 'IKU', width: 82 },
  { key: 'team', label: 'Tim', width: 48 },
  { key: 'nama', label: 'Kegiatan', width: 72 },
  { key: 'periode', label: 'Triwulan', width: 40 },
  { key: 'kendala', label: 'Kendala', width: 65 },
  { key: 'solusi', label: 'Solusi', width: 65 },
  { key: 'rtl', label: 'RTL', width: 55 },
  { key: 'pic', label: 'PIC TL', width: 40 },
  { key: 'batasWaktu', label: 'Batas Waktu TL', width: 42 },
  { key: 'buktiDukung', label: 'Bukti Dukung', width: 38 },
  { key: 'status', label: 'Status', width: 42 },
  { key: 'catatan', label: 'Catatan', width: 55 },
];

function buildPencapaianExportRows(user, query) {
  return baseKegiatanExportItems(query).map((k) => {
    const hasTl = !!(k.hasTlEvidence || (k.tlEvidenceFiles && k.tlEvidenceFiles.length));
    return {
      ikuNomor: k.ikuNomor || '-',
      iku: ikuNamaByNoAnyServer(k.ikuNomor) || '-',
      team: teamNameServer(k.timId),
      periode: k.periode || '-',
      rtl: k.rtl || '-',
      status: k.status || 'Belum Ditindaklanjuti',
      buktiDukung: hasTl ? 'Ada' : 'Belum ada',
      catatan: k.catatanTl || '-',
    };
  });
}
const PENCAPAIAN_EXPORT_COLUMNS = [
  { key: 'ikuNomor', label: 'No IKU', width: 40 },
  { key: 'iku', label: 'IKU', width: 140 },
  { key: 'team', label: 'Tim', width: 60 },
  { key: 'periode', label: 'Triwulan', width: 50 },
  { key: 'rtl', label: 'RTL', width: 150 },
  { key: 'status', label: 'Status', width: 55 },
  { key: 'buktiDukung', label: 'Bukti Dukung', width: 50 },
  { key: 'catatan', label: 'Catatan', width: 100 },
];

function buildExportDataset(user, query) {
  const collection = query.collection;
  const view = query.view;
  if (collection === 'fra') return { columns: FRA_EXPORT_COLUMNS, rows: buildFraExportRows(user, query) };
  if (collection === 'kegiatan' && view === 'pencapaian') return { columns: PENCAPAIAN_EXPORT_COLUMNS, rows: buildPencapaianExportRows(user, query) };
  if (collection === 'kegiatan') return { columns: KEGIATAN_EXPORT_COLUMNS, rows: buildKegiatanExportRows(user, query) };
  // fallback untuk koleksi lain (tugas/iku/teams) — pakai builder lama
  let items = filterItems(DB[collection] || [], user, collection, query);
  if (collection === 'fra') items = items.map(enrichFraWithLiveTarget);
  const columns = PDF_COLUMNS[collection] || PDF_COLUMNS.teams;
  const rows = items.map((item) => buildPdfRow(collection, item));
  return { columns, rows };
}
function formatDateIndo(dateStr) {
  if (!dateStr) return '-';
  const d = /T/.test(dateStr) ? new Date(dateStr) : new Date(`${dateStr}T00:00:00`);
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', timeZone: 'Asia/Jakarta' });
}

function buildPdfRow(collection, item) {
  if (collection === 'teams') return { name: item.name || '-' };
  const teamNm = teamNameServer(item.timId || item.id);
  if (collection === 'fra') {
    return {
      team: teamNm,
      iku: item.ikuNomor ? `IKU ${item.ikuNomor}` : '-',
      periode: item.periode || '-',
      target: String(Number(item.target) || 0),
      realisasi: String(Number(item.realisasi) || 0),
      persen: `${Number(item.persentase) || 0}%`,
      status: getItemStatus('fra', item),
    };
  }
  if (collection === 'kegiatan') {
    return {
      team: teamNm,
      iku: item.ikuNomor ? `IKU ${item.ikuNomor}` : '-',
      nama: item.nama || '-',
      kendala: item.kendala || '-',
      solusi: item.solusi || '-',
      rtl: item.rtl || '-',
      status: item.status || '-',
    };
  }
  if (collection === 'tugas') {
    const target = Number(item.target) || 0;
    const realisasi = Number(item.realisasi) || 0;
    const pct = target ? Math.min(999, Math.round((realisasi / target) * 100)) : 0;
    return {
      team: teamNm,
      nama: item.nama || '-',
      triwulan: TRIWULAN_LABELS_SERVER[item.triwulan] || item.triwulan || '-',
      target: String(target),
      realisasi: String(realisasi),
      persen: `${pct}%`,
      status: getItemStatus('tugas', item),
    };
  }
  if (collection === 'iku') {
    const target = Number(item.target) || 0;
    const capaian = Number(item.capaian) || 0;
    const pct = target ? Math.min(999, Math.round((capaian / target) * 100)) : 0;
    return {
      kode: item.kode || '-',
      nama: item.nama || '-',
      team: teamNm,
      target: String(target),
      capaian: String(capaian),
      persen: `${pct}%`,
      status: getItemStatus('iku', item),
    };
  }
  return {};
}

function drawPdfCover(doc, title) {
  const pageWidth = doc.page.width;
  const logoW = 230;

  // Judul laporan di paling atas.
  let cursorY = 130;
  doc.font('Helvetica-Bold').fontSize(23).fillColor('#0E2A44')
    .text(title, 40, cursorY, { align: 'center', width: pageWidth - 80 });

  // Logo besar, di tengah halaman, di bawah judul.
  cursorY = doc.y + 50;
  if (fs.existsSync(LOGO_PATH)) {
    try {
      // Hitung tinggi ASLI gambar berdasarkan rasio aspek sebenarnya,
      // supaya jarak ke teks di bawahnya presisi (tidak tumpang tindih).
      const img = doc.openImage(LOGO_PATH);
      const logoH = (img.height / img.width) * logoW;
      doc.image(LOGO_PATH, (pageWidth - logoW) / 2, cursorY, { width: logoW });
      cursorY += logoH + 40;
    } catch (e) {
      cursorY += logoW * 0.42 + 40;
    }
  } else {
    cursorY += 40;
  }

  // Waktu cetak (jam:menit saja, tanpa detik), di bawah logo.
  // timeZone dipaksa ke Asia/Jakarta (WIB) supaya konsisten, tidak mengikuti
  // timezone server hosting (Railway biasanya UTC).
  const printedAt = new Date().toLocaleString('id-ID', {
    day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit',
    timeZone: 'Asia/Jakarta',
  });
  doc.font('Helvetica').fontSize(10.5).fillColor('#5B6B78')
    .text(`Dicetak melalui SIMORA BPS pada ${printedAt}`, 40, cursorY, {
      align: 'center', width: pageWidth - 80,
    });

  doc.addPage();
}

function drawPdfTable(doc, columns, rows) {
  const startX = doc.page.margins.left;
  const usableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;

  // BARU: skalakan lebar tiap kolom secara proporsional supaya tabel selalu
  // mengisi penuh lebar halaman (usableWidth) — tidak lagi dipatok angka
  // fixed yang bisa menyisakan ruang kosong di kanan, dan tidak overflow
  // kalau totalnya lebih besar dari halaman.
  const definedTotal = columns.reduce((s, c) => s + c.width, 0);
  const scale = usableWidth / definedTotal;
  const scaledColumns = columns.map((c) => ({ ...c, width: c.width * scale }));

  const padX = 5, padY = 4, headerH = 22;

  function drawHeader(y) {
    doc.rect(startX, y, usableWidth, headerH).fill('#163B5C');
    let x = startX;
    doc.font('Helvetica-Bold').fontSize(8.5).fillColor('#ffffff');
    scaledColumns.forEach((col) => {
      doc.text(col.label, x + padX, y + 7, { width: col.width - padX * 2 });
      x += col.width;
    });
    return y + headerH;
  }

  function rowHeight(row) {
    doc.font('Helvetica').fontSize(8.5);
    let maxH = 18;
    scaledColumns.forEach((col) => {
      const text = String(row[col.key] ?? '-');
      const h = doc.heightOfString(text, { width: col.width - padX * 2 });
      if (h + padY * 2 > maxH) maxH = h + padY * 2;
    });
    return maxH;
  }

  let y = drawHeader(doc.y);
  rows.forEach((row, idx) => {
    const h = rowHeight(row);
    if (y + h > doc.page.height - doc.page.margins.bottom) {
      doc.addPage();
      y = drawHeader(doc.page.margins.top);
    }
    if (idx % 2 === 1) doc.rect(startX, y, usableWidth, h).fill('#F7F8F5');
    doc.strokeColor('#DFE3E0').lineWidth(0.5).rect(startX, y, usableWidth, h).stroke();
    let x = startX;
    doc.font('Helvetica').fontSize(8.5).fillColor('#16232E');
    scaledColumns.forEach((col) => {
      const text = String(row[col.key] ?? '-');
      doc.text(text, x + padX, y + padY, { width: col.width - padX * 2 });
      doc.moveTo(x, y).lineTo(x, y + h).strokeColor('#DFE3E0').lineWidth(0.5).stroke();
      x += col.width;
    });
    doc.moveTo(x, y).lineTo(x, y + h).strokeColor('#DFE3E0').lineWidth(0.5).stroke();
    y += h;
  });
  doc.y = y + 10;
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
  const scoped = filterItems(DB.tugas, user, 'tugas', {}).filter((t) => Number(t.tahun) === year && t.triwulan === quarterKey);
  const reminders = [];
  scoped.forEach((t) => {
    const realisasi = Number(t.realisasi) || 0;
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
  // Nama tim & capaian FRA sengaja dibuat bisa dibaca semua user yang login
  // (walau bukan admin), karena datanya diinput admin dan dipakai untuk
  // monitoring bersama. Pembatasan per tim tetap berlaku untuk koleksi lain
  // (kegiatan, tugas, iku) supaya operator hanya mengelola datanya sendiri.
  const READ_UNRESTRICTED = ['teams', 'fra', 'kegiatan'];
  if (user.role !== 'admin' && !READ_UNRESTRICTED.includes(collection)) {
    result = result.filter((item) => item.timId === user.teamId);
  }
  if (query.teamId) {
    result = result.filter((item) => item.timId === query.teamId || item.id === query.teamId);
  }
  if (query.tahun) {
    // Item yang memang tidak punya field "tahun" (mis. Analisis Kegiatan) tidak ikut
    // difilter berdasarkan tahun — supaya tidak hilang begitu saja saat export.
    result = result.filter((item) => item.tahun === undefined || item.tahun === null || Number(item.tahun) === Number(query.tahun));
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

/* ---------------- Realtime updates (Server-Sent Events) ---------------- */
let sseClients = [];

function broadcastChange(collection, action) {
  const payload = JSON.stringify({ collection, action, ts: Date.now() });
  sseClients.forEach((client) => {
    try {
      client.res.write(`data: ${payload}\n\n`);
    } catch (err) {
      // koneksi sudah putus, nanti dibersihkan otomatis lewat event 'close'
    }
  });
}

// Ping berkala supaya koneksi SSE tidak ditutup paksa oleh proxy/hosting saat idle (mis. Railway).
setInterval(() => {
  sseClients.forEach((client) => {
    try { client.res.write(': ping\n\n'); } catch (err) {}
  });
}, 25000);

const teamSchema = Joi.object({
  name: Joi.string().trim().min(3).max(120).required(),
});
const fraSchema = Joi.object({
  timId: Joi.string().required(),
  periode: Joi.string().trim().min(3).max(120).required(),
  ikuNomor: Joi.string().trim().min(1).max(20).required(),
  namaKegiatan: Joi.string().trim().max(250).allow('', null),
  tanggalKegiatan: Joi.string().isoDate().allow('', null),
  target: Joi.number().min(0.0001).required(),
  realisasi: Joi.number().min(0).required(),
  persentase: Joi.number().min(0).max(120).required(),
  tahun: Joi.number().integer().min(2020).max(2100).required(),
  catatan: Joi.string().trim().max(1000).allow('', null),
  hasEvidence: Joi.boolean().optional(),
  evidenceFiles: Joi.array().items(Joi.object({
    fileName: Joi.string().required(),
    originalName: Joi.string().required(),
    size: Joi.number().min(0).required(),
  })).default([]),
});
const kegiatanSchema = Joi.object({
  timId: Joi.string().required(),
  periode: Joi.string().trim().min(3).max(120).required(),
  nama: Joi.string().trim().max(3000).allow('', null),
  ikuNomor: Joi.string().trim().min(1).max(20).required(),
  kendala: Joi.string().trim().max(1000).allow('', null),
  solusi: Joi.string().trim().max(1000).allow('', null),
  rtl: Joi.string().trim().max(1000).allow('', null),
  catatan: Joi.string().trim().max(1000).allow('', null),        // BARU
  catatanTl: Joi.string().trim().max(1000).allow('', null),      // BARU
  status: Joi.string().valid('Belum Ditindaklanjuti', 'Dalam Proses', 'Selesai').required(),
  picTindakLanjut: Joi.string().trim().max(200).allow('', null),
  batasWaktuTindakLanjut: Joi.string().allow('', null),
  starred: Joi.boolean().optional(),
  tahun: Joi.number().integer().min(2020).max(2100).required(),
  hasEvidence: Joi.boolean().optional(),
  evidenceFiles: Joi.array().items(Joi.object({
    fileName: Joi.string().required(),
    originalName: Joi.string().required(),
    size: Joi.number().min(0).required(),
  })).default([]),
  hasTlEvidence: Joi.boolean().optional(),
  tlEvidenceFiles: Joi.array().items(Joi.object({
    fileName: Joi.string().required(),
    originalName: Joi.string().required(),
    size: Joi.number().min(0).required(),
  })).default([]),
});
const tugasSchema = Joi.object({
  timId: Joi.string().required(),
  tahun: Joi.number().integer().min(2020).max(2100).required(),
  nama: Joi.string().trim().min(3).max(250).required(),
  triwulan: Joi.string().valid('q1','q2','q3','q4').required(),
  target: Joi.number().min(1).required(),
  satuan: Joi.string().trim().min(1).max(50).required(),
  realisasi: Joi.number().min(0).allow(null,''),
  tanggal: Joi.string().allow('', null),
  hasEvidence: Joi.boolean().optional(),
  evidenceFiles: Joi.array().items(Joi.object({
    fileName: Joi.string().required(),
    originalName: Joi.string().required(),
    size: Joi.number().min(0).required(),
  })).default([]),
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
  evidenceFiles: Joi.array().items(Joi.object({
    fileName: Joi.string().required(),
    originalName: Joi.string().required(),
    size: Joi.number().min(0).required(),
  })).default([]),
});

const ikuTargetSchema = Joi.object({
  ikuNomor: Joi.string().trim().min(1).max(20).required(),
  ikuNama: Joi.string().trim().max(250).allow('', null).default(''),   // BARU
  tahun: Joi.number().integer().min(2020).max(2100).required(),
  timId: Joi.string().required(),
  target: Joi.number().min(0).required(),
  targetTw1: Joi.number().min(0).required(),
  targetTw2: Joi.number().min(0).required(),
  targetTw3: Joi.number().min(0).required(),
  targetTw4: Joi.number().min(0).required(),
});

const userCreateSchema = Joi.object({
  username: Joi.string().trim().min(3).max(50).required(),
  email: Joi.string().trim().email().required(),
  password: Joi.string().min(6).required(),
  role: Joi.string().valid('admin', 'operator').required(),
  teamId: Joi.string().allow(null, '').when('role', { is: 'operator', then: Joi.string().required() }),
});
const userUpdateSchema = Joi.object({
  username: Joi.string().trim().min(3).max(50).required(),
  email: Joi.string().trim().email().required(),
  password: Joi.string().min(6).allow('', null),
  role: Joi.string().valid('admin', 'operator').required(),
  teamId: Joi.string().allow(null, '').when('role', { is: 'operator', then: Joi.string().required() }),
});

const profileUpdateSchema = Joi.object({
  username: Joi.string().trim().min(3).max(50).required(),
  email: Joi.string().trim().email().required(),
  password: Joi.string().min(6).allow('', null),
  confirmPassword: Joi.string().allow('', null),
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
  const okExt = ['.pdf', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg', '.tif', '.tiff', '.heic', '.heif'];
  const ext = path.extname(file.originalname).toLowerCase();
  if (okExt.includes(ext)) cb(null, true);
  else cb(new Error('Format berkas harus PDF, Word, atau Gambar (.pdf, .doc, .docx, .jpg, .png, dll)'));
}

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: MAX_FILE_MB * 1024 * 1024 },
});

app.get('/api/setup-status', (req, res) => {
  res.json({ needsSetup: !hasAnyUser() });
});
app.get('/api/teams-public', (req, res) => {
  res.json((DB.teams || []).map(t => ({ id: t.id, name: t.name })));
});

const registerSchema = Joi.object({
  username: Joi.string().trim().min(3).max(50).required(),
  email: Joi.string().trim().email().required(),
  password: Joi.string().min(6).required(),
  teamId: Joi.string().required(),
});

app.post('/api/register', isJsonRequest, validateBody(registerSchema), async (req, res) => {
  const { username, email, password, teamId } = req.body;
  const emailLower = email.toLowerCase();
  const dup = (DB.users || []).find(
    (u) => u.username.toLowerCase() === username.toLowerCase() || u.email.toLowerCase() === emailLower
  );
  if (dup) {
    return res.status(400).json({ error: 'Username atau email sudah terdaftar. Silakan masuk.' });
  }
  const team = DB.teams.find((t) => t.id === teamId);
  if (!team) {
    return res.status(400).json({ error: 'Tim tidak ditemukan' });
  }
  const passwordHash = await bcrypt.hash(password, 10);
  const user = { id: uuidv4(), username, email: emailLower, passwordHash, role: 'operator', teamId };
  DB.users.push(user);
  await saveDB();
  req.session.user = { id: user.id, username: user.username, email: user.email, role: user.role, teamId: user.teamId };
  await auditLog(req.session.user, 'create', 'users', user.id, { username, email: emailLower, role: 'operator', note: 'daftar mandiri' });
  res.json(req.session.user);
});
app.post('/api/setup-first-admin', isJsonRequest, validateBody(Joi.object({
  username: Joi.string().trim().min(3).max(50).required(),
  email: Joi.string().trim().email().required(),
  password: Joi.string().min(6).required(),
})), async (req, res) => {
  if (hasAnyUser()) {
    return res.status(400).json({ error: 'Setup awal sudah pernah dilakukan. Gunakan halaman login biasa.' });
  }
  const { username, email, password } = req.body;
  const emailLower = email.toLowerCase();
  const passwordHash = await bcrypt.hash(password, 10);
  const user = { id: uuidv4(), username, email: emailLower, passwordHash, role: 'admin', teamId: null };
  DB.users = [user];
  await saveDB();
  req.session.user = { id: user.id, username: user.username, email: user.email, role: user.role, teamId: null };
  auditLog(req.session.user, 'create', 'users', user.id, { username, email: emailLower, role: 'admin', note: 'setup awal' });
  res.json(req.session.user);
});

app.post('/api/login', isJsonRequest, validateBody(Joi.object({ username: Joi.string().trim().required(), password: Joi.string().required() })), async (req, res) => {
  const { username, password } = req.body;
  const user = findUserByLogin(username);
  if (!user) {
    return res.status(401).json({ error: 'Username/email atau kata sandi tidak cocok' });
  }
  const valid = await bcrypt.compare(password, user.passwordHash);
  if (!valid) {
    return res.status(401).json({ error: 'Username/email atau kata sandi tidak cocok' });
  }
  req.session.user = { id: user.id, username: user.username, email: user.email, role: user.role, teamId: user.teamId || null };
  auditLog(req.session.user, 'login', 'auth', null, { ip: req.ip });
  res.json(req.session.user);
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

// User mengubah profilnya sendiri (username, email, kata sandi) — bukan role/tim, itu tetap wewenang admin.
app.put('/api/profile', requireLogin, isJsonRequest, validateBody(profileUpdateSchema), async (req, res) => {
  const { username, email, password, confirmPassword } = req.body;
  const emailLower = email.toLowerCase();
  const currentUserId = req.session.user.id;

  const dup = (DB.users || []).find(
    (u) => u.id !== currentUserId && (u.username.toLowerCase() === username.toLowerCase() || u.email.toLowerCase() === emailLower)
  );
  if (dup) return res.status(400).json({ error: 'Username atau email sudah digunakan oleh user lain' });

  if (password) {
    if (password !== confirmPassword) {
      return res.status(400).json({ error: 'Konfirmasi kata sandi baru tidak cocok' });
    }
  }

  const idx = (DB.users || []).findIndex((u) => u.id === currentUserId);
  if (idx === -1) return res.status(404).json({ error: 'User tidak ditemukan' });

  const existing = DB.users[idx];
  const passwordHash = password ? await bcrypt.hash(password, 10) : existing.passwordHash;
  DB.users[idx] = { ...existing, username, email: emailLower, passwordHash };

  req.session.user = { ...req.session.user, username, email: emailLower };
  await auditLog(req.session.user, 'update', 'users', currentUserId, { username, email: emailLower, note: 'ubah profil sendiri' });
  await saveDB();
  res.json(sanitizeUserForClient(DB.users[idx]));
});

app.get('/api/events', requireLogin, (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();
  res.write(': connected\n\n');

  const client = { id: uuidv4(), res };
  sseClients.push(client);

  req.on('close', () => {
    sseClients = sseClients.filter((c) => c.id !== client.id);
  });
});

function collectEvidenceFileIds(name, item) {
  if (!item) return [];
  let ids = [];
  if (Array.isArray(item.evidenceFiles)) {
    ids = ids.concat(item.evidenceFiles.map((f) => f.fileName).filter(Boolean));
  } else if (item.evidenceFileName) {
    // Kompatibel mundur dengan data lama yang masih pakai field tunggal.
    ids.push(item.evidenceFileName);
  }
  if (Array.isArray(item.tlEvidenceFiles)) {
    ids = ids.concat(item.tlEvidenceFiles.map((f) => f.fileName).filter(Boolean));
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
// Cek apakah sebuah berkas masih benar-benar ada di OneDrive.
async function fileExistsOnOneDrive(fileId) {
  try {
    const info = await onedrive.getDownloadUrl(fileId);
    return !!info;
  } catch (err) {
    const msg = String(err.message || '').toLowerCase();
    if (msg.includes('404') || msg.includes('not found') || msg.includes('itemnotfound')) {
      return false; // benar-benar sudah dihapus di OneDrive
    }
    return true; // error lain (jaringan, dsb) — jangan anggap terhapus, biar aman
  }
}

// Hapus satu referensi fileId tertentu dari SEMUA record (fra, kegiatan, tugas, iku)
// yang mengandungnya, lalu simpan & broadcast supaya web langsung ter-update.
async function pruneMissingFileReference(fileId) {
  const collectionsToCheck = ['fra', 'kegiatan', 'tugas', 'iku'];
  let changedAny = false;
  for (const name of collectionsToCheck) {
    const items = DB[name] || [];
    for (const item of items) {
      if (Array.isArray(item.evidenceFiles) && item.evidenceFiles.some((f) => f.fileName === fileId)) {
        item.evidenceFiles = item.evidenceFiles.filter((f) => f.fileName !== fileId);
        item.hasEvidence = item.evidenceFiles.length > 0;
        if (name === 'kegiatan') item.status = computeKegiatanStatusServer(item);
        changedAny = true;
      }
      if (Array.isArray(item.tlEvidenceFiles) && item.tlEvidenceFiles.some((f) => f.fileName === fileId)) {
        item.tlEvidenceFiles = item.tlEvidenceFiles.filter((f) => f.fileName !== fileId);
        item.hasTlEvidence = item.tlEvidenceFiles.length > 0;
        changedAny = true;
      }
    }
  }
  if (changedAny) {
    await saveDB();
    broadcastChange('fra', 'update');
    broadcastChange('kegiatan', 'update');
    broadcastChange('tugas', 'update');
    broadcastChange('iku', 'update');
  }
}

// Sinkronisasi berkala: cek SEMUA berkas bukti dukung yang tercatat, hapus
// referensinya kalau sudah tidak ada lagi di OneDrive (dihapus manual oleh siapa pun).
async function syncEvidenceFilesWithOneDrive() {
  if (!onedrive.isConnected()) return;
  const collectionsToCheck = ['fra', 'kegiatan', 'tugas', 'iku'];
  let totalRemoved = 0;
  for (const name of collectionsToCheck) {
    const items = DB[name] || [];
    for (const item of items) {
      let changed = false;
      if (Array.isArray(item.evidenceFiles) && item.evidenceFiles.length) {
        const kept = [];
        for (const f of item.evidenceFiles) {
          const exists = await fileExistsOnOneDrive(f.fileName);
          if (exists) kept.push(f); else { changed = true; totalRemoved++; }
        }
        if (changed) {
          item.evidenceFiles = kept;
          item.hasEvidence = kept.length > 0;
        }
      }
      if (Array.isArray(item.tlEvidenceFiles) && item.tlEvidenceFiles.length) {
        const kept = [];
        for (const f of item.tlEvidenceFiles) {
          const exists = await fileExistsOnOneDrive(f.fileName);
          if (exists) kept.push(f); else { changed = true; totalRemoved++; }
        }
        if (changed) {
          item.tlEvidenceFiles = kept;
          item.hasTlEvidence = kept.length > 0;
        }
      }
      if (changed && name === 'kegiatan') item.status = computeKegiatanStatusServer(item);
    }
  }
  if (totalRemoved > 0) {
    await saveDB();
    console.log(`[Sync OneDrive] ${totalRemoved} referensi berkas dihapus otomatis (sudah tidak ada di OneDrive).`);
    broadcastChange('fra', 'update');
    broadcastChange('kegiatan', 'update');
    broadcastChange('tugas', 'update');
    broadcastChange('iku', 'update');
  }
}
function hasIsiTeksServer(str) {
  return typeof str === 'string' && /[a-zA-Z]/.test(str);
}
function hasIsiTanggalServer(str) {
  return typeof str === 'string' && str.trim().length > 0;
}
function computeKegiatanStatusServer(body) {
  const textFields = [body.nama, body.kendala, body.solusi, body.rtl, body.picTindakLanjut];
  const filledText = textFields.filter(hasIsiTeksServer).length;
  const filledTanggal = hasIsiTanggalServer(body.batasWaktuTindakLanjut) ? 1 : 0;
  const hasEvidence = !!(body.hasEvidence || (Array.isArray(body.evidenceFiles) && body.evidenceFiles.length));
  const filledCount = filledText + filledTanggal + (hasEvidence ? 1 : 0);
  const total = textFields.length + 2; // 5 field teks + batas waktu + bukti dukung = 7
  if (filledCount === total) return 'Selesai';
  if (filledCount === 0) return 'Belum Ditindaklanjuti';
  return 'Dalam Proses';
}
function getCollectionRoute(name, schema, options = {}) {
  const isAdminOnly = options.adminOnly || false;
  const beforeSave = options.beforeSave || null;
  const enrichForRead = options.enrichForRead || null;

  app.get(`/api/${name}`, requireLogin, (req, res) => {
    const user = req.session.user;
    let items = filterItems(DB[name], user, name, req.query);
    if (enrichForRead) items = items.map(enrichForRead);
    res.json(items);
  });

  app.post(`/api/${name}`, requireLogin, isAdminOnly ? requireAdmin : authorizePayload, validateBody(schema), async (req, res) => {
    const user = req.session.user;
    if (user.role !== 'admin' && req.body.timId && req.body.timId !== user.teamId) {
      return res.status(403).json({ error: 'Tidak dapat menambahkan data untuk tim lain' });
    }
    // createdAt selalu diisi server dengan jam asli saat disimpan (bukan dari input user),
    // dipakai untuk menampilkan "jam upload" yang sebenarnya di frontend.
    // createdBy dicatat sekali saat pertama dibuat, dipakai untuk menargetkan
    // pengingat admin ke user yang benar-benar mengunggah data ini.
    const body = beforeSave ? beforeSave(req.body) : req.body;
    const item = { id: uuidv4(), ...body, createdBy: user.id, createdByUsername: user.username, createdAt: new Date().toISOString() };
    DB[name].push(item);
    await auditLog(user, 'create', name, item.id, { item });
    await saveDB();
    res.json(item);
    broadcastChange(name, 'create');
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
    const body = beforeSave ? beforeSave(req.body) : req.body;
    DB[name][idx] = { ...before, ...body, id: req.params.id };
    await auditLog(user, 'update', name, req.params.id, { newValue: DB[name][idx] });
    await saveDB();
    res.json(DB[name][idx]);
    broadcastChange(name, 'update');
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
    broadcastChange(name, 'delete');
    // Hapus berkas bukti terkait di OneDrive setelah respons dikirim,
    // supaya user tidak menunggu proses ini (tidak menghambat UI).
    deleteEvidenceFiles(collectEvidenceFileIds(name, deleted));
  });
}
function computeFraPersentase(body) {
  const target = Number(body.target) || 0;
  const realisasi = Number(body.realisasi) || 0;
  const persentase = target ? Math.min(120, Math.round((realisasi / target) * 10000) / 100) : 0;
  return { ...body, persentase };
}

const TW_ROMAN_TO_NUM = { i: 1, ii: 2, iii: 3, iv: 4 };
function periodeToQuarterNum(periode) {
  if (!periode) return null;
  const m = /triwulan\s+(iv|iii|ii|i)\b/i.exec(periode);
  return m ? TW_ROMAN_TO_NUM[m[1].toLowerCase()] : null;
}
function findIkuTargetRecord(ikuNomor, tahun, timId) {
  return (DB.ikuTargets || []).find(
    (t) => t.ikuNomor === ikuNomor && Number(t.tahun) === Number(tahun) && t.timId === timId
  );
}
function enrichFraWithLiveTarget(item) {
  const record = findIkuTargetRecord(item.ikuNomor, item.tahun, item.timId);
  if (!record) return item;
  const q = periodeToQuarterNum(item.periode);
  const liveTarget = q ? Number(record['targetTw' + q]) || 0 : Number(record.target) || 0;
  const realisasi = Number(item.realisasi) || 0;
  const persentase = liveTarget ? Math.min(120, Math.round((realisasi / liveTarget) * 10000) / 100) : 0;
  return { ...item, target: liveTarget, persentase };
}

getCollectionRoute('teams', teamSchema, { adminOnly: true });
getCollectionRoute('fra', fraSchema, { adminOnly: true, beforeSave: computeFraPersentase, enrichForRead: enrichFraWithLiveTarget });
getCollectionRoute('kegiatan', kegiatanSchema, {
  beforeSave: (body) => ({ ...body, status: computeKegiatanStatusServer(body) }),
});
getCollectionRoute('tugas', tugasSchema);
getCollectionRoute('iku', ikuSchema);

// Target per No IKU + Tahun — hanya admin yang boleh mengatur, semua user yang
// login boleh membaca (dipakai operator untuk mengunci field Target di form FRA).
app.get('/api/iku-targets', requireLogin, (req, res) => {
  res.json(DB.ikuTargets || []);
});

app.post('/api/iku-targets', requireLogin, requireAdmin, validateBody(ikuTargetSchema), async (req, res) => {
  const { ikuNomor, tahun, timId } = req.body;
  const dup = (DB.ikuTargets || []).find(t => t.ikuNomor === ikuNomor && Number(t.tahun) === Number(tahun) && t.timId === timId);
  if (dup) return res.status(400).json({ error: 'Target untuk No IKU, Tim, dan tahun ini sudah ada. Silakan ubah data yang sudah ada.' });
  const item = { id: uuidv4(), ...req.body };
  DB.ikuTargets.push(item);
  await auditLog(req.session.user, 'create', 'ikuTargets', item.id, { item });
  await saveDB();
  broadcastChange('fra', 'update');
  res.json(item);
});

app.put('/api/iku-targets/:id', requireLogin, requireAdmin, validateBody(ikuTargetSchema), async (req, res) => {
  const idx = (DB.ikuTargets || []).findIndex(x => x.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Data tidak ditemukan' });
  const { ikuNomor, tahun, timId } = req.body;
  const dup = DB.ikuTargets.find(t => t.id !== req.params.id && t.ikuNomor === ikuNomor && Number(t.tahun) === Number(tahun) && t.timId === timId);
  if (dup) return res.status(400).json({ error: 'Target untuk No IKU, Tim, dan tahun ini sudah ada pada data lain.' });
  DB.ikuTargets[idx] = { ...DB.ikuTargets[idx], ...req.body, id: req.params.id };
  await auditLog(req.session.user, 'update', 'ikuTargets', req.params.id, { newValue: DB.ikuTargets[idx] });
  await saveDB();
  broadcastChange('fra', 'update');
  res.json(DB.ikuTargets[idx]);
});

app.delete('/api/iku-targets/:id', requireLogin, requireAdmin, async (req, res) => {
  const idx = (DB.ikuTargets || []).findIndex(x => x.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Data tidak ditemukan' });
  const deleted = DB.ikuTargets.splice(idx, 1)[0];
  await auditLog(req.session.user, 'delete', 'ikuTargets', req.params.id, { deleted });
  await saveDB();
  broadcastChange('fra', 'update');
  res.json({ deleted: true });
});

app.get('/api/users', requireLogin, requireAdmin, (req, res) => {
  res.json((DB.users || []).map(sanitizeUserForClient));
});

app.post('/api/users', requireLogin, requireAdmin, validateBody(userCreateSchema), async (req, res) => {
  const { username, email, password, role, teamId } = req.body;
  const emailLower = email.toLowerCase();
  const dup = (DB.users || []).find(u => u.username.toLowerCase() === username.toLowerCase() || u.email.toLowerCase() === emailLower);
  if (dup) return res.status(400).json({ error: 'Username atau email sudah digunakan' });
  const passwordHash = await bcrypt.hash(password, 10);
  const user = { id: uuidv4(), username, email: emailLower, passwordHash, role, teamId: role === 'admin' ? null : teamId };
  DB.users.push(user);
  await auditLog(req.session.user, 'create', 'users', user.id, { username, email: emailLower, role });
  await saveDB();
  res.json(sanitizeUserForClient(user));
});

app.put('/api/users/:id', requireLogin, requireAdmin, validateBody(userUpdateSchema), async (req, res) => {
  const idx = (DB.users || []).findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'User tidak ditemukan' });
  const { username, email, password, role, teamId } = req.body;
  const emailLower = email.toLowerCase();
  const dup = DB.users.find(u => u.id !== req.params.id && (u.username.toLowerCase() === username.toLowerCase() || u.email.toLowerCase() === emailLower));
  if (dup) return res.status(400).json({ error: 'Username atau email sudah digunakan oleh user lain' });
  const existing = DB.users[idx];
  const passwordHash = password ? await bcrypt.hash(password, 10) : existing.passwordHash;
  DB.users[idx] = { ...existing, username, email: emailLower, passwordHash, role, teamId: role === 'admin' ? null : teamId };
  await auditLog(req.session.user, 'update', 'users', req.params.id, { username, email: emailLower, role });
  await saveDB();
  res.json(sanitizeUserForClient(DB.users[idx]));
});

app.delete('/api/users/:id', requireLogin, requireAdmin, async (req, res) => {
  const idx = (DB.users || []).findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'User tidak ditemukan' });
  if (DB.users[idx].id === req.session.user.id) return res.status(400).json({ error: 'Tidak dapat menghapus akun sendiri' });
  const deleted = DB.users.splice(idx, 1)[0];
  await auditLog(req.session.user, 'delete', 'users', req.params.id, { username: deleted.username });
  await saveDB();
  res.json({ deleted: true });
});

app.get('/api/audit', requireLogin, requireAdmin, (req, res) => {
  res.json(DB.auditLogs || []);
});

// Mengosongkan seluruh riwayat aktivitas (audit log) — hanya admin.
// Tindakan ini sendiri TIDAK dicatat di log (karena log-nya baru saja dikosongkan),
// tapi tetap aman: fitur audit log tetap aktif mencatat aktivitas berikutnya seperti biasa.
app.delete('/api/audit', requireLogin, requireAdmin, async (req, res) => {
  const totalBefore = (DB.auditLogs || []).length;
  DB.auditLogs = [];
  await saveDB();
  res.json({ deleted: true, totalDeleted: totalBefore });
});

app.get('/api/reminders', requireLogin, (req, res) => {
  res.json(computeReminders(req.session.user));
});

const notificationSchema = Joi.object({
  targetUserId: Joi.string().required(),
  collection: Joi.string().valid('fra', 'kegiatan', 'tugas', 'iku').required(),
  itemId: Joi.string().required(),
  message: Joi.string().trim().min(3).max(500).required(),
});

// Daftar pengingat pribadi milik user yang sedang login saja (bukan seluruh tim).
app.get('/api/notifications', requireLogin, (req, res) => {
  const list = (DB.notifications || [])
    .filter((n) => n.targetUserId === req.session.user.id)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  res.json(list);
});

// Admin mengirim pengingat bertarget ke satu user tertentu (mis. pengunggah data FRA).
app.post('/api/notifications', requireLogin, requireAdmin, validateBody(notificationSchema), async (req, res) => {
  const { targetUserId, collection, itemId, message } = req.body;
  const targetUser = (DB.users || []).find((u) => u.id === targetUserId);
  if (!targetUser) return res.status(404).json({ error: 'User tujuan tidak ditemukan' });
  const notif = {
    id: uuidv4(),
    targetUserId,
    targetUsername: targetUser.username,
    fromUsername: req.session.user.username,
    collection,
    itemId,
    message,
    createdAt: new Date().toISOString(),
  };
  DB.notifications.push(notif);
  await auditLog(req.session.user, 'create', 'notifications', notif.id, { targetUsername: targetUser.username, message });
  await saveDB();
  broadcastChange('notifications', 'create');   // ⬅️ TAMBAHKAN BARIS INI
  res.json(notif);
});

// User menutup/menyelesaikan pengingatnya sendiri.
app.delete('/api/notifications/:id', requireLogin, async (req, res) => {
  const idx = (DB.notifications || []).findIndex((n) => n.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Notifikasi tidak ditemukan' });
  if (DB.notifications[idx].targetUserId !== req.session.user.id) {
    return res.status(403).json({ error: 'Bukan notifikasi Anda' });
  }
  DB.notifications.splice(idx, 1);
  await saveDB();
  broadcastChange('notifications', 'delete');   // ⬅️ TAMBAHKAN BARIS INI
  res.json({ deleted: true });
});

app.get('/api/export/json', requireLogin, (req, res) => {
  const user = req.session.user;
  const collection = req.query.collection;
  const view = req.query.view;
  const targetCollections = collection ? [collection] : ['teams', 'fra', 'kegiatan', 'tugas', 'iku'];
  const payload = {};
  targetCollections.forEach((name) => {
    if (!DB[name]) return;
    let items = filterItems(DB[name], user, name, req.query);
    if (name === 'fra') items = items.map(enrichFraWithLiveTarget);
    payload[name] = items;
  });
  const { filename } = resolveExportMeta(collection, view);
  res.setHeader('Content-Disposition', `attachment; filename=${filename}.json`);
  res.setHeader('Content-Type', 'application/json');
  res.send(JSON.stringify(payload, null, 2));
});

app.get('/api/export/excel', requireLogin, async (req, res) => {
  try {
    const user = req.session.user;
    const collection = req.query.collection;
    const view = req.query.view;
    const workbook = new ExcelJS.Workbook();

    if (collection) {
      const { columns, rows } = buildExportDataset(user, req.query);
      const sheetName = (view === 'pencapaian' ? 'TL TW Sebelumnya' : collection).slice(0, 31);
      const sheet = workbook.addWorksheet(sheetName);
      sheet.columns = columns.map((col) => ({ header: col.label, key: col.key, width: Math.max(12, Math.round(col.width / 5)) }));
      sheet.getRow(1).font = { bold: true };
      if (rows.length) {
        rows.forEach((row) => sheet.addRow(sanitizeExcelRow(row)));
      } else {
        sheet.addRow(columns.reduce((acc, c) => ((acc[c.key] = '-'), acc), {}));
      }
    } else {
      for (const name of ['teams', 'fra', 'kegiatan', 'tugas', 'iku']) {
        const { columns, rows } = buildExportDataset(user, { ...req.query, collection: name });
        const sheet = workbook.addWorksheet(name);
        sheet.columns = columns.map((col) => ({ header: col.label, key: col.key, width: Math.max(12, Math.round(col.width / 5)) }));
        sheet.getRow(1).font = { bold: true };
        rows.forEach((row) => sheet.addRow(sanitizeExcelRow(row)));
      }
      if (!workbook.worksheets.length) workbook.addWorksheet('Data').addRow(['Tidak ada data']);
    }

    const { filename } = resolveExportMeta(collection, view);
    res.setHeader('Content-Disposition', `attachment; filename=${filename}.xlsx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await workbook.xlsx.write(res);
  } catch (err) {
    console.error('Gagal membuat file Excel:', err.message);
    if (!res.headersSent) res.status(500).json({ error: 'Gagal membuat file Excel: ' + err.message });
    else res.destroy();
  }
});

app.get('/api/export/pdf', requireLogin, (req, res) => {
  try {
    const user = req.session.user;
    const collection = req.query.collection || 'kegiatan';
    const view = req.query.view;

    // pakai dataset builder yang SAMA dengan Excel — jangan bikin ulang di sini
    const { columns, rows } = buildExportDataset(user, { ...req.query, collection });
    const { filename, title } = resolveExportMeta(collection, view);

    // FRA & Kegiatan/TL TW Sebelumnya kolomnya banyak -> landscape biar muat
    const isWide = collection === 'fra' || collection === 'kegiatan';
    const doc = new PDFDocument({ size: 'A4', margin: 30, layout: isWide ? 'landscape' : 'portrait' });

    res.setHeader('Content-Disposition', `attachment; filename=${filename}.pdf`);
    res.setHeader('Content-Type', 'application/pdf');
    doc.pipe(res);

    drawPdfCover(doc, title);

    if (!rows.length) {
      doc.font('Helvetica').fontSize(11).fillColor('#5B6B78')
        .text('Belum ada data untuk ditampilkan pada laporan ini.', { align: 'center' });
    } else {
      drawPdfTable(doc, columns, rows);
    }
    doc.end();
  } catch (err) {
    console.error('Gagal membuat file PDF:', err.message);
    if (!res.headersSent) res.status(500).json({ error: 'Gagal membuat file PDF: ' + err.message });
    else res.destroy();
  }
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
    res.send('<h2>OneDrive berhasil terhubung.</h2><p>Anda bisa menutup tab ini dan kembali ke aplikasi SIMORA.</p>');
  } catch (err) {
    console.error('Gagal menghubungkan OneDrive:', err.message);
    res.status(500).send(`Gagal menghubungkan OneDrive: ${err.message}`);
  }
});

// Pemetaan modul -> nama folder OneDrive. Kalau kind tidak dikenali/tidak
// dikirim, jatuh ke folder default (ONEDRIVE_FOLDER di .env / 'SIMORA-Uploads').
const UPLOAD_FOLDER_BY_KIND = {
  fra: 'SIMORA-FRA',
  iku: 'SIMORA-IKU',
  tugas: 'SIMORA-Tugas',
  kegiatan: 'SIMORA-Kegiatan',
  pencapaian: 'SIMORA-TL-Sebelumnya',
};

const TRIWULAN_ROMAN_TO_ARABIC = { I: 1, II: 2, III: 3, IV: 4 };

function periodeToTriwulanNumber(periode) {
  if (!periode) return null;
  const match = /triwulan\s+(IV|III|II|I)\b/i.exec(periode);
  if (!match) return null;
  return TRIWULAN_ROMAN_TO_ARABIC[match[1].toUpperCase()] || null;
}

// Map format triwulan modul Tugas ("q1".."q4") ke angka biasa.
const TUGAS_QKEY_TO_ARABIC = { q1: 1, q2: 2, q3: 3, q4: 4 };

// Untuk FRA: kalau periode & ikuNomor valid, arahkan ke folder bertingkat
// "SIMORA-FRA/Triwulan {n}/IKU {n}".
// Untuk Tugas: kalau triwulan valid, arahkan ke "SIMORA-Tugas/Triwulan {n}".
// Modul lain tetap folder flat seperti biasa.
// Nama folder OneDrive tidak boleh mengandung karakter tertentu (" \ / : * ? < > |)
function sanitizeFolderSegment(name) {
  return String(name || '')
    .replace(/[":\\/*?<>|]/g, '')
    .trim()
    .slice(0, 100) || 'Lainnya';
}

function buildUploadFolder(kind, meta) {
  const baseFolder = UPLOAD_FOLDER_BY_KIND[kind];
  if (!baseFolder) return undefined; // kind tak dikenal -> onedrive.js pakai folder default

  if (kind === 'fra' || kind === 'kegiatan' || kind === 'pencapaian') {
    // SIMORA-{tahun}/{baseFolder}/Triwulan {n}/{Nama Tim}/IKU {kode}
    const tahunNum = parseInt(meta.tahun, 10);
    const yearPrefix = tahunNum ? `SIMORA-${tahunNum}/` : '';
    const triwulanNum = periodeToTriwulanNumber(meta.periode);
    const ikuKode = String(meta.ikuNomor || '').trim();
    const teamSegment = meta.timId ? sanitizeFolderSegment(teamNameServer(meta.timId)) : '';

    if (triwulanNum && teamSegment && ikuKode) {
      return `${yearPrefix}${baseFolder}/Triwulan ${triwulanNum}/${teamSegment}/IKU ${ikuKode}`;
    }
    if (triwulanNum && ikuKode) {
      // fallback kalau timId belum terkirim (data lama / kasus tak terduga)
      return `${yearPrefix}${baseFolder}/Triwulan ${triwulanNum}/IKU ${ikuKode}`;
    }
    return `${yearPrefix}${baseFolder}`;
  }

  if (kind === 'tugas') {
    const triwulanNum = TUGAS_QKEY_TO_ARABIC[meta.triwulan];
    if (triwulanNum) {
      return `${baseFolder}/Triwulan ${triwulanNum}`;
    }
  }
  return baseFolder;
}

app.post('/api/upload', requireLogin, (req, res) => {
  upload.single('file')(req, res, async (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'Tidak ada berkas yang dikirim' });
    try {
      const originalName = sanitizeText(req.file.originalname);
      const ext = path.extname(req.file.originalname).toLowerCase();
      const baseName = path.basename(originalName, ext)
      .replace(/[^a-zA-Z0-9-_ ]/g, '')
      .trim()
      .slice(0, 100);
      const shortId = uuidv4().split('-')[0];
      const storedName = `${baseName}-${shortId}${ext}`;
      const kind = sanitizeText(req.body.kind || '');
      const periode = sanitizeText(req.body.periode || '');
      const ikuNomor = req.body.ikuNomor;
      const triwulan = sanitizeText(req.body.triwulan || '');
      const tahun = sanitizeText(req.body.tahun || '');
      const timId = sanitizeText(req.body.timId || '');           // BARU
      const folder = buildUploadFolder(kind, { periode, ikuNomor, triwulan, tahun, timId }); // timId ditambahkan
      console.log(`[Upload] kind="${kind}" periode="${periode}" ikuNomor="${ikuNomor}" triwulan="${triwulan}" -> folder tujuan: ${folder || '(default) ' + (process.env.ONEDRIVE_FOLDER || 'SIMORA-Uploads')}`);
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

// Endpoint JSON: kembalikan info berkas (termasuk downloadUrl sementara) supaya
// frontend bisa menampilkannya dalam modal viewer di dalam web sendiri (tidak pindah situs, tidak download).
app.get('/api/files/:filename/info', requireLogin, async (req, res) => {
  try {
    const itemId = req.params.filename;
    const info = await onedrive.getDownloadUrl(itemId);
    if (!info) {
      pruneMissingFileReference(itemId); // hapus referensi karena sudah tidak ada di OneDrive
      return res.status(404).json({ error: 'Berkas tidak ditemukan (mungkin sudah dihapus di OneDrive)' });
    }
    res.json({ downloadUrl: info.downloadUrl, webUrl: info.webUrl, name: info.name, mimeType: info.mimeType });
  } catch (err) {
    console.error('Gagal mengambil info berkas dari OneDrive:', err.message);
    const msg = String(err.message || '').toLowerCase();
    if (msg.includes('404') || msg.includes('not found') || msg.includes('itemnotfound')) {
      pruneMissingFileReference(req.params.filename);
      return res.status(404).json({ error: 'Berkas tidak ditemukan (mungkin sudah dihapus di OneDrive)' });
    }
    res.status(502).json({ error: `Gagal mengambil info berkas dari OneDrive: ${err.message}` });
  }
});

// Fallback: buka langsung viewer OneDrive Online di tab baru (dipakai tombol "Buka di tab baru" pada modal).
app.get('/api/files/:filename', requireLogin, async (req, res) => {
  try {
    const itemId = req.params.filename;
    const info = await onedrive.getDownloadUrl(itemId);
    if (!info) {
      pruneMissingFileReference(itemId);
      return res.status(404).json({ error: 'Berkas tidak ditemukan (mungkin sudah dihapus di OneDrive)' });
    }
    if (info.webUrl) return res.redirect(info.webUrl);
    if (info.downloadUrl) return res.redirect(info.downloadUrl);
    return res.status(404).json({ error: 'Berkas tidak ditemukan' });
  } catch (err) {
    console.error('Gagal mengambil berkas dari OneDrive:', err.message);
    const msg = String(err.message || '').toLowerCase();
    if (msg.includes('404') || msg.includes('not found') || msg.includes('itemnotfound')) {
      pruneMissingFileReference(req.params.filename);
      return res.status(404).json({ error: 'Berkas tidak ditemukan (mungkin sudah dihapus di OneDrive)' });
    }
    res.status(502).json({ error: `Gagal mengambil berkas dari OneDrive: ${err.message}` });
  }
});

// Endpoint terpisah untuk benar-benar MENGUNDUH berkas (dipakai kalau nanti perlu tombol Download).
app.get('/api/files/:filename/download', requireLogin, async (req, res) => {
  try {
    const itemId = req.params.filename;
    const info = await onedrive.getDownloadUrl(itemId);
    if (!info || !info.downloadUrl) return res.status(404).json({ error: 'Berkas tidak ditemukan' });
    res.redirect(info.downloadUrl);
  } catch (err) {
    console.error('Gagal mengunduh berkas dari OneDrive:', err.message);
    res.status(502).json({ error: `Gagal mengunduh berkas dari OneDrive: ${err.message}` });
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

app.post('/api/tanya-ai', requireLogin, isJsonRequest, async (req, res) => {
  try {
    const pertanyaan = sanitizeText(req.body.pertanyaan);
    if (!pertanyaan) {
      return res.status(400).json({ error: 'Pertanyaan tidak boleh kosong' });
    }

    const user = req.session.user;
    const fraScope = filterItems(DB.fra, user, 'fra', {}).map(enrichFraWithLiveTarget);
    const kegiatanScope = filterItems(DB.kegiatan, user, 'kegiatan', {});
    const ikuScope = filterItems(DB.iku, user, 'iku', {});
    const tugasScope = filterItems(DB.tugas, user, 'tugas', {});

    const contextData = {
      teams: DB.teams.map((team) => ({ id: team.id, name: team.name })),
      fra: fraScope,
      kegiatan: kegiatanScope,
      iku: ikuScope,
      tugas: tugasScope,
      totalFra: fraScope.length,
      totalKegiatan: kegiatanScope.length,
      totalIku: ikuScope.length,
      totalTugas: tugasScope.length,
      rataRataCapaianFra: fraScope.length
        ? Math.round(fraScope.reduce((s, f) => s + (Number(f.persentase) || 0), 0) / fraScope.length)
        : 0,
      rataRataCapaianIku: ikuScope.length
        ? Math.round(ikuScope.reduce((s, i) => s + (Number(i.capaian) || 0), 0) / ikuScope.length)
        : 0,
      tugasTerbaru: tugasScope.slice(0, 5).map(t => ({
        nama: t.nama, target: t.target, realisasi: t.realisasi, status: t.status, tim: teamNameServer(t.timId),
      })),
      kegiatanTerbaru: kegiatanScope.slice(0, 5).map(k => ({
        nama: k.nama, kendala: k.kendala, solusi: k.solusi, rtl: k.rtl, status: k.status, tim: teamNameServer(k.timId),
      })),
      ikuTerbaru: ikuScope.slice(0, 5).map(i => ({
        kode: i.kode || '-', nama: i.nama || '-', target: Number(i.target) || 0, capaian: Number(i.capaian) || 0, satuan: i.satuan || '-', team: teamNameServer(i.timId),
      })),
    };

    const jawaban = await tanyaAI(pertanyaan, contextData);
    res.json({ jawaban });
  } catch (err) {
    console.error('Error /api/tanya-ai:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.get('*', (req, res) => {
  if (req.path.startsWith('/api/')) return res.status(404).json({ error: 'Endpoint tidak ditemukan' });
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`✓ SIMORA BPS berjalan di http://localhost:${PORT}`);
  console.log(`  Data JSON   : ${DB_FILE}`);
  console.log(`  Berkas bukti: OneDrive (${process.env.ONEDRIVE_USER || 'ONEDRIVE_USER belum diisi di .env'} / ${process.env.ONEDRIVE_FOLDER || 'SIMORA-Uploads'})`);
});