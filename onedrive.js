/**
 * Modul integrasi OneDrive (Microsoft Graph API) — versi DELEGATED.
 *
 * Dipakai untuk akun Microsoft PRIBADI (mis. akun Gmail yang didaftarkan
 * sebagai Microsoft Account / Microsoft 365 Family, seperti bpskka@gmail.com).
 * Akun pribadi tidak bisa dipakai dengan alur "app-only / client credentials"
 * (itu hanya untuk akun organisasi Azure AD). Jadi di sini dipakai alur
 * OAuth "Authorization Code" standar:
 *
 *   1. Admin login SEKALI lewat browser ke /api/onedrive/connect
 *      (redirect ke halaman login Microsoft, login sbg bpskka@gmail.com, izinkan akses)
 *   2. Microsoft redirect balik ke /api/onedrive/callback membawa "code"
 *   3. Server tukar code itu dengan access_token + refresh_token,
 *      lalu simpan refresh_token di data/onedrive-token.json
 *   4. Setiap request berikutnya (upload/lihat/hapus berkas), server pakai
 *      refresh_token itu untuk minta access_token baru secara otomatis —
 *      TIDAK perlu login ulang, kecuali refresh_token dicabut/kedaluwarsa.
 *
 * Variabel environment yang dibutuhkan (isi di file .env):
 *   AZURE_CLIENT_ID     - Application (client) ID dari App Registration
 *   AZURE_CLIENT_SECRET - Client secret dari App Registration
 *   AZURE_REDIRECT_URI  - Harus SAMA PERSIS dengan yang didaftarkan di Azure
 *                         Portal, mis. http://localhost:3000/api/onedrive/callback
 *
 * Di Azure Portal, App Registration harus:
 *   - Supported account types: "Personal Microsoft accounts only"
 *     (atau "...and personal Microsoft accounts" kalau mau dukung dua-duanya)
 *   - Authentication > Platform: Web, Redirect URI = AZURE_REDIRECT_URI
 *   - API permissions (Delegated, BUKAN Application): Files.ReadWrite, offline_access
 *     -> untuk akun pribadi TIDAK ADA tombol admin consent, cukup Add permissions saja.
 */

const fs = require('fs');
const path = require('path');

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const REDIRECT_URI = process.env.AZURE_REDIRECT_URI;
const ONEDRIVE_FOLDER = process.env.ONEDRIVE_FOLDER || 'SIMORA-Uploads';

// "consumers" = khusus akun Microsoft pribadi. Kalau app registration Anda
// didaftarkan sebagai multi-tenant (organisasi + pribadi), ganti ke "common".
const AUTH_TENANT_SEGMENT = process.env.AZURE_AUTH_TENANT || 'consumers';
const SCOPES = 'offline_access Files.ReadWrite';

const AUTHORIZE_URL = `https://login.microsoftonline.com/${AUTH_TENANT_SEGMENT}/oauth2/v2.0/authorize`;
const TOKEN_URL = `https://login.microsoftonline.com/${AUTH_TENANT_SEGMENT}/oauth2/v2.0/token`;
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

const DATA_DIR = path.join(__dirname, 'data');
const TOKEN_FILE = path.join(DATA_DIR, 'onedrive-token.json');


function assertConfigured() {
  const missing = [];
  if (!CLIENT_ID) missing.push('AZURE_CLIENT_ID');
  if (!CLIENT_SECRET) missing.push('AZURE_CLIENT_SECRET');
  if (!REDIRECT_URI) missing.push('AZURE_REDIRECT_URI');
  if (missing.length) {
    throw new Error(`Konfigurasi OneDrive belum lengkap. Isi variabel berikut di .env: ${missing.join(', ')}`);
  }
}

function loadTokens() {
  if (!fs.existsSync(TOKEN_FILE)) return null;
  try {
    return JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf-8'));
  } catch {
    return null;
  }
}

function saveTokens(tokens) {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));
}

function isConnected() {
  const tokens = loadTokens();
  return Boolean(tokens && tokens.refresh_token);
}

/** Alamat untuk mengarahkan admin login ke Microsoft (dipakai di /api/onedrive/connect). */
function getAuthUrl(state) {
  assertConfigured();
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    response_type: 'code',
    redirect_uri: REDIRECT_URI,
    response_mode: 'query',
    scope: SCOPES,
    state: state || '',
  });
  return `${AUTHORIZE_URL}?${params.toString()}`;
}

/** Tukar "code" dari callback Microsoft dengan access_token + refresh_token, lalu simpan. */
async function exchangeCodeForTokens(code) {
  assertConfigured();
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    grant_type: 'authorization_code',
    code,
    redirect_uri: REDIRECT_URI,
    scope: SCOPES,
  });
  const res = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(`Gagal menukar kode login OneDrive: ${data.error_description || data.error || res.statusText}`);
  }
  saveTokens({
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    expires_at: Date.now() + Number(data.expires_in || 3600) * 1000,
  });
  return true;
}

async function refreshAccessToken(currentTokens) {
  assertConfigured();
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    grant_type: 'refresh_token',
    refresh_token: currentTokens.refresh_token,
    scope: SCOPES,
  });
  const res = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(
      `Gagal memperbarui akses OneDrive (mungkin perlu login ulang lewat /api/onedrive/connect): ${data.error_description || data.error || res.statusText}`
    );
  }
  const updated = {
    access_token: data.access_token,
    // Microsoft kadang mengganti refresh_token, kadang tidak — pakai yang baru kalau ada.
    refresh_token: data.refresh_token || currentTokens.refresh_token,
    expires_at: Date.now() + Number(data.expires_in || 3600) * 1000,
  };
  saveTokens(updated);
  return updated;
}

async function getAccessToken() {
  const tokens = loadTokens();
  if (!tokens || !tokens.refresh_token) {
    throw new Error(
      'OneDrive belum terhubung. Minta admin login lewat /api/onedrive/connect terlebih dahulu.'
    );
  }
  if (tokens.access_token && Date.now() < tokens.expires_at - 60000) {
    return tokens.access_token;
  }
  const refreshed = await refreshAccessToken(tokens);
  return refreshed.access_token;
}

async function graphFetch(url, options = {}) {
  const token = await getAccessToken();
  return fetch(url, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
  });
}

async function readGraphError(res) {
  try {
    const data = await res.json();
    return data.error?.message || res.statusText;
  } catch {
    return res.statusText;
  }
}

// Untuk delegated flow, "/me/drive" merujuk ke OneDrive milik akun yang login
// (bpskka@gmail.com), bukan /users/{email}/drive seperti pada alur app-only.
function driveRootUrl() {
  return `${GRAPH_BASE}/me/drive`;
}

const ensuredFolders = new Set();

/** Encode tiap segmen path secara terpisah — slash pemisah folder TIDAK ikut di-encode. */
function encodeGraphPath(p) {
  return p.split('/').filter(Boolean).map(encodeURIComponent).join('/');
}

/**
 * Pastikan folder ada, termasuk folder bertingkat (mis. "SIMORA-FRA/Triwulan 1/IKU 3").
 * Setiap level dicek & dibuat satu per satu kalau belum ada.
 */
async function ensureFolder(folderPath) {
  if (ensuredFolders.has(folderPath)) return;
  const segments = folderPath.split('/').filter(Boolean);
  let parentPath = '';
  let builtPath = '';

  for (const segment of segments) {
    parentPath = builtPath;
    builtPath = builtPath ? `${builtPath}/${segment}` : segment;
    if (ensuredFolders.has(builtPath)) continue;

    const checkUrl = `${driveRootUrl()}/root:/${encodeGraphPath(builtPath)}`;
    const res = await graphFetch(checkUrl);
    if (res.status === 200) {
      ensuredFolders.add(builtPath);
      continue;
    }
    if (res.status !== 404) {
      throw new Error(`Gagal memeriksa folder OneDrive "${builtPath}": ${await readGraphError(res)}`);
    }

    const createUrl = parentPath
      ? `${driveRootUrl()}/root:/${encodeGraphPath(parentPath)}:/children`
      : `${driveRootUrl()}/root/children`;
    const createRes = await graphFetch(createUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: segment,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'replace',
      }),
    });
    if (!createRes.ok) {
      throw new Error(`Gagal membuat folder OneDrive "${builtPath}": ${await readGraphError(createRes)}`);
    }
    ensuredFolders.add(builtPath);
  }
}

/**
 * Unggah buffer berkas ke OneDrive lewat upload session (aman untuk berbagai
 * ukuran berkas, tidak dibatasi ~4MB seperti PUT langsung).
 * `folderName` opsional — boleh berupa path bertingkat, mis. "SIMORA-FRA/Triwulan 1/IKU 3".
 * Kalau tidak diisi, pakai folder default (ONEDRIVE_FOLDER).
 * Mengembalikan metadata item OneDrive (termasuk `id`, dipakai sebagai
 * pengenal berkas di sisi aplikasi, menggantikan nama file lokal).
 */
async function uploadFile(buffer, filename, folderName) {
  const folder = folderName || ONEDRIVE_FOLDER;
  await ensureFolder(folder);
  const itemPath = `${folder}/${filename}`;

  const sessionRes = await graphFetch(
    `${driveRootUrl()}/root:/${encodeGraphPath(itemPath)}:/createUploadSession`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
    }
  );
  if (!sessionRes.ok) {
    throw new Error(`Gagal membuat sesi unggah OneDrive: ${await readGraphError(sessionRes)}`);
  }
  const session = await sessionRes.json();

  const uploadRes = await fetch(session.uploadUrl, {
    method: 'PUT',
    headers: {
      'Content-Length': String(buffer.length),
      'Content-Range': `bytes 0-${buffer.length - 1}/${buffer.length}`,
    },
    body: buffer,
  });
  if (!uploadRes.ok) {
    throw new Error(`Gagal mengunggah berkas ke OneDrive: ${await readGraphError(uploadRes)}`);
  }
  return uploadRes.json(); // { id, name, size, ... }
}

/** Ambil link unduh sementara (pre-authenticated) & link viewer OneDrive untuk sebuah item berdasarkan id-nya. */
async function getDownloadUrl(itemId) {
  const res = await graphFetch(`${driveRootUrl()}/items/${encodeURIComponent(itemId)}`);
  if (res.status === 404) return null;
  if (!res.ok) {
    throw new Error(`Gagal mengambil berkas dari OneDrive: ${await readGraphError(res)}`);
  }
  const data = await res.json();
  return {
    downloadUrl: data['@microsoft.graph.downloadUrl'], // link raw, otomatis download kalau dibuka
    webUrl: data.webUrl, // link viewer OneDrive Online, dibuka di tab tanpa mengunduh
    name: data.name,
    mimeType: data.file ? data.file.mimeType : undefined,
  };
}

async function deleteFile(itemId) {
  const res = await graphFetch(`${driveRootUrl()}/items/${encodeURIComponent(itemId)}`, {
    method: 'DELETE',
  });
  console.log(`[OneDrive] Graph API DELETE /items/${itemId} -> status ${res.status}`);
  if (res.status === 404) {
    console.log(`[OneDrive] Item ${itemId} tidak ditemukan di OneDrive (mungkin sudah terhapus sebelumnya, atau id salah).`);
  }
  if (!res.ok && res.status !== 404) {
    throw new Error(`Gagal menghapus berkas di OneDrive: ${await readGraphError(res)}`);
  }
  return true;
}

module.exports = {
  isConnected,
  getAuthUrl,
  exchangeCodeForTokens,
  uploadFile,
  getDownloadUrl,
  deleteFile,
};