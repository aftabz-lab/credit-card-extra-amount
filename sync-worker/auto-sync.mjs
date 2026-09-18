// Scheduled server-side sync for the Credit Card Extra Amount dashboard.
// Runs outside any browser (GitHub Actions), so it does not depend on a tab
// being open. It reuses the exact same parsing rules as the dashboard
// (../core.js, imported unmodified) so figures never drift between the
// manual "Read latest files" flow and this automatic one.
//
// Required environment variables (set as GitHub Actions secrets):
//   GOOGLE_SERVICE_ACCOUNT_JSON  - full JSON key of a service account that has
//                                  been shared as a Viewer on the source Drive folder
//   SUPABASE_SERVICE_ROLE_KEY    - Supabase service-role key (Project Settings -> API)
// Optional:
//   DRIVE_FOLDER_ID              - overrides the folder ID baked into sync.js
//   SUPABASE_URL                 - overrides the URL baked into supabase-sync.js

import { google } from 'googleapis';
import { readWorkbookBuffer } from './xlsx-reader.node.mjs';
import { parseCredit, parseZone, isCcolOutletsSource, isIgnoredRawSource } from '../core.js';

const SUPABASE_URL = process.env.SUPABASE_URL || 'https://wstxgbzmsbosinmhhjbl.supabase.co';
const SUPABASE_TABLE = 'dashboard_snapshots';
const SNAPSHOT_KEY = 'credit-card-extra-amount';
const ZONE_SNAPSHOT_KEY = 'zone-distribution';
const FOLDER_ID = process.env.DRIVE_FOLDER_ID || '16HTr8nfPz4P2PMr4QB0bjgwiD110Qd-0';
const DRIVE_FOLDER_MIME = 'application/vnd.google-apps.folder';
const SERVICE_ROLE_KEY = requireEnv('SUPABASE_SERVICE_ROLE_KEY');

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
}

function log(...args) { console.log(new Date().toISOString(), ...args); }

async function driveClient() {
  const keyJson = requireEnv('GOOGLE_SERVICE_ACCOUNT_JSON');
  const credentials = JSON.parse(keyJson);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/drive.readonly'],
  });
  return google.drive({ version: 'v3', auth });
}

async function listCandidateFiles(drive) {
  const isCandidate = (f) => !f.name.startsWith('~$') && (/\.(xlsx|xlsm|csv|tsv)$/i.test(f.name) || f.mimeType === 'application/vnd.google-apps.spreadsheet');
  const files = [];
  const queue = [{ id: FOLDER_ID, path: '' }];
  const seen = new Set();
  while (queue.length) {
    const folder = queue.shift();
    if (seen.has(folder.id)) continue;
    seen.add(folder.id);
    let pageToken;
    do {
      const res = await drive.files.list({
        q: `'${folder.id.replace(/'/g, "\\'")}' in parents and trashed = false`,
        fields: 'nextPageToken,files(id,name,mimeType,modifiedTime,size)',
        orderBy: 'modifiedTime desc,name',
        pageSize: 1000,
        pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      });
      for (const meta of res.data.files || []) {
        const drivePath = folder.path ? `${folder.path}/${meta.name}` : meta.name;
        if (meta.mimeType === DRIVE_FOLDER_MIME) queue.push({ id: meta.id, path: drivePath });
        else if (isCandidate(meta)) files.push({ ...meta, drivePath });
      }
      pageToken = res.data.nextPageToken || undefined;
    } while (pageToken);
  }
  return files.sort((a, b) => Date.parse(b.modifiedTime) - Date.parse(a.modifiedTime) || a.name.localeCompare(b.name));
}

async function downloadFile(drive, meta) {
  if (meta.mimeType === 'application/vnd.google-apps.spreadsheet') {
    const res = await drive.files.export({ fileId: meta.id, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }, { responseType: 'arraybuffer' });
    return Buffer.from(res.data);
  }
  const res = await drive.files.get({ fileId: meta.id, alt: 'media', supportsAllDrives: true }, { responseType: 'arraybuffer' });
  return Buffer.from(res.data);
}

function remoteSignature(meta) {
  return [meta.id || '', meta.name || '', meta.size || '', meta.modifiedTime || ''].join('|');
}

// The old reference workbook is retired as a source. It is never read again,
// even as a fallback, even if it is still sitting in the shared Drive folder
// or has a newer modified time than the current "Credit Card.xlsx" file.
const isRetiredSource = (meta) => /compiled[\s_-]*credit[\s_-]*card[\s_-]*extra[\s_-]*amount/i.test(meta.name);

async function readBothSources() {
  const drive = await driveClient();
  const listed = await listCandidateFiles(drive);
  const all = listed.filter((m) => !isRetiredSource(m) && !isIgnoredRawSource(m));
  if (!all.length) throw new Error('No Excel or Google Sheets files were found in the configured source folder.');
  const creditFiles = all.filter(isCcolOutletsSource);
  const zoneFiles = all.filter((m) => !isCcolOutletsSource(m));
  if (!creditFiles.length) throw new Error('No ccol_outlets_<start-date>_to_<end-date> source file was found in the configured Drive folder or its subfolders. The retired data file is intentionally ignored; the previous snapshot has been kept.');
  // The dated CCoL export is the only Credit Card source. Zone Distribution
  // remains structure-validated independently, including in subfolders.
  let credit = null, zone = null;
  const errors = [];
  const sheetsCache = new Map();
  async function tryFile(meta, { forCredit, forZone }) {
    if (Number(meta.size) > 30 * 1024 * 1024) return;
    const displayName = meta.drivePath || meta.name;
    log('Checking', displayName);
    const signature = remoteSignature(meta);
    let sheets = sheetsCache.get(signature);
    if (!sheets) {
      try {
        const buffer = await downloadFile(drive, meta);
        sheets = readWorkbookBuffer(buffer, meta.name);
        sheetsCache.set(signature, sheets);
      } catch (e) {
        errors.push({ name: displayName, error: e.message });
        return;
      }
    }
    const source = { fileName: meta.name, filePath: displayName, fileId: meta.id, modifiedTime: meta.modifiedTime, signature, folderId: FOLDER_ID };
    if (forCredit && !credit) { try { credit = parseCredit(sheets, source); } catch (e) { errors.push({ name: displayName, time: Date.parse(meta.modifiedTime), error: e.message, kind: 'credit' }); } }
    if (forZone && !zone) { try { zone = parseZone(sheets, source); } catch (e) { errors.push({ name: displayName, time: Date.parse(meta.modifiedTime), error: e.message, kind: 'zone' }); } }
  }
  for (const meta of creditFiles) { if (credit) break; await tryFile(meta, { forCredit: true, forZone: false }); }
  for (const meta of zoneFiles) { if (zone) break; await tryFile(meta, { forCredit: false, forZone: true }); }
  if (!credit || !zone) {
    throw new Error(`Could not validate ${!credit ? 'Credit Card' : ''}${!credit && !zone ? ' and ' : ''}${!zone ? 'Zone Distribution' : ''}. ${errors.slice(0, 2).map((e) => e.name + ': ' + e.error).join(' ')}`);
  }
  return { credit, zone };
}

async function supabaseRequest(path, { method = 'GET', headers = {}, body } = {}) {
  const res = await fetch(`${SUPABASE_URL}${path}`, {
    method,
    headers: { apikey: SERVICE_ROLE_KEY, Authorization: `Bearer ${SERVICE_ROLE_KEY}`, 'Content-Type': 'application/json', ...headers },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Supabase request failed (${res.status}): ${text.slice(0, 400)}`);
  }
  const text = await res.text();
  return text.trim() ? JSON.parse(text) : null;
}

async function readSnapshot(key) {
  const rows = await supabaseRequest(`/rest/v1/${SUPABASE_TABLE}?select=snapshot_key,payload,updated_at&snapshot_key=eq.${encodeURIComponent(key)}&limit=1`);
  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

async function upsertSnapshot(key, payload) {
  await supabaseRequest(`/rest/v1/${SUPABASE_TABLE}?on_conflict=snapshot_key`, {
    method: 'POST',
    headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
    body: { snapshot_key: key, payload, updated_at: new Date().toISOString() },
  });
}

async function main() {
  const current = await readSnapshot(SNAPSHOT_KEY);
  const currentZone = await readSnapshot(ZONE_SNAPSHOT_KEY);
  const currentPayload = current?.payload;

  const { credit, zone } = await readBothSources();

  const zoneSignatureChanged = zone.meta.signature !== currentPayload?.zone?.meta?.signature;

  if (credit.warnings?.some((w) => w.includes('reported Extra Amount different'))) {
    throw new Error('Resolve the Extra Amount / channel-sum mismatch (Column rules or source formulas) before this can auto-publish. No snapshot was written.');
  }

  // Preserve existing overrides/schemaOverrides made by reviewers in the dashboard.
  const saved = {
    format: 'credit-card-snapshot-v1',
    schemaVersion: 1,
    credit,
    zone,
    overrides: currentPayload?.overrides || {},
    schemaOverrides: currentPayload?.schemaOverrides || {},
    signature: await fingerprint({ credit, zone, overrides: currentPayload?.overrides || {}, schemaOverrides: currentPayload?.schemaOverrides || {} }),
    generatedAt: new Date().toISOString(),
  };

  if (current?.payload?.signature === saved.signature) {
    log('Computed signature unchanged after full parse. Nothing to publish.');
    return;
  }

  await upsertSnapshot(SNAPSHOT_KEY, saved);
  log(`Published new Credit Card snapshot. Credit source: ${credit.meta.fileName} (${credit.meta.modifiedTime}). Zone source: ${zone.meta.fileName} (${zone.meta.modifiedTime}).`);

  // Keep the shared Zone Distribution snapshot fresh too, the same way the
  // browser's poll() picks up a newer zone-distribution row, but only if this
  // run actually found a newer zone file than what's currently published.
  if (zoneSignatureChanged) {
    const zoneRows = zone.rows;
    const zonePayload = {
      snapshot: {
        rows: zoneRows,
        sheetName: zone.meta.sheetName,
        fileName: zone.meta.fileName,
        savedAt: new Date().toISOString(),
      },
    };
    const zoneChangedVsPublished = JSON.stringify(zonePayload.snapshot.rows) !== JSON.stringify(currentZone?.payload?.snapshot?.rows);
    if (zoneChangedVsPublished) {
      await upsertSnapshot(ZONE_SNAPSHOT_KEY, zonePayload);
      log('Published refreshed Zone Distribution snapshot.');
    }
  }
}

async function fingerprint(payload) {
  const bytes = new TextEncoder().encode(JSON.stringify(payload));
  const { createHash } = await import('node:crypto');
  return createHash('sha256').update(bytes).digest('hex');
}

main().catch((e) => {
  console.error('Auto-sync failed:', e.message);
  process.exitCode = 1;
});
