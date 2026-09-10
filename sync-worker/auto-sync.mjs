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
import { parseCredit, parseZone } from '../core.js';

const SUPABASE_URL = process.env.SUPABASE_URL || 'https://wstxgbzmsbosinmhhjbl.supabase.co';
const SUPABASE_TABLE = 'dashboard_snapshots';
const SNAPSHOT_KEY = 'credit-card-extra-amount';
const ZONE_SNAPSHOT_KEY = 'zone-distribution';
const FOLDER_ID = process.env.DRIVE_FOLDER_ID || '16HTr8nfPz4P2PMr4QB0bjgwiD110Qd-0';
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
  const res = await drive.files.list({
    q: `'${FOLDER_ID.replace(/'/g, "\\'")}' in parents and trashed = false`,
    fields: 'files(id,name,mimeType,modifiedTime,size)',
    orderBy: 'modifiedTime desc,name',
    pageSize: 1000,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  const isCandidate = (f) => !f.name.startsWith('~$') && (/\.(xlsx|xlsm|csv|tsv)$/i.test(f.name) || f.mimeType === 'application/vnd.google-apps.spreadsheet');
  return (res.data.files || []).filter(isCandidate);
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

async function readBothSources() {
  const drive = await driveClient();
  const files = await listCandidateFiles(drive);
  if (!files.length) throw new Error('No Excel or Google Sheets files were found in the configured source folder.');
  let credit = null, zone = null;
  const errors = [];
  for (const meta of files) {
    if (credit && zone) break;
    if (Number(meta.size) > 30 * 1024 * 1024) continue;
    log('Checking', meta.name);
    let sheets;
    try {
      const buffer = await downloadFile(drive, meta);
      sheets = readWorkbookBuffer(buffer, meta.name);
    } catch (e) {
      errors.push({ name: meta.name, error: e.message });
      continue;
    }
    const source = { fileName: meta.name, fileId: meta.id, modifiedTime: meta.modifiedTime, signature: remoteSignature(meta), folderId: FOLDER_ID };
    if (!credit) { try { credit = parseCredit(sheets, source); } catch (e) { if (/credit\s*card/i.test(meta.name)) errors.push({ name: meta.name, error: e.message, kind: 'credit' }); } }
    if (!zone) { try { zone = parseZone(sheets, source); } catch (e) { if (/zone.*distribut/i.test(meta.name)) errors.push({ name: meta.name, error: e.message, kind: 'zone' }); } }
  }
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
  return res.status === 204 ? null : res.json();
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

  const creditSignatureChanged = credit.meta.signature !== currentPayload?.credit?.meta?.signature;
  const zoneSignatureChanged = zone.meta.signature !== currentPayload?.zone?.meta?.signature;

  if (!creditSignatureChanged && !zoneSignatureChanged) {
    log('No change in either source file. Nothing to publish.');
    return;
  }

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
