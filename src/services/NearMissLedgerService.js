/**
 * ヒヤリハット周知・確認ログ（スプレッドシート＋localStorage）
 * 追記は Google Apps Script Web アプリ経由（APIキー単体では書込不可のため）
 */

import { CARELINK_FACILITIES, compactFacilityToken } from '../config/carelinkFacilities.js';
import {
  NEAR_MISS_ACK_SHEET_NAME,
  NEAR_MISS_REPORT_SHEET_NAME,
  REPORT_SHEET_HEADERS,
  REPORT_FIELD_KEYS,
  reportSheetColumnIndex,
  ACK_FIELD_KEYS,
  ackSheetColumnIndex,
} from '../config/nearMissLedgerConstants.js';
import { syncStaffRosterFromHrSheet } from './StaffRosterSheetService.js';
import { buildNearMissRosterPayloadFromShiftPreferences } from './shiftStaffRosterForNearMiss.js';
import {
  AWARENESS_LOG_SHEET_NAME,
  DEFAULT_AWARENESS_SPREADSHEET_ID,
} from '../config/hrSpreadsheetConstants.js';
import { NEAR_MISS_CATEGORY_LABELS } from './nearMissReportHtml.js';

const MONTH_CATEGORY_ORDER = [...NEAR_MISS_CATEGORY_LABELS, 'その他', '分類なし'];

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

export const NEAR_MISS_APPS_SCRIPT_URL = import.meta.env.VITE_NEAR_MISS_APPS_SCRIPT_URL ?? '';
export const NEAR_MISS_APP_SECRET = import.meta.env.VITE_NEAR_MISS_APP_SECRET ?? '';

/** ヒヤリのスプレッドシート追記（GAS）が使えるか（GAS Web アプリ URL が .env にあるか。追記は /api/near-miss-gas 経由でシークレットはサーバー側） */
export function isNearMissGasWriteConfigured() {
  return Boolean(String(NEAR_MISS_APPS_SCRIPT_URL ?? '').trim());
}

/** ヒヤリ台帳・周知ログの保存先 */
export function getLedgerSpreadsheetId() {
  const awareness = import.meta.env.VITE_AWARENESS_SPREADSHEET_ID?.trim();
  const legacy = import.meta.env.VITE_NEAR_MISS_LEDGER_SPREADSHEET_ID?.trim();
  return awareness || legacy || DEFAULT_AWARENESS_SPREADSHEET_ID;
}

const LS = {
  notices: 'carelink_os_near_miss_ledger_notices_v1',
  acks: 'carelink_os_near_miss_ledger_acks_v1',
  roster: 'carelink_os_near_miss_roster_v1',
  rosterSynced: 'carelink_os_hr_staff_roster_sync_v1',
  staffProfile: 'carelink_os_staff_profile_v1',
  ledgerSheetWriteLast: 'carelink_ledger_sheet_write_last_v1',
};

const MODEL = 'gemini-1.5-flash';

function readJson(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

function writeJson(key, val) {
  localStorage.setItem(key, JSON.stringify(val));
}

function notifyStaffProfileChanged() {
  if (typeof window === 'undefined') return;
  try {
    window.dispatchEvent(new CustomEvent('carelink-staff-profile'));
  } catch {
    // noop
  }
}

function newId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 12)}`;
}

function sheetRangeA1(sheetTitle, range) {
  const safe = `'${String(sheetTitle).replace(/'/g, "''")}'`;
  return `${safe}!${range}`;
}

/**
 * @param {string} label
 */
export function resolveLinkKeyFromFacilityLabel(label) {
  const s = String(label ?? '').trim();
  if (!s) return '';
  const hit = CARELINK_FACILITIES.find(
    (f) =>
      f.tabLabel === s ||
      f.linkKey === s ||
      f.sheetTitle === s ||
      compactFacilityToken(f.tabLabel) === compactFacilityToken(s) ||
      compactFacilityToken(f.sheetTitle) === compactFacilityToken(s)
  );
  return hit?.linkKey ?? s;
}

/**
 * @param {Record<string, unknown>} draft
 */
function inferImportance(draft) {
  const cats = Array.isArray(draft?.categories) ? draft.categories.map((x) => String(x)) : [];
  const high = new Set(['転倒', '転落', '窒息・誤嚥', 'やけど・火傷']);
  if (cats.some((c) => high.has(c))) return 'high';
  const text = `${draft?.situationContent ?? ''}${draft?.causeAndMeasures ?? ''}`;
  if (/転倒|転落|窒息|誤嚥|火傷|緊急|心肺|意識消失/u.test(String(text))) return 'high';
  return 'normal';
}

/**
 * @param {Record<string, unknown>} draft
 */
function makeTitle(draft) {
  const place = String(draft?.occurPlace ?? '').trim();
  const r = String(draft?.residentName ?? '').trim();
  const cat = (Array.isArray(draft?.categories) && draft.categories[0]) || 'ヒヤリハット';
  const base = `${cat}${place ? `（${place}）` : ''}${r ? ` ${r}様` : ''}`.trim();
  return base.slice(0, 200) || 'ヒヤリハット（気づき）';
}

/**
 * @param {Record<string, unknown>} draft
 */
function summarizeDraft(draft) {
  const t = String(draft?.situationContent ?? '').trim().replace(/\s+/g, ' ');
  return t.slice(0, 500);
}

/**
 * 事故報告を周知文へ短く整形
 * @param {Record<string, unknown>} draft
 */
function summarizeAccidentDraft(draft) {
  const situation = String(draft?.situation ?? '').trim().replace(/\s+/g, ' ');
  const causes = String(draft?.causes ?? '').trim().replace(/\s+/g, ' ');
  const improvements = String(draft?.improvements ?? '').trim().replace(/\s+/g, ' ');
  const body = [situation, causes, improvements].filter(Boolean).join(' / ');
  return body.slice(0, 500);
}

/** @returns {Record<string, unknown>[]} */
export function getLedgerNotices() {
  const raw = readJson(LS.notices, []);
  return Array.isArray(raw) ? raw : [];
}

/** @returns {Record<string, unknown>[]} */
export function getLedgerAcks() {
  const raw = readJson(LS.acks, []);
  return Array.isArray(raw) ? raw : [];
}

/**
 * @returns {Record<string, { id: string; name: string }[]>}
 */
export function getStaffRosterByFacility() {
  const raw = readJson(LS.roster, {});
  return raw && typeof raw === 'object' ? raw : {};
}

/**
 * @param {string} linkKey
 * @param {{ id: string; name: string }[]} rows
 */
export function setStaffRosterForFacility(linkKey, rows) {
  const k = String(linkKey ?? '').trim();
  if (!k) return;
  const cur = getStaffRosterByFacility();
  const cleaned = (Array.isArray(rows) ? rows : [])
    .map((r) => ({
      id: String(r?.id ?? '').trim() || newId(),
      name: String(r?.name ?? '').trim(),
    }))
    .filter((r) => r.name);
  cur[k] = cleaned;
  writeJson(LS.roster, cur);
}

/** @returns {Record<string, unknown> | null} */
export function getSyncedRosterPayload() {
  const raw = readJson(LS.rosterSynced, null);
  return raw && typeof raw === 'object' ? raw : null;
}

/**
 * 求人シート同期があれば優先。無ければ手入力名簿
 * @param {string} linkKey
 */
export function getEffectiveStaffRosterForFacility(linkKey) {
  const lk = String(linkKey ?? '').trim();
  const manual = getStaffRosterByFacility()[lk] ?? [];
  const sync = getSyncedRosterPayload();
  if (!sync?.syncedAt) return manual;

  if (Array.isArray(sync.global) && sync.global.length) {
    return /** @type {{ id: string; name: string }[]} */ (sync.global);
  }
  const per = sync.byFacility?.[lk];
  if (Array.isArray(per) && per.length) {
    return /** @type {{ id: string; name: string }[]} */ (per);
  }
  return manual;
}

/**
 * 求人・入退社シートからスタッフ名簿を読み込み、周知対象者を端末に保存
 * @param {string} sheetsApiKey
 * @param {{ preferredSheetTitle?: string }} [opts]
 */
export async function syncStaffRosterFromHrSheetAndStore(sheetsApiKey, opts = {}) {
  const data = await syncStaffRosterFromHrSheet(sheetsApiKey, opts);
  const prevMeta = data.meta && typeof data.meta === 'object' ? data.meta : {};
  writeJson(LS.rosterSynced, {
    sheetTitle: data.sheetTitle,
    spreadsheetId: data.spreadsheetId,
    byFacility: data.byFacility,
    global: data.global,
    syncedAt: data.syncedAt,
    meta: { ...prevMeta, source: 'hr_sheet' },
  });
  return data;
}

/**
 * 勤務表（記録アプリの「勤務希望・勤務表」画面で保存したスタッフ）から周知名簿を上書きする。
 * Google Sheets API は不要。
 */
export function syncStaffRosterFromShiftScheduleAndStore() {
  const data = buildNearMissRosterPayloadFromShiftPreferences();
  writeJson(LS.rosterSynced, {
    sheetTitle: data.sheetTitle,
    spreadsheetId: data.spreadsheetId,
    byFacility: data.byFacility,
    global: data.global,
    syncedAt: data.syncedAt,
    meta: data.meta,
  });
  return data;
}

/** @returns {{ staffId: string; displayName: string; lastFacilityLinkKey: string; nursingOfficeMode: boolean } | null} */
export function getStaffProfile() {
  const raw = readJson(LS.staffProfile, null);
  if (!raw || typeof raw !== 'object') return null;
  let staffId = String(raw.staffId ?? '').trim();
  // 未保存のときは毎回 newId() にならないよう、初回だけ生成して localStorage に固定する（氏名変更と独立）
  if (!staffId) {
    staffId = newId();
    writeJson(LS.staffProfile, {
      staffId,
      displayName: String(raw.displayName ?? '').trim(),
      lastFacilityLinkKey: String(raw.lastFacilityLinkKey ?? '').trim(),
      nursingOfficeMode: false,
    });
  }
  return {
    staffId,
    displayName: String(raw.displayName ?? '').trim(),
    lastFacilityLinkKey: String(raw.lastFacilityLinkKey ?? '').trim(),
    nursingOfficeMode: Boolean(raw.nursingOfficeMode),
  };
}

/**
 * @param {{ displayName: string; lastFacilityLinkKey?: string; nursingOfficeMode?: boolean }} p
 */
export function saveStaffProfile(p) {
  const prev = getStaffProfile();
  const displayName = String(p?.displayName ?? '').trim();
  // 氏名だけ変えても同じ職員として扱う（確認記録は staffId で紐づく）
  const staffId = String(prev?.staffId ?? '').trim() || newId();
  const nursingOfficeMode =
    typeof p?.nursingOfficeMode === 'boolean' ? p.nursingOfficeMode : Boolean(prev?.nursingOfficeMode);
  const o = {
    staffId,
    displayName,
    lastFacilityLinkKey: String(p?.lastFacilityLinkKey ?? prev?.lastFacilityLinkKey ?? '').trim(),
    nursingOfficeMode,
  };
  writeJson(LS.staffProfile, o);
  notifyStaffProfileChanged();
  return o;
}

/** 看護事務向け UI（訪問看護・特別指示の手動登録など）を表示するか */
export function isNursingOfficeUiEnabled() {
  return Boolean(getStaffProfile()?.nursingOfficeMode);
}

function upsertNotice(notice) {
  const list = getLedgerNotices().filter((n) => String(n?.id) !== String(notice.id));
  list.push(notice);
  writeJson(LS.notices, list);
}

function pushAck(ack) {
  const list = getLedgerAcks();
  list.push(ack);
  writeJson(LS.acks, list);
}

/**
 * @param {string} noticeId
 * @param {string} staffId
 */
function hasAck(noticeId, staffId) {
  return getLedgerAcks().some(
    (a) => String(a.noticeId) === String(noticeId) && String(a.staffId) === String(staffId)
  );
}

/**
 * 直近の台帳GAS追記結果（デバッグ・UI表示用）
 * @returns {Record<string, unknown> | null}
 */
export function getLastLedgerSheetWriteResult() {
  const raw = readJson(LS.ledgerSheetWriteLast, null);
  return raw && typeof raw === 'object' ? raw : null;
}

/**
 * ヒヤリ保存直後に呼ぶ。周知キューへ載せ、スプレッドシートへ追記を試みる。
 * @param {{ id: string; savedAt: string; facilityLabel: string; draft: Record<string, unknown> }} p
 * @returns {Promise<{ notice: Record<string, unknown>; sheetResult: Record<string, unknown> }>}
 */
export async function appendNoticeFromSavedReport(p) {
  const facilityLinkKey = resolveLinkKeyFromFacilityLabel(p.facilityLabel);
  const notice = {
    id: String(p.id),
    createdAt: String(p.savedAt ?? new Date().toISOString()),
    facilityLinkKey,
    importance: inferImportance(p.draft),
    categories: Array.isArray(p.draft?.categories) ? p.draft.categories.map((x) => String(x)) : [],
    title: makeTitle(p.draft),
    summary: summarizeDraft(p.draft),
    draftSnapshot: p.draft,
    archived: false,
    archivedAt: '',
    source: 'app',
    noticeType: 'near_miss',
  };
  upsertNotice(notice);
  const sheetResult = await postAppsScript('appendReport', {
    row: noticeToReportRow(notice),
  });
  writeJson(LS.ledgerSheetWriteLast, { ...sheetResult, action: 'appendReport', at: new Date().toISOString() });
  return { notice, sheetResult };
}

/**
 * 事故報告保存直後に呼ぶ。重要周知を自動作成して掲示する
 * @param {{ id: string; savedAt: string; facilityLabel: string; draft: Record<string, unknown> }} p
 * @returns {Promise<{ notice: Record<string, unknown>; sheetResult: Record<string, unknown> }>}
 */
export async function appendNoticeFromSavedAccidentReport(p) {
  const facilityLinkKey = resolveLinkKeyFromFacilityLabel(p.facilityLabel);
  const resident = String(p.draft?.residentName ?? '').trim();
  const place = String(p.draft?.occurPlace ?? '').trim();
  const title = `【重要】新しい事故報告と対策があります${resident ? `（${resident}様）` : ''}`;
  const notice = {
    id: `accident-${String(p.id)}`,
    createdAt: String(p.savedAt ?? new Date().toISOString()),
    facilityLinkKey,
    importance: 'high',
    categories: ['事故報告'],
    title,
    summary: `${place ? `発生場所: ${place} / ` : ''}${summarizeAccidentDraft(p.draft) || '詳細は事故報告書を確認してください。'}`,
    draftSnapshot: p.draft,
    archived: false,
    archivedAt: '',
    source: 'app',
    noticeType: 'accident',
  };
  upsertNotice(notice);
  const sheetResult = await postAppsScript('appendReport', { row: noticeToReportRow(notice) });
  writeJson(LS.ledgerSheetWriteLast, { ...sheetResult, action: 'appendReport', at: new Date().toISOString() });
  return { notice, sheetResult };
}

/**
 * @param {Record<string, unknown>} n
 * @returns {string[]}
 */
function noticeToReportRow(n) {
  const draftJson = JSON.stringify(n.draftSnapshot ?? {}).replace(/\r?\n/g, ' ');
  return [
    String(n.id),
    String(n.createdAt),
    String(n.facilityLinkKey),
    String(n.importance),
    (Array.isArray(n.categories) ? n.categories : []).join('、'),
    String(n.title),
    String(n.summary),
    draftJson,
    n.archived ? '1' : '',
  ];
}

/**
 * @param {Record<string, unknown>} ack
 * @param {string} noticeTitle
 */
function awarenessLogRow(ack, noticeTitle) {
  const t = String(ack.noticeType ?? '').trim();
  const typeLabel = t === 'accident' ? '事故' : 'ヒヤリ';
  const facility = CARELINK_FACILITIES.find((f) => f.linkKey === String(ack.facilityLinkKey ?? '').trim());
  const facilityLabel = facility?.tabLabel ?? String(ack.facilityLinkKey ?? '');
  const dept = String(ack.noticeDepartment ?? '').trim();
  return [
    String(ack.confirmedAt),
    typeLabel,
    String(noticeTitle ?? ''),
    String(ack.staffName),
    dept,
    facilityLabel,
  ];
}

/**
 * GAS が返す英語メッセージを、本番での対処が分かる日本語に寄せる（アラート表示用にも export）
 * @param {unknown} raw
 */
export function formatNearMissGasWriteError(raw) {
  if (raw != null && typeof raw === 'object') {
    try {
      return formatNearMissGasWriteError(
        /** @type {{ message?: unknown; error?: unknown }} */ (raw).message ??
          /** @type {{ message?: unknown; error?: unknown }} */ (raw).error
      );
    } catch {
      return '不明';
    }
  }
  const t = String(raw ?? '').trim();
  if (!t) return '不明';
  const lower = t.toLowerCase();
  if (lower === 'unauthorized' || /\bunauthorized\b/i.test(t)) {
    return [
      '認証エラー（unauthorized）: Google Apps Script 内の APP_SECRET と、',
      'Vercel の NEAR_MISS_APP_SECRET（または VITE_NEAR_MISS_APP_SECRET）が',
      '完全一致していません。スペースの混入・別の値を貼り付け・環境変数を保存したあと未デプロイ、などを確認してください。',
    ].join('');
  }
  return t;
}

/**
 * ヒヤリ保存が GAS まで成功したあと、「用紙」シートへミラーしなかった場合に案内を出す。
 * @param {{ ok?: boolean; skipped?: boolean; data?: Record<string, unknown> }} sheetResult
 */
export function alertIfNearMissYoushiMirrorSkipped(sheetResult) {
  if (!sheetResult || sheetResult.skipped || !sheetResult.ok) return;
  const data = sheetResult.data;
  if (!data || typeof data !== 'object') return;
  const mirror = data.mirror;
  if (!mirror || typeof mirror !== 'object') return;
  if (mirror.mirrored) {
    console.info('[CareLink] ヒヤリ: 「報告」と「用紙」の両方へ追記しました。');
    return;
  }
  if (mirror.skipped && mirror.reason) {
    alert(
      `「報告」シートへの保存は完了しました。\n\n【用紙シート】${String(mirror.reason)}\n\n` +
        '用紙にも同じ1行を付けたい場合: 同一スプレッドシートにタブ名「用紙」を用意し、1行目のA1を「レコードID」にしてください（「報告」と同じ見出し）。\n' +
        'マージセル入りの様式シートには自動追記しません。別シートで FILTER 等により「報告」を参照する運用もできます。'
    );
  }
}

/** 同一オリジン上の GAS 中継 API（Vite ミドルウェア / Vercel api） */
function getNearMissGasProxyPath() {
  const b = String(import.meta.env.BASE_URL || '/').replace(/\/?$/, '/');
  return `${b}api/near-miss-gas`;
}

/**
 * ブラウザから GAS へ直接 POST（CORS が通る環境・フォールバック用）
 * @param {string} action
 * @param {Record<string, unknown>} body
 */
async function postAppsScriptDirect(action, body) {
  const url = NEAR_MISS_APPS_SCRIPT_URL?.trim();
  const secret = NEAR_MISS_APP_SECRET?.trim();
  if (!url || !secret) {
    return { ok: false, error: 'GAS URL または VITE_NEAR_MISS_APP_SECRET が未設定です' };
  }
  try {
    const res = await fetch(url, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        secret,
        action,
        ...body,
      }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) return { ok: false, error: formatNearMissGasWriteError(data?.message || data?.error || res.statusText) };
    if (data && data.ok === false) {
      return { ok: false, error: formatNearMissGasWriteError(data?.message || data?.error || 'GAS error') };
    }
    return { ok: true, data };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : 'network' };
  }
}

/**
 * @param {string} action
 * @param {Record<string, unknown>} body
 */
async function postAppsScript(action, body) {
  const urlConfigured = Boolean(NEAR_MISS_APPS_SCRIPT_URL?.trim());
  if (!urlConfigured) {
    return {
      ok: false,
      skipped: true,
      reason:
        'VITE_NEAR_MISS_APPS_SCRIPT_URL が未設定（.env に GAS Web アプリの URL を追加してください）',
    };
  }

  const jsonBody = JSON.stringify({ action, ...body });
  const proxyPath = getNearMissGasProxyPath();

  let proxyRes;
  let rawText = '';
  try {
    proxyRes = await fetch(proxyPath, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: jsonBody,
    });
    rawText = await proxyRes.text();
  } catch (e) {
    const fb = await postAppsScriptDirect(action, body);
    if (fb.ok) return fb;
    return { ok: false, error: e instanceof Error ? e.message : 'network' };
  }

  let data = {};
  try {
    data = rawText ? JSON.parse(rawText) : {};
  } catch {
    const fb = await postAppsScriptDirect(action, body);
    if (fb.ok) return fb;
    return {
      ok: false,
      error:
        proxyRes.status === 404
          ? 'GAS 中継 API が見つかりません（facility-portal をルートに再デプロイしてください）。'
          : '中継 API の応答が JSON ではありません',
    };
  }

  if (proxyRes.ok) {
    if (data && data.ok === false) {
      return { ok: false, error: formatNearMissGasWriteError(data?.message || data?.error || 'GAS error') };
    }
    return { ok: true, data };
  }

  const fb = await postAppsScriptDirect(action, body);
  if (fb.ok) return fb;
  return { ok: false, error: formatNearMissGasWriteError(data?.error || data?.message || proxyRes.statusText) };
}

/**
 * @param {string} apiKey
 * @param {string} spreadsheetId
 * @param {string} sheetTitle
 * @param {string} rangeA1
 */
async function fetchSheetValues(apiKey, spreadsheetId, sheetTitle, rangeA1 = 'A:Z') {
  const sid = encodeURIComponent(spreadsheetId);
  const a1 = encodeURIComponent(sheetRangeA1(sheetTitle, rangeA1));
  const url = `${SHEETS_API}/${sid}/values/${a1}?key=${encodeURIComponent(apiKey)}`;
  const res = await fetch(url);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error?.message ?? `シート「${sheetTitle}」の取得に失敗しました`);
  return data.values ?? [];
}

/**
 * @param {string[][]} rows
 */
function parseReportSheetRows(rows) {
  if (!rows?.length) return [];
  const head = rows[0].map((c) => String(c ?? '').trim());
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every((c) => !String(c ?? '').trim())) continue;
    const o = {};
    for (let k = 0; k < REPORT_FIELD_KEYS.length; k++) {
      const key = REPORT_FIELD_KEYS[k];
      const j = reportSheetColumnIndex(head, k);
      o[key] = j >= 0 && j < row.length ? String(row[j] ?? '').trim() : '';
    }
    if (!o.recordId) continue;
    out.push(reportRowToNotice(o));
  }
  return out;
}

/**
 * @param {string[][]} rows
 */
function parseAckSheetRows(rows) {
  if (!rows?.length) return [];
  const head = rows[0].map((c) => String(c ?? '').trim());
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every((c) => !String(c ?? '').trim())) continue;
    const o = {};
    for (let k = 0; k < ACK_FIELD_KEYS.length; k++) {
      const key = ACK_FIELD_KEYS[k];
      const j = ackSheetColumnIndex(head, k);
      o[key] = j >= 0 && j < row.length ? String(row[j] ?? '').trim() : '';
    }
    if (!o.noticeId) continue;
    if (!o.logId) o.logId = `${o.noticeId}::${o.staffId || 'row'}`;
    out.push(ackRowToAck(o));
  }
  return out;
}

/**
 * 周知ログシート（日時, 種別, タイトル, 名前, 部署, 施設）を ACK 形式へ変換
 * @param {string[][]} rows
 */
function parseAwarenessLogSheetRows(rows) {
  if (!rows?.length) return [];
  const head = rows[0].map((c) => String(c ?? '').trim());
  const col = (name, fallback) => {
    const i = head.findIndex((h) => h === name);
    return i >= 0 ? i : fallback;
  };
  const ix = {
    confirmedAt: col('日時', 0),
    noticeType: col('種別', 1),
    noticeTitle: col('タイトル', 2),
    staffName: col('名前', 3),
  };
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every((c) => !String(c ?? '').trim())) continue;
    const noticeTitle = String(row[ix.noticeTitle] ?? '').trim();
    const staffName = String(row[ix.staffName] ?? '').trim();
    if (!noticeTitle || !staffName) continue;
    const noticeId = `title:${noticeTitle}`;
    const staffId = `name:${staffName}`;
    const confirmedAt = String(row[ix.confirmedAt] ?? '').trim();
    out.push({
      id: `${noticeId}::${staffId}::${r}`,
      noticeId,
      noticeTitle,
      facilityLinkKey: '',
      staffId,
      staffName,
      confirmedAt,
      noticeType: /事故/.test(String(row[ix.noticeType] ?? '')) ? 'accident' : 'near_miss',
      source: 'sheet',
    });
  }
  return out;
}

/**
 * @param {Record<string, string>} o
 */
function reportRowToNotice(o) {
  let draftSnapshot = {};
  try {
    draftSnapshot = o.draftJson ? JSON.parse(o.draftJson) : {};
  } catch {
    draftSnapshot = {};
  }
  let categories = [];
  if (o.categories) {
    categories = o.categories.split(/[、,]/).map((s) => s.trim()).filter(Boolean);
  }
  return {
    id: o.recordId,
    createdAt: o.createdAt || new Date().toISOString(),
    facilityLinkKey: o.facilityLinkKey || '',
    importance: o.importance === 'high' ? 'high' : 'normal',
    categories,
    title: o.title || '（無題）',
    summary: o.summary || '',
    draftSnapshot,
    archived: o.archived === '1' || o.archived === 'true' || o.archived === 'はい',
    archivedAt: '',
    noticeType:
      /^accident-/i.test(String(o.recordId ?? '')) || categories.some((c) => /事故/.test(String(c)))
        ? 'accident'
        : 'near_miss',
    source: 'sheet',
  };
}

/**
 * @param {Record<string, string>} o
 */
function ackRowToAck(o) {
  return {
    id: o.logId,
    noticeId: o.noticeId,
    facilityLinkKey: o.facilityLinkKey || '',
    staffId: o.staffId || '',
    staffName: o.staffName || '',
    confirmedAt: o.confirmedAt || '',
    source: 'sheet',
  };
}

/**
 * スプレッドシートから報告・確認を読み込み、ローカルとマージ
 * @param {string} sheetsApiKey VITE_GOOGLE_SHEETS_API_KEY
 */
export async function refreshLedgerFromSpreadsheet(sheetsApiKey) {
  const id = getLedgerSpreadsheetId().trim();
  const key = String(sheetsApiKey ?? '').trim();
  if (!id || !key) throw new Error('スプレッドシート ID（求人／台帳）と API キーが必要です');

  const reportRows = await fetchSheetValues(key, id, NEAR_MISS_REPORT_SHEET_NAME, 'A:I');
  const ackRows = await fetchSheetValues(key, id, NEAR_MISS_ACK_SHEET_NAME, 'A:F').catch(() => []);
  const awarenessRows = await fetchSheetValues(key, id, AWARENESS_LOG_SHEET_NAME, 'A:F').catch(() => []);

  const noticesFromSheet = parseReportSheetRows(reportRows);
  const acksFromSheet = [...parseAckSheetRows(ackRows), ...parseAwarenessLogSheetRows(awarenessRows)];

  const byId = new Map(getLedgerNotices().map((n) => [String(n.id), n]));
  for (const n of noticesFromSheet) {
    if (!n.id) continue;
    const prev = byId.get(String(n.id));
    if (!prev) {
      byId.set(String(n.id), n);
    } else {
      const keepDraft =
        prev.draftSnapshot &&
        typeof prev.draftSnapshot === 'object' &&
        Object.keys(prev.draftSnapshot).length
          ? prev.draftSnapshot
          : n.draftSnapshot;
      byId.set(String(n.id), {
        ...prev,
        ...n,
        draftSnapshot: keepDraft,
        archived: Boolean(prev.archived || n.archived),
      });
    }
  }
  writeJson(LS.notices, [...byId.values()]);

  const mergedAcks = [...acksFromSheet, ...getLedgerAcks()];
  const ackByKey = new Map();
  for (const a of mergedAcks) {
    if (!a.noticeId || !a.staffId) continue;
    const k = `${a.noticeId}::${a.staffId}`;
    const prev = ackByKey.get(k);
    if (!prev || String(a.confirmedAt) > String(prev.confirmedAt)) {
      ackByKey.set(k, a);
    }
  }
  writeJson(LS.acks, [...ackByKey.values()]);

  return { notices: getLedgerNotices().length, acks: getLedgerAcks().length };
}

/**
 * @param {string} noticeId
 * @param {string} facilityLinkKey
 */
/**
 * 周知の「確認しました」。端末に保存し、GAS 経由でスプレッドシートにも追記を試みる。
 * @returns {Promise<{ ok: boolean; duplicate?: boolean; ack?: Record<string, unknown>; sheetResult?: Record<string, unknown> }>}
 */
export async function confirmNotice(noticeId, facilityLinkKey) {
  const prof = getStaffProfile();
  if (!prof?.displayName?.trim()) {
    throw new Error('先にスタッフ名を入力・保存してください');
  }
  const fid = String(facilityLinkKey ?? '').trim() || prof.lastFacilityLinkKey;
  if (hasAck(noticeId, prof.staffId)) {
    return { ok: true, duplicate: true };
  }
  const n = getLedgerNotices().find((x) => String(x.id) === String(noticeId));
  const ack = {
    id: newId(),
    noticeId: String(noticeId),
    facilityLinkKey: fid,
    staffId: prof.staffId,
    staffName: prof.displayName.trim(),
    confirmedAt: new Date().toISOString(),
    noticeType: String(n?.noticeType ?? '').trim() || 'near_miss',
    noticeDepartment: String(n?.draftSnapshot?.reporterDept ?? '').trim(),
    noticeTitle: String(n?.title ?? '').trim(),
    source: 'app',
  };
  pushAck(ack);
  const noticeTitle = String(n?.title ?? '');
  const sheetResult = await postAppsScript('appendAwarenessLog', { row: awarenessLogRow(ack, noticeTitle) });
  writeJson(LS.ledgerSheetWriteLast, {
    ...sheetResult,
    action: 'appendAwarenessLog',
    at: new Date().toISOString(),
  });
  maybeAutoArchive(String(noticeId));
  return { ok: true, ack, sheetResult };
}

/**
 * 事故報告データを同一スプレッドシートの別タブへ保存（GAS経由）
 * @param {{ id: string; savedAt: string; facilityLabel: string; department: string; residentId: string; draft: Record<string, unknown> }} entry
 * @returns {Promise<Record<string, unknown>>}
 */
export async function appendAccidentReportToSpreadsheet(entry) {
  const row = [
    String(entry.savedAt ?? new Date().toISOString()),
    String(entry.id ?? ''),
    String(resolveLinkKeyFromFacilityLabel(entry.facilityLabel)),
    String(entry.facilityLabel ?? ''),
    String(entry.department ?? ''),
    String(entry.residentId ?? ''),
    String(entry.draft?.residentName ?? ''),
    String(entry.draft?.occurPlace ?? ''),
    String(entry.draft?.accidentTypeDetail ?? ''),
    JSON.stringify(entry.draft ?? {}).replace(/\r?\n/g, ' '),
  ];
  const sheetResult = await postAppsScript('appendAccidentReport', { row });
  writeJson(LS.ledgerSheetWriteLast, {
    ...sheetResult,
    action: 'appendAccidentReport',
    at: new Date().toISOString(),
  });
  return sheetResult;
}

/**
 * 周知1件に紐づく確認ログか（レコードID・旧形式の title: ・スプレッドシートのタイトル列）
 * @param {Record<string, unknown>} n
 * @param {Record<string, unknown>} a
 */
function ackMatchesNotice(n, a) {
  const nid = String(n?.id ?? '');
  const title = String(n?.title ?? '').trim();
  return (
    String(a?.noticeId ?? '') === nid ||
    String(a?.noticeId ?? '') === `title:${title}` ||
    String(a?.noticeTitle ?? '').trim() === title
  );
}

/**
 * @param {string} noticeId
 */
function maybeAutoArchive(noticeId) {
  const notices = getLedgerNotices();
  const n = notices.find((x) => String(x.id) === String(noticeId));
  if (!n || n.archived) return;
  const link = String(n.facilityLinkKey ?? '');
  const roster = getEffectiveStaffRosterForFacility(link);
  if (!roster.length) return;
  const names = new Set(
    getLedgerAcks()
      .filter((a) => ackMatchesNotice(n, a))
      .map((a) => String(a.staffName ?? '').trim())
  );
  const all = roster.every((r) => names.has(String(r.name).trim()));
  if (!all) return;
  const next = notices.map((x) =>
    String(x.id) === String(noticeId)
      ? { ...x, archived: true, archivedAt: new Date().toISOString() }
      : x
  );
  writeJson(LS.notices, next);
}

/**
 * @param {string} linkKey
 * @param {string} staffId
 */
export function getActiveNoticesForViewer(linkKey, staffId) {
  const lk = String(linkKey ?? '').trim();
  const sid = String(staffId ?? '').trim();
  const list = getLedgerNotices().filter(
    (n) => !n.archived && String(n.facilityLinkKey ?? '') === lk
  );
  const ownAcks = getLedgerAcks().filter(
    (a) =>
      String(a.staffId) === sid || String(a.staffName).trim() === String(getStaffProfile()?.displayName ?? '').trim()
  );
  const ackedIds = new Set(ownAcks.map((a) => String(a.noticeId)));
  const ackedTitles = new Set(
    ownAcks
      .map((a) => String(a.noticeTitle ?? '').trim())
      .filter(Boolean)
  );
  for (const id of ackedIds) {
    if (id.startsWith('title:')) ackedTitles.add(id.slice('title:'.length));
  }
  const pending = list.filter(
    (n) => !ackedIds.has(String(n.id)) && !ackedTitles.has(String(n.title ?? '').trim())
  );
  pending.sort((a, b) => {
    const ah = a.importance === 'high' ? 0 : 1;
    const bh = b.importance === 'high' ? 0 : 1;
    if (ah !== bh) return ah - bh;
    return String(b.createdAt).localeCompare(String(a.createdAt));
  });
  return pending;
}

export function getArchivedNoticesForFacility(linkKey) {
  const lk = String(linkKey ?? '').trim();
  return getLedgerNotices()
    .filter((n) => n.archived && String(n.facilityLinkKey ?? '') === lk)
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
}

/**
 * ポータルトップ用: 全施設の未確認周知件数
 * @param {string} staffId
 */
export function countPendingAllFacilities(staffId) {
  const sid = String(staffId ?? '').trim();
  if (!sid) return 0;
  let n = 0;
  for (const f of CARELINK_FACILITIES) {
    n += getActiveNoticesForViewer(f.linkKey, sid).length;
  }
  return n;
}

/**
 * 管理画面: 施設ごとに「誰がどの周知を未確認か」および名簿上の確認済み者
 * @param {string} linkKey
 * @returns {{ notice: Record<string, unknown>; missing: { id: string; name: string }[]; confirmed: { id: string; name: string }[]; extraAckNames: string[] }[]}
 */
export function buildUnconfirmedMatrix(linkKey) {
  const lk = String(linkKey ?? '').trim();
  const roster = getEffectiveStaffRosterForFacility(lk);
  const notices = getLedgerNotices().filter((n) => !n.archived && String(n.facilityLinkKey ?? '') === lk);
  const acks = getLedgerAcks();
  const rosterNameSet = new Set(roster.map((r) => String(r.name).trim()).filter(Boolean));
  return notices.map((n) => {
    const acksFor = acks.filter((a) => ackMatchesNotice(n, a));
    const confirmedNames = new Set(
      acksFor.map((a) => String(a.staffName ?? '').trim()).filter(Boolean)
    );
    const missing =
      roster.length > 0
        ? roster.filter((r) => !confirmedNames.has(String(r.name).trim()))
        : [{ id: '—', name: '（求人シート同期または手入力名簿を設定してください）' }];
    const confirmed =
      roster.length > 0 ? roster.filter((r) => confirmedNames.has(String(r.name).trim())) : [];
    const extraAckNames = Array.from(
      new Set(
        acksFor
          .map((a) => String(a.staffName ?? '').trim())
          .filter((name) => name && !rosterNameSet.has(name))
      )
    );
    return { notice: n, missing, confirmed, extraAckNames };
  });
}

/**
 * @param {string} yearMonth YYYY-MM
 */
export function aggregateLedgerCategoriesForMonth(yearMonth) {
  const counts = {};
  for (const label of MONTH_CATEGORY_ORDER) counts[label] = 0;
  for (const n of getLedgerNotices()) {
    const ca = String(n.createdAt ?? '');
    if (!ca.startsWith(yearMonth)) continue;
    for (const c of n.categories || []) {
      const k = String(c).trim();
      if (!k) continue;
      if (counts[k] === undefined) counts[k] = 0;
      counts[k] += 1;
    }
  }
  return counts;
}

/**
 * @param {string} apiKey Gemini
 * @param {string} yearMonth
 */
export async function fetchNearMissTrendAssessmentAi(apiKey, yearMonth) {
  const key = String(apiKey ?? '').trim();
  if (!key) return 'API キー未設定のため、件数バーのみ表示します。';
  const counts = aggregateLedgerCategoriesForMonth(yearMonth);
  const lines = Object.entries(counts)
    .filter(([, n]) => n > 0)
    .sort((a, b) => b[1] - a[1])
    .map(([k, v]) => `${k}: ${v}件`);
  const prompt = `あなたは介護施設の安全管理者です。次は「${yearMonth}」のヒヤリハット周知データから集計したカテゴリ別件数です（同一ブラウザ・スプレッドシート連携の合算）。

${lines.join('\n')}

日本語で3〜6文、箇条書き禁止で、多い傾向と現場での注意点を簡潔に述べてください。`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.25, maxOutputTokens: 1024 },
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(key)}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error(data?.error?.message ?? 'AI応答なし');
  return String(text).trim();
}

/**
 * 手動バックアップ用 CSV（報告）
 */
export function exportLedgerReportsCsv() {
  const rows = [REPORT_SHEET_HEADERS.join(',')];
  for (const n of getLedgerNotices()) {
    const line = noticeToReportRow({
      ...n,
      draftSnapshot: n.draftSnapshot,
    })
      .map((c) => `"${String(c).replace(/"/g, '""')}"`)
      .join(',');
    rows.push(line);
  }
  return rows.join('\r\n');
}
