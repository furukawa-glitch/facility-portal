/**
 * 求人・入退社シートからスタッフ名簿を読み、周知チェック対象者を同期する
 */

import { CARELINK_FACILITIES, compactFacilityToken, getShiftDepartmentsForLinkKey } from '../config/carelinkFacilities.js';
import { DEFAULT_HR_SPREADSHEET_ID } from '../config/hrSpreadsheetConstants.js';

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

/** VITE_HR_SPREADSHEET_ID が無ければデフォルト（ユーザー共有URL） */
export function getHrSpreadsheetId() {
  const a = import.meta.env.VITE_HR_SPREADSHEET_ID?.trim();
  return a || DEFAULT_HR_SPREADSHEET_ID;
}

function normCell(s) {
  return String(s ?? '')
    .trim()
    .replace(/\s/g, '');
}

/**
 * @param {string} statusRaw
 */
export function isActiveEmploymentStatus(statusRaw) {
  const t = String(statusRaw ?? '').trim();
  if (!t) return true;
  const n = normCell(t);
  if (/退職|離職|解雇|終了|不要|辞退|入社前のみ|内定のみ|不採用|見送|×|✕/u.test(t)) return false;
  if (/長期休職|休職中(?![のを])|育休(?!.*復帰)/u.test(t)) return false;
  if (/在籍|勤務中|現職|稼働中|入社済|契約中|本採用|ｏｋ|OK|○|正社員|パート|非常勤|嘱託|アルバイト|派遣/u.test(t)) return true;
  return true;
}

/**
 * 求人シートの「タグ」「#施設A」などをトークン化
 * @param {string} raw
 * @returns {string[]}
 */
export function splitHrTagTokens(raw) {
  return String(raw ?? '')
    .split(/[#、，,｜|\s\n\r\t／/]+/u)
    .map((x) => x.trim())
    .filter((x) => x && !/^例[：:]?$/u.test(x));
}

/**
 * 単一トークンまたは短文から linkKey を推定
 * @param {string} token
 */
function linkKeyFromSingleToken(token) {
  const raw = String(token ?? '').trim();
  if (!raw) return '';
  const hit = CARELINK_FACILITIES.find(
    (f) =>
      raw === f.tabLabel ||
      raw === f.linkKey ||
      raw === f.sheetTitle ||
      compactFacilityToken(raw) === compactFacilityToken(f.tabLabel) ||
      compactFacilityToken(raw) === compactFacilityToken(f.linkKey) ||
      raw.includes(f.tabLabel) ||
      f.tabLabel.includes(raw)
  );
  return hit?.linkKey ?? '';
}

/**
 * 施設セル・タグ列の全文から linkKey を推定（タグ複数可）
 * @param {string} facilityCell
 * @param {string} [tagCell]
 * @returns {string[]}
 */
/**
 * 施設・タグの文字列から、勤務表画面の「部署」候補（shiftDepartments）のいずれかを推定する。
 * タグに「千音寺介護」「#愛西デイ」のように**アプリの部署名が含まれる**と確実です。
 * @param {string} facilityLinkKey
 * @param {string} facilityCell
 * @param {string} tagCell
 * @returns {string}
 */
export function resolveShiftDepartmentForHrRow(facilityLinkKey, facilityCell, tagCell) {
  const lk = String(facilityLinkKey ?? '').trim();
  const options = getShiftDepartmentsForLinkKey(lk).filter(Boolean);
  if (!options.length) return '';

  const blobRaw = `${String(facilityCell ?? '')} ${String(tagCell ?? '')}`.trim();
  const blobNorm = normCell(blobRaw.replace(/\s+/g, ''));
  const sorted = [...options].sort((a, b) => b.length - a.length);

  for (const dep of sorted) {
    const d = normCell(dep);
    if (!d) continue;
    if (blobNorm.includes(d) || blobRaw.includes(dep)) return dep;
  }

  const tokens = [...splitHrTagTokens(facilityCell), ...splitHrTagTokens(tagCell)];
  for (const dep of sorted) {
    const dn = normCell(dep);
    if (!dn) continue;
    for (const tok of tokens) {
      const t = normCell(tok);
      if (!t) continue;
      if (t === dn || t.includes(dn) || dn.includes(t)) return dep;
    }
  }
  return '';
}

export function linkKeysFromFacilityAndTags(facilityCell, tagCell = '') {
  const combined = [String(facilityCell ?? '').trim(), String(tagCell ?? '').trim()]
    .filter(Boolean)
    .join(' ');
  const whole = linkKeyFromFacilityCell(combined);
  if (whole) return [whole];

  const tokens = [
    ...splitHrTagTokens(facilityCell),
    ...splitHrTagTokens(tagCell),
  ];
  const keys = new Set();
  for (const t of tokens) {
    const lk = linkKeyFromSingleToken(t);
    if (lk) keys.add(lk);
  }
  if (!keys.size && combined) {
    for (const piece of splitHrTagTokens(combined)) {
      const lk2 = linkKeyFromSingleToken(piece);
      if (lk2) keys.add(lk2);
    }
  }
  return [...keys];
}

function linkKeyFromFacilityCell(facilityCell) {
  const raw = String(facilityCell ?? '').trim();
  if (!raw) return '';
  return linkKeyFromSingleToken(raw);
}

/**
 * @param {string[]} headers
 */
function findStaffHeaderIndices(headers) {
  const h = headers.map((x) => normCell(String(x ?? '')));
  const nameIdx = h.findIndex((cell) =>
    /^(スタッフ名|氏名|名前|職員名|担当者名|スタッフ)$/u.test(cell) || /スタッフ名|氏名/u.test(cell)
  );
  let statusIdx = h.findIndex((cell) =>
    /^(在籍状況|状況|ステータス|入退社|雇用状況|勤務状況)$/u.test(cell) || /在籍|入退社|雇用/u.test(cell)
  );
  if (statusIdx < 0) {
    statusIdx = h.findIndex((cell) => cell.includes('在籍') || cell.includes('状況'));
  }
  const facilityIdx = h.findIndex((cell) =>
    /^(施設|事業所|拠点|所属|勤務先|ホーム|ブロック)$/u.test(cell) || /施設|事業所|所属/u.test(cell)
  );
  const tagIdx = h.findIndex((cell) =>
    /^(タグ|部署|部門|区分|カテゴリ|エリア|案件)$/u.test(cell) || /タグ|部署|ハッシュ|#/.test(cell)
  );
  return { nameIdx, statusIdx, facilityIdx, tagIdx };
}

/**
 * ヘッダ行を探す（1行目でなくても可）
 * @param {string[][]} rows
 */
function findHeaderRowIndex(rows) {
  const max = Math.min(45, rows?.length ?? 0);
  for (let r = 0; r < max; r++) {
    const headers = (rows[r] ?? []).map((c) => String(c ?? '').trim());
    const idx = findStaffHeaderIndices(headers);
    if (idx.nameIdx >= 0) return { headerRow: r, headers, idx };
  }
  throw new Error('スタッフ名（氏名）のヘッダ行が見つかりません。');
}

/**
 * @param {string[][]} rows
 * @returns {{ byFacility: Record<string, { id: string; name: string }[]>; global: { id: string; name: string }[] | null; meta: { rowCount: number; hasFacilityCol: boolean; hasTagCol: boolean } }}
 */
export function parseStaffRowsFromHrSheet(rows) {
  if (!rows?.length) {
    return {
      byFacility: {},
      global: [],
      meta: { rowCount: 0, hasFacilityCol: false, hasTagCol: false },
    };
  }
  const { headerRow, idx } = findHeaderRowIndex(rows);
  const hasFacilityCol = idx.facilityIdx >= 0;
  const hasTagCol = idx.tagIdx >= 0;
  /** 施設列またはタグ列のどちらかがあれば、施設別に振り分ける */
  const usePerFacility = hasFacilityCol || hasTagCol;
  const byFacility = {};
  const global = [];
  let rowCount = 0;

  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const name = String(row[idx.nameIdx] ?? '').trim();
    if (!name) continue;
    if (/合計|小計|^氏名$|^スタッフ名$|例[：:]/u.test(name)) continue;

    const status = idx.statusIdx >= 0 ? String(row[idx.statusIdx] ?? '').trim() : '';
    if (idx.statusIdx >= 0 && !isActiveEmploymentStatus(status)) continue;

    rowCount += 1;
    const baseId = `hr-${headerRow}-${r}-${normCell(name).slice(0, 20)}`;

    if (!usePerFacility) {
      global.push({ id: baseId, name });
      continue;
    }

    const fac = hasFacilityCol ? String(row[idx.facilityIdx] ?? '').trim() : '';
    const tag = hasTagCol ? String(row[idx.tagIdx] ?? '').trim() : '';
    if (!fac && !tag) {
      for (const f of CARELINK_FACILITIES) {
        if (!byFacility[f.linkKey]) byFacility[f.linkKey] = [];
        byFacility[f.linkKey].push({ id: `${baseId}-${f.linkKey}`, name });
      }
      continue;
    }

    const lkList = linkKeysFromFacilityAndTags(fac, tag);
    if (lkList.length) {
      for (const lk of lkList) {
        if (!byFacility[lk]) byFacility[lk] = [];
        byFacility[lk].push({ id: `${baseId}-${lk}`, name });
      }
    } else {
      if (!byFacility.__unmapped__) byFacility.__unmapped__ = [];
      byFacility.__unmapped__.push({ id: baseId, name, _facilityRaw: fac, _tagRaw: tag });
    }
  }

  return {
    byFacility,
    global: usePerFacility ? null : global,
    meta: { rowCount, hasFacilityCol: hasFacilityCol || hasTagCol, hasTagCol },
  };
}

/**
 * 求人シートの行から「勤務希望・勤務表」用の (施設linkKey, 氏名, 部署) を組み立てる。
 * 部署は resolveShiftDepartmentForHrRow（タグ／施設セルにアプリの部署名が含まれること）で推定。
 * @param {string[][]} rows
 * @returns {{ items: { linkKey: string; staffName: string; department: string }[]; warnings: string[] }}
 */
export function parseHrRowsForShiftPreferenceSeed(rows) {
  /** @type {{ items: { linkKey: string; staffName: string; department: string }[]; warnings: string[] }} */
  const out = { items: [], warnings: [] };
  const pushWarn = (msg) => {
    if (out.warnings.length < 35) out.warnings.push(msg);
  };
  if (!rows?.length) {
    pushWarn('シートが空です。');
    return out;
  }

  let headerRow;
  let idx;
  try {
    const found = findHeaderRowIndex(rows);
    headerRow = found.headerRow;
    idx = found.idx;
  } catch (e) {
    pushWarn(e instanceof Error ? e.message : 'ヘッダ行を読めませんでした。');
    return out;
  }

  const hasFacilityCol = idx.facilityIdx >= 0;
  const hasTagCol = idx.tagIdx >= 0;
  if (!hasFacilityCol && !hasTagCol) {
    pushWarn('「施設」または「タグ／部署」列がないため、施設・部署別の取り込みができません。');
    return out;
  }

  const seen = new Set();
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const name = String(row[idx.nameIdx] ?? '').trim();
    if (!name) continue;
    if (/合計|小計|^氏名$|^スタッフ名$|例[：:]/u.test(name)) continue;

    const status = idx.statusIdx >= 0 ? String(row[idx.statusIdx] ?? '').trim() : '';
    if (idx.statusIdx >= 0 && !isActiveEmploymentStatus(status)) continue;

    const fac = hasFacilityCol ? String(row[idx.facilityIdx] ?? '').trim() : '';
    const tag = hasTagCol ? String(row[idx.tagIdx] ?? '').trim() : '';
    if (!fac && !tag) continue;

    const lkList = linkKeysFromFacilityAndTags(fac, tag);
    if (!lkList.length) {
      pushWarn(`「${name}」: 施設・タグからアプリの施設にマップできませんでした（${fac || '—'} / ${tag || '—'}）。`);
      continue;
    }

    for (const lk of lkList) {
      let department = resolveShiftDepartmentForHrRow(lk, fac, tag);
      const deptOpts = getShiftDepartmentsForLinkKey(lk).filter(Boolean);
      if (!department) {
        if (deptOpts.length === 1) {
          department = deptOpts[0];
        } else {
          pushWarn(
            `「${name}」（${lk}）: タグに勤務表の部署名（例: ${deptOpts.slice(0, 3).join('、')}）を含めてください。`
          );
          continue;
        }
      }
      if (!deptOpts.includes(department)) {
        pushWarn(`「${name}」: 部署「${department}」は ${lk} の候補外です。`);
        continue;
      }

      const dedupe = `${lk}\t${name}\t${department}`;
      if (seen.has(dedupe)) continue;
      seen.add(dedupe);
      out.items.push({ linkKey: lk, staffName: name, department });
    }
  }

  if (!out.items.length && !out.warnings.length) {
    pushWarn('在籍として有効な行が見つかりませんでした。');
  }
  return out;
}

function sheetRangeA1(sheetTitle, range) {
  const safe = `'${String(sheetTitle).replace(/'/g, "''")}'`;
  return `${safe}!${range}`;
}

/**
 * @param {string} apiKey
 * @param {string} spreadsheetId
 */
export async function fetchSpreadsheetSheetTitles(apiKey, spreadsheetId) {
  const sid = encodeURIComponent(spreadsheetId);
  const url = `${SHEETS_API}/${sid}?fields=sheets(properties(title))&key=${encodeURIComponent(apiKey)}`;
  const res = await fetch(url);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error?.message ?? 'スプレッドシートのタブ一覧が取得できません');
  const sheets = data.sheets ?? [];
  return sheets.map((s) => String(s.properties?.title ?? '').trim()).filter(Boolean);
}

/**
 * @param {string[]} titles
 * @param {string} [preferredSheetName] env や UI の上書き
 */
export function pickStaffSheetTitle(titles, preferredSheetName = '') {
  const pref = String(preferredSheetName ?? '').trim();
  if (pref && titles.includes(pref)) return pref;
  const priority = [/求人|入退社|スタッフ|社員|採用|人事|名簿|在籍/u];
  for (const re of priority) {
    const hit = titles.find((t) => re.test(t));
    if (hit) return hit;
  }
  return titles[0] ?? '';
}

/**
 * @param {string} apiKey
 * @param {string} spreadsheetId
 * @param {string} sheetTitle
 */
async function fetchSheetValues(apiKey, spreadsheetId, sheetTitle, rangeA1 = 'A:ZZ') {
  const sid = encodeURIComponent(spreadsheetId);
  const a1 = encodeURIComponent(sheetRangeA1(sheetTitle, rangeA1));
  const url = `${SHEETS_API}/${sid}/values/${a1}?key=${encodeURIComponent(apiKey)}`;
  const res = await fetch(url);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data?.error?.message ?? `シート「${sheetTitle}」の取得に失敗しました`);
  return data.values ?? [];
}

/**
 * 求人スプレッドシートを1回取得し、名簿用パースと勤務表シード用パースの両方に使う。
 * @param {string} apiKey
 * @param {{ preferredSheetTitle?: string }} [opts]
 */
export async function fetchHrStaffSheetBundle(apiKey, opts = {}) {
  const key = String(apiKey ?? '').trim();
  const id = getHrSpreadsheetId();
  if (!key) throw new Error('VITE_GOOGLE_SHEETS_API_KEY が必要です');

  const titles = await fetchSpreadsheetSheetTitles(key, id);
  if (!titles.length) throw new Error('スプレッドシートにタブがありません');

  const envName = import.meta.env.VITE_HR_STAFF_SHEET_NAME?.trim();
  const sheetTitle = pickStaffSheetTitle(titles, opts.preferredSheetTitle || envName || '');
  if (!sheetTitle) throw new Error('スタッフ名簿のシート名を特定できません');

  const rows = await fetchSheetValues(key, id, sheetTitle, 'A:ZZ');
  const roster = parseStaffRowsFromHrSheet(rows);
  const shiftSeed = parseHrRowsForShiftPreferenceSeed(rows);

  return {
    sheetTitle,
    spreadsheetId: id,
    roster,
    shiftSeed,
  };
}

/**
 * 求人シートからスタッフを読み、localStorage の同期名簿を更新
 * @param {string} apiKey
 * @param {{ preferredSheetTitle?: string }} [opts]
 */
export async function syncStaffRosterFromHrSheet(apiKey, opts = {}) {
  const bundle = await fetchHrStaffSheetBundle(apiKey, opts);
  return {
    sheetTitle: bundle.sheetTitle,
    spreadsheetId: bundle.spreadsheetId,
    ...bundle.roster,
    shiftSeed: bundle.shiftSeed,
    syncedAt: new Date().toISOString(),
  };
}
