/**
 * CareLink OS — 監査ログ・異常検知・救急サマリー・月次出力・AIアドバイス補助
 * 永続化: localStorage（本番は API 差し替え想定）
 */

import { buildNearMissReportHtml, NEAR_MISS_CATEGORY_LABELS } from './nearMissReportHtml.js';
import { CARELINK_FACILITIES, facilityDefBySheetTitle } from '../config/carelinkFacilities.js';

export { buildNearMissReportHtml, NEAR_MISS_CATEGORY_LABELS };

/** @param {Record<string, unknown>} resident */
function nursingLinkKeyForResident(resident) {
  const st = String(resident?.sourceSheetTitle ?? '').trim();
  const fromTitle = facilityDefBySheetTitle(st);
  if (fromTitle) return fromTitle.linkKey;
  const fac = String(resident?.facility ?? '').trim();
  if (!fac) return '';
  const hit = CARELINK_FACILITIES.find(
    (f) =>
      fac === f.sheetTitle ||
      fac === f.tabLabel ||
      fac === f.linkKey ||
      fac.includes(f.tabLabel) ||
      f.tabLabel.includes(fac)
  );
  return hit?.linkKey ?? '';
}

const LS = {
  careEvents: 'carelink_os_care_events_v1',
  nursing: 'carelink_os_nursing_directives_v1',
  weeklyPlans: 'carelink_os_weekly_plans_v1',
  lastStool: 'carelink_os_last_stool_v1',
  /** 最終の排尿記録・トイレ誘導実施時刻（6時間アラート用） */
  lastUrine: 'carelink_os_last_urine_v1',
  /** 排便量「小」「少量」の連続回数（2回で排便1回相当として lastStool を更新） */
  smallStoolTally: 'carelink_os_small_stool_tally_v1',
  vitals: 'carelink_os_resident_vitals_v1',
  emergencyContact: 'carelink_os_emergency_contact_v1',
  seeded: 'carelink_os_demo_seeded_v1',
  accidentReports: 'carelink_os_accident_reports_v1',
  nearMissReports: 'carelink_os_near_miss_reports_v1',
  visitNursingSpecial: 'carelink_os_visit_nursing_special_v1',
};

const MAX_ACCIDENT_REPORTS = 2000;
const MAX_NEAR_MISS_REPORTS = 2000;

/** 月次分析で並べる事故種別の表示順 */
export const ACCIDENT_TYPE_ORDER = Object.freeze([
  '転倒',
  '転落',
  '落薬',
  '誤薬',
  '窒息・誤嚥',
  '徘徊',
  'やけど・火傷',
  '自傷行為',
  'その他',
]);

/** 時間帯スロットの表示順 */
export const ACCIDENT_SLOT_ORDER = Object.freeze([
  '深夜（0–5時）',
  '早朝（6–8時）',
  '午前（9–11時）',
  '昼（12–13時）',
  '午後（14–17時）',
  '夕方（18–20時）',
  '夜（21–23時）',
  '時間不明',
]);

/** ヒヤリ月次のカテゴリ表示順（複数選択は件数に重複加算） */
export const NEAR_MISS_MONTH_CATEGORY_ORDER = Object.freeze([
  ...NEAR_MISS_CATEGORY_LABELS,
  'その他',
  '分類なし',
]);

const MODEL = 'gemini-2.0-flash';

/**
 * Gemini generateContent の失敗を、画面上で「文字化け／英語だらけ」と誤解されないよう日本語中心にまとめる。
 * @param {any} data - JSON レスポンス
 * @param {number} httpStatus
 */
function formatGeminiGenerateContentErrorMessage(data, httpStatus) {
  const err = data?.error;
  const raw = String(err?.message ?? '').trim();
  const code = err?.code ?? err?.status;
  const lower = raw.toLowerCase();
  const st = Number(httpStatus) || 0;

  const retryMatch = raw.match(/retry\s+in\s+([\d.]+)\s*s/i);
  const retrySec = retryMatch ? Math.ceil(Number.parseFloat(retryMatch[1])) : null;

  const quotaLike =
    st === 429 ||
    code === 429 ||
    lower.includes('quota') ||
    lower.includes('resource_exhausted') ||
    lower.includes('rate limit') ||
    lower.includes('exceeded');

  if (quotaLike) {
    const waitLine =
      retrySec != null && retrySec > 0
        ? `目安として約 ${retrySec} 秒待ってから、再度「アセスメント生成」をお試しください。`
        : 'しばらく時間をおいてから、再度「アセスメント生成」をお試しください。';
    return [
      '【APIの利用上限に達しています】',
      '表示されていた英語は文字化けではなく、Google Gemini の「回数・トークン枠の超過」です。',
      '',
      waitLine,
      '',
      '確認のヒント:',
      '・ Google AI Studio でキー・プラン・利用枠（https://aistudio.google.com）',
      '・ このアプリの VITE_GEMINI_API_KEY が正しいか（無料枠の limit が 0 の場合は別キーや課金設定が必要なことがあります）',
      '',
      '────────',
      '（APIメッセージ）',
      raw || `HTTP ${st || '?'}`,
    ].join('\n');
  }

  if (
    st === 401 ||
    st === 403 ||
    lower.includes('api key not valid') ||
    lower.includes('invalid api key') ||
    lower.includes('permission denied')
  ) {
    return [
      '【APIキーまたは権限の問題です】',
      'キーが無効、またはこのモデル（' + MODEL + '）を呼び出す権限がありません。',
      '.env の VITE_GEMINI_API_KEY と Google 側の有効化を確認してください。',
      '',
      '────────',
      '（APIメッセージ）',
      raw || `HTTP ${st || '?'}`,
    ].join('\n');
  }

  return [
    '【AIの応答を取得できませんでした】',
    '通信やサーバー側の都合の可能性があります。時間をおいて再度お試しください。',
    '',
    '────────',
    '（APIメッセージ）',
    raw || `HTTP ${st || '?'} / code: ${code ?? '—'}`,
  ].join('\n');
}

/**
 * バイタル・排便・巡視の閾値（一覧カードのアラートに使用）
 *
 * - 体温 ${tempCMinFever}℃ 以上 → vital 異常
 * - 収縮期血圧 ${bpSystolicHigh} 以上、または拡張期 ${bpDiastolicLow} 以下 → vital 異常
 * - 最終排便（実効）から ${stoolHoursMax} 時間超 → 排便アラート（「小」「少量」は2回で1回として時刻更新）
 * - 最終排尿（尿量記録・トイレ誘導・排泄確認など）から ${urineHoursMax} 時間超 → 排尿アラート
 * - 名簿の巡視間隔（分）が ${patrolIntervalWarnMin} 超 → warn（赤 critical にはしない）
 */
export const VITAL_THRESHOLDS = Object.freeze({
  tempCMinFever: 37.5,
  bpSystolicHigh: 150,
  bpDiastolicLow: 80,
  stoolHoursMax: 72,
  urineHoursMax: 6,
  patrolIntervalWarnMin: 180,
});

/** @returns {Record<string, number>} */
function readSmallStoolTallyMap() {
  const raw = readJson(LS.smallStoolTally, {});
  return raw && typeof raw === 'object' ? raw : {};
}

function writeSmallStoolTallyMap(map) {
  writeJson(LS.smallStoolTally, map);
}

function getSmallStoolTally(residentId) {
  const m = readSmallStoolTallyMap();
  const n = m[String(residentId)];
  return typeof n === 'number' && n >= 0 ? n : 0;
}

function setSmallStoolTally(residentId, n) {
  const m = readSmallStoolTallyMap();
  if (n <= 0) delete m[String(residentId)];
  else m[String(residentId)] = n;
  writeSmallStoolTallyMap(m);
}

/**
 * 排便が記録されたとき、72時間アラート用の「最終排便時刻」をどう進めるか。
 * クイックの「多・中・小」／詳細画面の「多量・中等量・少量」に対応。
 * 「小」「少量」は2回に1回だけ時刻を更新（2回分で排便1回とみなす）。
 *
 * @param {string} residentId
 * @param {{ stoolVolume?: string; stoolAmount?: string; stoolCharacter?: string }} [opts]
 */
export function recordStoolForIntervalAlert(residentId, opts = {}) {
  const id = String(residentId ?? '').trim();
  if (!id) return;

  const sv = String(opts.stoolVolume ?? '').trim();
  const sa = String(opts.stoolAmount ?? '').trim();
  const sc = String(opts.stoolCharacter ?? '').trim();
  const vol = sv || sa;

  const isSmall = vol === '小' || vol === '少量';
  const isFull =
    vol === '多' ||
    vol === '中' ||
    vol === '多量' ||
    vol === '中等量';

  if (isFull) {
    setLastStoolNow(id);
    setSmallStoolTally(id, 0);
    return;
  }
  if (isSmall) {
    const next = getSmallStoolTally(id) + 1;
    if (next >= 2) {
      setLastStoolNow(id);
      setSmallStoolTally(id, 0);
    } else {
      setSmallStoolTally(id, next);
    }
    return;
  }
  if (!vol && sc) {
    setLastStoolNow(id);
    setSmallStoolTally(id, 0);
    return;
  }
}

/** 訪問看護・特別指示の人数がこの値以上のとき、減算管理上の注意喚起（算定要件は事業所・最新告示で確認） */
export const VISIT_NURSING_SPECIAL_WARN_THRESHOLD = 19;

/** 令和8年度報酬改定資料・実務相談Q&A の要約（本番は PDF 全文読込に差し替え可） */
export const REGULATORY_KNOWLEDGE_BASE = `
【令和8年度 介護報酬改定の考え方（委託資料・要約）】
- サービス提供記録（巡視・バイタル・食事・排泄等）は、提供実態の証跡として監査・検証で重視される。
- 身体拘束適正化・安全配慮義務に基づき、バイタル急変・排便異常時は観察記録と医療・看護への報告連携が求められる。
- 排便ケア: 長期無排便は腸閉塞・褥瘡悪化等のリスク。下剤・浣腸は医師・看護指示に基づき実施し結果を記録する。
- 感染症: 発熱時は隔離・消毒記録、必要に応じた受診・往診とその記録。

【実務相談Q&A（抜粋・要約）】
Q: 排便がないが食事はある。 A: 腹部症状・腸蠕動の観察、触診は指示のもとで。医師・看護へ相談。下剤の自己増量は避ける。
Q: 血圧が高い。 A: 安静・再測定、平常値との比較、指示薬の確認。基準超過は主治医報告を検討。
Q: 下剤が必要か。 A: 無排便日数・腹部所見を踏まえ医師・看護判断。実施したら種類・量・結果を記録。
`.trim();

function readJson(key, fallback) {
  try {
    const s = localStorage.getItem(key);
    if (!s) return fallback;
    return JSON.parse(s);
  } catch {
    return fallback;
  }
}

function writeJson(key, val) {
  localStorage.setItem(key, JSON.stringify(val));
}

/** @returns {{ temp?: string; bpUpper?: string; bpLower?: string; pulse?: string; spo2?: string; weight?: string; updatedAt?: string } | null} */
export function getResidentVitalSnapshot(residentId) {
  const all = readJson(LS.vitals, {});
  return all[String(residentId)] ?? null;
}

export function setResidentVitalSnapshot(residentId, patch) {
  const all = readJson(LS.vitals, {});
  const prev = all[String(residentId)] ?? {};
  all[String(residentId)] = { ...prev, ...patch, updatedAt: new Date().toISOString() };
  writeJson(LS.vitals, all);
}

const VISIT_NURSING_YMD = /^\d{4}-\d{2}-\d{2}$/;

function visitNursingSpecialLocalYmd(d = new Date()) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * 手動登録データが「今日の日付」でカウントに効くか（チェック ON かつ期間内）
 * @param {{ active?: boolean; periodStart?: string; periodEnd?: string }} vn
 */
function visitNursingManualActive(vn, now = new Date()) {
  if (!vn || !vn.active) return false;
  const today = visitNursingSpecialLocalYmd(now);
  const ps = String(vn.periodStart ?? '').trim();
  const pe = String(vn.periodEnd ?? '').trim();
  if (ps && VISIT_NURSING_YMD.test(ps) && today < ps) return false;
  if (pe && VISIT_NURSING_YMD.test(pe) && today > pe) return false;
  return true;
}

/**
 * 訪問看護で特別指示が付いた利用者のフラグ（同一ブラウザ・localStorage）
 * 終了日を過ぎた手動登録は読み取り時に active を false に戻す
 * @returns {{ active: boolean; note: string; periodStart: string; periodEnd: string; updatedAt?: string }}
 */
export function getVisitNursingSpecial(residentId) {
  const all = readJson(LS.visitNursingSpecial, {});
  const id = String(residentId ?? '').trim();
  const row = all[id];
  if (!row || typeof row !== 'object') {
    return { active: false, note: '', periodStart: '', periodEnd: '', updatedAt: '' };
  }
  const periodStart = String(row.periodStart ?? '').trim();
  const periodEnd = String(row.periodEnd ?? '').trim();
  let active = Boolean(row.active);
  const today = visitNursingSpecialLocalYmd();
  if (active && periodEnd && VISIT_NURSING_YMD.test(periodEnd) && today > periodEnd) {
    active = false;
    all[id] = { ...row, active: false, updatedAt: new Date().toISOString() };
    writeJson(LS.visitNursingSpecial, all);
  }
  return {
    active,
    note: String(row.note ?? ''),
    periodStart,
    periodEnd,
    updatedAt: row.updatedAt != null ? String(row.updatedAt) : '',
  };
}

/**
 * @param {string} residentId
 * @param {{ active?: boolean; note?: string; periodStart?: string; periodEnd?: string }} patch
 */
export function setVisitNursingSpecial(residentId, patch) {
  const id = String(residentId ?? '').trim();
  if (!id) return;
  const all = readJson(LS.visitNursingSpecial, {});
  const prev = all[id] && typeof all[id] === 'object' ? all[id] : {};
  const active = patch?.active !== undefined ? Boolean(patch.active) : Boolean(prev.active);
  const note =
    patch?.note !== undefined ? String(patch.note ?? '').trim() : String(prev.note ?? '').trim();
  const periodStart =
    patch?.periodStart !== undefined
      ? String(patch.periodStart ?? '').trim()
      : String(prev.periodStart ?? '').trim();
  const periodEnd =
    patch?.periodEnd !== undefined
      ? String(patch.periodEnd ?? '').trim()
      : String(prev.periodEnd ?? '').trim();
  all[id] = {
    ...prev,
    active,
    note,
    periodStart,
    periodEnd,
    updatedAt: new Date().toISOString(),
  };
  writeJson(LS.visitNursingSpecial, all);
}

/** 手動の「該当」が今の日付で集計に効いているか（名簿検出バッジの判定など） */
export function visitNursingManualRegistrationActive(residentId) {
  return visitNursingManualActive(getVisitNursingSpecial(String(residentId ?? '')));
}

/** 名簿の「医療保険」列から、訪問看護＋特別指示と読み取れるか（読み取りのみ） */
export function sheetSuggestsVisitNursingSpecial(insuranceLabelRaw) {
  const s = String(insuranceLabelRaw ?? '');
  if (!s.trim()) return false;
  return /訪問看護/u.test(s) && /特別指示|特指示/u.test(s);
}

/**
 * 訪問看護・特別指示としてカウント（アプリで「該当」登録した利用者、または名簿文言の自動検出）
 * @param {Record<string, unknown>} resident
 */
export function residentHasVisitNursingSpecial(resident) {
  const id = String(resident?.id ?? '');
  if (visitNursingManualActive(getVisitNursingSpecial(id))) return true;
  return sheetSuggestsVisitNursingSpecial(resident?.insuranceLabel);
}

/** @param {Record<string, unknown>[]} residents */
export function countVisitNursingSpecialAmong(residents) {
  let n = 0;
  for (const r of residents) {
    if (residentHasVisitNursingSpecial(r)) n += 1;
  }
  return n;
}

/** @param {{ temp?: string; bpUpper?: string; bpLower?: string }} v */
export function detectVitalAbnormal(v) {
  const flags = [];
  const t = parseFloat(String(v.temp ?? '').replace(',', '.'));
  const sys = parseFloat(String(v.bpUpper ?? '').replace(',', '.'));
  const dia = parseFloat(String(v.bpLower ?? '').replace(',', '.'));
  if (!Number.isNaN(t) && t >= VITAL_THRESHOLDS.tempCMinFever) {
    flags.push({ code: 'fever', label: `体温 ${t}℃（${VITAL_THRESHOLDS.tempCMinFever}℃以上）` });
  }
  if (!Number.isNaN(sys) && sys >= VITAL_THRESHOLDS.bpSystolicHigh) {
    flags.push({ code: 'bp_sys_high', label: `収縮期血圧 ${sys}（${VITAL_THRESHOLDS.bpSystolicHigh}以上）` });
  }
  if (!Number.isNaN(dia) && dia <= VITAL_THRESHOLDS.bpDiastolicLow) {
    flags.push({ code: 'bp_dia_low', label: `拡張期血圧 ${dia}（${VITAL_THRESHOLDS.bpDiastolicLow}以下）` });
  }
  return flags;
}

export function getLastStoolIso(residentId) {
  const all = readJson(LS.lastStool, {});
  return all[String(residentId)] ?? null;
}

export function setLastStoolNow(residentId) {
  setLastStoolIso(residentId, new Date().toISOString());
}

export function setLastStoolIso(residentId, iso) {
  const all = readJson(LS.lastStool, {});
  all[String(residentId)] = iso;
  writeJson(LS.lastStool, all);
}

/** @param {string} residentId */
export function getLastUrineIso(residentId) {
  const all = readJson(LS.lastUrine, {});
  return all[String(residentId)] ?? null;
}

/** 排尿記録・トイレ誘導・簡易排泄確認のいずれかがあったときに呼ぶ（6時間アラートの基準時刻を更新） */
export function setLastUrineNow(residentId) {
  const all = readJson(LS.lastUrine, {});
  all[String(residentId)] = new Date().toISOString();
  writeJson(LS.lastUrine, all);
}

/**
 * 最終排尿（本端末ログ）からの経過時間（時間）。一度も記録がないときは null（アラートなし）。
 * @param {Record<string, unknown>} resident
 * @param {Date} [now]
 */
export function getHoursSinceLastUrine(resident, now = new Date()) {
  const iso = getLastUrineIso(String(resident.id ?? ''));
  if (!iso) return null;
  return (now.getTime() - new Date(iso).getTime()) / 3600000;
}

/**
 * 最終排便からの経過時間（時間）。記録なしは null。
 * @param {Record<string, unknown>} resident
 * @param {Date} [now]
 */
export function getHoursSinceLastStool(resident, now = new Date()) {
  const id = String(resident.id ?? '');
  const iso = getLastStoolIso(id);
  if (iso) return (now.getTime() - new Date(iso).getTime()) / 3600000;
  const raw = resident.lastStoolDate;
  if (raw == null || raw === '' || raw === '—') return null;
  const parsed = parseSheetStoolDate(raw, now);
  if (!parsed) return null;
  return (now.getTime() - parsed.getTime()) / 3600000;
}

/** @param {unknown} raw 例 4/2, 2026/4/2 */
function parseSheetStoolDate(raw, now) {
  const s = String(raw).trim();
  const y = now.getFullYear();
  const m = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/.exec(s);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  const m2 = /^(\d{1,2})[\/\-](\d{1,2})$/.exec(s);
  if (m2) return new Date(y, Number(m2[1]) - 1, Number(m2[2]), 12, 0, 0);
  return null;
}

/**
 * @param {Record<string, unknown>} resident
 * @param {Date} [now]
 */
export function evaluateResidentMonitor(resident, now = new Date()) {
  const snap = getResidentVitalSnapshot(String(resident.id));
  const vitalFlags = snap ? detectVitalAbnormal(snap) : [];
  const stoolH = getHoursSinceLastStool(resident, now);
  const stoolBad = stoolH != null && stoolH >= VITAL_THRESHOLDS.stoolHoursMax;
  const urineH = getHoursSinceLastUrine(resident, now);
  const urineBad = urineH != null && urineH >= VITAL_THRESHOLDS.urineHoursMax;
  const patrolBad = Number(resident.patrolIntervalMinutes) > VITAL_THRESHOLDS.patrolIntervalWarnMin;
  const vitalBad = vitalFlags.length > 0;
  const critical = vitalBad || stoolBad || urineBad;
  const level = critical ? 'critical' : patrolBad ? 'warn' : 'ok';
  return {
    vitalFlags,
    vitalBad,
    stoolBad,
    stoolHours: stoolH,
    urineBad,
    urineHours: urineH,
    patrolBad,
    level,
    snapshot: snap,
  };
}

/**
 * 介護報酬の減算・監査観点の「要確認」候補（断定しない）
 * @param {Record<string, unknown>} resident
 * @param {ReturnType<typeof evaluateResidentMonitor>} monitorEv
 */
export function evaluateReimbursementDeductionAlerts(resident, monitorEv) {
  const lines = [];
  if (monitorEv.patrolBad) {
    lines.push('巡視間隔が長く空いています。サービス提供体制・減算の有無を事業所ルールで確認してください。');
  }
  if (monitorEv.vitalBad || monitorEv.stoolBad || monitorEv.urineBad) {
    lines.push(
      'バイタル異常、または排便・排尿の観察・記録の空白が長時間続いています。安全管理・記録の十分性（減算・監査）を確認してください。'
    );
  }
  return { hasAlert: lines.length > 0, lines };
}

/** @param {string} linkKey carelinkFacilities の linkKey */
export function getNursingDirectives(linkKey) {
  const all = readJson(LS.nursing, {});
  const list = Array.isArray(all[linkKey]) ? all[linkKey] : [];
  const today = localYmd(new Date());
  return list.filter((d) => {
    const from = String(d.startDate ?? '').trim();
    const to = String(d.endDate ?? '').trim();
    if (from && /^\d{4}-\d{2}-\d{2}$/.test(from) && today < from) return false;
    if (to && /^\d{4}-\d{2}-\d{2}$/.test(to) && today > to) return false;
    return true;
  });
}

export function addNursingDirective(linkKey, text, by = '看護', opts = {}) {
  const t = String(text ?? '').trim();
  if (!t || !linkKey) return false;
  const all = readJson(LS.nursing, {});
  const list = Array.isArray(all[linkKey]) ? all[linkKey] : [];
  const startDate = String(opts?.startDate ?? '').trim();
  const endDate = String(opts?.endDate ?? '').trim();
  list.unshift({
    id: `dir_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    text: t,
    ts: new Date().toISOString(),
    by,
    startDate: /^\d{4}-\d{2}-\d{2}$/.test(startDate) ? startDate : '',
    endDate: /^\d{4}-\d{2}-\d{2}$/.test(endDate) ? endDate : '',
  });
  all[linkKey] = list.slice(0, 30);
  writeJson(LS.nursing, all);
  return true;
}

export function removeNursingDirective(linkKey, directiveId, tsFallback = '') {
  const k = String(linkKey ?? '').trim();
  if (!k) return false;
  const all = readJson(LS.nursing, {});
  const list = Array.isArray(all[k]) ? all[k] : [];
  const next = list.filter(
    (d) =>
      String(d.id ?? '').trim() !== String(directiveId ?? '').trim() &&
      String(d.ts ?? '').trim() !== String(tsFallback ?? '').trim()
  );
  if (next.length === list.length) return false;
  all[k] = next;
  writeJson(LS.nursing, all);
  return true;
}

function localYmd(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * 施設ごとの週間予定（当日〜7日先）を返す
 * @param {string} linkKey
 * @param {Date} [anchor]
 */
export function getWeeklyPlans(linkKey, anchor = new Date()) {
  const k = String(linkKey ?? '').trim();
  if (!k) return [];
  const all = readJson(LS.weeklyPlans, {});
  const list = Array.isArray(all[k]) ? all[k] : [];
  const start = new Date(anchor);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 7);
  return list
    .filter((v) => {
      const d = new Date(String(v.date ?? ''));
      return Number.isFinite(d.getTime()) && d >= start && d < end;
    })
    .sort((a, b) => {
      const ad = `${a.date} ${a.time}`;
      const bd = `${b.date} ${b.time}`;
      return ad.localeCompare(bd);
    });
}

const WEEKDAY_JA_SHORT = Object.freeze(['日', '月', '火', '水', '木', '金', '土']);

/**
 * 当日0時基準の7日間それぞれに予定を割り当て（未登録日は空配列）。外出・受診の持ち物・服薬準備の俯瞰用。
 * @param {string} linkKey
 * @param {Date} [anchor]
 * @returns {{ date: string; weekdayShort: string; isToday: boolean; plans: unknown[] }[]}
 */
export function getWeeklyPlanDays(linkKey, anchor = new Date()) {
  const k = String(linkKey ?? '').trim();
  if (!k) return [];
  const all = readJson(LS.weeklyPlans, {});
  const list = Array.isArray(all[k]) ? all[k] : [];
  const start = new Date(anchor);
  start.setHours(0, 0, 0, 0);
  const endExclusive = new Date(start);
  endExclusive.setDate(endExclusive.getDate() + 7);

  const inWindow = list.filter((v) => {
    const ds = String(v.date ?? '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) return false;
    const d = new Date(`${ds}T12:00:00`);
    return Number.isFinite(d.getTime()) && d >= start && d < endExclusive;
  });

  const byDate = new Map();
  for (const p of inWindow) {
    const dkey = String(p.date ?? '').slice(0, 10);
    if (!byDate.has(dkey)) byDate.set(dkey, []);
    byDate.get(dkey).push(p);
  }
  for (const arr of byDate.values()) {
    arr.sort((a, b) => String(a.time ?? '').localeCompare(String(b.time ?? '')));
  }

  const todayKey = localYmd(new Date());
  const out = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(start);
    d.setDate(d.getDate() + i);
    const key = localYmd(d);
    out.push({
      date: key,
      weekdayShort: WEEKDAY_JA_SHORT[d.getDay()],
      isToday: key === todayKey,
      plans: byDate.get(key) ?? [],
    });
  }
  return out;
}

/**
 * 施設ごとの週間予定を追加
 * @param {string} linkKey
 * @param {{ date: string; time: string; title: string; type?: string }} plan
 */
export function addWeeklyPlan(linkKey, plan) {
  const k = String(linkKey ?? '').trim();
  const title = String(plan?.title ?? '').trim();
  if (!k || !title) return false;
  const dateRaw = String(plan?.date ?? '').trim();
  const date = /^\d{4}-\d{2}-\d{2}$/.test(dateRaw) ? dateRaw : localYmd(new Date());
  const timeRaw = String(plan?.time ?? '').trim();
  const time = /^\d{1,2}:\d{2}$/.test(timeRaw) ? timeRaw : '09:00';
  const type = String(plan?.type ?? 'その他').trim() || 'その他';

  const all = readJson(LS.weeklyPlans, {});
  const list = Array.isArray(all[k]) ? all[k] : [];
  list.push({
    id: `plan_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    date,
    time,
    title,
    type,
    ts: new Date().toISOString(),
  });
  // 古いものは肥大化防止で削る（直近90件）
  all[k] = list.sort((a, b) => `${a.date} ${a.time}`.localeCompare(`${b.date} ${b.time}`)).slice(-90);
  writeJson(LS.weeklyPlans, all);
  return true;
}

export function removeWeeklyPlan(linkKey, planId) {
  const k = String(linkKey ?? '').trim();
  if (!k) return false;
  const all = readJson(LS.weeklyPlans, {});
  const list = Array.isArray(all[k]) ? all[k] : [];
  const next = list.filter((p) => String(p.id ?? '') !== String(planId ?? ''));
  if (next.length === list.length) return false;
  all[k] = next;
  writeJson(LS.weeklyPlans, all);
  return true;
}

export function getAllCareEvents() {
  return readJson(LS.careEvents, []);
}

export function logCareEvent(payload) {
  const list = getAllCareEvents();
  const row = {
    id: `ev_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    ts: new Date().toISOString(),
    ...payload,
  };
  list.push(row);
  const max = 8000;
  writeJson(LS.careEvents, list.slice(-max));
  return row;
}

/** 月次イベント件数（監査集計用） */
export function aggregateMonthlyCareEvents(facilitySheetTitle, yearMonth) {
  const [y, m] = yearMonth.split('-').map(Number);
  const start = new Date(y, m - 1, 1).getTime();
  const end = new Date(y, m, 1).getTime();
  const events = getAllCareEvents().filter((e) => {
    const t = new Date(e.ts).getTime();
    return t >= start && t < end && String(e.facilitySheetTitle ?? '') === String(facilitySheetTitle);
  });
  const c = { patrol: 0, meal: 0, excretion: 0, vital_snapshot: 0, enteral: 0, other: 0 };
  for (const e of events) {
    const t = e.type;
    if (t === 'patrol') c.patrol++;
    else if (t === 'meal') c.meal++;
    else if (t === 'excretion') c.excretion++;
    else if (t === 'vital_snapshot') c.vital_snapshot++;
    else if (t === 'enteral') c.enteral++;
    else c.other++;
  }
  return { ...c, total: events.length, events };
}

/**
 * 利用者・対象月のケアイベント（時系列）
 * @param {string} residentId
 * @param {string} yearMonth YYYY-MM
 */
export function getCareEventsForResidentMonth(residentId, yearMonth) {
  const [y, m] = yearMonth.split('-').map(Number);
  if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) return [];
  const start = new Date(y, m - 1, 1).getTime();
  const end = new Date(y, m, 1).getTime();
  const rid = String(residentId);
  return getAllCareEvents()
    .filter((e) => {
      const t = new Date(e.ts).getTime();
      return String(e.residentId) === rid && t >= start && t < end;
    })
    .sort((a, b) => new Date(a.ts) - new Date(b.ts));
}

/**
 * 請求・月次集計: 利用者×暦月の食事ログ件数・経管実施ログ件数（このブラウザに保存された記録）
 * @param {string} residentId
 * @param {string} yearMonth YYYY-MM
 */
export function summarizeResidentMonthBilling(residentId, yearMonth) {
  const events = getCareEventsForResidentMonth(residentId, yearMonth);
  let mealLogged = 0;
  let enteralLogged = 0;
  for (const e of events) {
    if (e.type === 'meal') mealLogged++;
    if (e.type === 'enteral') enteralLogged++;
  }
  return { mealLogged, enteralLogged };
}

/**
 * 有料・監査向け: 利用者1人×1か月を1行にまとめた CSV（名簿は任意。未指定時は記録のある利用者のみ）
 * @param {string} facilitySheetTitle
 * @param {string} yearMonth YYYY-MM
 * @param {Array<{ id?: unknown; name?: unknown; room?: unknown }>} [roster]
 */
export function buildMonthlyAuditCsv(facilitySheetTitle, yearMonth, roster = []) {
  const q = (v) => {
    const t = String(v ?? '');
    if (/[",\n\r]/.test(t)) return `"${t.replace(/"/g, '""')}"`;
    return t;
  };
  const fmtTs = (ms) => {
    if (ms == null || !Number.isFinite(ms)) return '';
    return new Date(ms).toLocaleString('ja-JP', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false,
    });
  };

  const agg = aggregateMonthlyCareEvents(facilitySheetTitle, yearMonth);
  const { events } = agg;

  /** @type {Map<string, { residentId: string; residentName: string; room: string; patrol: number; meal: number; excretion: number; vital_snapshot: number; enteral: number; other: number; firstTs: number | null; lastTs: number | null; lastPatrolTs: number | null; lastMealTs: number | null; lastExcretionTs: number | null; lastVitalTs: number | null; lastEnteralTs: number | null }>} */
  const byResident = new Map();

  for (const e of events) {
    const rid = String(e.residentId ?? '').trim();
    if (!rid) continue;
    if (!byResident.has(rid)) {
      byResident.set(rid, {
        residentId: rid,
        residentName: String(e.residentName ?? '').trim(),
        room: '',
        patrol: 0,
        meal: 0,
        excretion: 0,
        vital_snapshot: 0,
        enteral: 0,
        other: 0,
        firstTs: null,
        lastTs: null,
        lastPatrolTs: null,
        lastMealTs: null,
        lastExcretionTs: null,
        lastVitalTs: null,
        lastEnteralTs: null,
      });
    }
    const row = byResident.get(rid);
    const rname = String(e.residentName ?? '').trim();
    if (rname) row.residentName = rname;
    const t = new Date(e.ts).getTime();
    if (!Number.isFinite(t)) continue;
    if (row.firstTs == null || t < row.firstTs) row.firstTs = t;
    if (row.lastTs == null || t > row.lastTs) row.lastTs = t;
    const typ = e.type;
    if (typ === 'patrol') {
      row.patrol++;
      if (row.lastPatrolTs == null || t > row.lastPatrolTs) row.lastPatrolTs = t;
    } else if (typ === 'meal') {
      row.meal++;
      if (row.lastMealTs == null || t > row.lastMealTs) row.lastMealTs = t;
    } else if (typ === 'excretion') {
      row.excretion++;
      if (row.lastExcretionTs == null || t > row.lastExcretionTs) row.lastExcretionTs = t;
    } else if (typ === 'vital_snapshot') {
      row.vital_snapshot++;
      if (row.lastVitalTs == null || t > row.lastVitalTs) row.lastVitalTs = t;
    } else if (typ === 'enteral') {
      row.enteral++;
      if (row.lastEnteralTs == null || t > row.lastEnteralTs) row.lastEnteralTs = t;
    } else {
      row.other++;
    }
  }

  const rosterArr = Array.isArray(roster) ? roster : [];
  for (const r of rosterArr) {
    const id = String(r?.id ?? '').trim();
    if (!id) continue;
    const room = String(r?.room ?? '').trim();
    const nm = String(r?.name ?? '').trim();
    if (!byResident.has(id)) {
      byResident.set(id, {
        residentId: id,
        residentName: nm,
        room,
        patrol: 0,
        meal: 0,
        excretion: 0,
        vital_snapshot: 0,
        enteral: 0,
        other: 0,
        firstTs: null,
        lastTs: null,
        lastPatrolTs: null,
        lastMealTs: null,
        lastExcretionTs: null,
        lastVitalTs: null,
        lastEnteralTs: null,
      });
    } else {
      const row = byResident.get(id);
      if (room) row.room = room;
      if (nm && (!row.residentName || row.residentName === '—')) row.residentName = nm;
    }
  }

  const rows = [...byResident.values()].sort((a, b) => {
    const an = a.residentName || a.residentId;
    const bn = b.residentName || b.residentId;
    return an.localeCompare(bn, 'ja');
  });

  const header = [
    '行種別',
    '施設名',
    '対象月',
    '利用者ID',
    '利用者名',
    '居室',
    '巡視回数',
    '食事回数',
    '経管回数',
    '排泄回数',
    'バイタル回数',
    'その他回数',
    '合計回数',
    '初回記録日時',
    '最終記録日時',
    '最終巡視日時',
    '最終食事日時',
    '最終排泄日時',
    '最終バイタル日時',
    '最終経管日時',
    '当月サマリー',
  ];

  const lines = [header.join(',')];

  const facSummary =
    `施設計（当月・本ブラウザ保存分）: 巡視${agg.patrol}・食事${agg.meal}・経管${agg.enteral}・排泄${agg.excretion}・バイタル${agg.vital_snapshot}・その他${agg.other}（総ログ${agg.total}件）／利用者行${rows.length}名`;
  lines.push(
    [
      '施設月次集計',
      q(facilitySheetTitle),
      q(yearMonth),
      '',
      '',
      '',
      agg.patrol,
      agg.meal,
      agg.enteral,
      agg.excretion,
      agg.vital_snapshot,
      agg.other,
      agg.total,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      q(facSummary),
    ].join(',')
  );

  for (const r of rows) {
    const total = r.patrol + r.meal + r.excretion + r.vital_snapshot + r.enteral + r.other;
    const monthStatus =
      total === 0
        ? '当月このブラウザに保存された提供記録なし（未入力・別端末の可能性あり）'
        : `巡視${r.patrol}回・食事${r.meal}回・経管${r.enteral}回・排泄${r.excretion}回・バイタル${r.vital_snapshot}回・その他${r.other}回（合計${total}件）`;
    lines.push(
      [
        '利用者月次',
        q(facilitySheetTitle),
        q(yearMonth),
        q(r.residentId),
        q(r.residentName),
        q(r.room),
        r.patrol,
        r.meal,
        r.enteral,
        r.excretion,
        r.vital_snapshot,
        r.other,
        total,
        q(fmtTs(r.firstTs)),
        q(fmtTs(r.lastTs)),
        q(fmtTs(r.lastPatrolTs)),
        q(fmtTs(r.lastMealTs)),
        q(fmtTs(r.lastExcretionTs)),
        q(fmtTs(r.lastVitalTs)),
        q(fmtTs(r.lastEnteralTs)),
        q(monthStatus),
      ].join(',')
    );
  }

  return lines.join('\n');
}

/**
 * @param {string} facilitySheetTitle
 * @param {string} yearMonth
 * @param {Array<{ id?: unknown; name?: unknown; room?: unknown }>} [roster] 画面上の入居者一覧（記録ゼロも行として出す）
 */
export function downloadMonthlyAuditSheet(facilitySheetTitle, yearMonth, roster = []) {
  const csv = buildMonthlyAuditCsv(facilitySheetTitle, yearMonth, roster);
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const safe = String(facilitySheetTitle).replace(/[\\/:*?"<>|]/g, '_');
  a.download = `有料月次_利用者別_${safe}_${yearMonth}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** 有料監査で想定する巡視間隔（分） */
export const PAID_AUDIT_PATROL_TARGET_MIN = 180;

/**
 * @param {unknown[]} events getCareEventsForResidentMonth の戻り（時系列）
 */
export function analyzePatrolIntervalsForMonth(events) {
  const patrols = (Array.isArray(events) ? events : [])
    .filter((e) => e.type === 'patrol')
    .map((e) => new Date(e.ts).getTime())
    .filter((t) => Number.isFinite(t))
    .sort((a, b) => a - b);
  if (patrols.length === 0) {
    return {
      count: 0,
      maxGapMin: null,
      avgGapMin: null,
      gapsOverTarget: 0,
      narrative:
        '当月の巡視ログなし（紙・別システムの可能性あり）。3時間おきの実施状況は記録と照合してください。',
    };
  }
  const gaps = [];
  for (let i = 1; i < patrols.length; i++) gaps.push((patrols[i] - patrols[i - 1]) / 60000);
  const maxGapMin = gaps.length ? Math.max(...gaps) : null;
  const avgGapMin = gaps.length ? gaps.reduce((a, b) => a + b, 0) / gaps.length : null;
  const gapsOverTarget = gaps.filter((g) => g > PAID_AUDIT_PATROL_TARGET_MIN).length;
  let narrative = `巡視ログ${patrols.length}件。`;
  if (maxGapMin != null) narrative += `記録間の最大空き約${Math.round(maxGapMin)}分`;
  if (avgGapMin != null) narrative += `、平均約${Math.round(avgGapMin)}分。`;
  narrative += `${PAID_AUDIT_PATROL_TARGET_MIN}分超の空きが${gapsOverTarget}回（記録ベース。実巡視と一致するかは現場確認）。`;
  return { count: patrols.length, maxGapMin, avgGapMin, gapsOverTarget, narrative };
}

/** @param {unknown} meta */
function mealValueFromMeta(meta) {
  const v = meta && typeof meta === 'object' ? meta.mealValue : null;
  if (v == null) return null;
  const s = String(v).trim();
  if (!s) return null;
  const n = parseInt(s, 10);
  return Number.isFinite(n) ? n : null;
}

/**
 * @param {unknown[]} events
 */
export function analyzeMealIntakeForMonth(events) {
  const meals = (Array.isArray(events) ? events : []).filter((e) => e.type === 'meal');
  const withVal = [];
  for (const e of meals) {
    const n = mealValueFromMeta(e.meta);
    if (n != null) withVal.push(n);
  }
  const tens = withVal.filter((n) => n >= 10).length;
  const avg = withVal.length ? withVal.reduce((a, b) => a + b, 0) / withVal.length : null;
  let narrative = `食事ログ${meals.length}件。`;
  if (withVal.length === 0) {
    narrative +=
      '摂取割合（◯割）の記録はありません（クイックの「食事確認」等のみの可能性）。10割摂取の評価は別記録と併せて確認してください。';
  } else {
    narrative += `割合記録${withVal.length}件のうち10割相当${tens}件、記録がある分の平均約${avg != null ? avg.toFixed(1) : '—'}割。`;
    if (meals.length > withVal.length) narrative += `（ログ${meals.length - withVal.length}件は割合未記入）`;
  }
  return { mealCount: meals.length, withValueCount: withVal.length, tenCount: tens, avgMealValue: avg, narrative };
}

/**
 * @param {unknown[]} events
 */
export function analyzeExcretionIntervalsForMonth(events) {
  const ex = (Array.isArray(events) ? events : [])
    .filter((e) => e.type === 'excretion')
    .map((e) => ({ t: new Date(e.ts).getTime(), note: String(e.meta?.note ?? '') }))
    .filter((x) => Number.isFinite(x.t))
    .sort((a, b) => a.t - b.t);
  if (ex.length === 0) {
    return {
      count: 0,
      avgGapHours: null,
      maxGapHours: null,
      avgDayGap: null,
      narrative:
        '排泄ログなし。排尿・排便の間隔は別記録（バイタル・排泄表等）と照合してください。※本システムのクイック記録は排尿・排便を区別しない場合があります。',
    };
  }
  const hourGaps = [];
  for (let i = 1; i < ex.length; i++) hourGaps.push((ex[i].t - ex[i - 1].t) / 3600000);
  const avgGapHours = hourGaps.length ? hourGaps.reduce((a, b) => a + b, 0) / hourGaps.length : null;
  const maxGapHours = hourGaps.length ? Math.max(...hourGaps) : null;

  const byDay = [...new Set(ex.map((x) => localYmd(new Date(x.t))))].sort();
  const dayGaps = [];
  for (let i = 1; i < byDay.length; i++) {
    const a = new Date(`${byDay[i - 1]}T12:00:00`);
    const b = new Date(`${byDay[i]}T12:00:00`);
    dayGaps.push((b - a) / 86400000);
  }
  const avgDayGap = dayGaps.length ? dayGaps.reduce((x, y) => x + y, 0) / dayGaps.length : null;

  let narrative = `排泄ログ${ex.length}件。記録間の平均約${avgGapHours != null ? avgGapHours.toFixed(1) : '—'}時間、最大約${maxGapHours != null ? maxGapHours.toFixed(1) : '—'}時間。`;
  narrative += `記録のあった日の間隔（目安）平均約${avgDayGap != null ? avgDayGap.toFixed(1) : '—'}日。`;
  narrative += '（排尿・排便の別・実際の排泄間隔は記録様式により異なります。）';
  return { count: ex.length, avgGapHours, maxGapHours, avgDayGap, narrative };
}

function escPaidHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * 監査HTML用: 同一利用者が複数施設で記録している場合、名簿の施設に合わせて絞り込む
 * @param {unknown[]} events
 * @param {string} facilitySheetTitle
 */
function filterEventsByFacilitySheet(events, facilitySheetTitle) {
  const fac = String(facilitySheetTitle ?? '').trim();
  if (!fac) return Array.isArray(events) ? events : [];
  return (Array.isArray(events) ? events : []).filter((e) => {
    const ef = e.facilitySheetTitle != null ? String(e.facilitySheetTitle).trim() : '';
    return !ef || ef === fac;
  });
}

/** @param {unknown} e */
function excretionContributesToUrine(e) {
  if (e.type !== 'excretion') return false;
  const m = e.meta && typeof e.meta === 'object' ? e.meta : {};
  const u = String(m.urineVolume ?? '').trim();
  const sv = String(m.stoolVolume ?? '').trim();
  const sc = String(m.stoolCharacter ?? '').trim();
  const hasDetail = Boolean(u || sv || sc);
  if (!hasDetail) return true;
  return Boolean(u);
}

/** @param {unknown} e */
function excretionContributesToStool(e) {
  if (e.type !== 'excretion') return false;
  const m = e.meta && typeof e.meta === 'object' ? e.meta : {};
  const u = String(m.urineVolume ?? '').trim();
  const sv = String(m.stoolVolume ?? '').trim();
  const sc = String(m.stoolCharacter ?? '').trim();
  const hasDetail = Boolean(u || sv || sc);
  if (!hasDetail) return true;
  return Boolean(sv || sc);
}

/**
 * 当月・ローカル時刻の「時」（0–23）ごとに件数を数える
 * @param {unknown[]} events
 * @param {string} type patrol | meal | excretion | enteral
 * @param {(e: unknown) => boolean} [predicate]
 */
function hourBucketsForMonthEvents(events, type, predicate) {
  const buckets = Array.from({ length: 24 }, () => 0);
  for (const e of Array.isArray(events) ? events : []) {
    if (e.type !== type) continue;
    if (predicate && !predicate(e)) continue;
    const t = new Date(e.ts).getTime();
    if (!Number.isFinite(t)) continue;
    const h = new Date(e.ts).getHours();
    if (h >= 0 && h < 24) buckets[h]++;
  }
  return buckets;
}

/**
 * バイタルチェック表風: 時刻帯（00–23）× 区分のグリッドHTML（印刷向け）
 * @param {string} facilitySheetTitle
 * @param {string} yearMonth
 * @param {Array<Record<string, unknown>>} roster
 */
function buildPaidAuditHourlySheetHtmlSection(facilitySheetTitle, yearMonth, roster = []) {
  const fac = String(facilitySheetTitle ?? '');
  const ym = String(yearMonth ?? '');
  const def = facilityDefBySheetTitle(fac);
  const displayName = def?.tabLabel ? String(def.tabLabel) : fac;
  const rosterArr = Array.isArray(roster) ? roster : [];
  const byId = new Map();
  for (const r of rosterArr) {
    const id = String(r?.id ?? '').trim();
    if (!id) continue;
    byId.set(id, r);
  }
  const rows = [...byId.values()].sort((a, b) =>
    String(a?.name ?? '').localeCompare(String(b?.name ?? ''), 'ja')
  );

  const hourLabels = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0'));

  const parts = [];
  parts.push(
    `<section class="hourly-sheet" style="page-break-before:always;margin-top:24px;">`
  );
  parts.push(`<h2 style="font-size:1.1rem;margin:0 0 8px;color:#0e7490;">バイタルチェック表形式（ログ発生時刻の時別・当月）</h2>`);
  parts.push(
    `<p style="font-size:0.86rem;color:#475569;margin:0 0 12px;">紙の様式に近い<strong>00〜23時のマス</strong>です。当月に記録されたイベントを<strong>発生時刻の時</strong>（この端末のタイムゾーン）に振り分けています。尿・便は記録内容（尿量・便量・性状）で振り分け、詳細のない「排泄確認」は尿・便の両方に反映します。</p>`
  );
  parts.push(`<div class="sheet-caption" style="font-size:0.95rem;font-weight:700;margin-bottom:10px;">${escPaidHtml(displayName)}　バイタルチェック表（${escPaidHtml(ym)}）</div>`);

  if (rows.length === 0) {
    parts.push(`<p style="font-size:0.88rem;">利用者がありません。</p></section>`);
    return parts.join('');
  }

  for (const r of rows) {
    const id = String(r?.id ?? '');
    const name = String(r?.name ?? '—');
    const room = String(r?.room ?? '—');
    const raw = getCareEventsForResidentMonth(id, ym);
    const events = filterEventsByFacilitySheet(raw, fac);

    const bPatrol = hourBucketsForMonthEvents(events, 'patrol');
    const bMeal = hourBucketsForMonthEvents(events, 'meal');
    const bEnteral = hourBucketsForMonthEvents(events, 'enteral');
    const bUrine = hourBucketsForMonthEvents(events, 'excretion', excretionContributesToUrine);
    const bStool = hourBucketsForMonthEvents(events, 'excretion', excretionContributesToStool);

    const rowSpecs = [
      { label: '巡視', buckets: bPatrol },
      { label: '食事', buckets: bMeal },
      { label: '経管', buckets: bEnteral },
      { label: '尿', buckets: bUrine },
      { label: '便', buckets: bStool },
    ];

    parts.push(`<div class="resident-hourly-block" style="break-inside:avoid;margin-bottom:20px;">`);
    parts.push(
      `<div style="font-size:0.9rem;font-weight:600;margin-bottom:6px;">${escPaidHtml(name)}　<span style="font-weight:400;color:#64748b;">居室 ${escPaidHtml(room)}</span></div>`
    );
    parts.push(`<div class="sheet-scroll" style="overflow-x:auto;-webkit-overflow-scrolling:touch;">`);
    parts.push(`<table class="hourly-grid" style="border-collapse:collapse;font-size:9px;width:100%;min-width:720px;">`);
    parts.push(`<thead><tr><th style="border:1px solid #334155;background:#f1f5f9;padding:4px 6px;min-width:3em;">区分</th>`);
    for (const h of hourLabels) {
      parts.push(
        `<th style="border:1px solid #334155;background:#f1f5f9;padding:2px 3px;width:2.2em;">${h}</th>`
      );
    }
    parts.push(`</tr></thead><tbody>`);
    for (const spec of rowSpecs) {
      parts.push(`<tr><th scope="row" style="border:1px solid #334155;background:#f8fafc;padding:4px 6px;text-align:left;white-space:nowrap;">${escPaidHtml(spec.label)}</th>`);
      for (let hi = 0; hi < 24; hi++) {
        const n = spec.buckets[hi];
        const cell = n > 0 ? (n > 1 ? String(n) : '●') : '';
        parts.push(
          `<td style="border:1px solid #94a3b8;padding:2px;text-align:center;min-height:1.4em;">${escPaidHtml(cell)}</td>`
        );
      }
      parts.push(`</tr>`);
    }
    parts.push(`</tbody></table></div></div>`);
  }

  parts.push(`</section>`);
  return parts.join('');
}

/**
 * 有料監査・請求説明用: 間隔・摂取状況の文章と請求用件数を利用者ごとにHTML化（印刷・提出のたたき台）
 * @param {string} facilitySheetTitle
 * @param {string} yearMonth YYYY-MM
 * @param {Array<Record<string, unknown>>} [roster] id, name, room, mealCountThisMonth, isEnteral
 */
export function buildPaidAuditMonthlyNarrativeHtml(facilitySheetTitle, yearMonth, roster = []) {
  const fac = String(facilitySheetTitle ?? '');
  const ym = String(yearMonth ?? '');
  const rosterArr = Array.isArray(roster) ? roster : [];
  const byId = new Map();
  for (const r of rosterArr) {
    const id = String(r?.id ?? '').trim();
    if (!id) continue;
    byId.set(id, r);
  }
  const rows = [...byId.values()].sort((a, b) =>
    String(a?.name ?? '').localeCompare(String(b?.name ?? ''), 'ja')
  );

  const parts = [];
  parts.push(`<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"/>`);
  parts.push(`<title>${escPaidHtml(fac)} ${escPaidHtml(ym)} 有料監査・請求用サマリー</title>`);
  parts.push(`<style>
    body{font-family:system-ui,-apple-system,sans-serif;margin:16px;line-height:1.55;color:#0f172a;max-width:1100px;}
    h1{font-size:1.2rem;margin:0 0 12px;}
    .warn{background:#fff7ed;border:1px solid #fdba74;padding:10px 12px;border-radius:8px;font-size:0.88rem;margin-bottom:16px;}
    .card{border:1px solid #cbd5e1;border-radius:10px;padding:12px 14px;margin-bottom:14px;break-inside:avoid;}
    h2{font-size:1.05rem;margin:0 0 8px;color:#0e7490;}
    .bill{font-size:0.88rem;color:#334155;margin-bottom:8px;}
    .one{font-weight:700;background:#f0fdfa;border-left:4px solid #0d9488;padding:8px 10px;margin:8px 0;font-size:0.95rem;}
    .detail{font-size:0.86rem;color:#475569;margin-top:6px;}
    @media print{.warn{break-inside:avoid;}.hourly-sheet .sheet-scroll{overflow:visible;}}
  </style></head><body>`);
  parts.push(`<h1>${escPaidHtml(fac)}／${escPaidHtml(ym)}　有料サービス監査・請求用サマリー（たたき台）</h1>`);
  parts.push(
    `<div class="warn"><strong>※重要</strong>　本書は<strong>この端末に保存されたログ</strong>から自動生成した草案です。<strong>CSV（件数・最終日時）</strong>と併用できます。巡視・排泄は「記録があった間隔」であり、実サービス実態・排尿排便の区別は原本記録と照合して追記・修正してください。提出用の<strong>一言</strong>は事業所の文面に合わせて調整してください。</div>`
  );

  for (const r of rows) {
    const id = String(r?.id ?? '');
    const name = String(r?.name ?? '—');
    const room = String(r?.room ?? '—');
    const events = filterEventsByFacilitySheet(getCareEventsForResidentMonth(id, ym), fac);
    const patrol = analyzePatrolIntervalsForMonth(events);
    const mealA = analyzeMealIntakeForMonth(events);
    const exc = analyzeExcretionIntervalsForMonth(events);
    const bill = summarizeResidentMonthBilling(id, ym);
    const sheetMeal = Number(r?.mealCountThisMonth) || 0;
    const mealTotal = sheetMeal + bill.mealLogged;
    const enteralFlag = Boolean(r?.isEnteral);

    const oneLine = [
      `【${name.replace(/様\s*$/u, '').trim()}様】`,
      `巡視は記録上${patrol.count}件（${PAID_AUDIT_PATROL_TARGET_MIN}分超の空き${patrol.gapsOverTarget}回）。`,
      `食事は名簿${sheetMeal}回＋ログ${bill.mealLogged}回＝合計${mealTotal}回、経管実施ログ${bill.enteralLogged}回。`,
      mealA.withValueCount
        ? `摂取割合のある食事記録は${mealA.withValueCount}件（10割相当${mealA.tenCount}件）。`
        : '食事は割合未記録のログ中心のため、10割摂取は別記録で確認。',
      exc.count
        ? `排泄ログは${exc.count}件（記録間隔の目安は本文参照）。`
        : '排泄ログなし（別記録要確認）。',
    ].join('');

    parts.push(`<div class="card">`);
    parts.push(`<h2>${escPaidHtml(name)}（居室 ${escPaidHtml(room)}）</h2>`);
    parts.push(
      `<div class="bill">【請求・集計】当月食事回数 名簿<strong>${sheetMeal}</strong>＋端末記録<strong>${bill.mealLogged}</strong>＝合計<strong>${mealTotal}</strong>回／経管栄養実施ログ <strong>${bill.enteralLogged}</strong>回／名簿の経管対象 <strong>${enteralFlag ? 'あり' : 'なし'}</strong></div>`
    );
    parts.push(`<div class="one">提出用・一言要約（調整可）<br/>${escPaidHtml(oneLine)}</div>`);
    parts.push(`<div class="detail"><strong>巡視（3時間おきの観点・記録ベース）</strong><br/>${escPaidHtml(patrol.narrative)}</div>`);
    parts.push(`<div class="detail"><strong>食事（全量・割合の記録ベース）</strong><br/>${escPaidHtml(mealA.narrative)}</div>`);
    parts.push(`<div class="detail"><strong>排泄（排尿・排便を分けないログの場合の間隔目安）</strong><br/>${escPaidHtml(exc.narrative)}</div>`);
    parts.push(`</div>`);
  }

  if (rows.length === 0) {
    parts.push(`<p>名簿が渡されていないか、利用者0件です。記録画面で施設を開いたうえで出力してください。</p>`);
  }

  parts.push(buildPaidAuditHourlySheetHtmlSection(facilitySheetTitle, yearMonth, roster));
  parts.push(`</body></html>`);
  return parts.join('');
}

/**
 * @param {string} facilitySheetTitle
 * @param {string} yearMonth
 * @param {Array<Record<string, unknown>>} [roster]
 */
export function downloadPaidAuditNarrativeHtml(facilitySheetTitle, yearMonth, roster = []) {
  const html = buildPaidAuditMonthlyNarrativeHtml(facilitySheetTitle, yearMonth, roster);
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const safe = String(facilitySheetTitle).replace(/[\\/:*?"<>|]/g, '_');
  a.download = `有料監査_月次要約_${safe}_${yearMonth}.html`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** @param {string} residentId */
export function getEmergencyContact(residentId) {
  const all = readJson(LS.emergencyContact, {});
  return all[String(residentId)] ?? { name: '（未登録）', tel: '—', relation: '—' };
}

export function setEmergencyContact(residentId, data) {
  const all = readJson(LS.emergencyContact, {});
  all[String(residentId)] = { ...data };
  writeJson(LS.emergencyContact, all);
}

/**
 * 直近7日のバイタルログ（スナップショット履歴は簡易: careEvents type vital から）
 * @param {string} residentId
 */
export function getWeeklyVitalTimeline(residentId) {
  const now = Date.now();
  const weekAgo = now - 7 * 24 * 3600000;
  return getAllCareEvents()
    .filter(
      (e) =>
        String(e.residentId) === String(residentId) &&
        e.type === 'vital_snapshot' &&
        new Date(e.ts).getTime() >= weekAgo
    )
    .sort((a, b) => new Date(a.ts) - new Date(b.ts));
}

export function logVitalSnapshot(residentId, residentName, facilitySheetTitle, vitals) {
  return logCareEvent({
    type: 'vital_snapshot',
    residentId,
    residentName,
    facilitySheetTitle,
    meta: vitals,
  });
}

/**
 * カレンダー用: 過去7日のイベントを日付キーで集約
 * @param {string} residentId
 */
export function getWeekCalendarBuckets(residentId, anchor = new Date()) {
  const buckets = {};
  for (let d = 0; d < 7; d++) {
    const dt = new Date(anchor);
    dt.setHours(0, 0, 0, 0);
    dt.setDate(dt.getDate() - d);
    const key = localYmd(dt);
    buckets[key] = { date: key, patrol: 0, meal: 0, excretion: 0, enteral: 0, notes: [] };
  }
  const start = new Date(anchor);
  start.setDate(start.getDate() - 6);
  start.setHours(0, 0, 0, 0);
  for (const e of getAllCareEvents()) {
    if (String(e.residentId) !== String(residentId)) continue;
    const ed = new Date(e.ts);
    if (ed.getTime() < start.getTime()) continue;
    const key = localYmd(ed);
    if (!buckets[key]) continue;
    if (e.type === 'patrol') buckets[key].patrol++;
    if (e.type === 'meal') buckets[key].meal++;
    if (e.type === 'excretion') buckets[key].excretion++;
    if (e.type === 'enteral') buckets[key].enteral++;
    if (e.meta?.note) buckets[key].notes.push(String(e.meta.note));
  }
  return Object.values(buckets).sort((a, b) => a.date.localeCompare(b.date));
}

const CARE_EVENT_TYPE_JA = Object.freeze({
  patrol: '巡視',
  meal: '食事',
  excretion: '排泄',
  vital_snapshot: 'バイタル',
  enteral: '経管栄養',
  fluid_intake: '水分',
});

/** @param {string} iso */
function careEventShortTs(iso) {
  try {
    return new Date(iso).toLocaleString('ja-JP', {
      month: 'numeric',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  } catch {
    return String(iso ?? '');
  }
}

/** @param {Record<string, unknown>} m */
function formatVitalMetaLine(m) {
  const parts = [];
  if (m.temp != null && String(m.temp).trim() !== '') parts.push(`体温 ${String(m.temp).trim()}℃`);
  if (m.bpUpper != null && String(m.bpUpper).trim() !== '')
    parts.push(
      `血圧 ${String(m.bpUpper).trim()}/${String(m.bpLower ?? '').trim() || '—'}`
    );
  if (m.pulse != null && String(m.pulse).trim() !== '') parts.push(`脈拍 ${String(m.pulse).trim()}`);
  if (m.spo2 != null && String(m.spo2).trim() !== '') parts.push(`SpO2 ${String(m.spo2).trim()}%`);
  if (m.weight != null && String(m.weight).trim() !== '') parts.push(`体重 ${String(m.weight).trim()}kg`);
  return parts.length ? parts.join('、') : '（数値なし）';
}

/** @param {{ type?: string; ts?: string; meta?: Record<string, unknown> }} e */
function formatCareEventOneLine(e) {
  const tj = CARE_EVENT_TYPE_JA[e.type] ?? String(e.type ?? '記録');
  const m = e.meta ?? {};
  if (e.type === 'vital_snapshot') return `${careEventShortTs(e.ts)} ${tj}: ${formatVitalMetaLine(m)}`;
  if (e.type === 'excretion') {
    const u = String(m.urineVolume ?? '').trim();
    const sv = String(m.stoolVolume ?? '').trim();
    const sc = String(m.stoolCharacter ?? '').trim();
    const tg = Boolean(m.toiletGuidance);
    if (u || sv || sc || tg) {
      const parts = [];
      if (u) parts.push(`尿量 ${u}`);
      if (sv) parts.push(`便量 ${sv}`);
      if (sc) parts.push(`性状 ${sc}`);
      if (tg) parts.push('トイレ誘導');
      return `${careEventShortTs(e.ts)} ${tj}: ${parts.join(' ')}`;
    }
  }
  if (e.type === 'fluid_intake') {
    const wm = String(m.waterMl ?? '').trim();
    if (wm) return `${careEventShortTs(e.ts)} ${tj}: ${wm}ml`;
  }
  if (e.type === 'meal') {
    const slot = String(m.mealSlot ?? '').trim();
    const amt = String(m.mealAmount ?? '').trim();
    const wm = String(m.waterMl ?? '').trim();
    const med = m.medicationTaken;
    if (slot || amt || wm || med === 'yes' || med === 'no') {
      const parts = [];
      if (slot) parts.push(slot);
      if (amt) parts.push(`量 ${amt}`);
      if (wm) parts.push(`水分 ${wm}ml`);
      if (med === 'yes') parts.push('内服 飲了');
      else if (med === 'no') parts.push('内服 未服');
      return `${careEventShortTs(e.ts)} ${tj}: ${parts.join(' ')}`;
    }
  }
  if (m.note) return `${careEventShortTs(e.ts)} ${tj}: ${String(m.note)}`;
  if (m.mealValue != null && String(m.mealValue).trim() !== '') {
    const mt = String(m.mealTime ?? '').trim();
    return `${careEventShortTs(e.ts)} ${tj}: ${[mt, `${String(m.mealValue).trim()}割`].filter(Boolean).join(' ')}`;
  }
  if (m.stool != null && String(m.stool).trim() !== '')
    return `${careEventShortTs(e.ts)} ${tj}: ${String(m.stool)}`;
  return `${careEventShortTs(e.ts)} ${tj}`;
}

/**
 * 救急搬送サマリー下段4欄を、localStorage のケアイベント・バイタル・名簿から組み立てる
 * @param {Record<string, unknown>} resident
 * @param {string} [facilitySheetTitle] 突合参考（現状は利用者ID中心で抽出）
 * @param {string} [linkKey] 施設の看護指示取得用（carelinkFacilities.linkKey）
 * @returns {{ dailyLife: string; nurseProblems: string; nurseContent: string; careNotes: string }}
 */
export function buildEmergencySummaryNarrativeFromRecords(resident, facilitySheetTitle = '', linkKey = '') {
  const id = String(resident?.id ?? '');
  const cond = String(resident?.condition ?? '').trim() || '—';
  const lastStoolCell = String(resident?.lastStoolDate ?? '').trim() || '—';
  const contact = id ? getEmergencyContact(id) : { name: '（未登録）', tel: '—', relation: '—' };
  const facHint = String(facilitySheetTitle ?? '').trim();

  const now = Date.now();
  const weekAgo = now - 7 * 24 * 3600000;

  const evalResult = id ? evaluateResidentMonitor(resident) : null;
  const buckets = id ? getWeekCalendarBuckets(id) : [];

  const lifeEvents = id
    ? getAllCareEvents().filter(
        (e) =>
          String(e.residentId) === id &&
          ['patrol', 'meal', 'excretion', 'enteral', 'fluid_intake'].includes(String(e.type)) &&
          new Date(e.ts).getTime() >= weekAgo
      )
    : [];
  lifeEvents.sort((a, b) => new Date(b.ts) - new Date(a.ts));

  const dailyLines = [];
  dailyLines.push(`【主疾患・状態（名簿）】${cond}`);
  dailyLines.push(`【名簿の排便欄】${lastStoolCell}`);
  if (facHint) dailyLines.push(`【参照タブ】${facHint}`);

  const bucketLines = [];
  for (const b of buckets) {
    const sum = b.patrol + b.meal + b.excretion;
    const noteStr = Array.isArray(b.notes) && b.notes.length ? ` メモ: ${[...new Set(b.notes)].join(' / ')}` : '';
    if (sum > 0 || noteStr)
      bucketLines.push(`${b.date} 巡視${b.patrol}・食事${b.meal}・排泄${b.excretion}${noteStr}`);
  }
  if (bucketLines.length) {
    dailyLines.push('【直近7日・提供記録件数（この端末に保存されたログ）】');
    dailyLines.push(...bucketLines);
  } else {
    dailyLines.push(
      '【直近7日・提供記録】この端末に保存された巡視・食事・排泄の件数ログはまだありません。'
    );
  }

  if (lifeEvents.length) {
    dailyLines.push('【直近の巡視・食事・排泄ログ（最大12件・新しい順）】');
    for (const e of lifeEvents.slice(0, 12)) dailyLines.push(`・${formatCareEventOneLine(e)}`);
  }

  const problemLines = [];
  if (evalResult) {
    if (evalResult.vitalBad && evalResult.vitalFlags.length) {
      problemLines.push('【バイタル自動検知】');
      for (const f of evalResult.vitalFlags) problemLines.push(`・${f.label}`);
    }
    if (evalResult.stoolBad) {
      const h = evalResult.stoolHours;
      problemLines.push(
        `【排便】最終排便から約 ${h != null ? Math.round(h) : '?'} 時間（${VITAL_THRESHOLDS.stoolHoursMax}時間超でフラグ）`
      );
    }
    if (evalResult.urineBad) {
      const uh = evalResult.urineHours;
      problemLines.push(
        `【排尿】最終排尿記録・トイレ誘導から約 ${uh != null ? Math.round(uh) : '?'} 時間（${VITAL_THRESHOLDS.urineHoursMax}時間超でフラグ）`
      );
    }
    if (evalResult.patrolBad) {
      const pm = Number(resident?.patrolIntervalMinutes);
      problemLines.push(`【巡視間隔】名簿ベース約 ${Number.isFinite(pm) ? pm : '—'} 分（要確認の目安）`);
    }
  }
  if (!problemLines.length)
    problemLines.push(
      '（直近のバイタル入力・排便・排尿の記録間隔から、システム上の明確な異常フラグはありません。臨床判断は担当者が行ってください。）'
    );

  const snap = id ? getResidentVitalSnapshot(id) : null;
  const vitalLog = id ? getWeeklyVitalTimeline(id).slice(-5) : [];
  const contentLines = [];
  contentLines.push('【緊急連絡先】');
  contentLines.push(
    `${String(contact.name ?? '')}（${String(contact.relation ?? '')}）${String(contact.tel ?? '')}`
  );
  contentLines.push('');
  contentLines.push('【現在のバイタル（最新入力値）】');
  if (snap && (snap.temp || snap.bpUpper || snap.pulse || snap.spo2 || snap.weight)) {
    contentLines.push(formatVitalMetaLine(/** @type {Record<string, unknown>} */ (snap)));
    if (snap.updatedAt) contentLines.push(`（更新: ${careEventShortTs(snap.updatedAt)}）`);
  } else {
    contentLines.push('（未入力）');
  }
  if (vitalLog.length) {
    contentLines.push('');
    contentLines.push('【直近のバイタル記録ログ（最大5件）】');
    for (const e of vitalLog) contentLines.push(`・${formatCareEventOneLine(e)}`);
  }

  const careLines = [];
  if (evalResult) {
    const adv = fallbackRegulatoryAdvice(evalResult);
    if (adv && !/^【システム】/.test(adv)) {
      careLines.push('【記録・連携上の配慮（自動検知に基づく参考）】');
      careLines.push(adv);
    }
  }
  const nDir = String(linkKey ?? '').trim() ? getNursingDirectives(String(linkKey)) : [];
  const recentN = Array.isArray(nDir) ? nDir.slice(0, 5) : [];
  if (recentN.length) {
    if (careLines.length) careLines.push('');
    careLines.push('【施設の看護指示メモ（直近・参考）】');
    for (const row of recentN) {
      const tx = String(row?.text ?? '').trim();
      if (tx) careLines.push(`・${tx}`);
    }
  }
  if (!careLines.length)
    careLines.push(
      '（看護指示メモの登録がなく、自動検知に基づく特記もありません。個別の注意事項があれば追記してください。）'
    );

  return {
    dailyLife: dailyLines.join('\n').trim(),
    nurseProblems: problemLines.join('\n').trim(),
    nurseContent: contentLines.join('\n').trim(),
    careNotes: careLines.join('\n').trim(),
  };
}

/**
 * 月次家族向け報告用AIプロンプトに埋め込むテキスト（同一ブラウザの記録）
 * @param {Record<string, unknown>} resident
 * @param {string} yearMonth YYYY-MM
 */
export function buildMonthlyResidentReportContextForAi(resident, yearMonth) {
  const id = String(resident?.id ?? '');
  const ym = String(yearMonth ?? '').trim();
  const events = id && ym ? getCareEventsForResidentMonth(id, ym) : [];
  const c = { patrol: 0, meal: 0, excretion: 0, vital_snapshot: 0, other: 0 };
  for (const e of events) {
    const t = e.type;
    if (t === 'patrol') c.patrol++;
    else if (t === 'meal') c.meal++;
    else if (t === 'excretion') c.excretion++;
    else if (t === 'vital_snapshot') c.vital_snapshot++;
    else c.other++;
  }
  const snap = id ? getResidentVitalSnapshot(id) : null;
  const evalR = id ? evaluateResidentMonitor(resident) : null;
  const contact = id ? getEmergencyContact(id) : null;
  const lk = nursingLinkKeyForResident(resident);
  const nDir = lk ? getNursingDirectives(lk).slice(0, 10) : [];
  const nursingLines = nDir.map((row) => String(row?.text ?? '').trim()).filter(Boolean);

  const excerpt = events
    .slice(-45)
    .map((e) => formatCareEventOneLine(/** @type {{ type?: string; ts?: string; meta?: Record<string, unknown> }} */ (e)))
    .join('\n');

  const lines = [];
  lines.push(`【対象月】${ym}`);
  lines.push(
    `【利用者】氏名: ${String(resident?.name ?? '')} / 居室: ${String(resident?.room ?? '')} / 主疾患・状態(名簿): ${String(resident?.condition ?? '—')}`
  );
  lines.push(
    `【施設・出所】facility列: ${String(resident?.facility ?? '—')} / 読込タブ: ${String(resident?.sourceSheetTitle ?? '—')}`
  );
  lines.push(
    `【当月記録件数（この端末）】巡視 ${c.patrol} / 食事 ${c.meal} / 排泄 ${c.excretion} / バイタル ${c.vital_snapshot} / その他 ${c.other} / 合計 ${events.length}`
  );
  if (snap && (snap.temp || snap.bpUpper || snap.pulse)) {
    lines.push(`【最新バイタル（入力済み）】${formatVitalMetaLine(/** @type {Record<string, unknown>} */ (snap))}`);
  }
  if (evalR) {
    lines.push(
      `【自動検知参考】バイタル注意: ${evalR.vitalBad ? evalR.vitalFlags.map((f) => f.label).join('、') : 'なし'} / 排便遅延: ${evalR.stoolBad ? `約${evalR.stoolHours != null ? Math.round(evalR.stoolHours) : '?'}h` : 'なし'} / 排尿間隔: ${evalR.urineBad ? `約${evalR.urineHours != null ? Math.round(evalR.urineHours) : '?'}h` : 'なし'}`
    );
  }
  lines.push(`【名簿の排便欄】${String(resident?.lastStoolDate ?? '—')}`);
  if (contact) {
    lines.push(
      `【緊急連絡先（登録値）】${String(contact.name ?? '')}（${String(contact.relation ?? '')}）${String(contact.tel ?? '')}`
    );
  }
  if (nursingLines.length) {
    lines.push('【施設の看護指示メモ（抜粋）】');
    nursingLines.forEach((t) => lines.push(`・${t}`));
  }
  lines.push('【当月のケアログ抜粋（時系列・最大45件）】');
  lines.push(excerpt || '（ログなし。クイック記録等が未登録の月です。）');
  return lines.join('\n');
}

function stripJsonFence(text) {
  const t = text.trim();
  const m = /^```(?:json)?\s*\n?([\s\S]*?)\n?```$/im.exec(t);
  return m ? m[1].trim() : t;
}

/**
 * 1か月の記録を参照し、家族向け月次報告の3文案を Gemini で生成
 * @param {string} apiKey
 * @param {Record<string, unknown>} resident
 * @param {string} yearMonth YYYY-MM
 * @returns {Promise<{ monthlyCondition: string; futureCarePoints: string; directorMessage: string }>}
 */
export async function fetchMonthlyResidentFamilyReportAi(apiKey, resident, yearMonth) {
  if (!apiKey?.trim()) throw new Error('VITE_GEMINI_API_KEY が未設定です。');
  const ym = String(yearMonth ?? '').trim();
  if (!/^\d{4}-\d{2}$/.test(ym)) throw new Error('対象月が不正です。');
  const ctx = buildMonthlyResidentReportContextForAi(resident, ym);
  const prompt = `あなたは有料老人ホームの施設長・ケアマネの補佐です。次の「根拠データ」は同一ブラウザに保存された提供記録ログと名簿情報です。データに書かれていないことは推測で補わず、不足時は「記録上は確認できません」と明記してください。推測が必要な場合は「〜の可能性」として断定しないでください。

【根拠データ】
${ctx}

【出力】
次のキーを持つJSONオブジェクト1つだけを返してください（説明文・Markdownのフェンス禁止）。各値は日本語の敬体（です・ます調）の文章です。
- monthlyCondition: 「1か月の状態と様子」（家族向け。200～600字程度。記録に基づく観察・バイタル傾向・食事排泄巡視の様子が分かるように）
- futureCarePoints: 「今後気を付けていくこと」（100～400字程度。安全・健康・連携の観点）
- directorMessage: 「施設長から一言」（短めでもよいが50～200字程度。温かみのある一言）

記録件数が0に近い月は、「記録蓄積が少ないため一般論のみ」であることを冒頭で述べてください。`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.35, maxOutputTokens: 4096 },
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!res.ok || !text) {
    const msg = String(data?.error?.message ?? '');
    if (isGeminiQuotaMessage(msg)) {
      const out = {};
      for (const k of ACCIDENT_VOICE_DRAFT_KEYS) out[k] = '';
      out.reportYear2 = String(ty).padStart(2, '0');
      out.reportMonth = tm;
      out.reportDay = td;
      out.occurYear2 = String(ty).padStart(2, '0');
      out.reporterDept = deptPreset || '';
      out.residentName = resName || '';
      out.occurPlace = room || '';
      out.situation = memo.slice(0, 400);
      out.response = '安全確保・観察・報告を実施。';
      out.causes = '状況要因を確認中。';
      out.improvements = '再発防止のため手順と環境を見直します。';
      return out;
    }
    throw new Error(msg || 'AI応答なし');
  }
  const raw = stripJsonFence(text);
  try {
    const parsed = JSON.parse(raw);
    return {
      monthlyCondition: String(parsed.monthlyCondition ?? '').trim(),
      futureCarePoints: String(parsed.futureCarePoints ?? '').trim(),
      directorMessage: String(parsed.directorMessage ?? '').trim(),
    };
  } catch {
    throw new Error('AIのJSONを解釈できませんでした。もう一度お試しください。');
  }
}

/** ルールベース即時アドバイス */
export function fallbackRegulatoryAdvice(evalResult) {
  const lines = [];
  if (evalResult.vitalBad) {
    for (const f of evalResult.vitalFlags) {
      if (f.code === 'fever')
        lines.push('【実務】発熱傾向: 再測定・水分・主治医／看護報告を検討（感染・脱水の観察記録を残す）。');
      if (f.code === 'bp_sys_high')
        lines.push('【実務】収縮期血圧高値: 安静後再測、服薬確認、閾値超過時は医療報告（報酬上も安全配慮の記録が重要）。');
      if (f.code === 'bp_dia_low')
        lines.push('【実務】拡張期血圧低値: めまい・失神の有無、脱水・服薬の確認、必要に応じ受診連絡。');
    }
  }
  if (evalResult.stoolBad) {
    lines.push(
      '【実務】排便遅延（72h超）: 腹部症状の観察、医師・看護へ報告。下剤は指示に基づき実施し結果を記録（褥瘡・腸閉塞リスク）。'
    );
  }
  if (evalResult.urineBad) {
    lines.push(
      `【実務】排尿の記録・トイレ誘導の間隔が空きすぎています（目安${VITAL_THRESHOLDS.urineHoursMax}時間超）。尿閉・尿路感染・転倒リスクの観察、必要に応じて医療・看護へ相談してください。`
    );
  }
  if (!lines.length) return '【システム】該当する自動アラートに紐づく定型アドバイスはありません。';
  return lines.join('\n');
}

/**
 * @param {string} apiKey
 * @param {ReturnType<typeof evaluateResidentMonitor>} evalResult
 * @param {Record<string, unknown>} resident
 */
export async function fetchAiRegulatoryAdvice(apiKey, evalResult, resident) {
  if (!apiKey?.trim()) return fallbackRegulatoryAdvice(evalResult);
  const prompt = `あなたは介護・看護の監査アドバイザーです。次の知識のみを根拠に、簡潔に日本語で答えてください（3〜6行）。

【参照知識】
${REGULATORY_KNOWLEDGE_BASE}

【ケース】
利用者: ${resident.name} / 居室 ${resident.room}
異常: ${JSON.stringify({
    vitalFlags: evalResult.vitalFlags,
    stoolBad: evalResult.stoolBad,
    stoolHours: evalResult.stoolHours,
    urineBad: evalResult.urineBad,
    urineHours: evalResult.urineHours,
  })}
最新バイタル保存値: ${JSON.stringify(evalResult.snapshot ?? {})}

上記に対し、下剤検討・主治医報告の要否など具体的な行動を根拠付きで。`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 512 },
  };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error(data?.error?.message ?? 'AI応答なし');
  return stripJsonFence(text);
}

/** デモ用: 初回のみサンプルイベント・バイタル・緊急連絡先 */
export function seedDemoIfEmpty(residents) {
  if (localStorage.getItem(LS.seeded)) return;
  if (!residents?.length) return;
  const r0 = residents[0];
  const id = String(r0.id);
  setResidentVitalSnapshot(id, { temp: '37.8', bpUpper: '158', bpLower: '72' });
  setLastStoolIso(id, new Date(Date.now() - 80 * 3600000).toISOString());
  setLastUrineNow(id);
  setEmergencyContact(id, { name: '山田太郎（長男）', tel: '090-0000-0000', relation: '長男' });
  const fac = String(r0.facility ?? '');
  logCareEvent({
    type: 'patrol',
    residentId: id,
    residentName: r0.name,
    facilitySheetTitle: fac,
    meta: { note: '3時間巡視: 異常なし' },
  });
  logCareEvent({
    type: 'meal',
    residentId: id,
    residentName: r0.name,
    facilitySheetTitle: fac,
    meta: { mealValue: '8', mealTime: '昼' },
  });
  logCareEvent({
    type: 'excretion',
    residentId: id,
    residentName: r0.name,
    facilitySheetTitle: fac,
    meta: { stool: '普通量' },
  });
  logVitalSnapshot(id, String(r0.name), fac, { temp: '37.8', bpUpper: '158', bpLower: '72' });
  localStorage.setItem(LS.seeded, '1');
}

function escapeHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function nl2br(s) {
  return escapeHtml(s).replace(/\n/g, '<br/>');
}

export function buildEmergencySummaryHtml(resident, evalResult, aiAdvice, contact, draft = {}) {
  const name = String(resident.name ?? '');
  const room = String(resident.room ?? '');
  const cond = String(resident.condition ?? '—');
  const today = new Date().toLocaleDateString('ja-JP');
  const week = getWeeklyVitalTimeline(String(resident.id));
  const rows = week
    .map(
      (e) =>
        `<tr><td>${escapeHtml(e.ts)}</td><td>${escapeHtml(JSON.stringify(e.meta ?? {}))}</td></tr>`
    )
    .join('');
  const senderOffice = String(draft.senderOffice ?? '').trim();
  const senderAddress = String(draft.senderAddress ?? '').trim();
  const senderTel = String(draft.senderTel ?? '').trim();
  const senderNurse = String(draft.senderNurse ?? '').trim();
  const primaryDoctor = String(draft.primaryDoctor ?? '').trim();
  const medicalAgency = String(draft.medicalAgency ?? '').trim();
  const medicalAddress = String(draft.medicalAddress ?? '').trim();
  const dailyLife = String(draft.dailyLife ?? '').trim();
  const nurseProblems = String(draft.nurseProblems ?? '').trim();
  const acuteChange = String(draft.acuteChange ?? '').trim();
  const nurseContent = String(draft.nurseContent ?? '').trim();
  const careNotes = String(draft.careNotes ?? '').trim();
  const other = String(draft.other ?? '').trim();

  return `
<!DOCTYPE html><html><head><meta charset="utf-8"/><title>救急搬送サマリー ${name}</title>
<style>
  body{font-family:system-ui,sans-serif;padding:16px;color:#111}
  h1{font-size:22px;border-bottom:2px solid #c00;padding-bottom:8px}
  h2{margin:16px 0 8px 0}
  table{border-collapse:collapse;width:100%;font-size:12px;margin-top:8px}
  th,td{border:1px solid #ccc;padding:6px;text-align:left}
  .box{border:1px solid #333;padding:12px;margin:12px 0;background:#fafafa}
  .muted{color:#666;font-size:12px}
  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
  .vtop{vertical-align:top}
  @media print { .no-print{display:none} }
</style></head><body>
  <h1>訪問看護の情報（療養に係る情報）提供書 / 救急搬送サマリー</h1>
  <div class="muted">作成日: ${escapeHtml(today)}</div>
  <div class="box">
    <div><strong>施設名</strong> ${escapeHtml(senderOffice || '（未入力）')}</div>
    <div><strong>住所</strong> ${escapeHtml(senderAddress || '（未入力）')}</div>
    <div><strong>電話</strong> ${escapeHtml(senderTel || '（未入力）')} &nbsp; <strong>担当看護師</strong> ${escapeHtml(senderNurse || '（未入力）')}</div>
  </div>
  <div class="box">
    <strong>氏名</strong> ${escapeHtml(name)} 様 &nbsp; <strong>居室</strong> ${escapeHtml(room)}<br/>
    <strong>主疾患/状態</strong> ${escapeHtml(cond)}<br/>
    <strong>緊急連絡先</strong> ${escapeHtml(contact.name)}（${escapeHtml(contact.relation)}） ${escapeHtml(contact.tel)}
  </div>
  <table>
    <tr><th style="width:160px">主治医氏名</th><td>${escapeHtml(primaryDoctor || '（未入力）')}</td></tr>
    <tr><th>医療機関名</th><td>${escapeHtml(medicalAgency || '（未入力）')}</td></tr>
    <tr><th>所在地</th><td>${escapeHtml(medicalAddress || '（未入力）')}</td></tr>
  </table>
  <table>
    <tr><th style="width:220px">日常生活等の状況</th><td class="vtop">${nl2br(dailyLife || '（未入力）')}</td></tr>
    <tr><th>看護上の問題等</th><td class="vtop">${nl2br(nurseProblems || '（未入力）')}</td></tr>
    <tr><th>急変の内容（看護師記入）</th><td class="vtop">${nl2br(acuteChange || '（未入力）')}</td></tr>
    <tr><th>看護の内容</th><td class="vtop">${nl2br(nurseContent || '（未入力）')}</td></tr>
    <tr><th>ケア時の注意点</th><td class="vtop">${nl2br(careNotes || '（未入力）')}</td></tr>
    <tr><th>その他</th><td class="vtop">${nl2br(other || '（未入力）')}</td></tr>
  </table>
  <h2>直近1週間 バイタル記録ログ</h2>
  <table><thead><tr><th>日時</th><th>内容</th></tr></thead><tbody>${rows || '<tr><td colspan="2">記録なし（記録蓄積後に表示）</td></tr>'}</tbody></table>
  <h2>現在の自動検知</h2>
  <div class="box"><pre style="white-space:pre-wrap;margin:0">${JSON.stringify(
    {
      vitalFlags: evalResult.vitalFlags,
      stoolHours: evalResult.stoolHours,
      stoolBad: evalResult.stoolBad,
      urineHours: evalResult.urineHours,
      urineBad: evalResult.urineBad,
    },
    null,
    2
  )}</pre></div>
  <h2>AIアドバイス（参考）</h2>
  <div class="box"><pre style="white-space:pre-wrap;margin:0">${escapeHtml(String(aiAdvice ?? ''))}</pre></div>
  <p class="no-print" style="margin-top:24px;font-size:12px;color:#666">ブラウザの印刷から PDF 保存可能です。</p>
</body></html>`;
}

/**
 * 月次家族向け報告（印刷用HTML）
 * @param {Record<string, unknown>} resident
 * @param {string} yearMonth YYYY-MM
 * @param {{ monthlyCondition?: string; futureCarePoints?: string; directorMessage?: string }} draft
 */
export function buildMonthlyFamilyReportHtml(resident, yearMonth, draft = {}) {
  const name = String(resident?.name ?? '');
  const room = String(resident?.room ?? '');
  const ym = String(yearMonth ?? '');
  const today = new Date().toLocaleDateString('ja-JP');
  const monthlyCondition = String(draft.monthlyCondition ?? '').trim();
  const futureCarePoints = String(draft.futureCarePoints ?? '').trim();
  const directorMessage = String(draft.directorMessage ?? '').trim();
  return `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"/><title>月次ご報告 ${name}</title>
<style>
  body{font-family:system-ui,sans-serif;padding:20px;color:#111;max-width:720px;margin:0 auto;line-height:1.65}
  h1{font-size:1.35rem;border-bottom:2px solid #1e40af;padding-bottom:8px}
  h2{font-size:1.05rem;margin-top:1.25rem;color:#1e3a8a}
  .meta{color:#64748b;font-size:0.85rem;margin-bottom:1rem}
  .box{border:1px solid #cbd5e1;border-radius:12px;padding:14px 16px;margin-top:8px;background:#f8fafc}
  @media print{.no-print{display:none}}
</style></head><body>
  <h1>月次ご報告</h1>
  <div class="meta">作成日: ${escapeHtml(today)} / 対象月: ${escapeHtml(ym)} / ${escapeHtml(name)} 様 / 居室 ${escapeHtml(room)}</div>
  <h2>1か月の状態と様子</h2>
  <div class="box">${nl2br(monthlyCondition || '（未入力）')}</div>
  <h2>今後気を付けていくこと</h2>
  <div class="box">${nl2br(futureCarePoints || '（未入力）')}</div>
  <h2>施設長から一言</h2>
  <div class="box">${nl2br(directorMessage || '（未入力）')}</div>
  <p class="no-print" style="margin-top:24px;font-size:12px;color:#64748b">ブラウザの印刷から PDF 保存できます。</p>
</body></html>`;
}

/**
 * 事故報告書（印刷用HTML）。略図は canvas の data URL を渡す。
 * @param {Record<string, string>} draft
 * @param {string | null | undefined} sketchDataUrl PNG data URL または空
 * @param {{ preview?: boolean }} [opts] preview 時は画面内 iframe 向けの縮小表示
 */
export function buildAccidentReportHtml(draft, sketchDataUrl, opts = {}) {
  const preview = Boolean(opts?.preview);
  const d = draft ?? {};
  const v = (k) => String(d[k] ?? '').trim();
  const cell = (k) => escapeHtml(v(k));
  const block = (k) => nl2br(v(k) || '（未入力）');
  const sketchUrl = String(sketchDataUrl ?? '');
  const sketch =
    sketchUrl.startsWith('data:image') && !/["<>]/.test(sketchUrl)
      ? `<img src="${sketchUrl}" alt="略図" style="max-width:100%;max-height:140px;object-fit:contain;display:block;margin:0 auto;"/>`
      : '<span style="color:#bbb;font-size:9pt;">（発生状況の図解記入欄）</span>';

  const previewCss = preview
    ? `
body.accident-report-preview{background:#e5e7eb;padding:8px;}
body.accident-report-preview .page{width:100%;max-width:210mm;min-height:auto;margin:0 auto;padding:10px;box-shadow:0 1px 4px rgba(0,0,0,.12);}
body.accident-report-preview .page th,body.accident-report-preview .page td{font-size:8.2pt;padding:3px 5px;}
body.accident-report-preview .page th{width:72px;}
body.accident-report-preview .title-area h1{font-size:16pt;}
`
    : '';

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>事故報告書</title>
<style>
:root{--primary-color:#000000;--text-dark:#333333;--border-color:#444444;--bg-light:#f0f0f0;}
body{font-family:"Helvetica Neue",Arial,"Hiragino Kaku Gothic ProN","Hiragino Sans",Meiryo,sans-serif;color:var(--text-dark);background:#fff;margin:0;padding:0;line-height:1.35;}
.page{width:210mm;min-height:285mm;margin:0 auto;background:#fff;padding:12mm;box-sizing:border-box;position:relative;}
${previewCss}
.header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:3px;}
.title-area h1{font-size:20pt;margin:0;color:var(--primary-color);letter-spacing:0.2em;border-bottom:2px solid var(--primary-color);padding-bottom:3px;}
.version{font-size:8pt;color:#666;margin-top:3px;}
.date-report{text-align:right;font-size:9pt;margin-bottom:5px;}
.approval-table{border-collapse:collapse;margin-left:auto;}
.approval-table td{border:1px solid var(--border-color);width:65px;text-align:center;font-size:7.5pt;padding:1px;}
.stamp-box{height:50px;}
table{width:100%;border-collapse:collapse;margin-bottom:5px;table-layout:fixed;}
th,td{border:1px solid var(--border-color);padding:5px 8px;font-size:9.5pt;vertical-align:middle;}
th{background-color:var(--bg-light);font-weight:bold;text-align:center;width:95px;}
.section-title{background-color:var(--primary-color);color:#fff;padding:3px 10px;font-size:9.5pt;font-weight:bold;margin-top:5px;}
.sketch-area{border:1px dashed var(--border-color);height:140px;display:flex;justify-content:center;align-items:center;background-color:#fafafa;}
.print-val{font-size:9.5pt;color:#000;white-space:pre-wrap;word-break:break-word;}
.inline-num{display:inline-block;min-width:1.2em;text-align:center;border-bottom:1px solid #ccc;}
@media print{
  @page{size:A4 portrait;margin:5mm;}
  body{background:#fff!important;margin:0!important;padding:0!important;}
  .page{margin:0!important;width:100%!important;min-height:0!important;padding:5mm 10mm!important;}
  *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}
}
</style>
</head>
<body${preview ? ' class="accident-report-preview"' : ''}>
<div class="page" id="report-content">
  <div class="date-report">
    報告日：20<span class="inline-num">${cell('reportYear2')}</span>年
    <span class="inline-num">${cell('reportMonth')}</span>月
    <span class="inline-num">${cell('reportDay')}</span>日
  </div>
  <div class="header">
    <div class="title-area">
      <h1>事故報告書</h1>
      <div class="version">ver.2</div>
    </div>
    <table class="approval-table">
      <tr><td>部長(又は代理)</td><td>管理者</td></tr>
      <tr><td class="stamp-box"></td><td class="stamp-box"></td></tr>
      <tr><td>（　 / 　）</td><td>（　 / 　）</td></tr>
    </table>
  </div>
  <table style="margin-top:5px;">
    <tr>
      <th>報告者</th><td><div class="print-val">${cell('reporterName')}</div></td>
      <th style="width:60px;">職種</th><td><div class="print-val">${cell('reporterJob')}</div></td>
      <th style="width:60px;">所属</th><td><div class="print-val">${cell('reporterDept')}</div></td>
    </tr>
    <tr>
      <th>発生日時</th>
      <td colspan="3">
        20<span class="inline-num">${cell('occurYear2')}</span>年
        <span class="inline-num">${cell('occurMonth')}</span>月
        <span class="inline-num">${cell('occurDay')}</span>日
        （<span class="inline-num">${cell('occurDayNote')}</span>）
        <span style="border:1px solid #ccc;padding:0 5px;margin:0 5px;">${cell('occurAmPm') || 'AM ・ PM'}</span>
        <span class="inline-num">${cell('occurHour')}</span>時
        <span class="inline-num">${cell('occurMinute')}</span>分 頃
      </td>
      <th>発生場所</th><td><div class="print-val">${cell('occurPlace')}</div></td>
    </tr>
    <tr>
      <th>利用者名</th><td colspan="2"><span class="print-val" style="font-size:11pt">${cell('residentName')}</span> 様</td>
      <th>性別・年齢</th><td colspan="2"><div class="print-val">${cell('genderAge')}</div></td>
    </tr>
    <tr>
      <th>事故の種類</th>
      <td colspan="5">
        <div>転倒 ・ 転落 ・ 落薬 ・ 誤薬 ・ その他（<span class="print-val">${cell('accidentTypeDetail')}</span>）</div>
      </td>
    </tr>
  </table>
  <div class="section-title">医療機関情報</div>
  <table>
    <tr>
      <th>医療機関名</th>
      <td colspan="5"><div class="print-val">${cell('medicalInstitutionName')}</div></td>
    </tr>
    <tr>
      <th>機関コード</th>
      <td colspan="5"><div class="print-val">${cell('medicalInstitutionCode')}</div></td>
    </tr>
    <tr>
      <th>所在地</th>
      <td colspan="5"><div class="print-val">${cell('medicalInstitutionAddress')}</div></td>
    </tr>
    <tr>
      <th>電話</th>
      <td colspan="5"><div class="print-val">${cell('medicalInstitutionTel')}</div></td>
    </tr>
  </table>
  <div style="display:flex;gap:8px;">
    <div style="flex:2;">
      <div class="section-title">事故の発生状況を具体的に記入</div>
      <div style="border:1px solid var(--border-color);border-top:none;padding:6px;min-height:140px;"><div class="print-val">${block('situation')}</div></div>
    </div>
    <div style="flex:1;">
      <div class="section-title">略図</div>
      <div class="sketch-area">${sketch}</div>
    </div>
  </div>
  <div class="section-title">発生直後の対応・処置</div>
  <div style="border:1px solid var(--border-color);border-top:none;padding:6px;">
    <div class="print-val" style="min-height:60px;">${block('response')}</div>
    <div style="text-align:right;font-size:8.5pt;margin-top:3px;">
      家族への報告：（ <span class="inline-num">${cell('familyReportMonth')}</span> 月 <span class="inline-num">${cell('familyReportDay')}</span> 日 ）
    </div>
  </div>
  <div class="section-title">原因として考えられる事</div>
  <div style="border:1px solid var(--border-color);border-top:none;padding:6px;min-height:55px;"><div class="print-val">${block('causes')}</div></div>
  <div class="section-title">今後の対応・改善策</div>
  <div style="border:1px solid var(--border-color);border-top:none;padding:6px;min-height:70px;"><div class="print-val">${block('improvements')}</div></div>
  <table style="margin-top:8px;">
    <tr>
      <th style="height:60px;">上司の所見</th>
      <td colspan="2" style="vertical-align:top;"><div class="print-val" style="min-height:55px;">${block('supervisorOpinion')}</div></td>
      <td style="width:150px;text-align:center;vertical-align:top;padding-top:8px;">
        <div style="font-size:8.5pt;font-weight:bold;margin-bottom:3px;">再検討の必要性</div>
        <div style="margin:5px 0;"><span style="border:1px solid #ccc;padding:2px 10px;font-size:9pt;">${cell('reviewNeeded') || 'あり ・ なし'}</span></div>
      </td>
    </tr>
    <tr><th>その他</th><td colspan="3"><div class="print-val">${block('otherNotes')}</div></td></tr>
  </table>
</div>
</body>
</html>`;
}

/**
 * メモから事故報告の主要欄を下書き（JSON）
 * @param {string} apiKey
 * @param {string} memo
 * @param {string} residentHint
 */
export async function fetchAccidentReportAssist(apiKey, memo, residentHint = '') {
  if (!apiKey?.trim()) {
    return { situation: '', response: '', causes: '', improvements: '' };
  }
  const prompt = `あなたは介護施設の事故報告書作成支援です。次のメモのみを根拠に、JSONオブジェクト1つだけを返してください（説明文不要）。
キー: situation, response, causes, improvements（各値は日本語の文字列。不明な点は「※要確認」と書く）
利用者・状況ヒント: ${String(residentHint || '').trim() || '（なし）'}
メモ:
${String(memo || '').trim() || '（なし）'}`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 1024 },
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error(data?.error?.message ?? 'AI応答なし');
  const raw = stripJsonFence(text);
  try {
    const parsed = JSON.parse(raw);
    return {
      situation: String(parsed.situation ?? ''),
      response: String(parsed.response ?? ''),
      causes: String(parsed.causes ?? ''),
      improvements: String(parsed.improvements ?? ''),
    };
  } catch {
    return { situation: raw, response: '', causes: '', improvements: '' };
  }
}

/** 事故報告下書きのキー（音声→AI 一括生成用） */
const ACCIDENT_VOICE_DRAFT_KEYS = [
  'reportYear2',
  'reportMonth',
  'reportDay',
  'reporterName',
  'reporterJob',
  'reporterDept',
  'occurYear2',
  'occurMonth',
  'occurDay',
  'occurDayNote',
  'occurAmPm',
  'occurHour',
  'occurMinute',
  'occurPlace',
  'residentName',
  'genderAge',
  'accidentTypeDetail',
  'situation',
  'response',
  'familyReportMonth',
  'familyReportDay',
  'causes',
  'improvements',
  'supervisorOpinion',
  'reviewNeeded',
  'otherNotes',
];

function padAccident2(v) {
  const n = parseInt(String(v ?? '').replace(/\D/g, ''), 10);
  if (!Number.isFinite(n)) return String(v ?? '').trim().padStart(2, '0').slice(-2);
  return String(n).padStart(2, '0').slice(-2);
}

function normalizeYear2Field(v) {
  const s = String(v ?? '').trim().replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
  if (/^\d{4}$/.test(s)) return String(parseInt(s, 10) % 100).padStart(2, '0');
  if (/^\d{2}$/.test(s)) return s;
  const d = s.replace(/\D/g, '');
  if (d.length >= 2) return d.slice(-2).padStart(2, '0');
  return s.slice(0, 2).padStart(2, '0');
}

/**
 * 音声メモから事故報告書の全欄を JSON で生成（1回の API）
 * @param {string} apiKey
 * @param {string} voiceMemo
 * @param {{ facilityLabel?: string; residentName?: string; room?: string; reporterDeptPreset?: string }} [context]
 */
export async function fetchAccidentReportFromVoiceMemo(apiKey, voiceMemo, context = {}) {
  if (!apiKey?.trim()) throw new Error('APIキーが設定されていません');
  const memo = String(voiceMemo ?? '').trim();
  if (!memo) throw new Error('話した内容がありません。音声入力してください。');

  const fac = String(context.facilityLabel ?? '').trim();
  const resName = String(context.residentName ?? '').trim();
  const room = String(context.room ?? '').trim();
  const deptPreset = String(context.reporterDeptPreset ?? '').trim();
  const today = new Date();
  const ty = today.getFullYear() % 100;
  const tm = String(today.getMonth() + 1).padStart(2, '0');
  const td = String(today.getDate()).padStart(2, '0');

  const keyList = ACCIDENT_VOICE_DRAFT_KEYS.join(', ');

  const prompt = `あなたは介護施設の事故報告書作成担当です。職員の「音声で話した内容」のみを根拠に、公式事故報告書に転記するJSONオブジェクトを1つだけ返してください（説明文・Markdownフェンス禁止）。

【今日の日付（報告日のデフォルト）】20${String(ty).padStart(2, '0')}年 ${tm}月 ${td}日

【コンテキスト（音声に明示が無いときは優先して埋める）】
- 施設: ${fac || '（不明）'}
- 利用者氏名: ${resName || '（未選択）'}
- 居室: ${room || '—'}
- 所属部署（音声で部署名が無ければ reporterDept にこの値を使う）: ${deptPreset || '（不明）'}

【JSONのキー（すべて文字列。必ずすべて含める）】
${keyList}

【各フィールドの意味】
- reportYear2 / occurYear2: 西暦の下2桁（例 26）
- reportMonth, reportDay, occurMonth, occurDay, familyReportMonth, familyReportDay: 2桁または1桁の月日（先頭0可）
- occurAmPm: 「午前」または「午後」または空
- occurHour, occurMinute: 数字のみの文字列が望ましい
- genderAge: 例「（ 男 ）　85 歳」や「（ 女 ・ 男 ）　 歳」のように書式に近づける
- accidentTypeDetail: 転倒・転落・誤薬等の補足（テンプレの「その他」括弧内）
- situation: 事故の発生状況（具体的に）
- response: 発生直後の対応・処置
- causes, improvements, supervisorOpinion, otherNotes: 該当が無ければ空
- reviewNeeded: 「あり」「なし」または空

不明・音声に無い項目は空文字。推測できない重要事項のみ「※要確認」。利用者名はコンテキストと矛盾させない。

音声の内容:
${memo}`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 8192 },
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error(data?.error?.message ?? 'AI応答なし');
  const raw = stripJsonFence(text);
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error('AIのJSONを解釈できませんでした');
  }

  /** @type {Record<string, string>} */
  const out = {};
  for (const k of ACCIDENT_VOICE_DRAFT_KEYS) {
    out[k] = String(parsed[k] ?? '').trim();
  }

  for (const yk of ['reportYear2', 'occurYear2']) {
    if (out[yk]) out[yk] = normalizeYear2Field(out[yk]);
  }
  for (const mk of ['reportMonth', 'reportDay', 'occurMonth', 'occurDay', 'familyReportMonth', 'familyReportDay']) {
    if (out[mk]) out[mk] = padAccident2(out[mk]);
  }

  if (!out.reporterDept && deptPreset) out.reporterDept = deptPreset;
  if (!out.residentName && resName) out.residentName = resName;

  return out;
}

function newNearMissReportId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function normalizeNearMissCategories(arr) {
  const allowed = new Set([...NEAR_MISS_CATEGORY_LABELS, 'その他']);
  if (!Array.isArray(arr)) return [];
  return [...new Set(arr.map((x) => String(x).trim()).filter((x) => allowed.has(x)))];
}

/** @param {string} msg */
function isGeminiQuotaMessage(msg) {
  return /quota|429|resource_exhausted|rate limit|free_tier|please retry/i.test(String(msg ?? ''));
}

/**
 * AI不可時でも最低限のヒヤリ下書きを返す
 * @param {string} memo
 * @param {number} ty
 * @param {number} tm
 * @param {number} td
 */
function buildNearMissFallbackDraft(memo, ty, tm, td) {
  const lines = String(memo ?? '')
    .split(/\r?\n/)
    .map((s) => s.replace(/^[・●\-\*\s]+/u, '').trim())
    .filter(Boolean);
  const firstLine = lines[0] || '状況確認中';
  const situation = lines.length
    ? `${lines.join('。')}。`
    : `${firstLine}です。`;
  const residentHit = lines.join(' ').match(/([一-龯々ぁ-んァ-ヶーA-Za-z]+)\s*様/u);
  const placeHit = lines.join(' ').match(/(居室|デイホール|食堂|廊下|トイレ|浴室|玄関|ベッドサイド|共有部|フロア)/u);
  return {
    reporterName: '',
    reporterDept: '',
    residentName: residentHit ? String(residentHit[1]) : '',
    occurPlace: placeHit ? String(placeHit[1]) : '',
    occurAmPm: '',
    occurHour: '',
    occurMinute: '',
    occurYear: null,
    occurMonth: null,
    occurDay: null,
    submitYear: ty,
    submitMonth: tm,
    submitDay: td,
    situationContent: situation,
    afterReportContent: '関係者へ共有し、再発防止のため見守りを強化しました。',
    causeAndMeasures: '要因を整理し、環境調整と声掛け手順の見直しを行います。',
    categories: [],
    categoryOther: '',
  };
}

/**
 * 箇条書きメモからヒヤリハット報告の下書き（JSON）。文体はです・ます調で統一。
 * @param {string} apiKey
 * @param {string} memo
 * @param {string} [facilityLabel]
 */
export async function fetchNearMissReportFromBullets(apiKey, memo, facilityLabel = '') {
  const today = new Date();
  const ty = today.getFullYear();
  const tm = today.getMonth() + 1;
  const td = today.getDate();
  const labelList = [...NEAR_MISS_CATEGORY_LABELS, 'その他'].join('、');

  if (!apiKey?.trim()) return buildNearMissFallbackDraft(memo, ty, tm, td);

  const prompt = `あなたは介護施設の専属文書アシスタントです。次の箇条書き業務メモのみを根拠に、ヒヤリハット(気づき)報告書用のJSONオブジェクト1つだけを返してください（説明文やMarkdownのフェンス禁止）。

【文体】situationContent, afterReportContent, causeAndMeasures の本文は、必ず敬体の「です・ます」調で統一してください。

【キーと型】
- reporterName: string
- reporterDept: string（所属事業所・部署）
- residentName: string
- occurPlace: string
- occurAmPm: "午前" | "午後" | "" のいずれか
- occurHour: string（時、数字のみ推奨）
- occurMinute: string（分）
- occurYear, occurMonth, occurDay: number | null（発生日・西暦。メモにない場合はnull）
- submitYear, submitMonth, submitDay: number（提出日。メモにない場合は本日: ${ty}年${tm}月${td}日）
- situationContent: string（セクション1【状況】＝「内容」）
- afterReportContent: string（セクション1【対応】＝報告後のフォロー・共有内容）
- causeAndMeasures: string（セクション2【原因と今後の対策】）
- categories: string[]（次のラベルから該当のみ厳密一致: ${labelList}）
- categoryOther: string（「その他」欄の補足。不要なら空文字）

施設タブ表示名（参考）: ${String(facilityLabel || '').trim() || '（不明）'}

メモ:
${String(memo || '').trim() || '（なし）'}`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 2048 },
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!res.ok || !text) {
    const msg = String(data?.error?.message ?? '');
    if (isGeminiQuotaMessage(msg)) return buildNearMissFallbackDraft(memo, ty, tm, td);
    throw new Error(msg || 'AI応答なし');
  }
  const raw = stripJsonFence(text);
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    throw new Error('AIのJSONを解釈できませんでした');
  }

  let cats = normalizeNearMissCategories(parsed.categories);
  const coOther = String(parsed.categoryOther ?? '').trim();
  if (coOther && !cats.includes('その他')) cats = [...cats, 'その他'];

  const num = (v, fallback) => {
    const n = typeof v === 'number' && Number.isFinite(v) ? v : parseInt(String(v ?? '').trim(), 10);
    return Number.isFinite(n) ? n : fallback;
  };

  return {
    reporterName: String(parsed.reporterName ?? ''),
    reporterDept: String(parsed.reporterDept ?? ''),
    residentName: String(parsed.residentName ?? ''),
    occurPlace: String(parsed.occurPlace ?? ''),
    occurAmPm: String(parsed.occurAmPm ?? ''),
    occurHour: String(parsed.occurHour ?? ''),
    occurMinute: String(parsed.occurMinute ?? ''),
    occurYear: parsed.occurYear != null ? num(parsed.occurYear, null) : null,
    occurMonth: parsed.occurMonth != null ? num(parsed.occurMonth, null) : null,
    occurDay: parsed.occurDay != null ? num(parsed.occurDay, null) : null,
    submitYear: num(parsed.submitYear, ty),
    submitMonth: num(parsed.submitMonth, tm),
    submitDay: num(parsed.submitDay, td),
    situationContent: String(parsed.situationContent ?? ''),
    afterReportContent: String(parsed.afterReportContent ?? ''),
    causeAndMeasures: String(parsed.causeAndMeasures ?? ''),
    categories: cats,
    categoryOther: String(parsed.categoryOther ?? ''),
  };
}

/**
 * @param {{ facilityLabel: string; department: string; residentId?: string; draft: Record<string, unknown> }} p
 * @returns {{ id: string; savedAt: string; facilityLabel: string; department: string; residentId: string; draft: Record<string, unknown> } | null}
 */
export function saveNearMissReport(p) {
  const facilityLabel = String(p?.facilityLabel ?? '').trim();
  const department = String(p?.department ?? '').trim();
  if (!facilityLabel || !department) return null;
  const list = getNearMissReports();
  const entry = {
    id: newNearMissReportId(),
    savedAt: new Date().toISOString(),
    facilityLabel,
    department,
    residentId: String(p?.residentId ?? '').trim(),
    draft: { ...(p?.draft && typeof p.draft === 'object' ? p.draft : {}) },
  };
  list.unshift(entry);
  writeJson(LS.nearMissReports, list.slice(0, MAX_NEAR_MISS_REPORTS));
  return entry;
}

/** @returns {unknown[]} */
export function getNearMissReports() {
  const raw = readJson(LS.nearMissReports, []);
  return Array.isArray(raw) ? raw : [];
}

function newAccidentReportId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

/** 発生日（下書き）を YYYY-MM-DD に。無効なら null */
export function occurrenceYmdFromDraft(draft) {
  const d = draft ?? {};
  const y2 = String(d.occurYear2 ?? '')
    .trim()
    .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .padStart(2, '0');
  const m = String(d.occurMonth ?? '')
    .trim()
    .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .padStart(2, '0');
  const day = String(d.occurDay ?? '')
    .trim()
    .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .padStart(2, '0');
  if (!/^\d{2}$/.test(y2) || !/^\d{2}$/.test(m) || !/^\d{2}$/.test(day)) return null;
  const year = 2000 + parseInt(y2, 10);
  const mo = parseInt(m, 10);
  const da = parseInt(day, 10);
  const dt = new Date(year, mo - 1, da);
  if (dt.getFullYear() !== year || dt.getMonth() !== mo - 1 || dt.getDate() !== da) return null;
  return `${year}-${m}-${day}`;
}

/**
 * ヒヤリ報告下書きの発生日（4桁年・月・日）。無効・未入力なら null
 * @param {Record<string, unknown>} draft
 */
export function nearMissOccurrenceYmdFromDraft(draft) {
  const d = draft ?? {};
  const norm = (v) =>
    String(v ?? '')
      .trim()
      .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
  const y = parseInt(norm(d.occurYear), 10);
  const mo = parseInt(norm(d.occurMonth), 10);
  const day = parseInt(norm(d.occurDay), 10);
  if (!Number.isFinite(y) || y < 1990 || y > 2100) return null;
  if (!Number.isFinite(mo) || mo < 1 || mo > 12) return null;
  if (!Number.isFinite(day) || day < 1 || day > 31) return null;
  const ms = String(mo).padStart(2, '0');
  const ds = String(day).padStart(2, '0');
  const dt = new Date(y, mo - 1, day);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== day) return null;
  return `${y}-${ms}-${ds}`;
}

/**
 * @param {Record<string, unknown>} draft
 * @returns {string[]}
 */
function classifyNearMissRecordCategories(draft) {
  const cats = Array.isArray(draft?.categories)
    ? draft.categories.map((x) => String(x).trim()).filter(Boolean)
    : [];
  const other = String(draft?.categoryOther ?? '').trim();
  const uniq = [...new Set(cats)];
  if (other && !uniq.includes('その他')) uniq.push('その他');
  if (!uniq.length) return ['分類なし'];
  return uniq;
}

/** 発生時刻を 0–23 時。判定できなければ null */
export function parseOccurHour24(draft) {
  const d = draft ?? {};
  const hStr = String(d.occurHour ?? '')
    .trim()
    .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
  const hourNum = parseInt(hStr, 10);
  if (!Number.isFinite(hourNum)) return null;
  const ap = String(d.occurAmPm ?? '').trim();
  const isPM = /午後|ＰＭ|PM|pm/.test(ap);
  const isAM = /午前|ＡＭ|AM|am/.test(ap);
  const hasAp = isPM || isAM;
  if (hourNum >= 0 && hourNum <= 23 && !hasAp) return hourNum;
  if (hourNum >= 1 && hourNum <= 12 && hasAp) {
    if (isPM) return hourNum === 12 ? 12 : hourNum + 12;
    return hourNum === 12 ? 0 : hourNum;
  }
  if (hourNum >= 0 && hourNum <= 23) return hourNum;
  return null;
}

export function hourToTimeSlot(hour24) {
  if (hour24 == null || !Number.isFinite(hour24)) return '時間不明';
  const h = Math.floor(Number(hour24));
  if (h >= 0 && h <= 5) return '深夜（0–5時）';
  if (h >= 6 && h <= 8) return '早朝（6–8時）';
  if (h >= 9 && h <= 11) return '午前（9–11時）';
  if (h >= 12 && h <= 13) return '昼（12–13時）';
  if (h >= 14 && h <= 17) return '午後（14–17時）';
  if (h >= 18 && h <= 20) return '夕方（18–20時）';
  if (h >= 21 && h <= 23) return '夜（21–23時）';
  return '時間不明';
}

export function classifyAccidentType(draft) {
  const text = `${String(draft?.accidentTypeDetail ?? '')}\n${String(draft?.situation ?? '')}`.replace(/\s/g, '');
  if (/転落/.test(text)) return '転落';
  if (/転倒/.test(text)) return '転倒';
  if (/誤薬/.test(text)) return '誤薬';
  if (/落薬/.test(text)) return '落薬';
  if (/誤嚥|窒息|むせ|嚥下異常/.test(text)) return '窒息・誤嚥';
  if (/徘徊/.test(text)) return '徘徊';
  if (/火傷|やけど/.test(text)) return 'やけど・火傷';
  if (/自傷/.test(text)) return '自傷行為';
  return 'その他';
}

/**
 * 事故報告を部署・施設単位でローカル保存（月次集計用）
 * @param {{ facilityLabel: string; department: string; residentId?: string; draft: Record<string, unknown> }} p
 * @returns {{ id: string; savedAt: string; facilityLabel: string; department: string; residentId: string; draft: Record<string, unknown> } | null}
 */
export function saveAccidentReport(p) {
  const facilityLabel = String(p?.facilityLabel ?? '').trim();
  const department = String(p?.department ?? '').trim();
  if (!facilityLabel || !department) return null;
  const list = getAccidentReports();
  const entry = {
    id: newAccidentReportId(),
    savedAt: new Date().toISOString(),
    facilityLabel,
    department,
    residentId: String(p?.residentId ?? '').trim(),
    draft: { ...(p?.draft && typeof p.draft === 'object' ? p.draft : {}) },
  };
  list.unshift(entry);
  writeJson(LS.accidentReports, list.slice(0, MAX_ACCIDENT_REPORTS));
  return entry;
}

/** @returns {Array<{ id: string; savedAt: string; facilityLabel: string; department: string; residentId: string; draft: Record<string, unknown> }>} */
export function getAccidentReports() {
  const raw = readJson(LS.accidentReports, []);
  return Array.isArray(raw) ? raw : [];
}

/**
 * 指定月の事故を種類・時間帯で集計
 * @param {string} yearMonth YYYY-MM
 * @param {{ facilityLabel?: string; department?: string }} [filters]
 */
export function aggregateAccidentMonthlySummary(yearMonth, filters = {}) {
  const ym = String(yearMonth ?? '').trim();
  const fac = String(filters.facilityLabel ?? '').trim();
  const dep = String(filters.department ?? '').trim();
  const prefix = ym.length === 7 ? `${ym}-` : '';
  const all = getAccidentReports();
  /** @type {Record<string, number>} */
  const byType = {};
  /** @type {Record<string, number>} */
  const bySlot = {};
  const records = [];
  let total = 0;

  for (const row of all) {
    if (fac && String(row.facilityLabel ?? '').trim() !== fac) continue;
    if (dep && String(row.department ?? '').trim() !== dep) continue;
    const ymd = occurrenceYmdFromDraft(row.draft);
    if (!ymd || (prefix && !ymd.startsWith(prefix))) continue;
    const hour = parseOccurHour24(row.draft);
    const slot = hourToTimeSlot(hour);
    const typ = classifyAccidentType(row.draft);
    total += 1;
    byType[typ] = (byType[typ] ?? 0) + 1;
    bySlot[slot] = (bySlot[slot] ?? 0) + 1;
    records.push({
      ...row,
      _occurrenceYmd: ymd,
      _hour24: hour,
      _slot: slot,
      _type: typ,
    });
  }
  records.sort((a, b) => {
    const c = String(b._occurrenceYmd).localeCompare(String(a._occurrenceYmd));
    if (c !== 0) return c;
    return String(b.savedAt).localeCompare(String(a.savedAt));
  });
  return {
    total,
    byType,
    bySlot,
    records,
    yearMonth: ym,
    filters: { facilityLabel: fac, department: dep },
  };
}

function sortedCountRows(map, order) {
  const seen = new Set();
  const rows = [];
  for (const k of order) {
    const n = map[k];
    if (n > 0) {
      rows.push({ key: k, count: n });
      seen.add(k);
    }
  }
  for (const [k, n] of Object.entries(map)) {
    if (!seen.has(k) && n > 0) rows.push({ key: k, count: n });
  }
  return rows;
}

/**
 * @param {ReturnType<typeof aggregateAccidentMonthlySummary>} agg
 * @param {string} [assessmentText]
 */
export function buildAccidentMonthlyAnalysisHtml(agg, assessmentText = '') {
  const typeRows = sortedCountRows(agg.byType, ACCIDENT_TYPE_ORDER);
  const slotRows = sortedCountRows(agg.bySlot, ACCIDENT_SLOT_ORDER);
  const fac = agg.filters?.facilityLabel ? escapeHtml(agg.filters.facilityLabel) : '全施設';
  const dep = agg.filters?.department ? escapeHtml(agg.filters.department) : '全部署';
  const ym = escapeHtml(agg.yearMonth);
  const typeTable =
    typeRows.length === 0
      ? '<tr><td colspan="2">該当なし</td></tr>'
      : typeRows.map((r) => `<tr><td>${escapeHtml(r.key)}</td><td style="text-align:right">${r.count}</td></tr>`).join('');
  const slotTable =
    slotRows.length === 0
      ? '<tr><td colspan="2">該当なし</td></tr>'
      : slotRows.map((r) => `<tr><td>${escapeHtml(r.key)}</td><td style="text-align:right">${r.count}</td></tr>`).join('');
  const listRows = agg.records
    .slice(0, 80)
    .map((r) => {
      const name = escapeHtml(String(r.draft?.residentName ?? '').trim() || '—');
      const typ = escapeHtml(String(r._type));
      const sl = escapeHtml(String(r._slot));
      const dept = escapeHtml(String(r.department ?? ''));
      return `<tr><td>${escapeHtml(r._occurrenceYmd)}</td><td>${dept}</td><td>${name}</td><td>${typ}</td><td>${sl}</td></tr>`;
    })
    .join('');
  const assess = nl2br(String(assessmentText ?? '').trim() || '（未生成。ブラウザ上で「アセスメント生成」を実行してください。）');

  return `<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/><title>事故月次分析 ${agg.yearMonth}</title>
<style>
body{font-family:system-ui,sans-serif;padding:16px;color:#111}
h1{font-size:20px;border-bottom:2px solid #333;padding-bottom:8px}
h2{font-size:16px;margin:20px 0 8px}
.meta{color:#555;font-size:13px;margin-bottom:16px}
table{border-collapse:collapse;width:100%;max-width:720px;font-size:13px;margin-top:8px}
th,td{border:1px solid #ccc;padding:8px;text-align:left}
th{background:#f0f0f0}
.box{border:1px solid #333;padding:12px;margin:16px 0;background:#fafafa;white-space:pre-wrap}
@media print{.no-print{display:none}}
</style></head><body>
<h1>事故報告 月次集計・アセスメント</h1>
<div class="meta">対象月: <strong>${ym}</strong> ／ 施設: <strong>${fac}</strong> ／ 部署: <strong>${dep}</strong> ／ 件数: <strong>${agg.total}</strong></div>
<h2>事故の種類別 件数</h2>
<table><thead><tr><th>種類</th><th>件数</th></tr></thead><tbody>${typeTable}</tbody></table>
<h2>発生時間帯別 件数</h2>
<table><thead><tr><th>時間帯</th><th>件数</th></tr></thead><tbody>${slotTable}</tbody></table>
<h2>アセスメント（参考）</h2>
<div class="box">${assess}</div>
<h2>明細（最大80件）</h2>
<table><thead><tr><th>発生日</th><th>部署</th><th>利用者</th><th>分類</th><th>時間帯</th></tr></thead><tbody>${
    listRows || '<tr><td colspan="5">該当なし</td></tr>'
  }</tbody></table>
<p class="no-print" style="margin-top:24px;font-size:12px;color:#666">CareLink OS — ブラウザの印刷から PDF 保存できます。</p>
</body></html>`;
}

function csvEscape(s) {
  const t = String(s ?? '');
  if (/[",\n\r]/.test(t)) return `"${t.replace(/"/g, '""')}"`;
  return t;
}

/** @param {ReturnType<typeof aggregateAccidentMonthlySummary>} agg */
export function buildAccidentMonthlyCsv(agg) {
  const header = ['発生日', '保存日時', '施設', '部署', '利用者名', '事故分類', '時間帯', '発生状況要約'];
  const lines = [header.join(',')];
  for (const r of agg.records) {
    const sit = String(r.draft?.situation ?? '').replace(/\s+/g, ' ').trim().slice(0, 200);
    lines.push(
      [
        r._occurrenceYmd,
        r.savedAt,
        r.facilityLabel,
        r.department,
        r.draft?.residentName ?? '',
        r._type,
        r._slot,
        sit,
      ]
        .map(csvEscape)
        .join(',')
    );
  }
  return lines.join('\r\n');
}

/**
 * 月次集計結果に基づくアセスメント文案（Gemini）
 * @param {string} apiKey
 * @param {ReturnType<typeof aggregateAccidentMonthlySummary>} agg
 */
export async function fetchAccidentMonthlyAssessmentAi(apiKey, agg) {
  if (!apiKey?.trim()) {
    return 'API キー未設定のため、集計表のみご利用ください。傾向は「種類別」「時間帯別」の件数からご判断ください。';
  }
  const typeRows = sortedCountRows(agg.byType, ACCIDENT_TYPE_ORDER);
  const slotRows = sortedCountRows(agg.bySlot, ACCIDENT_SLOT_ORDER);
  const samples = agg.records.slice(0, 12).map((r) => ({
    type: r._type,
    slot: r._slot,
    dept: r.department,
    situation: String(r.draft?.situation ?? '').slice(0, 400),
    causes: String(r.draft?.causes ?? '').slice(0, 200),
    improvements: String(r.draft?.improvements ?? '').slice(0, 200),
  }));
  const prompt = `あなたは介護・看護の安全管理者です。次の「1か月分の事故報告集計（同一ブラウザに保存された記録）」を踏まえ、施設内アセスメントとして日本語で簡潔にまとめてください。

【集計条件】対象月: ${agg.yearMonth} / 施設フィルタ: ${agg.filters?.facilityLabel || '全施設'} / 部署フィルタ: ${agg.filters?.department || '全部署'} / 合計件数: ${agg.total}

【種類別件数】
${typeRows.map((r) => `- ${r.key}: ${r.count}件`).join('\n') || '（なし）'}

【時間帯別件数】
${slotRows.map((r) => `- ${r.key}: ${r.count}件`).join('\n') || '（なし）'}

【参考メモ（抜粋・最大12件。個人名は出力に含めない）】
${JSON.stringify(samples, null, 2)}

出力は次の構成で、見出し付き箇条書きを中心に（総括・種類の傾向・時間帯の傾向・リスク要因の推測・今月の重点対策案・記録上の留意）。推測は「〜の可能性」として書き、断定しすぎないこと。`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.25, maxOutputTokens: 2048 },
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error(data?.error?.message ?? 'AI応答なし');
  return stripJsonFence(text);
}

/**
 * 指定月のヒヤリハット（気づき）をカテゴリ・時間帯で集計。発生日未入力は保存月で判定。
 * @param {string} yearMonth YYYY-MM
 * @param {{ facilityLabel?: string; department?: string }} [filters]
 */
export function aggregateNearMissMonthlySummary(yearMonth, filters = {}) {
  const ym = String(yearMonth ?? '').trim();
  const fac = String(filters.facilityLabel ?? '').trim();
  const dep = String(filters.department ?? '').trim();
  const all = getNearMissReports();
  /** @type {Record<string, number>} */
  const byCategory = {};
  /** @type {Record<string, number>} */
  const bySlot = {};
  const records = [];
  let total = 0;

  for (const row of all) {
    if (fac && String(row.facilityLabel ?? '').trim() !== fac) continue;
    if (dep && String(row.department ?? '').trim() !== dep) continue;

    const ymd = nearMissOccurrenceYmdFromDraft(row.draft);
    let rowYm;
    if (ymd) rowYm = ymd.slice(0, 7);
    else {
      const d = new Date(String(row.savedAt ?? ''));
      if (!Number.isFinite(d.getTime())) continue;
      rowYm = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    }
    if (rowYm !== ym) continue;

    const hour = parseOccurHour24(row.draft);
    const slot = hourToTimeSlot(hour);
    const cats = classifyNearMissRecordCategories(row.draft);
    total += 1;
    for (const c of cats) {
      byCategory[c] = (byCategory[c] ?? 0) + 1;
    }
    bySlot[slot] = (bySlot[slot] ?? 0) + 1;
    records.push({
      ...row,
      _occurrenceYmd: ymd || String(row.savedAt ?? '').slice(0, 10),
      _hour24: hour,
      _slot: slot,
      _categories: cats,
    });
  }
  records.sort((a, b) => {
    const c = String(b._occurrenceYmd).localeCompare(String(a._occurrenceYmd));
    if (c !== 0) return c;
    return String(b.savedAt).localeCompare(String(a.savedAt));
  });
  return {
    total,
    byCategory,
    bySlot,
    records,
    yearMonth: ym,
    filters: { facilityLabel: fac, department: dep },
  };
}

/**
 * @param {ReturnType<typeof aggregateNearMissMonthlySummary>} agg
 * @param {string} [assessmentText]
 */
export function buildNearMissMonthlyAnalysisHtml(agg, assessmentText = '') {
  const catRows = sortedCountRows(agg.byCategory, NEAR_MISS_MONTH_CATEGORY_ORDER);
  const slotRows = sortedCountRows(agg.bySlot, ACCIDENT_SLOT_ORDER);
  const fac = agg.filters?.facilityLabel ? escapeHtml(agg.filters.facilityLabel) : '全施設';
  const dep = agg.filters?.department ? escapeHtml(agg.filters.department) : '全部署';
  const ym = escapeHtml(agg.yearMonth);
  const catTable =
    catRows.length === 0
      ? '<tr><td colspan="2">該当なし</td></tr>'
      : catRows.map((r) => `<tr><td>${escapeHtml(r.key)}</td><td style="text-align:right">${r.count}</td></tr>`).join('');
  const slotTable =
    slotRows.length === 0
      ? '<tr><td colspan="2">該当なし</td></tr>'
      : slotRows.map((r) => `<tr><td>${escapeHtml(r.key)}</td><td style="text-align:right">${r.count}</td></tr>`).join('');
  const listRows = agg.records
    .slice(0, 80)
    .map((r) => {
      const name = escapeHtml(String(r.draft?.residentName ?? '').trim() || '—');
      const cats = escapeHtml((r._categories ?? []).join('・'));
      const sl = escapeHtml(String(r._slot));
      const dept = escapeHtml(String(r.department ?? ''));
      return `<tr><td>${escapeHtml(String(r._occurrenceYmd))}</td><td>${dept}</td><td>${name}</td><td>${cats}</td><td>${sl}</td></tr>`;
    })
    .join('');
  const assess = nl2br(String(assessmentText ?? '').trim() || '（未生成。ブラウザ上で「アセスメント生成」を実行してください。）');

  return `<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/><title>ヒヤリ月次分析 ${agg.yearMonth}</title>
<style>
body{font-family:system-ui,sans-serif;padding:16px;color:#111}
h1{font-size:20px;border-bottom:2px solid #0f766e;padding-bottom:8px}
h2{font-size:16px;margin:20px 0 8px}
.meta{color:#555;font-size:13px;margin-bottom:16px}
table{border-collapse:collapse;width:100%;max-width:720px;font-size:13px;margin-top:8px}
th,td{border:1px solid #ccc;padding:8px;text-align:left}
th{background:#ecfdf5}
.box{border:1px solid #333;padding:12px;margin:16px 0;background:#fafafa;white-space:pre-wrap}
@media print{.no-print{display:none}}
</style></head><body>
<h1>ヒヤリハット（気づき）月次集計・アセスメント</h1>
<div class="meta">対象月: <strong>${ym}</strong> ／ 施設: <strong>${fac}</strong> ／ 部署: <strong>${dep}</strong> ／ 件数: <strong>${agg.total}</strong>（カテゴリ別件数は複数選択で重複加算の場合あり）</div>
<h2>カテゴリ別 件数</h2>
<table><thead><tr><th>カテゴリ</th><th>件数</th></tr></thead><tbody>${catTable}</tbody></table>
<h2>発生時間帯別 件数</h2>
<table><thead><tr><th>時間帯</th><th>件数</th></tr></thead><tbody>${slotTable}</tbody></table>
<h2>アセスメント（参考）</h2>
<div class="box">${assess}</div>
<h2>明細（最大80件）</h2>
<table><thead><tr><th>発生日または保存日</th><th>部署</th><th>利用者</th><th>カテゴリ</th><th>時間帯</th></tr></thead><tbody>${
    listRows || '<tr><td colspan="5">該当なし</td></tr>'
  }</tbody></table>
<p class="no-print" style="margin-top:24px;font-size:12px;color:#666">CareLink OS — ブラウザの印刷から PDF 保存できます。</p>
</body></html>`;
}

/** @param {ReturnType<typeof aggregateNearMissMonthlySummary>} agg */
export function buildNearMissMonthlyCsv(agg) {
  const header = ['発生日または保存日', '保存日時', '施設', '部署', '利用者名', 'カテゴリ', '時間帯', '状況要約'];
  const lines = [header.join(',')];
  for (const r of agg.records) {
    const sit = String(r.draft?.situationContent ?? '')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 200);
    lines.push(
      [
        r._occurrenceYmd,
        r.savedAt,
        r.facilityLabel,
        r.department,
        r.draft?.residentName ?? '',
        (r._categories ?? []).join('・'),
        r._slot,
        sit,
      ]
        .map(csvEscape)
        .join(',')
    );
  }
  return lines.join('\r\n');
}

/**
 * @param {string} apiKey
 * @param {ReturnType<typeof aggregateNearMissMonthlySummary>} agg
 */
export async function fetchNearMissMonthlyAssessmentAi(apiKey, agg) {
  if (!apiKey?.trim()) {
    return 'API キー未設定のため、集計表のみご利用ください。傾向は「カテゴリ別」「時間帯別」の件数からご判断ください。';
  }
  const catRows = sortedCountRows(agg.byCategory, NEAR_MISS_MONTH_CATEGORY_ORDER);
  const slotRows = sortedCountRows(agg.bySlot, ACCIDENT_SLOT_ORDER);
  const samples = agg.records.slice(0, 12).map((r) => ({
    categories: r._categories,
    slot: r._slot,
    dept: r.department,
    situation: String(r.draft?.situationContent ?? '').slice(0, 400),
    causeAndMeasures: String(r.draft?.causeAndMeasures ?? '').slice(0, 200),
  }));
  const prompt = `あなたは介護・看護の安全管理者です。次の「1か月分のヒヤリハット（気づき）報告の集計（同一ブラウザに保存された記録）」を踏まえ、施設内アセスメントとして日本語で簡潔にまとめてください。

【集計条件】対象月: ${agg.yearMonth} / 施設フィルタ: ${agg.filters?.facilityLabel || '全施設'} / 部署フィルタ: ${agg.filters?.department || '全部署'} / 報告件数（1報告＝1件）: ${agg.total}

【カテゴリ別件数】※1件で複数カテゴリのときは複数カウントされる場合があります
${catRows.map((r) => `- ${r.key}: ${r.count}`).join('\n') || '（なし）'}

【時間帯別件数】
${slotRows.map((r) => `- ${r.key}: ${r.count}件`).join('\n') || '（なし）'}

【参考メモ（抜粋・最大12件。個人名は出力に含めない）】
${JSON.stringify(samples, null, 2)}

出力は次の構成で、見出し付き箇条書きを中心に（総括・カテゴリの傾向・時間帯の傾向・再発防止の観点・今月の重点教育・記録上の留意）。推測は「〜の可能性」として書き、断定しすぎないこと。`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.25, maxOutputTokens: 2048 },
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
  let res;
  /** @type {any} */
  let data;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    data = await res.json();
  } catch (e) {
    const hint = e instanceof Error ? e.message : String(e);
    throw new Error(
      [
        '【通信に失敗しました】',
        'インターネット接続やブラウザの制限（広告ブロック等）を確認してください。',
        '',
        '────────',
        '（技術）',
        hint,
      ].join('\n')
    );
  }

  if (data?.error) {
    throw new Error(formatGeminiGenerateContentErrorMessage(data, res.status));
  }
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!String(text ?? '').trim()) {
    throw new Error(formatGeminiGenerateContentErrorMessage(data, res.status));
  }
  return stripJsonFence(String(text));
}

/**
 * 最終勤務尿（シフト終了時の排尿記録）印刷用HTML
 * @param {{
 *   facilityLabel?: string;
 *   recordDate?: string;
 *   recordTime?: string;
 *   shiftKind?: string;
 *   residentName?: string;
 *   room?: string;
 *   urineMl?: string;
 *   appearance?: string;
 *   catheterNote?: string;
 *   note?: string;
 *   recorderName?: string;
 * }} [draft]
 */
export function buildLastShiftUrineFormHtml(draft = {}) {
  const d = draft && typeof draft === 'object' ? draft : {};
  const v = (k) => String(d[k] ?? '').trim();
  const facilityLabel = v('facilityLabel');
  const recordDate = v('recordDate');
  const recordTime = v('recordTime');
  const shiftKind = v('shiftKind');
  const residentName = v('residentName');
  const room = v('room');
  const urineMl = v('urineMl');
  const appearance = v('appearance');
  const catheterNote = v('catheterNote');
  const note = v('note');
  const recorderName = v('recorderName');
  const title = '最終勤務尿（排尿記録）';
  return `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"/><title>${escapeHtml(title)}</title>
<style>
  body{font-family:system-ui,sans-serif;padding:18px;color:#111;max-width:720px;margin:0 auto;line-height:1.55;font-size:13px}
  h1{font-size:1.25rem;border-bottom:2px solid #0f766e;padding-bottom:8px;margin:0 0 12px 0}
  .meta{color:#64748b;font-size:0.82rem;margin-bottom:14px}
  table{border-collapse:collapse;width:100%;margin-top:8px}
  th,td{border:1px solid #94a3b8;padding:8px 10px;text-align:left;vertical-align:top}
  th{width:9.5rem;background:#f1f5f9;font-weight:700}
  .free{min-height:4.5rem}
  @media print{.no-print{display:none}body{padding:12px}}
</style></head><body>
  <h1>${escapeHtml(title)}</h1>
  <div class="meta">紙の様式に合わせて運用・文言は現場で調整してください。ブラウザの印刷で PDF 保存できます。</div>
  <table>
    <tr><th>事業所</th><td>${escapeHtml(facilityLabel || '（未入力）')}</td></tr>
    <tr><th>記録日</th><td>${escapeHtml(recordDate || '（未入力）')}</td></tr>
    <tr><th>記録時刻</th><td>${escapeHtml(recordTime || '（未入力）')}</td></tr>
    <tr><th>勤務帯</th><td>${escapeHtml(shiftKind || '（未入力）')}</td></tr>
    <tr><th>利用者氏名</th><td>${escapeHtml(residentName || '（未入力）')}</td></tr>
    <tr><th>居室</th><td>${escapeHtml(room || '（未入力）')}</td></tr>
    <tr><th>排尿量</th><td>${escapeHtml(urineMl || '（未入力）')}</td></tr>
    <tr><th>性状・色など</th><td class="free">${nl2br(appearance || '（未入力）')}</td></tr>
    <tr><th>カテーテル・バルーン等</th><td class="free">${nl2br(catheterNote || '（未入力）')}</td></tr>
    <tr><th>特記事項</th><td class="free">${nl2br(note || '（未入力）')}</td></tr>
    <tr><th>記録者</th><td>${escapeHtml(recorderName || '（未入力）')}</td></tr>
  </table>
  <p class="no-print" style="margin-top:18px;font-size:11px;color:#94a3b8">CareLink Facility Portal — ${escapeHtml(title)}</p>
</body></html>`;
}

export function openPrintableSummary(html) {
  const w = window.open('', '_blank');
  if (!w) return false;
  w.document.write(html);
  w.document.close();
  w.focus();
  setTimeout(() => w.print(), 300);
  return true;
}

export function downloadSummaryHtml(filename, html) {
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}
