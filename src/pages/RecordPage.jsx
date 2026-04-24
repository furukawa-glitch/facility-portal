import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Activity,
  AlertCircle,
  AlertTriangle,
  Ambulance,
  Baby,
  BarChart3,
  CalendarClock,
  CalendarDays,
  Check,
  ChevronLeft,
  ClipboardList,
  Clock,
  ExternalLink,
  FileSpreadsheet,
  FileText,
  FileWarning,
  Home,
  LayoutGrid,
  Loader2,
  Megaphone,
  MessageSquarePlus,
  Mic,
  Monitor,
  RefreshCw,
  Save,
  Smartphone,
  Sparkles,
  Stethoscope,
  Table2,
  Scale,
  Upload,
  Utensils,
  Wind,
  X,
} from 'lucide-react';
import {
  CARELINK_RESIDENT_SPREADSHEET_ID,
  careLevelScoreForAverageCareLevel,
  fetchResidentsFromSheet,
  formatCareLevelForDisplay,
  getAverageCareLevelFromSheetSummary,
  getMedicalTargetCountFromSheetSummary,
  getResidentCountFromSheetSummary,
  normalizeCareLevelLabel,
  parseCsv,
} from '../services/GoogleSheetService.js';
import {
  CARELINK_FACILITIES,
  compactFacilityToken,
  facilityDefBySheetTitle,
  linkKeyForSheetTitle,
  residentBelongsToFacilityTab,
} from '../config/carelinkFacilities.js';
import { getExternalLinksForFacility } from '../config/facilityIntegrations.js';
import {
  STOOL_VOLUME_OPTIONS,
  STOOL_CHARACTER_OPTIONS,
  MEAL_WARI_OPTIONS,
  composeMealAmountForLog,
  parseVoiceToStoolVolume,
  parseVoiceToStoolCharacter,
  parseVoiceToMealWari,
  parseVoiceToWaterMl,
} from '../lib/careQuickCareFields.js';
import {
  PATROL_SLOT_HOURS,
  defaultPatrolSlotDateTimeLocal,
  joinPatrolDateTimeLocal,
  normalizePatrolDateTimeLocal,
  splitPatrolDateTimeLocal,
} from '../lib/patrolSlots.js';
import { AccidentMonthlyAnalysisModal } from '../components/AccidentMonthlyAnalysisModal.jsx';
import { AccidentReportModal } from '../components/AccidentReportModal.jsx';
import { NearMissAwarenessAdminModal } from '../components/NearMissAwarenessAdminModal.jsx';
import { NearMissAwarenessPanel } from '../components/NearMissAwarenessPanel.jsx';
import { NearMissMonthlyAnalysisModal } from '../components/NearMissMonthlyAnalysisModal.jsx';
import { NearMissReportModal } from '../components/NearMissReportModal.jsx';
import { ResidentBulkInputTable } from '../components/ResidentBulkInputTable.jsx';
import { isNursingOfficeUiEnabled } from '../services/NearMissLedgerService.js';
import * as Report from '../services/ReportService.js';

/** 名簿に「様」付きで入っているときの重複を避ける */
function residentNameWithoutSama(nameRaw) {
  return String(nameRaw ?? '')
    .replace(/様\s*$/u, '')
    .trim();
}

/** 一覧入力・保存後にクリアするケア項目（バイタル列は残す） */
const BULK_CARE_RESET = Object.freeze({
  patrol: false,
  patrolAt: '',
  meal: false,
  excretion: false,
  urineVolume: '',
  stoolVolume: '',
  stoolCharacter: '',
  mealSlot: '',
  mealStaple: '',
  mealSide: '',
  waterMl: '',
  medicationTaken: '',
  toiletGuidance: false,
});

/** @param {string} k */
function insuranceCategoryChipClass(k) {
  if (k === '医療保険特指示')
    return 'border-amber-500 bg-amber-100 text-amber-950 ring-2 ring-amber-400/70';
  if (k === '医療') return 'border-emerald-600 bg-emerald-50 text-emerald-950 ring-1 ring-emerald-300';
  if (k === '未設定') return 'border-slate-300 bg-slate-100 text-slate-800';
  return 'border-sky-500 bg-white text-sky-950 shadow-sm ring-1 ring-sky-200';
}

const GEMINI_KEY = import.meta.env.VITE_GEMINI_API_KEY ?? '';
const SHEETS_KEY = import.meta.env.VITE_GOOGLE_SHEETS_API_KEY ?? '';

const MONITOR_BOARD_BY_FACILITY = {};

const MONITOR_BOARD_FALLBACK = {
  notice:
    '【周知】インフルエンザ流行に伴い、面会はマスク着用でお願いします。異常時はナースステーションへ連絡ください。',
  handover: '【申し送り】夜勤より：特記事項なし（サンプル）。',
  schedule: [
    { time: '10:00', title: '面会（サンプル）' },
    { time: '14:30', title: '往診（サンプル）' },
  ],
};

function boardForFacilityLinkKey(linkKey) {
  const k = String(linkKey ?? '').trim();
  if (k && MONITOR_BOARD_BY_FACILITY[k]) return MONITOR_BOARD_BY_FACILITY[k];
  return MONITOR_BOARD_FALLBACK;
}

function ExternalToolButton({ href, icon: Icon, children, disabled, layout = 'stack' }) {
  const isHash = !href || href === '#';
  const inline = layout === 'inline';
  return (
    <a
      href={isHash ? undefined : href}
      target={isHash ? undefined : '_blank'}
      rel={isHash ? undefined : 'noopener noreferrer'}
      onClick={(e) => isHash && e.preventDefault()}
      className={`flex min-h-[3.25rem] items-center justify-center gap-2 rounded-2xl border-2 px-2 py-2 text-center font-bold transition-all xl:min-h-0 xl:py-4 ${
        inline ? 'flex-1 flex-row' : 'flex-1 flex-col gap-1 xl:flex-none'
      } ${
        isHash
          ? 'cursor-not-allowed border-slate-600 bg-slate-800/50 text-slate-500'
          : 'border-cyan-500/40 bg-slate-800 text-cyan-100 hover:border-cyan-400 hover:bg-slate-700'
      } ${disabled ? 'pointer-events-none opacity-50' : ''}`}
    >
      <Icon className={`shrink-0 stroke-[2] ${inline ? 'h-5 w-5' : 'h-6 w-6 xl:h-7 xl:w-7'}`} />
      <span className={`leading-tight ${inline ? 'text-sm' : 'text-sm xl:text-base'}`}>{children}</span>
      {!isHash && !inline && <ExternalLink className="h-3.5 w-3.5 opacity-60" aria-hidden />}
    </a>
  );
}

function currentYearMonth() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function currentYmd() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/** @param {unknown} value */
function toEventIsoOrNow(value) {
  const s = String(value ?? '').trim();
  if (!s) return new Date().toISOString();
  const t = new Date(s).getTime();
  if (!Number.isFinite(t)) return new Date().toISOString();
  return new Date(t).toISOString();
}

function emptyEmergencyDraft() {
  return {
    senderOffice: '',
    senderAddress: '',
    senderTel: '',
    senderNurse: '',
    primaryDoctor: '',
    medicalAgency: '',
    medicalAddress: '',
    dailyLife: '',
    nurseProblems: '',
    acuteChange: '',
    nurseContent: '',
    careNotes: '',
    other: '',
  };
}

/**
 * @param {{
 *   onSelectResident: (res: Record<string, unknown>) => void;
 *   onBack: () => void;
 *   onOpenMonthlyReport: () => void;
 *   onOpenNotionNewResidents?: () => void;
 *   onResidentsSync?: (list: Record<string, unknown>[]) => void;
 *   initialSheetTitle?: string;
 * }} props
 */
export function RecordPage({
  onSelectResident,
  onBack,
  onOpenMonthlyReport,
  onOpenNotionNewResidents,
  onResidentsSync,
  initialSheetTitle,
}) {
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState('');
  const [lastUpdated, setLastUpdated] = useState(/** @type {Date | null} */ (null));
  const [fetchSourceMeta, setFetchSourceMeta] = useState(
    /** @type {{ source: string; mode: string } | null} */ (null)
  );
  const [allResidents, setAllResidents] = useState(/** @type {Record<string, unknown>[]} */ ([]));
  const [selectedSheetTitle, setSelectedSheetTitle] = useState(() => {
    const t = String(initialSheetTitle ?? '').trim();
    if (t && CARELINK_FACILITIES.some((f) => f.sheetTitle === t)) return t;
    return CARELINK_FACILITIES[0].sheetTitle;
  });
  const [tick, setTick] = useState(0);
  const loadSeqRef = useRef(0);

  useEffect(() => {
    const bump = () => setTick((n) => n + 1);
    window.addEventListener('carelink-staff-profile', bump);
    return () => window.removeEventListener('carelink-staff-profile', bump);
  }, []);

  const [nursingDraft, setNursingDraft] = useState('');
  const [nursingStartDate, setNursingStartDate] = useState(currentYmd);
  const [nursingEndDate, setNursingEndDate] = useState('');
  const [nursingRev, setNursingRev] = useState(0);
  const [planDraftDate, setPlanDraftDate] = useState(currentYmd);
  const [planDraftTime, setPlanDraftTime] = useState('10:00');
  const [planDraftType, setPlanDraftType] = useState('外出');
  const [planDraftTitle, setPlanDraftTitle] = useState('');
  const [planRev, setPlanRev] = useState(0);

  const [emergencyOpen, setEmergencyOpen] = useState(false);
  const [emergencyPickId, setEmergencyPickId] = useState('');
  const [emergencyBusy, setEmergencyBusy] = useState(false);
  const [accidentReportOpen, setAccidentReportOpen] = useState(false);
  const [accidentMonthlyOpen, setAccidentMonthlyOpen] = useState(false);
  const [nearMissOpen, setNearMissOpen] = useState(false);
  const [nearMissMonthlyOpen, setNearMissMonthlyOpen] = useState(false);
  const [nearMissAwarenessAdminOpen, setNearMissAwarenessAdminOpen] = useState(false);
  const [emergencyDraft, setEmergencyDraft] = useState(emptyEmergencyDraft);
  const [dictatingField, setDictatingField] = useState('');
  const dictationRef = useRef(/** @type {SpeechRecognition | null} */ (null));
  const quickCareDictationRef = useRef(/** @type {SpeechRecognition | null} */ (null));
  const [dictatingQuickKey, setDictatingQuickKey] = useState('');

  const [calOpenId, setCalOpenId] = useState('');
  const [auditMonth, setAuditMonth] = useState(currentYearMonth);

  const [quickRes, setQuickRes] = useState(/** @type {Record<string, unknown> | null} */ (null));
  const [quickVisitNursing, setQuickVisitNursing] = useState(false);
  const [quickVisitNursingNote, setQuickVisitNursingNote] = useState('');
  const [quickVnPeriodStart, setQuickVnPeriodStart] = useState('');
  const [quickVnPeriodEnd, setQuickVnPeriodEnd] = useState('');
  const [quickTemp, setQuickTemp] = useState('');
  const [quickBpU, setQuickBpU] = useState('');
  const [quickBpL, setQuickBpL] = useState('');
  const [quickPulse, setQuickPulse] = useState('');
  const [quickWeight, setQuickWeight] = useState('');
  const [chkPatrol, setChkPatrol] = useState(false);
  const [quickPatrolAt, setQuickPatrolAt] = useState(() => defaultPatrolSlotDateTimeLocal());
  const [chkExcretion, setChkExcretion] = useState(false);
  const [chkMeal, setChkMeal] = useState(false);
  const [quickCareDetail, setQuickCareDetail] = useState(() => ({
    urineVolume: '',
    stoolVolume: '',
    stoolCharacter: '',
    mealSlot: '',
    mealStaple: '',
    mealSide: '',
    waterMl: '',
    medicationTaken: '',
    toiletGuidance: false,
  }));
  const [quickFlash, setQuickFlash] = useState(false);
  /** クイック記録の保存中（連打で食事ログが重複しないようガード） */
  const [quickSaveBusy, setQuickSaveBusy] = useState(false);
  const quickSaveLockRef = useRef(false);
  /** 'cards' | 'table' — 一覧表でバイタル・巡視等をまとめて入力 */
  const [residentInputView, setResidentInputView] = useState(/** @type {'cards' | 'table'} */ ('cards'));
  /** 一覧表：今回の食事区分（朝・昼・夜）を全員に共通適用 */
  const [bulkGlobalMealSlot, setBulkGlobalMealSlot] = useState('昼');
  const [bulkDraft, setBulkDraft] = useState(
    /** @type {Record<string, { temp: string; bpU: string; bpL: string; pulse: string; patrol: boolean; meal: boolean; excretion: boolean }>} */ ({})
  );
  const kaipokeCsvInputRef = useRef(/** @type {HTMLInputElement | null} */ (null));
  const quickPatrolSlot = useMemo(() => splitPatrolDateTimeLocal(quickPatrolAt), [quickPatrolAt]);

  const selectedDef = useMemo(
    () => facilityDefBySheetTitle(selectedSheetTitle),
    [selectedSheetTitle]
  );

  const { filteredResidents, residentFilterBanner } = useMemo(() => {
    const matched = allResidents.filter((r) => residentBelongsToFacilityTab(r, selectedSheetTitle));
    if (matched.length > 0) {
      return { filteredResidents: matched, residentFilterBanner: null };
    }
    if (allResidents.length === 0) {
      return { filteredResidents: matched, residentFilterBanner: null };
    }

    /** CSV 等: タブ名・施設列が付かない名簿はタブ照合で0件になるため全件表示 */
    const lacksTabBinding = (r) => {
      const f = String(r.facility ?? '').trim();
      return !String(r.sourceSheetTitle ?? '').trim() && (!f || f === '施設未設定');
    };
    if (allResidents.every(lacksTabBinding)) {
      return { filteredResidents: allResidents, residentFilterBanner: null };
    }

    /**
     * 名簿の読み込み元タブが1種類だけ（単一CSV／単一gid／VITE_CSV_DEFAULT_SHEET_TITLE）のとき、
     * 施設列の表記が UI の施設タブと一致しないと0件になるため全件表示する。
     */
    const sources = new Set(
      allResidents.map((r) => String(r.sourceSheetTitle ?? '').trim()).filter(Boolean)
    );
    if (sources.size <= 1) {
      return { filteredResidents: allResidents, residentFilterBanner: null };
    }

    /**
     * 複数タブ読込: タブ名の「核」で突き合わせ（表記ゆれ）
     */
    const core = compactFacilityToken(selectedSheetTitle);
    if (core) {
      const loose = allResidents.filter((r) => {
        const src = String(r.sourceSheetTitle ?? '').trim();
        const fac = String(r.facility ?? '').trim();
        return (
          (src && compactFacilityToken(src) === core) ||
          (fac && compactFacilityToken(fac) === core)
        );
      });
      if (loose.length > 0) {
        return { filteredResidents: loose, residentFilterBanner: null };
      }
    }

    /**
     * それでも0件なら全件表示（施設タブとスプレッドシートのタブ名がずれている場合の救済）
     */
    return {
      filteredResidents: allResidents,
      residentFilterBanner:
        '施設タブと名簿の照合ができなかったため、読み込んだ全利用者を表示しています。ポータルで施設を切り替えるか、carelinkFacilities.js の sheetTitle を実際のタブ名に合わせてください。',
    };
  }, [allResidents, selectedSheetTitle]);

  const insuranceBreakdown = useMemo(() => {
    const m = {};
    for (const r of filteredResidents) {
      const c = String(r.insuranceCategory ?? '未設定').trim() || '未設定';
      m[c] = (m[c] ?? 0) + 1;
    }
    return m;
  }, [filteredResidents]);

  const insuranceBreakdownLabel = useMemo(() => {
    const order = [
      '後期高齢',
      '国保',
      '協会けんぽ',
      '組合健保',
      '公費・その他',
      '医療保険特指示',
      '医療',
      'その他',
      '未設定',
    ];
    const parts = [];
    for (const k of order) {
      const n = insuranceBreakdown[k];
      if (n) parts.push(`${k} ${n}名`);
    }
    for (const k of Object.keys(insuranceBreakdown)) {
      if (!order.includes(k) && insuranceBreakdown[k]) parts.push(`${k} ${insuranceBreakdown[k]}名`);
    }
    return parts.length ? parts.join(' ・ ') : '—';
  }, [insuranceBreakdown]);

  const insuranceMedicalSummary = useMemo(() => {
    const total = filteredResidents.length;
    const unset = insuranceBreakdown['未設定'] ?? 0;
    const recorded = total - unset;
    const kohi = insuranceBreakdown['後期高齢'] ?? 0;
    const medicalNonKohi = Math.max(0, recorded - kohi);
    return { total, unset, recorded, kohi, medicalNonKohi };
  }, [filteredResidents, insuranceBreakdown]);

  /** 名簿の要介護1〜5の平均（要支援・自立は含めない）／医療保険対象列の入居済み医療対象人数 */
  const facilityCareStats = useMemo(() => {
    let scoreSum = 0;
    let scoreN = 0;
    const sheetMedicalTarget = getMedicalTargetCountFromSheetSummary(selectedSheetTitle);
    const sheetAvgCareLevel = getAverageCareLevelFromSheetSummary(selectedSheetTitle);
    let medicalTargetCount = 0;
    for (const r of filteredResidents) {
      const sc = careLevelScoreForAverageCareLevel(String(r.careLevelLabel ?? ''));
      if (sc != null) {
        scoreSum += sc;
        scoreN += 1;
      }
      if (sheetMedicalTarget == null && r.isMedicalInsuranceTarget) medicalTargetCount += 1;
    }
    if (sheetMedicalTarget != null) medicalTargetCount = sheetMedicalTarget;
    const averageCareLevelFromResidents =
      scoreN > 0 ? Math.round((scoreSum / scoreN) * 100) / 100 : null;
    const averageCareLevel =
      sheetAvgCareLevel != null
        ? Math.round(sheetAvgCareLevel * 100) / 100
        : averageCareLevelFromResidents;
    return { averageCareLevel, medicalTargetCount, careLevelScoreCount: scoreN };
  }, [filteredResidents, selectedSheetTitle, lastUpdated]);

  const headerResidentCountSubtitle = useMemo(() => {
    if (!selectedDef) return '施設を選択してください';
    const sheetN = getResidentCountFromSheetSummary(selectedSheetTitle);
    const n =
      sheetN != null && Number.isFinite(sheetN) ? Math.round(sheetN) : filteredResidents.length;
    return `${selectedDef.tabLabel}：${n} 名`;
  }, [selectedDef, selectedSheetTitle, filteredResidents.length, lastUpdated]);

  const visitNursingStats = useMemo(() => {
    const count = Report.countVisitNursingSpecialAmong(filteredResidents);
    const thr = Report.VISIT_NURSING_SPECIAL_WARN_THRESHOLD;
    return {
      count,
      warn: count >= thr,
      threshold: thr,
    };
  }, [filteredResidents, tick]);

  const nursingOfficeUi = useMemo(() => isNursingOfficeUiEnabled(), [tick]);

  const insuranceBreakdownChips = useMemo(() => {
    const order = [
      '後期高齢',
      '国保',
      '協会けんぽ',
      '組合健保',
      '公費・その他',
      '医療保険特指示',
      '医療',
      'その他',
      '未設定',
    ];
    const b = insuranceBreakdown;
    const rows = [];
    for (const k of order) {
      const n = b[k];
      if (n) rows.push({ k, n });
    }
    for (const k of Object.keys(b)) {
      if (!order.includes(k) && b[k]) rows.push({ k, n: b[k] });
    }
    return rows;
  }, [insuranceBreakdown]);

  const billingYearMonth = useMemo(() => {
    void tick;
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }, [tick]);

  const residentBillingById = useMemo(() => {
    void tick;
    const ym = billingYearMonth;
    const m = new Map();
    for (const r of filteredResidents) {
      m.set(String(r.id), Report.summarizeResidentMonthBilling(String(r.id), ym));
    }
    return m;
  }, [filteredResidents, billingYearMonth, tick]);

  const board = useMemo(
    () => boardForFacilityLinkKey(linkKeyForSheetTitle(selectedSheetTitle)),
    [selectedSheetTitle]
  );
  const extLinks = useMemo(
    () => getExternalLinksForFacility(selectedDef?.linkKey ?? ''),
    [selectedDef]
  );

  const nursingList = useMemo(() => {
    const k = selectedDef?.linkKey ?? '';
    return k ? Report.getNursingDirectives(k) : [];
  }, [selectedDef, nursingRev]);
  const weeklyPlanDays = useMemo(() => {
    void tick;
    const k = selectedDef?.linkKey ?? '';
    return k ? Report.getWeeklyPlanDays(k, new Date()) : [];
  }, [selectedDef, planRev, tick]);
  const todayPlans = useMemo(() => {
    const d = weeklyPlanDays.find((x) => x.isToday);
    if (!d || !Array.isArray(d.plans)) return [];
    return d.plans;
  }, [weeklyPlanDays]);
  const selectedEmergencyResident = useMemo(
    () => filteredResidents.find((r) => String(r.id) === String(emergencyPickId)) ?? null,
    [filteredResidents, emergencyPickId]
  );

  const buildEmergencyDraftFromResident = useCallback(
    (resident, prevDraft = null) => {
      if (!resident) return emptyEmergencyDraft();
      const narrative = Report.buildEmergencySummaryNarrativeFromRecords(
        resident,
        selectedSheetTitle,
        selectedDef?.linkKey ?? ''
      );
      const prev = prevDraft && typeof prevDraft === 'object' ? prevDraft : {};
      return {
        senderOffice: String(selectedDef?.emergencyFacilityName ?? selectedDef?.tabLabel ?? '').trim(),
        senderAddress: String(selectedDef?.emergencySenderAddress ?? '').trim(),
        senderTel: String(prev.senderTel ?? '').trim(),
        senderNurse: String(prev.senderNurse ?? '').trim(),
        primaryDoctor: String(prev.primaryDoctor ?? '').trim(),
        medicalAgency: String(prev.medicalAgency ?? '').trim(),
        medicalAddress: String(prev.medicalAddress ?? '').trim(),
        dailyLife: narrative.dailyLife,
        nurseProblems: narrative.nurseProblems,
        acuteChange: String(prev.acuteChange ?? '').trim(),
        nurseContent: narrative.nurseContent,
        careNotes: narrative.careNotes,
        other: String(prev.other ?? '').trim(),
      };
    },
    [selectedDef, selectedSheetTitle]
  );

  const load = useCallback(async (isManualRefresh) => {
    const seq = ++loadSeqRef.current;
    if (isManualRefresh) setRefreshing(true);
    else setLoading(true);
    setError('');
    try {
      const { residents, source, mode } = await fetchResidentsFromSheet({
        forceRefresh: Boolean(isManualRefresh),
      });
      if (seq !== loadSeqRef.current) return;
      setAllResidents(residents);
      setFetchSourceMeta({
        source: String(source ?? ''),
        mode: String(mode ?? ''),
      });
      Report.seedDemoIfEmpty(residents);
      setLastUpdated(new Date());
      setSelectedSheetTitle((prev) => {
        const prevOk =
          prev &&
          CARELINK_FACILITIES.some((def) => def.sheetTitle === prev) &&
          residents.some((r) => residentBelongsToFacilityTab(r, prev));
        if (prevOk) return prev;
        const firstWithData = CARELINK_FACILITIES.find((def) =>
          residents.some((r) => residentBelongsToFacilityTab(r, def.sheetTitle))
        );
        return firstWithData?.sheetTitle ?? CARELINK_FACILITIES[0].sheetTitle;
      });
    } catch (e) {
      if (seq !== loadSeqRef.current) return;
      const raw = e instanceof Error ? e.message : 'データの取得に失敗しました';
      const quotaLike = /quota exceeded|クォータ|429/i.test(raw);
      const hint = quotaLike ?
        ' Google 側の「1分あたりの読み取り」上限です。1〜2分待ってから「更新」を押すか、Cloud Console で Sheets API のクォータを確認してください。'
      : '';
      setError(raw + hint);
      setFetchSourceMeta(null);
      setAllResidents((prev) => (quotaLike && prev.length > 0 ? prev : []));
    } finally {
      if (seq === loadSeqRef.current) {
        setLoading(false);
        setRefreshing(false);
      }
    }
  }, []);

  /** 開いた直後・ポータルから施設が変わったときはキャッシュを使わず必ず再取得する */
  useEffect(() => {
    const t = String(initialSheetTitle ?? '').trim();
    if (t && CARELINK_FACILITIES.some((f) => f.sheetTitle === t)) {
      setSelectedSheetTitle(t);
    }
    void load(true);
  }, [initialSheetTitle, load]);

  useEffect(() => {
    if (!selectedEmergencyResident) {
      setEmergencyDraft(emptyEmergencyDraft());
      return;
    }
    setEmergencyDraft((prev) => buildEmergencyDraftFromResident(selectedEmergencyResident, prev));
  }, [selectedEmergencyResident, buildEmergencyDraftFromResident]);

  useEffect(
    () => () => {
      try {
        dictationRef.current?.stop();
        quickCareDictationRef.current?.stop();
      } catch {
        // noop
      }
    },
    []
  );

  useEffect(() => {
    onResidentsSync?.(filteredResidents);
  }, [filteredResidents, onResidentsSync]);

  useEffect(() => {
    const id = setInterval(() => setTick((n) => n + 1), 15000);
    return () => clearInterval(id);
  }, []);

  const [clock, setClock] = useState(() => new Date());
  useEffect(() => {
    const id = setInterval(() => setClock(new Date()), 30000);
    return () => clearInterval(id);
  }, []);
  const nowLabel = clock.toLocaleString('ja-JP', { dateStyle: 'medium', timeStyle: 'short' });

  const applyCareQuickRecord = useCallback((res, row) => {
    const {
      temp = '',
      bpU = '',
      bpL = '',
      pulse = '',
      weight = '',
      patrol = false,
      patrolAt = '',
      meal = false,
      excretion = false,
      urineVolume = '',
      stoolVolume = '',
      stoolCharacter = '',
      mealSlot = '',
      mealStaple = '',
      mealSide = '',
      mealAmount = '',
      waterMl = '',
      medicationTaken = '',
      toiletGuidance = false,
    } = row;
    const id = String(res.id);
    const fac = String(res.facility ?? selectedSheetTitle);
    const name = String(res.name ?? '');
    const weightTrim = String(weight ?? '').trim();
    const vitalPatch = {
      temp,
      bpUpper: bpU,
      bpLower: bpL,
      pulse,
    };
    if (weightTrim) vitalPatch.weight = weightTrim;
    Report.setResidentVitalSnapshot(id, vitalPatch);
    const snap = Report.getResidentVitalSnapshot(id);
    Report.logVitalSnapshot(id, name, fac, {
      temp: snap?.temp,
      bpUpper: snap?.bpUpper,
      bpLower: snap?.bpLower,
      pulse: snap?.pulse,
      spo2: snap?.spo2,
      weight: snap?.weight,
    });
    if (patrol) {
      const at = normalizePatrolDateTimeLocal(patrolAt);
      Report.logCareEvent({
        type: 'patrol',
        ts: toEventIsoOrNow(at),
        residentId: id,
        residentName: name,
        facilitySheetTitle: fac,
        meta: { note: '3時間おき巡回（クイック）' },
      });
    }
    const u = String(urineVolume ?? '').trim();
    const sv = String(stoolVolume ?? '').trim();
    const sc = String(stoolCharacter ?? '').trim();
    const tg = Boolean(toiletGuidance);
    const hasDetailedEx = u || sv || sc;
    if (hasDetailedEx) {
      Report.logCareEvent({
        type: 'excretion',
        residentId: id,
        residentName: name,
        facilitySheetTitle: fac,
        meta: {
          urineVolume: u,
          stoolVolume: sv,
          stoolCharacter: sc,
          ...(tg ? { toiletGuidance: true } : {}),
        },
      });
      if (sv || sc) Report.recordStoolForIntervalAlert(id, { stoolVolume: sv, stoolCharacter: sc });
      if (u || tg) Report.setLastUrineNow(id);
    } else if (excretion) {
      Report.logCareEvent({
        type: 'excretion',
        residentId: id,
        residentName: name,
        facilitySheetTitle: fac,
        meta: { note: '排泄確認（クイック）' },
      });
      Report.setLastStoolNow(id);
      Report.setLastUrineNow(id);
    } else if (tg) {
      Report.logCareEvent({
        type: 'excretion',
        residentId: id,
        residentName: name,
        facilitySheetTitle: fac,
        meta: { toiletGuidance: true, note: 'トイレ誘導' },
      });
      Report.setLastUrineNow(id);
    }
    const slot = String(mealSlot ?? '').trim();
    const composedMeal = composeMealAmountForLog(mealStaple, mealSide);
    const ma = composedMeal || String(mealAmount ?? '').trim();
    const wm = String(waterMl ?? '').trim();
    const med = medicationTaken === 'yes' || medicationTaken === 'no' ? medicationTaken : '';
    // 上で「朝・昼・夜」を共通指定しても、主食・副食・食事※が空なら水分のみ＝fluid_intake
    const waterOnly = Boolean(wm && !ma && !med && !meal);
    if (waterOnly) {
      Report.logCareEvent({
        type: 'fluid_intake',
        residentId: id,
        residentName: name,
        facilitySheetTitle: fac,
        meta: { waterMl: wm },
      });
    } else if (slot || ma || wm || med) {
      Report.logCareEvent({
        type: 'meal',
        residentId: id,
        residentName: name,
        facilitySheetTitle: fac,
        meta: { mealSlot: slot, mealAmount: ma, waterMl: wm, medicationTaken: med },
      });
    } else if (meal) {
      Report.logCareEvent({
        type: 'meal',
        residentId: id,
        residentName: name,
        facilitySheetTitle: fac,
        meta: { note: '食事確認（クイック）' },
      });
    }
  }, [selectedSheetTitle]);

  const openQuickPanel = useCallback((res) => {
    setQuickRes(res);
    const snap = Report.getResidentVitalSnapshot(String(res.id));
    setQuickTemp(snap?.temp != null ? String(snap.temp) : '');
    setQuickBpU(snap?.bpUpper != null ? String(snap.bpUpper) : '');
    setQuickBpL(snap?.bpLower != null ? String(snap.bpLower) : '');
    setQuickPulse(snap?.pulse != null ? String(snap.pulse) : '');
    setQuickWeight(snap?.weight != null ? String(snap.weight) : '');
    setChkPatrol(false);
    setQuickPatrolAt(defaultPatrolSlotDateTimeLocal());
    setChkExcretion(false);
    setChkMeal(false);
    setQuickCareDetail({
      urineVolume: '',
      stoolVolume: '',
      stoolCharacter: '',
      mealSlot: '',
      mealStaple: '',
      mealSide: '',
      waterMl: '',
      medicationTaken: '',
      toiletGuidance: false,
    });
    const vn = Report.getVisitNursingSpecial(String(res.id));
    setQuickVisitNursing(vn.active);
    setQuickVisitNursingNote(vn.note || '');
    setQuickVnPeriodStart(vn.periodStart || '');
    setQuickVnPeriodEnd(vn.periodEnd || '');
  }, []);

  const saveQuickPanel = useCallback(() => {
    if (!quickRes) return;
    if (quickSaveLockRef.current) return;

    const staple = String(quickCareDetail.mealStaple ?? '').trim();
    const side = String(quickCareDetail.mealSide ?? '').trim();
    const urine = String(quickCareDetail.urineVolume ?? '').trim();
    const stoolVol = String(quickCareDetail.stoolVolume ?? '').trim();
    const stoolChar = String(quickCareDetail.stoolCharacter ?? '').trim();

    if (chkMeal) {
      if (!staple || !side) {
        alert(
          '「食事の確認」にチェックする場合は、主食（割）と副食（割）の両方を選んでください（未入力のままでは保存できません）。'
        );
        return;
      }
    }
    if (chkExcretion) {
      if (!stoolChar) {
        alert('「排泄の確認」にチェックする場合は、排便性状を選んでください。');
        return;
      }
      if (!stoolVol && !urine) {
        alert(
          '「排泄の確認」にチェックする場合は、排尿量または排便量（多・中・小）のいずれかを入力してください。'
        );
        return;
      }
    }

    quickSaveLockRef.current = true;
    setQuickSaveBusy(true);
    const id = String(quickRes.id);
    try {
      applyCareQuickRecord(quickRes, {
        temp: quickTemp,
        bpU: quickBpU,
        bpL: quickBpL,
        pulse: quickPulse,
        weight: quickWeight,
        patrol: chkPatrol,
        patrolAt: quickPatrolAt,
        meal: chkMeal,
        excretion: chkExcretion,
        ...quickCareDetail,
        toiletGuidance: Boolean(quickCareDetail.toiletGuidance),
      });
      Report.setVisitNursingSpecial(id, {
        active: quickVisitNursing,
        note: quickVisitNursingNote,
        periodStart: quickVnPeriodStart,
        periodEnd: quickVnPeriodEnd,
      });
      setTick((n) => n + 1);
      setQuickFlash(true);
      setTimeout(() => setQuickFlash(false), 2200);
    } finally {
      window.setTimeout(() => {
        quickSaveLockRef.current = false;
        setQuickSaveBusy(false);
      }, 700);
    }
  }, [
    quickRes,
    applyCareQuickRecord,
    quickTemp,
    quickBpU,
    quickBpL,
    quickPulse,
    quickWeight,
    chkPatrol,
    quickPatrolAt,
    chkMeal,
    chkExcretion,
    quickCareDetail,
    quickVisitNursing,
    quickVisitNursingNote,
    quickVnPeriodStart,
    quickVnPeriodEnd,
  ]);

  const switchToTableInput = useCallback(() => {
    const init = {};
    for (const r of filteredResidents) {
      const id = String(r.id);
      const snap = Report.getResidentVitalSnapshot(id);
      init[id] = {
        temp: snap?.temp != null ? String(snap.temp) : '',
        bpU: snap?.bpUpper != null ? String(snap.bpUpper) : '',
        bpL: snap?.bpLower != null ? String(snap.bpLower) : '',
        pulse: snap?.pulse != null ? String(snap.pulse) : '',
        weight: snap?.weight != null ? String(snap.weight) : '',
        ...BULK_CARE_RESET,
        mealSlot: bulkGlobalMealSlot,
      };
    }
    setBulkDraft(init);
    setResidentInputView('table');
  }, [filteredResidents, bulkGlobalMealSlot]);

  const onBulkGlobalMealSlotChange = useCallback(
    (slot) => {
      setBulkGlobalMealSlot(slot);
      setBulkDraft((prev) => {
        const next = { ...prev };
        for (const r of filteredResidents) {
          const id = String(r.id);
          const cur = next[id];
          if (cur) next[id] = { ...cur, mealSlot: slot };
        }
        return next;
      });
    },
    [filteredResidents]
  );

  const patchBulkRow = useCallback((id, patch) => {
    setBulkDraft((prev) => {
      const base =
        prev[id] ??
        (() => {
          const snap = Report.getResidentVitalSnapshot(id);
          return {
            temp: snap?.temp != null ? String(snap.temp) : '',
            bpU: snap?.bpUpper != null ? String(snap.bpUpper) : '',
            bpL: snap?.bpLower != null ? String(snap.bpLower) : '',
            pulse: snap?.pulse != null ? String(snap.pulse) : '',
            weight: snap?.weight != null ? String(snap.weight) : '',
            ...BULK_CARE_RESET,
            mealSlot: bulkGlobalMealSlot,
          };
        })();
      return { ...prev, [id]: { ...base, ...patch } };
    });
  }, [bulkGlobalMealSlot]);

  const bulkRowHasInput = useCallback((row) => {
    if (!row) return false;
    return (
      String(row.temp ?? '').trim() !== '' ||
      String(row.bpU ?? '').trim() !== '' ||
      String(row.bpL ?? '').trim() !== '' ||
      String(row.pulse ?? '').trim() !== '' ||
      String(row.weight ?? '').trim() !== '' ||
      row.patrol ||
      row.meal ||
      row.excretion ||
      String(row.urineVolume ?? '').trim() !== '' ||
      String(row.stoolVolume ?? '').trim() !== '' ||
      String(row.stoolCharacter ?? '').trim() !== '' ||
      String(row.mealStaple ?? '').trim() !== '' ||
      String(row.mealSide ?? '').trim() !== '' ||
      String(row.mealAmount ?? '').trim() !== '' ||
      String(row.waterMl ?? '').trim() !== '' ||
      row.medicationTaken === 'yes' ||
      row.medicationTaken === 'no' ||
      row.toiletGuidance
    );
  }, []);

  const saveBulkRow = useCallback(
    (res) => {
      const id = String(res.id);
      const row = bulkDraft[id];
      if (!row || !bulkRowHasInput(row)) return;
      applyCareQuickRecord(res, row);
      setBulkDraft((prev) => ({
        ...prev,
        [id]: {
          ...prev[id],
          ...BULK_CARE_RESET,
          mealSlot: bulkGlobalMealSlot,
        },
      }));
      setTick((n) => n + 1);
    },
    [bulkDraft, applyCareQuickRecord, bulkRowHasInput, bulkGlobalMealSlot]
  );

  const saveBulkAllWithInput = useCallback(() => {
    const toSave = filteredResidents.filter((res) => {
      const row = bulkDraft[String(res.id)];
      return bulkRowHasInput(row);
    });
    if (toSave.length === 0) return;
    for (const res of toSave) {
      applyCareQuickRecord(res, bulkDraft[String(res.id)]);
    }
    setBulkDraft((prev) => {
      const next = { ...prev };
      for (const res of toSave) {
        const id = String(res.id);
        if (next[id])
          next[id] = { ...next[id], ...BULK_CARE_RESET, mealSlot: bulkGlobalMealSlot };
      }
      return next;
    });
    setTick((t) => t + 1);
  }, [filteredResidents, bulkDraft, applyCareQuickRecord, bulkRowHasInput, bulkGlobalMealSlot]);

  const setBulkPatrolForAllVisible = useCallback(
    (checked) => {
      setBulkDraft((prev) => {
        const next = { ...prev };
        for (const r of filteredResidents) {
          const id = String(r.id);
          const base =
            next[id] ??
            (() => {
              const snap = Report.getResidentVitalSnapshot(id);
              return {
                temp: snap?.temp != null ? String(snap.temp) : '',
                bpU: snap?.bpUpper != null ? String(snap.bpUpper) : '',
                bpL: snap?.bpLower != null ? String(snap.bpLower) : '',
                pulse: snap?.pulse != null ? String(snap.pulse) : '',
                weight: snap?.weight != null ? String(snap.weight) : '',
                ...BULK_CARE_RESET,
                mealSlot: bulkGlobalMealSlot,
              };
            })();
          next[id] = {
            ...base,
            patrol: Boolean(checked),
            ...(checked ? { patrolAt: normalizePatrolDateTimeLocal(base.patrolAt) } : {}),
          };
        }
        return next;
      });
    },
    [filteredResidents, bulkGlobalMealSlot]
  );

  useEffect(() => {
    if (residentInputView !== 'table') return;
    setBulkDraft((prev) => {
      const next = { ...prev };
      for (const r of filteredResidents) {
        const id = String(r.id);
        if (!next[id]) {
          const snap = Report.getResidentVitalSnapshot(id);
          next[id] = {
            temp: snap?.temp != null ? String(snap.temp) : '',
            bpU: snap?.bpUpper != null ? String(snap.bpUpper) : '',
            bpL: snap?.bpLower != null ? String(snap.bpLower) : '',
            pulse: snap?.pulse != null ? String(snap.pulse) : '',
            weight: snap?.weight != null ? String(snap.weight) : '',
            ...BULK_CARE_RESET,
            mealSlot: bulkGlobalMealSlot,
          };
        } else {
          const cur = next[id];
          for (const k of Object.keys(BULK_CARE_RESET)) {
            if (cur[k] === undefined) cur[k] = BULK_CARE_RESET[k];
          }
          if (cur.mealSlot === undefined || cur.mealSlot === '') {
            next[id] = { ...cur, mealSlot: bulkGlobalMealSlot };
          }
        }
      }
      const keep = new Set(filteredResidents.map((r) => String(r.id)));
      for (const k of Object.keys(next)) {
        if (!keep.has(k)) delete next[k];
      }
      return next;
    });
  }, [filteredResidents, residentInputView, bulkGlobalMealSlot]);

  const registerNursing = useCallback(() => {
    const k = selectedDef?.linkKey;
    if (!k) return;
    if (
      Report.addNursingDirective(k, nursingDraft, '看護', {
        startDate: nursingStartDate,
        endDate: nursingEndDate,
      })
    ) {
      setNursingDraft('');
      setNursingStartDate(currentYmd());
      setNursingEndDate('');
      setNursingRev((n) => n + 1);
    }
  }, [nursingDraft, nursingStartDate, nursingEndDate, selectedDef]);

  const removeNursing = useCallback(
    (d) => {
      const k = selectedDef?.linkKey;
      if (!k) return;
      if (Report.removeNursingDirective(k, String(d?.id ?? ''), String(d?.ts ?? ''))) {
        setNursingRev((n) => n + 1);
      }
    },
    [selectedDef]
  );

  const registerWeeklyPlan = useCallback(() => {
    const k = selectedDef?.linkKey;
    if (!k) return;
    const ok = Report.addWeeklyPlan(k, {
      date: planDraftDate,
      time: planDraftTime,
      type: planDraftType,
      title: planDraftTitle,
    });
    if (!ok) return;
    setPlanDraftTitle('');
    setPlanRev((n) => n + 1);
  }, [selectedDef, planDraftDate, planDraftTime, planDraftType, planDraftTitle]);

  const removeWeeklyPlan = useCallback(
    (planId) => {
      const k = selectedDef?.linkKey;
      if (!k) return;
      if (Report.removeWeeklyPlan(k, String(planId ?? ''))) {
        setPlanRev((n) => n + 1);
      }
    },
    [selectedDef]
  );

  const startDictation = useCallback((fieldKey) => {
    if (typeof window === 'undefined') return;
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) {
      alert('このブラウザは音声入力に対応していません');
      return;
    }
    try {
      dictationRef.current?.stop();
      quickCareDictationRef.current?.stop();
    } catch {
      // noop
    }
    const rec = new SR();
    rec.lang = 'ja-JP';
    rec.interimResults = false;
    rec.continuous = false;
    rec.onresult = (event) => {
      const text = String(event.results?.[0]?.[0]?.transcript ?? '').trim();
      if (!text) return;
      setEmergencyDraft((prev) => ({
        ...prev,
        [fieldKey]: prev[fieldKey] ? `${prev[fieldKey]}\n${text}` : text,
      }));
    };
    rec.onend = () => setDictatingField('');
    rec.onerror = () => setDictatingField('');
    dictationRef.current = rec;
    setDictatingField(fieldKey);
    rec.start();
  }, []);

  const stopQuickCareDictation = useCallback(() => {
    try {
      quickCareDictationRef.current?.stop();
    } catch {
      // noop
    } finally {
      setDictatingQuickKey('');
    }
  }, []);

  /** @param {'urineVolume'|'stoolVolume'|'stoolCharacter'|'mealStaple'|'mealSide'|'waterMl'} fieldKey */
  const startQuickCareDictation = useCallback((fieldKey) => {
    if (typeof window === 'undefined') return;
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) {
      alert('このブラウザは音声入力に対応していません');
      return;
    }
    try {
      dictationRef.current?.stop();
    } catch {
      // noop
    }
    try {
      quickCareDictationRef.current?.stop();
    } catch {
      // noop
    }
    const rec = new SR();
    rec.lang = 'ja-JP';
    rec.interimResults = false;
    rec.continuous = false;
    rec.onresult = (event) => {
      const text = String(event.results?.[0]?.[0]?.transcript ?? '').trim();
      if (!text) return;
      setQuickCareDetail((prev) => {
        if (fieldKey === 'urineVolume') {
          const next = text.replace(/\s+/g, ' ').trim();
          return {
            ...prev,
            urineVolume: prev.urineVolume ? `${prev.urineVolume} ${next}` : next,
          };
        }
        if (fieldKey === 'stoolVolume') {
          const v = parseVoiceToStoolVolume(text);
          if (!v) alert(`「${text}」から排便量（多・中・小）を判別できませんでした。`);
          return { ...prev, stoolVolume: v || prev.stoolVolume };
        }
        if (fieldKey === 'stoolCharacter') {
          const v = parseVoiceToStoolCharacter(text);
          if (!v) alert(`「${text}」から性状（普通便・硬便・軟便・水様便）を判別できませんでした。`);
          return { ...prev, stoolCharacter: v || prev.stoolCharacter };
        }
        if (fieldKey === 'mealStaple' || fieldKey === 'mealSide') {
          const v = parseVoiceToMealWari(text);
          if (!v) alert(`「${text}」から割合（10割〜0割）を判別できませんでした。`);
          return { ...prev, [fieldKey]: v || prev[fieldKey] };
        }
        if (fieldKey === 'waterMl') {
          const v = parseVoiceToWaterMl(text);
          if (!v) alert(`「${text}」から水分量（ml）の数字を判別できませんでした。`);
          return { ...prev, waterMl: v || prev.waterMl };
        }
        return prev;
      });
    };
    rec.onend = () => setDictatingQuickKey('');
    rec.onerror = () => setDictatingQuickKey('');
    quickCareDictationRef.current = rec;
    setDictatingQuickKey(fieldKey);
    rec.start();
  }, []);

  const stopDictation = useCallback(() => {
    try {
      dictationRef.current?.stop();
    } catch {
      // noop
    } finally {
      setDictatingField('');
    }
  }, []);

  const runEmergencySummary = useCallback(async () => {
    const res = filteredResidents.find((r) => String(r.id) === String(emergencyPickId));
    if (!res) return;
    setEmergencyBusy(true);
    const ev = Report.evaluateResidentMonitor(res);
    let advice = Report.fallbackRegulatoryAdvice(ev);
    if (GEMINI_KEY) {
      try {
        advice = await Report.fetchAiRegulatoryAdvice(GEMINI_KEY, ev, res);
      } catch {
        advice = Report.fallbackRegulatoryAdvice(ev);
      }
    }
    const contact = Report.getEmergencyContact(String(res.id));
    const html = Report.buildEmergencySummaryHtml(res, ev, advice, contact, emergencyDraft);
    Report.openPrintableSummary(html);
    setEmergencyBusy(false);
  }, [emergencyPickId, filteredResidents, emergencyDraft]);

  const downloadEmergencyHtml = useCallback(async () => {
    const res = filteredResidents.find((r) => String(r.id) === String(emergencyPickId));
    if (!res) return;
    setEmergencyBusy(true);
    const ev = Report.evaluateResidentMonitor(res);
    let advice = Report.fallbackRegulatoryAdvice(ev);
    if (GEMINI_KEY) {
      try {
        advice = await Report.fetchAiRegulatoryAdvice(GEMINI_KEY, ev, res);
      } catch {
        advice = Report.fallbackRegulatoryAdvice(ev);
      }
    }
    const contact = Report.getEmergencyContact(String(res.id));
    const html = Report.buildEmergencySummaryHtml(res, ev, advice, contact, emergencyDraft);
    Report.downloadSummaryHtml(`救急搬送サマリー_${String(res.name)}.html`, html);
    setEmergencyBusy(false);
  }, [emergencyPickId, filteredResidents, emergencyDraft]);

  const exportAudit = useCallback(() => {
    Report.downloadMonthlyAuditSheet(selectedSheetTitle, auditMonth, filteredResidents);
  }, [selectedSheetTitle, auditMonth, filteredResidents]);

  const exportAuditNarrative = useCallback(() => {
    Report.downloadPaidAuditNarrativeHtml(selectedSheetTitle, auditMonth, filteredResidents);
  }, [selectedSheetTitle, auditMonth, filteredResidents]);

  const importKaipokeVitalsCsv = useCallback(
    (file) => {
      if (!file) return;
      /** @param {string} [cell] */
      const parseCellToStoolIso = (cell) => {
        const t = String(cell ?? '').trim();
        if (!t) return null;
        const d = new Date(t.replace(/\//g, '-'));
        if (!Number.isNaN(d.getTime())) return d.toISOString();
        const m = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})[ T](\d{1,2}):(\d{2})/.exec(t);
        if (m)
          return new Date(
            Number(m[1]),
            Number(m[2]) - 1,
            Number(m[3]),
            Number(m[4]),
            Number(m[5])
          ).toISOString();
        const m2 = /^(\d{1,2})[\/\-](\d{1,2})[ T](\d{1,2}):(\d{2})/.exec(t);
        if (m2) {
          const y = new Date().getFullYear();
          return new Date(
            y,
            Number(m2[1]) - 1,
            Number(m2[2]),
            Number(m2[3]),
            Number(m2[4])
          ).toISOString();
        }
        const m3 = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/.exec(t);
        if (m3)
          return new Date(Number(m3[1]), Number(m3[2]) - 1, Number(m3[3]), 12, 0, 0).toISOString();
        const m4 = /^(\d{1,2})[\/\-](\d{1,2})$/.exec(t);
        if (m4) {
          const y = new Date().getFullYear();
          return new Date(y, Number(m4[1]) - 1, Number(m4[2]), 12, 0, 0).toISOString();
        }
        return null;
      };
      const reader = new FileReader();
      reader.onload = () => {
        const rows = parseCsv(String(reader.result ?? ''));
        if (rows.length < 2) {
          alert('CSV形式を認識できませんでした');
          return;
        }
        const headers = rows[0].map((x) => String(x ?? '').trim());
        const idx = {
          name: headers.findIndex((h) => /氏名|利用者名|入居者名/u.test(h)),
          temp: headers.findIndex((h) => {
            const hn = String(h);
            if (/血圧|目標/i.test(hn)) return false;
            return /体温/u.test(hn);
          }),
          bpU: headers.findIndex((h) => /血圧.*(上|高)|収縮|1回目：収縮/u.test(h)),
          bpL: headers.findIndex((h) => /血圧.*(下|低)|拡張|1回目：拡張/u.test(h)),
          pulse: headers.findIndex((h) => /脈拍|1回目：脈拍|pulse/i.test(h)),
          spo2: headers.findIndex((h) => /spo2|酸素|ｓｐｏ|ＳｐＯ2|SpO2/i.test(h)),
          weight: headers.findIndex((h) => /体重|weight/i.test(h) && !/血圧|目標/i.test(h)),
          stool: headers.findIndex((h) =>
            /排便.*(日時|時刻|時間)|最終排便|排便日時|排便（日時）|排便記録/u.test(h)
          ),
          urine: headers.findIndex((h) => /排尿.*(記録|内容|コメント)|排尿$/u.test(h)),
        };
        if (idx.name < 0) {
          alert('氏名列が見つかりません');
          return;
        }
        let applied = 0;
        for (let r = 1; r < rows.length; r++) {
          const row = rows[r];
          const nm = residentNameWithoutSama(String(row[idx.name] ?? ''));
          if (!nm) continue;
          const hit = filteredResidents.find(
            (res) => residentNameWithoutSama(String(res.name ?? '')) === nm
          );
          if (!hit) continue;
          const patch = {
            temp: idx.temp >= 0 ? String(row[idx.temp] ?? '').trim() : '',
            bpUpper: idx.bpU >= 0 ? String(row[idx.bpU] ?? '').trim() : '',
            bpLower: idx.bpL >= 0 ? String(row[idx.bpL] ?? '').trim() : '',
            pulse: idx.pulse >= 0 ? String(row[idx.pulse] ?? '').trim() : '',
            spo2: idx.spo2 >= 0 ? String(row[idx.spo2] ?? '').trim() : '',
            weight: idx.weight >= 0 ? String(row[idx.weight] ?? '').trim() : '',
            urineNote: idx.urine >= 0 ? String(row[idx.urine] ?? '').trim() : '',
          };
          Report.setResidentVitalSnapshot(String(hit.id), patch);
          Report.logVitalSnapshot(
            String(hit.id),
            String(hit.name ?? ''),
            String(hit.sourceSheetTitle ?? hit.facility ?? selectedSheetTitle),
            patch
          );
          if (idx.stool >= 0) {
            const iso = parseCellToStoolIso(String(row[idx.stool] ?? ''));
            if (iso) Report.setLastStoolIso(String(hit.id), iso);
          }
          applied += 1;
        }
        if (applied > 0) setTick((n) => n + 1);
        alert(applied > 0 ? `${applied}名に反映しました` : '一致する利用者がありませんでした');
      };
      reader.onerror = () => alert('CSV読み込みに失敗しました');
      reader.readAsText(file);
    },
    [filteredResidents, selectedSheetTitle]
  );

  const hdrBtn =
    'flex items-center gap-1 rounded-lg border-2 px-2 py-1.5 text-[11px] font-bold shadow-sm sm:gap-1.5 sm:px-2.5 sm:text-xs 2xl:px-3 2xl:text-sm';

  return (
    <div className="flex min-h-[100dvh] min-w-0 flex-col gap-2 bg-slate-300 p-2 pb-4 font-sans text-slate-900 sm:gap-2 sm:p-2 sm:pb-4">
      <header className="flex shrink-0 flex-wrap items-center justify-between gap-x-2 gap-y-1 rounded-2xl border border-slate-700 bg-slate-900 px-2 py-1.5 text-white shadow-lg sm:px-3 sm:py-2">
        <div className="flex min-w-0 flex-1 flex-wrap items-center gap-2">
          <button type="button" onClick={onBack} className="shrink-0 rounded-xl p-1.5 hover:bg-white/10" aria-label="戻る">
            <ChevronLeft className="h-6 w-6 text-slate-300 sm:h-7 sm:w-7" />
          </button>
          <div className="flex min-w-0 items-center gap-2 text-slate-300">
            <Monitor className="h-5 w-5 shrink-0 text-cyan-400 sm:h-6 sm:w-6" />
            <div className="min-w-0">
              <h1 className="truncate text-base font-black tracking-tight text-white sm:text-lg 2xl:text-2xl">
                {selectedDef?.tabLabel ?? '施設'}
              </h1>
              <p className="truncate text-[10px] text-slate-400 sm:text-xs">
                名簿・異常検知・周知
                <span className="text-slate-500">
                  {lastUpdated ? ` ・ 同期 ${lastUpdated.toLocaleTimeString('ja-JP')}` : ' ・ 未同期'}
                </span>
                {fetchSourceMeta ? (
                  <span className="text-cyan-400">
                    {` ・ ${fetchSourceMeta.source}${fetchSourceMeta.mode ? `(${fetchSourceMeta.mode})` : ''}`}
                  </span>
                ) : null}
              </p>
            </div>
          </div>
        </div>
        <div className="flex w-full flex-wrap items-center justify-end gap-1 sm:w-auto 2xl:flex-nowrap 2xl:gap-1.5">
          <button
            type="button"
            onClick={() => {
              setEmergencyPickId(String(filteredResidents[0]?.id ?? ''));
              setEmergencyOpen(true);
            }}
            className={`${hdrBtn} border-rose-500 bg-rose-600 text-white hover:bg-rose-500`}
          >
            <Ambulance className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
            救急
          </button>
          <button
            type="button"
            onClick={() => setAccidentReportOpen(true)}
            className={`${hdrBtn} border-slate-500 bg-slate-700 text-white hover:bg-slate-600`}
          >
            <ClipboardList className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
            事故
          </button>
          <button
            type="button"
            onClick={() => setAccidentMonthlyOpen(true)}
            className={`${hdrBtn} border-indigo-500 bg-indigo-700 text-white hover:bg-indigo-600`}
          >
            <BarChart3 className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
            月次
          </button>
          <button
            type="button"
            onClick={() => setNearMissOpen(true)}
            className={`${hdrBtn} border-teal-500 bg-teal-700 text-white hover:bg-teal-600`}
          >
            <FileWarning className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
            ヒヤリ
          </button>
          <button
            type="button"
            onClick={() => setNearMissMonthlyOpen(true)}
            className={`${hdrBtn} border-teal-600 bg-teal-900 text-white hover:bg-teal-800`}
          >
            <BarChart3 className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
            ヒヤリ月次
          </button>
          <div className="rounded-lg bg-slate-800 px-2 py-1 text-center">
            <p className="text-[9px] text-slate-400 leading-none">時刻</p>
            <p className="text-xs font-bold tabular-nums text-cyan-300 sm:text-sm 2xl:text-base">{nowLabel}</p>
          </div>
          <button
            type="button"
            disabled={loading || refreshing}
            onClick={() => load(true)}
            className={`${hdrBtn} border-slate-600 bg-slate-800 text-white hover:bg-slate-700 disabled:opacity-50`}
          >
            {refreshing ? <Loader2 className="h-4 w-4 animate-spin" /> : <RefreshCw className="h-4 w-4 shrink-0" />}
            更新
          </button>
          <input
            ref={kaipokeCsvInputRef}
            type="file"
            accept=".csv,text/csv"
            className="hidden"
            onChange={(e) => importKaipokeVitalsCsv(e.target.files?.[0] ?? null)}
          />
          <button
            type="button"
            onClick={() => kaipokeCsvInputRef.current?.click()}
            className={`${hdrBtn} border-cyan-700 bg-cyan-700 text-white hover:bg-cyan-600`}
            title="カイポケ等のCSVを一括反映（氏名・バイタル・SpO2・体重・排便日時・排尿メモの列を自動検出）"
          >
            <Upload className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
            CSV取込
          </button>
          <div className="flex flex-wrap items-center gap-0.5 rounded-lg border border-slate-600 bg-slate-800 px-1.5 py-0.5">
            <input
              type="month"
              value={auditMonth}
              onChange={(e) => setAuditMonth(e.target.value)}
              className="max-w-[7.5rem] rounded bg-slate-700 px-1 py-0.5 text-[10px] text-white sm:max-w-[9rem] sm:text-xs"
            />
            <button
              type="button"
              onClick={exportAudit}
              title="件数・最終日時の集計（Excel向け）"
              className="flex items-center gap-0.5 rounded-md bg-amber-600 px-1.5 py-1 text-[10px] font-bold text-white hover:bg-amber-500 sm:text-xs"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              CSV
            </button>
            <button
              type="button"
              onClick={exportAuditNarrative}
              title="有料監査用: 巡視の間隔・食事の割合・排泄の間隔・一言要約＋請求用食事・経管（HTML・印刷可）"
              className="flex items-center gap-0.5 rounded-md bg-teal-700 px-1.5 py-1 text-[10px] font-bold text-white hover:bg-teal-600 sm:text-xs"
            >
              <FileText className="h-3.5 w-3.5" />
              監査HTML
            </button>
          </div>
          {typeof onOpenNotionNewResidents === 'function' ? (
            <button
              type="button"
              onClick={onOpenNotionNewResidents}
              className={`${hdrBtn} border-violet-700 bg-violet-600 text-white hover:bg-violet-500`}
              title="営業が Notion に登録した新規入居一覧"
            >
              <Baby className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
              新規入居
            </button>
          ) : null}
          <button
            type="button"
            onClick={onOpenMonthlyReport}
            className={`${hdrBtn} border-emerald-700 bg-emerald-600 text-white hover:bg-emerald-500`}
          >
            <MessageSquarePlus className="h-4 w-4 shrink-0 sm:h-5 sm:w-5" />
            報告
          </button>
        </div>
      </header>

      {error && (
        <div className="shrink-0 rounded-2xl border-2 border-rose-400 bg-rose-50 px-4 py-3 text-base text-rose-900">
          <p className="font-bold">{error}</p>
          <p className="mt-2 text-sm font-normal text-rose-800">
            API キーあり: Sheets API。キーなし: 公開CSV（npm run dev 必須）。シートは閲覧可能な共有にしてください。
          </p>
        </div>
      )}

      <div className="flex min-w-0 flex-col gap-2 lg:flex-row lg:items-start lg:gap-3">
        <div className="flex min-w-0 flex-1 flex-col gap-2 lg:gap-3">
          <section className="order-2 flex flex-col gap-2">
            <div className="grid grid-cols-1 gap-2 md:grid-cols-2 xl:grid-cols-3 xl:gap-3 xl:items-stretch">
              <div className="flex min-h-0 min-w-0 flex-col rounded-2xl border-2 border-rose-300 bg-gradient-to-br from-rose-50 to-amber-50 p-2.5 shadow-md sm:p-3">
                <div className="mb-1 flex flex-wrap items-center gap-2 text-rose-900">
                  <Stethoscope className="h-6 w-6 shrink-0" />
                  <Megaphone className="h-5 w-5 shrink-0" />
                  <h2 className="text-base font-black sm:text-lg 2xl:text-xl">本日の重要周知（看護指示）</h2>
                </div>
                <div className="space-y-2 pr-1">
                  {nursingList.length === 0 ? (
                    <p className="text-base font-bold text-rose-700">看護からの処置・指示は未登録です。下欄に入力して掲示してください。</p>
                  ) : (
                    nursingList.map((d, i) => (
                      <div
                        key={`${d.ts}-${i}`}
                        className="rounded-xl border-2 border-rose-400 bg-white px-4 py-3 text-lg font-bold leading-snug text-rose-950 shadow-sm sm:text-xl"
                      >
                        <div className="flex items-start justify-between gap-2">
                          <div className="min-w-0">
                            <span className="mr-2 text-sm font-bold text-rose-500">{d.by}</span>
                            {d.text}
                            {(d.startDate || d.endDate) ? (
                              <div className="mt-1 text-[11px] font-bold text-rose-700">
                                表示期間: {d.startDate || '今日'} 〜 {d.endDate || '未設定'}
                              </div>
                            ) : null}
                          </div>
                          <button
                            type="button"
                            onClick={() => removeNursing(d)}
                            className="shrink-0 rounded border border-rose-300 bg-rose-50 px-2 py-1 text-xs font-black text-rose-700"
                          >
                            削除
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>
                <p className="mt-2 shrink-0 border-t border-rose-200 pt-2 text-sm font-bold text-amber-900 2xl:text-base">
                  {board.notice}
                </p>
                <div className="mt-2 flex shrink-0 flex-col gap-2 sm:flex-row">
                  <input
                    type="text"
                    value={nursingDraft}
                    onChange={(e) => setNursingDraft(e.target.value)}
                    placeholder="例: 下剤投与につき排便確認／褥瘡あり：右側臥位注意"
                    className="min-w-0 flex-1 rounded-xl border-2 border-rose-300 px-3 py-2.5 text-base font-bold text-slate-900 outline-none focus:ring-2 focus:ring-rose-400"
                  />
                  <button
                    type="button"
                    onClick={registerNursing}
                    className="shrink-0 rounded-xl bg-rose-600 px-5 py-2.5 text-base font-black text-white shadow-md hover:bg-rose-500"
                  >
                    看護指示を掲示
                  </button>
                </div>
                <div className="mt-2 grid grid-cols-1 gap-2 sm:grid-cols-2">
                  <input
                    type="date"
                    value={nursingStartDate}
                    onChange={(e) => setNursingStartDate(e.target.value)}
                    className="rounded-lg border border-rose-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-900"
                  />
                  <input
                    type="date"
                    value={nursingEndDate}
                    onChange={(e) => setNursingEndDate(e.target.value)}
                    className="rounded-lg border border-rose-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-900"
                    placeholder="終了日（任意）"
                  />
                </div>
              </div>
              <div className="flex min-h-0 min-w-0 flex-col rounded-2xl border-2 border-indigo-200/90 bg-indigo-50/95 p-2.5 shadow-md sm:p-3">
                <div className="mb-1 flex items-center gap-2 text-indigo-900">
                  <ClipboardList className="h-5 w-5 shrink-0 sm:h-6 sm:w-6" />
                  <h2 className="text-base font-bold sm:text-lg">申し送り</h2>
                </div>
                <p className="text-sm leading-relaxed text-indigo-950 sm:text-base 2xl:text-lg">
                  {board.handover}
                </p>
              </div>
              <div className="flex min-h-0 min-w-0 flex-col rounded-2xl border-2 border-teal-300/90 bg-teal-50/95 p-2.5 shadow-md sm:p-3 md:col-span-2 xl:col-span-1">
                <div className="mb-2 flex items-center gap-2 text-teal-900">
                  <CalendarClock className="h-5 w-5 shrink-0 sm:h-6 sm:w-6" />
                  <h2 className="text-base font-bold sm:text-lg">本日の予定</h2>
                </div>
                <ul className="space-y-1.5 pr-1 text-sm sm:text-base">
                  {(todayPlans.length > 0 ? todayPlans : board.schedule).map((item, i) => (
                    <li
                      key={String(item.id ?? i)}
                      className="flex gap-3 rounded-xl border border-teal-200/80 bg-white/90 px-3 py-2.5 shadow-sm"
                    >
                      <span className="w-16 shrink-0 font-mono font-bold text-teal-700">{item.time || '—'}</span>
                      <span className="font-bold leading-snug text-slate-800">{item.title || '予定'}</span>
                    </li>
                  ))}
                </ul>
              </div>
            </div>
            <div className="flex min-w-0 flex-col rounded-2xl border-2 border-teal-300/90 bg-teal-50/95 p-3 shadow-md">
                <div className="mb-1 flex flex-wrap items-center gap-2 text-teal-900">
                  <CalendarDays className="h-5 w-5 shrink-0" />
                  <h3 className="text-base font-black">今週の予定（7日分・一覧）</h3>
                </div>
                <p className="mb-2 text-[11px] font-bold leading-snug text-teal-900/85">
                  外出・外泊・受診ごとに、くすりや持ち物の準備がしやすいよう、今日から7日間を日付別に常時表示します。予定が無い日は「予定なし」です。下のフォームから追記できます（名簿連携は今後の拡張用）。
                </p>
                {weeklyPlanDays.length === 0 ? (
                  <p className="mb-2 rounded-lg border border-dashed border-teal-300 bg-white/70 px-2 py-3 text-center text-sm font-bold text-slate-600">
                    施設を選ぶと、ここに7日分の枠が表示されます。
                  </p>
                ) : (
                  <div className="carelink-resident-grid-scroll mb-2 flex gap-2 overflow-x-auto overflow-y-visible pb-2 pl-0.5 pr-1 pt-0.5">
                    {weeklyPlanDays.map((day) => (
                      <div
                        key={day.date}
                        className={`flex w-[min(100%,10.5rem)] shrink-0 flex-col rounded-xl border-2 shadow-sm ${
                          day.isToday
                            ? 'border-teal-600 bg-white ring-2 ring-teal-400/50'
                            : 'border-teal-200/90 bg-white/95'
                        }`}
                      >
                        <div className="shrink-0 border-b border-teal-100 bg-teal-600/10 px-2 py-1.5 text-center">
                          <div className="font-mono text-[11px] font-bold text-teal-800">{day.date}</div>
                          <div className="text-xs font-black text-teal-950 sm:text-sm">
                            {day.weekdayShort}曜
                            {day.isToday ? (
                              <span className="ml-1 inline-block rounded bg-teal-600 px-1 py-0.5 align-middle text-[9px] text-white">
                                今日
                              </span>
                            ) : null}
                          </div>
                        </div>
                        <ul className="min-h-[6.5rem] space-y-1.5 p-2 text-[11px] sm:min-h-[7.5rem] sm:text-xs">
                          {day.plans.length === 0 ? (
                            <li className="rounded-lg border border-dashed border-slate-200 bg-slate-50/90 px-2 py-3 text-center font-bold leading-snug text-slate-500">
                              予定なし
                            </li>
                          ) : (
                            day.plans.map((p) => (
                              <li
                                key={String(p.id)}
                                className="rounded-lg border border-teal-200 bg-teal-50/90 px-2 py-1.5 shadow-sm"
                              >
                                <div className="font-mono text-[11px] font-black text-teal-900">{p.time}</div>
                                <span className="mt-0.5 inline-block rounded bg-teal-600/90 px-1 py-0.5 text-[9px] font-black text-white">
                                  {p.type}
                                </span>
                                <div className="mt-1 font-bold leading-snug text-slate-900">{p.title}</div>
                                <button
                                  type="button"
                                  onClick={() => removeWeeklyPlan(p.id)}
                                  className="mt-1 rounded border border-teal-300 bg-white px-1.5 py-0.5 text-[10px] font-black text-teal-800"
                                >
                                  削除
                                </button>
                              </li>
                            ))
                          )}
                        </ul>
                      </div>
                    ))}
                  </div>
                )}
                <p className="mb-2 text-[10px] font-bold text-slate-500">予定を追加・更新する（任意）</p>
                <div className="mb-2 grid grid-cols-1 gap-2 sm:grid-cols-2">
                  <input
                    type="date"
                    value={planDraftDate}
                    onChange={(e) => setPlanDraftDate(e.target.value)}
                    className="rounded-lg border border-teal-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-900"
                  />
                  <input
                    type="time"
                    value={planDraftTime}
                    onChange={(e) => setPlanDraftTime(e.target.value)}
                    className="rounded-lg border border-teal-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-900"
                  />
                  <select
                    value={planDraftType}
                    onChange={(e) => setPlanDraftType(e.target.value)}
                    className="rounded-lg border border-teal-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-900"
                  >
                    <option value="外出">外出</option>
                    <option value="外泊">外泊</option>
                    <option value="受診">受診</option>
                    <option value="往診">往診</option>
                    <option value="面会">面会</option>
                    <option value="その他">その他</option>
                  </select>
                  <input
                    type="text"
                    value={planDraftTitle}
                    onChange={(e) => setPlanDraftTitle(e.target.value)}
                    placeholder="例: 〇〇様 14:00 内科（薬手帳・頓服）"
                    className="rounded-lg border border-teal-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-900"
                  />
                </div>
                <button
                  type="button"
                  onClick={registerWeeklyPlan}
                  className="mb-2 w-full rounded-lg bg-teal-600 px-3 py-2 text-sm font-black text-white hover:bg-teal-500"
                >
                  上記の日付に予定を1件追加
                </button>
            </div>
          </section>

          <section className="order-1 flex min-w-0 flex-col rounded-2xl border-2 border-slate-400 bg-white shadow-inner">
            <div className="shrink-0 border-b border-slate-200 bg-slate-100/90 px-2 py-1.5 sm:px-3 sm:py-2">
              <h2 className="text-base font-bold text-slate-900 sm:text-lg 2xl:text-xl">入居者一覧・異常監視</h2>
              {residentFilterBanner ? (
                <p className="mt-2 rounded-xl border-2 border-amber-500 bg-amber-50 px-3 py-2 text-xs font-bold leading-snug text-amber-950 sm:text-sm">
                  {residentFilterBanner}
                </p>
              ) : null}
              <p className="mt-0.5 text-sm font-bold text-blue-700 sm:text-base">{headerResidentCountSubtitle}</p>
            </div>
            <div className="shrink-0 border-b border-slate-200 bg-slate-100/90 px-2 py-1.5 sm:px-3 sm:py-2">
              <div className="flex flex-wrap items-end justify-between gap-2">
                <div className="min-w-0 flex-1">
                  <div className="mt-1 max-w-full rounded-2xl border-2 border-sky-500 bg-gradient-to-br from-sky-50 via-white to-indigo-50 px-3 py-2.5 shadow-md sm:px-4 sm:py-3">
                    <div className="flex flex-wrap items-center gap-2">
                      <Stethoscope className="h-5 w-5 shrink-0 text-sky-700 sm:h-6 sm:w-6" aria-hidden />
                      <span className="text-sm font-black tracking-tight text-sky-950 sm:text-base">
                        医療保険カウント（名簿）
                      </span>
                    </div>
                    <div
                      className={`mt-3 grid grid-cols-1 gap-2 ${nursingOfficeUi ? 'sm:grid-cols-3' : 'sm:grid-cols-2'}`}
                    >
                      <div className="rounded-xl border-2 border-rose-300 bg-rose-50/95 px-3 py-3 text-center shadow-sm">
                        <p className="text-xs font-black tracking-wide text-rose-800 sm:text-sm">平均介護度</p>
                        <p className="mt-1 font-mono text-4xl font-black leading-none text-rose-950 sm:text-5xl">
                          {facilityCareStats.averageCareLevel != null
                            ? facilityCareStats.averageCareLevel.toFixed(2)
                            : '—'}
                        </p>
                      </div>
                      <div className="rounded-xl border-2 border-cyan-500 bg-cyan-50 px-3 py-3 text-center shadow-sm">
                        <p className="text-xs font-black tracking-wide text-cyan-900 sm:text-sm">医療保険対象者</p>
                        <p className="mt-1 font-mono text-4xl font-black leading-none text-slate-950 sm:text-5xl">
                          {facilityCareStats.medicalTargetCount}
                        </p>
                        <p className="mt-1 text-xs font-bold text-cyan-900">名</p>
                      </div>
                      {nursingOfficeUi ? (
                        <div
                          className={`rounded-xl border-2 px-3 py-3 text-center shadow-sm ${
                            visitNursingStats.warn
                              ? 'border-amber-600 bg-amber-50'
                              : 'border-teal-500 bg-teal-50'
                          }`}
                        >
                          <p className="flex items-center justify-center gap-1 text-xs font-black tracking-wide text-teal-950 sm:text-sm">
                            <Home className="h-4 w-4 shrink-0" aria-hidden />
                            訪問看護・特別指示
                          </p>
                          <p className="mt-1 font-mono text-4xl font-black leading-none text-slate-950 sm:text-5xl">
                            {visitNursingStats.count}
                          </p>
                          <p className="mt-1 text-xs font-bold text-teal-900">名（名簿＋手動）</p>
                          {visitNursingStats.warn ? (
                            <p className="mt-2 rounded-lg border border-amber-500 bg-amber-100 px-2 py-1.5 text-[10px] font-black leading-snug text-amber-950">
                              {visitNursingStats.threshold}名以上：減算・体制の管理が必要です（算定要件は最新告示で確認）
                            </p>
                          ) : null}
                        </div>
                      ) : null}
                    </div>
                  </div>
                  <p className="mt-2 text-[10px] font-bold text-slate-500 sm:text-xs">
                    赤: バイタル・排便 / 黄: 巡視遅延 / 紫枠: 減算・監査の確認候補
                  </p>
                  <p className="mt-1 text-[10px] font-bold text-slate-600 sm:text-xs">
                    請求用: 当月食事は名簿の食事列＋この端末の記録を合算（{billingYearMonth}）。経管は名簿の「経管栄養」列と、生活記録保存時の経管実施ログを集計。
                  </p>
                </div>
                <div className="flex shrink-0 gap-2 text-[11px] font-bold sm:gap-3 sm:text-sm">
                  <span className="flex items-center gap-1 text-red-600">
                    <AlertTriangle className="h-5 w-5" /> バイタル・排便
                  </span>
                  <span className="flex items-center gap-1 text-amber-600">
                    <Wind className="h-5 w-5" /> 巡視
                  </span>
                </div>
              </div>
            </div>

            <div className="flex min-w-0 flex-col p-2 sm:p-3">
              {loading ? (
                <div className="flex h-40 flex-col items-center justify-center gap-4 text-slate-500">
                  <Loader2 className="h-12 w-12 animate-spin text-blue-600" />
                  <span className="text-xl font-bold">読み込み中…</span>
                </div>
              ) : filteredResidents.length === 0 ? (
                <div className="flex min-h-48 flex-col items-center justify-center rounded-2xl border-2 border-dashed border-slate-300 bg-slate-50 p-6 text-center text-lg text-slate-600">
                  <p className="font-bold">表示できる入居者がいません。</p>
                  {allResidents.length > 0 ? (
                    <p className="mt-3 max-w-lg text-base font-normal leading-relaxed text-slate-600">
                      名簿は {allResidents.length} 名読み込めていますが、いま選んでいる施設タブ（
                      {selectedDef?.tabLabel ?? '—'}）と一致する行がありません。ポータルから別の施設を選び直す・
                      <strong className="font-bold text-slate-800">更新</strong>
                      を押す・スプレッドシートのタブ名をアプリ設定（
                      <code className="rounded bg-slate-200 px-1 text-sm">carelinkFacilities.js</code> の sheetTitle）と揃える・「施設」列の表記を確認してください。
                    </p>
                  ) : (
                    <p className="mt-3 max-w-lg text-base font-normal leading-relaxed text-slate-600">
                      名簿が0件です。1行目に「氏名」列があるか、APIキー・スプレッドシートIDを確認してください。
                    </p>
                  )}
                </div>
              ) : (
                <>
                  <div className="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-xl border border-slate-200 bg-slate-50 px-2 py-2 sm:px-3">
                    <p className="max-w-[16rem] text-[11px] font-bold leading-snug text-slate-600 sm:text-xs">
                      カードは1人ずつ開く方式。一覧表は横スクロールで連続入力し、Tabキーで移動できます。
                    </p>
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setResidentInputView('cards')}
                        className={`rounded-xl border-2 px-3 py-2 text-xs font-black sm:text-sm ${
                          residentInputView === 'cards'
                            ? 'border-blue-600 bg-blue-600 text-white'
                            : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-100'
                        }`}
                      >
                        カード表示
                      </button>
                      <button
                        type="button"
                        onClick={switchToTableInput}
                        className={`inline-flex items-center gap-1.5 rounded-xl border-2 px-3 py-2 text-xs font-black sm:text-sm ${
                          residentInputView === 'table'
                            ? 'border-emerald-600 bg-emerald-600 text-white'
                            : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-100'
                        }`}
                      >
                        <Table2 className="h-4 w-4 shrink-0" aria-hidden />
                        一覧表で入力
                      </button>
                    </div>
                  </div>
                  {residentInputView === 'table' ? (
                    <ResidentBulkInputTable
                      filteredResidents={filteredResidents}
                      bulkDraft={bulkDraft}
                      bulkGlobalMealSlot={bulkGlobalMealSlot}
                      onBulkGlobalMealSlotChange={onBulkGlobalMealSlotChange}
                      residentNameWithoutSama={residentNameWithoutSama}
                      patchBulkRow={patchBulkRow}
                      setBulkPatrolForAllVisible={setBulkPatrolForAllVisible}
                      bulkRowHasInput={bulkRowHasInput}
                      saveBulkRow={saveBulkRow}
                      saveBulkAllWithInput={saveBulkAllWithInput}
                    />
                  ) : (
                    <div
                      className="grid min-w-0 max-w-full gap-2 pb-4 pl-0.5 pr-1 sm:gap-3 sm:pr-2 sm:pb-6"
                      style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(14rem, 1fr))' }}
                    >
                  {filteredResidents.map((res) => {
                    void tick;
                    const ev = Report.evaluateResidentMonitor(res);
                    const ded = Report.evaluateReimbursementDeductionAlerts(res, ev);
                    const critical = ev.level === 'critical';
                    const warn = ev.level === 'warn';
                    const adviceShort = critical ? Report.fallbackRegulatoryAdvice(ev) : '';
                    const cal = calOpenId === String(res.id) ? Report.getWeekCalendarBuckets(String(res.id)) : [];
                    const careCanonical = normalizeCareLevelLabel(String(res.careLevelLabel ?? '').trim());
                    const careDisplay = formatCareLevelForDisplay(res.careLevelLabel);
                    const careBadgeClass = (() => {
                      if (!careCanonical) return '';
                      const c = careCanonical.replace(/\s/g, '');
                      if (/^要介護[45]$/.test(c)) return 'bg-rose-700 text-white border-rose-900';
                      if (/^要介護[123]$/.test(c)) return 'bg-amber-600 text-white border-amber-800';
                      if (/^要支援/.test(c)) return 'bg-sky-600 text-white border-sky-800';
                      if (/自立/.test(careCanonical)) return 'bg-emerald-600 text-white border-emerald-800';
                      return 'bg-slate-700 text-white border-slate-900';
                    })();
                    const bill = residentBillingById.get(String(res.id)) ?? { mealLogged: 0, enteralLogged: 0 };
                    const sheetMeal = Number(res.mealCountThisMonth) || 0;
                    const mealTotal = sheetMeal + bill.mealLogged;
                    return (
                      <div
                        key={String(res.id)}
                        className={`flex min-w-0 flex-col rounded-2xl border-2 p-4 text-left shadow-sm ${
                          critical
                            ? 'animate-carelink-blink border-red-800 bg-red-600 text-white'
                            : warn
                              ? 'border-amber-500 bg-amber-100 text-slate-900'
                              : 'border-slate-200 bg-white text-slate-900'
                        }`}
                      >
                        <button
                          type="button"
                          onClick={() => openQuickPanel(res)}
                          className="w-full text-left"
                        >
                          <div className="mb-2 flex items-start justify-between gap-2">
                            <div className="min-w-0 flex-1">
                              <div className="line-clamp-2 text-2xl font-black leading-tight">
                                {residentNameWithoutSama(res.name)}
                                <span className="text-xl"> 様</span>
                              </div>
                              <div
                                className={`mt-2 rounded-xl border-2 px-3 py-2 ${
                                  critical ? 'border-white/50 bg-black/25' : 'border-slate-300 bg-slate-50'
                                }`}
                              >
                                <div
                                  className={`text-[10px] font-black uppercase tracking-wide ${
                                    critical ? 'text-red-100' : 'text-slate-500'
                                  }`}
                                >
                                  介護度
                                </div>
                                <div className="mt-1">
                                  {careDisplay ? (
                                    <span
                                      className={`inline-block rounded-lg border-2 px-2.5 py-1 text-sm font-black sm:text-base ${careBadgeClass}`}
                                    >
                                      {careDisplay}
                                    </span>
                                  ) : (
                                    <span
                                      className={`text-sm font-black ${critical ? 'text-red-100' : 'text-slate-400'}`}
                                    >
                                      名簿未登録
                                    </span>
                                  )}
                                </div>
                              </div>
                            </div>
                            <div className="flex shrink-0 flex-col items-end gap-1">
                              <span
                                className={`rounded-md px-2 py-0.5 text-sm font-bold uppercase tracking-wide ${
                                  critical ? 'text-red-100' : 'text-slate-500'
                                }`}
                              >
                                {String(res.room)}
                              </span>
                              <div className="flex shrink-0 gap-1">
                                {critical && <AlertCircle className="h-6 w-6 text-white" />}
                                {warn && !critical && <Clock className="h-5 w-5 animate-pulse text-amber-700" />}
                              </div>
                            </div>
                          </div>
                          <div
                            className={`line-clamp-2 text-sm font-bold sm:text-base ${
                              critical ? 'text-red-100' : 'text-slate-600'
                            }`}
                          >
                            {String(res.condition ?? '—')}
                          </div>
                          {String(res.insuranceLabel ?? '').trim() ? (
                            <div
                              className={`mt-1 line-clamp-2 text-[11px] font-bold sm:text-xs ${
                                critical ? 'text-red-100/90' : 'text-sky-800'
                              }`}
                            >
                              保険: {String(res.insuranceLabel)}
                            </div>
                          ) : null}
                          {nursingOfficeUi &&
                            (String(res.insuranceCategory ?? '') === '医療保険特指示' ||
                              /特指示|特別指示/u.test(String(res.insuranceLabel ?? ''))) && (
                            <div
                              className={`mt-1.5 rounded-lg border-2 px-2 py-1.5 text-[10px] font-black leading-snug sm:text-[11px] ${
                                critical
                                  ? 'border-amber-200 bg-black/25 text-amber-100'
                                  : 'border-amber-500 bg-amber-50 text-amber-950'
                              }`}
                            >
                              医療保険 特指示 → 名簿の医療保険列で内容を確認・更新してください（ポータルからは編集しません）
                            </div>
                          )}
                          {nursingOfficeUi && Report.residentHasVisitNursingSpecial(res) && (
                            <div
                              className={`mt-1.5 rounded-lg border-2 px-2 py-1.5 text-[10px] font-black leading-snug sm:text-[11px] ${
                                critical
                                  ? 'border-teal-200 bg-black/25 text-teal-100'
                                  : 'border-teal-600 bg-teal-50 text-teal-950'
                              }`}
                            >
                              <span className="flex flex-wrap items-center gap-1">
                                <Home className="h-3.5 w-3.5 shrink-0" aria-hidden />
                                訪問看護・特別指示（管理対象）
                                {Report.sheetSuggestsVisitNursingSpecial(res.insuranceLabel) &&
                                !Report.visitNursingManualRegistrationActive(String(res.id)) ? (
                                  <span className="font-bold opacity-90">・名簿検出</span>
                                ) : null}
                              </span>
                              {(() => {
                                const vn = Report.getVisitNursingSpecial(String(res.id));
                                return vn.periodStart || vn.periodEnd ? (
                                  <span className="mt-1 block font-mono tabular-nums text-[10px] font-bold opacity-90">
                                    手動登録の期間: {vn.periodStart || '—'} 〜 {vn.periodEnd || '—'}
                                  </span>
                                ) : null;
                              })()}
                            </div>
                          )}
                          {ded.hasAlert && (
                            <div
                              className={`mt-2 rounded-xl border-2 px-2 py-1.5 text-[10px] font-bold leading-snug ${
                                critical
                                  ? 'border-white/80 bg-black/25 text-white'
                                  : 'border-violet-600 bg-violet-100 text-violet-950'
                              }`}
                            >
                              <span
                                className={`flex items-center gap-1 font-black ${
                                  critical ? 'text-white' : 'text-violet-900'
                                }`}
                              >
                                <FileWarning className="h-3.5 w-3.5 shrink-0" />
                                減算・監査 要確認
                              </span>
                              <ul
                                className={`mt-1 list-inside list-disc ${
                                  critical ? 'text-red-50' : 'text-violet-900'
                                }`}
                              >
                                {ded.lines.map((line, i) => (
                                  <li key={i}>{line}</li>
                                ))}
                              </ul>
                            </div>
                          )}
                          {(ev.vitalBad || ev.stoolBad || ev.urineBad) && (
                            <ul className={`mt-2 list-inside list-disc text-xs font-bold ${critical ? 'text-white' : ''}`}>
                              {ev.vitalFlags.map((f) => (
                                <li key={f.code}>{f.label}</li>
                              ))}
                              {ev.stoolBad && (
                                <li>
                                  排便 {ev.stoolHours != null ? `${Math.round(ev.stoolHours)}h` : '—'} 未記録相当（72h超）
                                </li>
                              )}
                              {ev.urineBad && (
                                <li>
                                  排尿記録・トイレ誘導{' '}
                                  {ev.urineHours != null ? `${Math.round(ev.urineHours)}h` : '—'} 間隔（{Report.VITAL_THRESHOLDS.urineHoursMax}h超）
                                </li>
                              )}
                            </ul>
                          )}
                          {critical && adviceShort && (
                            <div className="mt-2 flex items-start gap-1 rounded-lg bg-black/20 px-2 py-1.5 text-[10px] font-bold leading-snug text-white">
                              <Sparkles className="mt-0.5 h-3.5 w-3.5 shrink-0" />
                              <span className="line-clamp-4">{adviceShort}</span>
                            </div>
                          )}
                          <div
                            className={`mt-2 space-y-1 border-t pt-2 text-[11px] font-bold sm:text-xs ${
                              critical ? 'border-red-400 text-red-100' : 'border-slate-100 text-slate-600'
                            }`}
                          >
                            <div className="flex flex-wrap items-center justify-between gap-x-2 gap-y-1">
                              <span className={critical ? '' : 'text-slate-900'}>
                                当月食事{' '}
                                <span className="font-mono text-base tabular-nums sm:text-lg">{mealTotal}</span> 回
                                {(sheetMeal > 0 || bill.mealLogged > 0) && (
                                  <span className="ml-1 block text-[10px] font-bold opacity-85 sm:ml-2 sm:inline">
                                    （名簿{sheetMeal}＋記録{bill.mealLogged}）
                                  </span>
                                )}
                              </span>
                              <span
                                className={
                                  warn || critical ? 'rounded-full bg-black/30 px-2 py-0.5' : 'text-blue-600'
                                }
                              >
                                巡視 {String(res.lastPatrol ?? '—')}
                              </span>
                            </div>
                            <div
                              className={`flex flex-wrap gap-x-3 gap-y-0.5 ${critical ? 'text-red-50' : 'text-slate-800'}`}
                            >
                              <span>
                                経管（名簿）{' '}
                                {res.isEnteral ? (
                                  <span className="rounded-md bg-amber-500 px-1.5 py-0.5 text-[10px] text-white sm:text-[11px]">
                                    管理対象
                                  </span>
                                ) : (
                                  <span className="opacity-75">—</span>
                                )}
                              </span>
                              <span>
                                当月経管実施{' '}
                                <span className="font-mono tabular-nums">{bill.enteralLogged}</span> 回
                              </span>
                            </div>
                          </div>
                        </button>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setCalOpenId((id) => (id === String(res.id) ? '' : String(res.id)));
                          }}
                          className={`mt-2 flex w-full items-center justify-center gap-1 rounded-xl border-2 py-2 text-xs font-black ${
                            critical
                              ? 'border-white/50 bg-white/10 text-white hover:bg-white/20'
                              : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                          }`}
                        >
                          <CalendarDays className="h-4 w-4" />
                          過去1週間
                        </button>
                        {calOpenId === String(res.id) && (
                          <div className="mt-2 max-w-full overflow-x-auto rounded-xl border border-slate-200 bg-slate-50 p-2 [-webkit-overflow-scrolling:touch]">
                            <table className="w-max border-collapse text-[9px] font-bold text-slate-700">
                              <thead>
                                <tr>
                                  <th className="sticky left-0 border border-slate-200 bg-slate-100 px-2 py-1 text-left">区分</th>
                                  {cal.map((day) => (
                                    <th
                                      key={`h-${day.date}`}
                                      className="border border-slate-200 bg-slate-100 px-2 py-1 text-center text-slate-500"
                                      title={day.date}
                                    >
                                      {day.date.slice(5)}
                                    </th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {[
                                  { key: 'patrol', label: '巡視', cls: 'accent-cyan-600' },
                                  { key: 'meal', label: '食事', cls: 'accent-orange-500' },
                                  { key: 'enteral', label: '経管', cls: 'accent-violet-600' },
                                  { key: 'excretion', label: '排泄', cls: 'accent-amber-600' },
                                ].map((rowDef) => (
                                  <tr key={rowDef.key}>
                                    <th className="sticky left-0 border border-slate-200 bg-white px-2 py-1 text-left">
                                      {rowDef.label}
                                    </th>
                                    {cal.map((day) => {
                                      const n = Number(day[rowDef.key] ?? 0);
                                      return (
                                        <td key={`${rowDef.key}-${day.date}`} className="border border-slate-200 px-2 py-1 text-center">
                                          <input
                                            type="checkbox"
                                            checked={n > 0}
                                            readOnly
                                            disabled
                                            className={`h-3.5 w-3.5 ${rowDef.cls}`}
                                            title={n > 0 ? `${rowDef.label} ${n}件` : `${rowDef.label} 0件`}
                                          />
                                        </td>
                                      );
                                    })}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    );
                  })}
                    </div>
                  )}
                </>
              )}
            </div>
          </section>

          <div className="order-3 shrink-0">
            <NearMissAwarenessPanel
              compact
              sheetsApiKey={SHEETS_KEY}
              facilityLinkKey={linkKeyForSheetTitle(selectedSheetTitle)}
              facilityTabLabel={selectedDef?.tabLabel ?? ''}
              onOpenAdmin={() => setNearMissAwarenessAdminOpen(true)}
            />
          </div>
        </div>

        {nursingOfficeUi ? (
        <aside className="hidden w-44 shrink-0 flex-col gap-2 xl:w-52 lg:flex">
          <div className="mb-0 w-full rounded-xl bg-slate-900/80 px-2 py-2 text-center text-xs font-bold text-cyan-400">
            外部連携
          </div>
          <ExternalToolButton href={extLinks.kaipoke} icon={Stethoscope}>
            カイポケ
          </ExternalToolButton>
          <ExternalToolButton href={extLinks.mcs} icon={LayoutGrid}>
            MCS
          </ExternalToolButton>
          <ExternalToolButton href={extLinks.line} icon={Smartphone}>
            公式LINE
          </ExternalToolButton>
        </aside>
        ) : null}
      </div>

      {nursingOfficeUi ? (
      <footer className="flex shrink-0 gap-2 rounded-2xl border border-slate-500 bg-slate-900 p-2 lg:hidden">
        <ExternalToolButton href={extLinks.kaipoke} icon={Stethoscope} layout="inline">
          カイポケ
        </ExternalToolButton>
        <ExternalToolButton href={extLinks.mcs} icon={LayoutGrid} layout="inline">
          MCS
        </ExternalToolButton>
        <ExternalToolButton href={extLinks.line} icon={Smartphone} layout="inline">
          LINE
        </ExternalToolButton>
      </footer>
      ) : null}

      {quickRes && (
        <>
          <button
            type="button"
            className="fixed inset-0 z-[160] bg-black/40"
            aria-label="背景をタップして閉じる"
            onClick={() => setQuickRes(null)}
          />
          <div
            className={`fixed inset-x-0 bottom-0 z-[161] max-h-[88vh] w-full overflow-y-auto rounded-t-3xl border-2 border-slate-300 bg-white shadow-2xl sm:bottom-auto sm:left-1/2 sm:top-1/2 sm:max-w-lg sm:-translate-x-1/2 sm:-translate-y-1/2 sm:rounded-3xl ${
              quickFlash ? 'ring-4 ring-emerald-400 ring-offset-2' : ''
            }`}
            onClick={(e) => e.stopPropagation()}
            role="dialog"
            aria-modal="true"
            aria-labelledby="quick-panel-title"
          >
            <div className="sticky top-0 z-10 flex items-center justify-between gap-2 border-b border-slate-200 bg-white px-3 py-2.5 sm:px-4">
              <div className="min-w-0 flex-1">
                <p className="text-[11px] font-bold text-slate-500">クイック記録</p>
                <p id="quick-panel-title" className="truncate text-base font-black leading-tight text-slate-900 sm:text-lg">
                  {residentNameWithoutSama(quickRes.name)} <span className="font-bold">様</span>
                  <span className="ml-1.5 text-sm font-bold text-slate-500">{String(quickRes.room ?? '')}</span>
                </p>
              </div>
              <div className="flex shrink-0 items-center gap-0.5">
                <button
                  type="button"
                  onClick={() => {
                    onSelectResident(quickRes);
                    setQuickRes(null);
                  }}
                  className="rounded-lg px-2.5 py-2 text-xs font-black text-blue-700 hover:bg-blue-50"
                >
                  詳細
                </button>
                <button
                  type="button"
                  onClick={() => setQuickRes(null)}
                  className="rounded-full p-2 text-slate-500 hover:bg-slate-100"
                  aria-label="閉じる"
                >
                  <X className="h-5 w-5 sm:h-6 sm:w-6" />
                </button>
              </div>
            </div>
            <div className="space-y-4 p-4">
              <div className="rounded-2xl border-2 border-sky-300 bg-sky-50/90 px-3 py-2.5">
                <p className="flex items-center gap-2 text-xs font-black text-sky-950">
                  <Stethoscope className="h-4 w-4 shrink-0" />
                  医療保険（名簿の列から表示）
                </p>
                <p className="mt-1 break-words text-sm font-bold text-slate-800">
                  {String(quickRes.insuranceLabel ?? '').trim() || '—（名簿に未記入）'}
                </p>
                {nursingOfficeUi &&
                  (String(quickRes.insuranceCategory ?? '') === '医療保険特指示' ||
                    /特指示|特別指示/u.test(String(quickRes.insuranceLabel ?? ''))) && (
                    <p className="mt-2 rounded-lg border border-amber-400 bg-amber-50 px-2 py-1.5 text-[11px] font-black leading-snug text-amber-950">
                      特指示の内容は Google スプレッドシートの「医療保険」列で修正し、再読込してください。ここから保険欄は保存されません。
                    </p>
                  )}
              </div>
              {nursingOfficeUi ? (
                <div className="rounded-xl border border-teal-300 bg-teal-50/70 px-3 py-3">
                  <p className="text-xs font-black text-teal-950">
                    訪問看護・特別指示（看護事務）
                    <span className="ml-1 font-bold text-teal-800">
                      ・{Report.VISIT_NURSING_SPECIAL_WARN_THRESHOLD}名以上で減算管理の確認が必要になりやすい
                    </span>
                  </p>
                  {Report.sheetSuggestsVisitNursingSpecial(quickRes.insuranceLabel) ? (
                    <p className="mt-1 text-[10px] font-bold text-teal-900">名簿で「訪問看護＋特別指示」と検出（集計対象）</p>
                  ) : null}
                  <label className="mt-2 flex cursor-pointer items-center gap-2">
                    <input
                      type="checkbox"
                      checked={quickVisitNursing}
                      onChange={(e) => setQuickVisitNursing(e.target.checked)}
                      className="h-4 w-4 shrink-0 accent-teal-600"
                    />
                    <span className="text-sm font-bold text-slate-800">手動で該当に登録（この端末）</span>
                  </label>
                  <div className="mt-2 flex flex-wrap items-end gap-3">
                    <label className="text-[11px] font-bold text-teal-900">
                      開始
                      <input
                        type="date"
                        value={quickVnPeriodStart}
                        onChange={(e) => setQuickVnPeriodStart(e.target.value)}
                        className="mt-0.5 block min-w-[10.5rem] rounded-lg border border-teal-200 bg-white px-2 py-1.5 text-sm font-bold text-slate-800"
                      />
                    </label>
                    <label className="text-[11px] font-bold text-teal-900">
                      終了
                      <input
                        type="date"
                        value={quickVnPeriodEnd}
                        onChange={(e) => setQuickVnPeriodEnd(e.target.value)}
                        className="mt-0.5 block min-w-[10.5rem] rounded-lg border border-teal-200 bg-white px-2 py-1.5 text-sm font-bold text-slate-800"
                      />
                    </label>
                  </div>
                  <p className="mt-1.5 text-[10px] font-bold text-teal-900">終了日の翌日から自動解除。未入力は手動オフまで継続。</p>
                  <textarea
                    value={quickVisitNursingNote}
                    onChange={(e) => setQuickVisitNursingNote(e.target.value)}
                    rows={2}
                    placeholder="備考（任意）"
                    className="mt-2 w-full rounded-lg border border-teal-200 bg-white px-2 py-1.5 text-sm font-bold text-slate-800"
                  />
                </div>
              ) : null}
              <div className="grid grid-cols-2 gap-3 sm:grid-cols-5">
                <label className="col-span-2 flex flex-col gap-1 sm:col-span-1">
                  <span className="flex items-center gap-1 text-xs font-black text-slate-600">
                    <Activity className="h-3.5 w-3.5" /> 体温 ℃
                  </span>
                  <input
                    value={quickTemp}
                    onChange={(e) => setQuickTemp(e.target.value)}
                    inputMode="decimal"
                    placeholder="36.5"
                    className="rounded-xl border-2 border-slate-200 px-3 py-2.5 text-base font-bold"
                  />
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-black text-slate-600">血圧 上</span>
                  <input
                    value={quickBpU}
                    onChange={(e) => setQuickBpU(e.target.value)}
                    inputMode="numeric"
                    placeholder="120"
                    className="rounded-xl border-2 border-slate-200 px-3 py-2.5 text-base font-bold"
                  />
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-black text-slate-600">血圧 下</span>
                  <input
                    value={quickBpL}
                    onChange={(e) => setQuickBpL(e.target.value)}
                    inputMode="numeric"
                    placeholder="80"
                    className="rounded-xl border-2 border-slate-200 px-3 py-2.5 text-base font-bold"
                  />
                </label>
                <label className="flex flex-col gap-1">
                  <span className="text-xs font-black text-slate-600">脈拍</span>
                  <input
                    value={quickPulse}
                    onChange={(e) => setQuickPulse(e.target.value)}
                    inputMode="numeric"
                    placeholder="72"
                    className="rounded-xl border-2 border-slate-200 px-3 py-2.5 text-base font-bold"
                  />
                </label>
                <label className="col-span-2 flex flex-col gap-1 sm:col-span-1">
                  <span className="flex items-center gap-1 text-xs font-black text-slate-600">
                    <Scale className="h-3.5 w-3.5" /> 体重 kg
                  </span>
                  <span className="text-[10px] font-bold text-slate-500">月1回測定</span>
                  <input
                    value={quickWeight}
                    onChange={(e) => setQuickWeight(e.target.value)}
                    inputMode="decimal"
                    placeholder="例: 52.3"
                    className="rounded-xl border-2 border-slate-200 px-3 py-2.5 text-base font-bold"
                    aria-label="体重（月1回）"
                  />
                </label>
              </div>
              <div className="rounded-2xl border-2 border-indigo-200 bg-indigo-50/90 px-3 py-2.5">
                <p className="text-xs font-black text-indigo-950">排尿・排便・食事（朝昼夜）・水分・内服</p>
                <p className="mt-0.5 text-[10px] font-bold text-indigo-800">
                  一覧表入力と同じ項目です。未入力の欄はログに出しません。右のマイクで音声入力できます（排便は多・中・小／性状／割・水分の数字などに変換します）。
                  <span className="mt-1 block text-indigo-950">
                    「食事の確認」ON のときは主食・副食の割が必須。「排泄の確認」ON のときは排尿または排便の量と排便性状が必須です。
                  </span>
                </p>
                <label className="mt-2 flex cursor-pointer items-center gap-2 rounded-xl border-2 border-sky-300 bg-sky-100/80 px-3 py-2">
                  <input
                    type="checkbox"
                    checked={Boolean(quickCareDetail.toiletGuidance)}
                    onChange={(e) =>
                      setQuickCareDetail((d) => ({ ...d, toiletGuidance: e.target.checked }))
                    }
                    className="h-5 w-5 shrink-0 accent-sky-700"
                  />
                  <span className="text-sm font-black text-sky-950">
                    トイレ誘導を実施した（排尿間隔アラートの基準を更新）
                  </span>
                </label>
                <div className="mt-2 grid grid-cols-1 gap-2 sm:grid-cols-2">
                  <div className="flex items-end gap-1">
                    <label className="flex min-w-0 flex-1 flex-col gap-0.5">
                      <span className="text-[10px] font-black text-slate-600">排尿量</span>
                      <input
                        value={quickCareDetail.urineVolume}
                        onChange={(e) =>
                          setQuickCareDetail((d) => ({ ...d, urineVolume: e.target.value }))
                        }
                        className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                        placeholder="例: 200ml"
                      />
                    </label>
                    <button
                      type="button"
                      title="音声で追記"
                      aria-label="排尿量 音声入力"
                      onClick={() =>
                        dictatingQuickKey === 'urineVolume'
                          ? stopQuickCareDictation()
                          : startQuickCareDictation('urineVolume')
                      }
                      className={`mb-0.5 shrink-0 rounded-lg border-2 px-2 py-2 ${
                        dictatingQuickKey === 'urineVolume'
                          ? 'border-rose-400 bg-rose-50 text-rose-800'
                          : 'border-slate-200 bg-white text-slate-600'
                      }`}
                    >
                      <Mic className="h-4 w-4" aria-hidden />
                    </button>
                  </div>
                  <div className="flex items-end gap-1">
                    <label className="flex min-w-0 flex-1 flex-col gap-0.5">
                      <span className="text-[10px] font-black text-slate-600">排便量</span>
                      <select
                        value={quickCareDetail.stoolVolume}
                        onChange={(e) =>
                          setQuickCareDetail((d) => ({ ...d, stoolVolume: e.target.value }))
                        }
                        className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                      >
                        {STOOL_VOLUME_OPTIONS.map((opt) => (
                          <option key={opt || '—'} value={opt}>
                            {opt || '—'}
                          </option>
                        ))}
                      </select>
                    </label>
                    <button
                      type="button"
                      title="音声（多・中・小）"
                      aria-label="排便量 音声入力"
                      onClick={() =>
                        dictatingQuickKey === 'stoolVolume'
                          ? stopQuickCareDictation()
                          : startQuickCareDictation('stoolVolume')
                      }
                      className={`mb-0.5 shrink-0 rounded-lg border-2 px-2 py-2 ${
                        dictatingQuickKey === 'stoolVolume'
                          ? 'border-rose-400 bg-rose-50 text-rose-800'
                          : 'border-slate-200 bg-white text-slate-600'
                      }`}
                    >
                      <Mic className="h-4 w-4" aria-hidden />
                    </button>
                  </div>
                  <div className="flex items-end gap-1 sm:col-span-2">
                    <label className="flex min-w-0 flex-1 flex-col gap-0.5">
                      <span className="text-[10px] font-black text-slate-600">排便性状</span>
                      <select
                        value={quickCareDetail.stoolCharacter}
                        onChange={(e) =>
                          setQuickCareDetail((d) => ({ ...d, stoolCharacter: e.target.value }))
                        }
                        className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                      >
                        {STOOL_CHARACTER_OPTIONS.map((opt) => (
                          <option key={opt || '—'} value={opt}>
                            {opt || '—'}
                          </option>
                        ))}
                      </select>
                    </label>
                    <button
                      type="button"
                      title="音声（普通便・硬便・軟便・水様便）"
                      aria-label="排便性状 音声入力"
                      onClick={() =>
                        dictatingQuickKey === 'stoolCharacter'
                          ? stopQuickCareDictation()
                          : startQuickCareDictation('stoolCharacter')
                      }
                      className={`mb-0.5 shrink-0 rounded-lg border-2 px-2 py-2 ${
                        dictatingQuickKey === 'stoolCharacter'
                          ? 'border-rose-400 bg-rose-50 text-rose-800'
                          : 'border-slate-200 bg-white text-slate-600'
                      }`}
                    >
                      <Mic className="h-4 w-4" aria-hidden />
                    </button>
                  </div>
                  <label className="flex flex-col gap-0.5">
                    <span className="text-[10px] font-black text-slate-600">食事区分</span>
                    <select
                      value={quickCareDetail.mealSlot}
                      onChange={(e) =>
                        setQuickCareDetail((d) => ({ ...d, mealSlot: e.target.value }))
                      }
                      className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                    >
                      <option value="">—</option>
                      <option value="朝">朝</option>
                      <option value="昼">昼</option>
                      <option value="夜">夜</option>
                    </select>
                  </label>
                  <div className="flex items-end gap-1">
                    <label className="flex min-w-0 flex-1 flex-col gap-0.5">
                      <span className="text-[10px] font-black text-slate-600">主食（割）</span>
                      <select
                        value={quickCareDetail.mealStaple}
                        onChange={(e) =>
                          setQuickCareDetail((d) => ({ ...d, mealStaple: e.target.value }))
                        }
                        className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                      >
                        {MEAL_WARI_OPTIONS.map((opt) => (
                          <option key={opt || '—'} value={opt}>
                            {opt || '—'}
                          </option>
                        ))}
                      </select>
                    </label>
                    <button
                      type="button"
                      title="音声（10割〜0割）"
                      aria-label="主食 音声入力"
                      onClick={() =>
                        dictatingQuickKey === 'mealStaple'
                          ? stopQuickCareDictation()
                          : startQuickCareDictation('mealStaple')
                      }
                      className={`mb-0.5 shrink-0 rounded-lg border-2 px-2 py-2 ${
                        dictatingQuickKey === 'mealStaple'
                          ? 'border-rose-400 bg-rose-50 text-rose-800'
                          : 'border-slate-200 bg-white text-slate-600'
                      }`}
                    >
                      <Mic className="h-4 w-4" aria-hidden />
                    </button>
                  </div>
                  <div className="flex items-end gap-1">
                    <label className="flex min-w-0 flex-1 flex-col gap-0.5">
                      <span className="text-[10px] font-black text-slate-600">副食（割）</span>
                      <select
                        value={quickCareDetail.mealSide}
                        onChange={(e) =>
                          setQuickCareDetail((d) => ({ ...d, mealSide: e.target.value }))
                        }
                        className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                      >
                        {MEAL_WARI_OPTIONS.map((opt) => (
                          <option key={`s-${opt || '—'}`} value={opt}>
                            {opt || '—'}
                          </option>
                        ))}
                      </select>
                    </label>
                    <button
                      type="button"
                      title="音声（10割〜0割）"
                      aria-label="副食 音声入力"
                      onClick={() =>
                        dictatingQuickKey === 'mealSide'
                          ? stopQuickCareDictation()
                          : startQuickCareDictation('mealSide')
                      }
                      className={`mb-0.5 shrink-0 rounded-lg border-2 px-2 py-2 ${
                        dictatingQuickKey === 'mealSide'
                          ? 'border-rose-400 bg-rose-50 text-rose-800'
                          : 'border-slate-200 bg-white text-slate-600'
                      }`}
                    >
                      <Mic className="h-4 w-4" aria-hidden />
                    </button>
                  </div>
                  <div className="flex items-end gap-1">
                    <label className="flex min-w-0 flex-1 flex-col gap-0.5">
                      <span className="text-[10px] font-black text-slate-600">水分量（ml）</span>
                      <input
                        value={quickCareDetail.waterMl}
                        onChange={(e) =>
                          setQuickCareDetail((d) => ({ ...d, waterMl: e.target.value }))
                        }
                        inputMode="numeric"
                        className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                        placeholder="例: 150"
                      />
                    </label>
                    <button
                      type="button"
                      title="音声で ml の数字"
                      aria-label="水分量 音声入力"
                      onClick={() =>
                        dictatingQuickKey === 'waterMl'
                          ? stopQuickCareDictation()
                          : startQuickCareDictation('waterMl')
                      }
                      className={`mb-0.5 shrink-0 rounded-lg border-2 px-2 py-2 ${
                        dictatingQuickKey === 'waterMl'
                          ? 'border-rose-400 bg-rose-50 text-rose-800'
                          : 'border-slate-200 bg-white text-slate-600'
                      }`}
                    >
                      <Mic className="h-4 w-4" aria-hidden />
                    </button>
                  </div>
                  <label className="flex flex-col gap-0.5">
                    <span className="text-[10px] font-black text-slate-600">内服</span>
                    <select
                      value={quickCareDetail.medicationTaken}
                      onChange={(e) =>
                        setQuickCareDetail((d) => ({ ...d, medicationTaken: e.target.value }))
                      }
                      className="rounded-lg border-2 border-white bg-white px-2 py-2 text-sm font-bold"
                    >
                      <option value="">—</option>
                      <option value="yes">飲了</option>
                      <option value="no">未服</option>
                    </select>
                  </label>
                </div>
              </div>
              <div className="space-y-2 rounded-2xl border-2 border-slate-200 bg-slate-50 p-3">
                <p className="text-xs font-black text-slate-500">ケア確認（チェックして保存）</p>
                <p className="text-[10px] font-bold leading-snug text-slate-600">
                  食事の確認: 主食・副食の割（何割）が両方必要。排泄の確認: 排便性状に加え、排尿量または排便量のどちらかが必要。
                </p>
                <label className="flex cursor-pointer items-center gap-3 rounded-xl border-2 border-white bg-white px-3 py-3 shadow-sm">
                  <input
                    type="checkbox"
                    checked={chkPatrol}
                    onChange={(e) => setChkPatrol(e.target.checked)}
                    className="h-5 w-5 accent-cyan-600"
                  />
                  <span className="flex flex-1 items-center gap-2 text-sm font-black text-slate-800">
                    <Wind className="h-5 w-5 text-cyan-600" />
                    巡回（巡視）3時間おき
                  </span>
                  {chkPatrol ? <Check className="h-5 w-5 text-emerald-600" /> : null}
                </label>
                <div className="flex flex-col gap-2 rounded-xl border-2 border-white bg-white px-3 py-2 shadow-sm sm:flex-row sm:items-end">
                  <label className="flex min-w-0 flex-1 flex-col gap-1">
                    <span className="text-[11px] font-black text-slate-600">巡視の日付</span>
                    <input
                      type="date"
                      value={quickPatrolSlot.date}
                      onChange={(e) =>
                        setQuickPatrolAt(joinPatrolDateTimeLocal(e.target.value, quickPatrolSlot.hour))
                      }
                      className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-700"
                    />
                  </label>
                  <label className="flex min-w-0 flex-1 flex-col gap-1">
                    <span className="text-[11px] font-black text-slate-600">時刻（0時から3時間おき）</span>
                    <select
                      value={quickPatrolSlot.hour}
                      onChange={(e) =>
                        setQuickPatrolAt(
                          joinPatrolDateTimeLocal(quickPatrolSlot.date, Number(e.target.value))
                        )
                      }
                      className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-sm font-bold text-slate-700"
                    >
                      {PATROL_SLOT_HOURS.map((h) => (
                        <option key={h} value={h}>
                          {String(h).padStart(2, '0')}:00
                        </option>
                      ))}
                    </select>
                  </label>
                </div>
                <p className="text-[10px] font-bold text-slate-500">当日・過去日・未来日のいずれでも選べます。時刻は 0・3・6…21 時のみです。</p>
                <label className="flex cursor-pointer items-center gap-3 rounded-xl border-2 border-white bg-white px-3 py-3 shadow-sm">
                  <input
                    type="checkbox"
                    checked={chkExcretion}
                    onChange={(e) => setChkExcretion(e.target.checked)}
                    className="h-5 w-5 accent-amber-600"
                  />
                  <span className="text-sm font-black text-slate-800">排泄の確認</span>
                  {chkExcretion ? <Check className="h-5 w-5 text-emerald-600" /> : null}
                </label>
                <label className="flex cursor-pointer items-center gap-3 rounded-xl border-2 border-white bg-white px-3 py-3 shadow-sm">
                  <input
                    type="checkbox"
                    checked={chkMeal}
                    onChange={(e) => setChkMeal(e.target.checked)}
                    className="h-5 w-5 accent-orange-500"
                  />
                  <span className="flex flex-1 items-center gap-2 text-sm font-black text-slate-800">
                    <Utensils className="h-5 w-5 text-orange-600" />
                    食事の確認
                  </span>
                  {chkMeal ? <Check className="h-5 w-5 text-emerald-600" /> : null}
                </label>
              </div>
              <div className="space-y-2">
                <button
                  type="button"
                  disabled={quickSaveBusy}
                  onClick={saveQuickPanel}
                  className="flex w-full items-center justify-center gap-2 rounded-2xl bg-emerald-600 py-4 text-base font-black text-white shadow-lg active:scale-[0.99] disabled:cursor-not-allowed disabled:opacity-60"
                >
                  {quickSaveBusy ? (
                    <Loader2 className="h-5 w-5 shrink-0 animate-spin" aria-hidden />
                  ) : (
                    <Save className="h-5 w-5 shrink-0" aria-hidden />
                  )}
                  {quickSaveBusy ? '保存中…' : '保存して反映'}
                </button>
                {quickFlash ? (
                  <p
                    className="text-center text-sm font-black text-emerald-700"
                    role="status"
                    aria-live="polite"
                  >
                    保存しました。下の一覧・集計に反映済みです。
                  </p>
                ) : null}
              </div>
            </div>
          </div>
        </>
      )}

      {emergencyOpen && (
        <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/60 p-4">
          <div className="max-h-[90vh] w-full max-w-5xl overflow-y-auto rounded-3xl border-4 border-rose-500 bg-white p-6 shadow-2xl">
            <div className="mb-4 flex items-center justify-between">
              <h3 className="flex items-center gap-2 text-xl font-black text-rose-700">
                <Ambulance className="h-7 w-7" />
                救急搬送サマリー
              </h3>
              <button type="button" onClick={() => setEmergencyOpen(false)} className="rounded-full p-2 hover:bg-slate-100">
                <X className="h-6 w-6" />
              </button>
            </div>
            <p className="mb-3 text-sm font-bold text-slate-600">
              現在の書式に合わせて追記できるようにしています。音声入力は各欄のマイクを押してください。
            </p>
            <div className="mb-4 grid gap-3 md:grid-cols-2">
              <select
                value={emergencyPickId}
                onChange={(e) => setEmergencyPickId(e.target.value)}
                className="w-full rounded-xl border-2 border-slate-300 px-3 py-3 text-base font-bold"
              >
                <option value="">— 利用者を選択 —</option>
                {filteredResidents.map((r) => (
                  <option key={String(r.id)} value={String(r.id)}>
                    {residentNameWithoutSama(r.name)} 様 {String(r.room)}
                  </option>
                ))}
              </select>
              <div className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-sm">
                <p className="font-bold text-slate-700">
                  {residentNameWithoutSama(selectedEmergencyResident?.name ?? '—')} 様
                </p>
                <p className="text-slate-500">居室: {String(selectedEmergencyResident?.room ?? '—')}</p>
              </div>
            </div>

            <div className="mb-4 grid gap-3 md:grid-cols-2">
              {[
                ['senderOffice', '施設名'],
                ['senderAddress', '住所'],
                ['senderTel', 'ステーション電話番号'],
                ['senderNurse', '担当看護師'],
                ['primaryDoctor', '主治医氏名'],
                ['medicalAgency', '医療機関名'],
                ['medicalAddress', '医療機関住所'],
              ].map(([k, label]) => (
                <label key={k} className="flex flex-col gap-1">
                  <span className="text-xs font-bold text-slate-600">{label}</span>
                  <input
                    value={emergencyDraft[k] ?? ''}
                    onChange={(e) => setEmergencyDraft((prev) => ({ ...prev, [k]: e.target.value }))}
                    className="rounded-lg border border-slate-300 px-3 py-2 text-sm font-bold"
                  />
                </label>
              ))}
            </div>

            <div className="mb-4 grid gap-3 md:grid-cols-2">
              {[
                ['dailyLife', '日常生活等の状況'],
                ['nurseProblems', '看護上の問題等'],
                ['acuteChange', '急変の内容（看護師記入）'],
                ['nurseContent', '看護の内容'],
                ['careNotes', 'ケア時の注意点'],
                ['other', 'その他'],
              ].map(([k, label]) => (
                <label key={k} className="flex flex-col gap-1 md:col-span-1">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-600">{label}</span>
                    <button
                      type="button"
                      onClick={() => (dictatingField === k ? stopDictation() : startDictation(k))}
                      className={`inline-flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-bold ${
                        dictatingField === k ? 'bg-rose-100 text-rose-700' : 'bg-slate-100 text-slate-700'
                      }`}
                    >
                      <Mic className="h-3.5 w-3.5" />
                      {dictatingField === k ? '停止' : '音声入力'}
                    </button>
                  </div>
                  <textarea
                    value={emergencyDraft[k] ?? ''}
                    onChange={(e) => setEmergencyDraft((prev) => ({ ...prev, [k]: e.target.value }))}
                    rows={4}
                    className="rounded-lg border border-slate-300 px-3 py-2 text-sm font-bold"
                  />
                </label>
              ))}
            </div>
            <div className="flex flex-col gap-2 sm:flex-row">
              <button
                type="button"
                disabled={!emergencyPickId || emergencyBusy}
                onClick={() =>
                  setEmergencyDraft((prev) =>
                    buildEmergencyDraftFromResident(selectedEmergencyResident, prev)
                  )
                }
                className="flex flex-1 items-center justify-center gap-2 rounded-2xl border-2 border-slate-300 py-4 text-base font-black text-slate-700 disabled:opacity-50"
              >
                自動作成（まとめて）
              </button>
              <button
                type="button"
                disabled={!emergencyPickId || emergencyBusy}
                onClick={() => void runEmergencySummary()}
                className="flex flex-1 items-center justify-center gap-2 rounded-2xl bg-rose-600 py-4 text-base font-black text-white disabled:opacity-50"
              >
                {emergencyBusy ? <Loader2 className="h-5 w-5 animate-spin" /> : null}
                印刷で開く
              </button>
              <button
                type="button"
                disabled={!emergencyPickId || emergencyBusy}
                onClick={() => void downloadEmergencyHtml()}
                className="flex flex-1 items-center justify-center gap-2 rounded-2xl border-2 border-rose-400 py-4 text-base font-black text-rose-700 disabled:opacity-50"
              >
                HTML保存
              </button>
            </div>
            {!GEMINI_KEY && (
              <p className="mt-3 text-xs text-amber-700">
                VITE_GEMINI_API_KEY 未設定時は定型の法令・実務アドバイスのみ添付されます。
              </p>
            )}
          </div>
        </div>
      )}

      <AccidentReportModal
        open={accidentReportOpen}
        onClose={() => setAccidentReportOpen(false)}
        geminiKey={GEMINI_KEY}
        facilityLabel={selectedDef?.tabLabel ?? ''}
        residents={filteredResidents}
      />
      <AccidentMonthlyAnalysisModal
        open={accidentMonthlyOpen}
        onClose={() => setAccidentMonthlyOpen(false)}
        geminiKey={GEMINI_KEY}
        defaultTabLabel={selectedDef?.tabLabel ?? ''}
        facilityDefs={CARELINK_FACILITIES}
      />
      <NearMissMonthlyAnalysisModal
        open={nearMissMonthlyOpen}
        onClose={() => setNearMissMonthlyOpen(false)}
        geminiKey={GEMINI_KEY}
        defaultTabLabel={selectedDef?.tabLabel ?? ''}
        facilityDefs={CARELINK_FACILITIES}
      />
      <NearMissReportModal
        open={nearMissOpen}
        onClose={() => setNearMissOpen(false)}
        geminiKey={GEMINI_KEY}
        facilityLabel={selectedDef?.tabLabel ?? ''}
        residents={filteredResidents}
      />
      <NearMissAwarenessAdminModal
        open={nearMissAwarenessAdminOpen}
        onClose={() => setNearMissAwarenessAdminOpen(false)}
        facilityLinkKey={linkKeyForSheetTitle(selectedSheetTitle)}
        facilityTabLabel={selectedDef?.tabLabel ?? ''}
        sheetsApiKey={SHEETS_KEY}
        geminiKey={GEMINI_KEY}
      />
    </div>
  );
}
