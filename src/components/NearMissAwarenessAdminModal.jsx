import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  BarChart3,
  Loader2,
  Plus,
  Trash2,
  UserPlus,
  X,
} from 'lucide-react';
import {
  aggregateLedgerCategoriesForMonth,
  buildUnconfirmedMatrix,
  exportLedgerReportsCsv,
  fetchNearMissTrendAssessmentAi,
  getEffectiveStaffRosterForFacility,
  getLedgerSpreadsheetId,
  getStaffRosterByFacility,
  getSyncedRosterPayload,
  refreshLedgerFromSpreadsheet,
  setStaffRosterForFacility,
  syncStaffRosterFromHrSheetAndStore,
  syncStaffRosterFromShiftScheduleAndStore,
} from '../services/NearMissLedgerService.js';

function CategoryBars({ counts }) {
  const entries = Object.entries(counts)
    .filter(([, v]) => v > 0)
    .sort((a, b) => b[1] - a[1]);
  const max = Math.max(1, ...entries.map(([, v]) => v));
  if (!entries.length) {
    return <p className="text-sm font-bold text-slate-500">今月のデータがありません</p>;
  }
  return (
    <div className="space-y-2">
      {entries.map(([k, v]) => (
        <div key={k}>
          <div className="flex justify-between text-xs font-black text-slate-800">
            <span>{k}</span>
            <span>{v} 件</span>
          </div>
          <div className="h-2.5 overflow-hidden rounded-full bg-slate-200">
            <div
              className="h-2.5 rounded-full bg-gradient-to-r from-teal-500 to-cyan-500"
              style={{ width: `${Math.min(100, (v / max) * 100)}%` }}
            />
          </div>
        </div>
      ))}
    </div>
  );
}

/**
 * @param {{
 *   open: boolean;
 *   onClose: () => void;
 *   facilityLinkKey: string;
 *   facilityTabLabel: string;
 *   sheetsApiKey: string;
 *   geminiKey: string;
 * }} props
 */
export function NearMissAwarenessAdminModal({
  open,
  onClose,
  facilityLinkKey,
  facilityTabLabel,
  sheetsApiKey,
  geminiKey,
}) {
  const [rev, setRev] = useState(0);
  const [loading, setLoading] = useState(false);
  const [aiText, setAiText] = useState('');
  const [aiBusy, setAiBusy] = useState(false);
  const [yearMonth, setYearMonth] = useState(() => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  });
  const [hrBusy, setHrBusy] = useState(false);
  const [shiftBusy, setShiftBusy] = useState(false);
  const [hrSheetHint, setHrSheetHint] = useState('');

  const lk = String(facilityLinkKey ?? '').trim();

  const syncedPayload = useMemo(() => getSyncedRosterPayload(), [rev, open]);
  const roster = useMemo(() => getStaffRosterByFacility()[lk] ?? [], [lk, rev]);
  const effectiveRoster = useMemo(() => getEffectiveStaffRosterForFacility(lk), [lk, rev]);

  const matrix = useMemo(() => buildUnconfirmedMatrix(lk), [lk, rev]);

  const counts = useMemo(() => aggregateLedgerCategoriesForMonth(yearMonth), [yearMonth, rev]);

  useEffect(() => {
    if (!open) return;
    setAiText('');
  }, [open, yearMonth]);

  const syncSheet = useCallback(async () => {
    const key = String(sheetsApiKey ?? '').trim();
    const id = getLedgerSpreadsheetId().trim();
    if (!key || !id) return;
    setLoading(true);
    try {
      await refreshLedgerFromSpreadsheet(key);
      setRev((x) => x + 1);
    } catch (e) {
      alert(e instanceof Error ? e.message : '同期に失敗しました');
    } finally {
      setLoading(false);
    }
  }, [sheetsApiKey]);

  const syncHrRoster = useCallback(async () => {
    const key = String(sheetsApiKey ?? '').trim();
    if (!key) {
      alert('VITE_GOOGLE_SHEETS_API_KEY が必要です');
      return;
    }
    setHrBusy(true);
    try {
      await syncStaffRosterFromHrSheetAndStore(key, { preferredSheetTitle: hrSheetHint.trim() });
      setRev((x) => x + 1);
    } catch (e) {
      alert(e instanceof Error ? e.message : 'スタッフ名簿の同期に失敗しました');
    } finally {
      setHrBusy(false);
    }
  }, [sheetsApiKey, hrSheetHint]);

  const syncShiftRoster = useCallback(() => {
    setShiftBusy(true);
    try {
      const data = syncStaffRosterFromShiftScheduleAndStore();
      const n = data.meta?.rowCount ?? 0;
      const here = (data.byFacility?.[lk] ?? []).length;
      setRev((x) => x + 1);
      if (n === 0) {
        alert(
          '勤務表にスタッフがいません。「勤務希望・勤務表」画面で、このブラウザに氏名と施設を保存してください。'
        );
      } else if (here === 0) {
        alert(
          `勤務表には他施設で合計 ${n} 名いますが、いまの施設キー「${lk}」では 0 名です。勤務表画面で施設を「${facilityTabLabel || lk}」側に合わせて登録し直してください。`
        );
      }
    } catch (e) {
      alert(e instanceof Error ? e.message : '勤務表からの反映に失敗しました');
    } finally {
      setShiftBusy(false);
    }
  }, [lk, facilityTabLabel]);

  const runAi = useCallback(async () => {
    if (!geminiKey?.trim()) {
      alert('VITE_GEMINI_API_KEY が未設定です');
      return;
    }
    setAiBusy(true);
    setAiText('');
    try {
      const t = await fetchNearMissTrendAssessmentAi(geminiKey, yearMonth);
      setAiText(t);
    } catch (e) {
      setAiText(e instanceof Error ? e.message : 'AI エラー');
    } finally {
      setAiBusy(false);
    }
  }, [geminiKey, yearMonth]);

  const updateRosterName = (idx, name) => {
    const next = roster.map((r, i) => (i === idx ? { ...r, name } : r));
    setStaffRosterForFacility(lk, next);
    setRev((x) => x + 1);
  };

  const addRoster = () => {
    const next = [...roster, { id: `s-${Date.now()}`, name: '' }];
    setStaffRosterForFacility(lk, next);
    setRev((x) => x + 1);
  };

  const removeRoster = (idx) => {
    const next = roster.filter((_, i) => i !== idx);
    setStaffRosterForFacility(lk, next);
    setRev((x) => x + 1);
  };

  const downloadCsv = () => {
    const csv = exportLedgerReportsCsv();
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `ヒヤリハット管理簿_報告_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
  };

  if (!open) return null;

  return (
    <div className="fixed inset-0 z-[220] flex items-center justify-center bg-black/60 p-3 sm:p-4">
      <div className="max-h-[92vh] w-full max-w-4xl overflow-y-auto rounded-3xl border-4 border-slate-700 bg-white shadow-2xl">
        <div className="sticky top-0 z-10 flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-white px-4 py-3">
          <div>
            <h2 className="text-lg font-black text-slate-900 sm:text-xl">ヒヤリハット周知 — 管理・統計</h2>
            <p className="text-xs font-bold text-slate-500">
              {facilityTabLabel || lk || '施設未選択'} — 未確認者の一覧と傾向分析
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              disabled={loading || !sheetsApiKey?.trim() || !getLedgerSpreadsheetId().trim()}
              onClick={() => void syncSheet()}
              className="rounded-xl border-2 border-slate-400 bg-slate-100 px-3 py-2 text-xs font-black text-slate-800 hover:bg-slate-200 disabled:opacity-50"
            >
              {loading ? <Loader2 className="inline h-4 w-4 animate-spin" /> : null} 台帳同期
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded-xl border-2 border-slate-300 p-2 text-slate-600 hover:bg-slate-100"
              aria-label="閉じる"
            >
              <X className="h-5 w-5" />
            </button>
          </div>
        </div>

        <div className="space-y-6 p-4 sm:p-6">
          <section className="rounded-2xl border-2 border-cyan-300 bg-cyan-50/40 p-4">
            <h3 className="mb-2 text-base font-black text-cyan-950">周知チェック用スタッフ名簿</h3>
            <p className="mb-3 text-xs font-bold text-cyan-950/90">
              未確認者一覧は<strong>下のどちらか一方</strong>を最後に押した内容で上書きされます。
              <strong>勤務表</strong>はアプリ内の「勤務希望・勤務表」に登録した氏名のみ（現場の当番表に近い）。
              <strong>求人シート</strong>は別スプレッドシートの名簿です。名簿が合わないときは勤務表を選んでください。
              求人と同じシートから<strong>勤務希望へ氏名・部署を反映</strong>するには、勤務表画面の「求人シートから勤務希望へ反映」とタグ内の部署名をご利用ください。
            </p>
            <div className="mb-4 rounded-xl border border-emerald-400 bg-emerald-50/80 p-3">
              <h4 className="mb-2 text-sm font-black text-emerald-950">① 勤務表（推奨・API 不要）</h4>
              <button
                type="button"
                disabled={shiftBusy}
                onClick={() => void syncShiftRoster()}
                className="rounded-xl bg-emerald-600 px-4 py-2 text-sm font-black text-white hover:bg-emerald-500 disabled:opacity-50"
              >
                {shiftBusy ? <Loader2 className="inline h-4 w-4 animate-spin" /> : null} 勤務表から名簿を反映
              </button>
            </div>
            <div className="mb-2">
              <h4 className="mb-2 text-sm font-black text-cyan-950">② 求人・入退社シート（従来）</h4>
              <p className="mb-2 text-xs font-bold text-cyan-950/90">
                「スタッフ名」「在籍状況」列を読み、在籍者だけを対象にします。
              </p>
            </div>
            <div className="mb-2 flex flex-wrap gap-2">
              <input
                type="text"
                value={hrSheetHint}
                onChange={(e) => setHrSheetHint(e.target.value)}
                placeholder="シート名（空欄で自動：求人・スタッフ等を優先）"
                className="min-w-[12rem] flex-1 rounded-lg border border-cyan-400 px-2 py-2 text-sm font-bold"
              />
              <button
                type="button"
                disabled={hrBusy || !sheetsApiKey?.trim()}
                onClick={() => void syncHrRoster()}
                className="rounded-xl bg-cyan-600 px-4 py-2 text-sm font-black text-white hover:bg-cyan-500 disabled:opacity-50"
              >
                {hrBusy ? <Loader2 className="inline h-4 w-4 animate-spin" /> : null} 求人シートから同期
              </button>
            </div>
            {syncedPayload?.syncedAt ? (
              <p className="text-xs font-bold text-cyan-900">
                最終更新: {String(syncedPayload.syncedAt).slice(0, 19).replace('T', ' ')} ／ ソース:{' '}
                {syncedPayload.meta?.source === 'shift_schedule' ? '勤務表' : '求人シート'} ／{' '}
                {String(syncedPayload.sheetTitle ?? '')} ／ 合計 {syncedPayload.meta?.rowCount ?? 0} 名
                {syncedPayload.meta?.source === 'hr_sheet' ? '（在籍のみ）' : ''}
              </p>
            ) : (
              <p className="text-xs font-bold text-amber-800">まだ名簿を取り込んでいません。上の①または②を押してください。</p>
            )}
            {effectiveRoster.length > 0 ? (
              <ul className="mt-2 max-h-32 overflow-y-auto rounded-lg border border-cyan-200 bg-white/90 p-2 text-xs font-bold text-slate-800">
                {effectiveRoster.map((r) => (
                  <li key={r.id}>・{r.name}</li>
                ))}
              </ul>
            ) : null}
          </section>

          <section className="rounded-2xl border-2 border-indigo-200 bg-indigo-50/50 p-4">
            <div className="mb-2 flex items-center gap-2">
              <UserPlus className="h-5 w-5 text-indigo-700" />
              <h3 className="text-base font-black text-indigo-950">手入力スタッフ名簿（同期が無いときの代替）</h3>
            </div>
            <p className="mb-3 text-xs font-bold text-indigo-900/90">
              勤務表・求人のどちらかで取り込んだ名簿が優先されます。取り込み前・補足用にだけ編集してください。
            </p>
            <ul className="space-y-2">
              {roster.map((r, i) => (
                <li key={r.id} className="flex flex-wrap items-center gap-2">
                  <input
                    value={r.name}
                    onChange={(e) => updateRosterName(i, e.target.value)}
                    placeholder={`スタッフ ${i + 1}`}
                    className="min-w-[12rem] flex-1 rounded-lg border border-indigo-300 px-2 py-2 text-sm font-bold"
                  />
                  <button
                    type="button"
                    onClick={() => removeRoster(i)}
                    className="rounded-lg border border-rose-300 bg-rose-50 p-2 text-rose-700 hover:bg-rose-100"
                    aria-label="削除"
                  >
                    <Trash2 className="h-4 w-4" />
                  </button>
                </li>
              ))}
            </ul>
            <button
              type="button"
              onClick={addRoster}
              className="mt-3 inline-flex items-center gap-1 rounded-xl bg-indigo-600 px-4 py-2 text-sm font-black text-white hover:bg-indigo-500"
            >
              <Plus className="h-4 w-4" /> 行を追加
            </button>
          </section>

          <section className="rounded-2xl border-2 border-amber-200 bg-amber-50/40 p-4">
            <h3 className="mb-2 text-base font-black text-amber-950">周知の確認状況（施設内）</h3>
            <p className="mb-3 text-xs font-bold text-amber-900/90">
              名簿は「勤務表から反映」または「求人シート」で取り込みます。台帳を同期すると、他端末・スプレッドシート経由の「確認しました」もここに集まります。
            </p>
            {matrix.length === 0 ? (
              <p className="text-sm font-bold text-slate-600">表示する周知がありません</p>
            ) : (
              <div className="max-h-72 overflow-auto rounded-lg border border-amber-200 bg-white">
                <table className="w-full text-left text-[10px] font-bold sm:text-xs">
                  <thead className="sticky top-0 bg-amber-100">
                    <tr>
                      <th className="px-2 py-2">周知タイトル</th>
                      <th className="px-2 py-2 text-emerald-900">確認済み（名簿）</th>
                      <th className="px-2 py-2 text-rose-900">未確認</th>
                      <th className="px-2 py-2 text-slate-700">名簿外の確認ログ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {matrix.map(({ notice, missing, confirmed, extraAckNames }) => (
                      <tr key={String(notice.id)} className="border-t border-amber-100">
                        <td className="px-2 py-2 align-top text-slate-800">{notice.title}</td>
                        <td className="px-2 py-2 align-top text-emerald-800">
                          {confirmed.map((c) => c.name).join('、') || '—'}
                        </td>
                        <td className="px-2 py-2 align-top text-rose-800">
                          {missing.map((m) => m.name).join('、') || '—'}
                        </td>
                        <td className="px-2 py-2 align-top text-slate-600">
                          {extraAckNames.length ? extraAckNames.join('、') : '—'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>

          <section className="rounded-2xl border-2 border-teal-200 bg-teal-50/30 p-4">
            <div className="mb-3 flex flex-wrap items-end justify-between gap-2">
              <div className="flex items-center gap-2">
                <BarChart3 className="h-6 w-6 text-teal-700" />
                <h3 className="text-base font-black text-teal-950">今月の傾向（カテゴリ別件数）</h3>
              </div>
              <label className="flex items-center gap-2 text-xs font-bold text-slate-600">
                対象月
                <input
                  type="month"
                  value={yearMonth}
                  onChange={(e) => setYearMonth(e.target.value)}
                  className="rounded-lg border border-slate-300 px-2 py-1 font-bold"
                />
              </label>
            </div>
            <CategoryBars counts={counts} />
            <div className="mt-4 rounded-xl border border-teal-200 bg-white/90 p-3">
              <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                <span className="text-xs font-black text-teal-900">AI コメント（傾向）</span>
                <button
                  type="button"
                  disabled={aiBusy}
                  onClick={() => void runAi()}
                  className="rounded-lg bg-teal-600 px-3 py-1.5 text-xs font-black text-white hover:bg-teal-500 disabled:opacity-50"
                >
                  {aiBusy ? <Loader2 className="inline h-4 w-4 animate-spin" /> : null} 分析を実行
                </button>
              </div>
              {aiText ? (
                <p className="whitespace-pre-wrap text-sm font-bold leading-relaxed text-slate-800">{aiText}</p>
              ) : (
                <p className="text-xs font-bold text-slate-500">「分析を実行」で Gemini が件数からコメントを生成します。</p>
              )}
            </div>
          </section>

          <section className="flex flex-wrap gap-2 border-t border-slate-200 pt-4">
            <button
              type="button"
              onClick={downloadCsv}
              className="rounded-xl border-2 border-slate-400 bg-slate-100 px-4 py-2 text-sm font-black text-slate-800 hover:bg-slate-200"
            >
              報告データを CSV 出力（バックアップ）
            </button>
          </section>
        </div>
      </div>
    </div>
  );
}
