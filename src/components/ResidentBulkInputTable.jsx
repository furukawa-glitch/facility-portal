import React from 'react';
import { Table2 } from 'lucide-react';
import {
  STOOL_VOLUME_OPTIONS,
  STOOL_CHARACTER_OPTIONS,
  MEAL_WARI_OPTIONS,
} from '../lib/careQuickCareFields.js';
import { PATROL_SLOT_HOURS, joinPatrolDateTimeLocal, splitPatrolDateTimeLocal } from '../lib/patrolSlots.js';

const DEFAULT_ROW = {
  temp: '',
  bpU: '',
  bpL: '',
  pulse: '',
  weight: '',
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
  mealAmount: '',
  waterMl: '',
  medicationTaken: '',
  toiletGuidance: false,
};

/**
 * @param {{
 *   filteredResidents: Record<string, unknown>[];
 *   bulkDraft: Record<string, typeof DEFAULT_ROW>;
 *   bulkGlobalMealSlot: string;
 *   onBulkGlobalMealSlotChange: (slot: string) => void;
 *   residentNameWithoutSama: (nameRaw: unknown) => string;
 *   patchBulkRow: (id: string, patch: Partial<typeof DEFAULT_ROW>) => void;
 *   setBulkPatrolForAllVisible: (checked: boolean) => void;
 *   bulkRowHasInput: (row: Partial<typeof DEFAULT_ROW> | undefined) => boolean;
 *   saveBulkRow: (res: Record<string, unknown>) => void;
 *   saveBulkAllWithInput: () => void;
 * }} props
 */
export function ResidentBulkInputTable({
  filteredResidents,
  bulkDraft,
  bulkGlobalMealSlot,
  onBulkGlobalMealSlotChange,
  residentNameWithoutSama,
  patchBulkRow,
  setBulkPatrolForAllVisible,
  bulkRowHasInput,
  saveBulkRow,
  saveBulkAllWithInput,
}) {
  return (
    <div className="min-w-0 pb-4">
      <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
        <p className="flex items-center gap-1.5 text-base font-black text-emerald-800">
          <Table2 className="h-4 w-4 shrink-0" aria-hidden />
          バイタル・体重（月1回）・巡視・排尿・排便・食事（朝昼夜）・水分・内服を一覧から
        </p>
        <div className="flex flex-wrap items-center gap-1.5">
          <button
            type="button"
            onClick={() => setBulkPatrolForAllVisible(true)}
            className="rounded-lg border border-cyan-600 bg-cyan-50 px-3 py-1.5 text-sm font-black text-cyan-900 hover:bg-cyan-100"
          >
            巡視を全員ON
          </button>
          <button
            type="button"
            onClick={() => setBulkPatrolForAllVisible(false)}
            className="rounded-lg border border-slate-500 bg-slate-50 px-3 py-1.5 text-sm font-black text-slate-700 hover:bg-slate-100"
          >
            巡視を全員OFF
          </button>
          <button
            type="button"
            onClick={saveBulkAllWithInput}
            className="rounded-xl border-2 border-emerald-600 bg-emerald-50 px-4 py-2.5 text-sm font-black text-emerald-900 hover:bg-emerald-100"
          >
            入力した行をまとめて保存
          </button>
        </div>
      </div>
      <p className="mb-2 text-sm font-bold leading-snug text-slate-500">
        下の表では<strong>主食・副食の割</strong>だけ行ごとに入力します（食事区分は上で統一）。「排」「食」は簡易確認。訪問看護・特別指示の手動登録は、ヒヤリ周知パネルで「看護事務メニューを表示」を ON にした端末のクイック記録のみ。
      </p>
      <div className="mb-3 flex flex-wrap items-center gap-2 rounded-2xl border-2 border-orange-300 bg-gradient-to-r from-orange-50 to-amber-50 px-3 py-2.5 shadow-sm">
        <span className="text-sm font-black text-orange-950 sm:text-base">今回の食事区分（全員共通）</span>
        <div className="flex flex-wrap gap-1.5" role="group" aria-label="食事区分 朝昼夜">
          {['朝', '昼', '夜'].map((slot) => {
            const on = bulkGlobalMealSlot === slot;
            return (
              <button
                key={slot}
                type="button"
                onClick={() => onBulkGlobalMealSlotChange(slot)}
                className={`min-w-[3.5rem] rounded-xl border-2 px-4 py-2 text-sm font-black transition sm:min-w-[4rem] sm:px-5 sm:text-base ${
                  on
                    ? 'border-orange-600 bg-orange-500 text-white shadow-md ring-2 ring-orange-400/80'
                    : 'border-orange-200 bg-white text-orange-900 hover:border-orange-400 hover:bg-orange-100/80'
                }`}
              >
                {slot}
              </button>
            );
          })}
        </div>
      </div>
      <div className="max-h-[min(70vh,720px)] overflow-auto rounded-xl border border-slate-200 shadow-inner">
        <table className="w-full min-w-[1860px] border-collapse text-left text-sm sm:text-base">
          <thead className="sticky top-0 z-10 bg-slate-100 text-sm font-black uppercase text-slate-700 sm:text-sm">
            <tr>
              <th className="border border-slate-200 px-0.5 py-1 text-slate-800">氏名</th>
              <th className="border border-slate-200 px-0.5 py-1">部屋</th>
              <th className="border border-slate-200 px-0.5 py-1">体温</th>
              <th className="border border-slate-200 px-0.5 py-1">上</th>
              <th className="border border-slate-200 px-0.5 py-1">下</th>
              <th className="border border-slate-200 px-0.5 py-1">脈</th>
              <th className="border border-slate-200 px-0.5 py-1 whitespace-nowrap text-teal-900" title="月1回の体重測定">
                体重kg<span className="block text-[9px] font-bold normal-case">月1回</span>
              </th>
              <th className="border border-slate-200 px-0.5 py-1">巡</th>
              <th className="border border-slate-200 px-0.5 py-1">巡視(3h)</th>
              <th className="border border-slate-200 px-0.5 py-1">排尿量</th>
              <th
                className="border border-slate-200 px-0.5 py-1 whitespace-nowrap text-sky-900"
                title="トイレ誘導を実施したらチェック（6時間アラートの基準を更新）"
              >
                誘導
              </th>
              <th className="border border-slate-200 px-0.5 py-1">排便量</th>
              <th className="border border-slate-200 px-0.5 py-1">性状</th>
              <th className="border border-slate-200 px-0.5 py-1">排※</th>
              <th className="border border-slate-200 bg-orange-50 px-0.5 py-1 whitespace-nowrap text-orange-950" title="上の「朝・昼・夜」が保存時に入ります">
                主食<span className="block text-[9px] font-bold normal-case">（区分は上）</span>
              </th>
              <th className="border border-slate-200 px-0.5 py-1">副食</th>
              <th className="border border-slate-200 px-0.5 py-1">水分ml</th>
              <th className="border border-slate-200 px-0.5 py-1">内服</th>
              <th className="border border-slate-200 px-0.5 py-1">食※</th>
              <th className="border border-slate-200 px-0.5 py-1"> </th>
            </tr>
          </thead>
          <tbody>
            {filteredResidents.map((res) => {
              const id = String(res.id);
              const nm = residentNameWithoutSama(res.name);
              const row = { ...DEFAULT_ROW, ...bulkDraft[id] };
              const pa = splitPatrolDateTimeLocal(row.patrolAt);
              return (
                <tr key={id} className="odd:bg-white even:bg-slate-50/50">
                  <td className="border border-slate-200 px-1 py-1 font-bold text-slate-900">{nm}</td>
                  <td className="border border-slate-200 px-1 py-1 text-center font-mono">{String(res.room ?? '')}</td>
                  <td className="border border-slate-200 p-0">
                    <input
                      value={row.temp}
                      onChange={(e) => patchBulkRow(id, { temp: e.target.value })}
                      inputMode="decimal"
                      className="w-full min-w-[2.75rem] bg-transparent px-1 py-1.5 font-mono text-sm font-bold sm:text-base"
                      aria-label={`${nm} 体温`}
                    />
                  </td>
                  <td className="border border-slate-200 p-0">
                    <input
                      value={row.bpU}
                      onChange={(e) => patchBulkRow(id, { bpU: e.target.value })}
                      inputMode="numeric"
                      className="w-full min-w-[2.25rem] bg-transparent px-1 py-1.5 font-mono text-sm sm:text-base"
                      aria-label={`${nm} 血圧上`}
                    />
                  </td>
                  <td className="border border-slate-200 p-0">
                    <input
                      value={row.bpL}
                      onChange={(e) => patchBulkRow(id, { bpL: e.target.value })}
                      inputMode="numeric"
                      className="w-full min-w-[2.25rem] bg-transparent px-1 py-1.5 font-mono text-sm sm:text-base"
                      aria-label={`${nm} 血圧下`}
                    />
                  </td>
                  <td className="border border-slate-200 p-0">
                    <input
                      value={row.pulse}
                      onChange={(e) => patchBulkRow(id, { pulse: e.target.value })}
                      inputMode="numeric"
                      className="w-full min-w-[2rem] bg-transparent px-1 py-1.5 font-mono text-sm sm:text-base"
                      aria-label={`${nm} 脈`}
                    />
                  </td>
                  <td className="border border-slate-200 bg-teal-50/40 p-0">
                    <input
                      value={row.weight}
                      onChange={(e) => patchBulkRow(id, { weight: e.target.value })}
                      inputMode="decimal"
                      placeholder="kg"
                      className="w-full min-w-[3rem] bg-transparent px-1 py-1.5 font-mono text-sm font-bold text-teal-950 sm:text-base"
                      aria-label={`${nm} 体重（月1回）`}
                    />
                  </td>
                  <td className="border border-slate-200 px-0.5 py-0 text-center">
                    <input
                      type="checkbox"
                      checked={row.patrol}
                      onChange={(e) => patchBulkRow(id, { patrol: e.target.checked })}
                      className="h-5 w-5 accent-cyan-600 sm:h-5 sm:w-5"
                      aria-label={`${nm} 巡視`}
                    />
                  </td>
                  <td className="border border-slate-200 p-0.5 align-top">
                    <div className="flex min-w-[8.5rem] flex-col gap-0.5">
                      <input
                        type="date"
                        value={pa.date}
                        onChange={(e) =>
                          patchBulkRow(id, { patrolAt: joinPatrolDateTimeLocal(e.target.value, pa.hour) })
                        }
                        className="w-full bg-white px-1 py-1 font-mono text-[11px] font-bold sm:text-xs"
                        aria-label={`${nm} 巡視の日付`}
                      />
                      <select
                        value={pa.hour}
                        onChange={(e) =>
                          patchBulkRow(id, {
                            patrolAt: joinPatrolDateTimeLocal(pa.date, Number(e.target.value)),
                          })
                        }
                        className="w-full bg-white px-1 py-1 text-[11px] font-bold sm:text-xs"
                        aria-label={`${nm} 巡視時刻（3時間おき）`}
                      >
                        {PATROL_SLOT_HOURS.map((h) => (
                          <option key={h} value={h}>
                            {String(h).padStart(2, '0')}:00
                          </option>
                        ))}
                      </select>
                    </div>
                  </td>
                  <td className="border border-slate-200 p-0">
                    <input
                      value={row.urineVolume}
                      onChange={(e) => patchBulkRow(id, { urineVolume: e.target.value })}
                      placeholder="ml等"
                      className="w-full min-w-[3rem] bg-transparent px-1 py-1.5 text-sm sm:text-base"
                      aria-label={`${nm} 排尿量`}
                    />
                  </td>
                  <td className="border border-slate-200 bg-sky-50/50 px-0.5 py-0 text-center">
                    <input
                      type="checkbox"
                      checked={Boolean(row.toiletGuidance)}
                      onChange={(e) => patchBulkRow(id, { toiletGuidance: e.target.checked })}
                      className="h-5 w-5 accent-sky-700 sm:h-5 sm:w-5"
                      title="トイレ誘導"
                      aria-label={`${nm} トイレ誘導`}
                    />
                  </td>
                  <td className="border border-slate-200 p-0">
                    <select
                      value={row.stoolVolume}
                      onChange={(e) => patchBulkRow(id, { stoolVolume: e.target.value })}
                      className="w-full min-w-[2.75rem] bg-white px-1 py-1.5 text-sm font-bold sm:text-base"
                      aria-label={`${nm} 排便量`}
                    >
                      {STOOL_VOLUME_OPTIONS.map((opt) => (
                        <option key={opt || 'empty'} value={opt}>
                          {opt || '—'}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td className="border border-slate-200 p-0">
                    <select
                      value={row.stoolCharacter}
                      onChange={(e) => patchBulkRow(id, { stoolCharacter: e.target.value })}
                      className="w-full min-w-[4rem] bg-white px-1 py-1.5 text-sm font-bold sm:text-base"
                      aria-label={`${nm} 排便性状`}
                    >
                      {STOOL_CHARACTER_OPTIONS.map((opt) => (
                        <option key={opt || 'empty'} value={opt}>
                          {opt || '—'}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td className="border border-slate-200 px-0.5 py-0 text-center">
                    <input
                      type="checkbox"
                      checked={row.excretion}
                      onChange={(e) => patchBulkRow(id, { excretion: e.target.checked })}
                      className="h-5 w-5 accent-amber-600 sm:h-5 sm:w-5"
                      title="排泄のみ確認（詳細なし）"
                      aria-label={`${nm} 排泄確認`}
                    />
                  </td>
                  <td className="border border-slate-200 bg-orange-50/40 p-0">
                    <select
                      value={row.mealStaple}
                      onChange={(e) => patchBulkRow(id, { mealStaple: e.target.value })}
                      className="w-full min-w-[3.25rem] bg-white px-1 py-1.5 text-sm font-bold sm:text-base"
                      aria-label={`${nm} 主食`}
                    >
                      {MEAL_WARI_OPTIONS.map((opt) => (
                        <option key={opt || 'st-empty'} value={opt}>
                          {opt || '—'}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td className="border border-slate-200 p-0">
                    <select
                      value={row.mealSide}
                      onChange={(e) => patchBulkRow(id, { mealSide: e.target.value })}
                      className="w-full min-w-[3.25rem] bg-white px-1 py-1.5 text-sm font-bold sm:text-base"
                      aria-label={`${nm} 副食`}
                    >
                      {MEAL_WARI_OPTIONS.map((opt) => (
                        <option key={`${opt}-side`} value={opt}>
                          {opt || '—'}
                        </option>
                      ))}
                    </select>
                  </td>
                  <td className="border border-slate-200 p-0">
                    <input
                      value={row.waterMl}
                      onChange={(e) => patchBulkRow(id, { waterMl: e.target.value })}
                      inputMode="numeric"
                      placeholder="ml"
                      className="w-full min-w-[2.5rem] bg-transparent px-1 py-1.5 font-mono text-sm sm:text-base"
                      aria-label={`${nm} 水分量`}
                    />
                  </td>
                  <td className="border border-slate-200 p-0">
                    <select
                      value={row.medicationTaken}
                      onChange={(e) => patchBulkRow(id, { medicationTaken: e.target.value })}
                      className="w-full min-w-[3.5rem] bg-white px-1 py-1.5 text-sm font-bold sm:text-base"
                      aria-label={`${nm} 内服`}
                    >
                      <option value="">—</option>
                      <option value="yes">飲了</option>
                      <option value="no">未服</option>
                    </select>
                  </td>
                  <td className="border border-slate-200 px-0.5 py-0 text-center">
                    <input
                      type="checkbox"
                      checked={row.meal}
                      onChange={(e) => patchBulkRow(id, { meal: e.target.checked })}
                      className="h-5 w-5 accent-orange-500 sm:h-5 sm:w-5"
                      title="食事のみ確認（詳細なし）"
                      aria-label={`${nm} 食事確認`}
                    />
                  </td>
                  <td className="border border-slate-200 px-0.5 py-0.5 text-center">
                    <button
                      type="button"
                      disabled={!bulkRowHasInput(row)}
                      onClick={() => saveBulkRow(res)}
                      className="rounded bg-emerald-600 px-3 py-1.5 text-sm font-black text-white disabled:opacity-40 sm:text-base"
                    >
                      保存
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
