/**
 * 施設ごとの外部アプリ URL
 * キーは carelinkFacilities.js の linkKey と完全一致（スペース・表記ゆれ不可）
 *
 * 公式LINE: 各施設の LINE公式アカウント管理画面（LINE Official Account Manager）で
 * 「ホーム」→ 友だち追加用の URL（https://line.me/R/ti/p/@xxxx や https://lin.ee/...）をコピーして line に貼る。
 * 未設定の施設は VITE_LINK_LINE_DEFAULT（.env）を使います。
 */

/** @type {Record<string, { kaipoke?: string; mcs?: string; line?: string }>} */
export const FACILITY_EXTERNAL_LINKS = {
  中川本館: {
    // line: 'https://line.me/R/ti/p/@xxxxxxxx',
  },
  愛西: {},
  北名古屋: {},
  千音寺: {},
  中村: {},
 '起': {
    line: 'https://line.me/R/ti/p/@732wunij',
  },
  一宮: {},
};

function firstEnv(...keys) {
  for (const k of keys) {
    const v = import.meta.env[k];
    if (v != null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

const DEFAULTS = {
  kaipoke: firstEnv('VITE_LINK_KAIPOKE_DEFAULT'),
  mcs: firstEnv('VITE_LINK_MCS_DEFAULT'),
  line: firstEnv('VITE_LINK_LINE_DEFAULT'),
};

/**
 * @param {string} facilityName
 * @returns {{ kaipoke: string; mcs: string; line: string; label: string }}
 */
export function getExternalLinksForFacility(facilityName) {
  const name = String(facilityName ?? '').trim();
  const ov = name ? FACILITY_EXTERNAL_LINKS[name] : undefined;
  return {
    label: name || '施設未選択',
    kaipoke: ov?.kaipoke?.trim() || DEFAULTS.kaipoke || '#',
    mcs: ov?.mcs?.trim() || DEFAULTS.mcs || '#',
    line: ov?.line?.trim() || DEFAULTS.line || '#',
  };
}
