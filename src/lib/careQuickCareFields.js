/** 排便量（クイック／一覧入力のプルダウン） */
export const STOOL_VOLUME_OPTIONS = ['', '多', '中', '小'];

/** 排便性状 */
export const STOOL_CHARACTER_OPTIONS = ['', '普通便', '硬便', '軟便', '水様便'];

/** 主食・副食の摂取割合 */
export const MEAL_WARI_OPTIONS = ['', '10割', '9割', '8割', '7割', '6割', '5割', '4割', '3割', '2割', '1割', '0割'];

function normVoiceChars(s) {
  return String(s ?? '')
    .trim()
    .replace(/\s/g, '')
    .replace(/[０-９]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xff10 + 0x30));
}

/**
 * @param {string} staple
 * @param {string} side
 * @returns {string} ログ用 1 行（例: 主食8割 副食7割）
 */
export function composeMealAmountForLog(staple, side) {
  const s = String(staple ?? '').trim();
  const d = String(side ?? '').trim();
  const parts = [];
  if (s) parts.push(`主食${s}`);
  if (d) parts.push(`副食${d}`);
  return parts.join(' ');
}

/** @param {string} text 音声認識結果 */
export function parseVoiceToStoolVolume(text) {
  const raw = String(text ?? '').trim();
  if (!raw) return '';
  if (STOOL_VOLUME_OPTIONS.includes(raw)) return raw;
  const n = normVoiceChars(raw);
  if (/(多|大|おお|だい|ダイ)/u.test(n)) return '多';
  if (/(中|ちゅう|チュウ|なか)/u.test(n)) return '中';
  if (/(小|しょう|ショウ|すくない)/u.test(n)) return '小';
  return '';
}

/** @param {string} text */
export function parseVoiceToStoolCharacter(text) {
  const raw = String(text ?? '').trim();
  if (!raw) return '';
  if (STOOL_CHARACTER_OPTIONS.includes(raw)) return raw;
  const n = normVoiceChars(raw);
  if (/水様|みずよう|スイヨ/u.test(n)) return '水様便';
  if (/硬便|こうべん|カタ|硬い/u.test(n)) return '硬便';
  if (/軟便|なんべん|ナン|やわらか|軟か/u.test(n)) return '軟便';
  if (/普通便|ふつうべん|ふつう|フツウ|普通/u.test(n)) return '普通便';
  for (const opt of STOOL_CHARACTER_OPTIONS) {
    if (opt && n.includes(opt)) return opt;
  }
  return '';
}

/** @param {string} text */
export function parseVoiceToMealWari(text) {
  const raw = String(text ?? '').trim();
  if (!raw) return '';
  if (MEAL_WARI_OPTIONS.includes(raw)) return raw;
  const n = normVoiceChars(raw);
  const m = n.match(/(\d{1,2})\s*割/u);
  if (m) {
    const v = Math.min(10, Math.max(0, parseInt(m[1], 10)));
    return `${v}割`;
  }
  const spokenWari = [
    [/いちわり|イチワリ|一割/u, '1割'],
    [/にわり|ニワリ|二割/u, '2割'],
    [/さんわり|サンワリ|三割/u, '3割'],
    [/よんわり|ヨンワリ|四割/u, '4割'],
    [/ごわり|ゴワリ|五割/u, '5割'],
    [/ろくわり|ロクワリ|六割/u, '6割'],
    [/ななわり|ナナワリ|シチワリ|七割/u, '7割'],
    [/はちわり|ハチワリ|八割/u, '8割'],
    [/きゅうわり|キュウワリ|九割/u, '9割'],
    [/じゅうわり|ジュウワリ|十割/u, '10割'],
    [/れいわり|レイワリ|ゼロわり/u, '0割'],
  ];
  for (const [re, val] of spokenWari) {
    if (re.test(n)) return val;
  }
  const jpNum = {
    零: 0,
    〇: 0,
    一: 1,
    二: 2,
    三: 3,
    四: 4,
    五: 5,
    六: 6,
    七: 7,
    八: 8,
    九: 9,
    十: 10,
  };
  for (const [ch, val] of Object.entries(jpNum)) {
    if (n.includes(`${ch}割`) || n.includes(`${ch}わり`)) return `${val}割`;
  }
  if (/全量|全部|ぜんぶ|ぜんりょう|満腹|マン|100パー|100%/u.test(n)) return '10割';
  if (/ゼロ|れい|未摂|無し|なし|0割|食べず|食べていない/u.test(n)) return '0割';
  return '';
}

/** 水分 ml 用（数字を拾う） */
export function parseVoiceToWaterMl(text) {
  const n = normVoiceChars(String(text ?? ''));
  const m = n.match(/(\d{2,4})/);
  if (m) return m[1];
  const m2 = n.match(/(\d+)/);
  return m2 ? m2[1] : '';
}
