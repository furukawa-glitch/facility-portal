/**
 * 求人・入退社・スタッフ名簿（1 スプレッドシートで運用）
 * デフォルト ID: ユーザー指定の求人シート
 */

/** @see https://docs.google.com/spreadsheets/d/1CQBElPFLr4vEAaKuYbGwUESFpSEcsvrSZnc_LK88Oqk */
export const DEFAULT_HR_SPREADSHEET_ID = '1CQBElPFLr4vEAaKuYbGwUESFpSEcsvrSZnc_LK88Oqk';

/** @see https://docs.google.com/spreadsheets/d/1iCVPq0-9JeK11mc3-d-YF9nOIAJq12_4A5vtxVnYbjY */
export const DEFAULT_AWARENESS_SPREADSHEET_ID = '1iCVPq0-9JeK11mc3-d-YF9nOIAJq12_4A5vtxVnYbjY';

/** 周知確認ログの追記先（ヒヤリハット/事故を共通管理） */
export const AWARENESS_LOG_SHEET_NAME = '周知確認ログ';

export const AWARENESS_LOG_COLS = Object.freeze([
  '日時',
  '種別',
  'タイトル',
  '名前',
  '部署',
  '施設',
]);
