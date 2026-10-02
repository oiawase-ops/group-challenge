/**
 * グループチャレンジ - データの保管・やり取りを行う裏方の仕組み。
 * このファイルをGoogle Apps Scriptのエディタに貼り付けて使う。
 *
 * 使い方（初回のみ）:
 * 1. ADMIN_PASSWORD を好きな文字列に変更する
 * 2. 関数選択で initializeSheets を選び、実行する（シートの土台を作る）
 * 3. 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 *    でデプロイし、発行されたURLを控える
 */

const ADMIN_PASSWORD = 'ここにパスワードを設定してください';

const SHEET_CONFIG = 'config';
const SHEET_PARTICIPANTS = 'participants';
const TIMEZONE = 'Asia/Tokyo';

const MAX_NICKNAME_LEN = 20;
const MAX_GOAL_LEN = 200;

function doGet(e) {
  return handle(e);
}

function doPost(e) {
  return handle(e);
}

function handle(e) {
  try {
    const params = parseParams(e);
    const action = String(params.action || '');
    let result;
    switch (action) {
      case 'state':
        result = getState();
        break;
      case 'join':
        result = join(params.nickname);
        break;
      case 'checkin':
        result = checkin(params.nickname);
        break;
      case 'setGoal':
        result = setGoal(params.password, params.text);
        break;
      default:
        result = { ok: false, error: 'unknown_action' };
    }
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

/** GETのクエリパラメータと、POST(text/plain)の本文の両方を読めるようにする */
function parseParams(e) {
  const params = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (k) {
      params[k] = e.parameter[k];
    });
  }
  if (e && e.postData && e.postData.contents) {
    e.postData.contents.split('&').forEach(function (pair) {
      const idx = pair.indexOf('=');
      if (idx === -1) return;
      const k = decodeURIComponent(pair.slice(0, idx));
      const v = decodeURIComponent(pair.slice(idx + 1).replace(/\+/g, ' '));
      params[k] = v;
    });
  }
  return params;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

/** 他人からは分からない匿名ID（並び順にも使う） */
function hashId(s) {
  let h = 0;
  for (let i = 0; i < s.length; i++) {
    h = (h * 31 + s.charCodeAt(i)) | 0;
  }
  return (h >>> 0).toString(36);
}

function normalizeName(name) {
  return String(name || '')
    .normalize('NFKC')
    .trim()
    .replace(/\s+/g, ' ')
    .slice(0, MAX_NICKNAME_LEN);
}

function todayStr() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}
function configSheet() {
  return ss().getSheetByName(SHEET_CONFIG);
}
function participantsSheet() {
  return ss().getSheetByName(SHEET_PARTICIPANTS);
}

/** 初回に1度だけ実行する。シートの土台（見出し・最初の目標）を作る */
function initializeSheets() {
  const spreadsheet = ss();

  let c = spreadsheet.getSheetByName(SHEET_CONFIG);
  if (!c) c = spreadsheet.insertSheet(SHEET_CONFIG);
  c.clear();
  c.getRange(1, 1, 1, 3).setValues([['goalText', 'periodId', 'updatedAt']]);
  c.getRange(2, 1, 1, 3).setValues([['まだ目標が決まっていません', 1, new Date()]]);

  let p = spreadsheet.getSheetByName(SHEET_PARTICIPANTS);
  if (!p) p = spreadsheet.insertSheet(SHEET_PARTICIPANTS);
  p.clear();
  p.getRange(1, 1, 1, 5).setValues([
    ['nickname', 'periodId', 'count', 'lastDoneDate', 'createdAt'],
  ]);
}

function getConfig() {
  const row = configSheet().getRange(2, 1, 1, 3).getValues()[0];
  return {
    goalText: row[0] || '',
    periodId: Number(row[1]) || 1,
    updatedAt: row[2],
  };
}

function getState() {
  const config = getConfig();
  const sheet = participantsSheet();
  const rows = sheet.getDataRange().getValues().slice(1);
  const participants = rows
    .filter(function (r) {
      return Number(r[1]) === config.periodId;
    })
    .map(function (r) {
      return { id: hashId(String(r[0])), count: Number(r[2]) || 0 };
    });
  return {
    ok: true,
    goalText: config.goalText,
    periodId: config.periodId,
    participants: participants,
  };
}

function findRowIndex(rows, nickname, periodId) {
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === nickname && Number(rows[i][1]) === periodId) return i;
  }
  return -1;
}

function join(nicknameRaw) {
  const nickname = normalizeName(nicknameRaw);
  if (!nickname) return { ok: false, error: 'empty_nickname' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = getConfig();
    const sheet = participantsSheet();
    const rows = sheet.getDataRange().getValues().slice(1);
    const idx = findRowIndex(rows, nickname, config.periodId);

    if (idx === -1) {
      sheet.appendRow([nickname, config.periodId, 0, '', new Date()]);
      SpreadsheetApp.flush();
      return {
        ok: true,
        nickname: nickname,
        count: 0,
        already: false,
        goalText: config.goalText,
        periodId: config.periodId,
      };
    }

    const today = todayStr();
    return {
      ok: true,
      nickname: nickname,
      count: Number(rows[idx][2]) || 0,
      already: rows[idx][3] === today,
      goalText: config.goalText,
      periodId: config.periodId,
    };
  } finally {
    lock.releaseLock();
  }
}

function checkin(nicknameRaw) {
  const nickname = normalizeName(nicknameRaw);
  if (!nickname) return { ok: false, error: 'empty_nickname' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = getConfig();
    const sheet = participantsSheet();
    const rows = sheet.getDataRange().getValues().slice(1);
    const idx = findRowIndex(rows, nickname, config.periodId);
    const today = todayStr();

    if (idx === -1) {
      sheet.appendRow([nickname, config.periodId, 1, today, new Date()]);
      SpreadsheetApp.flush();
      return { ok: true, count: 1, already: false };
    }

    if (rows[idx][3] === today) {
      return { ok: true, count: Number(rows[idx][2]) || 0, already: true };
    }

    const rowNum = idx + 2; // 見出し行(+1) + 1始まり(+1)
    const newCount = (Number(rows[idx][2]) || 0) + 1;
    sheet.getRange(rowNum, 3).setValue(newCount);
    sheet.getRange(rowNum, 4).setValue(today);
    SpreadsheetApp.flush();
    return { ok: true, count: newCount, already: false };
  } finally {
    lock.releaseLock();
  }
}

function setGoal(password, textRaw) {
  if (String(password || '') !== ADMIN_PASSWORD) {
    return { ok: false, error: 'wrong_password' };
  }
  const text = String(textRaw || '').trim().slice(0, MAX_GOAL_LEN);
  if (!text) return { ok: false, error: 'empty_text' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const current = getConfig();
    const newPeriodId = current.periodId + 1;
    configSheet().getRange(2, 1, 1, 3).setValues([[text, newPeriodId, new Date()]]);

    // 今までに参加した全員分の行を、新しい期間用に先に作っておく
    // （こうしないと切り替えた直後、一覧が誰もいないように見えてしまう）
    const sheet = participantsSheet();
    const rows = sheet.getDataRange().getValues().slice(1);
    const seen = {};
    rows.forEach(function (r) {
      seen[r[0]] = true;
    });
    const names = Object.keys(seen);
    if (names.length > 0) {
      const newRows = names.map(function (n) {
        return [n, newPeriodId, 0, '', new Date()];
      });
      sheet
        .getRange(sheet.getLastRow() + 1, 1, newRows.length, 5)
        .setValues(newRows);
    }
    SpreadsheetApp.flush();
    return { ok: true, goalText: text, periodId: newPeriodId };
  } finally {
    lock.releaseLock();
  }
}
