/**
 * 장안 골프 동호회 지표 대시보드 - Google Apps Script 백엔드
 * ============================================================
 * 스크린골프 · 코스라운드 · 선수명단, 3가지 데이터를 하나의 웹앱 URL로 처리합니다.
 * index.html 안의 SHEET_API_URL 값은 그대로 두고, 이 스크립트만 기존 프로젝트에
 * 덮어써서 다시 배포하면 됩니다 (웹앱 URL은 바뀌지 않습니다).
 *
 * ------------------------------------------------------------
 * [배포 방법]
 * 1. 기존에 SHEET_API_URL을 만들 때 사용했던 구글 스프레드시트를 엽니다.
 * 2. 상단 메뉴 "확장 프로그램 > Apps Script"로 들어갑니다.
 * 3. 기존 코드를 전부 지우고 이 파일 내용 전체를 붙여넣습니다.
 * 4. 저장(Ctrl+S) 후, 우측 상단 "배포 > 배포 관리"로 들어갑니다.
 * 5. 배포 목록에서 연필(편집) 아이콘 클릭 → "버전"을 "새 버전"으로 선택 → "배포".
 *    (⚠️ "새 배포"가 아니라 기존 배포를 "편집"해야 웹앱 URL이 그대로 유지됩니다.)
 * 6. 완료되면 시트 파일에 아래 탭들이 사용됩니다:
 *    - Rounds       : date, player, gHandicap, fir, gir, putt, distance  (기존 스크린골프 데이터 탭)
 *    - CourseRounds : date, player, handicap  (지표를 추가하면 열이 자동으로 늘어납니다. 신규 생성)
 *    - Players      : name  (신규 생성)
 *
 * 스크린골프 기존 데이터 탭 이름은 SHEET_NAMES.screenGolf 값과 반드시 일치해야 합니다.
 * (장안 골프 동호회 시트 기준 실제 탭 이름은 "Rounds" 입니다 — 이미 반영되어 있습니다.)
 * 만약 이전에 재배포하면서 "ScreenGolf"라는 빈 탭이 자동으로 하나 생겼다면,
 * 그 빈 탭은 사용되지 않으니 삭제하셔도 됩니다.
 * ------------------------------------------------------------
 */

const SHEET_NAMES = {
  screenGolf: 'Rounds',
  courseRounds: 'CourseRounds',
  players: 'Players'
};

const DEFAULT_HEADERS = {
  screenGolf: ['date', 'player', 'gHandicap', 'fir', 'gir', 'putt', 'distance'],
  courseRounds: ['date', 'player', 'handicap'],
  players: ['name']
};

function doGet(e) {
  const type = (e && e.parameter && e.parameter.type) || 'screenGolf';
  const sheet = getOrCreateSheet(type);
  const rows = sheetToObjects(sheet);
  return jsonResponse(rows);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const type = body.type || 'screenGolf';

  if (type === 'players') {
    savePlayers(body.players || []);
  } else {
    saveRoundRecords(type, body.date, body.records || []);
  }
  return jsonResponse({ status: 'ok' });
}

function getOrCreateSheet(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEET_NAMES[type] || SHEET_NAMES.screenGolf;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = DEFAULT_HEADERS[type] || DEFAULT_HEADERS.screenGolf;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(cell => cell === '' || cell === null)) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    rows.push(obj);
  }
  return rows;
}

// date + player 조합이 이미 있으면 해당 행을 덮어쓰고, 없으면 새 행을 추가합니다.
// records 안에 시트에 없는 새로운 필드(지표)가 있으면 열을 자동으로 추가합니다.
// -> 프론트엔드 COURSE_METRICS에 지표를 추가해도 이 스크립트는 수정할 필요가 없습니다.
function saveRoundRecords(type, date, records) {
  const sheet = getOrCreateSheet(type);
  ensureColumns(sheet, records);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const dateCol = headers.indexOf('date');
  const playerCol = headers.indexOf('player');

  records.forEach(record => {
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][dateCol]) === String(date) && String(values[i][playerCol]) === String(record.player)) {
        rowIndex = i;
        break;
      }
    }
    const rowValues = headers.map(h => {
      if (h === 'date') return date;
      if (h === 'player') return record.player;
      return (record[h] !== undefined && record[h] !== null) ? record[h] : '';
    });
    if (rowIndex === -1) {
      sheet.appendRow(rowValues);
      values.push(rowValues); // 같은 요청 안에서의 중복 추가 방지
    } else {
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([rowValues]);
    }
  });
}

// records에는 있는데 시트 헤더에는 없는 필드(새 지표)가 있으면 새 열로 추가합니다.
function ensureColumns(sheet, records) {
  let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
    .map(h => String(h).trim()).filter(h => h !== '');
  if (headers.length === 0) {
    headers = ['date', 'player'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const existing = new Set(headers);
  const newKeys = new Set();
  records.forEach(record => {
    Object.keys(record).forEach(k => {
      if (k !== 'player' && !existing.has(k)) newKeys.add(k);
    });
  });
  if (newKeys.size > 0) {
    const toAdd = Array.from(newKeys);
    sheet.getRange(1, headers.length + 1, 1, toAdd.length).setValues([toAdd]);
  }
}

function savePlayers(players) {
  const sheet = getOrCreateSheet('players');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 1).setValues([['name']]);
  if (players.length > 0) {
    sheet.getRange(2, 1, players.length, 1).setValues(players.map(p => [p]));
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
