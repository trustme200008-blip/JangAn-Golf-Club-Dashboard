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
 *    - Rounds         : date, player, gHandicap, fir, gir, putt, distance  (기존 스크린골프 데이터 탭.
 *      날짜는 "MM/DD"만 저장합니다 — 스크린골프는 연도 선택 기능이 없습니다)
 *    - CourseRounds   : date, player, handicap  (지표를 추가하면 열이 자동으로 늘어납니다.
 *      코스라운드는 연도 선택 기능이 있어 날짜를 "YYYY-MM-DD" 전체로 저장합니다 — 예전에
 *      "MM/DD"로만 저장됐던 기록은 읽거나 쓸 때 자동으로 2026년 기록으로 치유됩니다)
 *    - Players        : name  (신규 생성)
 *    - ScreenScoreCard: id, date, player, course, score  (스크린골프 스코어 카드, 신규 생성.
 *      날짜는 "YYYY-MM-DD" 전체 날짜를 그대로 저장합니다 — Rounds 탭의 "MM/DD"와 다르니 주의)
 *
 * 스크린골프 기존 데이터 탭 이름은 SHEET_NAMES.screenGolf 값과 반드시 일치해야 합니다.
 * (장안 골프 동호회 시트 기준 실제 탭 이름은 "Rounds" 입니다 — 이미 반영되어 있습니다.)
 * 만약 이전에 재배포하면서 "ScreenGolf"라는 빈 탭이 자동으로 하나 생겼다면,
 * 그 빈 탭은 사용되지 않으니 삭제하셔도 됩니다.
 * ------------------------------------------------------------
 */

// 관리자 쓰기(추가/수정/삭제) 요청마다 이 토큰이 body.token으로 함께 와야 처리됩니다.
// index.html의 ADMIN_PIN과 반드시 같은 값이어야 합니다 — PIN을 바꾸면 이 값도 같이 바꾸고
// 재배포해주세요. 조회(doGet)는 토큰 없이도 계속 누구나 가능합니다(화면 표시용이라 필요).
const ADMIN_TOKEN = '0243';

const SHEET_NAMES = {
  screenGolf: 'Rounds',
  courseRounds: 'CourseRounds',
  players: 'Players',
  screenScoreCard: 'ScreenScoreCard'
};

const DEFAULT_HEADERS = {
  screenGolf: ['date', 'player', 'gHandicap', 'fir', 'gir', 'putt', 'distance'],
  courseRounds: ['date', 'player', 'handicap'],
  players: ['name'],
  screenScoreCard: ['id', 'date', 'player', 'course', 'score', 'overPar']
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

  // 쓰기 요청은 전부 토큰 검증을 거칩니다. 프론트엔드가 관리자 PIN을 확인한 경우에만
  // 이 토큰을 실어 보내므로, PIN을 모르는 상태로 API를 직접 호출해 데이터를 추가/삭제하는
  // 것을 막습니다. (완전한 인증은 아니지만, 화면의 관리자 PIN 잠금을 서버 쪽에서도 강제합니다.)
  if (body.token !== ADMIN_TOKEN) {
    return jsonResponse({ status: 'error', message: '인증 실패: 관리자 토큰이 올바르지 않습니다.' });
  }

  if (body.action === 'delete') {
    if (type === 'screenScoreCard') {
      deleteScoreCardRecord(body.id);
    } else {
      deleteRoundRecords(type, body.date);
    }
  } else if (type === 'players') {
    savePlayers(body.players || []);
  } else if (type === 'screenScoreCard') {
    appendScoreCardRecords(body.records || []);
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
    // date 열은 항상 "일반 텍스트" 서식으로 고정해서, 구글 시트가 "01/01" 같은 값을
    // 실제 날짜 타입으로 자동 변환해버리는 것을 애초에 막습니다.
    // (date 열이 항상 1번 열이라고 가정하지 않고 헤더에서 실제 위치를 찾습니다 —
    // ScreenScoreCard는 id가 1열, date가 2열입니다.)
    const dateColIdx = headers.indexOf('date');
    if (dateColIdx !== -1) {
      sheet.getRange(2, dateColIdx + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
    }
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

// 구글 시트는 "01/01" 같은 문자열을 셀에 쓰면 스스로 실제 날짜(Date) 타입으로 자동 변환해버리는
// 경우가 있습니다. 이러면 그 셀 값을 String()으로 그대로 비교했을 때 우리가 보낸 "01/01"과
// 더 이상 일치하지 않아 수정/삭제 매칭이 실패합니다. 그래서 비교 전에 항상 이 함수로
// "MM/dd" 형태의 순수 문자열로 정규화한 뒤 비교합니다.
// (아래 ensureDateColumnAsText가 date 열을 텍스트로 고정/치유하므로 Date 분기는
// 혹시 남아있을 수 있는 기존 값을 위한 방어 코드입니다. 프론트엔드의 normalizeSheetDate와
// 동일하게 UTC 기준으로 월/일을 추출해 서로 어긋나지 않게 맞춥니다.)
function normalizeDateCell(value) {
  if (value instanceof Date) {
    const mm = ('0' + (value.getUTCMonth() + 1)).slice(-2);
    const dd = ('0' + value.getUTCDate()).slice(-2);
    return mm + '/' + dd;
  }
  let str = String(value).trim();
  if (str.charAt(0) === "'") str = str.slice(1); // 혹시 아포스트로피가 값에 그대로 남아있는 경우 방어
  const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[2]}/${isoMatch[3]}`;
  return str;
}

// date 열을 텍스트 서식으로 고정하고, 이미 실제 날짜 타입으로 저장돼버린 셀이 있으면
// 순수 문자열로 다시 써서 치유합니다. 저장/삭제할 때마다 먼저 실행합니다.
// 주의: 서식은 "현재 데이터가 있는 행"뿐 아니라 시트의 최대 행(sheet.getMaxRows(), 기본 1000행)까지
// 미리 걸어둬야 합니다. 그래야 이 함수 실행 시점 이후에 새로 append되는 행도(그 시점엔 아직
// 값이 없어 이 함수가 못 봄) 자동으로 날짜 타입으로 바뀌는 걸 막을 수 있습니다.
// normalizeFn은 기본값 normalizeDateCell("MM/dd")이며, ScreenScoreCard/CourseRounds처럼
// 연도까지 보존해야 하는 시트는 normalizeFullDateCell 계열을 넘겨줍니다.
// migrateLegacyText를 true로 넘기면, 이미 텍스트로 저장된 값도 normalizeFn을 거쳐 형식이
// 바뀌면 다시 써서 치유합니다 (CourseRounds가 "MM/DD"만 쓰던 시절 데이터를
// "YYYY-MM-DD"로 옮길 때 씁니다). 기본은 false로, 기존 시트들의 동작은 그대로 유지됩니다.
function ensureDateColumnAsText(sheet, headers, normalizeFn, migrateLegacyText) {
  normalizeFn = normalizeFn || normalizeDateCell;
  const dateCol = headers.indexOf('date');
  if (dateCol === -1) return;

  const totalRows = sheet.getMaxRows() - 1;
  if (totalRows > 0) {
    sheet.getRange(2, dateCol + 1, totalRows, 1).setNumberFormat('@');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, dateCol + 1, lastRow - 1, 1);
  const cellValues = range.getValues();
  let changed = false;
  const fixed = cellValues.map(row => {
    const v = row[0];
    if (v instanceof Date) {
      changed = true;
      return [normalizeFn(v)];
    }
    if (migrateLegacyText && typeof v === 'string' && v) {
      const stripped = v.charAt(0) === "'" ? v.slice(1) : v;
      const norm = normalizeFn(stripped);
      if (norm !== stripped) {
        changed = true;
        return ["'" + norm];
      }
    }
    return [v];
  });
  if (changed) range.setValues(fixed);
}

// normalizeDateCell과 같은 목적이지만 "YYYY-MM-DD" 전체 날짜(연도 포함)를 보존합니다.
// ScreenScoreCard는 스크린골프 Rounds/CourseRounds와 달리 연도가 다른 기록이 섞일 수 있어
// (스코어 카드에 여러 해가 쌓일 수 있음) 연도를 버리는 normalizeDateCell을 쓰면 안 됩니다.
function normalizeFullDateCell(value) {
  if (value instanceof Date) {
    const yyyy = value.getUTCFullYear();
    const mm = ('0' + (value.getUTCMonth() + 1)).slice(-2);
    const dd = ('0' + value.getUTCDate()).slice(-2);
    return `${yyyy}-${mm}-${dd}`;
  }
  let str = String(value).trim();
  if (str.charAt(0) === "'") str = str.slice(1);
  return str;
}

// CourseRounds 전용 날짜 정규화. normalizeFullDateCell과 같이 연도를 보존하되, 연도 선택
// 기능이 생기기 전(모든 기록이 "MM/DD"로만 저장되던 시절)의 값은 2026년 기록으로 간주해
// 연도를 붙여줍니다. ensureDateColumnAsText(migrateLegacyText=true)와 함께 쓰면 시트에
// 남아있는 옛 "MM/DD" 값도 "2026-MM-DD"로 자동 치유됩니다.
function normalizeCourseRoundDateCell(value) {
  const full = normalizeFullDateCell(value);
  if (/^\d{1,2}\/\d{1,2}$/.test(full)) {
    const parts = full.split('/');
    const mm = ('0' + parts[0]).slice(-2);
    const dd = ('0' + parts[1]).slice(-2);
    return `2026-${mm}-${dd}`;
  }
  return full;
}

// screenGolf(Rounds)는 지금도 "MM/DD"만 씁니다. courseRounds만 연도 선택 기능 때문에
// "YYYY-MM-DD" 전체 날짜를 씁니다(과거 "MM/DD" 기록은 자동으로 2026년으로 치유됩니다).
function getDateNormalizer(type) {
  return type === 'courseRounds' ? normalizeCourseRoundDateCell : normalizeDateCell;
}

// date + player 조합이 이미 있으면 해당 행을 덮어쓰고, 없으면 새 행을 추가합니다.
// records 안에 시트에 없는 새로운 필드(지표)가 있으면 열을 자동으로 추가합니다.
// -> 프론트엔드 COURSE_METRICS에 지표를 추가해도 이 스크립트는 수정할 필요가 없습니다.
function saveRoundRecords(type, date, records) {
  const sheet = getOrCreateSheet(type);
  ensureColumns(sheet, records);
  const normalizeFn = getDateNormalizer(type);
  const initialHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  ensureDateColumnAsText(sheet, initialHeaders, normalizeFn, type === 'courseRounds');

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const dateCol = headers.indexOf('date');
  const playerCol = headers.indexOf('player');
  const normDate = normalizeFn(date);

  records.forEach(record => {
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (normalizeFn(values[i][dateCol]) === normDate && String(values[i][playerCol]) === String(record.player)) {
        rowIndex = i;
        break;
      }
    }
    // date 앞에 아포스트로피(')를 붙여서 씁니다. setNumberFormat('@')만으로는 Apps Script의
    // appendRow/setValues가 "01/01" 같은 값을 여전히 실제 날짜로 자동 변환하는 경우가 있어서,
    // 구글 시트가 확실하게 텍스트로 인식하도록 강제하는 표준 방법입니다.
    // (읽어올 때는 아포스트로피가 값에 포함되지 않고 순수 문자열만 돌아옵니다.)
    const rowValues = headers.map(h => {
      if (h === 'date') return "'" + date;
      if (h === 'player') return record.player;
      return (record[h] !== undefined && record[h] !== null) ? record[h] : '';
    });
    if (rowIndex === -1) {
      sheet.appendRow(rowValues);
      // 같은 요청 안에서의 중복 추가 방지용 추적은 아포스트로피 없는 원본 date로 넣어둡니다.
      const trackingRow = rowValues.slice();
      trackingRow[dateCol] = date;
      values.push(trackingRow);
    } else {
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([rowValues]);
    }
  });
}

// 특정 날짜(date)에 해당하는 모든 선수의 행을 시트에서 삭제합니다. 되돌릴 수 없습니다.
function deleteRoundRecords(type, date) {
  const sheet = getOrCreateSheet(type);
  let values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  let headers = values[0].map(h => String(h).trim());
  const normalizeFn = getDateNormalizer(type);
  ensureDateColumnAsText(sheet, headers, normalizeFn, type === 'courseRounds');
  values = sheet.getDataRange().getValues(); // 치유 후 값 다시 읽기
  headers = values[0].map(h => String(h).trim());
  const dateCol = headers.indexOf('date');
  if (dateCol === -1) return;
  const normDate = normalizeFn(date);
  // 뒤에서부터 삭제해야 앞 행을 지워도 나머지 행 번호가 밀리지 않습니다.
  for (let i = values.length - 1; i >= 1; i--) {
    if (normalizeFn(values[i][dateCol]) === normDate) {
      sheet.deleteRow(i + 1);
    }
  }
}

// 스크린골프 스코어 카드: date+player가 유일하지 않으므로(같은 날 여러 회 플레이 가능) 항상
// 새 행으로 추가만 합니다(덮어쓰기 없음). 각 record는 프론트엔드에서 만든 고유 id를 포함합니다.
// records에 overPar처럼 시트에 아직 없는 새 필드가 오면 ensureColumns가 열을 자동으로 추가합니다
// (COURSE_METRICS 확장과 동일한 방식) — 그래서 이미 만들어진 시트에도 재배포 없이 반영됩니다.
function appendScoreCardRecords(records) {
  const sheet = getOrCreateSheet('screenScoreCard');
  ensureColumns(sheet, records);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  ensureDateColumnAsText(sheet, headers, normalizeFullDateCell);

  records.forEach(record => {
    const rowValues = headers.map(h => {
      if (h === 'date') return "'" + record.date; // 연도까지 텍스트로 강제 (apostrophe-prefix)
      return (record[h] !== undefined && record[h] !== null) ? record[h] : '';
    });
    sheet.appendRow(rowValues);
  });
}

// id로 스코어 카드 행 하나를 삭제합니다. 되돌릴 수 없습니다.
function deleteScoreCardRecord(id) {
  const sheet = getOrCreateSheet('screenScoreCard');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(h => String(h).trim());
  const idCol = headers.indexOf('id');
  if (idCol === -1) return;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
    }
  }
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
