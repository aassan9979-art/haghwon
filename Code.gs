/**
 * 학원 전화번호 관리 - 구글시트 공유 백엔드 (Apps Script)
 * 사용법: 구글시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣고
 *   [배포 > 배포 관리 > (기존 배포) 편집 > 버전: 새 버전 > 배포] 하면 같은 URL로 갱신됩니다.
 *   (처음이면 [배포 > 새 배포 > 웹 앱 > 액세스: 모든 사용자])
 * 헤더 이름 기준으로 읽고 쓰므로, 나중에 열이 늘어도 안전합니다.
 */
var SHEET_NAME = '변경사항';
var HEADERS = ['id','상태','학원명','시설종류','발송상태','주소','운영자','전화번호','전화번호2','보유차량수','차량형태','차량상세','상담이력','갱신시각'];

function getSheetAndMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < header.length; i++) { if (header[i] !== '') map[header[i]] = i + 1; }
  for (var j = 0; j < HEADERS.length; j++) {          // 누락된 헤더 자동 추가
    var h = HEADERS[j];
    if (!map[h]) { lastCol++; sh.getRange(1, lastCol).setValue(h); map[h] = lastCol; }
  }
  return { sh: sh, map: map };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 웹앱이 최신 변경분을 불러갈 때
function doGet(e) {
  var o = getSheetAndMap_(), sh = o.sh, map = o.map;
  var last = sh.getLastRow();
  if (last < 2) return json_({ ok: true, rows: [] });
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  function g(row, h) { var c = map[h]; return c ? row[c - 1] : ''; }
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i], id = g(row, 'id');
    if (!id) continue;
    rows.push({
      id: String(id), status: g(row, '상태'), name: g(row, '학원명'), kind: g(row, '시설종류'), send: g(row, '발송상태'), addr: g(row, '주소'),
      oper: g(row, '운영자'), tel: g(row, '전화번호'), tel2: g(row, '전화번호2'),
      carN: (g(row, '보유차량수') === '' ? '' : String(g(row, '보유차량수'))),
      carT: g(row, '차량형태'), carDetail: g(row, '차량상세'), consult: g(row, '상담이력')
    });
  }
  return json_({ ok: true, rows: rows });
}

// 웹앱이 변경 1건을 저장할 때 (같은 id 있으면 덮어쓰기, 없으면 추가)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var o = getSheetAndMap_(), sh = o.sh, map = o.map;
    var idCol = map['id'];
    var n = Math.max(sh.getLastRow() - 1, 0);
    var rowIndex = -1;
    if (n > 0) {
      var ids = sh.getRange(2, idCol, n, 1).getValues();
      for (var i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(d.id)) { rowIndex = i + 2; break; } }
    }
    if (rowIndex < 0) rowIndex = sh.getLastRow() + 1;
    function set(h, v) { sh.getRange(rowIndex, map[h]).setValue(v == null ? '' : v); }
    set('id', d.id); set('상태', d.status || ''); set('학원명', d.name || ''); set('시설종류', d.kind || ''); set('발송상태', d.send || ''); set('주소', d.addr || '');
    set('운영자', d.oper || ''); set('전화번호', d.tel || ''); set('전화번호2', d.tel2 || '');
    set('보유차량수', d.carN || ''); set('차량형태', d.carT || ''); set('차량상세', d.carDetail || '');
    set('상담이력', d.consult || '');
    set('갱신시각', new Date());
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
