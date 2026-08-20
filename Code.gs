/**
 * 학원 전화번호 관리 - 구글시트 공유 백엔드 (Apps Script)
 * 사용법: 구글시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣고
 *         "배포 > 새 배포 > 웹 앱 (액세스: 모든 사용자)" 으로 배포하세요.
 * 시트에는 변경분(추가/수정/삭제)만 기록됩니다. 원본 24,116건은 웹앱이 따로 갖고 있습니다.
 */
var SHEET_NAME = '변경사항';
var HEADERS = ['id','상태','학원명','주소','운영자','전화번호','전화번호2','보유차량수','차량형태','갱신시각'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 웹앱이 최신 변경분을 불러갈 때
function doGet(e) {
  var sh = getSheet_();
  var vals = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (!r[0]) continue;
    rows.push({
      id: String(r[0]), status: r[1], name: r[2], addr: r[3], oper: r[4],
      tel: r[5], tel2: r[6], carN: (r[7] === '' ? '' : String(r[7])), carT: r[8]
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
    var sh = getSheet_();
    var vals = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(d.id)) { rowIndex = i + 1; break; }
    }
    var row = [d.id, d.status || '', d.name || '', d.addr || '', d.oper || '',
               d.tel || '', d.tel2 || '', d.carN || '', d.carT || '', new Date()];
    if (rowIndex > 0) sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
