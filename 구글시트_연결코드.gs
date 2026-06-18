/**
 * 웅진 자막 검수 도구 → 구글 시트 연결 코드 (v2: 다른 기기에서 이어하기 지원)
 *
 * 설치 (한 번만):
 * 1) 새 구글 시트 만들기 (주소창에 sheets.new)
 * 2) 상단 메뉴: 확장 프로그램 → Apps Script
 * 3) 기존 코드를 전부 지우고, 이 파일 내용을 통째로 붙여넣은 뒤 저장
 * 4) 우측 상단 [배포] → [새 배포]
 *      - 유형(톱니바퀴): 웹 앱
 *      - 실행 계정: 나
 *      - 액세스 권한: 모든 사용자
 *    → [배포] → 권한 검토/허용
 * 5) 나온 "웹 앱 URL"(.../exec 로 끝남)을 복사
 * 6) 검수 도구의 [설정 → 구글 시트 주소] 칸에 붙여넣기
 *
 * ※ 코드를 바꾼 뒤에는 반드시 [배포] → [배포 관리] → 연필(편집) →
 *    버전 "새 버전" → [배포] 로 다시 적용해야 반영됩니다.
 *    (예전 v1 코드를 쓰고 계셨다면, 이 코드로 교체 후 재배포해 주세요.)
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    writeSheet(ss, '검수일정표', data.schedule);
    writeSheet(ss, '오류목록',   data.errors);
    writeSheet(ss, '정산표',     data.settle);
    if (data.raw) saveState(ss, JSON.stringify(data.raw));
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  }
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var payload = JSON.stringify({ raw: loadState(ss) });
  if (e && e.parameter && e.parameter.callback) {
    return ContentService
      .createTextOutput(e.parameter.callback + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

function writeSheet(ss, name, rows) {
  if (!rows || !rows.length) return;
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  var width = 0, i, j;
  for (i = 0; i < rows.length; i++) width = Math.max(width, rows[i].length);
  for (j = 0; j < rows.length; j++) while (rows[j].length < width) rows[j].push('');
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.getRange(1, 1, 1, width).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function stateSheet(ss) {
  var sh = ss.getSheetByName('_state');
  if (!sh) { sh = ss.insertSheet('_state'); sh.hideSheet(); }
  return sh;
}
function saveState(ss, jsonStr) { stateSheet(ss).getRange(1, 1).setValue(jsonStr); }
function loadState(ss) { var v = stateSheet(ss).getRange(1, 1).getValue(); return v ? String(v) : ''; }
