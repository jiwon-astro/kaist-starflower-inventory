// 별바라기 장비실: 웹 페이지의 반출 보드가 '반출' 탭에 기록하도록 하는 스크립트.
// 설치: 시트 → 확장 프로그램 → Apps Script → 이 내용을 붙여 넣기
//       → 프로젝트 설정 → 스크립트 속성에 PASSWORD (운영진 공용 비밀번호) 추가
//       → 배포 → 새 배포 → 웹 앱 (실행: 나, 액세스: 모든 사용자) → 주소를 index.html SCRIPT_URL에

function doPost(e) {
  const req = JSON.parse(e.postData.contents);
  if (req.password !== PropertiesService.getScriptProperties().getProperty('PASSWORD')) {
    return reply({ ok: false, error: '비밀번호가 틀렸습니다' });
  }
  const sh = SpreadsheetApp.getActive().getSheetByName('반출');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const n = sh.getMaxRows() - 1;
    const ids = sh.getRange(2, 3, n, 1).getValues().map(r => String(r[0]).trim());
    const back = sh.getRange(2, 6, n, 1).getValues().map(r => r[0] === true);
    if (req.action === 'out') {
      let row = ids.indexOf('') + 2; // 관리번호 열의 첫 빈 행 (물품명 수식이 미리 깔려 있어 appendRow를 쓰지 않음)
      if (row < 2) { sh.insertRowAfter(sh.getMaxRows()); row = sh.getMaxRows(); }
      sh.getRange(row, 1, 1, 3).setValues([[new Date(), req.event || '', req.id]]);
      sh.getRange(row, 5).setValue(req.who || '');
      sh.getRange(row, 6).insertCheckboxes().uncheck();
    } else if (req.action === 'in') {
      for (let i = ids.length - 1; i >= 0; i--) {
        if (ids[i] === req.id && !back[i]) { sh.getRange(i + 2, 6).insertCheckboxes().check(); break; }
      }
    } else {
      return reply({ ok: false, error: '알 수 없는 요청' });
    }
    return reply({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
