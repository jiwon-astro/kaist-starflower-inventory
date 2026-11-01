// 별바라기 장비실: 웹 페이지의 반출 보드가 '반출' 탭에 기록하고, 장비 사진 목록을 읽도록 하는 스크립트.
// 설치: 시트 → 확장 프로그램 → Apps Script → 이 내용을 붙여 넣기
//       → 프로젝트 설정 → 스크립트 속성에 PASSWORD (운영진 공용 비밀번호), PHOTO_FOLDER (사진 폴더 ID) 추가
//       → 배포 → 새 배포 → 웹 앱 (실행: 나, 액세스: 모든 사용자) → 주소를 index.html SCRIPT_URL에
// 코드를 고친 뒤에는 배포 → 배포 관리 → 수정 → 새 버전 (주소 유지)

// 사진 파일명: 관리번호로 시작하고 날짜(YYYYMMDD 또는 YYYY-MM-DD)가 들어 있으면 됨. 이름은 선택.
// 예: MT07_20261101.jpg, MT07_AZ-GTi_20261101.jpg, MT07-2026-11-01.png. 둘 중 하나라도 없으면 무시.
function parsePhotoName(name) {
  const m = name.match(/^([A-Z]{2}\d{2,})([_\-\s].*)?\.[a-z0-9]+$/i);
  const d = m && (m[2] || '').match(/(\d{4})-?(\d{2})-?(\d{2})(?!.*\d{4}-?\d{2}-?\d{2})/); // 마지막 날짜
  return d ? { code: m[1].toUpperCase(), date: d[1] + d[2] + d[3] } : null;
}

function doGet() {
  const raw = PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER') || '';
  const folder = DriveApp.getFolderById((raw.match(/[-\w]{25,}/) || [raw])[0]); // 폴더 주소 전체를 넣어도 ID만 꺼냄
  const out = {};
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const p = parsePhotoName(f.getName());
    if (!p || !f.getMimeType().startsWith('image/')) continue;
    (out[p.code] = out[p.code] || []).push({ id: f.getId(), name: f.getName(), date: p.date });
  }
  Object.values(out).forEach(list => list.sort((a, b) => b.date.localeCompare(a.date))); // 최신 사진 먼저
  return reply(out);
}

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
