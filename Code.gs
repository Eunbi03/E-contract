// @ts-nocheck

// ⭐ 3-1. doGet(e)
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.contractId = (e && e.parameter && e.parameter.id) ? e.parameter.id : "";
  template.viewMode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : "";
  return template.evaluate()
    .setTitle('전자 근로 계약서')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// ※ XFrameOptionsMode를 ALLOWALL → DEFAULT로 변경했습니다. 주민등록번호 등 민감정보를 입력받는 화면이라
//   다른 사이트가 iframe으로 몰래 감싸 클릭재킹하는 걸 막기 위함입니다. 이 앱을 의도적으로 다른 페이지에
//   iframe으로 삽입해 쓰고 있었다면 ALLOWALL로 되돌려야 합니다.

// ⭐ 동시 접근 시 시트 행이 꼬이지 않도록 잠그고 실행하는 헬퍼
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ⭐ 월초/월말처럼 200명 이상을 한 번에 발송하는 경우, 한 번의 실행에서 전부 보내려 하면
//   ①실행시간 제한(무료 6분/Workspace 30분), ②메일 하루 발송 한도(일반 계정 100통 등)에 걸릴 수 있다.
//   그래서 시트 기록은 한 번에 몰아서 처리(appendRow 반복 대신 setValues 1회)하고,
//   실제 발송은 "발송큐" 시트에 쌓아둔 뒤 일부만 즉시 보내고 나머지는 시간 기반 트리거로
//   몇 분 간격씩 나눠서 이어 보낸다(processSendQueue).
var SEND_CHUNK_SIZE = 30;               // 한 번에 처리할 인원
var SEND_CHUNK_INTERVAL_MS = 2 * 60 * 1000;   // 청크 사이 간격(2분)
var SEND_QUOTA_RETRY_MS = 60 * 60 * 1000;     // 일일 메일 한도 초과 시 재시도 간격(1시간)

// 3-2. sendBatchContracts(commonData, workerList)
function sendBatchContracts(commonData, workerList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var draftSheet = ss.getSheetByName('계약대기') || ss.insertSheet('계약대기');
  if (draftSheet.getLastRow() === 0) draftSheet.appendRow(['ID', '데이터', '발송시간', '대상자이메일', '전화번호']);

  var queueSheet = ss.getSheetByName('발송큐') || ss.insertSheet('발송큐');
  if (queueSheet.getLastRow() === 0) {
    queueSheet.appendRow(['ID', '성명', '이메일', '전화번호', '링크', '모드', '법인명', '담당자명', '담당자연락처', '담당자이메일', '상태', '생성시간']);
  }

  var urlBase = ScriptApp.getService().getUrl();
  var nowStr = new Date().toLocaleString();

  var draftRows = [];
  var queueRows = [];

  for (var i = 0; i < workerList.length; i++) {
    var worker = workerList[i];
    // ⭐ 발송 방식(이메일/문자)을 이제 전체 배치 공통이 아니라 근로자별로 받는다.
    var workerMode = worker.mode || 'email';
    var individualData = JSON.parse(JSON.stringify(commonData));
    individualData.contractPassword = worker.password;
    individualData.empName = worker.name;
    individualData.workerContactEmail = worker.email;
    individualData.status = 'DRAFT';
    individualData.sentAt = nowStr;

    // ⭐ 추측 불가능한 ID 발급 (기존 "법인코드-이름-순번-날짜" 방식은 URL을 쉽게 유추/전수조사할 수 있어 변경)
    var uniqueId = Utilities.getUuid();
    draftRows.push([uniqueId, JSON.stringify(individualData), nowStr, worker.email, worker.password]);

    var link = urlBase + "?id=" + uniqueId;
    queueRows.push([
      uniqueId, worker.name, worker.email || '', worker.password || '', link, workerMode,
      commonData.companyName, commonData.managerName || '', commonData.managerPhone || '', commonData.managerEmail || '',
      'PENDING', nowStr
    ]);
  }

  withLock(function () {
    if (draftRows.length) {
      draftSheet.getRange(draftSheet.getLastRow() + 1, 1, draftRows.length, draftRows[0].length).setValues(draftRows);
    }
    if (queueRows.length) {
      queueSheet.getRange(queueSheet.getLastRow() + 1, 1, queueRows.length, queueRows[0].length).setValues(queueRows);
    }
  });

  // ⭐ 첫 묶음은 바로 보내서 관리자가 빠르게 피드백을 받게 하고, 나머지는 큐에 남겨 트리거가 이어받는다.
  //   processSendQueue()는 시트 전체에서 가장 오래된 PENDING부터 처리하므로, 아직 안 빠진 이전
  //   배치가 남아있으면 이번 호출에서 실제로 처리되는 건 "이전 배치"일 수 있다. 그래서 응답은
  //   processSendQueue()의 결과를 그대로 쓰지 않고, 방금 추가한 이 배치의 id들만 다시 조회해서
  //   집계한다 — 그래야 배치를 연달아 보내도 안내 문구가 서로 섞이지 않는다.
  processSendQueue();

  var justAddedIds = {};
  draftRows.forEach(function (r) { justAddedIds[r[0]] = true; });
  var freshQueue = queueSheet.getDataRange().getValues();
  var sentNow = 0, failNow = 0, queuedRemaining = 0;
  for (var qi = 1; qi < freshQueue.length; qi++) {
    if (!justAddedIds[freshQueue[qi][0]]) continue;
    var st = freshQueue[qi][10];
    if (st === 'SENT') sentNow++;
    else if (st === 'FAILED') failNow++;
    else queuedRemaining++;
  }

  return { totalCount: workerList.length, sentNow: sentNow, failNow: failNow, queuedRemaining: queuedRemaining };
}

// ⭐ "발송큐" 시트에서 대기 중(PENDING)인 항목을 최대 SEND_CHUNK_SIZE개 꺼내 실제로 발송하고,
//   더 남아있으면 SEND_CHUNK_INTERVAL_MS 뒤에 자기 자신을 다시 호출하는 1회성 트리거를 예약한다.
//   (Apps Script의 after() 1회성 트리거는 실행되고 나면 자동으로 삭제된다.)
function processSendQueue() {
  return withLock(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('발송큐');
    if (!sheet) return { sentCount: 0, failCount: 0 };

    var data = sheet.getDataRange().getValues();
    var statusCol = 10; // '상태' 열 (0-based)

    // ⭐ 대기열을 앞에서부터 훑을 때, 이메일 한도가 소진돼 이번 회차에 못 보낼 이메일 건은
    //   건너뛰고 계속 다음 후보를 찾는다 — 그래야 뒤쪽에 있는 SMS 건(이메일 한도와 무관)이
    //   앞쪽에 막힌 이메일 건들 때문에 청크 30자리를 다 뺏기고 하루 종일 밀리지 않는다.
    var quotaExhausted = false;
    var simQuota = MailApp.getRemainingDailyQuota();
    var pendingRowIndexes = [];
    for (var i = 1; i < data.length && pendingRowIndexes.length < SEND_CHUNK_SIZE; i++) {
      if (data[i][statusCol] !== 'PENDING') continue;
      if (data[i][5] === 'email') {
        if (simQuota <= 0) { quotaExhausted = true; continue; }
        simQuota--;
      }
      pendingRowIndexes.push(i);
    }

    var sentCount = 0, failCount = 0;

    for (var k = 0; k < pendingRowIndexes.length; k++) {
      var rowIdx = pendingRowIndexes[k];
      var row = data[rowIdx];
      var name = row[1], email = row[2], phone = row[3], link = row[4], rowMode = row[5],
          companyName = row[6], managerName = row[7], managerPhone = row[8], managerEmail = row[9];

      // ⭐ 서명 요청 메세지 문구 (기존과 동일)
      var messageBody = "안녕하세요. " + name + "님.\n" + companyName + "입니다.\n\n입사를 진심으로 환영 드립니다.\n근로계약 체결을 위해 전자근로계약서를 발송드리오니, 내용을 충분히 검토하신 후 전자서명 진행 부탁드립니다.\n\n문의 사항이 있으면 서명 전 반드시 아래 연락처로 연락 부탁드립니다.\n\n감사합니다.\n\n접속 링크: " + link + "\n접속 비밀번호: 본인 휴대전화번호\n\n담당자: " + (managerName || "") + "\n전화번호: " + (managerPhone || "") + "\n이메일: " + (managerEmail || "");

      var newStatus = 'SENT';
      try {
        if (rowMode === 'email') {
          MailApp.sendEmail({
            to: email,
            replyTo: managerEmail,
            subject: "[전자계약] " + companyName + " 전자근로계약서 확인 및 서명 요청",
            body: messageBody
          });
        } else if (rowMode === 'sms') {
          sendSolapiMessage(phone, managerPhone || '0333400023', messageBody);
        }
        sentCount++;
      } catch (e) {
        Logger.log('발송 실패 (' + name + '): ' + e);
        newStatus = 'FAILED';
        failCount++;
      }
      data[rowIdx][statusCol] = newStatus;
      sheet.getRange(rowIdx + 1, statusCol + 1).setValue(newStatus);
    }

    var stillPending = false;
    for (var j = 1; j < data.length; j++) {
      if (data[j][statusCol] === 'PENDING') { stillPending = true; break; }
    }
    if (stillPending) {
      ensureQueueTriggerScheduled(quotaExhausted ? SEND_QUOTA_RETRY_MS : SEND_CHUNK_INTERVAL_MS);
    }

    return { sentCount: sentCount, failCount: failCount };
  });
}

// ⭐ processSendQueue용 1회성 트리거가 이미 예약돼 있으면 중복 예약하지 않는다
//   (여러 배치를 연달아 보내도 트리거가 쌓이지 않도록).
function ensureQueueTriggerScheduled(delayMs) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processSendQueue') return;
  }
  ScriptApp.newTrigger('processSendQueue').timeBased().after(delayMs).create();
}

// ⭐ 계약대기/승인대기 시트처럼 A열에 id가 그대로 들어있는 경우 공용으로 쓰는 조회 헬퍼.
//   (loadDraftData/loadSignedData/verifyIdentity가 각자 따로 구현하던 반복문을 여기로 모았다 —
//   예전에는 이 세 곳 중 한 곳만 고치고 나머지를 놓쳐서 "서명 제출 후 재접속 시 링크 오류"
//   같은 버그가 생겼었다.)
function findRowById(sheet, id, idCol) {
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return data[i];
  }
  return null;
}

// ⭐ 계약목록 시트는 id가 별도 열이 아니라 J열(인덱스 10) JSON 데이터 안 contractId 필드에
//   들어있어 findRowById와는 다른 방식으로 찾아야 한다.
function findContractIdInList(listSheet, id) {
  if (!listSheet) return null;
  var data = listSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    try {
      var parsed = JSON.parse(data[i][10]);
      if (String(parsed.contractId) === String(id)) return parsed;
    } catch (e) {}
  }
  return null;
}

// ⭐ findRowById와 짝을 이루는 삭제 헬퍼. saveContractData/approveContract가 각자 손으로
//   반복문을 짜다 보니 시작 인덱스가 0/1로 서로 달랐던 것도 여기로 모으면서 함께 정리된다.
function deleteRowById(sheet, id, idCol) {
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(id)) { sheet.deleteRow(i + 1); return; }
  }
}

// 3-3. loadDraftData(id)
function loadDraftData(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (findContractIdInList(ss.getSheetByName('계약목록'), id)) return "ALREADY_COMPLETED";

  // ⭐ 근로자가 서명 제출을 이미 마쳤고(승인대기로 이동됨) 아직 관리자 승인 전인 경우.
  //   중복 작성을 막는 목적은 유지하되, "유효하지 않은 링크"가 아니라 "이미 제출했고 승인
  //   대기 중"이라고 정확히 안내하도록 별도 신호값을 반환한다(ALREADY_COMPLETED와 같은 패턴).
  if (findRowById(ss.getSheetByName('승인대기'), id, 0)) return "ALREADY_SUBMITTED";

  var row = findRowById(ss.getSheetByName('계약대기'), id, 0);
  if (!row) return null;
  try {
    var parsed = JSON.parse(row[1]);
    // ⭐ 본인인증 전에 비밀번호(전화번호)를 클라이언트로 절대 내려보내지 않는다.
    //   verifyIdentity()가 서버에서만 비교하도록 분리했다.
    delete parsed.contractPassword;
    return parsed;
  } catch (e) {
    return null;
  }
}

// ⭐ (근로자) 본인인증 — 비밀번호 비교를 서버에서만 수행하고 결과(boolean)만 반환한다.
function verifyIdentity(id, inputPassword) {
  var cleanInput = String(inputPassword || '').replace(/[^0-9]/g, '');
  if (!cleanInput) return false;
  var row = findRowById(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('계약대기'), id, 0);
  if (!row) return false;
  try {
    var parsed = JSON.parse(row[1]);
    var cleanStored = String(parsed.contractPassword || '').replace(/[^0-9]/g, '');
    return !!cleanStored && cleanInput === cleanStored;
  } catch (e) {
    return false;
  }
}

// 3-4. loadSignedData(id)
function loadSignedData(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var row = findRowById(ss.getSheetByName('승인대기'), id, 0);
  if (row) {
    try { return JSON.parse(row[1]); } catch (e) {}
  }

  return findContractIdInList(ss.getSheetByName('계약목록'), id);
}

// 3-5. saveContractData(data)
function saveContractData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName('승인대기') || ss.insertSheet('승인대기');

  // ⭐ 승인대기 시트 목록
  if (targetSheet.getLastRow() === 0) targetSheet.appendRow(['ID', '완료데이터(JSON)', '현황', '수신시간', '승인링크']);

  var uniqueId = data.contractId;
  data.status = 'SIGNED_BY_WORKER';
  data.workerSignedAt = new Date().toLocaleString();
  var viewLink = ScriptApp.getService().getUrl() + "?id=" + uniqueId + "&mode=view";

  withLock(function () {
    // ⭐ 승인대기 시트 저장 데이터 설정
    targetSheet.appendRow([uniqueId, JSON.stringify(data), 'SIGNED_BY_WORKER', new Date(), viewLink]);
    deleteRowById(ss.getSheetByName('계약대기'), uniqueId, 0);
  });
  return "Success";
}

// 3-6. approveContract(data)
function approveContract(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  data.status = 'COMPLETED';
  data.approvedAt = new Date().toLocaleString();
  var listSheet = ss.getSheetByName('계약목록') || ss.insertSheet('계약목록');

  // ⭐ 계약목록 시트 목록
  if (listSheet.getLastRow() === 0) {
    listSheet.appendRow(['법인', '성명', '연락처', '주민등록번호', '주소', '입사일', '문서링크', '발송일시', '제출일시', '완료일시', '데이터(JSON)']);
  }

  var viewLink = ScriptApp.getService().getUrl() + "?id=" + data.contractId + "&mode=view";

  withLock(function () {
    // ⭐ 계약목록 시트 저장 데이터 설정
    listSheet.appendRow([
      data.companyName, data.empName, data.empPhone, data.empRegNumber, data.empAddress,
      data.joinDate, viewLink, data.sentAt || "", data.workerSignedAt || "", data.approvedAt, JSON.stringify(data)
    ]);
    deleteRowById(ss.getSheetByName('승인대기'), data.contractId, 0);
  });

  var todayObj = new Date();
  var expireDate = new Date(todayObj);
  expireDate.setMonth(expireDate.getMonth() + 3);
  var expireDateStr = Utilities.formatDate(expireDate, "Asia/Seoul", "yyyy년 MM월 dd일");

  // ⭐ 계약 완료 메세지 문구
  //   시트 커밋(승인 자체)은 위에서 이미 끝났으므로, 메일 발송만 실패해도(한도 초과, 잘못된
  //   주소 등) 승인 전체가 실패한 것처럼 클라이언트에 에러가 전달되지 않도록 별도로 감싼다.
  //   (클라이언트에 withFailureHandler가 붙어 있어, 여기서 그냥 던지면 이미 완료된 승인을
  //   "실패"로 오인하고 관리자가 재승인을 시도할 수 있다.)
  if (data.workerContactEmail) {
    try {
      MailApp.sendEmail({
        to: data.workerContactEmail,
        subject: "[계약완료] " + data.companyName + " 근로계약서 체결 완료 및 다운로드 안내",
        body: "안녕하세요. " + data.empName + "님.\n\n" +
          "전자근로계약이 최종 승인되어 체결이 완료되었습니다.\n" +
          "아래 링크로 접속하신 후, 화면의 [인쇄 및 PDF 저장] 버튼을 눌러 최종 계약서를 다운로드하여 보관해 주시기 바랍니다.\n\n" +
          "▶ 최종 계약서 확인 및 PDF 다운로드 링크:" + viewLink + "\n" +
          "▶ 다운로드 가능 기간: ~ " + expireDateStr + "\n\n" +
          "담당자: " + (data.managerName || "") + "\n전화번호: " + (data.managerPhone || "") + "\n이메일: " + (data.managerEmail || "") + "\n\n" +
          "감사합니다."
      });
    } catch (e) {
      Logger.log('완료 안내 메일 발송 실패(' + data.workerContactEmail + '): ' + e);
    }
  }
  return "Approved";
}

// ⭐ 관리자가 최종 승인한 직후, 화면에 렌더링된 계약서(직인 포함) 그대로를 넘겨받아
//   담당자 컴퓨터로 동기화되는 Drive 폴더에 자동 보관한다. Index.html의 #print-area outerHTML을
//   그대로 저장하므로 계약서 문구를 서버 코드에 다시 옮겨 적을 필요가 없고(오타/누락 위험 없음),
//   근로자가 서명한 화면과 100% 동일한 내용이 보관된다.
// ⭐ 이 함수는 공유 시트의 행 인덱스를 다루지 않고(고유 id 기준으로 항상 새 파일만 만든다),
//   Drive 문서 변환은 몇 초씩 걸릴 수 있어 다른 함수들과 같은 전역 락을 여기서 잡으면
//   그 시간 동안 승인/서명저장/발송 같은 무관한 작업들이 락 대기(30초)로 실패할 수 있다.
//   그래서 withLock으로 감싸지 않는다.
function archiveContractHtml(id, html) {
  var folder = getOrCreateArchiveFolder();
  var safeHtml = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>근로계약서 ' + id + '</title></head><body>' + html + '</body></html>';
  var htmlBlob = Utilities.newBlob(safeHtml, 'text/html', '근로계약서_' + id + '.html');
  folder.createFile(htmlBlob);

  // PDF 변환은 Apps Script 고급 서비스인 Drive API가 있어야 동작한다(스크립트 편집기 좌측 서비스(+)에서
  // "Drive API" 추가). 추가돼 있지 않으면 이 블록만 조용히 건너뛰고 위의 HTML 백업은 그대로 남는다.
  try {
    var resource = { title: '근로계약서_' + id, mimeType: MimeType.GOOGLE_DOCS };
    var docFile = Drive.Files.insert(resource, htmlBlob);
    var doc = DocumentApp.openById(docFile.id);
    var pdfBlob = doc.getAs('application/pdf').setName('근로계약서_' + id + '.pdf');
    folder.createFile(pdfBlob);
    DriveApp.getFileById(docFile.id).setTrashed(true);
  } catch (e) {
    Logger.log('PDF 자동 변환 건너뜀(고급 Drive 서비스 미설정 가능성): ' + e);
  }
  return 'OK';
}

function getOrCreateArchiveFolder() {
  var name = '전자근로계약_완료보관함';
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ⭐ 프론트엔드(Index.html)가 시작할 때 불러오는 법인 목록.
//   회사 정보가 바뀌거나 새 법인이 늘어나도 코드 배포 없이 "회사설정" 시트만 고치면 반영된다.
//   ※ 직인 이미지는 여기 포함하지 않는다 — getSealImage(companyName)을 따로 두고, 관리자가
//   최종 승인(직인 날인)한 시점에만 해당 법인 것 한 장만 내려받도록 분리했다. 그래야 아직
//   서명 중인 근로자를 포함해 이 링크를 여는 모든 사람에게 5개 법인 직인 이미지가 전부
//   전송되는 걸 막을 수 있다.
function getConfig() {
  return {
    companies: getCompanySettings(),
    clauses: getClauseTemplates(),
    defaults: getDefaults()
  };
}

// ⭐ 계약서 문구 중 회사명/근무조건 같은 가변 항목이 전혀 없는 순수 고정 조항만 시트로 뺐다.
//   (제2·6조처럼 정규직/계약직 체크박스나 근로시간 표 같은 조건부 구조가 있는 조항은 자유
//   텍스트로 바꾸면 오히려 깨지기 쉬워서 지금처럼 코드/사이드바 입력값으로 남겨둔다.)
//   담당자가 "조항템플릿" 시트의 본문 셀만 고치면 코드 배포 없이 바로 반영된다.
function getClauseTemplates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('조항템플릿');

  // ⭐ 변수(근무장소/담당업무)가 섞인 조항은 {{workPlace}}, {{duties}} 같은 자리표시자로
  //   저장해두고, 실제 계약서를 그릴 때 formData 값으로 치환한다.
  var seedDefaults = {
    '1': ['목적', '본 근로계약서는 "갑"과 "을"이 근로계약을 체결함에 있어 임금, 근로시간 등 근로조건을 정하는 것을 목적으로 한다.'],
    '3': ['근무장소', '"을"의 근무 장소는 {{workPlace}}(으)로 한다. 다만 "갑"은 인사명령에 의거 "을"의 근무 장소를 변경할 수 있으며 "을"은 이에 이의를 제기하지 않는다.'],
    '4': ['담당업무', '"을"의 담당업무는 {{duties}}(으)로 한다. 다만, "갑"은 인사명령에 의거 "을"의 담당업무를 변경할 수 있으며 "을"은 이에 이의를 제기하지 않는다.'],
    '5': ['임금', '(1) 임금형태 : "갑"과 "을"은 월급제로 임금계약을 체결한다. (기본급과 법정 제수당이 포함된 "포괄임금")<br/>(2) 임금내용 : 월급총액, 월급의 구성내역, 지급방법 등은 임금계약서에 의한다.<br/>(3) 사정변경 : "갑"은 계약기간 중 제반 경영여건상 처우조건이 불가피하거나, 인사상 처우조정 사유발생시에는 제반기준에 의거 임금을 조정할 수 있다.<br/>(4) 임금인상 : "갑"과 "을"은 근로계약기간이 종료되는 시점에서 업무상 필요성, "갑"의 경영사정에 의거 재계약을 체결할 시 상호 대등한 입장으로 임금인상 등을 협의할 수 있다. 다만 임금내용의 변동이 없는 경우, 또는 임금인상에 대한 합의가 성립 되지 않는 경우 이전과 동일한 금액으로 반복하여 체결하며 "을"은 이에 이의를 제기하지 않는다.<br/>(5) 임금산정일 및 지급일자 : 초일부터 말일까지 계산하여 익월 10일에 지급한다. (단 공휴일일경우에는 그 다음 영업일에 지급한다.)<br/>(6) 비밀유지 : 임금은 회사의 규정과 상사의 지시에 따른 본인의 근로제공에 대한 모든 금전적 보상임을 확인하며 직원상호간 비밀의무를 반드시 준수한다.'],
    '7': ['휴일 및 휴가', '(1) 휴일 : 주휴일(1주간 소정근로일수 개근시 주1회), 근로자의 날, 관공서공휴일규정에 따른 공휴일과 대체공휴일<br/> ※ 업무 특성 상 특정 주에 변경되는 주휴일은 휴일변경(대체)하는 것이며, 근로자는 이에 동의한다.<br/>(2) 휴가 : 연차유급휴가는 "근로기준법"에 의한다. (1년간 80퍼센트 이상 출근한 근로자에게 15일 / 1년 미만 근로자의 경우 1개월 개근시 1일)<br/> ※ 근로자가 본인에게 발생한 연차유급휴가일수 보다 더 많은 연차휴가를 사용한 경우 이는 "사용자"가 "근로자"에게 향후 발생할 연차휴가를 선사용하도록 허용한 것이므로 "근로자"가 퇴사하는 경우 더 많이 사용한 연차휴가 일수만큼 임금을 공제할 수 있다.<br/> ※ 1주간 소정근로시간 15시간 미만인 자는 휴일 및 휴가 규정은 적용하지 아니한다. (근로자의 날 제외)'],
    '8': ['근로계약의 종료 등', '(1) "을"이 정년에(만60세) 도달하거나 약정한 근로계약기간의 만료로 근로계약은 자동 종료된다.<br/>(2) 근로계약의 만료일이 지정된 경우 계약기간 만료시 본 계약은 자동해지 된다.<br/>(3) "갑"의 해지사유 : "갑"은 정년도래, 기간만료 이외 "갑"의 관련규정 등에서 정한 정당한 사유가 있는 경우에는 "을"의 의사에도 불구하고 근로계약을 해지할 수 있다. 다만, "갑"은 근로기준법에서 규정하고 있는 해고예고의 예외가 되는 근로자의 귀책사유 이외에는 해고예정일 30일전에 서면으로 통보해야 한다.<br/>(4) "을"의 해지사유 : "을"이 정년도래, 기간만료 이외 계약을 해지하고자 하는 경우에는 30일전에 "갑"에게 통보하여야 하고 후임자에게 업무의 인수인계를 하여야 하며, 이를 해태함으로 인하여 "갑"에게 손해가 발생하는 경우에는 이를 배상하여야 한다.'],
    '9': ['의무', '"갑"은 "을"의 근무조건 향상을 위하여 최선을 다하여야 하며, "을"은 신의성실의 원칙에 의하여 근로를 제공하여야 한다. 특히 "을"은 "갑"이 정한 안전에 관한 제 규칙과 지시사항을 위반하여 발생한 제반사고는 "을"의 귀책사유로 한다.'],
    '10': ['손해배상', '"을"이 계약기간 중 고의 또는 과실로 "갑"에게 손해를 입힌 때에는 이를 배상하여야 한다.'],
    '11': ['기타 근로조건', '본 계약서를 작성함에 있어 "을"은 "갑"의 취업규칙 및 제 규정을 열람하였으며 이 계약에 정함이 없는 사항은 관계법령 및 "갑"의 취업규칙 등에 정한 바에 따르며 상기사실을 확실히 하기 위하여 본 계약서를 2통 작성하여 사용자와 근로자가 각 1통씩 보관키로 한다.']
  };
  var order = ['1', '3', '4', '5', '7', '8', '9', '10', '11'];

  if (!sheet) {
    sheet = ss.insertSheet('조항템플릿');
    sheet.appendRow(['조번호', '제목', '본문']);
  }

  var data = sheet.getDataRange().getValues();
  var existingNums = {};
  for (var i = 1; i < data.length; i++) {
    var n = String(data[i][0]).trim();
    if (n) existingNums[n] = true;
  }

  // ⭐ 1단계(제1·9·10·11조만)에서 이미 만들어둔 시트라도, 이번에 추가된 조항(3·4·5·7·8)만
  //   모자라게 이어서 채워준다. 이미 있던 조항(관리자가 고쳤을 수 있는)은 건드리지 않는다.
  order.forEach(function (num) {
    if (!existingNums[num]) sheet.appendRow([num, seedDefaults[num][0], seedDefaults[num][1]]);
  });

  data = sheet.getDataRange().getValues();
  var map = {};
  for (var j = 1; j < data.length; j++) {
    var num2 = String(data[j][0]).trim();
    if (!num2) continue;
    map[num2] = { title: data[j][1], body: data[j][2] };
  }
  return map;
}

// ⭐ 수정 탭에서 조항 하나를 저장할 때 호출. 해당 조번호 행이 있으면 제목/본문을 덮어쓰고,
//   없으면(이론상 getClauseTemplates가 항상 먼저 만들어두므로 거의 없음) 새로 추가한다.
function saveClauseTemplate(num, title, body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('조항템플릿');
  if (!sheet) { getClauseTemplates(); sheet = ss.getSheetByName('조항템플릿'); }
  return withLock(function () {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(num)) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[title, body]]);
        return 'OK';
      }
    }
    sheet.appendRow([String(num), title, body]);
    return 'OK';
  });
}

function getCompanySettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('회사설정');

  // ⭐ 시트가 없으면 기존에 코드에 하드코딩돼 있던 값 그대로 최초 1회 생성해준다.
  if (!sheet) {
    sheet = ss.insertSheet('회사설정');
    sheet.appendRow(['법인명', '주소', '전화번호', '대표자', '담당자명', '담당자연락처', '담당자이메일']);
    var defaults = [
      ['(주)케이프라이드', '강원도 횡성군 우천면 우천제2농공단지로 65-50', 'T. 033-644-4467 F.033-644-1944', '김도영', '송문주', '033-340-0023', 'kpride@example.com'],
      ['㈜케이펙', '강원도 강릉시 정원로 54, 8층(교동, 주니어타운)', 'T. 033-644-4467 F.033-644-1944', '김도영', '하주연', '033-340-0012', 'kpride@example.com'],
      ['백두대간영농조합법인', '강원도 강릉시 정원로 54, 8층(교동, 주니어타운)', 'T. 033-644-4467 F.033-644-1944', '김도영', '서영민', '033-340-0021', 'kpride@example.com'],
      ['㈜보담', '강원도 원주시 저금어지길 456(가현동 강원LPC 2층)', 'T. 033-644-4467 F.033-644-1944', '이승수', '이선우', '033-340-0020', 'kpride@example.com'],
      ['(주)마시타', '강원도 강릉시 정원로 54, 8층(교통, 주니어타운)', 'T. 033-644-4467 F.033-644-1944', '김주원', '송문주', '033-340-0023', 'kpride@example.com']
    ];
    defaults.forEach(function (row) { sheet.appendRow(row); });
  }

  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    list.push({
      name: r[0], address: r[1], phone: r[2], rep: r[3],
      managerName: r[4], managerPhone: r[5], managerEmail: r[6]
    });
  }
  return list;
}

// ⭐ 설정 탭에서 회사를 추가/수정할 때 쓴다. originalName이 있으면(기존 회사 수정, 법인명 자체를
//   바꾸는 경우 포함) 그 이름의 행을 찾아 전체 필드를 덮어쓰고, 없으면(신규 추가) 새 행을 더한다.
function upsertCompany(company, originalName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('회사설정');
  if (!sheet) { getCompanySettings(); sheet = ss.getSheetByName('회사설정'); }

  var row = [
    company.name || '', company.address || '', company.phone || '', company.rep || '',
    company.managerName || '', company.managerPhone || '', company.managerEmail || ''
  ];

  return withLock(function () {
    if (originalName) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(originalName)) {
          sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
          return 'OK';
        }
      }
    }
    sheet.appendRow(row);
    return 'OK';
  });
}

// ⭐ 법인명으로 회사설정 시트에서 해당 행을 삭제한다.
function deleteCompany(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('회사설정');
  return withLock(function () {
    deleteRowById(sheet, name, 0);
    return 'OK';
  });
}

// ⭐ 새 계약서를 작성할 때 처음 채워지는 초기값(근무장소/담당업무/근로시간 등)을 시트에서 관리.
//   설정 탭에서 고치면 다음에 새로 작성하는 계약서부터 반영된다(이미 작성 중인 계약에는 영향 없음).
var DEFAULT_FIELDS = ['workPlace', 'duties', 'dailyWorkHours', 'weeklyWorkDays', 'startTime', 'endTime', 'breakTimeMinutes', 'breakTimeStart', 'breakTimeEnd'];
var HARDCODED_DEFAULTS = {
  workPlace: '새말센터', duties: '사무직', dailyWorkHours: '8', weeklyWorkDays: '5',
  startTime: '08:00', endTime: '17:00', breakTimeMinutes: '60', breakTimeStart: '11:30', breakTimeEnd: '12:30'
};

// ⭐ "08:00" 같은 시간 형식 문자열을 시트에 그대로 쓰면 구글시트가 자동으로 시간 값(Date)으로
//   바꿔버릴 수 있어서, 읽고 쓸 때 모두 이를 방지/보정한다.
function readDefaultValue(raw, fallback) {
  if (raw instanceof Date) return Utilities.formatDate(raw, "Asia/Seoul", "HH:mm");
  if (raw === '' || raw == null) return fallback;
  return String(raw);
}

function getDefaults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('기본값설정');
  if (!sheet) {
    sheet = ss.insertSheet('기본값설정');
    sheet.appendRow(DEFAULT_FIELDS);
    var seedRow = DEFAULT_FIELDS.map(function (f) { return HARDCODED_DEFAULTS[f]; });
    var seedRange = sheet.getRange(2, 1, 1, seedRow.length);
    seedRange.setNumberFormat('@'); // 일반 텍스트로 고정 — 시간 자동변환 방지
    seedRange.setValues([seedRow]);
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return HARDCODED_DEFAULTS;
  var result = {};
  DEFAULT_FIELDS.forEach(function (f, i) {
    result[f] = readDefaultValue(data[1][i], HARDCODED_DEFAULTS[f]);
  });
  return result;
}

function saveDefaults(defaultsInput) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('기본값설정');
  if (!sheet) { getDefaults(); sheet = ss.getSheetByName('기본값설정'); }
  var row = DEFAULT_FIELDS.map(function (f) { return defaultsInput[f] != null ? String(defaultsInput[f]) : ''; });
  return withLock(function () {
    var range = sheet.getRange(2, 1, 1, row.length);
    range.setNumberFormat('@');
    range.setValues([row]);
    return 'OK';
  });
}

// ⭐ 직인 이미지: Drive에 "전자근로계약_직인" 폴더를 만들고, 그 안에 "회사설정" 시트의 법인명과
//   정확히 같은 파일명(예: "(주)케이프라이드.png")으로 이미지를 올려두면 자동으로 매칭된다.
//   (예전 Seals.html 방식은 회사명 표기가 서로 달라 매칭이 안 되고, Index.html에 포함되지도 않아
//   직인이 아예 찍히지 않는 상태였다 — 이 방식으로 대체한다.)
//   getConfig()처럼 전체를 다 내려주지 않고 요청한 법인 한 곳의 이미지만 반환한다 — 관리자가
//   최종 승인한 뒤에만(Index.html 쪽에서 그 시점에만 호출) 실제로 쓰인다.
function getSealImage(companyName) {
  var folders = DriveApp.getFoldersByName('전자근로계약_직인');
  if (!folders.hasNext()) return '';
  var folder = folders.next();
  var extensions = ['png', 'jpg', 'jpeg'];
  for (var i = 0; i < extensions.length; i++) {
    var files = folder.getFilesByName(companyName + '.' + extensions[i]);
    if (files.hasNext()) {
      try {
        var blob = files.next().getBlob();
        return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
      } catch (e) {
        Logger.log('직인 이미지 로드 실패(' + companyName + '): ' + e);
        return '';
      }
    }
  }
  return '';
}

// ⭐ Solapi API 자격증명 — 코드에 직접 적지 않고 스크립트 속성에서 읽는다.
//   설정 방법: Apps Script 편집기 > 프로젝트 설정(톱니바퀴) > 스크립트 속성에서
//   SOLAPI_API_KEY, SOLAPI_API_SECRET 값을 등록해두면 된다.
function getSolapiCredentials() {
  var props = PropertiesService.getScriptProperties();
  return {
    apiKey: props.getProperty('SOLAPI_API_KEY'),
    apiSecret: props.getProperty('SOLAPI_API_SECRET')
  };
}

function sendSolapiMessage(to, from, text) {
  var creds = getSolapiCredentials();
  if (!creds.apiKey || !creds.apiSecret) {
    Logger.log('Solapi 자격증명이 스크립트 속성에 설정되어 있지 않습니다.');
    return null;
  }

  const url = "https://api.solapi.com/messages/v4/send-many/detail";

  // 인증 헤더 생성
  const date = new Date().toISOString();
  const salt = Math.random().toString(36).substring(2, 15);
  const hmacData = date + salt;
  const signature = Utilities.computeHmacSha256Signature(hmacData, creds.apiSecret)
    .map(function (chr) { return (chr + 256).toString(16).slice(-2) }).join('');

  const headers = {
    "Authorization": `HMAC-SHA256 apiKey=${creds.apiKey}, date=${date}, salt=${salt}, signature=${signature}`,
    "Content-Type": "application/json"
  };

  const payload = { "messages": [{ "to": to.replace(/-/g, ""), "from": from.replace(/-/g, ""), "text": text }] };

  const options = { "method": "post", "headers": headers, "payload": JSON.stringify(payload), "muteHttpExceptions": true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log("Solapi Response: " + response.getContentText());
    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("Solapi Error: " + e.toString());
    return null;
  }
}
