// ============================================
// 노잇(Knowit) - 통합 자동화 시스템
// ============================================
// 기능:
// 1. 자동 일정표 업데이트
// 2. 자동 이메일 발송 (입금 확인 시)
// 3. 리마인더 발송 (모임 1일 전)
// 4. 일괄 발송 기능
// 5. 통계 대시보드
// 6. 중복 신청 방지
// ============================================

// ============ [설정 구간] ============
var CONFIG = {
  // 시트 이름
  FORM_SHEET_NAME: "폼 응답 1",
  SCHEDULE_SHEET_NAME: "일정표",
  STATS_SHEET_NAME: "통계",
  
  // 관리자 설정
  ADMIN_EMAIL: "kimtaewook86@naver.com", // 관리자 이메일 (변경 가능)
  
  // 컬럼 번호 (폼 응답 시트)
  COL: {
    TIMESTAMP: 1,    // A열: 타임스탬프
    NAME: 2,         // B열: 성함
    PHONE: 3,        // C열: 연락처
    EMAIL: 4,        // D열: 이메일
    BIRTH: 5,        // E열: 생년월일
    GENDER: 6,       // F열: 성별
    JOB_TYPE: 7,     // G열: 직업군
    JOB_DETAIL: 8,   // H열: 소속 및 직무
    ATTACHMENT: 9,   // I열: 명함/사진
    PAYMENT: 10,     // J열: 참가비 동의
    PRIVACY: 11,     // K열: 개인정보 동의
    MEMBERSHIP: 12,  // L열: 멤버십 혜택
    NICKNAME: 13,    // M열: 닉네임
    LOCATION: 14,    // N열: 장소 선택
    PAYMENT_CHECK: 15, // O열: 입금확인
    EMAIL_SENT: 16,  // P열: 링크발송
    PARTICIPATION: 17 // Q열: 참여여부
  },
  
  // 카카오톡 오픈채팅방 링크
  KAKAO_LINK: "https://open.kakao.com/o/gtsueIai"
};
// ===================================

// ============================================
// 스프레드시트 열릴 때 메뉴 추가
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('👉 [노잇] 관리자 메뉴')
    .addItem('✉️ 선택한 사람에게 확정 메일 보내기', 'sendManualEmail')
    .addSeparator()
    .addItem('📧 일괄 발송 (체크된 모든 사람)', 'sendBulkEmails')
    .addSeparator()
    .addItem('📊 통계 대시보드 생성', 'createStatsDashboard')
    .addSeparator()
    .addItem('🔍 중복 신청 확인', 'checkDuplicates')
    .addSeparator()
    .addItem('📅 일정표 수동 업데이트 (모든 신청자)', 'manualUpdateSchedule')
    .addSeparator()
    .addItem('⚙️ 자동화 트리거 설정', 'setupTriggers')
    .addToUi();
}

// ============================================
// 1. 수동 이메일 발송 (기존 기능 개선)
// ============================================
function sendManualEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  var ui = SpreadsheetApp.getUi();

  // 헤더 행 체크
  if (row === 1) {
    ui.alert("⚠️ 안내", "헤더 행이 아닌 데이터 행을 선택해주세요.", ui.ButtonSet.OK);
    return;
  }

  // 이미 발송했는지 확인
  var statusValue = sheet.getRange(row, CONFIG.COL.EMAIL_SENT).getValue();
  if (String(statusValue).indexOf("완료") !== -1) {
    ui.alert("🚫 발송 중단", "이 참가자에게는 이미 메일을 보냈습니다.\n(P열에 '발송완료' 표시됨)", ui.ButtonSet.OK);
    return;
  }

  // 이메일 주소 가져오기
  var email = sheet.getRange(row, CONFIG.COL.EMAIL).getValue();
  var name = sheet.getRange(row, CONFIG.COL.NAME).getValue();

  // 이메일 검증
  if (!email || String(email).indexOf("@") === -1) {
    ui.alert("❌ 오류", "선택하신 줄(" + row + "행)에서 유효한 이메일 주소를 찾을 수 없습니다.", ui.ButtonSet.OK);
    return;
  }

  // 발송 확인
  var response = ui.alert("📧 발송 확인", 
      "받는 사람: " + name + " (" + email + ")\n\n참가 확정 메일을 보내시겠습니까?", 
      ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    try {
      sendConfirmationEmail(email, name);
      sheet.getRange(row, CONFIG.COL.EMAIL_SENT).setValue("수동발송완료 (" + new Date().toLocaleString('ko-KR') + ")");
      ui.alert("✅ 전송 성공!", "이메일이 성공적으로 발송되었습니다.");
    } catch (e) {
      ui.alert("💥 에러 발생", e.message, ui.ButtonSet.OK);
    }
  }
}

// ============================================
// 2. 자동 이메일 발송 (입금 확인 시)
// ============================================
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    
    // 헤더 행 제외
    if (row === 1) return;
    
    // 폼 응답 시트가 아니면 무시
    if (sheet.getName() !== CONFIG.FORM_SHEET_NAME) return;
    
    // O열(입금확인) 수정 시 자동 발송
    if (col === CONFIG.COL.PAYMENT_CHECK) {
      var value = range.getValue();
      
      // 입금 확인 체크되면 자동 발송
      if (value && (String(value).indexOf("확인") !== -1 || String(value).indexOf("완료") !== -1 || value === "O" || value === "o")) {
        autoSendConfirmationEmail(row, sheet);
      }
    }
  } catch (error) {
    Logger.log("onEdit 에러: " + error.message);
  }
}

function autoSendConfirmationEmail(row, sheet) {
  try {
    var email = sheet.getRange(row, CONFIG.COL.EMAIL).getValue();
    var name = sheet.getRange(row, CONFIG.COL.NAME).getValue();
    var status = sheet.getRange(row, CONFIG.COL.EMAIL_SENT).getValue();
    
    // 이미 발송했는지 확인
    if (String(status).indexOf("완료") !== -1) {
      Logger.log("이미 발송됨: " + email);
      return;
    }
    
    // 이메일 검증
    if (!email || String(email).indexOf("@") === -1) {
      Logger.log("유효하지 않은 이메일: " + email);
      return;
    }
    
    // 이메일 발송
    sendConfirmationEmail(email, name);
    
    // P열에 발송 완료 기록
    sheet.getRange(row, CONFIG.COL.EMAIL_SENT).setValue("자동발송완료 (" + new Date().toLocaleString('ko-KR') + ")");
    Logger.log("자동 발송 완료: " + email);
    
  } catch (error) {
    Logger.log("자동 이메일 발송 실패: " + error.message);
  }
}

// ============================================
// 이메일 발송 공통 함수
// ============================================
function sendConfirmationEmail(email, name) {
  var subject = "[노잇(Knowit) 참가 확정 안내]";
  var body = "안녕하세요, " + (name || "귀하") + "님!\n\n" +
             "프리미엄 연애 모임 '노잇(Knowit)'입니다.\n" +
             "입금이 정상적으로 확인되어 최종 참가가 확정되셨습니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "원활한 모임 안내와 일정 공유를 위해\n" +
             "아래 카톡방에 입장해 주세요.\n\n" +
             "▶ 참가자 전용 채팅방 입장하기\n" +
             CONFIG.KAKAO_LINK + "\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "📌 카톡방 입장 시 닉네임 설정 양식:\n" +
             "[년생] [직업/특징] [휴대폰 뒷번호 4자리]\n\n" +
             "예시)\n" +
             "• 96년생 변호사 4567\n" +
             "• 99년생 무용수 1234\n\n" +
             "(뒷번호는 오프라인 현장 본인 대조용입니다)\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "설레는 만남을 위해 정성껏 준비하겠습니다.\n" +
             "감사합니다.\n\n" +
             "- 노잇(Knowit) 운영팀 드림";
  
  MailApp.sendEmail(email, subject, body);
}

// ============================================
// 3. 리마인더 발송 (모임 1일 전)
// ============================================
function sendDailyReminders() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FORM_SHEET_NAME);
    if (!sheet) {
      Logger.log("시트를 찾을 수 없습니다: " + CONFIG.FORM_SHEET_NAME);
      return;
    }
    
    var tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    tomorrow.setHours(0, 0, 0, 0);
    
    var lastRow = sheet.getLastRow();
    var sentCount = 0;
    
    for (var i = 2; i <= lastRow; i++) {
      var participationDate = sheet.getRange(i, CONFIG.COL.PARTICIPATION).getValue();
      var email = sheet.getRange(i, CONFIG.COL.EMAIL).getValue();
      var name = sheet.getRange(i, CONFIG.COL.NAME).getValue();
      var location = sheet.getRange(i, CONFIG.COL.LOCATION).getValue();
      
      // 날짜 형식 확인 및 변환
      if (participationDate && isSameDay(participationDate, tomorrow)) {
        if (email && String(email).indexOf("@") !== -1) {
          sendReminderEmail(email, name, location, participationDate);
          sentCount++;
        }
      }
    }
    
    Logger.log("리마인더 발송 완료: " + sentCount + "명");
    
  } catch (error) {
    Logger.log("리마인더 발송 실패: " + error.message);
  }
}

function sendReminderEmail(email, name, location, date) {
  var subject = "[노잇] 내일 모임 안내 - 리마인더 📌";
  var dateStr = Utilities.formatDate(date, "GMT+9", "yyyy년 M월 d일");
  
  var body = "안녕하세요, " + (name || "귀하") + "님!\n\n" +
             "내일 저녁에 진행되는 노잇 모임을 상기시켜 드립니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "📅 일시: " + dateStr + " (내일) 오후 8시\n" +
             "📍 장소: " + (location || "카톡방에서 안내") + "\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "카카오톡 오픈채팅방에서 자세한 위치와\n" +
             "추가 안내사항을 확인해주세요.\n\n" +
             "▶ 참가자 전용 채팅방\n" +
             CONFIG.KAKAO_LINK + "\n\n" +
             "기대되는 만남이 되시길 바랍니다!\n" +
             "감사합니다.\n\n" +
             "- 노잇(Knowit) 운영팀 드림";
  
  MailApp.sendEmail(email, subject, body);
  Logger.log("리마인더 발송: " + email);
}

function isSameDay(date1, date2) {
  if (!date1 || !date2) return false;
  
  var d1 = new Date(date1);
  var d2 = new Date(date2);
  
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

// ============================================
// 4. 일괄 발송 기능
// ============================================
function sendBulkEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // O열(입금확인)이 체크되어 있고, P열(발송상태)이 비어있는 사람들 찾기
  var lastRow = sheet.getLastRow();
  var targets = [];
  
  for (var i = 2; i <= lastRow; i++) {
    var paymentCheck = sheet.getRange(i, CONFIG.COL.PAYMENT_CHECK).getValue();
    var emailSent = sheet.getRange(i, CONFIG.COL.EMAIL_SENT).getValue();
    var email = sheet.getRange(i, CONFIG.COL.EMAIL).getValue();
    var name = sheet.getRange(i, CONFIG.COL.NAME).getValue();
    
    // 입금 확인되었고, 아직 메일 안보낸 사람
    if (paymentCheck && String(emailSent).indexOf("완료") === -1 && email && String(email).indexOf("@") !== -1) {
      targets.push({
        row: i,
        email: email,
        name: name
      });
    }
  }
  
  if (targets.length === 0) {
    ui.alert("ℹ️ 안내", "발송할 대상이 없습니다.\n\n• O열(입금확인)이 체크되어 있고\n• P열(발송상태)이 비어있는\n• 유효한 이메일을 가진 참가자를 찾을 수 없습니다.", ui.ButtonSet.OK);
    return;
  }
  
  // 발송 확인
  var response = ui.alert("📧 일괄 발송 확인", 
      "총 " + targets.length + "명에게 이메일을 발송합니다.\n\n계속하시겠습니까?", 
      ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    var successCount = 0;
    var failCount = 0;
    
    for (var i = 0; i < targets.length; i++) {
      try {
        sendConfirmationEmail(targets[i].email, targets[i].name);
        sheet.getRange(targets[i].row, CONFIG.COL.EMAIL_SENT).setValue("일괄발송완료 (" + new Date().toLocaleString('ko-KR') + ")");
        successCount++;
        Utilities.sleep(1000); // API 제한 방지를 위한 1초 대기
      } catch (e) {
        Logger.log("발송 실패: " + targets[i].email + " - " + e.message);
        failCount++;
      }
    }
    
    ui.alert("✅ 일괄 발송 완료", 
        "성공: " + successCount + "명\n" +
        "실패: " + failCount + "명", 
        ui.ButtonSet.OK);
  }
}

// ============================================
// 5. 통계 대시보드 생성
// ============================================
function createStatsDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName(CONFIG.FORM_SHEET_NAME);
  var ui = SpreadsheetApp.getUi();
  
  if (!formSheet) {
    ui.alert("❌ 오류", "'" + CONFIG.FORM_SHEET_NAME + "' 시트를 찾을 수 없습니다.", ui.ButtonSet.OK);
    return;
  }
  
  // 통계 시트 생성 또는 가져오기
  var statsSheet = ss.getSheetByName(CONFIG.STATS_SHEET_NAME);
  if (statsSheet) {
    ss.deleteSheet(statsSheet);
  }
  statsSheet = ss.insertSheet(CONFIG.STATS_SHEET_NAME);
  
  // 데이터 수집
  var lastRow = formSheet.getLastRow();
  var stats = {
    total: lastRow - 1,
    male: 0,
    female: 0,
    paid: 0,
    emailSent: 0,
    jobs: {},
    dates: {},
    ageGroups: {}
  };
  
  for (var i = 2; i <= lastRow; i++) {
    var gender = formSheet.getRange(i, CONFIG.COL.GENDER).getValue();
    var payment = formSheet.getRange(i, CONFIG.COL.PAYMENT_CHECK).getValue();
    var emailSent = formSheet.getRange(i, CONFIG.COL.EMAIL_SENT).getValue();
    var job = formSheet.getRange(i, CONFIG.COL.JOB_TYPE).getValue();
    var date = formSheet.getRange(i, CONFIG.COL.PARTICIPATION).getValue();
    var birth = String(formSheet.getRange(i, CONFIG.COL.BIRTH).getValue());
    
    // 성별 통계
    if (String(gender).indexOf("남") !== -1) stats.male++;
    if (String(gender).indexOf("여") !== -1) stats.female++;
    
    // 입금 확인
    if (payment) stats.paid++;
    
    // 이메일 발송
    if (String(emailSent).indexOf("완료") !== -1) stats.emailSent++;
    
    // 직업 통계
    if (job) {
      stats.jobs[job] = (stats.jobs[job] || 0) + 1;
    }
    
    // 날짜별 통계
    if (date) {
      var dateStr = Utilities.formatDate(new Date(date), "GMT+9", "yyyy-MM-dd");
      stats.dates[dateStr] = (stats.dates[dateStr] || 0) + 1;
    }
    
    // 연령대 통계
    if (birth && birth.length >= 2) {
      var year = parseInt(birth.substring(0, 2));
      var ageGroup = "";
      if (year >= 90 && year <= 99) ageGroup = "20대 (90년대생)";
      else if (year >= 80 && year <= 89) ageGroup = "30대 (80년대생)";
      else if (year >= 70 && year <= 79) ageGroup = "40대 (70년대생)";
      else ageGroup = "기타";
      
      stats.ageGroups[ageGroup] = (stats.ageGroups[ageGroup] || 0) + 1;
    }
  }
  
  // 대시보드 작성
  var row = 1;
  
  // 헤더
  statsSheet.getRange(row, 1).setValue("📊 노잇(Knowit) 통계 대시보드");
  statsSheet.getRange(row, 1).setFontSize(16).setFontWeight("bold");
  row += 2;
  
  // 업데이트 시간
  statsSheet.getRange(row, 1).setValue("업데이트: " + new Date().toLocaleString('ko-KR'));
  statsSheet.getRange(row, 1).setFontColor("#666666");
  row += 2;
  
  // 전체 통계
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  statsSheet.getRange(row, 1).setValue("📌 전체 통계");
  statsSheet.getRange(row, 1).setFontWeight("bold").setFontSize(12);
  row++;
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  
  statsSheet.getRange(row, 1).setValue("총 신청자 수:");
  statsSheet.getRange(row, 2).setValue(stats.total + "명");
  row++;
  
  statsSheet.getRange(row, 1).setValue("남성:");
  statsSheet.getRange(row, 2).setValue(stats.male + "명");
  row++;
  
  statsSheet.getRange(row, 1).setValue("여성:");
  statsSheet.getRange(row, 2).setValue(stats.female + "명");
  row++;
  
  statsSheet.getRange(row, 1).setValue("입금 완료:");
  statsSheet.getRange(row, 2).setValue(stats.paid + "명");
  row++;
  
  statsSheet.getRange(row, 1).setValue("이메일 발송:");
  statsSheet.getRange(row, 2).setValue(stats.emailSent + "명");
  row += 2;
  
  // 직업별 통계
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  statsSheet.getRange(row, 1).setValue("💼 직업별 통계");
  statsSheet.getRange(row, 1).setFontWeight("bold").setFontSize(12);
  row++;
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  
  for (var job in stats.jobs) {
    statsSheet.getRange(row, 1).setValue(job);
    statsSheet.getRange(row, 2).setValue(stats.jobs[job] + "명");
    row++;
  }
  row++;
  
  // 연령대별 통계
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  statsSheet.getRange(row, 1).setValue("👥 연령대별 통계");
  statsSheet.getRange(row, 1).setFontWeight("bold").setFontSize(12);
  row++;
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  
  for (var age in stats.ageGroups) {
    statsSheet.getRange(row, 1).setValue(age);
    statsSheet.getRange(row, 2).setValue(stats.ageGroups[age] + "명");
    row++;
  }
  row++;
  
  // 날짜별 통계
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  statsSheet.getRange(row, 1).setValue("📅 날짜별 참가자 수");
  statsSheet.getRange(row, 1).setFontWeight("bold").setFontSize(12);
  row++;
  statsSheet.getRange(row, 1).setValue("━━━━━━━━━━━━━━━━━━━━");
  row++;
  
  for (var date in stats.dates) {
    statsSheet.getRange(row, 1).setValue(date);
    statsSheet.getRange(row, 2).setValue(stats.dates[date] + "명");
    row++;
  }
  
  // 열 너비 조정
  statsSheet.setColumnWidth(1, 250);
  statsSheet.setColumnWidth(2, 150);
  
  ui.alert("✅ 완료", "통계 대시보드가 생성되었습니다!\n\n'" + CONFIG.STATS_SHEET_NAME + "' 시트를 확인해주세요.", ui.ButtonSet.OK);
}

// ============================================
// 6. 중복 신청 방지
// ============================================
function checkDuplicates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FORM_SHEET_NAME);
  var ui = SpreadsheetApp.getUi();
  
  if (!sheet) {
    ui.alert("❌ 오류", "'" + CONFIG.FORM_SHEET_NAME + "' 시트를 찾을 수 없습니다.", ui.ButtonSet.OK);
    return;
  }
  
  var lastRow = sheet.getLastRow();
  var emails = {};
  var phones = {};
  var duplicates = [];
  
  for (var i = 2; i <= lastRow; i++) {
    var email = String(sheet.getRange(i, CONFIG.COL.EMAIL).getValue()).trim().toLowerCase();
    var phone = String(sheet.getRange(i, CONFIG.COL.PHONE).getValue()).trim().replace(/[^0-9]/g, "");
    var name = sheet.getRange(i, CONFIG.COL.NAME).getValue();
    
    // 이메일 중복 체크
    if (email && email.indexOf("@") !== -1) {
      if (emails[email]) {
        duplicates.push({
          type: "이메일",
          value: email,
          rows: [emails[email], i],
          names: [sheet.getRange(emails[email], CONFIG.COL.NAME).getValue(), name]
        });
      } else {
        emails[email] = i;
      }
    }
    
    // 전화번호 중복 체크
    if (phone && phone.length >= 10) {
      if (phones[phone]) {
        duplicates.push({
          type: "전화번호",
          value: phone,
          rows: [phones[phone], i],
          names: [sheet.getRange(phones[phone], CONFIG.COL.NAME).getValue(), name]
        });
      } else {
        phones[phone] = i;
      }
    }
  }
  
  if (duplicates.length === 0) {
    ui.alert("✅ 중복 없음", "중복 신청이 발견되지 않았습니다.", ui.ButtonSet.OK);
    return;
  }
  
  // 중복 결과 표시
  var message = "🔍 중복 신청이 " + duplicates.length + "건 발견되었습니다:\n\n";
  
  for (var i = 0; i < Math.min(duplicates.length, 10); i++) {
    var dup = duplicates[i];
    message += (i + 1) + ". " + dup.type + " 중복\n";
    message += "   " + dup.value + "\n";
    message += "   → " + dup.rows[0] + "행: " + dup.names[0] + "\n";
    message += "   → " + dup.rows[1] + "행: " + dup.names[1] + "\n\n";
  }
  
  if (duplicates.length > 10) {
    message += "... 외 " + (duplicates.length - 10) + "건\n\n";
  }
  
  message += "※ 로그에서 전체 목록을 확인할 수 있습니다.";
  
  // 로그에 전체 목록 기록
  Logger.log("=== 중복 신청 목록 ===");
  for (var i = 0; i < duplicates.length; i++) {
    Logger.log(JSON.stringify(duplicates[i]));
  }
  
  ui.alert("⚠️ 중복 발견", message, ui.ButtonSet.OK);
}

// ============================================
// 폼 제출 시 자동 실행 (트리거 설정 필요)
// ============================================
function onFormSubmit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    // 제출된 데이터 가져오기
    var email = sheet.getRange(lastRow, CONFIG.COL.EMAIL).getValue();
    var phone = String(sheet.getRange(lastRow, CONFIG.COL.PHONE).getValue()).replace(/[^0-9]/g, "");
    
    // 중복 체크 (이메일)
    for (var i = 2; i < lastRow; i++) {
      var existingEmail = sheet.getRange(i, CONFIG.COL.EMAIL).getValue();
      if (email && email === existingEmail) {
        sheet.getRange(lastRow, CONFIG.COL.EMAIL_SENT).setValue("⚠️ 중복 신청 (이메일)");
        Logger.log("중복 신청 감지: " + email);
        return;
      }
    }
    
    // 중복 체크 (전화번호)
    for (var i = 2; i < lastRow; i++) {
      var existingPhone = String(sheet.getRange(i, CONFIG.COL.PHONE).getValue()).replace(/[^0-9]/g, "");
      if (phone && phone === existingPhone) {
        sheet.getRange(lastRow, CONFIG.COL.EMAIL_SENT).setValue("⚠️ 중복 신청 (전화번호)");
        Logger.log("중복 신청 감지: " + phone);
        return;
      }
    }
    
    // 신규 신청 표시
    sheet.getRange(lastRow, CONFIG.COL.EMAIL_SENT).setValue("✨ 신규 신청");
    
    // ✨ 새로운 기능: 일정표 자동 업데이트
    updateScheduleSheet(lastRow);
    
    Logger.log("신규 신청 접수: " + email);
    
  } catch (error) {
    Logger.log("onFormSubmit 에러: " + error.message);
  }
}

// ============================================
// 일정표 시트 자동 업데이트 (신규 기능)
// ============================================
function updateScheduleSheet(submittedRow) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName(CONFIG.FORM_SHEET_NAME);
    var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    
    // 일정표 시트가 없으면 생성
    if (!scheduleSheet) {
      scheduleSheet = ss.insertSheet(CONFIG.SCHEDULE_SHEET_NAME);
      scheduleSheet.getRange(1, 1, 1, 4).setValues([["날짜", "장소", "참가자", "상태"]]);
      scheduleSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d4af37");
      Logger.log("일정표 시트 생성 완료");
    }
    
    // 제출된 데이터 가져오기
    var locationRaw = formSheet.getRange(submittedRow, CONFIG.COL.LOCATION).getValue();
    var nickname = formSheet.getRange(submittedRow, CONFIG.COL.NICKNAME).getValue();
    var birth = String(formSheet.getRange(submittedRow, CONFIG.COL.BIRTH).getValue());
    var jobType = formSheet.getRange(submittedRow, CONFIG.COL.JOB_TYPE).getValue();
    
    // 장소가 비어있으면 스킵
    if (!locationRaw || String(locationRaw).trim() === "") {
      Logger.log("장소 정보 없음 - 스킵");
      return;
    }
    
    // 장소에서 정보 추출
    // 예: "26.02.28 천호역 근처 5:5" → 날짜: "26.02.28", 장소: "천호역 근처 5:5"
    var dateMatch = String(locationRaw).match(/^(\d{2}\.\d{2}\.\d{2})\s+(.+)/);
    
    if (!dateMatch) {
      Logger.log("장소 형식 오류: " + locationRaw);
      return;
    }
    
    var date = dateMatch[1]; // "26.02.28"
    var location = dateMatch[2]; // "천호역 근처 5:5"
    
    // 생년 추출 (앞 2자리)
    var birthYear = "";
    if (birth && birth.length >= 2) {
      birthYear = birth.substring(0, 2);
    }
    
    // 참가자 정보 포맷: "닉네임{생년} 직업"
    var participantInfo = nickname + "{" + birthYear + "} " + jobType;
    
    // 일정표에서 같은 날짜 찾기
    var lastRow = scheduleSheet.getLastRow();
    var foundRow = -1;
    
    for (var i = 2; i <= lastRow; i++) {
      var existingDate = scheduleSheet.getRange(i, 1).getValue();
      var existingLocation = scheduleSheet.getRange(i, 2).getValue();
      
      if (existingDate === date && existingLocation === location) {
        foundRow = i;
        break;
      }
    }
    
    if (foundRow > 0) {
      // 기존 행에 참가자 추가
      var existingParticipants = scheduleSheet.getRange(foundRow, 3).getValue();
      var newParticipants = existingParticipants ? existingParticipants + " / " + participantInfo : participantInfo;
      scheduleSheet.getRange(foundRow, 3).setValue(newParticipants);
      scheduleSheet.getRange(foundRow, 4).setValue("참여");
      Logger.log("기존 일정에 참가자 추가: " + date + " - " + location);
    } else {
      // 새로운 행 추가
      var newRow = lastRow + 1;
      scheduleSheet.getRange(newRow, 1).setValue(date);
      scheduleSheet.getRange(newRow, 2).setValue(location);
      scheduleSheet.getRange(newRow, 3).setValue(participantInfo);
      scheduleSheet.getRange(newRow, 4).setValue("참여");
      Logger.log("새로운 일정 생성: " + date + " - " + location);
    }
    
  } catch (error) {
    Logger.log("일정표 업데이트 에러: " + error.message);
  }
}

// ============================================
// 수동으로 모든 신청자 일정표 업데이트
// ============================================
function manualUpdateSchedule() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert("📅 일정표 업데이트", 
      "모든 신청자의 정보를 일정표 시트에 업데이트합니다.\n\n" +
      "기존 일정표 데이터는 초기화되고 새로 생성됩니다.\n\n" +
      "계속하시겠습니까?", 
      ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName(CONFIG.FORM_SHEET_NAME);
    var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    
    // 일정표 시트가 없으면 생성
    if (!scheduleSheet) {
      scheduleSheet = ss.insertSheet(CONFIG.SCHEDULE_SHEET_NAME);
    } else {
      // 기존 데이터 삭제 (헤더 제외)
      scheduleSheet.clear();
    }
    
    // 헤더 설정
    scheduleSheet.getRange(1, 1, 1, 4).setValues([["날짜", "장소", "참가자", "상태"]]);
    scheduleSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d4af37");
    
    // 날짜별, 장소별로 참가자 그룹화
    var scheduleMap = {}; // key: "날짜|장소", value: [참가자1, 참가자2, ...]
    
    var lastRow = formSheet.getLastRow();
    var processedCount = 0;
    var skippedCount = 0;
    
    for (var i = 2; i <= lastRow; i++) {
      var locationRaw = formSheet.getRange(i, CONFIG.COL.LOCATION).getValue();
      
      // 장소가 비어있으면 스킵
      if (!locationRaw || String(locationRaw).trim() === "") {
        skippedCount++;
        continue;
      }
      
      // 장소에서 정보 추출
      var dateMatch = String(locationRaw).match(/^(\d{2}\.\d{2}\.\d{2})\s+(.+)/);
      
      if (!dateMatch) {
        Logger.log("행 " + i + ": 장소 형식 오류 - " + locationRaw);
        skippedCount++;
        continue;
      }
      
      var date = dateMatch[1];
      var location = dateMatch[2];
      var key = date + "|" + location;
      
      // 참가자 정보 생성
      var nickname = formSheet.getRange(i, CONFIG.COL.NICKNAME).getValue();
      var birth = String(formSheet.getRange(i, CONFIG.COL.BIRTH).getValue());
      var jobType = formSheet.getRange(i, CONFIG.COL.JOB_TYPE).getValue();
      
      var birthYear = "";
      if (birth && birth.length >= 2) {
        birthYear = birth.substring(0, 2);
      }
      
      var participantInfo = nickname + "{" + birthYear + "} " + jobType;
      
      // 그룹에 추가
      if (!scheduleMap[key]) {
        scheduleMap[key] = {
          date: date,
          location: location,
          participants: []
        };
      }
      
      scheduleMap[key].participants.push(participantInfo);
      processedCount++;
    }
    
    // 일정표 시트에 쓰기
    var row = 2;
    for (var key in scheduleMap) {
      var schedule = scheduleMap[key];
      var participantsStr = schedule.participants.join(" / ");
      
      scheduleSheet.getRange(row, 1).setValue(schedule.date);
      scheduleSheet.getRange(row, 2).setValue(schedule.location);
      scheduleSheet.getRange(row, 3).setValue(participantsStr);
      scheduleSheet.getRange(row, 4).setValue("참여");
      
      row++;
    }
    
    // 결과 알림
    var message = "✅ 일정표 업데이트 완료!\n\n";
    message += "처리된 신청자: " + processedCount + "명\n";
    message += "스킵된 항목: " + skippedCount + "개\n";
    message += "생성된 일정: " + Object.keys(scheduleMap).length + "개";
    
    ui.alert("📅 완료", message, ui.ButtonSet.OK);
    
    Logger.log("수동 일정표 업데이트 완료: " + processedCount + "명 처리");
    
  } catch (error) {
    ui.alert("❌ 오류", "업데이트 중 오류가 발생했습니다:\n" + error.message, ui.ButtonSet.OK);
    Logger.log("수동 일정표 업데이트 에러: " + error.message);
  }
}

// ============================================
// 자동화 트리거 설정
// ============================================
function setupTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert("⚙️ 자동화 트리거 설정", 
      "다음 자동화 기능을 설정합니다:\n\n" +
      "1. 폼 제출 시 중복 체크\n" +
      "2. 매일 오전 9시 리마인더 발송\n\n" +
      "계속하시겠습니까?\n\n" +
      "※ 기존 트리거는 모두 삭제됩니다.", 
      ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // 기존 트리거 모두 삭제
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    
    // 1. 폼 제출 트리거
    ScriptApp.newTrigger('onFormSubmit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onFormSubmit()
      .create();
    
    // 2. 매일 오전 9시 리마인더 트리거
    ScriptApp.newTrigger('sendDailyReminders')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();
    
    ui.alert("✅ 설정 완료", 
        "자동화 트리거가 성공적으로 설정되었습니다!\n\n" +
        "• 폼 제출 시 자동 중복 체크\n" +
        "• 매일 오전 9시 리마인더 자동 발송\n\n" +
        "트리거는 '확장 프로그램 > Apps Script > 트리거'에서 확인 가능합니다.", 
        ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert("❌ 오류", "트리거 설정 중 오류가 발생했습니다:\n\n" + error.message, ui.ButtonSet.OK);
    Logger.log("트리거 설정 오류: " + error.message);
  }
}

// ============================================
// 테스트 함수들
// ============================================
function testEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("테스트 이메일 발송", "받을 이메일 주소를 입력하세요:", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var email = response.getResponseText();
    try {
      sendConfirmationEmail(email, "테스트");
      ui.alert("✅ 테스트 이메일 발송 완료!");
    } catch (e) {
      ui.alert("❌ 오류: " + e.message);
    }
  }
}

function testReminder() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("테스트 리마인더 발송", "받을 이메일 주소를 입력하세요:", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var email = response.getResponseText();
    try {
      sendReminderEmail(email, "테스트", "테스트 장소", new Date());
      ui.alert("✅ 테스트 리마인더 발송 완료!");
    } catch (e) {
      ui.alert("❌ 오류: " + e.message);
    }
  }
}
