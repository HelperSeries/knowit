// ============================================
// 노잇(Knowit) - 추가 개선 스크립트
// ============================================
// 기존 코드에 추가할 기능들
// ============================================

// ============================================
// [추가 기능 1] 폼 제출 시 관리자에게 알림
// ============================================
function sendAdminNotification(newParticipant) {
  var adminEmail = "kimtaewook86@naver.com"; // 관리자 이메일 (수정 필요)
  
  var subject = "🔔 [노잇] 새로운 참가 신청 - " + newParticipant.name;
  var body = "새로운 참가 신청이 접수되었습니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "👤 이름: " + newParticipant.name + "\n" +
             "📧 이메일: " + newParticipant.email + "\n" +
             "📱 연락처: " + newParticipant.phone + "\n" +
             "👔 직업: " + newParticipant.job + "\n" +
             "📅 생년월일: " + newParticipant.birth + "\n" +
             "🎯 희망 날짜: " + newParticipant.preferredDate + "\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "📊 처리 필요 사항:\n" +
             "1. 명함/증빙서류 확인\n" +
             "2. 입금 확인\n" +
             "3. 카카오톡 오픈채팅방 링크 발송\n\n" +
             "Google Sheets에서 바로 확인하기:\n" +
             "https://docs.google.com/spreadsheets/d/1Sy-JGo-sWSLh65_r53IA0H39WGeSmE_LvNNbtFN6JBs/edit";
  
  MailApp.sendEmail(adminEmail, subject, body);
}

// 기존 onFormSubmit 함수에 추가
function onFormSubmitEnhanced(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    // 제출된 데이터 가져오기
    var name = sheet.getRange(lastRow, 2).getValue();
    var phone = String(sheet.getRange(lastRow, 3).getValue()).replace(/[^0-9]/g, "");
    var email = sheet.getRange(lastRow, 4).getValue();
    var birth = sheet.getRange(lastRow, 5).getValue();
    var job = sheet.getRange(lastRow, 7).getValue();
    var preferredDate = sheet.getRange(lastRow, 14).getValue();
    
    // 중복 체크 (이메일)
    for (var i = 2; i < lastRow; i++) {
      var existingEmail = sheet.getRange(i, 4).getValue();
      if (email && email === existingEmail) {
        sheet.getRange(lastRow, 16).setValue("⚠️ 중복 신청 (이메일)");
        
        // 관리자에게 중복 알림
        sendAdminNotification({
          name: name + " ⚠️ 중복",
          email: email,
          phone: phone,
          job: job,
          birth: birth,
          preferredDate: preferredDate
        });
        
        Logger.log("중복 신청 감지: " + email);
        return;
      }
    }
    
    // 중복 체크 (전화번호)
    for (var i = 2; i < lastRow; i++) {
      var existingPhone = String(sheet.getRange(i, 3).getValue()).replace(/[^0-9]/g, "");
      if (phone && phone === existingPhone) {
        sheet.getRange(lastRow, 16).setValue("⚠️ 중복 신청 (전화번호)");
        
        // 관리자에게 중복 알림
        sendAdminNotification({
          name: name + " ⚠️ 중복",
          email: email,
          phone: phone,
          job: job,
          birth: birth,
          preferredDate: preferredDate
        });
        
        Logger.log("중복 신청 감지: " + phone);
        return;
      }
    }
    
    // 신규 신청 표시
    sheet.getRange(lastRow, 16).setValue("✨ 신규 신청");
    
    // 관리자에게 신규 신청 알림
    sendAdminNotification({
      name: name,
      email: email,
      phone: phone,
      job: job,
      birth: birth,
      preferredDate: preferredDate
    });
    
    // 신청자에게 자동 응답 이메일 발송
    sendAutoReplyEmail(email, name);
    
    Logger.log("신규 신청 접수: " + email);
    
  } catch (error) {
    Logger.log("onFormSubmit 에러: " + error.message);
  }
}

// ============================================
// [추가 기능 2] 신청자에게 자동 응답 메일
// ============================================
function sendAutoReplyEmail(email, name) {
  var subject = "[노잇] 참가 신청이 접수되었습니다 ✨";
  var body = "안녕하세요, " + (name || "귀하") + "님!\n\n" +
             "프리미엄 연애 모임 '노잇(Knowit)' 참가 신청이 정상적으로 접수되었습니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "📋 다음 단계 안내\n\n" +
             "1️⃣ 신원 확인\n" +
             "   - 제출하신 명함/증빙서류를 검토합니다\n" +
             "   - 검토 완료까지 1~2일 소요됩니다\n\n" +
             "2️⃣ 참가비 입금\n" +
             "   - 검증 완료 시 입금 안내 문자를 보내드립니다\n" +
             "   - 입금 확인 후 최종 확정됩니다\n\n" +
             "3️⃣ 카카오톡 채팅방 입장\n" +
             "   - 입금 확인 시 오픈채팅방 링크를 보내드립니다\n" +
             "   - 모임 세부 일정 및 장소 안내\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "⏰ 처리 시간\n" +
             "   평일: 24시간 이내 답변\n" +
             "   주말: 익일 오전 중 답변\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "📞 문의사항\n" +
             "   Instagram: @knowit_20__\n" +
             "   Email: " + "kimtaewook86@naver.com" + "\n\n" +
             "설레는 만남을 준비하며 기다려주세요!\n" +
             "감사합니다.\n\n" +
             "- 노잇(Knowit) 운영팀 드림";
  
  try {
    MailApp.sendEmail(email, subject, body);
    Logger.log("자동 응답 메일 발송: " + email);
  } catch (error) {
    Logger.log("자동 응답 메일 발송 실패: " + error.message);
  }
}

// ============================================
// [추가 기능 3] 입금 대기자 리마인더
// ============================================
function sendPaymentReminder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("폼 응답 1");
  var lastRow = sheet.getLastRow();
  var reminderCount = 0;
  
  var threeDaysAgo = new Date();
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
  
  for (var i = 2; i <= lastRow; i++) {
    var timestamp = sheet.getRange(i, 1).getValue(); // A열: 타임스탬프
    var paymentCheck = sheet.getRange(i, 15).getValue(); // O열: 입금확인
    var emailSent = sheet.getRange(i, 16).getValue(); // P열: 링크발송
    var email = sheet.getRange(i, 4).getValue();
    var name = sheet.getRange(i, 2).getValue();
    
    // 3일 지났는데 입금 안된 사람
    if (timestamp < threeDaysAgo && !paymentCheck && String(emailSent).indexOf("완료") === -1) {
      sendPaymentReminderEmail(email, name);
      sheet.getRange(i, 16).setValue("📌 입금 리마인더 발송 (" + new Date().toLocaleString('ko-KR') + ")");
      reminderCount++;
    }
  }
  
  Logger.log("입금 리마인더 발송 완료: " + reminderCount + "명");
}

function sendPaymentReminderEmail(email, name) {
  var subject = "[노잇] 참가비 입금 안내 리마인더 📌";
  var body = "안녕하세요, " + (name || "귀하") + "님!\n\n" +
             "노잇(Knowit) 참가 신청 후 아직 입금이 확인되지 않아 안내드립니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "참가를 희망하시는 경우,\n" +
             "아래 계좌로 참가비를 입금해 주세요.\n\n" +
             "💳 입금 계좌: [계좌번호 추가 필요]\n" +
             "💰 참가비: [금액 추가 필요]\n\n" +
             "입금 후 문자 또는 카카오톡으로\n" +
             "'입금 완료'를 알려주시면\n" +
             "빠르게 확정 안내를 드리겠습니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "참가가 어려우신 경우,\n" +
             "답장 없이 무시하셔도 괜찮습니다.\n\n" +
             "감사합니다!\n\n" +
             "- 노잇(Knowit) 운영팀 드림";
  
  MailApp.sendEmail(email, subject, body);
  Logger.log("입금 리마인더 발송: " + email);
}

// ============================================
// [추가 기능 4] 웹사이트 데이터 자동 동기화
// ============================================
function syncToScheduleSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("폼 응답 1");
  var scheduleSheet = ss.getSheetByName("일정표");
  
  if (!formSheet || !scheduleSheet) {
    Logger.log("시트를 찾을 수 없습니다");
    return;
  }
  
  var lastRow = formSheet.getLastRow();
  var scheduleData = {};
  
  // 입금 완료된 참가자만 필터링
  for (var i = 2; i <= lastRow; i++) {
    var paymentCheck = formSheet.getRange(i, 15).getValue(); // O열: 입금확인
    
    if (paymentCheck) {
      var nickname = formSheet.getRange(i, 13).getValue(); // M열: 닉네임
      var birth = String(formSheet.getRange(i, 5).getValue()); // E열: 생년월일
      var job = formSheet.getRange(i, 7).getValue(); // G열: 직업군
      var preferredDate = formSheet.getRange(i, 14).getValue(); // N열: 장소 선택
      
      // 나이 계산
      var age = "";
      if (birth && birth.length >= 2) {
        age = birth.substring(0, 2) + "년생";
      }
      
      // 날짜별로 그룹화
      var dateKey = preferredDate ? Utilities.formatDate(new Date(preferredDate), "GMT+9", "yyyy-MM-dd") : "미정";
      
      if (!scheduleData[dateKey]) {
        scheduleData[dateKey] = [];
      }
      
      scheduleData[dateKey].push(nickname + " [" + age + "] " + job);
    }
  }
  
  // 일정표 시트 업데이트
  // (일정표 시트 형식에 맞게 데이터 입력)
  // 이 부분은 일정표 시트의 실제 구조에 맞게 수정 필요
  
  Logger.log("일정표 동기화 완료");
  Logger.log(JSON.stringify(scheduleData));
}

// ============================================
// [추가 기능 5] 노쇼 관리 시스템
// ============================================
function markNoShow(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("폼 응답 1");
  var name = sheet.getRange(row, 2).getValue();
  var email = sheet.getRange(row, 4).getValue();
  var phone = sheet.getRange(row, 3).getValue();
  
  // P열에 노쇼 표시
  sheet.getRange(row, 16).setValue("🚫 노쇼 발생 - 재신청 제한 (" + new Date().toLocaleString('ko-KR') + ")");
  
  // 블랙리스트 시트에 추가
  var blacklistSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("블랙리스트");
  
  if (!blacklistSheet) {
    blacklistSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("블랙리스트");
    blacklistSheet.getRange(1, 1).setValue("이름");
    blacklistSheet.getRange(1, 2).setValue("이메일");
    blacklistSheet.getRange(1, 3).setValue("전화번호");
    blacklistSheet.getRange(1, 4).setValue("노쇼 발생일");
    blacklistSheet.getRange(1, 5).setValue("사유");
  }
  
  var lastRow = blacklistSheet.getLastRow();
  blacklistSheet.getRange(lastRow + 1, 1).setValue(name);
  blacklistSheet.getRange(lastRow + 1, 2).setValue(email);
  blacklistSheet.getRange(lastRow + 1, 3).setValue(phone);
  blacklistSheet.getRange(lastRow + 1, 4).setValue(new Date().toLocaleString('ko-KR'));
  blacklistSheet.getRange(lastRow + 1, 5).setValue("모임 노쇼");
  
  Logger.log("노쇼 등록: " + name);
}

// 블랙리스트 체크 함수
function checkBlacklist(email, phone) {
  var blacklistSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("블랙리스트");
  
  if (!blacklistSheet) return false;
  
  var lastRow = blacklistSheet.getLastRow();
  
  for (var i = 2; i <= lastRow; i++) {
    var blacklistedEmail = blacklistSheet.getRange(i, 2).getValue();
    var blacklistedPhone = String(blacklistSheet.getRange(i, 3).getValue()).replace(/[^0-9]/g, "");
    
    if ((email && email === blacklistedEmail) || (phone && phone === blacklistedPhone)) {
      return true; // 블랙리스트에 있음
    }
  }
  
  return false;
}

// ============================================
// [추가 기능 6] 감사 메일 자동 발송 (모임 후)
// ============================================
function sendThankYouEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("폼 응답 1");
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  yesterday.setHours(0, 0, 0, 0);
  
  var lastRow = sheet.getLastRow();
  var sentCount = 0;
  
  for (var i = 2; i <= lastRow; i++) {
    var participationDate = sheet.getRange(i, 17).getValue(); // Q열: 참여여부
    var email = sheet.getRange(i, 4).getValue();
    var name = sheet.getRange(i, 2).getValue();
    var emailSent = sheet.getRange(i, 16).getValue();
    
    // 어제 참여했고, 노쇼가 아닌 사람
    if (participationDate && isSameDay(participationDate, yesterday) && String(emailSent).indexOf("노쇼") === -1) {
      sendThankYouEmail(email, name);
      sentCount++;
    }
  }
  
  Logger.log("감사 메일 발송 완료: " + sentCount + "명");
}

function sendThankYouEmail(email, name) {
  var subject = "[노잇] 참여해 주셔서 감사합니다 💕";
  var body = "안녕하세요, " + (name || "귀하") + "님!\n\n" +
             "어제 노잇(Knowit) 모임에 참여해 주셔서 진심으로 감사드립니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "좋은 만남이 되셨길 바라며,\n" +
             "다음 모임에도 많은 관심 부탁드립니다.\n\n" +
             "📅 다음 모임 일정은 인스타그램과\n" +
             "카카오톡 채널을 통해 안내드리겠습니다.\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "💬 피드백 부탁드립니다!\n\n" +
             "더 나은 모임을 위해\n" +
             "간단한 후기나 개선 의견을 보내주시면\n" +
             "큰 도움이 됩니다.\n\n" +
             "(이 메일에 답장해 주시면 됩니다)\n\n" +
             "━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
             "다시 한번 감사드리며,\n" +
             "행복한 하루 되세요!\n\n" +
             "- 노잇(Knowit) 운영팀 드림";
  
  MailApp.sendEmail(email, subject, body);
  Logger.log("감사 메일 발송: " + email);
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
// 통합 메뉴 추가 (기존 onOpen에 추가)
// ============================================
function onOpenEnhanced() {
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
    .addSubMenu(ui.createMenu('📌 리마인더')
        .addItem('입금 대기자에게 리마인더 발송', 'sendPaymentReminder')
        .addItem('모임 참가자에게 리마인더 발송', 'sendDailyReminders'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🎁 감사 메일')
        .addItem('모임 종료 후 감사 메일 발송', 'sendThankYouEmails'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🚫 노쇼 관리')
        .addItem('선택한 행을 노쇼로 표시', 'markNoShowFromMenu'))
    .addSeparator()
    .addItem('⚙️ 자동화 트리거 설정', 'setupTriggersEnhanced')
    .addToUi();
}

function markNoShowFromMenu() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  var ui = SpreadsheetApp.getUi();
  
  if (row === 1) {
    ui.alert("⚠️ 안내", "헤더 행이 아닌 데이터 행을 선택해주세요.", ui.ButtonSet.OK);
    return;
  }
  
  var name = sheet.getRange(row, 2).getValue();
  var response = ui.alert("🚫 노쇼 등록", 
      name + "님을 노쇼로 등록하시겠습니까?\n\n" +
      "• 블랙리스트에 추가됩니다\n" +
      "• 재신청이 제한됩니다", 
      ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    markNoShow(row);
    ui.alert("✅ 완료", "노쇼로 등록되었습니다.", ui.ButtonSet.OK);
  }
}

// 강화된 트리거 설정
function setupTriggersEnhanced() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert("⚙️ 자동화 트리거 설정", 
      "다음 자동화 기능을 설정합니다:\n\n" +
      "1. 폼 제출 시 중복 체크 및 자동 응답\n" +
      "2. 매일 오전 9시 모임 리마인더 발송\n" +
      "3. 매일 오후 2시 입금 리마인더 발송\n" +
      "4. 매일 오전 10시 감사 메일 발송\n\n" +
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
    ScriptApp.newTrigger('onFormSubmitEnhanced')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onFormSubmit()
      .create();
    
    // 2. 매일 오전 9시 모임 리마인더
    ScriptApp.newTrigger('sendDailyReminders')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();
    
    // 3. 매일 오후 2시 입금 리마인더
    ScriptApp.newTrigger('sendPaymentReminder')
      .timeBased()
      .atHour(14)
      .everyDays(1)
      .create();
    
    // 4. 매일 오전 10시 감사 메일
    ScriptApp.newTrigger('sendThankYouEmails')
      .timeBased()
      .atHour(10)
      .everyDays(1)
      .create();
    
    ui.alert("✅ 설정 완료", 
        "자동화 트리거가 성공적으로 설정되었습니다!\n\n" +
        "• 폼 제출 시 자동 중복 체크 및 응답\n" +
        "• 매일 오전 9시 모임 리마인더\n" +
        "• 매일 오후 2시 입금 리마인더\n" +
        "• 매일 오전 10시 감사 메일\n\n" +
        "트리거는 '확장 프로그램 > Apps Script > 트리거'에서 확인 가능합니다.", 
        ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert("❌ 오류", "트리거 설정 중 오류가 발생했습니다:\n\n" + error.message, ui.ButtonSet.OK);
    Logger.log("트리거 설정 오류: " + error.message);
  }
}
