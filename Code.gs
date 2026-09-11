/**
 * Google Apps Script - 나의 테니스 일지 백엔드 (Code.gs)
 * 다중 사진 업로드 + 댓글 + 텔레그램(Telegram) 봇 푸시 알림 + 스트링 교체 관리 + Gemini 3.6 Flash AI 스트링 교체 수명 진단
 */

// 🔔 텔레그램 봇 알림 설정 (마라톤 트레이닝과 동일 설정)
var TELEGRAM_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN") || "";
var TELEGRAM_CHAT_ID = PropertiesService.getScriptProperties().getProperty("TELEGRAM_CHAT_ID") || "";

// 💚 LINE(라인) 메신저 알림 봇 설정 (댓글 등록 시 라인 실시간 알림 발송)
var LINE_CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_CHANNEL_ACCESS_TOKEN") || "";
var LINE_USER_ID = PropertiesService.getScriptProperties().getProperty("LINE_USER_ID") || "";

// 🔑 Google Gemini AI API Key (신규 활성화 키)
var DEFAULT_GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || ["AQ", "Ab8RN6INIjH3Md1wSsrIG66nwP_70P-ImQK5eUZiLrfYj5j73g"].join("."); 

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result = { success: false };
    try {
      if (action === 'initDatabase') {
        result = initDatabase();
      } else if (action === 'getTennisData') {
        result = getTennisData();
      } else if (action === 'getTennisStringData') {
        result = getTennisStringData();
      } else if (action === 'testTennisTelegramAlert') {
        result = testTennisTelegramAlert();
      } else if (action === 'testTennisLineAlert') {
        result = testTennisLineAlert();
      }
    } catch(err) {
      result = { success: false, error: err.message };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('나의 테니스 일지 | Tennis Log & AI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function doPost(e) {
  var result = { success: false };
  try {
    var postData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch(pErr) {
        postData = e.parameter || {};
      }
    } else if (e && e.parameter) {
      postData = e.parameter;
    }

    var action = postData.action;
    var payload = postData.payload || postData;

    if (action === 'saveTennisLog') {
      result = saveTennisLog(payload);
    } else if (action === 'updateTennisLog') {
      result = updateTennisLog(payload);
    } else if (action === 'deleteTennisLog') {
      result = deleteTennisLog(payload.rowIndex || payload);
    } else if (action === 'addTennisComment') {
      result = addTennisComment(payload.rowIndex, payload.commentData);
    } else if (action === 'updateTennisComment') {
      result = updateTennisComment(payload.rowIndex, payload.commentId, payload.newText);
    } else if (action === 'deleteTennisComment') {
      result = deleteTennisComment(payload.rowIndex, payload.commentId);
    } else if (action === 'saveStringRecord') {
      result = saveStringRecord(payload);
    } else if (action === 'deleteStringRecord') {
      result = deleteStringRecord(payload.rowIndex || payload);
    } else if (action === 'analyzeTennisStringAI') {
      result = analyzeTennisStringAI(payload);
    } else if (action === 'generateGeminiTennisSummary') {
      result = generateGeminiTennisSummary(payload);
    } else if (action === 'analyzeTennisVideoFormAI') {
      result = analyzeTennisVideoFormAI(payload);
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// [1. 시트 헬퍼 함수군]
// -------------------------------------------------------------
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('테니스일지');
  if (!sheet) {
    sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    sheet.setName('테니스일지');
  }
  
  if (sheet.getLastRow() === 0) {
    var headers = [
      "일자", "구분", "운동제목", "시간(분)", "강도(RPE)", 
      "경기결과", "스코어", "함께한사람", "코트종류", "장소", 
      "집중기술", "컨디션/수기메모", "사진ID목록", "댓글", "AI코칭분석"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#16a34a').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // 기존 시트에 15열(AI코칭분석) 헤더가 없으면 자동 추가
    if (sheet.getLastColumn() < 15 || !sheet.getRange(1, 15).getValue()) {
      sheet.getRange(1, 15).setValue('AI코칭분석').setBackground('#16a34a').setFontColor('#ffffff').setFontWeight('bold');
    }
  }
  return sheet;
}

function getStringSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('스트링교체');
  if (!sheet) {
    sheet = ss.insertSheet('스트링교체');
  }
  if (sheet.getLastRow() === 0) {
    var headers = [
      "교체일자", "라켓모델", "스트링명", "스트링종류", "텐션(lbs)", 
      "교체비용", "스트링샵", "메모", "등록일시"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#16a34a').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizeDateString(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var str = String(dateVal).trim();
  var clean = str.replace(/[^\d]/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
  var parts = clean.split('-');
  if (parts.length >= 3) {
    var y = parts[0];
    var m = parts[1].length === 1 ? '0' + parts[1] : parts[1];
    var d = parts[2].length === 1 ? '0' + parts[2] : parts[2];
    return y + '-' + m + '-' + d;
  }
  return str;
}

function getOrCreatePhotoFolder() {
  var folderName = "Tennis_Log_Photos";
  var folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    folders = DriveApp.getFoldersByName("테니스일지_사진");
  }
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return folder;
}

// -------------------------------------------------------------
// [2. 테니스 일지 조회 & 통계]
// -------------------------------------------------------------
function getTennisData() {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) {
      return { logs: [], overview: getEmptyOverview(), monthly_trends: [] };
    }

    var logs = [];
    var totalHours = 0;
    var lessonCount = 0;
    var gameCount = 0;
    var practiceCount = 0;
    var wins = 0;
    var losses = 0;
    var monthlyMap = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[0]) continue;

      var dateStr = normalizeDateString(row[0]);
      var category = String(row[1] || '레슨').trim();
      var title = String(row[2] || '').trim();
      var dur = Number(row[3]) || 0;
      var rpe = Number(row[4]) || 5;
      var result = String(row[5] || '').trim();
      var score = String(row[6] || '').trim();
      var players = String(row[7] || '').trim();
      var courtType = String(row[8] || '하드코트').trim();
      var location = String(row[9] || '').trim();
      var focusSkill = String(row[10] || '').trim();
      var notes = String(row[11] || '').trim();
      var photoIdsStr = String(row[12] || '').trim();
      var commentsRaw = row[13] || '';
      var aiSummary = String(row[14] || '').trim();

      totalHours += dur / 60;

      if (category === '레슨') lessonCount++;
      else if (category === '클럽게임' || category === '게임') {
        gameCount++;
        if (result === '승' || result === '승리') wins++;
        else if (result === '패' || result === '패배') losses++;
      } else if (category === '랠리/연습' || category === '연습') {
        practiceCount++;
      }

      // 월별 집계
      var monthStr = dateStr.length >= 7 ? dateStr.substring(0, 7) : "기타";
      if (!monthlyMap[monthStr]) monthlyMap[monthStr] = { "레슨": 0, "클럽게임": 0, "랠리/연습": 0, "승": 0, "패": 0 };
      monthlyMap[monthStr][category] = (monthlyMap[monthStr][category] || 0) + 1;
      if (result === '승' || result === '승리') monthlyMap[monthStr]["승"]++;
      if (result === '패' || result === '패배') monthlyMap[monthStr]["패"]++;

      var photoIds = [];
      function isDriveId(str) {
        return /^(vid:)?[a-zA-Z0-9_\-]{20,50}$/.test(str);
      }

      if (photoIdsStr) {
        photoIds = photoIdsStr.split(',').map(function(id) { return id.trim(); }).filter(isDriveId);
      }

      // 만약 aiSummary에 사진/동영상 Drive ID가 잘못 들어가 있는 경우 자동 복구
      if (aiSummary && (aiSummary.indexOf('1') === 0 || aiSummary.indexOf('vid:') === 0 || aiSummary.indexOf(',') > 0)) {
        var potentialIds = aiSummary.split(',').map(function(id) { return id.trim(); }).filter(isDriveId);
        if (potentialIds.length > 0 && potentialIds.length === aiSummary.split(',').length) {
          if (photoIds.length === 0) {
            photoIds = potentialIds;
          }
          aiSummary = '';
        }
      }

      var comments = [];
      if (commentsRaw) {
        try {
          comments = typeof commentsRaw === 'string' ? JSON.parse(commentsRaw) : commentsRaw;
          if (!Array.isArray(comments)) comments = [];
        } catch(e) {
          comments = [];
        }
      }

      logs.push({
        rowIndex: i + 1,
        date: dateStr,
        category: category,
        title: title,
        duration_minutes: dur,
        intensity: rpe,
        result: result,
        score: score,
        players: players,
        court_type: courtType,
        location: location,
        focus_skill: focusSkill,
        notes: notes,
        photo_ids: photoIds,
        comments: comments,
        ai_summary: aiSummary
      });
    }

    logs.sort(function(a, b) {
      var dateCompare = (b.date || "").localeCompare(a.date || "");
      if (dateCompare !== 0) return dateCompare;
      return (b.rowIndex || 0) - (a.rowIndex || 0);
    });

    var totalGames = wins + losses;
    var winRate = totalGames > 0 ? Math.round((wins / totalGames) * 100) : 0;

    var months = Object.keys(monthlyMap).sort();
    var monthlyTrends = [];
    months.forEach(function(m) {
      monthlyTrends.push({ month: m, data: monthlyMap[m] });
    });

    return {
      logs: logs,
      overview: {
        total_activities: logs.length,
        lesson_count: lessonCount,
        game_count: gameCount,
        practice_count: practiceCount,
        wins: wins,
        losses: losses,
        win_rate: winRate,
        total_hours: Math.round(totalHours * 10) / 10
      },
      monthly_trends: monthlyTrends
    };
  } catch (e) {
    Logger.log("getTennisData error: " + e.message);
    return { logs: [], overview: getEmptyOverview(), monthly_trends: [], error: e.message };
  }
}

// -------------------------------------------------------------
// [3. 테니스 일지 CRUD]
// -------------------------------------------------------------
function saveTennisLog(item) {
  try {
    var sheet = getSheet();
    var dateStr = item.date;
    var category = item.category || '레슨';
    var title = item.title || (category === '레슨' ? (item.focus_skill ? item.focus_skill + ' 레슨' : '테니스 레슨') : '클럽 게임');
    var dur = item.duration_minutes || 0;
    var rpe = item.intensity || 5;
    var result = item.result || '';
    var score = item.score || '';
    var players = item.players || '';
    var courtType = item.court_type || '하드코트';
    var loc = item.location || '';
    var focusSkill = item.focus_skill || '';
    var notes = item.notes || '';
    var aiSummary = item.ai_summary || '';
    var photoIds = [];

    if (item.photos && item.photos.length > 0) {
      try {
        var folder = getOrCreatePhotoFolder();
        for (var p = 0; p < item.photos.length; p++) {
          var base64Data = item.photos[p];
          if (base64Data && base64Data.indexOf('data:') === 0) {
            var parts = base64Data.split(',');
            var contentType = parts[0].split(';')[0].replace('data:', '');
            var isVideo = contentType.indexOf('video/') === 0;
            var ext = 'jpg';
            if (isVideo) {
              if (contentType.indexOf('quicktime') !== -1 || contentType.indexOf('mov') !== -1) ext = 'mov';
              else if (contentType.indexOf('webm') !== -1) ext = 'webm';
              else ext = 'mp4';
            }
            var bytes = Utilities.base64Decode(parts[1]);
            var fileName = (isVideo ? 'tennis_video_' : 'tennis_photo_') + dateStr + '_' + (p + 1) + '_' + (new Date().getTime()) + '.' + ext;
            var blob = Utilities.newBlob(bytes, contentType, fileName);
            var file = folder.createFile(blob);
            try {
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch(eShare) {}
            var savedId = (isVideo ? 'vid:' : '') + file.getId();
            photoIds.push(savedId);
          }
        }
      } catch(e) {
        Logger.log('Tennis media upload error: ' + e.message);
      }
    }

    var photoIdsStr = photoIds.join(',');

    sheet.appendRow([
      dateStr, category, title, dur, rpe,
      result, score, players, courtType, loc,
      focusSkill, notes, photoIdsStr, "[]", aiSummary
    ]);

    sendTennisLogTelegramNotification(item);
    return { success: true };
  } catch (e) {
    Logger.log("saveTennisLog error: " + e.message);
    throw new Error("저장 실패: " + e.message);
  }
}

function updateTennisLog(item) {
  try {
    var rowIndex = item.rowIndex;
    if (!rowIndex || rowIndex < 2) throw new Error("유효하지 않은 일지 번호입니다.");
    var sheet = getSheet();
    var dateStr = item.date;
    var category = item.category || '레슨';
    var title = item.title || (category === '레슨' ? (item.focus_skill ? item.focus_skill + ' 레슨' : '테니스 레슨') : '클럽 게임');
    var dur = item.duration_minutes || 0;
    var rpe = item.intensity || 5;
    var result = item.result || '';
    var score = item.score || '';
    var players = item.players || '';
    var courtType = item.court_type || '하드코트';
    var loc = item.location || '';
    var focusSkill = item.focus_skill || '';
    var notes = item.notes || '';
    var aiSummary = item.ai_summary || '';
    var photoIds = item.existingPhotoIds || [];

    if (item.photos && item.photos.length > 0) {
      try {
        var folder = getOrCreatePhotoFolder();
        for (var p = 0; p < item.photos.length; p++) {
          var base64Data = item.photos[p];
          if (base64Data && base64Data.indexOf('data:') === 0) {
            var parts = base64Data.split(',');
            var contentType = parts[0].split(';')[0].replace('data:', '');
            var isVideo = contentType.indexOf('video/') === 0;
            var ext = 'jpg';
            if (isVideo) {
              if (contentType.indexOf('quicktime') !== -1 || contentType.indexOf('mov') !== -1) ext = 'mov';
              else if (contentType.indexOf('webm') !== -1) ext = 'webm';
              else ext = 'mp4';
            }
            var bytes = Utilities.base64Decode(parts[1]);
            var fileName = (isVideo ? 'tennis_video_' : 'tennis_photo_') + dateStr + '_' + (p + 1) + '_' + (new Date().getTime()) + '.' + ext;
            var blob = Utilities.newBlob(bytes, contentType, fileName);
            var file = folder.createFile(blob);
            try {
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch(eShare) {}
            var savedId = (isVideo ? 'vid:' : '') + file.getId();
            photoIds.push(savedId);
          }
        }
      } catch(e) {
        Logger.log('Tennis media upload error: ' + e.message);
      }
    }

    var photoIdsStr = photoIds.join(',');

    sheet.getRange(rowIndex, 1, 1, 13).setValues([[
      dateStr, category, title, dur, rpe,
      result, score, players, courtType, loc,
      focusSkill, notes, photoIdsStr
    ]]);
    sheet.getRange(rowIndex, 15).setValue(aiSummary);

    return { success: true };
  } catch (e) {
    Logger.log("updateTennisLog error: " + e.message);
    throw new Error("수정 실패: " + e.message);
  }
}

function deleteTennisLog(rowIndex) {
  try {
    if (!rowIndex || rowIndex < 2) throw new Error("유효하지 않은 일지 번호입니다.");
    var sheet = getSheet();
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    Logger.log("deleteTennisLog error: " + e.message);
    throw new Error("삭제 실패: " + e.message);
  }
}

// -------------------------------------------------------------
// [4. 댓글 관리 API]
// -------------------------------------------------------------
function addTennisComment(rowIndex, commentData) {
  try {
    if (!rowIndex || rowIndex < 2) throw new Error("유효하지 않은 일지 번호입니다.");
    var sheet = getSheet();
    var cell = sheet.getRange(rowIndex, 14);
    var rawVal = cell.getValue();
    var comments = [];
    if (rawVal) {
      try {
        comments = typeof rawVal === 'string' ? JSON.parse(rawVal) : rawVal;
        if (!Array.isArray(comments)) comments = [];
      } catch(e) { comments = []; }
    }

    var newComment = {
      id: Utilities.getUuid(),
      author: String(commentData.author || '작성자').trim(),
      text: String(commentData.text || '').trim(),
      createdAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
    };
    if (!newComment.text) throw new Error("댓글 내용을 입력해주세요.");

    comments.push(newComment);
    cell.setValue(JSON.stringify(comments));

    var rowData = sheet.getRange(rowIndex, 1, 1, 3).getValues()[0];
    var logDate = normalizeDateString(rowData[0]);
    var logTitle = rowData[2] || rowData[1] || '테니스 일지';
    
    // 💚 LINE(라인) 메신저: '징징이'가 댓글을 달았을 때는 라인으로만 알림 발송 (텔레그램 제외)
    if (newComment.author && newComment.author.indexOf('징징이') !== -1) {
      sendTennisCommentLineNotification(logDate, logTitle, newComment);
    } else {
      // 🤖 텔레그램: 징징이가 아닐 때(다람이 등)만 텔레그램으로 알림 발송
      sendTennisCommentTelegramNotification(logDate, logTitle, newComment);
    }

    return { success: true, comments: comments };
  } catch(e) {
    Logger.log("addTennisComment error: " + e.message);
    throw new Error("댓글 등록 실패: " + e.message);
  }
}

function updateTennisComment(rowIndex, commentId, newText) {
  try {
    if (!rowIndex || rowIndex < 2 || !commentId) throw new Error("유효하지 않은 댓글 정보입니다.");
    var sheet = getSheet();
    var cell = sheet.getRange(rowIndex, 14);
    var rawVal = cell.getValue();
    var comments = [];
    if (rawVal) {
      try {
        comments = typeof rawVal === 'string' ? JSON.parse(rawVal) : rawVal;
        if (!Array.isArray(comments)) comments = [];
      } catch(e) { comments = []; }
    }

    var updated = false;
    for (var i = 0; i < comments.length; i++) {
      if (comments[i].id === commentId) {
        comments[i].text = String(newText || '').trim();
        comments[i].updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        updated = true;
        break;
      }
    }

    if (!updated) throw new Error("수정할 댓글을 찾을 수 없습니다.");
    cell.setValue(JSON.stringify(comments));
    return { success: true, comments: comments };
  } catch(e) {
    Logger.log("updateTennisComment error: " + e.message);
    throw new Error("댓글 수정 실패: " + e.message);
  }
}

function deleteTennisComment(rowIndex, commentId) {
  try {
    if (!rowIndex || rowIndex < 2 || !commentId) throw new Error("유효하지 않은 댓글 정보입니다.");
    var sheet = getSheet();
    var cell = sheet.getRange(rowIndex, 14);
    var rawVal = cell.getValue();
    var comments = [];
    if (rawVal) {
      try {
        comments = typeof rawVal === 'string' ? JSON.parse(rawVal) : rawVal;
        if (!Array.isArray(comments)) comments = [];
      } catch(e) { comments = []; }
    }

    comments = comments.filter(function(c) { return c.id !== commentId; });
    cell.setValue(JSON.stringify(comments));
    return { success: true, comments: comments };
  } catch(e) {
    Logger.log("deleteTennisComment error: " + e.message);
    throw new Error("댓글 삭제 실패: " + e.message);
  }
}

// -------------------------------------------------------------
// [5. 🎾 스트링 교체 관리 & 데이터 분석]
// -------------------------------------------------------------
function getTennisStringData() {
  try {
    var stringSheet = getStringSheet();
    var sData = stringSheet.getDataRange().getValues();
    var stringLogs = [];

    if (sData && sData.length > 1) {
      for (var i = 1; i < sData.length; i++) {
        var r = sData[i];
        if (!r || !r[0]) continue;
        var dateStr = normalizeDateString(r[0]);
        stringLogs.push({
          rowIndex: i + 1,
          date: dateStr,
          racket: String(r[1] || '기본 라켓').trim(),
          string_name: String(r[2] || '미지정').trim(),
          string_type: String(r[3] || '폴리').trim(),
          tension: String(r[4] || '50').trim(),
          cost: Number(r[5]) || 0,
          shop_name: String(r[6] || '').trim(),
          notes: String(r[7] || '').trim(),
          createdAt: r[8] ? String(r[8]) : ''
        });
      }
    }

    // 날짜 내림차순 정렬
    stringLogs.sort(function(a, b) {
      var dComp = (b.date || '').localeCompare(a.date || '');
      if (dComp !== 0) return dComp;
      return (b.rowIndex || 0) - (a.rowIndex || 0);
    });

    var activeString = stringLogs.length > 0 ? stringLogs[0] : null;
    var statsSince = {
      daysElapsed: 0,
      totalHours: 0,
      totalMinutes: 0,
      sessionCount: 0,
      gameCount: 0,
      lessonCount: 0,
      practiceCount: 0,
      avgIntensity: 0,
      hardSessions: 0,
      courtTypes: {},
      logsSince: []
    };

    if (activeString && activeString.date) {
      var tennisSheet = getSheet();
      var tData = tennisSheet.getDataRange().getValues();
      var activeDate = activeString.date;
      var sumIntensity = 0;

      var tNow = new Date();
      var tStrung = new Date(activeDate);
      if (!isNaN(tStrung.getTime())) {
        var diffTime = Math.abs(tNow - tStrung);
        statsSince.daysElapsed = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      }

      if (tData && tData.length > 1) {
        for (var j = 1; j < tData.length; j++) {
          var row = tData[j];
          if (!row || !row[0]) continue;
          var logDate = normalizeDateString(row[0]);
          if (logDate >= activeDate) {
            var dur = Number(row[3]) || 0;
            var rpe = Number(row[4]) || 5;
            var cat = String(row[1] || '레슨').trim();
            var cType = String(row[8] || '하드코트').trim();
            var nts = String(row[11] || '').trim();

            statsSince.sessionCount++;
            statsSince.totalMinutes += dur;
            sumIntensity += rpe;
            if (rpe >= 7) statsSince.hardSessions++;

            if (cat === '레슨') statsSince.lessonCount++;
            else if (cat === '클럽게임' || cat === '게임') statsSince.gameCount++;
            else statsSince.practiceCount++;

            statsSince.courtTypes[cType] = (statsSince.courtTypes[cType] || 0) + 1;
            statsSince.logsSince.push({
              date: logDate,
              category: cat,
              duration: dur,
              intensity: rpe,
              notes: nts
            });
          }
        }
      }

      statsSince.totalHours = Math.round((statsSince.totalMinutes / 60) * 10) / 10;
      statsSince.avgIntensity = statsSince.sessionCount > 0 ? Math.round((sumIntensity / statsSince.sessionCount) * 10) / 10 : 0;
    }

    return {
      stringLogs: stringLogs,
      activeString: activeString,
      statsSince: statsSince
    };
  } catch(e) {
    Logger.log("getTennisStringData error: " + e.message);
    throw new Error("스트링 데이터 조회 실패: " + e.message);
  }
}

function saveStringRecord(item) {
  try {
    var sheet = getStringSheet();
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    sheet.appendRow([
      item.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      item.racket || '기본 라켓',
      item.string_name || '미지정',
      item.string_type || '폴리',
      item.tension || '50',
      Number(item.cost) || 0,
      item.shop_name || '',
      item.notes || '',
      nowStr
    ]);

    try {
      sendTennisStringTelegramNotification(item);
    } catch(err) {}

    return { success: true };
  } catch(e) {
    Logger.log("saveStringRecord error: " + e.message);
    throw new Error("스트링 기록 저장 실패: " + e.message);
  }
}

function deleteStringRecord(rowIndex) {
  try {
    var sheet = getStringSheet();
    var idx = parseInt(rowIndex);
    if (!idx || idx < 2) throw new Error("유효하지 않은 행입니다.");
    sheet.deleteRow(idx);
    return { success: true };
  } catch(e) {
    Logger.log("deleteStringRecord error: " + e.message);
    throw new Error("스트링 기록 삭제 실패: " + e.message);
  }
}

// -------------------------------------------------------------
// [6. 🤖 GEMINI AI 스트링 교체 수명 진단 & 일지 코칭 요약]
// -------------------------------------------------------------
function analyzeTennisStringAI(item) {
  try {
    var apiKey = item.geminiApiKey || DEFAULT_GEMINI_API_KEY;
    var stringInfo = item.activeString;
    var stats = item.statsSince;

    if (!stringInfo) {
      return {
        success: false,
        analysis: "등록된 스트링 교체 이력이 없습니다. 먼저 '➕ 새 스트링 교체 등록' 버튼을 눌러 현재 라켓의 스트링 교체일을 등록해주세요! 🎾"
      };
    }

    var prompt = "너는 세계적인 프로 테니스 스트링 기술자이자 동호인 맞춤 스포츠 의학/장비 전문 수석 테니스 코치야.\n" +
      "사용자(여성 테니스 동호인)가 입력한 라켓/스트링 사양과 스트링 장착 이후 기록된 '테니스 일지(플레이 시간, 강도, 게임 수, 일지 메모/후기)'를 종합 분석하여, **【지금 스트링을 교체해야 하는지 여부】**를 명확하고 친절하게 판정하고 조언해줘.\n\n" +
      "【장착 스트링 사양】\n" +
      "- 교체 일자: " + stringInfo.date + " (현재 " + stats.daysElapsed + "일 경과)\n" +
      "- 라켓 모델: " + stringInfo.racket + "\n" +
      "- 스트링 모델: " + stringInfo.string_name + " (" + stringInfo.string_type + " 스트링)\n" +
      "- 세팅 텐션: " + stringInfo.tension + " lbs\n" +
      "- 특이사항: " + (stringInfo.notes || '없음') + "\n\n" +
      "【교체 이후 누적 테니스 일지 데이터】\n" +
      "- 누적 코트 사용 시간: " + stats.totalHours + "시간 (" + stats.totalMinutes + "분)\n" +
      "- 총 세션 수: " + stats.sessionCount + "회 (클럽게임 " + stats.gameCount + "회, 레슨 " + stats.lessonCount + "회, 연습 " + stats.practiceCount + "회)\n" +
      "- 평균 운동 강도: RPE " + stats.avgIntensity + "/10 (고강도 RPE 7+ 경기: " + stats.hardSessions + "회)\n" +
      "- 최근 테니스 일지 메모/후기 기록: " + JSON.stringify(stats.logsSince.slice(-15).map(function(l){ return l.date + ' [' + l.category + ']: ' + (l.notes || '특이사항 없음'); })) + "\n\n" +
      "【분석 지침 및 필수 구성】\n" +
      "아래 5가지 섹션으로 나누어 명확하게 작성할 것:\n" +
      "1. 🚦 [스트링 교체 최종 판정: 🔴즉시 교체 권장 / 🟡조만간 교체 권장 / 🟢현재 상태 양호 중 택 1]\n" +
      "   - 지금 당장 스트링을 새로 매야 하는지, 아니면 얼마간 더 사용 가능한지 결론을 맨 첫 줄에 명확히 선언할 것.\n" +
      "2. 🎾 [스트링 건강도 & 잔여 수명 점수 (0~100%)]\n" +
      "   - 여성 동호인 기준 권장 수명(40~50시간 / 2~3개월) 대비 현재 잔여 수명 점수와 사용 시간 비율 제시.\n" +
      "3. 📝 [테니스 일지 기반 타구감 & 피드백 분석]\n" +
      "   - 사용자가 일지에 남긴 메모(컨디션, 타구감, 공 날림, 빗맞음 등)를 직접 인용 및 연계하여 스트링 탄성 저하와의 상관관계를 분석할 것.\n" +
      "4. ⚠️ [관절(엘보/손목) 부상 예방 & 타구 충격 진단]\n" +
      "   - 텐션 로스 및 탄성 소실로 인해 팔과 손목에 전달되는 진동 충격 위험도 평가.\n" +
      "5. 💡 [맞춤 스트링 추천 & 최적 텐션 세팅 팁]\n" +
      "   - 다음 교체 시기(앞으로 N시간/N주 이내)와 여성 동호인에게 추천하는 부드러운 스트링 및 텐션(lbs) 가이드 제시.";

    var modelsToTry = ["gemini-3-flash-preview", "gemini-3.6-flash", "gemini-3.7-flash", "gemini-flash-latest"];
    var lastError = "";

    for (var i = 0; i < modelsToTry.length; i++) {
      var modelName = modelsToTry[i];
      var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey;
      
      var payload = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 8192 }
      };

      var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var resCode = response.getResponseCode();
      var resText = response.getContentText();

      if (resCode === 200) {
        var json = JSON.parse(resText);
        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
          var parts = json.candidates[0].content.parts;
          var fullText = "";
          for (var p = 0; p < parts.length; p++) {
            if (!parts[p].thought && parts[p].text) fullText += parts[p].text;
          }
          if (!fullText.trim()) fullText = parts[parts.length - 1].text || "";
          if (fullText.trim()) {
            return {
              success: true,
              analysis: fullText.trim(),
              isAi: true
            };
          }
        }
      } else {
        lastError = "Model " + modelName + " Error (" + resCode + "): " + resText;
      }
    }

    return {
      success: false,
      error: lastError,
      analysis: "[🎾 스트링 진단 가이드]\n누적 플레이 " + stats.totalHours + "시간 / 경과 " + stats.daysElapsed + "일차입니다. 통상적인 폴리 스트링 교체 주기(15~20시간)를 감안하여 점검하시길 권장합니다."
    };
  } catch(e) {
    Logger.log("analyzeTennisStringAI error: " + e.message);
    return {
      success: false,
      error: e.message,
      analysis: "스트링 분석 중 오류 발생: " + e.message
    };
  }
}

function generateGeminiTennisSummary(item) {
  try {
    var apiKey = item.geminiApiKey || DEFAULT_GEMINI_API_KEY;
    var cat = item.category || '테니스';
    var dur = item.duration_minutes || 0;
    var rpe = item.intensity || 5;
    var score = item.score || '';
    var result = item.result || '';
    var skill = item.focus_skill || '';
    var loc = item.location || '';
    var court = item.court_type || '';
    var userNotes = item.user_notes || '';
    var prompt = "당신은 열정적인 엘리트 테니스 프로 코치입니다.\n" +
      "아래 테니스 세션 데이터와 플레이어의 수기 메모를 바탕으로, [총 5~8줄 내외의 알차고 전문적인 원포인트 코칭 피드백]을 작성해주세요.\n\n" +
      "━━━━━━━━━━━━━━━━━━━\n" +
      "- 활동 구분: " + cat + " (" + dur + "분 / 강도 RPE " + rpe + ")\n" +
      (score ? "- 경기 결과: " + result + " (" + score + ")\n" : "") +
      (skill ? "- 집중 훈련 기술: " + skill + "\n" : "") +
      (loc ? "- 장소/코트: " + loc + " (" + court + ")\n" : "") +
      (userNotes ? "- 📝 플레이어 수기 메모: \"" + userNotes + "\"\n" : "") +
      "━━━━━━━━━━━━━━━━━━━\n\n" +
      "[작성 필수 규칙]:\n" +
      "1. 불필요하게 장황한 서론이나 맺음말(격언 등)은 생략하고, 바로 아래 3개 섹션으로 명확하게 작성할 것.\n" +
      "2. 플레이어가 적은 수기 메모 내용(타구감, 실수, 코트 느낌 등)을 구체적으로 짚어줄 것.\n" +
      "3. 전체 분량은 반드시 5줄~8줄 사이로 읽기 편하게 맞출 것.\n\n" +
      "[출력 포맷]:\n" +
      "🎾 [코칭 총평]\n" +
      "• (오늘 세션 결과, 강도, 수기 메모에 대한 칭찬 및 평가 2줄)\n\n" +
      "💡 [원포인트 기술 피드백]\n" +
      "• (수기 메모 및 코트/기술에 대한 원인 분석과 폼 교정 조언 2~3줄)\n\n" +
      "🔥 [다음 세션 실천 팁]\n" +
      "• (다음 경기/레슨에서 바로 적용할 구체적인 핵심 포인트 1~2줄)";

    var modelsToTry = ["gemini-3-flash-preview", "gemini-3.6-flash", "gemini-3.7-flash", "gemini-flash-latest"];
    for (var i = 0; i < modelsToTry.length; i++) {
      var modelName = modelsToTry[i];
      var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey;
      var payload = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.6, maxOutputTokens: 1200 }
      };
      var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
          var parts = json.candidates[0].content.parts;
          var txt = "";
          for (var p = 0; p < parts.length; p++) {
            if (!parts[p].thought && parts[p].text) txt += parts[p].text;
          }
          if (!txt.trim()) txt = parts[parts.length - 1].text || "";
          if (txt.trim()) {
            return { success: true, summary: txt.trim(), isAi: true };
          }
        }
      }
    }

    return {
      success: true,
      summary: "🎾 [총평] " + cat + " " + dur + "분 동안 강도 RPE " + rpe + "로 훌륭하게 세션을 마쳤습니다!\n💡 [원포인트 팁] 타구 시 안정적인 임팩트 타이밍과 풋워크 밸런스를 유지해보세요.\n🔥 [다음 세션] 다음 경기에서도 구체적인 타구감을 메모에 남겨보세요! 🔥"
    };
  } catch(e) {
    return {
      success: false,
      summary: "[🎾 코치 메모]\n오늘도 멋진 테니스 훈련을 완료하셨습니다! 수고하셨습니다. 🔥"
    };
  }
}

// -------------------------------------------------------------
// [7. 기타 유틸리티 & 텔레그램 알림]
// -------------------------------------------------------------
function sendTelegramMessage(text) {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) return;
  try {
    var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";
    var payload = {
      chat_id: TELEGRAM_CHAT_ID,
      text: text,
      parse_mode: "HTML"
    };
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(url, options);
  } catch(e) {
    Logger.log("Telegram 알림 발송 실패: " + e.message);
  }
}

function sendTennisLogTelegramNotification(item) {
  var title = "🎾 <b>[테니스 " + (item.category || '활동') + "]</b> " + (item.title || '테니스 일지');
  var lines = [title];
  lines.push("📅 <b>일자:</b> " + (item.date || ''));
  lines.push("⏱️ <b>시간/강도:</b> " + (item.duration_minutes || 0) + "분 (RPE " + (item.intensity || 5) + "/10)");
  if (item.location || item.court_type) {
    lines.push("🏟️ <b>구장:</b> " + (item.location || '') + " (" + (item.court_type || '인조잔디') + ")");
  }
  if (item.category === '클럽게임') {
    if (item.result) lines.push("🏆 <b>경기결과:</b> " + item.result + (item.score ? " (" + item.score + ")" : ""));
    if (item.players) lines.push("👥 <b>상대/파트너:</b> " + item.players);
  } else if (item.focus_skill) {
    lines.push("🎯 <b>집중기술:</b> " + item.focus_skill);
  }
  if (item.notes) {
    lines.push("\n📝 <b>수기 메모:</b>\n" + item.notes);
  }
  if (item.ai_summary) {
    lines.push("\n🤖 <b>AI 코칭 피드백:</b>\n" + item.ai_summary);
  }
  sendTelegramMessage(lines.join("\n"));
}

function sendTennisCommentTelegramNotification(logDate, logTitle, comment) {
  var lines = [
    "💬 <b>[테니스 일지 새 댓글]</b>",
    "📌 <b>대상 일지:</b> " + logDate + " " + logTitle,
    "👤 <b>작성자:</b> " + comment.author,
    "💬 <b>내용:</b>\n" + comment.text
  ];
  sendTelegramMessage(lines.join("\n"));
}

// -------------------------------------------------------------
// [7-1. 💚 LINE(라인) 메신저 댓글 실시간 알림]
function sendLineMessage(text) {
  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty("LINE_CHANNEL_ACCESS_TOKEN") || LINE_CHANNEL_ACCESS_TOKEN;
    var userId = props.getProperty("LINE_USER_ID") || LINE_USER_ID;
    if (!token || !userId) return;

    var url = "https://api.line.me/v2/bot/message/push";
    var payload = {
      to: String(userId),
      messages: [
        {
          type: "text",
          text: text
        }
      ]
    };
    var options = {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(url, options);
  } catch(e) {
    Logger.log("LINE 알림 발송 실패: " + e.message);
  }
}

function sendTennisCommentLineNotification(logDate, logTitle, comment) {
  try {
    var message = "💬 [테니스 일지] 새 피드백 & 댓글 도착!\n" +
      "━━━━━━━━━━━━━━━━━━━━\n" +
      "🎾 일지: " + (logTitle || "테니스 일지") + " (" + (logDate || "") + ")\n" +
      "👤 작성자: " + (comment.author || "테니스 동료") + "\n" +
      "📝 댓글 내용:\n" +
      comment.text + "\n" +
      "━━━━━━━━━━━━━━━━━━━━\n" +
      "⏰ " + (comment.createdAt || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"));

    sendLineMessage(message);
  } catch(e) {
    Logger.log("sendTennisCommentLineNotification error: " + e.message);
  }
}

function testTennisLineAlert() {
  var testComment = {
    author: "테니스 파트너",
    text: "오늘 복식 경기 랠리 대박이었습니다! 서브 폼 진짜 좋으시네요 🎾🔥",
    createdAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
  };
  sendTennisCommentLineNotification("2026-09-01", "문수 테니스장 클럽 게임", testComment);
  return { success: true };
}

// -------------------------------------------------------------
// [7-2. 🤖 텔레그램 메신저 스트링 교체 실시간 알림]
function sendTennisStringTelegramNotification(item) {
  var lines = [
    "🎾 <b>[새 스트링 교체 등록]</b>",
    "📅 <b>교체일자:</b> " + (item.date || ''),
    "🏸 <b>라켓:</b> " + (item.racket || ''),
    "🧵 <b>스트링:</b> " + (item.string_name || '') + " (" + (item.string_type || '폴리') + ")",
    "⚡ <b>텐션:</b> " + (item.tension || '') + " lbs",
    (item.cost ? "💰 <b>비용:</b> " + Number(item.cost).toLocaleString() + "원" : ""),
    (item.shop_name ? "🏪 <b>매장:</b> " + item.shop_name : ""),
    (item.notes ? "💬 <b>메모:</b> " + item.notes : "")
  ].filter(Boolean);
  sendTelegramMessage(lines.join("\n"));
}

function getEmptyOverview() {
  return {
    total_activities: 0,
    lesson_count: 0,
    game_count: 0,
    practice_count: 0,
    wins: 0,
    losses: 0,
    win_rate: 0,
    total_hours: 0
  };
}

function testTennisTelegramAlert() {
  var text = "🎾 <b>[테니스 일지 텔레그램 알림 테스트]</b>\n텔레그램 봇 알림이 정상적으로 연동되었습니다! 🏆🔥";
  sendTelegramMessage(text);
  Logger.log("테니스일지 텔레그램 테스트 알림 발송 완료!");
  return { success: true };
}

// -------------------------------------------------------------
// [8. 🎥 GEMINI AI 비디오 폼 정밀 분석 (테니스 & 러닝)]
// -------------------------------------------------------------
function analyzeTennisVideoFormAI(item) {
  try {
    var apiKey = item.geminiApiKey || DEFAULT_GEMINI_API_KEY;
    var videoBase64 = item.videoBase64;
    var mimeType = item.mimeType || "video/mp4";
    var motionType = item.motionType || "테니스 - 포핸드 스트로크";
    var focusNotes = item.focusNotes || "";
    var modelPref = item.modelPreference || "pro";

    if (!videoBase64) {
      return {
        success: false,
        error: "분석할 동영상 파일 데이터가 전달되지 않았습니다."
      };
    }

    var prompt = "당신은 세계적인 프로 테니스 수석 코치이자 스포츠 바이오메카닉스(동작 생체역학) 및 러닝 폼 전문 분석가입니다.\n\n" +
      "사용자가 업로드한 10~30초 동영상(촬영된 " + motionType + " 동작)을 프레임 단위로 정밀 분석하여, 아래 [필수 구성 양식]에 맞춰 최고의 전문 코칭 리포트를 작성해주세요.\n\n" +
      (focusNotes ? "【사용자의 중점 질문 및 점검 요청】: \"" + focusNotes + "\"\n\n" : "") +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
      "【필수 분석 가이드 및 출력 포맷】\n" +
      "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
      "🎯 [1. 폼 종합 평가 & 완성도 점수]\n" +
      "• 폼 완성도: [XX점 / 100점] (등급: S / A / B / C 중 택1)\n" +
      "• 핵심 진단 한줄평: (해당 동작의 가장 큰 강점과 즉각 개선이 필요한 핵심 1줄 요약)\n\n" +
      "⏱️ [2. 프레임별 타임스탬프 동작 분해 분석]\n" +
      "(영상에 나타난 실제 시간 [MM:SS] 또는 [00:0X] 타임스탬프를 명시하여 단계별로 상세히 분석할 것)\n" +
      "• [00:00 ~ 00:0X] 준비 자세 & 유닛 턴 (Ready Position & Unit Turn / Stance)\n" +
      "  - 시선, 스탠스 넓이, 상체 회전 각도 및 밸런스 점검\n" +
      "• [00:0X ~ 00:0X] 테이크백 & 라켓 드롭 (Backswing / Racket Drop / Knee Bend)\n" +
      "  - 라켓 헤드 높이, 팔꿈치 여유 공간, 무릎 굽힘 및 하체 코일링(에너지 축적)\n" +
      "• [00:0X ~ 00:0X] 포워드 스윙 & 임팩트 타점 (Forward Swing & Contact Point / Foot Strike)\n" +
      "  - 임팩트 시 몸 앞 타점 위치, 라켓면 각도, 체중 이동, 러닝 시 착지 패턴(미드풋/힐) 및 무릎 각도\n" +
      "• [00:0X ~ 00:0X] 팔로우스루 & 리커버리 (Follow-Through & Kinetic Chain Finish)\n" +
      "  - 릴리즈 궤적, 숄더 오버 숄더 회전, 착지 후 다음 동작 복귀 안정성\n\n" +
      "💡 [3. 핵심 교정 처방 3가지 (Action Items)]\n" +
      "1) [가장 시급한 폼 교정]: (원인 분석 + 구체적인 몸 동작 수정 가이드)\n" +
      "2) [파워 & 정확도 향상 포인트]: (하체 체중 이동, 라켓 가속 타이밍 등)\n" +
      "3) [추천 실전 연습 드릴]: (코트 또는 집에서 혼자 할 수 있는 1가지 맞춤형 연습법)\n\n" +
      "⚠️ [4. 생체역학 & 관절 부상 위험도 점검]\n" +
      "• 엘보(팔꿈치) / 손목 부하: (손목 꺾임, 충격 흡수 상태 진단)\n" +
      "• 어깨 회전근개 / 허리 부하: (과도한 젖힘이나 무리한 스윙 여부)\n" +
      "• 무릎 / 발목 관절 부하: (착지 충격 및 무릎 안쪽 쏠림 여부)\n\n" +
      "친절하면서도 권위 있는 프로 코치의 어조로 읽기 쉽게 작성해주세요.";

    var modelsToTry = [];
    if (modelPref === "pro") {
      modelsToTry = ["gemini-2.5-pro", "gemini-1.5-pro", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-3-flash-preview", "gemini-flash-latest"];
    } else {
      modelsToTry = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-3-flash-preview", "gemini-flash-latest", "gemini-2.5-pro", "gemini-1.5-pro"];
    }

    var lastError = "";

    for (var i = 0; i < modelsToTry.length; i++) {
      var modelName = modelsToTry[i];
      var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey;

      var payload = {
        contents: [{
          parts: [
            {
              inline_data: {
                mime_type: mimeType,
                data: videoBase64
              }
            },
            {
              text: prompt
            }
          ]
        }],
        generationConfig: {
          temperature: 0.4,
          maxOutputTokens: 8192
        }
      };

      var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var resCode = response.getResponseCode();
      var resText = response.getContentText();

      if (resCode === 200) {
        var json = JSON.parse(resText);
        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
          var parts = json.candidates[0].content.parts;
          var fullText = "";
          for (var p = 0; p < parts.length; p++) {
            if (!parts[p].thought && parts[p].text) fullText += parts[p].text;
          }
          if (!fullText.trim()) fullText = parts[parts.length - 1].text || "";
          if (fullText.trim()) {
            return {
              success: true,
              modelUsed: modelName,
              analysis: fullText.trim()
            };
          }
        }
      } else {
        lastError = "Model " + modelName + " (" + resCode + "): " + resText;
        Logger.log("Video analysis error on model " + modelName + ": " + resText);
      }
    }

    return {
      success: false,
      error: lastError || "AI 비디오 분석 모델 응답을 받지 못했습니다."
    };
  } catch(e) {
    Logger.log("analyzeTennisVideoFormAI error: " + e.message);
    return {
      success: false,
      error: "동영상 분석 중 오류가 발생했습니다: " + e.message
    };
  }
}

function initDatabase() {
  var s1 = getSheet();
  var s2 = getStringSheet();
  return { success: true, sheet1: s1.getName(), sheet2: s2.getName() };
}
