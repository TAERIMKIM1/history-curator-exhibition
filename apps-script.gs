// ============================================================
// 역사 큐레이터 기획전 — Google Apps Script 백엔드 v2
// Google Apps Script 에디터에 붙여넣고 웹앱으로 재배포하세요
// ============================================================

var TEACHER_PASSWORD = "history2025"; // 교사 비밀번호 (원하는 대로 변경)

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // ── 스티커 투표 ──
    if (data.action === "vote") {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString() === data.id.toString()) {
          var currentVotes = Number(rows[i][12]) || 0;
          sheet.getRange(i + 1, 13).setValue(currentVotes + 1);
          return ok({ result: "voted", count: currentVotes + 1 });
        }
      }
      return ok({ result: "not_found" });
    }

    // ── 패널 삭제 (교사 전용) ──
    if (data.action === "delete") {
      if (data.password !== TEACHER_PASSWORD) {
        return ok({ result: "unauthorized" });
      }
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString() === data.id.toString()) {
          sheet.deleteRow(i + 1);
          return ok({ result: "deleted" });
        }
      }
      return ok({ result: "not_found" });
    }

    // ── 패널 수정 (학생 본인) ──
    if (data.action === "update") {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString() === data.id.toString()) {
          sheet.getRange(i + 1, 2).setValue(data.hall);
          sheet.getRange(i + 1, 6).setValue(data.relicName);
          sheet.getRange(i + 1, 7).setValue(data.relicSpec);
          sheet.getRange(i + 1, 8).setValue(data.historyContext);
          sheet.getRange(i + 1, 9).setValue(data.curatorVoice);
          sheet.getRange(i + 1, 10).setValue(data.source);
          sheet.getRange(i + 1, 11).setValue(data.imageUrl1);
          sheet.getRange(i + 1, 12).setValue(data.imageUrl2);
          return ok({ result: "updated", id: data.id });
        }
      }
      return ok({ result: "not_found" });
    }

    // ── 신규 패널 등록 ──
    var id = new Date().getTime();
    sheet.appendRow([
      id,
      data.hall,
      data.gradeClass,
      data.stdNum,
      data.stdName,
      data.relicName,
      data.relicSpec,
      data.historyContext,
      data.curatorVoice,
      data.source,
      data.imageUrl1,
      data.imageUrl2,
      0
    ]);
    return ok({ result: "success", id: id });

  } catch(err) {
    return ok({ result: "error", message: err.toString() });
  }
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    result.push({
      id:              rows[i][0],
      hall:            rows[i][1],
      gradeClass:      rows[i][2],
      stdNum:          rows[i][3],
      stdName:         rows[i][4],
      relicName:       rows[i][5],
      relicSpec:       rows[i][6],
      historyContext:  rows[i][7],
      curatorVoice:    rows[i][8],
      source:          rows[i][9],
      imageUrl1:       rows[i][10],
      imageUrl2:       rows[i][11],
      votes:           Number(rows[i][12]) || 0
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
