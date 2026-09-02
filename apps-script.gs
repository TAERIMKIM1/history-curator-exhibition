// ============================================================
// 역사 큐레이터 기획전 — Google Apps Script 백엔드
// 복사하여 https://script.google.com 에 붙여넣고 웹앱으로 배포하세요
// ============================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // 스티커 투표 처리
    if (data.action === "vote") {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString() === data.id.toString()) {
          var currentVotes = Number(rows[i][12]) || 0;
          sheet.getRange(i + 1, 13).setValue(currentVotes + 1);
          return ContentService
            .createTextOutput(JSON.stringify({result: "voted", count: currentVotes + 1}))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    // 신규 패널 등록
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

    return ContentService
      .createTextOutput(JSON.stringify({result: "success", id: id}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({result: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < rows.length; i++) {
    result.push({
      id: rows[i][0],
      hall: rows[i][1],
      gradeClass: rows[i][2],
      stdNum: rows[i][3],
      stdName: rows[i][4],
      relicName: rows[i][5],
      relicSpec: rows[i][6],
      historyContext: rows[i][7],
      curatorVoice: rows[i][8],
      source: rows[i][9],
      imageUrl1: rows[i][10],
      imageUrl2: rows[i][11],
      votes: Number(rows[i][12]) || 0
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
