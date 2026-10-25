// ============================================================
// 역사 큐레이터 기획전 — Google Apps Script 백엔드 v2
// Google Apps Script 에디터에 붙여넣고 웹앱으로 재배포하세요
// ============================================================

// ============================================================
// ★ 최초 1회 실행 함수: fillAllPanels
// 에디터에서 fillAllPanels 선택 후 ▶ 실행
// 유물명(F열)·기본 제원(G열)·출처(J열)를 행 ID 기준으로 채워줌
// ============================================================
function fillAllPanels() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('데이터가 없습니다.');
    return;
  }

  // 한 번에 읽기 (A~J열)
  var rows = sheet.getRange(1, 1, lastRow, 10).getValues();

  var DATA = {
    "1788614010173": ["청자 상감운학문 매병","12세기(고려 중기), 청자에 상감기법, 높이 43.9cm, 국보 제68호, 국립중앙박물관 소장","국립중앙박물관 e뮤지엄 (www.emuseum.go.kr), 우리역사넷 (contents.history.go.kr), 국가유산포털 (www.heritage.go.kr)"],
    "1788614013098": ["팔만대장경","1236~1251년(고려 고종 23~38년), 목판 인쇄본, 81,258판, 국보 제32호·유네스코 세계기록유산(1997), 해인사 장경판전 소장","해인사 공식 사이트, 우리역사넷 (contents.history.go.kr), 유네스코 세계기록유산 공식 자료"],
    "1788614015071": ["직지심체요절","1377년(고려 우왕 3년), 금속활자 인쇄본, 세계 최초 금속활자 인쇄본, 유네스코 세계기록유산(2001), 프랑스 국립도서관 소장","국립중앙박물관 e뮤지엄, 청주고인쇄박물관 (www.jikjiworld.cheongju.go.kr), 우리역사넷"],
    "1788614018426": ["고려 수월관음도","14세기(고려 후기), 비단에 금니·채색, 세로 419.5cm, 국보 제315호, 아모레퍼시픽미술관 소장","국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷"],
    "1788614021134": ["부석사 무량수전","9세기 창건·13세기(고려 중기) 중창, 목조 건물, 배흘림기둥 공포 양식, 국보 제18호, 경북 영주 부석사 소재","국가유산포털 (www.heritage.go.kr), 우리역사넷 (contents.history.go.kr)"],
    "1788614023244": ["고려 나전칠기 경함","13세기(고려 중기), 나무에 나전·옻칠, 국보 제102호, 국립중앙박물관 소장","국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷"],
    "1788614026063": ["경천사지 십층석탑","1348년(고려 충목왕 4년), 대리석 조각, 높이 13.5m, 국보 제86호, 국립중앙박물관 소장","국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷"],
    "1788614028546": ["청동 은입사 포류수금문 정병","12세기(고려 중기), 청동에 은입사기법, 높이 37.5cm, 국보 제92호, 국립중앙박물관 소장","국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷"],
    "1788614031168": ["삼국유사","1281년경(고려 충렬왕, 일연 저술), 목판 인쇄본, 국보 제306호, 연세대학교 박물관 소장(인각사본)","우리역사넷, 한국사데이터베이스 (db.history.go.kr), 국가유산포털"],
    "1788614033533": ["금동관음보살좌상","14세기(고려 후기), 금동 주조·도금, 국보 제124호, 국립중앙박물관 소장","국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷"],
    "1788614036176": ["훈민정음 해례본","1443년 창제·1446년(세종 28년) 반포, 목판 인쇄본(해례본), 국보 제70호·유네스코 세계기록유산(1997), 간송미술관 소장","국립한글박물관 (www.hangeul.go.kr), 우리역사넷, 유네스코 세계기록유산 공식 자료"],
    "1788614038218": ["측우기","1441년(세종 23년), 청동 원통형, 높이 30.0cm·지름 15.0cm, 세계 최초 강우량 측정기, 국보 제561호, 국립기상박물관 소장","우리역사넷, 국가유산포털, 기상청 국립기상박물관"],
    "1788614041116": ["자격루","1434년(세종 16년), 청동·목재, 물시계(자동 시보 장치), 국보 제229호, 국립고궁박물관 소장","국립고궁박물관, 우리역사넷, 국가유산포털"],
    "1788614043160": ["앙부일구","1434년(세종 16년) 제작·현존품은 17세기, 청동 주조, 국보 제845호, 국립고궁박물관 소장","국립고궁박물관 (www.gogung.go.kr), 우리역사넷, 국가유산포털"],
    "1788614045445": ["혼천의","15세기(세종~성종), 청동 주조, 천체 관측 기구, 국보 제230호, 고려대학교 박물관 소장","국립고궁박물관, 우리역사넷, 국가유산포털"],
    "1788614048145": ["조선왕조실록","1413~1865년(태조~철종), 지본 필사·목판본, 총 1,893권, 국보 제151호·유네스코 세계기록유산(1997), 서울대학교 규장각 소장","한국사데이터베이스 (db.history.go.kr), 우리역사넷, 유네스코 세계기록유산 자료"],
    "1788614050276": ["경복궁 근정전","1395년(태조 4년) 창건·1867년(고종 4년) 중건, 목조 건물, 국보 제223호, 서울 경복궁 소재","국립고궁박물관 (www.gogung.go.kr), 국가유산포털, 우리역사넷"],
    "1788614053224": ["칠정산 내편","1444년(세종 26년), 지본 필사·목판본, 조선 독자 역법서, 보물 제1248호, 서울대학교 규장각 소장","우리역사넷, 한국사데이터베이스, 규장각한국학연구원"],
    "1788614056032": ["용비어천가","1445년 지어 1447년(세종 29년) 간행, 목판 인쇄본, 125장, 보물 제445호, 서울대학교 규장각 소장","규장각한국학연구원 원문검색, 한국사데이터베이스 (db.history.go.kr), 우리역사넷"],
    "1788614058053": ["분청사기 박지 철화 어문 항아리","15~16세기(조선 전기), 분청사기 박지·철화 기법, 국보 제259호, 국립중앙박물관 소장","국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷"]
  };

  var dataRows = lastRow - 1;
  var colF = []; // F열: relicName
  var colG = []; // G열: relicSpec
  var colJ = []; // J열: source
  var updated = 0;

  for (var i = 1; i <= dataRows; i++) {
    var rowId = String(rows[i][0]).trim();
    var info = DATA[rowId];
    if (info) {
      colF.push([info[0]]);
      colG.push([info[1]]);
      var existing = String(rows[i][9]).trim();
      colJ.push([existing ? existing : info[2]]);
      updated++;
    } else {
      colF.push([rows[i][5]]);
      colG.push([rows[i][6]]);
      colJ.push([rows[i][9]]);
    }
  }

  // 셀별 setValue 대신 열 단위 일괄 쓰기 (3번만 호출)
  sheet.getRange(2, 6, dataRows, 1).setValues(colF);
  sheet.getRange(2, 7, dataRows, 1).setValues(colG);
  sheet.getRange(2, 10, dataRows, 1).setValues(colJ);

  console.log('✅ 완료: ' + updated + '개 패널에 문화유산 명칭·기본 정보·출처를 채웠습니다.');
}

var TEACHER_PASSWORD_HASH = "4c16016741f6c1ec8ab0afb40d66fb1a86f5a7906a02b8a18391f7c4dcdf0f46";
var ANTHROPIC_API_KEY_VALUE = ""; // ← Anthropic API 키를 여기에 직접 입력 (sk-ant-api03-...)

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
      if (data.password !== TEACHER_PASSWORD_HASH) {
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

    // ── 교사 채점 ──
    if (data.action === "score") {
      if (data.password !== TEACHER_PASSWORD_HASH) return ok({ result: "unauthorized" });
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString() === data.id.toString()) {
          sheet.getRange(i + 1, 16).setValue(Number(data.score1) || 0);
          sheet.getRange(i + 1, 17).setValue(Number(data.score2) || 0);
          sheet.getRange(i + 1, 18).setValue(Number(data.score3) || 0);
          sheet.getRange(i + 1, 19).setValue(Number(data.score4) || 0);
          sheet.getRange(i + 1, 20).setValue(data.feedback || '');
          var total = (Number(data.score1)||0)+(Number(data.score2)||0)+(Number(data.score3)||0)+(Number(data.score4)||0);
          return ok({ result: "scored", total: total });
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
          sheet.getRange(i + 1, 14).setValue(data.imgCaption1 || '');
          sheet.getRange(i + 1, 15).setValue(data.imgCaption2 || '');
          return ok({ result: "updated", id: data.id });
        }
      }
      return ok({ result: "not_found" });
    }

    // ── API 키 테스트 ──
    if (data.action === "testKey") {
      if (data.password !== TEACHER_PASSWORD_HASH) return ok({ result: "unauthorized" });
      var apiKey = data.apiKey || ANTHROPIC_API_KEY_VALUE || PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
      if (!apiKey) return ok({ result: "no_api_key" });
      var testPayload = { model: "claude-haiku-4-5", max_tokens: 30,
        messages: [{ role: "user", content: '{"r1":3,"r2":4,"r3":3,"r4":4,"feedback":"테스트"}만 출력' }] };
      var testOptions = { method: "post", contentType: "application/json",
        headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        payload: JSON.stringify(testPayload), muteHttpExceptions: true, deadline: 20 };
      try {
        var testResp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", testOptions);
        var testJson = JSON.parse(testResp.getContentText());
        if (testJson.error) return ok({ result: "api_error", message: testJson.error.message });
        return ok({ result: "ok", response: (testJson.content && testJson.content[0]) ? testJson.content[0].text : '' });
      } catch(testErr) { return ok({ result: "error", message: testErr.toString() }); }
    }

    // ── AI 자동채점 ──
    if (data.action === "autoScore") {
      if (data.password !== TEACHER_PASSWORD_HASH) return ok({ result: "unauthorized" });
      var apiKey = data.apiKey || ANTHROPIC_API_KEY_VALUE || PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
      if (!apiKey) return ok({ result: "no_api_key", message: "관리자 배너의 API 키 입력란에 Anthropic API 키를 입력하세요." });
      var allRows = sheet.getDataRange().getValues();
      var pRow = null;
      for (var i = 1; i < allRows.length; i++) {
        if (allRows[i][0].toString() === data.id.toString()) { pRow = allRows[i]; break; }
      }
      if (!pRow) return ok({ result: "not_found" });
      var hall = pRow[1]||'', gradeClass2 = pRow[2]||'', stdName2 = pRow[4]||'';
      var relicName2 = pRow[5]||'', relicSpec2 = pRow[6]||'';
      var histCtx = pRow[7]||'', curVoice = pRow[8]||'', src2 = pRow[9]||'';
      var img1 = pRow[10]||'', img2 = pRow[11]||'';
      var aiPrompt = '당신은 중학교 역사 교사입니다. 학생이 제출한 역사 큐레이터 기획전 패널을 채점해 주세요.\n\n'
        + '[학생 패널]\n전시관: ' + hall + '\n학급: ' + gradeClass2 + ' / 이름: ' + stdName2
        + '\n문화유산명: ' + relicName2 + '\n기본 제원: ' + relicSpec2
        + '\n큐레이터의 시선(시대적 배경): ' + histCtx
        + '\n큐레이터의 해석(역사적 의의): ' + curVoice
        + '\n사료 출처: ' + src2
        + '\n사진 첨부: ' + (img1 ? '사진1 있음' : '없음') + (img2 ? ', 사진2 있음' : '')
        + '\n\n[채점 기준] 각 항목 1~5점\n'
        + '① 문화유산 선정 및 기본 정보: 5=성취기준 부합·자료 풍부, 4=부합·자료 다소 부족, 3=연계성 부족, 2=무관·매우 빈약, 1=전혀 없음\n'
        + '② 시대적 배경: 5=구체적 사실 기반·정확, 4=배경 설명·세부 부족, 3=단편적 나열, 2=오류 다수·빈약, 1=서술 없음\n'
        + '③ 역사적 의의: 5=근거 있고 논리정연, 4=의의 설명·근거 부족, 3=단순 언급, 2=오류·근거 없음, 1=서술 없음\n'
        + '④ 형식 및 완성도: 5=출처·사진 완비·완성도 높음, 4=대체로 갖춤·일부 미흡, 3=일부 누락, 2=거의 미갖춤, 1=결과물 없음\n\n'
        + '다음 JSON 형식으로만 답변하세요 (다른 텍스트 없이):\n'
        + '{"score1":숫자,"score2":숫자,"score3":숫자,"score4":숫자,"reason1":"근거","reason2":"근거","reason3":"근거","reason4":"근거"}';
      var aiPayload = { model: "claude-haiku-4-5", max_tokens: 600,
        messages: [{ role: "user", content: aiPrompt }] };
      var aiOptions = { method: "post", contentType: "application/json",
        headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        payload: JSON.stringify(aiPayload), muteHttpExceptions: true, deadline: 25 };
      try {
        var aiResp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", aiOptions);
        var aiJson = JSON.parse(aiResp.getContentText());
        if (aiJson.error) return ok({ result: "api_error", message: aiJson.error.message });
        var aiText = (aiJson.content && aiJson.content[0]) ? aiJson.content[0].text : '';
        var m = aiText.match(/\{[\s\S]*\}/);
        if (!m) return ok({ result: "parse_error", raw: aiText });
        var sc = JSON.parse(m[0]);
        return ok({ result: "auto_scored",
          score1: parseInt(sc.score1)||1, score2: parseInt(sc.score2)||1,
          score3: parseInt(sc.score3)||1, score4: parseInt(sc.score4)||1,
          reason1: sc.reason1||'', reason2: sc.reason2||'',
          reason3: sc.reason3||'', reason4: sc.reason4||'' });
      } catch(aiErr) { return ok({ result: "error", message: aiErr.toString() }); }
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
      0,
      data.imgCaption1 || '',
      data.imgCaption2 || ''
    ]);
    // 반(gradeClass) 셀을 텍스트 형식으로 강제 지정 — Sheets 날짜 자동변환 방지
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setNumberFormat('@');
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
    // Sheets가 "3-4" 같은 반 표기를 날짜로 자동변환한 경우 복원
    var gc = rows[i][2];
    if (gc instanceof Date) {
      gc = (gc.getMonth() + 1) + '-' + gc.getDate();
    } else {
      gc = String(gc).trim();
    }
    result.push({
      id:              rows[i][0],
      hall:            rows[i][1],
      gradeClass:      gc,
      stdNum:          rows[i][3],
      stdName:         rows[i][4],
      relicName:       rows[i][5],
      relicSpec:       rows[i][6],
      historyContext:  rows[i][7],
      curatorVoice:    rows[i][8],
      source:          rows[i][9],
      imageUrl1:       rows[i][10],
      imageUrl2:       rows[i][11],
      votes:           Number(rows[i][12]) || 0,
      imgCaption1:     rows[i][13] || '',
      imgCaption2:     rows[i][14] || '',
      score1:          Number(rows[i][15]) || 0,
      score2:          Number(rows[i][16]) || 0,
      score3:          Number(rows[i][17]) || 0,
      score4:          Number(rows[i][18]) || 0,
      feedback:        rows[i][19] || '',
      totalScore:      (Number(rows[i][15])||0)+(Number(rows[i][16])||0)+(Number(rows[i][17])||0)+(Number(rows[i][18])||0)
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

// ============================================================
// ★ 교사용 일괄 보완 함수
// Apps Script 에디터에서 fillMissingInfo 선택 후 ▶ 실행
// 유물 기본 정보(G열)·출처(J열)가 비어 있는 행만 자동 채워줌
// ============================================================
function fillMissingInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();

  // 유물명 → {spec, src, img1, img2} 매핑 테이블
  var INFO = {
    // ── 고려관 ──
    "청자 상감운학문 매병":
      { spec: "12세기(고려 중기), 청자에 상감기법, 높이 43.9cm, 국보 제68호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄 (www.emuseum.go.kr), 우리역사넷 (contents.history.go.kr), 국가유산포털 (www.heritage.go.kr)",
        img1: "https://upload.wikimedia.org/wikipedia/commons/4/42/Korea-Goryeo_period-Celadon_maebyeong_with_inlaid_cloud_and_crane_design-02.jpg",
        img2: "https://upload.wikimedia.org/wikipedia/commons/2/2e/Korea-Goryeo_period-Celadon_maebyeong_with_inlaid_cloud_and_crane_design-01.jpg" },
    "팔만대장경":
      { spec: "1236~1251년(고려 고종 23~38년), 목판 인쇄본, 81,258판, 국보 제32호·유네스코 세계기록유산(1997), 해인사 장경판전 소장",
        src:  "해인사 공식 사이트, 우리역사넷 (contents.history.go.kr), 유네스코 세계기록유산 공식 자료",
        img1: "https://upload.wikimedia.org/wikipedia/commons/3/3d/Tripitaka_Koreana-Haeinsa.jpg",
        img2: "https://upload.wikimedia.org/wikipedia/commons/5/5e/Korea-Haeinsa-Tripitaka_Koreana-01.jpg" },
    "고려대장경":
      { spec: "1236~1251년(고려 고종 23~38년), 목판 인쇄본, 81,258판, 국보 제32호·유네스코 세계기록유산(1997), 해인사 장경판전 소장",
        src:  "해인사 공식 사이트, 우리역사넷 (contents.history.go.kr), 유네스코 세계기록유산 공식 자료",
        img1: "https://upload.wikimedia.org/wikipedia/commons/3/3d/Tripitaka_Koreana-Haeinsa.jpg",
        img2: "https://upload.wikimedia.org/wikipedia/commons/5/5e/Korea-Haeinsa-Tripitaka_Koreana-01.jpg" },
    "직지심체요절":
      { spec: "1377년(고려 우왕 3년), 금속활자 인쇄본, 세계 최초 금속활자 인쇄본, 유네스코 세계기록유산(2001), 프랑스 국립도서관 소장",
        src:  "국립중앙박물관 e뮤지엄, 청주고인쇄박물관 (www.jikjiworld.cheongju.go.kr), 우리역사넷",
        img1: "https://upload.wikimedia.org/wikipedia/commons/8/8b/Jikji.jpg",
        img2: "" },
    "삼국유사":
      { spec: "1281년경(고려 충렬왕, 일연 저술), 목판 인쇄본, 국보 제306호, 연세대학교 박물관 소장(인각사본)",
        src:  "우리역사넷, 한국사데이터베이스 (db.history.go.kr), 국가유산포털",
        img1: "https://upload.wikimedia.org/wikipedia/commons/8/83/Samguk_Yusa.jpg",
        img2: "" },
    "청자 투각 칠보문 향로":
      { spec: "12세기(고려 중기), 청자에 투각기법, 높이 15.3cm, 국보 제95호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털",
        img1: "https://upload.wikimedia.org/wikipedia/commons/3/31/Korea-Goryeo-Celadon_incense_burner-01.jpg",
        img2: "" },
    "청동 은입사 포류수금문 정병":
      { spec: "12세기(고려 중기), 청동에 은입사기법, 높이 37.5cm, 국보 제92호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷",
        img1: "https://upload.wikimedia.org/wikipedia/commons/f/f2/Korea-Goryeo-Bronze_kundika-01.jpg",
        img2: "" },
    "고려 수월관음도":
      { spec: "14세기(고려 후기), 비단에 금니·채색, 세로 419.5cm, 국보 제315호, 아모레퍼시픽미술관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷",
        img1: "https://upload.wikimedia.org/wikipedia/commons/5/5b/Goryeo-Avalokitesvara.jpg",
        img2: "" },
    "청자 참외형 병":
      { spec: "12세기(고려 중기), 순청자 음각기법, 높이 22.6cm, 국보 제94호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털",
        img1: "https://upload.wikimedia.org/wikipedia/commons/1/10/Korea-Goryeo-Celadon_melon_shaped_bottle-01.jpg",
        img2: "" },
    "고려 나전칠기 경함":
      { spec: "13세기(고려 중기), 나무에 나전·옻칠, 국보 제102호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷",
        img1: "https://upload.wikimedia.org/wikipedia/commons/9/91/Korea-Goryeo-Lacquerware_sutra_case-01.jpg",
        img2: "" },
    "고려 금동 대탑":
      { spec: "10~11세기(고려 전기), 금동 주조·도금, 높이 155cm, 국보 제213호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털",
        img1: "", img2: "" },
    // ── 조선관 ──
    "훈민정음 해례본":
      { spec: "1443년 창제·1446년(세종 28년) 반포, 목판 인쇄본(해례본), 국보 제70호·유네스코 세계기록유산(1997), 간송미술관 소장",
        src:  "국립한글박물관 (www.hangeul.go.kr), 우리역사넷, 유네스코 세계기록유산 공식 자료",
        img1: "https://upload.wikimedia.org/wikipedia/commons/d/d9/Hunminjeongeum.jpg",
        img2: "https://upload.wikimedia.org/wikipedia/commons/1/1e/Hunmin_jeongeum.jpg" },
    "훈민정음":
      { spec: "1443년 창제·1446년(세종 28년) 반포, 목판 인쇄본(해례본), 국보 제70호·유네스코 세계기록유산(1997), 간송미술관 소장",
        src:  "국립한글박물관 (www.hangeul.go.kr), 우리역사넷, 유네스코 세계기록유산 공식 자료",
        img1: "https://upload.wikimedia.org/wikipedia/commons/d/d9/Hunminjeongeum.jpg",
        img2: "https://upload.wikimedia.org/wikipedia/commons/1/1e/Hunmin_jeongeum.jpg" },
    "측우기":
      { spec: "1441년(세종 23년), 청동 원통형, 높이 30.0cm·지름 15.0cm, 세계 최초 강우량 측정기, 국보 제561호, 국립기상박물관 소장",
        src:  "우리역사넷, 국가유산포털, 기상청 국립기상박물관",
        img1: "https://upload.wikimedia.org/wikipedia/commons/4/40/Cheugugi.jpg",
        img2: "" },
    "앙부일구":
      { spec: "1434년(세종 16년) 제작·현존품은 17세기, 청동 주조, 국보 제845호, 국립고궁박물관 소장",
        src:  "국립고궁박물관 (www.gogung.go.kr), 우리역사넷, 국가유산포털",
        img1: "https://upload.wikimedia.org/wikipedia/commons/9/96/Angbuilgu.jpg",
        img2: "" },
    "자격루":
      { spec: "1434년(세종 16년), 청동·목재, 물시계(자동 시보 장치), 국보 제229호, 국립고궁박물관 소장",
        src:  "국립고궁박물관, 우리역사넷, 국가유산포털",
        img1: "https://upload.wikimedia.org/wikipedia/commons/3/30/Jagyeongnu.jpg",
        img2: "" },
    "조선왕조실록":
      { spec: "1413~1865년(태조~철종), 지본 필사·목판본, 총 1,893권, 국보 제151호·유네스코 세계기록유산(1997), 서울대학교 규장각 소장",
        src:  "한국사데이터베이스 (db.history.go.kr), 우리역사넷, 유네스코 세계기록유산 자료",
        img1: "https://upload.wikimedia.org/wikipedia/commons/c/c0/Joseon_Wangjo_Sillok.jpg",
        img2: "" },
    "경국대전":
      { spec: "1485년(성종 16년) 완성, 목판 인쇄본, 조선의 기본 법전, 보물 제1521호, 규장각한국학연구원 소장",
        src:  "한국사데이터베이스, 우리역사넷, 규장각한국학연구원 원문검색",
        img1: "https://upload.wikimedia.org/wikipedia/commons/4/45/Gyeonggukdaejeon.jpg",
        img2: "" },
    "동의보감":
      { spec: "1613년(광해군 5년), 목판 인쇄본, 총 25권, 국보 제319호·유네스코 세계기록유산(2009), 국립중앙도서관 소장",
        src:  "국립중앙도서관, 우리역사넷, 유네스코 세계기록유산 자료",
        img1: "https://upload.wikimedia.org/wikipedia/commons/5/5f/Dongui_Bogam.jpg",
        img2: "" },
    "혼일강리역대국도지도":
      { spec: "1402년(태종 2년), 채색 필사본, 세로 158cm·가로 168cm, 현존 동양 최고 세계지도 중 하나, 일본 류코쿠대학 소장",
        src:  "우리역사넷, 한국사데이터베이스, 국립중앙박물관 e뮤지엄",
        img1: "https://upload.wikimedia.org/wikipedia/commons/e/e5/Honil_Gangni_Yeokdae_Gukdo_Ji_Do.jpg",
        img2: "" },
    "신숙주 초상":
      { spec: "15세기(조선 전기), 비단에 채색, 세로 167cm·가로 109.5cm, 국보 제613호, 국립중앙박물관 기탁",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷",
        img1: "", img2: "" },
    "칠정산 내편":
      { spec: "1444년(세종 26년), 지본 필사·목판본, 조선 독자 역법서, 보물 제1248호, 서울대학교 규장각 소장",
        src:  "우리역사넷, 한국사데이터베이스, 규장각한국학연구원",
        img1: "", img2: "" },
    "혼천의":
      { spec: "15세기(세종~성종), 청동 주조, 천체 관측 기구, 국보 제230호, 고려대학교 박물관 소장",
        src:  "국립고궁박물관, 우리역사넷, 국가유산포털",
        img1: "https://upload.wikimedia.org/wikipedia/commons/7/70/Honcheoneui.jpg",
        img2: "" }
  };

  var updated = 0;
  for (var i = 1; i < rows.length; i++) {
    var relicName = String(rows[i][5]).trim();
    var relicSpec = String(rows[i][6]).trim();
    var source    = String(rows[i][9]).trim();
    var img1      = String(rows[i][10]).trim();
    var img2      = String(rows[i][11]).trim();

    if (!relicName) continue;

    var info = INFO[relicName];
    if (!info) continue;

    var needsUpdate = false;
    if (!relicSpec || relicSpec.length < 5) {
      sheet.getRange(i + 1, 7).setValue(info.spec);
      needsUpdate = true;
    }
    if (!source || source.length < 5) {
      sheet.getRange(i + 1, 10).setValue(info.src);
      needsUpdate = true;
    }
    if ((!img1 || img1.length < 5) && info.img1) {
      sheet.getRange(i + 1, 11).setValue(info.img1);
      needsUpdate = true;
    }
    if ((!img2 || img2.length < 5) && info.img2) {
      sheet.getRange(i + 1, 12).setValue(info.img2);
      needsUpdate = true;
    }
    if (needsUpdate) updated++;
  }

  console.log('✅ 보완 완료: ' + updated + '개 항목에 문화유산 기본 정보·출처·이미지를 추가했습니다.');
}
