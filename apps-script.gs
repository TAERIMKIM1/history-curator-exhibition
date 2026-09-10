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

  SpreadsheetApp.getUi().alert('✅ 완료!\n' + updated + '개 패널에 유물명·기본 제원·출처를 채웠습니다.');
}

var TEACHER_PASSWORD = "0070"; // 교사 비밀번호 (원하는 대로 변경)

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

// ============================================================
// ★ 교사용 일괄 보완 함수
// Apps Script 에디터에서 fillMissingInfo 선택 후 ▶ 실행
// 유물 기본 정보(G열)·출처(J열)가 비어 있는 행만 자동 채워줌
// ============================================================
function fillMissingInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();

  // 유물명 → {relicSpec, source} 매핑 테이블
  var INFO = {
    // ── 고려관 ──
    "청자 상감운학문 매병":
      { spec: "12세기(고려 중기), 청자에 상감기법, 높이 43.9cm, 국보 제68호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄 (www.emuseum.go.kr), 우리역사넷 (contents.history.go.kr), 국가유산포털 (www.heritage.go.kr)" },
    "팔만대장경":
      { spec: "1236~1251년(고려 고종 23~38년), 목판 인쇄본, 81,258판, 국보 제32호·유네스코 세계기록유산(1997), 해인사 장경판전 소장",
        src:  "해인사 공식 사이트, 우리역사넷 (contents.history.go.kr), 유네스코 세계기록유산 공식 자료" },
    "고려대장경":
      { spec: "1236~1251년(고려 고종 23~38년), 목판 인쇄본, 81,258판, 국보 제32호·유네스코 세계기록유산(1997), 해인사 장경판전 소장",
        src:  "해인사 공식 사이트, 우리역사넷 (contents.history.go.kr), 유네스코 세계기록유산 공식 자료" },
    "직지심체요절":
      { spec: "1377년(고려 우왕 3년), 금속활자 인쇄본, 세계 최초 금속활자 인쇄본, 유네스코 세계기록유산(2001), 프랑스 국립도서관 소장",
        src:  "국립중앙박물관 e뮤지엄, 청주고인쇄박물관 (www.jikjiworld.cheongju.go.kr), 우리역사넷" },
    "삼국유사":
      { spec: "1281년경(고려 충렬왕, 일연 저술), 목판 인쇄본, 국보 제306호, 연세대학교 박물관 소장(인각사본)",
        src:  "우리역사넷, 한국사데이터베이스 (db.history.go.kr), 국가유산포털" },
    "청자 투각 칠보문 향로":
      { spec: "12세기(고려 중기), 청자에 투각기법, 높이 15.3cm, 국보 제95호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털" },
    "청동 은입사 포류수금문 정병":
      { spec: "12세기(고려 중기), 청동에 은입사기법, 높이 37.5cm, 국보 제92호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷" },
    "고려 수월관음도":
      { spec: "14세기(고려 후기), 비단에 금니·채색, 세로 419.5cm, 국보 제315호, 아모레퍼시픽미술관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷" },
    "청자 참외형 병":
      { spec: "12세기(고려 중기), 순청자 음각기법, 높이 22.6cm, 국보 제94호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털" },
    "고려 나전칠기 경함":
      { spec: "13세기(고려 중기), 나무에 나전·옻칠, 국보 제102호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷" },
    "고려 금동 대탑":
      { spec: "10~11세기(고려 전기), 금동 주조·도금, 높이 155cm, 국보 제213호, 국립중앙박물관 소장",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털" },
    // ── 조선관 ──
    "훈민정음 해례본":
      { spec: "1443년 창제·1446년(세종 28년) 반포, 목판 인쇄본(해례본), 국보 제70호·유네스코 세계기록유산(1997), 간송미술관 소장",
        src:  "국립한글박물관 (www.hangeul.go.kr), 우리역사넷, 유네스코 세계기록유산 공식 자료" },
    "훈민정음":
      { spec: "1443년 창제·1446년(세종 28년) 반포, 목판 인쇄본(해례본), 국보 제70호·유네스코 세계기록유산(1997), 간송미술관 소장",
        src:  "국립한글박물관 (www.hangeul.go.kr), 우리역사넷, 유네스코 세계기록유산 공식 자료" },
    "측우기":
      { spec: "1441년(세종 23년), 청동 원통형, 높이 30.0cm·지름 15.0cm, 세계 최초 강우량 측정기, 국보 제561호, 국립기상박물관 소장",
        src:  "우리역사넷, 국가유산포털, 기상청 국립기상박물관" },
    "앙부일구":
      { spec: "1434년(세종 16년) 제작·현존품은 17세기, 청동 주조, 국보 제845호, 국립고궁박물관 소장",
        src:  "국립고궁박물관 (www.gogung.go.kr), 우리역사넷, 국가유산포털" },
    "자격루":
      { spec: "1434년(세종 16년), 청동·목재, 물시계(자동 시보 장치), 국보 제229호, 국립고궁박물관 소장",
        src:  "국립고궁박물관, 우리역사넷, 국가유산포털" },
    "조선왕조실록":
      { spec: "1413~1865년(태조~철종), 지본 필사·목판본, 총 1,893권, 국보 제151호·유네스코 세계기록유산(1997), 서울대학교 규장각 소장",
        src:  "한국사데이터베이스 (db.history.go.kr), 우리역사넷, 유네스코 세계기록유산 자료" },
    "경국대전":
      { spec: "1485년(성종 16년) 완성, 목판 인쇄본, 조선의 기본 법전, 보물 제1521호, 규장각한국학연구원 소장",
        src:  "한국사데이터베이스, 우리역사넷, 규장각한국학연구원 원문검색" },
    "동의보감":
      { spec: "1613년(광해군 5년), 목판 인쇄본, 총 25권, 국보 제319호·유네스코 세계기록유산(2009), 국립중앙도서관 소장",
        src:  "국립중앙도서관, 우리역사넷, 유네스코 세계기록유산 자료" },
    "혼일강리역대국도지도":
      { spec: "1402년(태종 2년), 채색 필사본, 세로 158cm·가로 168cm, 현존 동양 최고 세계지도 중 하나, 일본 류코쿠대학 소장",
        src:  "우리역사넷, 한국사데이터베이스, 국립중앙박물관 e뮤지엄" },
    "신숙주 초상":
      { spec: "15세기(조선 전기), 비단에 채색, 세로 167cm·가로 109.5cm, 국보 제613호, 국립중앙박물관 기탁",
        src:  "국립중앙박물관 e뮤지엄, 국가유산포털, 우리역사넷" },
    "칠정산 내편":
      { spec: "1444년(세종 26년), 지본 필사·목판본, 조선 독자 역법서, 보물 제1248호, 서울대학교 규장각 소장",
        src:  "우리역사넷, 한국사데이터베이스, 규장각한국학연구원" },
    "혼천의":
      { spec: "15세기(세종~성종), 청동 주조, 천체 관측 기구, 국보 제230호, 고려대학교 박물관 소장",
        src:  "국립고궁박물관, 우리역사넷, 국가유산포털" }
  };

  var updated = 0;
  for (var i = 1; i < rows.length; i++) {
    var relicName = String(rows[i][5]).trim();
    var relicSpec = String(rows[i][6]).trim();
    var source    = String(rows[i][9]).trim();

    if (!relicName) continue; // 유물명도 없으면 건너뜀

    var info = INFO[relicName];
    if (!info) continue; // 매핑 테이블에 없으면 건너뜀

    var needsUpdate = false;
    if (!relicSpec || relicSpec.length < 5) {
      sheet.getRange(i + 1, 7).setValue(info.spec);
      needsUpdate = true;
    }
    if (!source || source.length < 5) {
      sheet.getRange(i + 1, 10).setValue(info.src);
      needsUpdate = true;
    }
    if (needsUpdate) updated++;
  }

  SpreadsheetApp.getUi().alert(
    '✅ 보완 완료\n' + updated + '개 항목에 유물 기본 정보·출처를 추가했습니다.'
  );
}
