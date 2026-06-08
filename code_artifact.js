/**
 * 웹 앱 URL로 접속했을 때 HTML 페이지를 렌더링하는 함수입니다.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('학급 투표 앱')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 스프레드시트에서 투표 결과를 가져오는 함수입니다.
 * (클라이언트에서 google.script.run 으로 호출됨)
 */
function getVotes() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    // 기본 결과 객체 정의
    let result = {
      option1: 0,
      option2: 0,
      option3: 0
    };
    
    if (lastRow < 2) {
      // 데이터가 없는 경우 기본값 세팅 시도 (필요시 데이터 자동 추가)
      setupDefaultSheet(sheet);
      return result;
    }
    
    // A열(체험학습 장소)과 B열(득표수) 데이터를 모두 가져옴 (헤더 제외한 2행부터)
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    // 각 옵션 매핑용 텍스트 정의 (Index.html의 버튼 텍스트 내용과 맞춰줍니다)
    const mapping = {
      "option1": "놀이공원",
      "option2": "해수욕장",
      "option3": "박물관"
    };
    
    // 시트의 각 행을 돌며 매핑되는 항목의 득표수를 결과 객체에 반영
    data.forEach(function(row) {
      const place = String(row[0]).trim();
      const count = parseInt(row[1] || '0', 10);
      
      if (place.includes(mapping.option1)) {
        result.option1 = count;
      } else if (place.includes(mapping.option2)) {
        result.option2 = count;
      } else if (place.includes(mapping.option3)) {
        result.option3 = count;
      }
    });
    
    return result;
    
  } catch (error) {
    Logger.log("getVotes Error: " + error.toString());
    throw new Error("스프레드시트 데이터를 가져오는 데 실패했습니다: " + error.message);
  }
}

/**
 * 특정 항목에 투표를 추가하는 함수입니다.
 * 스프레드시트 A열에서 매핑 텍스트를 찾아 B열의 값을 +1 해줍니다.
 * (클라이언트에서 google.script.run 으로 호출됨)
 */
function addVote(option) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    const mapping = {
      "option1": "놀이공원",
      "option2": "해수욕장",
      "option3": "박물관"
    };
    
    const targetPlace = mapping[option];
    if (!targetPlace) {
      throw new Error("올바르지 않은 투표 항목입니다.");
    }
    
    let found = false;
    
    if (lastRow >= 2) {
      const range = sheet.getRange(2, 1, lastRow - 1, 2);
      const data = range.getValues();
      
      // 행을 돌며 일치하는 체험학습 장소를 검색
      for (let i = 0; i < data.length; i++) {
        const place = String(data[i][0]).trim();
        if (place.includes(targetPlace)) {
          // 일치하는 장소를 찾으면 해당 행의 B열(득표수) 값을 1 증가시킴
          // i는 0부터 시작하고 데이터는 2행부터 시작하므로 실제 행 번호는 i + 2 임
          const rowNum = i + 2;
          const currentCount = parseInt(data[i][1] || '0', 10);
          sheet.getRange(rowNum, 2).setValue(currentCount + 1);
          found = true;
          break;
        }
      }
    }
    
    // 만약 시트에 해당 체험학습 장소가 기록되어 있지 않다면 새롭게 행을 생성해 줌
    if (!found) {
      sheet.appendRow([targetPlace, 1]);
    }
    
    // 최신 결과를 반영하여 전달
    return getVotes();
    
  } catch (error) {
    Logger.log("addVote Error: " + error.toString());
    throw new Error("투표 저장 중 오류 발생: " + error.message);
  }
}

/**
 * 모든 투표 결과를 0으로 초기화하는 함수입니다.
 * 스프레드시트의 B열 득표수 값을 모두 0으로 만듭니다.
 * (클라이언트에서 google.script.run 으로 호출됨)
 */
function resetVotes() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow >= 2) {
      const range = sheet.getRange(2, 2, lastRow - 1, 1); // B열 전체 선택
      const resetValues = [];
      for (let i = 0; i < lastRow - 1; i++) {
        resetValues.push([0]);
      }
      range.setValues(resetValues);
    }
    
    return getVotes();
  } catch (error) {
    Logger.log("resetVotes Error: " + error.toString());
    throw new Error("초기화 중 오류 발생: " + error.message);
  }
}

/**
 * 시트에 데이터가 아예 없을 때 기본 양식을 만들어 주는 보조 함수입니다.
 */
function setupDefaultSheet(sheet) {
  sheet.clear();
  sheet.appendRow(["체험학습 장소", "득표수"]);
  sheet.appendRow(["놀이공원", 0]);
  sheet.appendRow(["해수욕장", 0]);
  sheet.appendRow(["박물관", 0]);
  
  // 가독성을 위한 헤더 스타일 지정
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#EEF2F6");
}

/**
 * 구글 앱스 스크립트 에디터에서 수동으로 테스트 실행해 볼 수 있는 함수입니다.
 */
function myFunction() {
  Logger.log("스프레드시트 연동 투표 앱 테스트를 실행합니다.");
  try {
    const currentVotes = getVotes();
    Logger.log("현재 스프레드시트 기록: " + JSON.stringify(currentVotes));
  } catch(e) {
    Logger.log("오류 발생: " + e.message);
  }
}
