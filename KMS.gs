// @ts-nocheck
// 
/*
      UPDATE LOG
    2023.12.26
     - NEXON Open API 적용

*/
const VERSION = "20231226-001";




// new variable

const MAPLESTORY_APIKEY = 'live_b27763e12fe1e782436f3eda65473a951b57c5a806ac1582a18ea38e922545ab5247416d40421cd8b096a14d6920c612';
const MAPLESTORY_API = 'https://open.api.nexon.com/maplestory/v1';
const GET_OCID = '/id';
const GET_CHAR_BASIC = '/character/basic';
const GET_CHAR_STAT = '/character/stat';

// old variable
const MAPLESTORY_HOME = 'https://maplestory.nexon.com';
const MAPLESTORY_RANKING_SEARCH = `${MAPLESTORY_HOME}/N23Ranking/World/Total`;
const CHARACTER_LINKS_SELECTOR = 'div.rank_table_wrap > table > tbody > tr > td.left > dl > dt > a';

const CHARACTER_NAME_SELECTOR = 'div.char_info_top > div.char_name > span';
const CHARACTER_INFO_SELECTOR = 'div.char_info_top > div.char_info > dl > dd';
const REBOOT_PARAM = '&w=254';
const WORLD_PARAM = '';

const JOB_INFO = [
  // 모험가
  ['히어로', '파이터', '크루세이더'],
  ['팔라딘', '페이지', '나이트'],
  ['다크나이트', '스피어맨', '버서커'],
  ['아크메이지(불,독)', '위자드(불,독)', '메이지(불,독)'],
  ['아크메이지(썬,콜)', '위자드(썬,콜)', '메이지(썬,콜)'],
  ['비숍', '클레릭', '프리스트'],
  ['보우마스터', '헌터', '레인저'],
  ['신궁', '사수', '저격수'],
  ['패스파인더', '에인션트아처', '체이서'],
  ['나이트로드', '어쌔신', '허밋'],
  ['섀도어', '시프', '시프마스터'],
  ['듀얼블레이더', '듀얼블레이드', '세미듀어러', '듀어러', '듀얼마스터', '슬래셔'],
  ['바이퍼', '인파이터', '버커니어'],
  ['캡틴', '건슬링거', '발키리'],
  ['캐논슈터', '캐논블래스터', '캐논마스터']
];




function createSpreadsheetOpenTrigger() {
  const ss = SpreadsheetApp.getActive();

  var check = 0;

  if (!checkIfTriggerExists(ScriptApp.EventType.ON_EDIT, 'unionEdit')) {
    try {
      ScriptApp.newTrigger('unionEdit')
          .forSpreadsheet(ss)
          .onEdit()
          .create();
      check++;
      // SpreadsheetApp.getUi().alert("권한이 정상적으로 적용되었습니다.");
    } catch (e) {
      Logger.log(`error : ${e}`);
      SpreadsheetApp.getUi().alert(`에러가 발생하였습니다 : ${e}`);
    }
  } 

  if (!checkIfTriggerExists(ScriptApp.EventType.CLOCK, 'autoRefresh')) {
    try {
      ScriptApp.newTrigger('autoRefresh')
        .timeBased().everyDays(1)
        .atHour(4)
        .create();
      check++;
    } catch (e) {
      Logger.log(`error : ${e}`);
      SpreadsheetApp.getUi().alert(`에러가 발생하였습니다 : ${e}`);
    }
  }

  if (check == 0) {
    SpreadsheetApp.getUi().alert('권한이 이미 적용되어 있습니다.');
  } else {
    SpreadsheetApp.getUi().alert('권한이 정상적으로 적용되었습니다.');
  }

}



// NEW API (Open API) =========================================================================================================
function test() {
  Logger.log(getCharInfoNew(getCharOCID('TEST NICKNAME')));
}


function getCharOCID(name) {
  if (typeof name == 'number') {
    name = name.toString();
  }

  let url = `${MAPLESTORY_API}${GET_OCID}`;

  let params = {
    'character_name': name
  };

  let headers = {
    'x-nxopen-api-key': MAPLESTORY_APIKEY
  };

  let options = {
    'method' : 'get',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'headers' : headers,
    'muteHttpExceptions' : true

  };

  let response = UrlFetchApp.fetch(buildUrl_(url, params), options);

  if (response.getResponseCode() != 200) {
    // 에러 처리
    Logger.log(`API response error : ${response.getResponseCode()}`);
    return null;
  }

  let data = JSON.parse(response.getContentText());

  Logger.log(`${name} ocid : ${data.ocid}`);

  return data.ocid;
}


function getCharInfoNew(ocid) {

  let url = `${MAPLESTORY_API}${GET_CHAR_BASIC}`;

  let thisTime = new Date();
  thisTime.setDate(thisTime.getDate()-1);
  let theDayBeforeStr = Utilities.formatDate(thisTime, 'Asia/Seoul', 'yyyy-MM-dd');

  Logger.log(`theDayBefore : ${theDayBeforeStr}`);

  let params = {
    'ocid': ocid,
    'date': theDayBeforeStr
  };

  let headers = {
    'x-nxopen-api-key': MAPLESTORY_APIKEY
  };

  let options = {
    'method' : 'get',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'headers' : headers,
    'muteHttpExceptions' : true

  };
  
  let response = UrlFetchApp.fetch(buildUrl_(url, params), options);

  if (response.getResponseCode() != 200) {
    // 에러 처리
    Logger.log(`API response error : ${response.getResponseCode()}`);
    return null;
  }

  let data = JSON.parse(response.getContentText());

  Logger.log(data);

  let charLv = Math.floor(data.character_level);
  let charJob = data.character_class;

  return [charLv, changeCharJob(charJob)];
}


// OLD API (웹 파싱=============================================================================================================


function testOld() {
  Logger.log(getCharInfo(getCharUrl('TEST NICKNAME', true)));
}


function getCharUrl(name, isReboot) {
  // Logger.log(`${MAPLESTORY_RANKING_SEARCH}?c=${encodeURI(name)}${isReboot ? REBOOT_PARAM : WORLD_PARAM}`);

  // name이 int형인지 검사 (2023.12.06)
  if (typeof name == 'number') {
    name = name.toString();
  }

  var url = `${MAPLESTORY_RANKING_SEARCH}?c=${encodeURI(name)}${isReboot ? REBOOT_PARAM : WORLD_PARAM}`;
  var success = false;
  var response = null;
 
  for(var attempts = 0; attempts < 10; attempts++) {
   
    if(!success) {
     
      try {
       
        // call UrlFetchApp.fetch..
        response = UrlFetchApp.fetch(url, {'validateHttpsCertificates': false, 'muteHttpExceptions': true});
        success = true;
        break;
       
      } catch (e) {
       
        Logger.log(`${+attempts+1} 실패 : ${e}`);
        // response = UrlFetchApp.fetch(url);
        // Utilities.sleep(1000);
      }
    }
  }
 
  if(!success) {
   
    Logger.log(`주소 접근 실패 : ${name}`);
    return;
  }

  
  // if (response.getResponseCode() != 200) {
  //   // 재시도
  //   response = UrlFetchApp.fetch(`${MAPLESTORY_RANKING_SEARCH}?c=${encodeURI(name)}${isReboot ? REBOOT_PARAM : WORLD_PARAM}`);
  // }


  var content = response.getContentText();

  const $ = Cheerio.load(content);
  // const node = HtmlParser.of(searchData);

  // const links = $(CHARACTER_LINKS_SELECTOR);
  // const link = links.find((linkNode) => linkNode.innerText.toLowerCase() === name.toLowerCase());
  // if (!link) throw new NotFoundError(name);
  

  const findUrl = $(CHARACTER_LINKS_SELECTOR)
  .filter(function (i, el) {
    // this === el
    return $(this).text().toLowerCase() === name.toLowerCase();
  })
  .attr('href'); //=> orange

  // Logger.log(findUrl);

  if (findUrl === null) {
    Logger.log('캐릭터 찾기 실패');
    return null;
  }

  return findUrl;
}

function getCharInfo(url) {
  var url = `${MAPLESTORY_HOME}${url}`;
  var success = false;
  var response = null;
 
  for(var attempts = 0; attempts < 10; attempts++) {
   
    if(!success) {
     
      try {
       
        // call UrlFetchApp.fetch..
        response = UrlFetchApp.fetch(url, {'validateHttpsCertificates': false, 'muteHttpExceptions': true});
        success = true;
        break;
       
      } catch (e) {
       
        Logger.log(`${+attempts+1} 실패 : ${e}`);
        // Utilities.sleep(1000);
        try {
          UrlFetchApp.fetch(MAPLESTORY_RANKING_SEARCH, {'validateHttpsCertificates': false, 'muteHttpExceptions': true});
        } catch (e) {
          Logger.log(`catch error : ${e}`);
        }
        
      }
    }
  }
 
  if(!success) {
   
    Logger.log(`주소 접근 실패 : ${name}`);
    return;
  }




  var content = response.getContentText();
  var $ = Cheerio.load(content);

  var charInfo = $(CHARACTER_INFO_SELECTOR);
  var charLv = charInfo.eq(0).text().substring(3);
  var charJob = charInfo.eq(1).text().split('/');

  // Logger.log(charLv);
  // Logger.log(charJob[1]);
  // if (charLv == '' || charJob == '' || charLv == null || charJob == null) {
  //   Logger.log(`캐릭터 레벨, 직업 정보 가져오기 오류 : ${charLv == null}, ${charJob == null}`);
  //   // errorCount++;
  //   content = UrlFetchApp.fetch(`${MAPLESTORY_HOME}${url}`).getContentText();
  //   $ = Cheerio.load(content);

  //   charInfo = $(CHARACTER_INFO_SELECTOR);
  //   charLv = charInfo.eq(0).text().substring(3);
  //   charJob = charInfo.eq(1).text().split('/');
  //   Logger.log(`재시도 결과 : ${charLv}, ${charJob}`);
  // }

  return [charLv, changeCharJob(charJob[1])];
}

function changeCharJob(job) {
  if (job == '' || job == null) {
    return '';
  }

  var searchJob = job.trim();

  for (var element of JOB_INFO) {
    if (element.includes(searchJob, 1)) {
      Logger.log(`직업 이름 변환 : ${searchJob} - ${element[0]}`);
      return element[0];
    }
  }
  return searchJob;
}

function autoRefreshUnionList() {
  try {
    var thisTime = new Date();
    var cnt = allRefreshUnionList();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('유니온');
    sheet.getRange('autoRefreshTime').setValue(thisTime);

  } catch (e) {
    Logger.log(`error : ${e}`);
    // SpreadsheetApp.getUi().alert(`에러가 발생하였습니다 : ${e}`);
  }
  

}

function allRefreshUnionList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('유니온');
  const isReboot = sheet.getRange('isReboot').getValues()[0][0];
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('데이터');
  const charNameStartCol = Number(dataSheet.getRange('charNameCol').getValues());
  const charNameStartRow = Number(dataSheet.getRange('charNameRow').getValues());
  const charCount = Number(dataSheet.getRange('charCount').getValues());
  const charUrlCol = Number(dataSheet.getRange('charUrlCol').getValues());
  const charJobCol = Number(dataSheet.getRange('charJobCol').getValues());
  const charLevelCol = Number(dataSheet.getRange('charLevelCol').getValues());

  var existList = sheet.getRange(charNameStartRow, charNameStartCol, charCount).getValues();
  existList = existList.map(function(element) {
    return element[0];
  }).filter(function(element) {
    if (element == '' || element == null) {
      return false;
    } else {
      return true;
    }
  });

  var totalRefreshCount = 0;

  for (var i in existList) {
    // var charUrl = sheet.getRange(+charNameStartRow + +i, charUrlCol).getValues()[0][0];

    // Logger.log(`${i} : ${charUrl}`);
    var isOldRank = false;

    // URL 새로 가져오기
    // var tempCharUrl = getCharUrl(existList[i], isReboot);
    var tempCharUrl = getCharOCID(existList[i]);
    if (tempCharUrl == null) {
      // API 가져오기 실패 -> 랭킹페이지 URL 가져오기 시도
      tempCharUrl = getCharUrl(existList[i], isReboot);
      if (tempCharUrl == null) {
        tempCharUrl = 'error';
      } else {
        isOldRank = true;
      }
    }
    sheet.getRange(+charNameStartRow + +i, charUrlCol).setValue(tempCharUrl);
    Logger.log(`${i} : ${existList[i]} : ${tempCharUrl}`);

    if (tempCharUrl != 'error') {
      var charInfo;
      if (isOldRank) {
        charInfo = getCharInfo(tempCharUrl);
      } else {
        charInfo = getCharInfoNew(tempCharUrl);
      }
      Logger.log(`${i} : ${charInfo}`);
      sheet.getRange(+charNameStartRow + +i, charJobCol).setValue(charInfo[1]);
      sheet.getRange(+charNameStartRow + +i, charLevelCol).setValue(charInfo[0]);

      totalRefreshCount++;
    }
  }

  return totalRefreshCount;
}



function unionEdit(e) {
  var range = e.range;
  var spreadSheet = e.source;
  var sheet = spreadSheet.getActiveSheet();
  var sheetName = spreadSheet.getActiveSheet().getName();
  var column = range.getColumn();
  var row = range.getRow();
  var lastRow = range.getLastRow();
  // var inputValue = typeof e.value === 'string' ? e.value.trim() : e.value;

  if (sheetName == '유니온') {

    SpreadsheetApp.flush();
    const isReboot = sheet.getRange('isReboot').getValues()[0][0];
    const dataSheet = spreadSheet.getSheetByName('데이터');
    const charNameStartCol = Number(dataSheet.getRange('charNameCol').getValues());
    const charNameStartRow = Number(dataSheet.getRange('charNameRow').getValues());
    const charCount = Number(dataSheet.getRange('charCount').getValues());
    const charUrlCol = Number(dataSheet.getRange('charUrlCol').getValues());
    const charJobCol = Number(dataSheet.getRange('charJobCol').getValues());
    const charLevelCol = Number(dataSheet.getRange('charLevelCol').getValues());

    const setAllRefreshCol = sheet.getRange('setAllRefresh').getColumn();
    const setAllRefreshRow = sheet.getRange('setAllRefresh').getRow();

    // Logger.log(row);

    // 캐릭터 닉네임 영역을 수정했는지 검사
    if (column == charNameStartCol && row >= charNameStartRow && row <= charNameStartRow + charCount - 1) {
      if (lastRow >= charNameStartRow + charCount - 1) lastRow = charNameStartRow + charCount - 1;

      var values = sheet.getRange(row, charNameStartCol, lastRow - row + 1).getValues();

      for (var i in values) {
        for (var j in values[i]) {
          
          var inputValue = typeof values[i][j] === 'string' ? values[i][j].trim() : values[i][j];
          Logger.log(inputValue);
          // Logger.log(+row + +i);
          if (inputValue == '' || inputValue == null) {
            sheet.getRange(+row + +i, charUrlCol).setValue('');
            sheet.getRange(+row + +i, charJobCol).setValue('');
            sheet.getRange(+row + +i, charLevelCol).setValue('');

          } else {
            var isOldRank = false;
            var charUrl = getCharOCID(inputValue);
            if (charUrl == null) {
              // API 가져오기 실패 -> 랭킹페이지 URL 가져오기 시도
              charUrl = getCharUrl(inputValue, isReboot);
              if (tempCharUrl == null) {
                charUrl = 'error';
              } else {
                isOldRank = true;
              }
            }
            
            sheet.getRange(+row + +i, charUrlCol).setValue(charUrl);
            if (charUrl != 'error') {
              var charInfo;
              if (isOldRank) {
                charInfo = getCharInfo(charUrl);
              } else {
                charInfo = getCharInfoNew(charUrl);
              }
              sheet.getRange(+row + +i, charJobCol).setValue(charInfo[1]);
              sheet.getRange(+row + +i, charLevelCol).setValue(charInfo[0]);
            }
          }

        }
      }

    } else if (column == setAllRefreshCol && row == setAllRefreshRow && e.value == "TRUE") {
      // 전체 새로고침을 체크한 경우
      try {
        // SpreadsheetApp.flush();
        // 유니온 시트에 이미 작성된 캐릭터 리스트
        var totalRefreshCount = allRefreshUnionList();
        var thisTime = new Date();
        sheet.getRange('autoRefreshTime').setValue(thisTime);
        SpreadsheetApp.getUi().alert(`총 ${totalRefreshCount}개의 캐릭터 정보를 갱신하였습니다.`);
      } catch (e) {
        Logger.log(`error : ${e}`);
        SpreadsheetApp.getUi().alert(`에러가 발생하였습니다 : ${e}`);
      } finally {
        sheet.getRange('setAllRefresh').setValue("FALSE");
      }
    }
    // var charNameRange = sheet.getRange(charNameStartRow, charNameStartCol, charCount);
    // var charNameValues = range.getValues();

  } else if (sheetName == "캐릭터 자동 입력") {
    // 캐릭터 자동 입력 칸에 값을 붙여넣은 경우
    SpreadsheetApp.flush();
    // const isReboot = sheet.getRange('autoInsertIsReboot').getValues()[0][0];
    var isReboot = false;
    const unionSheet = spreadSheet.getSheetByName('유니온');
    const dataSheet = spreadSheet.getSheetByName('데이터');
    const charUrlCol = Number(dataSheet.getRange('charUrlCol').getValues());
    const charJobCol = Number(dataSheet.getRange('charJobCol').getValues());
    const charLevelCol = Number(dataSheet.getRange('charLevelCol').getValues());

    const autoInsertCellCol = sheet.getRange('autoInsertCell').getColumn();
    const autoInsertCellRow = sheet.getRange('autoInsertCell').getRow();

    if (column == autoInsertCellCol && row == autoInsertCellRow) {

      var charIdx = -1;
      var charList = new Array();
      var values = sheet.getRange(autoInsertCellRow, autoInsertCellCol, 200).getValues();
      for (var i in values) {
        var j = 0;
        // Logger.log(`${i} ${j} ${typeof values[i][j]}`);
        if (typeof values[i][j] !== 'string') {
          continue;
        } else if (values[i][j].includes('월드/캐릭터 선택')) {
          // 한번 건너 뛰고 수집해야됨
          charIdx = +i + 2;
        } else if (values[i][j].includes('대표캐릭터는 10레벨 이상이어야 지정할 수 있습니다.')) {
          break;
        } else if (charIdx == +i + 1) {
          // 월드 감지
          var world = values[i][j].split('월드선택');
          if (world[0].includes('리부트')) {
            Logger.log(`리부트 월드 감지 : ${world[0]}`);
            isReboot = true;
          }
        } else if (charIdx != -1 && charIdx <= i) {
          // Logger.log(`${i} ${charIdx} ${values[i][j]}`);
          charList.push(values[i][j]);
        }
        
      }

      // Logger.log(charList);
      // 캐릭터 리스트 : charList
      if (charList.length == 0) return;

      // 리부트 체크 활성화
      unionSheet.getRange('isReboot').setValue(isReboot);

      const dataSheet = spreadSheet.getSheetByName('데이터');
      const charNameStartCol = Number(dataSheet.getRange('charNameCol').getValues());
      const charNameStartRow = Number(dataSheet.getRange('charNameRow').getValues());
      const charCount = Number(dataSheet.getRange('charCount').getValues());

      // 유니온 시트에 이미 작성된 캐릭터 리스트
      var existList = spreadSheet.getSheetByName('유니온').getRange(charNameStartRow, charNameStartCol, charCount).getValues();
      existList = existList.map(function(element) {
        return element[0];
      }).filter(function(element) {
        if (element == '' || element == null) {
          return false;
        } else {
          return true;
        }
      });

      // Logger.log(existList.length);


      // 유니온 시트에 없는 캐릭터 리스트 배열
      var nonExistList = charList.filter(x => !existList.includes(x));

      // Logger.log(nonExistList);
      for (var i in nonExistList) {
        unionSheet.getRange(+charNameStartRow + +existList.length + +i, charNameStartCol).setValue(nonExistList[i]);
        var isOldRank = false;
        // Logger.log(nonExistList[i]);

        // var charUrl = getCharUrl(nonExistList[i], isReboot);
        var charUrl = getCharOCID(nonExistList[i]);
        if (charUrl == null) {
          // API 가져오기 실패 -> 랭킹페이지 URL 가져오기 시도
          charUrl = getCharUrl(nonExistList[i], isReboot);
          if (charUrl == null) {
            charUrl = 'error';
          } else {
            isOldRank = true;
          }
        }
        
        unionSheet.getRange(+charNameStartRow + +existList.length + +i, charUrlCol).setValue(charUrl);

        if (charUrl != 'error') { 
          var charInfo;
          if (isOldRank) {
            charInfo = getCharInfo(charUrl);
          } else {
            charInfo = getCharInfoNew(charUrl);
          }
          unionSheet.getRange(+charNameStartRow + +existList.length + +i, charJobCol).setValue(charInfo[1]);
          unionSheet.getRange(+charNameStartRow + +existList.length + +i, charLevelCol).setValue(charInfo[0]);
        }


      }


      // 셀 초기화
      sheet.getRange(autoInsertCellRow, autoInsertCellCol, 200).clear();
      // sheet.getRange('autoInsertIsReboot').setValue('FALSE');
      sheet.getRange(autoInsertCellRow, autoInsertCellCol).setBorder(true /*top*/, true /*left*/, true /*bottom*/, true /*right*/, null  /*vertical*/, null  /*horizontal*/, "black" /*color*/, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
  }
}


function onOpen() {
  let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('데이터');
  let isGMS = false;

  if(dataSheet == null) {
    dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
    isGMS = true;
  }

  let sheetVersion;
  

  if(isGMS) {
    // GMS
    sheetVersion = dataSheet.getRange('D2').getValues()[0][0];
    // Logger.log(`OK ${sheetVersion}`)

    // 2023.12.26.001 버그 : 자동 업데이트 안되는 문제..
    if (sheetVersion != '2024.06.18.001') {

      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('Important Notice', '6월 18일 이슈사항 대응방법에 대한 안내입니다.\nThis is guidance on how to respond to the issues of June 18th.\n\nhttps://gall.dcinside.com/globalms/29500', ui.ButtonSet.OK);

      response = ui.alert('Important Notice', '공지사항을 확인하였습니까? 작업을 수행하였으면 확인을 눌러 이 메시지를 다시 보지 않습니다.\nDid you check the notice? When youre done, click YES to avoid seeing this message again.', ui.ButtonSet.YES_NO);

      // Process the user's response.
      if (response == ui.Button.YES) {
        dataSheet.getRange('D2').setValue('2024.06.18.001');
      } else {
        Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
      }
    }

  } else { 
    // KMS
    sheetVersion = dataSheet.getRange('B2').getValues()[0][0];

    if (sheetVersion != '2023.12.26.001') {

      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('중요 공지사항', 'Open API 적용 관련 공지사항을 확인하세요.\n조회되지 않는 캐릭터의 경우 반드시 접속을 한번 해주셔야 정상 조회가 가능합니다.\nhttps://www.inven.co.kr/board/maple/2304/36025', ui.ButtonSet.OK);

      response = ui.alert('중요 공지사항', '공지사항을 확인하였습니까? 작업을 수행하였으면 확인을 눌러 이 메시지를 다시 보지 않습니다.', ui.ButtonSet.YES_NO);

      // Process the user's response.
      if (response == ui.Button.YES) {
        dataSheet.getRange('B2').setValue('2023.12.26.001');
      } else {
        Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
      }
    }


    // 2023.10.12.001 버그 : 자동 업데이트 안되는 문제..
    // if (sheetVersion != '2023.10.17.001') {

    //   var ui = SpreadsheetApp.getUi();
    //   var response = ui.alert('중요 공지사항', '자동 업데이트 기능 문제로 인해 아래 링크의 조치사항을 확인하세요.\nhttps://www.inven.co.kr/board/maple/2304/36025', ui.ButtonSet.OK);

    //   response = ui.alert('중요 공지사항', '링크대로 조치하였습니까? 작업을 수행하였으면 확인을 눌러 이 메시지를 다시 보지 않습니다.', ui.ButtonSet.YES_NO);

    //   // Process the user's response.
    //   if (response == ui.Button.YES) {
    //     dataSheet.getRange('B2').setValue('2023.10.17.001');
    //   } else {
    //     Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
    //   }
    // }

  }


}

function getVersion() {
  Logger.log(VERSION);
  return VERSION;
}
