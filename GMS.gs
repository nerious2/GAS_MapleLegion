// GMS variable
/*
https://www.nexon.com/api/maplestory/no-auth/v1/ranking/na?type=job&reboot_index=0&page_index=1&character_name=

Magician
불독 12 / 2
썬콜 22 / 2
비숍 32 / 2

Warrior
히어로 12 / 1
팔라딘 22 / 1
다크나이트 32 / 1

Bowman
보우마스터 Bowmaster 12 / 3
신궁 Marksman 22 / 3
패스파인더 32 / 3  "jobName": "Pathfinder"

Thief
Night Lord 나이트로드 12 / 4
섀도어 22 / 4
듀얼블레이드 34 / 4 "jobName": "Dual Blade"

Pirate
Buccaneer 바이퍼 12 / 5
Corsair 캡틴 22 / 5
Cannoneer 캐논슈터 32 / 5

*/

const GMS_API = 'https://www.nexon.com/api/maplestory/no-auth/v1/ranking/';
const GMS_SERVER_NA = 'na';
const GMS_SERVER_EU = 'eu';
// const GMS_RANKING_SEARCH = `?type=job&reboot_index=0&page_index=1&character_name=`;


function convertUtf8ToEuckrToANSI(utf8Str) {
  var encoded = Utilities.newBlob('').setDataFromString(utf8Str, "EUC-KR");
  return encoded.getDataAsString('windows-1252');
}


function gms_createSpreadsheetOpenTrigger() {
  const ss = SpreadsheetApp.getActive();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Legion');
  const selLanguage = sheet.getRange('selLanguage').getValues()[0][0].toLowerCase();

  var check = 0;

  if (!checkIfTriggerExists(ScriptApp.EventType.ON_EDIT, 'legionEdit')) {
    try {
      ScriptApp.newTrigger('legionEdit')
          .forSpreadsheet(ss)
          .onEdit()
          .create();
      check++;
      // SpreadsheetApp.getUi().alert("권한이 정상적으로 적용되었습니다.");
    } catch (e) {
      Logger.log(`error : ${e}`);
      SpreadsheetApp.getUi().alert(`${selLanguage == 'english' ? 'An error occurred' : '에러가 발생하였습니다'} : ${e}`);
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
      SpreadsheetApp.getUi().alert(`${selLanguage == 'english' ? 'An error occurred' : '에러가 발생하였습니다'} : ${e}`);
    }
  }

  if (check == 0) {
    SpreadsheetApp.getUi().alert(selLanguage == 'english' ? 'Permissions are already applied.' : '권한이 이미 적용되어 있습니다.');
  } else {
    SpreadsheetApp.getUi().alert(selLanguage == 'english' ? 'Permissions have been applied successfully.' : '권한이 정상적으로 적용되었습니다.');
  }

}



function testGMS() {
  Logger.log(parseMapleStoryRanking("na", "테스트할닉네임입력"));
}

// GMS api =======================================================================
function gms_parseMapleStoryRanking(server, name) {
  var url = `${GMS_API}${server=='eu'? GMS_SERVER_EU : GMS_SERVER_NA}`;
// const GMS_RANKING_SEARCH = `?type=job&reboot_index=0&page_index=1&character_name=`;


  let params = {
    'type': 'job',
    'reboot_index': 0,
    'page_index': 1,
    'character_name': name
  };

  let options = {
    'method' : 'get',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    // 'headers' : headers,
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
  let charImage = data.character_image;

  if (data.totalCount == 0) {
    charLv = -1;
    charJob = "";
    charImage = "";
  } else {
    charLv = Math.floor(data.ranks[0].level);
    charImage = data.ranks[0].characterImgURL;

    switch(data.ranks[0].jobName) {
      case "Warrior" :
        switch(data.ranks[0].jobDetail) {
          case 12:
            charJob = 'Hero';
            break;
          case 22:
            charJob = 'Paladin';
            break;
          case 32:
            charJob = 'Dark Knight';
            break;
          default:
            charJob = 'Warrior';
            break;
        }
        break;

      case "Magician" :
        switch(data.ranks[0].jobDetail) {
          case 12:
            charJob = 'Magician (Fire, Poison)';
            break;
          case 22:
            charJob = 'Magician (Ice, Lightning)';
            break;
          case 32:
            charJob = 'Bishop';
            break;
          default:
            charJob = 'Magician';
            break;
        }
        break;
        
      case "Bowman" :
        switch(data.ranks[0].jobDetail) {
          case 12:
            charJob = 'Bowmaster';
            break;
          case 22:
            charJob = 'Marksman';
            break;
          default:
            charJob = 'Bowman';
            break;
        }
        break;

      case "Thief" :
        switch(data.ranks[0].jobDetail) {
          case 12:
            charJob = 'Night Lord';
            break;
          case 22:
            charJob = 'Shadower';
            break;
          default:
            charJob = 'Thief';
            break;
        }
        break;

      case "Pirate" :
        switch(data.ranks[0].jobDetail) {
          case 12:
            charJob = 'Buccaneer';
            break;
          case 22:
            charJob = 'Corsair';
            break;
          case 32:
            charJob = 'Cannoneer';
            break;
          default:
            charJob = 'Pirate';
            break;
        }
        break;

      default :
        charJob = data.ranks[0].jobName;
        break;
    }
  }

  return [charLv, charJob, charImage];

}



function gms_LegionEdit(e) {
  var range = e.range;
  var spreadSheet = e.source;
  var sheet = spreadSheet.getActiveSheet();
  var sheetName = spreadSheet.getActiveSheet().getName();
  var column = range.getColumn();
  var row = range.getRow();
  var lastRow = range.getLastRow();
  // var inputValue = typeof e.value === 'string' ? e.value.trim() : e.value;

  if (sheetName == 'Legion') {

    SpreadsheetApp.flush();
    const dataSheet = spreadSheet.getSheetByName('Data');
    const charNameStartCol = Number(dataSheet.getRange('charNameCol').getValues());
    const charNameStartRow = Number(dataSheet.getRange('charNameRow').getValues());
    const charCount = Number(dataSheet.getRange('charCount').getValues());
    const charUrlCol = Number(dataSheet.getRange('charUrlCol').getValues());
    const charJobCol = Number(dataSheet.getRange('charJobCol').getValues());
    const charJobRow = Number(dataSheet.getRange('charJobRow').getValues());
    const charLevelCol = Number(dataSheet.getRange('charLevelCol').getValues());
    const gmsServer = sheet.getRange('gmsServer').getValues()[0][0];
    let selLanguage = sheet.getRange('selLanguage').getValues()[0][0];

    if (selLanguage == null) {
      Logger.log('Language cell is empty. Force change Engish.');
      selLanguage = 'english';
    } else {
      selLanguage = selLanguage.toLowerCase();
    }

    

    const setAllRefreshCol = sheet.getRange('setAllRefresh').getColumn();
    const setAllRefreshRow = sheet.getRange('setAllRefresh').getRow();

    const selLanguageCol = sheet.getRange('selLanguage').getColumn();
    const selLanguageRow = sheet.getRange('selLanguage').getRow();
    

    // Logger.log(row);

    // 캐릭터 닉네임 영역을 수정했는지 검사
    if (column == charNameStartCol && row >= charNameStartRow && row <= charNameStartRow + charCount - 1) {
      if (lastRow >= charNameStartRow + charCount - 1) lastRow = charNameStartRow + charCount - 1;

      var values = sheet.getRange(row, charNameStartCol, lastRow - row + 1).getValues();

      Logger.log(`server : ${gmsServer} / lang : ${selLanguage}`);

      for (var i in values) {
        for (var j in values[i]) {
          
          var inputValue = typeof values[i][j] === 'string' ? values[i][j].trim() : values[i][j];
          Logger.log(inputValue);
          // Logger.log(+row + +i);
          if (inputValue == '' || inputValue == null) {
            sheet.getRange(+row + +i, +charUrlCol - 1).setValue('');
            sheet.getRange(+row + +i, charUrlCol).setValue('');
            sheet.getRange(+row + +i, charJobCol).setValue('');
            sheet.getRange(+row + +i, charLevelCol).setValue('');

          } else {

            // 한글이 포함되어 있는지 검사
            if(inputValue.match(/[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]/) != null) {
              inputValue = convertUtf8ToEuckrToANSI(inputValue);
              Logger.log(`change ANSI : ${inputValue}`);
              sheet.getRange(+row + +i, charNameStartCol+1).setValue(inputValue);
            } else {
              // 한글이 없으면 ANSI 셀 지우기
              sheet.getRange(+row + +i, +charUrlCol - 1).setValue('');
            }

            var charInfo = gms_parseMapleStoryRanking(gmsServer.toLowerCase(), inputValue);

            let charUrl = 'ok';
            if (charInfo[0] == -1) {
              charUrl = 'error';
            }
            
            sheet.getRange(+row + +i, charUrlCol).setValue(charUrl);
            if (charUrl != 'error') {

              if (selLanguage == 'korean') {
                let korJob = findEngJobToKor(false, charInfo[1]);
                if(korJob == null) {
                  Logger.log(`search failed job list : return ${charInfo[1]}`)
                  korJob = charInfo[1];
                } else {
                  Logger.log(` - change job name : ${korJob}`);
                }
                sheet.getRange(+row + +i, charJobCol).setValue(korJob);
              } else {
                sheet.getRange(+row + +i, charJobCol).setValue(charInfo[1]);
              }
              sheet.getRange(+row + +i, charLevelCol).setValue(charInfo[0]);
              sheet.getRange(+row + +i, charNameStartCol - 1).setValue("=image(\""+ charInfo[2] + "\")");
            }
          }

        }
      }

    } else if (column == setAllRefreshCol && row == setAllRefreshRow && e.value == "TRUE") {
      // 전체 새로고침을 체크한 경우
      try {
        // SpreadsheetApp.flush();
        // 유니온 시트에 이미 작성된 캐릭터 리스트
        var totalRefreshCount = gms_allRefreshLegionList();
        var thisTime = new Date();
        sheet.getRange('autoRefreshTime').setValue(thisTime);
        if (selLanguage == 'english') {
          SpreadsheetApp.getUi().alert(`A total of ${totalRefreshCount} character information has been updated.`);
        } else {
          SpreadsheetApp.getUi().alert(`총 ${totalRefreshCount}개의 캐릭터 정보를 갱신하였습니다.`);
        }
      } catch (e) {
        Logger.log(`error : ${e}`);
        
        if (selLanguage == 'english') {
          SpreadsheetApp.getUi().alert(`An error occurred : ${e}`);
        } else {
          SpreadsheetApp.getUi().alert(`에러가 발생하였습니다 : ${e}`);
        }
      } finally {
        sheet.getRange('setAllRefresh').setValue("FALSE");
      }

    } else if (column == selLanguageCol && row == selLanguageRow && e.value != null && e.value != '') {
      // 언어 변경을 한 경우 직업명을 변환해야함

      var jobList = sheet.getRange(charJobRow, charJobCol, charCount).getValues();
      jobList = jobList.map(function(element) {
        return element[0];
      }).filter(function(element) {
        if (element == '' || element == null) {
          return false;
        } else {
          return true;
        }
      });

      for (var i in jobList) {
        // 공백이면 건너뛰기
        if (jobList[i] == '' || jobList[i] == null) {
          Logger.log(`${i} skipped`);
          continue;
        }

        if (e.value.toLowerCase() == 'korean') {
          let korJob = findEngJobToKor(false, jobList[i]);
          if(korJob == null) {
            Logger.log(`search failed job list : return ${jobList[i]}`)
            korJob = jobList[i];
          } else {
            Logger.log(`${i}/${+charJobRow + +i} - change job name : ${jobList[i]} to ${korJob}`);
          }
          sheet.getRange(+charJobRow + +i, charJobCol).setValue(korJob);
        } if (e.value.toLowerCase() == 'english') {
          let engJob = findEngJobToKor(true, jobList[i]);
          if(engJob == null) {
            Logger.log(`search failed job list : return ${jobList[i]}`)
            engJob = jobList[i];
          } else {
            Logger.log(`${i}/${+charJobRow + +i} - change job name : ${jobList[i]} to ${engJob}`);
          }
          sheet.getRange(+charJobRow + +i, charJobCol).setValue(engJob);
        }
        
      }

    }

    // var charNameRange = sheet.getRange(charNameStartRow, charNameStartCol, charCount);
    // var charNameValues = range.getValues();

  }
}

function gms_autoRefreshLegionList() {
  try {
    var thisTime = new Date();
    var cnt = gms_allRefreshLegionList();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Legion');
    sheet.getRange('autoRefreshTime').setValue(thisTime);

  } catch (e) {
    Logger.log(`error : ${e}`);
    // SpreadsheetApp.getUi().alert(`에러가 발생하였습니다 : ${e}`);
  }
  
}

function gms_allRefreshLegionList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Legion');
  // const isReboot = sheet.getRange('isReboot').getValues()[0][0];
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const charNameStartCol = Number(dataSheet.getRange('charNameCol').getValues());
  const charNameStartRow = Number(dataSheet.getRange('charNameRow').getValues());
  const charCount = Number(dataSheet.getRange('charCount').getValues());
  const charUrlCol = Number(dataSheet.getRange('charUrlCol').getValues());
  const charJobCol = Number(dataSheet.getRange('charJobCol').getValues());
  const charLevelCol = Number(dataSheet.getRange('charLevelCol').getValues());
  const gmsServer = sheet.getRange('gmsServer').getValues()[0][0];
  let selLanguage = sheet.getRange('selLanguage').getValues()[0][0];

  if (selLanguage == null) {
    Logger.log('Language cell is empty. Force change Engish.');
    selLanguage = 'english';
  } else {
    selLanguage = selLanguage.toLowerCase();
  }

  // 레벨에 적힌 모든 메모 지우기
  sheet.getRange(charNameStartRow, charLevelCol, charCount).clearNote();

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


    var charName = typeof existList[i] === 'string' ? existList[i].trim() : existList[i];
    Logger.log(charName);
    // Logger.log(+row + +i);

    // 한글이 포함되어 있는지 검사
    if(charName.match(/[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]/) != null) {
      charName = convertUtf8ToEuckrToANSI(charName);
      Logger.log(`change ANSI : ${charName}`);
      sheet.getRange(+charNameStartRow + +i, charNameStartCol+1).setValue(charName);
    } else {
      // 한글이 없으면 ANSI 셀 지우기
      sheet.getRange(+charNameStartRow + +i, +charUrlCol - 1).setValue('');
    }
    Logger.log(`${i} searching : ${charName}`);
    var charInfo = gms_parseMapleStoryRanking(gmsServer.toLowerCase(), charName);

    let charUrl = 'ok';
    if (charInfo[0] == -1) {
      charUrl = 'error';
    }
    
    sheet.getRange(+charNameStartRow + +i, charUrlCol).setValue(charUrl);
    if (charUrl != 'error') {
      Logger.log(` - ${charInfo}`);

      if (selLanguage == 'korean') {
        let korJob = findEngJobToKor(false, charInfo[1]);
        if(korJob == null) {
          Logger.log(`search failed job list : return ${charInfo[1]}`)
          korJob = charInfo[1];
        } else {
          Logger.log(` - change job name : ${korJob}`);
        }
        sheet.getRange(+charNameStartRow + +i, charJobCol).setValue(korJob);
      } else {
        sheet.getRange(+charNameStartRow + +i, charJobCol).setValue(charInfo[1]);
      }
      sheet.getRange(+charNameStartRow + +i, charNameStartCol -1).setValue("=image(\""+ charInfo[2] + "\")");
      
      // 레벨이 입력된 레벨보다 낮을 경우 수정하지 않고 메모로 표시
      var lastLevel = Number(sheet.getRange(+charNameStartRow + +i, charLevelCol).getValues());
      if (lastLevel > charInfo[0]) {
        sheet.getRange(+charNameStartRow + +i, charLevelCol).setNote(`${selLanguage == 'korean' ? "서버와 값이 다름" : "different from server"} : ${charInfo[0]}`);
      } else {
        sheet.getRange(+charNameStartRow + +i, charLevelCol).setValue(charInfo[0]);
      }


      totalRefreshCount++;
    }

  }

  return totalRefreshCount;
}


function findEngJobToKor(isKor, searchString) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const jobListStartCol = Number(sheet.getRange('jobListStartCol').getValues());
  const jobListStartRow = Number(sheet.getRange('jobListStartRow').getValues());
  const jobCount = Number(sheet.getRange('jobCount').getValues());

  var data = sheet.getRange(jobListStartRow, jobListStartCol, jobCount, 2).getValues();
  
  // Logger.log(`length : ${data.length} / ${jobListStartRow} ${jobListStartCol} jobCount ${jobCount}`);
  for(var i = 0; i < data.length; i++) {
    // Logger.log(`${data[i][0]} ${data[i][1]}`);
    
    if(data[i][isKor ? 1 : 0] == searchString) {
      return data[i][isKor ? 0 : 1]; // F열의 값을 반환합니다.
    }
  }
  
  return null;
}




function gms_onOpen() {

  dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');

  const sheetVersion = dataSheet.getRange('D2').getValues()[0][0];
  

  // if (sheetVersion != '2023.12.26.001') {

  //   var ui = SpreadsheetApp.getUi();
  //   var response = ui.alert('중요 공지사항', 'Open API 적용 관련 공지사항을 확인하세요.\n조회되지 않는 캐릭터의 경우 반드시 접속을 한번 해주셔야 정상 조회가 가능합니다.\nhttps://www.inven.co.kr/board/maple/2304/36025', ui.ButtonSet.OK);

  //   response = ui.alert('중요 공지사항', '공지사항을 확인하였습니까? 작업을 수행하였으면 확인을 눌러 이 메시지를 다시 보지 않습니다.', ui.ButtonSet.YES_NO);

  //   // Process the user's response.
  //   if (response == ui.Button.YES) {
  //     dataSheet.getRange('B2').setValue('2023.12.26.001');
  //   } else {
  //     Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
  //   }
  // }


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
