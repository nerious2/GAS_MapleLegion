function gms_LegionEditV6(e) {
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
    const charLinkCol = Number(dataSheet.getRange('charLinkCol').getValues());
    const charLinkMemoCol = Number(dataSheet.getRange('charLinkMemoCol').getValues());
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
    

    Logger.log(`column ${column} row ${row} lastrow ${lastRow}`);

    // 캐릭터 닉네임 영역을 수정했는지 검사
    if (column == charNameStartCol && row >= charNameStartRow && row <= charNameStartRow + charCount - 1) {
      if (lastRow >= charNameStartRow + charCount - 1) lastRow = charNameStartRow + charCount - 1;

      let charInfoArray = [];
      // job
      let charJobArray = [];
      // level
      let charLevelArray = [];
      // level Memo
      let charLevelMemoArray = [];
      // Link Memo
      let charLinkMemoArray = [];

      var values = sheet.getRange(row, charNameStartCol, lastRow - row + 1).getValues();

      Logger.log(`server : ${gmsServer} / lang : ${selLanguage}`);

      for (var i in values) {
        for (var j in values[i]) {

          // temp 
          let tempCharInfo = [];
          let tempCharJob = [];
          let tempCharLevel = [];
          // let tempCharLevelMemo = [];
          
          var inputValue = typeof values[i][j] === 'string' ? values[i][j].trim() : values[i][j];
          Logger.log(inputValue);
          // Logger.log(+row + +i);
          if (inputValue == '' || inputValue == null) {
            // sheet.getRange(+row + +i, +charUrlCol - 1).setValue('');
            // sheet.getRange(+row + +i, charUrlCol).setValue('');
            // sheet.getRange(+row + +i, +charUrlCol + 1).setValue('');
            // sheet.getRange(+row + +i, charJobCol).setValue('');
            // sheet.getRange(+row + +i, charLevelCol).setValue('');

            tempCharInfo.push('');
            tempCharInfo.push('');
            tempCharInfo.push('');
            tempCharJob.push('');
            tempCharLevel.push('');

            // tempCharLevelMemo.push(null);

          } else {

            // 한글이 포함되어 있는지 검사
            if(inputValue.match(/[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]/) != null) {
              inputValue = convertUtf8ToEuckrToANSI(inputValue);
              Logger.log(`change ANSI : ${inputValue}`);
              // sheet.getRange(+row + +i, charNameStartCol+1).setValue(inputValue);
              tempCharInfo.push(inputValue);
            } else {
              // 한글이 없으면 ANSI 셀 지우기
              // sheet.getRange(+row + +i, +charUrlCol - 1).setValue('');
              tempCharInfo.push('');
            }

            var charInfo = gms_parseMapleStoryRanking(gmsServer.toLowerCase(), inputValue);

            let charUrl = 'ok';
            if (charInfo[0] == -1) {
              charUrl = 'error';
            }
            
            // sheet.getRange(+row + +i, charUrlCol).setValue(charUrl);
            tempCharInfo.push(charUrl);

            if (charUrl != 'error') {
              // char img 넣기
              tempCharInfo.push(charInfo[2]);

              if (selLanguage == 'korean') {
                let korJob = findEngJobToKor(false, charInfo[1]);
                if(korJob == null) {
                  Logger.log(`search failed job list : return ${charInfo[1]}`)
                  korJob = charInfo[1];
                } else {
                  Logger.log(` - change job name : ${korJob}`);
                }
                // sheet.getRange(+row + +i, charJobCol).setValue(korJob);
                tempCharJob.push(korJob);
              } else {
                // sheet.getRange(+row + +i, charJobCol).setValue(charInfo[1]);
                tempCharJob.push(charInfo[1]);
              }
              // sheet.getRange(+row + +i, charLevelCol).setValue(charInfo[0]);
              tempCharLevel.push(charInfo[0]);

              // sheet.getRange(+row + +i, charNameStartCol - 1).setValue("=image(\""+ charInfo[2] + "\")");
            } else {
              // 캐릭터 정보를 가져오지 못한 경우 : 기존 셀에 있는 데이터 넣기
              
              // 캐릭터 이미지
              tempCharInfo.push(sheet.getSheetValues(+row + +i, +charUrlCol + 1, 1, 1)[0][0]);
              tempCharJob.push(sheet.getSheetValues(+row + +i, charJobCol, 1, 1)[0][0]);
              tempCharLevel.push(sheet.getSheetValues(+row + +i, charLevelCol, 1, 1)[0][0]);
              
              // 레벨 메모 넣지 않기
              // tempCharLevelMemo.push(null);
              
            }
          }

          // Logger.log(`add [${inputValue}] info : ${tempCharInfo} // ${tempCharJob} // ${tempCharLevel}`);

          charInfoArray.push(tempCharInfo);
          charJobArray.push(tempCharJob);
          charLevelArray.push(tempCharLevel);
          // charLevelMemoArray.push(tempCharLevelMemo);

        }
      }

      // 최종적으로 시트에 일괄 반영
      // Logger.log(`[${lastRow - row + 1}]length info ${charInfoArray.length} job ${charJobArray.length} level ${charLevelArray.length} levelmemo ${charLevelMemoArray.length}`);

      sheet.getRange(row, +charUrlCol - 1, charInfoArray.length, charInfoArray[0].length).setValues(charInfoArray);
      sheet.getRange(row, charJobCol, charJobArray.length, 1).setValues(charJobArray);
      sheet.getRange(row, charLevelCol, charLevelArray.length, 1).setValues(charLevelArray);
      // sheet.getRange(row, charLevelCol, charLevelMemoArray.length, 1).setNotes(charLevelMemoArray);
      // 레벨 표기 메모 지우기
      sheet.getRange(row, charLevelCol, charLevelArray.length, 1).clearNote();


      // 링크 메모 추가
      // 링크에 적힌 모든 메모 지우기
      sheet.getRange(row, charLinkCol, lastRow - row + 1).clearNote();

      var existLinkMemoList = sheet.getRange(row, charLinkMemoCol, lastRow - row + 1).getValues();
      existLinkMemoList = existLinkMemoList.map(function(element) {
        return element[0];
      });

      // Logger.log("link memo col")
      // Logger.log(charLinkMemoCol);
      // Logger.log(existLinkMemoList.length);


      for (var i in existLinkMemoList) {
        let tempCharLinkMemo = [];
        if (existLinkMemoList[i] == "") {
          tempCharLinkMemo.push(null);
        } else {
          tempCharLinkMemo.push(existLinkMemoList[i]);
        }
        
        charLinkMemoArray.push(tempCharLinkMemo);
      }
      // Logger.log(charLinkMemoArray);

      // 실제로 반영
      sheet.getRange(row, charLinkCol, charLinkMemoArray.length, 1).setNotes(charLinkMemoArray);      


    } else if (column == setAllRefreshCol && row == setAllRefreshRow && e.value == "TRUE") {
      // 전체 새로고침을 체크한 경우
      try {
        // SpreadsheetApp.flush();
        // 유니온 시트에 이미 작성된 캐릭터 리스트
        var totalRefreshCount = gms_allRefreshLegionListV6();
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
      // Link Memo
      let charLinkMemoArray = [];

      
      var jobList = sheet.getRange(charJobRow, charJobCol, charCount).getValues();
      jobList = jobList.map(function(element) {
        return element[0];
      });

      // Logger.log(jobList);
      // Logger.log(jobList.length);

      let jobValueArray = [];

      for (var i in jobList) {
        let tempJobValue = [];
        // 공백이면 건너뛰기
        if (jobList[i] == '' || jobList[i] == null) {
          // Logger.log(`${i} skipped`);
          tempJobValue.push('');
          jobValueArray.push(tempJobValue);
          continue;
        }

        if (e.value.toLowerCase() == 'korean') {
          let korJob = findEngJobToKor(false, jobList[i]);
          if(korJob == null) {
            Logger.log(`search failed job list : return ${jobList[i]}`)
            korJob = jobList[i];
          } else {
            // Logger.log(`${i}/${+charJobRow + +i} - change job name : ${jobList[i]} to ${korJob}`);
          }
          // sheet.getRange(+charJobRow + +i, charJobCol).setValue(korJob);
          tempJobValue.push(korJob);
        } if (e.value.toLowerCase() == 'english') {
          let engJob = findEngJobToKor(true, jobList[i]);
          if(engJob == null) {
            Logger.log(`search failed job list : return ${jobList[i]}`)
            engJob = jobList[i];
          } else {
            // Logger.log(`${i}/${+charJobRow + +i} - change job name : ${jobList[i]} to ${engJob}`);
          }
          // sheet.getRange(+charJobRow + +i, charJobCol).setValue(engJob);
          tempJobValue.push(engJob);
        }
        jobValueArray.push(tempJobValue);
        
      }
      // Logger.log(jobValueArray.length);
      sheet.getRange(charJobRow, charJobCol, charCount, 1).setValues(jobValueArray);

      // 링크 메모 변환
      // 링크에 적힌 모든 메모 지우기
      sheet.getRange(charJobRow, charLinkCol, charCount).clearNote();

      var existLinkMemoList = sheet.getRange(charJobRow, charLinkMemoCol, charCount).getValues();
      existLinkMemoList = existLinkMemoList.map(function(element) {
        return element[0];
      });

      for (var i in existLinkMemoList) {
        let tempCharLinkMemo = [];
        if (existLinkMemoList[i] == "") {
          tempCharLinkMemo.push(null);
        } else {
          tempCharLinkMemo.push(existLinkMemoList[i]);
        }
        
        charLinkMemoArray.push(tempCharLinkMemo);
      }
      // Logger.log(charLinkMemoArray);

      // 실제로 반영
      sheet.getRange(charJobRow, charLinkCol, charLinkMemoArray.length, 1).setNotes(charLinkMemoArray);      

    // ========= 직업명 변환 끝 ===========
    }

    // var charNameRange = sheet.getRange(charNameStartRow, charNameStartCol, charCount);
    // var charNameValues = range.getValues();

  }
  else if (sheetName == 'Boss Crystals') {
    autoInsertBossCrystal(e);
  }
}

function gms_autoRefreshLegionListV6() {
  try {
    var thisTime = new Date();
    var cnt = gms_allRefreshLegionListV6();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Legion');
    sheet.getRange('autoRefreshTime').setValue(thisTime);

  } catch (e) {
    Logger.log(`error : ${e}`);
    // SpreadsheetApp.getUi().alert(`에러가 발생하였습니다 : ${e}`);
  }
  
}

// Legion Counter V6 : 2025.01.24
function gms_allRefreshLegionListV6() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Legion');
  // const isReboot = sheet.getRange('isReboot').getValues()[0][0];
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const charNameStartCol = Number(dataSheet.getRange('charNameCol').getValues());
  const charNameStartRow = Number(dataSheet.getRange('charNameRow').getValues());
  const charCount = Number(dataSheet.getRange('charCount').getValues());
  const charUrlCol = Number(dataSheet.getRange('charUrlCol').getValues());
  const charJobCol = Number(dataSheet.getRange('charJobCol').getValues());
  const charLevelCol = Number(dataSheet.getRange('charLevelCol').getValues());
  const charLinkCol = Number(dataSheet.getRange('charLinkCol').getValues());
  const charLinkMemoCol = Number(dataSheet.getRange('charLinkMemoCol').getValues());
  const gmsServer = sheet.getRange('gmsServer').getValues()[0][0];
  let selLanguage = sheet.getRange('selLanguage').getValues()[0][0];


  // sheet value array
  // ANSI / URL / IMG
  let charInfoArray = [];
  // job
  let charJobArray = [];
  // level
  let charLevelArray = [];
  // level Memo
  let charLevelMemoArray = [];
  // Link Memo
  let charLinkMemoArray = [];

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
  });

  // Logger.log(existList);
  // Logger.log(existList.length);

  var totalRefreshCount = 0;

  for (var i in existList) {
    let tempCharInfo = [];
    let tempCharJob = [];
    let tempCharLevel = [];
    let tempCharLevelMemo = [];
    

    // var charUrl = sheet.getRange(+charNameStartRow + +i, charUrlCol).getValues()[0][0];

    // Logger.log(`${i} : ${charUrl}`);


    var charName = typeof existList[i] === 'string' ? existList[i].trim() : existList[i];

    if (charName == '' || charName == null) {
      // 캐릭터 이름칸이 공백인 경우 : 기존 셀에 들어있는 값을 그대로 넣기
      tempCharInfo = sheet.getSheetValues(+charNameStartRow + +i, +charUrlCol - 1, 1, 3)[0];
      tempCharJob.push(sheet.getSheetValues(+charNameStartRow + +i, charJobCol, 1, 1)[0][0]);
      tempCharLevel.push(sheet.getSheetValues(+charNameStartRow + +i, charLevelCol, 1, 1)[0][0]);
      // Logger.log(`add NULL char info : ${tempCharInfo} // ${tempCharJob} // ${tempCharLevel}`);
      tempCharLevelMemo.push(null);

      charInfoArray.push(tempCharInfo);
      charJobArray.push(tempCharJob);
      charLevelArray.push(tempCharLevel);
      charLevelMemoArray.push(tempCharLevelMemo);
      
      continue;
    }

    // Logger.log(charName);



    // Logger.log(+row + +i);

    // 한글이 포함되어 있는지 검사
    if(charName.match(/[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]/) != null) {
      charName = convertUtf8ToEuckrToANSI(charName);
      Logger.log(`change ANSI : ${charName}`);
      tempCharInfo.push(charName);
    } else {
      // 한글이 없으면 ANSI 셀 지우기
      tempCharInfo.push('');
    }
    Logger.log(`${i} searching : ${charName}`);
    var charInfo = gms_parseMapleStoryRanking(gmsServer.toLowerCase(), charName);

    let charUrl = 'ok';
    if (charInfo[0] == -1) {
      charUrl = 'error';
    }
    tempCharInfo.push(charUrl);
    

    if (charUrl != 'error') {
      // char img 넣기
      tempCharInfo.push(charInfo[2]);
      Logger.log(` - ${charInfo}`);

      if (selLanguage == 'korean') {
        let korJob = findEngJobToKor(false, charInfo[1]);
        if(korJob == null) {
          Logger.log(`search failed job list : return ${charInfo[1]}`)
          korJob = charInfo[1];
        } else {
          Logger.log(` - change job name : ${korJob}`);
        }
        tempCharJob.push(korJob);
      } else {
        tempCharJob.push(charInfo[1]);
      }
      // sheet.getRange(+charNameStartRow + +i, charNameStartCol -1).setValue("=image(\""+ charInfo[2] + "\")");
      
      // 레벨이 입력된 레벨보다 낮을 경우 수정하지 않고 메모로 표시
      var lastLevel = Number(sheet.getRange(+charNameStartRow + +i, charLevelCol).getValues());
      if (lastLevel > charInfo[0]) {
        tempCharLevelMemo.push(`${selLanguage == 'korean' ? "서버와 값이 다름" : "different from server"} : ${charInfo[0]}`);
        tempCharLevel.push(lastLevel);
      } else {
        tempCharLevelMemo.push(null);
        tempCharLevel.push(charInfo[0]);
      }
      
      totalRefreshCount++;
    
    } else {
      // 캐릭터 정보를 가져오지 못한 경우 : 기존 셀에 있는 데이터 넣기
      
      // 캐릭터 이미지
      tempCharInfo.push(sheet.getSheetValues(+charNameStartRow + +i, +charUrlCol + 1, 1, 1)[0][0]);

      tempCharJob.push(sheet.getSheetValues(+charNameStartRow + +i, charJobCol, 1, 1)[0][0]);
      tempCharLevel.push(sheet.getSheetValues(+charNameStartRow + +i, charLevelCol, 1, 1)[0][0]);
      
      // 레벨 메모 넣지 않기
      tempCharLevelMemo.push(null);
      
    }

    // Logger.log(`add [${charName}] info : ${tempCharInfo} // ${tempCharJob} // ${tempCharLevel}`);

    charInfoArray.push(tempCharInfo);
    charJobArray.push(tempCharJob);
    charLevelArray.push(tempCharLevel);
    charLevelMemoArray.push(tempCharLevelMemo);

  }

  // 최종적으로 시트에 일괄 반영
  // Logger.log(`length info ${charInfoArray.length} job ${charJobArray.length} level ${charLevelArray.length} levelmemo ${charLevelMemoArray.length}`);

  sheet.getRange(charNameStartRow, +charUrlCol - 1, charInfoArray.length, charInfoArray[0].length).setValues(charInfoArray);
  sheet.getRange(charNameStartRow, charJobCol, charJobArray.length, 1).setValues(charJobArray);
  sheet.getRange(charNameStartRow, charLevelCol, charLevelArray.length, 1).setValues(charLevelArray);
  sheet.getRange(charNameStartRow, charLevelCol, charLevelMemoArray.length, 1).setNotes(charLevelMemoArray);


  // 링크 메모 추가
  // 링크에 적힌 모든 메모 지우기
  sheet.getRange(charNameStartRow, charLinkCol, charCount).clearNote();

  var existLinkMemoList = sheet.getRange(charNameStartRow, charLinkMemoCol, charCount).getValues();
  existLinkMemoList = existLinkMemoList.map(function(element) {
    return element[0];
  });

  // Logger.log("link memo col")
  // Logger.log(charLinkMemoCol);
  // Logger.log(existLinkMemoList.length);


  for (var i in existLinkMemoList) {
    let tempCharLinkMemo = [];
    if (existLinkMemoList[i] == "") {
      tempCharLinkMemo.push(null);
    } else {
      tempCharLinkMemo.push(existLinkMemoList[i]);
    }
    
    charLinkMemoArray.push(tempCharLinkMemo);
  }


  // 실제로 반영
  sheet.getRange(charNameStartRow, charLinkCol, charLinkMemoArray.length, 1).setNotes(charLinkMemoArray);

  return totalRefreshCount;
}


function autoInsertBossCrystal(e) {
  // Empty
  const Emptyboss = [
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1]
  ];

  // CRA
  const CRAboss = [
    ["Hilla (Hard)", "", 1],
    ["Pink Bean (Chaos)", "", 1],
    ["Cygnus (Normal)", "", 1],
    ["Crimson Queen (Chaos)", "", 1],
    ["Von Bon (Chaos)", "", 1],
    ["Pierre (Chaos)", "", 1],
    ["Zakum (Chaos)", "", 1],
    ["Princess No", "", 1],
    ["Magnus (Hard)", "", 1],
    ["Vellum (Chaos)", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1]
  ];

  // Akechi & CPop
  const CPapboss = [
    ["Hilla (Hard)", "", 1],
    ["Pink Bean (Chaos)", "", 1],
    ["Cygnus (Normal)", "", 1],
    ["Crimson Queen (Chaos)", "", 1],
    ["Von Bon (Chaos)", "", 1],
    ["Pierre (Chaos)", "", 1],
    ["Zakum (Chaos)", "", 1],
    ["Princess No", "", 1],
    ["Magnus (Hard)", "", 1],
    ["Vellum (Chaos)", "", 1],
    ["Papulatus (Chaos)", "", 1],
    ["Akechi Mitsuhide", "", 1],
    ["", "", 1],
    ["", "", 1],
    ["", "", 1]
  ];

  // NLomien
  const NLomienboss = [
    ["Hilla (Hard)", "", 1],
    ["Pink Bean (Chaos)", "", 1],
    ["Cygnus (Normal)", "", 1],
    ["Crimson Queen (Chaos)", "", 1],
    ["Von Bon (Chaos)", "", 1],
    ["Pierre (Chaos)", "", 1],
    ["Zakum (Chaos)", "", 1],
    ["Princess No", "", 1],
    ["Magnus (Hard)", "", 1],
    ["Vellum (Chaos)", "", 1],
    ["Papulatus (Chaos)", "", 1],
    ["Akechi Mitsuhide", "", 1],
    ["Lotus (Normal)", "", 1],
    ["Damien (Normal)", "", 1],
    ["", "", 1]
  ];

  // CTene
  const CTeneboss = [
    ["Zakum (Chaos)", "", 1],
    ["Princess No", "", 1],
    ["Magnus (Hard)", "", 1],
    ["Vellum (Chaos)", "", 1],
    ["Papulatus (Chaos)", "", 1],
    ["Akechi Mitsuhide", "", 1],
    ["Lotus (Hard)", "", 1],
    ["Damien (Hard)", "", 1],
    ["Guardian Angel Slime (Chaos)", "", 1],
    ["Lucid (Hard)", "", 1],
    ["Will (Hard)", "", 1],
    ["Gloom (Chaos)", "", 1],
    ["Darknell (Hard)", "", 1],
    ["Verus Hilla (Hard)", "", 1],
    ["", "", 1]
  ];


  var range = e.range;
  var spreadSheet = e.source;
  var sheet = spreadSheet.getActiveSheet();
  var column = range.getColumn();
  var row = range.getRow();
  var lastRow = range.getLastRow();
  // var inputValue = typeof e.value === 'string' ? e.value.trim() : e.value;

  SpreadsheetApp.flush();
  const bossResetCol = sheet.getRange('bossReset').getColumn();
  const bossResetRow = sheet.getRange('bossReset').getRow();

  const bossResetWithNameCol = sheet.getRange('bossResetWithName').getColumn();
  const bossResetWithNameRow = sheet.getRange('bossResetWithName').getRow();

  const bossCharAutoCol1 = sheet.getRange('bossCharAuto1').getColumn();
  const bossCharAutoCol2 = sheet.getRange('bossCharAuto2').getColumn();
  const bossCharAutoCol3 = sheet.getRange('bossCharAuto3').getColumn();

  const bossCharAutoRow1 = sheet.getRange('bossCharAuto1').getRow();
  const bossCharAutoRow2 = sheet.getRange('bossCharAuto4').getRow();
  const bossCharAutoRow3 = sheet.getRange('bossCharAuto7').getRow();
  const bossCharAutoRow4 = sheet.getRange('bossCharAuto10').getRow();
  const bossCharAutoRow5 = sheet.getRange('bossCharAuto13').getRow();

  // 리셋 체크했을 경우 -> 초기화 진행
  if (column == bossResetWithNameCol && row == bossResetWithNameRow && e.value == "TRUE") {

    for (var i = 1; i < 16; i++) {
      // 캐릭터 이름 초기화
      sheet.getRange('bossCharName' + i.toString()).setValue('');
      // 오토 셀렉트 초기화
      sheet.getRange('bossCharAuto' + i.toString()).setValue('');
      // 일일보스 체크 초기화
      sheet.getRange('bossDailyCheck' + i.toString()).setValue('FALSE');
      // 보스 리스트 초기화
      sheet.getRange('bossList' + i.toString()).setValues(Emptyboss);

    }
    e.range.setValue("FALSE");
  }
  else if (column == bossResetCol && row == bossResetRow && e.value == "TRUE") {
    
    for (var i = 1; i < 16; i++) {
      // 오토 셀렉트 초기화
      sheet.getRange('bossCharAuto' + i.toString()).setValue('');
      // 일일보스 체크 초기화
      sheet.getRange('bossDailyCheck' + i.toString()).setValue('FALSE');
      // 보스 리스트 초기화
      sheet.getRange('bossList' + i.toString()).setValues(Emptyboss);

    }
    e.range.setValue("FALSE");
  }
  else if ((column == bossCharAutoCol1 || column == bossCharAutoCol2 || column == bossCharAutoCol3) &&
   (row == bossCharAutoRow1 || row == bossCharAutoRow2 || row == bossCharAutoRow3 || row == bossCharAutoRow4 || row == bossCharAutoRow5)) {

    // if (column == setAllRefreshCol && row == setAllRefreshRow && e.value == "TRUE") {
    switch (e.value) {
      case "CRA":
        sheet.getRange(+row +3, +column -2, 15, 3).setValues(CRAboss);
        e.range.setValue("");
        break;
      case "Akechi/CPop":
      case "Akechi/CPap":
        sheet.getRange(+row +3, +column -2, 15, 3).setValues(CPapboss);
        e.range.setValue("");
        break;
      case "NLomien":
        sheet.getRange(+row +3, +column -2, 15, 3).setValues(NLomienboss);
        e.range.setValue("");
        break;
      case "CTene":
        sheet.getRange(+row +3, +column -2, 15, 3).setValues(CTeneboss);
        e.range.setValue("");
        break;
      default:
        e.range.setValue("");
        break;
    }
  }
}
