/**
 * 從最後學期的選課覆核單撈取學號&科目名稱
 */
function addLastSemesterData() {
  // 轉換 Excel 寫入 SpreadsheetID 試算表
  setSpreadsheetIDs();

  // 待分析門數
  // !僅修改這項變數即可
  const threshold = 5;

  // 取得原始分析資料
  let ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("分析(" + threshold + "門)");
  let ss1Data = ss1.getDataRange().getValues();
  // 建立學生資料字典
  let studentDicArr = [];
  for (let row = 2; row < ss1Data.length; row++) {
    studentDicArr.push({
      studentID: ss1Data[row][0],
      name: ss1Data[row][1],
      ms: ss1Data[row][2],
      msExp: ss1Data[row][3],
      msPt: ss1Data[row][4],
      commut: ss1Data[row][5],
      commutExp: ss1Data[row][6],
      commutPt: ss1Data[row][7],
      inf: ss1Data[row][8],
      infExp: ss1Data[row][9],
      infPt: ss1Data[row][10],
    });
  }

  let spreadsheetID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SpreadsheetID").getRange(1, 1).getValue();
  let spreadsheet = SpreadsheetApp.openById(spreadsheetID).getActiveSheet();
  // 取得標頭列
  let head = spreadsheet.getRange(1, 1, 1, spreadsheet.getLastColumn()).getValues();
  let insterestedColumn = [];
  let columnCount = 0;
  head[0].forEach(function (element, idx) {
    if (columnCount >= 2) {
      return;
    } else {
      if (element == "stu_id" || element == "sub_name") {
        columnCount += 1;
        insterestedColumn.push(idx + 1);
      }
    }
  });
  let targetStudentID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("分析(5門)").getRange(1, 2).getValue();
  let stu_idArr = spreadsheet.getRange(2, insterestedColumn[0], spreadsheet.getLastRow() - 1, 1).getValues();
  let sub_nameArr = spreadsheet.getRange(2, insterestedColumn[1], spreadsheet.getLastRow() - 1, 1).getValues();

  stu_idArr.forEach(function (studentID, idx) {
    let studentIDSplit = studentID[0].split("");
    let compID = studentIDSplit[0] + studentIDSplit[1] + studentIDSplit[2];
    if (compID == targetStudentID) {
      allCourse.forEach(function (courseType, courseTypeIdx) {
        courseType.forEach((course) => {
          if (sub_nameArr[idx][0] == course) {
            // writeSs2Arr.push([studentID[0], sub_nameArr[idx][0]]);
            for (let studentDicIdx = 0; studentDicIdx < studentDicArr.length; studentDicIdx++) {
              if (studentDicArr[studentDicIdx].studentID == studentID[0]) {
                addCoursePoint(courseTypeIdx, studentDicArr[studentDicIdx]);
                break;
              }
            }
          }
        });
      });
    }
  });

  // 確認專業領域資格
  studentDicArr.forEach((studentDic) => {
    preCheckPoint(studentDic, threshold);
  });

  // 確認預計可獲得專業領域資格之人數
  studentDicArr.forEach(function (studentDic, idx) {
    if (studentDic.msPt == "*") {
      ss1.getRange(idx + 3, 5).setValue("*");
    }
    if (studentDic.commutPt == "*") {
      ss1.getRange(idx + 3, 8).setValue("*");
    }
    if (studentDic.infPt == "*") {
      ss1.getRange(idx + 3, 11).setValue("*");
    }
  });

  // 確認獲得專業領域資格人數
  let professionalCount = 0;
  for (let i = 0; i < studentDicArr.length; i++) {
    if (studentDicArr[i].msPt == "*" || studentDicArr[i].commutPt == "*" || studentDicArr[i].infPt == "*") {
      let backgroundColor = ss1.getRange(i + 3, 1, 1, ss1.getLastColumn()).getBackground();
      if (backgroundColor != "#c9daf8") {
        professionalCount += 1;
        ss1.getRange(i + 3, 1, 1, ss1.getLastColumn()).setBackground("#ead1dc");
        continue;
      }
    }
  }
  ss1.getRange(1, 10).setValue(professionalCount);

  // 刪除使用過的資料
  clearOldFiles();
}

/**
 * 確認專業領域資格
 * @param {Dictionary} studentDic
 */
function preCheckPoint(studentDic, minPoint) {
  if (studentDic.msPt != 1 && studentDic.msExp >= 1 && studentDic.ms + studentDic.msExp >= minPoint) {
    studentDic.msPt = "*";
  }
  if (studentDic.commutPt != 1 && studentDic.commutExp >= 1 && studentDic.commut + studentDic.commutExp >= minPoint) {
    studentDic.commutPt = "*";
  }
  if (studentDic.infPt != 1 && studentDic.infExp >= 1 && studentDic.inf + studentDic.infExp >= minPoint) {
    studentDic.infPt = "*";
  }
}
