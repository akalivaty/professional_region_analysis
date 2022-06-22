/**
 * 主要流程: setSpreadsheetIDs() => main() => getLastSemesterData()
 */

/**
 * 建立 goolge sheet 自製選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("其他功能")
    .addItem("分析前5學期資料", "main")
    .addItem("加入第6學期資料", "addLastSemesterData")
    .addSeparator()
    .addItem("顯示操作說明", "show_sideBar")
    .addToUi();
}

/**
 * Show instruction.
 */
function show_sideBar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("instruction").setTitle("使用說明");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function main() {
  // 將 Excel 轉換為 Google 試算表
  setSpreadsheetIDs();

  // 待分析門數
  // !僅修改這項變數即可
  const threshold = 5;

  // 清除舊資料
  const ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("分析(" + threshold + "門)");
  ss1.getRange(3, 1, ss1.getLastRow() - 2, ss1.getLastColumn()).clear();

  // 確認各專業領域資格
  const arrDic = getInterestedData();
  arrDic.forEach((studentDic) => {
    checkPoint(studentDic, threshold);
  });

  // 確認獲得專業領域資格人數
  let professionalCount = 0;
  for (let i = 0; i < arrDic.length; i++) {
    if (arrDic[i].msPt == 1 || arrDic[i].commutPt == 1 || arrDic[i].infPt == 1) {
      professionalCount += 1;
      ss1.getRange(i + 3, 1, 1, ss1.getLastColumn()).setBackground("#c9daf8");
      continue;
    }
  }
  ss1.getRange(1, 6).setValue(professionalCount);

  // 拆開 dictionary，組合成二維陣列
  let sheetContent = [];
  arrDic.forEach((studentDic) => {
    const rowContent = [
      studentDic.studentID,
      studentDic.name,
      studentDic.ms,
      studentDic.msExp,
      studentDic.msPt,
      studentDic.commut,
      studentDic.commutExp,
      studentDic.commutPt,
      studentDic.inf,
      studentDic.infExp,
      studentDic.infPt,
    ];
    sheetContent.push(rowContent);
  });

  // 寫入表單
  const rows = sheetContent.length;
  logMessage(sheetContent);
  ss1.getRange(3, 1, rows, ss1.getLastColumn()).setValues(sheetContent);

  // 刪除使用過的資料
  clearOldFiles();
}

/**
 * 取得照學號排序後的完整資料，回傳 array[dictionary]
 * @param {Array} interestedData
 * @returns {Array}
 */
function getCompleteData(interestedData) {
  let allStudentDic = [];

  nextRow: for (let row = 0; row < interestedData.length; row++) {
    const rowData = interestedData[row];

    for (let dicIdx = 0; dicIdx < allStudentDic.length; dicIdx++) {
      // 若已存在該學生字典資料
      if (rowData[0] == allStudentDic[dicIdx].studentID) {
        // 判斷是否為專業領域課程
        allCourse.forEach(function (courseType, courseTypeIdx) {
          courseType.forEach((course) => {
            if (rowData[2] == course) {
              addCoursePoint(courseTypeIdx, allStudentDic[dicIdx]);
            }
          });
        });
        continue nextRow;
      }
    }

    // 若不存在該學生字典資料，則建立
    const studentDic = {
      studentID: rowData[0],
      name: rowData[1],
      ms: 0,
      msExp: 0,
      msPt: "",
      commut: 0,
      commutExp: 0,
      commutPt: "",
      inf: 0,
      infExp: 0,
      infPt: "",
    };
    // 判斷是否為專業領域課程
    allCourse.forEach(function (courseType, courseTypeIdx) {
      courseType.forEach((course) => {
        if (rowData[2] == course) {
          addCoursePoint(courseTypeIdx, studentDic);
        }
      });
    });
    allStudentDic.push(studentDic);
  }
  // 照學號排序字典
  allStudentDic.sort(function (dic1, dic2) {
    if (dic1.studentID.slice(-3) > dic2.studentID.slice(-3)) return 1;
    if (dic1.studentID.slice(-3) < dic2.studentID.slice(-3)) return -1;
    return 1;
  });

  return allStudentDic;
}

/**
 * 將感興趣資料撈出
 */
function getInterestedData() {
  const ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SpreadsheetID");
  const spreadsheetIDs = ss3.getDataRange().getValues();
  const targetStudentID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("分析(5門)").getRange(1, 2).getValue();
  let interestedData = [];
  spreadsheetIDs.forEach((spreadsheetID) => {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetID).getActiveSheet(); // 原始資料
    const lastColumn = spreadsheet.getLastColumn();
    const head = spreadsheet.getRange(1, 1, 1, lastColumn).getValues(); // 標頭列
    let targetColumns = [];
    head[0].forEach(function (item, column) {
      if (item == "學號" || item == "中文姓名" || item == "科目中文名稱" || item == "成績") {
        targetColumns.push(column + 1);
      }
    });

    const lastRow = spreadsheet.getLastRow(); // 原始資料列
    const studentIDs = spreadsheet.getRange(2, targetColumns[0], lastRow - 1, 1).getValues();
    const studentNames = spreadsheet.getRange(2, targetColumns[1], lastRow - 1, 1).getValues();
    const courses = spreadsheet.getRange(2, targetColumns[2], lastRow - 1, 1).getValues();
    const scores = spreadsheet.getRange(2, targetColumns[3], lastRow - 1, 1).getValues();

    console.time("complex");
    nextRow: for (let i = 0; i < lastRow - 1; i++) {
      // 學號開頭 U07 符合目標
      if (studentIDs[i][0].slice(0, 3) == targetStudentID) {
        // 成績大於 60
        if (scores[i][0] >= 60) {
          for (let j = 0; j < allCourse.length; j++) {
            for (let k = 0; k < allCourse[j].length; k++) {
              // 若課程名稱在清單中
              if (courses[i][0] == allCourse[j][k]) {
                let element = [studentIDs[i][0], studentNames[i][0], courses[i][0]];
                interestedData.push(element);
                continue nextRow;
              }
            }
          }
        }
      }
    }
    console.timeEnd("complex");
    const interestedDataLength = interestedData.length;
    logMessage(spreadsheetID + " :: 共 " + interestedDataLength + " 筆資料");
  });
  let arrDic = getCompleteData(interestedData);
  return arrDic;
}

/**
 * 確認專業領域資格
 * @param {Dictionary} studentDic
 */
function checkPoint(studentDic, minPoint) {
  if (studentDic.msExp >= 1 && studentDic.ms + studentDic.msExp >= minPoint) {
    studentDic.msPt = 1;
  }
  if (studentDic.commutExp >= 1 && studentDic.commut + studentDic.commutExp >= minPoint) {
    studentDic.commutPt = 1;
  }
  if (studentDic.infExp >= 1 && studentDic.inf + studentDic.infExp >= minPoint) {
    studentDic.infPt = 1;
  }
}

/**
 * 取得科目型態
 * @param {number} idxOfOneDimArr - from allCourse
 * @param {Dictionary} studentDic
 */
function addCoursePoint(idxOfOneDimArr, studentDic) {
  switch (idxOfOneDimArr) {
    case 0:
      studentDic.ms += 1;
      return;
    case 1:
      studentDic.msExp += 1;
      return;
    case 2:
      studentDic.commut += 1;
      return;
    case 3:
      studentDic.commutExp += 1;
      return;
    case 4:
      studentDic.inf += 1;
      return;
    case 5:
      studentDic.infExp += 1;
      return;
    default:
      return -1;
  }
}

/**
 * 將 Excel 轉換為 Goolge 試算表，並取得轉換後的 Google 試算表 ID
 */
function setSpreadsheetIDs() {
  const ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("spreadsheetID");
  const spreadsheetID = ss3.getParent().getId();
  const folderID = DriveApp.getFileById(spreadsheetID).getParents().next().getId();
  const excelFileIDs = getExcelFileIDs(folderID);
  let spreadsheetIDs = [];
  excelFileIDs.forEach((excelID) => {
    toastMessage(DriveApp.getFileById(excelID).getName() + " 轉換中");
    const spreadsheetID = convertExcel2GoogleSheet(excelID);
    spreadsheetIDs.push(spreadsheetID);
    const lastRow = ss3.getLastRow();
    ss3.getRange(lastRow + 1, 1).setValue(spreadsheetID);
  });
  toastMessage("試算表格式轉換完成");
}

/**
 * 取得原始資料夾底下的試算表ID (不包含操控分析檔)
 * @param {string} folderID
 */
function getExcelFileIDs(folderID) {
  const folder = DriveApp.getFolderById(folderID);
  const files = folder.getFiles();
  let excelFileIDs = [];
  while (files.hasNext()) {
    let file = files.next();
    if (file.getName() != "操控分析檔") {
      excelFileIDs.push(file.getId());
    }
  }
  return excelFileIDs;
}

/**
 * 將Excel檔轉換為Google試算表
 * @param {string} fileID
 */
function convertExcel2GoogleSheet(fileID) {
  const file = DriveApp.getFileById(fileID);
  const resourse = {
    title: file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: file.getParents().next().getId() }],
  };
  const spreadsheet = Drive.Files.insert(resourse, file.getBlob());
  return spreadsheet.id;
}

/**
 * 刪除 SpreadsheetID 試算表資料 & 刪除資料夾中檔案
 */
function clearOldFiles() {
  // 刪除SpreadsheetID試算表資料
  const ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SpreadsheetID");
  const ss3LastRow = ss3.getLastRow();
  ss3.deleteRows(1, ss3LastRow);

  // 刪除資料夾中檔案
  const spreadsheetID = ss3.getParent().getId();
  const folderID = DriveApp.getFileById(spreadsheetID).getParents().next().getId();
  const folder = DriveApp.getFolderById(folderID);
  const files = folder.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    if (file.getName() != "操控分析檔") {
      Drive.Files.remove(file.getId());
    }
  }
}

/**
 * 右下角顯示視窗訊息
 * @param {string} message
 */
function toastMessage(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, "目前狀態");
}
