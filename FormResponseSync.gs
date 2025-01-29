// 設定兩個試算表的 ID
const SOURCE_SPREADSHEET_ID = '11N3O17AWSOi61ASVbNUzk5bP_3jqwZvYNEFqreMKT_0';
const TARGET_SPREADSHEET_ID = '1Es9UH7JFOipQ0o-tyjFdTGg98ZHKC9y94cZnWoDVRG0';

// 常數定義
const STATUS_COLUMN_NAME = '處理狀態';
const SUCCESS_COLOR = '#E6FFE6';
const ERROR_COLOR = '#FFE6E6';

function updateUnprocessedData() {
  try {
    // 獲取來源試算表
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
                                    .getSheetByName('課前作業');
    
    // 獲取目標試算表
    const targetSpreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const targetSheet = targetSpreadsheet.getSheetByName('LIST');
    
    // 獲取所有資料
    const sourceData = sourceSheet.getDataRange().getValues();
    const headers = sourceData[0];
    
    // 獲取必要欄位的索引
    const nameIndex = headers.indexOf('Name');
    const urlIndex = headers.indexOf('URL');
    const courseIndex = headers.indexOf('Course');
    const statusIndex = headers.indexOf(STATUS_COLUMN_NAME);
    
    // 處理每一列未處理的資料
    let processCount = 0;
    console.log('開始處理未同步的資料...');
    
    for (let i = 1; i < sourceData.length; i++) {
      // 如果狀態欄是空的，才處理
      if (!sourceData[i][statusIndex]) {
        console.log(`處理第 ${i + 1} 列：教師 ${sourceData[i][nameIndex]}`);
        processRow(sourceSheet, targetSheet, sourceData[i], i + 1, 
                  nameIndex, urlIndex, courseIndex, statusIndex);
        processCount++;
      }
    }
    
    console.log(`處理完成，共處理 ${processCount} 筆資料`);
  } catch (error) {
    console.error('執行過程發生錯誤：', error);
    SpreadsheetApp.getUi().alert('執行過程發生錯誤：' + error.message);
  }
}

function processRow(sourceSheet, targetSheet, rowData, rowNum, 
                   nameIndex, urlIndex, courseIndex, statusIndex) {
  const errors = [];
  const timestamp = new Date().toLocaleString('zh-TW');
  const teacherName = rowData[nameIndex];
  
  console.log(`正在處理教師 ${teacherName} 的資料...`);
  
  // 驗證 URL
  if (!validateUrl(rowData[urlIndex])) {
    console.log(`- URL 格式錯誤：${rowData[urlIndex]}`);
    errors.push('URL 格式錯誤 - 需要包含 http:// 或 https://');
  }
  
  // 驗證教師姓名
  const nameValidation = validateTeacherName(teacherName);
  if (!nameValidation.isValid) {
    console.log(`- 教師姓名驗證失敗：${nameValidation.error}`);
    errors.push(nameValidation.error);
  }
  
  // 如果有錯誤
  if (errors.length > 0) {
    console.log(`- 處理失敗：${errors.join(', ')}`);
    markAsError(sourceSheet, rowNum, errors, timestamp);
    return;
  }
  
  // 在目標表格中尋找或新增教師
  try {
    updateTargetSheet(targetSheet, teacherName, 
                     rowData[urlIndex], rowData[courseIndex]);
    console.log(`- 處理成功`);
    markAsSuccess(sourceSheet, rowNum, timestamp);
  } catch (error) {
    console.log(`- 處理失敗：${error.message}`);
    markAsError(sourceSheet, rowNum, [error.message], timestamp);
  }
}

function validateUrl(url) {
  return url.toLowerCase().startsWith('http://') || 
         url.toLowerCase().startsWith('https://');
}

function validateTeacherName(name) {
  if (!name) {
    return { isValid: false, error: '教師姓名不能為空' };
  }
  if (name.length < 2 || name.length > 5) {
    return { isValid: false, error: '教師姓名長度必須在2-5個字之間' };
  }
  if (/[!@#$%^&*(),.?":{}|<>]/.test(name)) {
    return { isValid: false, error: '教師姓名不能包含特殊符號' };
  }
  return { isValid: true };
}

function markAsSuccess(sheet, rowNum, timestamp) {
  const range = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
  range.setBackground(SUCCESS_COLOR);
  sheet.getRange(rowNum, sheet.getLastColumn())
       .setValue(`${timestamp} 同步完成`);
}

function markAsError(sheet, rowNum, errors, timestamp) {
  const range = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
  range.setBackground(ERROR_COLOR);
  const errorMessage = `${timestamp} 同步失敗\n${errors.join('\n')}`;
  sheet.getRange(rowNum, sheet.getLastColumn()).setValue(errorMessage);
}

function updateTargetSheet(targetSheet, teacherName, url, courseName) {
  const data = targetSheet.getDataRange().getValues();
  const teacherCol = data[0].indexOf('教師');
  const urlCol = data[0].indexOf('[寒假]造課分享會簡報');
  const courseCol = data[0].indexOf('課程名稱');
  
  // 尋找教師
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][teacherCol] === teacherName) {
      // 更新現有資料
      targetSheet.getRange(i + 1, urlCol + 1).setValue(url);
      targetSheet.getRange(i + 1, courseCol + 1).setValue(courseName);
      found = true;
      break;
    }
  }
  
  // 如果找不到教師，新增一列
  if (!found) {
    const newRow = Array(data[0].length).fill('');
    newRow[teacherCol] = teacherName;
    newRow[urlCol] = url;
    newRow[courseCol] = courseName;
    targetSheet.appendRow(newRow);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('資料同步工具')
    .addItem('更新未處理資料', 'updateUnprocessedData')
    .addToUi();
}

function createTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('updateUnprocessedData')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
} 