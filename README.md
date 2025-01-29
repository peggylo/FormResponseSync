# Google Forms Response Auto-Sync

A Google Apps Script project that automatically synchronizes Google Form responses to another spreadsheet with data validation and status tracking.

## Features
- Auto-sync form responses to target spreadsheet
- Data validation for URLs and names
- Visual status tracking with color coding
- Detailed execution logs
- Support both automatic and manual triggers

## Setup
1. Create two Google Spreadsheets (source and target)
2. Open Apps Script editor in source spreadsheet
3. Copy this project's code
4. Replace spreadsheet IDs in the code
5. Set up trigger for form submission

## Usage

### Automatic Sync
- System automatically processes and syncs data when new form responses arrive

### Manual Sync
1. Click "Data Sync Tool" in the source spreadsheet menu
2. Select "Update Unprocessed Data"
3. System will process all unsynced data
4. A summary of results will be displayed upon completion

### Results
- Success: Row background turns light green
- Failure: Row background turns light red
- Detailed status is recorded in the status column

---

# Google 表單回應自動同步

這個 Google Apps Script 專案可以自動將 Google 表單的回應同步到另一個試算表，並進行資料驗證和狀態追蹤。

## 功能特點
- 自動同步表單回應到指定試算表
- 資料驗證功能
- 視覺化狀態追蹤（顏色標記）
- 詳細的執行記錄
- 支援自動和手動觸發

## 設定步驟
1. 建立兩個 Google 試算表（來源和目標）
2. 在來源試算表中開啟 Apps Script 編輯器
3. 複製此專案的程式碼
4. 替換程式碼中的試算表 ID
5. 設定表單提交觸發器

## 使用方式

### 自動同步
- 當有新的表單回應時，系統會自動處理並同步資料

### 手動同步
1. 在來源試算表的選單中點擊「資料同步工具」
2. 選擇「更新未處理資料」
3. 系統會處理所有未同步的資料
4. 完成後會顯示處理結果統計

### 執行結果查看
- 成功：該列背景變為淺綠色
- 失敗：該列背景變為淺紅色
- 詳細狀態會記錄在「處理狀態」欄位

