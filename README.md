# Sheet Connector

整合 Google Sheets API 與 Nodemailer 的全端應用程式。

## 功能

- 透過 Google Sheets API 讀取/寫入試算表資料。
- 整合 Nodemailer 寄送電子郵件。
- 使用 Chart.js 將資料視覺化展示於前端。

## 技術堆疊

- **Backend:** Node.js + Express
- **Frontend:** HTML/CSS/Vanilla JS + Chart.js
- **Integrations:** `googleapis` (Google Sheets API), `nodemailer`
- **Environment:** `dotenv`

## 環境變數 (.env)

請確保您的 `.env` 包含以下變數：

- `EMAIL_USER`: Gmail 帳號（例如：`example@gmail.com`）。
- `EMAIL_PASS`: Gmail 應用程式密碼（會自動移除所有空格）。
- `SPREADSHEET_ID`: Google 試算表的 ID。
- `GOOGLE_APPLICATION_CREDENTIALS`: GCP 服務帳號 JSON 金鑰路徑（預設 `./credentials.json`）。
- `PORT`: 伺服器埠號（預設 `3000`）。

## API 規格

### 1. 健康檢查

- **路徑:** `GET /health`
- **輸出:** `{ "ok": true }`

### 2. 資料預覽

- **路徑:** `GET /api/preview`
- **功能:** 讀取第一個工作表 A:Z 的內容。
- **輸出:**
  ```json
  {
    "headers": ["姓名", "Email", "是否自動回覆", ...],
    "data": [
      { "_rowIndex": 2, "姓名": "張三", "Email": "san@example.com", "是否自動回覆": "" },
      ...
    ]
  }
  ```

### 3. 執行任務

- **路徑:** `POST /api/execute`
- **功能:** 遍歷試算表，若「是否自動回覆」為空白且有姓名與 Email，則寄送郵件並回填「Y」。
- **輸出:** `{ "message": "執行完成", "processed": 5 }`

## 快速開始

### 1. 安裝依賴

```powershell
npm install
```

### 2. 啟動伺服器 (Windows PowerShell)

如果您是在 Windows 環境，可以使用以下腳本快速設定環境變數並啟動（或直接建立 `.env` 檔案）：

```powershell
$env:PORT="3000"; $env:EMAIL_USER="your-email@gmail.com"; $env:EMAIL_PASS="your-app-password"; $env:SPREADSHEET_ID="your-id"; $env:GOOGLE_APPLICATION_CREDENTIALS="./credentials.json"; node server.js
```

或是標準啟動（需先建立 `.env`）：

```bash
node server.js
```

啟動後可在瀏覽器開啟 `http://localhost:3000`。
