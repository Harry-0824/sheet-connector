require("dotenv").config();
const express = require("express");
const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static("public"));

// Google Sheets API Setup
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

// Nodemailer Setup
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: (process.env.EMAIL_PASS || "").replace(/\s+/g, ""), // 移除密碼中所有空白
  },
});

/**
 * Helper: 取得第一個工作表名稱
 */
async function getFirstSheetName() {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });
  return meta.data.sheets[0].properties.title;
}

/**
 * Helper: 將 0-based 欄位索引轉為 Excel 欄位字母 (0->A, 1->B, ...)
 */
function indexToColumn(index) {
  let column = "";
  while (index >= 0) {
    column = String.fromCharCode((index % 26) + 65) + column;
    index = Math.floor(index / 26) - 1;
  }
  return column;
}

// Health Check
app.get("/health", (req, res) => {
  res.json({ ok: true });
});

/**
 * GET /api/preview
 * 預覽試算表資料
 */
app.get("/api/preview", async (req, res) => {
  try {
    const sheetName = await getFirstSheetName();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A:Z`,
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return res.json({ headers: [], data: [] });
    }

    const headers = rows[0];
    const data = rows.slice(1).map((row, index) => {
      const rowObj = { _rowIndex: index + 2 }; // 第1列是 header，所以資料從第2列開始
      headers.forEach((header, i) => {
        rowObj[header] = row[i] || "";
      });
      return rowObj;
    });

    res.json({ headers, data });
  } catch (error) {
    console.error("Preview Error:", error);
    res.status(500).json({ error: "無法從 Google 試算表讀取資料" });
  }
});

/**
 * POST /api/execute
 * 執行自動回覆任務
 */
app.post("/api/execute", async (req, res) => {
  try {
    const sheetName = await getFirstSheetName();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A:Z`,
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return res.json({ message: "執行完成", processed: 0 });
    }

    const headers = rows[0];
    console.log("Headers found:", headers);
    const nameIdx = headers.indexOf("姓名");
    const emailIdx = headers.indexOf("Email");
    const statusIdx = headers.indexOf("是否自動回覆");
    console.log(`Indices - Name: ${nameIdx}, Email: ${emailIdx}, Status: ${statusIdx}`);

    if (nameIdx === -1 || emailIdx === -1 || statusIdx === -1) {
      return res
        .status(400)
        .json({ error: "找不到必要的欄位 (姓名, Email, 是否自動回覆)" });
    }

    let processedCount = 0;
    const statusColLetter = indexToColumn(statusIdx);

    // 資料從第2列開始 (index 1)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const rowIndex = i + 1;
      const name = row[nameIdx];
      const email = row[emailIdx];
      const status = row[statusIdx] ? row[statusIdx].trim() : "";

      console.log(`Checking Row ${rowIndex}: name=${name}, email=${email}, status="${status}"`);

      // 只有「是否自動回覆」為空白或為 'n'/'N' 才處理
      if ((status === "" || status.toLowerCase() === "n") && name && email) {
        try {
          // 寄送郵件
          await transporter.sendMail({
            from: process.env.EMAIL_USER,
            to: email,
            subject: "感謝您喜愛我們的產品！",
            text: `Hi ${name},

感謝您喜歡我們的某項產品！我們很高興能為您服務。

Best regards,
Yanwun`,
          });

          // 成功後更新試算表狀態為 'Y'
          console.log(`Updating Sheet: ${sheetName}!${statusColLetter}${rowIndex} to 'Y'`);
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `${sheetName}!${statusColLetter}${rowIndex}`,
            valueInputOption: "RAW",
            resource: {
              values: [["Y"]],
            },
          });

          processedCount++;
        } catch (mailError) {
          console.error(`Row ${rowIndex} fail:`, mailError.message);
          // 忽略單筆錯誤繼續處理下一筆
        }
      }
    }

    res.json({ message: "執行完成", processed: processedCount });
  } catch (error) {
    console.error("Execute Error:", error);
    res.status(500).json({ error: "執行任務失敗" });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
