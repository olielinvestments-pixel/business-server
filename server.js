const express = require('express');
const { google } = require('googleapis');
const path = require('path');

const app = express();
app.use(express.json({ limit: '15mb' })); // allow base64 images

const PORT = process.env.PORT || 3000;

// ---- Google Auth (Service Account) ----
const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
  ]
});

const SHEET_ID = process.env.SHEET_ID;          // הגיליון עם מעקב הזמנות / שירות
const CONFIRMATIONS_FOLDER_ID = process.env.CONFIRMATIONS_FOLDER_ID; // תיקיית דרייב לאישורי קבלה
const CONFIRMATIONS_SHEET_NAME = process.env.CONFIRMATIONS_SHEET_NAME || 'אישורי קבלה';

// ---- Static pages ----
app.get('/', (req, res) => {
  res.send('Server is running ✅');
});

app.get('/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'dashboard.html'));
});

app.get('/confirmation-form', (req, res) => {
  res.sendFile(path.join(__dirname, 'confirmation-form.html'));
});

// ---- Helper: upload a base64 image to Drive ----
async function uploadImageToDrive(drive, base64Data, fileName, folderId) {
  const matches = base64Data.match(/^data:(image\/\w+);base64,(.+)$/);
  if (!matches) throw new Error('Invalid base64 image');
  const mimeType = matches[1];
  const buffer = Buffer.from(matches[2], 'base64');

  const { Readable } = require('stream');
  const stream = Readable.from(buffer);

  const file = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [folderId]
    },
    media: {
      mimeType,
      body: stream
    },
    fields: 'id, webViewLink'
  });

  // make file viewable by anyone with the link
  await drive.permissions.create({
    fileId: file.data.id,
    requestBody: { role: 'reader', type: 'anyone' }
  });

  return file.data.webViewLink;
}

// ---- Helper: ensure a sheet/tab exists, create if not ----
async function ensureSheetExists(sheets, sheetId, sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  const exists = meta.data.sheets.some(s => s.properties.title === sheetName);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: sheetName } } }]
      }
    });
    // header row
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `${sheetName}!A1:F1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [['תאריך ושעה', 'מספר הזמנה', 'אישורים', 'קישור לתמונה', 'קישור לחתימה', 'סטטוס']]
      }
    });
  }
}

// ---- POST /confirmation : receive driver's confirmation form ----
app.post('/confirmation', async (req, res) => {
  try {
    const { orderNum, confirmations, photo, signature, timestamp } = req.body;

    if (!orderNum || !photo || !signature) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const client = await auth.getClient();
    const drive = google.drive({ version: 'v3', auth: client });
    const sheets = google.sheets({ version: 'v4', auth: client });

    const safeTimestamp = (timestamp || new Date().toISOString()).replace(/[:.]/g, '-');
    const baseName = `${orderNum}`;

    // Upload images
    const photoLink = await uploadImageToDrive(drive, photo, `${baseName}_תמונה.jpg`, CONFIRMATIONS_FOLDER_ID);
    const signatureLink = await uploadImageToDrive(drive, signature, `${baseName}_חתימה.png`, CONFIRMATIONS_FOLDER_ID);

    // Write row to sheet
    await ensureSheetExists(sheets, SHEET_ID, CONFIRMATIONS_SHEET_NAME);
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${CONFIRMATIONS_SHEET_NAME}!A:F`,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          new Date(timestamp || Date.now()).toLocaleString('he-IL'),
          orderNum,
          (confirmations || []).join(' | '),
          photoLink,
          signatureLink,
          'נשלח'
        ]]
      }
    });

    res.json({ success: true, photoLink, signatureLink });
  } catch (err) {
    console.error('Error in /confirmation:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
