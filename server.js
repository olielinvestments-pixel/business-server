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

const SHEET_ID = process.env.SHEET_ID;
const CONFIRMATIONS_FOLDER_ID = process.env.CONFIRMATIONS_FOLDER_ID;
const CONFIRMATIONS_SHEET_NAME = process.env.CONFIRMATIONS_SHEET_NAME || 'אישורי קבלה';
const ORDERS_SHEET_NAME = process.env.ORDERS_SHEET_NAME || 'מעקב הזמנות';
const SERVICE_SHEET_NAME = process.env.SERVICE_SHEET_NAME || 'שירות';

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

// ---- GET /api/dashboard-data : live numbers for dashboard ----
function isChecked(val) {
  if (val === true) return true;
  const s = String(val || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'V' || s === 'YES' || s === '1' || s === 'כן';
}

app.get('/api/dashboard-data', async (req, res) => {
  try {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });

    const ordersResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: ORDERS_SHEET_NAME + '!A:F'
    });
    const orderRows = (ordersResp.data.values || []).slice(1);

    let openOrders = 0, sumOpen = 0;
    const costByOrderNum = {};
    orderRows.forEach(row => {
      const orderNum = row[0];
      const cost = parseFloat(row[4]) || 0;
      if (orderNum) {
        costByOrderNum[String(orderNum).trim()] = cost;
        if (!isChecked(row[5])) {
          openOrders++;
          sumOpen += cost;
        }
      }
    });

    const serviceResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: SERVICE_SHEET_NAME + '!A:C'
    });
    const serviceRows = (serviceResp.data.values || []).slice(1);

    let openService = 0, serviceSum = 0;
    serviceRows.forEach(row => {
      const orderNum = row[0];
      if (orderNum && !isChecked(row[2])) {
        openService++;
        serviceSum += costByOrderNum[String(orderNum).trim()] || 0;
      }
    });

    res.json({
      openOrders: openOrders,
      sumOpen: sumOpen,
      openService: openService,
      serviceSum: serviceSum,
      updatedAt: new Date().toLocaleString('he-IL')
    });
  } catch (err) {
    console.error('Error in /api/dashboard-data:', err);
    res.status(500).json({ error: err.message });
  }
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
      mimeType: mimeType,
      body: stream
    },
    fields: 'id, webViewLink'
  });

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
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: sheetName + '!A1:F1',
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
    const baseName = orderNum + '';

    const photoLink = await uploadImageToDrive(drive, photo, baseName + '_תמונה.jpg', CONFIRMATIONS_FOLDER_ID);
    const signatureLink = await uploadImageToDrive(drive, signature, baseName + '_חתימה.png', CONFIRMATIONS_FOLDER_ID);

    await ensureSheetExists(sheets, SHEET_ID, CONFIRMATIONS_SHEET_NAME);
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: CONFIRMATIONS_SHEET_NAME + '!A:F',
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

    res.json({ success: true, photoLink: photoLink, signatureLink: signatureLink });
  } catch (err) {
    console.error('Error in /confirmation:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log('Server running on port ' + PORT);
});
