const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, HeadingLevel, BorderStyle } = require('docx');

const app = express();
app.use(express.json({ limit: '15mb' }));
const PORT = process.env.PORT || 3000;

const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
const serviceAuth = new google.auth.GoogleAuth({ credentials, scopes: ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/spreadsheets'] });

const oauthClient = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  'https://developers.google.com/oauthplayground'
);
oauthClient.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });

const SHEET_ID = process.env.SHEET_ID;
const CONFIRMATIONS_FOLDER_ID = process.env.CONFIRMATIONS_FOLDER_ID;
const CONFIRMATIONS_SHEET_NAME = process.env.CONFIRMATIONS_SHEET_NAME || 'אישורי קבלה';
const ORDERS_SHEET_NAME = process.env.ORDERS_SHEET_NAME || 'מעקב הזמנות';
const SERVICE_SHEET_NAME = process.env.SERVICE_SHEET_NAME || 'שירות';

app.get('/', (req, res) => res.send('Server is running ✅'));
app.get('/dashboard', (req, res) => res.sendFile(path.join(__dirname, 'dashboard.html')));
app.get('/confirmation-form', (req, res) => res.sendFile(path.join(__dirname, 'confirmation-form.html')));

function isChecked(val) {
  if (val === true) return true;
  const s = String(val || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'V' || s === 'YES' || s === '1' || s === 'כן';
}

app.get('/api/debug-rows', async (req, res) => {
  try {
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + '!A:G' });
    res.json({ orderSheetName: ORDERS_SHEET_NAME, rows: (resp.data.values || []).slice(0, 5) });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/dashboard-data', async (req, res) => {
  try {
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const ordersResp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + '!A:G' });
    const orderRows = (ordersResp.data.values || []).slice(1);
    let openOrders = 0, sumOpen = 0;
    const costByOrderNum = {};
    orderRows.forEach(row => {
      const orderNum = row[0];
      const cost = parseFloat(String(row[4] || '').replace(/[^0-9.]/g, '')) || 0;
      const gVal = row[6];
      if (orderNum) {
        costByOrderNum[String(orderNum).trim()] = cost;
        if (!isChecked(gVal)) { openOrders++; sumOpen += cost; }
      }
    });
    const serviceResp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: SERVICE_SHEET_NAME + '!A:C' });
    const serviceRows = (serviceResp.data.values || []).slice(1);
    let openService = 0, serviceSum = 0;
    serviceRows.forEach(row => {
      const orderNum = row[0];
      if (orderNum && !isChecked(row[2])) { openService++; serviceSum += costByOrderNum[String(orderNum).trim()] || 0; }
    });
    res.json({ openOrders, sumOpen, openService, serviceSum, updatedAt: new Date().toLocaleString('he-IL') });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

function base64ToBuffer(base64Data) {
  const matches = base64Data.match(/^data:(image\/(\w+));base64,(.+)$/);
  if (!matches) throw new Error('Invalid base64 image');
  return { mimeType: matches[1], ext: matches[2], buffer: Buffer.from(matches[3], 'base64') };
}

function dividerParagraph() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC', space: 1 } },
    spacing: { after: 200 }
  });
}

async function createConfirmationDocx(orderNum, confirmations, photoBase64, signatureBase64, timestamp) {
  const dateStr = new Date(timestamp || Date.now()).toLocaleString('he-IL');

  const photo = base64ToBuffer(photoBase64);
  const sig = base64ToBuffer(signatureBase64);

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'אישור קבלת מוצר', bold: true, size: 40, font: 'Arial' })],
      spacing: { after: 100 }
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: dateStr, size: 20, color: '888888', font: 'Arial' })],
      spacing: { after: 200 }
    }),
    dividerParagraph(),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({ text: 'מספר הזמנה: ', bold: true, size: 26, font: 'Arial' }),
        new TextRun({ text: String(orderNum), size: 26, font: 'Arial' })
      ],
      spacing: { after: 200 }
    }),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: 'הצהרות שאושרו:', bold: true, size: 24, font: 'Arial' })],
      spacing: { after: 100 }
    }),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: '✓  אני מאשר/ת כי קיבלתי את המוצר/ים בהתאם להזמנה שביצעתי. בדקתי את המוצר/ים בעת המסירה ואני מאשר/ת כי הם התקבלו במצב תקין, ללא נזק נראה לעין, ותואמים להזמנה שבוצעה. בחתימתי מטה אני מאשר/ת את קבלת המוצר/ים ואת שביעות רצוני ממצבם בעת המסירה.', size: 22, font: 'Arial', color: '333333' })],
      spacing: { after: 120 }
    }),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: '✓  ידוע לי כי יתרת התשלום בגין המוצר משולמת ישירות לספק, וכי באחריותי לקבל מהספק אסמכתא עבור התשלום.', size: 22, font: 'Arial', color: '333333' })],
      spacing: { after: 80 }
    }),
    dividerParagraph(),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: 'תמונת המוצר:', bold: true, size: 24, font: 'Arial' })],
      spacing: { after: 100 }
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({ data: photo.buffer, transformation: { width: 450, height: 300 }, type: photo.ext === 'jpg' ? 'jpg' : 'png' })],
      spacing: { after: 200 }
    }),
    dividerParagraph(),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: 'חתימת הלקוח:', bold: true, size: 24, font: 'Arial' })],
      spacing: { after: 100 }
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({ data: sig.buffer, transformation: { width: 400, height: 120 }, type: sig.ext === 'jpg' ? 'jpg' : 'png' })],
      spacing: { after: 200 }
    }),
    dividerParagraph(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'מסמך זה נוצר אוטומטית ומהווה אישור קבלת מוצר חתום דיגיטלית', size: 18, color: 'AAAAAA', font: 'Arial' })]
    })
  ];

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 } } },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}

async function uploadFileToDrive(buffer, fileName, mimeType) {
  const { Readable } = require('stream');
  const drive = google.drive({ version: 'v3', auth: oauthClient });
  const file = await drive.files.create({
    requestBody: { name: fileName, parents: [CONFIRMATIONS_FOLDER_ID], mimeType },
    media: { mimeType, body: Readable.from(buffer) },
    fields: 'id, webViewLink'
  });
  await drive.permissions.create({ fileId: file.data.id, requestBody: { role: 'reader', type: 'anyone' } });
  return file.data.webViewLink;
}

async function ensureSheetExists(sheets, sheetId, sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  if (!meta.data.sheets.some(s => s.properties.title === sheetName)) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: sheetId, requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] } });
    await sheets.spreadsheets.values.update({ spreadsheetId: sheetId, range: sheetName + '!A1:D1', valueInputOption: 'RAW', requestBody: { values: [['תאריך ושעה','מספר הזמנה','קישור לאישור','סטטוס']] } });
  }
}

app.post('/confirmation', async (req, res) => {
  try {
    const { orderNum, confirmations, photo, signature, timestamp } = req.body;
    if (!orderNum || !photo || !signature) return res.status(400).json({ error: 'Missing required fields' });
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const docBuffer = await createConfirmationDocx(orderNum, confirmations, photo, signature, timestamp);
    const docLink = await uploadFileToDrive(docBuffer, orderNum + '_אישור_קבלה.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    await ensureSheetExists(sheets, SHEET_ID, CONFIRMATIONS_SHEET_NAME);
    await sheets.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: CONFIRMATIONS_SHEET_NAME + '!A:D', valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [[new Date(timestamp || Date.now()).toLocaleString('he-IL'), orderNum, docLink, 'נשלח']] } });
    res.json({ success: true, pdfLink: docLink });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

app.listen(PORT, () => console.log('Server running on port ' + PORT));
