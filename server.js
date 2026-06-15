const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const PDFDocument = require('pdfkit');
const https = require('https');
const fs = require('fs');

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
const FONT_PATH = '/tmp/Alef-Regular.ttf';
const FONT_BOLD_PATH = '/tmp/Alef-Bold.ttf';

function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    if (fs.existsSync(dest)) return resolve();
    const file = fs.createWriteStream(dest);
    https.get(url, res => {
      if (res.statusCode === 302 || res.statusCode === 301) {
        file.close();
        return downloadFile(res.headers.location, dest).then(resolve).catch(reject);
      }
      res.pipe(file);
      file.on('finish', () => file.close(resolve));
    }).on('error', err => { fs.unlink(dest, () => {}); reject(err); });
  });
}

async function ensureFonts() {
  await Promise.all([
    downloadFile('https://fonts.gstatic.com/s/alef/v21/FeVfS0NQpLYgrjJbC5FxxbU.ttf', FONT_PATH),
    downloadFile('https://fonts.gstatic.com/s/alef/v21/FeVQS0NQpLYglo50L5la2bxii28.ttf', FONT_BOLD_PATH)
  ]);
}

ensureFonts().catch(console.error);

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
  const matches = base64Data.match(/^data:(image\/\w+);base64,(.+)$/);
  if (!matches) throw new Error('Invalid base64 image');
  return { mimeType: matches[1], buffer: Buffer.from(matches[2], 'base64') };
}

function reverseHebrew(text) {
  return text.split('').reverse().join('');
}

async function createConfirmationPDF(orderNum, confirmations, photoBase64, signatureBase64, timestamp) {
  await ensureFonts();
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', margin: 50 });
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const hasFont = fs.existsSync(FONT_PATH);
    if (hasFont) {
      doc.registerFont('Alef', FONT_PATH);
      doc.registerFont('Alef-Bold', fs.existsSync(FONT_BOLD_PATH) ? FONT_BOLD_PATH : FONT_PATH);
    }

    const regular = hasFont ? 'Alef' : 'Helvetica';
    const bold = hasFont ? 'Alef-Bold' : 'Helvetica-Bold';
    const dateStr = new Date(timestamp || Date.now()).toLocaleString('he-IL');

    doc.font(bold).fontSize(20).text('אישור קבלת מוצר', { align: 'center' });
    doc.moveDown(0.3);
    doc.font(regular).fontSize(11).fillColor('#888').text(dateStr, { align: 'center' });
    doc.moveDown(0.5);
    doc.moveTo(50, doc.y).lineTo(545, doc.y).strokeColor('#ddd').stroke();
    doc.moveDown(0.5);

    doc.fillColor('#000').font(bold).fontSize(12).text('מספר הזמנה: ' + String(orderNum), { align: 'right' });
    doc.moveDown(0.8);

    doc.font(bold).fontSize(12).text('הצהרות שאושרו:', { align: 'right' });
    doc.moveDown(0.3);
    (confirmations || []).forEach(c => {
      doc.font(regular).fontSize(11).fillColor('#333').text('✓  ' + c, { align: 'right' });
      doc.moveDown(0.3);
    });
    doc.moveDown(0.5);

    doc.moveTo(50, doc.y).lineTo(545, doc.y).strokeColor('#ddd').stroke();
    doc.moveDown(0.8);

    doc.font(bold).fontSize(12).fillColor('#000').text('תמונת המוצר:', { align: 'right' });
    doc.moveDown(0.3);
    try {
      const photo = base64ToBuffer(photoBase64);
      doc.image(photo.buffer, { fit: [495, 280], align: 'center' });
    } catch(e) { doc.font(regular).fontSize(11).fillColor('#888').text('תמונה לא זמינה'); }
    doc.moveDown(0.8);

    doc.moveTo(50, doc.y).lineTo(545, doc.y).strokeColor('#ddd').stroke();
    doc.moveDown(0.8);

    doc.font(bold).fontSize(12).fillColor('#000').text('חתימת הלקוח:', { align: 'right' });
    doc.moveDown(0.3);
    try {
      const sig = base64ToBuffer(signatureBase64);
      doc.image(sig.buffer, { fit: [495, 120], align: 'center' });
    } catch(e) { doc.font(regular).fontSize(11).fillColor('#888').text('חתימה לא זמינה'); }
    doc.moveDown(1);

    doc.moveTo(50, doc.y).lineTo(545, doc.y).strokeColor('#ddd').stroke();
    doc.moveDown(0.5);
    doc.font(regular).fontSize(9).fillColor('#aaa').text('מסמך זה נוצר אוטומטית ומהווה אישור קבלת מוצר חתום דיגיטלית', { align: 'center' });

    doc.end();
  });
}

async function uploadPDFToDrive(pdfBuffer, fileName) {
  const { Readable } = require('stream');
  const drive = google.drive({ version: 'v3', auth: oauthClient });
  const file = await drive.files.create({
    requestBody: { name: fileName, parents: [CONFIRMATIONS_FOLDER_ID], mimeType: 'application/pdf' },
    media: { mimeType: 'application/pdf', body: Readable.from(pdfBuffer) },
    fields: 'id, webViewLink'
  });
  await drive.permissions.create({ fileId: file.data.id, requestBody: { role: 'reader', type: 'anyone' } });
  return file.data.webViewLink;
}

async function ensureSheetExists(sheets, sheetId, sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId });
  if (!meta.data.sheets.some(s => s.properties.title === sheetName)) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: sheetId, requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] } });
    await sheets.spreadsheets.values.update({ spreadsheetId: sheetId, range: sheetName + '!A1:D1', valueInputOption: 'RAW', requestBody: { values: [['תאריך ושעה','מספר הזמנה','קישור לאישור PDF','סטטוס']] } });
  }
}

app.post('/confirmation', async (req, res) => {
  try {
    const { orderNum, confirmations, photo, signature, timestamp } = req.body;
    if (!orderNum || !photo || !signature) return res.status(400).json({ error: 'Missing required fields' });
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const pdfBuffer = await createConfirmationPDF(orderNum, confirmations, photo, signature, timestamp);
    const pdfLink = await uploadPDFToDrive(pdfBuffer, orderNum + '_אישור_קבלה.pdf');
    await ensureSheetExists(sheets, SHEET_ID, CONFIRMATIONS_SHEET_NAME);
    await sheets.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: CONFIRMATIONS_SHEET_NAME + '!A:D', valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [[new Date(timestamp || Date.now()).toLocaleString('he-IL'), orderNum, pdfLink, 'נשלח']] } });
    res.json({ success: true, pdfLink });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

app.listen(PORT, () => console.log('Server running on port ' + PORT));
