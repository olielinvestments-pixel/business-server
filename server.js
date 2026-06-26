const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const https = require('https');
const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, BorderStyle } = require('docx');

const app = express();
app.use(express.json({ limit: '50mb' }));
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
const INCOME_SHEET_ID = process.env.INCOME_SHEET_ID;
const INCOME_FOLDER_ID = process.env.INCOME_FOLDER_ID;

const ZOHO_CLIENT_ID = process.env.ZOHO_CLIENT_ID;
const ZOHO_CLIENT_SECRET = process.env.ZOHO_CLIENT_SECRET;
const ZOHO_ORG_ID = process.env.ZOHO_ORG_ID;
let ZOHO_REFRESH_TOKEN = process.env.ZOHO_REFRESH_TOKEN || '';
let zohoAccessToken = '';
let zohoTokenExpiry = 0;

const HEBREW_MONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];

app.get('/', (req, res) => res.send('Server is running ✅'));
app.get('/dashboard', (req, res) => res.sendFile(path.join(__dirname, 'dashboard.html')));
app.get('/confirmation-form', (req, res) => res.sendFile(path.join(__dirname, 'confirmation-form.html')));

// ── Zoho OAuth callback ────────────────────────────────────────────────────────
app.get('/zoho-callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.send('No code received');
  try {
    const result = await zohoPost('https://accounts.zoho.com/oauth/v2/token', new URLSearchParams({
      code, client_id: ZOHO_CLIENT_ID, client_secret: ZOHO_CLIENT_SECRET,
      redirect_uri: 'https://paypal-webhook-81m2.onrender.com/zoho-callback',
      grant_type: 'authorization_code'
    }).toString());
    ZOHO_REFRESH_TOKEN = result.refresh_token;
    res.send(`<h2>✅ Zoho connected!</h2><p>Refresh Token:</p><code style="word-break:break-all">${result.refresh_token}</code><p>Copy this token and add it to Render as <strong>ZOHO_REFRESH_TOKEN</strong></p>`);
  } catch (err) { res.send('Error: ' + err.message); }
});

// ── Zoho Webhook — fires when invoice is approved/paid ────────────────────────
app.post('/zoho-webhook', async (req, res) => {
  try {
    const event = req.body;
    const invoice = event.invoice || event.data && event.data.invoice;
    if (!invoice) return res.json({ received: true });

    const status = invoice.status || '';
    if (status !== 'paid' && status !== 'sent') return res.json({ received: true, skipped: true });

    const invoiceId = invoice.invoice_id;
    const invoiceNum = invoice.invoice_number || '';
    const amount = parseFloat(invoice.total || 0);
    const date = invoice.date || new Date().toLocaleDateString('he-IL');
    const referenceNum = invoice.po_number || invoice.reference_number || '';

    // Download PDF from Zoho
    const pdfBuffer = await downloadZohoPdf(invoiceId);

    // Upload to Drive in correct month folder
    const driveLink = await uploadIncomePdf(pdfBuffer, invoiceNum, date);

    // Add row to income sheet
    await addIncomeSheetRow(date, amount, driveLink, invoiceNum, referenceNum);

    console.log('✅ Invoice processed:', invoiceNum, '| Drive:', driveLink);
    res.json({ success: true });
  } catch (err) {
    console.error('Zoho webhook error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── Zoho helpers ───────────────────────────────────────────────────────────────
function zohoPost(url, body) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const options = { hostname: urlObj.hostname, path: urlObj.pathname, method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) } };
    const req = https.request(options, r => { let data = ''; r.on('data', d => data += d); r.on('end', () => { try { resolve(JSON.parse(data)); } catch(e) { reject(new Error(data)); } }); });
    req.on('error', reject); req.write(body); req.end();
  });
}

async function getZohoAccessToken() {
  if (zohoAccessToken && Date.now() < zohoTokenExpiry - 60000) return zohoAccessToken;
  const result = await zohoPost('https://accounts.zoho.com/oauth/v2/token', new URLSearchParams({
    refresh_token: ZOHO_REFRESH_TOKEN, client_id: ZOHO_CLIENT_ID, client_secret: ZOHO_CLIENT_SECRET, grant_type: 'refresh_token'
  }).toString());
  if (!result.access_token) throw new Error('Failed to get Zoho token: ' + JSON.stringify(result));
  zohoAccessToken = result.access_token;
  zohoTokenExpiry = Date.now() + (result.expires_in || 3600) * 1000;
  return zohoAccessToken;
}

async function zohoApiPost(path, body) {
  const token = await getZohoAccessToken();
  const bodyStr = JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const options = { hostname: 'invoice.zoho.com', path: '/api/v3' + path, method: 'POST', headers: { 'Authorization': 'Zoho-oauthtoken ' + token, 'X-com-zoho-invoice-organizationid': ZOHO_ORG_ID, 'Content-Type': 'application/json;charset=UTF-8', 'Content-Length': Buffer.byteLength(bodyStr) } };
    const req = https.request(options, r => { let data = ''; r.on('data', d => data += d); r.on('end', () => { try { resolve(JSON.parse(data)); } catch(e) { reject(new Error(data)); } }); });
    req.on('error', reject); req.write(bodyStr); req.end();
  });
}

async function zohoApiGet(apiPath) {
  const token = await getZohoAccessToken();
  return new Promise((resolve, reject) => {
    const options = { hostname: 'invoice.zoho.com', path: '/api/v3' + apiPath, method: 'GET', headers: { 'Authorization': 'Zoho-oauthtoken ' + token, 'X-com-zoho-invoice-organizationid': ZOHO_ORG_ID } };
    const req = https.request(options, r => { let data = ''; r.on('data', d => data += d); r.on('end', () => { try { resolve(JSON.parse(data)); } catch(e) { reject(new Error(data)); } }); });
    req.on('error', reject); req.end();
  });
}

async function downloadZohoPdf(invoiceId) {
  const token = await getZohoAccessToken();
  return new Promise((resolve, reject) => {
    const options = { hostname: 'invoice.zoho.com', path: `/api/v3/invoices/${invoiceId}?accept=pdf`, method: 'GET', headers: { 'Authorization': 'Zoho-oauthtoken ' + token, 'X-com-zoho-invoice-organizationid': ZOHO_ORG_ID } };
    const req = https.request(options, r => { const chunks = []; r.on('data', d => chunks.push(d)); r.on('end', () => resolve(Buffer.concat(chunks))); });
    req.on('error', reject); req.end();
  });
}

async function getOrCreateZohoContact(name) {
  const search = await zohoApiGet('/contacts?contact_name=' + encodeURIComponent(name));
  if (search.contacts && search.contacts.length > 0) return search.contacts[0].contact_id;
  const result = await zohoApiPost('/contacts', { contact_name: name, contact_type: 'customer' });
  if (result.contact) return result.contact.contact_id;
  throw new Error('Failed to create contact: ' + JSON.stringify(result));
}

async function createZohoDraftInvoice(orderNum, customerName, itemName, amountUSD) {
  const contactId = await getOrCreateZohoContact(customerName);
  const result = await zohoApiPost('/invoices', { customer_id: contactId, reference_number: String(orderNum), status: 'draft', line_items: [{ name: itemName || 'מוצר', quantity: 1, rate: parseFloat(amountUSD) || 0 }] });
  if (result.invoice) return result.invoice;
  throw new Error('Failed to create invoice: ' + JSON.stringify(result));
}

// ── Drive/Sheets income helpers ────────────────────────────────────────────────
async function getOrCreateMonthFolder(drive, dateStr) {
  const date = dateStr ? new Date(dateStr) : new Date();
  const monthName = HEBREW_MONTHS[date.getMonth()];
  const year = date.getFullYear();
  const folderName = `${monthName} ${year}`;

  const searchRes = await drive.files.list({ q: `'${INCOME_FOLDER_ID}' in parents and name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`, fields: 'files(id)' });
  if (searchRes.data.files.length > 0) return searchRes.data.files[0].id;

  const newFolder = await drive.files.create({ requestBody: { name: folderName, parents: [INCOME_FOLDER_ID], mimeType: 'application/vnd.google-apps.folder' }, fields: 'id' });
  return newFolder.data.id;
}

async function uploadIncomePdf(pdfBuffer, invoiceNum, dateStr) {
  const { Readable } = require('stream');
  const drive = google.drive({ version: 'v3', auth: oauthClient });
  const folderId = await getOrCreateMonthFolder(drive, dateStr);
  const fileName = `Invoice-${invoiceNum}.pdf`;
  const file = await drive.files.create({ requestBody: { name: fileName, parents: [folderId], mimeType: 'application/pdf' }, media: { mimeType: 'application/pdf', body: Readable.from(pdfBuffer) }, fields: 'id, webViewLink' });
  await drive.permissions.create({ fileId: file.data.id, requestBody: { role: 'reader', type: 'anyone' } });
  return file.data.webViewLink;
}

async function getOrCreateMonthSheet(sheets, dateStr) {
  const date = dateStr ? new Date(dateStr) : new Date();
  const monthName = HEBREW_MONTHS[date.getMonth()];
  const year = date.getFullYear();
  const sheetName = `${monthName} ${year}`;

  const meta = await sheets.spreadsheets.get({ spreadsheetId: INCOME_SHEET_ID });
  const existing = meta.data.sheets.find(s => s.properties.title === sheetName);
  if (existing) return sheetName;

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: INCOME_SHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] } });
  await sheets.spreadsheets.values.update({ spreadsheetId: INCOME_SHEET_ID, range: sheetName + '!A1:E1', valueInputOption: 'RAW', requestBody: { values: [['תאריך', 'סכום', 'קישור לחשבונית', '', 'מספר חשבונית']] } });
  return sheetName;
}

async function addIncomeSheetRow(dateStr, amount, driveLink, invoiceNum, referenceNum) {
  const client = await serviceAuth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const sheetName = await getOrCreateMonthSheet(sheets, dateStr);
  await sheets.spreadsheets.values.append({ spreadsheetId: INCOME_SHEET_ID, range: sheetName + '!A:E', valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [[dateStr.replace(/'/g,''), amount, driveLink, invoiceNum]] } });
}

// ── PayPal Webhook ─────────────────────────────────────────────────────────────
app.post('/webhook', async (req, res) => {
  try {
    const event = req.body;
    const eventType = event.event_type || '';
    if (!eventType.includes('PAYMENT') && !eventType.includes('SALE') && !eventType.includes('CAPTURE')) return res.json({ received: true });

    const resource = event.resource || {};
    const amountUSD = parseFloat((resource.seller_receivable_breakdown && resource.seller_receivable_breakdown.net_amount && resource.seller_receivable_breakdown.net_amount.value) || 0);
    const customerName = (resource.payer_name && (resource.payer_name.given_name + ' ' + resource.payer_name.surname)) || (resource.payer && resource.payer.name && (resource.payer.name.given_name + ' ' + resource.payer.name.surname)) || 'לקוח';
    const orderNum = resource.custom_id || resource.invoice_id || resource.id || '';
    const itemName = (resource.purchase_units && resource.purchase_units[0] && resource.purchase_units[0].items && resource.purchase_units[0].items[0] && resource.purchase_units[0].items[0].name) || 'מוצר';

    console.log('PayPal event:', eventType, '| Order:', orderNum, '| Customer:', customerName, '| USD:', amountUSD);

    if (ZOHO_REFRESH_TOKEN && orderNum) {
      try {
        const invoice = await createZohoDraftInvoice(orderNum, customerName, itemName, amountUSD);
        console.log('✅ Zoho draft invoice created:', invoice.invoice_number);
      } catch (zohoErr) { console.error('Zoho error (non-fatal):', zohoErr.message); }
    }

    res.json({ received: true });
  } catch (err) { console.error('Webhook error:', err); res.status(500).json({ error: err.message }); }
});

function isChecked(val) {
  if (val === true) return true;
  const s = String(val || '').trim().toUpperCase();
  return s === 'TRUE' || s === 'V' || s === 'YES' || s === '1' || s === 'כן';
}

app.get('/api/crm-data', async (req, res) => {
  try {
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const [ordersResp, serviceResp] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + '!A:K' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: SERVICE_SHEET_NAME + '!A:E' })
    ]);
    const orders = (ordersResp.data.values || []).slice(1).filter(r => r[0] && String(r[0]).trim()).map(r => ({ orderNum: r[0]||'', orderDate: r[1]||'', invoiceLink: r[2]||'', city: r[3]||'', cost: r[4]||'', immediate: r[5]||'', supplied: r[6]||'', deliveryDate: r[7]||'', notes: r[8]||'', confirmationLink: r[9]||'', zohoInvoiceLink: r[10]||'' }));
    const service = (serviceResp.data.values || []).slice(1).filter(r => r[0] && String(r[0]).trim()).map(r => ({ orderNum: r[0]||'', requestDate: r[1]||'', fixed: r[2]||'', fixDate: r[3]||'', notes: r[4]||'' }));
    res.json({ orders, service });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

app.get('/api/check-duplicate', async (req, res) => {
  try {
    const orderNum = String(req.query.orderNum || '').trim();
    if (!orderNum) return res.json({ isDuplicate: false });
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    // Check if order exists in orders sheet
    const ordersResp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + '!A:A' });
    const orderExists = (ordersResp.data.values || []).some(row => String(row[0] || '').includes(orderNum));
    if (!orderExists) return res.json({ isDuplicate: false, notFound: true });
    // Check if already confirmed
    const confResp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: CONFIRMATIONS_SHEET_NAME + '!B:B' });
    const isDuplicate = (confResp.data.values || []).some(row => String(row[0] || '').trim() === orderNum);
    res.json({ isDuplicate, notFound: false });
  } catch (err) { console.error('check-duplicate error:', err); res.json({ isDuplicate: false }); }
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
      const orderNum = row[0]; const cost = parseFloat(String(row[4]||'').replace(/[^0-9.]/g,''))||0; const gVal = row[6];
      if (orderNum) { costByOrderNum[String(orderNum).trim()] = cost; if (!isChecked(gVal)) { openOrders++; sumOpen += cost; } }
    });
    const serviceResp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: SERVICE_SHEET_NAME + '!A:C' });
    let openService = 0, serviceSum = 0;
    (serviceResp.data.values || []).slice(1).forEach(row => { if (row[0] && !isChecked(row[2])) { openService++; serviceSum += costByOrderNum[String(row[0]).trim()]||0; } });
    res.json({ openOrders, sumOpen, openService, serviceSum, updatedAt: new Date().toLocaleString('he-IL') });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

function base64ToBuffer(base64Data) {
  const matches = base64Data.match(/^data:(image\/(\w+));base64,(.+)$/);
  if (!matches) throw new Error('Invalid base64 image');
  return { mimeType: matches[1], ext: matches[2], buffer: Buffer.from(matches[3], 'base64') };
}

function dividerParagraph() {
  return new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC', space: 1 } }, spacing: { after: 200 } });
}

async function createConfirmationDocx(orderNum, confirmations, photosBase64, signatureBase64, timestamp) {
  const dateStr = new Date(timestamp || Date.now()).toLocaleString('he-IL');
  const sig = base64ToBuffer(signatureBase64);
  const photoElements = [];
  for (const photoBase64 of photosBase64) {
    const photo = base64ToBuffer(photoBase64);
    photoElements.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: photo.buffer, transformation: { width: 450, height: 300 }, type: photo.ext === 'jpg' ? 'jpg' : 'png' })], spacing: { after: 120 } }));
  }
  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'אישור קבלת מוצר', bold: true, size: 40, font: 'Arial' })], spacing: { after: 100 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: dateStr, size: 20, color: '888888', font: 'Arial' })], spacing: { after: 200 } }),
    dividerParagraph(),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'מספר הזמנה: ', bold: true, size: 26, font: 'Arial' }), new TextRun({ text: String(orderNum), size: 26, font: 'Arial' })], spacing: { after: 200 } }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'הצהרות שאושרו:', bold: true, size: 24, font: 'Arial' })], spacing: { after: 100 } }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: '✓ אני מאשר/ת כי קיבלתי את המוצר/ים בהתאם להזמנה שביצעתי. בדקתי את המוצר/ים בעת המסירה ואני מאשר/ת כי הם התקבלו במצב תקין, ללא נזק נראה לעין, ותואמים להזמנה שבוצעה. בחתימתי מטה אני מאשר/ת את קבלת המוצר/ים ואת שביעות רצוני ממצבם בעת המסירה.', size: 22, font: 'Arial', color: '333333' })], spacing: { after: 120 } }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: '✓ ידוע לי כי יתרת התשלום בגין המוצר משולמת ישירות לספק, וכי באחריותי לקבל מהספק אסמכתא עבור התשלום.', size: 22, font: 'Arial', color: '333333' })], spacing: { after: 80 } }),
    dividerParagraph(),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'תמונות המוצר:', bold: true, size: 24, font: 'Arial' })], spacing: { after: 100 } }),
    ...photoElements,
    dividerParagraph(),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'חתימת הלקוח:', bold: true, size: 24, font: 'Arial' })], spacing: { after: 100 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: sig.buffer, transformation: { width: 400, height: 120 }, type: sig.ext === 'jpg' ? 'jpg' : 'png' })], spacing: { after: 200 } }),
    dividerParagraph(),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'מסמך זה נוצר אוטומטית ומהווה אישור קבלת מוצר חתום דיגיטלית', size: 18, color: 'AAAAAA', font: 'Arial' })] })
  ];
  const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 } } }, children }] });
  return await Packer.toBuffer(doc);
}

async function uploadFileToDrive(buffer, fileName, mimeType) {
  const { Readable } = require('stream');
  const drive = google.drive({ version: 'v3', auth: oauthClient });
  const file = await drive.files.create({ requestBody: { name: fileName, parents: [CONFIRMATIONS_FOLDER_ID], mimeType }, media: { mimeType, body: Readable.from(buffer) }, fields: 'id, webViewLink' });
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

async function writeConfirmationLinkToOrders(sheets, orderNum, docLink) {
  try {
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + '!A:A' });
    const rows = resp.data.values || [];
    let targetRow = -1;
    for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]||'').includes(String(orderNum).trim())) { targetRow = i + 1; break; } }
    if (targetRow === -1) { console.log('Order not found:', orderNum); return; }
    await sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + '!J' + targetRow, valueInputOption: 'RAW', requestBody: { values: [[docLink]] } });
  } catch (err) { console.error('Failed to write link:', err.message); }
}

app.post('/confirmation', async (req, res) => {
  try {
    const { orderNum, confirmations, photos, photo, signature, timestamp } = req.body;
    if (!orderNum || !signature) return res.status(400).json({ error: 'Missing required fields' });
    const photosArray = photos && photos.length > 0 ? photos : (photo ? [photo] : []);
    if (photosArray.length === 0) return res.status(400).json({ error: 'Missing photo' });
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const docBuffer = await createConfirmationDocx(orderNum, confirmations, photosArray, signature, timestamp);
    const docLink = await uploadFileToDrive(docBuffer, orderNum + '_אישור_קבלה.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    await ensureSheetExists(sheets, SHEET_ID, CONFIRMATIONS_SHEET_NAME);
    await sheets.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: CONFIRMATIONS_SHEET_NAME + '!A:D', valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [[new Date(timestamp||Date.now()).toLocaleString('he-IL'), orderNum, docLink, 'נשלח']] } });
    await writeConfirmationLinkToOrders(sheets, orderNum, docLink);
    res.json({ success: true, pdfLink: docLink });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

// ── Upload Zoho Invoice to Drive + Sheets ─────────────────────────────────────
app.post('/api/upload-zoho-invoice', async (req, res) => {
  try {
    let { invoiceNumber } = req.body;
    if (!invoiceNumber) return res.status(400).json({ error: 'Missing invoiceNumber' });

    // Auto-add INV- prefix if only numbers provided
    invoiceNumber = invoiceNumber.trim();
    if (/^\d+$/.test(invoiceNumber)) invoiceNumber = 'INV-' + invoiceNumber;


    // Check for duplicate across ALL sheets in income spreadsheet
    try {
      const clientCheck = await serviceAuth.getClient();
      const sheetsCheck = google.sheets({ version: 'v4', auth: clientCheck });
      const meta = await sheetsCheck.spreadsheets.get({ spreadsheetId: INCOME_SHEET_ID });
      for (const sheet of meta.data.sheets) {
        const dupCheck = await sheetsCheck.spreadsheets.values.get({ spreadsheetId: INCOME_SHEET_ID, range: sheet.properties.title + '!D:D' }).catch(() => null);
        if (dupCheck && dupCheck.data.values) {
          const exists = dupCheck.data.values.some(row => String(row[0] || '').trim() === invoiceNumber);
          if (exists) return res.status(400).json({ error: 'חשבונית ' + invoiceNumber + ' כבר קיימת ב-Excel. למחיקה — מחק את השורה ישירות בגיליון.' });
        }
      }
    } catch (dupErr) { console.log('Duplicate check skipped:', dupErr.message); }
    // Find invoice in Zoho — try multiple search methods
    let invoice = null;
    const search1 = await zohoApiGet('/invoices?invoice_number=' + encodeURIComponent(invoiceNumber));
    console.log('Zoho search1:', JSON.stringify(search1).substring(0, 500));
    if (search1.invoices && search1.invoices.length > 0) {
      invoice = search1.invoices[0];
    } else {
      // Try searching all invoices and filter
      const search2 = await zohoApiGet('/invoices?per_page=200&sort_column=created_time&sort_order=D');
      console.log('Zoho search2 count:', search2.invoices ? search2.invoices.length : 0);
      if (search2.invoices) {
        console.log('First invoice numbers:', search2.invoices.slice(0,3).map(i=>i.invoice_number));
        invoice = search2.invoices.find(inv => inv.invoice_number === invoiceNumber);
      }
    }
    if (!invoice) return res.status(404).json({ error: 'חשבונית לא נמצאה ב-Zoho: ' + invoiceNumber });

    const invoiceId = invoice.invoice_id;
    const amount = parseFloat(invoice.total || 0);
    const date = invoice.date || new Date().toLocaleDateString('he-IL');
    const referenceNum = invoice.po_number || invoice.reference_number || '';

    // Download PDF from Zoho
    const pdfBuffer = await downloadZohoPdf(invoiceId);

    // Upload to Drive in correct month folder
    const driveLink = await uploadIncomePdf(pdfBuffer, invoiceNumber, date);

    // Add row to income sheet
    await addIncomeSheetRow(date, amount, driveLink, invoiceNumber, referenceNum);

    console.log('✅ Invoice uploaded:', invoiceNumber, '| Drive:', driveLink);
    // Write invoice link to orders sheet column K
    if (referenceNum) {
      try {
        const client2 = await serviceAuth.getClient();
        const sheets2 = google.sheets({ version: "v4", auth: client2 });
        const resp2 = await sheets2.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + "!A:A" });
        const rows2 = resp2.data.values || [];
        let targetRow2 = -1;
        for (let i = 1; i < rows2.length; i++) {
          if (String(rows2[i][0] || "").includes(String(referenceNum).trim())) { targetRow2 = i + 1; break; }
        }
        if (targetRow2 !== -1) {
          await sheets2.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + "!K" + targetRow2, valueInputOption: "RAW", requestBody: { values: [[driveLink]] } });
          console.log("Invoice link written to orders row", targetRow2);
        }
      } catch (e2) { console.error("Failed to write invoice link:", e2.message); }
    }
    res.json({ success: true, driveLink });
  } catch (err) {
    console.error('upload-zoho-invoice error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log('Server running on port ' + PORT));
