const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const https = require('https');
const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, HeadingLevel, BorderStyle } = require('docx');

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

const ZOHO_CLIENT_ID = process.env.ZOHO_CLIENT_ID;
const ZOHO_CLIENT_SECRET = process.env.ZOHO_CLIENT_SECRET;
const ZOHO_ORG_ID = process.env.ZOHO_ORG_ID;
let ZOHO_REFRESH_TOKEN = process.env.ZOHO_REFRESH_TOKEN || '';
let zohoAccessToken = '';
let zohoTokenExpiry = 0;

app.get('/', (req, res) => res.send('Server is running ✅'));
app.get('/dashboard', (req, res) => res.sendFile(path.join(__dirname, 'dashboard.html')));
app.get('/confirmation-form', (req, res) => res.sendFile(path.join(__dirname, 'confirmation-form.html')));

// ── Zoho OAuth callback (one-time setup) ──────────────────────────────────────
app.get('/zoho-callback', async (req, res) => {
  const code = req.query.code;
  if (!code) return res.send('No code received');
  try {
    const result = await zohoPost('https://accounts.zoho.com/oauth/v2/token', new URLSearchParams({
      code,
      client_id: ZOHO_CLIENT_ID,
      client_secret: ZOHO_CLIENT_SECRET,
      redirect_uri: 'https://paypal-webhook-81m2.onrender.com/zoho-callback',
      grant_type: 'authorization_code'
    }).toString());
    ZOHO_REFRESH_TOKEN = result.refresh_token;
    res.send(`
      <h2>✅ Zoho connected!</h2>
      <p>Refresh Token:</p>
      <code style="word-break:break-all">${result.refresh_token}</code>
      <p>Copy this token and add it to Render as <strong>ZOHO_REFRESH_TOKEN</strong></p>
    `);
  } catch (err) {
    res.send('Error: ' + err.message);
  }
});

// ── Zoho helpers ──────────────────────────────────────────────────────────────
function zohoPost(url, body) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
    };
    const req = https.request(options, r => {
      let data = '';
      r.on('data', d => data += d);
      r.on('end', () => { try { resolve(JSON.parse(data)); } catch(e) { reject(new Error(data)); } });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function getZohoAccessToken() {
  if (zohoAccessToken && Date.now() < zohoTokenExpiry - 60000) return zohoAccessToken;
  const result = await zohoPost('https://accounts.zoho.com/oauth/v2/token', new URLSearchParams({
    refresh_token: ZOHO_REFRESH_TOKEN,
    client_id: ZOHO_CLIENT_ID,
    client_secret: ZOHO_CLIENT_SECRET,
    grant_type: 'refresh_token'
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
    const options = {
      hostname: 'invoice.zoho.com',
      path: '/api/v3' + path,
      method: 'POST',
      headers: {
        'Authorization': 'Zoho-oauthtoken ' + token,
        'X-com-zoho-invoice-organizationid': ZOHO_ORG_ID,
        'Content-Type': 'application/json;charset=UTF-8',
        'Content-Length': Buffer.byteLength(bodyStr)
      }
    };
    const req = https.request(options, r => {
      let data = '';
      r.on('data', d => data += d);
      r.on('end', () => { try { resolve(JSON.parse(data)); } catch(e) { reject(new Error(data)); } });
    });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

async function zohoApiGet(path) {
  const token = await getZohoAccessToken();
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'invoice.zoho.com',
      path: '/api/v3' + path,
      method: 'GET',
      headers: {
        'Authorization': 'Zoho-oauthtoken ' + token,
        'X-com-zoho-invoice-organizationid': ZOHO_ORG_ID
      }
    };
    const req = https.request(options, r => {
      let data = '';
      r.on('data', d => data += d);
      r.on('end', () => { try { resolve(JSON.parse(data)); } catch(e) { reject(new Error(data)); } });
    });
    req.on('error', reject);
    req.end();
  });
}

async function getOrCreateZohoContact(name) {
  // Search for existing contact
  const search = await zohoApiGet('/contacts?contact_name=' + encodeURIComponent(name));
  if (search.contacts && search.contacts.length > 0) return search.contacts[0].contact_id;
  // Create new contact
  const result = await zohoApiPost('/contacts', { contact_name: name, contact_type: 'customer' });
  if (result.contact) return result.contact.contact_id;
  throw new Error('Failed to create contact: ' + JSON.stringify(result));
}

async function createZohoDraftInvoice(orderNum, customerName, itemName, amountUSD) {
  const contactId = await getOrCreateZohoContact(customerName);
  const result = await zohoApiPost('/invoices', {
    customer_id: contactId,
    reference_number: String(orderNum),
    status: 'draft',
    line_items: [{
      name: itemName || 'מוצר',
      quantity: 1,
      rate: parseFloat(amountUSD) || 0
    }]
  });
  if (result.invoice) return result.invoice;
  throw new Error('Failed to create invoice: ' + JSON.stringify(result));
}

// ── PayPal Webhook ─────────────────────────────────────────────────────────────
app.post('/webhook', async (req, res) => {
  try {
    const event = req.body;
    const eventType = event.event_type || '';
    if (!eventType.includes('PAYMENT') && !eventType.includes('SALE') && !eventType.includes('CAPTURE')) {
      return res.json({ received: true });
    }

    const resource = event.resource || {};
    const amountILS = parseFloat((resource.amount && resource.amount.total) || resource.gross_amount && resource.gross_amount.value || 0);
    const amountUSD = parseFloat((resource.seller_receivable_breakdown && resource.seller_receivable_breakdown.net_amount && resource.seller_receivable_breakdown.net_amount.value) || 0);
    const customerName = (resource.payer_name && (resource.payer_name.given_name + ' ' + resource.payer_name.surname)) ||
                         (resource.payer && resource.payer.name && (resource.payer.name.given_name + ' ' + resource.payer.name.surname)) || 'לקוח';
    const orderNum = resource.custom_id || resource.invoice_id || resource.id || '';
    const itemName = (resource.purchase_units && resource.purchase_units[0] && resource.purchase_units[0].items && resource.purchase_units[0].items[0] && resource.purchase_units[0].items[0].name) || 'מוצר';

    console.log('PayPal event:', eventType, '| Order:', orderNum, '| Customer:', customerName, '| USD:', amountUSD);

    // Create Zoho Draft Invoice if refresh token is set
    if (ZOHO_REFRESH_TOKEN && orderNum) {
      try {
        const invoice = await createZohoDraftInvoice(orderNum, customerName, itemName, amountUSD);
        console.log('✅ Zoho draft invoice created:', invoice.invoice_number);
      } catch (zohoErr) {
        console.error('Zoho error (non-fatal):', zohoErr.message);
      }
    }

    res.json({ received: true });
  } catch (err) {
    console.error('Webhook error:', err);
    res.status(500).json({ error: err.message });
  }
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
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ORDERS_SHEET_NAME + '!A:J' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: SERVICE_SHEET_NAME + '!A:E' })
    ]);
    const orderRows = (ordersResp.data.values || []).slice(1);
    const serviceRows = (serviceResp.data.values || []).slice(1);
    const orders = orderRows
      .filter(r => r[0] && String(r[0]).trim())
      .map(r => ({
        orderNum: r[0] || '',
        orderDate: r[1] || '',
        invoiceLink: r[2] || '',
        city: r[3] || '',
        cost: r[4] || '',
        immediate: r[5] || '',
        supplied: r[6] || '',
        deliveryDate: r[7] || '',
        notes: r[8] || '',
        confirmationLink: r[9] || ''
      }));
    const service = serviceRows
      .filter(r => r[0] && String(r[0]).trim())
      .map(r => ({
        orderNum: r[0] || '',
        requestDate: r[1] || '',
        fixed: r[2] || '',
        fixDate: r[3] || '',
        notes: r[4] || ''
      }));
    res.json({ orders, service });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

app.get('/api/check-duplicate', async (req, res) => {
  try {
    const orderNum = String(req.query.orderNum || '').trim();
    if (!orderNum) return res.json({ isDuplicate: false });
    const client = await serviceAuth.getClient();
    const sheets = google.sheets({ version: 'v4', auth: client });
    const resp = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: CONFIRMATIONS_SHEET_NAME + '!B:B' });
    const rows = resp.data.values || [];
    const exists = rows.some(row => String(row[0] || '').trim() === orderNum);
    res.json({ isDuplicate: exists });
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
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim() === String(orderNum).trim()) { targetRow = i + 1; break; }
    }
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
    await sheets.spreadsheets.values.append({ spreadsheetId: SHEET_ID, range: CONFIRMATIONS_SHEET_NAME + '!A:D', valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS', requestBody: { values: [[new Date(timestamp || Date.now()).toLocaleString('he-IL'), orderNum, docLink, 'נשלח']] } });
    await writeConfirmationLinkToOrders(sheets, orderNum, docLink);
    res.json({ success: true, pdfLink: docLink });
  } catch (err) { console.error(err); res.status(500).json({ error: err.message }); }
});

app.listen(PORT, () => console.log('Server running on port ' + PORT));
