const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });
const axios = require('axios');
const FormData = require('form-data');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

const getAuthClient = () => {
  return new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : '',
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
};

const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;
const SHEET_NAME = 'Sheet1';

// 1. Get All Items from Sheet1
app.get('/api/items', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:L`,
    });
    const rows = response.data.values;
    if (!rows || rows.length === 0) return res.json({ items: [] });

    const headers = rows[0];
    const items = rows.slice(1).map((row, index) => {
      let obj = { rowIndex: index + 2 };
      headers.forEach((h, i) => {
        const headerName = String(h).trim();
        if (headerName === 'Product ID') obj.productId = row[i] || '';
        if (headerName === 'Product Name') obj.productName = row[i] || '';
        if (headerName === 'Image') obj.imageUrl = row[i] || '';
        if (headerName === '3ml Price') obj.price3ml = row[i] ? Number(String(row[i]).replace(/,/g, '')) : 0;
        if (headerName === '5ml Price') obj.price5ml = row[i] ? Number(String(row[i]).replace(/,/g, '')) : 0;
        if (headerName === '10ml Price') obj.price10ml = row[i] ? Number(String(row[i]).replace(/,/g, '')) : 0;
        if (headerName === '30ml Price') obj.price30ml = row[i] ? Number(String(row[i]).replace(/,/g, '')) : 0;
        if (headerName === '100ml Price') obj.price100ml = row[i] ? Number(String(row[i]).replace(/,/g, '')) : 0;
        if (headerName === 'Category') obj.category = row[i] || '';
        if (headerName === 'Type') obj.type = row[i] || '';
        if (headerName === 'Description') obj.description = row[i] || '';
        if (headerName === 'Status') obj.status = row[i] || 'Active';
      });
      return obj;
    });

    res.json({ items, sysDate: new Date().toISOString().split('T')[0] });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

// 2. Excel Sync (Smart Duplicate Check)
app.post('/api/admin/sync-excel', async (req, res) => {
  try {
    const { newItems } = req.body;
    if (!newItems || !Array.isArray(newItems)) {
      return res.status(400).json({ error: 'Invalid items data' });
    }

    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:B`,
    });
    const rows = response.data.values || [];
    
    const existingKeys = new Set();
    rows.slice(1).forEach(r => {
      if (r[0]) existingKeys.add(String(r[0]).trim().toLowerCase());
      if (r[1]) existingKeys.add(String(r[1]).trim().toLowerCase());
    });

    let addedCount = 0;
    let skippedCount = 0;
    let rowsToAppend = [];

    newItems.forEach(item => {
      const pid = String(item.productId || item.productName || '').trim();
      const pname = String(item.productName || pid).trim();
      if (!pid) return;

      if (existingKeys.has(pid.toLowerCase()) || existingKeys.has(pname.toLowerCase())) {
        skippedCount++;
      } else {
        existingKeys.add(pid.toLowerCase());
        existingKeys.add(pname.toLowerCase());
        addedCount++;
        
        rowsToAppend.push([
          pid,
          pname,
          item.imageUrl || '',
          '', '', '', '', '',
          item.category || 'Unisex',
          item.type || 'Designer',
          item.description || '',
          item.status || 'Active'
        ]);
      }
    });

    if (rowsToAppend.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A:L`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: rowsToAppend },
      });
    }

    res.json({ success: true, addedCount, skippedCount });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

// 3. Upload or Replace Image via ImgBB API & Overwrite Link in Sheet1 Column C
app.post('/api/admin/upload-item-image', upload.single('image'), async (req, res) => {
  try {
    const { productId } = req.body;
    const apiKey = process.env.IMGBB_API_KEY;

    if (!req.file || !productId) {
      return res.status(400).json({ error: 'Image file and Product ID are required' });
    }
    if (!apiKey) {
      return res.status(500).json({ error: 'IMGBB_API_KEY is not configured in Vercel environment variables' });
    }

    const base64Image = req.file.buffer.toString('base64');
    const form = new FormData();
    form.append('image', base64Image);

    // Call ImgBB API to get new public URL
    const imgbbRes = await axios.post(`https://api.imgbb.com/1/upload?key=${apiKey}`, form, {
      headers: form.getHeaders(),
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });

    if (!imgbbRes.data || !imgbbRes.data.success) {
      throw new Error('ImgBB upload rejected');
    }

    const publicImageUrl = imgbbRes.data.data.url;

    // Find Row in Google Sheets
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:A`,
    });

    const rows = response.data.values || [];
    let targetRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] && String(rows[i][0]).trim().toLowerCase() === String(productId).trim().toLowerCase()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({ error: 'Product ID not found in Sheet1' });
    }

    // Overwrite old image link in Column C with new image link
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!C${targetRowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[publicImageUrl]]
      }
    });

    res.json({ success: true, imageUrl: publicImageUrl, productId });
  } catch (error) {
    console.error('ImgBB Upload Error:', error.response?.data || error.message);
    res.status(500).json({ error: error.response?.data?.error?.message || error.message });
  }
});

// 4. Update Prices & Details in Sheet1
app.post('/api/admin/update-item', async (req, res) => {
  try {
    const { productId, price3ml, price5ml, price10ml, price30ml, price100ml, category, type, description, status } = req.body;
    
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:A`,
    });
    const rows = response.data.values || [];
    let rowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] && String(rows[i][0]).trim().toLowerCase() === String(productId).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return res.status(404).json({ error: 'Product ID not found in Sheet1' });
    }

    const currentRowRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${rowIndex}:L${rowIndex}`,
    });
    const currentRow = currentRowRes.data.values ? currentRowRes.data.values[0] : [];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${rowIndex}:L${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          productId,
          currentRow[1] || '',
          currentRow[2] || '', // keep image link as is unless replaced
          price3ml !== undefined && price3ml !== '' ? price3ml : (currentRow[3] || ''),
          price5ml !== undefined && price5ml !== '' ? price5ml : (currentRow[4] || ''),
          price10ml !== undefined && price10ml !== '' ? price10ml : (currentRow[5] || ''),
          price30ml !== undefined && price30ml !== '' ? price30ml : (currentRow[6] || ''),
          price100ml !== undefined && price100ml !== '' ? price100ml : (currentRow[7] || ''),
          category !== undefined ? category : (currentRow[8] || ''),
          type !== undefined ? type : (currentRow[9] || ''),
          description !== undefined ? description : (currentRow[10] || ''),
          status !== undefined ? status : (currentRow[11] || 'Active')
        ]]
      }
    });

    res.json({ success: true, message: 'Item updated successfully' });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = app;
