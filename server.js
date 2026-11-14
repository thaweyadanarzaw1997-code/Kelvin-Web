const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });
const axios = require('axios');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

const getAuthClient = () => {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_CLIENT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : '',
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
};

const SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID;

// 1. Get All Items
app.get('/api/items', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Fragrance Catalog!A:L',
    });
    const rows = response.data.values;
    if (!rows || rows.length === 0) return res.json({ items: [] });

    const headers = rows[0];
    const items = rows.slice(1).map((row, index) => {
      let obj = { rowIndex: index + 2 };
      headers.forEach((h, i) => {
        if (h === 'Product ID') obj.productId = row[i] || '';
        if (h === 'Product Name') obj.productName = row[i] || '';
        if (h === 'Image') obj.imageUrl = row[i] || '';
        if (h === '3ml Price') obj.price3ml = row[i] ? Number(row[i]) : 0;
        if (h === '5ml Price') obj.price5ml = row[i] ? Number(row[i]) : 0;
        if (h === '10ml Price') obj.price10ml = row[i] ? Number(row[i]) : 0;
        if (h === '30ml Price') obj.price30ml = row[i] ? Number(row[i]) : 0;
        if (h === '100ml Price') obj.price100ml = row[i] ? Number(row[i]) : 0;
        if (h === 'Category') obj.category = row[i] || '';
        if (h === 'Type') obj.type = row[i] || '';
        if (h === 'Description') obj.description = row[i] || '';
        if (h === 'Status') obj.status = row[i] || 'Active';
      });
      return obj;
    });

    res.json({ items, sysDate: new Date().toISOString().split('T')[0] });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

// 2. Upload Image (Imgur API Integration for direct public image link)
app.post('/api/admin/upload-image', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No image file uploaded' });
    
    const formData = new URLSearchParams();
    formData.append('image', req.file.buffer.toString('base64'));

    const imgurRes = await axios.post('https://api.imgur.com/3/image', formData, {
      headers: {
        Authorization: `Client-ID ${process.env.IMGUR_CLIENT_ID || '560297f0a996f0e'}`,
      },
    });

    const imageUrl = imgurRes.data.data.link;
    res.json({ success: true, imageUrl });
  } catch (error) {
    console.error(error);
    if (req.file) {
      const b64 = req.file.buffer.toString('base64');
      const dataUrl = `data:${req.file.mimetype};base64,${b64}`;
      return res.json({ success: true, imageUrl: dataUrl });
    }
    res.status(500).json({ error: error.message });
  }
});

// 3. Excel Sync & Smart Duplicate Check (Skip existing Product IDs, add new ones)
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
      range: 'Fragrance Catalog!A:A',
    });
    const rows = response.data.values || [];
    const existingIds = new Set(rows.flat().map(id => String(id).trim().toLowerCase()));

    let addedCount = 0;
    let skippedCount = 0;
    let rowsToAppend = [];

    newItems.forEach(item => {
      const pid = String(item.productId || '').trim();
      if (!pid) return;

      if (existingIds.has(pid.toLowerCase())) {
        skippedCount++;
      } else {
        existingIds.add(pid.toLowerCase());
        addedCount++;
        rowsToAppend.push([
          pid,
          item.productName || '',
          item.imageUrl || '',
          item.price3ml || '',
          item.price5ml || '',
          item.price10ml || '',
          item.price30ml || '',
          item.price100ml || '',
          item.category || 'Unisex',
          item.type || 'Niche',
          item.description || '',
          item.status || 'Active'
        ]);
      }
    });

    if (rowsToAppend.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: 'Fragrance Catalog!A:L',
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

// 4. Update Single Item (Prices, Details, Image URL)
app.post('/api/admin/update-item', async (req, res) => {
  try {
    const { productId, productName, imageUrl, price3ml, price5ml, price10ml, price30ml, price100ml, category, type, description, status } = req.body;
    
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Fragrance Catalog!A:A',
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
      return res.status(404).json({ error: 'Product ID not found in Sheet' });
    }

    const updateRange = `Fragrance Catalog!A${rowIndex}:L${rowIndex}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: updateRange,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          productId,
          productName || '',
          imageUrl || '',
          price3ml || '',
          price5ml || '',
          price10ml || '',
          price30ml || '',
          price100ml || '',
          category || '',
          type || '',
          description || '',
          status || 'Active'
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
