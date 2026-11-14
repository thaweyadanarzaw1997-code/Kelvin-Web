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
const SHEET_NAME = 'Sheet1'; // Sheet1 တွင် Save လုပ်မည်

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

// 2. Excel Sync (Skip existing Product IDs, Add new ones without prices)
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
    
    // ရှိပြီးသား Product ID နှင့် Product Name များကို စစ်ဆေးရန် Set ထဲထည့်မည်
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
        
        // ဈေးနှုန်းများ မပါဝင်စေဘဲ (Empty string '') အတိုင်း Record လုပ်မည်
        rowsToAppend.push([
          pid,                         // A: Product ID
          pname,                       // B: Product Name
          item.imageUrl || '',         // C: Image
          '',                          // D: 3ml Price
          '',                          // E: 5ml Price
          '',                          // F: 10ml Price
          '',                          // G: 30ml Price
          '',                          // H: 100ml Price
          item.category || 'Unisex',   // I: Category
          item.type || 'Designer',     // J: Type
          item.description || '',      // K: Description
          item.status || 'Active'      // L: Status
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

// 3. Upload Image & Auto-Save Link directly to Sheet1 Image Column (Column C)
app.post('/api/admin/upload-item-image', upload.single('image'), async (req, res) => {
  try {
    const { productId } = req.body;
    if (!req.file || !productId) {
      return res.status(400).json({ error: 'Image file and Product ID are required' });
    }

    // ImgBB သို့မဟုတ် Imgur သို့ ပုံတင်၍ Public Link ရယူခြင်း
    let publicImageUrl = '';
    const base64Image = req.file.buffer.toString('base64');

    if (process.env.IMGBB_API_KEY) {
      const form = new FormData();
      form.append('image', base64Image);
      const imgbbRes = await axios.post(`https://api.imgbb.com/1/upload?key=${process.env.IMGBB_API_KEY}`, form, {
        headers: form.getHeaders()
      });
      publicImageUrl = imgbbRes.data.data.url;
    } else {
      // Fallback to Imgur Anonymous Upload
      const formData = new URLSearchParams();
      formData.append('image', base64Image);
      const imgurRes = await axios.post('https://api.imgur.com/3/image', formData, {
        headers: {
          Authorization: `Client-ID ${process.env.IMGUR_CLIENT_ID || '560297f0a996f0e'}`,
        },
      });
      publicImageUrl = imgurRes.data.data.link;
    }

    // Sheet1 ထဲရှိ သက်ဆိုင်ရာ Product ID ၏ Column C (Image) တွင် Link ကို Auto သွားထည့်ပေးခြင်း
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
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

// 4. Update Prices & Details in Sheet1
app.post('/api/admin/update-item', async (req, res) => {
  try {
    const { productId, productName, imageUrl, price3ml, price5ml, price10ml, price30ml, price100ml, category, type, description, status } = req.body;
    
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

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A${rowIndex}:L${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          productId,
          productName || '',
          imageUrl || '',
          price3ml > 0 ? price3ml : '',
          price5ml > 0 ? price5ml : '',
          price10ml > 0 ? price10ml : '',
          price30ml > 0 ? price30ml : '',
          price100ml > 0 ? price100ml : '',
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
