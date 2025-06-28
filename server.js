const express = require('express');
const ExcelJS = require('exceljs');
const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const app = express();
const port = 3001;

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

app.get('/export-orders-template', async (req, res) => {
  const userId = req.query.user_id;

  if (!userId) {
    return res.status(400).send('Missing user_id');
  }

  const [
    { data: platforms, error: platformsError },
    { data: creators, error: creatorsError },
    { data: products, error: productsError }
  ] = await Promise.all([
    supabase.from('platforms').select('name').eq('user_id', userId),
    supabase.from('creators').select('name').eq('user_id', userId),
    supabase.from('products').select('name').eq('user_id', userId),
  ]);
  
  console.log("🟡 user_id:", userId);
  console.log("📦 platforms:", platforms, "error:", platformsError);
  console.log("🎨 creators:", creators, "error:", creatorsError);
  console.log("🛒 products:", products, "error:", productsError);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Orders');

  ws.addRow([
    'Order ID', 'วันที่', 'Platform', 'Creator', 'ลูกค้า', 'แคมเปญ', 'สถานะ',
    'ชื่อสินค้า', 'หมวดหมู่', 'รหัสสินค้า', 'จำนวน', 'ต้นทุนต่อชิ้น',
    'ราคาขายต่อชิ้น','ค่าใช้จ่าย ชื่อ', 'ค่าใช้จ่าย จำนวน', 'ค่าใช้จ่าย หน่วย', 'หมายเหตุ'
  ]);

  const addDropdown = (ws, col, values) => {
    const escaped = values.map(v => `"${v.replace(/"/g, '""')}"`); // Escape " เป็น ""
    const validation = {
      type: 'list',
      allowBlank: true,
      formulae: [escaped.join(',')],
      showErrorMessage: true
    };
    for (let row = 2; row <= 100; row++) {
      ws.getCell(`${col}${row}`).dataValidation = validation;
    }
  };
  

  addDropdown(ws, 'C', platforms.map(p => p.name)); // Platform
  addDropdown(ws, 'D', creators.map(c => c.name));  // Creator
  addDropdown(ws, 'H', products.map(p => p.name));  // Product Name
  addDropdown(ws, 'G', ['เสร็จสิ้น', 'ยกเลิก', 'กำลังดำเนินการ']); // Status
  addDropdown(ws, 'Q', ['บาท', '%']); // Cost Unit

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=order_template.xlsx');
  await wb.xlsx.write(res);
  res.end();
});

app.listen(port, () => console.log(`Excel download API running at http://localhost:${port}`));
