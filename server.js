const express = require('express');
const ExcelJS = require('exceljs');
const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const app = express();
const port = 3001;

// Supabase client
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

// ช่วยกรอกข้อมูลอย่างปลอดภัย (ไม่กรอก null หรือ undefined)
const safeWrite = (sheet, col, list) => {
  list
    .map(x => x?.name)
    .filter(Boolean)
    .forEach((val, i) => {
      sheet.getCell(`${col}${i + 2}`).value = val;
    });
};

app.get('/export-orders-template', async (req, res) => {
  const userId = req.query.user_id;
  if (!userId) return res.status(400).send('Missing user_id');

  const [
    { data: platforms = [] },
    { data: creators = [] },
    { data: products = [] }
  ] = await Promise.all([
    supabase.from('platforms').select('name').eq('user_id', userId),
    supabase.from('creators').select('name').eq('user_id', userId),
    supabase.from('products').select('name').eq('user_id', userId),
  ]);

  const wb = new ExcelJS.Workbook();

  // Sheet 1: Orders
  const ws = wb.addWorksheet('Orders');
  ws.addRow([
    'Order ID', 'วันที่', 'Platform', 'Creator', 'ลูกค้า', 'แคมเปญ', 'สถานะ',
    'ชื่อสินค้า', 'หมวดหมู่', 'รหัสสินค้า', 'จำนวน', 'ต้นทุนต่อชิ้น',
    'ราคาขายต่อชิ้น', 'ค่าใช้จ่าย ชื่อ', 'ค่าใช้จ่าย จำนวน', 'ค่าใช้จ่าย หน่วย', 'หมายเหตุ'
  ]);

  // Sheet 2: Dictionary
  const dict = wb.addWorksheet('Dictionary');
  safeWrite(dict, 'A', platforms);  // Platform names
  safeWrite(dict, 'B', creators);   // Creator names
  safeWrite(dict, 'C', products);   // Product names
  ['เสร็จสิ้น', 'ยกเลิก', 'กำลังดำเนินการ'].forEach((v, i) => dict.getCell(`D${i + 2}`).value = v);
  ['บาท', '%'].forEach((v, i) => dict.getCell(`E${i + 2}`).value = v);

  // Set dropdowns in Orders sheet
  for (let row = 2; row <= 100; row++) {
    ws.getCell(`C${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!A2:A100']
    };
    ws.getCell(`D${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!B2:B100']
    };
    ws.getCell(`H${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!C2:C100']
    };
    ws.getCell(`G${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!D2:D100']
    };
    ws.getCell(`Q${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!E2:E100']
    };
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=order_template.xlsx');
  await wb.xlsx.write(res);
  res.end();
});

app.listen(port, () => {
  console.log(`✅ Excel download API running at http://localhost:${port}`);
});
