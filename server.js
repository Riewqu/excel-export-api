const express = require('express');
const ExcelJS = require('exceljs');
const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const app = express();
const port = 3001;

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

app.get('/export-orders-template', async (req, res) => {
  const userId = req.query.user_id;
  if (!userId) return res.status(400).send('Missing user_id');

  const [
    { data: platforms = [], error: pErr },
    { data: creators = [], error: cErr },
    { data: products = [], error: prErr }
  ] = await Promise.all([
    supabase.from('platforms').select('name').eq('user_id', userId),
    supabase.from('creators').select('name').eq('user_id', userId),
    supabase.from('products').select('name,category,sku,costprice,suggestedPrice').eq('user_id', userId),
  ]);

  console.log({ pErr, cErr, prErr });

  const wb = new ExcelJS.Workbook();

  // 1. Orders Sheet
  const ordersSheet = wb.addWorksheet('Orders');
  ordersSheet.addRow([
    'Order ID', 'วันที่', 'Platform', 'Creator', 'ลูกค้า', 'แคมเปญ', 'สถานะ',
    'ชื่อสินค้า', 'หมวดหมู่', 'รหัสสินค้า', 'จำนวน', 'ต้นทุนต่อชิ้น',
    'ราคาขายต่อชิ้น','ค่าใช้จ่าย ชื่อ', 'ค่าใช้จ่าย จำนวน', 'ค่าใช้จ่าย หน่วย', 'หมายเหตุ'
  ]);

  for (let row = 2; row <= 100; row++) {
    ordersSheet.getCell(`C${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!A2:A100']
    };
    ordersSheet.getCell(`D${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!B2:B100']
    };
    ordersSheet.getCell(`H${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!C2:C100']
    };
    ordersSheet.getCell(`G${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!D2:D100']
    };
    ordersSheet.getCell(`Q${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['=Dictionary!E2:E100']
    };

    // Autofill from ProductData
    ordersSheet.getCell(`I${row}`).value = { formula: `=IFERROR(VLOOKUP(H${row},ProductData!A:E,2,FALSE),"")` }; // หมวดหมู่
    ordersSheet.getCell(`J${row}`).value = { formula: `=IFERROR(VLOOKUP(H${row},ProductData!A:E,3,FALSE),"")` }; // รหัสสินค้า
    ordersSheet.getCell(`L${row}`).value = { formula: `=IFERROR(VLOOKUP(H${row},ProductData!A:E,4,FALSE),"")` }; // ต้นทุน
    ordersSheet.getCell(`M${row}`).value = { formula: `=IFERROR(VLOOKUP(H${row},ProductData!A:E,5,FALSE),"")` }; // ราคาขาย
  }

  // 2. Dictionary Sheet
  const dictSheet = wb.addWorksheet('Dictionary');
  platforms.forEach((p, i) => dictSheet.getCell(`A${i + 2}`).value = p.name);
  creators.forEach((c, i) => dictSheet.getCell(`B${i + 2}`).value = c.name);
  products.forEach((p, i) => dictSheet.getCell(`C${i + 2}`).value = p.name);
  ['เสร็จสิ้น', 'ยกเลิก', 'กำลังดำเนินการ'].forEach((v, i) => dictSheet.getCell(`D${i + 2}`).value = v);
  ['บาท', '%'].forEach((v, i) => dictSheet.getCell(`E${i + 2}`).value = v);

  // 3. ProductData Sheet
  const productSheet = wb.addWorksheet('ProductData');
  productSheet.addRow(['ชื่อสินค้า', 'หมวดหมู่', 'รหัสสินค้า', 'ต้นทุนต่อชิ้น', 'ราคาขายต่อชิ้น']);
  products.forEach(p =>
    productSheet.addRow([p.name, p.category, p.sku, p.costprice, p.suggestedPrice])
  );

  // Response
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=order_template.xlsx');
  await wb.xlsx.write(res);
  res.end();
});

app.listen(port, () => console.log(`✅ Excel download API running at http://localhost:${port}`));
