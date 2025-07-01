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

  if (pErr || cErr || prErr) {
    console.error({ pErr, cErr, prErr });
    return res.status(500).send('Error fetching data from Supabase');
  }

  const wb = new ExcelJS.Workbook();

  // Sheet 1: Orders
  const ordersSheet = wb.addWorksheet('Orders');
  ordersSheet.addRow([
    'Order ID', 'วันที่', 'Platform', 'Creator', 'ลูกค้า', 'แคมเปญ', 'สถานะ',
    'ชื่อสินค้า', 'หมวดหมู่', 'รหัสสินค้า', 'จำนวน', 'ต้นทุนต่อชิ้น',
    'ราคาขายต่อชิ้น', 'ค่าใช้จ่าย ชื่อ', 'ค่าใช้จ่าย จำนวน', 'ค่าใช้จ่าย หน่วย', 'หมายเหตุ'
  ]);

  // สร้าง Map เพื่อ lookup ชื่อสินค้า → รายละเอียดอื่น
  const productMap = {};
  products.forEach(p => {
    productMap[p.name] = p;
  });

  for (let row = 2; row <= 100; row++) {
    // Dropdown
    ordersSheet.getCell(`C${row}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['=Dictionary!$A$2:$A$100']
    };
    ordersSheet.getCell(`D${row}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['=Dictionary!$B$2:$B$100']
    };
    ordersSheet.getCell(`H${row}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['=Dictionary!$C$2:$C$100']
    };
    ordersSheet.getCell(`G${row}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['=Dictionary!$D$2:$D$100']
    };
    ordersSheet.getCell(`P${row}`).dataValidation = {
      type: 'list', allowBlank: true, formulae: ['=Dictionary!$E$2:$E$100']
    };
  
    // ✅ ใส่ VLOOKUP ทุกแถว โดยไม่เช็ค product
    ordersSheet.getCell(`I${row}`).value = {
      formula: `=IFERROR(VLOOKUP(H${row}, ProductData!A:E, 2, FALSE), "")`
    };
    ordersSheet.getCell(`J${row}`).value = {
      formula: `=IFERROR(VLOOKUP(H${row}, ProductData!A:E, 3, FALSE), "")`
    };
    ordersSheet.getCell(`L${row}`).value = {
      formula: `=IFERROR(VLOOKUP(H${row}, ProductData!A:E, 4, FALSE), "")`
    };
    ordersSheet.getCell(`M${row}`).value = {
      formula: `=IFERROR(VLOOKUP(H${row}, ProductData!A:E, 5, FALSE), "")`
    };
  }

  // Sheet 2: Dictionary (สำหรับ dropdown)
  const dictSheet = wb.addWorksheet('Dictionary');
  platforms.forEach((p, i) => dictSheet.getCell(`A${i + 2}`).value = p.name);
  creators.forEach((c, i) => dictSheet.getCell(`B${i + 2}`).value = c.name);
  products.forEach((p, i) => dictSheet.getCell(`C${i + 2}`).value = p.name);
  ['เสร็จสิ้น', 'ยกเลิก', 'กำลังดำเนินการ'].forEach((v, i) => dictSheet.getCell(`D${i + 2}`).value = v);
  ['บาท', '%'].forEach((v, i) => dictSheet.getCell(`E${i + 2}`).value = v);

  // Sheet 3: ProductData (สำหรับดูรายการสินค้า)
  const productSheet = wb.addWorksheet('ProductData');
  productSheet.addRow(['ชื่อสินค้า', 'หมวดหมู่', 'รหัสสินค้า', 'ต้นทุนต่อชิ้น', 'ราคาขายต่อชิ้น']);
  products.forEach(p =>
    productSheet.addRow([p.name, p.category, p.sku, p.costprice, p.suggestedPrice])
  );

  // Sheet 4: ตัวอย่าง
  const example = wb.addWorksheet('ตัวอย่างการกรอก');
  example.addRow([
    'Order ID', 'วันที่', 'Platform', 'Creator', 'ลูกค้า', 'แคมเปญ', 'สถานะ',
    'ชื่อสินค้า', 'หมวดหมู่', 'รหัสสินค้า', 'จำนวน', 'ต้นทุนต่อชิ้น',
    'ราคาขายต่อชิ้น', 'ค่าใช้จ่าย ชื่อ', 'ค่าใช้จ่าย จำนวน', 'ค่าใช้จ่าย หน่วย', 'หมายเหตุ'
  ]);
  example.addRow([
    'ORDER001', '2025-06-30', platforms[0]?.name || '', creators[0]?.name || '', 'ลูกค้า A', 'แคมเปญ X', 'เสร็จสิ้น',
    products[0]?.name || '', products[0]?.category || '', products[0]?.sku || '', 10,
    products[0]?.costprice || '', products[0]?.suggestedPrice || '',
    'ค่าขนส่ง', 50, 'บาท', 'กรอกสินค้าหลายชนิดได้โดยใช้ Order ID เดียวกัน'
  ]);
  example.addRow([
    'ORDER001', '', '', '', '', '', '',
    products[1]?.name || '', products[1]?.category || '', products[1]?.sku || '', 5,
    products[1]?.costprice || '', products[1]?.suggestedPrice || '',
    'ค่าธรรมเนียม', 20, 'บาท', ''
  ]);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=order_template.xlsx');
  await wb.xlsx.write(res);
  res.end();
});

app.listen(port, () => console.log(`✅ Excel download API running at http://localhost:${port}`));
