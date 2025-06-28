const express = require("express");
const ExcelJS = require("exceljs");
const { createClient } = require("@supabase/supabase-js");
require("dotenv").config();

const app = express();
const port = process.env.PORT || 3001;

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

app.get("/export-orders-template", async (req, res) => {
  const userId = String(req.query.user_id || "");
  console.log("🔎 Requested user_id:", userId);
  if (!userId) return res.status(400).send("Missing user_id");

  const [{ data: platforms }, { data: creators }, { data: products }] = await Promise.all([
    supabase.from("platforms").select("name").eq("user_id", userId),
    supabase.from("creators").select("name").eq("user_id", userId),
    supabase.from("products").select("name").eq("user_id", userId),
  ]);

  console.log("✅ platforms:", platforms);
  console.log("✅ creators:", creators);
  console.log("✅ products:", products);

  const wb = new ExcelJS.Workbook();
  const wsOrders = wb.addWorksheet("Orders");
  const wsDict = wb.addWorksheet("Dictionary");

  wsOrders.addRow([
    "Order ID", "วันที่", "Platform", "Creator", "ลูกค้า",
    "แคมเปญ", "สถานะ", "ชื่อสินค้า", "หมวดหมู่", "รหัสสินค้า",
    "จำนวน", "ต้นทุนต่อชิ้น", "ราคาขายต่อชิ้น",
    "ค่าใช้จ่าย ชื่อ", "ค่าใช้จ่าย จำนวน", "ค่าใช้จ่าย หน่วย", "หมายเหตุ"
  ]);

  platforms.forEach((p, i) => wsDict.getCell(`A${i + 2}`).value = p.name);
  creators.forEach((c, i) => wsDict.getCell(`B${i + 2}`).value = c.name);
  products.forEach((p, i) => wsDict.getCell(`C${i + 2}`).value = p.name);
  ["เสร็จสิ้น","ยกเลิก","กำลังดำเนินการ"].forEach((v,i) => wsDict.getCell(`D${i+2}`).value = v);
  ["บาท","%"].forEach((v,i)=> wsDict.getCell(`E${i+2}`).value = v);

  for (let r = 2; r <= 100; r++) {
    wsOrders.getCell(`C${r}`).dataValidation = {type:'list', formulae:['=Dictionary!A2:A100'], allowBlank:true};
    wsOrders.getCell(`D${r}`).dataValidation = {type:'list', formulae:['=Dictionary!B2:B100'], allowBlank:true};
    wsOrders.getCell(`H${r}`).dataValidation = {type:'list', formulae:['=Dictionary!C2:C100'], allowBlank:true};
    wsOrders.getCell(`G${r}`).dataValidation = {type:'list', formulae:['=Dictionary!D2:D100'], allowBlank:true};
    wsOrders.getCell(`Q${r}`).dataValidation = {type:'list', formulae:['=Dictionary!E2:E100'], allowBlank:true};
  }

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=order_template.xlsx");
  await wb.xlsx.write(res);
  res.end();
});

app.listen(port, () => console.log(`✅ Excel download API on port ${port}`));
