const express = require("express")
const ExcelJS = require("exceljs")
const { createClient } = require("@supabase/supabase-js")
require("dotenv").config()

const app = express()
const port = process.env.PORT || 3001

// Enable CORS for all routes
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*")
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept")
  next()
})

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY)

app.get("/export-orders-template", async (req, res) => {
  const userId = req.query.user_id
  if (!userId) return res.status(400).send("Missing user_id")

  try {
    const [
      { data: platforms = [], error: pErr },
      { data: creators = [], error: cErr },
      { data: products = [], error: prErr },
    ] = await Promise.all([
      supabase.from("platforms").select("name").eq("user_id", userId),
      supabase.from("creators").select("name, commission_rate").eq("user_id", userId),
      supabase
        .from("products")
        .select("name,category,sku,costprice,suggestedPrice,commissionRate")
        .eq("user_id", userId),
    ])

    if (pErr || cErr || prErr) {
      console.error({ pErr, cErr, prErr })
      return res.status(500).send("Error fetching data from Supabase")
    }

    const wb = new ExcelJS.Workbook()

    // Sheet 1: Orders (เพิ่มคอลัมน์ประเภทค่าคอม)
    const ordersSheet = wb.addWorksheet("Orders")
    ordersSheet.addRow([
      "Order ID",
      "วันที่",
      "Platform",
      "Creator",
      "ประเภทค่าคอม",
      "ลูกค้า",
      "แคมเปญ",
      "สถานะ",
      "ชื่อสินค้า",
      "หมวดหมู่",
      "รหัสสินค้า",
      "จำนวน",
      "ต้นทุนต่อชิ้น",
      "ราคาขายต่อชิ้���",
      "ค่าใช้จ่าย ชื่อ",
      "ค่าใช้จ่าย จำนวน",
      "ค่าใช้จ่าย หน่วย",
      "หมายเหตุ",
    ])

    // สร้าง Map เพื่อ lookup ข้อมูล
    const productMap = {}
    products.forEach((p) => {
      productMap[p.name] = p
    })

    const creatorMap = {}
    creators.forEach((c) => {
      creatorMap[c.name] = c
    })

    for (let row = 2; row <= 100; row++) {
      // Dropdown validations
      ordersSheet.getCell(`C${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ["=Dictionary!$A$2:$A$100"],
      }
      ordersSheet.getCell(`D${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ["=Dictionary!$B$2:$B$100"],
      }
      // เพิ่ม dropdown สำหรับประเภทค่าคอม
      ordersSheet.getCell(`E${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ["=Dictionary!$F$2:$F$100"],
      }
      ordersSheet.getCell(`I${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ["=Dictionary!$C$2:$C$100"],
      }
      ordersSheet.getCell(`H${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ["=Dictionary!$D$2:$D$100"],
      }
      ordersSheet.getCell(`Q${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: ["=Dictionary!$E$2:$E$100"],
      }

      // VLOOKUP formulas (ปรับตำแหน่งคอลัมน์)
      ordersSheet.getCell(`J${row}`).value = {
        formula: `=IFERROR(VLOOKUP(I${row}, ProductData!A:F, 2, FALSE), "")`,
      }
      ordersSheet.getCell(`K${row}`).value = {
        formula: `=IFERROR(VLOOKUP(I${row}, ProductData!A:F, 3, FALSE), "")`,
      }
      ordersSheet.getCell(`M${row}`).value = {
        formula: `=IFERROR(VLOOKUP(I${row}, ProductData!A:F, 4, FALSE), "")`,
      }
      ordersSheet.getCell(`N${row}`).value = {
        formula: `=IFERROR(VLOOKUP(I${row}, ProductData!A:F, 5, FALSE), "")`,
      }

      // เพิ่ม default value สำหรับประเภทค่าคอม
      if (row === 2) {
        ordersSheet.getCell(`E${row}`).value = "จากสินค้า"
      }
    }

    // Sheet 2: Dictionary (เพิ่มประเภทค่าคอม)
    const dictSheet = wb.addWorksheet("Dictionary")
    platforms.forEach((p, i) => (dictSheet.getCell(`A${i + 2}`).value = p.name))
    creators.forEach((c, i) => (dictSheet.getCell(`B${i + 2}`).value = c.name))
    products.forEach((p, i) => (dictSheet.getCell(`C${i + 2}`).value = p.name))
    ;["เสร็จสิ้น", "ยกเลิก", "กำลังดำเนินการ"].forEach((v, i) => (dictSheet.getCell(`D${i + 2}`).value = v))
    ;["บาท", "%"].forEach((v, i) => (dictSheet.getCell(`E${i + 2}`).value = v))
    // เพิ่มประเภทค่าคอม
    ;["จากสินค้า", "จากครีเอเตอร์"].forEach((v, i) => (dictSheet.getCell(`F${i + 2}`).value = v))

    // Sheet 3: ProductData (เพิ่มคอลัมน์ค่าคอม)
    const productSheet = wb.addWorksheet("ProductData")
    productSheet.addRow(["ชื่อสินค้า", "หมวดหมู่", "รหัสสินค้า", "ต้นทุนต่อชิ้น", "ราคาขายต่อชิ้น", "ค่าคอมสินค้า(%)"])
    products.forEach((p) =>
      productSheet.addRow([p.name, p.category, p.sku, p.costprice, p.suggestedPrice, p.commissionRate || 0]),
    )

    // Sheet 4: CreatorData (ข้อมูลครีเอเตอร์)
    const creatorSheet = wb.addWorksheet("CreatorData")
    creatorSheet.addRow(["ชื่อครีเอเตอร์", "ค่าคอมครีเอเตอร์(%)"])
    creators.forEach((c) => creatorSheet.addRow([c.name, c.commission_rate || 0]))

    // Sheet 5: ตัวอย่าง (อัปเดตตัวอย่าง)
    const example = wb.addWorksheet("ตัวอย่างการกรอก")
    example.addRow([
      "Order ID",
      "วันที่",
      "Platform",
      "Creator",
      "ประเภทค่าคอม",
      "ลูกค้า",
      "แคมเปญ",
      "สถานะ",
      "ชื่อสินค้า",
      "หมวดหมู่",
      "รหัสสินค้า",
      "จำนวน",
      "ต้นทุนต่อชิ้น",
      "ราคาขายต่อชิ้น",
      "ค่าใช้จ่าย ชื่อ",
      "ค่าใช้จ่าย จำนวน",
      "ค่าใช้จ่าย หน่วย",
      "หมายเหตุ",
    ])
    example.addRow([
      "ORDER001",
      "2025-06-30",
      platforms[0]?.name || "",
      creators[0]?.name || "",
      "จากครีเอเตอร์",
      "ลูกค้า A",
      "แคมเปญ X",
      "เสร็จสิ้น",
      products[0]?.name || "",
      products[0]?.category || "",
      products[0]?.sku || "",
      10,
      products[0]?.costprice || "",
      products[0]?.suggestedPrice || "",
      "ค่าขนส่ง",
      50,
      "บาท",
      "ใช้ค่าคอมจากครีเอเตอร์",
    ])
    example.addRow([
      "ORDER002",
      "2025-06-30",
      platforms[0]?.name || "",
      "ขายเอง",
      "จากสินค้า",
      "ลูกค้า B",
      "",
      "เสร็จสิ้น",
      products[1]?.name || "",
      products[1]?.category || "",
      products[1]?.sku || "",
      5,
      products[1]?.costprice || "",
      products[1]?.suggestedPrice || "",
      "",
      "",
      "",
      "ขายเองไม่มีค่าคอม",
    ])

    // Sheet 6: คำอธิบาย
    const instructionSheet = wb.addWorksheet("คำอธิบาย")
    instructionSheet.addRow(["คำอธิบายการใช้งาน Template"])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["1. ประเภทค่าคอม:"])
    instructionSheet.addRow(['   - "จากสินค้า" = ใช้ค่าคอมที่ตั้งไว้ในสินค้า'])
    instructionSheet.addRow(['   - "จากครีเอเตอร์" = ใช้ค่าคอมที่ตั้งไว้ในครีเอเตอร์'])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["2. การกรอกข้อมูล:"])
    instructionSheet.addRow(["   - กรอกข้อมูลในแถวที่ 2 เป็นต้นไป"])
    instructionSheet.addRow(["   - ใช้ Order ID เดียวกันสำหรับสินค้าหลายชนิด"])
    instructionSheet.addRow(["   - เลือกจาก Dropdown เท่านั้น"])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["3. หมายเหตุ:"])
    instructionSheet.addRow(['   - หาก Creator = "ขายเอง" จะไม่มีค่าคอม'])
    instructionSheet.addRow(["   - ค่าคอมจะคำนวณตามประเภทที่เลือก"])

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    res.setHeader("Content-Disposition", "attachment; filename=order_template_with_commission_type.xlsx")
    await wb.xlsx.write(res)
    res.end()
  } catch (error) {
    console.error("Error generating Excel:", error)
    res.status(500).send("Error generating Excel file")
  }
})

// Health check endpoint
app.get("/health", (req, res) => {
  res.json({ status: "OK", timestamp: new Date().toISOString() })
})

app.listen(port, () => {
  console.log(`✅ Excel download API running at http://localhost:${port}`)
  console.log(`🌐 Health check: http://localhost:${port}/health`)
})
