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

    // Sheet 1: Orders
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
      "ราคาขายต่อชิ้น",
      "ค่าใช้จ่าย ชื่อ",
      "ค่าใช้จ่าย จำนวน",
      "ค่าใช้จ่าย หน่วย",
      "หมายเหตุ",
    ])

    // ตั้งค่าความกว้างคอลัมน์
    ordersSheet.getColumn("A").width = 12 // Order ID
    ordersSheet.getColumn("B").width = 15 // วันที่
    ordersSheet.getColumn("C").width = 12 // Platform
    ordersSheet.getColumn("D").width = 15 // Creator
    ordersSheet.getColumn("E").width = 15 // ประเภทค่าคอม
    ordersSheet.getColumn("I").width = 25 // ชื่อสินค้า

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

      // VLOOKUP formulas
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

      // เพิ่ม default values
      if (row === 2) {
        ordersSheet.getCell(`E${row}`).value = "จากสินค้า"
        // ตั้งค่าวันที่เป็นวันปัจจุบันในรูปแบบที่ถูกต้อง
        ordersSheet.getCell(`B${row}`).value = new Date()
        ordersSheet.getCell(`B${row}`).numFmt = "dd/mm/yyyy"
      }
    }

    // Sheet 2: Dictionary
    const dictSheet = wb.addWorksheet("Dictionary")
    platforms.forEach((p, i) => (dictSheet.getCell(`A${i + 2}`).value = p.name))
    creators.forEach((c, i) => (dictSheet.getCell(`B${i + 2}`).value = c.name))
    products.forEach((p, i) => (dictSheet.getCell(`C${i + 2}`).value = p.name))
    ;["เสร็จสิ้น", "ยกเลิก", "กำลังดำเนินการ"].forEach((v, i) => (dictSheet.getCell(`D${i + 2}`).value = v))
    ;["บาท", "%"].forEach((v, i) => (dictSheet.getCell(`E${i + 2}`).value = v))
    ;["จากสินค้า", "จากครีเอเตอร์"].forEach((v, i) => (dictSheet.getCell(`F${i + 2}`).value = v))

    // Sheet 3: ProductData
    const productSheet = wb.addWorksheet("ProductData")
    productSheet.addRow(["ชื่อสินค้า", "หมวดหมู่", "รหัสสินค้า", "ต้นทุนต่อชิ้น", "ราคาขายต่อชิ้น", "ค่าคอมสินค้า(%)"])
    products.forEach((p) =>
      productSheet.addRow([p.name, p.category, p.sku, p.costprice, p.suggestedPrice, p.commissionRate || 0]),
    )

    // Sheet 4: CreatorData
    const creatorSheet = wb.addWorksheet("CreatorData")
    creatorSheet.addRow(["ชื่อครีเอเตอร์", "ค่าคอมครีเอเตอร์(%)"])
    creators.forEach((c) => creatorSheet.addRow([c.name, c.commission_rate || 0]))

    // Sheet 5: ตัวอย่างการกรอก (ปรับปรุงใหม่)
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

    // ตัวอย่างที่ 1: สินค้าเดียว + ค่าใช้จ่ายเดียว
    const today = new Date()
    example.addRow([
      "ORDER001",
      today,
      platforms[0]?.name || "TikTok",
      creators[0]?.name || "ขายเอง",
      "จากสินค้า",
      "ลูกค้า A",
      "แคมเปญ X",
      "เสร็จสิ้น",
      products[0]?.name || "สินค้า A",
      products[0]?.category || "หมวดหมู่ A",
      products[0]?.sku || "SKU001",
      10,
      products[0]?.costprice || 100,
      products[0]?.suggestedPrice || 200,
      "ค่าขนส่ง",
      50,
      "บาท",
      "สินค้าเดียว ค่าใช้จ่ายเดียว",
    ])

    // ตัวอย่างที่ 2: สินค้าหลายรายการ + ค่าใช้จ่ายเดียว
    example.addRow([
      "ORDER002",
      today,
      platforms[0]?.name || "TikTok",
      creators[0]?.name || "ขายเอง",
      "จากครีเอเตอร์",
      "ลูกค้า B",
      "แคมเปญ Y",
      "เสร็จสิ้น",
      products[0]?.name || "สินค้า A",
      products[0]?.category || "หมวดหมู่ A",
      products[0]?.sku || "SKU001",
      5,
      products[0]?.costprice || 100,
      products[0]?.suggestedPrice || 200,
      "ค่าขนส่ง",
      100,
      "บาท",
      "สินค้าหลายรายการ ค่าใช้จ่ายเดียว",
    ])
    example.addRow([
      "ORDER002", // Order ID เดียวกัน
      "", // วันที่ว่าง
      "", // Platform ว่าง
      "", // Creator ว่าง
      "", // ประเภทค่าคอม ว่าง
      "", // ลูกค้า ว่าง
      "", // แคมเปญ ว่าง
      "", // สถานะ ว่าง
      products[1]?.name || "สินค้า B",
      products[1]?.category || "หมวดหมู่ B",
      products[1]?.sku || "SKU002",
      3,
      products[1]?.costprice || 150,
      products[1]?.suggestedPrice || 300,
      "", // ค่าใช้จ่าย ชื่อ ว่าง
      "", // ค่าใช้จ่าย จำนวน ว่าง
      "", // ค่าใช้จ่าย หน่วย ว่าง
      "สินค้าที่ 2 ในออเดอร์เดียวกัน",
    ])

    // ตัวอย่างที่ 3: สินค้าเดียว + ค่าใช้จ่ายหลายรายการ
    example.addRow([
      "ORDER003",
      today,
      platforms[0]?.name || "TikTok",
      creators[0]?.name || "ขายเอง",
      "จากสินค้า",
      "ลูกค้า C",
      "แคมเปญ Z",
      "เสร็จสิ้น",
      products[0]?.name || "สินค้า A",
      products[0]?.category || "หมวดหมู่ A",
      products[0]?.sku || "SKU001",
      8,
      products[0]?.costprice || 100,
      products[0]?.suggestedPrice || 200,
      "ค่าขนส่ง",
      80,
      "บาท",
      "สินค้าเดียว ค่าใช้จ่ายหลายรายการ",
    ])
    example.addRow([
      "ORDER003", // Order ID เดียวกัน
      "", // ข้อมูลอื่นว่าง
      "",
      "",
      "",
      "",
      "",
      "",
      "", // สินค้าว่าง
      "",
      "",
      "",
      "",
      "",
      "ค่าธรรมเนียม", // ค่าใช้จ่ายรายการที่ 2
      30,
      "บาท",
      "ค่าใช้จ่ายรายการที่ 2",
    ])
    example.addRow([
      "ORDER003", // Order ID เดียวกัน
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "ค่าโฆษณา", // ค่าใช้จ่ายรายการที่ 3
      5,
      "%",
      "ค่าใช้จ่ายรายการที่ 3 (เป็น %)",
    ])

    // ตัวอย่างที่ 4: สินค้าหลายรายการ + ค่าใช้จ่ายหลายรายการ
    example.addRow([
      "ORDER004",
      today,
      platforms[0]?.name || "TikTok",
      creators[0]?.name || "ขายเอง",
      "จากครีเอเตอร์",
      "ลูกค้า D",
      "แคมเปญ W",
      "เสร็จสิ้น",
      products[0]?.name || "สินค้า A",
      products[0]?.category || "หมวดหมู่ A",
      products[0]?.sku || "SKU001",
      12,
      products[0]?.costprice || 100,
      products[0]?.suggestedPrice || 200,
      "ค่าขนส่ง",
      120,
      "บาท",
      "สินค้าหลายรายการ ค่าใช้จ่ายหลายรายการ",
    ])
    example.addRow([
      "ORDER004",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      products[1]?.name || "สินค้า B", // สินค้าที่ 2
      products[1]?.category || "หมวดหมู่ B",
      products[1]?.sku || "SKU002",
      6,
      products[1]?.costprice || 150,
      products[1]?.suggestedPrice || 300,
      "ค่าบรรจุภัณฑ์", // ค่าใช้จ่ายรายการที่ 2
      50,
      "บาท",
      "สินค้าที่ 2 + ค่าใช้จ่ายที่ 2",
    ])
    example.addRow([
      "ORDER004",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      products[2]?.name || "สินค้า C", // สินค้าที่ 3
      products[2]?.category || "หมวดหมู่ C",
      products[2]?.sku || "SKU003",
      4,
      products[2]?.costprice || 200,
      products[2]?.suggestedPrice || 400,
      "ค่าโฆษณา", // ค่าใช้จ่ายรายการที่ 3
      3,
      "%",
      "สินค้าที่ 3 + ค่าใช้จ่ายที่ 3 (เป็น %)",
    ])

    // ตั้งค่าการแสดงวันที่ให้ถูกต้อง
    example.getColumn("B").numFmt = "dd/mm/yyyy"

    // Sheet 6: คำอธิบาย (ปรับปรุงใหม่)
    const instructionSheet = wb.addWorksheet("คำอธิบาย")
    instructionSheet.addRow(["คำอธิบายการใช้งาน Template"])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["1. การกรอกวันที่:"])
    instructionSheet.addRow(["   - ใช้รูปแบบ dd/mm/yyyy เช่น 02/07/2025"])
    instructionSheet.addRow(["   - หรือ yyyy-mm-dd เช่น 2025-07-02"])
    instructionSheet.addRow(["   - หรือพิมพ์ตัวเลขวันที่ Excel จะแปลงให้อัตโนมัติ"])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["2. ประเภทค่าคอม:"])
    instructionSheet.addRow(['   - "จากสินค้า" = ใช้ค่าคอมที่ตั้งไว้ในสินค้า'])
    instructionSheet.addRow(['   - "จากครีเอเตอร์" = ใช้ค่าคอมที่ตั้งไว้ในครีเอเตอร์'])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["3. การกรอกออเดอร์หลายรายการ:"])
    instructionSheet.addRow(["   - ใช้ Order ID เดียวกันสำหรับสินค้าหลายชนิด"])
    instructionSheet.addRow(["   - กรอกข้อมูลหลักในแถวแรก"])
    instructionSheet.addRow(["   - แถวถัดไปใส่เฉพาะสินค้าหรือค่าใช้จ่ายเพิ่มเติม"])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["4. การกรอกค่าใช้จ่าย:"])
    instructionSheet.addRow(["   - สามารถมีหลายรายการต่อ 1 ออเดอร์"])
    instructionSheet.addRow(["   - หน่วยเป็น 'บาท' หรือ '%'"])
    instructionSheet.addRow(["   - ถ้าเป็น % จะคิดจากยอดขายรวม"])
    instructionSheet.addRow([""])
    instructionSheet.addRow(["5. ตัวอย่างในไฟล์:"])
    instructionSheet.addRow(["   - ORDER001: สินค้าเดียว + ค่าใช้จ่ายเดียว"])
    instructionSheet.addRow(["   - ORDER002: สินค้าหลายรายการ + ค่าใช้จ่ายเดียว"])
    instructionSheet.addRow(["   - ORDER003: สินค้าเดียว + ค่าใช้จ่ายหลายรายการ"])
    instructionSheet.addRow(["   - ORDER004: สินค้าหลายรายการ + ค่าใช้จ่ายหลายรายการ"])

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    res.setHeader("Content-Disposition", "attachment; filename=order_template_complete.xlsx")
    await wb.xlsx.write(res)
    res.end()
  } catch (error) {
    console.error("Error generating Excel:", error)
    res.status(500).send("Error generating Excel file")
  }
})

app.get("/health", (req, res) => {
  res.json({ status: "OK", timestamp: new Date().toISOString() })
})

app.listen(port, () => {
  console.log(`✅ Excel download API running at http://localhost:${port}`)
})
