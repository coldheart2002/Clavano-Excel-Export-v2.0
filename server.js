require("dotenv").config();
const express = require("express");
const ExcelJS = require("exceljs");
const axios = require("axios");
const path = require("path");
const cors = require("cors");
const fieldToExcelMap = require("./mapping");

const app = express();

// ✅ CONFIGURABLE VARIABLES
const SHEET_NAME = "COSTING  SHEET"; // Excel sheet name
const DIGITAL_UNIT_PRICE_CELL = "S54";
const OFFSET_UNIT_PRICE_CELL = "P54";
const CUSTOMER_NAME_FIELD = "customer_name";
const SKU_FIELD = "sku";

// ✅ Enable CORS for your Kintone domain
app.use(
  cors({
    origin: "https://clavano-printers.kintone.com",
    methods: ["GET", "POST"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

app.use(express.json());

// 🔹 Fetch record from Kintone
async function fetchKintoneRecord(recordId) {
  const url = `https://${process.env.KINTONE_DOMAIN}/k/v1/record.json`;
  const response = await axios.get(url, {
    params: { app: process.env.KINTONE_APP_ID, id: recordId },
    headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
  });
  return response.data.record;
}

app.get("/test", async (req, res) => {
  res.json({ success: true, message: "test successful" });
});

// 🔹 API route for export
app.post("/export", async (req, res) => {
  const { recordId } = req.body;

  if (!recordId) {
    return res.status(400).json({ error: "recordId is required" });
  }

  try {
    console.log(`📥 Export requested for recordId: ${recordId}`);

    // 1. Fetch record
    const record = await fetchKintoneRecord(recordId);

    // 2. Always use the same Excel template
    const templateFile = "QUOTATION TEMPLATE.xlsx";
    const templatePath = path.resolve(
      process.env.EXCEL_TEMPLATE_DIR || "./templates",
      templateFile
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // 3. Apply field mappings (Kintone -> Excel)
    for (const [fieldCode, mapping] of Object.entries(fieldToExcelMap)) {
      if (record[fieldCode]) {
        const ws = workbook.getWorksheet(mapping.sheet);
        if (ws) {
          let value = record[fieldCode].value;
          let handled = false;

          if (typeof mapping.extract === "function") {
            const result = mapping.extract(
              record[fieldCode].value,
              ws,
              mapping.cell,
              mapping.concat || false
            );
            if (result === null) handled = true;
            else value = result;
          }

          if (!handled) {
            // Handle date fields
            if (
              typeof value === "string" &&
              /^\d{4}-\d{2}-\d{2}$/.test(value)
            ) {
              const dateObj = new Date(value);
              ws.getCell(mapping.cell).value = dateObj;
              ws.getCell(mapping.cell).numFmt = "mmm dd, yyyy";
            } else {
              ws.getCell(mapping.cell).value = value;
            }
          }
        } else {
          console.warn(`⚠️ Worksheet "${mapping.sheet}" not found`);
        }
      } else {
        console.warn(`⚠️ Field "${fieldCode}" not found in record`);
      }
    }

    // 4. Send Excel file to client
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${templateFile}"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);

    console.log(`✅ ${templateFile} generated and downloaded`);
  } catch (err) {
    console.error("❌ Export failed:", err.message);
    res.status(500).json({ error: "Export failed" });
  }
});

// 🔹 Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Export server running at http://localhost:${PORT}`);
});
