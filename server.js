require("dotenv").config();
const express = require("express");
const ExcelJS = require("exceljs");
const axios = require("axios");
const path = require("path");
const fieldToExcelMap = require("./mapping");

const app = express();

// ✅ Allowed Origin (Kintone domain)
const allowedOrigin = "https://clavano-printers.kintone.com";

app.use(express.json());

// ✅ 1️⃣ Explicit preflight route — Vercel requires this
app.options("/export", (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", allowedOrigin);
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  return res.status(204).end();
});

// ✅ 2️⃣ Global CORS middleware for all other routes
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", allowedOrigin);
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.header("Access-Control-Allow-Credentials", "true");

  if (req.method === "OPTIONS") {
    return res.sendStatus(204);
  }

  next();
});

// 🔹 Fetch record from Kintone
async function fetchKintoneRecord(recordId) {
  const url = `https://${process.env.KINTONE_DOMAIN}/k/v1/record.json`;
  const response = await axios.get(url, {
    params: { app: process.env.KINTONE_APP_ID, id: recordId },
    headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
  });
  return response.data.record;
}

// 🔹 Health check route
app.get("/", (req, res) => {
  res.json({ success: true, message: "Server running successfully" });
});

// 🔹 Main export route
app.post("/export", async (req, res) => {
  try {
    const { recordId } = req.body;
    if (!recordId)
      return res.status(400).json({ error: "recordId is required" });

    console.log(`📥 Export requested for recordId: ${recordId}`);

    // 1. Fetch record from Kintone
    const record = await fetchKintoneRecord(recordId);

    // 2. Load Excel template
    const templateFile = "QUOTATION TEMPLATE.xlsx";
    const templatePath = path.resolve(
      process.env.EXCEL_TEMPLATE_DIR || "./templates",
      templateFile
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // 3. Apply field mappings
    for (const [fieldCode, mapping] of Object.entries(fieldToExcelMap)) {
      if (!record[fieldCode]) continue;
      const ws = workbook.getWorksheet(mapping.sheet);
      if (!ws) continue;

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
        if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
          const dateObj = new Date(value);
          ws.getCell(mapping.cell).value = dateObj;
          ws.getCell(mapping.cell).numFmt = "mmm dd, yyyy";
        } else {
          ws.getCell(mapping.cell).value = value;
        }
      }
    }

    // 4. Generate Excel buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // 5. Send file
    res.setHeader("Access-Control-Allow-Origin", allowedOrigin);
    res.setHeader("Access-Control-Allow-Credentials", "true");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${templateFile}"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.send(buffer);
    console.log(`✅ ${templateFile} generated and sent successfully`);
  } catch (err) {
    console.error("❌ Export failed:", err.message);
    res.status(500).json({ error: "Export failed", details: err.message });
  }
});

// 🔹 Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Export server running at http://localhost:${PORT}`);
});
