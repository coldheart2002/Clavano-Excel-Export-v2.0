require("dotenv").config();
const express = require("express");
const ExcelJS = require("exceljs");
const axios = require("axios");
const path = require("path");
const fieldToExcelMap = require("../mapping"); // adjust path if mapping.js is not one folder up

const app = express();

// ✅ Allowed Origin (your Kintone domain)
const allowedOrigin = "https://clavano-printers.kintone.com";

// ✅ Middleware
app.use(express.json());

// ✅ Preflight route (required for Vercel CORS)
app.options("/export", (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", allowedOrigin);
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  return res.status(204).end();
});

// ✅ Global CORS middleware
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

// 🔹 Fetch Kintone record by ID
async function fetchKintoneRecord(recordId) {
  const url = `https://${process.env.KINTONE_DOMAIN}/k/v1/record.json`;
  const response = await axios.get(url, {
    params: { app: process.env.KINTONE_APP_ID, id: recordId },
    headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
  });
  return response.data.record;
}

// 🔹 Root health check
app.get("/", (req, res) => {
  res.json({ success: true, message: "Server running successfully" });
});

// 🔹 Export route (Excel generation)
app.post("/export", async (req, res) => {
  const { recordId } = req.body;
  if (!recordId) {
    return res.status(400).json({ error: "recordId is required" });
  }

  try {
    console.log(`📥 Export requested for recordId: ${recordId}`);

    // 1️⃣ Fetch record from Kintone
    const record = await fetchKintoneRecord(recordId);

    // 2️⃣ Load Excel template
    const templateFile = "QUOTATION TEMPLATE.xlsx";
    const templatePath = path.resolve(
      process.env.EXCEL_TEMPLATE_DIR || "./templates",
      templateFile
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // 3️⃣ Apply field mappings (Kintone → Excel)
    for (const [fieldCode, mapping] of Object.entries(fieldToExcelMap)) {
      const field = record[fieldCode];
      if (!field) {
        console.warn(`⚠️ Field "${fieldCode}" not found in record`);
        continue;
      }

      const ws = workbook.getWorksheet(mapping.sheet);
      if (!ws) {
        console.warn(`⚠️ Worksheet "${mapping.sheet}" not found`);
        continue;
      }

      // ✅ Handle image fields (e.g., signature)
      if (
        mapping.isImage &&
        Array.isArray(field.value) &&
        field.value.length > 0
      ) {
        try {
          const fileInfo = field.value[0];
          const fileKey = fileInfo.fileKey;

          const fileUrl = `https://${process.env.KINTONE_DOMAIN}/k/v1/file.json?fileKey=${fileKey}`;
          const imgResponse = await axios.get(fileUrl, {
            responseType: "arraybuffer",
            headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
          });

          const imageId = workbook.addImage({
            buffer: imgResponse.data,
            extension: "png",
          });

          const cell = ws.getCell(mapping.cell);
          const col = cell.col;
          const row = cell.row;

          ws.addImage(imageId, {
            tl: { col: col - 1, row: row - 1 },
            ext: { width: mapping.width || 120, height: mapping.height || 50 },
          });

          console.log(`🖋️ Added image for ${fieldCode} at ${mapping.cell}`);
          continue;
        } catch (imgErr) {
          console.error(
            `❌ Failed to add image for ${fieldCode}:`,
            imgErr.message
          );
          continue;
        }
      }

      // ✅ Handle text/number/date fields
      let value = field.value;
      let handled = false;

      if (typeof mapping.extract === "function") {
        const result = mapping.extract(
          value,
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

    // 4️⃣ Send Excel buffer as response
    const buffer = await workbook.xlsx.writeBuffer();
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

// ✅ Export for Vercel serverless function
module.exports = app;

// ✅ Allow local testing
if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () =>
    console.log(`🚀 Local server running at http://localhost:${PORT}`)
  );
}
