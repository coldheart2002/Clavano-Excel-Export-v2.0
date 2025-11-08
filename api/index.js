require("dotenv").config();
const express = require("express");
const ExcelJS = require("exceljs");
const axios = require("axios");
const path = require("path");
const nodemailer = require("nodemailer");
const fieldToExcelMap = require("../mapping"); // adjust path if needed

const app = express();
const allowedOrigin = "https://clavano-printers.kintone.com";

app.use(express.json());

// Preflight CORS
app.options("/export", (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", allowedOrigin);
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  return res.status(204).end();
});

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", allowedOrigin);
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.header("Access-Control-Allow-Credentials", "true");

  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

// Fetch Kintone record
async function fetchKintoneRecord(recordId) {
  const url = `https://${process.env.KINTONE_DOMAIN}/k/v1/record.json`;
  const response = await axios.get(url, {
    params: { app: process.env.KINTONE_APP_ID, id: recordId },
    headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
  });
  return response.data.record;
}

// Send email with attachment
async function sendEmail(toEmail, buffer, fileName) {
  const transporter = nodemailer.createTransport({
    service: "gmail", // or SMTP
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  await transporter.sendMail({
    from: process.env.EMAIL_USER,
    to: toEmail,
    subject: "Your Quotation",
    text: "Please find attached your quotation.",
    attachments: [
      {
        filename: fileName,
        content: buffer,
      },
    ],
  });
}

// Export route
app.post("/export", async (req, res) => {
  const { recordId } = req.body;
  if (!recordId) return res.status(400).json({ error: "recordId required" });

  try {
    console.log(`📥 Export requested for recordId: ${recordId}`);
    const record = await fetchKintoneRecord(recordId);

    const templateFile = "QUOTATION TEMPLATE.xlsx";
    const templatePath = path.resolve(
      process.env.EXCEL_TEMPLATE_DIR || "./templates",
      templateFile
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // Apply mapping
    for (const [fieldCode, mapping] of Object.entries(fieldToExcelMap)) {
      const field = record[fieldCode];
      if (!field) continue;
      const ws = workbook.getWorksheet(mapping.sheet);
      if (!ws) continue;

      // Handle image fields
      if (
        mapping.isImage &&
        Array.isArray(field.value) &&
        field.value.length > 0
      ) {
        try {
          const fileKey = field.value[0].fileKey;
          const fileUrl = `https://${process.env.KINTONE_DOMAIN}/k/v1/file.json?fileKey=${fileKey}`;
          const imgResp = await axios.get(fileUrl, {
            responseType: "arraybuffer",
            headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
          });

          const imageId = workbook.addImage({
            buffer: imgResp.data,
            extension: "png",
          });

          const cell = ws.getCell(mapping.cell);
          ws.addImage(imageId, {
            tl: { col: cell.col - 1, row: cell.row - 1 },
            ext: { width: mapping.width || 120, height: mapping.height || 50 },
          });

          continue;
        } catch (err) {
          console.error(`❌ Image error for ${fieldCode}:`, err.message);
          continue;
        }
      }

      // Handle text/number/date fields
      let value = field.value;
      if (mapping.extract && typeof mapping.extract === "function") {
        mapping.extract(value, ws, mapping.cell);
        continue;
      }

      // Default assignment
      if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
        ws.getCell(mapping.cell).value = new Date(value);
        ws.getCell(mapping.cell).numFmt = "mmm dd, yyyy";
      } else {
        ws.getCell(mapping.cell).value = value;
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();

    // Get client email from mapping
    const clientEmailField = record.emailAddress?.value; // your mapping has 'emailAddress'
    if (clientEmailField) {
      await sendEmail(clientEmailField, buffer, "Quotation.xlsx");
      console.log(`✅ Quotation emailed to ${clientEmailField}`);
    }

    // Send Excel for download
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="Quotation.xlsx"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);
  } catch (err) {
    console.error("❌ Export failed:", err.message);
    res.status(500).json({ error: "Export failed", details: err.message });
  }
});

module.exports = app;

if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () =>
    console.log(`🚀 Server running at http://localhost:${PORT}`)
  );
}
