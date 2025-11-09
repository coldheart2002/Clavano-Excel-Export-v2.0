require("dotenv").config();
const express = require("express");
const fs = require("fs");
const path = require("path");
const { PDFDocument, rgb } = require("pdf-lib");
const fontkit = require("@pdf-lib/fontkit");
const axios = require("axios");
const nodemailer = require("nodemailer");
const fieldToPdfMap = require("../config/mappingPDF"); // PDF coordinates

const app = express();
const allowedOrigin = "https://clavano-printers.kintone.com";

app.use(express.json());

// Toggle for visualizing coordinates
const SHOW_COORDINATES = false; // set true for mapping boxes

// CORS setup
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

// Send email with PDF attachment
async function sendEmail(toEmail, buffer, fileName) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS },
  });

  await transporter.sendMail({
    from: process.env.EMAIL_USER,
    to: toEmail,
    subject: "Your Quotation",
    text: "Please find attached your quotation.",
    attachments: [{ filename: fileName, content: buffer }],
  });
}

// Export route
app.post("/export", async (req, res) => {
  const { recordId } = req.body;
  if (!recordId) return res.status(400).json({ error: "recordId required" });

  try {
    const record = await fetchKintoneRecord(recordId);

    const templatePath = path.resolve(
      process.env.PDF_TEMPLATE_DIR || "./templates",
      "QUOTATION TEMPLATE.pdf"
    );
    const templateBytes = fs.readFileSync(templatePath);
    const pdfDoc = await PDFDocument.load(templateBytes);

    // Register fontkit for custom fonts
    pdfDoc.registerFontkit(fontkit);

    // Embed Calibri font
    const calibriPath = path.resolve(__dirname, "../fonts/Roboto-Regular.ttf");
    const calibriBytes = fs.readFileSync(calibriPath);
    const font = await pdfDoc.embedFont(calibriBytes);

    const page = pdfDoc.getPages()[0];

    // Fill fields
    for (const [fieldCode, mapping] of Object.entries(fieldToPdfMap)) {
      const field = record[fieldCode];
      if (!field) continue;

      // Draw visualization box if toggle is on
      if (SHOW_COORDINATES) {
        page.drawRectangle({
          x: mapping.left,
          y: page.getHeight() - mapping.top - mapping.height,
          width: mapping.width,
          height: mapping.height,
          borderColor: rgb(1, 0, 0),
          borderWidth: 1,
          color: rgb(1, 1, 1, 0), // transparent
        });

        page.drawText(fieldCode, {
          x: mapping.left + 2,
          y: page.getHeight() - mapping.top - mapping.height / 2 - 4.5,
          size: 1,
          font: font,
          color: rgb(1, 0, 0),
        });
        continue; // skip normal drawing while visualizing
      }

      // Handle images (signature)
      if (
        mapping.isImage &&
        Array.isArray(field.value) &&
        field.value.length > 0
      ) {
        const fileKey = field.value[0].fileKey;
        const fileUrl = `https://${process.env.KINTONE_DOMAIN}/k/v1/file.json?fileKey=${fileKey}`;
        const imgResp = await axios.get(fileUrl, {
          responseType: "arraybuffer",
          headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
        });
        const pngImage = await pdfDoc.embedPng(imgResp.data);
        page.drawImage(pngImage, {
          x: mapping.left,
          y: page.getHeight() - mapping.top - mapping.height,
          width: mapping.width,
          height: mapping.height,
        });
        continue;
      }

      // Handle text
      let value = field.value;
      if (typeof value === "object" && value.name) value = value.name;
      if (Array.isArray(value) && value.length > 0)
        value = value[0].name || value[0];
      if (!value) continue;

      // Format fields
      if (fieldCode === "officialUnitPrice" || fieldCode === "totalAmount") {
        const numericValue = Number(value);
        if (!isNaN(numericValue)) {
          value = `₱ ${numericValue.toLocaleString()}`;
        } else {
          value = `₱ ${value}`;
        }
      } else if (fieldCode === "orderQuantity") {
        const numericValue = Number(value);
        if (!isNaN(numericValue)) {
          value = numericValue.toLocaleString();
        }
      } else if (fieldCode === "weight") {
        const numericValue = Number(value);
        if (!isNaN(numericValue)) {
          value = `${numericValue} gsm`;
        }
      } else if (fieldCode === "date" && value) {
        const dateObj = new Date(value);
        if (!isNaN(dateObj)) {
          const options = { month: "short", day: "numeric", year: "numeric" };
          value = dateObj.toLocaleDateString("en-US", options); // 👉 e.g. Jan 26, 2002
        }
      } else if (fieldCode === "salesRepresentative") {
        console.log(value);
      }

      // Precise left-center alignment
      const textHeight = font.heightAtSize(9);
      page.drawText(String(value), {
        x: mapping.left,
        y: page.getHeight() - mapping.top - (mapping.height + textHeight) / 2,
        size: 8,
        font: font,
        color: rgb(0, 0, 0),
      });
    }

    const pdfBytes = await pdfDoc.save();

    // Email if emailAddress exists (optional)
    const clientEmail = record.emailAddress?.value;
    if (clientEmail) {
      // await sendEmail(clientEmail, pdfBytes, "Quotation.pdf");
    }

    // Send PDF to client
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="Quotation.pdf"`
    );
    res.setHeader("Content-Type", "application/pdf");
    res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error("Export failed:", err.message);
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
