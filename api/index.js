require("dotenv").config();
const express = require("express");
const fs = require("fs");
const path = require("path");
const { PDFDocument, rgb } = require("pdf-lib");
const fontkit = require("@pdf-lib/fontkit");
const axios = require("axios");
const nodemailer = require("nodemailer");
const getPdfMap = require("../config/mappingPDF"); // PDF coordinates

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

// Health check route
app.get("/", (req, res) => {
  res.json({
    success: true,
    message: "Server running successfully 🚀",
  });
});

// Export route
app.post("/export", async (req, res) => {
  const type = req.query.type || "offset";
  console.log("Export type:", type);

  const fieldToPdfMap = getPdfMap(type);
  const { recordId } = req.body;

  if (!recordId) return res.status(400).json({ error: "recordId required" });

  try {
    // Fetch Kintone record
    const record = await fetchKintoneRecord(recordId);

    // Load PDF template - FIXED PATH for Vercel
    const templatePath = path.join(
      process.cwd(),
      process.env.PDF_TEMPLATE_DIR || "templates",
      "QUOTATION TEMPLATE.pdf"
    );

    // Check if template file exists
    if (!fs.existsSync(templatePath)) {
      console.error("Template not found at:", templatePath);
      console.error(
        "Current directory contents:",
        fs.readdirSync(process.cwd())
      );
      return res.status(500).json({
        error: "Template file not found",
        path: templatePath,
        cwd: process.cwd(),
        files: fs.readdirSync(process.cwd()).join(", "),
      });
    }

    const templateBytes = fs.readFileSync(templatePath);
    const pdfDoc = await PDFDocument.load(templateBytes);
    pdfDoc.registerFontkit(fontkit);

    // Load font - FIXED PATH for Vercel
    const fontPath = path.join(process.cwd(), "fonts", "Roboto-Regular.ttf");

    // Check if font file exists
    if (!fs.existsSync(fontPath)) {
      console.error("Font not found at:", fontPath);
      return res.status(500).json({
        error: "Font file not found",
        path: fontPath,
      });
    }

    const font = await pdfDoc.embedFont(fs.readFileSync(fontPath));
    const page = pdfDoc.getPages()[0];

    // Determine price fields based on export type
    const priceFields =
      type === "offset"
        ? ["offset_unit_selling_price_official", "offset_total_amount"]
        : ["digital_unit_selling_price_official", "digital_total_amount"];

    // Loop through mapping
    for (const [fieldCode, mapping] of Object.entries(fieldToPdfMap)) {
      const field = record[fieldCode];
      if (!field) continue;

      // Handle signature IMAGE
      if (mapping.isImage) {
        if (Array.isArray(field.value) && field.value.length > 0) {
          const fileKey = field.value[0].fileKey;
          const fileUrl = `https://${process.env.KINTONE_DOMAIN}/k/v1/file.json?fileKey=${fileKey}`;

          try {
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
          } catch (imgError) {
            console.error(
              `Failed to load image for field ${fieldCode}:`,
              imgError.message
            );
            // Continue with other fields even if image fails
          }
        }
        continue;
      }

      // Handle TEXT FIELDS
      let value = field.value;

      // Normalize dropdown/user fields
      if (typeof value === "object" && value !== null && value.name) {
        value = value.name;
      }

      if (Array.isArray(value) && value.length > 0) {
        value = value[0].name || value[0];
      }

      if (!value) continue;

      // Format PRICE fields dynamically
      if (priceFields.includes(fieldCode)) {
        const numericValue = Number(value);
        if (!isNaN(numericValue)) {
          value = `₱ ${numericValue.toLocaleString(undefined, {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          })}`;
        } else {
          value = `₱ ${value}`;
        }
      }

      // Format date
      if (fieldCode === "date") {
        const d = new Date(value);
        if (!isNaN(d)) {
          value = d.toLocaleDateString("en-US", {
            month: "short",
            day: "numeric",
            year: "numeric",
          });
        }
      }

      // Format order qty
      if (fieldCode === "order_qty") {
        const n = Number(value);
        if (!isNaN(n)) value = n.toLocaleString();
      }

      // Draw text
      const textHeight = font.heightAtSize(9);

      page.drawText(String(value), {
        x: mapping.left,
        y: page.getHeight() - mapping.top - (mapping.height + textHeight) / 2,
        size: 6,
        font,
        color: rgb(0, 0, 0),
      });
    }

    // Save final PDF
    const pdfBytes = await pdfDoc.save();

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="Quotation_${type}.pdf"`
    );
    res.setHeader("Content-Type", "application/pdf");
    res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error("Export failed:", err.message);
    console.error("Stack trace:", err.stack);
    res.status(500).json({
      error: "Export failed",
      details: err.message,
      stack: err.stack,
    });
  }
});
module.exports = app;

if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () =>
    console.log(`🚀 Server running at http://localhost:${PORT}`)
  );
}
