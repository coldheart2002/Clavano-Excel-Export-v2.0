/*  api/index.js  */
require("dotenv").config();
const express = require("express");
const path = require("path");
const { PDFDocument, rgb } = require("pdf-lib");
const fontkit = require("@pdf-lib/fontkit");
const axios = require("axios");
const nodemailer = require("nodemailer");
const getPdfMap = require("../config/mappingPDF");

/* ---------- 1.  READ STATIC ASSETS ONCE  ---------- */
const TEMPLATE_BUF = require("fs").readFileSync(
  path.join(__dirname, "..", "templates", "QUOTATION TEMPLATE.pdf")
);
const ROBOTO_BUF = require("fs").readFileSync(
  path.join(__dirname, "..", "fonts", "Roboto-Regular.ttf")
);

/* ---------- 2.  EXPRESS SET-UP  ---------- */
const app = express();
const allowedOrigin = "https://clavano-printers.kintone.com"; // ← space removed

app.use(express.json());

/* ---------- 3.  CORS  ---------- */
app.options("/export", (_, res) => {
  res.setHeader("Access-Control-Allow-Origin", allowedOrigin);
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.status(204).end();
});
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", allowedOrigin);
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.header("Access-Control-Allow-Credentials", "true");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

/* ---------- 4.  HELPERS  ---------- */
async function fetchKintoneRecord(recordId) {
  const url = `https://${process.env.KINTONE_DOMAIN}/k/v1/record.json`;
  const { data } = await axios.get(url, {
    params: { app: process.env.KINTONE_APP_ID, id: recordId },
    headers: { "X-Cybozu-API-Token": process.env.KINTONE_API_TOKEN },
  });
  return data.record;
}

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

/* ---------- 5.  ROUTES  ---------- */
app.get("/", (_, res) =>
  res.json({ success: true, message: "Server running successfully 🚀" })
);

app.post("/export", async (req, res) => {
  const type = req.query.type || "offset";
  const { recordId } = req.body;
  if (!recordId) return res.status(400).json({ error: "recordId required" });

  try {
    const record = await fetchKintoneRecord(recordId);
    const fieldToPdfMap = getPdfMap(type);

    /* ---- PDF  ---- */
    const pdfDoc = await PDFDocument.load(TEMPLATE_BUF);
    pdfDoc.registerFontkit(fontkit);
    const font = await pdfDoc.embedFont(ROBOTO_BUF);
    const page = pdfDoc.getPages()[0];

    const priceFields =
      type === "offset"
        ? ["offset_unit_selling_price_official", "offset_total_amount"]
        : ["digital_unit_selling_price_official", "digital_total_amount"];

    /* ---- LOOP OVER FIELDS  ---- */
    for (const [fieldCode, mapping] of Object.entries(fieldToPdfMap)) {
      const field = record[fieldCode];
      if (!field) continue;

      /* IMAGE */
      if (mapping.isImage) {
        if (Array.isArray(field.value) && field.value.length) {
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
          } catch (e) {
            console.error(`Image embed failed for ${fieldCode}:`, e.message);
          }
        }
        continue;
      }

      /* TEXT */
      let value = field.value;
      if (value && typeof value === "object" && value.name) value = value.name;
      if (Array.isArray(value) && value.length)
        value = value[0].name || value[0];
      if (!value) continue;

      if (priceFields.includes(fieldCode)) {
        const n = Number(value);
        value = `₱ ${
          !isNaN(n)
            ? n.toLocaleString(undefined, {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2,
              })
            : value
        }`;
      }
      if (fieldCode === "date") {
        const d = new Date(value);
        if (!isNaN(d))
          value = d.toLocaleDateString("en-US", {
            month: "short",
            day: "numeric",
            year: "numeric",
          });
      }
      if (fieldCode === "order_qty") {
        const n = Number(value);
        if (!isNaN(n)) value = n.toLocaleString();
      }

      const textSize = 6;
      page.drawText(String(value), {
        x: mapping.left,
        y:
          page.getHeight() -
          mapping.top -
          (mapping.height + font.heightAtSize(textSize)) / 2,
        size: textSize,
        font,
        color: rgb(0, 0, 0),
      });
    }

    /* ---- RETURN PDF  ---- */
    const pdfBytes = await pdfDoc.save();
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="Quotation_${type}.pdf"`
    );
    res.setHeader("Content-Type", "application/pdf");
    res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error("Export failed:", err.message, err.stack);
    res.status(500).json({ error: "Export failed", details: err.message });
  }
});

/* ---------- 6.  START OR EXPORT  ---------- */
module.exports = app;
if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () =>
    console.log(`🚀 Server running at http://localhost:${PORT}`)
  );
}
