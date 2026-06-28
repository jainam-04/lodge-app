const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");
const multer = require("multer");
const Tesseract = require("tesseract.js");
const sharp = require("sharp");
const path = require("path");
const fs = require("fs");
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// ─── Multer — store uploaded card images temporarily ─────────────────────────
const upload = multer({
  dest: "uploads/",
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB max
});

// ─── Google Sheets Auth ───────────────────────────────────────────────────────
// credentials.json must be placed in the same folder as server.js
// Download it from Google Cloud Console → Service Accounts → Keys
const auth = new google.auth.GoogleAuth({
  keyFile: path.join(__dirname, "credentials.json"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const SPREADSHEET_ID = process.env.SPREADSHEET_ID; // Set in .env file

// ─── Helper: Get or create sheet for current month ───────────────────────────
async function getOrCreateMonthSheet(sheets) {
  const now = new Date();
  const monthName = now.toLocaleString("en-IN", {
    month: "long",
    year: "numeric",
  }); // e.g. "June 2025"

  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
  });

  const existingSheets = spreadsheet.data.sheets.map(
    (s) => s.properties.title
  );

  if (!existingSheets.includes(monthName)) {
    // Create new sheet for this month
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: monthName },
            },
          },
        ],
      },
    });

    // Add header row
    const headers = [
      [
        "Sr No",
        "Check-in Date",
        "Check-in Time",
        "Room No",
        "Guest Name",
        "Aadhaar No",
        "DOB",
        "Gender",
        "Address",
        "Total Members",
        "Member 2 Name",
        "Member 2 Aadhaar",
        "Member 3 Name",
        "Member 3 Aadhaar",
        "Member 4 Name",
        "Member 4 Aadhaar",
        "Member 5 Name",
        "Member 5 Aadhaar",
        "Payment Type",
        "Cash Amount (₹)",
        "Online Amount (₹)",
        "Total Amount (₹)",
        "UPI/Ref No",
        "Remarks",
      ],
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${monthName}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: headers },
    });
  }

  return monthName;
}

// ─── Helper: Preprocess image for better OCR ─────────────────────────────────
async function preprocessImage(inputPath) {
  const outputPath = inputPath + "_processed.png";
  await sharp(inputPath)
    .resize({ width: 1600, withoutEnlargement: true }) // scale up small images
    .grayscale()
    .normalize() // boost contrast
    .sharpen()
    .png()
    .toFile(outputPath);
  return outputPath;
}

// ─── Helper: Parse Aadhaar text extracted by OCR ─────────────────────────────
function parseAadhaarText(text) {
  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);

  const result = {
    name: "",
    dob: "",
    gender: "",
    aadhaarNo: "",
    address: "",
  };

  // Aadhaar number: 12 digits (may appear as XXXX XXXX XXXX)
  const aadhaarMatch = text.match(/\d{4}\s?\d{4}\s?\d{4}/);
  if (aadhaarMatch) {
    result.aadhaarNo = aadhaarMatch[0].replace(/\s/g, "");
  }

  // DOB: DD/MM/YYYY or DD-MM-YYYY
  const dobMatch = text.match(/\d{2}[\/\-]\d{2}[\/\-]\d{4}/);
  if (dobMatch) result.dob = dobMatch[0];

  // Gender
  if (/\bMALE\b/i.test(text)) result.gender = "Male";
  else if (/\bFEMALE\b/i.test(text)) result.gender = "Female";

  // Name: usually the line after "Government of India" or before DOB line
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (
      /government of india/i.test(line) ||
      /भारत सरकार/.test(line) ||
      /unique identification/i.test(line)
    ) {
      // Name is typically 1-2 lines after this header
      if (lines[i + 1] && /^[A-Z][a-zA-Z\s]+$/.test(lines[i + 1])) {
        result.name = lines[i + 1];
      } else if (lines[i + 2] && /^[A-Z][a-zA-Z\s]+$/.test(lines[i + 2])) {
        result.name = lines[i + 2];
      }
      break;
    }
  }

  // Address: lines containing S/O, D/O, W/O, or PIN code
  const addrLines = [];
  let capturing = false;
  for (const line of lines) {
    if (/\b[SDWCsdwc]\/[Oo]\b/.test(line) || /\bVill\b|\bPost\b|\bDist\b/i.test(line)) {
      capturing = true;
    }
    if (capturing) {
      addrLines.push(line);
      if (/\b\d{6}\b/.test(line)) break; // PIN code signals end of address
    }
  }
  if (addrLines.length > 0) result.address = addrLines.join(", ");

  return result;
}

// ─── ROUTE: Scan Aadhaar card image ──────────────────────────────────────────
// POST /api/scan-aadhaar
// Body: multipart/form-data with field "card" (image file)
// Returns parsed Aadhaar details
app.post("/api/scan-aadhaar", upload.single("card"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No image uploaded" });
  }

  let processedPath = null;
  try {
    // Step 1: Preprocess image
    processedPath = await preprocessImage(req.file.path);

    // Step 2: Run OCR (English + Hindi)
    const { data } = await Tesseract.recognize(processedPath, "eng+hin", {
      logger: () => {}, // suppress logs
    });

    // Step 3: Parse the extracted text
    const parsed = parseAadhaarText(data.text);

    res.json({
      success: true,
      rawText: data.text, // send raw too, for debugging
      parsed,
    });
  } catch (err) {
    console.error("OCR error:", err);
    res.status(500).json({ error: "OCR failed", details: err.message });
  } finally {
    // Cleanup temp files
    if (req.file?.path) fs.unlink(req.file.path, () => {});
    if (processedPath) fs.unlink(processedPath, () => {});
  }
});

// ─── ROUTE: Submit check-in data to Google Sheets ────────────────────────────
// POST /api/checkin
// Body: JSON with guest + members + payment details
app.post("/api/checkin", async (req, res) => {
  const {
    roomNo,
    guest,        // { name, aadhaarNo, dob, gender, address }
    members,      // array of { name, aadhaarNo }
    payment,      // { type: "cash"|"online"|"half", cashAmount, onlineAmount, upiRef }
    remarks,
  } = req.body;

  // Basic validation
  if (!guest?.name || !guest?.aadhaarNo) {
    return res.status(400).json({ error: "Guest name and Aadhaar are required" });
  }
  if (!roomNo) {
    return res.status(400).json({ error: "Room number is required" });
  }

  try {
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: authClient });

    const sheetName = await getOrCreateMonthSheet(sheets);

    // Get current row count to generate Sr No
    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A:A`,
    });
    const srNo = (existing.data.values?.length || 1); // row 1 = header, so sr no = rows - 1

    const now = new Date();
    const date = now.toLocaleDateString("en-IN");
    const time = now.toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });

    const cashAmt = parseFloat(payment?.cashAmount || 0);
    const onlineAmt = parseFloat(payment?.onlineAmount || 0);
    const totalAmt = cashAmt + onlineAmt;

    // Build row — up to 5 members supported (guest + 4 extras)
    const extraMembers = members || [];
    const row = [
      srNo,
      date,
      time,
      roomNo,
      guest.name,
      guest.aadhaarNo,
      guest.dob || "",
      guest.gender || "",
      guest.address || "",
      1 + extraMembers.length,
      extraMembers[0]?.name || "",
      extraMembers[0]?.aadhaarNo || "",
      extraMembers[1]?.name || "",
      extraMembers[1]?.aadhaarNo || "",
      extraMembers[2]?.name || "",
      extraMembers[2]?.aadhaarNo || "",
      extraMembers[3]?.name || "",
      extraMembers[3]?.aadhaarNo || "",
      payment?.type || "cash",
      cashAmt || "",
      onlineAmt || "",
      totalAmt,
      payment?.upiRef || "",
      remarks || "",
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [row] },
    });

    res.json({ success: true, message: "Check-in saved!", srNo });
  } catch (err) {
    console.error("Sheets error:", err);
    res.status(500).json({ error: "Failed to save to Google Sheets", details: err.message });
  }
});

// ─── ROUTE: Get all records for a given month ─────────────────────────────────
// GET /api/records?month=June 2025
app.get("/api/records", async (req, res) => {
  const monthName = req.query.month;
  if (!monthName) {
    return res.status(400).json({ error: "month query param required" });
  }

  try {
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: authClient });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${monthName}!A1:Z1000`,
    });

    res.json({ success: true, data: response.data.values || [] });
  } catch (err) {
    res.status(500).json({ error: "Failed to fetch records", details: err.message });
  }
});

// ─── ROUTE: List all available months ─────────────────────────────────────────
// GET /api/months
app.get("/api/months", async (req, res) => {
  try {
    const authClient = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: authClient });

    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
    });

    const months = spreadsheet.data.sheets.map((s) => s.properties.title);
    res.json({ success: true, months });
  } catch (err) {
    res.status(500).json({ error: "Failed to fetch months", details: err.message });
  }
});

// ─── Start server ─────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Lodge backend running on port ${PORT}`);
  if (!fs.existsSync(path.join(__dirname, "credentials.json"))) {
    console.warn("⚠️  WARNING: credentials.json not found! Google Sheets will not work.");
  }
  if (!process.env.SPREADSHEET_ID) {
    console.warn("⚠️  WARNING: SPREADSHEET_ID not set in .env!");
  }
});
