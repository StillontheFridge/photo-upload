// ============================================================
// STILL ON THE FRIDGE — Photo Upload Receiver
// Google Apps Script Web App
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Go to script.google.com
// 2. Create a new project, name it "Still on the Fridge Upload"
// 3. Paste this entire file into the editor
// 4. Click Deploy > New Deployment
//    - Type: Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Click Deploy, authorize when prompted
//    (it will ask for Drive, Sheets, AND Gmail permissions)
// 6. Copy the Web App URL and paste it into index.html
//    where it says YOUR_APPS_SCRIPT_URL_HERE
// ============================================================

const FOLDER_NAME   = "Still on the Fridge — Market Uploads";
const SHEET_NAME    = "Still on the Fridge — Orders";
const NOTIFY_EMAIL  = "stillonthefridge@gmail.com";

function doPost(e) {
  try {
    const data        = JSON.parse(e.postData.contents);
    const marketDate  = data.marketDate || getTodayString();
    const orderNumber = data.orderNumber;
    const name        = data.name  || "";
    const phone       = data.phone || "";
    const email       = data.email || "";
    const photos      = data.photos; // [{ base64, filename, quantity }]

    // ── Save photos to Drive ──
    const rootFolder  = getOrCreateFolder(FOLDER_NAME);
    const dayFolder   = getOrCreateFolder(marketDate, rootFolder);
    const orderFolder = getOrCreateFolder("Order-" + orderNumber, dayFolder);

    for (let i = 0; i < photos.length; i++) {
      const photo      = photos[i];
      const base64Data = photo.base64.split(",")[1];
      const blob       = Utilities.newBlob(
        Utilities.base64Decode(base64Data),
        "image/jpeg",
        photo.filename   // e.g. photo-1_qty3.jpg
      );
      const file = orderFolder.createFile(blob);
      file.setDescription("Quantity: " + photo.quantity);
    }

    // Write order summary text file
    const summary = buildSummary(orderNumber, marketDate, name, phone, email, photos);
    orderFolder.createFile("order-summary.txt", summary, MimeType.PLAIN_TEXT);

    // ── Log to Google Sheet ──
    logToSheet(orderNumber, marketDate, name, phone, email, photos);

    // ── Send email notification ──
    sendNotification(orderNumber, marketDate, name, phone, email, photos);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, orderNumber }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Google Sheets logging ──
function logToSheet(orderNumber, marketDate, name, phone, email, photos) {
  let sheet;

  // Find or create the spreadsheet
  const files = DriveApp.getFilesByName(SHEET_NAME);
  let spreadsheet;
  if (files.hasNext()) {
    spreadsheet = SpreadsheetApp.open(files.next());
  } else {
    spreadsheet = SpreadsheetApp.create(SHEET_NAME);
  }

  // Find or create the sheet tab for this market date
  sheet = spreadsheet.getSheetByName(marketDate);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(marketDate);
    // Write header row
    sheet.appendRow([
      "Order #",
      "Time",
      "Name",
      "Phone",
      "Email",
      "Photos Summary",
      "Total Magnets",
      "Date"
    ]);
    // Bold the header
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  // Build photos summary string e.g. "Photo 1: 3, Photo 2: 1, Photo 3: 2"
  const photosSummary = photos
    .map((p, i) => `Photo ${i + 1}: ${p.quantity}`)
    .join(", ");

  const totalMagnets = photos.reduce((sum, p) => sum + parseInt(p.quantity), 0);

  // Format time from order number (HHMMSS)
  const timeFormatted = formatOrderTime(orderNumber);

  sheet.appendRow([
    "#" + orderNumber,
    timeFormatted,
    name,
    phone,
    email,
    photosSummary,
    totalMagnets,
    marketDate
  ]);

  // Auto-resize columns for readability
  sheet.autoResizeColumns(1, 8);
}

// ── Email notification ──
function sendNotification(orderNumber, marketDate, name, phone, email, photos) {
  const totalMagnets  = photos.reduce((sum, p) => sum + parseInt(p.quantity), 0);
  const magnetWord    = totalMagnets === 1 ? "magnet" : "magnets";
  const subject       = `New order #${orderNumber} — ${name} — ${totalMagnets} ${magnetWord}`;

  const photoLines = photos
    .map((p, i) => `  Photo ${i + 1}: ${p.quantity} ${parseInt(p.quantity) === 1 ? "magnet" : "magnets"}  (${p.filename})`)
    .join("\n");

  const body = [
    "STILL ON THE FRIDGE — NEW ORDER",
    "=================================",
    "",
    `Order Number : #${orderNumber}`,
    `Time         : ${formatOrderTime(orderNumber)}`,
    `Market Date  : ${marketDate}`,
    "",
    `Name         : ${name}`,
    `Phone        : ${phone}`,
    `Email        : ${email || "(not provided)"}`,
    "",
    `Photos: ${photos.length}   Total Magnets: ${totalMagnets}`,
    "",
    photoLines,
    "",
    "Drive location:",
    `${FOLDER_NAME} / ${marketDate} / Order-${orderNumber}`,
  ].join("\n");

  GmailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

// ── Helpers ──
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "Still on the Fridge upload service is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder(name, parent) {
  const search = parent
    ? parent.getFoldersByName(name)
    : DriveApp.getFoldersByName(name);
  if (search.hasNext()) return search.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function getTodayString() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

function formatOrderTime(orderNumber) {
  // orderNumber is HHMMSS
  if (orderNumber.length === 6) {
    const h   = parseInt(orderNumber.substring(0, 2));
    const min = orderNumber.substring(2, 4);
    const sec = orderNumber.substring(4, 6);
    const ampm = h >= 12 ? "pm" : "am";
    const h12  = h % 12 === 0 ? 12 : h % 12;
    return `${h12}:${min}:${sec} ${ampm}`;
  }
  return orderNumber;
}

function buildSummary(orderNumber, marketDate, name, phone, email, photos) {
  const totalMagnets = photos.reduce((sum, p) => sum + parseInt(p.quantity), 0);
  const lines = [
    "STILL ON THE FRIDGE — ORDER SUMMARY",
    "=====================================",
    `Order Number : #${orderNumber}`,
    `Time         : ${formatOrderTime(orderNumber)}`,
    `Market Date  : ${marketDate}`,
    "",
    `Name         : ${name}`,
    `Phone        : ${phone}`,
    `Email        : ${email || "(not provided)"}`,
    "",
    `Total Photos : ${photos.length}`,
    `Total Magnets: ${totalMagnets}`,
    "",
    "PHOTOS:",
  ];
  photos.forEach((p, i) => {
    lines.push(`  ${i + 1}. ${p.filename}  |  Qty: ${p.quantity}`);
  });
  return lines.join("\n");
}
