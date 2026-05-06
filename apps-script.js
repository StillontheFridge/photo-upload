// ============================================================
// STILL ON THE FRIDGE — Photo Upload Receiver
// Google Apps Script Web App
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Go to script.google.com
// 2. Create a new project, name it "Still on the Fridge Upload"
// 3. Paste this entire file into the editor
// 4. Change FOLDER_NAME below if you want a different root folder name
// 5. Click Deploy > New Deployment
//    - Type: Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Click Deploy, authorize when prompted
//    (it will ask for Drive AND Gmail permissions — both are needed)
// 7. Copy the Web App URL — that goes into your customer webpage
// ============================================================

const FOLDER_NAME = "Still on the Fridge — Market Uploads";
const NOTIFY_EMAIL = "stillonthefridge@gmail.com";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const marketDate = data.marketDate || getTodayString();
    const orderNumber = data.orderNumber;
    const photos = data.photos; // array of { base64, filename, quantity }

    // Get or create root folder
    const rootFolder = getOrCreateFolder(FOLDER_NAME);

    // Get or create market day subfolder (e.g. "2026-05-06")
    const dayFolder = getOrCreateFolder(marketDate, rootFolder);

    // Get or create order subfolder (e.g. "Order-143022")
    const orderFolder = getOrCreateFolder("Order-" + orderNumber, dayFolder);

    // Save each photo
    for (let i = 0; i < photos.length; i++) {
      const photo = photos[i];
      const base64Data = photo.base64.split(",")[1]; // strip data:image/... prefix
      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64Data),
        "image/jpeg",
        photo.filename
      );
      const file = orderFolder.createFile(blob);
      file.setDescription("Quantity: " + photo.quantity);
    }

    // Write a summary text file for the order
    const summary = buildSummary(orderNumber, marketDate, photos);
    orderFolder.createFile("order-summary.txt", summary, MimeType.PLAIN_TEXT);

    // Send email notification
    sendNotification(orderNumber, marketDate, photos);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, orderNumber: orderNumber }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendNotification(orderNumber, marketDate, photos) {
  const totalMagnets = photos.reduce((sum, p) => sum + parseInt(p.quantity), 0);
  const photoWord = photos.length === 1 ? "photo" : "photos";
  const magnetWord = totalMagnets === 1 ? "magnet" : "magnets";

  const subject = `New order #${orderNumber} — ${totalMagnets} ${magnetWord}`;

  const photoLines = photos.map((p, i) =>
    `  Photo ${i + 1}: ${p.quantity} ${parseInt(p.quantity) === 1 ? "magnet" : "magnets"}`
  ).join("\n");

  const body = [
    `Still on the Fridge — New Order`,
    ``,
    `Order Number : #${orderNumber}`,
    `Market Date  : ${marketDate}`,
    ``,
    `${photos.length} ${photoWord}, ${totalMagnets} ${magnetWord} total`,
    ``,
    photoLines,
    ``,
    `Photos are saved in your Drive under:`,
    `${FOLDER_NAME} / ${marketDate} / Order-${orderNumber}`,
  ].join("\n");

  GmailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

function doGet(e) {
  // Health check endpoint
  return ContentService
    .createTextOutput(JSON.stringify({ status: "Still on the Fridge upload service is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder(name, parent) {
  const search = parent
    ? parent.getFoldersByName(name)
    : DriveApp.getFoldersByName(name);

  if (search.hasNext()) {
    return search.next();
  }

  return parent
    ? parent.createFolder(name)
    : DriveApp.createFolder(name);
}

function getTodayString() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function buildSummary(orderNumber, marketDate, photos) {
  const lines = [
    "STILL ON THE FRIDGE — ORDER SUMMARY",
    "=====================================",
    `Order Number : ${orderNumber}`,
    `Market Date  : ${marketDate}`,
    `Total Photos : ${photos.length}`,
    `Total Magnets: ${photos.reduce((sum, p) => sum + parseInt(p.quantity), 0)}`,
    "",
    "PHOTOS:",
  ];

  photos.forEach((p, i) => {
    lines.push(`  ${i + 1}. ${p.filename}  |  Qty: ${p.quantity}`);
  });

  return lines.join("\n");
}

