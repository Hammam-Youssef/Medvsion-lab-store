/**
 * LabStore backend — Google Apps Script bound to your Google Sheet.
 *
 * SHEET STRUCTURE EXPECTED:
 *
 * Tab "Products": Department | Brand | Category | ProductName | CatalogNumber | Price | ImageURL | Audience | Stock
 *   - Category can be left blank (product sits directly under the Brand)
 *   - Brand can be left blank if a Department has no brand level
 *   - Audience: set to "Public" for products guests (Specialty Departments) can see/order
 *   - Stock: optional. Leave blank if you don't track stock for a product.
 *
 * Tab "Users": Username | Password | ClientName | DiscountPercent | Role
 *
 * Tab "CustomPricing" (optional, create it if you need special pricing):
 *   Username | CatalogNumber | Price
 *   For a given user + product combo listed here, this exact Price is used
 *   instead of the normal Price × discount calculation. Any product not
 *   listed for that user just falls back to their normal DiscountPercent.
 *
 * Tab "DepartmentEmails" (optional — create it to route Guest/Specialty
 * orders to the right purchasing contact automatically):
 *   Department | Email
 *   When a guest order contains items from a given Department, that part
 *   of the order is emailed to the address listed here instead of the
 *   default EMAIL_TO. If a Department has no row here, it falls back to
 *   EMAIL_TO. To add a new department, just add one row — no code changes.
 *   NOTE: this only affects Guest/"Specialty Departments" orders. Lab
 *   (logged-in) orders always go to EMAIL_TO, unchanged.
 *
 * Tab "Orders" (created automatically if missing): log of every submitted order line
 *
 * SETUP:
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Delete any starter code, paste this whole file in
 * 3. Update EMAIL_TO / EMAIL_CC below
 * 4. Click Deploy → New deployment → type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Copy the deployment URL — paste it into APPS_SCRIPT_URL in the web app's app.js
 */

const PRODUCTS_SHEET = "Products";
const USERS_SHEET = "Users";
const ORDERS_SHEET = "Orders";
const CUSTOM_PRICING_SHEET = "CustomPricing";
const DEPARTMENT_EMAILS_SHEET = "DepartmentEmails";

// Default/fallback address for order emails — used for Lab orders always,
// and for Guest/Specialty orders whose Department isn't listed in
// DepartmentEmails.
const EMAIL_TO = "h.yousef@medvision-kw.com";
const EMAIL_CC = []; // e.g. ["someone@medvision-kw.com"]

function doGet(e) {
  try {
    const action = e.parameter.action || "products";
    if (action === "products") {
      // Public/unauthenticated endpoint — only ever returns products
      // explicitly marked Audience=Public (the guest-facing specialty catalog).
      return jsonOut({ success: true, products: getProducts().filter(isPublicProduct) });
    }
    return jsonOut({ success: false, error: "Unknown action" });
  } catch (err) {
    return jsonOut({ success: false, error: String(err) });
  }
}

function isPublicProduct(p) {
  return String(p.Audience || "").trim().toLowerCase() === "public";
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === "login") {
      const login = handleLogin(body.username, body.password);
      if (!login.success) return jsonOut(login);
      const products = getPricedProductsForUser(body.username, login.discount);
      return jsonOut(Object.assign({}, login, { products: products }));
    }
    if (body.action === "submitOrder") {
      return jsonOut(handleSubmitOrder(body));
    }
    if (body.action === "adminStats") {
      return jsonOut(handleAdminStats(body.username, body.password));
    }
    if (body.action === "guestOrder") {
      return jsonOut(handleGuestOrder(body));
    }
    return jsonOut({ success: false, error: "Unknown action" });
  } catch (err) {
    return jsonOut({ success: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getProducts() {
  const sheet = getSheet(PRODUCTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1).filter(r => r[headers.indexOf("ProductName")]);
  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = r[i]));
    return obj;
  });
}

function getUsers() {
  const sheet = getSheet(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = r[i]));
    return obj;
  });
}

/**
 * Reads the optional CustomPricing sheet into a lookup map keyed by
 * "username||catalognumber" (lowercased username). Returns {} if the
 * sheet doesn't exist yet — custom pricing is entirely opt-in.
 */
function getCustomPricingMap() {
  const sheet = getSheet(CUSTOM_PRICING_SHEET);
  const map = {};
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  const headers = data[0].map(h => String(h).trim());
  const uIdx = headers.indexOf("Username");
  const cIdx = headers.indexOf("CatalogNumber");
  const pIdx = headers.indexOf("Price");
  if (uIdx === -1 || cIdx === -1 || pIdx === -1) return map;
  data.slice(1).forEach(row => {
    if (!row[uIdx] || !row[cIdx] || row[pIdx] === "") return;
    const key = String(row[uIdx]).trim().toLowerCase() + "||" + String(row[cIdx]).trim();
    map[key] = Number(row[pIdx]);
  });
  return map;
}

/**
 * Reads the optional DepartmentEmails sheet into a lookup map keyed by
 * Department name. Returns {} if the sheet doesn't exist yet — routing
 * is entirely opt-in and falls back to EMAIL_TO per department.
 */
function getDepartmentEmailMap() {
  const sheet = getSheet(DEPARTMENT_EMAILS_SHEET);
  const map = {};
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  const headers = data[0].map(h => String(h).trim());
  const dIdx = headers.indexOf("Department");
  const eIdx = headers.indexOf("Email");
  if (dIdx === -1 || eIdx === -1) return map;
  data.slice(1).forEach(row => {
    if (!row[dIdx] || !row[eIdx]) return;
    map[String(row[dIdx]).trim()] = String(row[eIdx]).trim();
  });
  return map;
}

/**
 * For a given username, returns the full product list with an
 * EffectivePrice on each item: their CustomPricing override if one
 * exists for that product, otherwise Price × (1 − their discount).
 */
function getPricedProductsForUser(username, discount) {
  const products = getProducts().filter(p => !isPublicProduct(p));
  const customMap = getCustomPricingMap();
  const uname = String(username).trim().toLowerCase();
  return products.map(p => {
    const key = uname + "||" + String(p.CatalogNumber).trim();
    const hasCustom = customMap.hasOwnProperty(key);
    const effectivePrice = hasCustom ? customMap[key] : Number(p.Price) * (1 - discount / 100);
    return Object.assign({}, p, {
      EffectivePrice: Number(effectivePrice.toFixed(3)),
      HasCustomPrice: hasCustom
    });
  });
}

function handleLogin(username, password) {
  const users = getUsers();
  const match = users.find(
    u => String(u.Username).trim() === String(username).trim() &&
         String(u.Password) === String(password)
  );
  if (!match) {
    return { success: false, error: "Incorrect username or password." };
  }
  const isAdmin = String(match.Role || "").trim().toLowerCase() === "admin" ||
                  String(match.Username).trim().toLowerCase() === "labadmin";
  return {
    success: true,
    clientName: match.ClientName,
    discount: Number(match.DiscountPercent) || 0,
    isAdmin: isAdmin
  };
}

/**
 * Order submission: re-validates the user and re-computes every price
 * from the sheet itself, rather than trusting numbers sent from the app.
 * body = { username, password, items: [{ catalogNumber, qty }] }
 */
function handleSubmitOrder(body) {
  const login = handleLogin(body.username, body.password);
  if (!login.success) return login;

  const { sheet: productsSheet, headers, map } = getProductsIndexed();
  const stockColIdx = headers.indexOf("Stock"); // -1 if the sheet has no Stock column yet
  const customMap = getCustomPricingMap();
  const uname = String(body.username).trim().toLowerCase();

  const lines = [];
  const stockErrors = [];
  let grandTotal = 0;

  (body.items || []).forEach(item => {
    const p = map[String(item.catalogNumber).trim()];
    if (!p || !item.qty || item.qty <= 0) return;

    if (stockColIdx !== -1 && p.Stock !== "" && p.Stock !== undefined) {
      const available = Number(p.Stock);
      if (!isNaN(available) && item.qty > available) {
        stockErrors.push(`${p.ProductName} (only ${available} left)`);
        return;
      }
    }

    const priceKey = uname + "||" + String(p.CatalogNumber).trim();
    const unitPrice = customMap.hasOwnProperty(priceKey)
      ? customMap[priceKey]
      : Number(p.Price) * (1 - login.discount / 100);
    const lineTotal = unitPrice * item.qty;
    grandTotal += lineTotal;
    lines.push({
      department: p.Department || "Unassigned",
      productName: p.ProductName,
      catalogNumber: p.CatalogNumber,
      qty: item.qty,
      unitPrice: unitPrice,
      lineTotal: lineTotal,
      _row: p._row
    });
  });

  if (stockErrors.length) {
    return { success: false, error: "Not enough stock for: " + stockErrors.join(", ") };
  }
  if (lines.length === 0) {
    return { success: false, error: "Order has no valid items." };
  }

  if (stockColIdx !== -1) {
    lines.forEach(l => {
      const cell = productsSheet.getRange(l._row, stockColIdx + 1);
      const current = Number(cell.getValue());
      if (!isNaN(current)) cell.setValue(Math.max(0, current - l.qty));
    });
  }

  logOrder(login.clientName, body.username, lines, grandTotal);
  sendLabOrderEmail(login.clientName, lines, grandTotal);

  return { success: true, grandTotal: grandTotal };
}

/**
 * Reads the Products sheet and returns it indexed by CatalogNumber,
 * with each product carrying its real sheet row number (_row) so stock
 * levels can be updated in place after an order.
 */
function getProductsIndexed() {
  const sheet = getSheet(PRODUCTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const nameIdx = headers.indexOf("ProductName");
  const catalogIdx = headers.indexOf("CatalogNumber");
  const map = {};
  data.slice(1).forEach((row, i) => {
    if (!row[nameIdx]) return;
    const obj = {};
    headers.forEach((h, idx) => (obj[h] = row[idx]));
    obj._row = i + 2; // +2: 1 for header row, 1 because arrays are 0-indexed
    map[String(row[catalogIdx]).trim()] = obj;
  });
  return { sheet, headers, map };
}

/**
 * Guest checkout ("Specialty Departments"): no username/password.
 * Requires name, phone, and a full broken-out address (city, area, block,
 * street, building number) — all fields must be filled in or the order is
 * rejected. Prices are the plain catalog Price (no discount, since there's
 * no client account to look up a discount for).
 *
 * The order is split by each product's Department and one email is sent
 * per department to whoever is listed for it in DepartmentEmails (falling
 * back to EMAIL_TO if that department isn't listed there yet).
 *
 * body = { guestName, guestPhone, city, area, block, street, buildingNumber,
 *          items: [{ catalogNumber, qty }] }
 */
function handleGuestOrder(body) {
  const name = String(body.guestName || "").trim();
  const phone = String(body.guestPhone || "").trim();
  const city = String(body.city || "").trim();
  const area = String(body.area || "").trim();
  const block = String(body.block || "").trim();
  const street = String(body.street || "").trim();
  const buildingNumber = String(body.buildingNumber || "").trim();

  if (!name || !phone || !city || !area || !block || !street || !buildingNumber) {
    return { success: false, error: "Please fill in all required fields." };
  }

  const { sheet: productsSheet, headers, map } = getProductsIndexed();
  const stockColIdx = headers.indexOf("Stock");

  const lines = [];
  const stockErrors = [];
  let grandTotal = 0;

  (body.items || []).forEach(item => {
    const p = map[String(item.catalogNumber).trim()];
    if (!p || !isPublicProduct(p) || !item.qty || item.qty <= 0) return;

    if (stockColIdx !== -1 && p.Stock !== "" && p.Stock !== undefined) {
      const available = Number(p.Stock);
      if (!isNaN(available) && item.qty > available) {
        stockErrors.push(`${p.ProductName} (only ${available} left)`);
        return;
      }
    }

    const unitPrice = Number(p.Price);
    const lineTotal = unitPrice * item.qty;
    grandTotal += lineTotal;
    lines.push({
      department: p.Department || "Unassigned",
      productName: p.ProductName,
      catalogNumber: p.CatalogNumber,
      qty: item.qty,
      unitPrice: unitPrice,
      lineTotal: lineTotal,
      _row: p._row
    });
  });

  if (stockErrors.length) {
    return { success: false, error: "Not enough stock for: " + stockErrors.join(", ") };
  }
  if (lines.length === 0) {
    return { success: false, error: "Order has no valid items." };
  }

  if (stockColIdx !== -1) {
    lines.forEach(l => {
      const cell = productsSheet.getRange(l._row, stockColIdx + 1);
      const current = Number(cell.getValue());
      if (!isNaN(current)) cell.setValue(Math.max(0, current - l.qty));
    });
  }

  const guestInfo = { phone, city, area, block, street, buildingNumber };

  logOrder(name, "GUEST", lines, grandTotal, guestInfo);

  // Split the order by Department and email each department's contact
  // separately, so an order mixing e.g. Radiology + Dental supplies
  // reaches both purchasing owners automatically.
  const emailMap = getDepartmentEmailMap();
  const byDept = {};
  lines.forEach(l => {
    if (!byDept[l.department]) byDept[l.department] = [];
    byDept[l.department].push(l);
  });
  Object.keys(byDept).forEach(dept => {
    const deptLines = byDept[dept];
    const subtotal = deptLines.reduce((s, l) => s + l.lineTotal, 0);
    const toEmail = emailMap[dept] || EMAIL_TO;
    sendGuestDepartmentEmail(dept, toEmail, name, guestInfo, deptLines, subtotal);
  });

  return { success: true, grandTotal: grandTotal };
}

/**
 * Appends one row per order line to the Orders sheet.
 * guestInfo is omitted (undefined) for Lab orders.
 */
function logOrder(clientName, username, lines, grandTotal, guestInfo) {
  let sheet = getSheet(ORDERS_SHEET);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(ORDERS_SHEET);
    sheet.appendRow([
      "Timestamp", "Username", "ClientName", "Department", "ProductName",
      "CatalogNumber", "Qty", "UnitPrice", "LineTotal", "OrderTotal",
      "Phone", "City", "Area", "Block", "Street", "BuildingNumber"
    ]);
  }
  const timestamp = new Date();
  const g = guestInfo || {};
  lines.forEach(l => {
    sheet.appendRow([
      timestamp, username, clientName, l.department || "", l.productName,
      l.catalogNumber, l.qty, l.unitPrice.toFixed(2),
      l.lineTotal.toFixed(2), grandTotal.toFixed(2),
      g.phone || "", g.city || "", g.area || "", g.block || "", g.street || "", g.buildingNumber || ""
    ]);
  });
}

/**
 * Admin-only: aggregates the Orders sheet into monthly totals and
 * per-customer-per-month totals. Requires the logging-in user to be an admin.
 */
function handleAdminStats(username, password) {
  const login = handleLogin(username, password);
  if (!login.success) return login;
  if (!login.isAdmin) return { success: false, error: "Not authorized." };

  const sheet = getSheet(ORDERS_SHEET);
  if (!sheet) return { success: true, monthly: [], byCustomer: [] };

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);
  const idx = name => headers.indexOf(name);

  const monthlyMap = {}; // "YYYY-MM" -> total
  const customerMap = {}; // "ClientName||YYYY-MM" -> total

  rows.forEach(r => {
    const ts = r[idx("Timestamp")];
    if (!ts) return;
    const d = new Date(ts);
    const month = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM");
    const clientName = r[idx("ClientName")];
    const lineTotal = Number(r[idx("LineTotal")]) || 0;

    monthlyMap[month] = (monthlyMap[month] || 0) + lineTotal;
    const key = clientName + "||" + month;
    customerMap[key] = (customerMap[key] || 0) + lineTotal;
  });

  const monthly = Object.keys(monthlyMap).sort().reverse().map(month => ({
    month: month, total: Number(monthlyMap[month].toFixed(2))
  }));

  const byCustomer = Object.keys(customerMap).sort().reverse().map(key => {
    const [clientName, month] = key.split("||");
    return { clientName: clientName, month: month, total: Number(customerMap[key].toFixed(2)) };
  });

  return { success: true, monthly: monthly, byCustomer: byCustomer };
}

/**
 * Lab (logged-in) order confirmation — always goes to EMAIL_TO / EMAIL_CC,
 * unaffected by DepartmentEmails.
 */
function sendLabOrderEmail(clientName, lines, grandTotal) {
  const rowsHtml = lines.map(l => `
    <tr>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;">${l.productName}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;">${l.catalogNumber}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:center;">${l.qty}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:right;">${l.unitPrice.toFixed(2)} KWD</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:right;">${l.lineTotal.toFixed(2)} KWD</td>
    </tr>`).join("");

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:640px;">
      <h2 style="margin-bottom:2px;">New Order — ${clientName}</h2>
      <p style="color:#666;margin-top:0;">Submitted ${new Date().toLocaleString()}</p>
      <table style="border-collapse:collapse;width:100%;font-size:14px;">
        <thead>
          <tr style="background:#f2f2f2;text-align:left;">
            <th style="padding:6px 10px;">Product</th>
            <th style="padding:6px 10px;">Catalog #</th>
            <th style="padding:6px 10px;text-align:center;">Qty</th>
            <th style="padding:6px 10px;text-align:right;">Unit Price</th>
            <th style="padding:6px 10px;text-align:right;">Line Total</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <p style="font-size:16px;margin-top:14px;"><strong>Grand Total: ${grandTotal.toFixed(2)} KWD</strong></p>
    </div>`;

  const options = {};
  if (EMAIL_CC.length) options.cc = EMAIL_CC.join(",");

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: `New LabStore Order — ${clientName}`,
    htmlBody: html,
    cc: options.cc
  });
}

/**
 * Guest/Specialty department order confirmation — one call per Department
 * present in the order, sent to that department's mapped email (or
 * EMAIL_TO as a fallback if the department isn't in DepartmentEmails yet).
 */
function sendGuestDepartmentEmail(department, toEmail, guestName, guestInfo, lines, subtotal) {
  const rowsHtml = lines.map(l => `
    <tr>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;">${l.productName}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;">${l.catalogNumber}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:center;">${l.qty}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:right;">${l.unitPrice.toFixed(2)} KWD</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:right;">${l.lineTotal.toFixed(2)} KWD</td>
    </tr>`).join("");

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:640px;">
      <h2 style="margin-bottom:2px;">New ${department} Order — ${guestName}</h2>
      <p style="color:#666;margin-top:0;">Submitted ${new Date().toLocaleString()}</p>
      <p style="color:#333;margin:4px 0;"><strong>Phone:</strong> ${guestInfo.phone || "—"}</p>
      <p style="color:#333;margin:4px 0 14px;">
        <strong>Address:</strong>
        City: ${guestInfo.city || "—"} · Area: ${guestInfo.area || "—"} · Block: ${guestInfo.block || "—"} ·
        Street: ${guestInfo.street || "—"} · Building No.: ${guestInfo.buildingNumber || "—"}
      </p>
      <table style="border-collapse:collapse;width:100%;font-size:14px;">
        <thead>
          <tr style="background:#f2f2f2;text-align:left;">
            <th style="padding:6px 10px;">Product</th>
            <th style="padding:6px 10px;">Catalog #</th>
            <th style="padding:6px 10px;text-align:center;">Qty</th>
            <th style="padding:6px 10px;text-align:right;">Unit Price</th>
            <th style="padding:6px 10px;text-align:right;">Line Total</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <p style="font-size:16px;margin-top:14px;"><strong>Total: ${subtotal.toFixed(2)} KWD</strong></p>
    </div>`;

  MailApp.sendEmail({
    to: toEmail,
    subject: `New ${department} Order — ${guestName}`,
    htmlBody: html
  });
}
