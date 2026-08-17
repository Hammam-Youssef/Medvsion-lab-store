/**
 * LabStore backend — Google Apps Script bound to your Google Sheet.
 *
 * SHEET STRUCTURE EXPECTED:
 *
 * Tab "Products": Department | Brand | Category | ProductName | CatalogNumber | Price | ImageURL
 *   - Category can be left blank (product sits directly under the Brand)
 *   - Brand can be left blank if a Department has no brand level
 *
 * Tab "Users": Username | Password | ClientName | DiscountPercent
 *
 * Tab "CustomPricing" (optional, create it if you need special pricing):
 *   Username | CatalogNumber | Price
 *   For a given user + product combo listed here, this exact Price is used
 *   instead of the normal Price × discount calculation. Any product not
 *   listed for that user just falls back to their normal DiscountPercent.
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

// Only this address receives order emails for now.
// Add more addresses to EMAIL_CC once you've confirmed test orders look right.
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
  sendOrderEmail(login.clientName, lines, grandTotal);

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
 * Guest checkout: no username/password. Requires a name, phone, and
 * detailed address instead — all three must be filled in or the order
 * is rejected. Prices are the plain catalog Price (no discount, since
 * there's no client account to look up a discount for).
 */
function handleGuestOrder(body) {
  const name = String(body.guestName || "").trim();
  const phone = String(body.guestPhone || "").trim();
  const address = String(body.guestAddress || "").trim();

  if (!name || !phone || !address) {
    return { success: false, error: "Please fill in your name, phone number, and address." };
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

  logOrder(name, "GUEST", lines, grandTotal, phone, address);
  sendOrderEmail(name, lines, grandTotal, phone, address);

  return { success: true, grandTotal: grandTotal };
}

function logOrder(clientName, username, lines, grandTotal, guestPhone, guestAddress) {
  let sheet = getSheet(ORDERS_SHEET);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(ORDERS_SHEET);
    sheet.appendRow([
      "Timestamp", "Username", "ClientName", "ProductName",
      "CatalogNumber", "Qty", "UnitPrice", "LineTotal", "OrderTotal",
      "Phone", "Address"
    ]);
  }
  const timestamp = new Date();
  lines.forEach(l => {
    sheet.appendRow([
      timestamp, username, clientName, l.productName,
      l.catalogNumber, l.qty, l.unitPrice.toFixed(2),
      l.lineTotal.toFixed(2), grandTotal.toFixed(2),
      guestPhone || "", guestAddress || ""
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

function sendOrderEmail(clientName, lines, grandTotal, guestPhone, guestAddress) {
  const rowsHtml = lines.map(l => `
    <tr>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;">${l.productName}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;">${l.catalogNumber}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:center;">${l.qty}</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:right;">${l.unitPrice.toFixed(2)} KWD</td>
      <td style="padding:6px 10px;border-bottom:1px solid #e2e2e2;text-align:right;">${l.lineTotal.toFixed(2)} KWD</td>
    </tr>`).join("");

  const contactHtml = (guestPhone || guestAddress) ? `
      <p style="color:#333;margin:4px 0;"><strong>Phone:</strong> ${guestPhone || "—"}</p>
      <p style="color:#333;margin:4px 0 14px;"><strong>Address:</strong> ${guestAddress || "—"}</p>
  ` : "";

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:640px;">
      <h2 style="margin-bottom:2px;">New Order — ${clientName}</h2>
      <p style="color:#666;margin-top:0;">Submitted ${new Date().toLocaleString()}</p>
      ${contactHtml}
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

  const options = { htmlBody: html };
  if (EMAIL_CC.length) options.cc = EMAIL_CC.join(",");

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: `New LabStore Order — ${clientName}`,
    htmlBody: html,
    cc: options.cc
  });
}
