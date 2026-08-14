// ====== CONFIG ======
// Paste your Google Apps Script Web App deployment URL here (ends in /exec)
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxp9fBnMDYgOF0wjG_Zzeh1U2bsP6rGolO0--IpcP_JQU4x3_IqoSvmSt-J2XbugRdc/exec";

// ====== ICONS (inline SVG, no external deps) ======
const ICON_FLASK = `<svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M9 2H15" stroke="#fff" stroke-width="1.8" stroke-linecap="round"/><path d="M10 2V8.5L4.5 18C3.8 19.3 4.7 21 6.2 21H17.8C19.3 21 20.2 19.3 19.5 18L14 8.5V2" stroke="#fff" stroke-width="1.8" stroke-linejoin="round"/><path d="M7.5 15H16.5" stroke="#D9A441" stroke-width="1.8"/></svg>`;
const ICON_CART = `<svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" width="18" height="18"><path d="M3 4H5L6.5 15.5C6.6 16.3 7.3 17 8.2 17H17.5C18.4 17 19.1 16.4 19.3 15.5L21 7H6" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/><circle cx="9" cy="20" r="1.4" fill="#fff"/><circle cx="17" cy="20" r="1.4" fill="#fff"/></svg>`;

// Tries to show logo.png (upload your own file with this name next to index.html).
// Falls back to the flask icon automatically if logo.png isn't there yet.
const BRAND_MARK = `<img src="logo.png" alt="Medvision" onerror="this.outerHTML='${ICON_FLASK.replace(/'/g, "\\'")}'" />`;

// ====== STATE ======
let session = JSON.parse(localStorage.getItem("labstore_session") || "null");
let products = [];
let path = [];
let cart = {}; // catalogNumber -> qty
let view = "loading"; // loading | login | browse | summary | admin
let loadError = "";
let searchQuery = "";
let adminStats = null; // { monthly, byCustomer }

const root = document.getElementById("app");

// ====== API HELPERS ======
async function apiGet(action) {
  const res = await fetch(`${APPS_SCRIPT_URL}?action=${action}`);
  return res.json();
}
async function apiPost(body) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" }, // avoids CORS preflight on Apps Script
    body: JSON.stringify(body)
  });
  return res.json();
}

// ====== BOOT ======
async function boot() {
  if (APPS_SCRIPT_URL.includes("PASTE_YOUR")) {
    loadError = "Configuration needed: set APPS_SCRIPT_URL in app.js to your deployed Apps Script web app URL.";
    view = "login";
    render();
    return;
  }
  if (session) {
    // Re-validate the cached session and refresh pricing (discount/custom pricing may have changed).
    try {
      const res = await apiPost({ action: "login", username: session.username, password: session.password });
      if (!res.success) throw new Error(res.error || "Session expired");
      session = { username: session.username, password: session.password, clientName: res.clientName, discount: res.discount, isAdmin: !!res.isAdmin };
      localStorage.setItem("labstore_session", JSON.stringify(session));
      products = res.products || [];
      view = "browse";
    } catch (err) {
      session = null;
      localStorage.removeItem("labstore_session");
      view = "login";
    }
  } else {
    view = "login";
  }
  render();
}

// ====== TREE HELPERS ======
function uniq(arr) { return [...new Set(arr)]; }

function currentLevel() {
  if (searchQuery.trim()) {
    const q = searchQuery.trim().toLowerCase();
    const matches = products.filter(p =>
      String(p.ProductName).toLowerCase().includes(q) ||
      String(p.CatalogNumber).toLowerCase().includes(q) ||
      String(p.Brand).toLowerCase().includes(q) ||
      String(p.Category).toLowerCase().includes(q)
    );
    return { type: "product", items: matches, isSearch: true };
  }
  const filtered = products.filter((p, _, __) =>
    path.every((val, i) => [p.Department, p.Brand, p.Category][i] === val)
  );
  if (path.length === 0) return { type: "dept", items: uniq(filtered.map(p => p.Department)) };
  if (path.length === 1) {
    const brands = uniq(filtered.filter(p => p.Brand).map(p => p.Brand));
    if (brands.length) return { type: "brand", items: brands };
    return { type: "product", items: filtered };
  }
  if (path.length === 2) {
    const cats = uniq(filtered.filter(p => p.Category).map(p => p.Category));
    if (cats.length) return { type: "cat", items: cats };
    return { type: "product", items: filtered };
  }
  return { type: "product", items: filtered };
}

function cartCount() { return Object.values(cart).reduce((a, b) => a + b, 0); }

function cartTotal() {
  return Object.entries(cart).reduce((sum, [cat, qty]) => {
    const p = products.find(pr => String(pr.CatalogNumber) === cat);
    if (!p || qty <= 0) return sum;
    return sum + Number(p.EffectivePrice) * qty;
  }, 0);
}

// ====== RENDER ======
function render() {
  if (view === "loading") return renderLoading();
  if (view === "login") return renderLogin();
  if (view === "browse") return renderBrowse();
  if (view === "summary") return renderSummary();
  if (view === "admin") return renderAdmin();
}

function renderLoading() {
  root.innerHTML = `<div class="loading-state">Loading store…</div>`;
}

function renderLogin() {
  root.innerHTML = `
    <div class="login-wrap">
      <div class="login-card">
        <div class="login-flask">${BRAND_MARK}</div>
        <p class="login-title">Medvision Lab Store</p>
        <p class="login-sub">Sign in with your client account</p>
      </div>
      <div class="login-form">
        <input id="username" type="text" placeholder="Username" autocomplete="username" />
        <input id="password" type="password" placeholder="Password" autocomplete="current-password" />
        <button class="btn-primary" id="loginBtn">Sign in</button>
        <p class="login-error" id="loginError"></p>
        <div class="login-divider"><span>or</span></div>
        <button class="btn-secondary" id="guestBtn">Continue as Guest</button>
      </div>
    </div>`;
  if (loadError) {
    const err = document.getElementById("loginError");
    err.textContent = loadError;
    err.style.display = "block";
  }
  document.getElementById("loginBtn").onclick = doLogin;
  document.getElementById("guestBtn").onclick = enterGuestMode;
  document.getElementById("password").addEventListener("keydown", e => { if (e.key === "Enter") doLogin(); });
}

async function enterGuestMode() {
  const btn = document.getElementById("guestBtn");
  btn.disabled = true; btn.textContent = "Loading…";
  try {
    const res = await apiGet("products");
    if (!res.success) throw new Error(res.error || "Failed to load products");
    // Guests see plain catalog prices — normalize EffectivePrice so all
    // the shared rendering code (cards, cart, summary) works unchanged.
    products = res.products.map(p => Object.assign({}, p, { EffectivePrice: Number(p.Price), HasCustomPrice: false }));
    session = { isGuest: true, clientName: "Guest", discount: 0, isAdmin: false };
    view = "browse";
    render();
  } catch (err) {
    btn.disabled = false; btn.textContent = "Continue as Guest";
    loadError = "Could not load the store. Please try again.";
    render();
  }
}

async function doLogin() {
  const username = document.getElementById("username").value.trim();
  const password = document.getElementById("password").value;
  const errEl = document.getElementById("loginError");
  const btn = document.getElementById("loginBtn");
  if (!username || !password) return;
  btn.disabled = true; btn.textContent = "Signing in…";
  try {
    const res = await apiPost({ action: "login", username, password });
    if (!res.success) {
      errEl.textContent = res.error || "Incorrect username or password.";
      errEl.style.display = "block";
      btn.disabled = false; btn.textContent = "Sign in";
      return;
    }
    session = { username, password, clientName: res.clientName, discount: res.discount, isAdmin: !!res.isAdmin };
    localStorage.setItem("labstore_session", JSON.stringify(session));
    products = res.products || [];
    view = "browse";
    render();
  } catch (err) {
    errEl.textContent = "Something went wrong. Please try again.";
    errEl.style.display = "block";
    btn.disabled = false; btn.textContent = "Sign in";
  }
}

function logout() {
  session = null; cart = {}; path = [];
  localStorage.removeItem("labstore_session");
  view = "login";
  render();
}

function renderBrowse() {
  const level = currentLevel();
  const crumbs = searchQuery.trim() ? ["Store", `Search: "${searchQuery.trim()}"`] : ["Store", ...path];

  let html = `
    <div class="topbar">
      <div class="brand-mark">${BRAND_MARK}</div>
      <p class="brand-name">Medvision Lab Store</p>
      <div class="topbar-spacer"></div>
      ${session.isAdmin ? `<button class="btn-ghost" id="dashboardBtn">Dashboard</button>` : ''}
      <button class="btn-ghost" id="logoutBtn">Sign out</button>
    </div>
    <div class="client-strip">
      <span class="name">${session.clientName}</span>
      ${session.isGuest ? `<span class="discount guest-badge">Browsing as guest</span>` : `<span class="discount">${session.discount}% off</span>`}
    </div>
    <div class="search-row">
      <input id="searchInput" type="text" placeholder="Search products or catalog #..." value="${searchQuery.replace(/"/g, '&quot;')}" />
      ${searchQuery.trim() ? `<button class="btn-round" id="clearSearchBtn">✕</button>` : ''}
    </div>
    <div class="crumbs">
      ${crumbs.map((c, i) => `<span ${i < crumbs.length - 1 && !searchQuery.trim() ? `data-crumb="${i}"` : ''} class="${i === crumbs.length - 1 ? 'crumb-current' : ''}">${c}</span>${i < crumbs.length - 1 ? '<span class="crumb-sep">/</span>' : ''}`).join("")}
    </div>`;

  if (level.items.length === 0) {
    html += `<div class="empty-state">No items here yet.</div>`;
  } else if (level.type !== "product") {
    html += `<div class="nav-grid">`;
    level.items.forEach(item => {
      html += `<div class="nav-card" data-nav="${item}"><div class="label">${item}</div></div>`;
    });
    html += `</div>`;
  } else {
    html += `<div class="product-grid">`;
    level.items.forEach(p => {
      const discounted = Number(p.EffectivePrice);
      const qty = cart[p.CatalogNumber] || 0;
      const hasStockInfo = p.Stock !== undefined && p.Stock !== "" && !isNaN(Number(p.Stock));
      const stock = hasStockInfo ? Number(p.Stock) : null;
      const outOfStock = stock !== null && stock <= 0;
      const atStockLimit = stock !== null && qty >= stock;
      html += `
        <div class="product-card ${outOfStock ? 'out-of-stock' : ''}">
          <div class="img-wrap">
            <img src="${p.ImageURL || ''}" alt="${p.ProductName}" data-zoom="${(p.ImageURL || '').replace(/"/g, '&quot;')}" data-zoom-name="${String(p.ProductName).replace(/"/g, '&quot;')}" onerror="this.style.opacity=0" />
            ${outOfStock ? `<span class="stock-badge">Out of stock</span>` : ''}
          </div>
          <div class="product-body">
            <p class="product-name">${p.ProductName}</p>
            <span class="catalog-chip">${p.CatalogNumber}</span>
            <div class="price-row">
              ${p.HasCustomPrice ? '' : `<span class="price-strike">${Number(p.Price).toFixed(2)}</span>`}
              <span class="price-final">${discounted.toFixed(2)} KWD</span>
            </div>
            ${outOfStock ? `
            <div class="qty-row"><span class="stock-note">Unavailable</span></div>
            ` : `
            <div class="qty-row">
              <button class="btn-round" data-dec="${p.CatalogNumber}">−</button>
              <span class="qty-val">${qty}</span>
              <button class="btn-round filled" data-inc="${p.CatalogNumber}" data-stock="${stock === null ? '' : stock}" ${atStockLimit ? 'disabled' : ''}>+</button>
            </div>
            ${stock !== null && stock <= 5 ? `<p class="stock-note">${stock} left</p>` : ''}
            `}
          </div>
        </div>`;
    });
    html += `</div>`;
  }

  html += `
    <div class="cart-fab ${cartCount() === 0 ? 'hidden' : ''}" id="cartFab">
      <span class="left">${ICON_CART} ${cartCount()} item${cartCount() === 1 ? '' : 's'}</span>
      <span class="total">${cartTotal().toFixed(2)} KWD</span>
    </div>`;

  root.innerHTML = html;

  document.getElementById("logoutBtn").onclick = logout;
  const dashBtn = document.getElementById("dashboardBtn");
  if (dashBtn) dashBtn.onclick = openDashboard;
  const searchInput = document.getElementById("searchInput");
  searchInput.oninput = e => { searchQuery = e.target.value; render(); };
  searchInput.focus();
  searchInput.setSelectionRange(searchInput.value.length, searchInput.value.length);
  const clearBtn = document.getElementById("clearSearchBtn");
  if (clearBtn) clearBtn.onclick = () => { searchQuery = ""; render(); };
  root.querySelectorAll("[data-crumb]").forEach(el => el.onclick = () => { path = path.slice(0, Number(el.dataset.crumb)); render(); });
  root.querySelectorAll("[data-nav]").forEach(el => el.onclick = () => { path.push(el.dataset.nav); render(); });
  root.querySelectorAll("img[data-zoom]").forEach(img => {
    if (!img.dataset.zoom) return;
    img.style.cursor = "zoom-in";
    img.onclick = () => openImageModal(img.dataset.zoom, img.dataset.zoomName);
  });
  root.querySelectorAll("[data-inc]").forEach(el => el.onclick = () => {
    const stockAttr = el.dataset.stock;
    const stockLimit = stockAttr === "" ? null : Number(stockAttr);
    const current = cart[el.dataset.inc] || 0;
    if (stockLimit !== null && current >= stockLimit) return;
    cart[el.dataset.inc] = current + 1;
    render();
  });
  root.querySelectorAll("[data-dec]").forEach(el => el.onclick = () => { cart[el.dataset.dec] = Math.max(0, (cart[el.dataset.dec] || 0) - 1); render(); });
  const fab = document.getElementById("cartFab");
  if (fab) fab.onclick = () => { view = "summary"; render(); };
}

function renderSummary() {
  const lines = Object.entries(cart)
    .filter(([, qty]) => qty > 0)
    .map(([catalogNumber, qty]) => {
      const p = products.find(pr => String(pr.CatalogNumber) === catalogNumber);
      const unitPrice = Number(p.EffectivePrice);
      return { ...p, qty, unitPrice, lineTotal: unitPrice * qty };
    });
  const grand = lines.reduce((s, l) => s + l.lineTotal, 0);

  let html = `
    <div class="topbar"><button class="btn-ghost" id="backBtn">← Back</button></div>
    <div class="summary-header">
      <p class="title">Order Summary</p>
      <p class="sub">${session.clientName} — this is what gets emailed on submit</p>
    </div>`;

  if (!lines.length) {
    html += `<div class="empty-state">Your cart is empty.</div>`;
  } else {
    html += `<table class="order-table"><thead><tr>
      <th>Product</th><th>Cat #</th><th class="num">Qty</th><th class="num">Total</th>
    </tr></thead><tbody>`;
    lines.forEach(l => {
      html += `<tr>
        <td>${l.ProductName}</td>
        <td style="font-family:var(--font-mono);font-size:11px;">${l.CatalogNumber}</td>
        <td class="num">${l.qty}</td>
        <td class="num">${l.lineTotal.toFixed(2)}</td>
      </tr>`;
    });
    html += `</tbody></table>
    <div class="grand-total-row"><span class="label">Grand Total</span><span class="value">${grand.toFixed(2)} KWD</span></div>`;

    if (session.isGuest) {
      html += `
      <div class="guest-form">
        <p class="guest-form-title">Your details</p>
        <input id="guestName" type="text" placeholder="First name" />
        <input id="guestPhone" type="tel" placeholder="Phone number" />
        <input id="guestAddress" type="text" placeholder="Address (as detailed as possible)" />
      </div>`;
    }

    html += `
    <div class="summary-actions">
      <button class="btn-secondary" id="editBtn">Edit order</button>
      <button class="btn-primary" id="submitBtn">Submit order</button>
    </div>
    <p class="status-msg" id="statusMsg"></p>`;
  }

  root.innerHTML = html;
  document.getElementById("backBtn").onclick = () => { view = "browse"; render(); };
  const editBtn = document.getElementById("editBtn");
  if (editBtn) editBtn.onclick = () => { view = "browse"; render(); };
  const submitBtn = document.getElementById("submitBtn");
  if (submitBtn) submitBtn.onclick = () => submitOrder(lines);
}

async function submitOrder(lines) {
  const btn = document.getElementById("submitBtn");
  const status = document.getElementById("statusMsg");

  let guestName, guestPhone, guestAddress;
  if (session.isGuest) {
    guestName = document.getElementById("guestName").value.trim();
    guestPhone = document.getElementById("guestPhone").value.trim();
    guestAddress = document.getElementById("guestAddress").value.trim();
    if (!guestName || !guestPhone || !guestAddress) {
      status.textContent = "Please fill in your name, phone number, and address.";
      status.className = "status-msg error";
      status.style.display = "block";
      return;
    }
  }

  btn.disabled = true; btn.textContent = "Submitting…";
  try {
    const payload = session.isGuest
      ? { action: "guestOrder", guestName, guestPhone, guestAddress, items: lines.map(l => ({ catalogNumber: l.CatalogNumber, qty: l.qty })) }
      : { action: "submitOrder", username: session.username, password: session.password, items: lines.map(l => ({ catalogNumber: l.CatalogNumber, qty: l.qty })) };
    const res = await apiPost(payload);
    if (!res.success) throw new Error(res.error || "Order failed");
    status.textContent = "✓ Order submitted — confirmation email sent.";
    status.className = "status-msg success";
    status.style.display = "block";
    cart = {};
    btn.textContent = "Submitted";
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    console.error("Submit order failed:", msg);
    status.textContent = "Could not submit order: " + msg;
    status.className = "status-msg error";
    status.style.display = "block";
    btn.disabled = false; btn.textContent = "Submit order";
  }
}

// ====== IMAGE LIGHTBOX ======
function openImageModal(url, name) {
  if (!url) return;
  const existing = document.getElementById("imgModal");
  if (existing) existing.remove();

  const overlay = document.createElement("div");
  overlay.id = "imgModal";
  overlay.className = "img-modal-overlay";
  overlay.innerHTML = `
    <button class="img-modal-close" aria-label="Close">✕</button>
    <img src="${url}" alt="${(name || '').replace(/"/g, '&quot;')}" class="img-modal-img" />
  `;
  overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
  overlay.querySelector(".img-modal-close").onclick = () => overlay.remove();
  document.body.appendChild(overlay);
}

// ====== ADMIN DASHBOARD ======
async function openDashboard() {
  view = "admin";
  adminStats = null;
  render();
  try {
    const res = await apiPost({ action: "adminStats", username: session.username, password: session.password });
    if (!res.success) throw new Error(res.error || "Failed to load stats");
    adminStats = res;
  } catch (err) {
    adminStats = { error: true };
  }
  render();
}

function monthLabel(ym) {
  const [y, m] = ym.split("-");
  const d = new Date(Number(y), Number(m) - 1, 1);
  return d.toLocaleDateString("en-US", { month: "short", year: "numeric" });
}

function renderAdmin() {
  let html = `
    <div class="topbar">
      <button class="btn-ghost" id="backBtn">← Back</button>
      <div class="topbar-spacer"></div>
      <p class="brand-name" style="font-size:16px;">Sales Dashboard</p>
    </div>`;

  if (!adminStats) {
    html += `<div class="loading-state">Loading stats…</div>`;
  } else if (adminStats.error) {
    html += `<div class="empty-state">Could not load stats. Please try again.</div>`;
  } else {
    const { monthly, byCustomer } = adminStats;

    html += `<p style="font-size:13px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.04em;margin:4px 0 8px;">Total sales per month</p>`;
    if (!monthly.length) {
      html += `<div class="empty-state" style="padding:24px;">No orders yet.</div>`;
    } else {
      html += `<table class="order-table" style="margin-bottom:20px;"><thead><tr><th>Month</th><th class="num">Total</th></tr></thead><tbody>`;
      monthly.forEach(m => {
        html += `<tr><td>${monthLabel(m.month)}</td><td class="num">${m.total.toFixed(2)} KWD</td></tr>`;
      });
      html += `</tbody></table>`;
    }

    html += `<p style="font-size:13px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.04em;margin:4px 0 8px;">Sales per customer, per month</p>`;
    if (!byCustomer.length) {
      html += `<div class="empty-state" style="padding:24px;">No orders yet.</div>`;
    } else {
      html += `<table class="order-table"><thead><tr><th>Customer</th><th>Month</th><th class="num">Total</th></tr></thead><tbody>`;
      byCustomer.forEach(c => {
        html += `<tr><td>${c.clientName}</td><td>${monthLabel(c.month)}</td><td class="num">${c.total.toFixed(2)} KWD</td></tr>`;
      });
      html += `</tbody></table>`;
    }
  }

  root.innerHTML = html;
  document.getElementById("backBtn").onclick = () => { view = "browse"; render(); };
}

// ====== SERVICE WORKER ======
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  });
}

boot();
