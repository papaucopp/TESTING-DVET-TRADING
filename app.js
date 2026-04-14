const COMPANY_CONFIG = {
  trading: {
    label: "DVET TRADING SDN BHD",
    name: "Dvet Trading Sdn Bhd",
    registrationNo: "1482481-M",
    addressLines: [
      "TINGKAT 4, WISMA KAWAN, 36,",
      "JALAN SERI UTARA 1, OFF JALAN IPOH, SRI UTARA,",
      "68100 BATU CAVES,",
      "WILAYAH PERSEKUTUAN KUALA LUMPUR.",
    ],
    phone: "03-6243 7276",
    bankDetails: "AL RAJHI BANK A/C 129001881556134",
  },
  holding: {
    label: "DVET HOLDING SDN BHD",
    name: "Dvet Holding Sdn Bhd",
    registrationNo: "707154-K",
    addressLines: [
      "TINGKAT 4, WISMA KAWAN, 36,",
      "JALAN SERI UTARA 1, OFF JALAN IPOH, SRI UTARA,",
      "68100 BATU CAVES,",
      "WILAYAH PERSEKUTUAN KUALA LUMPUR.",
    ],
    phone: "03-6243 7276",
    bankDetails: "AL RAJHI BANK A/C 129001881556134",
  },
};

function createBlankItem() {
  return { description: "", code: "", quantity: 1, unitPrice: 0, amount: 0 };
}

const CATEGORY_CONFIG = {
  ubat: { label: "UBAT", code: "U" },
  susu: { label: "SUSU", code: "S" },
  peralatan: { label: "PERALATAN", code: "P" },
  lampin_pakai_buang: { label: "LAMPIN PAKAI BUANG", code: "LPB" },
};

const defaultState = {
  company: "",
  category: "",
  poNumber: "",
  poDate: "",
  poSupplier: "",
  poSuppliers: [""],
  poShipTo: "",
  poPreparedBy: "",
  poItems: [createBlankItem()],
  invoiceNumber: "",
  invoiceDate: "",
  invoiceBillTo: "",
  invoiceShipTo: "",
  invoicePreparedBy: "",
  invoicePoJhev: "",
  doNumber: "",
  invoiceStatus: "Pending",
  doStatus: "Delivered",
  doBillTo: "",
  doShipTo: "",
  doShipDate: "",
  doInvoiceNumber: "",
  doInvoiceDate: "",
  doPoJhev: "",
  invoiceItems: [createBlankItem()],
  doItems: [createBlankItem()],
};
const BILL_TO_STORAGE_KEY = "po-complete-set-bill-to-history";
const SUPPLIER_STORAGE_KEY = "po-complete-set-supplier-history";
const COUNTERS_STORAGE_KEY = "po-complete-set-running-counters";
const RECORDS_STORAGE_KEY = "po-complete-set-saved-records";
const PO_ITEMS_STORAGE_KEY = "po-complete-set-saved-po-items";
const AUTH_SESSION_STORAGE_KEY = "po-complete-set-auth-session";
const LPB_INVOICE_FIXED_AMOUNT = 300;
const ACCESS_USERS = {
  ALIF: { password: "ALIF@DVET", displayName: "Alif" },
  ROSLIMAN: { password: "ROSLIMAN@DVET", displayName: "Rosliman" },
  SYAWAL: { password: "SYAWAL@DVET", displayName: "Syawal" },
};
const state = {
  data: structuredClone(defaultState),
  extractedText: "",
  billToHistory: [],
  supplierHistory: [],
  runningCounters: {},
  savedRecords: [],
  savedPoItems: [],
  currentRecordId: "",
  logoDataUrl: "",
  authUser: "",
};

const form = document.getElementById("documentForm");
const statusBox = document.getElementById("statusBox");
const authOverlay = document.getElementById("authOverlay");
const loginForm = document.getElementById("loginForm");
const loginUserIdField = document.getElementById("loginUserId");
const loginPasswordField = document.getElementById("loginPassword");
const loginNotice = document.getElementById("loginNotice");
const activeUserBadge = document.getElementById("activeUserBadge");
const logoutButton = document.getElementById("logoutButton");
const companyInputs = Array.from(document.querySelectorAll('input[name="documentCompany"]'));
const categorySelector = document.getElementById("categorySelector");
const categoryInputs = Array.from(document.querySelectorAll('input[name="documentCategory"]'));
const uploadCard = document.getElementById("uploadCard");
const poFileInput = document.getElementById("poFile");
const poSupplierField = form.elements.namedItem("poSupplier");
const poSupplierLabel = document.getElementById("poSupplierLabel");
const supplierHistoryLabel = document.getElementById("supplierHistoryLabel");
const poSupplierList = document.getElementById("poSupplierList");
const addPoSupplierButton = document.getElementById("addPoSupplierButton");
const poItemsEditor = document.getElementById("poItemsEditor");
const invoiceItemsEditor = document.getElementById("invoiceItemsEditor");
const doItemsEditor = document.getElementById("doItemsEditor");
const itemEditorTemplate = document.getElementById("itemEditorTemplate");
const poPreview = document.getElementById("poPreview");
const invoicePreview = document.getElementById("invoicePreview");
const doPreview = document.getElementById("doPreview");
const billToHistoryList = document.getElementById("billToHistory");
const supplierHistoryList = document.getElementById("supplierHistory");
const documentHistoryList = document.getElementById("documentHistory");
const saveBillToHistoryButton = document.getElementById("saveBillToHistoryButton");
const editBillToHistoryButton = document.getElementById("editBillToHistoryButton");
const deleteBillToHistoryButton = document.getElementById("deleteBillToHistoryButton");
const saveSupplierHistoryButton = document.getElementById("saveSupplierHistoryButton");
const editSupplierHistoryButton = document.getElementById("editSupplierHistoryButton");
const deleteSupplierHistoryButton = document.getElementById("deleteSupplierHistoryButton");
const saveRecordButton = document.getElementById("saveRecordButton");
const exportRecordsButton = document.getElementById("exportRecordsButton");
const deleteRecordButton = document.getElementById("deleteRecordButton");
const deleteAllRecordsButton = document.getElementById("deleteAllRecordsButton");
const resetCountersButton = document.getElementById("resetCountersButton");
const newPoButton = document.getElementById("newPoButton");
const syncPoButton = document.getElementById("syncPoButton");
const addPoItemButton = document.getElementById("addPoItemButton");
const saveRecordNotice = document.getElementById("saveRecordNotice");
let saveRecordNoticeTimeoutId = null;

const fieldNames = [
  "company",
  "poNumber",
  "poDate",
  "poSupplier",
  "poShipTo",
  "poPreparedBy",
  "invoiceNumber",
  "invoiceDate",
  "invoiceBillTo",
  "invoiceShipTo",
  "invoicePreparedBy",
  "invoicePoJhev",
  "doNumber",
  "invoiceStatus",
  "doStatus",
  "doBillTo",
  "doShipTo",
  "doShipDate",
  "doInvoiceNumber",
  "doInvoiceDate",
  "doPoJhev",
];
const dateFieldNames = new Set(["poDate", "invoiceDate", "doShipDate", "doInvoiceDate"]);

async function ensurePdfJs() {
  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.worker.min.mjs";
    return window.pdfjsLib;
  }

  const pdfjsModule = await import(
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.min.mjs"
  );
  pdfjsModule.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.worker.min.mjs";
  return pdfjsModule;
}

async function ensureLogoDataUrl() {
  if (state.logoDataUrl) {
    return state.logoDataUrl;
  }

  const response = await fetch("./assets/dvet-logo.png");
  if (!response.ok) {
    throw new Error("Failed to load logo image.");
  }

  const blob = await response.blob();
  state.logoDataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Failed to convert logo image."));
    reader.readAsDataURL(blob);
  });

  return state.logoDataUrl;
}

function setStatus(message, tone = "neutral") {
  statusBox.textContent = message;
  statusBox.dataset.tone = tone;
}

function setLoginNotice(message, tone = "neutral") {
  if (!loginNotice) return;
  loginNotice.textContent = message;
  loginNotice.dataset.tone = tone;
}

function updateAuthUi() {
  const isAuthenticated = Boolean(state.authUser);
  document.body.classList.toggle("auth-locked", !isAuthenticated);
  document.querySelector(".app-shell")?.classList.toggle("is-locked", !isAuthenticated);
  authOverlay?.classList.toggle("is-hidden", isAuthenticated);
  if (activeUserBadge) {
    activeUserBadge.textContent = isAuthenticated ? `Logged in: ${state.authUser}` : "Not logged in";
  }
}

function saveAuthSession(userId) {
  try {
    localStorage.setItem(AUTH_SESSION_STORAGE_KEY, JSON.stringify({ userId }));
  } catch (error) {
    console.error("Failed to save auth session", error);
  }
}

function clearAuthSession() {
  try {
    localStorage.removeItem(AUTH_SESSION_STORAGE_KEY);
  } catch (error) {
    console.error("Failed to clear auth session", error);
  }
}

function restoreAuthSession() {
  try {
    const raw = localStorage.getItem(AUTH_SESSION_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    const savedUserId = String(parsed?.userId || "").toUpperCase();
    if (ACCESS_USERS[savedUserId]) {
      state.authUser = ACCESS_USERS[savedUserId].displayName;
    }
  } catch (error) {
    console.error("Failed to restore auth session", error);
    state.authUser = "";
  }
}

function loginUser(userId, password) {
  const normalizedUserId = String(userId || "").trim().toUpperCase();
  const normalizedPassword = String(password || "").trim();
  const account = ACCESS_USERS[normalizedUserId];
  if (!account || account.password !== normalizedPassword) {
    setLoginNotice("Invalid ID or password. Please try again.", "error");
    return false;
  }

  state.authUser = account.displayName;
  saveAuthSession(normalizedUserId);
  applyAuthenticatedPreparedBy(state.data);
  syncFormFromState();
  updateAuthUi();
  setLoginNotice("", "neutral");
  if (loginPasswordField) {
    loginPasswordField.value = "";
  }
  setStatus(
    `Welcome ${account.displayName}. Please select one company first, then choose UBAT, SUSU, PERALATAN, or LAMPIN PAKAI BUANG.`,
    "success",
  );
  return true;
}

function logoutCurrentUser() {
  state.authUser = "";
  clearAuthSession();
  updateAuthUi();
  setLoginNotice("Session locked. Please log in again.", "warn");
  setStatus("System locked. Please log in to continue.", "warn");
}

function getAuthenticatedPreparedBy() {
  return String(state.authUser || "").trim();
}

function applyAuthenticatedPreparedBy(target = state.data) {
  const preparedBy = getAuthenticatedPreparedBy();
  if (!preparedBy) return target;
  target.poPreparedBy = preparedBy;
  target.invoicePreparedBy = preparedBy;
  return target;
}

function setSaveRecordNotice(message = "", tone = "neutral") {
  if (!saveRecordNotice) return;
  saveRecordNotice.textContent = message;
  saveRecordNotice.dataset.tone = tone;

  if (saveRecordNoticeTimeoutId) {
    clearTimeout(saveRecordNoticeTimeoutId);
    saveRecordNoticeTimeoutId = null;
  }

  if (message) {
    saveRecordNoticeTimeoutId = window.setTimeout(() => {
      saveRecordNotice.textContent = "";
      saveRecordNotice.dataset.tone = "neutral";
      saveRecordNoticeTimeoutId = null;
    }, 5000);
  }
}

function getSaveRecordMessage(sourceData = state.data) {
  const preparedBy = String(sourceData.poPreparedBy || sourceData.invoicePreparedBy || "").trim();
  if (!preparedBy) {
    return "HANTU! SIMPAN!";
  }
  return `${preparedBy.toUpperCase()}! KERJA KERAS ANDA TELAH DISIMPAN`;
}

function currency(value) {
  return Number(value || 0).toLocaleString("en-MY", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatNumber(value) {
  if (!Number.isFinite(Number(value))) return "0";
  return `${Number(value)}`;
}

function sanitizeFilenamePart(value = "") {
  return String(value)
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function safeText(text = "") {
  return String(text).replace(/[&<>"]/g, (char) => {
    const map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" };
    return map[char] || char;
  });
}

function normalizeMultiline(text = "") {
  return String(text)
    .replace(/\s+\n/g, "\n")
    .replace(/\n\s+/g, "\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

function sumItems(items) {
  return items.reduce((total, item) => total + Number(item.amount || 0), 0);
}

function totalQuantity(items) {
  return items.reduce((total, item) => total + Number(item.quantity || 0), 0);
}

function displayDateToIso(value = "") {
  const match = String(value).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return "";
  const [, dd, mm, yyyy] = match;
  return `${yyyy}-${mm}-${dd}`;
}

function isoDateToDisplay(value = "") {
  const match = String(value).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return value || "";
  const [, yyyy, mm, dd] = match;
  return `${dd}/${mm}/${yyyy}`;
}

function getYearSuffix(dateValue = "") {
  const displayMatch = String(dateValue).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (displayMatch) {
    return displayMatch[3].slice(-2);
  }

  const isoMatch = String(dateValue).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return isoMatch[1].slice(-2);
  }

  return String(new Date().getFullYear()).slice(-2);
}

function getCounterBucket(type, yearSuffix, categoryCode = "", company = state.data.company) {
  return `${company || "unknown"}-${type}-${categoryCode || "X"}-${yearSuffix}`;
}

function formatTradingRunningNumber(type, yearSuffix, categoryCode, counter) {
  const paddedCounter = String(counter).padStart(5, "0");
  if (type === "po") {
    return `DT/PO-${categoryCode}/${yearSuffix}-${paddedCounter}`;
  }
  if (type === "invoice") {
    return `DT/INV-${categoryCode}/${yearSuffix}-${paddedCounter}`;
  }
  return `DT/DO-${categoryCode}/${yearSuffix}-${paddedCounter}`;
}

function formatHoldingRunningNumber(type, yearSuffix, categoryCode, counter) {
  const paddedCounter = String(counter).padStart(5, "0");
  if (type === "po") {
    return `DH/PO-${categoryCode}/${yearSuffix}-${paddedCounter}`;
  }
  if (type === "invoice") {
    return `DH/INV-${categoryCode}/${yearSuffix}-${paddedCounter}`;
  }
  return `DH/DO-${categoryCode}/${yearSuffix}-${paddedCounter}`;
}

function formatRunningNumberByCompany(type, company, yearSuffix, categoryCode, counter) {
  if (company === "holding") {
    return formatHoldingRunningNumber(type, yearSuffix, categoryCode, counter);
  }
  return formatTradingRunningNumber(type, yearSuffix, categoryCode, counter);
}

function getCategoryMeta(category = state.data.category) {
  return CATEGORY_CONFIG[category] || null;
}

function isLampinCategory(category = state.data.category) {
  return category === "lampin_pakai_buang";
}

function getPoSubNumber(basePoNumber, supplierIndex) {
  const normalizedBase = String(basePoNumber || "").trim();
  if (!normalizedBase) return "";
  return `${normalizedBase}(${supplierIndex + 1})`;
}

function getPoSupplierEntries(data = state.data) {
  if (Array.isArray(data.poSuppliers) && data.poSuppliers.length) {
    return data.poSuppliers.map((entry) => String(entry ?? ""));
  }
  const fallback = String(data.poSupplier || "");
  return [fallback];
}

function buildPoSupplierSummary(data = state.data) {
  if (!isLampinCategory(data.category)) {
    return normalizeMultiline(data.poSupplier || getPoSupplierEntries(data)[0] || "");
  }

  const supplierEntries = getPoSupplierEntries(data).map((supplier) => normalizeMultiline(supplier));
  const hasMultiple = supplierEntries.length > 1;
  const basePoNumber = data.poNumber || "";
  const sections = supplierEntries
    .map((supplier, index) => {
      if (!supplier) return "";
      return `${hasMultiple ? getPoSubNumber(basePoNumber, index) : basePoNumber}\n${supplier}`;
    })
    .filter(Boolean);

  return sections.join("\n\n");
}

function syncPoSupplierDerivedText(target = state.data) {
  const supplierEntries = getPoSupplierEntries(target);
  target.poSuppliers = supplierEntries.length ? supplierEntries : [""];
  target.poSupplier = buildPoSupplierSummary(target);
  return target.poSupplier;
}

function getPoSupplierPageEntries(data = state.data) {
  const suppliers = getPoSupplierEntries(data).map((supplier) => normalizeMultiline(supplier));
  const hasMultiple = suppliers.length > 1;
  const basePoNumber = String(data.poNumber || "").trim();
  return suppliers.map((supplier, index) => ({
    supplier,
    poNumber: hasMultiple ? getPoSubNumber(basePoNumber, index) : basePoNumber,
    index,
  }));
}

function getCompanyMeta(company = state.data.company) {
  return COMPANY_CONFIG[company] || null;
}

function getActiveCompanyProfile() {
  return getCompanyMeta() || COMPANY_CONFIG.trading;
}

function syncCompanyControlsFromState() {
  companyInputs.forEach((input) => {
    const isActive = input.value === state.data.company;
    input.checked = isActive;
    input.closest(".category-option")?.classList.toggle("is-active", isActive);
  });
}

function syncCategoryControlsFromState() {
  categoryInputs.forEach((input) => {
    const isActive = input.value === state.data.category;
    input.checked = isActive;
    input.closest(".category-option")?.classList.toggle("is-active", isActive);
  });
}

function setCategoryGateState() {
  const enabled = Boolean(state.data.company && state.data.category);
  const gatedButtons = [
    document.getElementById("parseButton"),
    newPoButton,
    addPoItemButton,
    addPoSupplierButton,
    syncPoButton,
    saveRecordButton,
    document.getElementById("downloadPoButton"),
    document.getElementById("downloadInvoiceButton"),
    document.getElementById("downloadDoButton"),
    document.getElementById("downloadCompleteSetButton"),
  ];

  [...form.elements].forEach((field) => {
    field.disabled = !enabled;
  });

  gatedButtons.forEach((button) => {
    if (button) button.disabled = !enabled;
  });

  if (poFileInput) {
    poFileInput.disabled = !enabled;
  }
  uploadCard?.classList.toggle("is-disabled", !enabled);
  categorySelector?.classList.toggle("is-disabled", !state.data.company);
}

function syncFormFromState() {
  syncPoSupplierDerivedText(state.data);
  fieldNames.forEach((name) => {
    const field = form.elements.namedItem(name);
    if (field) {
      if (dateFieldNames.has(name) && field.type === "date") {
        field.value = displayDateToIso(state.data[name] ?? "");
      } else {
        field.value = state.data[name] ?? "";
      }
    }
  });
  renderPoSupplierEditor();
}

function syncStateFromForm() {
  fieldNames.forEach((name) => {
    const field = form.elements.namedItem(name);
    if (field) {
      if (dateFieldNames.has(name) && field.type === "date") {
        state.data[name] = isoDateToDisplay(field.value);
      } else {
        state.data[name] = field.value;
      }
    }
  });
  if (isLampinCategory()) {
    state.data.poSuppliers = getPoSupplierEntries(state.data);
  } else {
    state.data.poSuppliers = [String(state.data.poSupplier || "")];
  }
  syncPoSupplierDerivedText(state.data);
}

function generateSequence(prefix, source, fallback = "0001") {
  const digits = String(source || "").match(/\d+/g)?.join("") || fallback;
  return `${prefix}${digits.slice(-6).padStart(6, "0")}`;
}

function coerceNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function inferItemAmounts(items) {
  return items.map((item) => {
    const quantity = coerceNumber(item.quantity || 0);
    const unitPrice = coerceNumber(item.unitPrice || 0);
    const amount = coerceNumber(item.amount || quantity * unitPrice);
    return {
      description: item.description || "",
      code: item.code || "",
      quantity,
      unitPrice,
      amount,
    };
  });
}

function normalizeItemsWithoutFormula(items) {
  return items.map((item) => ({
    description: item.description || "",
    code: item.code || "",
    quantity: coerceNumber(item.quantity || 0),
    unitPrice: coerceNumber(item.unitPrice || 0),
    amount: coerceNumber(item.amount || 0),
  }));
}

function applyLampinInvoiceAmount(items) {
  return normalizeItemsWithoutFormula(items).map((item) => ({
    ...item,
    amount: LPB_INVOICE_FIXED_AMOUNT,
  }));
}

function loadBillToHistory() {
  try {
    const raw = localStorage.getItem(BILL_TO_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    state.billToHistory = parsed
      .filter((entry) => entry && entry.value)
      .sort((a, b) => Number(b.savedAt || 0) - Number(a.savedAt || 0));
  } catch (error) {
    console.error("Failed to load Bill To history", error);
    state.billToHistory = [];
  }
}

function saveBillToHistory() {
  try {
    localStorage.setItem(BILL_TO_STORAGE_KEY, JSON.stringify(state.billToHistory));
  } catch (error) {
    console.error("Failed to save Bill To history", error);
  }
}

function loadSupplierHistory() {
  try {
    const raw = localStorage.getItem(SUPPLIER_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    state.supplierHistory = parsed
      .filter((entry) => entry && entry.value)
      .sort((a, b) => Number(b.savedAt || 0) - Number(a.savedAt || 0));
  } catch (error) {
    console.error("Failed to load Supplier history", error);
    state.supplierHistory = [];
  }
}

function saveSupplierHistory() {
  try {
    localStorage.setItem(SUPPLIER_STORAGE_KEY, JSON.stringify(state.supplierHistory));
  } catch (error) {
    console.error("Failed to save Supplier history", error);
  }
}

function loadRunningCounters() {
  try {
    const raw = localStorage.getItem(COUNTERS_STORAGE_KEY);
    state.runningCounters = raw ? JSON.parse(raw) : {};
  } catch (error) {
    console.error("Failed to load running counters", error);
    state.runningCounters = {};
  }
}

function saveRunningCounters() {
  try {
    localStorage.setItem(COUNTERS_STORAGE_KEY, JSON.stringify(state.runningCounters));
  } catch (error) {
    console.error("Failed to save running counters", error);
  }
}

function loadSavedRecords() {
  try {
    const raw = localStorage.getItem(RECORDS_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    state.savedRecords = parsed
      .filter((record) => record && record.id && record.data)
      .sort((a, b) => Number(b.updatedAt || b.createdAt || 0) - Number(a.updatedAt || a.createdAt || 0));
  } catch (error) {
    console.error("Failed to load saved records", error);
    state.savedRecords = [];
  }
}

function saveSavedRecords() {
  try {
    localStorage.setItem(RECORDS_STORAGE_KEY, JSON.stringify(state.savedRecords));
  } catch (error) {
    console.error("Failed to save records", error);
  }
}

function loadSavedPoItems() {
  try {
    const raw = localStorage.getItem(PO_ITEMS_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    state.savedPoItems = parsed.filter((item) => item && normalizeMultiline(item.description));
  } catch (error) {
    console.error("Failed to load saved PO items", error);
    state.savedPoItems = [];
  }
}

function saveSavedPoItems() {
  try {
    localStorage.setItem(PO_ITEMS_STORAGE_KEY, JSON.stringify(state.savedPoItems));
  } catch (error) {
    console.error("Failed to save PO item history", error);
  }
}

function getPoItemHistoryLabel(item) {
  return `${sanitizeFilenamePart(item.description)} | Qty ${formatNumber(item.quantity)} | RM ${currency(item.unitPrice)}`;
}

function rememberPoItems(items = state.data.poItems) {
  const normalizedItems = inferItemAmounts(items)
    .filter((item) => normalizeMultiline(item.description))
    .map((item) => ({
      description: normalizeMultiline(item.description),
      quantity: coerceNumber(item.quantity),
      unitPrice: coerceNumber(item.unitPrice),
      amount: coerceNumber(item.amount),
    }));

  if (!normalizedItems.length) return;

  const merged = [...normalizedItems, ...state.savedPoItems].reduce((accumulator, item) => {
    const key = JSON.stringify([
      item.description,
      coerceNumber(item.quantity),
      coerceNumber(item.unitPrice),
      coerceNumber(item.amount),
    ]);
    if (!accumulator.some((entry) => entry.key === key)) {
      accumulator.push({ key, item });
    }
    return accumulator;
  }, []);

  state.savedPoItems = merged.map((entry) => entry.item).slice(0, 100);
  saveSavedPoItems();
}

function deleteSavedPoItem(index) {
  if (!Number.isInteger(index) || !state.savedPoItems[index]) return false;
  state.savedPoItems.splice(index, 1);
  saveSavedPoItems();
  return true;
}

function deleteSavedPoItems(indexes = []) {
  const normalizedIndexes = [...new Set(indexes.filter((index) => Number.isInteger(index) && state.savedPoItems[index]))]
    .sort((a, b) => b - a);
  if (!normalizedIndexes.length) return false;
  normalizedIndexes.forEach((index) => {
    state.savedPoItems.splice(index, 1);
  });
  saveSavedPoItems();
  return true;
}

function buildRecordLabel(record) {
  const recipient =
    splitLines(record.data?.invoiceShipTo)[0] ||
    splitLines(record.data?.doShipTo)[0] ||
    splitLines(record.data?.poShipTo)[0] ||
    "Untitled";
  return `${record.data?.poNumber || "PO"} | ${record.data?.invoiceNumber || "INV"} | ${record.data?.doNumber || "DO"} | ${recipient}`;
}

function renderSavedRecords() {
  if (!documentHistoryList) return;
  const currentCompany = state.data.company;
  const currentCategory = state.data.category;
  const filteredRecords = currentCompany && currentCategory
    ? state.savedRecords.filter(
        (record) =>
          record.data?.company === currentCompany && record.data?.category === currentCategory,
      )
    : [];

  documentHistoryList.innerHTML = currentCompany && currentCategory
    ? '<option value="">Select saved record</option>'
    : '<option value="">Select company and category first</option>';

  filteredRecords.forEach((record) => {
    const option = document.createElement("option");
    option.value = record.id;
    option.textContent = buildRecordLabel(record);
    documentHistoryList.appendChild(option);
  });

  if (state.currentRecordId && filteredRecords.some((record) => record.id === state.currentRecordId)) {
    documentHistoryList.value = state.currentRecordId;
  } else {
    documentHistoryList.value = "";
  }
}

function peekNextRunningNumber(type, yearSuffix, categoryCode, company) {
  const bucket = getCounterBucket(type, yearSuffix, categoryCode, company);
  const nextCounter = Number(state.runningCounters[bucket] || 0) + 1;
  return formatRunningNumberByCompany(type, company, yearSuffix, categoryCode, nextCounter);
}

function assignDraftRunningNumbers(target = state.data) {
  const categoryMeta = getCategoryMeta(target.category);
  const company = target.company || state.data.company;
  if (!categoryMeta || !company) {
    target.poNumber = "";
    target.invoiceNumber = "";
    target.doNumber = "";
    target.doInvoiceNumber = "";
    return;
  }
  const yearSuffix = getYearSuffix(target.poDate || target.invoiceDate || target.doShipDate || target.doInvoiceDate);
  target.poNumber = peekNextRunningNumber("po", yearSuffix, categoryMeta.code, company);
  target.invoiceNumber = peekNextRunningNumber("invoice", yearSuffix, categoryMeta.code, company);
  target.doNumber = peekNextRunningNumber("do", yearSuffix, categoryMeta.code, company);
  target.doInvoiceNumber = target.invoiceNumber;
}

function commitCurrentDraftNumbers(target = state.data) {
  const categoryMeta = getCategoryMeta(target.category);
  const company = target.company || state.data.company;
  if (!categoryMeta || !company) return;
  const yearSuffix = getYearSuffix(target.poDate || target.invoiceDate || target.doShipDate || target.doInvoiceDate);
  state.runningCounters[getCounterBucket("po", yearSuffix, categoryMeta.code, company)] =
    Number(state.runningCounters[getCounterBucket("po", yearSuffix, categoryMeta.code, company)] || 0) + 1;
  state.runningCounters[getCounterBucket("invoice", yearSuffix, categoryMeta.code, company)] =
    Number(state.runningCounters[getCounterBucket("invoice", yearSuffix, categoryMeta.code, company)] || 0) + 1;
  state.runningCounters[getCounterBucket("do", yearSuffix, categoryMeta.code, company)] =
    Number(state.runningCounters[getCounterBucket("do", yearSuffix, categoryMeta.code, company)] || 0) + 1;
  saveRunningCounters();
}

function hasMeaningfulDocumentData(data = state.data) {
  return Boolean(
    normalizeMultiline(data.poNumber) ||
      normalizeMultiline(data.invoiceNumber) ||
      normalizeMultiline(data.doNumber) ||
      normalizeMultiline(data.poSupplier) ||
      normalizeMultiline(data.poShipTo) ||
      normalizeMultiline(data.invoiceBillTo) ||
      normalizeMultiline(data.invoiceShipTo) ||
      data.poItems.some((item) => normalizeMultiline(item.description)),
  );
}

function saveCurrentRecord({ showStatus = true } = {}) {
  syncStateFromForm();

  if (!hasMeaningfulDocumentData(state.data)) {
    if (showStatus) {
      setSaveRecordNotice("Please fill in the document details before saving a record.", "warn");
    }
    return false;
  }

  const now = Date.now();
  const recordId = state.currentRecordId || `record-${now}`;
  const existingRecord = state.savedRecords.find((record) => record.id === recordId);
  const record = {
    id: recordId,
    createdAt: existingRecord?.createdAt || now,
    updatedAt: now,
    data: structuredClone(state.data),
  };
  const saveMessage = getSaveRecordMessage(record.data);

  state.savedRecords = [record, ...state.savedRecords.filter((entry) => entry.id !== recordId)]
    .sort((a, b) => Number(b.updatedAt || 0) - Number(a.updatedAt || 0))
    .slice(0, 100);
  const wasExistingRecord = Boolean(existingRecord);
  state.currentRecordId = recordId;
  rememberPoItems(state.data.poItems);
  rememberSupplier(state.data.poSupplier);
  saveSavedRecords();
  renderSavedRecords();

  if (!wasExistingRecord) {
    commitCurrentDraftNumbers(record.data);
  }

  syncCompanyControlsFromState();
  syncCategoryControlsFromState();
  setCategoryGateState();
  syncFormFromState();
  renderItemsEditor(poItemsEditor, "poItems");
  renderItemsEditor(invoiceItemsEditor, "invoiceItems");
  renderItemsEditor(doItemsEditor, "doItems");
  renderDocuments();
  renderSavedRecords();

  if (showStatus) {
    setSaveRecordNotice(saveMessage, "success");
  }
  return true;
}

function loadRecordById(recordId) {
  const record = state.savedRecords.find((entry) => entry.id === recordId);
  if (!record) return false;

  state.currentRecordId = record.id;
  state.data = applyAuthenticatedPreparedBy({
    ...structuredClone(defaultState),
    ...structuredClone(record.data),
  });
  syncFormFromState();
  syncCompanyControlsFromState();
  syncCategoryControlsFromState();
  setCategoryGateState();
  renderItemsEditor(poItemsEditor, "poItems");
  renderItemsEditor(invoiceItemsEditor, "invoiceItems");
  renderItemsEditor(doItemsEditor, "doItems");
  renderDocuments();
  renderSavedRecords();
  return true;
}

function deleteRecordById(recordId) {
  if (!recordId) return false;
  const nextRecords = state.savedRecords.filter((record) => record.id !== recordId);
  if (nextRecords.length === state.savedRecords.length) return false;
  state.savedRecords = nextRecords;
  if (state.currentRecordId === recordId) {
    state.currentRecordId = "";
  }
  saveSavedRecords();
  renderSavedRecords();
  return true;
}

function deleteAllSavedRecords() {
  state.savedRecords = [];
  state.currentRecordId = "";
  saveSavedRecords();
  renderSavedRecords();
}

function escapeCsv(value = "") {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function buildExportRows() {
  return state.savedRecords.flatMap((record) => {
    const data = record.data || {};
    const categoryLabel = getCategoryMeta(data.category)?.label || "";
    const items = inferItemAmounts(data.poItems?.length ? data.poItems : [createBlankItem()]);

    return items.map((item, index) => ({
      category: categoryLabel,
      po_no: data.poNumber || "",
      po_date: data.poDate || "",
      invoice_no: data.invoiceNumber || "",
      invoice_date: data.invoiceDate || "",
      do_no: data.doNumber || "",
      do_date: data.doShipDate || "",
      supplier: normalizeMultiline(data.poSupplier || ""),
      bill_to: normalizeMultiline(data.invoiceBillTo || ""),
      ship_to: normalizeMultiline(data.poShipTo || data.invoiceShipTo || data.doShipTo || ""),
      prepared_by: data.poPreparedBy || data.invoicePreparedBy || "",
      status_invoice: data.invoiceStatus || "",
      status_do: data.doStatus || "",
      item_no: index + 1,
      description: normalizeMultiline(item.description || ""),
      qty: coerceNumber(item.quantity),
      unit_price: coerceNumber(item.unitPrice),
      amount: coerceNumber(item.amount),
      saved_at: record.updatedAt ? new Date(record.updatedAt).toLocaleString("en-MY") : "",
    }));
  });
}

function exportSavedRecordsToExcel() {
  const rows = buildExportRows();
  if (!rows.length) {
    setStatus("No saved records available to export.", "warn");
    return;
  }

  const headers = [
    "Category",
    "PO No",
    "PO Date",
    "Invoice No",
    "Invoice Date",
    "DO No",
    "DO Date",
    "Supplier",
    "Bill To",
    "Ship To",
    "Prepared By",
    "Invoice Status",
    "DO Status",
    "Item No",
    "Description",
    "Qty",
    "Unit Price",
    "Amount",
    "Saved At",
  ];

  const csvLines = [
    headers.join(","),
    ...rows.map((row) =>
      [
        row.category,
        row.po_no,
        row.po_date,
        row.invoice_no,
        row.invoice_date,
        row.do_no,
        row.do_date,
        row.supplier,
        row.bill_to,
        row.ship_to,
        row.prepared_by,
        row.status_invoice,
        row.status_do,
        row.item_no,
        row.description,
        row.qty,
        row.unit_price,
        row.amount,
        row.saved_at,
      ]
        .map(escapeCsv)
        .join(","),
    ),
  ];

  const blob = new Blob([`\uFEFF${csvLines.join("\n")}`], {
    type: "text/csv;charset=utf-8;",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `DVET_Saved_Records_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  setStatus("Saved records exported successfully for Excel.", "success");
}

function createNewDocumentState() {
  const nextState = structuredClone(defaultState);
  nextState.invoiceStatus = "Pending";
  nextState.doStatus = "Delivered";
  nextState.poSuppliers = [""];
  return applyAuthenticatedPreparedBy(nextState);
}

function createNewDocumentStateForCategory(category) {
  const nextState = structuredClone(defaultState);
  nextState.company = state.data.company;
  nextState.category = category;
  nextState.invoiceStatus = "Pending";
  nextState.doStatus = "Delivered";
  nextState.poSuppliers = [""];
  assignDraftRunningNumbers(nextState);
  syncPoSupplierDerivedText(nextState);
  return applyAuthenticatedPreparedBy(nextState);
}

function createBlankDraftWithCurrentNumbers(current = state.data) {
  return applyAuthenticatedPreparedBy({
    ...structuredClone(defaultState),
    company: current.company,
    category: current.category,
    poNumber: current.poNumber,
    invoiceNumber: current.invoiceNumber,
    doNumber: current.doNumber,
    doInvoiceNumber: current.invoiceNumber,
    invoiceStatus: "Pending",
    doStatus: "Delivered",
    poSuppliers: [""],
  });
}

function resetRunningCounters() {
  state.runningCounters = {};
  saveRunningCounters();
}

function renderBillToHistory() {
  if (!billToHistoryList) return;
  billToHistoryList.innerHTML = '<option value="">Select previous Bill To</option>';
  state.billToHistory.forEach((entry) => {
    const option = document.createElement("option");
    option.value = entry.value;
    option.textContent = entry.value.split("\n")[0];
    billToHistoryList.appendChild(option);
  });
}

function renderSupplierHistory() {
  if (!supplierHistoryList) return;
  supplierHistoryList.innerHTML = '<option value="">Select previous Supplier</option>';
  state.supplierHistory.forEach((entry) => {
    const option = document.createElement("option");
    option.value = entry.value;
    option.textContent = entry.value.split("\n")[0];
    supplierHistoryList.appendChild(option);
  });
}

function renderPoSupplierEditor() {
  if (!poSupplierField || !poSupplierList || !poSupplierLabel || !supplierHistoryLabel || !addPoSupplierButton) {
    return;
  }

  const isLpb = isLampinCategory();
  poSupplierField.readOnly = isLpb;
  poSupplierLabel.hidden = isLpb;
  supplierHistoryLabel.hidden = isLpb;
  poSupplierList.hidden = !isLpb;
  addPoSupplierButton.hidden = !isLpb;

  if (!isLpb) {
    poSupplierField.value = state.data.poSupplier || "";
    poSupplierList.innerHTML = "";
    return;
  }

  state.data.poSuppliers = getPoSupplierEntries(state.data);
  if (!state.data.poSuppliers.length) {
    state.data.poSuppliers = [""];
  }
  syncPoSupplierDerivedText(state.data);
  poSupplierField.value = state.data.poSupplier || "";

  poSupplierList.innerHTML = "";
  const hasMultipleSuppliers = state.data.poSuppliers.length > 1;
  state.data.poSuppliers.forEach((supplier, index) => {
    const row = document.createElement("div");
    row.className = "po-supplier-row";

    const header = document.createElement("div");
    header.className = "po-supplier-row-header";

    const badge = document.createElement("span");
    badge.className = "po-supplier-badge";
    badge.textContent =
      (hasMultipleSuppliers ? getPoSubNumber(state.data.poNumber, index) : state.data.poNumber) ||
      `Supplier ${index + 1}`;
    header.appendChild(badge);

    if (state.data.poSuppliers.length > 1) {
      const removeButton = document.createElement("button");
      removeButton.type = "button";
      removeButton.className = "ghost-button danger compact-button";
      removeButton.textContent = "Remove";
      removeButton.addEventListener("click", () => {
        state.data.poSuppliers.splice(index, 1);
        if (!state.data.poSuppliers.length) {
          state.data.poSuppliers = [""];
        }
        syncPoSupplierDerivedText(state.data);
        renderPoSupplierEditor();
        renderDocuments();
      });
      header.appendChild(removeButton);
    }

    const textarea = document.createElement("textarea");
    textarea.rows = 4;
    textarea.value = supplier || "";
    textarea.placeholder = `Enter supplier ${index + 1}`;
    textarea.addEventListener("input", () => {
      state.data.poSuppliers[index] = textarea.value;
      syncPoSupplierDerivedText(state.data);
      poSupplierField.value = state.data.poSupplier || "";
      renderDocuments();
    });

    row.appendChild(header);
    row.appendChild(textarea);
    poSupplierList.appendChild(row);
  });
}

function rememberBillTo(value) {
  const normalized = normalizeMultiline(value);
  if (!normalized) return;

  const now = Date.now();
  state.billToHistory = [
    { value: normalized, savedAt: now },
    ...state.billToHistory.filter((entry) => entry.value !== normalized),
  ].slice(0, 20);

  saveBillToHistory();
  renderBillToHistory();
}

function rememberSupplier(value) {
  const normalized = normalizeMultiline(value);
  if (!normalized) return;

  const now = Date.now();
  state.supplierHistory = [
    { value: normalized, savedAt: now },
    ...state.supplierHistory.filter((entry) => entry.value !== normalized),
  ].slice(0, 20);

  saveSupplierHistory();
  renderSupplierHistory();
}

function deleteBillToHistoryEntry(value) {
  if (!value) return;
  state.billToHistory = state.billToHistory.filter((entry) => entry.value !== value);
  saveBillToHistory();
  renderBillToHistory();
}

function deleteSupplierHistoryEntry(value) {
  if (!value) return;
  state.supplierHistory = state.supplierHistory.filter((entry) => entry.value !== value);
  saveSupplierHistory();
  renderSupplierHistory();
}

function renderItemsEditor(container, stateKey) {
  if (!container) return;
  container.innerHTML = "";
  const allowRowRemoval = stateKey === "poItems";
  const allowSavedItems = stateKey === "poItems";
  const isStaticInvoicePricing = stateKey === "invoiceItems" && isLampinCategory();

  state.data[stateKey].forEach((item, index) => {
    const fragment = itemEditorTemplate.content.cloneNode(true);
    const row = fragment.querySelector(".item-row");
    const removeButton = fragment.querySelector(".remove-item-button");
    const savedItemField = row.querySelector('[data-field="savedItem"]');
    const savedItemDeleteButton = row.querySelector(".saved-item-delete-button");
    const lineNumber = fragment.querySelector(".item-line-number");
    if (lineNumber) {
      lineNumber.textContent = `Bil ${index + 1}`;
      lineNumber.hidden = !isLampinCategory();
    }

    if (allowSavedItems && savedItemField) {
      savedItemField.innerHTML = '<option value="">Select saved item</option>';
      state.savedPoItems.forEach((savedItem, savedIndex) => {
        const option = document.createElement("option");
        option.value = String(savedIndex);
        option.textContent = getPoItemHistoryLabel(savedItem);
        savedItemField.appendChild(option);
      });

      savedItemField.addEventListener("change", () => {
        const selectedIndex = Number(savedItemField.value);
        if (!Number.isInteger(selectedIndex) || !state.savedPoItems[selectedIndex]) return;
        const selectedItem = state.savedPoItems[selectedIndex];
        state.data[stateKey][index] = {
          ...state.data[stateKey][index],
          description: selectedItem.description,
          quantity: selectedItem.quantity,
          unitPrice: selectedItem.unitPrice,
          amount: selectedItem.amount,
        };
        renderItemsEditor(container, stateKey);
        renderDocuments();
      });

      savedItemDeleteButton?.addEventListener("click", () => {
        const selectedIndex = Number(savedItemField.value);
        if (!Number.isInteger(selectedIndex) || !state.savedPoItems[selectedIndex]) return;
        deleteSavedPoItem(selectedIndex);
        renderItemsEditor(container, stateKey);
        renderDocuments();
        setStatus("Saved item deleted.", "success");
      });
    } else {
      savedItemField?.closest(".item-field")?.remove();
      savedItemDeleteButton?.remove();
    }

    row.querySelectorAll("input, textarea").forEach((fieldElement) => {
      const field = fieldElement.dataset.field;
      if (isStaticInvoicePricing && field === "amount") {
        state.data[stateKey][index].amount = LPB_INVOICE_FIXED_AMOUNT;
        fieldElement.value = formatNumber(LPB_INVOICE_FIXED_AMOUNT);
        fieldElement.disabled = true;
      } else {
        fieldElement.value = item[field] ?? "";
        fieldElement.disabled = false;
      }
      fieldElement.addEventListener("input", () => {
        const value = field === "description" ? fieldElement.value : coerceNumber(fieldElement.value);

        state.data[stateKey][index][field] = value;

        if (!isStaticInvoicePricing && (field === "quantity" || field === "unitPrice")) {
          state.data[stateKey][index].amount =
            coerceNumber(state.data[stateKey][index].quantity) *
            coerceNumber(state.data[stateKey][index].unitPrice);
          const amountInput = row.querySelector('[data-field="amount"]');
          if (amountInput) {
            amountInput.value = formatNumber(state.data[stateKey][index].amount);
          }
        }

        if (isStaticInvoicePricing) {
          state.data[stateKey][index].amount = LPB_INVOICE_FIXED_AMOUNT;
          const amountInput = row.querySelector('[data-field="amount"]');
          if (amountInput) {
            amountInput.value = formatNumber(LPB_INVOICE_FIXED_AMOUNT);
          }
        }

        renderDocuments();
      });
    });

    if (allowRowRemoval) {
      removeButton.addEventListener("click", () => {
        if (state.data[stateKey].length <= 1) {
          state.data[stateKey] = [createBlankItem()];
        } else {
          state.data[stateKey].splice(index, 1);
        }
        renderItemsEditor(container, stateKey);
        renderDocuments();
      });
    } else {
      removeButton.remove();
    }

    container.appendChild(fragment);
  });
}

function companyBlockMarkup(title) {
  const companyProfile = getActiveCompanyProfile();
  return `
    <div class="document-top">
      <div class="logo-box">
        <img src="./assets/dvet-logo.png" alt="DVet logo" />
      </div>
      <div class="company-block">
        <h1 class="document-title">${safeText(title)}</h1>
        <p class="company-name">${safeText(companyProfile.name)}</p>
        <p>${safeText(companyProfile.registrationNo)}</p>
        ${companyProfile.addressLines.map((line) => `<p>${safeText(line)}</p>`).join("")}
        <p>${safeText(companyProfile.phone)}</p>
      </div>
    </div>
  `;
}

function itemRowsMarkup(items, { includePricing, includeCheckBox }) {
  return items
    .map((item, index) => {
      return `
        <tr>
          <td>${index + 1}</td>
          <td>
            <div class="item-description">
              <strong>${safeText(item.description)}</strong>
              <span class="item-code">(${safeText(item.code)})</span>
            </div>
          </td>
          <td class="numeric">${formatNumber(item.quantity)}</td>
          ${
            includePricing
              ? `
                <td class="numeric">${currency(item.unitPrice)}</td>
                <td class="numeric">${currency(item.amount)}</td>
              `
              : ""
          }
          ${includeCheckBox ? `<td class="numeric"><span class="check-square"></span></td>` : ""}
        </tr>
      `;
    })
    .join("");
}

function poMarkup() {
  const total = sumItems(state.data.poItems);
  const poPages = isLampinCategory()
    ? getPoSupplierPageEntries(state.data).filter((entry) => entry.supplier || entry.index === 0)
    : [{ supplier: state.data.poSupplier, poNumber: state.data.poNumber, index: 0 }];

  return poPages
    .map(
      (entry, pageIndex) => `
        <article class="document-sheet po-sheet">
          ${companyBlockMarkup("PURCHASE ORDER")}
          <section class="party-grid">
            <div class="party-box">
              <h4>Supplier:</h4>
              <p>${safeText(entry.supplier)}</p>
            </div>
            <div class="party-box">
              <h4>Ship To:</h4>
              <p>${safeText(state.data.poShipTo)}</p>
            </div>
            <div class="party-box">
              <h4>Details:</h4>
              <div class="detail-grid">
                <strong>PO#:</strong><span>${safeText(entry.poNumber)}</span>
                <strong>Date:</strong><span>${safeText(state.data.poDate)}</span>
                <strong>By:</strong><span>${safeText(state.data.poPreparedBy)}</span>
              </div>
            </div>
          </section>

          <table class="document-table">
            <thead>
              <tr>
                <th>No</th>
                <th>Description</th>
                <th class="numeric">Unit</th>
                <th class="numeric">Price</th>
                <th class="numeric">Amount</th>
              </tr>
            </thead>
            <tbody>${itemRowsMarkup(state.data.poItems, { includePricing: true, includeCheckBox: false })}</tbody>
          </table>

          <div class="summary-grid">
            <div class="summary-line"><span></span><strong>Total (MYR): ${currency(total)}</strong></div>
          </div>

          <p class="footer-note">Generated automatically. No signature is required.</p>
          <p class="page-note">Page ${pageIndex + 1}/${poPages.length}</p>
        </article>
      `,
    )
    .join("");
}

function invoiceMarkup() {
  const companyProfile = getActiveCompanyProfile();
  const total = sumItems(state.data.invoiceItems);

  return `
    <article class="document-sheet invoice-sheet">
      ${companyBlockMarkup("INVOICE")}
      <section class="party-grid">
        <div class="party-box">
          <h4>Bill To:</h4>
          <p>${safeText(state.data.invoiceBillTo)}</p>
        </div>
        <div class="party-box">
          <h4>Ship To:</h4>
          <p>${safeText(state.data.invoiceShipTo)}</p>
        </div>
        <div class="party-box">
          <h4>Details:</h4>
          <div class="detail-grid invoice-detail-grid">
            <strong>INV#:</strong><span>${safeText(state.data.invoiceNumber)}</span>
            <strong>Date:</strong><span>${safeText(state.data.invoiceDate || state.data.poDate)}</span>
            <strong>By:</strong><span>${safeText(state.data.invoicePreparedBy)}</span>
            <strong>e-Invoice Type:</strong><span>Consolidate</span>
            <strong>e-Invoice:</strong><span>-</span>
            <strong>Status:</strong><span>${safeText(state.data.invoiceStatus)}</span>
            <strong>PO JHEV:</strong><span>${safeText(state.data.invoicePoJhev)}</span>
          </div>
        </div>
      </section>

      <table class="document-table">
        <thead>
          <tr>
            <th>No</th>
            <th>Description</th>
            <th class="numeric">Unit</th>
            <th class="numeric">Price</th>
            <th class="numeric">Amount</th>
          </tr>
        </thead>
        <tbody>${itemRowsMarkup(state.data.invoiceItems, { includePricing: true, includeCheckBox: false })}</tbody>
      </table>

      <div class="summary-grid">
        <div class="summary-line"><span>${totalQuantity(state.data.invoiceItems)}</span><strong>Total (MYR): ${currency(total)}</strong></div>
        <div class="summary-line"><span>Payment:</span><strong>0.00</strong></div>
        <div class="summary-line"><span>Balance:</span><strong class="total-accent">${currency(total)}</strong></div>
      </div>

      <div class="terms-layout">
        <div class="terms-block">
          <h4>Terms &amp; Conditions</h4>
          <p>All cheque should be crossed and made payable to DVET TRADING SDN BHD
Our Bank Details: ${safeText(companyProfile.bankDetails)}</p>
          <p class="footer-note">Generated automatically. No signature is required.</p>
        </div>
      </div>

      <p class="page-note">Page 1/1</p>
    </article>
  `;
}

function deliveryOrderMarkup() {
  const companyProfile = getActiveCompanyProfile();
  return `
    <article class="document-sheet do-sheet">
      ${companyBlockMarkup("DELIVERY ORDER")}
      <section class="party-grid">
        <div class="party-box">
          <h4>Bill To:</h4>
          <p>${safeText(state.data.doBillTo)}</p>
        </div>
        <div class="party-box">
          <h4>Ship To:</h4>
          <p>${safeText(state.data.doShipTo)}</p>
        </div>
        <div class="party-box">
          <h4>Details:</h4>
          <div class="detail-grid">
            <strong>DO#:</strong><span>${safeText(state.data.doNumber)}</span>
            <strong>Ship date:</strong><span>${safeText(state.data.doShipDate || state.data.poDate)}</span>
            <strong>Ship by:</strong><span>LORRY</span>
            <strong>Status:</strong><span>${safeText(state.data.doStatus)}</span>
            <strong>INV#:</strong><span>${safeText(state.data.doInvoiceNumber || state.data.invoiceNumber)}</span>
            <strong>INV# date:</strong><span>${safeText(state.data.doInvoiceDate || state.data.invoiceDate || state.data.poDate)}</span>
            <strong>PO JHEV:</strong><span>${safeText(state.data.doPoJhev)}</span>
          </div>
        </div>
      </section>

      <table class="document-table">
        <thead>
          <tr>
            <th>No</th>
            <th>Description</th>
            <th class="numeric">Unit</th>
            <th class="numeric">Check</th>
          </tr>
        </thead>
        <tbody>${itemRowsMarkup(state.data.doItems, { includePricing: false, includeCheckBox: true })}</tbody>
      </table>

      <div class="summary-grid">
        <div class="summary-line"><span>Total :</span><strong>${totalQuantity(state.data.doItems)}</strong></div>
      </div>

      <div class="terms-block">
        <h4>Terms &amp; Conditions</h4>
        <p>- GOODS SOLD ARE NOT RETURNABLE
- PLEASE KEEP THE INVOICE FOR WARRANTY PURPOSE
- WE WILL NOT ENTERTAINED ANY CLAIM WITHOUT OFFICIAL INVOICE</p>
      </div>

      <p class="footer-note">Generated automatically. No signature is required.</p>
      <p class="page-note">Page 1/1</p>
    </article>
  `;
}

function renderDocuments() {
  syncStateFromForm();
  syncPoSupplierDerivedText(state.data);
  state.data.poItems = inferItemAmounts(state.data.poItems);
  state.data.invoiceItems = isLampinCategory()
    ? applyLampinInvoiceAmount(state.data.invoiceItems)
    : inferItemAmounts(state.data.invoiceItems);
  state.data.doItems = inferItemAmounts(state.data.doItems);
  poPreview.innerHTML = poMarkup();
  invoicePreview.innerHTML = invoiceMarkup();
  doPreview.innerHTML = deliveryOrderMarkup();
}

function applyParsedData(parsed) {
  const nextPoSupplier = parsed.poSupplier ?? state.data.poSupplier;
  state.data = {
    ...state.data,
    ...parsed,
    poSuppliers:
      parsed.poSuppliers?.length
        ? [...parsed.poSuppliers]
        : isLampinCategory(parsed.category || state.data.category)
          ? [String(nextPoSupplier || "")]
          : [String(nextPoSupplier || "")],
    poItems: inferItemAmounts(parsed.poItems?.length ? parsed.poItems : state.data.poItems),
    invoiceItems: inferItemAmounts(parsed.invoiceItems?.length ? parsed.invoiceItems : state.data.invoiceItems),
    doItems: inferItemAmounts(parsed.doItems?.length ? parsed.doItems : state.data.doItems),
  };
  syncPoSupplierDerivedText(state.data);
  syncFormFromState();
  renderItemsEditor(poItemsEditor, "poItems");
  renderItemsEditor(invoiceItemsEditor, "invoiceItems");
  renderItemsEditor(doItemsEditor, "doItems");
  renderDocuments();
}

function extractTextLines(items) {
  const grouped = new Map();

  items.forEach((item) => {
    const value = String(item.str || "").trim();
    if (!value) return;
    const y = Math.round(item.transform?.[5] || 0);
    const bucket = grouped.get(y) || [];
    bucket.push({ x: item.transform?.[4] || 0, value });
    grouped.set(y, bucket);
  });

  return [...grouped.entries()]
    .sort((a, b) => b[0] - a[0])
    .map(([, lineItems]) =>
      lineItems
        .sort((a, b) => a.x - b.x)
        .map((entry) => entry.value)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim(),
    )
    .filter(Boolean);
}

function extractRegionLines(items, bounds) {
  const grouped = new Map();

  items.forEach((item) => {
    const value = String(item.str || "").trim();
    if (!value) return;
    const x = item.transform?.[4] || 0;
    const y = item.transform?.[5] || 0;
    if (x < bounds.minX || x > bounds.maxX || y > bounds.maxY || y < bounds.minY) return;

    const bucketY = Math.round(y);
    const bucket = grouped.get(bucketY) || [];
    bucket.push({ x, value });
    grouped.set(bucketY, bucket);
  });

  return [...grouped.entries()]
    .sort((a, b) => b[0] - a[0])
    .map(([, lineItems]) =>
      lineItems
        .sort((a, b) => a.x - b.x)
        .map((entry) => entry.value)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim(),
    )
    .filter(Boolean);
}

function findTextAnchor(items, matcher) {
  return items.find((item) => matcher(String(item.str || "").trim()));
}

function extractSection(lines, title, stopTitles) {
  const startIndex = lines.findIndex((line) => line.toUpperCase().includes(title.toUpperCase()));
  if (startIndex === -1) return "";

  const content = [];
  for (let index = startIndex + 1; index < lines.length; index += 1) {
    const line = lines[index].trim();
    const shouldStop = stopTitles.some((stop) => line.toUpperCase().includes(stop.toUpperCase()));
    if (shouldStop) break;
    if (line) content.push(line);
  }
  return normalizeMultiline(content.join("\n"));
}

function cleanShipToSection(text) {
  const lines = normalizeMultiline(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const cleaned = lines.filter((line) => {
    const upper = line.toUpperCase();
    if (upper.startsWith("DETAILS")) return false;
    if (upper.startsWith("PO#")) return false;
    if (upper.startsWith("DATE")) return false;
    if (upper.startsWith("BY")) return false;
    if (upper === "NO") return false;
    return true;
  });

  return normalizeMultiline(cleaned.join("\n"));
}

function cleanSupplierSection(text) {
  const lines = normalizeMultiline(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  const cleaned = lines.filter((line) => {
    const upper = line.toUpperCase();
    if (upper.startsWith("SHIP TO")) return false;
    if (upper.startsWith("DETAILS")) return false;
    if (upper.startsWith("PO#")) return false;
    if (upper.startsWith("DATE")) return false;
    if (upper.startsWith("BY")) return false;
    if (upper.startsWith("INV#")) return false;
    if (upper.startsWith("DO#")) return false;
    return true;
  });

  return normalizeMultiline(cleaned.join("\n"));
}

function extractShipToFromTextItems(items) {
  const shipAnchor =
    findTextAnchor(items, (text) => /^ship$/i.test(text)) ||
    findTextAnchor(items, (text) => /^ship to:?$/i.test(text)) ||
    findTextAnchor(items, (text) => /ship/i.test(text));
  const detailsAnchor = findTextAnchor(items, (text) => /^details:?$/i.test(text) || /^details$/i.test(text));
  const tableAnchor = findTextAnchor(items, (text) => /^no$/i.test(text));

  if (!shipAnchor || !detailsAnchor || !tableAnchor) {
    return "";
  }

  const shipX = shipAnchor.transform?.[4] || 0;
  const shipY = shipAnchor.transform?.[5] || 0;
  const detailsX = detailsAnchor.transform?.[4] || 9999;
  const tableY = tableAnchor.transform?.[5] || 0;

  const lines = extractRegionLines(items, {
    minX: shipX - 8,
    maxX: detailsX - 10,
    maxY: shipY - 8,
    minY: tableY + 12,
  });

  return cleanShipToSection(lines.join("\n"));
}

function extractSupplierFromTextItems(items) {
  const supplierAnchor =
    findTextAnchor(items, (text) => /^supplier:?$/i.test(text)) ||
    findTextAnchor(items, (text) => /^supplier$/i.test(text));
  const shipAnchor =
    findTextAnchor(items, (text) => /^ship$/i.test(text)) ||
    findTextAnchor(items, (text) => /^ship to:?$/i.test(text)) ||
    findTextAnchor(items, (text) => /ship/i.test(text));
  const tableAnchor = findTextAnchor(items, (text) => /^no$/i.test(text));

  if (!supplierAnchor || !shipAnchor || !tableAnchor) {
    return "";
  }

  const supplierX = supplierAnchor.transform?.[4] || 0;
  const supplierY = supplierAnchor.transform?.[5] || 0;
  const shipX = shipAnchor.transform?.[4] || 9999;
  const tableY = tableAnchor.transform?.[5] || 0;

  const lines = extractRegionLines(items, {
    minX: Math.max(0, supplierX - 8),
    maxX: shipX - 14,
    maxY: supplierY - 8,
    minY: tableY + 12,
  });

  return cleanSupplierSection(lines.join("\n"));
}

function extractItems(lines) {
  const items = [];

  for (const line of lines) {
    const match = line.match(/^(\d+)\s+(.+?)\s+\(?(\d{6,})\)?\s+(\d+)\s+([\d,]+(?:\.\d{2})?)\s+([\d,]+(?:\.\d{2})?)$/);

    if (!match) continue;

    items.push({
      description: match[2].trim(),
      code: match[3].trim(),
      quantity: coerceNumber(match[4].replace(/,/g, "")),
      unitPrice: coerceNumber(match[5].replace(/,/g, "")),
      amount: coerceNumber(match[6].replace(/,/g, "")),
    });
  }

  return items;
}

function extractItemsFromTextItems(items) {
  const rows = new Map();

  items.forEach((item) => {
    const value = String(item.str || "").trim();
    if (!value) return;
    const y = Math.round(item.transform?.[5] || 0);
    const row = rows.get(y) || [];
    row.push({ x: item.transform?.[4] || 0, value });
    rows.set(y, row);
  });

  return [...rows.entries()]
    .sort((a, b) => b[0] - a[0])
    .map(([, rowItems]) => rowItems.sort((a, b) => a.x - b.x))
    .filter((rowItems) => rowItems.length >= 5 && /^\d+$/.test(rowItems[0]?.value || ""))
    .map((rowItems) => {
      const values = rowItems.map((entry) => entry.value);
      const quantity = values.at(-3);
      const unitPrice = values.at(-2);
      const amount = values.at(-1);
      const code = values.find((value) => /^\d{6,}$/.test(value)) || "";
      const descriptionParts = values.slice(1, Math.max(2, values.length - 3)).filter((value) => value !== code);

      return {
        description: descriptionParts.join(" ").trim(),
        code,
        quantity: coerceNumber(String(quantity || "").replace(/,/g, "")),
        unitPrice: coerceNumber(String(unitPrice || "").replace(/,/g, "")),
        amount: coerceNumber(String(amount || "").replace(/,/g, "")),
      };
    })
    .filter((item) => item.description && item.quantity >= 0);
}

function extractPoDataFromLines(lines, rawItems = []) {
  const rawText = lines.join("\n");
  const poNumber = state.data.poNumber || rawText.match(/\bPO\d{5,}\b/i)?.[0] || "";
  const poDate = rawText.match(/\b\d{2}\/\d{2}\/\d{4}\b/)?.[0] || state.data.poDate;
  const preparedBy = rawText.match(/\b(?:By|Prepared By)\s*:?\s*([A-Za-z][A-Za-z ]{2,})/i)?.[1]?.trim() || "System";

  const supplierFromRegion = extractSupplierFromTextItems(rawItems);
  const supplier =
    supplierFromRegion || cleanSupplierSection(extractSection(lines, "Supplier", ["Ship To", "Details", "No"]));
  const shipToFromRegion = extractShipToFromTextItems(rawItems);
  const shipTo = shipToFromRegion || cleanShipToSection(extractSection(lines, "Ship To", ["Details", "No"]));
  const items = extractItems(lines);
  const normalizedItems = items.length
    ? items.map((item) => ({ ...item, unitPrice: item.unitPrice, amount: item.amount }))
    : state.data.poItems;

  const parsed = {
    poNumber,
    poDate,
    poSupplier: supplier || "",
    poShipTo: shipTo || "",
    poPreparedBy: preparedBy,
    invoiceNumber: state.data.invoiceNumber || "",
    invoiceDate: poDate,
    invoiceBillTo: state.data.invoiceBillTo || "",
    invoiceShipTo: shipTo || "",
    invoicePreparedBy: preparedBy,
    doNumber: state.data.doNumber || "",
    doBillTo: state.data.doBillTo || "",
    doShipTo: shipTo || "",
    doShipDate: poDate,
    doInvoiceNumber: state.data.invoiceNumber || "",
    doInvoiceDate: poDate,
    invoiceStatus: "Pending",
    doStatus: "Delivered",
    poItems: normalizedItems,
    invoiceItems: structuredClone(normalizedItems),
    doItems: structuredClone(normalizedItems),
  };

  return parsed;
}

async function parsePdf(file) {
  const pdfjsLib = await ensurePdfJs();
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  let allLines = [];
  let rawItems = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    rawItems = rawItems.concat(content.items);
    const lines = extractTextLines(content.items);
    allLines = allLines.concat(lines);
  }

  state.extractedText = allLines.join("\n");
  const parsed = extractPoDataFromLines(allLines, rawItems);
  const tableItems = extractItemsFromTextItems(rawItems);
  if (tableItems.length) {
    parsed.poItems = inferItemAmounts(tableItems);
    parsed.invoiceItems = inferItemAmounts(structuredClone(tableItems));
    parsed.doItems = inferItemAmounts(structuredClone(tableItems));
  }
  return parsed;
}

async function handleParseClick() {
  const file = poFileInput.files?.[0];

  if (!file) {
    setStatus("Please select a PO PDF file first.", "warn");
    return;
  }

  setStatus(`Reading file: ${file.name}`, "progress");

  try {
    const parsed = await parsePdf(file);
    applyParsedData(parsed);
    setStatus(
      "PO was read successfully. Review the information and items before downloading the invoice and delivery order.",
      "success",
    );
  } catch (error) {
    console.error(error);
    setStatus(
      "Failed to read the PDF automatically. Please review and edit the fields manually.",
      "error",
    );
  }
}

function syncPoToDocuments() {
  if (!state.data.poNumber || !state.data.invoiceNumber || !state.data.doNumber) {
    assignDraftRunningNumbers(state.data);
  }
  state.data.invoiceDate = state.data.invoiceDate || state.data.poDate;
  state.data.invoiceShipTo = state.data.poShipTo;
  state.data.invoicePreparedBy = state.data.invoicePreparedBy || state.data.poPreparedBy;
  state.data.invoiceItems = isLampinCategory()
    ? applyLampinInvoiceAmount(structuredClone(state.data.poItems))
    : inferItemAmounts(structuredClone(state.data.poItems));

  state.data.doShipTo = state.data.poShipTo;
  state.data.doShipDate = state.data.doShipDate || state.data.poDate;
  state.data.doInvoiceNumber = state.data.invoiceNumber;
  state.data.doInvoiceDate = state.data.doInvoiceDate || state.data.invoiceDate || state.data.poDate;
  state.data.doItems = inferItemAmounts(structuredClone(state.data.poItems));

  syncFormFromState();
  renderItemsEditor(invoiceItemsEditor, "invoiceItems");
  renderItemsEditor(doItemsEditor, "doItems");
  renderDocuments();
}

function splitLines(text = "") {
  return String(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
}

function getRecipientName() {
  return sanitizeFilenamePart(
    splitLines(state.data.invoiceShipTo)[0] ||
      splitLines(state.data.doShipTo)[0] ||
      splitLines(state.data.poShipTo)[0] ||
      splitLines(state.data.invoiceBillTo)[0] ||
      "RECIPIENT",
  );
}

function getInvoiceDigits() {
  const raw = String(state.data.invoiceNumber || "").trim();
  const digits = raw.match(/\d+/g)?.join("") || sanitizeFilenamePart(raw) || "0000000";
  return digits;
}

function getInvoiceFilename() {
  return sanitizeFilenamePart(`INV${getInvoiceDigits()} ${getRecipientName()}`) + ".pdf";
}

function getDoFilename() {
  return sanitizeFilenamePart(`DO${getInvoiceDigits()} ${getRecipientName()}`) + ".pdf";
}

function getCompleteSetFilename() {
  return sanitizeFilenamePart(`INV & DO (${getRecipientName()})`) + ".pdf";
}

function getPoFilename() {
  return sanitizeFilenamePart(`PO${String(state.data.poNumber || "").replace(/^PO/i, "")} ${getRecipientName()}`) + ".pdf";
}

function drawWrappedText(doc, text, x, y, options = {}) {
  const lines = Array.isArray(text) ? text : doc.splitTextToSize(String(text || ""), options.maxWidth || 60);
  doc.text(lines, x, y);
  return y + lines.length * (options.lineHeight || 5);
}

function getWrappedHeight(doc, text, maxWidth, lineHeight = 5) {
  const lines = Array.isArray(text) ? text : doc.splitTextToSize(String(text || ""), maxWidth);
  return { lines, height: Math.max(lineHeight, lines.length * lineHeight) };
}

function getPdfColumns() {
  return {
    left: { x: 10, width: 44 },
    middle: { x: 80, width: 44 },
    right: { x: 150, width: 46 },
  };
}

function createPdfDocument(title) {
  if (!window.jspdf?.jsPDF) {
    throw new Error("jsPDF library is not loaded.");
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
  drawPdfHeader(doc, title);
  return doc;
}

function drawPdfHeader(doc, title) {
  const companyProfile = getActiveCompanyProfile();
  doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, 210, 297, "F");
  if (state.logoDataUrl) {
    doc.addImage(state.logoDataUrl, "PNG", 10, 10, 28, 18);
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(20);
  doc.text(title, 190, 18, { align: "right" });
  doc.setFontSize(10);
  doc.text(companyProfile.name, 190, 26, { align: "right" });
  doc.setFont("helvetica", "normal");
  doc.text(companyProfile.registrationNo, 190, 30, { align: "right" });
  companyProfile.addressLines.forEach((line, index) => {
    doc.text(line, 190, 35 + index * 4.5, { align: "right" });
  });
  doc.text(companyProfile.phone, 190, 53, { align: "right" });
}

function addPdfPage(doc) {
  doc.addPage("a4", "portrait");
}

function drawItemTable(doc, items, options) {
  const { includePricing = false, includeCheck = false, startY = 86 } = options;
  const rowHeight = 14;
  const colX = includePricing
    ? { no: 10, desc: 22, qty: 132, unitPrice: 156, amount: 188 }
    : { no: 10, desc: 22, qty: 160, check: 188 };

  doc.setDrawColor(212, 212, 212);
  doc.setLineWidth(0.2);
  doc.line(10, startY - 6, 200, startY - 6);
  doc.line(10, startY, 200, startY);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("No", colX.no, startY - 1.5);
  doc.text("Description", colX.desc, startY - 1.5);
  doc.text("Unit", colX.qty, startY - 1.5, { align: "right" });
  if (includePricing) {
    doc.text("Price", colX.unitPrice, startY - 1.5, { align: "right" });
    doc.text("Amount", colX.amount, startY - 1.5, { align: "right" });
  }
  if (includeCheck) {
    doc.text("Check", colX.check, startY - 1.5, { align: "right" });
  }

  let currentY = startY + 8;
  doc.setFont("helvetica", "normal");
  items.forEach((item, index) => {
    const description = `${item.description}${item.code ? ` (${item.code})` : ""}`;
    doc.text(String(index + 1), colX.no, currentY);
    const wrapped = doc.splitTextToSize(description, includePricing ? 95 : 120);
    doc.text(wrapped, colX.desc, currentY);
    doc.text(formatNumber(item.quantity), colX.qty, currentY, { align: "right" });
    if (includePricing) {
      doc.text(currency(item.unitPrice), colX.unitPrice, currentY, { align: "right" });
      doc.text(currency(item.amount), colX.amount, currentY, { align: "right" });
    }
    if (includeCheck) {
      doc.rect(colX.check - 5, currentY - 4, 4.5, 4.5);
    }
    const usedHeight = Math.max(rowHeight, wrapped.length * 5.2);
    doc.line(10, currentY + usedHeight - 7, 200, currentY + usedHeight - 7);
    currentY += usedHeight;
  });

  return currentY;
}

function renderInvoicePage(doc) {
  const companyProfile = getActiveCompanyProfile();
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  const columns = getPdfColumns();
  const sectionTop = 64;
  const contentTop = 70;
  doc.text("Bill To:", columns.left.x, sectionTop);
  doc.text("Ship To:", columns.middle.x, sectionTop);
  doc.text("Details:", columns.right.x, sectionTop);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  const billToBlock = getWrappedHeight(doc, splitLines(state.data.invoiceBillTo), columns.left.width, 4.5);
  const shipToBlock = getWrappedHeight(doc, splitLines(state.data.invoiceShipTo), columns.middle.width, 4.5);
  const detailsBlock = getWrappedHeight(
    doc,
    [
      `INV#: ${state.data.invoiceNumber}`,
      `Date: ${state.data.invoiceDate || state.data.poDate}`,
      `By: ${state.data.invoicePreparedBy}`,
      "e-Invoice Type: Consolidate",
      "e-Invoice: -",
      `Status: ${state.data.invoiceStatus}`,
      `PO JHEV: ${state.data.invoicePoJhev || ""}`,
    ],
    columns.right.width,
    4.5,
  );

  drawWrappedText(doc, billToBlock.lines, columns.left.x, contentTop, { lineHeight: 4.5 });
  drawWrappedText(doc, shipToBlock.lines, columns.middle.x, contentTop, { lineHeight: 4.5 });
  drawWrappedText(doc, detailsBlock.lines, columns.right.x, contentTop, { lineHeight: 4.5 });

  const tableStartY = Math.max(
    contentTop + billToBlock.height,
    contentTop + shipToBlock.height,
    contentTop + detailsBlock.height,
  ) + 12;

  let currentY = drawItemTable(doc, state.data.invoiceItems, { includePricing: true, startY: tableStartY });
  const total = sumItems(state.data.invoiceItems);
  currentY += 4;
  doc.setFont("helvetica", "bold");
  doc.text(`${totalQuantity(state.data.invoiceItems)}`, 132, currentY, { align: "right" });
  doc.text(`Total (MYR):`, 170, currentY, { align: "right" });
  doc.text(currency(total), 196, currentY, { align: "right" });
  currentY += 6;
  doc.setFont("helvetica", "normal");
  doc.text("Payment:", 170, currentY, { align: "right" });
  doc.text("0.00", 196, currentY, { align: "right" });
  currentY += 6;
  doc.text("Balance:", 170, currentY, { align: "right" });
  doc.setTextColor(198, 50, 50);
  doc.text(currency(total), 196, currentY, { align: "right" });
  doc.setTextColor(0, 0, 0);
  currentY += 12;
  doc.setFont("helvetica", "bold");
  doc.text("Terms & Conditions", 10, currentY);
  doc.setFont("helvetica", "normal");
  currentY += 6;
  drawWrappedText(
    doc,
    [
      "All cheque should be crossed and made payable to DVET TRADING SDN BHD",
      `Our Bank Details: ${companyProfile.bankDetails}`,
      "",
      "Generated automatically. No signature is required.",
    ],
    10,
    currentY,
    { maxWidth: 150, lineHeight: 4.8 },
  );
  doc.setFontSize(9);
  doc.text("Page 1/1", 105, 289, { align: "center" });
}

function renderPoPage(doc, poPageEntry = null, pageNumber = 1, totalPages = 1) {
  const activePoEntry = poPageEntry || {
    supplier: state.data.poSupplier,
    poNumber: state.data.poNumber,
    index: 0,
  };
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  const columns = getPdfColumns();
  const sectionTop = 64;
  const contentTop = 70;
  doc.text("Supplier:", columns.left.x, sectionTop);
  doc.text("Ship To:", columns.middle.x, sectionTop);
  doc.text("Details:", columns.right.x, sectionTop);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  const supplierBlock = getWrappedHeight(doc, splitLines(activePoEntry.supplier), columns.left.width, 4.5);
  const shipToBlock = getWrappedHeight(doc, splitLines(state.data.poShipTo), columns.middle.width, 4.5);
  const detailsBlock = getWrappedHeight(
    doc,
    [
      `PO#: ${activePoEntry.poNumber}`,
      `Date: ${state.data.poDate}`,
      `By: ${state.data.poPreparedBy}`,
    ],
    columns.right.width,
    4.5,
  );

  drawWrappedText(doc, supplierBlock.lines, columns.left.x, contentTop, { lineHeight: 4.5 });
  drawWrappedText(doc, shipToBlock.lines, columns.middle.x, contentTop, { lineHeight: 4.5 });
  drawWrappedText(doc, detailsBlock.lines, columns.right.x, contentTop, { lineHeight: 4.5 });

  const tableStartY = Math.max(
    contentTop + supplierBlock.height,
    contentTop + shipToBlock.height,
    contentTop + detailsBlock.height,
  ) + 12;

  let currentY = drawItemTable(doc, state.data.poItems, { includePricing: true, startY: tableStartY });
  const total = sumItems(state.data.poItems);
  currentY += 6;
  doc.setFont("helvetica", "bold");
  doc.text("Total (MYR):", 170, currentY, { align: "right" });
  doc.text(currency(total), 196, currentY, { align: "right" });
  doc.setFontSize(9);
  doc.text(`Page ${pageNumber}/${totalPages}`, 105, 289, { align: "center" });
}

function getRenderablePoPages() {
  if (!isLampinCategory()) {
    return [{ supplier: state.data.poSupplier, poNumber: state.data.poNumber, index: 0 }];
  }
  return getPoSupplierPageEntries(state.data).filter((entry) => entry.supplier || entry.index === 0);
}

function renderAllPoPages(doc) {
  const poPages = getRenderablePoPages();
  poPages.forEach((pageEntry, index) => {
    if (index > 0) {
      addPdfPage(doc);
      drawPdfHeader(doc, "PURCHASE ORDER");
    }
    renderPoPage(doc, pageEntry, index + 1, poPages.length);
  });
  return poPages.length;
}

async function downloadPoPdf(filename) {
  await ensureLogoDataUrl();
  const doc = createPdfDocument("PURCHASE ORDER");
  renderAllPoPages(doc);
  return doc.save(filename, { returnPromise: true });
}

async function downloadInvoicePdf(filename) {
  await ensureLogoDataUrl();
  const doc = createPdfDocument("INVOICE");
  renderInvoicePage(doc);
  return doc.save(filename, { returnPromise: true });
}

function renderDeliveryOrderPage(doc) {
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  const columns = getPdfColumns();
  const sectionTop = 64;
  const contentTop = 70;
  doc.text("Bill To:", columns.left.x, sectionTop);
  doc.text("Ship To:", columns.middle.x, sectionTop);
  doc.text("Details:", columns.right.x, sectionTop);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  const billToBlock = getWrappedHeight(doc, splitLines(state.data.doBillTo), columns.left.width, 4.5);
  const shipToBlock = getWrappedHeight(doc, splitLines(state.data.doShipTo), columns.middle.width, 4.5);
  const detailsBlock = getWrappedHeight(
    doc,
    [
      `DO#: ${state.data.doNumber}`,
      `Ship date: ${state.data.doShipDate || state.data.poDate}`,
      "Ship by: LORRY",
      `Status: ${state.data.doStatus}`,
      `INV#: ${state.data.doInvoiceNumber || state.data.invoiceNumber}`,
      `INV# date: ${state.data.doInvoiceDate || state.data.invoiceDate || state.data.poDate}`,
      `PO JHEV: ${state.data.doPoJhev || ""}`,
    ],
    columns.right.width,
    4.5,
  );

  drawWrappedText(doc, billToBlock.lines, columns.left.x, contentTop, { lineHeight: 4.5 });
  drawWrappedText(doc, shipToBlock.lines, columns.middle.x, contentTop, { lineHeight: 4.5 });
  drawWrappedText(doc, detailsBlock.lines, columns.right.x, contentTop, { lineHeight: 4.5 });

  const tableStartY = Math.max(
    contentTop + billToBlock.height,
    contentTop + shipToBlock.height,
    contentTop + detailsBlock.height,
  ) + 12;

  let currentY = drawItemTable(doc, state.data.doItems, { includeCheck: true, startY: tableStartY });
  currentY += 5;
  doc.setFont("helvetica", "bold");
  doc.text("Total :", 170, currentY, { align: "right" });
  doc.text(`${totalQuantity(state.data.doItems)}`, 196, currentY, { align: "right" });
  currentY += 12;
  doc.text("Terms & Conditions", 10, currentY);
  doc.setFont("helvetica", "normal");
  drawWrappedText(
    doc,
    [
      "- GOODS SOLD ARE NOT RETURNABLE",
      "- PLEASE KEEP THE INVOICE FOR WARRANTY PURPOSE",
      "- WE WILL NOT ENTERTAINED ANY CLAIM WITHOUT OFFICIAL INVOICE",
      "",
      "Generated automatically. No signature is required.",
    ],
    10,
    currentY + 6,
    { maxWidth: 150, lineHeight: 4.8 },
  );
  doc.setFontSize(9);
  doc.text("Page 1/1", 105, 289, { align: "center" });
}

async function downloadDeliveryOrderPdf(filename) {
  await ensureLogoDataUrl();
  const doc = createPdfDocument("DELIVERY ORDER");
  renderDeliveryOrderPage(doc);
  return doc.save(filename, { returnPromise: true });
}

async function downloadCompleteSetPdf(filename) {
  await ensureLogoDataUrl();
  const doc = createPdfDocument("PURCHASE ORDER");
  renderAllPoPages(doc);
  addPdfPage(doc);
  drawPdfHeader(doc, "INVOICE");
  renderInvoicePage(doc);
  addPdfPage(doc);
  drawPdfHeader(doc, "DELIVERY ORDER");
  renderDeliveryOrderPage(doc);
  return doc.save(filename, { returnPromise: true });
}

document.getElementById("parseButton").addEventListener("click", handleParseClick);
companyInputs.forEach((input) => {
  input.addEventListener("change", () => {
    if (!input.checked) {
      input.checked = true;
      return;
    }

    companyInputs.forEach((otherInput) => {
      if (otherInput !== input) {
        otherInput.checked = false;
      }
    });

    state.currentRecordId = "";
    state.data = {
      ...structuredClone(defaultState),
      company: input.value,
    };
    syncCompanyControlsFromState();
    syncCategoryControlsFromState();
    setCategoryGateState();
    syncFormFromState();
    renderItemsEditor(poItemsEditor, "poItems");
    renderItemsEditor(invoiceItemsEditor, "invoiceItems");
    renderItemsEditor(doItemsEditor, "doItems");
    renderDocuments();
    renderSavedRecords();
    setStatus(`${getCompanyMeta(input.value)?.label || "Company"} selected. Please choose a category next.`, "success");
  });
});

categoryInputs.forEach((input) => {
  input.addEventListener("change", () => {
    if (!state.data.company) {
      input.checked = false;
      setStatus("Please select one company first.", "warn");
      return;
    }
    if (!input.checked) {
      input.checked = true;
      return;
    }

    categoryInputs.forEach((otherInput) => {
      if (otherInput !== input) {
        otherInput.checked = false;
      }
    });

    state.currentRecordId = "";
    state.data = createNewDocumentStateForCategory(input.value);
    syncCompanyControlsFromState();
    syncCategoryControlsFromState();
    setCategoryGateState();
    syncFormFromState();
    renderItemsEditor(poItemsEditor, "poItems");
    renderItemsEditor(invoiceItemsEditor, "invoiceItems");
    renderItemsEditor(doItemsEditor, "doItems");
    renderDocuments();
    renderSavedRecords();
    setStatus(`Category ${getCategoryMeta(input.value)?.label || ""} selected. You can now fill in the document.`, "success");
  });
});

newPoButton?.addEventListener("click", () => {
  if (!state.data.category) {
    setStatus("Please select one main category first.", "warn");
    return;
  }
  state.currentRecordId = "";
  state.data = createBlankDraftWithCurrentNumbers(state.data);
  syncCategoryControlsFromState();
  setCategoryGateState();
  syncFormFromState();
  renderItemsEditor(poItemsEditor, "poItems");
  renderItemsEditor(invoiceItemsEditor, "invoiceItems");
  renderItemsEditor(doItemsEditor, "doItems");
  renderDocuments();
  renderSavedRecords();
  setStatus("New PO form is ready with fresh running numbers.", "success");
});

syncPoButton?.addEventListener("click", () => {
  syncStateFromForm();
  syncPoToDocuments();
  setStatus("PO data synced to Invoice and Delivery Order.", "success");
});

addPoItemButton?.addEventListener("click", () => {
  syncStateFromForm();
  state.data.poItems.push(createBlankItem());
  renderItemsEditor(poItemsEditor, "poItems");
  renderDocuments();
  setStatus("New PO item row added.", "success");
});

addPoSupplierButton?.addEventListener("click", () => {
  if (!isLampinCategory()) return;
  syncStateFromForm();
  state.data.poSuppliers = [...getPoSupplierEntries(state.data), ""];
  syncPoSupplierDerivedText(state.data);
  renderPoSupplierEditor();
  renderDocuments();
  setStatus("New supplier row added for Lampin Pakai Buang.", "success");
});

saveRecordButton?.addEventListener("click", () => {
  saveCurrentRecord({ showStatus: true });
});

exportRecordsButton?.addEventListener("click", () => {
  exportSavedRecordsToExcel();
});

documentHistoryList?.addEventListener("change", (event) => {
  const recordId = event.target.value;
  if (!recordId) return;
  if (!loadRecordById(recordId)) {
    setStatus("Failed to load the selected record.", "error");
    return;
  }
  setStatus("Saved record loaded successfully.", "success");
});

deleteRecordButton?.addEventListener("click", () => {
  const recordId = documentHistoryList?.value;
  if (!recordId) {
    setStatus("Please select a saved record first.", "warn");
    return;
  }
  if (!deleteRecordById(recordId)) {
    setStatus("Failed to delete the selected record.", "error");
    return;
  }
  setStatus("Saved record deleted.", "success");
});

deleteAllRecordsButton?.addEventListener("click", () => {
  deleteAllSavedRecords();
  setStatus("All saved records were deleted.", "success");
});

resetCountersButton?.addEventListener("click", () => {
  resetRunningCounters();
  setStatus("Running numbers were reset. The next new document will start again from 00001.", "success");
});

document.getElementById("refreshPoPreview").addEventListener("click", renderDocuments);
document.getElementById("refreshInvoicePreview").addEventListener("click", renderDocuments);
document.getElementById("refreshDoPreview").addEventListener("click", renderDocuments);

document.getElementById("downloadPoButton").addEventListener("click", async () => {
  try {
    renderDocuments();
    rememberBillTo(state.data.invoiceBillTo);
    await downloadPoPdf(getPoFilename());
    setStatus("PO PDF generated successfully.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Failed to generate the PO PDF.", "error");
  }
});

document.getElementById("downloadInvoiceButton").addEventListener("click", async () => {
  try {
    renderDocuments();
    rememberBillTo(state.data.invoiceBillTo);
    await downloadInvoicePdf(getInvoiceFilename());
    setStatus("Invoice PDF generated successfully.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Failed to generate the invoice PDF.", "error");
  }
});

document.getElementById("downloadDoButton").addEventListener("click", async () => {
  try {
    renderDocuments();
    rememberBillTo(state.data.invoiceBillTo);
    await downloadDeliveryOrderPdf(getDoFilename());
    setStatus("Delivery order PDF generated successfully.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Failed to generate the delivery order PDF.", "error");
  }
});

document.getElementById("downloadCompleteSetButton").addEventListener("click", async () => {
  try {
    renderDocuments();
    rememberBillTo(state.data.invoiceBillTo);
    setStatus("Generating the complete set PDF...", "progress");
    await downloadCompleteSetPdf(getCompleteSetFilename());
    setStatus("Complete set PDF was generated successfully.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Failed to generate the complete set PDF.", "error");
  }
});

fieldNames.forEach((name) => {
  const field = form.elements.namedItem(name);
  if (field) {
    field.addEventListener("input", () => {
      renderDocuments();
      if (name === "poNumber") {
        renderPoSupplierEditor();
      }
    });
  }
});

const poPreparedByField = form.elements.namedItem("poPreparedBy");
const invoicePoJhevField = form.elements.namedItem("invoicePoJhev");
const doPoJhevField = form.elements.namedItem("doPoJhev");

function syncPoJhevAcrossDocuments(value = "") {
  state.data.invoicePoJhev = value;
  state.data.doPoJhev = value;
  if (invoicePoJhevField) {
    invoicePoJhevField.value = value;
  }
  if (doPoJhevField) {
    doPoJhevField.value = value;
  }
}

poPreparedByField?.addEventListener("change", () => {
  state.data.poPreparedBy = poPreparedByField.value;
  state.data.invoicePreparedBy = poPreparedByField.value;
  syncFormFromState();
  renderDocuments();
});

invoicePoJhevField?.addEventListener("input", () => {
  syncPoJhevAcrossDocuments(invoicePoJhevField.value);
  renderDocuments();
});

invoicePoJhevField?.addEventListener("change", () => {
  syncPoJhevAcrossDocuments(invoicePoJhevField.value);
  renderDocuments();
});

doPoJhevField?.addEventListener("input", () => {
  syncPoJhevAcrossDocuments(doPoJhevField.value);
  renderDocuments();
});

doPoJhevField?.addEventListener("change", () => {
  syncPoJhevAcrossDocuments(doPoJhevField.value);
  renderDocuments();
});

billToHistoryList?.addEventListener("change", (event) => {
  const selectedValue = event.target.value;
  if (!selectedValue) return;
  state.data.invoiceBillTo = selectedValue;
  state.data.doBillTo = selectedValue;
  syncFormFromState();
  renderDocuments();
});

supplierHistoryList?.addEventListener("change", (event) => {
  const selectedValue = event.target.value;
  if (!selectedValue) return;
  state.data.poSupplier = selectedValue;
  syncFormFromState();
  renderDocuments();
});

saveBillToHistoryButton?.addEventListener("click", () => {
  const currentBillTo = form.elements.namedItem("invoiceBillTo")?.value || "";
  if (!normalizeMultiline(currentBillTo)) {
    setStatus("Please enter a Bill To address before saving.", "warn");
    return;
  }
  rememberBillTo(currentBillTo);
  billToHistoryList.value = normalizeMultiline(currentBillTo);
  setStatus("Bill To address saved to history.", "success");
});

saveSupplierHistoryButton?.addEventListener("click", () => {
  const currentSupplier = form.elements.namedItem("poSupplier")?.value || "";
  if (!normalizeMultiline(currentSupplier)) {
    setStatus("Please enter a Supplier address before saving.", "warn");
    return;
  }
  rememberSupplier(currentSupplier);
  supplierHistoryList.value = normalizeMultiline(currentSupplier);
  setStatus("Supplier saved to history.", "success");
});

editBillToHistoryButton?.addEventListener("click", () => {
  const selectedValue = billToHistoryList?.value;
  if (!selectedValue) {
    setStatus("Please select a Bill To history entry first.", "warn");
    return;
  }
  state.data.invoiceBillTo = selectedValue;
  state.data.doBillTo = selectedValue;
  syncFormFromState();
  renderDocuments();
  setStatus("Selected Bill To history loaded into the editor.", "success");
});

editSupplierHistoryButton?.addEventListener("click", () => {
  const selectedValue = supplierHistoryList?.value;
  if (!selectedValue) {
    setStatus("Please select a Supplier history entry first.", "warn");
    return;
  }
  state.data.poSupplier = selectedValue;
  syncFormFromState();
  renderDocuments();
  setStatus("Selected Supplier history loaded into the editor.", "success");
});

deleteBillToHistoryButton?.addEventListener("click", () => {
  const selectedValue = billToHistoryList?.value;
  if (!selectedValue) {
    setStatus("Please select a Bill To history entry first.", "warn");
    return;
  }
  deleteBillToHistoryEntry(selectedValue);
  if (state.data.invoiceBillTo === selectedValue) {
    state.data.invoiceBillTo = "";
    state.data.doBillTo = "";
    syncFormFromState();
    renderDocuments();
  }
  setStatus("Selected Bill To history entry deleted.", "success");
});

deleteSupplierHistoryButton?.addEventListener("click", () => {
  const selectedValue = supplierHistoryList?.value;
  if (!selectedValue) {
    setStatus("Please select a Supplier history entry first.", "warn");
    return;
  }
  deleteSupplierHistoryEntry(selectedValue);
  if (state.data.poSupplier === selectedValue) {
    state.data.poSupplier = "";
    syncFormFromState();
    renderDocuments();
  }
  setStatus("Selected Supplier history entry deleted.", "success");
});

loginForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  loginUser(loginUserIdField?.value, loginPasswordField?.value);
});

logoutButton?.addEventListener("click", () => {
  logoutCurrentUser();
});

loadBillToHistory();
loadSupplierHistory();
loadRunningCounters();
loadSavedRecords();
loadSavedPoItems();
restoreAuthSession();
renderBillToHistory();
renderSupplierHistory();
renderSavedRecords();
applyParsedData(createNewDocumentState());
syncCompanyControlsFromState();
syncCategoryControlsFromState();
setCategoryGateState();
updateAuthUi();
if (state.authUser) {
  setStatus("Please select one company first, then choose UBAT, SUSU, PERALATAN, or LAMPIN PAKAI BUANG.", "warn");
} else {
  setLoginNotice("Please log in with your assigned ID and password.", "warn");
  setStatus("System locked. Please log in to continue.", "warn");
}
