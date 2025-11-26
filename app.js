/******************************
 *   PARTIE 1 / 4 – CONFIG
 ******************************/

// === CONFIG & ÉTAT GLOBAL ===

const LINES = [
  "Râpé",
  "T2",
  "RT",
  "OMORI",
  "T1",
  "Sticks",
  "Emballage",
  "Dés",
  "Filets",
  "Prédécoupé",
];

// Stockage version
const STORAGE_KEY = "atelier_ppnc_state_v2";  
const ARCHIVES_KEY = "atelier_ppnc_archives_v2";

let archives = []; // [{ id, label, savedAt, equipe, week, quantieme, state }]

// Sous-lignes pour Râpé
const ARRET_SUBLINES = {
  "Râpé": ["R1", "R2", "R1/R2"],
};

// Machines déclarées
const ARRET_MACHINES = {
  "Râpé": [
    "Cubeuse", "Cheesix", "Liftvrac", "Associative", "Ensacheuse",
    "Encaisseuse", "Smartdate", "Bizerba", "DPM", "Scotcheuse",
    "Markem", "Ascenseur",
  ],
  "T2": [
    "Selvex", "Trieuse", "Robots", "Tiromat", "Vision",
    "Convoyeur", "DPM", "Bizerba", "Suremballage", "Markem",
    "Scotcheuse", "Balance cartons", "Formeuse caisse", "Ascenseur",
  ],
  "OMORI": [
    "BFR", "Accumulateur", "OMORI", "Videojet", "DPM",
    "Encaisseuse", "Balance cartons", "Ascenseur",
  ],
  "Emballage": [
    "Brinkman", "Encaisseuse", "Bizerba", "Palettiseur",
    "Paraffineuse", "Râpé", "Ecroûtage", "Alpma",
  ],
  "T1": [
    "Slicer", "AES", "Tiromat", "Préhenseur",
    "DPM", "Encaisseuse T1", "Encaisseuse David",
    "Balance cartons", "Ascenseur",
  ],
  "Dés": ["Cheesix", "Meca 2002", "DPM", "Bizerba", "Scotcheuse"],
  "Filets": ["Lieuse", "C-Pack", "Etiqueteuse", "Scotcheuse"],
  "Prédécoupé": ["DPM", "Selvex", "Bizerba", "Quartivac", "Scotcheuse"],
  "RT": ["Autre"],
  "Sticks": ["Autre"],
};

// === ÉTAT GLOBAL ===
let state = {
  currentSection: "atelier",
  currentLine: LINES[0],
  currentEquipe: "M",
  production: {},
  arrets: [],
  organisation: [],
  personnel: [],
  excelRecords: [],
  excelFiles: [],
  formDraft: {},
};

// Références pour page Arrêts
let arretLineEl = null;
let arretSousZoneEl = null;
let arretSousZoneRowEl = null;
let arretMachineEl = null;

// Graphiques
let atelierChart = null;
let historyChart = null;

// Base manager (classeur Excel historisé)
const MANAGER_MASTER_KEY = "manager_master_excel_b64";
const MANAGER_IMPORT_LOG_KEY = "manager_import_log_v1";
const MANAGER_PASSWORD_KEY = "manager_password_hash_v1";

const MANAGER_FIELDS = [
  { key: "type", label: "Type" },
  { key: "dateHeure", label: "Date/Heure" },
  { key: "equipe", label: "Équipe" },
  { key: "ligne", label: "Ligne" },
  { key: "sousLigne", label: "Sous-ligne" },
  { key: "machine", label: "Machine" },
  { key: "heureDebut", label: "Heure Début" },
  { key: "heureFin", label: "Heure Fin" },
  { key: "quantite", label: "Quantité" },
  { key: "arretMinutes", label: "Arrêt (min)" },
  { key: "cadence", label: "Cadence" },
  { key: "tempsRestant", label: "Temps Restant" },
  { key: "commentaire", label: "Commentaire" },
  { key: "article", label: "Article" },
  { key: "nomPersonnel", label: "Personnel" },
  { key: "motifPersonnel", label: "Motif personnel" },
  { key: "visa", label: "Visa" },
  { key: "validee", label: "Validée" },
  { key: "fileName", label: "Fichier" },
  { key: "quantieme", label: "Quantième" },
  { key: "semaine", label: "Semaine" },
  { key: "heureFichier", label: "Heure fichier" },
];

const MANAGER_COLUMNS_FOR_EXCEL = [
  { key: "type", header: "Type" },
  { key: "dateHeure", header: "Date/Heure" },
  { key: "equipe", header: "Équipe" },
  { key: "ligne", header: "Ligne" },
  { key: "sousLigne", header: "Sous-ligne" },
  { key: "machine", header: "Machine" },
  { key: "heureDebut", header: "Heure Début" },
  { key: "heureFin", header: "Heure Fin" },
  { key: "quantite", header: "Quantité" },
  { key: "arretMinutes", header: "Arrêt (min)" },
  { key: "cadence", header: "Cadence" },
  { key: "tempsRestant", header: "Temps Restant" },
  { key: "commentaire", header: "Commentaire" },
  { key: "article", header: "Article" },
  { key: "nomPersonnel", header: "Nom Personnel" },
  { key: "motifPersonnel", header: "Motif Personnel" },
  { key: "visa", header: "Visa" },
  { key: "validee", header: "Validée" },
  { key: "fileName", header: "Fichier source" },
  { key: "quantieme", header: "Quantième" },
  { key: "semaine", header: "Semaine" },
  { key: "heureFichier", header: "Heure du fichier" },
  { key: "importedAt", header: "Importé le" },
];

const managerSearchState = {
  text: "",
  equipe: "",
  ligne: "",
  fields: new Set(MANAGER_FIELDS.map(f => f.key)),
  sortField: "dateHeure",
  sortDir: "desc",
};

let managerDataset = [];
let managerImportLog = [];
let managerUnlocked = false;

/******************************
 *   PARTIE 1 / 4 – FONCTIONS DATE
 ******************************/

function getNow() {
  return new Date();
}

function getWeekNumber(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  return Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
}

function getQuantieme(d) {
  const start = new Date(d.getFullYear(), 0, 0);
  const diff = d - start;
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

function formatDateTime(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()} ${hh}h${mi}`;
}

function formatTimeRemaining(min) {
  if (!min || min <= 0) return "-";
  const h = Math.floor(min / 60);
  const m = Math.round(min % 60);
  if (h === 0) return `${m} min`;
  if (m === 0) return `${h} h`;
  return `${h} h ${m} min`;
}

function parseTimeToMinutes(t) {
  if (!t) return null;
  const [h, m] = t.split(":").map(Number);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

// Conversion quantième → Date
function quantiemeToDate(quantieme, year) {
  const date = new Date(year, 0);
  date.setDate(quantieme);
  return date;
}

// Formater date pour DDM
function formatDateDDM(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

/******************************
 *   PARTIE 1 / 4 – LOCAL STORAGE
 ******************************/

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      LINES.forEach(l => state.production[l] = []);
      return;
    }
    const parsed = JSON.parse(raw);

    const base = {
      currentSection: "atelier",
      currentLine: LINES[0],
      currentEquipe: "M",
      production: {},
      arrets: [],
      organisation: [],
      personnel: [],
      excelRecords: [],
      excelFiles: [],
      formDraft: {},
    };

    state = Object.assign(base, parsed);

    LINES.forEach(l => {
      if (!state.production[l]) state.production[l] = [];
      if (!state.formDraft[l]) state.formDraft[l] = {};
    });

    if (!state.excelRecords) state.excelRecords = [];
    if (!state.excelFiles) state.excelFiles = [];

  } catch (e) {
    console.error("loadState error", e);
    LINES.forEach(l => state.production[l] = []);
  }
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    console.error("saveState error", e);
  }
}

function loadArchives() {
  try {
    const raw = localStorage.getItem(ARCHIVES_KEY);
    archives = raw ? JSON.parse(raw) : [];
  } catch {
    archives = [];
  }
}

function saveArchives() {
  try {
    localStorage.setItem(ARCHIVES_KEY, JSON.stringify(archives));
  } catch {}
}

/******************************
 *   PARTIE 1B – IMPORT EXCEL
 ******************************/

function generateId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function parseFileNameInfo(fileName) {
  const match = fileName.match(/Q(\d{3}).*?S(\d{2}).*?(\d{2})h(\d{2})/i);
  if (!match) return { fileName, quantieme: null, semaine: null, heureFichier: null, formattedDate: null };

  const quantieme = Number(match[1]);
  const semaine = Number(match[2]);
  const heureFichier = `${match[3]}:${match[4]}`;

  let formattedDate = null;
  try {
    const year = new Date().getFullYear();
    const date = quantiemeToDate(quantieme, year);
    formattedDate = formatDateDDM(date);
  } catch {
    formattedDate = null;
  }

  return { fileName, quantieme, semaine, heureFichier, formattedDate };
}

function normalizeExcelRow(row, meta) {
  const getStr = (...keys) => {
    for (const key of keys) {
      if (row[key] !== undefined && row[key] !== null) {
        return String(row[key]).trim();
      }
    }
    return "";
  };

  const getNum = (...keys) => {
    for (const key of keys) {
      if (row[key] === undefined || row[key] === null || row[key] === "") continue;
      const num = Number(row[key]);
      if (!Number.isNaN(num)) return num;
    }
    return null;
  };

  return {
    id: generateId(),
    fileName: meta.fileName,
    importedAt: meta.importedAt,
    quantieme: meta.quantieme,
    semaine: meta.semaine,
    heureFichier: meta.heureFichier,
    fichierDate: meta.formattedDate,
    type: getStr("Type"),
    dateHeure: getStr("Date/Heure", "Date", "Heure"),
    equipe: getStr("Équipe", "Equipe"),
    ligne: getStr("Ligne"),
    sousLigne: getStr("Sous-ligne", "Sous ligne", "Sous_ligne"),
    machine: getStr("Machine"),
    heureDebut: getStr("Heure Début", "Heure debut"),
    heureFin: getStr("Heure Fin"),
    quantite: getNum("Quantité", "Quantite"),
    arretMinutes: getNum("Arrêt (min)", "Arret (min)", "Arrêt"),
    cadence: getNum("Cadence"),
    tempsRestant: getNum("Temps Restant"),
    commentaire: getStr("Commentaire"),
    article: getStr("Article"),
    nomPersonnel: getStr("Nom Personnel"),
    motifPersonnel: getStr("Motif Personnel"),
    visa: getStr("Visa"),
    validee: getStr("Validée", "Validee"),
  };
}

function setImportStatus(message, variant = "") {
  const el = document.getElementById("importStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "import-status helper-text";
  if (variant) {
    el.classList.add(variant);
  }
}

async function processExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheet];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  const meta = { ...parseFileNameInfo(file.name), fileName: file.name, importedAt: new Date().toISOString() };

  const normalizedRows = rawRows
    .filter(row => Object.values(row).some(v => v !== null && v !== ""))
    .map(row => normalizeExcelRow(row, meta));

  if (!normalizedRows.length) return { count: 0 };

  state.excelRecords.push(...normalizedRows);
  state.excelFiles.push({
    id: generateId(),
    fileName: file.name,
    importedAt: meta.importedAt,
    quantieme: meta.quantieme,
    semaine: meta.semaine,
    heureFichier: meta.heureFichier,
    fichierDate: meta.formattedDate,
    rowCount: normalizedRows.length,
  });

  saveState();
  return { count: normalizedRows.length };
}

async function importExcelFiles(files) {
  if (!files.length) {
    setImportStatus("Sélectionne au moins un fichier Excel.", "warning");
    return;
  }

  let totalRows = 0;
  const skipped = [];

  for (const file of files) {
    if (state.excelFiles.some(f => f.fileName === file.name)) {
      skipped.push(file.name);
      continue;
    }

    try {
      const { count } = await processExcelFile(file);
      totalRows += count;
    } catch (e) {
      console.error("Import excel error", e);
      setImportStatus(`Erreur sur ${file.name} : ${e.message || e}`, "error");
    }
  }

  if (totalRows > 0 || skipped.length) {
    const parts = [];
    if (totalRows > 0) parts.push(`${totalRows} ligne(s) ajoutée(s)`);
    if (skipped.length) parts.push(`fichiers ignorés (déjà importés) : ${skipped.join(", ")}`);
    const variant = totalRows > 0 ? "success" : "warning";
    setImportStatus(`Import terminé : ${parts.join(" • ")}`, variant);
  } else {
    setImportStatus("Aucune ligne importée.", "warning");
  }

  refreshImportLog();
  refreshImportSearch();
}

/********************************************
 *   PARTIE 1B – BASE MANAGER (CLASSEUR EXCEL)
 ********************************************/

let managerSignatureSet = new Set();

function setManagerStatus(message, variant = "") {
  const el = document.getElementById("managerImportStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "import-status helper-text";
  if (variant) el.classList.add(variant);
}

function setManagerSecurityStatus(message, variant = "") {
  const el = document.getElementById("managerSecurityStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "import-status helper-text";
  if (variant) el.classList.add(variant);
}

function recordToSheetRow(rec) {
  const row = {};
  MANAGER_COLUMNS_FOR_EXCEL.forEach(col => {
    row[col.header] = rec[col.key] ?? "";
  });
  return row;
}

function sheetRowToRecord(row) {
  const rec = {};
  MANAGER_COLUMNS_FOR_EXCEL.forEach(col => {
    rec[col.key] = row[col.header] ?? "";
  });
  return rec;
}

function persistManagerWorkbook() {
  try {
    const rows = managerDataset.map(recordToSheetRow);
    const ws = XLSX.utils.json_to_sheet(rows, { header: MANAGER_COLUMNS_FOR_EXCEL.map(c => c.header) });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Historique");
    const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
    localStorage.setItem(MANAGER_MASTER_KEY, b64);
  } catch (e) {
    console.error("Persist manager workbook failed", e);
  }
}

function loadManagerWorkbookFromStorage() {
  managerDataset = [];
  managerSignatureSet = new Set();
  const b64 = localStorage.getItem(MANAGER_MASTER_KEY);
  if (!b64) return;
  try {
    const wb = XLSX.read(b64, { type: "base64" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    managerDataset = rows.map(sheetRowToRecord);
    managerDataset.forEach(r => managerSignatureSet.add(getManagerSignature(r)));
  } catch (e) {
    console.error("Lecture du classeur manager impossible", e);
  }
}

function loadManagerImportLog() {
  try {
    const raw = localStorage.getItem(MANAGER_IMPORT_LOG_KEY);
    managerImportLog = raw ? JSON.parse(raw) : [];
  } catch {
    managerImportLog = [];
  }
}

function saveManagerImportLog() {
  localStorage.setItem(MANAGER_IMPORT_LOG_KEY, JSON.stringify(managerImportLog));
}

async function hashPassword(pwd) {
  const enc = new TextEncoder().encode(pwd);
  const hashBuffer = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(hashBuffer))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

function getStoredPasswordHash() {
  return localStorage.getItem(MANAGER_PASSWORD_KEY);
}

async function verifyManagerPassword(pwd) {
  const stored = getStoredPasswordHash();
  if (!stored) return true; // aucun mot de passe configuré
  const hashed = await hashPassword(pwd);
  return hashed === stored;
}

async function setManagerPassword(pwd) {
  if (!pwd || pwd.length < 4) {
    setManagerSecurityStatus("Choisis un mot de passe d'au moins 4 caractères.", "warning");
    return false;
  }
  const hashed = await hashPassword(pwd);
  localStorage.setItem(MANAGER_PASSWORD_KEY, hashed);
  setManagerSecurityStatus("Mot de passe enregistré localement.", "success");
  return true;
}

function updateManagerLockUI() {
  const content = document.getElementById("managerContent");
  const overlay = document.getElementById("managerLockedOverlay");
  const unlockForm = document.getElementById("managerUnlockForm");

  const hasPassword = Boolean(getStoredPasswordHash());
  if (unlockForm) unlockForm.style.display = hasPassword ? "block" : "none";

  if (managerUnlocked || !hasPassword) {
    overlay?.classList.add("hidden");
    if (content) content.style.display = "grid";
  } else {
    overlay?.classList.remove("hidden");
    if (content) content.style.display = "none";
  }
}

async function handleUnlock(password) {
  const ok = await verifyManagerPassword(password);
  if (!ok) {
    setManagerSecurityStatus("Mot de passe incorrect.", "error");
    managerUnlocked = false;
    updateManagerLockUI();
    return;
  }
  managerUnlocked = true;
  setManagerSecurityStatus("Accès manager déverrouillé.", "success");
  updateManagerLockUI();
  refreshManagerImports();
  refreshManagerResults();
}

function getManagerSignature(row) {
  return [
    row.fileName,
    row.dateHeure,
    row.machine,
    row.article,
    row.commentaire,
    row.quantite,
    row.heureDebut,
    row.heureFin,
  ]
    .map(v => (v === undefined || v === null ? "" : String(v).toLowerCase().trim()))
    .join("|");
}

function refreshManagerImports() {
  const tbody = document.getElementById("managerImportsTable")?.querySelector("tbody");
  if (!tbody) return;

  const imports = [...managerImportLog].sort((a, b) => new Date(b.importedAt) - new Date(a.importedAt));
  tbody.innerHTML = "";

  imports.forEach(file => {
    const tr = document.createElement("tr");
    const dateImport = file.importedAt ? formatDateTime(new Date(file.importedAt)) : "-";
    const qsh = [
      file.quantieme ? `Q${String(file.quantieme).padStart(3, "0")}` : "?",
      file.semaine ? `S${file.semaine}` : "?",
      file.heureFichier || "-",
    ]
      .filter(Boolean)
      .join(" / ");

    tr.innerHTML = `
      <td>${file.fileName}</td>
      <td>${qsh}</td>
      <td>${file.rowCount || 0}</td>
      <td>${dateImport}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderManagerFieldSelector() {
  const container = document.getElementById("managerFieldSelector");
  if (!container) return;

  container.innerHTML = "";

  MANAGER_FIELDS.forEach(field => {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = managerSearchState.fields.has(field.key);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) managerSearchState.fields.add(field.key);
      else managerSearchState.fields.delete(field.key);
      refreshManagerResults();
    });

    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(field.label));
    container.appendChild(label);
  });
}

async function populateManagerLineFilter() {
  const select = document.getElementById("managerLineFilter");
  if (!select) return;

  const previous = select.value;
  select.innerHTML = "";

  const baseOption = document.createElement("option");
  baseOption.value = "";
  baseOption.textContent = "Toutes";
  select.appendChild(baseOption);

  const uniqueLines = new Set(LINES);
  managerDataset.forEach(r => {
    if (r.ligne) uniqueLines.add(r.ligne);
  });

  Array.from(uniqueLines)
    .sort()
    .forEach(line => {
      const opt = document.createElement("option");
      opt.value = line;
      opt.textContent = line;
      select.appendChild(opt);
    });

  select.value = previous;
}

function populateManagerSortOptions() {
  const sortSelect = document.getElementById("managerSortField");
  if (!sortSelect) return;

  sortSelect.innerHTML = "";
  MANAGER_FIELDS.forEach(field => {
    const opt = document.createElement("option");
    opt.value = field.key;
    opt.textContent = field.label;
    sortSelect.appendChild(opt);
  });

  sortSelect.value = managerSearchState.sortField;
}

function refreshManagerResults() {
  const table = document.getElementById("managerResultsTable");
  if (!table) return;

  let rows = [...managerDataset];

  const query = managerSearchState.text.trim().toLowerCase();
  if (query && managerSearchState.fields.size) {
    rows = rows.filter(r => {
      return Array.from(managerSearchState.fields).some(key => {
        const val = r[key];
        return val !== undefined && val !== null && String(val).toLowerCase().includes(query);
      });
    });
  }

  if (managerSearchState.equipe) {
    rows = rows.filter(
      r => (r.equipe || "").toUpperCase() === managerSearchState.equipe.toUpperCase()
    );
  }

  if (managerSearchState.ligne) {
    rows = rows.filter(r => (r.ligne || "") === managerSearchState.ligne);
  }

  const { sortField, sortDir } = managerSearchState;
  rows.sort((a, b) => {
    const av = a[sortField] ?? "";
    const bv = b[sortField] ?? "";

    if (av === bv) return 0;
    if (av > bv) return sortDir === "asc" ? 1 : -1;
    return sortDir === "asc" ? -1 : 1;
  });

  const tbody = table.querySelector("tbody");
  tbody.innerHTML = "";

  const max = 500;
  rows.slice(0, max).forEach(row => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.dateHeure || ""}</td>
      <td>${row.equipe || ""}</td>
      <td>${row.ligne || ""}</td>
      <td>${row.sousLigne || ""}</td>
      <td>${row.machine || ""}</td>
      <td>${row.quantite ?? ""}</td>
      <td>${row.arretMinutes ?? ""}</td>
      <td>${row.cadence ?? ""}</td>
      <td>${row.commentaire || ""}</td>
      <td>${row.article || ""}</td>
      <td>${row.nomPersonnel || ""}</td>
      <td>${row.motifPersonnel || ""}</td>
      <td>${row.fileName || ""}</td>
    `;
    tbody.appendChild(tr);
  });

  const infoEl = document.getElementById("managerResultsCount");
  if (infoEl) {
    infoEl.textContent = `${rows.length} résultat(s) / ${managerDataset.length} dans le classeur (affichage limité à ${max}).`;
  }
}

async function importManagerFiles(files) {
  if (!managerUnlocked && getStoredPasswordHash()) {
    setManagerStatus("Déverrouille l'accès manager avant d'importer.", "error");
    return;
  }
  if (!files.length) {
    setManagerStatus("Sélectionne au moins un fichier Excel.", "warning");
    return;
  }

  let totalRows = 0;
  const skipped = [];

  for (const file of files) {
    const existing = managerImportLog.find(log => log.fileName === file.name);
    if (existing) {
      skipped.push(file.name);
      continue;
    }

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheet = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheet];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const meta = {
        ...parseFileNameInfo(file.name),
        fileName: file.name,
        importedAt: new Date().toISOString(),
      };

      const normalizedRows = rawRows
        .filter(row => Object.values(row).some(v => v !== null && v !== ""))
        .map(row => normalizeExcelRow(row, meta));

      const freshRows = normalizedRows.filter(r => {
        const sig = getManagerSignature(r);
        if (managerSignatureSet.has(sig)) return false;
        managerSignatureSet.add(sig);
        return true;
      });

      if (!freshRows.length) continue;

      managerDataset.push(...freshRows);
      managerImportLog.push({
        fileName: file.name,
        importedAt: meta.importedAt,
        quantieme: meta.quantieme,
        semaine: meta.semaine,
        heureFichier: meta.heureFichier,
        rowCount: freshRows.length,
      });

      totalRows += freshRows.length;
    } catch (e) {
      console.error("Manager import error", e);
      setManagerStatus(`Erreur sur ${file.name} : ${e.message || e}`, "error");
    }
  }

  persistManagerWorkbook();
  saveManagerImportLog();

  const parts = [];
  if (totalRows > 0) parts.push(`${totalRows} ligne(s) ajoutée(s)`);
  if (skipped.length) parts.push(`fichiers ignorés : ${skipped.join(", ")}`);

  const variant = totalRows > 0 ? "success" : skipped.length ? "warning" : "";
  setManagerStatus(parts.length ? parts.join(" • ") : "Aucune ligne ajoutée.", variant);

  refreshManagerImports();
  refreshManagerResults();
  populateManagerLineFilter();
}

function resetManagerStore() {
  managerDataset = [];
  managerImportLog = [];
  managerSignatureSet = new Set();
  localStorage.removeItem(MANAGER_MASTER_KEY);
  localStorage.removeItem(MANAGER_IMPORT_LOG_KEY);
  setManagerStatus("Base vidée.", "warning");
  refreshManagerImports();
  refreshManagerResults();
  populateManagerLineFilter();
}

function downloadManagerWorkbook() {
  try {
    const rows = managerDataset.map(recordToSheetRow);
    const ws = XLSX.utils.json_to_sheet(rows, { header: MANAGER_COLUMNS_FOR_EXCEL.map(c => c.header) });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Historique");
    const buffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "manager_historique.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setManagerStatus("Classeur exporté.", "success");
  } catch (e) {
    console.error("Export manager impossible", e);
    setManagerStatus("Impossible de générer le classeur.", "error");
  }
}

function bindManagerArea() {
  loadManagerWorkbookFromStorage();
  loadManagerImportLog();

  renderManagerFieldSelector();
  populateManagerSortOptions();
  populateManagerLineFilter();
  updateManagerLockUI();

  const input = document.getElementById("managerExcelInput");
  const btn = document.getElementById("managerImportBtn");
  const resetBtn = document.getElementById("managerResetBtn");
  const exportBtn = document.getElementById("managerExportBtn");

  if (btn && input) {
    btn.addEventListener("click", async () => {
      const files = Array.from(input.files || []);
      setManagerStatus("Import en cours...");
      await importManagerFiles(files);
      input.value = "";
    });
  }

  if (resetBtn) {
    resetBtn.addEventListener("click", async () => {
      const confirmReset = confirm("Vider la base manager ? (sans toucher aux autres formulaires)");
      if (confirmReset) resetManagerStore();
    });
  }

  if (exportBtn) {
    exportBtn.addEventListener("click", downloadManagerWorkbook);
  }

  const searchInput = document.getElementById("managerSearchInput");
  if (searchInput) {
    searchInput.addEventListener("input", () => {
      managerSearchState.text = searchInput.value;
      refreshManagerResults();
    });
  }

  const equipeFilter = document.getElementById("managerEquipeFilter");
  if (equipeFilter) {
    equipeFilter.addEventListener("change", () => {
      managerSearchState.equipe = equipeFilter.value;
      refreshManagerResults();
    });
  }

  const lineFilter = document.getElementById("managerLineFilter");
  if (lineFilter) {
    lineFilter.addEventListener("change", () => {
      managerSearchState.ligne = lineFilter.value;
      refreshManagerResults();
    });
  }

  const sortSelect = document.getElementById("managerSortField");
  if (sortSelect) {
    sortSelect.addEventListener("change", () => {
      managerSearchState.sortField = sortSelect.value;
      refreshManagerResults();
    });
  }

  const sortDir = document.getElementById("managerSortDir");
  if (sortDir) {
    sortDir.addEventListener("change", () => {
      managerSearchState.sortDir = sortDir.value;
      refreshManagerResults();
    });
  }

  const unlockForm = document.getElementById("managerUnlockForm");
  const setPwdForm = document.getElementById("managerSetPasswordForm");
  const unlockInput = document.getElementById("managerUnlockInput");
  const pwdInput = document.getElementById("managerPasswordInput");

  if (unlockForm) {
    unlockForm.addEventListener("submit", async e => {
      e.preventDefault();
      const pwd = unlockInput?.value || "";
      await handleUnlock(pwd);
      if (unlockInput) unlockInput.value = "";
    });
  }

  if (setPwdForm) {
    setPwdForm.addEventListener("submit", async e => {
      e.preventDefault();
      const pwd = pwdInput?.value || "";
      const ok = await setManagerPassword(pwd);
      if (ok && !managerUnlocked) await handleUnlock(pwd);
      if (pwdInput) pwdInput.value = "";
      updateManagerLockUI();
    });
  }

  refreshManagerImports();
  refreshManagerResults();
}

function refreshImportLog() {
  const tbody = document.getElementById("importsTableBody");
  if (!tbody) return;

  tbody.innerHTML = "";

  [...(state.excelFiles || [])]
    .sort((a, b) => new Date(b.importedAt) - new Date(a.importedAt))
    .forEach(file => {
      const tr = document.createElement("tr");
      const dateImport = file.importedAt ? formatDateTime(new Date(file.importedAt)) : "-";
      const qsh = [file.quantieme ? `Q${String(file.quantieme).padStart(3, "0")}` : "?", file.semaine ? `S${file.semaine}` : "?", file.heureFichier || "-"]
        .filter(Boolean)
        .join(" / ");

      tr.innerHTML = `
        <td>${file.fileName}</td>
        <td>${qsh}</td>
        <td>${file.rowCount || 0}</td>
        <td>${dateImport}</td>
      `;
      tbody.appendChild(tr);
    });

  const countEl = document.getElementById("importResultsCount");
  if (countEl) {
    const total = state.excelRecords ? state.excelRecords.length : 0;
    countEl.textContent = `${total} enregistrement(s) importés.`;
  }
}

function populateImportFilters() {
  const lineSelect = document.getElementById("importLineFilter");
  if (!lineSelect || lineSelect.dataset.populated === "1") return;
  lineSelect.dataset.populated = "1";

  LINES.forEach(line => {
    const opt = document.createElement("option");
    opt.value = line;
    opt.textContent = line;
    lineSelect.appendChild(opt);
  });
}

function refreshImportSearch() {
  const table = document.getElementById("importResultsTable");
  if (!table) return;

  const query = (document.getElementById("importSearchInput")?.value || "").trim().toLowerCase();
  const equipe = (document.getElementById("importEquipeFilter")?.value || "").toUpperCase();
  const ligne = document.getElementById("importLineFilter")?.value || "";

  let results = state.excelRecords || [];

  if (equipe) {
    results = results.filter(r => (r.equipe || "").toUpperCase() === equipe);
  }

  if (ligne) {
    results = results.filter(r => (r.ligne || "") === ligne);
  }

  if (query) {
    results = results.filter(r => {
      return [r.article, r.machine, r.commentaire, r.type, r.dateHeure, r.nomPersonnel]
        .filter(Boolean)
        .some(val => String(val).toLowerCase().includes(query));
    });
  }

  const tbody = table.querySelector("tbody");
  tbody.innerHTML = "";

  results.slice(0, 200).forEach(rec => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${rec.dateHeure || ""}</td>
      <td>${rec.equipe || ""}</td>
      <td>${rec.ligne || ""}</td>
      <td>${rec.machine || ""}</td>
      <td>${rec.quantite ?? ""}</td>
      <td>${rec.arretMinutes ?? ""}</td>
      <td>${rec.cadence ?? ""}</td>
      <td>${rec.commentaire || ""}</td>
      <td>${rec.fileName}</td>
    `;
    tbody.appendChild(tr);
  });

  const infoEl = document.getElementById("importResultsCount");
  if (infoEl) {
    const total = state.excelRecords ? state.excelRecords.length : 0;
    infoEl.textContent = `${results.length} résultat(s) filtrés / ${total} importés (max 200 affichés).`;
  }
}

function bindExcelImport() {
  const btn = document.getElementById("importExcelBtn");
  const input = document.getElementById("excelInput");
  if (!btn || !input) return;

  populateImportFilters();

  btn.addEventListener("click", async () => {
    const files = Array.from(input.files || []);
    setImportStatus("Import en cours...", "");
    await importExcelFiles(files);
    input.value = "";
  });

  const searchInput = document.getElementById("importSearchInput");
  const equipeFilter = document.getElementById("importEquipeFilter");
  const lineFilter = document.getElementById("importLineFilter");

  if (searchInput) searchInput.addEventListener("input", refreshImportSearch);
  if (equipeFilter) equipeFilter.addEventListener("change", refreshImportSearch);
  if (lineFilter) lineFilter.addEventListener("change", refreshImportSearch);

  refreshImportLog();
  refreshImportSearch();
}

/******************************
 *   PARTIE 1 / 4 – HEADER
 ******************************/

function initHeaderDate() {
  const el = document.getElementById("header-datetime");
  if (!el) return;

  function update() {
    const now = getNow();
    const week = getWeekNumber(now);
    const q = getQuantieme(now);

    const d = String(now.getDate()).padStart(2, "0");
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");

    el.textContent =
      `Quantième ${q} • ${d}/${m}/${now.getFullYear()} • ` +
      `S${week} • ${hh}:${mm}`;
  }

  update();
  setInterval(update, 30000);
}

function initEquipeSelector() {
  const selector = document.getElementById("equipeSelector");
  if (!selector) return;

  // Initialiser avec l'équipe actuelle
  selector.value = state.currentEquipe;

  // Écouter les changements
  selector.addEventListener("change", () => {
    state.currentEquipe = selector.value;
    saveState();
    
    // Confirmation visuelle
    const btn = document.createElement("div");
    btn.textContent = `✓ Équipe ${selector.value} activée`;
    btn.style.cssText = `
      position: fixed;
      top: 80px;
      right: 20px;
      background: #62d38b;
      color: white;
      padding: 10px 20px;
      border-radius: 8px;
      z-index: 9999;
      animation: fadeIn 0.3s ease-out;
    `;
    document.body.appendChild(btn);
    
    setTimeout(() => {
      btn.style.opacity = "0";
      btn.style.transition = "opacity 0.3s";
      setTimeout(() => btn.remove(), 300);
    }, 2000);
  });
}

/********************************************
 *   PARTIE 2 / 4 – NAVIGATION + PRODUCTION
 ********************************************/

// === NAVIGATION ===

function showSection(section) {
  state.currentSection = section;
  saveState();

  document.querySelectorAll(".section").forEach(sec =>
    sec.classList.toggle("visible", sec.id === `section-${section}`)
  );

  document.querySelectorAll(".nav-btn").forEach(btn =>
    btn.classList.toggle("active", btn.dataset.section === section)
  );

  if (section === "atelier") refreshAtelierView();
  else if (section === "production") refreshProductionView();
  else if (section === "arrets") refreshArretsView();
  else if (section === "organisation") refreshOrganisationView();
  else if (section === "personnel") refreshPersonnelView();
  else if (section === "manager") {
    populateManagerLineFilter();
    refreshManagerImports();
    refreshManagerResults();
  }
}

function initNav() {
  document.querySelectorAll(".nav-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      showSection(btn.dataset.section);
    });
  });
}

// === LIGNES (Production) ===

function initLinesSidebar() {
  const container = document.getElementById("linesList");
  if (!container) return;
  
  container.innerHTML = "";

  LINES.forEach(line => {
    const btn = document.createElement("button");
    btn.className = "line-btn";
    btn.textContent = line;
    btn.dataset.line = line;

    btn.addEventListener("click", () => selectLine(line, true));

    container.appendChild(btn);
  });

  selectLine(state.currentLine || LINES[0], false);
}

function selectLine(line, scrollToForm) {
  state.currentLine = line;
  saveState();

  const titleEl = document.getElementById("currentLineTitle");
  if (titleEl) titleEl.textContent = `Ligne ${line}`;

  document.querySelectorAll(".line-btn").forEach(b =>
    b.classList.toggle("active", b.dataset.line === line)
  );

  refreshProductionForm();
  refreshProductionHistoryTable();

  if (scrollToForm) {
    const card = document.querySelector("#section-production .card");
    if (card) card.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

// === PRODUCTION : OUTILS ===

function getCurrentLineRecords() {
  return state.production[state.currentLine] || [];
}

function computeCadenceFromInputs() {
  const s = document.getElementById("prodStartTime")?.value;
  const e = document.getElementById("prodEndTime")?.value;
  const q = Number(document.getElementById("prodQuantity")?.value) || 0;

  const sMin = parseTimeToMinutes(s);
  const eMin = parseTimeToMinutes(e);
  if (sMin == null || eMin == null || q <= 0) return null;

  let dur = eMin - sMin;
  if (dur <= 0) dur += 24 * 60;

  const hours = dur / 60;
  return hours > 0 ? q / hours : null;
}

function computeRefCadenceForLine(line) {
  const manualEl = document.getElementById("prodCadenceManual");
  const mana = Number(manualEl?.value);
  if (mana > 0) return mana;

  const recs = state.production[line] || [];
  const withCad = recs.filter(r => r.cadence && r.cadence > 0);
  if (!withCad.length) return null;

  return withCad[withCad.length - 1].cadence;
}

// === MISE À JOUR AFFICHAGE CADENCE + TEMPS RESTANT ===

function updateCadenceDisplay() {
  const cad = computeCadenceFromInputs();
  const el = document.getElementById("prodCadenceDisplay");
  if (el) el.textContent = cad ? cad.toFixed(2) : "-";
}

function updateRemainingTimeDisplay() {
  const remainingEl = document.getElementById("prodRemaining");
  const remaining = Number(remainingEl?.value) || 0;
  const line = state.currentLine;
  const cadRef = computeRefCadenceForLine(line);
  const el = document.getElementById("prodRemainingTimeDisplay");

  if (!el) return;

  if (!remaining || !cadRef || cadRef <= 0) {
    el.textContent = "-";
    return;
  }

  const min = (remaining / cadRef) * 60;
  el.textContent = formatTimeRemaining(min);
}

// === PERSISTENCE AUTO DU FORMULAIRE ===

function saveDraft() {
  const L = state.currentLine;

  state.formDraft[L] = {
    start: document.getElementById("prodStartTime")?.value || "",
    end: document.getElementById("prodEndTime")?.value || "",
    qty: document.getElementById("prodQuantity")?.value || "",
    remaining: document.getElementById("prodRemaining")?.value || "",
    cadenceMan: document.getElementById("prodCadenceManual")?.value || "",
    arret: document.getElementById("prodArretMinutes")?.value || "",
    comment: document.getElementById("prodComment")?.value || "",
    article: document.getElementById("prodCodeArticle")?.value || "",
  };

  saveState();
}

function loadDraft() {
  const L = state.currentLine;
  const d = state.formDraft[L] || {};

  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val || "";
  };

  setVal("prodStartTime", d.start);
  setVal("prodEndTime", d.end);
  setVal("prodQuantity", d.qty);
  setVal("prodRemaining", d.remaining);
  setVal("prodCadenceManual", d.cadenceMan);
  setVal("prodArretMinutes", d.arret);
  setVal("prodComment", d.comment);
  setVal("prodCodeArticle", d.article);

  updateCadenceDisplay();
  updateRemainingTimeDisplay();
}

// === REFRESH FORMULAIRE PRODUCTION ===

function refreshProductionForm() {
  loadDraft();
}

// === TABLE HISTORIQUE DES PRODUCTIONS ===

function refreshProductionHistoryTable() {
  const table = document.getElementById("prodHistoryTable");
  if (!table) return;
  
  const tbody = table.querySelector("tbody");
  if (!tbody) return;
  
  tbody.innerHTML = "";

  getCurrentLineRecords().forEach((rec, idx) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${rec.dateTime}</td>
      <td>${rec.equipe}</td>
      <td>${rec.start || "-"}</td>
      <td>${rec.end || "-"}</td>
      <td>${rec.quantity}</td>
      <td>${rec.arret || 0}</td>
      <td>${rec.cadence ? rec.cadence.toFixed(2) : "-"}</td>
      <td>${rec.remainingTime || "-"}</td>
      <td>${rec.comment || ""}</td>
      <td>${rec.article || ""}</td>
      <td><button class="secondary-btn" data-idx="${idx}">✕</button></td>
    `;

    tbody.appendChild(tr);
  });

  tbody.querySelectorAll("button[data-idx]").forEach(btn => {
    btn.addEventListener("click", () => {
      const arr = getCurrentLineRecords();
      const i = Number(btn.dataset.idx);
      arr.splice(i, 1);
      saveState();
      refreshProductionHistoryTable();
      refreshAtelierView();
    });
  });
}

// === FORM BINDING ===

function bindProductionForm() {
  const formIds = [
    "prodStartTime",
    "prodEndTime",
    "prodQuantity",
    "prodRemaining",
    "prodCadenceManual",
    "prodArretMinutes",
    "prodComment",
    "prodCodeArticle"
  ];

  formIds.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    
    el.addEventListener("input", () => {
      if (id !== "prodComment") updateCadenceDisplay();
      if (id !== "prodComment") updateRemainingTimeDisplay();
      saveDraft();
    });
  });

  const saveBtn = document.getElementById("prodSaveBtn");
  if (saveBtn) {
    saveBtn.addEventListener("click", () => {
      const now = getNow();
      const L = state.currentLine;
      const equipe = state.currentEquipe;

      const rec = {
        dateTime: formatDateTime(now),
        equipe,
        start: document.getElementById("prodStartTime")?.value || "",
        end: document.getElementById("prodEndTime")?.value || "",
        quantity: Number(document.getElementById("prodQuantity")?.value) || 0,
        remainingTime:
          document.getElementById("prodRemainingTimeDisplay")?.textContent || "-",
        arret: Number(document.getElementById("prodArretMinutes")?.value) || 0,
        cadence:
          Number(document.getElementById("prodCadenceManual")?.value) ||
          computeCadenceFromInputs() ||
          null,
        comment: document.getElementById("prodComment")?.value || "",
        article: document.getElementById("prodCodeArticle")?.value || "",
      };

      state.production[L].push(rec);

      // effacer brouillon
      state.formDraft[L] = {};
      saveState();

      refreshProductionForm();
      refreshProductionHistoryTable();
      refreshAtelierView();

      if (rec.arret > 0) {
        openArretFromProduction(L, rec.arret);
      }
    });
  }

  const undoBtn = document.getElementById("prodUndoBtn");
  if (undoBtn) {
    undoBtn.addEventListener("click", () => {
      const arr = getCurrentLineRecords();
      if (arr.length) arr.pop();
      saveState();
      refreshProductionHistoryTable();
      refreshAtelierView();
    });
  }
}

// === SCROLL HORIZONTAL PRODUCTION ===

function refreshProductionView() {
  refreshProductionForm();
  refreshProductionHistoryTable();

  const table = document.getElementById("prodHistoryTable");
  if (table && table.parentElement) {
    table.parentElement.style.overflowX = "auto";
  }
}

/********************************************
 *   PARTIE 3 / 4 – ARRETS + ORGANISATION + PERSONNEL
 ********************************************/

// === ARRETS – Mise à jour des menus déroulants ===

function updateArretControlsForLine() {
  if (!arretLineEl || !arretMachineEl) return;

  const line = arretLineEl.value;

  // Sous-zones pour Râpé
  if (arretSousZoneRowEl && arretSousZoneEl) {
    if (line === "Râpé") {
      arretSousZoneRowEl.style.display = "";
      arretSousZoneEl.innerHTML = "";
      ARRET_SUBLINES["Râpé"].forEach(s => {
        const opt = document.createElement("option");
        opt.value = s;
        opt.textContent = s;
        arretSousZoneEl.appendChild(opt);
      });
    } else {
      arretSousZoneRowEl.style.display = "none";
      arretSousZoneEl.innerHTML = "";
    }
  }

  // Machines
  arretMachineEl.innerHTML = "";
  (ARRET_MACHINES[line] || ["Autre"]).forEach(m => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    arretMachineEl.appendChild(opt);
  });
}

// === ARRETS INIT ===

function initArretsForm() {
  arretLineEl = document.getElementById("arretLine");
  arretSousZoneEl = document.getElementById("arretSousZone");
  arretMachineEl = document.getElementById("arretMachine");
  
  if (!arretLineEl || !arretSousZoneEl || !arretMachineEl) return;
  
  arretSousZoneRowEl = arretSousZoneEl.closest(".form-row");

  // Lignes dans la liste
  arretLineEl.innerHTML = "";
  LINES.forEach(l => {
    const opt = document.createElement("option");
    opt.value = l;
    opt.textContent = l;
    arretLineEl.appendChild(opt);
  });

  arretLineEl.addEventListener("change", updateArretControlsForLine);
  updateArretControlsForLine();

  // Enregistrer un arrêt
  const saveBtn = document.getElementById("arretSaveBtn");
  if (saveBtn) {
    saveBtn.addEventListener("click", () => {
      const now = getNow();

      const rec = {
        dateTime: formatDateTime(now),
        line: arretLineEl.value,
        sousLigne:
          arretSousZoneRowEl && arretSousZoneRowEl.style.display !== "none"
            ? arretSousZoneEl.value
            : "",
        machine: arretMachineEl.value,
        duration: Number(document.getElementById("arretDuration")?.value) || 0,
        comment: document.getElementById("arretComment")?.value || "",
        article: document.getElementById("arretArticle")?.value || "",
      };

      state.arrets.push(rec);
      saveState();

      const durationEl = document.getElementById("arretDuration");
      const commentEl = document.getElementById("arretComment");
      const articleEl = document.getElementById("arretArticle");
      
      if (durationEl) durationEl.value = "";
      if (commentEl) commentEl.value = "";
      if (articleEl) articleEl.value = "";

      refreshArretsView();
      refreshAtelierView();
    });
  }
}

// === OUVERTURE PAGE ARRÊT depuis Production ===

function openArretFromProduction(line, duration) {
  if (arretLineEl) {
    arretLineEl.value = line;
    updateArretControlsForLine();
  }
  
  const durationEl = document.getElementById("arretDuration");
  if (durationEl) durationEl.value = duration || "";
  
  showSection("arrets");
}

// === TABLEAU HISTORIQUE DES ARRÊTS ===

function refreshArretsView() {
  const table = document.getElementById("arretsHistoryTable");
  if (!table) return;
  
  const tbody = table.querySelector("tbody");
  if (!tbody) return;

  tbody.innerHTML = "";

  state.arrets.forEach((rec, idx) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${rec.dateTime}</td>
      <td>${rec.line}</td>
      <td>${rec.sousLigne || "-"}</td>
      <td>${rec.machine}</td>
      <td>${rec.article || ""}</td>
      <td>${rec.duration}</td>
      <td>${rec.comment || ""}</td>
      <td><button class="secondary-btn" data-idx="${idx}">✕</button></td>
    `;

    tbody.appendChild(tr);
  });

  // Suppressions
  tbody.querySelectorAll("button[data-idx]").forEach(btn => {
    btn.addEventListener("click", () => {
      const i = Number(btn.dataset.idx);
      state.arrets.splice(i, 1);
      saveState();
      refreshArretsView();
      refreshAtelierView();
    });
  });
}

// === ORGANISATION ===

function bindOrganisationForm() {
  const saveBtn = document.getElementById("orgSaveBtn");
  if (!saveBtn) return;

  saveBtn.addEventListener("click", () => {
    const now = getNow();

    const rec = {
      dateTime: formatDateTime(now),
      equipe: state.currentEquipe,
      consigne: document.getElementById("orgConsigne")?.value || "",
      visa: document.getElementById("orgVisa")?.value || "",
      valide: false,
    };

    state.organisation.push(rec);
    saveState();

    const consigneEl = document.getElementById("orgConsigne");
    const visaEl = document.getElementById("orgVisa");
    
    if (consigneEl) consigneEl.value = "";
    if (visaEl) visaEl.value = "";

    refreshOrganisationView();
  });
}

function refreshOrganisationView() {
  const table = document.getElementById("orgHistoryTable");
  if (!table) return;
  
  const tbody = table.querySelector("tbody");
  if (!tbody) return;
  
  tbody.innerHTML = "";

  state.organisation.forEach((rec, idx) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${rec.dateTime}</td>
      <td>${rec.equipe}</td>
      <td>${rec.consigne}</td>
      <td>${rec.visa}</td>
      <td>
        <button class="secondary-btn" data-idx="${idx}">
          ${rec.valide ? "✅" : "❌"}
        </button>
      </td>
    `;

    tbody.appendChild(tr);
  });

  tbody.querySelectorAll("button[data-idx]").forEach(btn => {
    btn.addEventListener("click", () => {
      const i = Number(btn.dataset.idx);
      state.organisation[i].valide = !state.organisation[i].valide;
      saveState();
      refreshOrganisationView();
    });
  });
}

// === PERSONNEL ===

function bindPersonnelForm() {
  const saveBtn = document.getElementById("persSaveBtn");
  if (!saveBtn) return;

  saveBtn.addEventListener("click", () => {
    const now = getNow();

    const rec = {
      dateTime: formatDateTime(now),
      equipe: state.currentEquipe,
      nom: document.getElementById("persNom")?.value || "",
      motif: document.getElementById("persMotif")?.value || "",
      comment: document.getElementById("persComment")?.value || "",
    };

    state.personnel.push(rec);
    saveState();

    const nomEl = document.getElementById("persNom");
    const motifEl = document.getElementById("persMotif");
    const commentEl = document.getElementById("persComment");
    
    if (nomEl) nomEl.value = "";
    if (motifEl) motifEl.value = "";
    if (commentEl) commentEl.value = "";

    refreshPersonnelView();
  });
}

function refreshPersonnelView() {
  const table = document.getElementById("persHistoryTable");
  if (!table) return;
  
  const tbody = table.querySelector("tbody");
  if (!tbody) return;
  
  tbody.innerHTML = "";

  state.personnel.forEach(rec => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${rec.dateTime}</td>
      <td>${rec.equipe}</td>
      <td>${rec.nom}</td>
      <td>${rec.motif}</td>
      <td>${rec.comment}</td>
    `;

    tbody.appendChild(tr);
  });
}

/********************************************
 *   PARTIE 3 / 4 – GRAPHIQUES ATELIER
 ********************************************/

function refreshAtelierChart() {
  const canvas = document.getElementById("atelierChart");
  if (!canvas || typeof Chart === 'undefined') return;

  // Détruire l'ancien graphique
  if (atelierChart) {
    atelierChart.destroy();
    atelierChart = null;
  }

  // Préparer les données : dernières cadences par ligne
  // ✅ FILTRER uniquement les lignes avec des données
  const datasets = LINES.map(line => {
    const recs = state.production[line] || [];
    const cadences = recs
      .filter(r => r.cadence && r.cadence > 0)
      .map(r => r.cadence);

    return {
      label: line,
      data: cadences.slice(-10), // 10 dernières valeurs
      borderWidth: 2,
      fill: false,
      tension: 0.1,
      hasData: cadences.length > 0  // ✅ Marqueur de données
    };
  }).filter(dataset => dataset.hasData);  // ✅ Ne garder que les lignes avec données

  // Créer le graphique
  atelierChart = new Chart(canvas, {
    type: 'line',
    data: {
      labels: Array.from({ length: 10 }, (_, i) => `#${i + 1}`),
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      scales: {
        y: {
          beginAtZero: true,
          title: { display: true, text: 'Cadence (colis/h)' }
        }
      },
      plugins: {
        legend: { 
          display: datasets.length > 0,  // ✅ Masquer si aucune donnée
          position: 'bottom' 
        }
      }
    }
  });
}

function refreshAtelierView() {
  const container = document.getElementById("atelier-lines-summary");
  if (!container) return;

  container.innerHTML = "";

  refreshImportLog();
  refreshImportSearch();

  // cartes lignes
  LINES.forEach(line => {
    const recs = state.production[line] || [];
    const total = recs.reduce((s, r) => s + (r.quantity || 0), 0);
    const cad = recs
      .map(r => r.cadence)
      .filter(c => c && c > 0);
    const avg = cad.length
      ? cad.reduce((s, c) => s + c) / cad.length
      : 0;

    const arts = new Set(recs.map(r => r.article).filter(a => a));
    const artStr = arts.size ? [...arts].join(", ") : "-";

    const div = document.createElement("div");
    div.className = "summary-card";

    div.innerHTML = `
      <div class="summary-card-title">${line}</div>
      <div class="summary-main">${total} colis</div>
      <div class="summary-sub">Articles : ${artStr}</div>
      <div class="summary-sub">Cadence moy. ${avg ? avg.toFixed(1) : "-"} h</div>
    `;

    container.appendChild(div);
  });

  // Arrêts majeurs
  const table = document.getElementById("atelier-arrets-table");
  if (!table) return;
  
  
  const tbody = table.querySelector("tbody");
  if (!tbody) return;
  
  tbody.innerHTML = "";

  [...state.arrets]
    .sort((a, b) => (b.duration || 0) - (a.duration || 0))
    .forEach(rec => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${rec.line}</td>
        <td>${rec.sousLigne || "-"}</td>
        <td>${rec.machine}</td>
        <td>${rec.duration}</td>
        <td>${rec.comment || ""}</td>
      `;
      tbody.appendChild(tr);
    });

  // scroll horizontal ATELIER
  if (table.parentElement) {
    table.parentElement.style.overflowX = "auto";
  }

  // Rafraîchir le graphique
  refreshAtelierChart();
}

/********************************************
 *   PARTIE 3 / 4 – HISTORIQUE ÉQUIPES
 ********************************************/

function initHistoriqueEquipes() {
  const select = document.getElementById("historySelect");
  if (!select) return;
  
  refreshHistorySelect();

  select.addEventListener("change", () => {
    if (select.value === "") {
      clearHistoryView();
    } else {
      const idx = Number(select.value);
      if (archives[idx]) {
        refreshHistoryView(archives[idx]);
      }
    }
  });
}

function refreshHistorySelect() {
  const select = document.getElementById("historySelect");
  if (!select) return;
  
  select.innerHTML = `<option value="">-- Sélectionner --</option>`;

  archives.forEach((snap, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    opt.textContent = snap.label;
    select.appendChild(opt);
  });
}

function clearHistoryView() {
  const summaryEl = document.getElementById("history-lines-summary");
  if (summaryEl) summaryEl.innerHTML = "";

  const table = document.getElementById("history-arrets-table");
  if (table) {
    const tbody = table.querySelector("tbody");
    if (tbody) tbody.innerHTML = "";
  }

  if (historyChart) {
    historyChart.destroy();
    historyChart = null;
  }
}

function refreshHistoryView(snapshot) {
  if (!snapshot || !snapshot.state) return;

  const savedState = snapshot.state;

  // Afficher les lignes de production
  const summaryEl = document.getElementById("history-lines-summary");
  if (summaryEl) {
    summaryEl.innerHTML = "";

    LINES.forEach(line => {
      const recs = savedState.production[line] || [];
      const total = recs.reduce((s, r) => s + (r.quantity || 0), 0);
      const cad = recs.map(r => r.cadence).filter(c => c && c > 0);
      const avg = cad.length ? cad.reduce((s, c) => s + c) / cad.length : 0;

      const div = document.createElement("div");
      div.className = "summary-card";
      div.innerHTML = `
        <div class="summary-card-title">${line}</div>
        <div class="summary-main">${total} colis</div>
        <div class="summary-sub">Cadence moy. ${avg ? avg.toFixed(1) : "-"} h</div>
      `;
      summaryEl.appendChild(div);
    });
  }

  // Afficher les arrêts
  const table = document.getElementById("history-arrets-table");
  if (table) {
    const tbody = table.querySelector("tbody");
    if (tbody) {
      tbody.innerHTML = "";

      (savedState.arrets || [])
        .sort((a, b) => (b.duration || 0) - (a.duration || 0))
        .forEach(rec => {
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td>${rec.line}</td>
            <td>${rec.sousLigne || "-"}</td>
            <td>${rec.machine}</td>
            <td>${rec.duration}</td>
            <td>${rec.comment || ""}</td>
          `;
          tbody.appendChild(tr);
        });
    }
  }

  // Graphique Chart.js
  const chartCanvas = document.getElementById("historyChart");
  if (chartCanvas && typeof Chart !== 'undefined') {
    if (historyChart) historyChart.destroy();

    const labels = LINES;
    const data = LINES.map(line => {
      const recs = savedState.production[line] || [];
      return recs.reduce((s, r) => s + (r.quantity || 0), 0);
    });

    historyChart = new Chart(chartCanvas, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: 'Production (colis)',
          data: data,
          backgroundColor: 'rgba(54, 162, 235, 0.6)',
          borderColor: 'rgba(54, 162, 235, 1)',
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        scales: {
          y: { beginAtZero: true }
        }
      }
    });
  }

  // ✅ AFFICHER ORGANISATION
  const orgContainer = document.getElementById("history-organisation");
  if (orgContainer && savedState.organisation && savedState.organisation.length > 0) {
    orgContainer.style.display = "block";
    const orgTable = orgContainer.querySelector("tbody");
    if (orgTable) {
      orgTable.innerHTML = "";
      savedState.organisation.forEach(r => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${r.dateTime}</td>
          <td>${r.equipe}</td>
          <td>${r.consigne}</td>
          <td>${r.visa}</td>
          <td>${r.valide ? "✅" : "❌"}</td>
        `;
        orgTable.appendChild(tr);
      });
    }
  } else if (orgContainer) {
    orgContainer.style.display = "none";
  }

  // ✅ AFFICHER PERSONNEL
  const persContainer = document.getElementById("history-personnel");
  if (persContainer && savedState.personnel && savedState.personnel.length > 0) {
    persContainer.style.display = "block";
    const persTable = persContainer.querySelector("tbody");
    if (persTable) {
      persTable.innerHTML = "";
      savedState.personnel.forEach(r => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${r.dateTime}</td>
          <td>${r.equipe}</td>
          <td>${r.nom}</td>
          <td>${r.motif}</td>
          <td>${r.comment}</td>
        `;
        persTable.appendChild(tr);
      });
    }
  } else if (persContainer) {
    persContainer.style.display = "none";
  }
}

/********************************************
 *   PARTIE 4 / 4 – EXPORT GLOBAL
 ********************************************/

// ===== EXPORT 1 : DATA (Base de données) =====
function exportDataToExcel(srcState, filename) {
  if (typeof XLSX === 'undefined') {
    alert("Bibliothèque XLSX non chargée.");
    return;
  }

  const wb = XLSX.utils.book_new();

  const rows = [[
    "Type",
    "Date/Heure",
    "Équipe",
    "Ligne",
    "Sous-ligne",
    "Machine",
    "Heure Début",
    "Heure Fin",
    "Quantité",
    "Arrêt (min)",
    "Cadence",
    "Temps Restant",
    "Commentaire",
    "Article",
    "Nom Personnel",
    "Motif Personnel",
    "Visa",
    "Validée"
  ]];

  // PRODUCTION
  LINES.forEach(line => {
    const recs = srcState.production[line] || [];
    recs.forEach(r => {
      rows.push([
        "PRODUCTION",
        r.dateTime,
        r.equipe,
        line,
        "",
        "",
        r.start || "",
        r.end || "",
        r.quantity || 0,
        r.arret || 0,
        r.cadence || "",
        r.remainingTime || "",
        r.comment || "",
        r.article || "",
        "",
        "",
        "",
        ""
      ]);
    });
  });

  // ARRETS
  srcState.arrets.forEach(r => {
    rows.push([
      "ARRET",
      r.dateTime,
      "",
      r.line,
      r.sousLigne || "",
      r.machine || "",
      "",
      "",
      "",
      r.duration || 0,
      "",
      "",
      r.comment || "",
      r.article || "",
      "",
      "",
      "",
      ""
    ]);
  });

  // ORGANISATION
  srcState.organisation.forEach(r => {
    rows.push([
      "ORGANISATION",
      r.dateTime,
      r.equipe,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      r.consigne || "",
      "",
      "",
      "",
      r.visa || "",
      r.valide ? "Oui" : "Non"
    ]);
  });

  // PERSONNEL
  srcState.personnel.forEach(r => {
    rows.push([
      "PERSONNEL",
      r.dateTime,
      r.equipe,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      r.comment || "",
      "",
      r.nom || "",
      r.motif || "",
      "",
      ""
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  
  // Auto-largeur des colonnes
  const colWidths = rows[0].map((_, i) => {
    const maxLen = Math.max(...rows.map(row => String(row[i] || "").length));
    return { wch: Math.min(maxLen + 2, 50) };
  });
  ws['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, ws, "DATA");
  XLSX.writeFile(wb, filename);
}

// ===== EXPORT 2 : PRÉSENTATION (Réunion) - VERSION AMÉLIORÉE =====
function exportPresentationToExcel(srcState, filename) {
  if (typeof XLSX === 'undefined') {
    alert("Bibliothèque XLSX non chargée.");
    return;
  }

  const wb = XLSX.utils.book_new();

  // ===== ONGLET 1 : SYNTHÈSE EXÉCUTIVE =====
  const synthRows = [
    ["📊 RAPPORT DE PRODUCTION - ATELIER PPNC"],
    [""],
    ["📅 Date d'export", new Date().toLocaleDateString("fr-FR") + " à " + new Date().toLocaleTimeString("fr-FR")],
    ["👥 Équipe", srcState.currentEquipe],
    ["📍 Semaine", "S" + getWeekNumber(new Date())],
    ["📍 Quantième", getQuantieme(new Date())],
    [""],
    [""],
    ["═══════════════════════════════════════════════════════════════"],
    ["📦 PRODUCTION PAR LIGNE"],
    ["═══════════════════════════════════════════════════════════════"],
    [""],
    ["Ligne", "Quantité Totale (colis)", "Cadence Moyenne (colis/h)", "Articles produits"]
  ];

  let totalProduction = 0;
  LINES.forEach(line => {
    const recs = srcState.production[line] || [];
    const total = recs.reduce((s, r) => s + (r.quantity || 0), 0);
    totalProduction += total;
    const cadences = recs.map(r => r.cadence).filter(c => c && c > 0);
    const avgCad = cadences.length ? (cadences.reduce((s, c) => s + c, 0) / cadences.length).toFixed(1) : "-";
    const articles = [...new Set(recs.map(r => r.article).filter(a => a))].join(", ") || "-";
    
    synthRows.push([line, total, avgCad, articles]);
  });

  synthRows.push([]);
  synthRows.push(["🎯 TOTAL PRODUCTION", totalProduction + " colis", "", ""]);
  synthRows.push([]);
  synthRows.push([]);

  synthRows.push(["═══════════════════════════════════════════════════════════════"]);
  synthRows.push(["⚠️ TOP 10 ARRÊTS MAJEURS"]);
  synthRows.push(["═══════════════════════════════════════════════════════════════"]);
  synthRows.push([]);
  synthRows.push(["Ligne", "Sous-ligne", "Machine", "Durée (min)", "Commentaire"]);

  const totalArrets = srcState.arrets.reduce((s, r) => s + (r.duration || 0), 0);
  [...srcState.arrets]
    .sort((a, b) => (b.duration || 0) - (a.duration || 0))
    .slice(0, 10)
    .forEach(r => {
      synthRows.push([r.line, r.sousLigne || "-", r.machine, r.duration, r.comment || ""]);
    });

  synthRows.push([]);
  synthRows.push(["⏱️ DURÉE TOTALE ARRÊTS", totalArrets + " minutes", "", "", ""]);
  
  // Statistiques Organisation & Personnel
  if (srcState.organisation && srcState.organisation.length > 0) {
    synthRows.push([]);
    synthRows.push([]);
    synthRows.push(["═══════════════════════════════════════════════════════════════"]);
    synthRows.push(["📋 RÉSUMÉ ORGANISATION"]);
    synthRows.push(["═══════════════════════════════════════════════════════════════"]);
    synthRows.push([]);
    synthRows.push(["Nombre de consignes", srcState.organisation.length]);
    synthRows.push(["Consignes validées", srcState.organisation.filter(o => o.valide).length]);
  }

  if (srcState.personnel && srcState.personnel.length > 0) {
    synthRows.push([]);
    synthRows.push([]);
    synthRows.push(["═══════════════════════════════════════════════════════════════"]);
    synthRows.push(["👤 RÉSUMÉ PERSONNEL"]);
    synthRows.push(["═══════════════════════════════════════════════════════════════"]);
    synthRows.push([]);
    synthRows.push(["Nombre d'événements", srcState.personnel.length]);
    const motifs = {};
    srcState.personnel.forEach(p => {
      motifs[p.motif] = (motifs[p.motif] || 0) + 1;
    });
    Object.entries(motifs).forEach(([motif, count]) => {
      synthRows.push([motif, count]);
    });
  }

  const wsSynth = XLSX.utils.aoa_to_sheet(synthRows);
  
  wsSynth['!cols'] = [
    { wch: 25 },
    { wch: 25 },
    { wch: 25 },
    { wch: 20 },
    { wch: 50 }
  ];

  XLSX.utils.book_append_sheet(wb, wsSynth, "📊 SYNTHÈSE");

  // ===== ONGLET 2 : PRODUCTION DÉTAILLÉE =====
  const prodRows = [
    ["📦 PRODUCTION DÉTAILLÉE PAR LIGNE"],
    [""],
    ["Tous les enregistrements de production par ordre chronologique"],
    [""],
    ["Date/Heure", "Équipe", "Ligne", "Heure Début", "Heure Fin", "Quantité (colis)", "Arrêt (min)", "Cadence (colis/h)", "Code Article", "Commentaire"]
  ];

  LINES.forEach(line => {
    const recs = srcState.production[line] || [];
    if (recs.length > 0) {
      prodRows.push([]);
      prodRows.push(["▶ " + line, "", "", "", "", "", "", "", "", ""]);
      recs.forEach(r => {
        prodRows.push([
          r.dateTime,
          r.equipe,
          line,
          r.start || "-",
          r.end || "-",
          r.quantity || 0,
          r.arret || 0,
          r.cadence ? r.cadence.toFixed(1) : "-",
          r.article || "-",
          r.comment || ""
        ]);
      });
    }
  });

  const wsProd = XLSX.utils.aoa_to_sheet(prodRows);
  wsProd['!cols'] = [
    { wch: 20 },
    { wch: 10 },
    { wch: 15 },
    { wch: 12 },
    { wch: 12 },
    { wch: 18 },
    { wch: 12 },
    { wch: 18 },
    { wch: 15 },
    { wch: 45 }
  ];

  XLSX.utils.book_append_sheet(wb, wsProd, "📦 PRODUCTION");

  // ===== ONGLET 3 : ARRÊTS MACHINES =====
  const arretRows = [
    ["⚠️ HISTORIQUE COMPLET DES ARRÊTS MACHINES"],
    [""],
    ["Détail de tous les arrêts enregistrés avec durées et causes"],
    [""],
    ["Date/Heure", "Ligne", "Sous-ligne", "Machine", "Durée (min)", "Article concerné", "Commentaire / Cause"]
  ];

  // Tri par durée décroissante pour mettre en évidence les arrêts les plus longs
  [...srcState.arrets]
    .sort((a, b) => (b.duration || 0) - (a.duration || 0))
    .forEach((r, idx) => {
      arretRows.push([
        r.dateTime,
        r.line,
        r.sousLigne || "-",
        r.machine,
        r.duration || 0,
        r.article || "-",
        r.comment || ""
      ]);
    });

  arretRows.push([]);
  arretRows.push([]);
  const totalDuree = srcState.arrets.reduce((s, r) => s + (r.duration || 0), 0);
  arretRows.push(["📊 STATISTIQUES"]);
  arretRows.push(["Nombre total d'arrêts", srcState.arrets.length]);
  arretRows.push(["Durée cumulée", totalDuree + " minutes"]);
  arretRows.push(["Durée moyenne par arrêt", srcState.arrets.length > 0 ? (totalDuree / srcState.arrets.length).toFixed(1) + " minutes" : "-"]);

  const wsArrets = XLSX.utils.aoa_to_sheet(arretRows);
  wsArrets['!cols'] = [
    { wch: 20 },
    { wch: 15 },
    { wch: 12 },
    { wch: 18 },
    { wch: 12 },
    { wch: 18 },
    { wch: 50 }
  ];

  XLSX.utils.book_append_sheet(wb, wsArrets, "⚠️ ARRÊTS");

  // ===== ONGLET 4 : ORGANISATION & CONSIGNES =====
  if (srcState.organisation && srcState.organisation.length > 0) {
    const orgRows = [
      ["📋 ORGANISATION & CONSIGNES ATELIER"],
      [""],
      ["Toutes les consignes et instructions transmises pendant la période"],
      [""],
      ["Date/Heure", "Équipe", "Consigne / Instruction", "Visa Responsable", "Statut Validation"]
    ];

    srcState.organisation.forEach(r => {
      orgRows.push([
        r.dateTime,
        r.equipe,
        r.consigne,
        r.visa || "-",
        r.valide ? "✅ Validée" : "❌ En attente"
      ]);
    });

    orgRows.push([]);
    orgRows.push([]);
    orgRows.push(["📊 STATISTIQUES"]);
    orgRows.push(["Nombre de consignes", srcState.organisation.length]);
    orgRows.push(["Consignes validées", srcState.organisation.filter(o => o.valide).length]);
    orgRows.push(["En attente de validation", srcState.organisation.filter(o => !o.valide).length]);

    const wsOrg = XLSX.utils.aoa_to_sheet(orgRows);
    wsOrg['!cols'] = [
      { wch: 20 },
      { wch: 10 },
      { wch: 60 },
      { wch: 20 },
      { wch: 18 }
    ];

    XLSX.utils.book_append_sheet(wb, wsOrg, "📋 ORGANISATION");
  }

  // ===== ONGLET 5 : PERSONNEL & ÉVÉNEMENTS =====
  if (srcState.personnel && srcState.personnel.length > 0) {
    const persRows = [
      ["👤 ÉVÉNEMENTS PERSONNEL"],
      [""],
      ["Absences, retards et autres événements liés au personnel"],
      [""],
      ["Date/Heure", "Équipe", "Nom Collaborateur", "Motif", "Détails / Commentaire"]
    ];

    srcState.personnel.forEach(r => {
      const motifIcon = {
        "Absence": "❌",
        "Retard": "⏰",
        "Départ": "🚪",
        "Autre": "📝"
      };
      persRows.push([
        r.dateTime,
        r.equipe,
        r.nom,
        (motifIcon[r.motif] || "") + " " + r.motif,
        r.comment || "-"
      ]);
    });

    persRows.push([]);
    persRows.push([]);
    persRows.push(["📊 STATISTIQUES PAR MOTIF"]);
    const motifCount = {};
    srcState.personnel.forEach(p => {
      motifCount[p.motif] = (motifCount[p.motif] || 0) + 1;
    });
    Object.entries(motifCount).forEach(([motif, count]) => {
      persRows.push([motif, count + " événement(s)"]);
    });

    const wsPers = XLSX.utils.aoa_to_sheet(persRows);
    wsPers['!cols'] = [
      { wch: 20 },
      { wch: 10 },
      { wch: 25 },
      { wch: 18 },
      { wch: 50 }
    ];

    XLSX.utils.book_append_sheet(wb, wsPers, "👤 PERSONNEL");
  }

  XLSX.writeFile(wb, filename);
}

function bindExportGlobal() {
  const presentationBtn = document.getElementById("exportPresentationBtn");
  
  if (presentationBtn) {
    presentationBtn.addEventListener("click", () => {
      const now = getNow();
      const hh = String(now.getHours()).padStart(2, "0");
      const mm = String(now.getMinutes()).padStart(2, "0");
      const ss = String(now.getSeconds()).padStart(2, "0");

      const filename = `Atelier_PRESENTATION_${hh}h${mm}_${ss}.xlsx`;
      exportPresentationToExcel(state, filename);
    });
  }
}

/********************************************
 *   PARTIE RAZ ÉQUIPE
 ********************************************/

function bindRAZEquipe() {
  const btn = document.getElementById("razBtn");
  if (!btn) return;

  btn.addEventListener("click", () => {
    if (!confirm("RAZ + changement d'équipe + export ?")) return;

    const now = getNow();
    const week = getWeekNumber(now);
    const quantieme = getQuantieme(now);
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");

    let finished = prompt(
      "Quelle équipe vient de finir ? (M, AM, N)",
      state.currentEquipe
    );
    if (!finished) return;
    finished = finished.toUpperCase().trim();

    if (!["M", "AM", "N"].includes(finished)) {
      alert("Équipe invalide.");
      return;
    }

    const snap = {
      id: now.toISOString(),
      savedAt: formatDateTime(now),
      equipe: finished,
      week,
      quantieme,
      label: `Eq ${finished} • Q${quantieme} • S${week} • ${hh}h${mm}`,
      state: JSON.parse(JSON.stringify(state))
    };

    archives.push(snap);
    saveArchives();
    refreshHistorySelect();

    // ✅ DOUBLE EXPORT : DATA + PRÉSENTATION
    const filenameData = `Atelier_DATA_EQ${finished}_Q${quantieme}_S${week}_${hh}h${mm}.xlsx`;
    const filenamePres = `Atelier_PRESENTATION_EQ${finished}_Q${quantieme}_S${week}_${hh}h${mm}.xlsx`;
    
    exportDataToExcel(snap.state, filenameData);
    
    // Petit délai pour éviter que les 2 téléchargements se chevauchent
    setTimeout(() => {
      exportPresentationToExcel(snap.state, filenamePres);
    }, 500);

    let next = "M";
    if (finished === "M") next = "AM";
    else if (finished === "AM") next = "N";
    else next = "M";

    state.currentEquipe = next;

    LINES.forEach(l => {
      state.production[l] = [];
      state.formDraft[l] = {};
    });
    state.arrets = [];
    state.organisation = [];
    state.personnel = [];

    saveState();

    refreshProductionHistoryTable();
    refreshAtelierView();
    refreshArretsView();
    refreshOrganisationView();
    refreshPersonnelView();

    alert(`RAZ OK. Nouvelle équipe active : ${next}`);
  });
}

/********************************************
 *   DDM - VALIDATION DATE DURABILITÉ MINIMALE
 ********************************************/

function bindDDM() {
  const calcBtn = document.getElementById("ddmCalcBtn");
  if (!calcBtn) return;

  // Pré-remplir avec l'année courante
  const anneeEl = document.getElementById("ddmAnnee");
  if (anneeEl && !anneeEl.value) {
    anneeEl.value = getNow().getFullYear();
  }

  calcBtn.addEventListener("click", () => {
    const quantieme = Number(document.getElementById("ddmQuantieme")?.value);
    const annee = Number(document.getElementById("ddmAnnee")?.value);
    const duree = Number(document.getElementById("ddmDuree")?.value);

    const resultEl = document.getElementById("ddmResult");
    if (!resultEl) return;

    // Validation
    if (!quantieme || quantieme < 1 || quantieme > 366) {
      resultEl.textContent = "Quantième invalide (1-366)";
      resultEl.style.color = "red";
      return;
    }

    if (!annee || annee < 2000 || annee > 2100) {
      resultEl.textContent = "Année invalide";
      resultEl.style.color = "red";
      return;
    }

    if (!duree || duree < 0) {
      resultEl.textContent = "Durée invalide";
      resultEl.style.color = "red";
      return;
    }

    // Calcul : Quantième → Date → + durée
    const dateFab = quantiemeToDate(quantieme, annee);
    const dateDDM = new Date(dateFab);
    dateDDM.setDate(dateDDM.getDate() + duree);

    const ddmStr = formatDateDDM(dateDDM);
    const quantiemeDDM = getQuantieme(dateDDM);

    resultEl.textContent = `${ddmStr} (Quantième ${quantiemeDDM})`;
    resultEl.style.color = "green";
  });
}

/********************************************
 *   CALCULATRICE FLOTTANTE
 ********************************************/

function initCalculator() {
  const calcWidget = document.getElementById("calculator");
  const calcToggle = document.getElementById("calcToggle");
  const calcClose = document.getElementById("calcCloseBtn");
  const calcDisplay = document.getElementById("calcDisplay");

  if (!calcWidget || !calcToggle || !calcClose || !calcDisplay) return;

  let currentValue = "0";
  let operator = null;
  let previousValue = null;
  let waitingForOperand = false;

  function updateDisplay() {
    calcDisplay.value = currentValue;
  }

  function clear() {
    currentValue = "0";
    operator = null;
    previousValue = null;
    waitingForOperand = false;
    updateDisplay();
  }

  function inputDigit(digit) {
    if (waitingForOperand) {
      currentValue = String(digit);
      waitingForOperand = false;
    } else {
      currentValue = currentValue === "0" ? String(digit) : currentValue + digit;
    }
    updateDisplay();
  }

  function inputDecimal() {
    if (waitingForOperand) {
      currentValue = "0.";
      waitingForOperand = false;
    } else if (currentValue.indexOf(".") === -1) {
      currentValue += ".";
    }
    updateDisplay();
  }

  function performOperation(nextOperator) {
    const inputValue = parseFloat(currentValue);

    if (previousValue === null) {
      previousValue = inputValue;
    } else if (operator) {
      const result = calculate(previousValue, inputValue, operator);
      currentValue = String(result);
      previousValue = result;
    }

    waitingForOperand = true;
    operator = nextOperator;
    updateDisplay();
  }

  function calculate(prev, current, op) {
    switch (op) {
      case "+": return prev + current;
      case "-": return prev - current;
      case "*": return prev * current;
      case "/": return current !== 0 ? prev / current : 0;
      default: return current;
    }
  }

  // Toggle affichage
  calcToggle.addEventListener("click", () => {
    calcWidget.classList.toggle("hidden");
  });

  calcClose.addEventListener("click", () => {
    calcWidget.classList.add("hidden");
  });

  // Boutons calculatrice
  document.querySelectorAll(".calc-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const value = btn.dataset.value;
      const action = btn.dataset.action;

      if (value && !action) {
        if (value === ".") {
          inputDecimal();
        } else if (["+", "-", "*", "/"].includes(value)) {
          performOperation(value);
        } else {
          inputDigit(value);
        }
      } else if (action === "clear") {
        clear();
      } else if (action === "equals") {
        performOperation(null);
        operator = null;
        waitingForOperand = true;
      }
    });
  });

  updateDisplay();
}

/********************************************
 *   THÈME CLAIR / SOMBRE
 ********************************************/

function initTheme() {
  const btn = document.getElementById("themeToggleBtn");
  if (!btn) return;

  const saved = localStorage.getItem("themeMode");
  if (saved === "light") {
    document.body.classList.add("light-mode");
  }

  btn.addEventListener("click", () => {
    document.body.classList.toggle("light-mode");

    if (document.body.classList.contains("light-mode")) {
      localStorage.setItem("themeMode", "light");
    } else {
      localStorage.setItem("themeMode", "dark");
    }
  });
}

/********************************************
 *   ORIENTATION
 ********************************************/

function updateOrientationLayout() {
  const isLandscape = window.innerWidth > window.innerHeight;
  document.body.classList.toggle("is-landscape", isLandscape);
}

/********************************************
 *   INIT GLOBALE
 ********************************************/

document.addEventListener("DOMContentLoaded", () => {
  loadState();
  loadArchives();
  initHeaderDate();
  initEquipeSelector();
  initNav();
  bindExcelImport();
  bindManagerArea();
  initLinesSidebar();
  bindProductionForm();
  initArretsForm();
  bindOrganisationForm();
  bindPersonnelForm();
  bindExportGlobal();
  initCalculator();
  bindDDM();
  bindRAZEquipe();
  initHistoriqueEquipes();
  initTheme();

  updateOrientationLayout();
  window.addEventListener("resize", updateOrientationLayout);
  window.addEventListener("orientationchange", updateOrientationLayout);

  showSection(state.currentSection || "atelier");
});
