/******************************
 *   PARTIE 1 / 4 – CONFIG
 ******************************/

// === CONFIG & ÉTAT GLOBAL ===

const LINES = [
  "Râpé",
  "T2",
  "OMORI",
  "T1",
  "Emballage",
  "Dés",
  "Filets",
  "Prédécoupé",
];

// Stockage version
const STORAGE_KEY = "atelier_ppnc_state_v2";
const ARCHIVES_KEY = "atelier_ppnc_archives_v2";
const PLANNING_SNAPSHOTS_KEY = "planning_snapshots_v1";
let planningSnapshotCache = [];
let planningSnapshots = [];

let archives = []; // [{ id, label, savedAt, equipe, week, quantieme, state }]
let historyFiles = []; // Fichiers Excel présents dans honor200/Documents

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
  planning: {
    weekNumber: "",
    weekStart: "",
    orders: [],
    arretsPlanifies: [],
    cadences: [],
    savedPlans: [],
    selectedPlanWeek: "",
    activeOrders: [],
    activeArrets: [],
  },
};

const PLANNING_DAY_LABELS = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam"];

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
const MANAGER_DIR_DB = "manager_dir_handle_db";
const MANAGER_DIR_STORE = "handles";
const MANAGER_DIR_KEY = "manager_import_dir";
const MANAGER_AUTO_SCAN_MS = 30 * 60 * 1000;
const DEFAULT_MANAGER_PASSWORD = "3005";

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

const managerParetoFilters = {
  dateStart: "",
  dateEnd: "",
  ligne: "",
};

let managerDataset = [];
let managerImportLog = [];
let managerUnlocked = false;
let managerDirHandle = null;
let managerAutoScanTimer = null;
let managerFilteredRows = [];
let managerParetoChart = null;
let managerUnlockPromiseResolve = null;
// Historique équipes

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

function parseRecordDate(row) {
  const value = row?.dateHeure || "";
  const match = value.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!match) return null;
  const [, dd, mm, yyyy] = match;
  const year = Number(yyyy.length === 2 ? `20${yyyy}` : yyyy);
  const month = Number(mm) - 1;
  const day = Number(dd);
  const timeMatch = value.match(/(\d{1,2})h(\d{2})/i);
  const hours = timeMatch ? Number(timeMatch[1]) : 0;
  const minutes = timeMatch ? Number(timeMatch[2]) : 0;
  const d = new Date(year, month, day, hours, minutes);
  return Number.isNaN(d.getTime()) ? null : d;
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
      planning: {
        weekNumber: "",
        weekStart: "",
        orders: [],
        arretsPlanifies: [],
        cadences: [],
        savedPlans: [],
        selectedPlanWeek: "",
      },
    };

  state = Object.assign(base, parsed);

  if (!LINES.includes(state.currentLine)) {
    state.currentLine = LINES[0];
  }

  LINES.forEach(l => {
    if (!state.production[l]) state.production[l] = [];
    if (!state.formDraft[l]) state.formDraft[l] = {};
    });

    if (!state.excelRecords) state.excelRecords = [];
    if (!state.excelFiles) state.excelFiles = [];
    if (!state.planning) {
      state.planning = { ...base.planning };
    }
    if (!state.planning.orders) state.planning.orders = [];
    if (!state.planning.arretsPlanifies) state.planning.arretsPlanifies = [];
    if (!state.planning.cadences) state.planning.cadences = [];
    if (!state.planning.savedPlans) state.planning.savedPlans = [];
    if (!state.planning.activeOrders) state.planning.activeOrders = [];
    if (!state.planning.activeArrets) state.planning.activeArrets = [];

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

function readPlanningSnapshotsFromStorage() {
  try {
    const raw = localStorage.getItem(PLANNING_SNAPSHOTS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    if (Array.isArray(parsed) && parsed.length) planningSnapshotCache = parsed;
    return parsed.length ? parsed : planningSnapshotCache;
  } catch (e) {
    console.error("readPlanningSnapshotsFromStorage error", e);
    return planningSnapshotCache;
  }
}

function savePlanningSnapshots(snaps) {
  planningSnapshotCache = Array.isArray(snaps) ? [...snaps] : [];
  planningSnapshots = Array.isArray(snaps) ? [...snaps] : [];
  try {
    localStorage.setItem(PLANNING_SNAPSHOTS_KEY, JSON.stringify(snaps || []));
  } catch (e) {
    console.error("savePlanningSnapshots error", e);
  }
}

function getMergedPlanningSnapshots() {
  const stored = readPlanningSnapshotsFromStorage();
  const fromState = state?.planning?.savedPlans || [];
  const merged = [];

  if (planningSnapshots.length) {
    planningSnapshots.forEach(snap => {
      if (!snap || !snap.week) return;
      const exists = merged.some(m => `${m.week}` === `${snap.week}`);
      if (!exists) merged.push(snap);
    });
  }

  [ ...(Array.isArray(stored) ? stored : []), ...(Array.isArray(fromState) ? fromState : []) ].forEach(snap => {
    if (!snap || !snap.week) return;
    const exists = merged.some(m => `${m.week}` === `${snap.week}`);
    if (!exists) merged.push(snap);
  });

  return merged;
}

function loadPlanningSnapshots() {
  try {
    const merged = getMergedPlanningSnapshots();
    planningSnapshots = [...merged];
    if (!state.planning) state.planning = { savedPlans: [] };
    state.planning.savedPlans = merged;
    savePlanningSnapshots(merged);
  } catch (e) {
    console.error("loadPlanningSnapshots error", e);
    if (!state.planning) state.planning = { savedPlans: [] };
    state.planning.savedPlans = state.planning.savedPlans || [];
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

function setManagerFolderStatus(message, variant = "") {
  setSharedFolderStatus(message, variant);
}

function setManagerSecurityStatus(message, variant = "") {
  ["settingsSecurityStatus", "managerUnlockStatus"].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = message;
    el.className = "import-status helper-text";
    if (variant) el.classList.add(variant);
  });
}

function refreshSettingsFolderPath() {
  const el = document.getElementById("settingsFolderPath");
  if (!el) return;
  if (managerDirHandle) {
    el.textContent = `Dossier actuel : ${managerDirHandle.name}`;
  } else {
    el.textContent = "Dossier actuel : Documents (par défaut)";
  }
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

function setHistoryFolderStatus(message, variant = "") {
  const el = document.getElementById("historyFolderStatus");
  if (!el) return;
  el.textContent = message;
  el.className = "import-status helper-text";
  if (variant) el.classList.add(variant);
}

function setSharedFolderStatus(message, variant = "") {
  ["managerFolderStatus", "settingsFolderStatus"].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = message;
    el.className = "import-status helper-text";
    if (variant) el.classList.add(variant);
  });
}

function openDirHandleDB() {
  return new Promise(resolve => {
    const req = indexedDB.open(MANAGER_DIR_DB, 1);
    req.onerror = () => resolve(null);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(MANAGER_DIR_STORE)) {
        db.createObjectStore(MANAGER_DIR_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
  });
}

async function saveManagerDirectoryHandle(handle) {
  try {
    const db = await openDirHandleDB();
    if (!db) return false;
    return await new Promise(resolve => {
      const tx = db.transaction(MANAGER_DIR_STORE, "readwrite");
      tx.objectStore(MANAGER_DIR_STORE).put(handle, MANAGER_DIR_KEY);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    });
  } catch (e) {
    console.error("Impossible d'enregistrer le dossier import", e);
    return false;
  }
}

async function loadManagerDirectoryHandle() {
  try {
    const db = await openDirHandleDB();
    if (!db) return null;
    return await new Promise(resolve => {
      const tx = db.transaction(MANAGER_DIR_STORE, "readonly");
      const req = tx.objectStore(MANAGER_DIR_STORE).get(MANAGER_DIR_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (e) {
    console.error("Impossible de relire le dossier import", e);
    return null;
  }
}

async function getPreferredDirectoryHandle({ prompt = false, write = false, silent = false } = {}) {
  const permMode = write ? "readwrite" : "read";

  if (managerDirHandle) {
    const ok = await ensureDirectoryPermission(managerDirHandle, permMode);
    if (ok) return managerDirHandle;
  }

  const savedHandle = await loadManagerDirectoryHandle();
  if (savedHandle) {
    const ok = await ensureDirectoryPermission(savedHandle, permMode);
    if (ok) {
      managerDirHandle = savedHandle;
      setManagerFolderStatus(`Dossier mémorisé : ${savedHandle.name}`, "success");
      refreshSettingsFolderPath();
      return savedHandle;
    }
  }

  if (!prompt || !window.showDirectoryPicker) return null;

  try {
    const picked = await window.showDirectoryPicker({ startIn: "documents" });
    const ok = await ensureDirectoryPermission(picked, permMode);
    if (!ok) return null;
    managerDirHandle = picked;
    await saveManagerDirectoryHandle(picked);
    setManagerFolderStatus(`Dossier mémorisé : ${picked.name}`, "success");
    setHistoryFolderStatus(`Dossier prêt : ${picked.name}`, "success");
    refreshSettingsFolderPath();
    return picked;
  } catch (e) {
    if (!silent) console.warn("Sélection dossier annulée", e);
    return null;
  }
}

async function saveBlobToDirectory(dirHandle, filename, blob) {
  try {
    const ok = await ensureDirectoryPermission(dirHandle, "readwrite");
    if (!ok) return false;
    const fileHandle = await dirHandle.getFileHandle(filename, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return true;
  } catch (e) {
    console.error("Écriture impossible dans le dossier choisi", e);
    return false;
  }
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
  if (!stored) return false;
  const hashed = await hashPassword(pwd);
  return hashed === stored;
}

async function ensureDefaultManagerPassword() {
  if (getStoredPasswordHash()) return;
  const hashed = await hashPassword(DEFAULT_MANAGER_PASSWORD);
  localStorage.setItem(MANAGER_PASSWORD_KEY, hashed);
}

async function storeManagerPassword(pwd) {
  const hashed = await hashPassword(pwd);
  localStorage.setItem(MANAGER_PASSWORD_KEY, hashed);
}

function isManagerLocked() {
  return Boolean(getStoredPasswordHash()) && !managerUnlocked;
}

async function createManagerPassword(pwd, confirm) {
  if (getStoredPasswordHash()) {
    setManagerSecurityStatus("Un mot de passe existe déjà. Utilise le formulaire de modification.", "warning");
    return false;
  }
  if (!pwd || pwd.length < 4) {
    setManagerSecurityStatus("Choisis un mot de passe d'au moins 4 caractères.", "warning");
    return false;
  }
  if (pwd !== confirm) {
    setManagerSecurityStatus("Les deux mots de passe ne correspondent pas.", "error");
    return false;
  }
  await storeManagerPassword(pwd);
  managerUnlocked = true;
  setManagerSecurityStatus("Mot de passe créé et accès déverrouillé.", "success");
  updateManagerLockUI();
  return true;
}

async function changeManagerPassword(oldPwd, newPwd, confirm) {
  if (!newPwd || newPwd.length < 4) {
    setManagerSecurityStatus("Choisis un nouveau mot de passe d'au moins 4 caractères.", "warning");
    return false;
  }
  if (newPwd !== confirm) {
    setManagerSecurityStatus("La confirmation du nouveau mot de passe ne correspond pas.", "error");
    return false;
  }
  const ok = await verifyManagerPassword(oldPwd);
  if (!ok) {
    setManagerSecurityStatus("Ancien mot de passe incorrect.", "error");
    return false;
  }
  await storeManagerPassword(newPwd);
  managerUnlocked = true;
  setManagerSecurityStatus("Mot de passe mis à jour et accès déverrouillé.", "success");
  updateManagerLockUI();
  return true;
}

function closeManagerUnlockModal(success = false) {
  const modal = document.getElementById("managerUnlockModal");
  const input = document.getElementById("managerUnlockInput");
  if (modal) modal.classList.add("hidden");
  if (input) input.value = "";
  if (managerUnlockPromiseResolve) {
    managerUnlockPromiseResolve(success);
    managerUnlockPromiseResolve = null;
  }
}

function openManagerUnlockModal() {
  const modal = document.getElementById("managerUnlockModal");
  const input = document.getElementById("managerUnlockInput");
  if (!modal || !input) return;
  modal.classList.remove("hidden");
  setManagerSecurityStatus("", "");
  setTimeout(() => input.focus(), 20);
}

async function promptManagerUnlock() {
  await ensureDefaultManagerPassword();
  openManagerUnlockModal();
  return new Promise(resolve => {
    managerUnlockPromiseResolve = resolve;
  });
}

function updateManagerLockUI() {
  const content = document.getElementById("managerContent");
  const resultsCard = document.getElementById("managerResultsCard");
  const paretoCard = document.getElementById("managerParetoCard");

  const locked = isManagerLocked();

  if (locked) {
    if (content) content.style.display = "none";
    if (resultsCard) resultsCard.style.display = "none";
    if (paretoCard) paretoCard.style.display = "none";
    setManagerStatus("Accès manager verrouillé. Clique sur Manager pour saisir le mot de passe.", "warning");
  } else {
    if (content) content.style.display = "grid";
    if (resultsCard) resultsCard.style.display = "block";
    if (paretoCard && managerParetoChart) paretoCard.style.display = "block";
  }
}

async function handleUnlock(password) {
  const ok = await verifyManagerPassword(password);
  if (!ok) {
    setManagerSecurityStatus("Mot de passe incorrect.", "error");
    managerUnlocked = false;
    updateManagerLockUI();
    return false;
  }
  managerUnlocked = true;
  setManagerSecurityStatus("Accès manager déverrouillé.", "success");
  updateManagerLockUI();
  refreshManagerImports();
  refreshManagerResults();
  closeManagerUnlockModal(true);
  return true;
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

async function populateManagerParetoLineFilter() {
  const select = document.getElementById("managerParetoLine");
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

function escapeHtml(str = "") {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function highlightText(str, query) {
  const safe = escapeHtml(str ?? "");
  if (!query) return safe;
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const regex = new RegExp(`(${escaped})`, "gi");
  return safe.replace(regex, "<mark>$1</mark>");
}

function formatQSH(meta) {
  const q = meta.quantieme ? `Q${String(meta.quantieme).padStart(3, "0")}` : "Q???";
  const s = meta.semaine ? `S${meta.semaine}` : "S??";
  const h = meta.heureFichier || "-";
  return `${q} / ${s} / ${h}`;
}

function renderFilterChips(filters) {
  const container = document.getElementById("managerActiveFilters");
  if (!container) return;
  container.innerHTML = "";

  if (!filters.length) {
    const chip = document.createElement("div");
    chip.className = "filter-chip";
    chip.textContent = "Aucun filtre (tous les champs)";
    container.appendChild(chip);
    return;
  }

  filters.forEach(f => {
    const chip = document.createElement("div");
    chip.className = "filter-chip";
    chip.innerHTML = `<span>${f.label}</span> ${escapeHtml(f.value)}`;
    container.appendChild(chip);
  });
}

function renderManagerBadges(filtered, total) {
  const wrap = document.getElementById("managerResultBadges");
  if (!wrap) return;
  wrap.innerHTML = "";

  const badges = [
    { label: "Affichés", value: filtered },
    { label: "Total", value: total },
    { label: "Tri", value: `${managerSearchState.sortField} (${managerSearchState.sortDir === "asc" ? "↗" : "↘"})` },
  ];

  badges.forEach(b => {
    const div = document.createElement("div");
    div.className = "result-badge";
    div.textContent = `${b.label} : ${b.value}`;
    wrap.appendChild(div);
  });
}

function renderManagerStats(filtered) {
  const stats = document.getElementById("managerStats");
  if (!stats) return;
  stats.innerHTML = "";

  const sortedLog = [...managerImportLog].sort((a, b) => new Date(b.importedAt) - new Date(a.importedAt));
  const last = sortedLog[0];
  const items = [
    { label: "Lignes totales", value: managerDataset.length },
    { label: "Résultats filtrés", value: filtered },
    { label: "Fichiers importés", value: managerImportLog.length },
    { label: "Dernier import", value: last?.importedAt ? formatDateTime(new Date(last.importedAt)) : "-" },
  ];

  items.forEach(item => {
    const card = document.createElement("div");
    card.className = "stat-card";
    card.innerHTML = `<span class="stat-label">${item.label}</span><span class="stat-value">${item.value}</span>`;
    stats.appendChild(card);
  });
}

function renderManagerParetoBadges(count, totalMinutes) {
  const wrap = document.getElementById("managerParetoBadges");
  if (!wrap) return;
  wrap.innerHTML = "";

  const badges = [
    { label: "Arrêts comptés", value: count },
    { label: "Minutes", value: Math.round(totalMinutes) },
  ];

  badges.forEach(b => {
    const div = document.createElement("div");
    div.className = "result-badge";
    div.textContent = `${b.label} : ${b.value}`;
    wrap.appendChild(div);
  });
}

function renderManagerPareto(options = {}) {
  const { scroll = false } = options;
  const card = document.getElementById("managerParetoCard");
  const canvas = document.getElementById("managerPareto");
  const titleEl = document.getElementById("managerParetoTitle");
  const subtitleEl = document.getElementById("managerParetoSubtitle");
  const clickInfo = document.getElementById("managerParetoClickInfo");
  if (!card || !canvas || typeof Chart === "undefined") return;

  const resetTexts = () => {
    if (titleEl) titleEl.textContent = "Pareto des arrêts";
    if (subtitleEl) subtitleEl.textContent = "Camembert basé sur les arrêts (minutes) filtrés par période ou ligne.";
    if (clickInfo) clickInfo.textContent = "";
  };

  const { dateStart, dateEnd, ligne } = managerParetoFilters;
  if (!dateStart && !dateEnd && !ligne) {
    resetTexts();
    card.style.display = "none";
    renderManagerParetoBadges(0, 0);
    if (managerParetoChart) {
      managerParetoChart.destroy();
      managerParetoChart = null;
    }
    return;
  }

  let rows = (managerFilteredRows || []).filter(r => Number(r.arretMinutes) > 0);

  if (ligne) rows = rows.filter(r => (r.ligne || "") === ligne);

  const startDate = dateStart ? new Date(dateStart) : null;
  const endDate = dateEnd ? new Date(`${dateEnd}T23:59:59`) : null;

  if (startDate || endDate) {
    rows = rows.filter(r => {
      const d = parseRecordDate(r);
      if (!d) return false;
      if (startDate && d < startDate) return false;
      if (endDate && d > endDate) return false;
      return true;
    });
  }

  if (!rows.length) {
    resetTexts();
    card.style.display = "none";
    renderManagerParetoBadges(0, 0);
    if (managerParetoChart) {
      managerParetoChart.destroy();
      managerParetoChart = null;
    }
    return;
  }

  const formatDateLabel = value => {
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return value;
    return d.toLocaleDateString("fr-FR");
  };

  const lineLabel = ligne ? `Ligne ${ligne}` : "Toutes lignes";
  const periodLabel = dateStart || dateEnd
    ? `${dateStart ? formatDateLabel(dateStart) : "Début"} → ${dateEnd ? formatDateLabel(dateEnd) : "Fin"}`
    : "Période complète";

  if (titleEl) titleEl.textContent = `Pareto des arrêts · ${lineLabel}`;
  if (subtitleEl) subtitleEl.textContent = `Filtre : ${lineLabel} | ${periodLabel}`;
  if (clickInfo) clickInfo.textContent = "Clique sur une part pour voir le détail complet.";

  const buckets = {};
  rows.forEach(r => {
    const key = r.commentaire || r.motifPersonnel || r.machine || r.ligne || "Autre";
    const val = Number(r.arretMinutes) || 0;
    buckets[key] = (buckets[key] || 0) + val;
  });

  const entries = Object.entries(buckets)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 12);

  const labels = entries.map(([k]) => k);
  const data = entries.map(([, v]) => v);

  if (managerParetoChart) managerParetoChart.destroy();

  const titleText = [`${lineLabel}`, periodLabel];

  managerParetoChart = new Chart(canvas, {
    type: "pie",
    data: {
      labels,
      datasets: [
        {
          label: "Arrêts (min)",
          data,
          backgroundColor: labels.map((_, idx) => `hsl(${(idx * 45) % 360} 70% 55%)`),
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: "right" },
        title: { display: true, text: titleText },
        tooltip: {
          callbacks: {
            label: ctx => `${ctx.label} : ${Math.round(ctx.parsed)} min`,
          },
        },
      },
    },
  });

  const totalMinutes = data.reduce((s, v) => s + v, 0);
  const showSliceDetail = (idx, evt) => {
    if (!managerParetoChart) return;
    const value = data[idx] || 0;
    const pct = totalMinutes ? ((value / totalMinutes) * 100).toFixed(1) : 0;
    if (managerParetoChart.tooltip?.setActiveElements) {
      managerParetoChart.setActiveElements([{ datasetIndex: 0, index: idx }]);
      managerParetoChart.tooltip.setActiveElements(
        [{ datasetIndex: 0, index: idx }],
        evt ? { x: evt.offsetX, y: evt.offsetY } : undefined
      );
      managerParetoChart.update();
    }
    if (clickInfo) {
      clickInfo.innerHTML = `<strong>${labels[idx]}</strong><br>${Math.round(value)} min (${pct} %)` +
        (pct ? ` (${pct} %)` : "");
    }
  };

  canvas.onclick = evt => {
    const points = managerParetoChart.getElementsAtEventForMode(evt, "nearest", { intersect: true }, true);
    if (!points.length) return;
    showSliceDetail(points[0].index, evt);
  };

  renderManagerParetoBadges(rows.length, totalMinutes);
  card.style.display = "block";
  if (scroll) {
    card.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function refreshManagerResults() {
  const table = document.getElementById("managerResultsTable");
  if (!table) return;

  let rows = [...managerDataset];
  const totalRows = rows.length;
  const query = managerSearchState.text.trim();
  const queryLower = query.toLowerCase();

  const filters = [];

  if (query && managerSearchState.fields.size) {
    filters.push({ label: "Texte", value: query });
    rows = rows.filter(r => {
      return Array.from(managerSearchState.fields).some(key => {
        const val = r[key];
        return val !== undefined && val !== null && String(val).toLowerCase().includes(queryLower);
      });
    });
  }

  if (managerSearchState.equipe) {
    filters.push({ label: "Équipe", value: managerSearchState.equipe });
    rows = rows.filter(
      r => (r.equipe || "").toUpperCase() === managerSearchState.equipe.toUpperCase()
    );
  }

  if (managerSearchState.ligne) {
    filters.push({ label: "Ligne", value: managerSearchState.ligne });
    rows = rows.filter(r => (r.ligne || "") === managerSearchState.ligne);
  }

  if (managerSearchState.fields.size !== MANAGER_FIELDS.length) {
    filters.push({ label: "Champs", value: `${managerSearchState.fields.size}/${MANAGER_FIELDS.length}` });
  }

  const { sortField, sortDir } = managerSearchState;
  rows.sort((a, b) => {
    const av = a[sortField] ?? "";
    const bv = b[sortField] ?? "";

    if (av === bv) return 0;
    if (av > bv) return sortDir === "asc" ? 1 : -1;
    return sortDir === "asc" ? -1 : 1;
  });

  renderFilterChips(filters);
  renderManagerBadges(rows.length, totalRows);
  renderManagerStats(rows.length);

  managerFilteredRows = rows;

  renderManagerPareto();

  const tbody = table.querySelector("tbody");
  tbody.innerHTML = "";

  rows.forEach(row => {
    const tr = document.createElement("tr");
    const typeLabel = row.type || "Production / Arrêt";
    const qsh = formatQSH(row);
    tr.innerHTML = `
      <td><span class="type-pill">${escapeHtml(typeLabel)}</span></td>
      <td>${highlightText(row.dateHeure || "", query)}</td>
      <td>${highlightText(row.equipe || "", query)}</td>
      <td>${highlightText(row.ligne || "", query)}</td>
      <td>${highlightText(row.sousLigne || "", query)}</td>
      <td>${highlightText(row.machine || "", query)}</td>
      <td>${highlightText(row.quantite ?? "", query)}</td>
      <td>${highlightText(row.arretMinutes ?? "", query)}</td>
      <td>${highlightText(row.cadence ?? "", query)}</td>
      <td>${highlightText(row.commentaire || "", query)}</td>
      <td>${highlightText(row.article || "", query)}</td>
      <td>${highlightText(row.nomPersonnel || "", query)}</td>
      <td>${highlightText(row.motifPersonnel || "", query)}</td>
      <td>${highlightText(`${row.fileName || ""} — ${qsh}`, query)}</td>
    `;
    tbody.appendChild(tr);
  });

  const infoEl = document.getElementById("managerResultsCount");
  if (infoEl) {
    infoEl.textContent = `${rows.length} résultat(s) filtrés / ${managerDataset.length} dans le classeur.`;
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
  populateManagerParetoLineFilter();
}

async function ensureDirectoryPermission(dirHandle, mode = "read") {
  if (!dirHandle) return false;
  const opts = { mode };
  if ((await dirHandle.queryPermission(opts)) === "granted") return true;
  const perm = await dirHandle.requestPermission(opts);
  return perm === "granted";
}

async function applyDirectoryHandle(handle, { persist = true, silent = false } = {}) {
  if (!handle) return false;
  const ok = await ensureDirectoryPermission(handle);
  if (!ok) return false;

  managerDirHandle = handle;
  if (persist) await saveManagerDirectoryHandle(handle);

  refreshSettingsFolderPath();
  setManagerFolderStatus(`Dossier mémorisé : ${handle.name}`, "success");
  setHistoryFolderStatus(`Dossier prêt : ${handle.name}`, "success");

  scheduleManagerAutoScan();
  if (!silent) {
    scanFolderForManagerFiles(handle, true);
    scanHistoryFolder({ promptIfMissing: false });
  }
  return true;
}

async function requestManagerDirectory({ silent = false } = {}) {
  if (!window.showDirectoryPicker) return null;
  try {
    const handle = await window.showDirectoryPicker({ startIn: "documents" });
    const applied = await applyDirectoryHandle(handle, { silent });
    return applied ? handle : null;
  } catch (e) {
    if (!silent) console.error("Sélection du dossier import annulée", e);
    return null;
  }
}

function scheduleManagerAutoScan() {
  if (managerAutoScanTimer) clearInterval(managerAutoScanTimer);
  if (!managerDirHandle) return;
  managerAutoScanTimer = setInterval(() => {
    scanFolderForManagerFiles(managerDirHandle, true);
  }, MANAGER_AUTO_SCAN_MS);
}

async function restoreSavedManagerFolder() {
  const savedHandle = await loadManagerDirectoryHandle();
  if (!savedHandle) return;
  const ok = await ensureDirectoryPermission(savedHandle);
  if (!ok) {
    setManagerFolderStatus("Le dossier mémorisé nécessite une nouvelle autorisation.", "warning");
    refreshSettingsFolderPath();
    return;
  }
  await applyDirectoryHandle(savedHandle, { silent: true });
}

async function scanFolderForManagerFiles(dirHandle = null, silent = false) {
  const input = document.getElementById("managerExcelInput");
  if (!window.showDirectoryPicker) {
    if (!silent) {
      setManagerStatus("Ton navigateur ne permet pas de scanner un dossier. Utilise l'import manuel ci-dessous.", "warning");
      input?.click();
    }
    return;
  }

  let handle = dirHandle || managerDirHandle;
  if (!handle) {
    if (!silent) setManagerStatus("Choisis le dossier (ex : honor200 ➜ Documents) pour trouver les fichiers.", "warning");
    handle = await requestManagerDirectory();
  }
  if (!handle) return;

  const authorized = await ensureDirectoryPermission(handle);
  if (!authorized) {
    if (!silent) setManagerStatus("Autorise l'accès au dossier pour continuer.", "error");
    return;
  }

  const applied = await applyDirectoryHandle(handle, { silent: true });
  if (!applied) return;

  try {
    const files = [];
    for await (const entry of handle.values()) {
      if (entry.kind === "file" && /\.xlsx?$/i.test(entry.name)) {
        const file = await entry.getFile();
        files.push(file);
      }
    }

    if (!files.length) {
      if (!silent) setManagerStatus("Aucun fichier Excel trouvé dans ce dossier.", "warning");
      return;
    }

    const importedNames = new Set(managerImportLog.map(log => log.fileName));
    const freshFiles = files.filter(f => !importedNames.has(f.name));

    if (!freshFiles.length) {
      if (!silent) setManagerStatus("Tout est déjà importé pour ce dossier.", "success");
      return;
    }

    if (!silent) setManagerStatus(`Import en cours (${freshFiles.length} nouveau(x) fichier(s))...`);
    await importManagerFiles(freshFiles);
  } catch (e) {
    console.error("Scan dossier échoué", e);
    if (!silent) setManagerStatus("Impossible de scanner ce dossier (permissions ?)", "error");
  }
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
  populateManagerParetoLineFilter();
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
  populateManagerParetoLineFilter();
  updateManagerLockUI();
  setManagerFolderStatus("Sélectionne honor200 ➜ Documents pour activer la mise à jour auto.");

  const input = document.getElementById("managerExcelInput");
  const btn = document.getElementById("managerImportBtn");
  const scanBtn = document.getElementById("managerScanBtn");
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

  if (scanBtn) {
    scanBtn.addEventListener("click", async () => {
      setManagerStatus("Recherche des fichiers non importés...");
      await scanFolderForManagerFiles();
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

  const dateStart = document.getElementById("managerDateStart");
  if (dateStart) {
    dateStart.addEventListener("change", () => {
      managerParetoFilters.dateStart = dateStart.value;
      renderManagerPareto();
    });
  }

  const dateEnd = document.getElementById("managerDateEnd");
  if (dateEnd) {
    dateEnd.addEventListener("change", () => {
      managerParetoFilters.dateEnd = dateEnd.value;
      renderManagerPareto();
    });
  }

  const paretoLine = document.getElementById("managerParetoLine");
  if (paretoLine) {
    paretoLine.addEventListener("change", () => {
      managerParetoFilters.ligne = paretoLine.value;
      renderManagerPareto();
    });
  }

  const paretoBtn = document.getElementById("managerParetoBtn");
  paretoBtn?.addEventListener("click", () => renderManagerPareto({ scroll: true }));

  const paretoBackBtn = document.getElementById("managerParetoBackToTop");
  if (paretoBackBtn) {
    paretoBackBtn.addEventListener("click", () => {
      const target = document.getElementById("managerSearchCard") || document.getElementById("managerContent");
      target?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  const infoToggle = document.getElementById("managerInfoToggle");
  const infoContent = document.getElementById("managerInfoContent");
  if (infoToggle && infoContent) {
    infoToggle.addEventListener("click", () => {
      infoContent.classList.toggle("hidden");
    });
  }

  restoreSavedManagerFolder();
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
  if (section === "manager" && isManagerLocked()) {
    promptManagerUnlock();
    return;
  }
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
  else if (section === "historique") {
    scanHistoryFolder({ promptIfMissing: false });
    refreshHistorySelect();
    clearHistoryView();
  }
  else if (section === "manager") {
    populateManagerLineFilter();
    populateManagerParetoLineFilter();
    refreshManagerImports();
    refreshManagerResults();
  }
  else if (section === "planning") {
    refreshPlanningGantt();
    refreshPlanningDelays();
  }
}

function initNav() {
  document.querySelectorAll(".nav-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const target = btn.dataset.section;
      if (target === "manager" && isManagerLocked()) {
        const unlocked = await promptManagerUnlock();
        if (!unlocked) return;
      }
      showSection(target);
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

  if (!d.start) {
    const recs = getCurrentLineRecords();
    const last = recs[recs.length - 1];
    if (last?.end) d.start = last.end;
  }

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

      updatePlanningStatusFromProduction(L, rec);

      // pré-remplir la prochaine saisie avec l'heure de fin actuelle
      state.formDraft[L] = {
        start: rec.end || "",
      };
      saveState();

      refreshProductionForm();
      refreshProductionHistoryTable();
      refreshAtelierView();
      refreshPlanningGantt();

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
      refreshPlanningGantt();
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
      <td>
        <button class="secondary-btn danger" data-delete="${idx}">✕</button>
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

  tbody.querySelectorAll("button[data-delete]").forEach(btn => {
    btn.addEventListener("click", () => {
      const i = Number(btn.dataset.delete);
      if (!Number.isInteger(i)) return;
      state.organisation.splice(i, 1);
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
    const downtime = recs.reduce((s, r) => s + (r.arret || 0), 0);
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

    const tooltip = document.createElement("div");
    tooltip.className = "line-tooltip hidden";
    tooltip.innerHTML = `
      <h4>Ligne ${line}</h4>
      <p>Quantité cumulée : <strong>${total}</strong> colis</p>
      <p>Temps d'arrêts : <strong>${downtime}</strong> min</p>
      <div class="tooltip-actions">
        <button class="primary-btn access-line-btn" data-line="${line}">Accès ligne</button>
      </div>
    `;

    div.appendChild(tooltip);
    container.appendChild(div);

    div.addEventListener("click", e => {
      if (e.target.closest(".access-line-btn")) return;
      document.querySelectorAll(".line-tooltip").forEach(t => t.classList.add("hidden"));
      tooltip.classList.toggle("hidden");
    });

    tooltip.querySelector(".access-line-btn")?.addEventListener("click", e => {
      e.stopPropagation();
      document.querySelectorAll(".line-tooltip").forEach(t => t.classList.add("hidden"));
      showSection("production");
      selectLine(line, true);
    });
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

function formatHistoryLabel(meta, fileName) {
  const parts = [];
  if (meta.formattedDate) parts.push(meta.formattedDate);
  if (meta.heureFichier) parts.push(meta.heureFichier);
  parts.push(fileName);
  return parts.filter(Boolean).join(" • ");
}

function rebuildHistoryStateFromRows(rows) {
  const toStr = v => (v === undefined || v === null ? "" : String(v).trim());
  const toNum = v => {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  };

  const rebuilt = { production: {}, arrets: [], organisation: [], personnel: [] };
  LINES.forEach(l => rebuilt.production[l] = []);

  rows.forEach(row => {
    const type = toStr(row["Type"]).toUpperCase();
    if (!type) return;

    if (type.includes("PRODUCTION")) {
      const line = toStr(row["Ligne"]) || "Divers";
      if (!rebuilt.production[line]) rebuilt.production[line] = [];

      rebuilt.production[line].push({
        line,
        dateTime: toStr(row["Date/Heure"]),
        equipe: toStr(row["Équipe"] || row["Equipe"]),
        start: toStr(row["Heure Début"] || row["Heure debut"]),
        end: toStr(row["Heure Fin"]),
        quantity: toNum(row["Quantité"] || row["Quantite"]),
        arret: toNum(row["Arrêt (min)"] || row["Arret (min)"] || row["Arrêt"]),
        cadence: toNum(row["Cadence"]),
        remainingTime: toNum(row["Temps Restant"]),
        comment: toStr(row["Commentaire"]),
        article: toStr(row["Article"]),
      });
    } else if (type.includes("ARRET")) {
      rebuilt.arrets.push({
        line: toStr(row["Ligne"]),
        sousLigne: toStr(row["Sous-ligne"] || row["Sous ligne"] || row["Sous_ligne"]),
        machine: toStr(row["Machine"]),
        duration: toNum(row["Arrêt (min)"] || row["Arret (min)"] || row["Arrêt"]),
        comment: toStr(row["Commentaire"]),
        article: toStr(row["Article"]),
      });
    } else if (type.includes("ORGANISATION")) {
      rebuilt.organisation.push({
        dateTime: toStr(row["Date/Heure"]),
        equipe: toStr(row["Équipe"] || row["Equipe"]),
        consigne: toStr(row["Commentaire"]),
        visa: toStr(row["Visa"]),
        valide: toStr(row["Validée"] || row["Validee"]).toLowerCase().startsWith("o"),
      });
    } else if (type.includes("PERSONNEL")) {
      rebuilt.personnel.push({
        dateTime: toStr(row["Date/Heure"]),
        equipe: toStr(row["Équipe"] || row["Equipe"]),
        nom: toStr(row["Nom Personnel"]),
        motif: toStr(row["Motif Personnel"]),
        comment: toStr(row["Commentaire"]),
      });
    }
  });

  return rebuilt;
}

async function loadHistorySnapshotFromFileHandle(entry) {
  try {
    const file = entry.getFile ? await entry.getFile() : entry;
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheet = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheet];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const state = rebuildHistoryStateFromRows(rows);
    const meta = parseFileNameInfo(file.name || "");
    const label = formatHistoryLabel(meta, file.name || "Classeur historique");

    return { label, state };
  } catch (e) {
    console.error("Lecture historique dossier Documents impossible", e);
    setHistoryFolderStatus(`Lecture impossible : ${e.message || e}` , "error");
    return null;
  }
}

async function scanHistoryFolder({ promptIfMissing = false } = {}) {
  setHistoryFolderStatus("Scan du dossier Documents en cours…");

  const dirHandle = await getPreferredDirectoryHandle({ prompt: promptIfMissing, silent: !promptIfMissing });
  if (!dirHandle) {
    historyFiles = [];
    setHistoryFolderStatus("Choisis le dossier Excel dans Paramètres pour charger les archives.", "warning");
    refreshHistorySelect();
    return;
  }

  const hasPerm = await ensureDirectoryPermission(dirHandle, "read");
  if (!hasPerm) {
    historyFiles = [];
    setHistoryFolderStatus("Accès refusé : ouvre Paramètres et sélectionne le dossier Excel.", "error");
    refreshHistorySelect();
    return;
  }

  const collected = [];
  try {
    for await (const entry of dirHandle.values()) {
      if (entry.kind !== "file") continue;
      const lower = entry.name.toLowerCase();
      if (!lower.endsWith(".xlsx")) continue;
      if (!lower.startsWith("atelier")) continue;
      collected.push(entry);
    }
  } catch (e) {
    console.error("Scan dossier Documents impossible", e);
    setHistoryFolderStatus("Scan du dossier Documents impossible.", "error");
  }

  historyFiles = collected.sort((a, b) => b.name.localeCompare(a.name));

  if (historyFiles.length) {
    setHistoryFolderStatus(`${historyFiles.length} fichier(s) trouvés dans ${dirHandle.name}.`, "success");
  } else {
    setHistoryFolderStatus(`Aucun fichier Atelier trouvé dans ${dirHandle.name}.`, "warning");
  }

  refreshHistorySelect();
}

function initHistoriqueEquipes() {
  const select = document.getElementById("historySelect");
  if (!select) return;

  select.addEventListener("change", async () => {
    const value = select.value;
    clearHistoryView();

    if (!value) return;

    if (value.startsWith("file:")) {
      const idx = Number(value.split(":")[1]);
      const handle = historyFiles[idx];
      if (!handle) {
        setHistoryFolderStatus("Fichier introuvable dans le dossier Documents.", "warning");
        return;
      }
      const snap = await loadHistorySnapshotFromFileHandle(handle);
      if (snap) refreshHistoryView(snap);
      return;
    }

    if (value.startsWith("archive:")) {
      const snap = archives[Number(value.split(":")[1])];
      if (snap) refreshHistoryView(snap);
    }
  });

  refreshHistorySelect();
  scanHistoryFolder({ promptIfMissing: true });
}

function refreshHistorySelect() {
  const select = document.getElementById("historySelect");
  if (!select) return;

  const previous = select.value;
  select.innerHTML = `<option value="">-- Sélectionner --</option>`;

  historyFiles.forEach((entry, idx) => {
    const opt = document.createElement("option");
    opt.value = `file:${idx}`;
    opt.textContent = formatHistoryLabel(parseFileNameInfo(entry.name), entry.name);
    select.appendChild(opt);
  });

  archives.forEach((snap, idx) => {
    const opt = document.createElement("option");
    opt.value = `archive:${idx}`;
    opt.textContent = snap.label;
    select.appendChild(opt);
  });

  if (select.querySelector(`option[value="${previous}"]`)) {
    select.value = previous;
  } else if (select.options.length > 1) {
    select.selectedIndex = 1;
  }

  if (select.value) {
    select.dispatchEvent(new Event("change"));
  } else {
    clearHistoryView();
  }
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

  const orgContainer = document.getElementById("history-organisation");
  if (orgContainer) orgContainer.style.display = "none";

  const persContainer = document.getElementById("history-personnel");
  if (persContainer) persContainer.style.display = "none";
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
async function persistWorkbookToPreferredFolder(wb, filename) {
  const buffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

  const handle = await getPreferredDirectoryHandle({ prompt: false, write: true, silent: true });
  if (handle && await saveBlobToDirectory(handle, filename, blob)) return true;

  const prompted = await getPreferredDirectoryHandle({ prompt: true, write: true });
  if (prompted && await saveBlobToDirectory(prompted, filename, blob)) return true;

  return false;
}

async function exportDataToExcel(srcState, filename) {
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

  const saved = await persistWorkbookToPreferredFolder(wb, filename);
  if (!saved) {
    XLSX.writeFile(wb, filename);
  }
}

// ===== EXPORT 2 : PRÉSENTATION (Réunion) - VERSION AMÉLIORÉE =====
async function exportPresentationToExcel(srcState, filename) {
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

  const saved = await persistWorkbookToPreferredFolder(wb, filename);
  if (!saved) {
    XLSX.writeFile(wb, filename);
  }
}

function bindExportGlobal() {
  const presentationBtn = document.getElementById("exportPresentationBtn");

  if (presentationBtn) {
    presentationBtn.addEventListener("click", async () => {
      const now = getNow();
      const hh = String(now.getHours()).padStart(2, "0");
      const mm = String(now.getMinutes()).padStart(2, "0");
      const ss = String(now.getSeconds()).padStart(2, "0");

      const filename = `Atelier_PRESENTATION_${hh}h${mm}_${ss}.xlsx`;
      await exportPresentationToExcel(state, filename);
    });
  }
}

/********************************************
 *   PARTIE RAZ ÉQUIPE
 ********************************************/

function bindRAZEquipe() {
  const btn = document.getElementById("razBtn");
  if (!btn) return;

  btn.addEventListener("click", async () => {
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
    
    await exportDataToExcel(snap.state, filenameData);
    await exportPresentationToExcel(snap.state, filenamePres);

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

function initManagerUnlockModal() {
  const form = document.getElementById("managerUnlockForm");
  const cancelBtn = document.getElementById("managerUnlockCancel");
  const closeBtn = document.getElementById("managerUnlockClose");
  const input = document.getElementById("managerUnlockInput");

  form?.addEventListener("submit", async e => {
    e.preventDefault();
    const ok = await handleUnlock(input?.value || "");
    if (!ok) {
      setManagerSecurityStatus("Mot de passe incorrect.", "error");
    }
  });

  [cancelBtn, closeBtn].forEach(btn =>
    btn?.addEventListener("click", () => {
      closeManagerUnlockModal(false);
    })
  );
}

function initSettingsPanel() {
  const panel = document.getElementById("settingsPanel");
  const toggle = document.getElementById("settingsToggle");
  const close = document.getElementById("settingsClose");
  if (!panel || !toggle || !close) return;

  const hide = () => panel.classList.add("hidden");
  const show = () => panel.classList.remove("hidden");

  toggle.addEventListener("click", () => {
    panel.classList.toggle("hidden");
  });

  close.addEventListener("click", hide);

  const themeBtn = document.getElementById("settingsThemeBtn");
  themeBtn?.addEventListener("click", () => document.getElementById("themeToggleBtn")?.click());

  const headerToggle = document.getElementById("headerCompactToggle");
  const applyHeaderSize = () => {
    if (headerToggle?.checked) {
      document.body.classList.add("compact-header");
    } else {
      document.body.classList.remove("compact-header");
    }
  };
  headerToggle?.addEventListener("change", applyHeaderSize);
  applyHeaderSize();

  const folderBtn = document.getElementById("settingsFolderBtn");
  folderBtn?.addEventListener("click", async () => {
    const handle = await requestManagerDirectory();
    if (!handle) {
      setHistoryFolderStatus("Aucun dossier choisi. Sélectionne un dossier valide.", "warning");
    }
  });
  refreshSettingsFolderPath();
  const createForm = document.getElementById("settingsCreatePasswordForm");
  const createInput = document.getElementById("settingsCreatePassword");
  const createConfirm = document.getElementById("settingsCreatePasswordConfirm");
  createForm?.addEventListener("submit", async e => {
    e.preventDefault();
    const ok = await createManagerPassword(createInput?.value || "", createConfirm?.value || "");
    if (ok) hide();
    createForm.reset();
  });

  const changeForm = document.getElementById("settingsChangePasswordForm");
  const oldPwd = document.getElementById("settingsOldPassword");
  const newPwd = document.getElementById("settingsNewPassword");
  const newConfirm = document.getElementById("settingsNewPasswordConfirm");
  changeForm?.addEventListener("submit", async e => {
    e.preventDefault();
    const ok = await changeManagerPassword(oldPwd?.value || "", newPwd?.value || "", newConfirm?.value || "");
    if (ok) hide();
    changeForm.reset();
  });
}

/********************************************
 *   PARTIE 4B – PLANNING (GANTT)
 ********************************************/

const WEEK_TOTAL_MINUTES = (5 * 24 + 12) * 60; // lundi 00:00 → samedi 12:00
let planningEditId = null;

function getMonday(d = new Date()) {
  const date = new Date(d);
  const day = (date.getDay() + 6) % 7; // 0 = lundi
  date.setHours(0, 0, 0, 0);
  date.setDate(date.getDate() - day);
  return date;
}

function formatDateInput(d) {
  return d.toISOString().slice(0, 10);
}

function ensurePlanningDefaults() {
  if (!state.planning) return;
  if (!Array.isArray(state.planning.savedPlans)) state.planning.savedPlans = [];
  if (!Array.isArray(state.planning.orders)) state.planning.orders = [];
  if (!Array.isArray(state.planning.arretsPlanifies)) state.planning.arretsPlanifies = [];
  if (!Array.isArray(state.planning.activeOrders)) state.planning.activeOrders = [];
  if (!Array.isArray(state.planning.activeArrets)) state.planning.activeArrets = [];
  if (!state.planning.selectedPlanWeek) state.planning.selectedPlanWeek = "";
  if (!state.planning.weekStart) {
    const monday = getMonday();
    state.planning.weekStart = formatDateInput(monday);
    state.planning.weekNumber = getISOWeek(monday);
  }
  if (!state.planning.weekNumber) {
    state.planning.weekNumber = getISOWeek(new Date(state.planning.weekStart));
  }
}

function getPlanningWeekStartDate() {
  ensurePlanningDefaults();
  const d = new Date(state.planning.weekStart || new Date());
  d.setHours(0, 0, 0, 0);
  return d;
}

function getPlanningWeekEndDate() {
  const start = getPlanningWeekStartDate();
  const end = new Date(start);
  end.setMinutes(end.getMinutes() + WEEK_TOTAL_MINUTES);
  return end;
}

function getISOWeek(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function toDateFromDay(dayIndex, timeValue) {
  const start = getPlanningWeekStartDate();
  const d = new Date(start);
  d.setDate(d.getDate() + Number(dayIndex || 0));
  if (timeValue) {
    const [h, m] = timeValue.split(":").map(Number);
    d.setHours(h || 0, m || 0, 0, 0);
  }
  return d;
}

function getCadenceForArticle(code, line) {
  const lower = (code || "").toLowerCase();
  const match = state.planning.cadences.find(c => c.code.toLowerCase() === lower && (!line || c.line === line));
  if (match) return match;
  return state.planning.cadences.find(c => c.code.toLowerCase() === lower) || null;
}

function computeOrderDurationMinutes(order) {
  const qty = Number(order.quantity) || 0;
  const cad = Number(order.cadence) || 0;
  if (cad <= 0) return Math.max(30, qty);
  return (qty / cad) * 60;
}

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && aEnd > bStart;
}

function splitOrderWithStops(order, startDate, baseEnd, stops) {
  const relevant = (stops || [])
    .filter(s => new Date(s.end) > startDate && new Date(s.start) < baseEnd)
    .sort((a, b) => new Date(a.start) - new Date(b.start));

  const segments = [];
  let cursor = new Date(startDate);
  let extensionMs = 0;

  relevant.forEach(stop => {
    const ss = new Date(stop.start);
    const se = new Date(stop.end);
    if (se <= cursor) return;

    if (ss > cursor) {
      segments.push({
        id: order.id,
        parentId: order.id,
        code: order.code,
        label: order.label,
        line: order.line,
        quantity: order.quantity,
        status: order.status,
        blockedReason: order.blockedReason,
        start: cursor.toISOString(),
        end: ss.toISOString(),
      });
    }

    extensionMs += Math.max(0, se - ss);
    cursor = new Date(Math.max(cursor.getTime(), se.getTime()));
  });

  const finalEnd = new Date(baseEnd.getTime() + extensionMs);
  if (!segments.length || cursor < finalEnd) {
    segments.push({
      id: order.id,
      parentId: order.id,
      code: order.code,
      label: order.label,
      line: order.line,
      quantity: order.quantity,
      status: order.status,
      blockedReason: order.blockedReason,
      start: cursor.toISOString(),
      end: finalEnd.toISOString(),
    });
  }

  return { segments, finalEnd };
}

function recalibrateLine(line, dataset = {}) {
  const orders = (dataset.orders || state.planning.orders || []).filter(o => o.line === line);
  const plannedStops = dataset.plannedStops || state.planning.arretsPlanifies || [];
  const stops = getAllStopsForLine(line, plannedStops).sort((a, b) => new Date(a.start) - new Date(b.start));
  const sortedOrders = orders.sort((a, b) => new Date(a.start) - new Date(b.start));

  let cursor = null;
  sortedOrders.forEach(order => {
    let start = new Date(order.start);
    if (cursor && start < cursor) start = new Date(cursor);
    const overlappingStop = stops.find(s => new Date(s.start) <= start && start < new Date(s.end));
    if (overlappingStop) start = new Date(overlappingStop.end);

    const duration = computeOrderDurationMinutes(order);
    const baseEnd = new Date(start.getTime() + duration * 60000);
    const { segments, finalEnd } = splitOrderWithStops(order, start, baseEnd, stops);

    order.start = start.toISOString();
    order.end = finalEnd.toISOString();
    order.segments = segments;
    cursor = finalEnd;
  });
}

function getAllStopsForLine(line, plannedStops = state.planning.arretsPlanifies || []) {
  const startWeek = getPlanningWeekStartDate();
  const endWeek = getPlanningWeekEndDate();

  const plannedStopsForLine = (plannedStops || [])
    .filter(a => a.line === line)
    .map(a => ({
      ...a,
      start: a.start,
      end: new Date(new Date(a.start).getTime() + (Number(a.duration) || 0) * 60000).toISOString(),
      status: "stop",
      label: a.type,
    }));

  const recStops = state.arrets
    .filter(r => r.line === line)
    .map(r => ({
      start: new Date(r.dateTime).toISOString(),
      end: new Date(new Date(r.dateTime).getTime() + (Number(r.duration) || 0) * 60000).toISOString(),
      type: r.comment ? `Arrêt : ${r.comment}` : "Arrêt",
      line: r.line,
      duration: r.duration,
      status: "stop",
      id: `arret-${r.dateTime}-${r.duration}`,
      comment: r.comment || "",
    }))
    .filter(s => new Date(s.start) >= startWeek && new Date(s.start) <= endWeek);

  return [...plannedStopsForLine, ...recStops];
}

function renderPlanningCadences() {
  const container = document.getElementById("planningCadenceList");
  const datalist = document.getElementById("planningArticleCodesList");
  if (!container) return;
  const rows = state.planning.cadences
    .map(c => `<tr><td>${c.code}</td><td>${c.label || ""}</td><td>${c.cadence || "-"}</td><td>${c.poids || "-"} kg</td><td>${c.line || "-"}</td></tr>`)
    .join("");
  container.innerHTML = `<table><thead><tr><th>Code</th><th>Libellé</th><th>Cadence</th><th>Poids</th><th>Ligne</th></tr></thead><tbody>${rows || "<tr><td colspan=5>Aucune cadence enregistrée</td></tr>"}</tbody></table>`;

  if (datalist) {
    datalist.innerHTML = state.planning.cadences.map(c => `<option value="${c.code}"></option>`).join("");
  }

  updatePlanningOFPrereq();
}

function updatePlanningOFPrereq() {
  const btn = document.getElementById("planningOFAddBtn");
  const helper = document.getElementById("planningOFPrereq");
  const hasCadences = state.planning.cadences.length > 0;
  if (btn) btn.disabled = !hasCadences;
  if (helper) helper.textContent = hasCadences
    ? "Choisis un code article enregistré pour générer un OF."
    : "Enregistre d'abord les cadences articles pour activer la création d'OF.";
}

function findArticleByCode(code) {
  if (!code) return null;
  const lower = code.toLowerCase();
  return state.planning.cadences.find(c => c.code.toLowerCase() === lower) || null;
}

function autofillOFFieldsFromCode(code) {
  const article = findArticleByCode(code);
  const lineSelect = document.getElementById("planningOFLine");
  const cadenceInput = document.getElementById("planningOFcadence");
  if (article) {
    if (lineSelect) lineSelect.value = article.line || "";
    if (cadenceInput && (!cadenceInput.value || cadenceInput.value === "0")) {
      cadenceInput.value = article.cadence || "";
    }
  }
}

function renderPlanningArticlePicker() {
  const body = document.getElementById("planningArticlePickerBody");
  if (!body) return;
  if (!state.planning.cadences.length) {
    body.innerHTML = "<p class=\"helper-text\">Aucun article enregistré pour le moment.</p>";
    return;
  }
  const grouped = state.planning.cadences.reduce((acc, item) => {
    acc[item.line] = acc[item.line] || [];
    acc[item.line].push(item);
    return acc;
  }, {});
  const sections = Object.keys(grouped).map(line => {
    const rows = grouped[line]
      .map(a => `<button class="list-btn" data-code="${a.code}" data-line="${a.line}" data-cadence="${a.cadence}" data-poids="${a.poids || 0}" data-label="${a.label || ""}"><strong>${a.code}</strong> — ${a.label || ""}<br><small>${a.cadence || ""} colis/h • ${a.poids || 0} kg/colis</small></button>`)
      .join("");
    return `<div class="picker-section"><h4>${line}</h4><div class="picker-grid">${rows}</div></div>`;
  }).join("");
  body.innerHTML = sections;
  body.querySelectorAll(".list-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const codeInput = document.getElementById("planningOFCode");
      const lineSelect = document.getElementById("planningOFLine");
      const cadenceInput = document.getElementById("planningOFcadence");
      if (codeInput) codeInput.value = btn.dataset.code || "";
      if (lineSelect) lineSelect.value = btn.dataset.line || "";
      if (cadenceInput) cadenceInput.value = btn.dataset.cadence || "";
      closePlanningArticlePicker();
    });
  });
}

function openPlanningArticlePicker() {
  renderPlanningArticlePicker();
  const pop = document.getElementById("planningArticlePicker");
  if (pop) pop.classList.remove("hidden");
}

function closePlanningArticlePicker() {
  const pop = document.getElementById("planningArticlePicker");
  if (pop) pop.classList.add("hidden");
}

function refreshPlanningLineSelectors() {
  const selects = [
    document.getElementById("planningArretLine"),
    document.getElementById("planningArticleLine"),
    document.getElementById("planningOFLine"),
  ];
  selects.forEach(sel => {
    if (!sel) return;
    sel.innerHTML = "";
    if (sel.id === "planningArretLine") {
      const allOpt = document.createElement("option");
      allOpt.value = "ALL";
      allOpt.textContent = "Toutes les lignes";
      sel.appendChild(allOpt);
    }
    LINES.forEach(line => {
      const opt = document.createElement("option");
      opt.value = line;
      opt.textContent = line;
      sel.appendChild(opt);
    });
  });
}

function addPlanningCadence() {
  const code = document.getElementById("planningArticleCode")?.value.trim();
  if (!code) return;
  const label = document.getElementById("planningArticleLabel")?.value.trim() || "";
  const poids = Number(document.getElementById("planningArticlePoids")?.value) || 0;
  const cadence = Number(document.getElementById("planningArticleCadence")?.value) || 0;
  const line = document.getElementById("planningArticleLine")?.value || "";

  const existingIdx = state.planning.cadences.findIndex(c => c.code.toLowerCase() === code.toLowerCase() && c.line === line);
  const payload = { code, label, poids, cadence, line };
  if (existingIdx >= 0) state.planning.cadences[existingIdx] = payload;
  else state.planning.cadences.push(payload);
  saveState();
  renderPlanningCadences();
  updatePlanningOFPrereq();
  resetPlanningArticleForm();
}

function addPlanningStop() {
  const type = document.getElementById("planningArretType")?.value || "Autre";
  const line = document.getElementById("planningArretLine")?.value || LINES[0];
  const day = document.getElementById("planningArretDay")?.value || 0;
  const startTime = document.getElementById("planningArretStart")?.value || "00:00";
  const duration = Number(document.getElementById("planningArretDuration")?.value) || 0;
  const comment = document.getElementById("planningArretComment")?.value || "";

  const targetLines = line === "ALL" ? [...LINES] : [line];
  const targetDays = `${day}` === "ALL" ? [0, 1, 2, 3, 4, 5] : [Number(day)];

  targetLines.forEach(L => {
    targetDays.forEach(d => {
      const startDate = toDateFromDay(d, startTime);
      const stop = {
        id: generateId(),
        type,
        line: L,
        duration,
        start: startDate.toISOString(),
        comment,
      };
      state.planning.arretsPlanifies.push(stop);
    });
  });
  saveState();
  LINES.forEach(recalibrateLine);
  refreshPlanningGantt();
  closePlanningStopPopover();
}

function addPlanningOF() {
  const code = document.getElementById("planningOFCode")?.value || "";
  autofillOFFieldsFromCode(code);
  const lineSelect = document.getElementById("planningOFLine");
  let line = (lineSelect && lineSelect.value) || LINES[0];
  const qty = Number(document.getElementById("planningOFQty")?.value) || 0;
  const day = document.getElementById("planningOFDay")?.value || 0;
  const startTime = document.getElementById("planningOFStart")?.value || "00:00";
  const manualCad = Number(document.getElementById("planningOFcadence")?.value) || null;

  if (!state.planning.cadences.length) {
    alert("Commence par enregistrer les codes articles et cadences avant de créer un OF.");
    return;
  }

  const cadenceInfo = getCadenceForArticle(code, line);
  if (!cadenceInfo) {
    alert("Code article inconnu : ajoute-le dans 'Cadences articles' avant de créer l'OF.");
    return;
  }
  line = cadenceInfo.line || line;
  if (lineSelect) lineSelect.value = line;
  const cadence = manualCad || (cadenceInfo ? cadenceInfo.cadence : 0);
  const poids = cadenceInfo ? cadenceInfo.poids : 0;
  const label = cadenceInfo ? cadenceInfo.label : "";

  const startDate = toDateFromDay(day, startTime);
  const duration = (qty && cadence) ? (qty / cadence) * 60 : 60;
  const endDate = new Date(startDate.getTime() + duration * 60000);

  LINES.forEach(recalibrateLine);
  const conflict = state.planning.orders.some(o => {
    if (o.line !== line) return false;
    return rangesOverlap(startDate, endDate, new Date(o.start), new Date(o.end));
  });
  if (conflict) {
    alert("Un OF occupe déjà ce créneau sur cette ligne. Choisis un autre horaire ou déplace l'OF existant.");
    return;
  }

  const of = {
    id: generateId(),
    code,
    label,
    line,
    quantity: qty,
    cadence,
    poids,
    start: startDate.toISOString(),
    end: endDate.toISOString(),
    status: "planned",
    blockedReason: "",
  };

  state.planning.orders.push(of);
  recalibrateLine(line);
  saveState();
  refreshPlanningGantt();
  resetPlanningOFForm();
}

function resetPlanningArticleForm() {
  const ids = [
    "planningArticleCode",
    "planningArticleLabel",
    "planningArticlePoids",
    "planningArticleCadence",
  ];
  ids.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  const line = document.getElementById("planningArticleLine");
  if (line) line.value = line.querySelector("option")?.value || LINES[0] || "";
}

function resetPlanningOFForm() {
  const code = document.getElementById("planningOFCode");
  const qty = document.getElementById("planningOFQty");
  const start = document.getElementById("planningOFStart");
  const day = document.getElementById("planningOFDay");
  const cad = document.getElementById("planningOFcadence");
  if (code) code.value = "";
  if (qty) qty.value = "";
  if (start) start.value = "";
  if (cad) cad.value = "";
  if (day) day.value = "0";
  const line = document.getElementById("planningOFLine");
  if (line) line.value = line.querySelector("option")?.value || LINES[0] || "";
}

function refreshPlanningGantt() {
  renderPlanningGantt("planningGantt", { interactive: true });
  renderPlanningGantt("planningPreview", { interactive: false });
}

function resetPlanningPreview() {
  const preview = document.getElementById("planningPreview");
  if (preview) {
    preview.innerHTML = "<p class=\"helper-text\">Aucune prévisualisation en attente. Ajoute des OF pour voir l'aperçu.</p>";
  }
}

function renderPlanningGantt(targetId, { interactive = true } = {}) {
  const gantt = document.getElementById(targetId);
  if (!gantt) return;
  const start = getPlanningWeekStartDate();
  const end = getPlanningWeekEndDate();

  const isRunView = targetId === "planningGantt";
  const orders = isRunView ? (state.planning.activeOrders || []) : (state.planning.orders || []);
  const plannedStops = isRunView ? (state.planning.activeArrets || []) : (state.planning.arretsPlanifies || []);

  if (isRunView && !state.planning.selectedPlanWeek) {
    gantt.innerHTML = "<p class=\"helper-text\">Choisis un planning validé puis clique sur Lancer pour l'afficher ici.</p>";
    return;
  }

  const hasOrders = orders.length > 0;
  if (!interactive && !hasOrders) {
    resetPlanningPreview();
    return;
  }
  if (interactive && !hasOrders) {
    gantt.innerHTML = "<p class=\"helper-text\">Aucun planning validé n'est lancé. Sélectionne une semaine et clique sur Lancer.</p>";
    return;
  }

  // recalage à la volée pour le jeu de données rendu
  LINES.forEach(line => recalibrateLine(line, { orders, plannedStops }));

  gantt.innerHTML = "";

  const axis = document.createElement("div");
  axis.className = "gantt-axis";
  const axisLabel = document.createElement("div");
  axisLabel.className = "gantt-axis-label";
  axisLabel.textContent = "Jours / heures";
  const axisTimeline = document.createElement("div");
  axisTimeline.className = "gantt-axis-timeline";

  let minutesCursor = 0;
  const dayDurations = [24, 24, 24, 24, 24, 12];
  dayDurations.forEach((hours, idx) => {
    const minutes = hours * 60;
    const leftPct = (minutesCursor / WEEK_TOTAL_MINUTES) * 100;
    const widthPct = (minutes / WEEK_TOTAL_MINUTES) * 100;
    const band = document.createElement("div");
    band.className = "gantt-axis-day";
    band.style.left = `${leftPct}%`;
    band.style.width = `${widthPct}%`;
    band.textContent = PLANNING_DAY_LABELS[idx] || "";
    axisTimeline.appendChild(band);

    for (let h = 0; h <= hours; h += 6) {
      const tick = document.createElement("div");
      tick.className = "gantt-axis-tick";
      const minuteOffset = minutesCursor + h * 60;
      tick.style.left = `${(minuteOffset / WEEK_TOTAL_MINUTES) * 100}%`;
      tick.textContent = `${String(h % 24).padStart(2, "0")}h`;
      axisTimeline.appendChild(tick);
    }

    minutesCursor += minutes;
  });

  axis.appendChild(axisLabel);
  axis.appendChild(axisTimeline);
  gantt.appendChild(axis);

  LINES.forEach(line => {
    const row = document.createElement("div");
    row.className = "gantt-row";

    const label = document.createElement("div");
    label.className = "gantt-line-label";
    label.textContent = line;

    const timeline = document.createElement("div");
    timeline.className = "gantt-line-timeline";
    timeline.dataset.line = line;
    if (interactive) {
      timeline.addEventListener("dragover", e => e.preventDefault());
      timeline.addEventListener("drop", handlePlanningDrop);
    }

    for (let i = 1; i <= 5; i++) {
      const split = document.createElement("div");
      split.className = "gantt-day-split";
      split.style.left = `${(i * 24 * 60) / WEEK_TOTAL_MINUTES * 100}%`;
      timeline.appendChild(split);
    }

    const stops = getAllStopsForLine(line, plannedStops).map(s => ({ ...s, kind: "stop" }));
    const ordersForLine = orders.filter(o => o.line === line);
    const orderSegments = [];

    ordersForLine.forEach(o => {
      const segs = (o.segments && o.segments.length)
        ? o.segments
        : [{ ...o, start: o.start, end: o.end }];
      segs.forEach((seg, idx) => {
        orderSegments.push({
          ...seg,
          kind: "of",
          parentId: o.id,
          segmentIndex: idx,
          status: seg.status || o.status || "planned",
          blockedReason: seg.blockedReason || o.blockedReason,
          code: seg.code || o.code,
          label: seg.label || o.label,
          quantity: seg.quantity || o.quantity,
        });
      });
    });

    const events = [...orderSegments, ...stops];

    events.forEach(ev => {
      const s = new Date(ev.start);
      const e = new Date(ev.end || ev.start);
      if (e < start || s > end) return;
      const durationMin = Math.max(5, (e - s) / 60000);
      const offsetMin = Math.max(0, (s - start) / 60000);
      const leftPct = (offsetMin / WEEK_TOTAL_MINUTES) * 100;
      const widthPct = Math.min(100 - leftPct, (durationMin / WEEK_TOTAL_MINUTES) * 100);

      const block = document.createElement("div");
      block.className = `gantt-block ${ev.kind === "stop" ? "stop" : (ev.status || "planned")}`;
      block.style.left = `${leftPct}%`;
      block.style.width = `${Math.max(widthPct, 0.5)}%`;
      block.draggable = interactive && ev.kind !== "stop";
      block.dataset.id = ev.id;
      block.dataset.line = line;
      block.dataset.kind = ev.kind;
      if (interactive) {
        block.addEventListener("click", () => {
          if (ev.kind === "stop") return;
          openPlanningEditor(ev.id);
        });
        block.addEventListener("dragstart", e => {
          e.dataTransfer.setData("text/planning-order", ev.id);
        });
      }

      const title = document.createElement("div");
      title.className = "title";
      title.textContent = ev.kind === "stop"
        ? ev.type
        : [ev.code || "OF", ev.label].filter(Boolean).join(" — ");

      const meta = document.createElement("div");
      meta.className = "meta";
      const labelLine = ev.kind === "stop"
        ? `${formatPlanningDay(ev.start)} ${formatTimeShort(ev.start)} → ${formatTimeShort(ev.end)} • ${(ev.duration || 0)} min`
        : `${formatPlanningDay(ev.start)} ${formatTimeShort(ev.start)} → ${formatPlanningDay(ev.end)} ${formatTimeShort(ev.end)} • ${ev.quantity || ""} colis`;
      meta.textContent = labelLine;

      block.appendChild(title);
      block.appendChild(meta);
      if (ev.kind !== "stop" && ev.label) {
        const desc = document.createElement("div");
        desc.className = "meta meta-secondary";
        desc.textContent = ev.label;
        block.appendChild(desc);
      }
      if (ev.kind === "stop" && ev.comment) {
        const desc = document.createElement("div");
        desc.className = "meta meta-secondary";
        desc.textContent = ev.comment;
        block.appendChild(desc);
      }
      if (ev.blockedReason) {
        const reason = document.createElement("div");
        reason.className = "meta";
        reason.textContent = `Blocage : ${ev.blockedReason}`;
        block.appendChild(reason);
      }
      timeline.appendChild(block);
    });

    row.appendChild(label);
    row.appendChild(timeline);
    gantt.appendChild(row);
  });
}

function formatTimeShort(dateLike) {
  const d = new Date(dateLike);
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

function formatPlanningDay(dateLike) {
  const base = getPlanningWeekStartDate();
  const diff = Math.floor((new Date(dateLike) - base) / 86400000);
  if (diff < 0) return "";
  return PLANNING_DAY_LABELS[Math.min(diff, PLANNING_DAY_LABELS.length - 1)] || "";
}

function updatePlanningStatusFromProduction(line, rec) {
  if (!state.planning || !state.planning.activeOrders?.length) return;
  const baseDate = rec.dateTime ? new Date(rec.dateTime) : new Date();
  const makeDate = (t) => {
    if (!t) return null;
    const d = new Date(baseDate);
    const [h, m] = t.split(":").map(Number);
    d.setHours(h || 0, m || 0, 0, 0);
    return d;
  };
  const startD = makeDate(rec.start);
  const endD = makeDate(rec.end);

  const orders = state.planning.activeOrders.filter(o => o.line === line);
  const matchStart = orders.find(o => startD && new Date(o.start) <= startD && startD <= new Date(o.end));
  const matchEnd = orders.find(o => endD && new Date(o.start) <= endD && endD <= new Date(o.end));
  if (matchStart) {
    matchStart.status = "running";
    matchStart.produced = (matchStart.produced || 0) + (Number(rec.quantity) || 0);
  }
  if (matchEnd) {
    matchEnd.status = "done";
  }
  saveState();
}

function handlePlanningDrop(e) {
  e.preventDefault();
  const id = e.dataTransfer.getData("text/planning-order");
  if (!id) return;
  const line = e.currentTarget.dataset.line;
  const rect = e.currentTarget.getBoundingClientRect();
  const ratio = (e.clientX - rect.left) / rect.width;
  const minutes = Math.max(0, Math.min(WEEK_TOTAL_MINUTES, ratio * WEEK_TOTAL_MINUTES));
  const startDate = new Date(getPlanningWeekStartDate().getTime() + minutes * 60000);

  const order = (state.planning.activeOrders || []).find(o => o.id === id);
  if (!order) return;
  order.line = line;
  order.start = startDate.toISOString();
  const duration = computeOrderDurationMinutes(order);
  order.end = new Date(startDate.getTime() + duration * 60000).toISOString();
  recalibrateLine(line, { orders: state.planning.activeOrders, plannedStops: state.planning.activeArrets });
  saveState();
  refreshPlanningGantt();
}

function setPlanningTab(tab) {
  const panes = document.querySelectorAll("#section-planning .planning-pane");
  const tabs = document.querySelectorAll("#section-planning .planning-tab-btn");
  tabs.forEach(btn => btn.classList.toggle("active", btn.dataset.tab === tab));
  panes.forEach(pane => pane.classList.toggle("hidden", pane.id !== `planning-tab-${tab}`));
  if (tab === "run") {
    refreshPlanningGantt();
    refreshPlanningDelays();
  }
}

function openPlanningStopPopover() {
  const pop = document.getElementById("planningStopPopover");
  if (pop) pop.classList.remove("hidden");
}

function closePlanningStopPopover() {
  const pop = document.getElementById("planningStopPopover");
  if (pop) pop.classList.add("hidden");
}

function openPlanningEditor(id) {
  const modal = document.getElementById("planningBlockEditor");
  const of = (state.planning.activeOrders || []).find(o => o.id === id);
  if (!modal || !of) return;
  planningEditId = id;
  document.getElementById("planningEditQty").value = of.quantity || 0;
  document.getElementById("planningEditStart").value = of.start ? of.start.slice(0, 16) : "";
  document.getElementById("planningEditEnd").value = of.end ? of.end.slice(0, 16) : "";
  document.getElementById("planningEditStatus").value = of.status || "planned";
  document.getElementById("planningEditBlockReason").value = of.blockedReason || "";
  modal.hidden = false;
  modal.classList.remove("hidden");
}

function closePlanningEditor() {
  const modal = document.getElementById("planningBlockEditor");
  if (modal) {
    modal.hidden = true;
    modal.classList.add("hidden");
  }
  planningEditId = null;
}

function bindPlanningEditor() {
  const saveBtn = document.getElementById("planningEditSaveBtn");
  const cancelBtn = document.getElementById("planningEditCancelBtn");
  const modal = document.getElementById("planningBlockEditor");
  const modalContent = modal?.querySelector(".modal-content");
  closePlanningEditor();

  cancelBtn?.addEventListener("click", closePlanningEditor);
  saveBtn?.addEventListener("click", () => {
    if (!planningEditId) return;
    const of = (state.planning.activeOrders || []).find(o => o.id === planningEditId);
    if (!of) return;
    of.quantity = Number(document.getElementById("planningEditQty")?.value) || of.quantity;
    const newStart = document.getElementById("planningEditStart")?.value;
    const newEnd = document.getElementById("planningEditEnd")?.value;
    if (newStart) of.start = new Date(newStart).toISOString();
    if (newEnd) of.end = new Date(newEnd).toISOString();
    of.status = document.getElementById("planningEditStatus")?.value || of.status;
    of.blockedReason = document.getElementById("planningEditBlockReason")?.value || "";
    recalibrateLine(of.line, { orders: state.planning.activeOrders, plannedStops: state.planning.activeArrets });
    saveState();
    refreshPlanningGantt();
    closePlanningEditor();
  });

  modal?.addEventListener("click", (e) => {
    if (e.target === modal) closePlanningEditor();
  });

  modalContent?.addEventListener("click", e => e.stopPropagation());
}

function bindPlanning() {
  ensurePlanningDefaults();
  const knownWeeks = (state.planning.savedPlans || []).map(p => `${p.week}`);
  if (!knownWeeks.includes(`${state.planning.selectedPlanWeek || ""}`)) {
    state.planning.selectedPlanWeek = "";
    state.planning.activeOrders = [];
    state.planning.activeArrets = [];
    saveState();
  }
  refreshPlanningLineSelectors();
  renderPlanningCadences();
  const weekNumberInput = document.getElementById("planningWeekNumber");
  const weekStartInput = document.getElementById("planningWeekStart");
  if (weekNumberInput) weekNumberInput.value = state.planning.weekNumber || "";
  if (weekStartInput) weekStartInput.value = state.planning.weekStart || formatDateInput(getMonday());

  document.querySelectorAll("#section-planning .planning-tab-btn")?.forEach(btn => {
    btn.addEventListener("click", () => setPlanningTab(btn.dataset.tab));
  });
  setPlanningTab("articles");

  document.getElementById("planningArticlePickerBtn")?.addEventListener("click", openPlanningArticlePicker);
  document.getElementById("planningArticlePickerClose")?.addEventListener("click", closePlanningArticlePicker);
  document.getElementById("planningArticlePicker")?.addEventListener("click", e => {
    if (e.target.id === "planningArticlePicker") closePlanningArticlePicker();
  });

  const codeInput = document.getElementById("planningOFCode");
  codeInput?.addEventListener("change", () => autofillOFFieldsFromCode(codeInput.value));
  codeInput?.addEventListener("blur", () => autofillOFFieldsFromCode(codeInput.value));

  weekNumberInput?.addEventListener("input", () => {
    state.planning.weekNumber = weekNumberInput.value;
    saveState();
  });
  weekStartInput?.addEventListener("change", () => {
    state.planning.weekStart = weekStartInput.value;
    saveState();
    refreshPlanningGantt();
  });

  document.getElementById("planningArticleAddBtn")?.addEventListener("click", addPlanningCadence);
  document.getElementById("planningStopSaveBtn")?.addEventListener("click", addPlanningStop);
  document.getElementById("planningOpenStopPopover")?.addEventListener("click", openPlanningStopPopover);
  document.getElementById("planningStopCloseBtn")?.addEventListener("click", closePlanningStopPopover);
  document.getElementById("planningOFAddBtn")?.addEventListener("click", addPlanningOF);
  const stopPopover = document.getElementById("planningStopPopover");
  const stopContent = stopPopover?.querySelector(".popover-content");
  stopPopover?.addEventListener("click", e => { if (e.target === stopPopover) closePlanningStopPopover(); });
  stopContent?.addEventListener("click", e => e.stopPropagation());

  document.getElementById("planningRebuildBtn")?.addEventListener("click", () => {
    LINES.forEach(recalibrateLine);
    saveState();
    refreshPlanningGantt();
  });

  document.getElementById("planningScrollTopBtn")?.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  document.getElementById("planningShowDelaysBtn")?.addEventListener("click", refreshPlanningDelays);

  document.getElementById("planningValidateBtn")?.addEventListener("click", savePlanningSnapshot);
  document.getElementById("planningLaunchBtn")?.addEventListener("click", () => launchPlanningSnapshot());
  document.getElementById("planningEditLoadBtn")?.addEventListener("click", loadPlanningForEditing);

  bindPlanningEditor();
  refreshSavedPlanningList(true);
  refreshPlanningGantt();
  refreshPlanningDelays();
}

function savePlanningSnapshot() {
  let week = document.getElementById("planningWeekNumber")?.value || state.planning.weekNumber || "";
  const start = document.getElementById("planningWeekStart")?.value || state.planning.weekStart;
  if (!week && start) {
    const monday = new Date(start);
    week = getWeekNumber(monday);
  }
  if (!week || !start) {
    alert("Renseigne la semaine et le lundi de référence avant de valider le planning.");
    return;
  }
  if (!state.planning.savedPlans) state.planning.savedPlans = [];
  const snapshot = {
    week,
    start,
    savedAt: new Date().toISOString(),
    orders: JSON.parse(JSON.stringify(state.planning.orders || [])),
    arretsPlanifies: JSON.parse(JSON.stringify(state.planning.arretsPlanifies || [])),
    cadences: JSON.parse(JSON.stringify(state.planning.cadences || [])),
  };

  // Recharge le stockage pour éviter les divergences puis merge/écrase la semaine en cours
  const stored = getMergedPlanningSnapshots();
  const merged = [...stored];
  const existingIdx = merged.findIndex(p => `${p.week}` === `${week}`);
  if (existingIdx >= 0) merged[existingIdx] = snapshot;
  else merged.push(snapshot);

  planningSnapshots = merged;
  state.planning.savedPlans = merged;
  state.planning.weekNumber = week;
  state.planning.weekStart = start;
  state.planning.selectedPlanWeek = "";
  state.planning.activeOrders = [];
  state.planning.activeArrets = [];
  savePlanningSnapshots(merged);
  saveState();
  refreshSavedPlanningList(true);
  const editSelect = document.getElementById("planningEditSelect");
  const launchSelect = document.getElementById("planningLaunchSelect");
  if (editSelect) editSelect.value = `${week}`;
  if (launchSelect) launchSelect.value = `${week}`;

  clearPlanningDraft();
  LINES.forEach(recalibrateLine);
  saveState();
  refreshPlanningGantt();
  refreshPlanningDelays();
}

function clearPlanningDraft() {
  state.planning.orders = [];
  state.planning.arretsPlanifies = [];
  resetPlanningOFForm();
  resetPlanningPreview();
  const preview = document.getElementById("planningPreview");
  if (preview) preview.classList.add("is-empty");
}

function refreshSavedPlanningList(forceReload = false) {
  const container = document.getElementById("planningSavedList");
  const editSelect = document.getElementById("planningEditSelect");
  const launchSelect = document.getElementById("planningLaunchSelect");
  if (container) container.innerHTML = "";

  // Recharge depuis le stockage et la mémoire à chaque rafraîchissement pour éviter les listes vides
  if (forceReload || !state.planning.savedPlans.length) {
    state.planning.savedPlans = getMergedPlanningSnapshots();
    planningSnapshots = [...state.planning.savedPlans];
  }

  if (!state.planning.savedPlans.length) {
    if (container) container.textContent = "Aucun planning validé pour l'instant.";
    if (editSelect) editSelect.innerHTML = "<option value=\"\">Aucun planning</option>";
    if (launchSelect) launchSelect.innerHTML = "<option value=\"\">Aucun planning</option>";
    return;
  }

  const sorted = [...state.planning.savedPlans].sort((a, b) => (b.week || 0) - (a.week || 0));
  if (container) {
    const rows = sorted
      .map(p => `<div class="helper-text">Semaine ${p.week} – validé le ${formatDateTime(p.savedAt)}</div>`)
      .join("");
    container.innerHTML = rows;
  }

  const options = sorted
    .map(p => `<option value="${p.week}">Semaine ${p.week} (${formatDateInput(p.start || "") || ""})</option>`)
    .join("");
  if (editSelect) editSelect.innerHTML = `<option value="">Choisir…</option>${options}`;
  if (launchSelect) {
    launchSelect.innerHTML = `<option value="">Choisir…</option>${options}`;
    if (state.planning.selectedPlanWeek) {
      launchSelect.value = `${state.planning.selectedPlanWeek}`;
    }
  }
}

function loadPlanningForEditing() {
  const targetWeek = document.getElementById("planningEditSelect")?.value || "";
  if (!targetWeek) {
    alert("Sélectionne un planning validé à charger.");
    return;
  }
  const snap = state.planning.savedPlans.find(p => `${p.week}` === `${targetWeek}`);
  if (!snap) return;
  state.planning.weekNumber = snap.week;
  state.planning.weekStart = snap.start;
  state.planning.orders = JSON.parse(JSON.stringify(snap.orders || []));
  state.planning.arretsPlanifies = JSON.parse(JSON.stringify(snap.arretsPlanifies || []));
  saveState();
  const weekNumberInput = document.getElementById("planningWeekNumber");
  const weekStartInput = document.getElementById("planningWeekStart");
  if (weekNumberInput) weekNumberInput.value = state.planning.weekNumber;
  if (weekStartInput) weekStartInput.value = state.planning.weekStart;
  renderPlanningCadences();
  updatePlanningOFPrereq();
  LINES.forEach(recalibrateLine);
  refreshPlanningGantt();
  refreshPlanningDelays();
  setPlanningTab("build");
}

function launchPlanningSnapshot(targetWeekOverride) {
  const targetWeek = targetWeekOverride || document.getElementById("planningLaunchSelect")?.value || state.planning.weekNumber;
  if (!targetWeek) {
    alert("Choisis un planning validé dans la liste.");
    return;
  }
  const snap = state.planning.savedPlans.find(p => `${p.week}` === `${targetWeek}`);
  if (!snap) return;
  state.planning.weekNumber = snap.week;
  state.planning.weekStart = snap.start;
  state.planning.selectedPlanWeek = `${snap.week}`;
  state.planning.activeOrders = JSON.parse(JSON.stringify(snap.orders || []));
  state.planning.activeArrets = JSON.parse(JSON.stringify(snap.arretsPlanifies || []));
  saveState();
  const weekNumberInput = document.getElementById("planningWeekNumber");
  const weekStartInput = document.getElementById("planningWeekStart");
  if (weekNumberInput) weekNumberInput.value = state.planning.weekNumber;
  if (weekStartInput) weekStartInput.value = state.planning.weekStart;
  renderPlanningCadences();
  updatePlanningOFPrereq();
  LINES.forEach(line => recalibrateLine(line, { orders: state.planning.activeOrders, plannedStops: state.planning.activeArrets }));
  refreshPlanningGantt();
  refreshPlanningDelays();
  setPlanningTab("run");
}

function refreshPlanningDelays() {
  const container = document.getElementById("planningDelays");
  if (!container) return;
  const weekStart = getPlanningWeekStartDate();
  const weekEnd = getPlanningWeekEndDate();
  const now = new Date();
  container.innerHTML = "";

  const activeOrders = state.planning.activeOrders || [];
  if (!activeOrders.length) {
    container.innerHTML = "<p class=\"helper-text\">Lance un planning validé pour calculer l'avance/retard.</p>";
    return;
  }

  LINES.forEach(line => {
    const orders = activeOrders.filter(o => o.line === line);
    let expectedQty = 0;
    let expectedWeight = 0;
    orders.forEach(of => {
      const start = new Date(of.start);
      const end = new Date(of.end);
      const duration = Math.max(1, (end - start) / 60000);
      const qty = Number(of.quantity) || 0;
      const weight = (Number(of.poids) || 0) * qty;
      if (now <= start) return;
      const ref = Math.min(now.getTime(), end.getTime());
      const elapsed = Math.max(0, ref - start.getTime());
      const ratio = Math.min(1, elapsed / (duration * 60000));
      expectedQty += qty * ratio;
      expectedWeight += weight * ratio;
    });

    const prods = state.production[line] || [];
    let actualQty = 0;
    let actualWeight = 0;
    prods.forEach(p => {
      const when = new Date(p.dateTime);
      if (when < weekStart || when > weekEnd) return;
      const qty = Number(p.quantity) || 0;
      actualQty += qty;
      const art = getCadenceForArticle(p.article || "", line);
      if (art && art.poids) actualWeight += qty * art.poids;
    });

    const deltaQty = actualQty - expectedQty;
    const deltaWeight = actualWeight - expectedWeight;

    const card = document.createElement("div");
    card.className = "delay-card";
    card.innerHTML = `
      <h4>${line}</h4>
      <div>Avance/retard (colis) : <span class="${deltaQty >= 0 ? "delta-positive" : "delta-negative"}">${deltaQty.toFixed(1)}</span></div>
      <div>Avance/retard (kg) : <span class="${deltaWeight >= 0 ? "delta-positive" : "delta-negative"}">${deltaWeight.toFixed(1)}</span></div>
    `;
    container.appendChild(card);
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

document.addEventListener("DOMContentLoaded", async () => {
  await ensureDefaultManagerPassword();
  loadState();
  loadPlanningSnapshots();
  loadArchives();
  initHeaderDate();
  initEquipeSelector();
  initNav();
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
  initSettingsPanel();
  initManagerUnlockModal();
  bindPlanning();

  updateOrientationLayout();
  window.addEventListener("resize", updateOrientationLayout);
  window.addEventListener("orientationchange", updateOrientationLayout);

  const initialSection = state.currentSection || "atelier";
  if (initialSection === "manager" && isManagerLocked()) {
    showSection("atelier");
    const unlocked = await promptManagerUnlock();
    if (unlocked) showSection("manager");
  } else {
    showSection(initialSection);
  }
});
