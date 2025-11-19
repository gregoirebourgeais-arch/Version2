"/******************************
 *   PARTIE 1 / 4 – CONFIG
 ******************************/

// === CONFIG & ÉTAT GLOBAL ===

const LINES = [
  \"Râpé\",
  \"T2\",
  \"RT\",
  \"OMORI\",
  \"T1\",
  \"Sticks\",
  \"Emballage\",
  \"Dés\",
  \"Filets\",
  \"Prédécoupé\",
];

// Stockage version
const STORAGE_KEY = \"atelier_ppnc_state_v2\";  
const ARCHIVES_KEY = \"atelier_ppnc_archives_v2\";

let archives = []; // [{ id, label, savedAt, equipe, week, quantieme, state }]

// Sous-lignes pour Râpé
const ARRET_SUBLINES = {
  \"Râpé\": [\"R1\", \"R2\", \"R1/R2\"],
};

// Machines déclarées
const ARRET_MACHINES = {
  \"Râpé\": [
    \"Cubeuse\", \"Cheesix\", \"Liftvrac\", \"Associative\", \"Ensacheuse\",
    \"Encaisseuse\", \"Smartdate\", \"Bizerba\", \"DPM\", \"Scotcheuse\",
    \"Markem\", \"Ascenseur\",
  ],
  \"T2\": [
    \"Selvex\", \"Trieuse\", \"Robots\", \"Tiromat\", \"Vision\",
    \"Convoyeur\", \"DPM\", \"Bizerba\", \"Suremballage\", \"Markem\",
    \"Scotcheuse\", \"Balance cartons\", \"Formeuse caisse\", \"Ascenseur\",
  ],
  \"OMORI\": [
    \"BFR\", \"Accumulateur\", \"OMORI\", \"Videojet\", \"DPM\",
    \"Encaisseuse\", \"Balance cartons\", \"Ascenseur\",
  ],
  \"Emballage\": [
    \"Brinkman\", \"Encaisseuse\", \"Bizerba\", \"Palettiseur\",
    \"Paraffineuse\", \"Râpé\", \"Ecroûtage\", \"Alpma\",
  ],
  \"T1\": [
    \"Slicer\", \"AES\", \"Tiromat\", \"Préhenseur\",
    \"DPM\", \"Encaisseuse T1\", \"Encaisseuse David\",
    \"Balance cartons\", \"Ascenseur\",
  ],
  \"Dés\": [\"Cheesix\", \"Meca 2002\", \"DPM\", \"Bizerba\", \"Scotcheuse\"],
  \"Filets\": [\"Lieuse\", \"C-Pack\", \"Etiqueteuse\", \"Scotcheuse\"],
  \"Prédécoupé\": [\"DPM\", \"Selvex\", \"Bizerba\", \"Quartivac\", \"Scotcheuse\"],
  \"RT\": [\"Autre\"],
  \"Sticks\": [\"Autre\"],
};

// === ÉTAT GLOBAL ===
// 👇 Version \"persistance par ligne\" (Option 1)
let state = {
  currentSection: \"atelier\",
  currentLine: LINES[0],
  currentEquipe: \"M\", // Ne dépend plus de l'heure
  production: {},     // { line: [records] }
  arrets: [],
  organisation: [],
  personnel: [],

  // Champs en cours (pendant saisie non enregistrée)
  formDraft: {},      // { line: { start, end, qty, remaining, cadenceMan, arret, comment, article } }
};

// Références pour page Arrêts
let arretLineEl = null;
let arretSousZoneEl = null;
let arretSousZoneRowEl = null;
let arretMachineEl = null;

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
  const dd = String(d.getDate()).padStart(2, \"0\");
  const mm = String(d.getMonth() + 1).padStart(2, \"0\");
  const hh = String(d.getHours()).padStart(2, \"0\");
  const mi = String(d.getMinutes()).padStart(2, \"0\");
  return `${dd}/${mm}/${d.getFullYear()} ${hh}h${mi}`;
}

function formatTimeRemaining(min) {
  if (!min || min <= 0) return \"-\";
  const h = Math.floor(min / 60);
  const m = Math.round(min % 60);
  if (h === 0) return `${m} min`;
  if (m === 0) return `${h} h`;
  return `${h} h ${m} min`;
}

function parseTimeToMinutes(t) {
  if (!t) return null;
  const [h, m] = t.split(\":\").map(Number);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
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

    // Reconstruction propre
    const base = {
      currentSection: \"atelier\",
      currentLine: LINES[0],
      currentEquipe: \"M\",
      production: {},
      arrets: [],
      organisation: [],
      personnel: [],
      formDraft: {},
    };

    state = Object.assign(base, parsed);

    LINES.forEach(l => {
      if (!state.production[l]) state.production[l] = [];
      if (!state.formDraft[l]) state.formDraft[l] = {};
    });

  } catch (e) {
    console.error(\"loadState error\", e);
    LINES.forEach(l => state.production[l] = []);
  }
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    console.error(\"saveState error\", e);
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
 *   PARTIE 1 / 4 – HEADER
 ******************************/

function initHeaderDate() {
  const el = document.getElementById(\"header-datetime\");
  if (!el) return;

  function update() {
    const now = getNow();
    const week = getWeekNumber(now);
    const q = getQuantieme(now);

    const d = String(now.getDate()).padStart(2, \"0\");
    const m = String(now.getMonth() + 1).padStart(2, \"0\");
    const hh = String(now.getHours()).padStart(2, \"0\");
    const mm = String(now.getMinutes()).padStart(2, \"0\");

    el.textContent =
      `Quantième ${q} • ${d}/${m}/${now.getFullYear()} • ` +
      `S${week} • ${hh}:${mm} • Équipe ${state.currentEquipe}`;
  }

  update();
  setInterval(update, 30000);
}

/********************************************
 *   PARTIE 2 / 4 – NAVIGATION + PRODUCTION
 ********************************************/

// === NAVIGATION ===

function showSection(section) {
  state.currentSection = section;
  saveState();

  document.querySelectorAll(\".section\").forEach(sec =>
    sec.classList.toggle(\"visible\", sec.id === `section-${section}`)
  );

  document.querySelectorAll(\".nav-btn\").forEach(btn =>
    btn.classList.toggle(\"active\", btn.dataset.section === section)
  );

  if (section === \"atelier\") refreshAtelierView();
  else if (section === \"production\") refreshProductionView();
  else if (section === \"arrets\") refreshArretsView();
  else if (section === \"organisation\") refreshOrganisationView();
  else if (section === \"personnel\") refreshPersonnelView();
}

function initNav() {
  document.querySelectorAll(\".nav-btn\").forEach(btn => {
    btn.addEventListener(\"click\", () => {
      showSection(btn.dataset.section);
    });
  });
}

// === LIGNES (Production) ===

function initLinesSidebar() {
  const container = document.getElementById(\"linesList\");
  if (!container) return;
  
  container.innerHTML = \"\";

  LINES.forEach(line => {
    const btn = document.createElement(\"button\");
    btn.className = \"line-btn\";
    btn.textContent = line;
    btn.dataset.line = line;

    btn.addEventListener(\"click\", () => selectLine(line, true));

    container.appendChild(btn);
  });

  selectLine(state.currentLine || LINES[0], false);
}

function selectLine(line, scrollToForm) {
  state.currentLine = line;
  saveState();

  const titleEl = document.getElementById(\"currentLineTitle\");
  if (titleEl) titleEl.textContent = `Ligne ${line}`;

  document.querySelectorAll(\".line-btn\").forEach(b =>
    b.classList.toggle(\"active\", b.dataset.line === line)
  );

  refreshProductionForm();
  refreshProductionHistoryTable();

  if (scrollToForm) {
    const card = document.querySelector(\"#section-production .card\");
    if (card) card.scrollIntoView({ behavior: \"smooth\", block: \"start\" });
  }
}

// === PRODUCTION : OUTILS ===

function getCurrentLineRecords() {
  return state.production[state.currentLine] || [];
}

function computeCadenceFromInputs() {
  const s = document.getElementById(\"prodStartTime\")?.value;
  const e = document.getElementById(\"prodEndTime\")?.value;
  const q = Number(document.getElementById(\"prodQuantity\")?.value) || 0;

  const sMin = parseTimeToMinutes(s);
  const eMin = parseTimeToMinutes(e);
  if (sMin == null || eMin == null || q <= 0) return null;

  let dur = eMin - sMin;
  if (dur <= 0) dur += 24 * 60;

  const hours = dur / 60;
  return hours > 0 ? q / hours : null;
}

// Cadence de référence = dernière cadence valide OU cadence manuelle
function computeRefCadenceForLine(line) {
  const manualEl = document.getElementById(\"prodCadenceManual\");
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
  const el = document.getElementById(\"prodCadenceDisplay\");
  if (el) el.textContent = cad ? cad.toFixed(2) : \"-\";
}

function updateRemainingTimeDisplay() {
  const remainingEl = document.getElementById(\"prodRemaining\");
  const remaining = Number(remainingEl?.value) || 0;
  const line = state.currentLine;
  const cadRef = computeRefCadenceForLine(line);
  const el = document.getElementById(\"prodRemainingTimeDisplay\");

  if (!el) return;

  if (!remaining || !cadRef || cadRef <= 0) {
    el.textContent = \"-\";
    return;
  }

  const min = (remaining / cadRef) * 60;
  el.textContent = formatTimeRemaining(min);
}

// === PERSISTENCE AUTO DU FORMULAIRE ===

function saveDraft() {
  const L = state.currentLine;

  state.formDraft[L] = {
    start: document.getElementById(\"prodStartTime\")?.value || \"\",
    end: document.getElementById(\"prodEndTime\")?.value || \"\",
    qty: document.getElementById(\"prodQuantity\")?.value || \"\",
    remaining: document.getElementById(\"prodRemaining\")?.value || \"\",
    cadenceMan: document.getElementById(\"prodCadenceManual\")?.value || \"\",
    arret: document.getElementById(\"prodArretMinutes\")?.value || \"\",
    comment: document.getElementById(\"prodComment\")?.value || \"\",
    article: document.getElementById(\"prodCodeArticle\")?.value || \"\",
  };

  saveState();
}

function loadDraft() {
  const L = state.currentLine;
  const d = state.formDraft[L] || {};

  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val || \"\";
  };

  setVal(\"prodStartTime\", d.start);
  setVal(\"prodEndTime\", d.end);
  setVal(\"prodQuantity\", d.qty);
  setVal(\"prodRemaining\", d.remaining);
  setVal(\"prodCadenceManual\", d.cadenceMan);
  setVal(\"prodArretMinutes\", d.arret);
  setVal(\"prodComment\", d.comment);
  setVal(\"prodCodeArticle\", d.article);

  updateCadenceDisplay();
  updateRemainingTimeDisplay();
}

// === REFRESH FORMULAIRE PRODUCTION ===

function refreshProductionForm() {
  loadDraft(); // ← persistance OK
}

// === TABLE HISTORIQUE DES PRODUCTIONS ===

function refreshProductionHistoryTable() {
  const table = document.getElementById(\"prodHistoryTable\");
  if (!table) return;
  
  const tbody = table.querySelector(\"tbody\");
  if (!tbody) return;
  
  tbody.innerHTML = \"\";

  getCurrentLineRecords().forEach((rec, idx) => {
    const tr = document.createElement(\"tr\");

    tr.innerHTML = `
      <td>${rec.dateTime}</td>
      <td>${rec.equipe}</td>
      <td>${rec.start || \"-\"}</td>
      <td>${rec.end || \"-\"}</td>
      <td>${rec.quantity}</td>
      <td>${rec.arret || 0}</td>
      <td>${rec.cadence ? rec.cadence.toFixed(2) : \"-\"}</td>
      <td>${rec.remainingTime || \"-\"}</td>
      <td>${rec.comment || \"\"}</td>
      <td>${rec.article || \"\"}</td>
      <td><button class=\"secondary-btn\" data-idx=\"${idx}\">✕</button></td>
    `;

    tbody.appendChild(tr);
  });

  tbody.querySelectorAll(\"button[data-idx]\").forEach(btn => {
    btn.addEventListener(\"click\", () => {
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
    \"prodStartTime\",
    \"prodEndTime\",
    \"prodQuantity\",
    \"prodRemaining\",
    \"prodCadenceManual\",
    \"prodArretMinutes\",
    \"prodComment\",
    \"prodCodeArticle\"
  ];

  formIds.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    
    el.addEventListener(\"input\", () => {
      if (id !== \"prodComment\") updateCadenceDisplay();
      if (id !== \"prodComment\") updateRemainingTimeDisplay();
      saveDraft();
    });
  });

  const saveBtn = document.getElementById(\"prodSaveBtn\");
  if (saveBtn) {
    saveBtn.addEventListener(\"click\", () => {
      const now = getNow();
      const L = state.currentLine;
      const equipe = state.currentEquipe;

      // → rec complet
      const rec = {
        dateTime: formatDateTime(now),
        equipe,
        start: document.getElementById(\"prodStartTime\")?.value || \"\",
        end: document.getElementById(\"prodEndTime\")?.value || \"\",
        quantity: Number(document.getElementById(\"prodQuantity\")?.value) || 0,
        remainingTime:
          document.getElementById(\"prodRemainingTimeDisplay\")?.textContent || \"-\",
        arret: Number(document.getElementById(\"prodArretMinutes\")?.value) || 0,
        cadence:
          Number(document.getElementById(\"prodCadenceManual\")?.value) ||
          computeCadenceFromInputs() ||
          null,
        comment: document.getElementById(\"prodComment\")?.value || \"\",
        article: document.getElementById(\"prodCodeArticle\")?.value || \"\",
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

  const undoBtn = document.getElementById(\"prodUndoBtn\");
  if (undoBtn) {
    undoBtn.addEventListener(\"click\", () => {
      const arr = getCurrentLineRecords();
      if (arr.length) arr.pop();
      saveState();
      refreshProductionHistoryTable();
      refreshAtelierView();
    });
  }
}

// === SCROLL HORIZONTAL PRODUCTION (fix) ===

function refreshProductionView() {
  refreshProductionForm();
  refreshProductionHistoryTable();

  const table = document.getElementById(\"prodHistoryTable\");
  if (table && table.parentElement) {
    table.parentElement.style.overflowX = \"auto\"; // ← scroll latéral ACTIVÉ
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
    if (line === \"Râpé\") {
      arretSousZoneRowEl.style.display = \"\";
      arretSousZoneEl.innerHTML = \"\";
      ARRET_SUBLINES[\"Râpé\"].forEach(s => {
        const opt = document.createElement(\"option\");
        opt.value = s;
        opt.textContent = s;
        arretSousZoneEl.appendChild(opt);
      });
    } else {
      arretSousZoneRowEl.style.display = \"none\";
      arretSousZoneEl.innerHTML = \"\";
    }
  }

  // Machines
  arretMachineEl.innerHTML = \"\";
  (ARRET_MACHINES[line] || [\"Autre\"]).forEach(m => {
    const opt = document.createElement(\"option\");
    opt.value = m;
    opt.textContent = m;
    arretMachineEl.appendChild(opt);
  });
}

// === ARRETS INIT ===

function initArretsForm() {
  arretLineEl = document.getElementById(\"arretLine\");
  arretSousZoneEl = document.getElementById(\"arretSousZone\");
  arretMachineEl = document.getElementById(\"arretMachine\");
  
  if (!arretLineEl || !arretSousZoneEl || !arretMachineEl) return;
  
  arretSousZoneRowEl = arretSousZoneEl.closest(\".form-row\");

  // Lignes dans la liste
  arretLineEl.innerHTML = \"\";
  LINES.forEach(l => {
    const opt = document.createElement(\"option\");
    opt.value = l;
    opt.textContent = l;
    arretLineEl.appendChild(opt);
  });

  arretLineEl.addEventListener(\"change\", updateArretControlsForLine);
  updateArretControlsForLine();

  // Enregistrer un arrêt
  const saveBtn = document.getElementById(\"arretSaveBtn\");
  if (saveBtn) {
    saveBtn.addEventListener(\"click\", () => {
      const now = getNow();

      const rec = {
        dateTime: formatDateTime(now),
        line: arretLineEl.value,
        sousLigne:
          arretSousZoneRowEl && arretSousZoneRowEl.style.display !== \"none\"
            ? arretSousZoneEl.value
            : \"\",
        machine: arretMachineEl.value,
        duration: Number(document.getElementById(\"arretDuration\")?.value) || 0,
        comment: document.getElementById(\"arretComment\")?.value || \"\",
        article: document.getElementById(\"arretArticle\")?.value || \"\",
      };

      state.arrets.push(rec);
      saveState();

      const durationEl = document.getElementById(\"arretDuration\");
      const commentEl = document.getElementById(\"arretComment\");
      const articleEl = document.getElementById(\"arretArticle\");
      
      if (durationEl) durationEl.value = \"\";
      if (commentEl) commentEl.value = \"\";
      if (articleEl) articleEl.value = \"\";

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
  
  const durationEl = document.getElementById(\"arretDuration\");
  if (durationEl) durationEl.value = duration || \"\";
  
  showSection(\"arrets\");
}

// === TABLEAU HISTORIQUE DES ARRÊTS ===

function refreshArretsView() {
  const table = document.getElementById(\"arretsHistoryTable\");
  if (!table) return;
  
  const tbody = table.querySelector(\"tbody\");
  if (!tbody) return;

  tbody.innerHTML = \"\";

  state.arrets.forEach((rec, idx) => {
    const tr = document.createElement(\"tr\");

    tr.innerHTML = `
      <td>${rec.dateTime}</td>
      <td>${rec.line}</td>
      <td>${rec.sousLigne || \"-\"}</td>
      <td>${rec.machine}</td>
      <td>${rec.article || \"\"}</td>
      <td>${rec.duration}</td>
      <td>${rec.comment || \"\"}</td>
      <td><button class=\"secondary-btn\" data-idx=\"${idx}\">✕</button></td>
    `;

    tbody.appendChild(tr);
  });

  // Suppressions
  tbody.querySelectorAll(\"button[data-idx]\").forEach(btn => {
    btn.addEventListener(\"click\", () => {
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
  const saveBtn = document.getElementById(\"orgSaveBtn\");
  if (!saveBtn) return;

  saveBtn.addEventListener(\"click\", () => {
    const now = getNow();

    const rec = {
      dateTime: formatDateTime(now),
      equipe: state.currentEquipe,
      consigne: document.getElementById(\"orgConsigne\")?.value || \"\",
      visa: document.getElementById(\"orgVisa\")?.value || \"\",
      valide: false,
    };

    state.organisation.push(rec);
    saveState();

    const consigneEl = document.getElementById(\"orgConsigne\");
    const visaEl = document.getElementById(\"orgVisa\");
    
    if (consigneEl) consigneEl.value = \"\";
    if (visaEl) visaEl.value = \"\";

    refreshOrganisationView();
  });
}

function refreshOrganisationView() {
  const table = document.getElementById(\"orgHistoryTable\");
  if (!table) return;
  
  const tbody = table.querySelector(\"tbody\");
  if (!tbody) return;
  
  tbody.innerHTML = \"\";

  state.organisation.forEach((rec, idx) => {
    const tr = document.createElement(\"tr\");

    tr.innerHTML = `
      <td>${rec.dateTime}</td>
      <td>${rec.equipe}</td>
      <td>${rec.consigne}</td>
      <td>${rec.visa}</td>
      <td>
        <button class=\"secondary-btn\" data-idx=\"${idx}\">
          ${rec.valide ? \"✅\" : \"❌\"}
        </button>
      </td>
    `;

    tbody.appendChild(tr);
  });

  tbody.querySelectorAll(\"button[data-idx]\").forEach(btn => {
    btn.addEventListener(\"click\", () => {
      const i = Number(btn.dataset.idx);
      state.organisation[i].valide = !state.organisation[i].valide;
      saveState();
      refreshOrganisationView();
    });
  });
}

// === PERSONNEL ===

function bindPersonnelForm() {
  const saveBtn = document.getElementById(\"persSaveBtn\");
  if (!saveBtn) return;

  saveBtn.addEventListener(\"click\", () => {
    const now = getNow();

    const rec = {
      dateTime: formatDateTime(now),
      equipe: state.currentEquipe,
      nom: document.getElementById(\"persNom\")?.value || \"\",
      motif: document.getElementById(\"persMotif\")?.value || \"\",
      comment: document.getElementById(\"persComment\")?.value || \"\",
    };

    state.personnel.push(rec);
    saveState();

    const nomEl = document.getElementById(\"persNom\");
    const motifEl = document.getElementById(\"persMotif\");
    const commentEl = document.getElementById(\"persComment\");
    
    if (nomEl) nomEl.value = \"\";
    if (motifEl) motifEl.value = \"\";
    if (commentEl) commentEl.value = \"\";

    refreshPersonnelView();
  });
}

function refreshPersonnelView() {
  const table = document.getElementById(\"persHistoryTable\");
  if (!table) return;
  
  const tbody = table.querySelector(\"tbody\");
  if (!tbody) return;
  
  tbody.innerHTML = \"\";

  state.personnel.forEach(rec => {
    const tr = document.createElement(\"tr\");

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
 *   PARTIE 3 / 4 – SCROLL ATELIER
 ********************************************/

function refreshAtelierView() {
  const container = document.getElementById(\"atelier-lines-summary\");
  if (!container) return;
  
  container.innerHTML = \"\";

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
    const artStr = arts.size ? [...arts].join(\", \") : \"-\";

    const div = document.createElement(\"div\");
    div.className = \"summary-card\";

    div.innerHTML = `
      <div class=\"summary-card-title\">${line}</div>
      <div class=\"summary-main\">${total} colis</div>
      <div class=\"summary-sub\">Articles : ${artStr}</div>
      <div class=\"summary-sub\">Cadence moy. ${avg ? avg.toFixed(1) : \"-\"} h</div>
    `;

    container.appendChild(div);
  });

  // Arrêts majeurs
  const table = document.getElementById(\"atelier-arrets-table\");
  if (!table) return;
  
  const tbody = table.querySelector(\"tbody\");
  if (!tbody) return;
  
  tbody.innerHTML = \"\";

  [...state.arrets]
    .sort((a, b) => (b.duration || 0) - (a.duration || 0))
    .forEach(rec => {
      const tr = document.createElement(\"tr\");
      tr.innerHTML = `
        <td>${rec.line}</td>
        <td>${rec.sousLigne || \"-\"}</td>
        <td>${rec.machine}</td>
        <td>${rec.duration}</td>
        <td>${rec.comment || \"\"}</td>
      `;
      tbody.appendChild(tr);
    });

  // scroll horizontal ATELIER (pour les arrêts)
  if (table.parentElement) {
    table.parentElement.style.overflowX = \"auto\";
  }
}

/********************************************
 *   PARTIE 3 / 4 – HISTORIQUE ÉQUIPES
 ********************************************/

let historyChart = null;

function initHistoriqueEquipes() {
  const select = document.getElementById(\"historySelect\");
  if (!select) return;
  
  refreshHistorySelect();

  select.addEventListener(\"change\", () => {
    if (select.value === \"\") {
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
  const select = document.getElementById(\"historySelect\");
  if (!select) return;
  
  select.innerHTML = `<option value=\"\">-- Sélectionner --</option>`;

  archives.forEach((snap, idx) => {
    const opt = document.createElement(\"option\");
    opt.value = idx;
    opt.textContent = snap.label;
    select.appendChild(opt);
  });
}

function clearHistoryView() {
  const summaryEl = document.getElementById(\"history-lines-summary\");
  if (summaryEl) summaryEl.innerHTML = \"\";

  const table = document.getElementById(\"history-arrets-table\");
  if (table) {
    const tbody = table.querySelector(\"tbody\");
    if (tbody) tbody.innerHTML = \"\";
  }

  if (historyChart) {
    historyChart.destroy();
    historyChart = null;
  }
}

// ✅ FONCTION MANQUANTE AJOUTÉE
function refreshHistoryView(snapshot) {
  if (!snapshot || !snapshot.state) return;

  const savedState = snapshot.state;

  // Afficher les lignes de production
  const summaryEl = document.getElementById(\"history-lines-summary\");
  if (summaryEl) {
    summaryEl.innerHTML = \"\";

    LINES.forEach(line => {
      const recs = savedState.production[line] || [];
      const total = recs.reduce((s, r) => s + (r.quantity || 0), 0);
      const cad = recs.map(r => r.cadence).filter(c => c && c > 0);
      const avg = cad.length ? cad.reduce((s, c) => s + c) / cad.length : 0;

      const div = document.createElement(\"div\");
      div.className = \"summary-card\";
      div.innerHTML = `
        <div class=\"summary-card-title\">${line}</div>
        <div class=\"summary-main\">${total} colis</div>
        <div class=\"summary-sub\">Cadence moy. ${avg ? avg.toFixed(1) : \"-\"} h</div>
      `;
      summaryEl.appendChild(div);
    });
  }

  // Afficher les arrêts
  const table = document.getElementById(\"history-arrets-table\");
  if (table) {
    const tbody = table.querySelector(\"tbody\");
    if (tbody) {
      tbody.innerHTML = \"\";

      (savedState.arrets || []).forEach(rec => {
        const tr = document.createElement(\"tr\");
        tr.innerHTML = `
          <td>${rec.line}</td>
          <td>${rec.sousLigne || \"-\"}</td>
          <td>${rec.machine}</td>
          <td>${rec.duration}</td>
          <td>${rec.comment || \"\"}</td>
        `;
        tbody.appendChild(tr);
      });
    }
  }

  // Graphique Chart.js (si disponible)
  const chartCanvas = document.getElementById(\"historyChart\");
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
        scales: {
          y: { beginAtZero: true }
        }
      }
    });
  }
}

/********************************************
 *   PARTIE 4 / 4 – EXPORT GLOBAL (1 SEUL ONGLET)
 ********************************************/

function exportStateToExcel(srcState, filename) {
  // Vérifier que la bibliothèque XLSX est chargée
  if (typeof XLSX === 'undefined') {
    alert(\"Bibliothèque XLSX non chargée. Impossible d'exporter.\");
    console.error(\"XLSX library not found. Please include SheetJS library.\");
    return;
  }

  const wb = XLSX.utils.book_new();

  // === UN SEUL ONGLET ===
  const rows = [[
    \"Type\",
    \"Date/Heure\",
    \"Équipe\",
    \"Ligne\",
    \"Sous-ligne\",
    \"Machine\",
    \"Début\",
    \"Fin\",
    \"Qté\",
    \"Arrêt (min)\",
    \"Cadence\",
    \"Temps restant\",
    \"Commentaire\",
    \"Article\",
    \"Nom personnel\",
    \"Motif personnel\",
    \"Visa organisation\",
    \"Validée organisation\"
  ]];

  // === PRODUCTION ===
  LINES.forEach(line => {
    const recs = srcState.production[line] || [];
    recs.forEach(r => {
      rows.push([
        \"PRODUCTION\",
        r.dateTime,
        r.equipe,
        line,
        \"\",
        \"\",
        r.start || \"\",
        r.end || \"\",
        r.quantity || \"\",
        r.arret || \"\",
        r.cadence || \"\",
        r.remainingTime || \"\",
        r.comment || \"\",
        r.article || \"\",
        \"\",
        \"\",
        \"\",
        \"\"
      ]);
    });
  });

  // === ARRETS ===
  srcState.arrets.forEach(r => {
    rows.push([
      \"ARRET\",
      r.dateTime,
      \"\",
      r.line,
      r.sousLigne || \"\",
      r.machine || \"\",
      \"\",
      \"\",
      \"\",
      r.duration || \"\",
      \"\",
      \"\",
      r.comment || \"\",
      r.article || \"\",
      \"\",
      \"\",
      \"\",
      \"\"
    ]);
  });

  // === ORGANISATION ===
  srcState.organisation.forEach(r => {
    rows.push([
      \"ORGANISATION\",
      r.dateTime,
      r.equipe,
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      r.consigne || \"\",
      \"\",
      \"\",
      \"\",
      r.visa || \"\",
      r.valide ? \"Oui\" : \"Non\"
    ]);
  });

  // === PERSONNEL ===
  srcState.personnel.forEach(r => {
    rows.push([
      \"PERSONNEL\",
      r.dateTime,
      r.equipe,
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      \"\",
      r.comment || \"\",
      \"\",
      r.nom || \"\",
      r.motif || \"\",
      \"\",
      \"\"
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, \"GLOBAL\");
  XLSX.writeFile(wb, filename);
}

// === BIND EXPORT ===
function bindExportGlobal() {
  const btn = document.getElementById(\"exportGlobalBtn\");
  if (!btn) return;

  btn.addEventListener(\"click\", () => {
    const now = getNow();
    const hh = String(now.getHours()).padStart(2, \"0\");
    const mm = String(now.getMinutes()).padStart(2, \"0\");
    const ss = String(now.getSeconds()).padStart(2, \"0\");

    const filename = `Atelier_PPNC_${hh}h${mm}_${ss}.xlsx`;
    exportStateToExcel(state, filename);
  });
}

/********************************************
 *   PARTIE RAZ ÉQUIPE  – changement manuel
 ********************************************/

function bindRAZEquipe() {
  const btn = document.getElementById(\"razBtn\");
  if (!btn) return;

  btn.addEventListener(\"click\", () => {
    if (!confirm(\"RAZ + changement d'équipe + export ?\")) return;

    const now = getNow();
    const week = getWeekNumber(now);
    const quantieme = getQuantieme(now);
    const hh = String(now.getHours()).padStart(2, \"0\");
    const mm = String(now.getMinutes()).padStart(2, \"0\");

    // Choisir équipe finissante
    let finished = prompt(
      \"Quelle équipe vient de finir ? (M, AM, N)\",
      state.currentEquipe
    );
    if (!finished) return;
    finished = finished.toUpperCase().trim();

    if (![\"M\", \"AM\", \"N\"].includes(finished)) {
      alert(\"Équipe invalide.\");
      return;
    }

    // Snapshot archive
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

    // Export du snapshot
    const filename =
      `Atelier_EQ${finished}_Q${quantieme}_S${week}_${hh}h${mm}.xlsx`;
    exportStateToExcel(snap.state, filename);

    // Calcul équipe suivante
    let next = \"M\";
    if (finished === \"M\") next = \"AM\";
    else if (finished === \"AM\") next = \"N\";
    else next = \"M\";

    state.currentEquipe = next;

    // Reset complet
    LINES.forEach(l => {
      state.production[l] = [];
      state.formDraft[l] = {}; // vider brouillons
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
} // ✅ ACCOLADE FERMANTE AJOUTÉE

// === THÈME CLAIR / SOMBRE ===  

function initTheme() {
  const btn = document.getElementById(\"themeToggleBtn\");
  if (!btn) return;

  // Charger le thème sauvegardé
  const saved = localStorage.getItem(\"themeMode\");
  if (saved === \"light\") {
    document.body.classList.add(\"light-mode\");
  }

  // Toggle du thème
  btn.addEventListener(\"click\", () => {
    document.body.classList.toggle(\"light-mode\");

    if (document.body.classList.contains(\"light-mode\")) {
      localStorage.setItem(\"themeMode\", \"light\");
    } else {
      localStorage.setItem(\"themeMode\", \"dark\");
    }
  });
}

// ✅ FONCTIONS MANQUANTES AJOUTÉES (stubs basiques)

function initCalculator() {
  // Fonction pour initialiser une calculatrice (si nécessaire)
  // Ajoutez votre logique ici si cette fonctionnalité existe
  console.log(\"Calculator initialized (stub)\");
}

function bindDDM() {
  // Fonction pour gérer les DDM (Date de Durabilité Minimale?)
  // Ajoutez votre logique ici si cette fonctionnalité existe
  console.log(\"DDM binding initialized (stub)\");
}

// === INIT GLOBALE / ORIENTATION ===

function updateOrientationLayout() {
  const isLandscape = window.innerWidth > window.innerHeight;
  document.body.classList.toggle(\"is-landscape\", isLandscape);
}

document.addEventListener(\"DOMContentLoaded\", () => {
  loadState();
  loadArchives();
  initHeaderDate();
  initNav();
  initLinesSidebar();
  bindProductionForm();
  initArretsForm();
  bindOrganisationForm();
  bindPersonnelForm();
  bindExportGlobal();
  initCalculator(); // ✅ Fonction ajoutée
  bindDDM(); // ✅ Fonction ajoutée
  bindRAZEquipe();
  initHistoriqueEquipes();
  initTheme();

  // Layout en fonction de l'orientation réelle du téléphone
  updateOrientationLayout();
  window.addEventListener(\"resize\", updateOrientationLayout);
  window.addEventListener(\"orientationchange\", updateOrientationLayout);

  showSection(state.currentSection || \"atelier\");
});
"
