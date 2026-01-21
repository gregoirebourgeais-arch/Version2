/********************************************
 *   MODULE PLANNING - VERSION 4.0
 *   Système de planification de production
 *   Drag&Drop, Arrêts planifiés, Échelle précise
 ********************************************/

const PlanningModule = (function() {
  // Configuration
  const DAYS = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
  const TOTAL_HOURS = 5 * 24 + 12; // Lun 00:00 -> Sam 12:00 = 132 heures
  const PX_PER_HOUR = 12; // Pixels par heure pour l'échelle
  const GANTT_WIDTH_PX = TOTAL_HOURS * PX_PER_HOUR; // 1584px

  // État local du planning
  let planningState = {
    articles: [],
    currentOFs: [],
    plannedStops: [], // Arrêts planifiés (pauses, maintenance, etc.)
    savedPlannings: [],
    activePlanning: null,
    weekNumber: null,
    weekStart: null,
    editingOFId: null,
    dragData: null
  };

  let nowMarkerInterval = null;

  // ==========================================
  // UTILITAIRES
  // ==========================================

  function loadPlanningState() {
    try {
      const saved = localStorage.getItem('planning_v4');
      if (saved) {
        const data = JSON.parse(saved);
        planningState = { ...planningState, ...data };
      }
    } catch (e) {
      console.error('Erreur chargement planning:', e);
    }
    if (!planningState.weekStart) {
      setCurrentWeek();
    }
  }

  function savePlanningState() {
    try {
      localStorage.setItem('planning_v4', JSON.stringify(planningState));
    } catch (e) {
      console.error('Erreur sauvegarde planning:', e);
    }
  }

  // Correction du calcul du lundi - Utilise getDay() correctement
  function getMonday(date = new Date()) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    const dayOfWeek = d.getDay(); // 0 = Dimanche, 1 = Lundi, ..., 6 = Samedi
    
    // Si dimanche (0), reculer de 6 jours pour avoir lundi précédent
    // Sinon, reculer de (dayOfWeek - 1) jours
    const daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    d.setDate(d.getDate() - daysToSubtract);
    
    return d;
  }

  function getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  }

  // Calcule le lundi d'une semaine donnée (ISO 8601)
  function getMondayOfWeek(weekNumber, year = new Date().getFullYear()) {
    // Le 4 janvier est toujours dans la semaine 1 (ISO 8601)
    const jan4 = new Date(year, 0, 4);
    const jan4Day = jan4.getDay() || 7; // 1=Lundi, 7=Dimanche
    
    // Trouver le lundi de la semaine 1
    const mondayWeek1 = new Date(jan4);
    mondayWeek1.setDate(jan4.getDate() - jan4Day + 1);
    
    // Ajouter (weekNumber - 1) semaines
    const targetMonday = new Date(mondayWeek1);
    targetMonday.setDate(mondayWeek1.getDate() + (weekNumber - 1) * 7);
    
    return targetMonday;
  }

  function setCurrentWeek() {
    const monday = getMonday(new Date());
    planningState.weekStart = formatDateISO(monday);
    planningState.weekNumber = getWeekNumber(monday);
    savePlanningState();
  }

  // Met à jour la date du lundi quand le numéro de semaine change
  function onWeekNumberChange() {
    const weekNumInput = document.getElementById('planningWeekNumber');
    const weekStartInput = document.getElementById('planningWeekStart');
    
    if (!weekNumInput || !weekStartInput) return;
    
    const weekNum = parseInt(weekNumInput.value);
    if (isNaN(weekNum) || weekNum < 1 || weekNum > 53) return;
    
    const monday = getMondayOfWeek(weekNum);
    const isoDate = formatDateISO(monday);
    
    weekStartInput.value = isoDate;
    planningState.weekNumber = weekNum;
    planningState.weekStart = isoDate;
    savePlanningState();
  }

  function formatDateISO(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function formatDateFR(dateStr) {
    if (!dateStr) return '';
    const [y, m, d] = dateStr.split('-');
    return `${d}/${m}/${y}`;
  }

  function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
  }

  function formatTime(hours, minutes = 0) {
    const h = Math.floor(hours) % 24;
    const m = Math.round(minutes);
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }

  function formatDuration(hours) {
    const h = Math.floor(hours);
    const m = Math.round((hours - h) * 60);
    if (m === 0) return `${h}h`;
    return `${h}h${String(m).padStart(2, '0')}`;
  }

  function calculateDuration(quantity, cadence) {
    if (!cadence || cadence <= 0) return 1;
    return quantity / cadence;
  }

  function getPositionHours(dayIndex, timeStr) {
    const [h, m] = (timeStr || '06:00').split(':').map(Number);
    return dayIndex * 24 + h + (m / 60);
  }

  function positionToDayTime(hours) {
    const totalH = Math.max(0, Math.min(hours, TOTAL_HOURS));
    const day = Math.floor(totalH / 24);
    const hourOfDay = totalH % 24;
    const h = Math.floor(hourOfDay);
    const m = Math.round((hourOfDay - h) * 60);
    return {
      day: Math.min(day, 5),
      time: formatTime(h, m)
    };
  }

  function hoursToPixels(hours) {
    return hours * PX_PER_HOUR;
  }

  function pixelsToHours(px) {
    return px / PX_PER_HOUR;
  }

  // Position actuelle en heures depuis le début de la semaine
  function getCurrentHoursInWeek() {
    if (!planningState.activePlanning?.weekStart) return null;
    
    const now = new Date();
    const weekStartStr = planningState.activePlanning.weekStart;
    const [y, m, d] = weekStartStr.split('-').map(Number);
    const weekStart = new Date(y, m - 1, d, 0, 0, 0, 0);
    
    const diffMs = now - weekStart;
    const diffHours = diffMs / (1000 * 60 * 60);
    
    if (diffHours < 0 || diffHours > TOTAL_HOURS) return null;
    return diffHours;
  }

  // ==========================================
  // ARRÊTS PLANIFIÉS
  // ==========================================

  function renderPlannedStopsList() {
    const container = document.getElementById('plannedStopsList');
    if (!container) return;

    const stops = planningState.plannedStops || [];
    
    if (stops.length === 0) {
      container.innerHTML = '<p class="helper-text">Aucun arrêt planifié.</p>';
      return;
    }

    container.innerHTML = stops.map(stop => `
      <div class="stop-card" data-id="${stop.id}">
        <div class="stop-card-info">
          <div class="stop-card-type">${stop.type}</div>
          <div class="stop-card-details">
            ${stop.line === 'ALL' ? 'Toutes lignes' : stop.line} • 
            ${stop.dayPattern === 'ALL' ? 'Tous les jours' : DAYS[stop.dayPattern]} • 
            ${stop.startTime} - ${formatDuration(stop.duration / 60)}
          </div>
          ${stop.comment ? `<div class="stop-card-comment">${stop.comment}</div>` : ''}
        </div>
        <button class="danger-btn btn-delete-stop" data-id="${stop.id}">🗑️</button>
      </div>
    `).join('');

    container.querySelectorAll('.btn-delete-stop').forEach(btn => {
      btn.addEventListener('click', () => deletePlannedStop(btn.dataset.id));
    });
  }

  function addPlannedStop() {
    const type = document.getElementById('plannedStopType')?.value || 'Pause';
    const line = document.getElementById('plannedStopLine')?.value || 'ALL';
    const dayPattern = document.getElementById('plannedStopDay')?.value || 'ALL';
    const startTime = document.getElementById('plannedStopStart')?.value || '12:00';
    const durationMin = parseInt(document.getElementById('plannedStopDuration')?.value) || 30;
    const comment = document.getElementById('plannedStopComment')?.value || '';

    const stop = {
      id: 'stop_' + generateId(),
      type,
      line,
      dayPattern,
      startTime,
      duration: durationMin,
      comment
    };

    planningState.plannedStops = planningState.plannedStops || [];
    planningState.plannedStops.push(stop);
    savePlanningState();
    
    renderPlannedStopsList();
    renderPreviewGantt();
    
    // Reset form
    document.getElementById('plannedStopComment').value = '';
  }

  function deletePlannedStop(id) {
    planningState.plannedStops = (planningState.plannedStops || []).filter(s => s.id !== id);
    savePlanningState();
    renderPlannedStopsList();
    renderPreviewGantt();
  }

  // Obtenir les arrêts planifiés pour un jour donné
  function getPlannedStopsForGantt() {
    const stops = planningState.plannedStops || [];
    const result = [];
    
    for (let dayIdx = 0; dayIdx < 6; dayIdx++) {
      stops.forEach(stop => {
        if (stop.dayPattern === 'ALL' || parseInt(stop.dayPattern) === dayIdx) {
          const [h, m] = stop.startTime.split(':').map(Number);
          const startHours = dayIdx * 24 + h + m / 60;
          const durationHours = stop.duration / 60;
          
          // Pour samedi, vérifier qu'on ne dépasse pas 12h
          if (dayIdx === 5 && startHours >= 5 * 24 + 12) return;
          
          const lines = stop.line === 'ALL' 
            ? (window.LINES || ['Râpé', 'T2', 'OMORI', 'T1', 'Emballage', 'Dés', 'Filets', 'Prédécoupé'])
            : [stop.line];
          
          lines.forEach(line => {
            result.push({
              id: `${stop.id}_${dayIdx}_${line}`,
              type: 'stop',
              stopType: stop.type,
              line,
              startHours,
              duration: durationHours,
              label: stop.type,
              comment: stop.comment
            });
          });
        }
      });
    }
    
    return result;
  }

  // ==========================================
  // INTÉGRATION ARRÊTS NON PRÉVUS (app.js)
  // ==========================================

  function getUnplannedStops() {
    const arrets = [];
    
    if (window.state?.arrets && Array.isArray(window.state.arrets)) {
      const weekStartStr = planningState.activePlanning?.weekStart || planningState.weekStart;
      if (!weekStartStr) return arrets;
      
      const [y, m, d] = weekStartStr.split('-').map(Number);
      const weekStart = new Date(y, m - 1, d, 0, 0, 0, 0);
      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 5);
      weekEnd.setHours(12, 0, 0, 0);

      window.state.arrets.forEach(arret => {
        const arretDate = new Date(arret.dateTime);
        if (arretDate >= weekStart && arretDate <= weekEnd) {
          const dayIndex = Math.floor((arretDate - weekStart) / (24 * 60 * 60 * 1000));
          const hourOfDay = arretDate.getHours() + arretDate.getMinutes() / 60;
          const startHours = dayIndex * 24 + hourOfDay;
          const durationHours = (arret.duration || 30) / 60;

          arrets.push({
            id: 'unplanned_' + arret.dateTime,
            type: 'unplanned',
            line: arret.line,
            startHours,
            duration: durationHours,
            label: arret.sousZone || 'Arrêt imprévu',
            comment: arret.comment || ''
          });
        }
      });
    }

    return arrets;
  }

  // ==========================================
  // INTÉGRATION PRODUCTION (app.js)
  // ==========================================

  function updateOFsFromProduction() {
    if (!planningState.activePlanning || !window.state?.production) return;

    const weekStartStr = planningState.activePlanning.weekStart;
    if (!weekStartStr) return;
    
    const [y, m, d] = weekStartStr.split('-').map(Number);
    const weekStart = new Date(y, m - 1, d, 0, 0, 0, 0);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 5);
    weekEnd.setHours(12, 0, 0, 0);

    planningState.activePlanning.ofs.forEach(of => {
      const lineProduction = window.state.production[of.line] || [];
      let totalProduced = 0;

      lineProduction.forEach(prod => {
        const prodDate = new Date(prod.dateTime);
        if (prodDate >= weekStart && prodDate <= weekEnd) {
          const prodHours = (prodDate - weekStart) / (1000 * 60 * 60);
          const ofEnd = of.startHours + of.duration;

          if ((prodHours >= of.startHours && prodHours <= ofEnd) ||
              (prod.article && prod.article.toUpperCase() === of.articleCode.toUpperCase())) {
            totalProduced += Number(prod.quantity) || 0;
          }
        }
      });

      of.produced = totalProduced;

      if (totalProduced >= of.quantity) {
        of.status = 'done';
      } else if (totalProduced > 0) {
        of.status = 'running';
      }
    });

    savePlanningState();
  }

  // ==========================================
  // CONFLITS ET CRÉNEAUX
  // ==========================================

  function findNextAvailableSlot(line, duration, startFromHours = 6) {
    const ofs = planningState.currentOFs.filter(of => of.line === line);
    const plannedStops = getPlannedStopsForGantt().filter(s => s.line === line);
    
    const occupied = [
      ...ofs.map(of => ({ start: of.startHours, end: of.startHours + of.duration })),
      ...plannedStops.map(s => ({ start: s.startHours, end: s.startHours + s.duration }))
    ].sort((a, b) => a.start - b.start);

    let cursor = startFromHours;
    
    for (const block of occupied) {
      if (cursor + duration <= block.start) {
        return positionToDayTime(cursor);
      }
      cursor = Math.max(cursor, block.end);
    }

    if (cursor + duration <= TOTAL_HOURS) {
      return positionToDayTime(cursor);
    }

    return null;
  }

  function checkConflict(line, startHours, duration, excludeId = null) {
    const ofs = planningState.currentOFs.filter(of => of.line === line && of.id !== excludeId);
    const endHours = startHours + duration;
    
    for (const of of ofs) {
      const ofEnd = of.startHours + of.duration;
      if (startHours < ofEnd && endHours > of.startHours) {
        return { conflict: true, with: 'OF', item: of };
      }
    }

    return { conflict: false };
  }

  // ==========================================
  // GESTION DES ARTICLES
  // ==========================================

  function renderArticlesList() {
    const container = document.getElementById('planningArticlesList');
    if (!container) return;

    if (planningState.articles.length === 0) {
      container.innerHTML = '<p class="helper-text">Aucun article enregistré.</p>';
      return;
    }

    container.innerHTML = planningState.articles.map(article => `
      <div class="article-card" data-code="${article.code}">
        <div class="article-card-info">
          <div class="article-card-code">${article.code}</div>
          <div class="article-card-label">${article.label}</div>
          <div class="article-card-meta">
            <span>🏭 ${article.line}</span>
            <span>⚡ ${article.cadence} colis/h</span>
            ${article.poids ? `<span>⚖️ ${article.poids} kg</span>` : ''}
          </div>
        </div>
        <div class="article-card-actions">
          <button class="secondary-btn btn-edit-article" data-code="${article.code}">✏️</button>
          <button class="danger-btn btn-delete-article" data-code="${article.code}">🗑️</button>
        </div>
      </div>
    `).join('');

    container.querySelectorAll('.btn-edit-article').forEach(btn => {
      btn.addEventListener('click', () => editArticle(btn.dataset.code));
    });
    container.querySelectorAll('.btn-delete-article').forEach(btn => {
      btn.addEventListener('click', () => deleteArticle(btn.dataset.code));
    });

    updateArticleSelect();
  }

  function addOrUpdateArticle() {
    const code = document.getElementById('planningArticleCode')?.value?.trim()?.toUpperCase() || '';
    const label = document.getElementById('planningArticleLabel')?.value?.trim() || '';
    const line = document.getElementById('planningArticleLine')?.value || '';
    const cadence = parseFloat(document.getElementById('planningArticleCadence')?.value) || 0;
    const poids = parseFloat(document.getElementById('planningArticlePoids')?.value) || 0;

    if (!code || !label || !line || cadence <= 0) {
      alert('Veuillez remplir tous les champs obligatoires.');
      return;
    }

    const existingIndex = planningState.articles.findIndex(a => a.code === code);
    const article = { code, label, line, cadence, poids };

    if (existingIndex >= 0) {
      planningState.articles[existingIndex] = article;
    } else {
      planningState.articles.push(article);
    }

    savePlanningState();
    renderArticlesList();
    clearArticleForm();
    updateOFFormVisibility();
  }

  function editArticle(code) {
    const article = planningState.articles.find(a => a.code === code);
    if (!article) return;

    document.getElementById('planningArticleCode').value = article.code;
    document.getElementById('planningArticleLabel').value = article.label;
    document.getElementById('planningArticleLine').value = article.line;
    document.getElementById('planningArticleCadence').value = article.cadence;
    document.getElementById('planningArticlePoids').value = article.poids || '';
  }

  function deleteArticle(code) {
    if (!confirm(`Supprimer l'article ${code} ?`)) return;
    planningState.articles = planningState.articles.filter(a => a.code !== code);
    savePlanningState();
    renderArticlesList();
    updateOFFormVisibility();
  }

  function clearArticleForm() {
    ['planningArticleCode', 'planningArticleLabel', 'planningArticleCadence', 'planningArticlePoids'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
  }

  // ==========================================
  // GESTION DES OFs
  // ==========================================

  function updateArticleSelect() {
    const select = document.getElementById('planningOFArticle');
    if (!select) return;

    select.innerHTML = '<option value="">-- Sélectionner --</option>' +
      planningState.articles.map(a => 
        `<option value="${a.code}" data-line="${a.line}" data-cadence="${a.cadence}">${a.code} - ${a.label}</option>`
      ).join('');
  }

  function updateOFFormVisibility() {
    const noArticles = document.getElementById('planningOFNoArticles');
    const form = document.getElementById('planningOFForm');
    
    if (planningState.articles.length === 0) {
      if (noArticles) noArticles.classList.remove('hidden');
      if (form) form.style.display = 'none';
    } else {
      if (noArticles) noArticles.classList.add('hidden');
      if (form) form.style.display = 'block';
    }
  }

  function onArticleSelected() {
    const select = document.getElementById('planningOFArticle');
    const option = select?.selectedOptions[0];
    
    if (!option || !option.value) {
      document.getElementById('planningOFLine').value = '';
      document.getElementById('planningOFCadence').value = '';
      return;
    }

    document.getElementById('planningOFLine').value = option.dataset.line;
    document.getElementById('planningOFCadence').value = option.dataset.cadence;

    updateOFPreview();
    suggestNextSlot();
  }

  function updateOFPreview() {
    const qty = parseInt(document.getElementById('planningOFQty')?.value) || 0;
    const cadence = parseFloat(document.getElementById('planningOFCadence')?.value) || 0;
    const day = parseInt(document.getElementById('planningOFDay')?.value) || 0;
    const startTime = document.getElementById('planningOFStart')?.value || '06:00';

    const durationEl = document.getElementById('planningOFDuration');
    const endEl = document.getElementById('planningOFEnd');

    if (qty > 0 && cadence > 0) {
      const duration = calculateDuration(qty, cadence);
      const startHours = getPositionHours(day, startTime);
      const endHours = startHours + duration;
      const endPos = positionToDayTime(endHours);

      if (durationEl) durationEl.textContent = formatDuration(duration);
      if (endEl) endEl.textContent = `${DAYS[endPos.day]} ${endPos.time}`;
    } else {
      if (durationEl) durationEl.textContent = '--';
      if (endEl) endEl.textContent = '--';
    }
  }

  function suggestNextSlot() {
    const line = document.getElementById('planningOFLine')?.value;
    const qty = parseInt(document.getElementById('planningOFQty')?.value) || 0;
    const cadence = parseFloat(document.getElementById('planningOFCadence')?.value) || 0;
    const nextSlotEl = document.getElementById('planningOFNextSlot');

    if (!line || qty <= 0 || cadence <= 0) {
      if (nextSlotEl) nextSlotEl.textContent = '--';
      return;
    }

    const duration = calculateDuration(qty, cadence);
    const slot = findNextAvailableSlot(line, duration);

    if (slot && nextSlotEl) {
      nextSlotEl.textContent = `${DAYS[slot.day]} ${slot.time}`;
    } else if (nextSlotEl) {
      nextSlotEl.textContent = 'Aucun créneau';
    }
  }

  function autoPlaceOF() {
    const line = document.getElementById('planningOFLine')?.value;
    const qty = parseInt(document.getElementById('planningOFQty')?.value) || 0;
    const cadence = parseFloat(document.getElementById('planningOFCadence')?.value) || 0;

    if (!line || qty <= 0 || cadence <= 0) {
      alert('Sélectionnez un article et entrez une quantité.');
      return;
    }

    const duration = calculateDuration(qty, cadence);
    const slot = findNextAvailableSlot(line, duration);

    if (slot) {
      document.getElementById('planningOFDay').value = slot.day;
      document.getElementById('planningOFStart').value = slot.time;
      updateOFPreview();
    } else {
      alert('Aucun créneau disponible.');
    }
  }

  function addOF() {
    const articleCode = document.getElementById('planningOFArticle')?.value;
    const line = document.getElementById('planningOFLine')?.value;
    const qty = parseInt(document.getElementById('planningOFQty')?.value) || 0;
    const cadence = parseFloat(document.getElementById('planningOFCadence')?.value) || 0;
    const day = parseInt(document.getElementById('planningOFDay')?.value) || 0;
    const startTime = document.getElementById('planningOFStart')?.value || '06:00';

    if (!articleCode || qty <= 0) {
      alert('Sélectionnez un article et entrez une quantité.');
      return;
    }

    const article = planningState.articles.find(a => a.code === articleCode);
    if (!article) return;

    const duration = calculateDuration(qty, cadence);
    const startHours = getPositionHours(day, startTime);

    const conflict = checkConflict(line, startHours, duration);
    if (conflict.conflict) {
      alert(`Conflit avec un ${conflict.with} sur cette ligne !`);
      return;
    }

    if (startHours + duration > TOTAL_HOURS) {
      alert('L\'OF dépasse la fin de la semaine.');
      return;
    }

    const of = {
      id: 'of_' + generateId(),
      articleCode,
      articleLabel: article.label,
      line,
      quantity: qty,
      cadence,
      duration,
      day,
      startTime,
      startHours,
      status: 'planned',
      produced: 0
    };

    planningState.currentOFs.push(of);
    savePlanningState();
    
    renderOFList();
    renderPreviewGantt();
    clearOFForm();
  }

  function clearOFForm() {
    document.getElementById('planningOFArticle').value = '';
    document.getElementById('planningOFLine').value = '';
    document.getElementById('planningOFQty').value = '';
    document.getElementById('planningOFCadence').value = '';
    document.getElementById('planningOFDay').value = '0';
    document.getElementById('planningOFStart').value = '06:00';
    document.getElementById('planningOFDuration').textContent = '--';
    document.getElementById('planningOFEnd').textContent = '--';
    document.getElementById('planningOFNextSlot').textContent = '--';
  }

  function renderOFList() {
    const container = document.getElementById('planningOFList');
    if (!container) return;

    if (planningState.currentOFs.length === 0) {
      container.innerHTML = '<p class="helper-text">Aucun OF ajouté.</p>';
      return;
    }

    const sorted = [...planningState.currentOFs].sort((a, b) => a.startHours - b.startHours);

    container.innerHTML = sorted.map(of => {
      const endPos = positionToDayTime(of.startHours + of.duration);
      return `
        <div class="of-card" data-id="${of.id}">
          <div class="of-card-status ${of.status}"></div>
          <div class="of-card-info">
            <div class="of-card-title">${of.articleCode} - ${of.articleLabel}</div>
            <div class="of-card-details">${of.line} • ${of.quantity} colis • ${formatDuration(of.duration)}</div>
          </div>
          <div class="of-card-time">
            <div class="of-card-day">${DAYS[of.day]}</div>
            <div class="of-card-hours">${of.startTime} → ${endPos.time}</div>
          </div>
          <div class="of-card-actions">
            <button class="secondary-btn btn-edit-of" data-id="${of.id}">✏️</button>
            <button class="danger-btn btn-delete-of" data-id="${of.id}">🗑️</button>
          </div>
        </div>
      `;
    }).join('');

    container.querySelectorAll('.btn-edit-of').forEach(btn => {
      btn.addEventListener('click', () => openOFEditor(btn.dataset.id));
    });
    container.querySelectorAll('.btn-delete-of').forEach(btn => {
      btn.addEventListener('click', () => deleteOF(btn.dataset.id));
    });
  }

  function deleteOF(id) {
    planningState.currentOFs = planningState.currentOFs.filter(of => of.id !== id);
    savePlanningState();
    renderOFList();
    renderPreviewGantt();
  }

  function clearAllOFs() {
    if (!confirm('Supprimer tous les OFs ?')) return;
    planningState.currentOFs = [];
    savePlanningState();
    renderOFList();
    renderPreviewGantt();
  }

  // ==========================================
  // ÉDITEUR D'OF
  // ==========================================

  function openOFEditor(id) {
    const of = planningState.currentOFs.find(o => o.id === id) || 
               (planningState.activePlanning?.ofs || []).find(o => o.id === id);
    if (!of) return;

    planningState.editingOFId = id;

    document.getElementById('planningEditOFQty').value = of.quantity;
    document.getElementById('planningEditOFDay').value = of.day;
    document.getElementById('planningEditOFStart').value = of.startTime;
    document.getElementById('planningEditOFStatus').value = of.status;

    document.getElementById('planningOFEditModal').classList.remove('hidden');
  }

  function closeOFEditor() {
    document.getElementById('planningOFEditModal').classList.add('hidden');
    planningState.editingOFId = null;
  }

  function saveOFEdit() {
    const id = planningState.editingOFId;
    if (!id) return;

    let of = planningState.currentOFs.find(o => o.id === id);
    let isActive = false;
    
    if (!of && planningState.activePlanning) {
      of = planningState.activePlanning.ofs.find(o => o.id === id);
      isActive = true;
    }
    
    if (!of) return;

    const newQty = parseInt(document.getElementById('planningEditOFQty').value) || of.quantity;
    const newDay = parseInt(document.getElementById('planningEditOFDay').value);
    const newTime = document.getElementById('planningEditOFStart').value;
    const newStatus = document.getElementById('planningEditOFStatus').value;

    of.quantity = newQty;
    of.duration = calculateDuration(newQty, of.cadence);
    of.day = newDay;
    of.startTime = newTime;
    of.startHours = getPositionHours(newDay, newTime);
    of.status = newStatus;

    savePlanningState();
    closeOFEditor();

    if (isActive) {
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
    } else {
      renderOFList();
      renderPreviewGantt();
    }
  }

  function deleteOFFromEditor() {
    const id = planningState.editingOFId;
    if (!id || !confirm('Supprimer cet OF ?')) return;

    planningState.currentOFs = planningState.currentOFs.filter(o => o.id !== id);
    if (planningState.activePlanning) {
      planningState.activePlanning.ofs = planningState.activePlanning.ofs.filter(o => o.id !== id);
    }

    savePlanningState();
    closeOFEditor();
    
    renderOFList();
    renderPreviewGantt();
    renderActiveGantt();
    renderActiveOFList();
  }

  // ==========================================
  // DRAG AND DROP
  // ==========================================

  function initDragDrop(container, isActive) {
    container.querySelectorAll('.gantt-of-block[data-draggable="true"]').forEach(block => {
      block.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return; // Left click only
        
        const ofId = block.dataset.ofId;
        const rect = block.getBoundingClientRect();
        const containerRect = block.closest('.gantt-line-track').getBoundingClientRect();
        
        planningState.dragData = {
          ofId,
          startX: e.clientX,
          originalLeft: block.offsetLeft,
          block,
          isActive,
          containerWidth: containerRect.width
        };
        
        block.classList.add('dragging');
        document.addEventListener('mousemove', handleDragMove);
        document.addEventListener('mouseup', handleDragEnd);
        e.preventDefault();
      });
    });
  }

  function handleDragMove(e) {
    if (!planningState.dragData) return;
    
    const { block, startX, originalLeft, containerWidth } = planningState.dragData;
    const deltaX = e.clientX - startX;
    let newLeft = originalLeft + deltaX;
    
    // Limiter au conteneur
    newLeft = Math.max(0, Math.min(newLeft, containerWidth - block.offsetWidth));
    
    block.style.left = newLeft + 'px';
  }

  function handleDragEnd(e) {
    if (!planningState.dragData) return;
    
    const { ofId, block, isActive, containerWidth } = planningState.dragData;
    
    document.removeEventListener('mousemove', handleDragMove);
    document.removeEventListener('mouseup', handleDragEnd);
    
    block.classList.remove('dragging');
    
    // Calculer la nouvelle position en heures
    const newLeft = parseInt(block.style.left) || 0;
    const newStartHours = pixelsToHours(newLeft);
    
    // Arrondir à 15 minutes (0.25h)
    const roundedHours = Math.round(newStartHours * 4) / 4;
    
    // Trouver et mettre à jour l'OF
    let of = isActive 
      ? planningState.activePlanning?.ofs.find(o => o.id === ofId)
      : planningState.currentOFs.find(o => o.id === ofId);
    
    if (of && roundedHours >= 0 && roundedHours + of.duration <= TOTAL_HOURS) {
      of.startHours = roundedHours;
      const dayTime = positionToDayTime(roundedHours);
      of.day = dayTime.day;
      of.startTime = dayTime.time;
      
      savePlanningState();
    }
    
    planningState.dragData = null;
    
    // Rafraîchir
    if (isActive) {
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
    } else {
      renderOFList();
      renderPreviewGantt();
    }
  }

  // ==========================================
  // GANTT CHART - VERSION CORRIGÉE
  // ==========================================

  function renderGantt(containerId, ofs, interactive = false) {
    const container = document.getElementById(containerId);
    if (!container) return;

    const lines = window.LINES || ['Râpé', 'T2', 'OMORI', 'T1', 'Emballage', 'Dés', 'Filets', 'Prédécoupé'];
    const plannedStops = getPlannedStopsForGantt();
    const unplannedStops = interactive ? getUnplannedStops() : [];
    const allStops = [...plannedStops, ...unplannedStops];
    
    if (interactive && planningState.activePlanning) {
      updateOFsFromProduction();
    }

    if ((!ofs || ofs.length === 0) && allStops.length === 0) {
      container.innerHTML = `<div class="gantt-empty">${
        containerId === 'planningGantt' 
          ? 'Sélectionnez un planning'
          : 'Ajoutez des OFs'
      }</div>`;
      return;
    }

    // Filtrer les lignes avec contenu
    const usedLines = [...new Set([
      ...(ofs || []).map(of => of.line),
      ...allStops.map(s => s.line)
    ])];
    const sortedLines = lines.filter(l => usedLines.includes(l));
    if (sortedLines.length === 0) {
      sortedLines.push(...lines.slice(0, 3));
    }

    let html = `<div class="gantt-wrapper-v4" style="width: ${GANTT_WIDTH_PX + 130}px">`;
    
    // En-tête avec jours
    html += '<div class="gantt-header-v4">';
    html += '<div class="gantt-corner-v4">Lignes</div>';
    html += '<div class="gantt-timeline-header" style="width: ' + GANTT_WIDTH_PX + 'px">';
    
    // Labels des jours
    for (let d = 0; d < 6; d++) {
      const dayHours = d < 5 ? 24 : 12;
      const left = hoursToPixels(d * 24);
      const width = hoursToPixels(dayHours);
      html += `<div class="gantt-day-header-v4" style="left:${left}px;width:${width}px">${DAYS[d]}</div>`;
    }
    
    // Graduation des heures
    html += '<div class="gantt-hours-row">';
    for (let h = 0; h <= TOTAL_HOURS; h += 6) {
      const left = hoursToPixels(h);
      const hourOfDay = h % 24;
      html += `<span class="gantt-hour-label" style="left:${left}px">${String(hourOfDay).padStart(2,'0')}h</span>`;
    }
    html += '</div>';
    
    html += '</div></div>';

    // Lignes du Gantt
    sortedLines.forEach(line => {
      const lineOFs = (ofs || []).filter(of => of.line === line);
      const lineStops = allStops.filter(s => s.line === line);
      
      html += `<div class="gantt-row-v4">`;
      html += `<div class="gantt-line-name-v4">${line}</div>`;
      html += `<div class="gantt-line-track" data-line="${line}" style="width: ${GANTT_WIDTH_PX}px">`;
      
      // Grille de fond avec lignes verticales pour chaque heure
      for (let h = 0; h <= TOTAL_HOURS; h += 6) {
        const left = hoursToPixels(h);
        html += `<div class="gantt-grid-line" style="left:${left}px"></div>`;
      }
      
      // Fond alterné par jour
      for (let d = 0; d < 6; d++) {
        const left = hoursToPixels(d * 24);
        const width = hoursToPixels(d < 5 ? 24 : 12);
        html += `<div class="gantt-day-bg-v4 ${d % 2 === 0 ? 'even' : ''}" style="left:${left}px;width:${width}px"></div>`;
      }
      
      // Marqueur "maintenant"
      if (interactive) {
        const nowHours = getCurrentHoursInWeek();
        if (nowHours !== null && nowHours >= 0 && nowHours <= TOTAL_HOURS) {
          const nowLeft = hoursToPixels(nowHours);
          html += `<div class="gantt-now-line" style="left:${nowLeft}px"><span class="now-label">Maintenant</span></div>`;
        }
      }
      
      // Arrêts planifiés
      lineStops.filter(s => s.type === 'stop').forEach(stop => {
        const left = hoursToPixels(stop.startHours);
        const width = Math.max(hoursToPixels(stop.duration), 30);
        html += `<div class="gantt-stop-block planned" style="left:${left}px;width:${width}px">
          <span class="stop-icon">⏸️</span>
          <span class="stop-label">${stop.label}</span>
        </div>`;
      });
      
      // Arrêts imprévus
      lineStops.filter(s => s.type === 'unplanned').forEach(stop => {
        const left = hoursToPixels(stop.startHours);
        const width = Math.max(hoursToPixels(stop.duration), 30);
        html += `<div class="gantt-stop-block unplanned" style="left:${left}px;width:${width}px">
          <span class="stop-icon">🛑</span>
          <span class="stop-label">${stop.label}</span>
        </div>`;
      });
      
      // OFs
      lineOFs.forEach(of => {
        const left = hoursToPixels(of.startHours);
        const width = Math.max(hoursToPixels(of.duration), 50);
        const endPos = positionToDayTime(of.startHours + of.duration);
        const progress = of.quantity > 0 ? Math.min(100, (of.produced || 0) / of.quantity * 100) : 0;
        
        // Statut visuel
        let statusClass = of.status;
        if (interactive) {
          const nowHours = getCurrentHoursInWeek();
          if (nowHours !== null && of.status !== 'done' && of.startHours + of.duration < nowHours && progress < 100) {
            statusClass = 'late';
          }
        }
        
        html += `<div class="gantt-of-block ${statusClass}" 
          style="left:${left}px;width:${width}px"
          data-of-id="${of.id}"
          data-draggable="${interactive}">
          ${progress > 0 ? `<div class="gantt-progress-bar" style="width:${progress}%"></div>` : ''}
          <div class="gantt-of-content">
            <div class="gantt-of-code">${of.articleCode}</div>
            <div class="gantt-of-qty">${of.quantity} colis</div>
            <div class="gantt-of-time">${of.startTime}→${endPos.time}</div>
          </div>
        </div>`;
      });
      
      html += '</div></div>';
    });
    
    html += '</div>';
    container.innerHTML = html;

    // Initialiser le drag & drop
    if (interactive) {
      initDragDrop(container, true);
      
      // Click pour éditer
      container.querySelectorAll('.gantt-of-block').forEach(block => {
        block.addEventListener('dblclick', () => {
          openOFEditor(block.dataset.ofId);
        });
      });
    } else {
      initDragDrop(container, false);
    }
  }

  function renderPreviewGantt() {
    renderGantt('planningPreviewGantt', planningState.currentOFs, false);
  }

  function renderActiveGantt() {
    if (!planningState.activePlanning) {
      const container = document.getElementById('planningGantt');
      if (container) {
        container.innerHTML = '<div class="gantt-empty">Sélectionnez un planning</div>';
      }
      return;
    }
    renderGantt('planningGantt', planningState.activePlanning.ofs, true);
  }

  function updateNowMarkers() {
    const nowHours = getCurrentHoursInWeek();
    if (nowHours === null) return;
    
    document.querySelectorAll('.gantt-now-line').forEach(marker => {
      const left = hoursToPixels(nowHours);
      marker.style.left = left + 'px';
    });
  }

  // ==========================================
  // VALIDATION ET LANCEMENT
  // ==========================================

  function validatePlanning() {
    if (planningState.currentOFs.length === 0) {
      alert('Ajoutez au moins un OF.');
      return;
    }

    const weekNum = document.getElementById('planningWeekNumber')?.value;
    const weekStart = document.getElementById('planningWeekStart')?.value;

    if (!weekNum || !weekStart) {
      alert('Définissez la semaine.');
      return;
    }

    const existingIdx = planningState.savedPlannings.findIndex(p => p.weekNumber == weekNum);
    
    if (existingIdx >= 0 && !confirm(`Remplacer le planning semaine ${weekNum} ?`)) {
      return;
    }

    const planning = {
      id: 'plan_' + Date.now(),
      weekNumber: parseInt(weekNum),
      weekStart,
      createdAt: new Date().toISOString(),
      ofs: JSON.parse(JSON.stringify(planningState.currentOFs)),
      plannedStops: JSON.parse(JSON.stringify(planningState.plannedStops || []))
    };

    if (existingIdx >= 0) {
      planningState.savedPlannings[existingIdx] = planning;
    } else {
      planningState.savedPlannings.push(planning);
    }

    planningState.currentOFs = [];
    savePlanningState();
    
    alert(`Planning semaine ${weekNum} validé !`);
    
    renderOFList();
    renderPreviewGantt();
    updatePlanningSelect();
  }

  function updatePlanningSelect() {
    const select = document.getElementById('planningSelectWeek');
    if (!select) return;

    select.innerHTML = '<option value="">-- Sélectionner --</option>' +
      planningState.savedPlannings
        .sort((a, b) => b.weekNumber - a.weekNumber)
        .map(p => `<option value="${p.id}">Semaine ${p.weekNumber} (${formatDateFR(p.weekStart)})</option>`)
        .join('');
  }

  function launchPlanning() {
    const select = document.getElementById('planningSelectWeek');
    const planningId = select?.value;

    if (!planningId) {
      alert('Sélectionnez un planning.');
      return;
    }

    const planning = planningState.savedPlannings.find(p => p.id === planningId);
    if (!planning) return;

    planningState.activePlanning = JSON.parse(JSON.stringify(planning));
    
    // Restaurer les arrêts planifiés du planning
    if (planning.plannedStops) {
      planningState.plannedStops = JSON.parse(JSON.stringify(planning.plannedStops));
    }
    
    updateOFsFromProduction();
    savePlanningState();

    const activeInfo = document.getElementById('planningActiveInfo');
    const activeWeek = document.getElementById('planningActiveWeek');
    if (activeInfo) activeInfo.classList.remove('hidden');
    if (activeWeek) activeWeek.textContent = planning.weekNumber;

    renderActiveGantt();
    renderActiveOFList();
    updateDelays();
    startNowMarkerTimer();
  }

  function startNowMarkerTimer() {
    if (nowMarkerInterval) clearInterval(nowMarkerInterval);
    nowMarkerInterval = setInterval(updateNowMarkers, 60000);
  }

  function renderActiveOFList() {
    const container = document.getElementById('planningActiveOFList');
    if (!container) return;

    if (!planningState.activePlanning || planningState.activePlanning.ofs.length === 0) {
      container.innerHTML = '<p class="helper-text">Aucun planning chargé.</p>';
      return;
    }

    const sorted = [...planningState.activePlanning.ofs].sort((a, b) => a.startHours - b.startHours);

    container.innerHTML = sorted.map(of => {
      const endPos = positionToDayTime(of.startHours + of.duration);
      const progress = of.quantity > 0 ? Math.round((of.produced || 0) / of.quantity * 100) : 0;
      
      return `
        <div class="of-card" data-id="${of.id}">
          <div class="of-card-status ${of.status}"></div>
          <div class="of-card-info">
            <div class="of-card-title">${of.articleCode} - ${of.articleLabel}</div>
            <div class="of-card-details">
              ${of.line} • ${of.quantity} colis
              ${of.produced > 0 ? `• ✓ ${of.produced} (${progress}%)` : ''}
            </div>
          </div>
          <div class="of-card-time">
            <div class="of-card-day">${DAYS[of.day]}</div>
            <div class="of-card-hours">${of.startTime} → ${endPos.time}</div>
          </div>
          <div class="of-card-actions">
            <button class="secondary-btn btn-edit-of" data-id="${of.id}">✏️</button>
          </div>
        </div>
      `;
    }).join('');

    container.querySelectorAll('.btn-edit-of').forEach(btn => {
      btn.addEventListener('click', () => openOFEditor(btn.dataset.id));
    });
  }

  // ==========================================
  // AVANCE / RETARD
  // ==========================================

  function updateDelays() {
    const container = document.getElementById('planningDelaysContainer');
    if (!container) return;

    if (!planningState.activePlanning) {
      container.innerHTML = '<p class="helper-text">Chargez un planning.</p>';
      return;
    }

    updateOFsFromProduction();

    const lines = [...new Set(planningState.activePlanning.ofs.map(of => of.line))];
    const nowHours = getCurrentHoursInWeek() || 0;

    container.innerHTML = lines.map(line => {
      const lineOFs = planningState.activePlanning.ofs.filter(of => of.line === line);
      
      let totalPlanned = 0;
      let totalProduced = 0;
      let totalExpected = 0;

      lineOFs.forEach(of => {
        totalPlanned += of.quantity;
        totalProduced += of.produced || 0;
        
        const ofEnd = of.startHours + of.duration;
        if (nowHours >= ofEnd) {
          totalExpected += of.quantity;
        } else if (nowHours > of.startHours) {
          const ratio = (nowHours - of.startHours) / of.duration;
          totalExpected += Math.round(of.quantity * ratio);
        }
      });

      const delta = totalProduced - totalExpected;
      const deltaClass = delta > 0 ? 'positive' : delta < 0 ? 'negative' : 'neutral';
      const deltaText = delta > 0 ? `+${delta}` : delta.toString();

      return `
        <div class="delay-card-new">
          <h4>${line}</h4>
          <div class="delay-stats">
            <div class="delay-stat">
              <span class="delay-stat-label">Planifié</span>
              <span class="delay-stat-value">${totalPlanned}</span>
            </div>
            <div class="delay-stat">
              <span class="delay-stat-label">Attendu</span>
              <span class="delay-stat-value">${totalExpected}</span>
            </div>
            <div class="delay-stat">
              <span class="delay-stat-label">Produit</span>
              <span class="delay-stat-value">${totalProduced}</span>
            </div>
            <div class="delay-stat">
              <span class="delay-stat-label">Écart</span>
              <span class="delay-stat-value ${deltaClass}">${deltaText}</span>
            </div>
          </div>
        </div>
      `;
    }).join('');
  }

  // ==========================================
  // INITIALISATION
  // ==========================================

  function initLineSelects() {
    const lines = window.LINES || ['Râpé', 'T2', 'OMORI', 'T1', 'Emballage', 'Dés', 'Filets', 'Prédécoupé'];
    
    ['planningArticleLine', 'planningOFLine', 'plannedStopLine'].forEach(id => {
      const select = document.getElementById(id);
      if (select) {
        if (id === 'plannedStopLine') {
          select.innerHTML = '<option value="ALL">Toutes les lignes</option>' + 
            lines.map(l => `<option value="${l}">${l}</option>`).join('');
        } else {
          select.innerHTML = lines.map(l => `<option value="${l}">${l}</option>`).join('');
        }
      }
    });
  }

  function initWeekInputs() {
    const weekNumInput = document.getElementById('planningWeekNumber');
    const weekStartInput = document.getElementById('planningWeekStart');

    if (!planningState.weekNumber || !planningState.weekStart) {
      setCurrentWeek();
    }

    if (weekNumInput) weekNumInput.value = planningState.weekNumber || '';
    if (weekStartInput) weekStartInput.value = planningState.weekStart || '';
  }

  function bindEvents() {
    // Navigation onglets
    document.querySelectorAll('#section-planning .planning-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const tab = btn.dataset.tab;
        document.querySelectorAll('#section-planning .planning-tab-btn').forEach(b => 
          b.classList.toggle('active', b.dataset.tab === tab)
        );
        document.querySelectorAll('#section-planning .planning-pane').forEach(p => 
          p.classList.toggle('hidden', p.id !== `planning-tab-${tab}`)
        );
        
        if (tab === 'run') {
          updateOFsFromProduction();
          renderActiveGantt();
          updateDelays();
        }
      });
    });

    // Articles
    document.getElementById('planningArticleAddBtn')?.addEventListener('click', addOrUpdateArticle);
    document.getElementById('planningArticleClearBtn')?.addEventListener('click', clearArticleForm);

    // Semaine - mise à jour automatique du lundi quand le numéro de semaine change
    document.getElementById('planningWeekNumber')?.addEventListener('change', onWeekNumberChange);
    document.getElementById('planningWeekNumber')?.addEventListener('input', onWeekNumberChange);
    
    document.getElementById('planningAutoWeekBtn')?.addEventListener('click', () => {
      setCurrentWeek();
      initWeekInputs();
    });

    // OFs
    document.getElementById('planningOFArticle')?.addEventListener('change', onArticleSelected);
    document.getElementById('planningOFQty')?.addEventListener('input', () => {
      updateOFPreview();
      suggestNextSlot();
    });
    document.getElementById('planningOFDay')?.addEventListener('change', updateOFPreview);
    document.getElementById('planningOFStart')?.addEventListener('change', updateOFPreview);
    document.getElementById('planningOFAddBtn')?.addEventListener('click', addOF);
    document.getElementById('planningOFSuggestBtn')?.addEventListener('click', autoPlaceOF);
    document.getElementById('planningClearAllOFBtn')?.addEventListener('click', clearAllOFs);

    // Arrêts planifiés
    document.getElementById('plannedStopAddBtn')?.addEventListener('click', addPlannedStop);

    // Validation
    document.getElementById('planningValidateBtn')?.addEventListener('click', validatePlanning);

    // Lancement
    document.getElementById('planningLaunchBtn')?.addEventListener('click', launchPlanning);

    // Éditeur OF
    document.getElementById('planningOFEditClose')?.addEventListener('click', closeOFEditor);
    document.getElementById('planningEditOFSave')?.addEventListener('click', saveOFEdit);
    document.getElementById('planningEditOFDelete')?.addEventListener('click', deleteOFFromEditor);

    // Semaine
    document.getElementById('planningWeekNumber')?.addEventListener('change', (e) => {
      planningState.weekNumber = parseInt(e.target.value);
      savePlanningState();
    });
    document.getElementById('planningWeekStart')?.addEventListener('change', (e) => {
      planningState.weekStart = e.target.value;
      savePlanningState();
    });

    // Rafraîchir
    document.getElementById('planningRefreshBtn')?.addEventListener('click', () => {
      updateOFsFromProduction();
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
    });
  }

  function init() {
    console.log('Initialisation module Planning v4...');
    
    loadPlanningState();
    initLineSelects();
    initWeekInputs();
    bindEvents();
    
    renderArticlesList();
    updateOFFormVisibility();
    renderPlannedStopsList();
    renderOFList();
    renderPreviewGantt();
    updatePlanningSelect();
    
    if (planningState.activePlanning) {
      const activeInfo = document.getElementById('planningActiveInfo');
      const activeWeek = document.getElementById('planningActiveWeek');
      if (activeInfo) activeInfo.classList.remove('hidden');
      if (activeWeek) activeWeek.textContent = planningState.activePlanning.weekNumber;
      
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
      startNowMarkerTimer();
    }

    console.log('Module Planning v4 initialisé.');
  }

  return {
    init,
    openOFEditor,
    refresh: () => {
      updateOFsFromProduction();
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
    },
    getState: () => planningState
  };
})();

document.addEventListener('DOMContentLoaded', () => setTimeout(PlanningModule.init, 300));
