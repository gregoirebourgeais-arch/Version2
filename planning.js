/********************************************
 *   MODULE PLANNING - VERSION 3.0
 *   Système de planification de production
 *   Avec Drag&Drop, intégration arrêts/production
 ********************************************/

const PlanningModule = (function() {
  // Configuration
  const DAYS = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
  const DAYS_SHORT = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam'];
  const TOTAL_HOURS = 5 * 24 + 12; // Lun 00:00 -> Sam 12:00 = 132 heures
  const GANTT_WIDTH_PX = 1320; // 10px par heure

  // État local du planning
  let planningState = {
    articles: [],
    currentOFs: [],
    savedPlannings: [],
    activePlanning: null,
    weekNumber: null,
    weekStart: null,
    editingOFId: null,
    draggedOFId: null
  };

  // Intervalle pour le marqueur temps réel
  let nowMarkerInterval = null;

  // ==========================================
  // UTILITAIRES
  // ==========================================

  function loadPlanningState() {
    try {
      const saved = localStorage.getItem('planning_v3');
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
      localStorage.setItem('planning_v3', JSON.stringify(planningState));
    } catch (e) {
      console.error('Erreur sauvegarde planning:', e);
    }
  }

  function getMonday(date = new Date()) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    d.setDate(diff);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  }

  function setCurrentWeek() {
    const monday = getMonday();
    planningState.weekStart = monday.toISOString().split('T')[0];
    planningState.weekNumber = getWeekNumber(monday);
    savePlanningState();
  }

  function generateId() {
    return 'of_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  function formatTime(hours, minutes = 0) {
    return `${String(Math.floor(hours)).padStart(2, '0')}:${String(Math.round(minutes)).padStart(2, '0')}`;
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
    const day = Math.floor(hours / 24);
    const hourOfDay = hours % 24;
    const h = Math.floor(hourOfDay);
    const m = Math.round((hourOfDay - h) * 60);
    return {
      day: Math.min(Math.max(day, 0), 5),
      time: formatTime(h, m)
    };
  }

  function hoursToPixels(hours) {
    return (hours / TOTAL_HOURS) * GANTT_WIDTH_PX;
  }

  function pixelsToHours(px) {
    return (px / GANTT_WIDTH_PX) * TOTAL_HOURS;
  }

  // Obtenir la date/heure actuelle en heures depuis le début de la semaine
  function getCurrentHoursInWeek() {
    if (!planningState.weekStart) return null;
    
    const now = new Date();
    const weekStart = new Date(planningState.weekStart);
    weekStart.setHours(0, 0, 0, 0);
    
    const diffMs = now - weekStart;
    const diffHours = diffMs / (1000 * 60 * 60);
    
    // Si on est avant ou après la semaine, retourner null
    if (diffHours < 0 || diffHours > TOTAL_HOURS) return null;
    
    return diffHours;
  }

  // ==========================================
  // INTÉGRATION AVEC ARRÊTS (app.js state.arrets)
  // ==========================================

  function getArretsPlanifies() {
    const arrets = [];
    
    // Récupérer les arrêts depuis l'état global (app.js)
    if (window.state && window.state.arrets && Array.isArray(window.state.arrets)) {
      const weekStart = new Date(planningState.weekStart || new Date());
      weekStart.setHours(0, 0, 0, 0);
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
            id: 'arret_' + arret.dateTime,
            type: 'arret',
            line: arret.line,
            startHours,
            duration: durationHours,
            label: arret.sousZone || arret.comment || 'Arrêt',
            comment: arret.comment || ''
          });
        }
      });
    }

    return arrets;
  }

  // ==========================================
  // INTÉGRATION AVEC PRODUCTION (app.js state.production)
  // ==========================================

  function updateOFsFromProduction() {
    if (!planningState.activePlanning || !window.state || !window.state.production) return;

    const weekStart = new Date(planningState.weekStart || new Date());
    weekStart.setHours(0, 0, 0, 0);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 5);
    weekEnd.setHours(12, 0, 0, 0);

    // Pour chaque ligne, récupérer la production
    planningState.activePlanning.ofs.forEach(of => {
      const lineProduction = window.state.production[of.line] || [];
      let totalProduced = 0;

      lineProduction.forEach(prod => {
        const prodDate = new Date(prod.dateTime);
        if (prodDate >= weekStart && prodDate <= weekEnd) {
          // Vérifier si cette production correspond à cet OF (par article ou par créneau horaire)
          const prodHours = (prodDate - weekStart) / (1000 * 60 * 60);
          const ofEnd = of.startHours + of.duration;

          // Si la production est dans le créneau de l'OF
          if (prodHours >= of.startHours && prodHours <= ofEnd) {
            totalProduced += Number(prod.quantity) || 0;
          }
          // Ou si l'article correspond
          else if (prod.article && prod.article.toUpperCase() === of.articleCode.toUpperCase()) {
            totalProduced += Number(prod.quantity) || 0;
          }
        }
      });

      of.produced = totalProduced;

      // Mettre à jour le statut automatiquement
      if (totalProduced >= of.quantity) {
        of.status = 'done';
      } else if (totalProduced > 0) {
        of.status = 'running';
      }
    });

    savePlanningState();
  }

  // ==========================================
  // GESTION DES CONFLITS ET CRÉNEAUX
  // ==========================================

  function findNextAvailableSlot(line, duration, startFromHours = 6) {
    const ofs = planningState.currentOFs.filter(of => of.line === line);
    const arrets = getArretsPlanifies().filter(a => a.line === line);
    
    // Combiner OFs et arrêts
    const occupied = [
      ...ofs.map(of => ({ start: of.startHours, end: of.startHours + of.duration })),
      ...arrets.map(a => ({ start: a.startHours, end: a.startHours + a.duration }))
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
    const ofs = planningState.currentOFs.filter(of => 
      of.line === line && of.id !== excludeId
    );
    const arrets = getArretsPlanifies().filter(a => a.line === line);
    
    const endHours = startHours + duration;
    
    // Vérifier conflits avec autres OFs
    for (const of of ofs) {
      const ofEnd = of.startHours + of.duration;
      if (startHours < ofEnd && endHours > of.startHours) {
        return { conflict: true, with: 'OF', item: of };
      }
    }

    // Vérifier conflits avec arrêts
    for (const arret of arrets) {
      const arretEnd = arret.startHours + arret.duration;
      if (startHours < arretEnd && endHours > arret.startHours) {
        return { conflict: true, with: 'Arrêt', item: arret };
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
    const codeEl = document.getElementById('planningArticleCode');
    const labelEl = document.getElementById('planningArticleLabel');
    const lineEl = document.getElementById('planningArticleLine');
    const cadenceEl = document.getElementById('planningArticleCadence');
    const poidsEl = document.getElementById('planningArticlePoids');
    
    const code = codeEl?.value?.trim()?.toUpperCase() || '';
    const label = labelEl?.value?.trim() || '';
    const line = lineEl?.value || '';
    const cadence = parseFloat(cadenceEl?.value) || 0;
    const poids = parseFloat(poidsEl?.value) || 0;

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

    const line = option.dataset.line;
    const cadence = option.dataset.cadence;

    document.getElementById('planningOFLine').value = line;
    document.getElementById('planningOFCadence').value = cadence;

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
      id: generateId(),
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

    const sorted = [...planningState.currentOFs].sort((a, b) => {
      if (a.line !== b.line) return a.line.localeCompare(b.line);
      return a.startHours - b.startHours;
    });

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

    const newDuration = calculateDuration(newQty, of.cadence);
    const newStartHours = getPositionHours(newDay, newTime);

    of.quantity = newQty;
    of.duration = newDuration;
    of.day = newDay;
    of.startTime = newTime;
    of.startHours = newStartHours;
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

  function handleDragStart(e, ofId) {
    planningState.draggedOFId = ofId;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', ofId);
    e.target.classList.add('dragging');
  }

  function handleDragEnd(e) {
    planningState.draggedOFId = null;
    e.target.classList.remove('dragging');
    document.querySelectorAll('.gantt-line-blocks').forEach(el => {
      el.classList.remove('drag-over');
    });
  }

  function handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    e.currentTarget.classList.add('drag-over');
  }

  function handleDragLeave(e) {
    e.currentTarget.classList.remove('drag-over');
  }

  function handleDrop(e, targetLine) {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    
    const ofId = planningState.draggedOFId;
    if (!ofId) return;

    // Calculer la nouvelle position en heures
    const rect = e.currentTarget.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const newStartHours = pixelsToHours(x);

    // Arrondir à 15 minutes
    const roundedHours = Math.round(newStartHours * 4) / 4;

    // Trouver l'OF
    let of = planningState.currentOFs.find(o => o.id === ofId);
    let isActive = false;
    
    if (!of && planningState.activePlanning) {
      of = planningState.activePlanning.ofs.find(o => o.id === ofId);
      isActive = true;
    }
    
    if (!of) return;

    // Vérifier les conflits (exclure l'OF qu'on déplace)
    const ofsToCheck = isActive ? planningState.activePlanning.ofs : planningState.currentOFs;
    const conflict = ofsToCheck.some(other => {
      if (other.id === ofId) return false;
      if (other.line !== targetLine) return false;
      const otherEnd = other.startHours + other.duration;
      return roundedHours < otherEnd && (roundedHours + of.duration) > other.startHours;
    });

    if (conflict) {
      alert('Position occupée par un autre OF !');
      return;
    }

    // Vérifier les limites
    if (roundedHours < 0 || roundedHours + of.duration > TOTAL_HOURS) {
      alert('Position hors limites !');
      return;
    }

    // Mettre à jour l'OF
    of.line = targetLine;
    of.startHours = roundedHours;
    const dayTime = positionToDayTime(roundedHours);
    of.day = dayTime.day;
    of.startTime = dayTime.time;

    savePlanningState();
    
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
  // GANTT CHART AMÉLIORÉ
  // ==========================================

  function renderGantt(containerId, ofs, interactive = false) {
    const container = document.getElementById(containerId);
    if (!container) return;

    // Récupérer les lignes depuis l'état global
    const lines = window.LINES || ['Râpé', 'T2', 'OMORI', 'T1', 'Emballage', 'Dés', 'Filets', 'Prédécoupé'];
    
    // Récupérer les arrêts
    const arrets = interactive ? getArretsPlanifies() : [];
    
    // Si planning actif, mettre à jour depuis la production
    if (interactive && planningState.activePlanning) {
      updateOFsFromProduction();
    }

    if (!ofs || ofs.length === 0) {
      container.innerHTML = `<div class="gantt-empty">${
        containerId === 'planningGantt' 
          ? 'Sélectionnez un planning pour l\'afficher'
          : 'Ajoutez des OFs pour voir la prévisualisation'
      }</div>`;
      return;
    }

    // Filtrer les lignes avec des OFs ou des arrêts
    const usedLines = [...new Set([
      ...ofs.map(of => of.line),
      ...arrets.map(a => a.line)
    ])];

    // Trier les lignes selon l'ordre original
    const sortedLines = lines.filter(l => usedLines.includes(l));

    let html = `<div class="gantt-chart-v3" style="min-width: ${GANTT_WIDTH_PX + 140}px">`;
    
    // En-tête avec jours et heures
    html += '<div class="gantt-header-v3">';
    html += '<div class="gantt-corner">Lignes</div>';
    html += '<div class="gantt-time-header">';
    
    // Jours
    DAYS.forEach((day, idx) => {
      const dayHours = idx < 5 ? 24 : 12;
      const width = hoursToPixels(dayHours);
      const left = hoursToPixels(idx * 24);
      html += `<div class="gantt-day-label" style="left: ${left}px; width: ${width}px">${day}</div>`;
    });
    
    // Heures (tous les 6h)
    for (let h = 0; h <= TOTAL_HOURS; h += 6) {
      const left = hoursToPixels(h);
      const hourOfDay = h % 24;
      html += `<div class="gantt-hour-tick" style="left: ${left}px">${String(hourOfDay).padStart(2, '0')}h</div>`;
    }
    
    html += '</div></div>';

    // Lignes du Gantt
    sortedLines.forEach(line => {
      const lineOFs = ofs.filter(of => of.line === line);
      const lineArrets = arrets.filter(a => a.line === line);
      
      html += `<div class="gantt-row-v3">`;
      html += `<div class="gantt-line-label-v3">${line}</div>`;
      html += `<div class="gantt-line-blocks-v3" data-line="${line}">`;
      
      // Grille de fond (jours)
      for (let d = 0; d < 6; d++) {
        const left = hoursToPixels(d * 24);
        const width = hoursToPixels(d < 5 ? 24 : 12);
        html += `<div class="gantt-day-bg ${d % 2 === 0 ? 'even' : 'odd'}" style="left: ${left}px; width: ${width}px"></div>`;
      }
      
      // Marqueur "maintenant"
      const nowHours = getCurrentHoursInWeek();
      if (interactive && nowHours !== null && nowHours >= 0 && nowHours <= TOTAL_HOURS) {
        const nowLeft = hoursToPixels(nowHours);
        html += `<div class="gantt-now-marker" style="left: ${nowLeft}px" data-now-marker></div>`;
      }
      
      // Arrêts
      lineArrets.forEach(arret => {
        const left = hoursToPixels(arret.startHours);
        const width = Math.max(hoursToPixels(arret.duration), 20);
        html += `<div class="gantt-block-v3 stop" style="left: ${left}px; width: ${width}px">
          <div class="gantt-block-title">🛑 ${arret.label}</div>
          <div class="gantt-block-meta">${formatDuration(arret.duration)}</div>
        </div>`;
      });
      
      // OFs
      lineOFs.forEach(of => {
        const left = hoursToPixels(of.startHours);
        const width = Math.max(hoursToPixels(of.duration), 40);
        const endPos = positionToDayTime(of.startHours + of.duration);
        const progress = of.quantity > 0 ? Math.min(100, (of.produced || 0) / of.quantity * 100) : 0;
        
        // Déterminer le statut visuel
        let statusClass = of.status;
        if (interactive && nowHours !== null) {
          if (of.status !== 'done' && of.startHours + of.duration < nowHours && progress < 100) {
            statusClass = 'late';
          }
        }
        
        html += `<div class="gantt-block-v3 ${statusClass}" 
          style="left: ${left}px; width: ${width}px"
          data-of-id="${of.id}"
          ${interactive ? 'draggable="true"' : ''}>
          ${progress > 0 ? `<div class="gantt-progress" style="width: ${progress}%"></div>` : ''}
          <div class="gantt-block-content">
            <div class="gantt-block-title">${of.articleCode}</div>
            <div class="gantt-block-meta">${of.quantity} colis</div>
            <div class="gantt-block-time">${of.startTime} → ${endPos.time}</div>
            ${of.produced > 0 ? `<div class="gantt-block-produced">✓ ${of.produced}</div>` : ''}
          </div>
        </div>`;
      });
      
      html += '</div></div>';
    });
    
    html += '</div>';
    container.innerHTML = html;

    // Ajouter les event listeners
    if (interactive) {
      // Drag & Drop
      container.querySelectorAll('.gantt-block-v3[draggable="true"]').forEach(block => {
        const ofId = block.dataset.ofId;
        block.addEventListener('dragstart', (e) => handleDragStart(e, ofId));
        block.addEventListener('dragend', handleDragEnd);
        block.addEventListener('click', () => openOFEditor(ofId));
      });
      
      container.querySelectorAll('.gantt-line-blocks-v3').forEach(lineBlocks => {
        const line = lineBlocks.dataset.line;
        lineBlocks.addEventListener('dragover', handleDragOver);
        lineBlocks.addEventListener('dragleave', handleDragLeave);
        lineBlocks.addEventListener('drop', (e) => handleDrop(e, line));
      });
    }
  }

  function renderPreviewGantt() {
    renderGantt('planningPreviewGantt', planningState.currentOFs, false);
  }

  function renderActiveGantt() {
    if (!planningState.activePlanning) {
      const container = document.getElementById('planningGantt');
      if (container) {
        container.innerHTML = '<div class="gantt-empty">Sélectionnez un planning pour l\'afficher</div>';
      }
      return;
    }
    renderGantt('planningGantt', planningState.activePlanning.ofs, true);
  }

  // Mettre à jour le marqueur "maintenant" en temps réel
  function updateNowMarkers() {
    const nowHours = getCurrentHoursInWeek();
    if (nowHours === null) return;

    document.querySelectorAll('[data-now-marker]').forEach(marker => {
      const left = hoursToPixels(nowHours);
      marker.style.left = `${left}px`;
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
      ofs: JSON.parse(JSON.stringify(planningState.currentOFs))
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
        .map(p => `<option value="${p.id}">Semaine ${p.weekNumber} (${p.weekStart})</option>`)
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
    planningState.weekNumber = planning.weekNumber;
    planningState.weekStart = planning.weekStart;
    
    // Mettre à jour depuis la production
    updateOFsFromProduction();
    
    savePlanningState();

    const activeInfo = document.getElementById('planningActiveInfo');
    const activeWeek = document.getElementById('planningActiveWeek');
    if (activeInfo) activeInfo.classList.remove('hidden');
    if (activeWeek) activeWeek.textContent = planning.weekNumber;

    renderActiveGantt();
    renderActiveOFList();
    updateDelays();
    
    // Démarrer le timer pour le marqueur temps réel
    startNowMarkerTimer();
  }

  function startNowMarkerTimer() {
    if (nowMarkerInterval) clearInterval(nowMarkerInterval);
    nowMarkerInterval = setInterval(updateNowMarkers, 60000); // Mise à jour toutes les minutes
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

    // Mettre à jour depuis la production
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
        
        // Calcul de l'attendu
        const ofEnd = of.startHours + of.duration;
        if (nowHours >= ofEnd) {
          totalExpected += of.quantity;
        } else if (nowHours > of.startHours) {
          const elapsed = nowHours - of.startHours;
          const ratio = elapsed / of.duration;
          totalExpected += Math.round(of.quantity * ratio);
        }
      });

      const delta = totalProduced - totalExpected;
      const deltaClass = delta > 0 ? 'positive' : delta < 0 ? 'negative' : 'neutral';
      const deltaText = delta > 0 ? `+${delta}` : delta.toString();
      const pct = totalExpected > 0 ? Math.round(totalProduced / totalExpected * 100) : 0;

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
              <span class="delay-stat-value ${deltaClass}">${deltaText} (${pct}%)</span>
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
    
    ['planningArticleLine', 'planningOFLine'].forEach(id => {
      const select = document.getElementById(id);
      if (select) {
        select.innerHTML = lines.map(l => `<option value="${l}">${l}</option>`).join('');
      }
    });
  }

  function initWeekInputs() {
    const weekNumInput = document.getElementById('planningWeekNumber');
    const weekStartInput = document.getElementById('planningWeekStart');

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
        
        // Rafraîchir le Gantt actif quand on va sur l'onglet
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

    // Semaine
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

    // Validation
    document.getElementById('planningValidateBtn')?.addEventListener('click', validatePlanning);

    // Lancement
    document.getElementById('planningLaunchBtn')?.addEventListener('click', launchPlanning);

    // Éditeur OF
    document.getElementById('planningOFEditClose')?.addEventListener('click', closeOFEditor);
    document.getElementById('planningEditOFSave')?.addEventListener('click', saveOFEdit);
    document.getElementById('planningEditOFDelete')?.addEventListener('click', deleteOFFromEditor);

    // Sauvegarder semaine
    document.getElementById('planningWeekNumber')?.addEventListener('change', (e) => {
      planningState.weekNumber = parseInt(e.target.value);
      savePlanningState();
    });
    document.getElementById('planningWeekStart')?.addEventListener('change', (e) => {
      planningState.weekStart = e.target.value;
      savePlanningState();
    });

    // Bouton rafraîchir
    document.getElementById('planningRefreshBtn')?.addEventListener('click', () => {
      updateOFsFromProduction();
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
    });
  }

  function init() {
    console.log('Initialisation module Planning v3...');
    
    loadPlanningState();
    initLineSelects();
    initWeekInputs();
    bindEvents();
    
    renderArticlesList();
    updateOFFormVisibility();
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

    console.log('Module Planning v3 initialisé.');
  }

  // API publique
  return {
    init,
    openOFEditor,
    refreshFromProduction: () => {
      updateOFsFromProduction();
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
    },
    getState: () => planningState
  };
})();

// Initialiser
document.addEventListener('DOMContentLoaded', function() {
  setTimeout(PlanningModule.init, 500);
});
