/********************************************
 *   MODULE PLANNING - VERSION 2.0
 *   Système de planification de production
 ********************************************/

const PlanningModule = (function() {
  // Configuration
  const DAYS = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
  const HOURS_PER_DAY = 24;
  const TOTAL_HOURS = 5 * 24 + 12; // Lun 00:00 -> Sam 12:00

  // État local du planning
  let planningState = {
    articles: [],
    currentOFs: [],
    savedPlannings: [],
    activePlanning: null,
    weekNumber: null,
    weekStart: null,
    editingOFId: null
  };

  // Charger l'état depuis localStorage
  function loadPlanningState() {
    try {
      const saved = localStorage.getItem('planning_v2');
      if (saved) {
        const data = JSON.parse(saved);
        planningState = { ...planningState, ...data };
      }
    } catch (e) {
      console.error('Erreur chargement planning:', e);
    }
    // Initialiser la semaine actuelle si pas définie
    if (!planningState.weekStart) {
      setCurrentWeek();
    }
  }

  // Sauvegarder l'état
  function savePlanningState() {
    try {
      localStorage.setItem('planning_v2', JSON.stringify(planningState));
    } catch (e) {
      console.error('Erreur sauvegarde planning:', e);
    }
  }

  // Obtenir le lundi de la semaine
  function getMonday(date = new Date()) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    d.setDate(diff);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // Obtenir le numéro de semaine ISO
  function getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  }

  // Définir la semaine actuelle
  function setCurrentWeek() {
    const monday = getMonday();
    planningState.weekStart = monday.toISOString().split('T')[0];
    planningState.weekNumber = getWeekNumber(monday);
    savePlanningState();
  }

  // Générer un ID unique
  function generateId() {
    return 'of_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  // Formater l'heure
  function formatTime(hours, minutes = 0) {
    return `${String(Math.floor(hours)).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }

  // Calculer la durée en heures à partir de quantité et cadence
  function calculateDuration(quantity, cadence) {
    if (!cadence || cadence <= 0) return 1;
    return quantity / cadence;
  }

  // Convertir jour + heure en position (heures depuis lundi 00:00)
  function getPositionHours(dayIndex, timeStr) {
    const [h, m] = (timeStr || '06:00').split(':').map(Number);
    return dayIndex * 24 + h + (m / 60);
  }

  // Convertir position en jour + heure
  function positionToDayTime(hours) {
    const day = Math.floor(hours / 24);
    const hourOfDay = hours % 24;
    const h = Math.floor(hourOfDay);
    const m = Math.round((hourOfDay - h) * 60);
    return {
      day: Math.min(day, 5),
      time: formatTime(h, m)
    };
  }

  // Trouver le prochain créneau disponible pour une ligne
  function findNextAvailableSlot(line, duration, startFromHours = 6) {
    const ofs = planningState.currentOFs.filter(of => of.line === line);
    
    if (ofs.length === 0) {
      return { day: 0, time: formatTime(startFromHours) };
    }

    // Trier les OFs par position de début
    ofs.sort((a, b) => a.startHours - b.startHours);

    // Chercher un créneau libre
    let cursor = startFromHours;
    
    for (const of of ofs) {
      if (cursor + duration <= of.startHours) {
        // Espace trouvé avant cet OF
        return positionToDayTime(cursor);
      }
      cursor = Math.max(cursor, of.startHours + of.duration);
    }

    // Placer après le dernier OF
    if (cursor + duration <= TOTAL_HOURS) {
      return positionToDayTime(cursor);
    }

    // Pas assez de place
    return null;
  }

  // Vérifier les conflits
  function checkConflict(line, startHours, duration, excludeId = null) {
    const ofs = planningState.currentOFs.filter(of => 
      of.line === line && of.id !== excludeId
    );
    
    const endHours = startHours + duration;
    
    for (const of of ofs) {
      const ofEnd = of.startHours + of.duration;
      if (startHours < ofEnd && endHours > of.startHours) {
        return true; // Conflit détecté
      }
    }
    return false;
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
          <button class="secondary-btn btn-edit-article" data-code="${article.code}">✏️ Modifier</button>
          <button class="danger-btn btn-delete-article" data-code="${article.code}">🗑️</button>
        </div>
      </div>
    `).join('');

    // Event listeners
    container.querySelectorAll('.btn-edit-article').forEach(btn => {
      btn.addEventListener('click', () => editArticle(btn.dataset.code));
    });
    container.querySelectorAll('.btn-delete-article').forEach(btn => {
      btn.addEventListener('click', () => deleteArticle(btn.dataset.code));
    });

    // Mettre à jour la liste déroulante des articles dans l'onglet OF
    updateArticleSelect();
  }

  function addOrUpdateArticle() {
    console.log('addOrUpdateArticle called');
    const codeEl = document.getElementById('planningArticleCode');
    const labelEl = document.getElementById('planningArticleLabel');
    const lineEl = document.getElementById('planningArticleLine');
    const cadenceEl = document.getElementById('planningArticleCadence');
    const poidsEl = document.getElementById('planningArticlePoids');
    
    console.log('Elements:', { codeEl, labelEl, lineEl, cadenceEl, poidsEl });
    
    const code = codeEl?.value?.trim()?.toUpperCase() || '';
    const label = labelEl?.value?.trim() || '';
    const line = lineEl?.value || '';
    const cadence = parseFloat(cadenceEl?.value) || 0;
    const poids = parseFloat(poidsEl?.value) || 0;

    console.log('Values:', { code, label, line, cadence, poids });

    if (!code || !label || !line || cadence <= 0) {
      alert('Veuillez remplir tous les champs obligatoires (code, libellé, ligne, cadence).');
      return;
    }

    const existingIndex = planningState.articles.findIndex(a => a.code === code);
    const article = { code, label, line, cadence, poids };

    if (existingIndex >= 0) {
      planningState.articles[existingIndex] = article;
    } else {
      planningState.articles.push(article);
    }
    
    console.log('Article added:', article);
    console.log('All articles:', planningState.articles);

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

    select.innerHTML = '<option value="">-- Sélectionner un article --</option>' +
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

      if (durationEl) {
        const h = Math.floor(duration);
        const m = Math.round((duration - h) * 60);
        durationEl.textContent = `${h}h${m > 0 ? m.toString().padStart(2, '0') : ''}`;
      }
      if (endEl) {
        endEl.textContent = `${DAYS[endPos.day]} ${endPos.time}`;
      }
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
      nextSlotEl.textContent = 'Aucun créneau disponible';
    }
  }

  function autoPlaceOF() {
    const line = document.getElementById('planningOFLine')?.value;
    const qty = parseInt(document.getElementById('planningOFQty')?.value) || 0;
    const cadence = parseFloat(document.getElementById('planningOFCadence')?.value) || 0;

    if (!line || qty <= 0 || cadence <= 0) {
      alert('Sélectionnez un article et entrez une quantité d\'abord.');
      return;
    }

    const duration = calculateDuration(qty, cadence);
    const slot = findNextAvailableSlot(line, duration);

    if (slot) {
      document.getElementById('planningOFDay').value = slot.day;
      document.getElementById('planningOFStart').value = slot.time;
      updateOFPreview();
    } else {
      alert('Aucun créneau disponible sur cette ligne pour cette durée.');
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
    if (!article) {
      alert('Article non trouvé.');
      return;
    }

    const duration = calculateDuration(qty, cadence);
    const startHours = getPositionHours(day, startTime);

    // Vérifier les conflits
    if (checkConflict(line, startHours, duration)) {
      alert('Conflit détecté ! Un autre OF occupe déjà ce créneau sur cette ligne.');
      return;
    }

    // Vérifier les limites
    if (startHours + duration > TOTAL_HOURS) {
      alert('L\'OF dépasse la fin de la semaine (samedi 12:00).');
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
      container.innerHTML = '<p class="helper-text">Aucun OF ajouté. Créez des OFs ci-dessus.</p>';
      return;
    }

    // Trier par ligne puis par heure de début
    const sorted = [...planningState.currentOFs].sort((a, b) => {
      if (a.line !== b.line) return a.line.localeCompare(b.line);
      return a.startHours - b.startHours;
    });

    container.innerHTML = sorted.map(of => {
      const endHours = of.startHours + of.duration;
      const endPos = positionToDayTime(endHours);
      const durationH = Math.floor(of.duration);
      const durationM = Math.round((of.duration - durationH) * 60);

      return `
        <div class="of-card" data-id="${of.id}">
          <div class="of-card-status ${of.status}"></div>
          <div class="of-card-info">
            <div class="of-card-title">${of.articleCode} - ${of.articleLabel}</div>
            <div class="of-card-details">${of.line} • ${of.quantity} colis • ${durationH}h${durationM > 0 ? durationM : ''}</div>
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

    // Event listeners
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
    if (!confirm('Supprimer tous les OFs du planning en cours ?')) return;
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

    // Chercher dans les OFs courants ou actifs
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

    // Vérifier les conflits
    const ofsToCheck = isActive ? planningState.activePlanning.ofs : planningState.currentOFs;
    const hasConflict = ofsToCheck.some(other => {
      if (other.id === id || other.line !== of.line) return false;
      const otherEnd = other.startHours + other.duration;
      return newStartHours < otherEnd && (newStartHours + newDuration) > other.startHours;
    });

    if (hasConflict) {
      alert('Conflit détecté avec un autre OF !');
      return;
    }

    // Mettre à jour l'OF
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
    if (!id) return;

    if (!confirm('Supprimer cet OF ?')) return;

    // Chercher et supprimer
    let found = false;
    const idx = planningState.currentOFs.findIndex(o => o.id === id);
    if (idx >= 0) {
      planningState.currentOFs.splice(idx, 1);
      found = true;
    }

    if (!found && planningState.activePlanning) {
      const activeIdx = planningState.activePlanning.ofs.findIndex(o => o.id === id);
      if (activeIdx >= 0) {
        planningState.activePlanning.ofs.splice(activeIdx, 1);
        found = true;
      }
    }

    savePlanningState();
    closeOFEditor();
    
    renderOFList();
    renderPreviewGantt();
    renderActiveGantt();
    renderActiveOFList();
  }

  // ==========================================
  // GANTT CHART
  // ==========================================

  function renderGantt(containerId, ofs, interactive = false) {
    const container = document.getElementById(containerId);
    if (!container) return;

    if (!ofs || ofs.length === 0) {
      container.innerHTML = `<div class="gantt-empty">${
        containerId === 'planningGantt' 
          ? 'Sélectionnez un planning pour l\'afficher'
          : 'Ajoutez des OFs pour voir la prévisualisation'
      }</div>`;
      return;
    }

    // Obtenir les lignes utilisées
    const lines = [...new Set(ofs.map(of => of.line))].sort();

    // Créer le Gantt
    let html = '<div class="gantt-chart">';
    
    // En-tête avec jours
    html += '<div class="gantt-header">';
    html += '<div class="gantt-header-label">Lignes</div>';
    html += '<div class="gantt-time-axis">';
    
    DAYS.forEach((day, idx) => {
      const width = idx < 5 ? (24 / TOTAL_HOURS * 100) : (12 / TOTAL_HOURS * 100);
      html += `<div class="gantt-day-header" style="width: ${width}%">
        ${day}
        <div class="gantt-hours">
          ${idx < 5 ? '06h | 12h | 18h | 00h' : '06h | 12h'}
        </div>
      </div>`;
    });
    
    html += '</div></div>';

    // Lignes
    lines.forEach(line => {
      const lineOFs = ofs.filter(of => of.line === line);
      
      html += `<div class="gantt-line-row">
        <div class="gantt-line-name">${line}</div>
        <div class="gantt-line-blocks">`;
      
      lineOFs.forEach(of => {
        const leftPct = (of.startHours / TOTAL_HOURS) * 100;
        const widthPct = (of.duration / TOTAL_HOURS) * 100;
        const endPos = positionToDayTime(of.startHours + of.duration);
        
        html += `<div class="gantt-of-block ${of.status}" 
          style="left: ${leftPct}%; width: ${Math.max(widthPct, 1)}%"
          data-id="${of.id}"
          ${interactive ? 'onclick="PlanningModule.openOFEditor(\'' + of.id + '\')"' : ''}>
          <div class="gantt-of-title">${of.articleCode}</div>
          <div class="gantt-of-meta">${of.quantity} colis • ${of.startTime}-${endPos.time}</div>
        </div>`;
      });
      
      html += '</div></div>';
    });
    
    html += '</div>';
    container.innerHTML = html;
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

  // ==========================================
  // VALIDATION ET LANCEMENT
  // ==========================================

  function validatePlanning() {
    if (planningState.currentOFs.length === 0) {
      alert('Ajoutez au moins un OF avant de valider le planning.');
      return;
    }

    const weekNum = document.getElementById('planningWeekNumber')?.value;
    const weekStart = document.getElementById('planningWeekStart')?.value;

    if (!weekNum || !weekStart) {
      alert('Définissez la semaine et le lundi de référence.');
      return;
    }

    // Vérifier si un planning existe déjà pour cette semaine
    const existingIdx = planningState.savedPlannings.findIndex(p => p.weekNumber == weekNum);
    
    if (existingIdx >= 0) {
      if (!confirm(`Un planning existe déjà pour la semaine ${weekNum}. Le remplacer ?`)) {
        return;
      }
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

    // Vider les OFs courants
    planningState.currentOFs = [];
    
    savePlanningState();
    
    alert(`Planning semaine ${weekNum} validé avec succès !`);
    
    renderOFList();
    renderPreviewGantt();
    updatePlanningSelect();
  }

  function updatePlanningSelect() {
    const select = document.getElementById('planningSelectWeek');
    if (!select) return;

    select.innerHTML = '<option value="">-- Sélectionner un planning --</option>' +
      planningState.savedPlannings
        .sort((a, b) => b.weekNumber - a.weekNumber)
        .map(p => `<option value="${p.id}">Semaine ${p.weekNumber} (${p.weekStart})</option>`)
        .join('');
  }

  function launchPlanning() {
    const select = document.getElementById('planningSelectWeek');
    const planningId = select?.value;

    if (!planningId) {
      alert('Sélectionnez un planning à charger.');
      return;
    }

    const planning = planningState.savedPlannings.find(p => p.id === planningId);
    if (!planning) {
      alert('Planning non trouvé.');
      return;
    }

    planningState.activePlanning = JSON.parse(JSON.stringify(planning));
    planningState.weekNumber = planning.weekNumber;
    planningState.weekStart = planning.weekStart;
    
    savePlanningState();

    // Mettre à jour l'affichage
    const activeInfo = document.getElementById('planningActiveInfo');
    const activeWeek = document.getElementById('planningActiveWeek');
    if (activeInfo) activeInfo.classList.remove('hidden');
    if (activeWeek) activeWeek.textContent = planning.weekNumber;

    renderActiveGantt();
    renderActiveOFList();
    updateDelays();
  }

  function renderActiveOFList() {
    const container = document.getElementById('planningActiveOFList');
    if (!container) return;

    if (!planningState.activePlanning || planningState.activePlanning.ofs.length === 0) {
      container.innerHTML = '<p class="helper-text">Aucun planning chargé.</p>';
      return;
    }

    const sorted = [...planningState.activePlanning.ofs].sort((a, b) => {
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
            <div class="of-card-details">
              ${of.line} • ${of.quantity} colis 
              ${of.produced > 0 ? `• Produit: ${of.produced}` : ''}
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
      container.innerHTML = '<p class="helper-text">Chargez un planning pour voir l\'avance/retard.</p>';
      return;
    }

    const lines = [...new Set(planningState.activePlanning.ofs.map(of => of.line))];
    const now = new Date();

    container.innerHTML = lines.map(line => {
      const lineOFs = planningState.activePlanning.ofs.filter(of => of.line === line);
      
      let totalPlanned = 0;
      let totalProduced = 0;
      let totalExpected = 0;

      lineOFs.forEach(of => {
        totalPlanned += of.quantity;
        totalProduced += of.produced || 0;
        
        // Calculer la quantité attendue à ce moment
        const startDate = new Date(planningState.weekStart);
        startDate.setHours(Math.floor(of.startHours % 24), (of.startHours % 1) * 60);
        startDate.setDate(startDate.getDate() + of.day);
        
        const endDate = new Date(startDate.getTime() + of.duration * 3600000);
        
        if (now >= endDate) {
          totalExpected += of.quantity;
        } else if (now > startDate) {
          const elapsed = (now - startDate) / (endDate - startDate);
          totalExpected += Math.round(of.quantity * elapsed);
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
              <span class="delay-stat-label">Produit</span>
              <span class="delay-stat-value">${totalProduced}</span>
            </div>
            <div class="delay-stat">
              <span class="delay-stat-label">Attendu</span>
              <span class="delay-stat-value">${totalExpected}</span>
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
    // Récupérer les lignes depuis l'état global ou utiliser des valeurs par défaut
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
      });
    });

    // Articles
    document.getElementById('planningArticleAddBtn')?.addEventListener('click', addOrUpdateArticle);
    document.getElementById('planningArticleClearBtn')?.addEventListener('click', clearArticleForm);

    // Semaine auto
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

    // Sauvegarde semaine
    document.getElementById('planningWeekNumber')?.addEventListener('change', (e) => {
      planningState.weekNumber = parseInt(e.target.value);
      savePlanningState();
    });
    document.getElementById('planningWeekStart')?.addEventListener('change', (e) => {
      planningState.weekStart = e.target.value;
      savePlanningState();
    });
  }

  function init() {
    console.log('Initialisation module Planning v2...');
    
    loadPlanningState();
    initLineSelects();
    initWeekInputs();
    bindEvents();
    
    // Rendu initial
    renderArticlesList();
    updateOFFormVisibility();
    renderOFList();
    renderPreviewGantt();
    updatePlanningSelect();
    
    // Si un planning actif existe
    if (planningState.activePlanning) {
      const activeInfo = document.getElementById('planningActiveInfo');
      const activeWeek = document.getElementById('planningActiveWeek');
      if (activeInfo) activeInfo.classList.remove('hidden');
      if (activeWeek) activeWeek.textContent = planningState.activePlanning.weekNumber;
      
      renderActiveGantt();
      renderActiveOFList();
      updateDelays();
    }

    console.log('Module Planning v2 initialisé.');
  }

  // API publique
  return {
    init,
    openOFEditor,
    // Expose pour debug
    getState: () => planningState
  };
})();

// Initialiser quand le DOM est prêt
document.addEventListener('DOMContentLoaded', function() {
  // Attendre un peu que tout soit chargé
  setTimeout(function() {
    PlanningModule.init();
  }, 500);
});
