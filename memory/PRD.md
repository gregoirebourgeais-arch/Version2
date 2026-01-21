# Atelier PPNC - Production Requirements Document

## Problème Original
Application PWA de gestion de production pour un atelier de transformation fromagère. L'utilisateur a demandé une refonte complète du module Planning inspiré de SAP, avec des fonctionnalités avancées de planification de production.

## Architecture Technique
- **Stack** : PWA statique (HTML, CSS, Vanilla JavaScript)
- **Stockage** : localStorage (clé `planning_v4`)
- **Fichiers principaux** :
  - `/app/index.html` - Interface utilisateur
  - `/app/style.css` - Styles (thème sombre/clair)
  - `/app/app.js` - Logique principale (Production, Arrêts, Manager)
  - `/app/planning.js` - Module Planning V4
  - `/app/service-worker.js` - PWA

## Fonctionnalités Implémentées

### ✅ Module Planning V4 (Complété le 21/01/2026)

#### Articles
- CRUD complet pour les articles de production
- Propriétés : Code, Libellé, Ligne de production, Cadence (colis/h), Poids (kg)
- Persistance via localStorage

#### Création de Planning
- Configuration semaine (numéro, date de début)
- **Arrêts Planifiés** : Pauses, maintenance, etc.
  - Type, Ligne (ou toutes), Jour (ou tous), Heure, Durée, Commentaire
- **Ordres de Fabrication (OFs)** :
  - Sélection article, quantité, jour, heure de début
  - Calcul automatique de la durée basé sur la cadence
  - Bouton "Placer automatiquement" pour trouver un créneau

#### Planning Actif
- Sélection et chargement d'un planning validé
- **Diagramme de Gantt** avec :
  - Échelle de temps précise (Lundi 00h → Samedi 12h)
  - Blocs OFs colorés selon statut (planifié, en cours, terminé, en retard)
  - Arrêts planifiés en orange
  - Arrêts non planifiés (depuis app.js) en rouge
  - **Marqueur "Maintenant"** en temps réel sur chaque ligne
  - **Drag & Drop** pour repositionner les OFs
  - Double-clic pour éditer un OF
- **Table Avance/Retard** par ligne :
  - Planifié, Attendu, Produit, Écart

#### Intégration Temps Réel
- Connexion avec `window.state.production` (données de production)
- Connexion avec `window.state.arrets` (arrêts non planifiés)
- Mise à jour automatique des statuts OFs
- Bouton "Rafraîchir depuis Production/Arrêts"

### ✅ Autres Corrections (21/01/2026)
- **Calculatrice** : Pop-up moderne
- **Toggle Thème** : Sombre/Clair fonctionnel
- **Onglet Manager** : Modal mot de passe (3005)

## Credentials
- **Mot de passe Manager** : 3005

## Backlog / Tâches Futures

### P1 - Priorité Haute
- [ ] Tests automatisés pour le module Planning
- [ ] Export du planning (Excel/PDF)

### P2 - Priorité Moyenne
- [ ] Notifications pour OFs en retard
- [ ] Historique des plannings validés
- [ ] Statistiques de production vs planning

### P3 - Améliorations
- [ ] Mode hors-ligne amélioré (sync)
- [ ] Import/Export de configuration articles
- [ ] Personnalisation des lignes de production

## Notes Techniques

### localStorage Keys
- `planning_v4` : État complet du module Planning
- `state` : État principal de l'application (app.js)

### Intégration app.js ↔ planning.js
```javascript
// app.js expose l'état globalement
window.state = state;

// planning.js lit les données
window.state.production  // Données de production par ligne
window.state.arrets      // Arrêts non planifiés
```

### Configuration Gantt
- 132 heures totales (Lundi 00h → Samedi 12h)
- 12 pixels par heure
- Arrondi des positions à 15 minutes
