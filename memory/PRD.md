# Atelier PPNC - Product Requirements Document

## Problème Original
Application PWA de gestion de production pour un atelier de transformation fromagère. L'utilisateur a demandé une refonte complète du module Planning inspiré de SAP.

## Architecture Technique
- **Stack** : PWA statique (HTML, CSS, Vanilla JavaScript)
- **Stockage** : localStorage (clé `planning_v4`)
- **Fichiers principaux** :
  - `/app/index.html` - Interface utilisateur
  - `/app/style.css` - Styles
  - `/app/app.js` - Logique principale
  - `/app/planning.js` - Module Planning V4
  - `/app/frontend/` - Copie pour le serveur

## Fonctionnalités Implémentées

### Module Planning V4 (Janvier 2026)

#### Articles
- CRUD complet pour les articles de production
- Propriétés : Code, Libellé, Ligne, Cadence (colis/h), Poids (kg)

#### Arrêts Planifiés
- Types : Pause, Maintenance, etc.
- Configuration : Ligne (ou toutes), Jour (ou tous), Heure, Durée
- **Les arrêts coupent les OFs et décalent leur fin**

#### Changements (Nouveau - 21/01/2026)
- Types : Intermédiaire (15min), Format (30min), Produit (45min), Couleur (60min)
- Affichage sur le Gantt avec couleurs distinctes
- Peuvent être coupés par arrêts, ne coupent pas les OFs

#### Ordres de Fabrication (OFs)
- Création avec calcul automatique de durée
- Affichage en segments quand coupés par arrêts
- Drag & drop (OF reste sur sa ligne)

#### Diagramme de Gantt
- Échelle : Lundi 00h → Samedi 12h (132 heures)
- OFs en segments colorés selon statut
- Arrêts en jaune (⏸️)
- Changements en couleur selon type
- Marqueur "Maintenant" en temps réel
- Drag & drop interactif

#### Planning Actif
- Table Avance/Retard par ligne
- Intégration temps réel avec Production/Arrêts (window.state)

### Corrections antérieures
- Calculatrice en pop-up moderne
- Toggle thème sombre/clair
- Calcul automatique lundi selon n° de semaine

## Credentials
- **Mot de passe Manager** : 3005

## Backlog

### P0 - Critique
- [ ] Tester drag & drop manuellement

### P1 - Haute
- [ ] Règles drag & drop changements
- [ ] Export planning (Excel/PDF)

### P2 - Moyenne
- [ ] Notifications OFs en retard
- [ ] Historique plannings

### P3 - Future
- [ ] Mode hors-ligne amélioré
- [ ] Import/Export configuration

## Notes Techniques

### Calcul segments OF
Quand un arrêt chevauche un OF :
1. Segment 1 : OF start → arrêt start
2. Arrêt visible entre les segments
3. Segment 2 : arrêt end → OF end + durée arrêt

### localStorage
- `planning_v4` : État complet Planning
- `state` : État app.js (exposé via window.state)

### Service Worker
- Cache name : `atelier-ppnc-v2`
- Assets : index.html, style.css, app.js, planning.js
