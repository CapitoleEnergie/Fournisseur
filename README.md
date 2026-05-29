# Sélection Fournisseurs — Capitole Énergie

Application Vercel de sélection, classification et suivi des fournisseurs d'énergie.

---

## 🧭 Objectif

Cette application permet de :

- Classifier dynamiquement tous les fournisseurs selon le profil d'un dossier client (énergie, segment, note, volume, DDF/DFF, syndic…)
- Visualiser les règles d'octroi issues du fichier SharePoint en temps réel
- Consulter le panel fournisseurs (Gold Premium, Gold, Silver, Bronze)
- Suivre les news et challenges fournisseurs via ClickUp
- Enregistrer les transmissions de dossiers avec référence et commentaire vers le Hub Capitole

---

## 🏗️ Architecture

```
Frontend (public/index.html)
         │
         ├── Auth Guard (public/auth-guard.js)
         │         └── Hub Capitole (hub-capitole.vercel.app)
         │                   └── /api/verify-app-token  ← vérifie le token
         │                   └── /api/log-transmis      ← enregistre les transmissions
         │
         ├── /api/fournisseur-rules      ← règles d'octroi globales (SharePoint)
         ├── /api/fournisseur-selection  ← moteur de sélection + panel (SharePoint)
         ├── /api/fournisseur-pricing    ← données prix par segment (SharePoint)
         └── /api/clickup               ← news & challenges (ClickUp)
```

---

## 🔐 Authentification (auth-guard)

Fichier : `public/auth-guard.js`

Chaque page vérifie un token à l'ouverture :

1. L'utilisateur arrive avec un `?token=xxx` dans l'URL (généré par le Hub)
2. Le token est vérifié via `POST https://hub-capitole.vercel.app/api/verify-app-token`
3. En cas de succès, une session est sauvegardée en `sessionStorage` (validité 8h)
4. En cas d'échec, redirection automatique vers le Hub

La session est stockée sous la clé `cap_auth_fournisseur` et contient l'email de l'utilisateur.

---

## 📊 Fichiers SharePoint — Source de vérité

Tous les fichiers sont lus via l'API **Microsoft Graph** (OAuth2 client credentials).

| Fichier | Onglet | Endpoint | Usage |
|---|---|---|---|
| `regles_all_fournisseurs.xlsx` | `Résumé global` | `/api/fournisseur-rules` | Règles d'octroi globales (résumé) |
| `regles_panel_fournisseurs.xlsx` | `Fournisseurs Export` + `Fournisseurs Panel` | `/api/fournisseur-selection` | Règles détaillées par fournisseur + panel |
| `Chiffres_fournisseurs.xlsx` | `1.Top fournisseurs C4` | `/api/fournisseur-pricing` | Prix marché par segment (30 derniers jours) |

Chemin SharePoint : `PARTAGE/Team/3. Fournisseurs/Data Fournisseurs`

Un **cache mémoire de 5 minutes** est appliqué sur chaque endpoint pour éviter de re-télécharger le fichier à chaque requête.

---

## 🧠 Moteur de sélection fournisseur

Fichier : `api/fournisseur-selection.js`

À partir d'un profil dossier (énergie, segment, note, volume, DDF, DFF, syndic, 1ère MES…), chaque fournisseur reçoit :

### Score métier /10

| Critère | Points max |
|---|---|
| Paiement UPFRONT | 3 pts |
| Marge (vs seuil fournisseur) | 2 pts |
| Horizon (couverture de la DFF) | 2 pts |
| Scoring minimum (note Ellipro) | 2 pts |
| Régularisation | 1 pt |

### Critères d'exclusion (hors score)

Les critères suivants peuvent exclure un fournisseur indépendamment du score :
- Segment non couvert
- Syndic refusé
- Volume sous le minimum
- Note crédit insuffisante

### Classification finale

- **Vert** (éligible, top 5) — fournisseurs recommandés
- **Orange** (éligible, hors top 5) — fournisseurs envisageables
- **Rouge** (non éligible) — fournisseurs exclus

---

## 💰 Données prix

Fichier : `api/fournisseur-pricing.js`

Chargement à la demande par segment (ex : `C4`, `T2`). Pour chaque fournisseur :
- Prix moyen sur 30 jours (€/MWh)
- Moyenne de marché sur le segment
- Écart en % vs la moyenne

Les données prix s'affichent en colonne dans les résultats de sélection et entrent dans l'évaluation (bonus/malus sur le score).

---

## 🗂️ Onglets de l'application

### Sélection Fournisseur
Formulaire avec : énergie, segment, note, volume, DDF, DFF, profil, état PDL, syndic, 1ère MES, fournisseur actuel, marge, UPFRONT.
Résultat : tableau classé avec score /10, prix, justification et case à cocher pour transmission.

### Résumé Fournisseurs
Vue tableau de tous les fournisseurs × tous les critères, alimentée dynamiquement via `/api/fournisseur-rules`.
Filtres : Tous / Gaz / Élec + recherche textuelle.

### Fournisseurs Panel
Classement panel (Gold Premium → Bronze) issu de `regles_panel_fournisseurs.xlsx`, onglet `Fournisseurs Panel`.

### News Fournisseurs
Tâches ClickUp de la liste `CLICKUP_LIST_NEWS`.
Filtres : statut (ouvert/fermé), mois de création, fournisseur.
Affichage : pièces jointes, logos fournisseurs, badges de statut.

### Challenges
Tâches ClickUp de la liste `CLICKUP_LIST_CHALLENGES`.
Affichage en cartes avec description, pièces jointes, logos fournisseurs.

---

## 📤 Système de transmission

Depuis la vue résultats, l'utilisateur peut cocher des fournisseurs et enregistrer une transmission :

1. Sélection des fournisseurs via cases à cocher
2. Saisie d'une référence dossier et d'un commentaire (optionnel)
3. Récapitulatif dans une modale de confirmation
4. Envoi vers `POST https://hub-capitole.vercel.app/api/log-transmis` avec :
   - Email de l'utilisateur (depuis la session auth)
   - Liste des fournisseurs transmis
   - Paramètres de recherche (énergie, segment, note, volume…)
   - Référence dossier + commentaire

---

## ⚙️ Variables d'environnement requises

```env
# Microsoft Graph / SharePoint
MICROSOFT_TENANT_ID=
MICROSOFT_CLIENT_ID=
MICROSOFT_CLIENT_SECRET=
SP_DRIVE_ID=
SP_FOLDER_PATH=PARTAGE/Team/3. Fournisseurs/Data Fournisseurs

# ClickUp
CLICKUP_TOKEN=
CLICKUP_LIST_NEWS=
CLICKUP_LIST_CHALLENGES=
```

---

## 📁 Fichiers clés

| Fichier | Rôle |
|---|---|
| `public/index.html` | Interface principale (SPA) |
| `public/auth-guard.js` | Authentification par token (Hub Capitole) |
| `api/fournisseur_rules.js` | Règles d'octroi globales depuis SharePoint |
| `api/fournisseur-selection.js` | Moteur de scoring + panel depuis SharePoint |
| `api/fournisseur-pricing.js` | Données prix par segment depuis SharePoint |
| `api/_clickup.js` | Récupération des tâches ClickUp (News + Challenges) |
| `vercel.json` | Configuration Vercel (routes, cleanUrls) |
| `package.json` | Dépendances Node.js (xlsx) |

---

## ⚠️ Points d'attention

- Les fichiers SharePoint doivent être accessibles et au bon chemin/format attendu
- En cas d'erreur API SharePoint, les règles embarquées dans le HTML servent de fallback pour le résumé fournisseurs
- En cas d'erreur pricing, la classification continue sans données prix
- La session auth expire après 8h ; une reconnexion via le Hub est alors nécessaire
- Toutes les variables d'environnement sont obligatoires pour le bon fonctionnement des endpoints
