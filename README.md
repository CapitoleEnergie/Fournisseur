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
- Pré-remplir le formulaire de sélection depuis une opportunité Salesforce (avec sélection du compteur si l'opportunité en contient plusieurs)

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
         ├── /api/salesforce-opportunity ← pré-remplissage depuis une opportunité (Salesforce / jsforce)
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

> 📁 **Accès direct au dossier SharePoint :**
> [Ouvrir Data Fournisseurs](https://sascapitoleenergie.sharepoint.com/sites/PARTAGECAPITOLE-ENERGIE/Documents%20partages/Forms/AllItems.aspx?id=%2Fsites%2FPARTAGECAPITOLE%2DENERGIE%2FDocuments%20partages%2FPARTAGE%2FTeam%2F3%2E%20Fournisseurs%2FData%20Fournisseurs&viewid=11f04f50%2D66d5%2D4342%2D8695%2D7aade14ee934)
>
> C'est ici que vous devez déposer/modifier les fichiers pour que l'application les prenne en compte au prochain rechargement (cache 5 min).
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

## 🔄 Pré-remplissage Salesforce

Fichier : `api/salesforce-opportunity.js` (dépendance `jsforce`)

Depuis l'onglet **Sélection Fournisseur**, l'utilisateur saisit un **ID d'opportunité Salesforce** et clique sur « Pré-remplir ». L'endpoint `GET /api/salesforce-opportunity?opportunityId=006XXX` :

1. Lit l'opportunité (`Energie__c`, `NoteCredit__c`)
2. Récupère **tous les compteurs** associés via `Offre__c WHERE Opportunity__c = id` (dédoublonnés par PRM)
3. Normalise les valeurs SF vers le format du formulaire (énergie → `elec`/`gaz`, segment → `C1`–`C5`/`T1`–`T4`, état PDL → `premiere_mes`/`mes`/`en_service`/`ferme`)
4. Retourne `{ opportunityName, energie, note_credit, compteurs: [...], warnings }`

### Comportement selon le nombre de compteurs

- **1 compteur** → remplissage direct du formulaire + message de confirmation
- **N compteurs** → affichage d'un sélecteur (carte par compteur : PRM, segment, volume, adresse) ; l'utilisateur choisit celui à utiliser

Tous les champs restent **modifiables manuellement** après le pré-remplissage. Les champs non couverts par Salesforce (DDF, DFF, syndic, marge, commission) restent à saisir à la main.

### Mapping des champs

| Champ SF | Objet SF | Champ formulaire |
|---|---|---|
| `Energie__c` | Opportunité | énergie |
| `NoteCredit__c` | Opportunité | note |
| `Segment__c` | Compteur | segment |
| `ProfilCompteurGaz__c` | Compteur (gaz) | profil |
| `Fournisseur_Actuel_Nom__c` | Compteur | fournisseur actuel |
| `EtatPDL__c` | Compteur | état PDL |
| `VolumeTotalAnnuel__c` | Compteur (élec) | volume |
| `VolumeReel__c` | Compteur (gaz) | volume |

---

## 🗂️ Onglets de l'application

### Sélection Fournisseur
Formulaire avec : énergie, segment, note, volume, DDF, DFF, profil, état PDL, syndic, 1ère MES, fournisseur actuel, marge, UPFRONT.
Pré-remplissage possible depuis une opportunité Salesforce (voir section dédiée).
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

# Salesforce (pré-remplissage opportunité)
SF_LOGIN_URL=https://login.salesforce.com
SF_USERNAME=
SF_PASSWORD=
SF_SECURITY_TOKEN=
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
| `api/salesforce-opportunity.js` | Pré-remplissage du formulaire depuis une opportunité Salesforce (jsforce) |
| `api/_clickup.js` | Récupération des tâches ClickUp (News + Challenges) |
| `vercel.json` | Configuration Vercel (routes, cleanUrls) |
| `package.json` | Dépendances Node.js (xlsx, jsforce) |

---

## ⚠️ Points d'attention

- Les fichiers SharePoint doivent être accessibles et au bon chemin/format attendu
- En cas d'erreur API SharePoint, les règles embarquées dans le HTML servent de fallback pour le résumé fournisseurs
- En cas d'erreur pricing, la classification continue sans données prix
- La session auth expire après 8h ; une reconnexion via le Hub est alors nécessaire
- Toutes les variables d'environnement sont obligatoires pour le bon fonctionnement des endpoints
