# Sélection Fournisseurs — projet Vercel

🧭 Objectif
Cette application permet de :

Visualiser les règles d’attribution des fournisseurs
Sélectionner dynamiquement un fournisseur selon des critères métier
Gérer et suivre les demandes de modification via un questionnaire
Centraliser les informations issues de Salesforce
🚀 Nouveautés (dernière mise à jour)
🔗 1. Intégration API – Règles fournisseurs

Les données ne sont plus codées en dur dans le HTML.

Elles sont désormais récupérées dynamiquement via l’endpoint :

/api/fournisseur-rules
✅ Comportement
Chargement automatique au démarrage de l’application
Stockage des règles dans une variable globale : FOURNISSEUR_RULES
Gestion d’erreur si l’API ne répond pas
📊 2. Résumé Fournisseurs dynamique

Le tableau "Résumé Fournisseurs" :

est maintenant alimenté par l’API
s’adapte automatiquement aux données retournées
ne dépend plus de RESUME_DATA
🧠 3. Moteur de sélection fournisseur

La logique de sélection :

utilise désormais les règles issues de l’API
fonctionne en temps réel selon les critères utilisateur
garantit une cohérence avec les règles métier centralisées
📝 4. Formulaire “Règles d’octroi” connecté

Dans le questionnaire :

la liste des fournisseurs est maintenant dynamique
elle est alimentée automatiquement depuis l’API
plus besoin de maintenir une liste statique côté front
📰 5. Correction – Onglet News

Correction du filtre :

le statut "En cours" fonctionne désormais correctement
amélioration du filtrage des demandes
🏗️ Architecture simplifiée
Frontend (HTML / JS)
        ↓
Fetch API (/api/fournisseur-rules)
        ↓
Backend (Vercel Serverless Function)
        ↓
Source de vérité (Salesforce / JSON / autre)
⚙️ Fonctionnement technique
Chargement des règles
async function loadFournisseurRules() {
  const response = await fetch('/api/fournisseur-rules');
  const data = await response.json();
  window.FOURNISSEUR_RULES = data;
}
Utilisation dans l’application
Résumé fournisseurs → basé sur FOURNISSEUR_RULES
Sélection fournisseur → moteur dynamique
Formulaire → liste générée automatiquement
⚠️ Points d’attention
L’API /api/fournisseur-rules doit être déployée et accessible
En cas d’erreur API :
affichage d’un message utilisateur
aucune donnée ne sera chargée
🔜 Prochaines évolutions possibles
🔄 Connexion directe à Salesforce (API REST)
📈 Historisation des règles fournisseurs
🔔 Notifications en cas de modification
👤 Gestion des droits utilisateurs
📁 Fichiers clés
index.html → interface principale
/api/fournisseur-rules.js → récupération des règles
questionnaire.js → gestion des demandes
