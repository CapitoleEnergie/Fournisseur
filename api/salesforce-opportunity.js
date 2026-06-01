/**
 * api/salesforce-opportunity.js
 * Récupère les champs Salesforce d'une opportunité pour pré-remplir
 * le formulaire de sélection fournisseur.
 *
 * GET /api/salesforce-opportunity?opportunityId=006XXXX
 */

const jsforce = require('jsforce');

const SF_CONFIG = {
  loginUrl:      process.env.SF_LOGIN_URL      || 'https://login.salesforce.com',
  username:      process.env.SF_USERNAME        || '',
  password:      process.env.SF_PASSWORD        || '',
  securityToken: process.env.SF_SECURITY_TOKEN  || '',
};

// ── Normalisation des valeurs Salesforce → valeurs formulaire ─────────────

const ENERGIE_MAP = {
  'electricite':   'elec',
  'électricité':   'elec',
  'electricité':   'elec',
  'electricity':   'elec',
  'elec':          'elec',
  'élec':          'elec',
  'gaz':           'gaz',
  'gas':           'gaz',
  'gaz naturel':   'gaz',
};

const SEGMENT_MAP = {
  'c1': 'C1', 'c2': 'C2', 'c3': 'C3', 'c4': 'C4', 'c5': 'C5',
  't1': 'T1', 't2': 'T2', 't3': 'T3', 't4': 'T4', 'tp': 'TP',
};

const ETAT_PDL_MAP = {
  'première mise en service': 'premiere_mes',
  'premiere mise en service': 'premiere_mes',
  '1ere mise en service':     'premiere_mes',
  '1ère mise en service':     'premiere_mes',
  'premiere mes':              'premiere_mes',
  '1ere mes':                  'premiere_mes',
  'mise en service':           'mes',
  'mes':                       'mes',
  'en service':                'en_service',
  'ferme':                     'ferme',
  'fermé':                     'ferme',
};

function norm(v) {
  return String(v ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function mapEnergy(raw)   { return ENERGIE_MAP[norm(raw)]   || null; }
function mapSegment(raw)  { return SEGMENT_MAP[norm(raw)]   || null; }
function mapEtatPdl(raw)  { return ETAT_PDL_MAP[norm(raw)]  || null; }

function safeNum(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

// ── Handler ───────────────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ message: 'Méthode non autorisée' });
  }

  const opportunityId = String(req.query.opportunityId || '').trim();
  if (!opportunityId) {
    return res.status(400).json({ message: 'Paramètre opportunityId manquant' });
  }

  const conn = new jsforce.Connection({ loginUrl: SF_CONFIG.loginUrl });

  try {
    await conn.login(SF_CONFIG.username, SF_CONFIG.password + SF_CONFIG.securityToken);

    // 1. Opportunité — NoteCredit__c et Energie__c
    const oppRes = await conn.query(`
      SELECT Id, Name, NoteCredit__c, Energie__c
      FROM Opportunity
      WHERE Id = '${opportunityId}'
      LIMIT 1
    `);

    if (!oppRes.records.length) {
      return res.status(404).json({ message: `Opportunity ${opportunityId} introuvable dans Salesforce.` });
    }

    const opp = oppRes.records[0];

    // 2. Premier Compteur via Offre__c (trié par nom de compteur)
    const offresRes = await conn.query(`
      SELECT
        Compteur__r.Segment__c,
        Compteur__r.ProfilCompteurGaz__c,
        Compteur__r.Fournisseur_Actuel_Nom__c,
        Compteur__r.EtatPDL__c,
        Compteur__r.VolumeTotalAnnuel__c,
        Compteur__r.VolumeReel__c
      FROM Offre__c
      WHERE Opportunity__c = '${opportunityId}'
      ORDER BY Compteur__r.Name
      LIMIT 1
    `);

    const compteur = offresRes.records.length > 0
      ? offresRes.records[0].Compteur__r
      : null;

    // 3. Mapper les valeurs SF vers le format formulaire
    const energieRaw  = opp.Energie__c || '';
    const energieMapped = mapEnergy(energieRaw);

    const payload = {
      opportunityId:    opp.Id,
      opportunityName:  opp.Name || '',

      // Opportunité
      energie:          energieMapped,
      energie_raw:      energieRaw,
      note_credit:      safeNum(opp.NoteCredit__c),

      // Compteur
      segment:          compteur ? mapSegment(compteur.Segment__c) : null,
      segment_raw:      compteur?.Segment__c || '',
      profil:           compteur?.ProfilCompteurGaz__c || '',
      fournisseur_actuel: compteur?.Fournisseur_Actuel_Nom__c || '',
      etat_pdl:         compteur ? mapEtatPdl(compteur.EtatPDL__c) : null,
      etat_pdl_raw:     compteur?.EtatPDL__c || '',

      // Volume selon énergie
      volume_elec:      safeNum(compteur?.VolumeTotalAnnuel__c),
      volume_gaz:       safeNum(compteur?.VolumeReel__c),
    };

    // Warnings si valeurs non mappées
    const warnings = [];
    if (energieRaw && !energieMapped)   warnings.push(`Énergie non reconnue : "${energieRaw}"`);
    if (compteur?.Segment__c && !payload.segment) warnings.push(`Segment non reconnu : "${compteur.Segment__c}"`);
    if (compteur?.EtatPDL__c && !payload.etat_pdl) warnings.push(`État PDL non reconnu : "${compteur.EtatPDL__c}"`);

    return res.status(200).json({ ...payload, warnings });

  } catch (err) {
    console.error('[salesforce-opportunity]', err.message || err);
    return res.status(500).json({ message: err.message || 'Erreur Salesforce' });
  } finally {
    try { await conn.logout(); } catch (_) {}
  }
};
