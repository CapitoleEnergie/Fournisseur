const jsforce = require("jsforce");

const SF_CONFIG = {
  loginUrl: process.env.SF_LOGIN_URL || "https://login.salesforce.com",
  username: process.env.SF_USERNAME || "",
  password: process.env.SF_PASSWORD || "",
  securityToken: process.env.SF_SECURITY_TOKEN || "",
};

const ENERGIE_MAP = {
  electricite: "elec",
  "électricité": "elec",
  electricity: "elec",
  elec: "elec",
  gaz: "gaz",
  gas: "gaz",
  "gaz naturel": "gaz",
};

const SEGMENT_MAP = {
  c1: "C1", c2: "C2", c3: "C3", c4: "C4", c5: "C5",
  t1: "T1", t2: "T2", t3: "T3", t4: "T4", tp: "TP",
};

const ETAT_PDL_MAP = {
  "première mise en service": "premiere_mes",
  "premiere mise en service": "premiere_mes",
  "1ere mise en service": "premiere_mes",
  "1ère mise en service": "premiere_mes",
  "mise en service": "mes",
  mes: "mes",
  "en service": "en_service",
  ferme: "ferme",
  "fermé": "ferme",
};

function norm(v) {
  return String(v ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function safeNum(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function buildAdresse(c) {
  return [c.Voie__c, c.CodePostal__c, c.Commune__c]
    .map(v => String(v || "").replace(/[\r\n]+/g, " ").trim())
    .filter(Boolean)
    .join(", ");
}

module.exports = async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  const opportunityId = String(req.query.opportunityId || "").trim();

  if (!opportunityId) {
    return res.status(400).json({ message: "Paramètre opportunityId manquant" });
  }

  if (!SF_CONFIG.username || !SF_CONFIG.password || !SF_CONFIG.securityToken) {
    return res.status(500).json({
      message: "Variables Salesforce manquantes dans Vercel : SF_USERNAME, SF_PASSWORD ou SF_SECURITY_TOKEN.",
    });
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_CONFIG.loginUrl,
  });

  let loggedIn = false;

  try {
    await conn.login(
      SF_CONFIG.username,
      SF_CONFIG.password + SF_CONFIG.securityToken
    );

    loggedIn = true;

    const safeOpportunityId = opportunityId.replace(/'/g, "\\'");

    const oppRes = await conn.query(`
      SELECT Id, Name, NoteCredit__c, Energie__c
      FROM Opportunity
      WHERE Id = '${safeOpportunityId}'
      LIMIT 1
    `);

    if (!oppRes.records.length) {
      return res.status(404).json({
        message: `Opportunity ${opportunityId} introuvable dans Salesforce.`,
      });
    }

    const opp = oppRes.records[0];
    const energieMapped = ENERGIE_MAP[norm(opp.Energie__c)] || null;

    const offresRes = await conn.query(`
      SELECT
        Compteur__r.Name,
        Compteur__r.Segment__c,
        Compteur__r.ProfilCompteurGaz__c,
        Compteur__r.Fournisseur_Actuel_Nom__c,
        Compteur__r.EtatPDL__c,
        Compteur__r.VolumeTotalAnnuel__c,
        Compteur__r.VolumeReel__c,
        Compteur__r.Voie__c,
        Compteur__r.CodePostal__c,
        Compteur__r.Commune__c
      FROM Offre__c
      WHERE Opportunity__c = '${safeOpportunityId}'
      ORDER BY Compteur__r.Name
    `);

    const seen = new Set();
    const compteurs = [];

    for (const record of offresRes.records) {
      const c = record.Compteur__r;
      if (!c || seen.has(c.Name)) continue;

      seen.add(c.Name);

      compteurs.push({
        prm: c.Name || "",
        segment: SEGMENT_MAP[norm(c.Segment__c)] || null,
        segment_raw: c.Segment__c || "",
        profil: c.ProfilCompteurGaz__c || "",
        fournisseur_actuel: c.Fournisseur_Actuel_Nom__c || "",
        etat_pdl: ETAT_PDL_MAP[norm(c.EtatPDL__c)] || null,
        etat_pdl_raw: c.EtatPDL__c || "",
        volume_elec: safeNum(c.VolumeTotalAnnuel__c),
        volume_gaz: safeNum(c.VolumeReel__c),
        adresse: buildAdresse(c),
      });
    }

    const warnings = [];

    if (opp.Energie__c && !energieMapped) {
      warnings.push(`Énergie non reconnue : "${opp.Energie__c}"`);
    }

    compteurs.forEach(c => {
      if (c.segment_raw && !c.segment) {
        warnings.push(`Segment non reconnu : "${c.segment_raw}" (${c.prm})`);
      }
      if (c.etat_pdl_raw && !c.etat_pdl) {
        warnings.push(`État PDL non reconnu : "${c.etat_pdl_raw}" (${c.prm})`);
      }
    });

    return res.status(200).json({
      opportunityId: opp.Id,
      opportunityName: opp.Name || "",
      energie: energieMapped,
      energie_raw: opp.Energie__c || "",
      note_credit: safeNum(opp.NoteCredit__c),
      compteurs,
      warnings,
    });
  } catch (err) {
    console.error("[salesforce-opportunity]", err);

    return res.status(500).json({
      message: err.message || "Erreur Salesforce",
      name: err.name || "SalesforceError",
    });
  } finally {
    if (loggedIn) {
      try {
        await conn.logout();
      } catch (_) {}
    }
  }
};
