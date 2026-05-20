const xlsx = require("xlsx");

// ============ SHAREPOINT CONFIG ============
const TENANT_ID     = process.env.MICROSOFT_TENANT_ID;
const CLIENT_ID     = process.env.MICROSOFT_CLIENT_ID;
const CLIENT_SECRET = process.env.MICROSOFT_CLIENT_SECRET;
const DRIVE_ID      = process.env.SP_DRIVE_ID;
const FOLDER_PATH   = process.env.SP_FOLDER_PATH || "PARTAGE/Team/3. Fournisseurs/Data Fournisseurs";
const FILE_NAME     = "regles_panel_fournisseurs.xlsx";
const RULES_SHEET   = "Fournisseurs Export";
const PANEL_SHEET   = "Fournisseurs Panel";

const PANEL_PRIORITY = {
  "gold premium": 1,
  "gold": 2,
  "silver": 3,
  "bronze": 4,
  "": 99,
  "non classe": 99
};

// ============ CACHE MÉMOIRE (5 minutes) ============
let _cache = null;
let _cacheAt = 0;
const CACHE_TTL_MS = 5 * 60 * 1000;

// ============ TOKEN MICROSOFT ============
async function getMsToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default"
      })
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error("Impossible d'obtenir le token Microsoft Graph");
  return data.access_token;
}

// ============ TÉLÉCHARGEMENT EXCEL DEPUIS SHAREPOINT ============
async function downloadExcelFromSharePoint() {
  const token = await getMsToken();

  const fileRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${FOLDER_PATH}/${FILE_NAME}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const fileMeta = await fileRes.json();

  if (!fileMeta["@microsoft.graph.downloadUrl"]) {
    throw new Error(`Fichier SharePoint introuvable : ${FILE_NAME}`);
  }

  const dlRes = await fetch(fileMeta["@microsoft.graph.downloadUrl"]);
  const arrayBuffer = await dlRes.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

function normalizeText(value = "") {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function slugify(value = "") {
  return normalizeText(value)
    .toLowerCase()
    .replace(/&/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function safeNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;

  let s = String(value)
    .replace(/\u00A0/g, " ")
    .replace(/€/g, "")
    .replace(/MWh/gi, "")
    .trim();

  if (!s) return null;

  s = s.replace(/\s+/g, "");

  const hasComma = s.includes(",");
  const hasDot = s.includes(".");

  if (hasComma && hasDot) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (hasComma) {
    s = s.replace(",", ".");
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function parseFrenchDate(value) {
  if (value === null || value === undefined || value === "") return null;

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = xlsx.SSF.parse_date_code(value);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
  }

  const s = normalizeText(value);
  if (!s) return null;

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    const parts = s.split("/");
    return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  }

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return new Date(s.slice(0, 10));
  }

  if (/^\d{5}(?:\.\d+)?$/.test(s)) {
    const parsed = xlsx.SSF.parse_date_code(Number(s));
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
  }

  return null;
}

function getYearFromHorizon(value) {
  const s = normalizeText(value);
  if (!s) return null;
  const match = s.match(/\b(20\d{2})\b/);
  return match ? Number(match[1]) : null;
}

function normalizeSupplierName(raw = "") {
  const cleaned = normalizeText(raw);
  const key = slugify(cleaned);

  const aliases = new Map([
    ["endesa energia succursal france", "ENDESA"],
    ["sefe energy sas", "SEFE"],
    ["totalenergies electricite et gaz france", "TOTALENERGIES"],
    ["alterna", "ALTERNA"],
    ["alterna energie", "ALTERNA"],
    ["engie", "ENGIE"],
    ["energem", "ENERGEM"],
    ["mint", "MINT Energie"],
    ["mint energie", "MINT Energie"],
    ["ilek", "ILEK"],
    ["gedia", "GEDIA"],
    ["gedia energies services", "GEDIA"],
    ["enovos", "ENOVOS"],
    ["enovos luxembourg s a", "ENOVOS"],
    ["primeo", "PRIMEO"],
    ["primeo energie france", "PRIMEO"],
    ["ohm", "OHM (mid)"],
    ["ohm energie", "OHM (mid)"],
    ["ohm gc", "OHM (GC)"],
    ["ohm mid", "OHM (mid)"],
    ["dyneff", "DYNEFF"],
    ["dyneff s a s", "DYNEFF"],
    ["bcm energy", "BCM ENERGY"],
    ["hellio", "HELLIO"],
    ["hellio solutions", "HELLIO"],
    ["la bellenergie", "LA BELLENERGIE"],
    ["save", "SAVE"],
    ["societe d approvisionnement et de vente d energies", "SAVE"],
    ["geg", "GEG"],
    ["geg source d energies", "GEG"],
    ["synelva", "SYNELVA"],
    ["gaz europeen", "GAZ Européen"],
    ["gaz de bordeaux", "GAZ DE BORDEAUX"],
    ["energie d ici", "ENERGIE D'ICI"],
    ["vattenfall c5", "VATTENFALL (C5)"],
    ["vattenfall gc", "VATTENFALL (GC)"],
    ["vattenfall", "VATTENFALL (C5)"],
    ["picoty", "PICOTY"],
    ["natgas", "NATGAS"],
    ["elmy", "ELMY"],
    ["ekwateur", "EKWATEUR"]
  ]);

  return aliases.get(key) || cleaned;
}

function normalizeEnergy(value = "") {
  const s = slugify(value);
  if (s === "gaz") return "gaz";
  if (s === "elec" || s === "electricite" || s === "electricity") return "elec";
  return "";
}

function normalizeSegment(value = "") {
  const s = normalizeText(value).toUpperCase();
  const allowed = ["TP", "T1", "T2", "T3", "T4", "C1", "C2", "C3", "C4", "C5"];
  return allowed.includes(s) ? s : "";
}

function normalizeOuiNon(value = "") {
  const s = slugify(value);
  if (s === "oui" || s === "true" || s === "1") return "oui";
  if (s === "non" || s === "false" || s === "0") return "non";
  return "";
}

function extractMinVolume(value) {
  const s = normalizeText(value);
  if (!s) return null;
  const match = s.match(/(\d+(?:[.,]\d+)?)\s*MWh/i);
  return match ? safeNumber(match[1]) : null;
}

function evaluateSegmentRule(ruleValue) {
  const value = normalizeText(ruleValue);
  const upper = value.toUpperCase();

  if (!value) return { eligible: false, status: "ko", reason: "Segment non renseigne" };
  if (upper === "OUI") return { eligible: true, status: "ok", reason: "Segment couvert" };
  if (upper === "NON") return { eligible: false, status: "ko", reason: "Segment non couvert" };
  if (upper.includes("EN PAUSE")) return { eligible: false, status: "ko", reason: value };
  if (upper.includes("CAS PAR CAS") || upper.startsWith("OUI") || upper.startsWith("NON")) {
    return { eligible: true, status: "warn", reason: value };
  }
  return { eligible: true, status: "warn", reason: value };
}

function evaluateSyndicRule(ruleValue, syndic, energie, supplierName) {
  const value = normalizeText(ruleValue);
  const upper = value.toUpperCase();

  // GAZ Européen fait UNIQUEMENT du Syndic → vert si syndic = oui, ko si non
  if (supplierName && slugify(supplierName).includes("gaz europeen")) {
    if (syndic === "oui") {
      return { eligible: true, status: "ok", reason: "Gaz Européen — spécialiste Syndic/SDC" };
    }
    // Le check "gaz europeen syndic obligatoire" est déjà géré séparément, ici on renvoie juste ok si pas demandé
  }

  if (syndic !== "oui") {
    return { eligible: true, status: "ok", reason: "Critere syndic non demande" };
  }

  if (!value) return { eligible: true, status: "warn", reason: "Regle syndic non renseignee" };
  if (upper === "NON") return { eligible: false, status: "ko", reason: "Syndic non accepte" };
  if (upper === "OUI") return { eligible: true, status: "ok", reason: "Syndic accepte" };

  // Vérifier condition d'énergie : "OUI (élec uniquement)", "OUI (seulement gaz si SIRET)"
  if (upper.startsWith("OUI") && energie) {
    const parenthesisMatch = value.match(/\(([^)]+)\)/i);
    if (parenthesisMatch) {
      const condition = normalizeText(parenthesisMatch[1]).toUpperCase();
      const isGazCondition = condition.includes("GAZ") || condition.includes("SEULEMENT GAZ");
      const isElecCondition = condition.includes("ELEC") || condition.includes("ELECTRICITE");

      if (isElecCondition && !isGazCondition) {
        if (energie === "elec") {
          return { eligible: true, status: "ok", reason: `${value} — énergie élec correspond` };
        } else {
          return { eligible: false, status: "ko", reason: `${value} — uniquement élec, énergie du dossier : gaz` };
        }
      }
      if (isGazCondition && !isElecCondition) {
        if (energie === "gaz") {
          return { eligible: true, status: "ok", reason: `${value} — énergie gaz correspond` };
        } else {
          return { eligible: false, status: "ko", reason: `${value} — uniquement gaz, énergie du dossier : élec` };
        }
      }
    }
  }

  if (upper.includes("CAS PAR CAS") || upper.startsWith("OUI")) {
    return { eligible: true, status: "warn", reason: value };
  }
  return { eligible: true, status: "warn", reason: value };
}

function evaluateScoringRule(ruleValue, note) {
  const value = normalizeText(ruleValue);
  if (!value || note === null || note === undefined) {
    return { eligible: true, status: "neutral", reason: value || "Scoring non renseigne" };
  }

  // Nouveau format : juste un nombre (ex: "4", "5", "6") ou ancien format "X/10"
  const match = value.match(/(\d+)(?:\s*\/\s*10)?/);
  if (!match) return { eligible: true, status: "neutral", reason: value };

  const minimum = Number(match[1]);
  if (!Number.isFinite(minimum)) return { eligible: true, status: "neutral", reason: value };

  if (note >= minimum) {
    return { eligible: true, status: "ok", reason: `Note ${note} ≥ minimum ${minimum}` };
  }
  // note < minimum → alerte orange mais pas rédhibitoire
  return { eligible: true, status: "warn", reason: `Note ${note} < minimum ${minimum}` };
}

function evaluateVolumeRule(ruleValue, volume) {
  const normalized = normalizeText(ruleValue);

  // "Aucun" ou "aucun" → pas de volume minimal → vert ok
  if (slugify(normalized) === "aucun") {
    return { eligible: true, status: "ok", reason: "Aucun volume minimal requis" };
  }

  const minVolume = extractMinVolume(ruleValue);
  if (minVolume === null || volume === null || volume === undefined) {
    return { eligible: true, status: "neutral", reason: normalized || "Pas de volume minimum exploitable" };
  }

  if (volume < minVolume) {
    return { eligible: true, status: "warn", reason: `Volume ${volume} MWh < minimum ${minVolume} MWh` };
  }
  return { eligible: true, status: "ok", reason: `Volume ${volume} MWh ≥ minimum ${minVolume} MWh` };
}

function evaluateUpfrontPaymentRule(ruleValue, commissionEstimee, ddfDate, supplierName, dffDate) {
  const value = normalizeText(ruleValue);
  const upper = value.toUpperCase();

  if (!value) {
    return { eligible: true, status: "neutral", reason: "Paiement UPFRONT non renseigné" };
  }

  if (upper === "NON") {
    return { eligible: true, status: "warn", reason: "Paiement UPFRONT non proposé" };
  }

  // --- Vattenfall GC règle spéciale : "OUI sur 2025/2026" → vérifier DFF ≤ N+1 ---
  if (supplierName && slugify(supplierName).includes("vattenfall") && slugify(supplierName).includes("gc")) {
    if (upper.includes("2025/2026") || upper.startsWith("OUI")) {
      const currentYear = new Date().getFullYear();
      const maxYear = currentYear + 1; // N+1
      if (dffDate) {
        const dffYear = dffDate.getFullYear();
        if (dffYear <= maxYear) {
          return { eligible: true, status: "ok", reason: `${value} — DFF ${dffYear} ≤ N+1 (${maxYear}), UPFRONT validé` };
        } else {
          return { eligible: true, status: "warn", reason: `${value} — DFF ${dffYear} > N+1 (${maxYear}), UPFRONT incertain` };
        }
      }
      return { eligible: true, status: "warn", reason: `${value} — DFF non renseignée, impossible de vérifier` };
    }
  }

  if (upper.startsWith("OUI")) {
    // --- Vérification seuil commission ---
    // Supporte les formats : "< 50K", "< 50k", "< 10 000 €", "< 10000€", "< 2,5 GWh"
    let threshold = null;

    // Format "< XK" ou "< X k"
    const kMatches = [...upper.matchAll(/[<≤]\s*(\d+(?:[.,]\d+)?)\s*K/gi)];
    if (kMatches.length > 0) {
      const lastK = kMatches[kMatches.length - 1];
      threshold = parseFloat(lastK[1].replace(",", ".")) * 1000;
    }

    // Format "< X XXX €" ou "< XXXXX €" (nombre avec espaces + symbole €)
    if (threshold === null) {
      const euroMatches = [...normalizeText(ruleValue).toUpperCase().matchAll(/[<≤]\s*([\d\s]+(?:[.,]\d+)?)\s*€/gi)];
      if (euroMatches.length > 0) {
        const lastEuro = euroMatches[euroMatches.length - 1];
        const cleaned = lastEuro[1].replace(/\s/g, "").replace(",", ".");
        const val = parseFloat(cleaned);
        if (Number.isFinite(val)) threshold = val;
      }
    }

    let commOk = true;
    let commReason = "";
    if (threshold !== null) {
      if (commissionEstimee === null || commissionEstimee === undefined) {
        commOk = false;
        commReason = `commission estimée non renseignée (seuil ${threshold.toLocaleString("fr-FR")} €)`;
      } else if (commissionEstimee <= threshold) {
        commReason = `commission ${commissionEstimee.toLocaleString("fr-FR")} € ≤ seuil ${threshold.toLocaleString("fr-FR")} €`;
      } else {
        commOk = false;
        commReason = `commission ${commissionEstimee.toLocaleString("fr-FR")} € > seuil ${threshold.toLocaleString("fr-FR")} €`;
      }
    }

    // --- Vérification DDF M+X ou N+X dans la règle upfront ---
    const moisMatch = upper.match(/DDF\s*[<≤]\s*M\s*\+\s*(\d+)/i);
    const anneesMatch = upper.match(/DDF\s*[<≤]\s*N\s*\+\s*(\d+)/i);
    let ddfOk = true;
    let ddfReason = "";

    if (moisMatch || anneesMatch) {
      if (!ddfDate) {
        ddfOk = false;
        ddfReason = "DDF non renseignée";
      } else {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const dateMax = new Date(today);
        if (moisMatch) {
          dateMax.setMonth(dateMax.getMonth() + parseInt(moisMatch[1], 10));
        } else {
          dateMax.setFullYear(dateMax.getFullYear() + parseInt(anneesMatch[1], 10));
        }
        const label = dateMax.toLocaleDateString("fr-FR");
        if (dateMax >= ddfDate) {
          ddfReason = `DDF compatible (limite ${label})`;
        } else {
          ddfOk = false;
          ddfReason = `DDF ${ddfDate.toLocaleDateString("fr-FR")} > limite ${label}`;
        }
      }
    }

    // --- Résultat combiné ---
    const hasDdfCondition = !!(moisMatch || anneesMatch);
    const hasCommCondition = threshold !== null;

    if (!hasCommCondition && !hasDdfCondition) {
      return { eligible: true, status: "ok", reason: value };
    }

    const parts = [commReason, ddfReason].filter(Boolean);
    const allOk = commOk && ddfOk;
    return {
      eligible: true,
      status: allOk ? "ok" : "warn",
      reason: `${value} — ${parts.join(" | ")}`
    };
  }

  return { eligible: true, status: "warn", reason: value };
}

function evaluateRegularisationCommissionsRule(ruleValue) {
  const value = normalizeText(ruleValue);
  const upper = value.toUpperCase();

  if (!value) return { eligible: true, status: "neutral", reason: "Régularisation commissions non renseignée" };
  if (upper === "OUI") return { eligible: true, status: "warn", reason: "Oui" };
  if (upper === "NON") return { eligible: true, status: "ok", reason: "Non" };
  return { eligible: true, status: "warn", reason: value };
}

function evaluateMesRule(ruleValue, mesType, energie) {
  // mesType = "premiere" | "remise" | null
  // Si non pertinent (En service, Fermé, non renseigné) → neutral sans impact
  if (!mesType) {
    return null; // null = ne pas afficher la règle
  }

  const value = normalizeText(ruleValue);
  const upper = value.toUpperCase();

  if (!value) {
    return { eligible: true, status: "neutral", reason: "Règle MES non renseignée" };
  }

  // "NON" → exclusion rouge
  if (upper === "NON") {
    return { eligible: false, status: "ko", reason: `Non accepté` };
  }

  // "OUI" strict → vert
  if (upper === "OUI") {
    return { eligible: true, status: "ok", reason: "Accepté" };
  }

  // "OUI (...)" → vérifier si la condition porte sur l'énergie ou le segment
  if (upper.startsWith("OUI")) {
    // Détecter condition d'énergie : "OUI (gaz)", "OUI (élec)", "OUI (elec)", "OUI (gaz et C5)"
    const parenthesisMatch = value.match(/\(([^)]+)\)/i);
    if (parenthesisMatch && energie) {
      const condition = normalizeText(parenthesisMatch[1]).toUpperCase();
      const isGazCondition = condition.includes("GAZ");
      const isElecCondition = condition.includes("ELEC") || condition.includes("ELECTRICITE");
      const isC5Condition = condition.includes("C5");

      // Si la condition mentionne une énergie spécifique
      if (isGazCondition && !isElecCondition) {
        // Condition "gaz" uniquement
        if (energie === "gaz") {
          return { eligible: true, status: "ok", reason: `${value} — énergie gaz correspond` };
        } else {
          return { eligible: false, status: "ko", reason: `${value} — uniquement gaz, énergie du dossier : élec` };
        }
      }
      if (isElecCondition && !isGazCondition) {
        // Condition "élec" uniquement
        if (energie === "elec") {
          return { eligible: true, status: "ok", reason: `${value} — énergie élec correspond` };
        } else {
          return { eligible: false, status: "ko", reason: `${value} — uniquement élec, énergie du dossier : gaz` };
        }
      }
      // Condition segment uniquement ex: "OUI (C5 uniquement)"
      if (isC5Condition && !isGazCondition && !isElecCondition) {
        return { eligible: true, status: "warn", reason: value };
      }
    }

    return { eligible: true, status: "warn", reason: value };
  }

  return { eligible: true, status: "warn", reason: value };
}

function evaluateMargeRule(ruleValue, energie, segment, volume, margeGlobale) {
  const raw = normalizeText(ruleValue);
  if (!raw) return { eligible: true, status: "neutral", reason: "Marge non renseignée", margeImpact: 0 };

  // Cas neutres immédiats
  const neutralKeywords = ["grille", "pas de limite", "pas de marge maximum", "cas par cas", "prm", "marge integree", "marge dans l abonnement"];
  const rawSlug = slugify(raw);
  if (neutralKeywords.some(k => rawSlug.includes(slugify(k)))) {
    return { eligible: true, status: "ok", reason: raw, margeImpact: 0 };
  }

  // Découper le texte en blocs par énergie
  // On cherche le bloc correspondant à l'énergie saisie
  function extractEnergyBlock(text, eng) {
    const norm = normalizeText(text);
    // Identifier les blocs "Electricite= ..." et "Gaz= ..."
    const elecPattern = /Electricite\s*(?:et\s*Gaz\s*)?=\s*/i;
    const gazPattern = /Gaz\s*(?:et\s*Electricite\s*)?=\s*/i;
    const bothPattern = /Electricite\s*et\s*Gaz\s*=\s*/i;

    // Tester "Electricité et Gaz" d'abord (s'applique aux deux)
    if (/electricite\s*(?:ou\s*)?(?:et\s*)?gaz\s*=/i.test(norm)) {
      return norm.replace(/electricite\s*(?:ou\s*)?(?:et\s*)?gaz\s*=\s*/i, "").trim();
    }

    const elecIdx = norm.search(/electricite\s*=/i);
    const gazIdx = norm.search(/gaz\s*=/i);

    if (eng === "elec") {
      if (elecIdx >= 0) {
        const end = gazIdx > elecIdx ? gazIdx : norm.length;
        return norm.slice(elecIdx).replace(/electricite\s*=\s*/i, "").slice(0, end - elecIdx).trim();
      }
      return null;
    }

    if (eng === "gaz") {
      if (gazIdx >= 0) {
        const end = elecIdx > gazIdx ? elecIdx : norm.length;
        return norm.slice(gazIdx).replace(/gaz\s*=\s*/i, "").slice(0, end - gazIdx).trim();
      }
      return null;
    }

    return norm;
  }

  const block = extractEnergyBlock(raw, energie) || raw;

  // Cas neutre dans le bloc extrait
  const blockSlug = slugify(block);
  if (neutralKeywords.some(k => blockSlug.includes(slugify(k)))) {
    return { eligible: true, status: "ok", reason: block, margeImpact: 0 };
  }

  // Si marge non saisie → informatif
  if (margeGlobale === null || margeGlobale === undefined) {
    return { eligible: true, status: "neutral", reason: `Marge non saisie — règle : ${block}`, margeImpact: 0 };
  }

  // Chercher la ligne du segment dans le bloc
  function findSegmentLine(text, seg) {
    const lines = text.split(/\n/);
    for (const line of lines) {
      if (line.toUpperCase().includes(seg)) return line.trim();
    }
    return null;
  }

  // Chercher la ligne selon le volume (CAR)
  function findVolumeLine(text, vol) {
    if (vol === null || vol === undefined) return null;
    const lines = text.split(/\n/);
    const volMwh = vol;

    for (const line of lines) {
      const normLine = normalizeText(line);
      const rangeMatch = normLine.match(/(\d+(?:[.,]\d+)?)\s*(MWh|GWh)\s*<\s*(?:CAR)?\s*<\s*(\d+(?:[.,]\d+)?)\s*(MWh|GWh)/i);
      if (rangeMatch) {
        const lo = safeNumber(rangeMatch[1]) * (rangeMatch[2].toLowerCase() === "gwh" ? 1000 : 1);
        const hi = safeNumber(rangeMatch[3]) * (rangeMatch[4].toLowerCase() === "gwh" ? 1000 : 1);
        if (lo !== null && hi !== null && volMwh > lo && volMwh <= hi) return line.trim();
        continue;
      }
      const ltMatch = normLine.match(/[<≤]\s*(\d+(?:[.,]\d+)?)\s*(MWh|GWh)/i);
      if (ltMatch) {
        const limit = safeNumber(ltMatch[1]) * (ltMatch[2].toLowerCase() === "gwh" ? 1000 : 1);
        if (limit !== null && volMwh <= limit) return line.trim();
        continue;
      }
      const gtMatch = normLine.match(/[>≥]\s*(\d+(?:[.,]\d+)?)\s*(MWh|GWh)/i);
      if (gtMatch) {
        const limit = safeNumber(gtMatch[1]) * (gtMatch[2].toLowerCase() === "gwh" ? 1000 : 1);
        if (limit !== null && volMwh > limit) return line.trim();
        continue;
      }
    }
    return null;
  }

  // Trouver la ligne pertinente
  let targetLine = null;

  // 1. Chercher segment d'abord
  if (segment) targetLine = findSegmentLine(block, segment);

  // 2. Si pas trouvé et bloc contient CAR ou seuils de volume → chercher par volume
  if (!targetLine && /CAR|MWh|GWh/i.test(block) && volume !== null) {
    targetLine = findVolumeLine(block, volume);
  }

  // 3. Fallback : prendre le bloc entier
  if (!targetLine) targetLine = block;

  // Extraire le seuil €/MWh de la ligne cible
  const seuilMatch = normalizeText(targetLine).match(/(\d+(?:[.,]\d+)?)\s*€\s*\/\s*MWh/i);
  if (!seuilMatch) {
    // Pas de seuil extractible → neutre
    return { eligible: true, status: "ok", reason: targetLine, margeImpact: 0 };
  }

  const seuil = safeNumber(seuilMatch[1]);
  if (seuil === null) {
    return { eligible: true, status: "neutral", reason: targetLine, margeImpact: 0 };
  }

  if (margeGlobale <= seuil) {
    return {
      eligible: true,
      status: "ok",
      reason: `Marge ${margeGlobale} €/MWh ≤ seuil ${seuil} €/MWh`,
      margeImpact: 5
    };
  }

  return {
    eligible: true,
    status: "warn",
    reason: `Marge ${margeGlobale} €/MWh > seuil ${seuil} €/MWh`,
    margeImpact: -5
  };
}

function evaluateDdfRule(ruleValue, ddfDate) {
  const value = normalizeText(ruleValue);
  const upper = value.toUpperCase();

  if (!value) {
    return { eligible: true, status: "neutral", reason: "DDF max non renseignee" };
  }

  if (!ddfDate) {
    return { eligible: true, status: "neutral", reason: value };
  }

  // M+X (mois) ou N+X (années) → calculer la date limite depuis aujourd'hui
  const moisMatch = upper.match(/\bM\s*\+\s*(\d+)\b/i);
  const anneesMatch = upper.match(/\bN\s*\+\s*(\d+)\b/i);

  if (moisMatch || anneesMatch) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dateMax = new Date(today);

    if (moisMatch) {
      dateMax.setMonth(dateMax.getMonth() + parseInt(moisMatch[1], 10));
    } else {
      dateMax.setFullYear(dateMax.getFullYear() + parseInt(anneesMatch[1], 10));
    }

    const label = dateMax.toLocaleDateString("fr-FR");

    if (dateMax >= ddfDate) {
      return {
        eligible: true,
        status: "ok",
        reason: `${value} — limite le ${label}, compatible avec la DDF`
      };
    }

    return {
      eligible: true,
      status: "warn",
      reason: `${value} — limite le ${label}, DDF ${ddfDate.toLocaleDateString("fr-FR")} trop lointaine`
    };
  }

  if (slugify(value) === "pas de limite") {
    return { eligible: true, status: "ok", reason: "Pas de limite" };
  }

  const maxDate = parseFrenchDate(ruleValue);
  if (!maxDate) {
    return { eligible: true, status: "warn", reason: value };
  }

  if (ddfDate > maxDate) {
    return {
      eligible: true,
      status: "warn",
      reason: `DDF ${ddfDate.toLocaleDateString("fr-FR")} > DDF max ${maxDate.toLocaleDateString("fr-FR")}`
    };
  }

  return {
    eligible: true,
    status: "ok",
    reason: `DDF compatible jusqu'au ${maxDate.toLocaleDateString("fr-FR")}`
  };
}

function evaluateHorizonRule(ruleValue, dffDate) {
  const year = getYearFromHorizon(ruleValue);
  if (!year || !dffDate) {
    return { eligible: true, status: "neutral", reason: normalizeText(ruleValue) || "Horizon non renseigne" };
  }

  const dffYear = dffDate.getFullYear();
  if (year < dffYear) {
    return { eligible: true, status: "warn", reason: `Horizon ${year} < fin fourniture ${dffYear}` };
  }
  if (year === dffYear) {
    return { eligible: true, status: "ok", reason: `Horizon ${year} couvre la fin de fourniture ${dffYear}` };
  }
  return { eligible: true, status: "ok", reason: `Horizon ${year} couvre la periode` };
}

function parseRulesSheet(workbook) {
  const sheet = workbook.Sheets[RULES_SHEET];
  if (!sheet) {
    throw new Error(`Onglet introuvable: ${RULES_SHEET}`);
  }

  const data = xlsx.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  if (data.length < 3) {
    throw new Error("Onglet regles vide ou format inattendu.");
  }

  const headerRowIndex = 2;
  const headerRow = data[headerRowIndex];
  const supplierIndexes = [];

  for (let col = 3; col < headerRow.length; col += 1) {
    const supplierName = normalizeSupplierName(headerRow[col]);
    if (!supplierName) continue;
    supplierIndexes.push({ col, supplierName });
  }

  const fournisseurs = supplierIndexes.map((x) => x.supplierName);
  const rulesBySupplier = {};
  fournisseurs.forEach((f) => {
    rulesBySupplier[f] = {};
  });

  let currentCategory = "";

  for (let rowIdx = headerRowIndex + 1; rowIdx < data.length; rowIdx += 1) {
    const row = data[rowIdx];
    const category = normalizeText(row[0]);
    const critere = normalizeText(row[1]);

    if (category) currentCategory = category;
    if (!category && !critere) continue;

    let ruleKey = "";
    if (currentCategory.toUpperCase().includes("SEGMENTS") && critere) {
      ruleKey = critere.toUpperCase();
    } else if (critere) {
      ruleKey = critere;
    } else if (category) {
      ruleKey = category;
    }

    if (!ruleKey) continue;

    supplierIndexes.forEach(({ col, supplierName }) => {
      const value = normalizeText(row[col]);
      rulesBySupplier[supplierName][ruleKey] = value;
    });
  }

  return {
    fournisseurs,
    rulesBySupplier
  };
}

function parsePanelSheet(workbook) {
  const sheet = workbook.Sheets[PANEL_SHEET];
  if (!sheet) {
    throw new Error(`Onglet introuvable: ${PANEL_SHEET}`);
  }

  const data = xlsx.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  if (data.length < 2) {
    throw new Error("Onglet panel vide ou format inattendu.");
  }

  const panelBySupplier = {};

  for (let i = 1; i < data.length; i += 1) {
    const rawSupplier = normalizeText(data[i][0]);
    const rawPanel = normalizeText(data[i][1]);

    if (!rawSupplier) continue;

    const supplier = normalizeSupplierName(rawSupplier);
    const panel = rawPanel || "";

    panelBySupplier[supplier] = {
      supplier,
      panel,
      panelPriority: PANEL_PRIORITY[panel.toLowerCase()] ?? 99
    };
  }

  return panelBySupplier;
}

function getSegmentRuleKey(segment) {
  return segment;
}

function getHorizonRuleKey(energie) {
  return energie === "gaz"
    ? "HORIZON GAZ (date fin fourniture)"
    : "HORIZON ELECTRICITE (date fin fourniture)";
}

function getRuleValue(rules, candidates = []) {
  const entries = Object.entries(rules || {});
  for (const candidate of candidates) {
    const wanted = slugify(candidate);
    const found = entries.find(([key]) => slugify(key) === wanted || slugify(key).includes(wanted));
    if (found && found[1] !== undefined && found[1] !== null && String(found[1]).trim() !== "") {
      return found[1];
    }
  }
  return "";
}

function getDdfRuleValue(rules) {
  return getRuleValue(rules, [
    "DDF MAX (date début fourniture)",
    "DDF MAX (date debut fourniture)",
    "DDF MAX",
    "date début fourniture",
    "date debut fourniture"
  ]);
}

function evaluateSupplier(input) {
  const supplier = input.supplier;
  const rules = input.rules || {};
  const panelInfo = input.panelInfo || null;
  const params = input.params;

  const evaluations = [];

  const segmentEval = evaluateSegmentRule(rules[getSegmentRuleKey(params.segment)]);
  evaluations.push({ criterion: `Segment ${params.segment}`, ...segmentEval });

  const syndicEval = evaluateSyndicRule(rules["SYNDIC ?"], params.syndic, params.energie, supplier);
  evaluations.push({ criterion: "Syndic", ...syndicEval });

  // Gaz Européen fait UNIQUEMENT du Syndic → hors cible si syndic ≠ oui
  if (slugify(supplier).includes("gaz europeen") && params.syndic !== "oui") {
    evaluations.push({
      criterion: "Syndic obligatoire (Gaz Européen)",
      eligible: false,
      status: "ko",
      reason: "Gaz Européen fait uniquement du Syndic — hors cible"
    });
  }

  const horizonEval = evaluateHorizonRule(rules[getHorizonRuleKey(params.energie)], params.dffDate);
  evaluations.push({ criterion: "Horizon", ...horizonEval });

  const ddfRuleValue = getDdfRuleValue(rules);
  const ddfEval = evaluateDdfRule(ddfRuleValue, params.ddfDate);
  evaluations.push({ criterion: "DDF max", ...ddfEval });

  const scoringEval = evaluateScoringRule(rules["SCORING MINIMUM"], params.note);
  evaluations.push({ criterion: "Scoring minimum", ...scoringEval });

  const volumeEval = evaluateVolumeRule(rules["VOLUME MINIMAL (CAR en MWh)"], params.volume);
  evaluations.push({ criterion: "Volume minimal", ...volumeEval });

  const upfrontPaymentEval = evaluateUpfrontPaymentRule(rules["Paiement UPFRONT"], params.commissionEstimee, params.ddfDate, supplier, params.dffDate);
  evaluations.push({ criterion: "Paiement UPFRONT", ...upfrontPaymentEval });

  const regCommValue = getRuleValue(rules, [
  "Régularisation sur consommation",
  "Regularisation sur consommation",
  "Régularisation des commissions",
  "Regularisation des commissions"
  ]);
  const regCommEval = evaluateRegularisationCommissionsRule(regCommValue);
  evaluations.push({ criterion: "Régularisation commissions", ...regCommEval });

  // MES : uniquement si Première mise en service ou Mise en service
  if (params.mesType) {
    const mesRuleKey = params.mesType === "premiere" ? "1ère MES" : "re-MES";
    const mesLabel = params.mesType === "premiere" ? "1ère mise en service" : "Remise en service";
    const mesRuleValue = getRuleValue(rules, [mesRuleKey, mesRuleKey.replace("è", "e")]);
    const mesEval = evaluateMesRule(mesRuleValue, params.mesType, params.energie);
    if (mesEval !== null) {
      evaluations.push({ criterion: mesLabel, ...mesEval });
    }
  }

  // Marge
  const margeRuleValue = getRuleValue(rules, ["Marge", "MARGE"]);
  const margeEval = evaluateMargeRule(margeRuleValue, params.energie, params.segment, params.volume, params.margeGlobale);
  evaluations.push({ criterion: "Marge", ...margeEval });

  const eligible = evaluations.every((e) => e.eligible !== false);
  const warnings = evaluations.filter((e) => e.status === "warn").length;

  // ── SCORING MÉTIER (contextuel au dossier) ─────────────────────
  // Chaque composante évalue si le fournisseur CONVIENT à CE dossier précis.

  // 1. PAIEMENT UPFRONT (max 3 pts)
  //    OUI sans condition → 3 pts
  //    OUI avec conditions → on évalue vs le dossier :
  //      toutes respectées → 3 pts | 1 sur 2 → 2 pts | aucune → 0 pts
  //    NON → 0 pts
  let scoreUpfront = 0;
  const upfrontRaw = normalizeText(rules["Paiement UPFRONT"] || "").toUpperCase();
  if (upfrontRaw.startsWith("OUI")) {
    // Détecter les conditions — mêmes formats que evaluateUpfrontPaymentRule
    let thresholdScore = null;

    const kMatchesScore = [...upfrontRaw.matchAll(/[<≤]\s*(\d+(?:[.,]\d+)?)\s*K/gi)];
    if (kMatchesScore.length > 0) {
      const lastK = kMatchesScore[kMatchesScore.length - 1];
      thresholdScore = parseFloat(lastK[1].replace(",", ".")) * 1000;
    }
    if (thresholdScore === null) {
      const euroMatchesScore = [...normalizeText(rules["Paiement UPFRONT"] || "").toUpperCase().matchAll(/[<≤]\s*([\d\s]+(?:[.,]\d+)?)\s*€/gi)];
      if (euroMatchesScore.length > 0) {
        const lastEuro = euroMatchesScore[euroMatchesScore.length - 1];
        const cleaned = lastEuro[1].replace(/\s/g, "").replace(",", ".");
        const val = parseFloat(cleaned);
        if (Number.isFinite(val)) thresholdScore = val;
      }
    }

    const moisMatchUp = upfrontRaw.match(/DDF\s*[<≤]\s*M\s*\+\s*(\d+)/i);
    const anneesMatchUp = upfrontRaw.match(/DDF\s*[<≤]\s*N\s*\+\s*(\d+)/i);

    const hasCommCondition = thresholdScore !== null;
    const hasDdfCondition = !!(moisMatchUp || anneesMatchUp);

    if (!hasCommCondition && !hasDdfCondition) {
      // OUI sans condition → 3 pts
      scoreUpfront = 3;
    } else {
      // Évaluer chaque condition vs le dossier
      let conditionsTotal = 0;
      let conditionsOk = 0;

      if (hasCommCondition) {
        conditionsTotal++;
        if (params.commissionEstimee !== null && params.commissionEstimee !== undefined && params.commissionEstimee <= thresholdScore) {
          conditionsOk++;
        }
      }

      if (hasDdfCondition && params.ddfDate) {
        conditionsTotal++;
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const dateMax = new Date(today);
        if (moisMatchUp) {
          dateMax.setMonth(dateMax.getMonth() + parseInt(moisMatchUp[1], 10));
        } else {
          dateMax.setFullYear(dateMax.getFullYear() + parseInt(anneesMatchUp[1], 10));
        }
        if (dateMax >= params.ddfDate) {
          conditionsOk++;
        }
      }

      if (conditionsTotal > 0 && conditionsOk === conditionsTotal) {
        scoreUpfront = 3; // toutes les conditions respectées
      } else if (conditionsTotal > 1 && conditionsOk >= 1) {
        scoreUpfront = 2; // une condition sur deux respectée
      } else {
        scoreUpfront = 0; // aucune condition respectée
      }
    }
  }

  // 2. MARGE (max 2 pts) — contextuelle au dossier
  //    Marge saisie ≤ seuil → 2 pts | Marge > seuil → 0 pts
  //    Pas de seuil / grille / pas de limite → 2 pts (pas de contrainte = favorable)
  //    Marge non saisie → 2 pts (pas d'info pour pénaliser)
  let scoreMarge = 0;
  if (margeEval.status === "ok") {
    scoreMarge = 2;
  } else if (margeEval.status === "warn") {
    scoreMarge = 0;
  } else {
    // neutral = pas de contrainte ou non renseigné → favorable
    scoreMarge = 2;
  }

  // 3. HORIZON (max 2 pts) — vs la DFF du dossier
  //    Horizon ≥ année DFF → 2 pts | Horizon < DFF → 0 pts
  let scoreHorizon = 0;
  const horizonYear = getYearFromHorizon(rules[getHorizonRuleKey(params.energie)] || "");
  if (horizonYear !== null && params.dffDate) {
    const dffYear = params.dffDate.getFullYear();
    if (horizonYear >= dffYear) {
      scoreHorizon = 2;
    } else {
      scoreHorizon = 0;
    }
  } else if (horizonYear !== null) {
    // Pas de DFF saisie → on ne peut pas comparer, on reste neutre favorable
    scoreHorizon = 2;
  }

  // 4. SCORING MINIMUM (max 2 pts) — vs la note du dossier
  //    Note dossier ≥ minimum fournisseur → 2 pts | Note < minimum → 0 pts
  //    Pas de scoring renseigné → 2 pts (pas de contrainte)
  let scoringMin = 0;
  const scoringRawValue = normalizeText(rules["SCORING MINIMUM"] || "");
  const scoringMinMatch = scoringRawValue.match(/(\d+)/);
  const scoringMinVal = scoringMinMatch ? safeNumber(scoringMinMatch[1]) : null;
  if (scoringMinVal !== null && params.note !== null && params.note !== undefined) {
    scoringMin = params.note >= scoringMinVal ? 2 : 0;
  } else {
    scoringMin = 2; // pas de contrainte ou note non saisie → favorable
  }

  // 5. RÉGULARISATION (max 1 pt) : Non = 1 / Oui = 0
  let scoreRegul = 0;
  const regulRaw = normalizeText(regCommValue || "").toUpperCase();
  if (regulRaw === "NON" || regulRaw.startsWith("NON")) scoreRegul = 1;

  // Score métier total (max 10, min 0)
  // Si non éligible (au moins 1 critère KO) → forcer à 0/10
  const scoreMetier = eligible ? Math.min(Math.max(scoreUpfront + scoreMarge + scoreHorizon + scoringMin + scoreRegul, 0), 10) : 0;

  // Score de tri : scoreMetier prioritaire, warnings comme départage (moins = mieux)
  // Jamais négatif : non éligible = 0
  const score = eligible ? Math.max(scoreMetier * 1000 - warnings, 0) : 0;

  return {
    supplier,
    eligible,
    panel: panelInfo?.panel || "",
    panelPriority: panelInfo?.panelPriority ?? 99,
    score,
    scoreMetier,
    scoreDetail: { scoreUpfront, scoreMarge, scoreHorizon, scoringMin, scoreRegul },
    evaluations,
    rulesUsed: {
      segment: rules[getSegmentRuleKey(params.segment)] || "",
      upfrontPayment: rules["Paiement UPFRONT"] || "",
      syndic: rules["SYNDIC ?"] || "",
      horizon: rules[getHorizonRuleKey(params.energie)] || "",
      ddfMax: ddfRuleValue || "",
      scoring: rules["SCORING MINIMUM"] || "",
      volumeMinimal: rules["VOLUME MINIMAL (CAR en MWh)"] || ""
    }
  };
}

// ============ CHARGEMENT AVEC CACHE ============
async function loadSelectionEngine() {
  const now = Date.now();
  if (_cache && now - _cacheAt < CACHE_TTL_MS) return _cache;

  const buffer = await downloadExcelFromSharePoint();
  const workbook = xlsx.read(buffer, { type: "buffer", cellDates: false });
  const rulesData = parseRulesSheet(workbook);
  const panelData = parsePanelSheet(workbook);

  _cache = {
    fournisseurs:    rulesData.fournisseurs,
    rulesBySupplier: rulesData.rulesBySupplier,
    panelBySupplier: panelData
  };
  _cacheAt = now;
  return _cache;
}

// ============ HANDLER ============
module.exports = async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  try {
    const query = req.query || {};

    const normalizedEnergy  = normalizeEnergy(query.energie || "");
    const normalizedSegment = normalizeSegment(query.segment || "");
    const normalizedSyndic  = normalizeOuiNon(query.syndic || "");
    const currentSupplier   = normalizeSupplierName(query.fournisseur_actuel || "");

    if (!normalizedEnergy || !normalizedSegment) {
      return res.status(400).json({
        message: "Les paramètres energie et segment sont obligatoires."
      });
    }

    const engine = await loadSelectionEngine();

    const etatPdl = normalizeText(query.etat_pdl || "").toLowerCase().replace(/\s+/g, "_");
    const mesType = etatPdl === "premiere_mes" ? "premiere"
      : etatPdl === "mes" ? "remise"
      : null;

    const params = {
      energie:           normalizedEnergy,
      segment:           normalizedSegment,
      syndic:            normalizedSyndic,
      note:              safeNumber(query.note),
      volume:            safeNumber(query.volume),
      commissionEstimee: safeNumber(query.commission_estimee),
      margeGlobale:      safeNumber(query.marge_globale),
      ddfDate:           parseFrenchDate(query.ddf),
      dffDate:           parseFrenchDate(query.dff),
      mesType
    };

    const results = engine.fournisseurs.map((supplier) =>
      evaluateSupplier({
        supplier,
        rules:     engine.rulesBySupplier[supplier] || {},
        panelInfo: engine.panelBySupplier[supplier] || null,
        params
      })
    );

    const eligibleResults = results
      .filter(r => r.eligible)
      .sort((a, b) => b.score - a.score);

    const topSuppliers = eligibleResults.slice(0, 5);

    const partnerSupplier =
      currentSupplier && engine.rulesBySupplier[currentSupplier]
        ? evaluateSupplier({
            supplier:  currentSupplier,
            rules:     engine.rulesBySupplier[currentSupplier] || {},
            panelInfo: engine.panelBySupplier[currentSupplier] || null,
            params
          })
        : null;

    return res.status(200).json({
      meta: {
        fileName:       FILE_NAME,
        rulesSheet:     RULES_SHEET,
        panelSheet:     PANEL_SHEET,
        totalSuppliers: engine.fournisseurs.length
      },
      input: {
        energie:            normalizedEnergy,
        segment:            normalizedSegment,
        syndic:             normalizedSyndic,
        note:               params.note,
        volume:             params.volume,
        commissionEstimee:  params.commissionEstimee,
        ddf:                query.ddf || "",
        dff:                query.dff || "",
        fournisseur_actuel: currentSupplier || ""
      },
      allSuppliers:  results,
      topSuppliers,
      eligibleCount: eligibleResults.length,
      partnerSupplier: partnerSupplier
        ? {
            label:       "FOURNISSEUR ACTUEL",
            supplier:    partnerSupplier.supplier,
            eligible:    partnerSupplier.eligible,
            panel:       partnerSupplier.panel || "",
            evaluations: partnerSupplier.evaluations,
            score:       partnerSupplier.score,
            scoreMetier: partnerSupplier.scoreMetier,
            scoreDetail: partnerSupplier.scoreDetail
          }
        : null
    });

  } catch (error) {
    console.error("fournisseur-selection error:", error);
    return res.status(500).json({
      message: error.message || "Erreur serveur",
      detail:  String(error)
    });
  }
};
