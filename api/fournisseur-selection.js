const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const EXCEL_FILE = path.join(process.cwd(), "data", "regle_fournisseur.xlsx");
const RULES_SHEET = "Fournisseurs Export";
const PANEL_SHEET = "Fournisseurs Panel";

const PANEL_PRIORITY = {
  "gold premium": 1,
  "gold": 2,
  "silver": 3,
  "bronze": 4,
  "": 99,
  "non classe": 99
};

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
  const s = normalizeText(value);
  if (!s) return null;

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    const parts = s.split("/");
    return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
  }

  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return new Date(s.slice(0, 10));
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
    ["ohm", "OHM"],
    ["ohm energie", "OHM"],
    ["ohm gc", "OHM"],
    ["ohm mid", "OHM"],
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
    ["vattenfall c5", "VATTENFALL"],
    ["vattenfall gc", "VATTENFALL"],
    ["vattenfall", "VATTENFALL"],
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

function evaluateSyndicRule(ruleValue, syndic) {
  const value = normalizeText(ruleValue);
  const upper = value.toUpperCase();

  if (syndic !== "oui") {
    return { eligible: true, status: "ok", reason: "Critere syndic non demande" };
  }

  if (!value) return { eligible: true, status: "warn", reason: "Regle syndic non renseignee" };
  if (upper === "NON") return { eligible: false, status: "ko", reason: "Syndic non accepte" };
  if (upper === "OUI") return { eligible: true, status: "ok", reason: "Syndic accepte" };
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

  const match = value.match(/(\d+)\s*\/\s*10/);
  if (!match) return { eligible: true, status: "neutral", reason: value };

  const minimum = Number(match[1]);
  if (!Number.isFinite(minimum)) return { eligible: true, status: "neutral", reason: value };

  if (note < minimum) {
    return { eligible: false, status: "ko", reason: `Note ${note}/10 < minimum ${minimum}/10` };
  }
  if (note === minimum) {
    return { eligible: true, status: "warn", reason: `Note ${note}/10 au minimum requis` };
  }
  return { eligible: true, status: "ok", reason: `Note ${note}/10 ≥ minimum ${minimum}/10` };
}

function evaluateVolumeRule(ruleValue, volume) {
  const minVolume = extractMinVolume(ruleValue);
  if (minVolume === null || volume === null || volume === undefined) {
    return { eligible: true, status: "neutral", reason: normalizeText(ruleValue) || "Pas de volume minimum exploitable" };
  }

  if (volume < minVolume) {
    return { eligible: false, status: "ko", reason: `Volume ${volume} MWh < minimum ${minVolume} MWh` };
  }
  return { eligible: true, status: "ok", reason: `Volume ${volume} MWh ≥ minimum ${minVolume} MWh` };
}

function evaluateDdfRule(ruleValue, ddfDate) {
  const value = normalizeText(ruleValue);
  if (!value || !ddfDate) {
    return { eligible: true, status: "neutral", reason: value || "DDF max non renseignee" };
  }

  if (value.includes("M+") || value.includes("N+")) {
    return { eligible: true, status: "warn", reason: value };
  }

  if (slugify(value) === "pas de limite") {
    return { eligible: true, status: "ok", reason: "Pas de limite" };
  }

  const maxDate = parseFrenchDate(value);
  if (!maxDate) return { eligible: true, status: "warn", reason: value };

  if (ddfDate > maxDate) {
    return {
      eligible: false,
      status: "ko",
      reason: `DDF ${ddfDate.toLocaleDateString("fr-FR")} > DDF max ${maxDate.toLocaleDateString("fr-FR")}`
    };
  }

  return { eligible: true, status: "ok", reason: "DDF compatible" };
}

function evaluateHorizonRule(ruleValue, dffDate) {
  const year = getYearFromHorizon(ruleValue);
  if (!year || !dffDate) {
    return { eligible: true, status: "neutral", reason: normalizeText(ruleValue) || "Horizon non renseigne" };
  }

  const dffYear = dffDate.getFullYear();
  if (year < dffYear) {
    return { eligible: false, status: "ko", reason: `Horizon ${year} < fin fourniture ${dffYear}` };
  }
  if (year === dffYear) {
    return { eligible: true, status: "warn", reason: `Horizon ${year} = fin fourniture` };
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

function evaluateSupplier(input) {
  const supplier = input.supplier;
  const rules = input.rules || {};
  const panelInfo = input.panelInfo || null;
  const params = input.params;

  const evaluations = [];

  const segmentEval = evaluateSegmentRule(rules[getSegmentRuleKey(params.segment)]);
  evaluations.push({ criterion: `Segment ${params.segment}`, ...segmentEval });

  const syndicEval = evaluateSyndicRule(rules["SYNDIC ?"], params.syndic);
  evaluations.push({ criterion: "Syndic", ...syndicEval });

  const horizonEval = evaluateHorizonRule(rules[getHorizonRuleKey(params.energie)], params.dffDate);
  evaluations.push({ criterion: "Horizon", ...horizonEval });

  const ddfEval = evaluateDdfRule(rules["DDF MAX (date début fourniture)"], params.ddfDate);
  evaluations.push({ criterion: "DDF max", ...ddfEval });

  const scoringEval = evaluateScoringRule(rules["SCORING MINIMUM"], params.note);
  evaluations.push({ criterion: "Scoring minimum", ...scoringEval });

  const volumeEval = evaluateVolumeRule(rules["VOLUME MINIMAL (CAR en MWh)"], params.volume);
  evaluations.push({ criterion: "Volume minimal", ...volumeEval });

  const eligible = evaluations.every((e) => e.eligible !== false);
  const warnings = evaluations.filter((e) => e.status === "warn").length;

  const score = (eligible ? 100 : 0) - warnings * 5 - ((panelInfo?.panelPriority ?? 99) * 2);

  return {
    supplier,
    eligible,
    panel: panelInfo?.panel || "",
    panelPriority: panelInfo?.panelPriority ?? 99,
    score,
    evaluations,
    rulesUsed: {
      segment: rules[getSegmentRuleKey(params.segment)] || "",
      syndic: rules["SYNDIC ?"] || "",
      horizon: rules[getHorizonRuleKey(params.energie)] || "",
      ddfMax: rules["DDF MAX (date début fourniture)"] || "",
      scoring: rules["SCORING MINIMUM"] || "",
      volumeMinimal: rules["VOLUME MINIMAL (CAR en MWh)"] || ""
    }
  };
}

function loadSelectionEngine() {
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(`Fichier introuvable: ${EXCEL_FILE}`);
  }

  const workbook = xlsx.readFile(EXCEL_FILE, { cellDates: false });
  const rulesData = parseRulesSheet(workbook);
  const panelData = parsePanelSheet(workbook);

  return {
    fournisseurs: rulesData.fournisseurs,
    rulesBySupplier: rulesData.rulesBySupplier,
    panelBySupplier: panelData
  };
}

module.exports = function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  try {
    const query = req.query || {};

    const normalizedEnergy = normalizeEnergy(query.energie || "");
    const normalizedSegment = normalizeSegment(query.segment || "");
    const normalizedSyndic = normalizeOuiNon(query.syndic || "");
    const currentSupplier = normalizeSupplierName(query.fournisseur_actuel || "");

    if (!normalizedEnergy || !normalizedSegment) {
      return res.status(400).json({
        message: "Les paramètres energie et segment sont obligatoires."
      });
    }

    const engine = loadSelectionEngine();

    const params = {
      energie: normalizedEnergy,
      segment: normalizedSegment,
      syndic: normalizedSyndic,
      note: safeNumber(query.note),
      volume: safeNumber(query.volume),
      ddfDate: parseFrenchDate(query.ddf),
      dffDate: parseFrenchDate(query.dff)
    };

    const results = engine.fournisseurs.map((supplier) =>
      evaluateSupplier({
        supplier,
        rules: engine.rulesBySupplier[supplier] || {},
        panelInfo: engine.panelBySupplier[supplier] || null,
        params
      })
    );

    const eligibleResults = results
      .filter((r) => r.eligible)
      .sort((a, b) => {
        if (a.panelPriority !== b.panelPriority) {
          return a.panelPriority - b.panelPriority;
        }
        return b.score - a.score;
      });

    const topSuppliers = eligibleResults.slice(0, 5);

    const partnerSupplier =
      currentSupplier && engine.rulesBySupplier[currentSupplier]
        ? evaluateSupplier({
            supplier: currentSupplier,
            rules: engine.rulesBySupplier[currentSupplier] || {},
            panelInfo: engine.panelBySupplier[currentSupplier] || null,
            params
          })
        : null;

    return res.status(200).json({
  meta: {
    fileName: path.basename(EXCEL_FILE),
    rulesSheet: RULES_SHEET,
    panelSheet: PANEL_SHEET,
    totalSuppliers: engine.fournisseurs.length
  },
  input: {
    energie: normalizedEnergy,
    segment: normalizedSegment,
    syndic: normalizedSyndic,
    note: params.note,
    volume: params.volume,
    ddf: query.ddf || "",
    dff: query.dff || "",
    fournisseur_actuel: currentSupplier || ""
  },

  // 🔥 NOUVEAU : tous les fournisseurs
  allSuppliers: results,

  // 🔥 existant
  topSuppliers,
  eligibleCount: eligibleResults.length,

  partnerSupplier: partnerSupplier
    ? {
        label: "FOURNISSEUR PARTENAIRE",
        supplier: partnerSupplier.supplier,
        eligible: partnerSupplier.eligible,
        panel: partnerSupplier.panel || "",
        evaluations: partnerSupplier.evaluations,
        score: partnerSupplier.score // 👈 bonus utile
      }
    : null
});
  } catch (error) {
    console.error("fournisseur-selection error:", error);
    return res.status(500).json({
      message: error.message || "Erreur serveur"
    });
  }
};
