import fs from "fs";
import path from "path";
import xlsx from "xlsx";

const EXCEL_FILE = path.join(process.cwd(), "data", "fournisseurs_export new .xlsx");
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

function isBlank(value) {
  return value === null || value === undefined || normalizeText(value) === "";
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
    const [d, m, y] = s.split("/");
    return new Date(Number(y), Number(m) - 1, Number(d));
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

function normalizeOfferValue(value) {
  const s = normalizeText(value);
  if (!s) return "";

  if (
    s.toUpperCase() === "OUI" ||
    s.toUpperCase() === "NON" ||
    s.toUpperCase().startsWith("OUI ") ||
    s.toUpperCase().startsWith("OUI(") ||
    s.toUpperCase().startsWith("NON ") ||
    s.toUpperCase().startsWith("NON(") ||
    s.toUpperCase().includes("CAS PAR CAS") ||
    s.toUpperCase().includes("EN PAUSE")
  ) {
    return s;
  }

  return s;
}

function isStrictNo(value) {
  const s = normalizeText(value).toUpperCase();
  return s === "NON";
}

function isYesLike(value) {
  const s = normalizeText(value).toUpperCase();
  return s === "OUI" || s.startsWith("OUI");
}

function isConditional(value) {
  const s = normalizeText(value).toUpperCase();
  if (!s) return false;
  return (
    s.includes("CAS PAR CAS") ||
    s.includes("EN PAUSE") ||
    s.startsWith("OUI") && s !== "OUI" ||
    s.startsWith("NON") && s !== "NON"
  );
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
    ["societe d approvisionnement et de vente d energies", "SAVE"],
    ["save", "SAVE"],
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
  const value = normalizeOfferValue(ruleValue);
  const upper = value.toUpperCase();

  if (!value) {
    return {
      eligible: false,
      status: "ko",
      reason: "Segment non renseigné"
    };
  }

  if (upper === "OUI") {
    return {
      eligible: true,
      status: "ok",
      reason: "Segment couvert"
    };
  }

  if (upper === "NON") {
    return {
      eligible: false,
      status: "ko",
      reason: "Segment non couvert"
    };
  }

  if (upper.includes("EN PAUSE")) {
    return {
      eligible: false,
      status: "ko",
      reason: value
    };
  }

  if (upper.includes("CAS PAR CAS") || upper.startsWith("OUI") || upper.startsWith("NON")) {
    return {
      eligible: true,
      status: "warn",
      reason: value
    };
  }

  return {
    eligible: true,
    status: "warn",
    reason: value
  };
}

function evaluateSyndicRule(ruleValue, syndic) {
  const value = normalizeOfferValue(ruleValue);
  const upper = value.toUpperCase();

  if (syndic !== "oui") {
    return {
      eligible: true,
      status: "ok",
      reason: "Critère syndic non demandé"
    };
  }

  if (!value) {
    return {
      eligible: true,
      status: "warn",
      reason: "Règle syndic non renseignée"
    };
  }

  if (upper === "NON") {
    return {
      eligible: false,
      status: "ko",
      reason: "Syndic non accepté"
    };
  }

  if (upper === "OUI") {
    return {
      eligible: true,
      status: "ok",
      reason: "Syndic accepté"
    };
  }

  if (upper.includes("CAS PAR CAS") || upper.startsWith("OUI")) {
    return {
      eligible: true,
      status: "warn",
      reason: value
    };
  }

  return {
    eligible: true,
    status: "warn",
    reason: value
  };
}

function evaluateScoringRule(ruleValue, note) {
  const value = normalizeText(ruleValue);
  if (!value || note === null || note === undefined) {
    return { eligible: true, status: "neutral", reason: value || "Scoring non renseigné" };
  }

  const match = value.match(/(\d+)\s*\/\s*10/);
  if (!match) {
    return { eligible: true, status: "neutral", reason: value };
  }

  const minimum = Number(match[1]);
  if (!Number.isFinite(minimum)) {
    return { eligible: true, status: "neutral", reason: value };
  }

  if (note < minimum) {
    return {
      eligible: false,
      status: "ko",
      reason: `Note ${note}/10 < minimum ${minimum}/10`
    };
  }

  if (note === minimum) {
    return {
      eligible: true,
      status: "warn",
      reason: `Note ${note}/10 au minimum requis`
    };
  }

  return {
    eligible: true,
    status: "ok",
    reason: `Note ${note}/10 ≥ minimum ${minimum}/10`
  };
}

function evaluateVolumeRule(ruleValue, volume) {
  const minVolume = extractMinVolume(ruleValue);
  if (minVolume === null || volume === null || volume === undefined) {
    return { eligible: true, status: "neutral", reason: normalizeText(ruleValue) || "Pas de volume minimum exploitable" };
  }

  if (volume < minVolume) {
    return {
      eligible: false,
      status: "ko",
      reason: `Volume ${volume} MWh < minimum ${minVolume} MWh`
    };
  }

  return {
    eligible: true,
    status: "ok",
    reason: `Volume ${volume} MWh ≥ minimum ${minVolume} MWh`
  };
}

function evaluateDdfRule(ruleValue, ddfDate) {
  const value = normalizeText(ruleValue);
  if (!value || !ddfDate) {
    return { eligible: true, status: "neutral", reason: value || "DDF max non renseignée" };
  }

  if (value.includes("M+") || value.includes("N+")) {
    return { eligible: true, status: "warn", reason: value };
  }

  if (slugify(value) === "pas de limite") {
    return { eligible: true, status: "ok", reason: "Pas de limite" };
  }

  const maxDate = parseFrenchDate(value);
  if (!maxDate) {
    return { eligible: true, status: "warn", reason: value };
  }

  if (ddfDate > maxDate) {
    return {
      eligible: false,
      status: "ko",
      reason: `DDF ${ddfDate.toLocaleDateString("fr-FR")} > DDF max ${maxDate.toLocaleDateString("fr-FR")}`
    };
  }

  return {
    eligible: true,
    status: "ok",
    reason: `DDF compatible`
  };
}

function evaluateHorizonRule(ruleValue, dffDate) {
  const year = getYearFromHorizon(ruleValue);
  if (!year || !dffDate) {
    return { eligible: true, status: "neutral", reason: normalizeText(ruleValue) || "Horizon non renseigné" };
  }

  const dffYear = dffDate.getFullYear();
  if (year < dffYear) {
    return {
      eligible: false,
      status: "ko",
      reason: `Horizon ${year} < fin fourniture ${dffYear}`
    };
  }

  if (year === dffYear) {
    return {
      eligible: true,
      status: "warn",
      reason: `Horizon ${year} = fin fourniture`
    };
  }

  return {
    eligible: true,
    status: "ok",
    reason: `Horizon ${year} couvre la période`
  };
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
    throw new Error("Onglet règles vide ou format inattendu.");
  }

  const headerRowIndex = 2;
  const headerRow = data[headerRowIndex];
  const fournisseurs = headerRow.slice(3).map((name) => normalizeSupplierName(name)).filter(Boolean);

  const supplierIndexes = [];
  for (let col = 3; col < headerRow.length; col += 1) {
    const supplierName = normalizeSupplierName(headerRow[col]);
    if (!supplierName) continue;
    supplierIndexes.push({ col, supplierName });
  }

  const rulesBySupplier = {};
  fournisseurs.forEach((f) => {
    rulesBySupplier[f] = {};
  });

  let currentCategory = "";

  for (let rowIdx = headerRowIndex + 1; rowIdx < data.length; rowIdx += 1) {
    const row = data[rowIdx];
    const category = normalizeText(row[0]);
    const critere = normalizeText(row[1]);
    const sousCritere = normalizeText(row[2]);

    if (category) {
      currentCategory = category;
    }

    if (!category && !critere) continue;

    let ruleKey = "";
    if (currentCategory.toUpperCase().includes("SEGMENTS") && critere) {
      ruleKey = critere.toUpperCase();
    } else if (category && !critere) {
      ruleKey = category;
    } else if (critere) {
      ruleKey = critere;
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

  // On démarre à 1 pour skip le header
  for (let i = 1; i < data.length; i++) {
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
    console.log("Panel loaded:", panelBySupplier);
  return panelBySupplier;
  
}
  const data = xlsx.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  if (!data.length) {
    throw new Error("Onglet panel vide ou format inattendu.");
  }

  const headerIndex = data.findIndex((row) => {
    const first = slugify(row[0]);
    const second = slugify(row[1]);
    return first.includes("fournisseur") && second.includes("panel");
  });

  if (headerIndex === -1) {
    throw new Error("Impossible de trouver l'en-tête de l'onglet Fournisseurs Panel.");
  }

  const panelBySupplier = {};

  for (let i = headerIndex + 1; i < data.length; i += 1) {
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

  const data = xlsx.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  const panelBySupplier = {};

  for (let i = 0; i < data.length; i += 1) {
    const rawSupplier = normalizeText(data[i][0]);
    const rawPanel = normalizeText(data[i][5]);

    if (!rawSupplier) continue;
    if (slugify(rawSupplier) === "fournisseur classement") continue;
    if (slugify(rawSupplier) === "couleur") continue;

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

function getSegmentRuleKey(energie, segment) {
  return segment;
}

function getHorizonRuleKey(energie) {
  return energie === "gaz"
    ? "HORIZON GAZ (date fin fourniture)"
    : "HORIZON ELECTRICITE (date fin fourniture)";
}

function evaluateSupplier({ supplier, rules, panelInfo, params }) {
  const {
    energie,
    segment,
    syndic,
    note,
    volume,
    ddfDate,
    dffDate
  } = params;

  const evaluations = [];

  const segmentEval = evaluateSegmentRule(rules[getSegmentRuleKey(energie, segment)]);
  evaluations.push({ criterion: `Segment ${segment}`, ...segmentEval });

  const syndicEval = evaluateSyndicRule(rules["SYNDIC ?"], syndic);
  evaluations.push({ criterion: "Syndic", ...syndicEval });

  const horizonEval = evaluateHorizonRule(rules[getHorizonRuleKey(energie)], dffDate);
  evaluations.push({ criterion: "Horizon", ...horizonEval });

  const ddfEval = evaluateDdfRule(rules["DDF MAX (date début fourniture)"], ddfDate);
  evaluations.push({ criterion: "DDF max", ...ddfEval });

  const scoringEval = evaluateScoringRule(rules["SCORING MINIMUM"], note);
  evaluations.push({ criterion: "Scoring minimum", ...scoringEval });

  const volumeEval = evaluateVolumeRule(rules["VOLUME MINIMAL (CAR en MWh)"], volume);
  evaluations.push({ criterion: "Volume minimal", ...volumeEval });

  const eligible = evaluations.every((e) => e.eligible !== false);
  const warnings = evaluations.filter((e) => e.status === "warn").length;

  const score =
    (eligible ? 100 : 0) -
    warnings * 5 -
    ((panelInfo?.panelPriority ?? 99) * 2);

  return {
    supplier,
    eligible,
    panel: panelInfo?.panel || "",
    panelPriority: panelInfo?.panelPriority ?? 99,
    score,
    evaluations,
    rulesUsed: {
      segment: rules[getSegmentRuleKey(energie, segment)] || "",
      syndic: rules["SYNDIC ?"] || "",
      horizon: rules[getHorizonRuleKey(energie)] || "",
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

export default function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  try {
    const {
      energie = "",
      segment = "",
      syndic = "",
      note = "",
      volume = "",
      ddf = "",
      dff = "",
      fournisseur_actuel = ""
    } = req.query;

    const normalizedEnergy = normalizeEnergy(energie);
    const normalizedSegment = normalizeSegment(segment);
    const normalizedSyndic = normalizeOuiNon(syndic);
    const currentSupplier = normalizeSupplierName(fournisseur_actuel);

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
      note: safeNumber(note),
      volume: safeNumber(volume),
      ddfDate: parseFrenchDate(ddf),
      dffDate: parseFrenchDate(dff)
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
        ddf,
        dff,
        fournisseur_actuel: currentSupplier || ""
      },
      topSuppliers,
      eligibleCount: eligibleResults.length,
      partnerSupplier: partnerSupplier
        ? {
            label: "FOURNISSEUR PARTENAIRE",
            supplier: partnerSupplier.supplier,
            eligible: partnerSupplier.eligible,
            panel: partnerSupplier.panel || "",
            evaluations: partnerSupplier.evaluations
          }
        : null
    });
  } catch (error) {
    console.error("fournisseur-selection error:", error);
    return res.status(500).json({
      message: error.message || "Erreur serveur"
    });
  }
}
