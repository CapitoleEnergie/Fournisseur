import fs from "fs";
import path from "path";
import xlsx from "xlsx";

const EXCEL_FILE = path.join(
  process.cwd(),
  "data",
  "regles_all_fournisseurs.xlsx"
);

const SHEET_NAME = "Résumé global";

// Normalisation des clés pour le front
function normalizeRuleKey(label = "") {
  const key = label
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w\s]/g, "")
    .trim();

  if (key.includes("HORIZON ELECTRICITE")) return "HORIZON ELEC";
  if (key.includes("HORIZON GAZ")) return "HORIZON GAZ";
  if (key.includes("DDF MAX")) return "DDF MAX";
  if (key.includes("SCORING")) return "SCORING MINIMUM";
  if (key.includes("MARGE")) return "MARGE";
  if (key.includes("UPFRONT")) return "UPFRONT";
  if (key.includes("DECOMMISSION")) return "DÉCOMMISSION CONSO";
  if (key.includes("VOLUME MIN")) return "VOLUME MIN";
  if (key.includes("REMES")) return "re-MES";
  if (key.includes("1ERE MES")) return "1ère MES";
  if (key.includes("SYNDIC")) return "SYNDIC";

  return key;
}

// Détection segments
function normalizeSegment(label = "") {
  const clean = label.toUpperCase().trim();

  if (["C1", "C2", "C3", "C4", "C5"].includes(clean)) {
    return `ELEC_${clean}`;
  }
  if (["T1", "T2", "T3", "T4"].includes(clean)) {
    return `GAZ_${clean}`;
  }

  return null;
}

function loadRules() {
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error("Fichier Excel introuvable");
  }

  const workbook = xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.Sheets[SHEET_NAME];

  if (!sheet) {
    throw new Error("Onglet 'Résumé global' introuvable");
  }

  const data = xlsx.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  // Ligne fournisseurs (ligne 0 ou 1 selon fichier)
  const headerRow = data.find(row =>
    row.some(cell => String(cell).toUpperCase().includes("ALTERNA"))
  );

  if (!headerRow) {
    throw new Error("Impossible de trouver les fournisseurs");
  }

  const fournisseurs = headerRow.slice(1).map(f => String(f).trim());

  const rules = {};
  const normalizedRules = {};

  let currentSection = null;

  for (const row of data) {
    const label = String(row[0] || "").trim();
    if (!label) continue;

    // Détection section segments
    if (label.toUpperCase().includes("SEGMENTS ELECTRICITE")) {
      currentSection = "ELEC";
      continue;
    }
    if (label.toUpperCase().includes("SEGMENTS GAZ")) {
      currentSection = "GAZ";
      continue;
    }

    // Gestion segments
    const segmentKey = normalizeSegment(label);
    if (segmentKey) {
      normalizedRules[segmentKey] = {};
      fournisseurs.forEach((f, i) => {
        normalizedRules[segmentKey][f] = row[i + 1] || "";
      });
      continue;
    }

    // Règles générales
    const ruleKey = normalizeRuleKey(label);

    rules[label] = {};
    normalizedRules[ruleKey] = {};

    fournisseurs.forEach((f, i) => {
      const value = row[i + 1] || "";
      rules[label][f] = value;
      normalizedRules[ruleKey][f] = value;
    });
  }

  return {
    meta: {
      fileName: path.basename(EXCEL_FILE),
      sheetName: SHEET_NAME
    },
    fournisseurs,
    rules,
    normalizedRules
  };
}

export default function handler(req, res) {
  try {
    const data = loadRules();

    const fournisseur = req.query.fournisseur;

    if (!fournisseur) {
      return res.status(200).json(data);
    }

    const result = {};

    Object.keys(data.normalizedRules).forEach(rule => {
      result[rule] = data.normalizedRules[rule][fournisseur] || null;
    });

    return res.status(200).json({
      fournisseur,
      rules: result
    });

  } catch (error) {
    console.error(error);
    return res.status(500).json({
      message: error.message
    });
  }
}
