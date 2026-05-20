import path from "path";
import xlsx from "xlsx";

// ============ SHAREPOINT CONFIG ============
const TENANT_ID     = process.env.MICROSOFT_TENANT_ID;
const CLIENT_ID     = process.env.MICROSOFT_CLIENT_ID;
const CLIENT_SECRET = process.env.MICROSOFT_CLIENT_SECRET;
const DRIVE_ID      = process.env.SP_DRIVE_ID;
const FOLDER_PATH   = process.env.SP_FOLDER_PATH || "PARTAGE/Team/3. Fournisseurs/Data Fournisseurs";
const FILE_NAME     = "regles_all_fournisseurs.xlsx";
const SHEET_NAME    = "Résumé global";

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

  // Récupérer l'URL de téléchargement du fichier
  const fileRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${FOLDER_PATH}/${FILE_NAME}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const fileMeta = await fileRes.json();

  if (!fileMeta["@microsoft.graph.downloadUrl"]) {
    throw new Error(`Fichier SharePoint introuvable : ${FILE_NAME}`);
  }

  // Télécharger le contenu binaire
  const dlRes = await fetch(fileMeta["@microsoft.graph.downloadUrl"]);
  const arrayBuffer = await dlRes.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

// ============ NORMALISATION ============
function normalizeRuleKey(label = "") {
  const key = label
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w\s]/g, "")
    .trim();

  if (key.includes("HORIZON ELECTRICITE")) return "HORIZON ELEC";
  if (key.includes("HORIZON GAZ"))         return "HORIZON GAZ";
  if (key.includes("DDF MAX"))             return "DDF MAX";
  if (key.includes("SCORING"))             return "SCORING MINIMUM";
  if (key.includes("MARGE"))               return "MARGE";
  if (key.includes("UPFRONT"))             return "UPFRONT";
  if (key.includes("DECOMMISSION"))        return "DÉCOMMISSION CONSO";
  if (key.includes("VOLUME MIN"))          return "VOLUME MIN";
  if (key.includes("REMES"))               return "re-MES";
  if (key.includes("1ERE MES"))            return "1ère MES";
  if (key.includes("SYNDIC"))              return "SYNDIC";

  return key;
}

function normalizeSegment(label = "") {
  const clean = label.toUpperCase().trim();
  if (["C1", "C2", "C3", "C4", "C5"].includes(clean)) return `ELEC_${clean}`;
  if (["T1", "T2", "T3", "T4"].includes(clean))        return `GAZ_${clean}`;
  return null;
}

// ============ LECTURE DU WORKBOOK ============
function parseRules(buffer) {
  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheet = workbook.Sheets[SHEET_NAME];

  if (!sheet) throw new Error("Onglet 'Résumé global' introuvable");

  const data = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const headerRow = data.find(row =>
    row.some(cell => String(cell).toUpperCase().includes("ALTERNA"))
  );
  if (!headerRow) throw new Error("Impossible de trouver les fournisseurs");

  const fournisseurs = headerRow.slice(1).map(f => String(f).trim());
  const rules = {};
  const normalizedRules = {};
  let currentSection = null;

  for (const row of data) {
    const label = String(row[0] || "").trim();
    if (!label) continue;

    if (label.toUpperCase().includes("SEGMENTS ELECTRICITE")) { currentSection = "ELEC"; continue; }
    if (label.toUpperCase().includes("SEGMENTS GAZ"))         { currentSection = "GAZ";  continue; }

    const segmentKey = normalizeSegment(label);
    if (segmentKey) {
      normalizedRules[segmentKey] = {};
      fournisseurs.forEach((f, i) => {
        normalizedRules[segmentKey][f] = row[i + 1] || "";
      });
      continue;
    }

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
    meta: { fileName: FILE_NAME, sheetName: SHEET_NAME },
    fournisseurs,
    rules,
    normalizedRules
  };
}

// ============ CHARGEMENT AVEC CACHE ============
async function loadRules() {
  const now = Date.now();
  if (_cache && now - _cacheAt < CACHE_TTL_MS) return _cache;

  const buffer = await downloadExcelFromSharePoint();
  _cache = parseRules(buffer);
  _cacheAt = now;
  return _cache;
}

// ============ HANDLER ============
export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  try {
    const data = await loadRules();
    const fournisseur = req.query.fournisseur;

    if (!fournisseur) {
      return res.status(200).json(data);
    }

    const result = {};
    Object.keys(data.normalizedRules).forEach(rule => {
      result[rule] = data.normalizedRules[rule][fournisseur] || null;
    });

    return res.status(200).json({ fournisseur, rules: result });

  } catch (error) {
    console.error("fournisseur-rules error:", error);
    return res.status(500).json({ message: error.message });
  }
}
