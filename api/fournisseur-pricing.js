import fs from "fs";
import path from "path";
import xlsx from "xlsx";

const EXCEL_FILE = path.join(process.cwd(), "data", "Top fournisseurs.xlsx");
const SHEET_NAME = "1.Top fournisseurs C4";

function safeNumber(value) {
  if (value === null || value === undefined || value === "") return null;

  const cleaned = String(value)
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, "")
    .replace(/€/g, "")
    .replace(/MWh/gi, "")
    .replace(",", ".")
    .trim();

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function round(value, digits = 2) {
  if (value === null || value === undefined || Number.isNaN(value)) return null;
  return Number(Number(value).toFixed(digits));
}

function slugify(str = "") {
  return String(str)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/&/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function normalizeSupplierName(raw = "") {
  const cleaned = String(raw).trim();
  const key = slugify(cleaned);

  const aliases = new Map([
    ["endesa energia succursal france", "ENDESA"],
    ["sefe energy sas", "SEFE"],
    ["totalenergies electricite et gaz france", "TOTALENERGIES"],
    ["alterna energie", "ALTERNA"],
    ["engie", "ENGIE"],
    ["energem", "ENERGEM"],
    ["mint", "MINT"],
    ["ilek", "ILEK"],
    ["gedia energies services", "GEDIA"],
    ["enovos luxembourg s a", "ENOVOS"],
    ["primeo energie france", "PRIMEO"],
    ["ohm energie", "OHM"],
    ["dyneff s a s", "DYNEFF"],
    ["bcm energy", "BCM ENERGY"],
    ["hellio solutions", "HELLIO"],
    ["la bellenergie", "LA BELLENERGIE"],
    ["societe d approvisionnement et de vente d energies", "SAVE"],
    ["geg source d energies", "GEG"],
    ["synelva", "SYNELVA"]
  ]);

  return aliases.get(key) || cleaned.toUpperCase();
}

function extractDateRange(text = "") {
  const match = String(text).match(/\((\d{2}\/\d{2}\/\d{4})\s+to\s+(\d{2}\/\d{2}\/\d{4})\)/i);
  if (!match) return { startDate: null, endDate: null };
  return { startDate: match[1], endDate: match[2] };
}

function extractGeneratedAt(text = "") {
  const match = String(text).match(/As of (\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2})/i);
  if (!match) return { generatedAt: null, updatedAtLabel: null };

  const [, isoDate, time] = match;
  const [year, month, day] = isoDate.split("-");
  return {
    generatedAt: `${isoDate}T${time}`,
    updatedAtLabel: `${day}/${month}/${year}`
  };
}

function buildMarketAverage(rows) {
  const valid = rows.filter(
    (r) =>
      Number.isFinite(r.price30j) &&
      Number.isFinite(r.avgVolume) &&
      Number.isFinite(r.offerCount)
  );

  const denominator = valid.reduce(
    (sum, r) => sum + r.avgVolume * r.offerCount,
    0
  );

  if (!denominator) return null;

  const numerator = valid.reduce(
    (sum, r) => sum + r.price30j * r.avgVolume * r.offerCount,
    0
  );

  return numerator / denominator;
}

function loadWorkbookData() {
  if (!fs.existsSync(EXCEL_FILE)) {
    throw new Error(`Fichier introuvable: ${EXCEL_FILE}`);
  }

  const workbook = xlsx.readFile(EXCEL_FILE, { cellDates: false });
  const sheet = workbook.Sheets[SHEET_NAME];

  if (!sheet) {
    throw new Error(`Onglet introuvable: ${SHEET_NAME}`);
  }

  const matrix = xlsx.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: false
  });

  const generatedInfo = extractGeneratedAt(matrix?.[2]?.[1] || "");
  const periodInfo = extractDateRange(matrix?.[7]?.[1] || "");

  const rows = xlsx.utils.sheet_to_json(sheet, {
    range: 14,
    defval: null,
    raw: false
  });

  const segmentKey = Object.keys(rows[0] || {}).find((k) =>
    k.includes("Compteur segment")
  );
  const supplierKey = Object.keys(rows[0] || {}).find((k) =>
    k.includes("Nom Fournisseur")
  );
  const priceKey = Object.keys(rows[0] || {}).find((k) =>
    k.includes("Average Prix Moyen Pondéré Non Margé")
  );
  const volumeKey = Object.keys(rows[0] || {}).find((k) =>
    k.includes("Average Volume du compteur")
  );
  const countKey = Object.keys(rows[0] || {}).find((k) =>
    k.includes("Record Count")
  );

  if (!segmentKey || !supplierKey || !priceKey || !volumeKey || !countKey) {
    throw new Error("Colonnes Excel introuvables ou format inattendu.");
  }

  let currentSegment = null;
  const normalizedRows = [];

  for (const row of rows) {
    const rawSegment = row[segmentKey];
    const rawSupplier = row[supplierKey];

    if (rawSegment && String(rawSegment).trim()) {
      currentSegment = String(rawSegment).trim();
    }

    if (!rawSupplier || !String(rawSupplier).trim()) {
      continue;
    }

    if (
      String(currentSegment).toLowerCase().includes("subtotal") ||
      String(rawSupplier).toLowerCase().includes("subtotal")
    ) {
      continue;
    }

    const price30j = (row[priceKey]);
    const avgVolume = (row[volumeKey]);
    const offerCount = (row[countKey]);

    normalizedRows.push({
      segment: currentSegment,
      supplierRaw: String(rawSupplier).trim(),
      supplier: normalizeSupplierName(rawSupplier),
      price30j,
      avgVolume,
      offerCount
    });
  }

  const bySegment = new Map();

  for (const row of normalizedRows) {
    if (!bySegment.has(row.segment)) {
      bySegment.set(row.segment, []);
    }
    bySegment.get(row.segment).push(row);
  }

  const segments = {};

  for (const [segment, segmentRows] of bySegment.entries()) {
    const marketAverage = buildMarketAverage(segmentRows);

    segments[segment] = segmentRows
      .map((row) => {
        const deltaPct =
          marketAverage && row.price30j
            ? ((row.price30j - marketAverage) / marketAverage) * 100
            : null;

        return {
          ...row,
          marketAverage: round(marketAverage, 2),
          deltaPct: round(deltaPct, 1)
        };
      })
      .sort((a, b) => {
        if (a.price30j === null) return 1;
        if (b.price30j === null) return -1;
        return a.price30j - b.price30j;
      });
  }

  return {
    meta: {
      fileName: path.basename(EXCEL_FILE),
      sheetName: SHEET_NAME,
      ...periodInfo,
      ...generatedInfo,
      segments: Object.keys(segments).sort()
    },
    segments
  };
}

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  try {
    const data = loadWorkbookData();
    const segment = String(req.query.segment || "").trim().toUpperCase();
    const fournisseur = String(req.query.fournisseur || "").trim();

    if (!segment) {
      return res.status(200).json(data);
    }

    const segmentRows = data.segments[segment] || [];

    if (!fournisseur) {
      return res.status(200).json({
        meta: data.meta,
        segment,
        rows: segmentRows
      });
    }

    const supplierNeedle = normalizeSupplierName(fournisseur);
    const row =
      segmentRows.find((r) => r.supplier === supplierNeedle) ||
      segmentRows.find((r) => slugify(r.supplierRaw) === slugify(fournisseur)) ||
      null;

    return res.status(200).json({
      meta: data.meta,
      segment,
      fournisseur: supplierNeedle,
      row,
      rows: segmentRows
    });
  } catch (error) {
    console.error("fournisseur-pricing error:", error);
    return res.status(500).json({
      message: error.message || "Erreur serveur"
    });
  }
}
