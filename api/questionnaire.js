const CLICKUP_BASE = "https://api.clickup.com/api/v2";

function getEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Variable d'environnement manquante: ${name}`);
  }
  return value;
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  try {
    const token = getEnv("CLICKUP_TOKEN");
    const listId = getEnv("CLICKUP_LIST_QUESTIONNAIRE");

    const {
      fournisseur,
      energie,
      segment,
      contact,
      commentaire,
      priorite,
      typeDemande
    } = req.body || {};

    if (!fournisseur || !energie) {
      return res.status(400).json({
        message: "Les champs fournisseur et énergie sont obligatoires."
      });
    }

    const description = [
      `Type de demande : ${typeDemande || "-"}`,
      `Fournisseur : ${fournisseur}`,
      `Énergie : ${energie}`,
      `Segment : ${segment || "-"}`,
      `Contact : ${contact || "-"}`,
      `Priorité : ${priorite || "-"}`,
      ``,
      `Commentaire :`,
      `${commentaire || "-"}`
    ].join("\n");

    const response = await fetch(`${CLICKUP_BASE}/list/${listId}/task`, {
      method: "POST",
      headers: {
        Authorization: token,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        name: `Questionnaire fournisseur - ${fournisseur}`,
        description
      })
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      return res.status(response.status).json({
        message: data.err || data.error || "Erreur ClickUp"
      });
    }

    return res.status(200).json({
      ok: true,
      taskId: data.id || null
    });
  } catch (e) {
    return res.status(500).json({
      message: e.message || "Erreur serveur"
    });
  }
}
