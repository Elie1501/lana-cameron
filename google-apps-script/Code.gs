/**
 * Réponses du mariage Lana & Cameron → Google Sheet
 * Sheet : « Réponses mariage — Lana & Cameron »
 *
 * Installation (une seule fois) :
 *  1. Ouvrir le Google Sheet → Extensions → Apps Script
 *  2. Remplacer tout le code par ce fichier, enregistrer
 *  3. Déployer → Nouveau déploiement → type « Application Web »
 *     Exécuter en tant que : Moi — Accès : Tout le monde
 *  4. Autoriser, puis copier l'URL qui finit par /exec
 */

const SHEET_ID = "16EsuQgkQ_2PTlFOBsbl0H42MM95OSeUP_P6mxLpIJLc";
const ONGLET = "Réponses";
const VERSION = 2;

function doPost(e) {
  const p = (e && e.parameter) || {};
  // Champ piège invisible : rempli seulement par les robots
  if (p.website) return json({ ok: true });

  const nom = clean(p.nom), prenom = clean(p.prenom);
  if (!nom && !prenom) return json({ ok: false, error: "nom manquant" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ONGLET);
    sheet.appendRow([
      new Date(),
      nom,
      prenom,
      clean(p.mairie),
      num(p.nb_mairie),
      clean(p.henne),
      num(p.nb_henne),
      clean(p.houppa),
      num(p.nb_houppa),   // un seul nombre pour la houppa et la soirée
      clean(p.message)
    ]);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

// Ouvrir l'URL /exec dans un navigateur affiche la version en ligne
function doGet() {
  return json({ ok: true, version: VERSION });
}

function clean(v) {
  // Empêche qu'un texte commençant par = + - @ soit lu comme une formule
  const s = String(v || "").trim().slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function num(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? "" : n;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
