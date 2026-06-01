// js/pdfRename.js

export function buildPdfFileName({
  esiNiv3,
  diagnosticCode,
  dateJJMMAAAA,
  pcOuPp,
  reperes
}) {

  // Règle ESI selon le type de diagnostic
  // Niveau bâtiment (3 segments) : DTA, FRDTA, RADTA
  // Niveau PC (4 segments) ou PP (code complet) : tous les autres
  const segments = esiNiv3.split("-");
  const NIVEAU_BATIMENT = ["DTA", "FRDTA", "RADTA"];

  let esiBase;
  if (NIVEAU_BATIMENT.includes(diagnosticCode)) {
    esiBase = segments.slice(0, 3).join("-");
  } else if (pcOuPp === "PC") {
    esiBase = segments.slice(0, 4).join("-");
  } else {
    // PP ou valeur non reconnue → code complet
    esiBase = esiNiv3;
  }

  // Ordre de priorité (du plus défavorable au moins)
  const PRIORITY = [
    "PA3", "PA2", "PA1",
    "PB2", "PB1", "PB",
    "PC",
    "N"
  ];

  let detected = new Set();

  reperes.forEach(r => {
    const liste    = (r.liste             || "").toUpperCase();
    const etat     = (r.etatConservation  || "").toUpperCase();
    const resultat = (r.resultatAmiante   || "").toUpperCase();

    // Liste B/C → prioriser B
    const isB = liste.includes("B");
    const isC = liste.includes("C") && !isB;

    if (liste.includes("A")) {
      if      (etat.includes("3")) detected.add("PA3");
      else if (etat.includes("2")) detected.add("PA2");
      else if (etat.includes("1")) detected.add("PA1");
    }

    if (isB) {
      if      (etat.includes("AC2")) detected.add("PB2");
      else if (etat.includes("AC1")) detected.add("PB1");
      else if (etat.includes("EP"))  detected.add("PB");
    }

    // Liste C : amiante détecté seulement si le résultat
    // contient "AMIANTE" mais PAS "ABSENCE"
    // (ex : "Absence d'amiante" ne doit pas déclencher PC)
    if (isC && resultat.includes("AMIANTE") && !resultat.includes("ABSENCE")) {
      detected.add("PC");
    }
  });

  if (detected.size === 0) {
    detected.add("N");
  }

  const finalResult = PRIORITY.find(code => detected.has(code)) || "N";

  return `${esiBase}_${diagnosticCode}_${dateJJMMAAAA}_${pcOuPp}_${finalResult}.pdf`;
}
