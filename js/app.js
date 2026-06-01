// Export Amiante SIA — Prototype v0.1
// Tout local (navigateur). Règles conformes au cadrage défini dans la conversation.
import { buildPdfFileName } from "./pdfRename.js";

const state = {
  // --- fichiers & XML ---
  files: [],
  xmlByKey: {},
  extracted: null,

  // --- Identification SIA (éditable UI) ---
  identification: {
    dateDiag: "",
    esiNiv3: "",
    agence: "",
    commune: "",
    codePostal: "",
    adresse: "",
    residence: ""
  }
};

let pdfOriginalSpan = null;
let pdfFinalSpan = null;


const $ = (id) => document.getElementById(id);

// -------------------------------
// Nomenclature 46.020 (TSV embarqué)
// Colonnes: Famille \t Ouvrages \t Parties
// -------------------------------

const NOMENCLATURE_46020_TSV = `Famille de composant\tOuvrages ou Composants de la construction\tParties d’ouvrages ou de composants à inspecter ou à sonder
1 - Couvertures, Toitures, Terrasses et étanchéités\tPlaques ondulées et planes\tPlaques en fibres‐ciment (y compris plaques « sous tuiles »)
1 - Couvertures, Toitures, Terrasses et étanchéités\tPlaques ondulées et planes\tPlaques en matériaux bitumineux
1 - Couvertures, Toitures, Terrasses et étanchéités\tPlaques ondulées et planes\tRevêtements anti condensation sous bac acier
1 - Couvertures, Toitures, Terrasses et étanchéités\tArdoises, bardeaux bitumineux\tArdoises composites hors fibro ciment
1 - Couvertures, Toitures, Terrasses et étanchéités\tArdoises, bardeaux bitumineux\tArdoises en fibro ciment
1 - Couvertures, Toitures, Terrasses et étanchéités\tArdoises, bardeaux bitumineux\tBardeaux bitumineux (« shingles »)
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tConduits de fumée, de cheminée, de ventilation
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tConduits d'eaux pluviales
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tGarnissage des joints de dilatation
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tJoints de dilatation
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tCouvre‐joints
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tTresses d'étanchéité à l'air
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tÉléments complémentaires de toiture (chéneaux, rives, closoirs, faitages, mîtres, costières, etc.)
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tJonctions bitumineuses
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tSolins en fibre ciment
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments associés à la toiture\tColle des solins en fibre ciment
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments sous toiture\tPare‐vapeur, pare pluie
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments sous toiture\tIsolants fibreux en sous toiture
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉléments sous toiture\tFlocages, enduits projetés
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉtanchéité de toiture terrasse\tParties planes : revêtements bitumineux (bandes, lés…), écrans de semi indépendance, pare‐vapeur
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉtanchéité de toiture terrasse\tRelevés : revêtements bitumineux (bandes, lés…)
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉtanchéité de toiture terrasse\tParties planes ou relevés : complexes asphaltés
1 - Couvertures, Toitures, Terrasses et étanchéités\tÉtanchéité de toiture terrasse\tColles, produits d'accrochage
1 - Couvertures, Toitures, Terrasses et étanchéités\tFenêtres de toit, lanternaux, verrières\tMastics (vitriers, bitumineux…)
1 - Couvertures, Toitures, Terrasses et étanchéités\tFenêtres de toit, lanternaux, verrières\tJoints d'étanchéité entre menuiserie et ossature
1 - Couvertures, Toitures, Terrasses et étanchéités\tFenêtres de toit, lanternaux, verrières\tGarnitures de friction sur fenêtres basculantes
2 - Parois verticales extérieures et Façades\tFaçades légères, murs rideaux, bardages, panneaux sandwich\tPlaques, panneaux, bacs en fibres‐ciment, éléments de remplissage
2 - Parois verticales extérieures et Façades\tFaçades légères, murs rideaux, bardages, panneaux sandwich\tArdoises composites hors fibro ciment
2 - Parois verticales extérieures et Façades\tFaçades légères, murs rideaux, bardages, panneaux sandwich\tArdoises en fibro ciment
2 - Parois verticales extérieures et Façades\tFaçades légères, murs rideaux, bardages, panneaux sandwich\tJoints d'assemblage ou d'étanchéité, mastics, tresses
2 - Parois verticales extérieures et Façades\tFaçades légères, murs rideaux, bardages, panneaux sandwich\tRevêtements intérieurs anti condensation (hors peintures)
2 - Parois verticales extérieures et Façades\tFaçades légères, murs rideaux, bardages, panneaux sandwich\tPeintures des bardages métalliques
2 - Parois verticales extérieures et Façades\tIsolant et protection thermique ou acoustique sous bardage\tFlocages, enduits projetés
2 - Parois verticales extérieures et Façades\tIsolant et protection thermique ou acoustique sous bardage\tCarton‐amiante
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tEnduits extérieurs (projetés, lissés ou talochés), crépis extérieurs
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tRevêtements plastiques épais (RPE)
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tPeintures sur béton
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tEnduits pelliculaires de lissage/débullage
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tColles et joints (faience, pâte de verre, carrelage), ragréages, primaires d'accrochage, Imperméabilisants
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tÉléments en maçonnerie silico‐calcaire (1880‐1940) briques blanches silico‐calcaire
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tGarnissage des joints de dilatation
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tJoints de dilatation
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tCouvre‐joints
2 - Parois verticales extérieures et Façades\tFaçades lourdes y compris poteaux\tAppuis de fenêtres en fibres‐ciment
2 - Parois verticales extérieures et Façades\tMenuiseries extérieures\tJoints de mastic de vitrage (notamment châssis aluminium)
2 - Parois verticales extérieures et Façades\tMenuiseries extérieures\tJoints d'étanchéité entre menuiserie et structure
2 - Parois verticales extérieures et Façades\tMenuiseries extérieures\tGarnitures de friction sur fenêtres basculantes
2 - Parois verticales extérieures et Façades\tMenuiseries extérieures\tPlaques de fibres‐ciment (allèges, coffres, etc.)
2 - Parois verticales extérieures et Façades\tMenuiseries extérieures\tPeintures décoratives
2 - Parois verticales extérieures et Façades\tÉléments associés aux façades\tConduits de fumées, de cheminée, de ventilation
2 - Parois verticales extérieures et Façades\tÉléments associés aux façades\tConduits d'eaux (pluviales et usées)
2 - Parois verticales extérieures et Façades\tÉléments associés aux façades\tÉléments ponctuels : chéneaux, rives, corniches
3 - Parois verticales intérieures\tMurs et cloisons maçonnés\tFlocages
3 - Parois verticales intérieures\tMurs et cloisons maçonnés\tEnduits à base de plâtre ou ciment projetés, lissés ou talochés
3 - Parois verticales intérieures\tMurs et cloisons maçonnés\tEnduits de ragréage, débullage, lissage
3 - Parois verticales intérieures\tMurs et cloisons maçonnés\tJoints de dilatation, d'assemblage, joints coupe‐ feu
3 - Parois verticales intérieures\tMurs et cloisons maçonnés\tFourreaux (carton, fibres‐ciment…)
3 - Parois verticales intérieures\tPoteaux\tFlocages
3 - Parois verticales intérieures\tPoteaux\tEnduits à base de plâtre projetés, lissés ou talochés
3 - Parois verticales intérieures\tPoteaux\tEnduits à base de ciment, lissés ou talochés (ragréage, débullage, lissage)
3 - Parois verticales intérieures\tPoteaux\tJoints de dilatation, d'assemblage avec poutraison
3 - Parois verticales intérieures\tPoteaux\tEntourages de poteau (carton‐amiante, fibres‐ ciment, matériaux sandwich…), coffrages perdus
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tPanneaux de cloisons lisses ou moulurées, préfabriquées ou non
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tEnduits à base de plâtre ou ciment projetés, lissés ou talochés
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tFlocages
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tPlots de colle fixant les cloisons au mur
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tBandes calicot
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tenduits de jointoiement des plaques de plâtre
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tSous couches des tissus muraux
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tIsolants intérieurs fibreux, bourre en vrac
3 - Parois verticales intérieures\tCloisons sèches (assemblées, préfabriquées)\tJonctions entre panneaux préfabriqués et pieds / têtes de cloisons (notamment IGH et ERP): tresse, carton, fibres‐ciment
3 - Parois verticales intérieures\tGaines et coffres verticaux\tFlocages
3 - Parois verticales intérieures\tGaines et coffres verticaux\tEnduits à base de plâtre (projetés, lissés ou talochés)
3 - Parois verticales intérieures\tGaines et coffres verticaux\tEnduits à base de ciment, lissés ou talochés (ragréage, débullage, lissage)
3 - Parois verticales intérieures\tGaines et coffres verticaux\tBandes calicot,
3 - Parois verticales intérieures\tGaines et coffres verticaux\tenduits de jointoiement des plaques de plâtre cartonné
3 - Parois verticales intérieures\tGaines et coffres verticaux\tPanneaux (fibres‐ciment, …)
3 - Parois verticales intérieures\tGaines et coffres verticaux\tJonctions entre panneaux (tresses, étanchéité entre panneaux)
3 - Parois verticales intérieures\tPortes coupe‐feu, pare‐flamme, isothermiques, frigorifiques\tJoints des portes coupe‐feu, phoniques ou pare flammes (sur battant et dormant y compris occulus, et sur serrurerie)
3 - Parois verticales intérieures\tPortes coupe‐feu, pare‐flamme, isothermiques, frigorifiques\tPanneaux, plaques en fibres‐ciment des vantaux, bakelite
3 - Parois verticales intérieures\tPortes coupe‐feu, pare‐flamme, isothermiques, frigorifiques\tIsolants intérieurs des portes
3 - Parois verticales intérieures\tRevêtements de murs, poteaux, cloisons, gaines, coffres\tSous couches des tissus muraux, moquettes murales ou les vinyles
3 - Parois verticales intérieures\tRevêtements de murs, poteaux, cloisons, gaines, coffres\tPanneaux décoratifs en fibre‐ciment (lambris), revêtements durs en fibres‐ciment
3 - Parois verticales intérieures\tRevêtements de murs, poteaux, cloisons, gaines, coffres\tColles et joints de carrelage ou de faïence, ragréage, primaire d'accrochage
3 - Parois verticales intérieures\tRevêtements de murs, poteaux, cloisons, gaines, coffres\tPeintures décoratives (pailletées, gouttelettes, …)
3 - Parois verticales intérieures\tRevêtements de murs, poteaux, cloisons, gaines, coffres\tRevêtements bitumineux
3 - Parois verticales intérieures\tRevêtements de murs, poteaux, cloisons, gaines, coffres\tPeintures intumescentes
4 - Plafonds et faux plafonds\tPlafonds\tFlocages
4 - Plafonds et faux plafonds\tPlafonds\tEnduits à base de plâtre ou ciment projetés, lissés ou talochés
4 - Plafonds et faux plafonds\tPlafonds\tPanneaux collés vissés ou cloués
4 - Plafonds et faux plafonds\tPlafonds\tCoffrages perdus (carton‐amiante, fibres‐ciment, composite)
4 - Plafonds et faux plafonds\tPlafonds\tBandes calicot
4 - Plafonds et faux plafonds\tPlafonds\tEnduits de jointoiement et plots de colle des plaques de plâtre
4 - Plafonds et faux plafonds\tPlafonds\tSous couches des tissus muraux
4 - Plafonds et faux plafonds\tPlafonds\tPeintures intumescentes
4 - Plafonds et faux plafonds\tPlafonds\tRevêtements bitumineux
4 - Plafonds et faux plafonds\tPlafonds\tPeintures décoratives (pailletées, gouttelettes…)
4 - Plafonds et faux plafonds\tPlafonds\tRésines
4 - Plafonds et faux plafonds\tPlafonds\tColles de carrelage, ragréages, primaires d'accrochage et joints de carrelage
4 - Plafonds et faux plafonds\tPoutres et charpentes\tFlocages
4 - Plafonds et faux plafonds\tPoutres et charpentes\tEnduits à base de plâtre ou ciment (projetés, lissés ou talochés)
4 - Plafonds et faux plafonds\tPoutres et charpentes\tEntourages de poutres (carton‐amiante, fibres‐ ciment, matériaux sandwich)
4 - Plafonds et faux plafonds\tPoutres et charpentes\tPeintures intumescentes
4 - Plafonds et faux plafonds\tPoutres et charpentes\tRevêtements bitumineux
4 - Plafonds et faux plafonds\tPoutres et charpentes\tPeintures décoratives (pailletées, gouttelettes…)
4 - Plafonds et faux plafonds\tPoutres et charpentes\tJonctions avec la façade, calfeutrements, joints (coupe‐feu, de dilatation, de structure)
4 - Plafonds et faux plafonds\tGaines et coffres horizontaux\tFlocages
4 - Plafonds et faux plafonds\tGaines et coffres horizontaux\tEnduits à base de plâtre ou ciment (projetés, lissés ou talochés)
4 - Plafonds et faux plafonds\tGaines et coffres horizontaux\tBandes calicot
4 - Plafonds et faux plafonds\tGaines et coffres horizontaux\tEnduits de jointoiement des plaques de plâtre cartonné
4 - Plafonds et faux plafonds\tGaines et coffres horizontaux\tPanneaux (fibres‐ciment, …)
4 - Plafonds et faux plafonds\tGaines et coffres horizontaux\tJonctions entre panneaux (tresses, étanchéité entre panneaux)
4 - Plafonds et faux plafonds\tFaux plafonds\tPanneaux et plaques
4 - Plafonds et faux plafonds\tFaux plafonds\tJonctions entre faux plafond et structure, joints entre panneaux
4 - Plafonds et faux plafonds\tFaux plafonds\tPare vapeur
4 - Plafonds et faux plafonds\tFaux plafonds\tIsolants posés dans le plénum au‐dessus du panneau de faux plafond
4 - Plafonds et faux plafonds\tFaux plafonds\tÉcrans de cantonnement et leurs joints (dans le plénum entre le faux plafond et le plancher supérieur)
4 - Plafonds et faux plafonds\tSuspentes et contrevents\tFlocages
4 - Plafonds et faux plafonds\tSuspentes et contrevents\tProtections en plâtre
4 - Plafonds et faux plafonds\tSuspentes et contrevents\tPeintures intumescentes
5 - Planchers et planchers techniques\tRevêtements de sols\tDalles de sol
5 - Planchers et planchers techniques\tRevêtements de sols\tNez de marche
5 - Planchers et planchers techniques\tRevêtements de sols\tDalles moquettes avec entrecouche noire
5 - Planchers et planchers techniques\tRevêtements de sols\tSous‐couches (carton, feutre, …) des revêtements souples
5 - Planchers et planchers techniques\tRevêtements de sols\tColles bitumineuses
5 - Planchers et planchers techniques\tRevêtements de sols\tColles non bitumineuses
5 - Planchers et planchers techniques\tRevêtements de sols\tMoquette
5 - Planchers et planchers techniques\tRevêtements de sols\tSols coulés à base ciment (terrazolith, etc.)
5 - Planchers et planchers techniques\tRevêtements de sols\tPeintures de sol
5 - Planchers et planchers techniques\tRevêtements de sols\tColles et joints de carrelage, ragréages, primaires d'accrochage
5 - Planchers et planchers techniques\tRevêtements de sols\tRevêtements de sols sportifs
5 - Planchers et planchers techniques\tRevêtements de sols\tJoints de dilatation et d'assemblage
5 - Planchers et planchers techniques\tRevêtements de sols\tJoints de cantonnement sur faux planchers
5 - Planchers et planchers techniques\tRevêtements de sols\tEnduit de cuvelage
5 - Planchers et planchers techniques\tRevêtements de sols\tRebouchages autour de conduits (principalement IGH et ERP), fourreaux en carton ou fibres‐ciment
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tCalorifuges (tresses, coquilles, matelas…)
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tMatelas
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tEnveloppes (bandes tissées enduites ou non), colles de calorifugeage
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tJoints entre éléments, joints plats prédécoupés pour brides
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tRubans adhésifs
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tMastics
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tConduits en fibres‐ciment, manchons
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tConduits en fibres‐bitumes (conduits de drainage)
6 - Conduits et accessoires intérieurs\tConduits de fluides (air, eau, vapeur, fumée, échappement, autres fluides)\tTresses dans câbles électriques d'alimentation, (notamment de secours, souvent orange), résistants au feu
6 - Conduits et accessoires intérieurs\tClapets / volets coupe‐feu\tClapets (tunnels, lames, joints)
6 - Conduits et accessoires intérieurs\tClapets / volets coupe‐feu\tVolets coupe‐feu y compris ossature
6 - Conduits et accessoires intérieurs\tClapets / volets coupe‐feu\tRebouchages et calfeutrements de clapets et volets coupe‐feu
6 - Conduits et accessoires intérieurs\tVide‐ordures\tConduits et vidoirs en fibres‐ciment
6 - Conduits et accessoires intérieurs\tVide‐ordures\tJoints d'étanchéité des trappes
7 - Ascenseurs, monte-charges et escaliers mécaniques\tPortes et cloisons palières\tPanneaux dans les portes palières
7 - Ascenseurs, monte-charges et escaliers mécaniques\tPortes et cloisons palières\tPanneaux des cloisons palières
7 - Ascenseurs, monte-charges et escaliers mécaniques\tparois des équipements\tPlaques, panneaux décoratifs (habillages cabines, joues des escaliers mécaniques…)
7 - Ascenseurs, monte-charges et escaliers mécaniques\tparois des équipements\tCalfeutrement entre mur et plancher (joint, bourre)
7 - Ascenseurs, monte-charges et escaliers mécaniques\tparois des équipements\tIsolants
7 - Ascenseurs, monte-charges et escaliers mécaniques\tparois des équipements\tColles
7 - Ascenseurs, monte-charges et escaliers mécaniques\tparois des équipements\tJoints
7 - Ascenseurs, monte-charges et escaliers mécaniques\tMatériels en machinerie\tFreins d'ascenseurs
7 - Ascenseurs, monte-charges et escaliers mécaniques\tMatériels en machinerie\tÉléments de protection contre les arcs électriques intégrés dans des équipements de type contacteurs, sélecteurs, coupe‐circuits…
7 - Ascenseurs, monte-charges et escaliers mécaniques\tMatériels en machinerie\tTresses
7 - Ascenseurs, monte-charges et escaliers mécaniques\tMatériels en machinerie\tJoints plats
8 - Équipements divers et accessoires\tChaudières (mixtes, collectives), chauffe bains, radiateurs gaz modulables, Poêles à bois à fuel, à charbon, Groupes électrogènes\tFlocages
8 - Équipements divers et accessoires\tChaudières (mixtes, collectives), chauffe bains, radiateurs gaz modulables, Poêles à bois à fuel, à charbon, Groupes électrogènes\tBourres
8 - Équipements divers et accessoires\tChaudières (mixtes, collectives), chauffe bains, radiateurs gaz modulables, Poêles à bois à fuel, à charbon, Groupes électrogènes\tTresses
8 - Équipements divers et accessoires\tChaudières (mixtes, collectives), chauffe bains, radiateurs gaz modulables, Poêles à bois à fuel, à charbon, Groupes électrogènes\tCalorifugeages
8 - Équipements divers et accessoires\tChaudières (mixtes, collectives), chauffe bains, radiateurs gaz modulables, Poêles à bois à fuel, à charbon, Groupes électrogènes\tJoints d'étanchéité, joints plats prédécoupés pour brides
8 - Équipements divers et accessoires\tChaudières (mixtes, collectives), chauffe bains, radiateurs gaz modulables, Poêles à bois à fuel, à charbon, Groupes électrogènes\tDispositifs anti condensation (peintures, films, etc.)
8 - Équipements divers et accessoires\tChaudières (mixtes, collectives), chauffe bains, radiateurs gaz modulables, Poêles à bois à fuel, à charbon, Groupes électrogènes\tTissus, soufflets amortisseurs acoustiques
8 - Équipements divers et accessoires\tConvecteurs et radiateurs électriques\tIsolants thermiques cartonnés
8 - Équipements divers et accessoires\tConvecteurs et radiateurs électriques\tTresses des diffuseurs
8 - Équipements divers et accessoires\tfusibles à broche\tCarton, tresse
8 - Équipements divers et accessoires\tcanalisations électriques préfabriquées\tIsolants
8 - Équipements divers et accessoires\tCoffres‐forts\tPortes et parois
8 - Équipements divers et accessoires\tPortes de placard, baignoires et éviers métalliques\tPlaques souples bitumineuses antivibratiles
8 - Équipements divers et accessoires\tJardinières, bac à sable incendie\tÉléments en fibres‐ciment
9 - Fondations et soubassements\tÉtanchéité des murs enterrés\tEnduits bitumineux des ouvrages enterrés
9 - Fondations et soubassements\tParois verticales et horizontales enterrées\tJoints de fractionnement, de rupture, de dilatation
9 - Fondations et soubassements\tConduits et fourreaux\tFourreaux en fibres‐ciment dans maçonnerie
10 - Aménagements, voiries et réseaux divers\tConduits, Siphons\tÉléments de canalisations enterrés en fibres‐ ciment
10 - Aménagements, voiries et réseaux divers\tVoiries\tEnrobés bitumineux des couches de voirie (juste partie bitume), asphaltes
10 - Aménagements, voiries et réseaux divers\tEspaces sportifs\tRevêtements de sols
10 - Aménagements, voiries et réseaux divers\tAménagements extérieurs\tÉléments en fibres‐ciment (jardinières, bordures…)`;

function formatDateJJMMAAAA(dateStr) {
  // Attend idéalement "dd/mm/yyyy" (ce que renvoie ton parseDateFR)
  if (!dateStr) return "";

  const m = dateStr.trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const dd = m[1], mm = m[2], yyyy = m[3];
    return `${dd}${mm}${yyyy}`; // JJMMAAAA
  }

  // Fallback : tente de parser comme Date
  const d = new Date(dateStr);
  if (!isNaN(d)) {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = String(d.getFullYear());
    return `${dd}${mm}${yyyy}`;
  }

  // Sinon, renvoie brut (mais ça évite le crash)
  return dateStr.replace(/\D/g, "");
}

function updatePdfFinalName() {
  if (!state.extracted) return;
  if (!pdfFinalSpan) return;

  const diagSelect = document.getElementById("diagnosticType");
  if (!diagSelect) return;

  const identification = state.extracted.identification;
  const reperes = state.extracted.reperes || [];

  const dateJJMMAAAA = formatDateJJMMAAAA(identification.dateDiag);
  if (!dateJJMMAAAA) {
    pdfFinalSpan.textContent = "—";
    return;
  }

  const pcOuPp = reperes[0]?.pcOuPp || "PC";

  const baseName = getClientBaseFileName();
if (!baseName) {
  pdfFinalSpan.textContent = "—";
  return;
}

pdfFinalSpan.textContent = baseName + ".pdf";

}

function getClientBaseFileName() {
  if (!state.extracted) return null;

  const diagSelect = document.getElementById("diagnosticType");
  if (!diagSelect) return null;

  const identification = state.extracted.identification;
  const reperes = state.extracted.reperes || [];

  const dateJJMMAAAA = formatDateJJMMAAAA(identification.dateDiag);
  if (!dateJJMMAAAA) return null;

  const pcOuPp = reperes[0]?.pcOuPp || "PC";

  // ⚠️ buildPdfFileName retourne un nom AVEC .pdf
  const pdfName = buildPdfFileName({
    esiNiv3: identification.esiNiv3,
    diagnosticCode: diagSelect.value,
    dateJJMMAAAA,
    pcOuPp,
    reperes
  });

  // On enlève l’extension .pdf
  return pdfName.replace(/\.pdf$/i, "");
}

function handlePdfRename(pdfFile, diagnosticCode) {
  const identification = state.extracted.identification;
  const reperes = state.extracted.reperes;

  const date = formatDateJJMMAAAA(identification.dateDiag);

  const pcOuPp = reperes[0]?.pcOuPp || "PC";

  const newName = buildPdfFileName({
    esiNiv3: identification.esiNiv3,
    diagnosticCode,
    dateJJMMAAAA: date,
    pcOuPp,
    reperes
  });

  // Affichage du nom final
  if (pdfFinalSpan) {
    pdfFinalSpan.textContent = newName;
  }

  // Téléchargement
  const blob = new Blob([pdfFile], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = newName;
  a.click();

  URL.revokeObjectURL(url);
}

function parseNomenclatureTSV(tsv) {
  const lines = tsv.split(/\r?\n/).filter(Boolean);
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split("\t");
    if (cols.length < 3) continue;
    out.push({
      famille: cols[0].trim(),
      ouvrage: cols[1].trim(),
      partie: cols.slice(2).join("\t").trim(),
    });
  }
  return out;
}

// -------------------------------
// File handling
// -------------------------------

function setStatus(targetId, msg, kind = "") {
  const el = document.getElementById(targetId);
  if (!el) return; // sécurité

  el.className = "status" + (kind ? " " + kind : "");
  el.textContent = msg;
}



// -------------------------------
// Decoding + XML parsing
// -------------------------------

async function readFileSmart(file) {
  const buf = await file.arrayBuffer();

  // 1) UTF-8
  let text = new TextDecoder("utf-8").decode(buf);
  // If replacement chars appear, try windows-1252 (common in legacy exports)
  if (text.includes("\uFFFD")) {
    const alt = new TextDecoder("windows-1252").decode(buf);
    // Keep alt if it looks better
    if (!alt.includes("\uFFFD")) text = alt;
  }
  return text;
}

function parseXmlString(text) {
  const parser = new DOMParser();
  const xml = parser.parseFromString(text, "application/xml");
  const err = xml.querySelector("parsererror");
  if (err) throw new Error("XML invalide ou mal formé");
  return xml;
}

function pickXmlFile(nameContains) {
  const needle = nameContains.toLowerCase();
  return state.files.find((f) => f.name.toLowerCase().includes(needle));
}

async function loadRequiredXml() {
  const general = pickXmlFile("table_general_bien");
  const amiante = pickXmlFile("table_z_amiante.xml");
  const docRemis = pickXmlFile("table_z_amiante_doc_remis");

  const missing = [];
  if (!general) missing.push("Table_General_Bien.xml");
  if (!amiante) missing.push("Table_Z_Amiante.xml");
  if (!docRemis) missing.push("Table_Z_Amiante_doc_remis.xml");

  if (missing.length) {
    setStatus("analysisStatus", `Fichiers manquants : ${missing.join(", ")}`, "err");
    throw new Error("Fichiers requis manquants");
  }

  const [generalText, amianteText, docText] = await Promise.all([
    readFileSmart(general),
    readFileSmart(amiante),
    readFileSmart(docRemis),
  ]);

  state.xmlByKey.general = parseXmlString(generalText);
  state.xmlByKey.amiante = parseXmlString(amianteText);
  state.xmlByKey.doc = parseXmlString(docText);
}

function qText(xml, selector) {
  const n = xml.querySelector(selector);
  return n && n.textContent ? n.textContent.trim() : "";
}

// -------------------------------
// Extractors — Onglet 1
// -------------------------------

function parseDateFR(value) {
  // Liciel sample: 29/09/2025. We keep as is if already fr.
  if (!value) return "";

  // If already dd/mm/yyyy
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) return value;

  // Try Date parsing
  const d = new Date(value);
  if (!isNaN(d)) return d.toLocaleDateString("fr-FR");

  return value;
}

function extractIdentificationSIA(generalXml) {
  return {
    dateDiag: parseDateFR(qText(generalXml, "LiColonne_Mission_Date_Visite")),
    esiNiv3: qText(generalXml, "LiColonne_Immeuble_Lot"),
    agence: "", // à définir plus tard
    commune: qText(generalXml, "LiColonne_Immeuble_Commune"),
    codePostal: qText(generalXml, "LiColonne_Immeuble_Departement"),
    adresse: qText(generalXml, "LiColonne_Immeuble_Adresse1"),
    residence: "", // à définir plus tard
  };
}

// -------------------------------
// Extractors — Onglet 2
// -------------------------------

function extractEtudeDocumentaire(docXml) {
  const rows = [];
  const items = docXml.querySelectorAll("LiItem_table_Z_Amiante_doc_remis");
  items.forEach((item) => {
    const typeRapport = qText(item, "LiColonne_Doc_Remis");
    const commentaires = qText(item, "LiColonne_Doc_Demandes");

    rows.push({
      nomRapport: "",
      typeRapport,
      date: "",
      editeur: "",
      commentaires,
    });
  });
  return rows;
}

// -------------------------------
// Extractors — Onglet 3
// -------------------------------

function mapPcOuPp(generalXml) {
  const v = qText(generalXml, "LiColonne_Immeuble_Nature_bien");
  if (v === "Habitation (partie privative d'immeuble)") return "PP";
  if (v === "Habitation (parties communes)") return "PC";
  return "";
}

function extractCodeFromEsi(esi) {
  if (!esi) return "";
  const m = esi.trim().match(/-(\d{3})$/);
  return m ? m[1] : "";
}

function splitOuvrages(ouvrages) {
  const v = (ouvrages || "").trim();
  if (!v) return { famille: "", composant: "" };

  // Split on '-' and rebuild: family = part0 + ' - ' + part1, component = rest
  const parts = v.split("-").map((s) => s.trim()).filter(Boolean);
  if (parts.length >= 3) {
    return {
      famille: `${parts[0]} - ${parts[1]}`,
      composant: parts.slice(2).join(" - "),
    };
  }
  if (parts.length === 2) {
    return { famille: `${parts[0]} - ${parts[1]}`, composant: "" };
  }
  return { famille: v, composant: "" };
}

function volumesFromLocalisation(localisation) {
  const raw = (localisation || "").trim();
  if (!raw) return [];

  return raw
    .split(";")
    .map((seg) => {
      const s = seg.trim();
      if (!s) return "";
      const idx = s.indexOf("-");
      if (idx === -1) return s;
      return s.slice(idx + 1).trim();
    })
    .filter((x) => x.length > 0);
}

function extractReperesAmiante(generalXml, amianteXml) {
  const esi = qText(generalXml, "LiColonne_Immeuble_Lot");
  const pcOuPp = mapPcOuPp(generalXml);
  const code = extractCodeFromEsi(esi);

  const out = [];
  const items = amianteXml.querySelectorAll("LiItem_table_Z_Amiante");

  items.forEach((item) => {
    const localisation = qText(item, "LiColonne_Localisation");
    const volumes = volumesFromLocalisation(localisation);
    const ouvrages = qText(item, "LiColonne_Ouvrages");
    const split = splitOuvrages(ouvrages);

    const partieInspectee = qText(item, "LiColonne_Partie_Inspectee");
    const description = qText(item, "LiColonne_Description");
    const liste = qText(item, "LiColonne_ListeCSP_amiante");
    const resultats = qText(item, "LiColonne_Resultats");
    const justification = qText(item, "LiColonne_Justification");
    const etatConservation = qText(item, "LiColonne_Etat_Conservation");

    // If no ';' and no '-', volumesFromLocalisation may still return 1 entry.
    const vols = volumes.length ? volumes : [""];

    vols.forEach((nomVolume) => {
      out.push({
        // A-R
        esiComplet: esi,
        entree: "", // à venir
        pcOuPp,
        etage: "", // à venir
        code,
        nomVolume,
        familleComposant: split.famille,
        ouvragesComposants: split.composant,
        partiesInspecter: partieInspectee,
        coucheAnalysee: description,
        commentaire: "", // colonne K (à venir)
        liste,
        resultatAmiante: resultats,
        origineConclusion: justification,
        etatConservation,
        score: "", // à venir
        mesures: "", // à venir
        remarque: "", // à venir
      });
    });
  });

  return out;
}

// -------------------------------
// Workbook builders
// -------------------------------

function sheetFromAOA(aoa) {
  return XLSX.utils.aoa_to_sheet(aoa);
}

function buildWorkbook(extracted) {
  const wb = XLSX.utils.book_new();

  // Onglet 1 — Identification SIA
  const idHeaders = [
    "Date diag",
    "ESI NIV 3",
    "Agence",
    "Commune",
    "Code postal",
    "Adresse",
    "Résidence",
  ];
  const idRow = [
    extracted.identification.dateDiag,
    extracted.identification.esiNiv3,
    extracted.identification.agence,
    extracted.identification.commune,
    extracted.identification.codePostal,
    extracted.identification.adresse,
    extracted.identification.residence,
  ];
  XLSX.utils.book_append_sheet(wb, sheetFromAOA([idHeaders, idRow]), "Identification SIA");

  // Onglet 2 — Étude documentaire
  const docHeaders = ["Nom du rapport", "Type de rapport", "Date", "Editeur", "Commentaires"]; // Editeur sans accent selon capture
  const docRows = extracted.etude.map((r) => [r.nomRapport, r.typeRapport, r.date, r.editeur, r.commentaires]);
  XLSX.utils.book_append_sheet(wb, sheetFromAOA([docHeaders, ...docRows]), "Etude documentaire");

  // Onglet 3 — Repères amiante
  const repHeaders = [
    "ESI COMPLET",
    "Entrée",
    "PC ou PP",
    "Etage",
    "Code",
    "Nom du volume",
    "Famille de composant",
    "Ouvrages ou Composants de la construction",
    "Parties d’ouvrages ou de composants à inspecter ou à sonder",
    "Couche analysée",
    "Commentaire",
    "Liste",
    "Résultat amiante",
    "Origine de la conclusion",
    "Etat de conservation",
    "Score",
    "Mesures",
    "Remarque",
  ];
  const repRows = extracted.reperes.map((r) => [
    r.esiComplet,
    r.entree,
    r.pcOuPp,
    r.etage,
    r.code,
    r.nomVolume,
    r.familleComposant,
    r.ouvragesComposants,
    r.partiesInspecter,
    r.coucheAnalysee,
    r.commentaire,
    r.liste,
    r.resultatAmiante,
    r.origineConclusion,
    r.etatConservation,
    r.score,
    r.mesures,
    r.remarque,
  ]);
  XLSX.utils.book_append_sheet(wb, sheetFromAOA([repHeaders, ...repRows]), "Repères amiante");

  // Onglet 4 — Nomenclature 46.020
  const nom = extracted.nomenclature;
  const nomHeaders = [
    "Famille de composant",
    "Ouvrages ou Composants de la construction",
    "Parties d’ouvrages ou de composants à inspecter ou à sonder",
  ];
  const nomRows = nom.map((x) => [x.famille, x.ouvrage, x.partie]);
  XLSX.utils.book_append_sheet(wb, sheetFromAOA([nomHeaders, ...nomRows]), "Nomenclature 46.020");

  return wb;
}

function fileNameFromId(identification) {
  const safe = (s) => (s || "").toString().replace(/[^a-zA-Z0-9_-]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
  const esi = safe(identification.esiNiv3);
  const date = safe(identification.dateDiag);
  const base = ["Export_SIA", esi || "ESI", date || "DATE"].join("_");
  return base + ".xlsx";
}

// -------------------------------
// Analyze + Export
// -------------------------------

function renderIdentificationPreview() {
  const container = document.getElementById("previewArea");

  container.innerHTML = `
    <table class="preview-table">
      <thead>
        <tr>
          <th>Date diag</th>
          <th>ESI NIV 3</th>
          <th>Agence</th>
          <th>Commune</th>
          <th>Code postal</th>
          <th>Adresse</th>
          <th>Résidence</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td class="locked">${state.identification.dateDiag}</td>
          <td class="locked">${state.identification.esiNiv3}</td>
          <td class="editable" contenteditable data-field="agence">${state.identification.agence}</td>
          <td class="locked">${state.identification.commune}</td>
          <td class="locked">${state.identification.codePostal}</td>
          <td class="locked">${state.identification.adresse}</td>
          <td class="editable" contenteditable data-field="residence">${state.identification.residence}</td>
        </tr>
      </tbody>
    </table>
  `;

  document.querySelectorAll("[contenteditable][data-field]").forEach(cell => {
    cell.addEventListener("input", () => {
      const field = cell.dataset.field;
      state.identification[field] = cell.textContent.trim();
      updatePdfFinalName(); // 👈 MAJ immédiate du nom PDF
    });
  });
}

function handleMissionFolder(e) {
  const files = [...e.target.files];

  // On ne garde que les fichiers dans /XML
  const xmlFiles = files.filter(f =>
    f.webkitRelativePath && f.webkitRelativePath.includes("/XML/")
  );

  if (!xmlFiles.length) {
    setStatus("xmlStatus", "Aucun fichier XML trouvé dans le dossier /XML.", "err");
    return;
  }

  state.files = xmlFiles;
  setStatus("xmlStatus", `${xmlFiles.length} fichier(s) XML détecté(s).`, "ok");

  // Active l'étape Analyse
  document.getElementById("step2").classList.remove("disabled");
}

async function analyzeXml() {
  try {
   setStatus("analysisStatus", "Analyse des fichiers XML en cours…");

    await loadRequiredXml();

    const identification = extractIdentificationSIA(state.xmlByKey.general);
    const etude = extractEtudeDocumentaire(state.xmlByKey.doc);
    const reperes = extractReperesAmiante(state.xmlByKey.general, state.xmlByKey.amiante);
    const nomenclature = parseNomenclatureTSV(NOMENCLATURE_46020_TSV);

    state.extracted = {
      identification,
      etude,
      reperes,
      nomenclature,
    };

    // Synchronise Identification avec l’UI éditable
    state.identification = { ...identification };

    renderIdentificationPreview();
updatePdfFinalName();

    document.getElementById("step3").classList.remove("disabled");
    document.getElementById("step4").classList.remove("disabled");
document.getElementById("step5").classList.remove("disabled");

   setStatus("analysisStatus", "Analyse terminée. Prévisualisation disponible.", "ok");


  } catch (err) {
    console.error(err);
    setStatus("analysisStatus", "Erreur lors de l’analyse XML.", "err");
  }
}

function exportExcel() {
  if (!state.extracted) {
    alert("Aucune donnée à exporter.");
    return;
  }

  // Nom client calculé dynamiquement
  const baseName = getClientBaseFileName();
  if (!baseName) {
    alert("Nom client non disponible.");
    return;
  }

  // Injecte les valeurs éditées
  state.extracted.identification = { ...state.identification };

  const wb = buildWorkbook(state.extracted);

  // 🔥 même nom que le PDF client
  const excelFileName = baseName + ".xlsx";

  XLSX.writeFile(wb, excelFileName);
}



// -------------------------------
// UI — Previews par onglet
// -------------------------------

function renderDocsPreview() {
  const c = document.getElementById("previewArea");

  if (!state.extracted) {
    c.innerHTML = "<p>Aucune donnée.</p>";
    return;
  }

  const rows = state.extracted.etude.map(r => `
    <tr>
      <td>${r.nomRapport}</td>
      <td>${r.typeRapport}</td>
      <td>${r.date}</td>
      <td>${r.editeur}</td>
      <td>${r.commentaires}</td>
    </tr>
  `).join("");

  c.innerHTML = `
    <table class="preview-table">
      <thead>
        <tr>
          <th>Nom du rapport</th>
          <th>Type</th>
          <th>Date</th>
          <th>Éditeur</th>
          <th>Commentaires</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderReperesPreview() {
  const c = document.getElementById("previewArea");

  if (!state.extracted) {
    c.innerHTML = "<p>Aucune donnée.</p>";
    return;
  }

  const rows = state.extracted.reperes.slice(0, 100).map(r => `
    <tr>
      <td>${r.nomVolume}</td>
      <td>${r.familleComposant}</td>
      <td>${r.ouvragesComposants}</td>
      <td>${r.partiesInspecter}</td>
      <td>${r.resultatAmiante}</td>
      <td>${r.etatConservation}</td>
    </tr>
  `).join("");

  c.innerHTML = `
    <p><strong>${state.extracted.reperes.length} lignes</strong> (aperçu 100)</p>
    <table class="preview-table">
      <thead>
        <tr>
          <th>Volume</th>
          <th>Famille</th>
          <th>Ouvrage</th>
          <th>Partie inspectée</th>
          <th>Résultat</th>
          <th>État</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}


function renderNomenclaturePreview() {
  const c = document.getElementById("previewArea");

  if (!state.extracted) {
    c.innerHTML = "<p>Aucune donnée.</p>";
    return;
  }

  const rows = state.extracted.nomenclature.slice(0, 100).map(n => `
    <tr>
      <td>${n.famille}</td>
      <td>${n.ouvrage}</td>
      <td>${n.partie}</td>
    </tr>
  `).join("");

  c.innerHTML = `
    <p><strong>${state.extracted.nomenclature.length} lignes</strong> (aperçu 100)</p>
    <table class="preview-table">
      <thead>
        <tr>
          <th>Famille</th>
          <th>Ouvrage</th>
          <th>Partie</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}


function handleTabClick(tabName) {
  document.querySelectorAll(".tab").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.tab === tabName);
  });

  switch (tabName) {
    case "identification":
      renderIdentificationPreview();
      break;
    case "docs":
      renderDocsPreview();
      break;
    case "reperes":
      renderReperesPreview();
      break;
    case "nomenclature":
      renderNomenclaturePreview();
      break;
  }
}

// -------------------------------
// DOM READY
// -------------------------------

document.addEventListener("DOMContentLoaded", () => {

  // --------------------
  // Références PDF
  // --------------------
  pdfOriginalSpan = document.getElementById("pdfOriginalName");
  pdfFinalSpan = document.getElementById("pdfFinalName");

  const pdfInput = document.getElementById("pdfInput");
  const diagSelect = document.getElementById("diagnosticType");

  if (diagSelect) {
  diagSelect.addEventListener("change", () => {
    updatePdfFinalName();
  });
}

  const renameBtn = document.getElementById("renamePdfBtn");

 if (pdfInput) {
  pdfInput.addEventListener("change", () => {
    if (pdfInput.files.length && pdfOriginalSpan) {
      pdfOriginalSpan.textContent = pdfInput.files[0].name;
      updatePdfFinalName();
    }
  });
}

  if (renameBtn) {
    renameBtn.addEventListener("click", () => {
      if (!state.extracted) {
        alert("Analyse XML requise avant le renommage PDF.");
        return;
      }

      if (!pdfInput.files.length) {
        alert("Sélectionne un fichier PDF.");
        return;
      }

      handlePdfRename(pdfInput.files[0], diagSelect.value);
    });
  }

  // --------------------
  // Sélection mission
  // --------------------
  const missionFolder = document.getElementById("missionFolder");
  if (missionFolder) {
    missionFolder.addEventListener("change", handleMissionFolder);
  }

  // --------------------
  // Analyse XML
  // --------------------
  const analyzeBtn = document.getElementById("analyzeBtn");
  if (analyzeBtn) {
    analyzeBtn.addEventListener("click", analyzeXml);
  }

  // --------------------
  // Export Excel
  // --------------------
  const exportBtn = document.getElementById("exportBtn");
  if (exportBtn) {
    exportBtn.addEventListener("click", exportExcel);
  }

  // --------------------
  // Onglets preview
  // --------------------
  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => {
      handleTabClick(btn.dataset.tab);
    });
  });

});
