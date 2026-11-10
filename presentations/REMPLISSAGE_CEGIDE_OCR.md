# Remplissage Cegide OCR

## Objectif

Automatiser la preparation des lignes de facturation dans Excel a partir de factures PDF d'huissiers : lire les operations, associer leurs montants HT et debours, retrouver les articles dans le referentiel Excel et renseigner les informations de module, client et signataire.

Ce projet complete les projets Excel VBA du depot. Il ne remplace pas **GLOIRE V2**, qui se concentre sur le renommage et le classement des factures. Il prepare une feuille Excel ; il ne transmet pas automatiquement les ecritures a un service Cegid.

## Source principale

[Nouvelle_Macro_Remplissage_Cegide_FINAL_TSV_MONTANTS.bas](../src/Nouvelle_Macro_Remplissage_Cegide_FINAL_TSV_MONTANTS.bas)

Module VBA : `GloireCegideOCR`. Le fichier est conserve en Windows-1252 avec fins de ligne CRLF pour son import dans l'editeur VBA Windows. Il contient un pont OCR Windows integre en Base64, extrait en executable local lors de l'installation. Ce binaire n'est pas un service cloud ; son code source de compilation n'est pas fourni dans cette publication.

## Fonctionnement

1. Choisir le dossier des factures et l'emplacement du pont OCR au premier lancement.
2. Selectionner le PDF le plus recent du dossier configure, sans le deplacer.
3. Analyser uniquement la premiere page, y compris lorsqu'il s'agit d'un scan.
4. Combiner le texte TXT et les positions du TSV pour reconstruire les lignes et colonnes.
5. Rechercher les operations dans le referentiel Excel et determiner les montants.
6. Completer les metadonnees manquantes avec une saisie utilisateur lorsque necessaire.
7. Controler le module puis remplir les lignes disponibles ou enregistrer le blocage dans MODUP.

Le TXT de disposition conserve une organisation spatiale approximative. Il ne constitue pas une reproduction visuelle identique du PDF et ne garantit pas que chaque mot a ete reconnu.

## Regles de detection

- **HT prioritaire** : reconnaissance de HT, Base HT, Hors taxe et variantes. Le TTC sert de repli seulement si la colonne HT est absente, jamais pour remplacer une cellule HT vide. La colonne Debit est exclue.
- **Debours** : Debours, NS, Non soumis, Exo et Exonere sont traites comme non soumis. Une operation comportant HT et debours produit deux lignes. Les montants explicitement nuls ne produisent pas de ligne de montant.
- **Codes articles** : recherche dans le referentiel Excel. Pour les debours, un code d'origine terminant par 3 recherche sa correspondance en 2 ; un code d'origine terminant par 1 reste en 1. Aucun code n'est invente.
- **Libelle integral** : le texte reconnu sur la facture est conserve, meme si la recherche d'article utilise une expression plus courte.
- **Montants decales** : une ligne de montants proche peut completer le libelle precedent encore sans montant, en l'absence de nouveau libelle ou de nouvelle date. Un avertissement invite a verifier l'association.
- **Montant non lu** : la ligne peut rester presente avec la mention `Montant non lu - a verifier`. Cette mention n'est pas un montant numerique et doit etre corrigee avant import comptable.
- **Module** : numero de 10 chiffres proche des references. La repetition de ses quatre premiers chiffres est un indice facultatif.
- **Client** : numero de 1 a 7 chiffres dans la zone des references. La saisie manuelle accepte `12345`, `L12345` ou `L/12345`. Q conserve `L/12345` et S contient `L12345`, avec conservation des zeros initiaux.
- **Signataire** : nom proche du module ; a defaut, correspondance du service dans Code organisation. La saisie partielle est acceptee uniquement si elle identifie une seule ligne du referentiel. Une saisie ambigue ou inconnue doit etre precisee.

## Structure du classeur attendue

Le classeur metier n'est pas inclus. Ne pas utiliser un modele dont les colonnes different sans adapter et tester le code.

| Feuille | Colonnes ou role |
| --- | --- |
| TYPE d'article a refacturer | C : operation de reference ; F : code article |
| Code organisation | M : organisation ; N : destination ; O : signataire ; P : service DTG ; R : identifiant du signataire |
| Remplissage Cegide | A : numero ; B : libelle ; C : article ; D : quantite ; G : montant ; H/I : organisation/destination ; L : module ; M/N : dates ; O : signataire ; P : etape ; Q : client ; R : identifiant signataire ; S : client corrige |
| EXPORT HUISSIER | A : numero de facture ; B : recherche/fournisseur ; D : reference, utilises pour le transfert MODUP |
| GTDTECGR | E : liste des modules provoquant le blocage |
| MODUP | A : numero de facture ; B : fournisseur/recherche ; C : reference ; M : redacteur ; N : module ; O : motif |

La feuille de resultat doit avoir une ligne numerotee disponible, avec les champs de resultat et l'etape vides. La macro ajoute les lignes necessaires. Le signataire conserve sa liste de validation. L'etape normale est `Traiter`.

**Attention aux dates** : cette version ecrit les dates fixes du 01/01/2026 au 31/12/2026. Les adapter avant une utilisation sur un autre exercice.

## Controle des modules

Selon la regle metier de ce projet, **la presence du module dans GTDTECGR!E bloque le remplissage**, avec le motif `Module non reference`. Ce n'est pas une recherche d'absence dans un catalogue de modules valides.

La macro recherche la facture correspondante dans EXPORT HUISSIER et ajoute les informations dans MODUP. Si les feuilles ou la correspondance necessaires manquent, elle signale l'erreur au lieu de choisir une facture au hasard. Un module vide ne peut pas etre compare a cette liste et doit faire l'objet d'un controle humain.

## Installation et utilisation

1. Utiliser Excel de bureau sous Windows, sur une copie de travail du classeur au format `.xlsm`.
2. Verifier les feuilles et correspondances de colonnes ci-dessus.
3. Telecharger le fichier `.bas` brut et l'importer dans l'editeur VBA via **Fichier > Importer un fichier**. Eviter les doublons du module `GloireCegideOCR`.
4. Executer `Gloire_Cegide_Choisir_Dossier_Factures` pour choisir le dossier source.
5. Executer `Gloire_Cegide_Choisir_Pont_OCR` pour configurer le pont integre et, si besoin, `Gloire_Cegide_Tester_Pont_OCR`.
6. Lancer `Gloire_Cegide_Traiter_Derniere_Facture` et verifier le PDF effectivement selectionne.
7. Controler chaque ligne, les codes, les montants et les avertissements avant toute utilisation comptable.

Le pont necessite les composants OCR Windows et les langues de reconnaissance adaptees aux factures. Respecter la politique de securite de l'organisation pour les macros et executables ; ne pas desactiver les protections du poste.

La version publiee contient aussi `Gloire_Lire_Message_Boucle`, destinee a recuperer les messages du traitement dans une automatisation externe. Cette automatisation externe n'est pas incluse.

## Validation et limites

Les essais de developpement precedents ont couvert les synonymes d'en-tetes, les codes HT/debours, les noms partiels, les clients, le blocage MODUP et les montants decales. Sur les TXT/TSV de 11 factures de travail : modules et montants HT conformes sur 11/11, clients sur 10/11 et lignes/debours sur 10/11. Un client illisible et un debours masque restaient non reconnus.

Ces resultats historiques ne sont pas une mesure de precision generale ni une certification de cette publication. Les documents de travail ne sont pas publics et aucun nouveau test Excel n'a ete lance pour la mise en ligne. Les modifications locales intervenues depuis les essais doivent etre revalidees sur le classeur cible.

Une zone masquee, une photo floue ou un mauvais alignement peut encore entrainer une omission ou une association incorrecte. La verification humaine reste indispensable, notamment pour les montants ambigus et les codes absents.

## Confidentialite

Cette publication contient uniquement le module et sa documentation. Les factures, photos, fichiers TXT/TSV issus de documents reels, classeurs, referentiels clients et journaux de production sont exclus. Ne pas ajouter ces fichiers dans les commits ou les issues.

## Competences mobilisees

Excel VBA, OCR local, extraction documentaire, coordonnees spatiales, expressions regulieres, rapprochement de referentiels, automatisation comptable et controles de coherence.
