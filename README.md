# Gloire V2 — Classement automatique des factures fournisseurs

**Automatiser la lecture, l'identification, le renommage, le classement et le suivi des factures PDF avec Excel VBA et un OCR local.**

## Présentation

Gloire V2 est un projet d'automatisation documentaire développé sous **Excel VBA**. L'objectif est de réduire les opérations manuelles nécessaires au traitement des factures fournisseurs : ouverture des PDF, lecture des références, renommage des fichiers, création des dossiers fournisseurs et classement.

La solution fonctionne **localement sous Windows**, sans API payante ni service cloud. Un pont OCR Windows est embarqué dans le module VBA afin d'extraire le texte des factures PDF.

## Problématique métier

Dans un processus classique, un utilisateur doit ouvrir chaque facture, identifier le fournisseur et les références utiles, renommer le PDF puis le déplacer dans le bon dossier. Sur un volume important de documents, ces opérations deviennent répétitives et augmentent le risque d'erreur de classement.

Gloire V2 automatise cette chaîne tout en conservant un **contrôle humain lorsqu'une information est incertaine**.

## Fonctionnalités principales

- Analyse automatique des factures PDF.
- OCR Windows exécuté localement.
- Détection du **fournisseur**.
- Extraction du **numéro de contrat / référence client**.
- Extraction du **numéro de facture**.
- Détection de la **date d'échéance**.
- Dictionnaire de fournisseurs et gestion des alias.
- Utilisation de connecteurs textuels et d'expressions régulières.
- Mémoire des formats propres à chaque fournisseur.
- Priorité donnée aux formats déjà appris.
- Contrôle utilisateur en cas de doute.
- Correction manuelle avec mise à jour de la mémoire fournisseur.
- Renommage automatique des fichiers.
- Création automatique des dossiers fournisseurs.
- Détection des doublons.
- Conservation des PDF non validés dans le dossier de départ.
- Journalisation dans une feuille **STATISTIQUES**.
- Bouton permettant d'arrêter proprement le traitement.

## Chaîne de traitement

```text
Factures PDF
     |
     v
OCR Windows local
     |
     v
Extraction du texte
     |
     v
Identification fournisseur
     |
     v
Recherche contrat / facture / échéance
     |
     +------ information incertaine ------> contrôle utilisateur
     |                                          |
     |                                     validation / correction
     |                                          |
     v                                          v
Renommage standardisé <-------------------------+
     |
     v
Création du dossier fournisseur
     |
     v
Classement du PDF
     |
     v
Historique et statistiques Excel
```

## Exemple

Une facture identifiée avec :

- fournisseur : `ENGIE`
- contrat : `123456`
- facture : `F2026-854`

peut être renommée sous la forme :

```text
123456_F2026-854_ENGIE.pdf
```

puis classée automatiquement dans :

```text
Dossier_Arrivee\ENGIE\123456_F2026-854_ENGIE.pdf
```

## Organisation dans Excel

Le module peut créer et exploiter plusieurs feuilles de travail :

| Feuille | Rôle |
|---|---|
| `ARRETER_TRAITEMENT` | Interface permettant d'interrompre proprement la macro |
| `FOURNISSEURS` | Référentiel et fournisseurs personnalisés |
| `MEMOIRE_FORMATS` | Mémoire des formats détectés/corrigés par fournisseur |
| `STATISTIQUES` | Historique des traitements et anomalies |

## Principales macros

```text
Gloire_V2_Tester_OCR
Gloire_V2_Classement_Factures
Gloire_V2_Initialiser_Memoire
Gloire_V2_Arreter_Le_Traitement
```

La macro principale est :

```text
Gloire_V2_Classement_Factures
```

## Technologies utilisées

- **Microsoft Excel**
- **VBA**
- **Windows OCR**
- **Win32 API** pour certaines interactions avec les fenêtres
- **Microsoft Edge** pour le contrôle visuel des factures
- **VBScript RegExp** / expressions régulières
- **FileSystemObject** pour la gestion des fichiers et dossiers
- **ADODB.Stream** pour les flux et fichiers UTF-8
- **MSXML** pour le décodage Base64 du pont OCR embarqué

## Contraintes de conception

Le projet a été conçu pour fonctionner dans un environnement où l'installation de solutions externes peut être limitée.

Il n'utilise pas :

- d'API OCR payante ;
- de cloud pour analyser les factures ;
- Power Query pour le traitement ;
- PowerShell pour exécuter le processus principal.

Les factures et les textes OCR restent traités localement sur la machine de l'utilisateur.

## Gestion des erreurs et contrôle

Si une facture ne peut pas être classée avec suffisamment de certitude, le programme ne doit pas inventer l'information. Le document reste dans le dossier de départ et le traitement continue avec les factures suivantes.

Le système prend notamment en compte :

- OCR indisponible ;
- fournisseur non identifié ;
- contrat ou référence introuvable ;
- numéro de facture absent ;
- date d'échéance absente ;
- doublon de fichier ;
- erreur de déplacement ;
- demande d'arrêt utilisateur.

## Autre projet du portfolio — Contrôle fournisseurs & SIREN

Le dépôt contient également les archives retrouvées d'un projet antérieur de **contrôle du référentiel fournisseurs** :

- [`FOURNISSEUR_V1`](Fournisseurs_SIREN/FOURNISSEUR_V1_Reconstruit.bas) : recherche d'entreprise, récupération du SIREN, de la raison sociale, du statut et de la date de cessation via l'API Recherche Entreprises ;
- [`NOUVEAU_SIREN_V1`](Fournisseurs_SIREN/README.md) : évolution utilisant le BODACC et Recherche Entreprises pour analyser les cessations, fusions, TUP, scissions et autres événements juridiques afin d'identifier une éventuelle continuité vers un nouveau SIREN.

➡️ [Voir le projet Fournisseurs & SIREN](Fournisseurs_SIREN/README.md)

## Source

Le module VBA original contient également le **pont OCR compilé encodé en Base64**. Pour conserver le fichier original à l'identique, une archive complète du module est fournie dans le dossier `source/` sous forme Base64 découpée en deux parties, avec les instructions de reconstruction.

Un extrait VBA lisible directement sur GitHub est également disponible dans `src/` afin de présenter les points d'entrée et l'architecture du module.

## Compétences démontrées

Ce projet met en pratique plusieurs compétences techniques et métier :

- automatisation de processus ;
- VBA avancé ;
- OCR et traitement documentaire ;
- manipulation de fichiers PDF ;
- expressions régulières ;
- gestion d'erreurs ;
- contrôle et qualité des données ;
- conception d'un système semi-automatique avec validation humaine ;
- amélioration continue d'un processus administratif et financier.

Le projet **Fournisseurs & SIREN** complète ces compétences avec : appels d'API REST publiques, analyse JSON en VBA, exploitation du BODACC, contrôle de référentiels et analyse de continuité juridique.

## Statut du projet

**Version : Gloire V2 — MAJ 5**

Le code source est disponible pour démonstration et amélioration. Les gains de temps et le taux de reconnaissance OCR ne sont pas présentés comme des résultats mesurés tant qu'ils n'ont pas été évalués sur un corpus de factures de référence.

---

**Projet personnel — Automatisation / Finance & SI / RPA / OCR**
