# PRÉSENTATION DU PROJET — GLOIRE V2

## Automatisation Finance & SI avec Excel VBA

Ce dépôt présente un ensemble d'outils développés sous **Excel VBA** pour automatiser des tâches répétitives liées au traitement des factures fournisseurs et au contrôle du référentiel fournisseurs.

Le projet est organisé autour de trois modules complémentaires :

1. **FOURNISSEUR_V1** — recherche et contrôle des informations d'une entreprise.
2. **NOUVEAU_SIREN_V1** — recherche d'une continuité juridique lorsqu'un fournisseur cesse son activité.
3. **Gloire V2** — lecture OCR, identification, renommage et classement automatique des factures PDF.

L'objectif général est de réduire les contrôles manuels, fiabiliser les données et conserver une intervention humaine lorsqu'une situation nécessite une validation.

---

## 1. FOURNISSEUR_V1

### Objectif

FOURNISSEUR_V1 automatise la recherche d'informations sur une entreprise à partir d'un nom ou d'une information présente dans Excel.

La macro interroge l'**API Recherche d'Entreprises** et récupère notamment :

- le SIREN ;
- la raison sociale ;
- le statut administratif ;
- la date de fermeture ou de cessation lorsqu'elle est disponible.

### Fonctionnement

```text
Donnée fournisseur dans Excel
        ↓
Requête API Recherche Entreprises
        ↓
Lecture de la réponse JSON
        ↓
Extraction du SIREN / statut / société / date
        ↓
Mise à jour automatique du tableau Excel
```

Ce module permet d'accélérer le contrôle d'un référentiel fournisseurs et de repérer les entreprises actives ou cessées sans effectuer chaque recherche manuellement.

---

## 2. NOUVEAU_SIREN_V1

### Objectif

NOUVEAU_SIREN_V1 intervient lorsqu'un fournisseur est cessé, radié ou concerné par une opération juridique.

L'outil analyse des données publiques afin de rechercher une éventuelle entreprise ayant repris, absorbé ou reçu tout ou partie de son activité ou de son patrimoine.

### Sources principales

- **BODACC** ;
- **API Recherche d'Entreprises**.

### Événements analysés

Le moteur contient des règles permettant notamment d'identifier :

- fusion ;
- fusion-absorption ;
- transmission universelle de patrimoine (TUP) ;
- scission ;
- apport partiel d'actif ;
- cession de fonds ;
- dissolution ;
- liquidation judiciaire ;
- redressement judiciaire ;
- procédures collectives.

### Fonctionnement

```text
Ancien SIREN
    ↓
Contrôle du statut
    ↓
Recherche BODACC
    ↓
Analyse des annonces juridiques
    ↓
Détection de la cause / opération
    ↓
Identification d'une entreprise bénéficiaire
    ↓
Recherche de son SIREN
    ↓
Résultat ou contrôle humain
```

Le système ne doit pas inventer de nouveau SIREN lorsqu'aucune continuité fiable ne peut être établie.

---

## 3. GLOIRE V2 — CLASSEMENT AUTOMATIQUE DES FACTURES

### Objectif

Gloire V2 automatise la lecture, l'identification, le renommage, le classement et le suivi de factures fournisseurs au format PDF.

La solution utilise un **OCR Windows exécuté localement** afin d'extraire les informations présentes sur les documents sans envoyer les factures vers un service OCR cloud.

### Informations recherchées

- fournisseur ;
- numéro de contrat ou référence client ;
- numéro de facture ;
- date d'échéance.

### Chaîne de traitement

```text
Facture PDF
    ↓
OCR local Windows
    ↓
Extraction du texte
    ↓
Identification du fournisseur
    ↓
Recherche contrat / facture / échéance
    ↓
Contrôle des valeurs détectées
    ↓
Validation humaine si nécessaire
    ↓
Renommage du PDF
    ↓
Classement dans le dossier fournisseur
    ↓
Journalisation du traitement
```

### Mémoire fournisseur

Gloire V2 mémorise des exemples de formats propres aux fournisseurs. Lorsqu'un format connu est retrouvé sur une nouvelle facture, cette mémoire est utilisée en priorité avant les règles génériques.

Il s'agit d'un **système de reconnaissance adaptatif basé sur des exemples et des règles**, et non d'un modèle de machine learning.

### Contrôle humain

Lorsque le fournisseur ou une référence importante n'est pas suffisamment fiable, le PDF peut être ouvert pour validation. L'utilisateur peut valider, corriger ou reporter le traitement. Un document non validé reste dans son dossier de départ.

### Renommage

```text
CONTRAT_NUMEROFACTURE_FOURNISSEUR.pdf
```

Exemple :

```text
123456_F2026-854_ENGIE.pdf
```

---

## Technologies utilisées

- Microsoft Excel ;
- VBA ;
- Windows OCR ;
- API Recherche d'Entreprises ;
- API BODACC ;
- HTTP / JSON ;
- VBScript.RegExp ;
- Scripting.Dictionary ;
- FileSystemObject ;
- Microsoft Edge ;
- API Windows.

---

## Compétences démontrées

- automatisation de processus ;
- amélioration continue ;
- développement VBA ;
- intégration d'API publiques ;
- traitement de données JSON ;
- OCR et gestion documentaire ;
- contrôle et qualité des données ;
- expressions régulières ;
- gestion des erreurs ;
- conception d'un processus semi-automatisé avec contrôle humain ;
- compréhension des problématiques Finance & SI.

---

## Organisation du dépôt

```text
PRESENTATION_DU_PROJET.md
src/
├── FOURNISSEUR_V1.bas
├── NOUVEAU_SIREN_V1_CORE.bas
└── GLOIRE_V2_CORE.bas
```

Le fichier `FOURNISSEUR_V1.bas` correspond à la macro reconstruite depuis les archives retrouvées.

Les fichiers `NOUVEAU_SIREN_V1_CORE.bas` et `GLOIRE_V2_CORE.bas` présentent les points d'entrée, paramètres et logiques principales vérifiés à partir des archives complètes du projet. Les archives originales ont servi de référence pour éviter de présenter du code inventé ou des fonctionnalités non présentes.

---

**Positionnement : Automatisation de processus · Finance & SI · VBA · OCR · RPA · Gestion documentaire · Qualité des données**
