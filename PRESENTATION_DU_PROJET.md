# Automatisation Finance & SI — Excel VBA

Ce dépôt regroupe trois projets que j’ai développés autour d’un même objectif : **réduire les tâches manuelles, fiabiliser les informations fournisseurs et faciliter le traitement des factures**.

Je suis parti de besoins très concrets rencontrés dans des processus de gestion : retrouver rapidement les informations d’une entreprise, comprendre ce qu’il se passe lorsqu’un fournisseur cesse son activité, puis automatiser le traitement de factures PDF.

Les trois projets se complètent :

1. **FOURNISSEUR_V1** — contrôle automatique des informations d’un fournisseur.
2. **NOUVEAU_SIREN_V1** — recherche d’une éventuelle continuité juridique après une cessation d’activité.
3. **GLOIRE V2** — lecture OCR, identification, renommage et classement de factures fournisseurs.

---

## Les projets

### 1. FOURNISSEUR_V1 — Contrôle du référentiel fournisseurs

Le premier besoin était simple : éviter de rechercher manuellement chaque entreprise sur Internet.

À partir d’un nom de fournisseur présent dans Excel, la macro interroge l’API Recherche d’Entreprises et récupère automatiquement le **SIREN, la raison sociale, le statut administratif et la date de fermeture lorsqu’elle existe**.

➡️ [Voir la présentation détaillée de FOURNISSEUR_V1](presentations/FOURNISSEUR_V1.md)

---

### 2. NOUVEAU_SIREN_V1 — Recherche de continuité juridique

Le deuxième projet répond à une question plus complexe : **que devient un fournisseur lorsqu’il est radié, absorbé, fusionné ou concerné par une autre opération juridique ?**

Le programme analyse les informations disponibles dans le BODACC, identifie la nature de l’événement et cherche, lorsqu’il existe suffisamment d’éléments, la société qui pourrait avoir repris l’activité ou le patrimoine de l’entreprise initiale.

➡️ [Voir la présentation détaillée de NOUVEAU_SIREN_V1](presentations/NOUVEAU_SIREN_V1.md)

---

### 3. GLOIRE V2 — Classement automatique des factures

Le troisième projet porte sur le traitement documentaire.

GLOIRE V2 analyse des factures PDF avec un OCR local, identifie le fournisseur et plusieurs informations utiles, propose une validation lorsque nécessaire, renomme le document et le classe automatiquement dans le bon dossier.

Le programme conserve également une mémoire des formats déjà rencontrés afin de mieux reconnaître certaines factures fournisseurs au fil des corrections validées.

➡️ [Voir la présentation détaillée de GLOIRE V2](presentations/GLOIRE_V2.md)

---

## Logique d’ensemble

```text
Référentiel fournisseurs
        │
        ├── Vérifier l’identité et le statut de l’entreprise
        │        ↓
        │   FOURNISSEUR_V1
        │
        ├── Si l’entreprise est cessée ou radiée
        │        ↓
        │   NOUVEAU_SIREN_V1
        │
        └── Pour les factures reçues
                 ↓
             GLOIRE V2
                 ↓
        OCR → contrôle → renommage → classement
```

---

## Technologies utilisées

- Microsoft Excel
- VBA
- API Recherche d’Entreprises
- données BODACC
- requêtes HTTP
- traitement JSON
- expressions régulières
- `Scripting.Dictionary`
- `FileSystemObject`
- OCR Windows local
- Microsoft Edge pour certains contrôles visuels

---

## Ce que ces projets m’ont permis de travailler

- automatisation de processus métier ;
- amélioration continue ;
- développement VBA ;
- utilisation d’API publiques ;
- traitement et contrôle de données ;
- lecture de réponses JSON ;
- expressions régulières ;
- gestion des erreurs et des cas particuliers ;
- OCR et gestion documentaire ;
- conception d’un processus avec validation humaine lorsque l’automatisation ne suffit pas.

---

## Organisation du dépôt

```text
PRESENTATION_DU_PROJET.md

presentations/
├── FOURNISSEUR_V1.md
├── NOUVEAU_SIREN_V1.md
└── GLOIRE_V2.md

src/
├── FOURNISSEUR_V1.bas
├── NOUVEAU_SIREN_V1_BLOCS_PRINCIPAUX.bas
└── GLOIRE_V2_BLOCS_PRINCIPAUX.bas
```

Les fichiers VBA présents dans `src/` permettent de voir directement les parties les plus importantes de la logique de chaque projet.

---

**Positionnement : Automatisation de processus · Finance & SI · Excel VBA · OCR · Qualité des données · Gestion documentaire**