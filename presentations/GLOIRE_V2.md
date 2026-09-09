# GLOIRE V2

## Automatisation du traitement et du classement des factures fournisseurs

### Pourquoi j’ai créé ce projet

Le traitement manuel de factures PDF demande plusieurs petites actions répétitives : ouvrir le document, retrouver le fournisseur, identifier la référence ou le contrat, relever le numéro de facture, vérifier certaines informations, renommer le fichier puis le déplacer dans le bon dossier.

J’ai développé **GLOIRE V2** pour regrouper ces étapes dans un seul processus automatisé sous Excel VBA, tout en conservant une validation humaine lorsqu’une information n’est pas assez fiable.

---

## Objectif

GLOIRE V2 doit pouvoir :

- lire une facture PDF avec un OCR local ;
- identifier le fournisseur ;
- rechercher le numéro de contrat ou une référence client ;
- rechercher le numéro de facture ;
- rechercher la date d’échéance ;
- proposer une validation lorsque nécessaire ;
- renommer le PDF selon une nomenclature définie ;
- créer ou utiliser le bon dossier fournisseur ;
- déplacer le fichier sans écraser un document existant ;
- enregistrer le résultat du traitement dans Excel.

---

## Principe de fonctionnement

```text
Factures PDF à traiter
        ↓
Lecture OCR locale
        ↓
Extraction du texte
        ↓
Identification du fournisseur
        ↓
Recherche des références importantes
        ↓
Contrôle de la fiabilité
        ↓
Si nécessaire : validation / correction humaine
        ↓
Renommage du fichier
        ↓
Classement dans le dossier fournisseur
        ↓
Enregistrement du résultat dans STATISTIQUES
```

---

## Pourquoi un OCR local

Le programme utilise un pont OCR Windows exécuté localement.

Le choix d’un traitement local permet de garder la lecture du document sur la machine utilisée pour le traitement, sans dépendre d’un service OCR cloud payant pour chaque facture.

Le chemin du programme OCR peut être enregistré dans une feuille dédiée afin que GLOIRE V2 sache où le retrouver.

---

## Bloc principal 1 — Les points d’entrée du projet

```vb
Public Sub Gloire_V2_Tester_OCR()
    ExecuterTestOCRGloireV2
End Sub

Public Sub Gloire_V2_Classement_Factures()
    ExecuterClassementFacturesGloireV2
End Sub

Public Sub Gloire_V2_Initialiser_Memoire()
    Dim ws As Worksheet

    Set ws = PreparerFeuilleMemoireFormats(ThisWorkbook)
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub
```

Ces procédures correspondent aux actions principales accessibles à l’utilisateur : tester l’OCR, lancer le classement et préparer la mémoire de formats fournisseurs.

---

## Bloc principal 2 — Configuration du pont OCR

```vb
Public Sub Gloire_V2_Configurer_Pont_OCR()
    Dim fso As Object
    Dim chemin As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    chemin = RecupererCheminPontOCR(fso, True)

    If chemin <> "" Then
        MsgBox "Chemin du pont OCR enregistré :" & _
               vbCrLf & vbCrLf & chemin, vbInformation
    End If
End Sub
```

Cette partie permet de retrouver le programme utilisé pour réaliser la lecture OCR avant de commencer le traitement des factures.

---

## Bloc principal 3 — Arrêt propre du traitement

```vb
Public Sub Gloire_V2_Arreter_Le_Traitement()
    ArretTraitementDemande = True

    Application.StatusBar = _
        "Arrêt demandé : la facture en cours restera dans le dossier de départ."

    MettreAJourEtatFeuilleArret _
        ThisWorkbook, _
        "ARRET DEMANDE", _
        RGB(192, 0, 0)

    Beep
End Sub
```

Je voulais éviter qu’un arrêt utilisateur laisse un document dans un état intermédiaire. La facture en cours reste donc dans le dossier de départ si le traitement est interrompu.

---

## Les informations recherchées dans une facture

Le programme cherche principalement :

- le fournisseur ;
- le contrat ou la référence client ;
- le numéro de facture ;
- la date d’échéance.

La détection s’appuie sur plusieurs méthodes : mots-clés, connecteurs textuels, expressions régulières et formats déjà validés pour certains fournisseurs.

---

## Mémoire des formats fournisseurs

Toutes les factures ne présentent pas leurs informations au même endroit.

Pour éviter de repartir entièrement de zéro à chaque document, GLOIRE V2 possède une feuille `MEMOIRE_FORMATS`.

Lorsqu’une correction manuelle est validée, le programme peut conserver un format utile pour ce fournisseur. Lors d’un prochain traitement, cette information peut être utilisée en priorité avant les règles plus générales.

L’objectif n’est pas de remplacer le contrôle humain, mais de rendre le système plus efficace sur des formats déjà rencontrés.

---

## Validation humaine

Une automatisation documentaire doit savoir gérer le doute.

Lorsque le fournisseur ou une information importante n’est pas suffisamment claire, GLOIRE V2 peut laisser l’utilisateur :

- **valider** la proposition ;
- **corriger** l’information ;
- **reporter** le traitement.

Un document qui n’est pas validé n’est pas déplacé automatiquement comme s’il avait été correctement traité.

---

## Renommage des fichiers

Une fois les informations contrôlées, le document peut être renommé selon une structure homogène.

Exemple de logique :

```text
CONTRAT_NUMEROFACTURE_FOURNISSEUR.pdf
```

Exemple :

```text
123456_F2026-854_ENGIE.pdf
```

Cette nomenclature facilite ensuite la recherche et le classement des documents.

---

## Classement et gestion des doublons

Après validation :

1. le dossier fournisseur est recherché ou créé ;
2. le nom de fichier est préparé ;
3. le programme contrôle qu’un fichier portant le même nom n’est pas déjà présent ;
4. le PDF est déplacé dans le dossier d’arrivée ;
5. le traitement est enregistré dans la feuille `STATISTIQUES`.

Cette étape évite d’écraser silencieusement un fichier existant.

---

## Suivi dans Excel

La feuille `STATISTIQUES` sert de trace du traitement.

Elle permet de conserver les informations utiles sur les documents passés dans le processus et de mieux identifier les cas qui ont nécessité une correction ou une intervention humaine.

---

## Ce que ce projet apporte

GLOIRE V2 transforme une succession d’actions manuelles en une chaîne de traitement structurée :

**lecture → identification → contrôle → validation → renommage → classement → suivi**.

Le projet montre surtout comment Excel VBA peut être utilisé au-delà de simples formules ou macros de mise en forme, pour piloter un véritable processus documentaire sur Windows.

---

## Compétences mobilisées

- Excel VBA ;
- automatisation de processus ;
- OCR ;
- traitement de fichiers PDF ;
- expressions régulières ;
- gestion de fichiers et dossiers ;
- `FileSystemObject` ;
- dictionnaires ;
- gestion des erreurs ;
- interface avec Windows ;
- contrôle humain dans un processus semi-automatisé ;
- amélioration continue ;
- gestion documentaire.

---

### Code associé

[`src/GLOIRE_V2_BLOCS_PRINCIPAUX.bas`](../src/GLOIRE_V2_BLOCS_PRINCIPAUX.bas)