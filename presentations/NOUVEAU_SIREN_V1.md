# NOUVEAU_SIREN_V1

## Recherche d’une continuité juridique après la cessation d’un fournisseur

### Pourquoi j’ai créé ce projet

Identifier qu’un fournisseur est fermé ou radié ne suffit pas toujours. Dans certains cas, l’activité peut avoir été reprise par une autre société à la suite d’une fusion, d’une transmission universelle de patrimoine, d’une scission, d’une cession ou d’une autre opération juridique.

Le vrai problème devient alors : **est-ce qu’une autre entreprise a repris tout ou partie de l’activité, et si oui, quel est son SIREN ?**

J’ai développé **NOUVEAU_SIREN_V1** pour automatiser une partie de cette recherche à partir de données publiques et limiter les vérifications manuelles ligne par ligne.

---

## Objectif

Le programme analyse un ancien SIREN et cherche à déterminer :

- la cause possible de cessation ;
- l’événement juridique associé ;
- l’existence éventuelle d’une société bénéficiaire ou absorbante ;
- le nouveau SIREN lorsqu’un lien suffisamment cohérent peut être établi ;
- le nom de la nouvelle société.

Le programme reste volontairement prudent : lorsqu’il n’y a pas assez d’éléments pour conclure, il laisse le résultat à vérifier plutôt que de produire un SIREN au hasard.

---

## Sources utilisées

Le projet s’appuie principalement sur :

- les annonces du **BODACC** ;
- l’**API Recherche d’Entreprises**.

Le BODACC permet d’analyser les événements juridiques publiés autour d’une entreprise, tandis que l’API Recherche d’Entreprises permet de retrouver ou confirmer les informations de la société identifiée.

---

## Événements recherchés

Le moteur prend notamment en compte des situations telles que :

- fusion ;
- fusion-absorption ;
- transmission universelle de patrimoine (TUP) ;
- scission ;
- apport partiel d’actif ;
- cession de fonds ;
- dissolution ;
- liquidation judiciaire ;
- redressement judiciaire ;
- plan de cession ;
- procédures collectives.

Le but n’est pas seulement de repérer un mot, mais de replacer l’information dans le contexte de l’annonce et de vérifier si elle peut mener vers une autre entreprise.

---

## Fonctionnement général

```text
Ancien SIREN
    ↓
Contrôle du format du SIREN
    ↓
Lecture du statut et de la date de cessation
    ↓
Recherche d’annonces BODACC
    ↓
Analyse de l’événement juridique
    ↓
Recherche d’une société bénéficiaire / absorbante
    ↓
Contrôle dans l’API Recherche d’Entreprises
    ↓
Nouveau SIREN trouvé
OU
Résultat à vérifier
```

---

## Bloc principal 1 — Lancer le traitement sur plusieurs lignes

```vb
Set lignes = LignesSelectionnees(Selection)
Set cacheBodacc = CreateObject("Scripting.Dictionary")
Set cacheAnnuaire = CreateObject("Scripting.Dictionary")

total = lignes.Count

For Each cle In lignes.Keys
    traitees = traitees + 1
    Application.StatusBar = "NOUVEAU SIREN V1 : " & traitees & "/" & total

    resultat = TraiterLigne( _
        ws, _
        CLng(cle), _
        cacheBodacc, _
        cacheAnnuaire)

    DoEvents
Next cle
```

Ce bloc permet de traiter plusieurs fournisseurs à la suite tout en évitant de répéter certaines recherches grâce aux dictionnaires utilisés comme cache.

---

## Bloc principal 2 — Vérifier l’ancien SIREN

```vb
ancienSiren = ExtraireSirenValide( _
    CStr(ws.Cells(ligne, COL_SIREN).Text))

If Len(ancienSiren) <> 9 Then
    EcrireResultat ws, ligne, _
        "SIREN absent ou invalide en colonne B", _
        "A VERIFIER", _
        ""

    TraiterLigne = CreerRetour("ERREUR", "", "", "")
    Exit Function
End If
```

Avant d’interroger une source externe, le programme vérifie que le SIREN est exploitable. Cela évite de lancer des recherches inutiles sur des valeurs incorrectes.

---

## Bloc principal 3 — Choisir le traitement selon le statut

```vb
If StatutCesseOuRadie(statut) Or dateOk Then

    If dateOk Then
        resultat = ChercherBodacc( _
            ancienSiren, _
            dateCessation, _
            cacheBodacc, _
            cacheAnnuaire)

        resultat = CorrigerResultatCauseCessation( _
            resultat, _
            ancienSiren, _
            dateCessation, _
            cacheBodacc)
    Else
        cause = ChercherCauseCessationBodaccSansDate( _
            ancienSiren, _
            cacheBodacc)
    End If

End If
```

La recherche est adaptée à la situation : si une date de cessation est disponible, elle sert de point de repère pour cibler les annonces pertinentes.

---

## Pourquoi utiliser un cache

Les appels externes peuvent être coûteux en temps lorsqu’un tableau contient beaucoup de lignes.

Le projet utilise des `Scripting.Dictionary` afin de conserver certaines réponses déjà récupérées pendant l’exécution. Si une information a déjà été demandée, elle peut être réutilisée sans refaire immédiatement la même requête.

---

## Contrôle humain

Ce projet touche à des informations juridiques qui peuvent être complexes. L’automatisation sert donc à **réduire le volume de recherche**, pas à supprimer tout contrôle.

Lorsqu’une correspondance n’est pas assez claire, le programme doit privilégier un résultat à vérifier plutôt qu’une conclusion trop rapide.

---

## Ce que ce projet apporte

NOUVEAU_SIREN_V1 transforme une recherche manuelle souvent longue en un processus structuré :

- le fournisseur est contrôlé ;
- les annonces sont analysées ;
- les événements sont classés ;
- une société cible peut être recherchée ;
- le résultat revient directement dans Excel.

C’est une extension logique de FOURNISSEUR_V1 : le premier projet indique qu’une société n’est plus active, le second cherche à comprendre ce qu’elle est devenue.

---

## Compétences mobilisées

- Excel VBA ;
- intégration d’API ;
- données BODACC ;
- traitement de texte ;
- logique de règles ;
- expressions régulières ;
- gestion de dictionnaires et de cache ;
- contrôle de cohérence ;
- automatisation de recherches juridiques et administratives ;
- gestion des cas incertains.

---

### Code associé

[`src/NOUVEAU_SIREN_V1_BLOCS_PRINCIPAUX.bas`](../src/NOUVEAU_SIREN_V1_BLOCS_PRINCIPAUX.bas)