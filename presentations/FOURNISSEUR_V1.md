# FOURNISSEUR_V1

## Contrôle automatique des informations fournisseurs dans Excel

### Pourquoi j’ai créé ce projet

Lorsqu’un tableau contient plusieurs fournisseurs, vérifier manuellement leur identité et leur statut peut vite devenir répétitif. Il faut rechercher l’entreprise, contrôler son SIREN, vérifier si elle est toujours active et parfois retrouver sa date de fermeture.

J’ai donc développé **FOURNISSEUR_V1** pour automatiser cette première étape directement depuis Excel.

L’idée est simple : l’utilisateur sélectionne les lignes à contrôler, la macro lit le texte de recherche présent dans le tableau, interroge l’API Recherche d’Entreprises puis renseigne automatiquement les informations utiles.

---

## Objectif

Le projet permet de récupérer automatiquement :

- le **SIREN** ;
- la **raison sociale** ;
- le **statut administratif** ;
- la **date de fermeture ou de cessation** lorsqu’elle est disponible.

Dans la version utilisée ici :

- colonne `H` : texte utilisé pour rechercher l’entreprise ;
- colonne `B` : SIREN ;
- colonne `C` : date de cessation ;
- colonne `D` : raison sociale ;
- colonne `E` : statut.

---

## Fonctionnement

```text
Nom ou information fournisseur dans Excel
                ↓
Lecture des lignes sélectionnées
                ↓
Construction de la requête HTTP
                ↓
API Recherche d’Entreprises
                ↓
Réponse JSON
                ↓
Extraction des informations utiles
                ↓
Mise à jour du tableau Excel
```

Le traitement est volontairement simple pour l’utilisateur : il sélectionne les lignes concernées puis lance la macro.

---

## Bloc principal 1 — Parcourir les fournisseurs sélectionnés

```vb
For Each zone In Selection.Areas
    For Each ligne In zone.Rows

        texteRecherche = Trim$(CStr(ws.Cells(ligne.Row, "H").Value))
        ws.Range("B" & ligne.Row & ":E" & ligne.Row).ClearContents

        If texteRecherche <> "" Then
            trouve = ChercherInfosDataGouv( _
                texteRecherche, _
                siren, _
                dateCessation, _
                raisonSociale, _
                statut)

            If trouve Then
                ws.Cells(ligne.Row, "B").Value = ValeurOuAucune(siren)
                ws.Cells(ligne.Row, "C").Value = ValeurOuAucune(dateCessation)
                ws.Cells(ligne.Row, "D").Value = ValeurOuAucune(raisonSociale)
                ws.Cells(ligne.Row, "E").Value = ValeurOuAucune(statut)
            End If
        End If

    Next ligne
Next zone
```

Ce bloc relie directement le tableau Excel à la fonction qui effectue la recherche.

---

## Bloc principal 2 — Interroger l’API

```vb
url = "https://recherche-entreprises.api.gouv.fr/search" & _
      "?page=1" & _
      "&per_page=1" & _
      "&minimal=true" & _
      "&sort_by_size=true" & _
      "&q=" & UrlEncode(recherche)

Set http = CreateObject("MSXML2.XMLHTTP.6.0")
http.Open "GET", url, False
http.setRequestHeader "Accept", "application/json"
http.setRequestHeader "User-Agent", "Excel VBA Recherche Entreprise"
http.Send
```

La macro construit l’URL de recherche puis effectue une requête HTTP depuis VBA.

---

## Bloc principal 3 — Transformer le statut reçu

```vb
etat = ExtraireValeurJson(reponse, "etat_administratif")

Select Case UCase$(etat)
    Case "A"
        statut = "Active"
    Case "C", "F"
        statut = "Fermée / cessée"
    Case Else
        statut = ""
End Select
```

L’objectif est de restituer une information immédiatement compréhensible dans le tableau Excel.

---

## Points importants du projet

### Gestion des informations manquantes

Lorsqu’une entreprise est trouvée mais qu’une information n’est pas disponible, la macro affiche `Aucune` plutôt que de laisser une valeur ambiguë.

### Gestion des erreurs HTTP

Le programme contrôle également le statut de la réponse HTTP. En cas de problème de communication, l’utilisateur peut identifier qu’il ne s’agit pas simplement d’une entreprise introuvable.

### Encodage des recherches

Les noms contenant des espaces ou certains caractères spéciaux sont encodés avant l’envoi de la requête.

---

## Ce que ce projet apporte

FOURNISSEUR_V1 permet surtout de **gagner du temps sur les contrôles répétitifs** et de centraliser le résultat directement dans Excel.

Il constitue également la première brique du projet global : une fois qu’un fournisseur est identifié comme cessé ou radié, le projet **NOUVEAU_SIREN_V1** peut prendre le relais pour analyser une éventuelle continuité juridique.

---

## Compétences mobilisées

- Excel VBA ;
- appels HTTP ;
- utilisation d’une API publique ;
- traitement JSON ;
- expressions régulières ;
- automatisation de contrôles ;
- gestion des erreurs ;
- structuration d’un traitement par lots.

---

### Code associé

[`src/FOURNISSEUR_V1.bas`](../src/FOURNISSEUR_V1.bas)