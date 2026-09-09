# Automatisation du contrôle fournisseurs et de la continuité SIREN

Cette partie du portfolio regroupe deux générations de macros Excel VBA développées autour du contrôle du référentiel fournisseurs.

## 1. FOURNISSEUR_V1

`FOURNISSEUR_V1` constitue la première étape du projet. La macro interroge l'API publique **Recherche Entreprises** à partir d'un texte de recherche présent dans Excel.

### Données traitées

- Colonne H : texte / fournisseur à rechercher.
- Colonne B : SIREN retourné.
- Colonne C : date de cessation.
- Colonne D : raison sociale.
- Colonne E : statut de l'entreprise.

### Fonctionnement

1. L'utilisateur sélectionne les lignes à contrôler.
2. La macro récupère le texte de recherche en colonne H.
3. Une requête HTTP est envoyée à l'API Recherche Entreprises.
4. Le JSON retourné est analysé avec des expressions régulières VBA.
5. Le SIREN, la raison sociale, la date de cessation et le statut sont inscrits dans Excel.
6. Les lignes sans résultat restent vides afin de ne pas inventer d'information.

Le fichier `FOURNISSEUR_V1_Reconstruit.bas` a été **reconstitué à partir de l'archive de conversation retrouvée**. Le contenu fonctionnel de la macro et ses règles ont pu être récupérés, mais il ne s'agit pas d'une copie binaire garantie identique au fichier `.bas` d'origine.

## 2. NOUVEAU_SIREN_V1

`NOUVEAU_SIREN_V1` est l'évolution du projet. Il intervient lorsqu'un fournisseur est cessé, radié ou concerné par une opération juridique et cherche à déterminer s'il existe une continuité vers une autre entreprise.

### Sources utilisées

- **BODACC / DILA** pour les annonces juridiques et événements de la vie des entreprises.
- **API Recherche Entreprises** pour identifier et vérifier les entreprises et leurs SIREN.

### Informations exploitées dans Excel

| Colonne | Information |
|---|---|
| B | Ancien SIREN |
| D | Statut |
| E | Date de cessation |
| F | Cause de cessation / événement |
| G | Nouveau SIREN proposé |
| H | Nouvelle société |

### Traitements retrouvés dans l'archive

Le moteur historique comprend notamment :

- validation du format du SIREN ;
- gestion des lignes Excel sélectionnées ;
- caches de requêtes BODACC et Recherche Entreprises ;
- recherche d'annonces autour de la date de cessation ;
- analyse des sociétés actives et cessées ;
- détection et normalisation d'événements juridiques ;
- recherche d'une entreprise bénéficiaire ou absorbante ;
- recherche du SIREN correspondant ;
- distinction entre résultat trouvé, absence de successeur et erreur de consultation ;
- contrôle humain lorsque la situation juridique n'est pas suffisamment certaine.

Le moteur lexical retrouvé sait notamment reconnaître ou rapprocher des événements tels que :

- fusion et fusion-absorption ;
- transmission universelle du patrimoine (TUP) ;
- scission ;
- apport partiel d'actif ;
- cession de fonds ;
- dissolution ;
- liquidation judiciaire ;
- redressement judiciaire ;
- procédures de sauvegarde et plans de cession.

## Évolution du projet

```text
FOURNISSEUR_V1
      |
      |  contrôle de l'identité et du statut
      v
Entreprise active / cessée
      |
      |  si cessation ou événement juridique
      v
NOUVEAU_SIREN_V1
      |
      +--> analyse BODACC
      |
      +--> identification du bénéficiaire / absorbant
      |
      +--> Recherche Entreprises
      |
      v
Nouveau SIREN proposé + justification
```

## Fichiers publiés

- [`FOURNISSEUR_V1_Reconstruit.bas`](FOURNISSEUR_V1_Reconstruit.bas) : version récupérée et reconstruite de la première macro.
- [`NOUVEAU_SIREN_V1_Apercu.bas`](NOUVEAU_SIREN_V1_Apercu.bas) : constantes et point d'entrée authentifiés de la version historique avancée.

### Archive NOUVEAU_SIREN retrouvée

L'archive d'origine retrouvée dans la bibliothèque est un document `NOUVEAU_SIREN_MOTEUR_LEXICAL_30_POURCENT_RAPIDE(1).docx` d'environ 97 pages contenant le code du moteur VBA. Une extraction locale du code représente environ **2 860 lignes**.

Le connecteur GitHub utilisé ici ne permet pas de transférer directement ce fichier binaire DOCX depuis la bibliothèque. Pour cette raison, ce dépôt publie pour l'instant l'aperçu vérifié et documente précisément l'archive retrouvée, sans prétendre que l'aperçu correspond au moteur complet.

## Confidentialité

Le tableau Excel de fournisseurs retrouvé avec ces archives **n'est volontairement pas publié** : le dépôt GitHub est public et ce fichier peut contenir des données fournisseur réelles. Seul le code et la documentation technique sont destinés au portfolio.

## Compétences démontrées

- Excel VBA ;
- appels HTTP vers des API REST ;
- analyse JSON en VBA ;
- expressions régulières ;
- automatisation du contrôle d'un référentiel fournisseurs ;
- exploitation de données publiques d'entreprises ;
- analyse d'événements juridiques ;
- qualité et contrôle des données ;
- mise en place d'un contrôle humain en cas d'incertitude.
