# Gloire — Portfolio Data & Automatisation

**Transformer des documents en données structurées, puis préparer leur analyse.**

Ce portfolio rassemble deux projets complémentaires : un prototype OCR relié à Excel et un environnement de formation SQL Server construit autour d’un grand groupe fictif. Chaque projet présente son besoin métier, ses composants, les vérifications réalisées et son périmètre actuel.

## Les projets

| Projet | Besoin métier | Solution disponible | Technologies |
|---|---|---|---|
| [OCR de factures](OCR/README.md) | Limiter la ressaisie d’informations issues de factures PDF | Pont local entre un moteur OCR et des modules Excel/VBA | C#, .NET Framework, VBA, Python |
| [Analyse d’un grand groupe](SQL/README.md) | Disposer de données cohérentes pour pratiquer l’analyse commerciale, achats et recouvrement | Base de 700 000 lignes, scripts de contrôle et parcours de 60 exercices | SQL Server, T-SQL, Python ; parcours Excel et Power BI |

Ces projets sont indépendants : les résultats OCR ne sont pas injectés dans la base SQL de formation.

## 01 — OCR de factures vers Excel

Le projet vise à préparer des lignes de saisie à partir de la première page d’une facture PDF. Excel transmet un référentiel d’opérations et de rédacteurs ; le pont C# renvoie des informations structurées accompagnées d’indices de confiance. Le traitement s’effectue localement.

**Disponible :** code du pont C#, deux modules VBA, protocole d’échange, exemple anonymisé et contrôles automatiques de structure.

**État :** prototype d’intégration. Les deux modules ont passé les contrôles de structure et de cohérence du moteur embarqué. Le classeur métier n’est pas fourni et la précision OCR n’a pas été mesurée sur un corpus de référence.

[Lire l’étude de cas OCR](OCR/ETUDE_DE_CAS.md) · [Consulter le code C#](OCR/GLOIRE_LOG_LECTUREFACTURE/Program.cs)

## 02 — SQL Server : Groupe Galaxie

Un groupe français fictif de 50 sociétés et 500 établissements sert de cadre à l’analyse des ventes, achats, salariés et paiements. Le modèle relie 13 tables et permet d’étudier le chiffre d’affaires, la marge commerciale et les retards de règlement.

**Disponible :** 700 000 lignes synthétiques, 20 CSV, scripts de création et d’import, vues analytiques, 50 exercices progressifs et 10 demandes avancées avec corrigés.

**État :** base chargée et contrôlée sur SQL Server 2025. Les vérifications n’ont détecté aucune anomalie de montants ou de relations ; les 60 requêtes de correction ont été exécutées. Les tableaux de bord Excel et Power BI restent des livrables à réaliser dans le parcours.

[Lire l’étude de cas SQL](SQL/ETUDE_DE_CAS.md) · [Voir la feuille de mission](SQL/Feuille_de_mission.md) · [Consulter le modèle SQL](SQL/01_Create_Database.sql)

## Compétences mobilisées dans les projets

- **Structuration des données :** référentiels, clés primaires et étrangères, granularité des tables.
- **Qualité :** vérification des montants, des dates, des relations et du contenu embarqué.
- **Automatisation :** échanges entre Excel/VBA et un programme C#, génération de données avec Python.
- **SQL :** jointures, agrégations, CTE, fonctions de dates et fonctions de classement.
- **Lecture métier :** distinction ventes/achats, gestion des annulations, paiements partiels et encours.
- **Documentation :** procédures de démarrage, règles de calcul, limites et preuves de validation.

## Parcourir le portfolio

Pour une lecture rapide, commencer par les études de cas. Pour examiner l’implémentation, ouvrir les fichiers sources. Pour reproduire le travail, suivre les instructions propres à chaque projet et utiliser les archives complètes.

[Paquet OCR](OCR_Projet.zip) · [Paquet SQL — 700 000 lignes](SQL_Projet_700000.zip)

## Ce qui reste à démontrer

Les gains de temps OCR et son taux d’exactitude restent à mesurer. Le parcours SQL doit être complété par des analyses personnelles, un classeur Excel et un rapport Power BI. Aucun gain de productivité, résultat en production ou niveau de maîtrise n’est déduit automatiquement des tests techniques.

