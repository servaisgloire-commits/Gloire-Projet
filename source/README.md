# Source complète — Gloire V2

Le module VBA original `Gloire_V2_Classement_Factures_Automatique_MAJ_5.bas` embarque également un petit pont OCR Windows compilé sous forme Base64. Pour conserver le fichier original sans modifier son contenu, il est archivé dans un ZIP puis encodé en Base64 en deux parties.

## Fichiers

- `Gloire_V2_archive_part1.b64`
- `Gloire_V2_archive_part2.b64`

Après reconstruction, l'archive contient :

```text
Gloire_V2_Classement_Factures_Automatique_MAJ_5.bas
```

## Reconstruction sous Windows

Depuis le dossier contenant les deux fichiers :

```bat
copy /b Gloire_V2_archive_part1.b64+Gloire_V2_archive_part2.b64 Gloire_V2.zip.b64
certutil -decode Gloire_V2.zip.b64 Gloire_V2_Classement_Factures_Automatique_MAJ_5.zip
```

Décompresser ensuite le fichier ZIP pour récupérer le module `.bas` original.

## Reconstruction sous Linux / macOS

```bash
cat Gloire_V2_archive_part1.b64 Gloire_V2_archive_part2.b64 | base64 -d > Gloire_V2_Classement_Factures_Automatique_MAJ_5.zip
unzip Gloire_V2_Classement_Factures_Automatique_MAJ_5.zip
```

## Pourquoi cette présentation ?

Le module contient du VBA lisible mais aussi un exécutable OCR embarqué encodé en Base64. L'archive permet de préserver exactement le module original, tandis que le dossier `src/` présente une vue lisible de l'architecture du code pour un lecteur GitHub.
