Option Explicit

' ==========================================================
' NOUVEAU_SIREN_V1 - aperçu de l'archive historique
' ==========================================================
' L'archive retrouvée contient un moteur beaucoup plus long
' intégrant BODACC, Recherche Entreprises, cache et moteur
' lexical pour les événements juridiques.
'
' Cet aperçu documente les constantes et le point d'entrée
' authentifiés dans l'archive. Le fichier historique complet
' est décrit dans README.md.
' ==========================================================

Private Const API_BODACC As String = "https://bodacc-datadila.opendatasoft.com/api/explore/v2.1/catalog/datasets/annonces-commerciales/records"
Private Const API_RECHERCHE As String = "https://recherche-entreprises.api.gouv.fr/search?q="
Private Const MAX_OFFSET_BODACC_ACTIF As Long = 100
Private Const MAX_OFFSET_BODACC_CESSE As Long = 300
Private Const LIGNE_DEBUT_DONNEES As Long = 11

Private Const COL_SIREN As String = "B"
Private Const COL_STATUT As String = "D"
Private Const COL_DATE_CESSATION As String = "E"
Private Const COL_CAUSE As String = "F"
Private Const COL_NOUVEAU_SIREN As String = "G"
Private Const COL_NOUVELLE_SOCIETE As String = "H"

Public Sub NOUVEAU_SIREN_V1()
    ' Point d'entrée de la version historique retrouvée.
    ' Le moteur complet :
    ' - analyse les lignes sélectionnées ;
    ' - valide le SIREN ;
    ' - interroge le BODACC ;
    ' - recherche les événements juridiques ;
    ' - tente d'identifier une entreprise bénéficiaire ;
    ' - recherche son SIREN dans l'API Recherche Entreprises ;
    ' - écrit la cause, le nouveau SIREN et la nouvelle société.
    '
    ' Voir README.md pour la provenance et les limites de cet aperçu.
End Sub
