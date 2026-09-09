Option Explicit

' ==========================================================
' NOUVEAU_SIREN_V1
' Blocs principaux de la logique de traitement
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
    Dim ws As Worksheet
    Dim lignes As Object
    Dim cacheBodacc As Object
    Dim cacheAnnuaire As Object
    Dim cle As Variant
    Dim total As Long
    Dim traitees As Long
    Dim trouves As Long
    Dim sansNouveau As Long
    Dim erreurs As Long
    Dim resultat As Variant
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldStatusBar As Variant

    If TypeName(Selection) <> "Range" Then
        MsgBox "Sélectionne les cellules à analyser puis relance NOUVEAU_SIREN_V1.", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set lignes = LignesSelectionnees(Selection)
    Set cacheBodacc = CreateObject("Scripting.Dictionary")
    Set cacheAnnuaire = CreateObject("Scripting.Dictionary")

    If lignes.Count = 0 Then
        MsgBox "Aucune cellule exploitable dans la sélection.", vbExclamation
        Exit Sub
    End If

    InstallerEntetes ws

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    On Error GoTo FinAvecErreur
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    total = lignes.Count

    For Each cle In lignes.Keys
        traitees = traitees + 1
        Application.StatusBar = "NOUVEAU SIREN V1 : " & traitees & "/" & total

        resultat = TraiterLigne(ws, CLng(cle), cacheBodacc, cacheAnnuaire)

        Select Case CStr(resultat(1))
            Case "NOUVEAU": trouves = trouves + 1
            Case "ERREUR": erreurs = erreurs + 1
            Case Else: sansNouveau = sansNouveau + 1
        End Select

        DoEvents
    Next cle

FinPropre:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.StatusBar = oldStatusBar

    MsgBox "NOUVEAU SIREN V1 terminé." & vbCrLf & _
           "Lignes traitées : " & traitees & vbCrLf & _
           "Nouveaux SIREN trouvés : " & trouves & vbCrLf & _
           "Sans nouveau SIREN : " & sansNouveau & vbCrLf & _
           "Erreurs : " & erreurs, vbInformation
    Exit Sub

FinAvecErreur:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.StatusBar = oldStatusBar
    MsgBox "NOUVEAU SIREN V1 s'est arrêté : " & Err.Description, vbCritical
End Sub

Private Function TraiterLigne(ByVal ws As Worksheet, ByVal ligne As Long, _
                              ByVal cacheBodacc As Object, ByVal cacheAnnuaire As Object) As Variant
    Dim ancienSiren As String
    Dim statut As String
    Dim dateCessation As Date
    Dim dateOk As Boolean
    Dim resultat As Variant
    Dim cause As String

    If LigneVideRecherche(ws, ligne) Then
        TraiterLigne = CreerRetour("VIDE", "", "", "")
        Exit Function
    End If

    ancienSiren = ExtraireSirenValide(CStr(ws.Cells(ligne, COL_SIREN).Text))
    statut = NettoyerTexte(CStr(ws.Cells(ligne, COL_STATUT).Text) & " " & _
                           CStr(ws.Cells(ligne, COL_DATE_CESSATION).Text))

    If Len(ancienSiren) <> 9 Then
        EcrireResultat ws, ligne, "SIREN absent ou invalide en colonne B", "A VERIFIER", ""
        TraiterLigne = CreerRetour("ERREUR", "", "", "")
        Exit Function
    End If

    dateOk = ExtraireDateCellule(ws.Cells(ligne, COL_DATE_CESSATION), dateCessation)

    If StatutCesseOuRadie(statut) Or dateOk Then

        If dateOk Then
            resultat = ChercherBodacc(ancienSiren, dateCessation, cacheBodacc, cacheAnnuaire)
            resultat = CorrigerResultatCauseCessation(resultat, ancienSiren, dateCessation, cacheBodacc)
        Else
            cause = ChercherCauseCessationBodaccSansDate(ancienSiren, cacheBodacc)

            If Len(cause) > 0 Then
                resultat = CreerRetour("CAUSE", cause, "", "")
            Else
                resultat = CreerRetour("SANS", "", "", "")
            End If
        End If

        EcrireResultat ws, ligne, CStr(resultat(2)), CStr(resultat(3)), CStr(resultat(4))
        TraiterLigne = resultat
        Exit Function
    End If

    resultat = ChercherActualiteBodacc(ancienSiren, cacheBodacc)
    EcrireResultat ws, ligne, CStr(resultat(2)), CStr(resultat(3)), CStr(resultat(4))
    TraiterLigne = resultat
End Function

' Les autres fonctions du projet gèrent notamment :
' - l'analyse des annonces BODACC ;
' - la détection des événements juridiques ;
' - la recherche d'une société bénéficiaire ou absorbante ;
' - la recherche et le contrôle du SIREN cible ;
' - la normalisation du texte ;
' - les requêtes HTTP et le traitement JSON ;
' - la mise en cache des résultats.