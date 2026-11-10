Attribute VB_Name = "GloireCegideOCR"
Option Explicit

' Notifications non bloquantes et interface avec la saisie automatique.
Private messagesRemplissageAutomatique As String

'==========================================================
' GLOIRE - REMPLISSAGE CEGIDE AVEC PONT OCR INTEGRE
'==========================================================
' - Le dossier d'installation du pont est choisi au premier lancement.
' - S'il est deja present et valide, il n'est pas reinstalle.
' - Le dernier PDF de Telechargements est analyse sans etre deplace.
' - Le resultat est ecrit sur la premiere ligne numerotee dont
'   la colonne ETAPE est vide dans le tableau Remplissage Cegide.
' - Le tableau est agrandi sous cette ligne si plusieurs operations
'   sont detectees, sans ecraser les traitements precedents.
'==========================================================

Private Const NOM_PARAM_PONT As String = "_GLOIRE_CEGIDE_PONT_OCR_INTEGRE_V1"
Private Const REG_APPLICATION As String = "GloireCegide"
Private Const REG_SECTION As String = "OCR"
Private Const REG_CLE_PONT As String = "CheminPontIntegreV1"
Private Const NOM_PARAM_DOSSIER As String = "_GLOIRE_CEGIDE_DOSSIER_FACTURES"
Private Const REG_CLE_DOSSIER As String = "DossierFactures"
Private Const FEUILLE_RESULTAT As String = "Remplissage Cegide"
Private Const FEUILLE_ARTICLES As String = "TYPE d'article a refacturer"
Private Const FEUILLE_ORGANISATION As String = "Code organisation"
Private Const NOMBRE_CHIFFRES_MODULE As Long = 10
Private Const DELAI_OCR_SECONDES As Long = 240
Private Const NOM_FICHIER_PONT_INTEGRE As String = "WinOcrBridge_Spatial_MAJ5_Corrige.exe"
Private Const TAILLE_PONT_INTEGRE As Long = 14848

Private baseOperationsExactesCegide As Object
Private baseOperationsPrefixeCegide As Object

Public Sub Gloire_Cegide_Traiter_Derniere_Facture()
    Dim fso As Object
    Dim wsResultat As Worksheet
    Dim wsArticles As Worksheet
    Dim wsOrganisation As Worksheet
    Dim pontOCR As String
    Dim dossierTelechargements As String
    Dim dossierTextes As String
    Dim cheminPDF As String
    Dim cheminTXT As String
    Dim motsTSV As Collection
    Dim texteOCR As String
    Dim erreurOCR As String
    Dim probleme As String
    Dim articles As Collection
    Dim organisations As Collection
    Dim lignesFacture As Collection
    Dim redacteur As Object
    Dim moduleFacture As String
    Dim clientFacture As String
    Dim nomRedacteur As String
    Dim problemeDebours As String
    Dim nombreLignes As Long
    Dim messageFin As String
    Dim calculPrecedent As XlCalculation

    messagesRemplissageAutomatique = ""
    On Error GoTo GestionErreur
    calculPrecedent = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Verification du pont OCR integre..."

    pontOCR = ObtenirPontOCRCegide(False)
    If pontOCR = "" Then GoTo SortiePropre

    Application.StatusBar = "Preparation du remplissage Cegide..."

    Set wsResultat = TrouverFeuilleCegide(ThisWorkbook, FEUILLE_RESULTAT)
    Set wsArticles = TrouverFeuilleCegide(ThisWorkbook, FEUILLE_ARTICLES)
    Set wsOrganisation = TrouverFeuilleCegide(ThisWorkbook, FEUILLE_ORGANISATION)
    If wsResultat Is Nothing Then probleme = AjouterProblemeCegide(probleme, "Feuille Remplissage Cegide introuvable")
    If wsArticles Is Nothing Then probleme = AjouterProblemeCegide(probleme, "Feuille TYPE d'article a refacturer introuvable")
    If wsOrganisation Is Nothing Then probleme = AjouterProblemeCegide(probleme, "Feuille Code organisation introuvable")
    If probleme <> "" Then Err.Raise vbObjectError + 800, , probleme

    Set fso = CreateObject("Scripting.FileSystemObject")
    dossierTelechargements = ObtenirDossierFacturesCegide(fso)
    If dossierTelechargements = "" Then GoTo SortiePropre

    cheminPDF = DernierPDFCegide(fso, dossierTelechargements)
    If cheminPDF = "" Then Err.Raise vbObjectError + 802, , "Aucun fichier PDF trouve dans Telechargements."

    dossierTextes = fso.BuildPath(dossierTelechargements, "Textes OCR Cegide")
    If Not fso.FolderExists(dossierTextes) Then fso.CreateFolder dossierTextes
    cheminTXT = fso.BuildPath(dossierTextes, NettoyerNomFichierCegide(fso.GetBaseName(cheminPDF)) & ".txt")

    Application.StatusBar = "OCR du dernier PDF telecharge : " & fso.GetFileName(cheminPDF)
    texteOCR = ExecuterOCRCegide(pontOCR, cheminPDF, cheminTXT, fso, erreurOCR)
    If Trim$(texteOCR) = "" Then
        If erreurOCR = "" Then erreurOCR = "Le pont OCR n'a retourne aucun texte."
        Err.Raise vbObjectError + 803, , erreurOCR
    End If

    Set articles = ChargerArticlesCegide(wsArticles)
    Set organisations = ChargerOrganisationsCegide(wsOrganisation)
    If articles.Count = 0 Then Err.Raise vbObjectError + 804, , "Aucune correspondance exploitable dans les colonnes C et F des types d'article."
    If organisations.Count = 0 Then probleme = AjouterProblemeCegide(probleme, "Aucun redacteur exploitable dans les colonnes M, N et O de Code organisation.")

    Application.StatusBar = "Recherche complete : redacteur, module, operations et montants..."
    Set motsTSV = LirePositionsCegide(Left$(cheminTXT, Len(cheminTXT) - 4) & ".layout.tsv")
    moduleFacture = ModuleSpatialCegide(motsTSV, probleme)
    clientFacture = ClientSpatialCegide(motsTSV, probleme)
    Set redacteur = TrouverRedacteurCegide(texteOCR, organisations, motsTSV, moduleFacture)
    If redacteur Is Nothing Then
        Set redacteur = RedacteurServiceCegide(texteOCR, organisations)
        If Not redacteur Is Nothing Then probleme = AjouterProblemeCegide(probleme, "Redacteur par defaut du service " & CStr(redacteur("Service")) & " (premiere correspondance colonne P) : " & CStr(redacteur("Nom")))
    End If
    CompleterSaisieCegide clientFacture, moduleFacture, redacteur, organisations
    If redacteur Is Nothing Then
        probleme = AjouterProblemeCegide(probleme, "Aucun redacteur de la colonne O n'a ete retrouve.")
        nomRedacteur = ""
    Else
        nomRedacteur = CStr(redacteur("Nom"))
    End If

    If moduleFacture = "" Then
        probleme = AjouterProblemeCegide(probleme, "Le module de 10 chiffres, par exemple 0321010063, est introuvable.")
    End If

    If BloquerModuleCegide(wsResultat, moduleFacture, redacteur) Then
        Gloire_MessageAutomatique "Module non reference : annulation de la commande. Facture enregistree dans MODUP.", vbExclamation
        GoTo SortiePropre
    End If

    Set lignesFacture = AnalyserCombineCegide(motsTSV, texteOCR, articles, probleme)
    If lignesFacture.Count = 0 Then
        probleme = AjouterProblemeCegide(probleme, "Aucune operation equivalente a la colonne C n'a ete retrouvee dans ce PDF.")
    End If

    problemeDebours = VerifierCodesDeboursCegide(lignesFacture, articles)
    If problemeDebours <> "" Then probleme = AjouterProblemeCegide(probleme, problemeDebours)

    Application.StatusBar = "Remplissage des lignes Cegide..."
    If lignesFacture.Count > 0 Then
        nombreLignes = CompterLignesSortieCegide(lignesFacture)
        RemplirResultatCegide wsResultat, lignesFacture, redacteur, moduleFacture, articles, clientFacture
    Else
        nombreLignes = 1
        RemplirMetadonneesPartiellesCegide wsResultat, redacteur, moduleFacture, clientFacture
    End If

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = calculPrecedent
    messageFin = CStr(nombreLignes) & " ligne(s) remplie(s) dans Remplissage Cegide."
    messageFin = messageFin & vbCrLf & vbCrLf & "Facture : " & fso.GetFileName(cheminPDF)
    If redacteur Is Nothing Then
        messageFin = messageFin & vbCrLf & "Redacteur contentieux : non trouve"
    Else
        messageFin = messageFin & vbCrLf & "Redacteur contentieux : " & CStr(redacteur("Nom"))
    End If
    If moduleFacture = "" Then
        messageFin = messageFin & vbCrLf & "Module : non trouve"
    Else
        messageFin = messageFin & vbCrLf & "Module : " & moduleFacture
    End If
    If probleme = "" Then
        Gloire_MessageAutomatique messageFin, vbInformation
    Else
        messageFin = messageFin & vbCrLf & vbCrLf & "Elements a completer :" & vbCrLf & "- " & probleme
        Gloire_MessageAutomatique messageFin, vbExclamation
    End If
    Exit Sub

GestionErreur:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = calculPrecedent
    messageFin = "Remplissage Cegide interrompu." & vbCrLf & vbCrLf & Err.Description
    If cheminPDF <> "" Then messageFin = messageFin & vbCrLf & vbCrLf & "PDF analyse : " & cheminPDF
    Gloire_MessageAutomatique messageFin, vbExclamation
    Exit Sub

SortiePropre:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = calculPrecedent
End Sub

Public Sub Gloire_Cegide_Choisir_Pont_OCR()
    Dim chemin As String
    chemin = ObtenirPontOCRCegide(True)
    If chemin <> "" Then Gloire_MessageAutomatique "Pont OCR integre reinstalle :" & vbCrLf & chemin, vbInformation
End Sub

Public Sub Gloire_Cegide_Tester_Pont_OCR()
    Dim chemin As String
    chemin = ObtenirPontOCRCegide(False)
    If chemin <> "" Then Gloire_MessageAutomatique "Pont OCR disponible :" & vbCrLf & chemin, vbInformation
End Sub

Public Sub Gloire_Cegide_Choisir_Dossier_Factures()
    Dim fso As Object
    Dim chemin As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    chemin = ChoisirEtEnregistrerDossierCegide(fso)
    If chemin <> "" Then Gloire_MessageAutomatique "Dossier des factures memorise :" & vbCrLf & chemin, vbInformation
End Sub

Private Function ObtenirPontOCRCegide(ByVal forcerInstallation As Boolean) As String
    Dim chemin As String
    On Error GoTo Echec

    If Not forcerInstallation Then
        chemin = LirePontMemoriseCegide()
        If chemin <> "" Then
            If PontIntegreValideCegide(chemin) Then
                ObtenirPontOCRCegide = chemin
                Exit Function
            End If
            EffacerPontMemoriseCegide
        End If
    End If

    chemin = InstallerPontIntegreCegide(forcerInstallation)
    If chemin = "" Then Exit Function

    EnregistrerPontCegide chemin
    ObtenirPontOCRCegide = chemin
    Exit Function
Echec:
    Gloire_MessageAutomatique "Installation automatique du pont OCR impossible." & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Function

Private Function InstallerPontIntegreCegide(ByVal forcerInstallation As Boolean) As String
    Dim fso As Object
    Dim dossierOCR As String
    Dim chemin As String
    Dim selecteur As FileDialog
    Dim documentXML As Object
    Dim noeudBase64 As Object
    Dim flux As Object
    Dim donnees As Variant
    Dim numeroErreur As Long
    Dim descriptionErreur As String

    On Error GoTo Echec
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set selecteur = Application.FileDialog(msoFileDialogFolderPicker)
    With selecteur
        .Title = "Choisissez le dossier d'installation du pont OCR Cegide"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        dossierOCR = CStr(.SelectedItems(1))
    End With
    If Not fso.FolderExists(dossierOCR) Then Err.Raise vbObjectError + 820, , "Le dossier d'installation selectionne est introuvable."
    chemin = fso.BuildPath(dossierOCR, NOM_FICHIER_PONT_INTEGRE)

    If Not forcerInstallation Then
        If fso.FileExists(chemin) Then
            If fso.GetFile(chemin).Size = TAILLE_PONT_INTEGRE Then
                If PontDisponibleCegide(chemin) Then
                    InstallerPontIntegreCegide = chemin
                    Exit Function
                End If
            End If
        End If
    End If

    If fso.FileExists(chemin) Then fso.DeleteFile chemin, True
    Set documentXML = CreateObject("MSXML2.DOMDocument.6.0")
    Set noeudBase64 = documentXML.createElement("pont")
    noeudBase64.DataType = "bin.base64"
    noeudBase64.Text = PontIntegreBase64Cegide()
    donnees = noeudBase64.nodeTypedValue

    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 1
    flux.Open
    flux.Write donnees
    flux.SaveToFile chemin, 2
    flux.Close

    If Not fso.FileExists(chemin) Then Err.Raise vbObjectError + 821, , "Le fichier du pont n'a pas ete cree."
    If fso.GetFile(chemin).Size <> TAILLE_PONT_INTEGRE Then Err.Raise vbObjectError + 822, , "Le pont extrait est incomplet."
    If Not PontDisponibleCegide(chemin) Then Err.Raise vbObjectError + 823, , "Le pont integre est bloque ou l'OCR Windows est indisponible."

    InstallerPontIntegreCegide = chemin
    Exit Function
Echec:
    numeroErreur = Err.Number
    descriptionErreur = Err.Description
    On Error Resume Next
    If Not flux Is Nothing Then flux.Close
    On Error GoTo 0
    If numeroErreur = 0 Then numeroErreur = vbObjectError + 824
    Err.Raise numeroErreur, , descriptionErreur
End Function

Private Function PontIntegreValideCegide(ByVal chemin As String) As Boolean
    Dim fso As Object

    On Error GoTo Fin
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(chemin) Then Exit Function
    If fso.GetFile(chemin).Size <> TAILLE_PONT_INTEGRE Then Exit Function
    PontIntegreValideCegide = PontDisponibleCegide(chemin)
Fin:
End Function

Private Function PontIntegreBase64Cegide() As String
    Dim contenu As String

    contenu = contenu & "TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAA4fug4AtAnNIbgB"
    contenu = contenu & "TM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAABQRQAATAEDAOYXkGoAAAAAAAAAAOAA"
    contenu = contenu & "AgELAQsAADAAAAAIAAAAAAAATk8AAAAgAAAAYAAAAABAAAAgAAAAAgAABAAAAAAAAAAEAAAAAAAAAACgAAAAAgAAAAAAAAMAQIUA"
    contenu = contenu & "ABAAABAAAAAAEAAAEAAAAAAAABAAAAAAAAAAAAAAAPhOAABTAAAAAGAAAEAFAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAwAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAACAAAAAAAAAAAAAAACCAAAEgAAAAAAAAAAAAA"
    contenu = contenu & "AC50ZXh0AAAAVC8AAAAgAAAAMAAAAAIAAAAAAAAAAAAAAAAAACAAAGAucnNyYwAAAEAFAAAAYAAAAAYAAAAyAAAAAAAAAAAAAAAA"
    contenu = contenu & "AABAAABALnJlbG9jAAAMAAAAAIAAAAACAAAAOAAAAAAAAAAAAAAAAAAAQAAAQgAAAAAAAAAAAAAAAAAAAAAwTwAAAAAAAEgAAAAC"
    contenu = contenu & "AAUAoDAAAFgeAAABAAAABgAABgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGICewMAAAQC"
    contenu = contenu & "ewUAAAQjAAAAAAAAAEBbWCo6AnsDAAAEAnsFAAAEWCoeAigDAAAKKhMwBABcAAAAAQAAEQJ7BgAABG8EAAAKbAoCAnsHAAAEBloD"
    contenu = contenu & "bwEAAAZYBiMAAAAAAADwP1hbfQcAAAQCAnsIAAAEBloDewUAAARYBiMAAAAAAADwP1hbfQgAAAQCewYAAAQDbwUAAAoqSgJzBgAA"
    contenu = contenu & "Cn0GAAAEAigDAAAKKgAbMAQAugAAAAIAABECjmkXMz0CFppyAQAAcBsoCAAACiwtKAkAAAoKBi0WKAoAAApyEQAAcG8LAAAKGQ3d"
    contenu = contenu & "gwAAAHIxAABwKAwAAAoWDd51Ao5pGC8TKAoAAApyTQAAcG8LAAAKGA3eXBYLAo5pGTIRAhiaHSgNAAAKEgEoDgAACiYCFpooDwAA"
    contenu = contenu & "CgIXmigPAAAKBygHAAAGDd4qDCgKAAAKCG8QAAAKbxEAAApyxQAAcAhvEgAACigTAAAKbwsAAAoXDd4ACSoAAAEQAAAAAAAAjo4A"
    contenu = contenu & "KhYAAAEbMAUAxwEAAAMAABECKBQAAAotDHLLAABwAnMVAAAKeigJAAAKCgYtC3LrAABwcxYAAAp6AigXAAAKKAEAACsLBygYAAAK"
    contenu = contenu & "KAIAACsMCG8ZAAAKDQQWMQgJbgRqMQIEDXMaAAAKEwRzGgAAChMFEQVybgEAcG8bAAAKJhYTBjj+AAAACBEGbxwAAAoTB3MdAAAK"
    contenu = contenu & "EwgRBygeAAAKKAoAAAYTCREHEQgRCW8fAAAKKAkAAAYRCBZqbyAAAAoRCCghAAAKKAMAACsTChEKH1cWbyIAAAooBAAAKxMLBhEL"
    contenu = contenu & "byMAAAooBQAAKxMMEQYWNggRBG8kAAAKJhEEcvMBAHARBhdYEw8SDygNAAAKKCUAAApyCwIAcCgTAAAKbxsAAAomEQQRDBELbyYA"
    contenu = contenu & "AAoRC28nAAAKKAsAAAYRBREMEQYXWBELbyYAAAoRC28nAAAKKA8AAAbeDBELLAcRC28oAAAK3N4MEQgsBxEIbygAAArc3gwRBywH"
    contenu = contenu & "EQdvKAAACtwRBhdYEwYRBglE+v7//wMoKQAAChMNEQ0oKgAACi0IEQ0oKwAACiYWcywAAAoTDgMRBG8tAAAKEQ4oLgAACgNyGQIA"
    contenu = contenu & "cCgvAAAKEQVvLQAAChEOKC4AAAoWKgABKAAAAgDNAHlGAQwAAAAAAgCHAM1UAQwAAAAAAgCAAOJiAQwAAAAAEzACAGEAAAAEAAAR"
    contenu = contenu & "KDAAAAoKKyYGbzEAAAogwNQBAGoxEQJvMgAACnIxAgBwczMAAAp6HwooNAAACgJvNQAACizSAm81AAAKFzMHAm82AAAKKnJnAgBw"
    contenu = contenu & "Am81AAAKjC4AAAEoNwAACnMWAAAKegAAABMwAgBhAAAABAAAESgwAAAKCismBm8xAAAKIMDUAQBqMRECbzIAAApyqQIAcHMzAAAK"
    contenu = contenu & "eh8KKDQAAAoCbzUAAAos0gJvNQAAChczBwJvOAAACipy4QIAcAJvNQAACowuAAABKDcAAApzFgAACnoAAAATMAoA4QEAAAUAABEC"
    contenu = contenu & "bzkAAAqMMgAAAQojAAAAAAAA8D9+DQAABC0lHxDQMwAAASg6AAAK0AQAAAIoOgAACig7AAAKKDwAAAqADQAABH4NAAAEez0AAAp+"
    contenu = contenu & "DQAABH4OAAAELTQWch0DAHDQBAAAAig6AAAKF404AAABEwoRChYWFCg+AAAKohEKKD8AAAooQAAACoAOAAAEfg4AAAR7QQAACn4O"
    contenu = contenu & "AAAEBm9CAAAKb0MAAAooRAAACgsjAAAAAAAA8D9+DwAABC0lHxDQMwAAASg6AAAK0AQAAAIoOgAACig7AAAKKDwAAAqADwAABH4P"
    contenu = contenu & "AAAEez0AAAp+DwAABH4QAAAELTQWcikDAHDQBAAAAig6AAAKF404AAABEwsRCxYWFCg+AAAKohELKD8AAAooQAAACoAQAAAEfhAA"
    contenu = contenu & "AAR7QQAACn4QAAAEBm9CAAAKb0MAAAooRAAACgwHCChEAAAKDQMfIDUFA3ZsKw0DdmwjAAAAAAAAQEBZEwQjAAAAAAAA8D8jAAAA"
    contenu = contenu & "AADAokARBChFAAAKKEQAAAoTBREFCVsTBiMAAAAAAADwPwcRBlooRgAACihEAAAKbRMHIwAAAAAAAPA/CBEGWihGAAAKKEQAAApt"
    contenu = contenu & "EwhzRwAAChMJEQkRB29IAAAKEQkRCG9JAAAKEQkqAAAAGzADADYBAAAGAAARAygMAAAGCgZvBAAACi0BKgYoDQAABgsGKBEAAAYM"
    contenu = contenu & "IwAAAAAAAPA/BGwjAAAAAACAdkBbKEQAAAoNCCPNzMzMzMzsP1oJIwAAAAAAAOA/WyhEAAAKEwQFbBEEWyhKAAAKaRMFEQUHb0sA"
    contenu = contenu & "AAovCAdvSwAAChMFEQUgGAEAADEHIBgBAAATBRUTBgdvTAAAChMJOIQAAAASCShNAAAKEwcRB3sGAAAEFP4GFAAABnNOAAAKb08A"
    contenu = contenu & "AAoRB3sHAAAEIwAAAAAAAPA/BWwoRAAAClsRBRdZbFooRgAACmkTCBEIEQYwFREGF1gTCCsNAm8kAAAKJhEGF1gTBhEGF1gRCDLr"
    contenu = contenu & "AhEHewYAAAQEKA4AAAZvGwAACiYRBhdYEwYSCShQAAAKOnD////eDhIJ/hYIAAAbbygAAArcKgAAARAAAAIAkACXJwEOAAAAABsw"
    contenu = contenu & "DgANBAAABwAAEXMGAAAKCgJvUQAACm9SAAAKEwY4xwMAABEGb1MAAAoLB29UAAAKb1UAAAoTBziTAwAAEQdvVgAACgwIb1cAAAoo"
    contenu = contenu & "EgAABg0Jb1gAAAo5dAMAAAhvWQAACoxDAAABEwRzAwAABhMFEQUJfQEAAAQRBX4RAAAELSQW0DMAAAEoOgAACtAEAAACKDoAAAoo"
    contenu = contenu & "OwAACig8AAAKgBEAAAR+EQAABHs9AAAKfhEAAAR+EgAABC00FnI3AwBw0AQAAAIoOgAACheNOAAAARMIEQgWFhQoPgAACqIRCCg/"
    contenu = contenu & "AAAKKEAAAAqAEgAABH4SAAAEe0EAAAp+EgAABBEEb0IAAApvQwAACn0CAAAEEQV+EwAABC0kFtAzAAABKDoAAArQBAAAAig6AAAK"
    contenu = contenu & "KDsAAAooPAAACoATAAAEfhMAAAR7PQAACn4TAAAEfhQAAAQtNBZyOwMAcNAEAAACKDoAAAoXjTgAAAETCREJFhYUKD4AAAqiEQko"
    contenu = contenu & "PwAACihAAAAKgBQAAAR+FAAABHtBAAAKfhQAAAQRBG9CAAAKb0MAAAp9AwAABBEFfhUAAAQtJBbQMwAAASg6AAAK0AQAAAIoOgAA"
    contenu = contenu & "Cig7AAAKKDwAAAqAFQAABH4VAAAEez0AAAp+FQAABH4WAAAELUwWcj8DAHAU0AQAAAIoOgAAChmNOAAAARMKEQoWHyEUKD4AAAqi"
    contenu = contenu & "EQoXGRQoPgAACqIRChgWFCg+AAAKohEKKFoAAAooWwAACoAWAAAEfhYAAAR7XAAACn4WAAAE0DsAAAEoOgAACiMAAAAAAADwP34X"
    contenu = contenu & "AAAELTQWch0DAHDQBAAAAig6AAAKF404AAABEwsRCxYWFCg+AAAKohELKD8AAAooQAAACoAXAAAEfhcAAAR7QQAACn4XAAAEEQRv"
    contenu = contenu & "QgAACm9dAAAKb0MAAAp9BAAABBEFfhgAAAQtJBbQMwAAASg6AAAK0AQAAAIoOgAACig7AAAKKDwAAAqAGAAABH4YAAAEez0AAAp+"
    contenu = contenu & "GAAABH4ZAAAELUwWcj8DAHAU0AQAAAIoOgAAChmNOAAAARMMEQwWHyEUKD4AAAqiEQwXGRQoPgAACqIRDBgWFCg+AAAKohEMKFoA"
    contenu = contenu & "AAooWwAACoAZAAAEfhkAAAR7XAAACn4ZAAAE0DsAAAEoOgAACiMAAAAAAADwP34aAAAELTQWcikDAHDQBAAAAig6AAAKF404AAAB"
    contenu = contenu & "Ew0RDRYWFCg+AAAKohENKD8AAAooQAAACoAaAAAEfhoAAAR7QQAACn4aAAAEEQRvQgAACm9dAAAKb0MAAAp9BQAABAYRBW8FAAAK"
    contenu = contenu & "EQdvXgAACjph/P//3gwRBywHEQdvKAAACtwRBm9eAAAKOi38///eDBEGLAcRBm8oAAAK3AYU/gYTAAAGc04AAApvTwAACgYqAAAA"
    contenu = contenu & "QTQAAAIAAAAtAAAApgMAANMDAAAMAAAAAAAAAAIAAAATAAAA2gMAAO0DAAAMAAAAAAAAABswAwDeAAAACAAAEXNfAAAKCgJvYAAA"
    contenu = contenu & "ChMHOJsAAAASByhhAAAKCxQMI////////+9/DQZvTAAAChMIK00SCChNAAAKEwQHewUAAAQRBHsIAAAEKEQAAAojj8L1KFyP4j9a"
    contenu = contenu & "EwUHbwEAAAYRBHsHAAAEWShiAAAKEwYRBhEFNQsRBgk0BhEEDBEGDRIIKFAAAAotqt4OEgj+FggAABtvKAAACtwILQ1zBQAABgwG"
    contenu = contenu & "CG9jAAAKCAdvBAAABhIHKGQAAAo6Wf///94OEgf+FhAAABtvKAAACtwGFP4GFQAABnNlAAAKb2YAAAoGKgAAARwAAAIALwBaiQAO"
    contenu = contenu & "AAAAAAIADgCuvAAOAAAAABswBADQAAAACQAAESCIAQAAc2cAAAoKAm9gAAAKEwU4igAAABIFKGEAAAoLB3sCAAAEIwAAAAAAAPA/"
    contenu = contenu & "A2woRAAAClsjAAAAAABwdkBaKEYAAAppDAgWLwIWDAggaAEAADIGIGcBAAAMBm9oAAAKLAoGb2gAAAoXWCsBCA0ICShpAAAKEwQR"
    contenu = contenu & "BAZvaAAACjESBh8gEQQGb2gAAApZb2oAAAomBgd7AQAABG9rAAAKJhIFKGQAAAo6av///94OEgX+FhAAABtvKAAACtwGby0AAAoW"
    contenu = contenu & "jUUAAAFvbAAACioBEAAAAgATAJ2wAA4AAAAAGzAEAAwBAAAKAAARAygMAAAGCgZvYAAACgw43AAAABICKGEAAAoLAg8CKA0AAAoo"
    contenu = contenu & "JQAACm9rAAAKHwlvbQAACiYCDwMoDQAACihuAAAKb2sAAAofCW9tAAAKJgIPBCgNAAAKKG4AAApvawAACh8Jb20AAAomAgd7AgAA"
    contenu = contenu & "BAUoEAAABm9rAAAKHwlvbQAACiYCB3sDAAAEDgQoEAAABm9rAAAKHwlvbQAACiYCB3sEAAAEBSgQAAAGb2sAAAofCW9tAAAKJgIH"
    contenu = contenu & "ewUAAAQOBCgQAAAGb2sAAAofCW9tAAAKJgIHewEAAARyRwMAcHJLAwBwb28AAApvGwAACiYSAihkAAAKOhj////eDhIC/hYQAAAb"
    contenu = contenu & "bygAAArcKgEQAAACAA4A7/0ADgAAAAATMAMALwAAAAEAABEjAAAAAAAAWUACWiMAAAAAAADwPwNsKEQAAApbChIAck8DAHAoDQAA"
    contenu = contenu & "CihwAAAKKgAbMAMAUQAAAAsAABFzcQAACgoCb2AAAAoMKxQSAihhAAAKCwYHewUAAARvcgAAChICKGQAAAot494OEgL+FhAAABtv"
    contenu = contenu & "KAAACtwGb3MAAAoGBm90AAAKGFtvdQAACioAAAABEAAAAgANACEuAA4AAAAAwgIodgAACiwGfncAAAoqAh8NHyBveAAACh8KHyBv"
    contenu = contenu & "eAAACh8JHyBveAAACm95AAAKKgAAABMwAgAsAAAADAAAEQJvAQAABgsSAQNvAQAABih6AAAKCgYsAgYqAnwCAAAEA3sCAAAEKHoA"
    contenu = contenu & "AAoqSgJ8AgAABAN7AgAABCh6AAAKKkoCfAcAAAQDewcAAAQoegAACioAAEJTSkIBAAEAAAAAAAwAAAB2NC4wLjMwMzE5AAAAAAUA"
    contenu = contenu & "bAAAAFwJAAAjfgAAyAkAACAMAAAjU3RyaW5ncwAAAADoFQAAWAMAACNVUwBAGQAAEAAAACNHVUlEAAAAUBkAAAgFAAAjQmxvYgAA"
    contenu = contenu & "AAAAAAACAAABVx2iCQkOAAAA+iUzABYAAAEAAABGAAAABgAAABoAAAAVAAAAIAAAAHsAAAAEAAAABQAAAAwAAAABAAAAAgAAAAIA"
    contenu = contenu & "AAASAAAAAQAAAAoAAAACAAAAAQAAAAUAAAAAAAoAAQAAAAAABgByAGsABgDbAMAACgB8AWkBCgCbAWkBDgDRAcABDgDmAcABBgAN"
    contenu = contenu & "AgECEgA7AikCBgCxA5EDBgDRA5EDBgAbBGsABgAuBGsABgA1BGsAEgBNBCkCBgB5BGsABgCLBIEEBgC/BKoEBgDgBGsABgDmBKoE"
    contenu = contenu & "BgDzBGsABgAMBYEEBgAdBWsABgAnBWsABgBUBYEEBgBgBYEEBgB2BWsAFgCgBZAFDgDBBcABFgDNBZAFFgAlBg0GFgBWBg0GGgCt"
    contenu = contenu & "BpQGGgDHBpQGGgDWBpQGGgDoBpQGBgAeB2sABgBNB2sABgCAB4EEBgCKB4EEBgCoBwECBgC1BwECHgDuB9sHCgAZCGkBBgArCGsA"
    contenu = contenu & "BgBNCDwICgBaCGkBIgCuCJEDBgC5CGsAIgDACJEDJgATCWkBBgAhCWsABgAoCWsAKgB8CV0JIgCDCZEDKgCSCV0JKgC6CV0JKgDN"
    contenu = contenu & "CV0JBgDlCcAABgAECmsACwBKCgAABgBvCmsABgDeCmsABgAcC8AAEgAsCykCBgA+C8AAEgBMCykCJgByC2kBBgCoC5ULBgC/C2sA"
    contenu = contenu & "BgAFDJEDAAAAAAEAAAAAAAEAAQAAARAAMAA4AAUAAQABAAABEABQADgABQAGAAQAgAEQAFoAOAAFAAkABgCDARAAfAgAAAUADQAW"
    contenu = contenu & "AIMBEACKCgAABQARABYABgB5AAoABgB+AA0ABgCDAA0ABgCHAA0ABgCNAA0AJgDiABwABgDoAA0ABgD4AA0AUYAKASoAUYAYASoA"
    contenu = contenu & "UYAqASoAUYBHAQ0AFgDJCCkCFgDUCDoCFgDfCCkCFgDqCDoCFgCnCikCFgCyCjoCFgC9CikCFgDICjoCFgDTCikCFgDlClcDFgDw"
    contenu = contenu & "CjoCFgD7CikCFgAGC1cDFgARCzoCUCAAAAAAhgiUABAAAQBpIAAAAACGCKAAEAABAHggAAAAAIYYqwAUAAEAgCAAAAAAhgAGASQA"
    contenu = contenu & "AQDoIAAAAACGGKsAFAACAPwgAAAAAJEAYAFFAAIA1CEAAAAAkQBlAUsAAwDQIwAAAACRAI4BUgAGAEAkAAAAAJEAqAFeAAcAsCQA"
    contenu = contenu & "AAAAkQDuAWQACACgJgAAAACRAEUCbAAKAPQnAAAAAJEAVwJ2AA4ARCwAAAAAkQBhAoEADwBMLQAAAACRAGsCkAAQADguAAAAAJEA"
    contenu = contenu & "fAKbABIAYC8AAAAAkQCNAqYAFwCcLwAAAACRAJUCrAAZAAwwAAAAAJEApQK2ABoAQDAAAAAAkQCvArsAGwB4MAAAAACRAMYCuwAd"
    contenu = contenu & "AIswAAAAAJEA2QLDAB8AAAABAO8CAAABAPQCAAABAPkCAAACAAEDAAADAAwDAAABABUDAAABAB8DAAABACYDAAACACsDAAABADwD"
    contenu = contenu & "AAACAEMDAAADAEoDAAAEAFQDAAABAEMDAAABAF8DAAABAF8DAAACAEoDAAABAGUDAAACAEMDAAADAGoDAAAEAEoDAAAFAFQDAAAB"
    contenu = contenu & "AHUDAAACAHsDAAABAF8DAAABAIEDAAABAIYDAAACAIsDAAABAIYDAAACAIsDAAABAIYDAAACAIsDSQCrAMsAUQCrABQACQCrABQA"
    contenu = contenu & "DAARBNcADAAGAdsADACrABQAWQCrABQAYQBGBOoAcQBXBPIAeQCWBPcAgQCgBPwAeQCgBAEBiQDLBAYBkQADBQsBqQARBbYAsQAs"
    contenu = contenu & "BRYBuQA0BRsBsQBBBRsBYQBNBR8BwQBZBS8ByQCrADQB0QCrAPwA2QCsBToB4QDaBUkB4QDsBVkBOQCrABQAOQD6BV0B4QAFBmMB"
    contenu = contenu & "8QCrABQAcQBABmkBMQBqBm0B8QB+BnYBAQG7BnsBAQH4Bo0BcQAPB6MBOQD6BbQBIQElB7kBCQEuB9cACQE9B9cAKQFZBxQAqQBh"
    contenu = contenu & "B7YAYQByBy8BMQGYB78BQQGrAMYBCQAlBxsBwQC+B8sBqQDLB9QBUQH4B/wBUQEBCAICWQEkCBQAYQGrAPwAaQFUCAYCWQFmCAsC"
    contenu = contenu & "FABxCBgCYQBNBR0CIQBxCBQAMQAYCUsCuQA6CVECqQGkCWICHACsCX8CHACzCYwCwQGsCZACqQHzCZoCJACsCX8CJACzCYwCLAD9"
    contenu = contenu & "CckCNAD9CckC2QEJCt0C2QENCt0C2QERCuMCKQCrABQAKQAXCugCKQAsCugC2QFCCuMCPAARBNcAPABVCgoDRABjChgCTACrACQD"
    contenu = contenu & "DAB8CioDRACBCjUDQQA0C2sDVABVCn8DXABjChgCAQJUC5IDZABVCn8DbABjChgCEQJeCxsBYQBnC9cAEQJ3C68DqQGIC7UDdACs"
    contenu = contenu & "CX8CdACzCYwCfAD9CfEDIQKBCjUDPACrABQADABVCgoDhABjChgC2QG0C+MCPAAGAdsAhACBCjUDjACrACQDPAB8CioDOQCrAMsA"
    contenu = contenu & "OQBnC9cA2QEJCmkEOQC4C28EOQC4C10BYQDEC3YEOQC4C40EkQAlB7kBYQDMC5MEmQElB6sElACrABQAlAAGAdsAlAB8ChQAlAAR"
    contenu = contenu & "BNcAlADUC7gEYQDdCy8BYQDwCwoAYQDMC88EYQD2CxsBmQH7C9UEMQKrABQACAAkAC0ACAAoADIACAAsADcADQAwADwALgALAN8E"
    contenu = contenu & "LgATAOgEowDbA+UAwAA7AOUAwwDbA+UA4QAmAdoBIwLtAjkD/gNJBHwEmQS+BNoEAgABAAAAsQAYAAAAuQAYAAIAAQADAAIAAgAF"
    contenu = contenu & "ANAAEQJvAq4CvgLSAgMDFAMcA3YDiQOdA6YD0APjAzkEQQSyBASAAAAAAAAAAAAAAAAAAAAAAO8DAAAEAAAAAAAAAAAAAAABAGIA"
    contenu = contenu & "AAAAAP8A/wD/AP8AAAIAAAAAaQEAAAAA/wD/AP8A/wAAAgAAAACzAQAAAAD/AP8A/wD/AAACAAAAABsCAAAAAP8A/wD/AP8AAAIA"
    contenu = contenu & "AAAAkAUAAAAA/wD/AP8A/wAAAgAAAACDBgAAAAAEAAAAAAAAAAAAAAABAGsAAAAAAAQAAAAAAAAAAAAAAAEAoggAAAAABAAAAAAA"
    contenu = contenu & "AAAAAAAAAQD1CAAAAAAEAAAAAAAAAAAAAABZAkwJAAAAAAUABAAGAAQAAAAAABEAmQEQAEQBEABUARAAhwEQAJ0BEACvAQAAADxN"
    contenu = contenu & "b2R1bGU+AFdpbk9jckJyaWRnZV9TcGF0aWFsX01BSjVfQ29ycmlnZS5leGUAV29yZEJveABXaW5PY3JCcmlkZ2VTcGF0aWFsTWFq"
    contenu = contenu & "NQBMYXlvdXRSb3cAUHJvZ3JhbQBtc2NvcmxpYgBTeXN0ZW0AT2JqZWN0AFRleHQATGVmdABUb3AAV2lkdGgASGVpZ2h0AGdldF9D"
    contenu = contenu & "ZW50ZXJZAGdldF9Cb3R0b20ALmN0b3IAQ2VudGVyWQBCb3R0b20AU3lzdGVtLkNvbGxlY3Rpb25zLkdlbmVyaWMATGlzdGAxAFdv"
    contenu = contenu & "cmRzAFdlaWdodGVkQ2VudGVyWQBBdmVyYWdlSGVpZ2h0AEFkZABMYXlvdXRDb2x1bW5zAE1heGltdW1MYXlvdXRSb3dzAE9wZXJh"
    contenu = contenu & "dGlvblRpbWVvdXRNaWxsaXNlY29uZHMAQ2hhcmFjdGVyQ2VsbEFzcGVjdFJhdGlvAE1haW4AUnVuAFdpbmRvd3MuRm91bmRhdGlv"
    contenu = contenu & "bgBJQXN5bmNPcGVyYXRpb25gMQBXYWl0UmVzdWx0AFQASUFzeW5jQWN0aW9uAFdhaXRBY3Rpb24AV2luZG93cy5EYXRhAFdpbmRv"
    contenu = contenu & "d3MuRGF0YS5QZGYAUGRmUGFnZVJlbmRlck9wdGlvbnMAUGRmUGFnZQBCdWlsZFJlbmRlck9wdGlvbnMAU3lzdGVtLlRleHQAU3Ry"
    contenu = contenu & "aW5nQnVpbGRlcgBXaW5kb3dzLk1lZGlhAFdpbmRvd3MuTWVkaWEuT2NyAE9jclJlc3VsdABBcHBlbmRTcGF0aWFsUGFnZQBSZWFk"
    contenu = contenu & "V29yZHMAQnVpbGRSb3dzAEJ1aWxkU3BhdGlhbExpbmUAQXBwZW5kTGF5b3V0RGF0YQBQZXJjZW50AEdldE1lZGlhbkhlaWdodABD"
    contenu = contenu & "bGVhbldvcmQAQ29tcGFyZVdvcmRzQnlQb3NpdGlvbgBDb21wYXJlV29yZHNCeUxlZnQAQ29tcGFyZVJvd3NCeVBvc2l0aW9uAHdv"
    contenu = contenu & "cmQAYXJncwBwZGZQYXRoAG91dHB1dFBhdGgAbWF4UGFnZXMAb3BlcmF0aW9uAGFjdGlvbgBwYWdlAG1heGltdW1EaW1lbnNpb24A"
    contenu = contenu & "b3V0cHV0AHJlc3VsdABwYWdlV2lkdGgAcGFnZUhlaWdodAB3b3JkcwBkYXRhAHBhZ2VOdW1iZXIAdmFsdWUAdG90YWwAdGV4dABs"
    contenu = contenu & "ZWZ0AHJpZ2h0AFN5c3RlbS5SdW50aW1lLkNvbXBpbGVyU2VydmljZXMAQ29tcGlsYXRpb25SZWxheGF0aW9uc0F0dHJpYnV0ZQBS"
    contenu = contenu & "dW50aW1lQ29tcGF0aWJpbGl0eUF0dHJpYnV0ZQBXaW5PY3JCcmlkZ2VfU3BhdGlhbF9NQUo1X0NvcnJpZ2UAZ2V0X0NvdW50AFNU"
    contenu = contenu & "QVRocmVhZEF0dHJpYnV0ZQBTdHJpbmcAU3RyaW5nQ29tcGFyaXNvbgBFcXVhbHMAT2NyRW5naW5lAFRyeUNyZWF0ZUZyb21Vc2Vy"
    contenu = contenu & "UHJvZmlsZUxhbmd1YWdlcwBDb25zb2xlAFN5c3RlbS5JTwBUZXh0V3JpdGVyAGdldF9FcnJvcgBXcml0ZUxpbmUAU3lzdGVtLkds"
    contenu = contenu & "b2JhbGl6YXRpb24AQ3VsdHVyZUluZm8AZ2V0X0ludmFyaWFudEN1bHR1cmUASW50MzIATnVtYmVyU3R5bGVzAElGb3JtYXRQcm92"
    contenu = contenu & "aWRlcgBUcnlQYXJzZQBQYXRoAEdldEZ1bGxQYXRoAEV4Y2VwdGlvbgBUeXBlAEdldFR5cGUAZ2V0X0Z1bGxOYW1lAGdldF9NZXNz"
    contenu = contenu & "YWdlAENvbmNhdABGaWxlAEV4aXN0cwBGaWxlTm90Rm91bmRFeGNlcHRpb24ASW52YWxpZE9wZXJhdGlvbkV4Y2VwdGlvbgBXaW5k"
    contenu = contenu & "b3dzLlN0b3JhZ2UAU3RvcmFnZUZpbGUAR2V0RmlsZUZyb21QYXRoQXN5bmMAUGRmRG9jdW1lbnQASVN0b3JhZ2VGaWxlAExvYWRG"
    contenu = contenu & "cm9tRmlsZUFzeW5jAGdldF9QYWdlQ291bnQAQXBwZW5kTGluZQBHZXRQYWdlAFdpbmRvd3MuU3RvcmFnZS5TdHJlYW1zAEluTWVt"
    contenu = contenu & "b3J5UmFuZG9tQWNjZXNzU3RyZWFtAGdldF9NYXhJbWFnZURpbWVuc2lvbgBJUmFuZG9tQWNjZXNzU3RyZWFtAFJlbmRlclRvU3Ry"
    contenu = contenu & "ZWFtQXN5bmMAU2VlawBXaW5kb3dzLkdyYXBoaWNzAFdpbmRvd3MuR3JhcGhpY3MuSW1hZ2luZwBCaXRtYXBEZWNvZGVyAENyZWF0"
    contenu = contenu & "ZUFzeW5jAFNvZnR3YXJlQml0bWFwAEJpdG1hcFBpeGVsRm9ybWF0AEJpdG1hcEFscGhhTW9kZQBHZXRTb2Z0d2FyZUJpdG1hcEFz"
    contenu = contenu & "eW5jAFJlY29nbml6ZUFzeW5jAFVJbnQzMgBUb1N0cmluZwBnZXRfUGl4ZWxXaWR0aABnZXRfUGl4ZWxIZWlnaHQASURpc3Bvc2Fi"
    contenu = contenu & "bGUARGlzcG9zZQBHZXREaXJlY3RvcnlOYW1lAElzTnVsbE9yRW1wdHkARGlyZWN0b3J5AERpcmVjdG9yeUluZm8AQ3JlYXRlRGly"
    contenu = contenu & "ZWN0b3J5AFVURjhFbmNvZGluZwBFbmNvZGluZwBXcml0ZUFsbFRleHQAQ2hhbmdlRXh0ZW5zaW9uAFN5c3RlbS5EaWFnbm9zdGlj"
    contenu = contenu & "cwBTdG9wd2F0Y2gAU3RhcnROZXcAZ2V0X0VsYXBzZWRNaWxsaXNlY29uZHMASUFzeW5jSW5mbwBDYW5jZWwAVGltZW91dEV4Y2Vw"
    contenu = contenu & "dGlvbgBTeXN0ZW0uVGhyZWFkaW5nAFRocmVhZABTbGVlcABBc3luY1N0YXR1cwBnZXRfU3RhdHVzAEdldFJlc3VsdHMAPEJ1aWxk"
    contenu = contenu & "UmVuZGVyT3B0aW9ucz5vX19TaXRlQ29udGFpbmVyMABTeXN0ZW0uQ29yZQBDYWxsU2l0ZWAxAEZ1bmNgMwBDYWxsU2l0ZQA8PnBf"
    contenu = contenu & "X1NpdGUxADw+cF9fU2l0ZTIAPD5wX19TaXRlMwA8PnBfX1NpdGU0AFN5c3RlbS5SdW50aW1lLldpbmRvd3NSdW50aW1lAFNpemUA"
    contenu = contenu & "Z2V0X1NpemUARG91YmxlAFJ1bnRpbWVUeXBlSGFuZGxlAEdldFR5cGVGcm9tSGFuZGxlAE1pY3Jvc29mdC5DU2hhcnAATWljcm9z"
    contenu = contenu & "b2Z0LkNTaGFycC5SdW50aW1lQmluZGVyAEJpbmRlcgBDYWxsU2l0ZUJpbmRlcgBDU2hhcnBCaW5kZXJGbGFncwBDb252ZXJ0AENy"
    contenu = contenu & "ZWF0ZQBUYXJnZXQAQ1NoYXJwQXJndW1lbnRJbmZvAENTaGFycEFyZ3VtZW50SW5mb0ZsYWdzAElFbnVtZXJhYmxlYDEAR2V0TWVt"
    contenu = contenu & "YmVyAEludm9rZQBNYXRoAE1heABNaW4AUm91bmQAcHV0X0Rlc3RpbmF0aW9uV2lkdGgAcHV0X0Rlc3RpbmF0aW9uSGVpZ2h0AENl"
    contenu = contenu & "aWxpbmcARW51bWVyYXRvcgBHZXRFbnVtZXJhdG9yAGdldF9DdXJyZW50AENvbXBhcmlzb25gMQBTb3J0AE1vdmVOZXh0ADxSZWFk"
    contenu = contenu & "V29yZHM+b19fU2l0ZUNvbnRhaW5lcjUAPD5wX19TaXRlNgA8PnBfX1NpdGU3ADw+cF9fU2l0ZTgAPD5wX19TaXRlOQA8PnBfX1Np"
    contenu = contenu & "dGVhAEZ1bmNgNQA8PnBfX1NpdGViADw+cF9fU2l0ZWMAPD5wX19TaXRlZAA8PnBfX1NpdGVlADw+cF9fU2l0ZWYASVJlYWRPbmx5"
    contenu = contenu & "TGlzdGAxAE9jckxpbmUAZ2V0X0xpbmVzAElFbnVtZXJhdG9yYDEAT2NyV29yZABnZXRfV29yZHMAZ2V0X1RleHQAZ2V0X0xlbmd0"
    contenu = contenu & "aABSZWN0AGdldF9Cb3VuZGluZ1JlY3QASW52b2tlTWVtYmVyAFN5c3RlbS5Db2xsZWN0aW9ucwBJRW51bWVyYXRvcgBBYnMAQXBw"
    contenu = contenu & "ZW5kAENoYXIAVHJpbUVuZABSZXBsYWNlAGdldF9JdGVtAElzTnVsbE9yV2hpdGVTcGFjZQBFbXB0eQBUcmltAENvbXBhcmVUbwBD"
    contenu = contenu & "b21waWxlckdlbmVyYXRlZEF0dHJpYnV0ZQAADy0ALQBjAGgAZQBjAGsAAR9PAEMAUgBfAFUATgBBAFYAQQBJAEwAQQBCAEwARQAA"
    contenu = contenu & "G08AQwBSAF8AQQBWAEEASQBMAEEAQgBMAEUAAHdVAHMAYQBnAGUAOgAgAFcAaQBuAE8AYwByAEIAcgBpAGQAZwBlAC4AZQB4AGUA"
    contenu = contenu & "IAA8AGkAbgBwAHUAdAAuAHAAZABmAD4AIAA8AG8AdQB0AHAAdQB0AC4AdAB4AHQAPgAgAFsAbQBhAHgAUABhAGcAZQBzAF0AAAU6"
    contenu = contenu & "ACAAAB9QAEQARgAgAGkAbgB0AHIAbwB1AHYAYQBiAGwAZQAAgIFPAEMAUgAgAFcAaQBuAGQAbwB3AHMAIABpAG4AZABpAHMAcABv"
    contenu = contenu & "AG4AaQBiAGwAZQAgAHAAbwB1AHIAIABsAGUAcwAgAGwAYQBuAGcAdQBlAHMAIABkAHUAIABwAHIAbwBmAGkAbAAgAHUAdABpAGwA"
    contenu = contenu & "aQBzAGEAdABlAHUAcgAuAACAg3AAYQBnAGUACQBwAGEAZwBlAF8AdwBpAGQAdABoAAkAcABhAGcAZQBfAGgAZQBpAGcAaAB0AAkA"
    contenu = contenu & "eABfAHAAYwB0AAkAeQBfAHAAYwB0AAkAdwBpAGQAdABoAF8AcABjAHQACQBoAGUAaQBnAGgAdABfAHAAYwB0AAkAdABlAHgAdAAA"
    contenu = contenu & "Fy0ALQAtAC0ALQAgAFAAQQBHAEUAIAABDSAALQAtAC0ALQAtAAEXLgBsAGEAeQBvAHUAdAAuAHQAcwB2AAA1RABlAGwAYQBpACAA"
    contenu = contenu & "TwBDAFIAIABXAGkAbgBkAG8AdwBzACAAZABlAHAAYQBzAHMAZQAuAABBTwBwAGUAcgBhAHQAaQBvAG4AIABXAGkAbgBkAG8AdwBz"
    contenu = contenu & "ACAAaQBuAHQAZQByAHIAbwBtAHAAdQBlACAAOgAgAAA3RABlAGwAYQBpACAAZABlACAAcgBlAG4AZAB1ACAAUABEAEYAIABkAGUA"
    contenu = contenu & "cABhAHMAcwBlAC4AADtBAGMAdABpAG8AbgAgAFcAaQBuAGQAbwB3AHMAIABpAG4AdABlAHIAcgBvAG0AcAB1AGUAIAA6ACAAAAtX"
    contenu = contenu & "AGkAZAB0AGgAAA1IAGUAaQBnAGgAdAAAA1gAAANZAAAHTQBhAHgAAAMJAAADIAAABUYANAAAAAAA/AiX4KuRHE6h0rCGJ85cvAAI"
    contenu = contenu & "t3pcVhk04IkCBg4CBg0DIAANAyAAAQMoAA0HBhUSCQESCAUgAQESCAIGCARoAQAABBgBAAAEwNQBAAgAAAAAAADgPwUAAQgdDgYA"
    contenu = contenu & "AwgODggLEAEBHgAVEg0BHgAFAAEBEhEHAAISFRIZCQkABAESHRIhCAgKAAEVEgkBEggSIQ4AARUSCQESDBUSCQESCAoAAg4VEgkB"
    contenu = contenu & "EggICgAFARIdEiEJCAgFAAIODQgJAAENFRIJARIIBAABDg4HAAIIEggSCAcAAggSDBIMBCABAQgGFRIJARIIAyAACAUgAQETAAMH"
    contenu = contenu & "AQ0EAQAAAAcAAwIODhE1BAAAEjkEAAASQQQgAQEOBAABAQ4EAAASRQoABAIOEU0SURAIBCAAEl0DIAAOBgADDg4ODggHBBI5CBJZ"
    contenu = contenu & "CAQAAQIOBSACAQ4OCQABFRINARJtDgQKARJtCgABFRINARJxEnUECgEScQMgAAkFIAESHQ4FIAESGQkDAAAJCCACEhESfRIVBCAB"
    contenu = contenu & "AQsLAAEVEg0BEoCBEn0FCgESgIEPIAIVEg0BEoCFEYCJEYCNBQoBEoCFCyABFRINARIhEoCFBAoBEiEEIAASHQUgAQ4SUQYAARKA"
    contenu = contenu & "nQ4EIAEBAggAAwEODhKApQUAAg4ODiEHEBI5Em0ScQkSHRIdCRIZEnkSFRKAgRKAhRIhDhKAoQkFAAASgKkDIAAKBAABAQgFIAAR"
    contenu = contenu & "gLkGFRINAR4ABCAAEwAFAAIOHBwFBwESgKkQBhUSgL0BFRKAwQMSgMUcDRAGFRKAvQEVEoDBAxKAxRwcBSAAEYDJBwABEl0RgNEI"
    contenu = contenu & "sD9ffxHVCjoMAAMSgNkRgN0SXRJdDxUSgL0BFRKAwQMSgMUcDQwAARUSgL0BEwASgNkDBhMACQACEoDhEYDlDhMABBKA2RGA3Q4S"
    contenu = contenu & "XRUSgOkBEoDhDxUSgL0BFRKAwQMSgMUcHAoVEoDBAxKAxRwcCCACEwITABMBChUSgMEDEoDFHA0FAAINDQ0EAAENDQQgAQEJFQcM"
    contenu = contenu & "HA0NDQ0NDQkJEhUdEoDhHRKA4QYVEgkBEgwJIAAVEYDxARMABxURgPEBEgwHFRKA9QESCAUgAgEcGAogAQEVEoD1ARMAAyAAAh0H"
    contenu = contenu & "ChUSCQESCBUSCQESDA0NDQgIEgwIFRGA8QESDBMGFRKAvQEVEoD5BRKAxRJdDRwcCiAAFRKA/QESgQEIFRKA6QESgQEJIAAVEoEF"
    contenu = contenu & "ARMACBUSgQUBEoEBCiAAFRKA/QESgQkIFRKA6QESgQkIFRKBBQESgQkFIAARgQ0aAAUSgNkRgN0OFRKA6QESXRJdFRKA6QESgOES"
    contenu = contenu & "FRKAvQEVEoD5BRKAxRJdDRwcDRUSgPkFEoDFEl0NHBwMIAQTBBMAEwETAhMDOgcOFRIJARIIEoEBEoEJDhwSCBUSgQUBEoEBFRKB"
    contenu = contenu & "BQESgQkdEoDhHRKA4R0SgOEdEoDhHRKA4R0SgOEHFRGA8QESCAcVEoD1ARIMHwcJFRIJARIMEggSDA0SDA0NFRGA8QESCBURgPEB"
    contenu = contenu & "EgwFAAIICAgGIAISHQMIBSABDh0DEAcGEh0SCAgICBURgPEBEggFIAESHQMFIAIODg4RBwMVEgkBEggSCBURgPEBEggGIAIODhJR"
    contenu = contenu & "BRUSCQENBSABEwAIEAcDFRIJAQ0SCBURgPEBEggFIAIOAwMEIAEIDQQHAggNCAEACAAAAAAAHgEAAQBUAhZXcmFwTm9uRXhjZXB0"
    contenu = contenu & "aW9uVGhyb3dzAQAgTwAAAAAAAAAAAAA+TwAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAME8AAAAAAAAAAAAAAAAAAAAAX0NvckV4"
    contenu = contenu & "ZU1haW4AbXNjb3JlZS5kbGwAAAAAAP8lACBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAEAAAACAAAIAY"
    contenu = contenu & "AAAAOAAAgAAAAAAAAAAAAAAAAAAAAQABAAAAUAAAgAAAAAAAAAAAAAAAAAAAAQABAAAAaAAAgAAAAAAAAAAAAAAAAAAAAQAAAAAA"
    contenu = contenu & "gAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAkAAAAKBgAACsAgAAAAAAAAAAAABQYwAA6gEAAAAAAAAAAAAArAI0AAAAVgBTAF8AVgBF"
    contenu = contenu & "AFIAUwBJAE8ATgBfAEkATgBGAE8AAAAAAL0E7/4AAAEAAAAAAAAAAAAAAAAAAAAAAD8AAAAAAAAABAAAAAEAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AABEAAAAAQBWAGEAcgBGAGkAbABlAEkAbgBmAG8AAAAAACQABAAAAFQAcgBhAG4AcwBsAGEAdABpAG8AbgAAAAAAAACwBAwCAAAB"
    contenu = contenu & "AFMAdAByAGkAbgBnAEYAaQBsAGUASQBuAGYAbwAAAOgBAAABADAAMAAwADAAMAA0AGIAMAAAACwAAgABAEYAaQBsAGUARABlAHMA"
    contenu = contenu & "YwByAGkAcAB0AGkAbwBuAAAAAAAgAAAAMAAIAAEARgBpAGwAZQBWAGUAcgBzAGkAbwBuAAAAAAAwAC4AMAAuADAALgAwAAAAbAAm"
    contenu = contenu & "AAEASQBuAHQAZQByAG4AYQBsAE4AYQBtAGUAAABXAGkAbgBPAGMAcgBCAHIAaQBkAGcAZQBfAFMAcABhAHQAaQBhAGwAXwBNAEEA"
    contenu = contenu & "SgA1AF8AQwBvAHIAcgBpAGcAZQAuAGUAeABlAAAAKAACAAEATABlAGcAYQBsAEMAbwBwAHkAcgBpAGcAaAB0AAAAIAAAAHQAJgAB"
    contenu = contenu & "AE8AcgBpAGcAaQBuAGEAbABGAGkAbABlAG4AYQBtAGUAAABXAGkAbgBPAGMAcgBCAHIAaQBkAGcAZQBfAFMAcABhAHQAaQBhAGwA"
    contenu = contenu & "XwBNAEEASgA1AF8AQwBvAHIAcgBpAGcAZQAuAGUAeABlAAAANAAIAAEAUAByAG8AZAB1AGMAdABWAGUAcgBzAGkAbwBuAAAAMAAu"
    contenu = contenu & "ADAALgAwAC4AMAAAADgACAABAEEAcwBzAGUAbQBiAGwAeQAgAFYAZQByAHMAaQBvAG4AAAAwAC4AMAAuADAALgAwAAAAAAAAAO+7"
    contenu = contenu & "vzw/eG1sIHZlcnNpb249IjEuMCIgZW5jb2Rpbmc9IlVURi04IiBzdGFuZGFsb25lPSJ5ZXMiPz4NCjxhc3NlbWJseSB4bWxucz0i"
    contenu = contenu & "dXJuOnNjaGVtYXMtbWljcm9zb2Z0LWNvbTphc20udjEiIG1hbmlmZXN0VmVyc2lvbj0iMS4wIj4NCiAgPGFzc2VtYmx5SWRlbnRp"
    contenu = contenu & "dHkgdmVyc2lvbj0iMS4wLjAuMCIgbmFtZT0iTXlBcHBsaWNhdGlvbi5hcHAiLz4NCiAgPHRydXN0SW5mbyB4bWxucz0idXJuOnNj"
    contenu = contenu & "aGVtYXMtbWljcm9zb2Z0LWNvbTphc20udjIiPg0KICAgIDxzZWN1cml0eT4NCiAgICAgIDxyZXF1ZXN0ZWRQcml2aWxlZ2VzIHht"
    contenu = contenu & "bG5zPSJ1cm46c2NoZW1hcy1taWNyb3NvZnQtY29tOmFzbS52MyI+DQogICAgICAgIDxyZXF1ZXN0ZWRFeGVjdXRpb25MZXZlbCBs"
    contenu = contenu & "ZXZlbD0iYXNJbnZva2VyIiB1aUFjY2Vzcz0iZmFsc2UiLz4NCiAgICAgIDwvcmVxdWVzdGVkUHJpdmlsZWdlcz4NCiAgICA8L3Nl"
    contenu = contenu & "Y3VyaXR5Pg0KICA8L3RydXN0SW5mbz4NCjwvYXNzZW1ibHk+DQoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAQAAADAAAAFA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    contenu = contenu & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=="
    PontIntegreBase64Cegide = contenu
End Function

Private Function LirePontMemoriseCegide() As String
    Dim nomParametre As Name
    Dim valeur As String

    On Error Resume Next
    Set nomParametre = ThisWorkbook.Names(NOM_PARAM_PONT)
    On Error GoTo 0
    If Not nomParametre Is Nothing Then
        valeur = CStr(nomParametre.RefersTo)
        If Left$(valeur, 2) = "=""" And Right$(valeur, 1) = """" Then
            valeur = Mid$(valeur, 3, Len(valeur) - 3)
            valeur = Replace(valeur, """""", """")
        End If
    End If
    If valeur = "" Then valeur = GetSetting(REG_APPLICATION, REG_SECTION, REG_CLE_PONT, "")
    LirePontMemoriseCegide = valeur
End Function

Private Sub EnregistrerPontCegide(ByVal chemin As String)
    On Error Resume Next
    ThisWorkbook.Names(NOM_PARAM_PONT).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NOM_PARAM_PONT, RefersTo:="=""" & Replace(chemin, """", """""") & """", Visible:=False
    SaveSetting REG_APPLICATION, REG_SECTION, REG_CLE_PONT, chemin
End Sub

Private Sub EffacerPontMemoriseCegide()
    On Error Resume Next
    ThisWorkbook.Names(NOM_PARAM_PONT).Delete
    DeleteSetting REG_APPLICATION, REG_SECTION, REG_CLE_PONT
    On Error GoTo 0
End Sub

Private Function PontDisponibleCegide(ByVal chemin As String) As Boolean
    Dim fso As Object
    Dim shell As Object
    Dim codeRetour As Long
    Dim nomFichier As String

    On Error GoTo Fin
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(chemin) Then Exit Function
    nomFichier = UCase$(fso.GetBaseName(chemin))
    If InStr(1, nomFichier, "WINOCRBRIDGE", vbTextCompare) = 0 Then Exit Function
    If LCase$(fso.GetExtensionName(chemin)) <> "exe" Then Exit Function
    Set shell = CreateObject("WScript.Shell")
    codeRetour = shell.Run(EncadrerCegide(chemin) & " --check", 0, True)
    PontDisponibleCegide = (codeRetour = 0)
Fin:
End Function

Private Function ExecuterOCRCegide(ByVal pontOCR As String, ByVal cheminPDF As String, ByVal cheminTXT As String, ByVal fso As Object, ByRef messageErreur As String) As String
    Dim shell As Object
    Dim commande As String
    Dim codeRetour As Long

    On Error GoTo Echec
    messageErreur = ""
    If fso.FileExists(cheminTXT) Then fso.DeleteFile cheminTXT, True
    If fso.FileExists(Left$(cheminTXT, Len(cheminTXT) - 4) & ".layout.tsv") Then fso.DeleteFile Left$(cheminTXT, Len(cheminTXT) - 4) & ".layout.tsv", True
    commande = EncadrerCegide(pontOCR) & " " & EncadrerCegide(cheminPDF) & " " & EncadrerCegide(cheminTXT) & " 1"
    Set shell = CreateObject("WScript.Shell")
    codeRetour = shell.Run(commande, 0, True)
    If codeRetour <> 0 Then
        messageErreur = "OCR code " & CStr(codeRetour) & ". Verifiez que le fichier selectionne est bien WinOcrBridge.exe."
        Exit Function
    End If
    If Not fso.FileExists(cheminTXT) Then
        messageErreur = "Le pont OCR n'a pas cree le fichier TXT."
        Exit Function
    End If
    ExecuterOCRCegide = LireUTF8Cegide(cheminTXT)
    PreparerDispositionCegide cheminPDF, cheminTXT, fso.GetParentFolderName(pontOCR)
    Exit Function
Echec:
    messageErreur = "Erreur OCR " & CStr(Err.Number) & " : " & Err.Description
End Function

Private Sub SauverTexteUTF8Cegide(ByVal chemin As String, ByVal texte As String)
    Dim flux As Object
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 2: flux.Charset = "utf-8": flux.Open
    flux.WriteText texte
    flux.SaveToFile chemin, 2
    flux.Close
End Sub

Private Function PontSecondeLectureCegide(ByVal dossier As String, Optional ByVal lectureClient As Boolean = False, Optional ByVal basseResolution As Boolean = False) As String
    Dim xml As Object, noeud As Object, flux As Object, donnees As Variant, attendu As Variant, nouveau As Variant, i As Long
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set noeud = xml.createElement("pont")
    noeud.DataType = "bin.base64": noeud.Text = PontIntegreBase64Cegide()
    donnees = noeud.nodeTypedValue
    attendu = Array(0, 0, 0, 0, 0, 192, 162, 64)
    nouveau = Array(0, 0, 0, 0, 0, 32, 156, 64)
    If lectureClient Then nouveau = Array(0, 0, 0, 0, 0, 192, 146, 64)
    If basseResolution Then nouveau = Array(0, 0, 0, 0, 0, 32, 140, 64)
    ' Variante du pont embarque connu : seule la constante ldc.r8 de resolution change.
    If UBound(donnees) + 1 <> TAILLE_PONT_INTEGRE Then Err.Raise vbObjectError + 870, , "Pont OCR incompatible"
    If donnees(2096) <> 35 Then Err.Raise vbObjectError + 870, , "Instruction OCR incompatible"
    For i = 0 To 7
        If donnees(2097 + i) <> attendu(i) Then Err.Raise vbObjectError + 870, , "Resolution OCR incompatible"
        donnees(2097 + i) = CByte(nouveau(i))
    Next i
    PontSecondeLectureCegide = dossier & "\WinOcrBridge_lecture2.exe"
    If lectureClient Then PontSecondeLectureCegide = dossier & "\WinOcrBridge_client.exe"
    If basseResolution Then PontSecondeLectureCegide = dossier & "\WinOcrBridge_client900.exe"
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 1: flux.Open: flux.Write donnees
    flux.SaveToFile PontSecondeLectureCegide, 2: flux.Close
End Function

Private Function EnteteHTManquanteCegide(ByVal mots As Collection) As Boolean
    Dim lignes As Collection, ligne As Object, mot As Object, colonnes As Collection, colonne As Object
    Dim libelle As Boolean, ht As Boolean, montant As Boolean, n As String
    Set lignes = LignesPositionsCegide(mots)
    For Each ligne In lignes
        libelle = False: ht = False: montant = False
        For Each mot In ligne("m")
            n = NormaliserCompactCegide(CStr(mot("t")))
            If n = "LIBELLE" Or n = "NATURE" Then libelle = True
        Next mot
        If libelle Then
            Set colonnes = ColonnesPositionsCegide(ligne, mots)
            For Each colonne In colonnes
                If colonne("type") = "HT" Then ht = True
                If colonne("type") = "TTC" Or colonne("type") = "NS" Or colonne("type") = "TVA" Then montant = True
            Next colonne
            If montant And Not ht Then EnteteHTManquanteCegide = True: Exit Function
        End If
    Next ligne
End Function

Private Function EnteteDansZoneTableauCegide(ByVal candidat As Object, ByVal mots As Collection) As Boolean
    Dim lignes As Collection, ligne As Object, mot As Object, colonne As Object, colonnes As Collection
    Dim ancre As Boolean, montant As Boolean
    Set lignes = LignesPositionsCegide(mots)
    For Each ligne In lignes
        If ligne("p") = candidat("p") And Abs(ligne("y") - candidat("y")) < Application.Max(ligne("h"), candidat("h")) * 1.5 Then
            ancre = False: montant = False
            For Each mot In ligne("m")
                If NormaliserCompactCegide(CStr(mot("t"))) = "LIBELLE" Or NormaliserCompactCegide(CStr(mot("t"))) = "NATURE" Then
                    If candidat("x") > mot("x") + mot("w") Then ancre = True
                End If
            Next mot
            Set colonnes = ColonnesPositionsCegide(ligne, mots)
            For Each colonne In colonnes
                If colonne("type") = "TVA" Or colonne("type") = "TTC" Or colonne("type") = "NS" Then
                    If candidat("x") < colonne("x") Then montant = True
                End If
            Next colonne
            If ancre And montant Then EnteteDansZoneTableauCegide = True: Exit Function
        End If
    Next ligne
End Function

Private Sub PreparerDispositionCegide(ByVal pdf As String, ByVal txt As String, ByVal dossierPont As String)
    Dim base As String, mots As Collection, seconde As Collection, mot As Object, original As Object
    Dim shell As Object, fso As Object, pont As String, code As Long, present As Boolean, ajouts As Long, journal As String
    On Error GoTo EchecDisposition
    base = Left$(txt, Len(txt) - 4)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mots = LirePositionsCegide(base & ".layout.tsv")
    If EnteteHTManquanteCegide(mots) Then
        pont = PontSecondeLectureCegide(dossierPont)
        If fso.FileExists(base & ".lecture2.layout.tsv") Then fso.DeleteFile base & ".lecture2.layout.tsv", True
        Set shell = CreateObject("WScript.Shell")
        code = shell.Run(EncadrerCegide(pont) & " " & EncadrerCegide(pdf) & " " & EncadrerCegide(base & ".lecture2.txt") & " 1", 0, True)
        If code = 0 And fso.FileExists(base & ".lecture2.layout.tsv") Then
            Set seconde = LirePositionsCegide(base & ".lecture2.layout.tsv")
            For Each mot In seconde
                If mot("p") = 1 And TypeColonneSpatialCegide(CStr(mot("t"))) = "HT" Then
                    present = False
                    For Each original In mots
                        If original("p") = 1 And TypeColonneSpatialCegide(CStr(original("t"))) = "HT" Then
                            If Abs(original("x") - mot("x")) < 2 And Abs(original("y") - mot("y")) < 2 Then present = True: Exit For
                        End If
                    Next original
                    If Not present And EnteteDansZoneTableauCegide(mot, mots) Then mots.Add mot: ajouts = ajouts + 1
                End If
            Next mot
        Else
            journal = "Seconde lecture indisponible : sortie OCR initiale conservee." & vbCrLf
        End If
    End If
    If ajouts > 0 Then
        fso.CopyFile base & ".layout.tsv", base & ".original.layout.tsv", True
        SauverTexteUTF8Cegide base & ".layout.tsv", PositionsTexteCegide(mots)
    End If
    CompleterClientLectureCegide pdf, base, dossierPont, mots, journal
    SauverTexteUTF8Cegide base & ".disposition.txt", TexteDispositionCegide(mots)
    SauverTexteUTF8Cegide base & ".lecture.log.txt", journal & "En-tetes HT recuperes par seconde lecture : " & ajouts & vbCrLf & "Aucun montant de seconde lecture substitue. TXT original conserve."
    Exit Sub
EchecDisposition:
    On Error Resume Next
    SauverTexteUTF8Cegide base & ".lecture.log.txt", "Lecture complementaire indisponible : " & Err.Description & ". OCR initial conserve."
End Sub

Private Sub CompleterClientLectureCegide(ByVal pdf As String, ByVal base As String, ByVal dossier As String, ByVal mots As Collection, ByRef journal As String)
    Dim avis As String, module As String, client As String, seconde As Collection, mot As Object
    Dim fso As Object, shell As Object, pont As String, code As Long, resolution As Variant, sortie As String
    On Error GoTo Indisponible
    If ClientSpatialCegide(mots, avis) <> "" Then Exit Sub
    module = ModuleSpatialCegide(mots, avis)
    If Len(module) <> 10 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    For Each resolution In Array(1200, 900)
        pont = PontSecondeLectureCegide(dossier, True, CLng(resolution) = 900)
        sortie = base & ".client" & CStr(resolution)
        If fso.FileExists(sortie & ".layout.tsv") Then fso.DeleteFile sortie & ".layout.tsv", True
        code = shell.Run(EncadrerCegide(pont) & " " & EncadrerCegide(pdf) & " " & EncadrerCegide(sortie & ".txt") & " 1", 0, True)
        If code = 0 And fso.FileExists(sortie & ".layout.tsv") Then
            Set seconde = LirePositionsCegide(sortie & ".layout.tsv")
            Set mot = ClientExpliciteReluCegide(seconde, module)
            If Not mot Is Nothing Then
                mots.Add mot
                If Not fso.FileExists(base & ".original.layout.tsv") Then fso.CopyFile base & ".layout.tsv", base & ".original.layout.tsv", False
                SauverTexteUTF8Cegide base & ".layout.tsv", PositionsTexteCegide(mots)
                journal = journal & "Client relu a " & resolution & " pixels, module identique : " & mot("t") & vbCrLf
                Exit Sub
            End If
        End If
    Next resolution
    Exit Sub
Indisponible:
    journal = journal & "Lecture client complementaire indisponible ; donnees initiales conservees." & vbCrLf
End Sub

Private Function ClientExpliciteReluCegide(ByVal mots As Collection, ByVal module As String) As Object
    Dim ancre As Object, mot As Object, choisi As Object, copie As Object, cle As Variant
    Dim reg As Object, correspondances As Object, client As String, retenu As String
    If Len(module) <> 10 Then Exit Function
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "^[: -]*L(?:/|[,']{1,2})([0-9]{1,7})-?$"
    reg.IgnoreCase = True
    ' Le module initial a deja ete valide dans la zone des references.
    For Each ancre In mots
        If ancre("p") = 1 And InStr(CStr(ancre("t")), module) > 0 Then
            For Each mot In mots
                If mot("p") = 1 And Abs(mot("y") - ancre("y")) <= 4 And Abs((mot("x") + mot("w") / 2) - (ancre("x") + ancre("w") / 2)) <= 30 Then
                    Set correspondances = reg.Execute(Trim$(CStr(mot("t"))))
                    If correspondances.Count = 1 Then
                        client = "L/" & correspondances(0).SubMatches(0)
                        If retenu <> "" And client <> retenu Then Exit Function
                        retenu = client
                        Set choisi = mot
                    End If
                End If
            Next mot
        End If
    Next ancre
    If choisi Is Nothing Then Exit Function
    Set copie = CreateObject("Scripting.Dictionary")
    For Each cle In choisi.Keys
        copie.Add cle, choisi(cle)
    Next cle
    copie("t") = retenu
    Set ClientExpliciteReluCegide = copie
End Function

Private Function PositionsTexteCegide(ByVal mots As Collection) As String
    Dim mot As Object, champ As Variant, resultat As String
    resultat = "page" & vbTab & "width" & vbTab & "height" & vbTab & "x_pct" & vbTab & "y_pct" & vbTab & "w_pct" & vbTab & "h_pct" & vbTab & "text" & vbCrLf
    For Each mot In mots
        If mot("p") = 1 Then
            resultat = resultat & "1" & vbTab & "1000" & vbTab & "1000"
            For Each champ In Array("x", "y", "w", "h")
                resultat = resultat & vbTab & Replace(CStr(mot(champ)), ",", ".")
            Next champ
            resultat = resultat & vbTab & Replace(Replace(CStr(mot("t")), vbTab, " "), vbCrLf, " ") & vbCrLf
        End If
    Next mot
    PositionsTexteCegide = resultat
End Function

Private Function TexteDispositionCegide(ByVal mots As Collection) As String
    Dim lignes As Collection, ligne As Object, mot As Object, largeur As Long, position As Long, dernier As Long
    Dim texte As String, contenu As String, resultat As String, y As Double, blancs As Long
    largeur = 240
    For Each mot In mots
        If mot("p") = 1 And mot("w") > 0 Then largeur = Application.Max(largeur, Int((Len(CStr(mot("t"))) + 1) * 100 / mot("w")) + 1)
    Next mot
    largeur = Application.Min(largeur, 1000)
    Set lignes = LignesPositionsCegide(mots)
    For Each ligne In lignes
        If ligne("p") = 1 Then
            blancs = Application.Min(20, Application.Max(0, Int((ligne("y") - y) / 1.5) - 1))
            resultat = resultat & String$(blancs, vbLf)
            contenu = Space$(largeur): dernier = 0
            For Each mot In ligne("m")
                texte = Replace(Replace(CStr(mot("t")), vbCr, " "), vbLf, " ")
                position = Int(mot("x") * largeur / 100) + 1
                ' En cas de chevauchement, garder la colonne et une ligne distincte plutot que de deplacer le mot.
                If position <= dernier Then resultat = resultat & RTrim$(contenu) & vbLf: contenu = Space$(largeur)
                If position + Len(texte) > Len(contenu) Then contenu = contenu & Space$(position + Len(texte) - Len(contenu))
                Mid$(contenu, position, Len(texte)) = texte
                dernier = position + Len(texte)
            Next mot
            resultat = resultat & RTrim$(contenu) & vbLf
            y = ligne("y")
        End If
    Next ligne
    TexteDispositionCegide = Replace(resultat, vbLf, vbCrLf)
End Function

Private Function LireUTF8Cegide(ByVal chemin As String) As String
    Dim flux As Object
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 2
    flux.Charset = "utf-8"
    flux.Open
    flux.LoadFromFile chemin
    LireUTF8Cegide = flux.ReadText
    flux.Close
End Function

Private Function EncadrerCegide(ByVal valeur As String) As String
    EncadrerCegide = Chr$(34) & valeur & Chr$(34)
End Function

Private Function DossierTelechargementsCegide() As String
    Dim applicationShell As Object
    Dim dossier As Object
    Dim chemin As String
    On Error Resume Next
    Set applicationShell = CreateObject("Shell.Application")
    Set dossier = applicationShell.Namespace("shell:Downloads")
    If Not dossier Is Nothing Then chemin = CStr(dossier.Self.Path)
    On Error GoTo 0
    If chemin = "" Then chemin = Environ$("USERPROFILE") & "\Downloads"
    DossierTelechargementsCegide = chemin
End Function

Private Function ObtenirDossierFacturesCegide(ByVal fso As Object) As String
    Dim chemin As String

    chemin = LireDossierMemoriseCegide()
    If chemin <> "" Then
        If fso.FolderExists(chemin) Then
            ObtenirDossierFacturesCegide = chemin
            Exit Function
        End If
    End If

    chemin = DossierTelechargementsCegide()
    If chemin <> "" Then
        If fso.FolderExists(chemin) Then
            ObtenirDossierFacturesCegide = chemin
            Exit Function
        End If
    End If

    ObtenirDossierFacturesCegide = ChoisirEtEnregistrerDossierCegide(fso)
End Function

Private Function ChoisirEtEnregistrerDossierCegide(ByVal fso As Object) As String
    Dim selecteur As FileDialog
    Dim chemin As String

    Set selecteur = Application.FileDialog(msoFileDialogFolderPicker)
    With selecteur
        .Title = "Indiquez le dossier contenant les factures telechargees"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        chemin = CStr(.SelectedItems(1))
    End With
    If Not fso.FolderExists(chemin) Then
        Gloire_MessageAutomatique "Le dossier selectionne n'est pas accessible. Il ne sera pas memorise.", vbExclamation
        Exit Function
    End If

    EnregistrerDossierCegide chemin
    ChoisirEtEnregistrerDossierCegide = chemin
End Function

Private Function LireDossierMemoriseCegide() As String
    Dim nomParametre As Name
    Dim valeur As String

    On Error Resume Next
    Set nomParametre = ThisWorkbook.Names(NOM_PARAM_DOSSIER)
    On Error GoTo 0
    If Not nomParametre Is Nothing Then
        valeur = CStr(nomParametre.RefersTo)
        If Left$(valeur, 2) = "=""" And Right$(valeur, 1) = """" Then
            valeur = Mid$(valeur, 3, Len(valeur) - 3)
            valeur = Replace(valeur, """""", """")
        End If
    End If
    If valeur = "" Then valeur = GetSetting(REG_APPLICATION, REG_SECTION, REG_CLE_DOSSIER, "")
    LireDossierMemoriseCegide = valeur
End Function

Private Sub EnregistrerDossierCegide(ByVal chemin As String)
    On Error Resume Next
    ThisWorkbook.Names(NOM_PARAM_DOSSIER).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NOM_PARAM_DOSSIER, RefersTo:="=""" & Replace(chemin, """", """""") & """", Visible:=False
    SaveSetting REG_APPLICATION, REG_SECTION, REG_CLE_DOSSIER, chemin
End Sub

Private Function DernierPDFCegide(ByVal fso As Object, ByVal dossier As String) As String
    Dim fichier As Object
    Dim plusRecent As Date
    For Each fichier In fso.GetFolder(dossier).Files
        If LCase$(fso.GetExtensionName(CStr(fichier.Path))) = "pdf" And Left$(CStr(fichier.Name), 2) <> "~$" Then
            If DernierPDFCegide = "" Or fichier.DateCreated > plusRecent Then
                plusRecent = fichier.DateCreated
                DernierPDFCegide = CStr(fichier.Path)
            End If
        End If
    Next fichier
End Function

Private Function NettoyerNomFichierCegide(ByVal valeur As String) As String
    Dim caractere As Variant
    NettoyerNomFichierCegide = Trim$(valeur)
    For Each caractere In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        NettoyerNomFichierCegide = Replace(NettoyerNomFichierCegide, CStr(caractere), "_")
    Next caractere
End Function

Private Function TrouverFeuilleCegide(ByVal wb As Workbook, ByVal nomAttendu As String) As Worksheet
    Dim ws As Worksheet
    Dim cible As String
    cible = NormaliserCompactCegide(nomAttendu)
    For Each ws In wb.Worksheets
        If NormaliserCompactCegide(ws.Name) = cible Then
            Set TrouverFeuilleCegide = ws
            Exit Function
        End If
    Next ws
End Function

Private Function ChargerArticlesCegide(ByVal ws As Worksheet) As Collection
    Dim resultat As New Collection
    Dim candidats As Collection
    Dim derniereLigne As Long
    Dim ligne As Long
    Dim operationTexte As String
    Dim operationNorm As String
    Dim codeArticle As String
    Dim info As Object

    Set baseOperationsExactesCegide = CreateObject("Scripting.Dictionary")
    baseOperationsExactesCegide.CompareMode = vbTextCompare
    Set baseOperationsPrefixeCegide = CreateObject("Scripting.Dictionary")
    baseOperationsPrefixeCegide.CompareMode = vbTextCompare
    derniereLigne = Application.Max(ws.Cells(ws.Rows.Count, 3).End(xlUp).Row, ws.Cells(ws.Rows.Count, 6).End(xlUp).Row)
    For ligne = 2 To derniereLigne
        operationTexte = Trim$(CStr(ws.Cells(ligne, 3).Value2))
        codeArticle = Trim$(CStr(ws.Cells(ligne, 6).Value2))
        If operationTexte <> "" And codeArticle <> "" Then
            operationNorm = NormaliserCompactCegide(operationTexte)
            If operationNorm <> "OPERATION" And NormaliserCompactCegide(codeArticle) <> "CODEARTICLE" And OperationReferenceValideCegide(operationTexte) Then
                Set info = CreateObject("Scripting.Dictionary")
                info.CompareMode = vbTextCompare
                info.Add "Operation", operationTexte
                info.Add "OperationNorm", operationNorm
                info.Add "Code", codeArticle
                info.Add "TypeFournisseur", Trim$(CStr(ws.Cells(ligne, 1).Value2))
                info.Add "Article", Trim$(CStr(ws.Cells(ligne, 2).Value2))
                info.Add "ModeRefacturation", Trim$(CStr(ws.Cells(ligne, 4).Value2))
                info.Add "Cle", Trim$(CStr(ws.Cells(ligne, 5).Value2))
                info.Add "LigneSource", ligne
                resultat.Add info
                If Not baseOperationsExactesCegide.Exists(operationNorm) Then
                    Set candidats = New Collection
                    baseOperationsExactesCegide.Add operationNorm, candidats
                Else
                    Set candidats = baseOperationsExactesCegide(operationNorm)
                End If
                candidats.Add info
                IndexerPrefixesOperationCegide info
            End If
        End If
    Next ligne
    Set ChargerArticlesCegide = resultat
End Function

Private Function OperationReferenceValideCegide(ByVal operationTexte As String) As Boolean
    Dim operationNorm As String
    Dim caractere As String
    Dim nombreLettres As Long
    Dim i As Long

    operationNorm = NormaliserCompactCegide(operationTexte)
    If Len(operationNorm) < 3 Then Exit Function

    For i = 1 To Len(operationNorm)
        caractere = Mid$(operationNorm, i, 1)
        If caractere Like "[A-Z]" Then nombreLettres = nombreLettres + 1
        If nombreLettres >= 3 Then
            OperationReferenceValideCegide = True
            Exit Function
        End If
    Next i
End Function

Private Sub IndexerPrefixesOperationCegide(ByVal article As Object)
    Dim mots As Variant
    Dim mot As Variant
    Dim motTexte As String
    Dim prefixe As String
    Dim identifiant As String
    Dim candidats As Object

    identifiant = CStr(article("LigneSource"))
    mots = Split(NormaliserEspacesCegide(CStr(article("Operation"))), " ")
    For Each mot In mots
        motTexte = Trim$(CStr(mot))
        If Len(motTexte) >= 4 And Left$(motTexte, 1) Like "[A-Z]" And Not EstMotNeutreOperationCegide(motTexte) Then
            prefixe = Left$(motTexte, 3)
            If Not baseOperationsPrefixeCegide.Exists(prefixe) Then
                Set candidats = CreateObject("Scripting.Dictionary")
                candidats.CompareMode = vbTextCompare
                baseOperationsPrefixeCegide.Add prefixe, candidats
            Else
                Set candidats = baseOperationsPrefixeCegide(prefixe)
            End If
            If Not candidats.Exists(identifiant) Then candidats.Add identifiant, article
        End If
    Next mot
End Sub

Private Function ChargerOrganisationsCegide(ByVal ws As Worksheet) As Collection
    Dim resultat As New Collection
    Dim derniereLigne As Long
    Dim ligne As Long
    Dim nom As String
    Dim codeOrganisation As String
    Dim destination As String
    Dim info As Object

    derniereLigne = ws.Cells(ws.Rows.Count, 15).End(xlUp).Row
    For ligne = 1 To derniereLigne
        nom = Trim$(CStr(ws.Cells(ligne, 15).Value2))
        codeOrganisation = Trim$(CStr(ws.Cells(ligne, 13).Value2))
        destination = Trim$(CStr(ws.Cells(ligne, 14).Value2))
        If nom <> "" And NormaliserCompactCegide(nom) <> "SIGNATAIRE" Then
            If InStr(1, NormaliserCompactCegide(nom), "REDACTEUR", vbTextCompare) = 0 Then
                Set info = CreateObject("Scripting.Dictionary")
                info.CompareMode = vbTextCompare
                info.Add "Nom", nom
                info.Add "NomNorm", NormaliserEspacesCegide(nom)
                info.Add "Alias", AliasRedacteurCegide(nom)
                info.Add "CodeOrganisation", codeOrganisation
                info.Add "Destination", destination
                info.Add "Service", Trim$(CStr(ws.Cells(ligne, 16).Value2))
                info.Add "Identifiant", Trim$(CStr(ws.Cells(ligne, 18).Value2))
                resultat.Add info
            End If
        End If
    Next ligne
    Set ChargerOrganisationsCegide = resultat
End Function

Private Function AnalyserTableauFactureFlexibleCegide(ByVal texteOCR As String, ByVal articles As Collection, ByRef probleme As String) As Collection
    Dim resultat As New Collection
    Dim lignes As Variant
    Dim ligneBrute As String
    Dim ligneAnalyse As String
    Dim ligneNorm As String
    Dim segmentOperation As String
    Dim segmentHT As String
    Dim segmentDebours As String
    Dim article As Object
    Dim articleDebours As Object
    Dim info As Object
    Dim i As Long
    Dim ligneEntete As Long
    Dim pNature As Long
    Dim pHT As Long
    Dim pTVA As Long
    Dim pDebours As Long
    Dim pTTC As Long
    Dim montantHT As Double
    Dim montantDebours As Double
    Dim montantHTTrouve As Boolean
    Dim montantDeboursTrouve As Boolean
    Dim montantOperation As Double
    Dim montantOperationTrouve As Boolean
    Dim codeDeboursOperation As String
    Dim ligneCandidateOperation As Boolean
    Dim ambigu As Boolean
    Dim enteteTrouvee As Boolean

    texteOCR = SeparerLignesMultidatesCegide(texteOCR)
    texteOCR = Replace(Replace(texteOCR, vbCrLf, vbLf), vbCr, vbLf)
    lignes = Split(texteOCR, vbLf)
    ligneEntete = LBound(lignes) - 1

    For i = LBound(lignes) To UBound(lignes)
        ligneBrute = CStr(lignes(i))
        If EstEnteteTableauFlexibleCegide(ligneBrute) Then
            enteteTrouvee = True
            ligneEntete = i
            pNature = PositionPremierMotifCegide(ligneBrute, Array("NATURE[ ]*(DE|D')?[ ]*L?[ ]*OPERATION", "LIBELLE", "DESIGNATION", "DESCRIPTION", "PRESTATION", "ACTE", "DETAIL", "OPERATION"))
            pHT = PositionPremierMotifCegide(ligneBrute, Array("HORS[ ]*TAXE", "MONTANT[ ]*H[ .]*T", "PRIX[ ]*H[ .]*T", "(^|[^A-Z])H[ .]*T([^A-Z]|$)", "MONTANT", "PRIX"))
            pTVA = PositionMotifCegide(ligneBrute, "T[ .]*V[ .]*A")
            pDebours = PositionMotifCegide(ligneBrute, "DEBOURS")
            pTTC = PositionMotifCegide(ligneBrute, "T[ .]*T[ .]*C")
            Exit For
        End If
    Next i

    For i = ligneEntete + 1 To UBound(lignes)
        If i Mod 20 = 0 Then DoEvents
        ligneBrute = CStr(lignes(i))
        ligneNorm = NormaliserCompactCegide(ligneBrute)
        If EstFinTableauFlexibleCegide(ligneNorm) And Not LigneContientDateCegide(ligneBrute) Then GoTo LigneSuivanteFlexible

        If Trim$(ligneBrute) <> "" And Not EstEnteteTableauFlexibleCegide(ligneBrute) Then
            ligneAnalyse = ligneBrute
            segmentOperation = ""
            segmentHT = ""
            segmentDebours = ""
            If pNature > 0 And pHT > pNature Then segmentOperation = Trim$(Mid$(ligneBrute, pNature, pHT - pNature))
            If pHT > 0 And pTVA > pHT Then segmentHT = Mid$(ligneBrute, pHT, pTVA - pHT)
            If pDebours > 0 And pTTC > pDebours Then segmentDebours = Mid$(ligneBrute, pDebours, pTTC - pDebours)

            montantHT = LireMontantSegmentCegide(segmentHT, montantHTTrouve)
            montantDebours = LireMontantSegmentCegide(segmentDebours, montantDeboursTrouve)
            ligneCandidateOperation = LigneContientDateCegide(ligneBrute) _
                Or montantHTTrouve _
                Or montantDeboursTrouve _
                Or OperationReferenceValideCegide(segmentOperation)
            If Not ligneCandidateOperation Then GoTo LigneSuivanteFlexible

            ambigu = False
            Set article = Nothing
            Set articleDebours = Nothing

            If InStr(1, ligneNorm, "DEBOUR", vbTextCompare) > 0 Then
                Set articleDebours = TrouverArticleDeboursGeneriqueCegide(articles)
                If Not articleDebours Is Nothing Then
                    montantOperation = montantDebours
                    montantOperationTrouve = montantDeboursTrouve
                    If Not montantOperationTrouve Then
                        montantOperation = montantHT
                        montantOperationTrouve = montantHTTrouve
                    End If

                    Set info = CreateObject("Scripting.Dictionary")
                    info.CompareMode = vbTextCompare
                    info.Add "OperationReference", CStr(articleDebours("Operation"))
                    info.Add "TexteDetecte", Trim$(ligneBrute)
                    info.Add "Code", CStr(articleDebours("Code"))
                    info.Add "TypeFournisseur", CStr(articleDebours("TypeFournisseur"))
                    info.Add "ArticleReference", CStr(articleDebours("Article"))
                    info.Add "MontantHT", 0#
                    info.Add "Debours", montantOperation
                    info.Add "MontantTrouve", montantOperationTrouve
                    info.Add "DeboursSeulement", True
                    resultat.Add info
                    If Not montantOperationTrouve Then probleme = AjouterProblemeCegide(probleme, "Montant introuvable pour l'operation de debours.")
                End If
            Else
                Set article = TrouverArticleCegide(ligneAnalyse, segmentOperation, articles, ambigu)
                If article Is Nothing And i < UBound(lignes) Then
                    ligneAnalyse = ligneBrute & " " & CStr(lignes(i + 1))
                    ambigu = False
                    Set article = TrouverArticleCegide(ligneAnalyse, segmentOperation, articles, ambigu)
                End If
                If Not article Is Nothing Then
                    Set info = CreateObject("Scripting.Dictionary")
                    info.CompareMode = vbTextCompare
                    info.Add "OperationReference", CStr(article("Operation"))
                    If Trim$(segmentOperation) <> "" Then
                        info.Add "TexteDetecte", Trim$(segmentOperation)
                    Else
                        info.Add "TexteDetecte", CStr(article("Operation"))
                    End If
                    info.Add "Code", CStr(article("Code"))
                    info.Add "TypeFournisseur", CStr(article("TypeFournisseur"))
                    info.Add "ArticleReference", CStr(article("Article"))
                    If montantHTTrouve Then
                        info.Add "MontantHT", montantHT
                        info.Add "MontantTrouve", True
                    ElseIf montantDeboursTrouve Then
                        info.Add "MontantHT", montantDebours
                        info.Add "MontantTrouve", True
                        codeDeboursOperation = TrouverCodeDeboursCegide(articles, info)
                        If codeDeboursOperation <> "" Then
                            info("Code") = codeDeboursOperation
                        Else
                            probleme = AjouterProblemeCegide(probleme, "Code article de debours terminant par 2 introuvable pour : " & CStr(article("Operation")))
                        End If
                    Else
                        info.Add "MontantHT", 0#
                        info.Add "MontantTrouve", False
                    End If
                    info.Add "Debours", 0#
                    info.Add "DeboursSeulement", False
                    resultat.Add info
                    If Not montantHTTrouve And Not montantDeboursTrouve Then probleme = AjouterProblemeCegide(probleme, "Montant introuvable pour l'operation : " & CStr(article("Operation")))
                End If
            End If
        End If
LigneSuivanteFlexible:
    Next i
    Set AnalyserTableauFactureFlexibleCegide = resultat
End Function

Private Function EstEnteteTableauFlexibleCegide(ByVal texte As String) As Boolean
    Dim normalise As String
    Dim descriptionTrouvee As Boolean
    Dim montantTrouve As Boolean

    normalise = NormaliserCompactCegide(texte)
    If InStr(1, normalise, "NATUREOPERATION", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "OPERATION", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "LIBELLE", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "DESIGNATION", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "DESCRIPTION", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "PRESTATION", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "ACTE", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "DETAIL", vbTextCompare) > 0 Then descriptionTrouvee = True
    If InStr(1, normalise, "HORSTAXE", vbTextCompare) > 0 Then montantTrouve = True
    If InStr(1, normalise, "MONTANTHT", vbTextCompare) > 0 Then montantTrouve = True
    If PositionMotifCegide(texte, "(^|[^A-Z])H[ .]*T([^A-Z]|$)") > 0 Then montantTrouve = True
    If InStr(1, normalise, "MONTANT", vbTextCompare) > 0 Then montantTrouve = True
    If InStr(1, normalise, "PRIX", vbTextCompare) > 0 Then montantTrouve = True
    If InStr(1, normalise, "DEBOUR", vbTextCompare) > 0 Then montantTrouve = True
    If InStr(1, normalise, "TTC", vbTextCompare) > 0 Then montantTrouve = True
    EstEnteteTableauFlexibleCegide = descriptionTrouvee And montantTrouve
End Function

Private Function EstFinTableauFlexibleCegide(ByVal texteNormalise As String) As Boolean
    If InStr(1, texteNormalise, "TOTAL", vbTextCompare) > 0 Then EstFinTableauFlexibleCegide = True
    If InStr(1, texteNormalise, "SOUSTOTAL", vbTextCompare) > 0 Then EstFinTableauFlexibleCegide = True
    If InStr(1, texteNormalise, "NETAPAYER", vbTextCompare) > 0 Then EstFinTableauFlexibleCegide = True
    If InStr(1, texteNormalise, "TAUXTVA", vbTextCompare) > 0 Then EstFinTableauFlexibleCegide = True
    If InStr(1, texteNormalise, "ACOMPTES", vbTextCompare) > 0 Then EstFinTableauFlexibleCegide = True
End Function

Private Function PositionPremierMotifCegide(ByVal texte As String, ByVal motifs As Variant) As Long
    Dim motif As Variant
    Dim position As Long
    Dim meilleurePosition As Long
    For Each motif In motifs
        position = PositionMotifCegide(texte, CStr(motif))
        If position > 0 Then
            If meilleurePosition = 0 Or position < meilleurePosition Then meilleurePosition = position
        End If
    Next motif
    PositionPremierMotifCegide = meilleurePosition
End Function

Private Function TrouverArticleDeboursGeneriqueCegide(ByVal articles As Collection) As Object
    Dim article As Object
    Dim code As String
    For Each article In articles
        code = Trim$(CStr(article("Code")))
        If Right$(code, 1) = "2" Then
            If InStr(1, CStr(article("OperationNorm")), "DEBOUR", vbTextCompare) > 0 Or InStr(1, NormaliserCompactCegide(CStr(article("Article"))), "DEBOUR", vbTextCompare) > 0 Then
                Set TrouverArticleDeboursGeneriqueCegide = article
                Exit Function
            End If
        End If
    Next article
End Function

Private Function AnalyserTableauFactureCegide(ByVal texteOCR As String, ByVal articles As Collection, ByRef probleme As String) As Collection
    Dim resultat As New Collection
    Dim lignes As Variant
    Dim ligneBrute As String
    Dim ligneNorm As String
    Dim segmentOperation As String
    Dim segmentHT As String
    Dim segmentDebours As String
    Dim montants As Collection
    Dim article As Object
    Dim info As Object
    Dim i As Long
    Dim ligneEntete As Long
    Dim pNature As Long
    Dim pHT As Long
    Dim pTVA As Long
    Dim pDebours As Long
    Dim pTTC As Long
    Dim montantHT As Double
    Dim montantDebours As Double
    Dim montantHTTrouve As Boolean
    Dim montantDeboursTrouve As Boolean
    Dim ambigu As Boolean

    texteOCR = Replace(Replace(texteOCR, vbCrLf, vbLf), vbCr, vbLf)
    lignes = Split(texteOCR, vbLf)
    ligneEntete = LBound(lignes) - 1

    For i = LBound(lignes) To UBound(lignes)
        ligneBrute = CStr(lignes(i))
        ligneNorm = NormaliserCompactCegide(ligneBrute)
        If InStr(1, ligneNorm, "NATURE", vbTextCompare) > 0 Then
            If InStr(1, ligneNorm, "OPERATION", vbTextCompare) > 0 Then
                If InStr(1, ligneNorm, "HORSTAXE", vbTextCompare) > 0 Then
                    ligneEntete = i
                    pNature = PositionMotifCegide(ligneBrute, "NATURE")
                    pHT = PositionMotifCegide(ligneBrute, "HORS[ ]*TAXE")
                    pTVA = PositionMotifCegide(ligneBrute, "T[ .]*V[ .]*A")
                    pDebours = PositionMotifCegide(ligneBrute, "DEBOURS")
                    pTTC = PositionMotifCegide(ligneBrute, "T[ .]*T[ .]*C")
                    Exit For
                End If
            End If
        End If
    Next i

    For i = ligneEntete + 1 To UBound(lignes)
        ligneBrute = CStr(lignes(i))
        ligneNorm = NormaliserCompactCegide(ligneBrute)
        If resultat.Count > 0 Then
            If InStr(1, ligneNorm, "TAUXTVA", vbTextCompare) > 0 Then Exit For
            If InStr(1, ligneNorm, "TOTALTTC", vbTextCompare) > 0 Then Exit For
            If InStr(1, ligneNorm, "ACOMPTES", vbTextCompare) > 0 Then Exit For
        End If

        If LigneContientDateCegide(ligneBrute) Then
            segmentOperation = ""
            segmentHT = ""
            segmentDebours = ""
            If pNature > 0 And pHT > pNature Then segmentOperation = Trim$(Mid$(ligneBrute, pNature, pHT - pNature))
            If pHT > 0 And pTVA > pHT Then segmentHT = Mid$(ligneBrute, pHT, pTVA - pHT)
            If pDebours > 0 And pTTC > pDebours Then segmentDebours = Mid$(ligneBrute, pDebours, pTTC - pDebours)

            ambigu = False
            Set article = TrouverArticleCegide(ligneBrute, segmentOperation, articles, ambigu)
            If article Is Nothing Or ambigu Then
                probleme = AjouterProblemeCegide(probleme, "Operation non reconnue ou ambigue : " & Trim$(ligneBrute))
            Else
                montantHT = LireMontantSegmentCegide(segmentHT, montantHTTrouve)
                montantDebours = LireMontantSegmentCegide(segmentDebours, montantDeboursTrouve)

                If Not montantHTTrouve Then
                    probleme = AjouterProblemeCegide(probleme, "Montant hors taxes introuvable pour " & CStr(article("Operation")))
                Else
                    Set info = CreateObject("Scripting.Dictionary")
                    info.CompareMode = vbTextCompare
                    info.Add "OperationReference", CStr(article("Operation"))
                    info.Add "Code", CStr(article("Code"))
                    info.Add "TypeFournisseur", CStr(article("TypeFournisseur"))
                    info.Add "ArticleReference", CStr(article("Article"))
                    info.Add "MontantHT", montantHT
                    info.Add "MontantTrouve", True
                    info.Add "Debours", montantDebours
                    resultat.Add info
                End If
            End If
        End If
    Next i
    Set AnalyserTableauFactureCegide = resultat
End Function

Private Function PositionMotifCegide(ByVal texte As String, ByVal motif As String) As Long
    Dim reg As Object
    Dim correspondances As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = False
    reg.IgnoreCase = True
    reg.Pattern = motif
    Set correspondances = reg.Execute(UCase$(RemplacerAccentsCegide(texte)))
    If correspondances.Count > 0 Then PositionMotifCegide = CLng(correspondances(0).FirstIndex) + 1
End Function

Private Function LigneContientDateCegide(ByVal texte As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = False
    reg.IgnoreCase = True
    reg.Pattern = "(^|[^0-9])[0-3]?[0-9][./-][01]?[0-9][./-](20)?[0-9]{2}([^0-9]|$)"
    LigneContientDateCegide = reg.Test(texte)
End Function

Private Function SeparerLignesMultidatesCegide(ByVal texte As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.Pattern = "([0-3]?[0-9][./-][01]?[0-9][./-](20)?[0-9]{2})"
    SeparerLignesMultidatesCegide = reg.Replace(texte, vbLf & "$1")
End Function

Private Function SupprimerDatesCegide(ByVal texte As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.Pattern = "(^|[^0-9])[0-3]?[0-9][./-][01]?[0-9][./-](20)?[0-9]{2}([^0-9]|$)"
    SupprimerDatesCegide = reg.Replace(texte, " ")
End Function

Private Function ExtraireMontantsCegide(ByVal texte As String) As Collection
    Dim resultat As New Collection
    Dim reg As Object
    Dim correspondances As Object
    Dim correspondance As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.Pattern = "-?[0-9]+([ .][0-9]{3})*[,.][0-9]{2}"
    Set correspondances = reg.Execute(texte)
    For Each correspondance In correspondances
        resultat.Add ConvertirMontantCegide(CStr(correspondance.Value))
    Next correspondance
    Set ExtraireMontantsCegide = resultat
End Function

Private Function LireMontantSegmentCegide(ByVal segment As String, ByRef trouve As Boolean) As Double
    Dim montants As Collection

    trouve = False
    If Trim$(segment) = "" Then Exit Function

    Set montants = ExtraireMontantsCegide(segment)
    If montants.Count > 0 Then
        LireMontantSegmentCegide = CDbl(montants(1))
        trouve = True
    End If
End Function

Private Function ConvertirMontantCegide(ByVal texte As String) As Double
    Dim valeur As String
    Dim dernierePosition As Long
    valeur = Replace(Replace(Trim$(texte), " ", ""), ChrW$(160), "")
    If InStr(1, valeur, ",", vbBinaryCompare) > 0 Then
        valeur = Replace(valeur, ".", "")
        valeur = Replace(valeur, ",", ".")
    ElseIf Len(valeur) - Len(Replace(valeur, ".", "")) > 1 Then
        dernierePosition = InStrRev(valeur, ".")
        valeur = Replace(Left$(valeur, dernierePosition - 1), ".", "") & Mid$(valeur, dernierePosition)
    End If
    ConvertirMontantCegide = Val(valeur)
End Function

Private Function TrouverArticleCegide(ByVal ligneComplete As String, ByVal operationLue As String, ByVal articles As Collection, ByRef ambigu As Boolean) As Object
    Dim article As Object
    Dim meilleur As Object
    Dim cible As String
    Dim ligneNorm As String
    Dim operationNorm As String
    Dim score As Long
    Dim meilleurScore As Long
    Dim candidatsRapides As Object
    Dim identifiant As Variant

    cible = NormaliserCompactCegide(operationLue)
    ligneNorm = NormaliserCompactCegide(ligneComplete)
    If cible <> "" Then
        Set meilleur = TrouverArticleExactBaseCegide(cible, ambigu)
        If Not meilleur Is Nothing Then
            Set TrouverArticleCegide = meilleur
            Exit Function
        End If
    End If
    For Each article In articles
        operationNorm = CStr(article("OperationNorm"))
        score = 0
        If cible <> "" And cible = operationNorm Then
            score = 10000 + Len(operationNorm)
        ElseIf cible <> "" And InStr(1, cible, operationNorm, vbTextCompare) > 0 Then
            score = 9000 + Len(operationNorm)
        ElseIf cible <> "" And Len(cible) >= 5 And InStr(1, operationNorm, cible, vbTextCompare) > 0 Then
            score = 8000 + Len(cible)
        ElseIf InStr(1, ligneNorm, operationNorm, vbTextCompare) > 0 Then
            score = 7000 + Len(operationNorm)
        End If

        ' Une abreviation courte ne doit pas correspondre a un fragment d'un autre mot.
        If score > 0 And Len(operationNorm) <= 3 Then
            If InStr(1, " " & NormaliserEspacesCegide(ligneComplete) & " ", " " & operationNorm & " ", vbTextCompare) = 0 Then score = 0
        End If
        If score > 0 Then score = score + ScoreArticleNormalCegide(article)

        If score > meilleurScore Then
            meilleurScore = score
            Set meilleur = article
            ambigu = False
        ElseIf score > 0 And score = meilleurScore Then
            If Not meilleur Is Nothing Then
                If CStr(meilleur("Code")) <> CStr(article("Code")) Then ambigu = True
            End If
        End If
    Next article
    If meilleurScore > 0 Then
        Set TrouverArticleCegide = meilleur
        Exit Function
    End If

    Set meilleur = Nothing
    ambigu = False
    Set candidatsRapides = CandidatsAbreviationRapideCegide(ligneComplete)
    For Each identifiant In candidatsRapides.Keys
        Set article = candidatsRapides(identifiant)
        score = ScoreAbreviationRapideCegide(ligneComplete, CStr(article("Operation")))
        If score > 0 Then score = score + ScoreArticleNormalCegide(article)
        If score > meilleurScore Then
            meilleurScore = score
            Set meilleur = article
            ambigu = False
        ElseIf score > 0 And score = meilleurScore Then
            If Not meilleur Is Nothing Then
                If CStr(meilleur("Code")) <> CStr(article("Code")) Then ambigu = True
            End If
        End If
    Next identifiant
    If meilleurScore > 0 Then Set TrouverArticleCegide = meilleur
End Function

Private Function TrouverArticleExactBaseCegide(ByVal operationNorm As String, ByRef ambigu As Boolean) As Object
    Dim candidats As Collection
    Dim article As Object
    Dim meilleur As Object
    Dim score As Long
    Dim meilleurScore As Long

    If baseOperationsExactesCegide Is Nothing Then Exit Function
    If Not baseOperationsExactesCegide.Exists(operationNorm) Then Exit Function
    Set candidats = baseOperationsExactesCegide(operationNorm)
    For Each article In candidats
        score = 1000 + ScoreArticleNormalCegide(article)
        If score > meilleurScore Then
            meilleurScore = score
            Set meilleur = article
            ambigu = False
        ElseIf score = meilleurScore Then
            If Not meilleur Is Nothing Then
                If CStr(meilleur("Code")) <> CStr(article("Code")) Then ambigu = True
            End If
        End If
    Next article
    Set TrouverArticleExactBaseCegide = meilleur
End Function

Private Function ScoreArticleNormalCegide(ByVal article As Object) As Long
    Dim typeNorm As String
    Dim codeArticle As String

    typeNorm = NormaliserCompactCegide(CStr(article("TypeFournisseur")))
    codeArticle = Trim$(CStr(article("Code")))
    If InStr(1, typeNorm, "20", vbTextCompare) > 0 Or Right$(codeArticle, 1) = "2" Then
        ScoreArticleNormalCegide = -300
    Else
        ScoreArticleNormalCegide = 300
    End If
End Function

Private Function CandidatsAbreviationRapideCegide(ByVal ligneTexte As String) As Object
    Dim resultat As Object
    Dim candidats As Object
    Dim mots As Variant
    Dim mot As Variant
    Dim motTexte As String
    Dim prefixe As String
    Dim identifiant As Variant

    Set resultat = CreateObject("Scripting.Dictionary")
    resultat.CompareMode = vbTextCompare
    If baseOperationsPrefixeCegide Is Nothing Then
        Set CandidatsAbreviationRapideCegide = resultat
        Exit Function
    End If
    mots = Split(NormaliserEspacesCegide(ligneTexte), " ")
    For Each mot In mots
        motTexte = Trim$(CStr(mot))
        If Len(motTexte) >= 3 And Left$(motTexte, 1) Like "[A-Z]" And Not EstMotNeutreOperationCegide(motTexte) Then
            prefixe = Left$(motTexte, 3)
            If baseOperationsPrefixeCegide.Exists(prefixe) Then
                Set candidats = baseOperationsPrefixeCegide(prefixe)
                For Each identifiant In candidats.Keys
                    If Not resultat.Exists(CStr(identifiant)) Then resultat.Add CStr(identifiant), candidats(identifiant)
                Next identifiant
            End If
        End If
    Next mot
    Set CandidatsAbreviationRapideCegide = resultat
End Function

Private Function ScoreAbreviationRapideCegide(ByVal ligneTexte As String, ByVal operationTexte As String) As Long
    Dim motsLigne As Variant
    Dim motsOperation As Variant
    Dim motLigne As Variant
    Dim motOperation As Variant
    Dim ligneMot As String
    Dim operationMot As String
    Dim trouves As Long
    Dim exacts As Long
    Dim dejaTrouve As Boolean

    If Len(ligneTexte) > 220 Then Exit Function
    motsLigne = Split(NormaliserEspacesCegide(ligneTexte), " ")
    motsOperation = Split(NormaliserEspacesCegide(operationTexte), " ")
    For Each motOperation In motsOperation
        operationMot = Trim$(CStr(motOperation))
        If Len(operationMot) >= 4 And Left$(operationMot, 1) Like "[A-Z]" And Not EstMotNeutreOperationCegide(operationMot) Then
            dejaTrouve = False
            For Each motLigne In motsLigne
                ligneMot = Trim$(CStr(motLigne))
                If Len(ligneMot) >= 3 And Left$(ligneMot, 1) Like "[A-Z]" And Not EstMotNeutreOperationCegide(ligneMot) Then
                    If ligneMot = operationMot Then
                        trouves = trouves + 1
                        exacts = exacts + 1
                        dejaTrouve = True
                    ElseIf Len(ligneMot) < Len(operationMot) And Left$(operationMot, Len(ligneMot)) = ligneMot Then
                        trouves = trouves + 1
                        dejaTrouve = True
                    ElseIf MotsOCRProchesRapideCegide(ligneMot, operationMot) Then
                        trouves = trouves + 1
                        dejaTrouve = True
                    End If
                End If
                If dejaTrouve Then Exit For
            Next motLigne
        End If
    Next motOperation
    If trouves > 0 Then ScoreAbreviationRapideCegide = 6000 + trouves * 100 + exacts * 20
End Function

Private Function MotsOCRProchesRapideCegide(ByVal motOCR As String, ByVal motReference As String) As Boolean
    Dim longueurMin As Long
    Dim differences As Long
    Dim i As Long

    If Len(motOCR) < 5 Or Len(motReference) < 5 Then Exit Function
    If Left$(motOCR, 3) <> Left$(motReference, 3) Then Exit Function
    If Abs(Len(motOCR) - Len(motReference)) > 2 Then Exit Function

    longueurMin = Application.Min(Len(motOCR), Len(motReference))
    differences = Abs(Len(motOCR) - Len(motReference))
    For i = 4 To longueurMin
        If Mid$(motOCR, i, 1) <> Mid$(motReference, i, 1) Then
            differences = differences + 1
            If differences > 2 Then Exit Function
        End If
    Next i
    MotsOCRProchesRapideCegide = True
End Function

Private Function EstMotNeutreOperationCegide(ByVal mot As String) As Boolean
    Select Case UCase$(Trim$(mot))
        Case "DE", "DU", "DES", "LA", "LE", "LES", "D", "L", "ET", "OU", "A", "AU", "AUX", "EN", "UN", "UNE", "POUR", "SUR", "PAR", "AVEC"
            EstMotNeutreOperationCegide = True
    End Select
End Function

Private Function TrouverRedacteurCegide(ByVal texteOCR As String, ByVal organisations As Collection, Optional ByVal mots As Collection = Nothing, Optional ByVal moduleFacture As String = "") As Object
    Dim texteEspace As String
    Dim texteCompact As String
    Dim info As Object
    Dim meilleur As Object
    Dim score As Long
    Dim meilleurScore As Long
    Dim ambigu As Boolean

    texteEspace = " " & NormaliserEspacesCegide(texteOCR) & " "
    texteCompact = NormaliserCompactCegide(texteOCR)
    For Each info In organisations
        score = ScoreRedacteurCegide(CStr(info("Nom")), texteEspace, texteCompact)
        If score > 0 Then
            If Not NomProcheReferencesCegide(CStr(info("Nom")), mots, moduleFacture) Then score = 0
        End If
        If score > meilleurScore Then
            meilleurScore = score
            Set meilleur = info
            ambigu = False
        ElseIf score > 0 And score = meilleurScore Then
            If Not meilleur Is Nothing Then
                If CStr(meilleur("Nom")) <> CStr(info("Nom")) Then ambigu = True
                If CStr(meilleur("CodeOrganisation")) <> CStr(info("CodeOrganisation")) Then ambigu = True
                If CStr(meilleur("Destination")) <> CStr(info("Destination")) Then ambigu = True
            End If
        End If
    Next info
    If meilleurScore > 0 And Not ambigu Then Set TrouverRedacteurCegide = meilleur
End Function

Private Function RedacteurServiceCegide(ByVal texteOCR As String, ByVal organisations As Collection) As Object
    Dim info As Object, service As String, texte As String
    texte = " " & NormaliserEspacesCegide(texteOCR) & " "
    ' La collection conserve l'ordre des lignes de Code organisation.
    For Each info In organisations
        If info.Exists("Service") Then
            service = NormaliserEspacesCegide(CStr(info("Service")))
            If Len(service) >= 2 Then
                If InStr(1, texte, " " & service & " ", vbTextCompare) > 0 Then
                    Set RedacteurServiceCegide = info
                    Exit Function
                End If
            End If
        End If
    Next info
End Function

Private Function ScoreRedacteurCegide(ByVal nomRedacteur As String, ByVal texteEspace As String, ByVal texteCompact As String) As Long
    Dim nomEspace As String
    Dim nomCompact As String
    Dim mots As Variant
    Dim mot As Variant
    Dim motTexte As String
    Dim motsTrouves As Long
    Dim longueurTrouvee As Long
    Dim vus As Object

    nomEspace = NormaliserEspacesCegide(nomRedacteur)
    nomCompact = Replace(nomEspace, " ", "")
    If nomCompact <> "" And InStr(1, texteCompact, nomCompact, vbTextCompare) > 0 Then
        ScoreRedacteurCegide = 100000 + Len(nomCompact)
        Exit Function
    End If

    Set vus = CreateObject("Scripting.Dictionary")
    vus.CompareMode = vbTextCompare
    mots = Split(nomEspace, " ")
    For Each mot In mots
        motTexte = Trim$(CStr(mot))
        If Len(motTexte) >= 3 And Not EstCodeServiceCegide(motTexte) Then
            If Not vus.Exists(motTexte) Then
                vus.Add motTexte, True
                If InStr(1, texteEspace, " " & motTexte & " ", vbTextCompare) > 0 Then
                    motsTrouves = motsTrouves + 1
                    longueurTrouvee = longueurTrouvee + Len(motTexte)
                End If
            End If
        End If
    Next mot
    If motsTrouves > 0 Then ScoreRedacteurCegide = motsTrouves * 10000 + longueurTrouvee
End Function

Private Function TrouverModuleCegide(ByVal texteOCR As String, ByVal nomRedacteur As String) As String
    Dim lignes As Variant
    Dim ligne As String
    Dim ligneNorm As String
    Dim reg As Object
    Dim correspondances As Object
    Dim correspondance As Object
    Dim candidat As String
    Dim score As Long
    Dim meilleurScore As Long
    Dim i As Long

    lignes = Split(Replace(Replace(texteOCR, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.Pattern = "(^|[^0-9])([0-9]{" & CStr(NOMBRE_CHIFFRES_MODULE) & "})([^0-9]|$)"
    For i = LBound(lignes) To UBound(lignes)
        ligne = CStr(lignes(i))
        ligneNorm = NormaliserEspacesCegide(ligne)
        Set correspondances = reg.Execute(ligne)
        For Each correspondance In correspondances
            candidat = CStr(correspondance.SubMatches(1))
            score = 1
            If Left$(candidat, 1) = "0" Then score = score + 20
            If LigneContientNomRedacteurCegide(ligneNorm, nomRedacteur) Then score = score + 200
            If InStr(1, NormaliserCompactCegide(ligne), "VOSREF", vbTextCompare) > 0 Then score = score + 80
            If InStr(1, NormaliserCompactCegide(ligne), "MODULE", vbTextCompare) > 0 Then score = score + 80
            If InStr(1, NormaliserCompactCegide(ligne), "SIRET", vbTextCompare) > 0 Then score = score - 100
            If InStr(1, NormaliserCompactCegide(ligne), "TVA", vbTextCompare) > 0 Then score = score - 100
            If score > meilleurScore Then
                meilleurScore = score
                TrouverModuleCegide = candidat
            End If
        Next correspondance
    Next i
End Function

Private Function LigneContientNomRedacteurCegide(ByVal ligneNormalisee As String, ByVal nomRedacteur As String) As Boolean
    Dim mots As Variant
    Dim mot As Variant
    Dim motTexte As String
    Dim ligneEntouree As String

    If Trim$(nomRedacteur) = "" Then Exit Function
    ligneEntouree = " " & NormaliserEspacesCegide(ligneNormalisee) & " "
    mots = Split(NormaliserEspacesCegide(nomRedacteur), " ")
    For Each mot In mots
        motTexte = Trim$(CStr(mot))
        If Len(motTexte) >= 3 And Not EstCodeServiceCegide(motTexte) Then
            If InStr(1, ligneEntouree, " " & motTexte & " ", vbTextCompare) > 0 Then
                LigneContientNomRedacteurCegide = True
                Exit Function
            End If
        End If
    Next mot
End Function

Private Function VerifierCodesDeboursCegide(ByVal lignesFacture As Collection, ByVal articles As Collection) As String
    Dim ligneFacture As Object
    Dim codeDebours As String
    For Each ligneFacture In lignesFacture
        If CDbl(ligneFacture("Debours")) <> 0 Then
            codeDebours = TrouverCodeDeboursCegide(articles, ligneFacture)
            If codeDebours = "" Then
                VerifierCodesDeboursCegide = AjouterProblemeCegide(VerifierCodesDeboursCegide, "Code de debours terminant par 2 introuvable pour " & CStr(ligneFacture("OperationReference")))
            ElseIf ligneFacture.Exists("CodeDebours") Then
                ligneFacture("CodeDebours") = codeDebours
            Else
                ligneFacture.Add "CodeDebours", codeDebours
            End If
        End If
    Next ligneFacture
End Function

Private Function TrouverCodeDeboursCegide(ByVal articles As Collection, ByVal ligneFacture As Object) As String
    Dim article As Object
    Dim code As String
    Dim operationReference As String
    Dim typeFournisseur As String
    Dim articleReference As String
    ' Les articles d'origine en 1 restent en 1, meme pour une ligne NS/Exo.
    If ligneFacture.Exists("Code") Then
        code = Trim$(CStr(ligneFacture("Code")))
        If Right$(code, 1) = "1" Then
            TrouverCodeDeboursCegide = code
            Exit Function
        End If
        If Right$(code, 1) <> "3" And Right$(code, 1) <> "2" Then Exit Function
    Else
        Exit Function
    End If
    operationReference = NormaliserCompactCegide(CStr(ligneFacture("OperationReference")))
    typeFournisseur = NormaliserCompactCegide(CStr(ligneFacture("TypeFournisseur")))
    articleReference = NormaliserCompactCegide(CStr(ligneFacture("ArticleReference")))

    For Each article In articles
        code = Trim$(CStr(article("Code")))
        If Right$(code, 1) = "2" And CStr(article("OperationNorm")) = operationReference Then
            TrouverCodeDeboursCegide = code
            Exit Function
        End If
    Next article
    If typeFournisseur = "" And articleReference = "" Then Exit Function
    For Each article In articles
        code = Trim$(CStr(article("Code")))
        If Right$(code, 1) = "2" Then
            If NormaliserCompactCegide(CStr(article("TypeFournisseur"))) = typeFournisseur Then
                If NormaliserCompactCegide(CStr(article("Article"))) = articleReference Then
                    TrouverCodeDeboursCegide = code
                    Exit Function
                End If
            End If
        End If
    Next article
    For Each article In articles
        code = Trim$(CStr(article("Code")))
        If Right$(code, 1) = "2" Then
            If InStr(1, CStr(article("OperationNorm")), "DEBOUR", vbTextCompare) > 0 Or InStr(1, NormaliserCompactCegide(CStr(article("Article"))), "DEBOUR", vbTextCompare) > 0 Then
                TrouverCodeDeboursCegide = code
                Exit Function
            End If
        End If
    Next article
End Function

Private Function CompterLignesSortieCegide(ByVal lignesFacture As Collection) As Long
    CompterLignesSortieCegide = lignesFacture.Count
End Function

Private Function EstDeboursSeulementCegide(ByVal ligneFacture As Object) As Boolean
    On Error Resume Next
    EstDeboursSeulementCegide = CBool(ligneFacture("DeboursSeulement"))
    On Error GoTo 0
End Function

Private Function SaisirIdentifiantCegide(ByVal titre As String, ByVal motif As String, ByVal exemple As String) As String
    Dim saisie As Variant, re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = motif: re.IgnoreCase = True
    Do
        saisie = Application.InputBox(titre & " non detecte. Saisir la valeur de la facture (" & exemple & "). Annuler pour laisser vide.", "Donnee manquante", Type:=2)
        If VarType(saisie) = vbBoolean Then Exit Function
        saisie = UCase$(Trim$(CStr(saisie)))
        If re.Test(saisie) Then
            SaisirIdentifiantCegide = saisie
            Exit Function
        End If
        Gloire_MessageAutomatique "Format attendu : " & exemple, vbExclamation
    Loop
End Function

Private Sub CompleterSaisieCegide(ByRef client As String, ByRef module As String, ByRef redacteur As Object, ByVal organisations As Collection)
    Dim saisie As Variant, nombre As Long
    If client = "" Then client = ClientSaisiCegide(SaisirIdentifiantCegide("Client", "^(L/?)?[0-9]{1,7}$", "12345, L12345 ou L/12345, de 1 a 7 chiffres"))
    If module = "" Then module = SaisirIdentifiantCegide("Module", "^[0-9]{10}$", "0321010063, 10 chiffres")
    If Not redacteur Is Nothing Then Exit Sub
    Do
        saisie = Application.InputBox("Signataire non detecte. Saisir tout ou partie du nom de la colonne O (exemple : NOWA). Annuler pour laisser vide.", "Donnee manquante", Type:=2)
        If VarType(saisie) = vbBoolean Then Exit Sub
        Set redacteur = SignataireSaisiCegide(CStr(saisie), organisations, nombre)
        If Not redacteur Is Nothing Then Exit Sub
        If nombre > 1 Then
            Gloire_MessageAutomatique "Plusieurs correspondances en colonne O. Precisez le nom ou le prenom.", vbExclamation
        Else
            Gloire_MessageAutomatique "Aucune correspondance en colonne O. Saisissez un autre nom.", vbExclamation
        End If
    Loop
End Sub

Private Function ClientSaisiCegide(ByVal valeur As String) As String
    Dim re As Object
    valeur = UCase$(Trim$(valeur))
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(L/?)?([0-9]{1,7})$"
    If re.Test(valeur) Then ClientSaisiCegide = "L/" & re.Execute(valeur)(0).SubMatches(1)
End Function

Private Function SignataireSaisiCegide(ByVal saisie As String, ByVal organisations As Collection, ByRef nombre As Long) As Object
    Dim item As Object, candidat As Object, mot As Variant, correspond As Boolean, nom As String
    nombre = 0
    saisie = NormaliserEspacesCegide(saisie)
    If saisie = "" Then Exit Function
    For Each item In organisations
        nom = NormaliserEspacesCegide(CStr(item("Nom")))
        correspond = True
        For Each mot In Split(saisie, " ")
            If InStr(1, nom, CStr(mot), vbTextCompare) = 0 Then correspond = False: Exit For
        Next mot
        If correspond Then
            nombre = nombre + 1
            Set candidat = item
        End If
    Next item
    If nombre = 1 Then Set SignataireSaisiCegide = candidat
End Function

Private Function ModuleInterditCegide(ByVal ws As Worksheet, ByVal module As String) As Boolean
    Dim valeurs As Variant, i As Long, valeur As String
    If module = "" Then Exit Function
    valeurs = ws.Range("E1:E" & Application.Max(2, ws.Cells(ws.Rows.Count, 5).End(xlUp).Row)).Value2
    For i = 1 To UBound(valeurs, 1)
        If Not IsError(valeurs(i, 1)) Then
            valeur = Trim$(CStr(valeurs(i, 1)))
            If VarType(valeurs(i, 1)) <> vbString And IsNumeric(valeurs(i, 1)) Then valeur = Format$(valeurs(i, 1), "0000000000")
            If valeur = module Then ModuleInterditCegide = True: Exit Function
        End If
    Next i
End Function

Private Function BloquerModuleCegide(ByVal resultat As Worksheet, ByVal module As String, ByVal redacteur As Object) As Boolean
    Dim interdit As Worksheet, export As Worksheet, suivi As Worksheet
    Dim r As Long, entete As Long, etape As Long, numero As String, source As Long, cible As Long, col As Variant
    Set interdit = TrouverFeuilleCegide(resultat.Parent, "GTDTECGR")
    If interdit Is Nothing Then Err.Raise vbObjectError + 870, , "Controle impossible : feuille GTDTECGR absente. Aucun remplissage effectue."
    If Not ModuleInterditCegide(interdit, module) Then Exit Function
    Set export = TrouverFeuilleCegide(resultat.Parent, "EXPORT HUISSIER")
    Set suivi = TrouverFeuilleCegide(resultat.Parent, "MODUP")
    If export Is Nothing Or suivi Is Nothing Then Err.Raise vbObjectError + 871, , "Module non reference : commande annulee. EXPORT HUISSIER ou MODUP absent."
    entete = TrouverLigneEnteteResultatCegide(resultat)
    etape = TrouverColonneEnteteCegide(resultat, entete, "ETAPE")
    If entete < 1 Or etape < 1 Then Err.Raise vbObjectError + 872, , "Module non reference : commande annulee. Ligne de facture introuvable."
    For r = entete + 1 To resultat.Cells(resultat.Rows.Count, 1).End(xlUp).Row
        If Not resultat.Rows(r).Hidden Then
            If LigneResultatDisponibleCegide(resultat, r, etape) Then
                numero = Trim$(CStr(resultat.Cells(r, 1).Value2))
                If numero <> "" Then Exit For
            End If
        End If
    Next r
    If numero = "" Then Err.Raise vbObjectError + 873, , "Module non reference : commande annulee. Numero de facture absent."
    For r = 2 To export.Cells(export.Rows.Count, 1).End(xlUp).Row
        If Trim$(CStr(export.Cells(r, 1).Value2)) = numero Then
            If source > 0 Then Err.Raise vbObjectError + 874, , "Module non reference : commande annulee. Plusieurs lignes EXPORT HUISSIER pour ce numero. MODUP non modifie."
            source = r
        End If
    Next r
    If source = 0 Then Err.Raise vbObjectError + 875, , "Module non reference : commande annulee. Facture absente de EXPORT HUISSIER. MODUP non modifie."
    ' Append after all existing data; never overwrite another invoice or its formulas.
    cible = suivi.UsedRange.Row + suivi.UsedRange.Rows.Count
    If cible < 2 Then cible = 2
    For Each col In Array(1, 2, 3, 13, 14)
        suivi.Cells(cible, CLng(col)).NumberFormat = "@"
    Next col
    suivi.Cells(cible, 1).Value2 = CStr(export.Cells(source, 1).Value2)
    suivi.Cells(cible, 2).Value2 = CStr(export.Cells(source, 2).Value2)
    suivi.Cells(cible, 3).Value2 = CStr(export.Cells(source, 4).Value2)
    If Not redacteur Is Nothing Then suivi.Cells(cible, 13).Value2 = CStr(redacteur("Nom"))
    suivi.Cells(cible, 14).Value2 = module
    suivi.Cells(cible, 15).Value2 = "Module non reference"
    BloquerModuleCegide = True
End Function

Private Sub RemplirResultatCegide(ByVal ws As Worksheet, ByVal lignesFacture As Collection, ByVal redacteur As Object, ByVal moduleFacture As String, ByVal articles As Collection, Optional ByVal clientFacture As String = "")
    Dim nombreLignes As Long
    Dim lignesVisibles As Collection
    Dim ligneFacture As Object
    Dim numero As Variant
    Dim codeDebours As String
    Dim indexSortie As Long
    Dim ligneSortie As Long
    Dim ligneEntete As Long
    Dim colonneEtape As Long

    nombreLignes = CompterLignesSortieCegide(lignesFacture)
    Set lignesVisibles = AjouterLignesSortieCegide(ws, nombreLignes, ligneEntete, numero)
    If lignesVisibles.Count = 0 Then
        Err.Raise vbObjectError + 820, , "Aucune ligne numerotee avec ETAPE vide n'a ete trouvee sous l'entete N (numero)."
    End If
    If lignesVisibles.Count < nombreLignes Then
        Err.Raise vbObjectError + 824, , "Le tableau Remplissage Cegide n'a pas pu etre agrandi pour recevoir toutes les lignes."
    End If
    If Trim$(CStr(numero)) = "" Then Err.Raise vbObjectError + 821, , "La ligne disponible n'a aucun numero sous l'entete N (numero)."
    colonneEtape = TrouverColonneEnteteCegide(ws, ligneEntete, "ETAPE")
    If colonneEtape = 0 Then Err.Raise vbObjectError + 822, , "La colonne Etape est introuvable."

    For Each ligneFacture In lignesFacture
        If EstDeboursSeulementCegide(ligneFacture) Then
            codeDebours = ""
            If ligneFacture.Exists("CodeDebours") Then codeDebours = CStr(ligneFacture("CodeDebours"))
            If codeDebours = "" And ligneFacture.Exists("Code") Then codeDebours = CStr(ligneFacture("Code"))
            indexSortie = indexSortie + 1
            ligneSortie = CLng(lignesVisibles(indexSortie))
            EcrireLigneCegide ws, ligneSortie, numero, TexteDetecteCegide(ligneFacture), codeDebours, CDbl(ligneFacture("Debours")), redacteur, moduleFacture, colonneEtape, MontantTrouveCegide(ligneFacture), clientFacture
        Else
            indexSortie = indexSortie + 1
            ligneSortie = CLng(lignesVisibles(indexSortie))
            EcrireLigneCegide ws, ligneSortie, numero, TexteDetecteCegide(ligneFacture), CStr(ligneFacture("Code")), CDbl(ligneFacture("MontantHT")), redacteur, moduleFacture, colonneEtape, MontantTrouveCegide(ligneFacture), clientFacture
        End If
    Next ligneFacture
    ws.Activate
    Application.Goto ws.Cells(CLng(lignesVisibles(1)), 1), True
End Sub

Private Sub RemplirMetadonneesPartiellesCegide(ByVal ws As Worksheet, ByVal redacteur As Object, ByVal moduleFacture As String, Optional ByVal clientFacture As String = "")
    Dim lignesVisibles As Collection
    Dim ligneEntete As Long
    Dim ligneSortie As Long
    Dim colonneEtape As Long
    Dim numero As Variant

    Set lignesVisibles = AjouterLignesSortieCegide(ws, 1, ligneEntete, numero)
    If lignesVisibles.Count = 0 Then Err.Raise vbObjectError + 820, , "Aucune ligne numerotee avec ETAPE vide n'a ete trouvee sous l'entete N (numero)."
    ligneSortie = CLng(lignesVisibles(1))
    colonneEtape = TrouverColonneEnteteCegide(ws, ligneEntete, "ETAPE")
    If colonneEtape = 0 Then Err.Raise vbObjectError + 822, , "La colonne Etape est introuvable."

    ws.Cells(ligneSortie, 1).Value2 = numero
    ws.Cells(ligneSortie, 2).ClearContents
    ws.Cells(ligneSortie, 4).Value2 = 1
    ws.Cells(ligneSortie, 7).Value2 = "Montant non lu - a verifier"
    EcrireChampsCommunsCegide ws, ligneSortie, redacteur, moduleFacture, clientFacture
    AppliquerValidationEtapeCegide ws.Cells(ligneSortie, colonneEtape), True
    ws.Activate
    Application.Goto ws.Cells(ligneSortie, 1), True
End Sub

Private Function AjouterLignesSortieCegide(ByVal ws As Worksheet, ByVal nombreDemande As Long, ByRef ligneEntete As Long, ByRef numeroMaitre As Variant) As Collection
    Dim resultat As New Collection
    Dim tableau As ListObject
    Dim ligneTableau As Range
    Dim nouvelleLigne As ListRow
    Dim indexTableau As Long
    Dim indexAncrage As Long
    Dim ligne As Long
    Dim ligneAncrage As Long
    Dim derniereLigneNumero As Long
    Dim derniereColonneResultat As Long
    Dim colonneEtape As Long
    Dim decalage As Long
    Dim plageModele As Range
    Dim plageCible As Range
    Set tableau = TrouverTableauResultatCegide(ws)
    If Not tableau Is Nothing Then
        ligneEntete = tableau.HeaderRowRange.Row
        PreparerColonneClientCegide ws, ligneEntete, tableau
        colonneEtape = TrouverColonneEnteteCegide(ws, ligneEntete, "ETAPE")
        If Not tableau.DataBodyRange Is Nothing Then
            For indexTableau = 1 To tableau.ListRows.Count
                Set ligneTableau = tableau.ListRows(indexTableau).Range
                ligne = ligneTableau.Row
                If Not ws.Rows(ligne).Hidden Then
                    If Trim$(CStr(ws.Cells(ligne, 1).Value2)) <> "" Then
                        If LigneResultatDisponibleCegide(ws, ligne, colonneEtape) Then
                            indexAncrage = indexTableau
                            ligneAncrage = ligne
                            numeroMaitre = ws.Cells(ligne, 1).Value2
                            Exit For
                        End If
                    End If
                End If
            Next indexTableau
        End If

        If indexAncrage > 0 Then
            resultat.Add ligneAncrage
            If nombreDemande > 1 Then
                Set plageModele = tableau.ListRows(indexAncrage).Range
                For decalage = 1 To nombreDemande - 1
                    Set nouvelleLigne = tableau.ListRows.Add(Position:=indexAncrage + decalage, AlwaysInsert:=True)
                    plageModele.Copy Destination:=nouvelleLigne.Range
                    resultat.Add nouvelleLigne.Range.Row
                Next decalage
                Application.CutCopyMode = False
            End If
        End If
    Else
        ligneEntete = TrouverLigneEnteteResultatCegide(ws)
        PreparerColonneClientCegide ws, ligneEntete, tableau
        colonneEtape = TrouverColonneEnteteCegide(ws, ligneEntete, "ETAPE")
        derniereColonneResultat = ws.Cells(ligneEntete, ws.Columns.Count).End(xlToLeft).Column
        derniereLigneNumero = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For ligne = ligneEntete + 1 To derniereLigneNumero
            If Not ws.Rows(ligne).Hidden Then
                If Trim$(CStr(ws.Cells(ligne, 1).Value2)) <> "" And LigneResultatDisponibleCegide(ws, ligne, colonneEtape) Then
                    ligneAncrage = ligne
                    numeroMaitre = ws.Cells(ligne, 1).Value2
                    Exit For
                End If
            End If
        Next ligne

        If ligneAncrage > 0 Then
            resultat.Add ligneAncrage
            If nombreDemande > 1 Then
                ws.Rows(ligneAncrage + 1).Resize(nombreDemande - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Set plageModele = ws.Range(ws.Cells(ligneAncrage, 1), ws.Cells(ligneAncrage, derniereColonneResultat))
                For decalage = 1 To nombreDemande - 1
                    Set plageCible = ws.Range(ws.Cells(ligneAncrage + decalage, 1), ws.Cells(ligneAncrage + decalage, derniereColonneResultat))
                    plageModele.Copy Destination:=plageCible
                    resultat.Add ligneAncrage + decalage
                Next decalage
                Application.CutCopyMode = False
            End If
        End If
    End If
    Set AjouterLignesSortieCegide = resultat
End Function

Private Sub PreparerColonneClientCegide(ByVal ws As Worksheet, ByVal ligneEntete As Long, ByVal tableau As ListObject)
    Dim titre As String, derniereLigne As Long, derniereColonne As Long
    If ligneEntete < 1 Then Err.Raise vbObjectError + 860, , "En-tete introuvable pour ajouter Client en Q."
    titre = NormaliserCompactCegide(CStr(ws.Cells(ligneEntete, 17).Value2))
    If titre <> "" And titre <> "CLIENT" And titre <> "CLIEN" And titre <> "CODECLIENT" Then
        Err.Raise vbObjectError + 861, , "La colonne Q contient deja un autre champ : " & ws.Cells(ligneEntete, 17).Value2
    End If
    If titre = "" Then
        derniereLigne = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row
        If derniereLigne > ligneEntete Then Err.Raise vbObjectError + 862, , "La colonne Q contient des donnees sans en-tete Client. Aucune donnee n'a ete remplacee."
    End If
    titre = NormaliserCompactCegide(CStr(ws.Cells(ligneEntete, 19).Value2))
    If titre <> "" And titre <> "CLIENTCORRIGE" Then Err.Raise vbObjectError + 863, , "La colonne S contient deja un autre champ."
    If titre = "" And ws.Cells(ws.Rows.Count, 19).End(xlUp).Row > ligneEntete Then Err.Raise vbObjectError + 864, , "La colonne S contient des donnees sans en-tete."
    If Not tableau Is Nothing Then
        derniereColonne = tableau.Range.Column + tableau.Range.Columns.Count - 1
        If derniereColonne < 19 Then tableau.Resize ws.Range(tableau.Range.Cells(1, 1), ws.Cells(tableau.Range.Row + tableau.Range.Rows.Count - 1, 19))
    End If
    ws.Cells(ligneEntete, 17).Value2 = "Client"
    If Trim$(CStr(ws.Cells(ligneEntete, 18).Value2)) = "" Then ws.Cells(ligneEntete, 18).Value2 = "Identifiant signataire"
    ws.Cells(ligneEntete, 19).Value2 = "Client corrige"
End Sub

Private Function LigneResultatDisponibleCegide(ByVal ws As Worksheet, ByVal ligne As Long, ByVal colonneEtape As Long) As Boolean
    If colonneEtape = 0 Then Exit Function
    LigneResultatDisponibleCegide = (Trim$(CStr(ws.Cells(ligne, colonneEtape).Value2)) = "" _
        And Trim$(CStr(ws.Cells(ligne, 2).Value2)) = "" _
        And Trim$(CStr(ws.Cells(ligne, 3).Value2)) = "" _
        And Trim$(CStr(ws.Cells(ligne, 7).Value2)) = "")
End Function

Private Function TrouverTableauResultatCegide(ByVal ws As Worksheet) As ListObject
    Dim tableau As ListObject
    Dim cellule As Range
    Dim entetes As String
    For Each tableau In ws.ListObjects
        entetes = ""
        For Each cellule In tableau.HeaderRowRange.Cells
            entetes = entetes & "|" & NormaliserCompactCegide(CStr(cellule.Value2))
        Next cellule
        If EstEnteteNumeroCegide(CStr(ws.Cells(tableau.HeaderRowRange.Row, 1).Value2)) Then
            If InStr(1, entetes, "|ARTICLE", vbTextCompare) > 0 Then
                If InStr(1, entetes, "|MONTANT", vbTextCompare) > 0 Then
                    If InStr(1, entetes, "|MODULE", vbTextCompare) > 0 Then
                        Set TrouverTableauResultatCegide = tableau
                        Exit Function
                    End If
                End If
            End If
        End If
    Next tableau
End Function

Private Function TrouverLigneEnteteResultatCegide(ByVal ws As Worksheet) As Long
    Dim ligne As Long
    For ligne = 1 To Application.Min(50, ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1)
        If EstEnteteNumeroCegide(CStr(ws.Cells(ligne, 1).Value2)) Then
            If NormaliserCompactCegide(CStr(ws.Cells(ligne, 3).Value2)) = "ARTICLE" Then
                If NormaliserCompactCegide(CStr(ws.Cells(ligne, 7).Value2)) = "MONTANT" Then
                    If NormaliserCompactCegide(CStr(ws.Cells(ligne, 12).Value2)) = "MODULE" Then
                        TrouverLigneEnteteResultatCegide = ligne
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ligne
    Err.Raise vbObjectError + 823, , "L'entete du tableau Remplissage Cegide est introuvable."
End Function

Private Function EstEnteteNumeroCegide(ByVal valeur As String) As Boolean
    Dim entete As String
    entete = NormaliserCompactCegide(valeur)
    EstEnteteNumeroCegide = (entete = "N" Or entete = "NO" Or entete = "NUMERO")
End Function

Private Function TrouverColonneEnteteCegide(ByVal ws As Worksheet, ByVal ligneEntete As Long, ByVal entete As String) As Long
    Dim derniereColonne As Long
    Dim colonne As Long
    derniereColonne = ws.Cells(ligneEntete, ws.Columns.Count).End(xlToLeft).Column
    For colonne = 1 To derniereColonne
        If NormaliserCompactCegide(CStr(ws.Cells(ligneEntete, colonne).Value2)) = NormaliserCompactCegide(entete) Then
            TrouverColonneEnteteCegide = colonne
            Exit Function
        End If
    Next colonne
End Function

Private Function MontantTrouveCegide(ByVal ligneFacture As Object) As Boolean
    MontantTrouveCegide = True
    If ligneFacture.Exists("MontantTrouve") Then MontantTrouveCegide = CBool(ligneFacture("MontantTrouve"))
End Function

Private Function TexteDetecteCegide(ByVal ligneFacture As Object) As String
    If ligneFacture.Exists("TexteDetecte") Then TexteDetecteCegide = Trim$(CStr(ligneFacture("TexteDetecte")))
    If TexteDetecteCegide = "" Then TexteDetecteCegide = Trim$(CStr(ligneFacture("OperationReference")))
End Function

Private Sub EcrireLigneCegide(ByVal ws As Worksheet, ByVal ligne As Long, ByVal numero As Variant, ByVal texteDetecte As String, ByVal codeArticle As String, ByVal montant As Double, ByVal redacteur As Object, ByVal moduleFacture As String, ByVal colonneEtape As Long, Optional ByVal montantTrouve As Boolean = True, Optional ByVal clientFacture As String = "")
    ws.Cells(ligne, 1).Value2 = numero
    ws.Cells(ligne, 2).NumberFormat = "@"
    ws.Cells(ligne, 2).Value2 = texteDetecte
    ws.Cells(ligne, 3).NumberFormat = "@"
    ws.Cells(ligne, 3).Value2 = codeArticle
    ws.Cells(ligne, 4).Value2 = 1
    If montantTrouve Then
        ws.Cells(ligne, 7).Value2 = montant
        ws.Cells(ligne, 7).NumberFormat = "0.00"
    Else
        ws.Cells(ligne, 7).Value2 = "Montant non lu - a verifier"
    End If
    EcrireChampsCommunsCegide ws, ligne, redacteur, moduleFacture, clientFacture
    AppliquerValidationEtapeCegide ws.Cells(ligne, colonneEtape), True
End Sub

Private Sub EcrireChampsCommunsCegide(ByVal ws As Worksheet, ByVal ligne As Long, ByVal redacteur As Object, ByVal moduleFacture As String, Optional ByVal clientFacture As String = "")
    ws.Cells(ligne, 18).NumberFormat = "@"
    ws.Cells(ligne, 18).ClearContents
    If redacteur Is Nothing Then
        ws.Cells(ligne, 8).ClearContents
        ws.Cells(ligne, 9).ClearContents
        ws.Cells(ligne, 15).ClearContents
    Else
        ws.Cells(ligne, 8).Value2 = CStr(redacteur("CodeOrganisation"))
        ws.Cells(ligne, 9).Value2 = CStr(redacteur("Destination"))
        ws.Cells(ligne, 15).Value2 = CStr(redacteur("Nom"))
        If redacteur.Exists("Identifiant") Then ws.Cells(ligne, 18).Value2 = CStr(redacteur("Identifiant"))
    End If
    ws.Cells(ligne, 10).NumberFormat = "@"
    ws.Cells(ligne, 10).Value2 = "00000000"
    ws.Cells(ligne, 11).Value2 = "GLOC"
    ws.Cells(ligne, 12).NumberFormat = "@"
    ws.Cells(ligne, 12).Value2 = moduleFacture
    ws.Cells(ligne, 17).NumberFormat = "@"
    ws.Cells(ligne, 17).Value2 = clientFacture
    ws.Cells(ligne, 19).NumberFormat = "@"
    ws.Cells(ligne, 19).Value2 = Replace(ClientSaisiCegide(clientFacture), "/", "")
    ws.Cells(ligne, 13).Value = DateSerial(2026, 1, 1)
    ws.Cells(ligne, 13).NumberFormat = "dd/mm/yyyy"
    ws.Cells(ligne, 14).Value = DateSerial(2026, 12, 31)
    ws.Cells(ligne, 14).NumberFormat = "dd/mm/yyyy"
    ValidationSignataireCegide ws.Cells(ligne, 15)
End Sub

Private Sub ValidationSignataireCegide(ByVal cellule As Range)
    Dim source As Worksheet, derniere As Long, premiere As Long, plage As String
    Set source = TrouverFeuilleCegide(cellule.Worksheet.Parent, FEUILLE_ORGANISATION)
    If source Is Nothing Then Exit Sub
    derniere = source.Cells(source.Rows.Count, 15).End(xlUp).Row
    premiere = 1
    If InStr(NormaliserCompactCegide(CStr(source.Cells(1, 15).Value2)), "REDACTEUR") > 0 Or NormaliserCompactCegide(CStr(source.Cells(1, 15).Value2)) = "SIGNATAIRE" Then premiere = 2
    If derniere < premiere Then derniere = premiere
    plage = "='" & Replace(source.Name, "'", "''") & "'!$O$" & premiere & ":$O$" & derniere
    cellule.Worksheet.Parent.Names.Add Name:="_GLOIRE_CEGIDE_SIGNATAIRES", RefersTo:=plage
    With cellule.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=_GLOIRE_CEGIDE_SIGNATAIRES"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Sub AppliquerValidationEtapeCegide(ByVal cellule As Range, ByVal ligneComplete As Boolean)
    Dim separateur As String
    Dim conditionVerte As FormatCondition
    separateur = Application.International(xlListSeparator)
    On Error Resume Next
    cellule.Validation.Delete
    cellule.FormatConditions.Delete
    On Error GoTo 0
    cellule.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Traiter" & separateur & "OK" & separateur & "Valider"
    cellule.Validation.IgnoreBlank = True
    cellule.Validation.InCellDropdown = True
    Set conditionVerte = cellule.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & cellule.Address & "=""Traiter""")
    conditionVerte.Interior.Color = RGB(198, 239, 206)
    conditionVerte.Font.Color = RGB(0, 97, 0)
    If ligneComplete Then cellule.Value2 = "Traiter" Else cellule.ClearContents
End Sub

Private Function NormaliserEspacesCegide(ByVal texte As String) As String
    Dim source As String
    Dim resultat As String
    Dim caractere As String
    Dim i As Long
    source = UCase$(RemplacerAccentsCegide(texte))
    For i = 1 To Len(source)
        caractere = Mid$(source, i, 1)
        If caractere Like "[A-Z0-9]" Then resultat = resultat & caractere Else resultat = resultat & " "
    Next i
    Do While InStr(1, resultat, "  ", vbBinaryCompare) > 0
        resultat = Replace(resultat, "  ", " ")
    Loop
    NormaliserEspacesCegide = Trim$(resultat)
End Function

Private Function NormaliserCompactCegide(ByVal texte As String) As String
    NormaliserCompactCegide = Replace(NormaliserEspacesCegide(texte), " ", "")
End Function

Private Function AliasRedacteurCegide(ByVal nom As String) As String
    Dim mots As Variant
    Dim mot As Variant
    Dim candidat As String
    mots = Split(NormaliserEspacesCegide(nom), " ")
    For Each mot In mots
        candidat = CStr(mot)
        If Len(candidat) >= 3 And Not EstCodeServiceCegide(candidat) Then AliasRedacteurCegide = candidat
    Next mot
End Function

Private Function EstCodeServiceCegide(ByVal mot As String) As Boolean
    Select Case UCase$(Trim$(mot))
        Case "DTGE", "DTGO", "DTMGP", "MVH", "SERVICE", "CONTENTIEUX"
            EstCodeServiceCegide = True
    End Select
End Function

Private Function RemplacerAccentsCegide(ByVal texte As String) As String
    Dim codes As Variant
    Dim lettres As Variant
    Dim i As Long
    codes = Array(192, 193, 194, 195, 196, 197, 199, 200, 201, 202, 203, 204, 205, 206, 207, 209, 210, 211, 212, 213, 214, 217, 218, 219, 220, 221, 224, 225, 226, 227, 228, 229, 231, 232, 233, 234, 235, 236, 237, 238, 239, 241, 242, 243, 244, 245, 246, 249, 250, 251, 252, 253, 255)
    lettres = Array("A", "A", "A", "A", "A", "A", "C", "E", "E", "E", "E", "I", "I", "I", "I", "N", "O", "O", "O", "O", "O", "U", "U", "U", "U", "Y", "A", "A", "A", "A", "A", "A", "C", "E", "E", "E", "E", "I", "I", "I", "I", "N", "O", "O", "O", "O", "O", "U", "U", "U", "U", "Y", "Y")
    RemplacerAccentsCegide = texte
    For i = LBound(codes) To UBound(codes)
        RemplacerAccentsCegide = Replace(RemplacerAccentsCegide, ChrW$(CLng(codes(i))), CStr(lettres(i)))
    Next i
    RemplacerAccentsCegide = Replace(RemplacerAccentsCegide, ChrW$(338), "OE")
    RemplacerAccentsCegide = Replace(RemplacerAccentsCegide, ChrW$(339), "OE")
End Function

Private Function AjouterProblemeCegide(ByVal problemes As String, ByVal nouveauProbleme As String) As String
    If problemes = "" Then
        AjouterProblemeCegide = nouveauProbleme
    Else
        AjouterProblemeCegide = problemes & vbCrLf & "- " & nouveauProbleme
    End If
End Function

' TSV du pont embarque : coordonnees en pourcentage, independantes du DPI.
Private Function LirePositionsCegide(ByVal chemin As String) As Collection
    Dim resultat As New Collection, lignes As Variant, champs As Variant
    Dim i As Long, mot As Object, champ As Long, reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "^[0-9]+([.][0-9]+)?$"
    If Dir$(chemin) = "" Then Err.Raise vbObjectError + 850, , "TSV spatial absent : " & chemin
    lignes = Split(Replace(LireUTF8Cegide(chemin), vbCr, ""), vbLf)
    If InStr(1, CStr(lignes(0)), "x_pct", vbTextCompare) = 0 Then Err.Raise vbObjectError + 851, , "Format TSV spatial incompatible."
    For i = 1 To UBound(lignes)
        If Trim$(CStr(lignes(i))) <> "" Then
            champs = Split(CStr(lignes(i)), vbTab)
            If UBound(champs) < 7 Then Err.Raise vbObjectError + 852, , "Ligne TSV incomplete : " & i + 1
            For champ = 0 To 6
                If Not reg.Test(CStr(champs(champ))) Then Err.Raise vbObjectError + 854, , "Coordonnee TSV invalide ligne " & i + 1
            Next champ
            If Val(champs(0)) < 1 Or Val(champs(1)) <= 0 Or Val(champs(2)) <= 0 Or Val(champs(5)) <= 0 Or Val(champs(6)) <= 0 Then Err.Raise vbObjectError + 854, , "Dimensions TSV invalides ligne " & i + 1
            If Val(champs(3)) + Val(champs(5)) > 101 Or Val(champs(4)) + Val(champs(6)) > 101 Then Err.Raise vbObjectError + 854, , "Coordonnees TSV hors page ligne " & i + 1
            Set mot = CreateObject("Scripting.Dictionary")
            mot.Add "p", CLng(Val(champs(0)))
            mot.Add "x", Val(champs(3))
            mot.Add "y", Val(champs(4))
            mot.Add "w", Val(champs(5))
            mot.Add "h", Val(champs(6))
            mot.Add "t", CStr(champs(7))
            resultat.Add mot
        End If
    Next i
    If resultat.Count = 0 Then Err.Raise vbObjectError + 853, , "TSV sans mots."
    Set LirePositionsCegide = resultat
End Function

Private Function LignesPositionsCegide(ByVal mots As Collection) As Collection
    Dim resultat As New Collection, mot As Object, ligne As Object, retenue As Object
    Dim groupe As Collection, j As Long, distance As Double, meilleure As Double
    Dim pentes As Object, ancre As Object, fin As Object, clePage As String, pente As Double, yCorrige As Double, n As String
    Dim bornes As Object, ordre As New Collection, prioritaire As Boolean, passe As Long, regMontant As Object, bornesPage As Variant, typeAncre As Variant
    Set pentes = CreateObject("Scripting.Dictionary")
    Set bornes = CreateObject("Scripting.Dictionary")
    Set regMontant = CreateObject("VBScript.RegExp")
    regMontant.Pattern = "^[+-]?[0-9]+[.,][0-9]{2}$"
    ' Redresser les lignes de lecture des photos inclinees, sans modifier le PDF ni le TSV.
    For Each ancre In mots
        n = NormaliserCompactCegide(CStr(ancre("t")))
        clePage = CStr(ancre("p"))
        If (n = "NATURE" Or n = "LIBELLE") And Not pentes.Exists(clePage) Then
          For Each typeAncre In Array("TTC", "TVA", "NS", "HT")
            For Each fin In mots
                If fin("p") = ancre("p") And TypeColonneSpatialCegide(CStr(fin("t"))) = typeAncre Then
                    If fin("x") - ancre("x") > 20 And Abs(fin("y") - ancre("y")) < 3 Then
                        pente = ((fin("y") + fin("h") / 2) - (ancre("y") + ancre("h") / 2)) / (fin("x") - ancre("x"))
                        pentes.Add clePage, pente
                        bornes.Add clePage, Array(ancre("x") + (fin("x") - ancre("x")) * 0.45, fin("x") + 4)
                        Exit For
                    End If
                End If
            Next fin
            If pentes.Exists(clePage) Then Exit For
          Next typeAncre
        End If
    Next ancre
    ' Ancrer les lignes sur les montants avant d'y rattacher le texte.
    ' Les textes des barres laterales ne peuvent ainsi pas deplacer une ligne de facture.
    For passe = 1 To 2
        For Each mot In mots
            prioritaire = False
            If bornes.Exists(CStr(mot("p"))) Then
                bornesPage = bornes(CStr(mot("p")))
                If mot("x") >= bornesPage(0) And mot("x") <= bornesPage(1) Then prioritaire = regMontant.Test(CStr(mot("t")))
            End If
            If (passe = 1 And prioritaire) Or (passe = 2 And Not prioritaire) Then ordre.Add mot
        Next mot
    Next passe
    For Each mot In ordre
        pente = 0
        If pentes.Exists(CStr(mot("p"))) Then pente = pentes(CStr(mot("p")))
        yCorrige = mot("y") + mot("h") / 2 - pente * mot("x")
        Set retenue = Nothing
        meilleure = 1000
        For Each ligne In resultat
            If ligne("p") = mot("p") Then
                distance = Abs(CDbl(ligne("y")) - yCorrige)
                If distance < Application.Max(0.25, Application.Min(ligne("h"), mot("h")) * 0.85) And distance < meilleure Then
                    Set retenue = ligne
                    meilleure = distance
                End If
            End If
        Next ligne
        If retenue Is Nothing Then
            Set retenue = CreateObject("Scripting.Dictionary")
            retenue.Add "p", mot("p")
            retenue.Add "y", yCorrige
            retenue.Add "h", mot("h")
            Set groupe = New Collection
            retenue.Add "m", groupe
            For j = 1 To resultat.Count
                If resultat(j)("p") > retenue("p") Then Exit For
                If resultat(j)("p") = retenue("p") And resultat(j)("y") > retenue("y") Then Exit For
            Next j
            If j > resultat.Count Then resultat.Add retenue Else resultat.Add retenue, Before:=j
        End If
        Set groupe = retenue("m")
        For j = 1 To groupe.Count
            If groupe(j)("x") > mot("x") Then Exit For
        Next j
        If j > groupe.Count Then groupe.Add mot Else groupe.Add mot, Before:=j
    Next mot
    Set LignesPositionsCegide = resultat
End Function

Private Function ControlerTXTAncienCegide(ByVal mots As Collection, ByVal texteTXT As String, ByVal articles As Collection, ByRef probleme As String) As Collection
    Dim sorties As Collection, brute As Variant, libelle As String, norme As String, actif As Boolean
    Dim reg As Object, correspondances As Object, article As Object, ambigu As Boolean, operation As Object
    Dim attendues As Object, cle As Variant, position As Long, sansDate As String
    Set sorties = AnalyserPositionsCegide(mots, articles, probleme)
    Set attendues = CreateObject("Scripting.Dictionary")
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "[+-]?[0-9]+[.,][0-9]{2}([^0-9]|$)"
    ' Le TXT controle les omissions, jamais les positions ni l'affectation des montants.
    For Each brute In Split(Replace(texteTXT, vbCr, ""), vbLf)
        norme = NormaliserCompactCegide(CStr(brute))
        If EstEnteteTableauFlexibleCegide(CStr(brute)) Or InStr(norme, "NATUREDELOPERATION") > 0 Then
            actif = True
        ElseIf actif Then
            If EstFinTableauFlexibleCegide(norme) Or Left$(norme, 5) = "SOLDE" Then
                actif = False
            ElseIf Left$(norme, 3) <> "TVA" Then
              sansDate = SupprimerDatesCegide(CStr(brute))
              If reg.Test(sansDate) Then
                Set correspondances = reg.Execute(sansDate)
                position = correspondances(0).FirstIndex
                libelle = Trim$(Left$(sansDate, position))
                If Len(libelle) > 3 Then
                    Set article = ArticleSpatialCegide(libelle, articles, ambigu)
                    If Not article Is Nothing Then
                        cle = NormaliserCompactCegide(CStr(article("Operation")))
                        If Not attendues.Exists(cle) Then attendues.Add cle, 0
                        attendues(cle) = attendues(cle) + 1
                    End If
                End If
              End If
            End If
        End If
    Next brute
    For Each cle In attendues.Keys
        position = 0
        For Each operation In sorties
            norme = NormaliserCompactCegide(CStr(operation("OperationReference")))
            If norme = CStr(cle) Then
                ' Les deux ecritures HT/NS d'une meme operation ne sont pas deux lignes source.
                If Not EstDeboursSeulementCegide(operation) Then position = position + 1
            End If
        Next operation
        If position = 0 Then
            For Each operation In sorties
                norme = NormaliserCompactCegide(CStr(operation("OperationReference")))
                If norme = CStr(cle) Then position = position + 1
            Next operation
        End If
        If position < attendues(cle) Then probleme = AjouterProblemeCegide(probleme, "Controle TXT/TSV : ligne(s) potentiellement manquante(s) pour " & CStr(cle) & ". Verification necessaire.")
    Next cle
    Set ControlerTXTAncienCegide = sorties
End Function

Private Function MotsTableauTXTCegide(ByVal texte As String) As Collection
    Dim resultat As New Collection, lignes As Variant, brute As Variant, reg As Object, occurrence As Object, mot As Object
    Dim largeur As Long, hauteur As Long, numero As Long
    lignes = Split(Replace(Replace(texte, vbCr, ""), vbTab, "    "), vbLf)
    hauteur = UBound(lignes) + 2
    For Each brute In lignes
        If Len(CStr(brute)) > largeur Then largeur = Len(CStr(brute))
    Next brute
    If largeur = 0 Then Set MotsTableauTXTCegide = resultat: Exit Function
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True: reg.Pattern = "\S+"
    For Each brute In lignes
        numero = numero + 1
        For Each occurrence In reg.Execute(CStr(brute))
            Set mot = CreateObject("Scripting.Dictionary")
            mot.Add "p", 1
            mot.Add "x", occurrence.FirstIndex * 100# / largeur
            mot.Add "y", numero * 100# / hauteur
            mot.Add "w", occurrence.Length * 100# / largeur
            mot.Add "h", 90# / hauteur
            mot.Add "t", occurrence.Value
            resultat.Add mot
        Next occurrence
    Next brute
    ' Ces positions de caracteres servent uniquement au TXT, jamais a modifier le TSV.
    Set MotsTableauTXTCegide = resultat
End Function

Private Function CleLibelleFusionCegide(ByVal texte As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "^\s*[0-9][0-9./ -]*\s*"
    CleLibelleFusionCegide = NormaliserCompactCegide(reg.Replace(SupprimerDatesCegide(texte), ""))
End Function

Private Function AnalyserCombineCegide(ByVal mots As Collection, ByVal texteTXT As String, ByVal articles As Collection, ByRef probleme As String) As Collection
    Dim spatial As Collection, principal As Collection, texteMots As Collection, resultat As New Collection
    Dim ligne As Object, autre As Object, utilise As Object, i As Long, retenu As Long, cle As String, cleAutre As String
    Dim champ As String, avisTXT As String, avisTSV As String, ttcTXT As Boolean, ttcTSV As Boolean
    Set spatial = AnalyserPositionsCegide(mots, articles, avisTSV)
    Set texteMots = MotsTableauTXTCegide(texteTXT)
    Set principal = AnalyserPositionsCegide(texteMots, articles, avisTXT)
    Set utilise = CreateObject("Scripting.Dictionary")
    For Each ligne In principal
        If RepliContreditCegide(ligne, spatial) Then GoTo LigneTXTSuivante
        cle = CleLibelleFusionCegide(CStr(ligne("TexteDetecte")))
        retenu = 0
        For i = 1 To spatial.Count
            If Not utilise.Exists(CStr(i)) Then
                Set autre = spatial(i)
                If EstDeboursSeulementCegide(ligne) = EstDeboursSeulementCegide(autre) Then
                    cleAutre = CleLibelleFusionCegide(CStr(autre("TexteDetecte")))
                    If cle = cleAutre Or (Len(cle) >= 8 And Len(cleAutre) >= 8 And (InStr(cle, cleAutre) > 0 Or InStr(cleAutre, cle) > 0)) Then
                        retenu = i: Exit For
                    End If
                End If
            End If
        Next i
        If retenu > 0 Then
            utilise.Add CStr(retenu), True
            Set autre = spatial(retenu)
            cleAutre = CleLibelleFusionCegide(CStr(autre("TexteDetecte")))
            If Len(cleAutre) > Len(cle) And InStr(cleAutre, cle) > 0 Then ligne("TexteDetecte") = autre("TexteDetecte")
            champ = "MontantHT"
            If EstDeboursSeulementCegide(ligne) Then champ = "Debours"
            ttcTXT = (CStr(ligne("BaseMontant")) = "TTC")
            ttcTSV = (CStr(autre("BaseMontant")) = "TTC")
            If ttcTXT And Not ttcTSV Then
                ligne(champ) = autre(champ): ligne("MontantTrouve") = MontantTrouveCegide(autre): ligne("BaseMontant") = autre("BaseMontant")
            ElseIf Not ttcTXT And ttcTSV Then
                ' Une colonne HT detectee dans une source interdit le repli TTC de l'autre.
            ElseIf MontantTrouveCegide(ligne) And MontantTrouveCegide(autre) Then
                If Abs(CDbl(ligne(champ)) - CDbl(autre(champ))) > 0.005 Then
                    If CoherenceOperationCegide(ligne) = -1 And CoherenceOperationCegide(autre) = 1 Then
                        probleme = AjouterProblemeCegide(probleme, "Conflit TXT/TSV : montant TSV retenu " & CStr(autre(champ)) & " au lieu de " & CStr(ligne(champ)) & " ; controle HT+TVA+NS=TTC coherent uniquement en TSV : " & ligne("TexteDetecte"))
                        ligne(champ) = autre(champ)
                    Else
                        probleme = AjouterProblemeCegide(probleme, "Conflit TXT/TSV, montant TXT conserve " & CStr(ligne(champ)) & " (TSV : " & CStr(autre(champ)) & ") : " & ligne("TexteDetecte"))
                    End If
                End If
            ElseIf Not MontantTrouveCegide(ligne) And MontantTrouveCegide(autre) Then
                ligne(champ) = autre(champ): ligne("MontantTrouve") = True
            End If
            If CStr(ligne("Code")) = "" And CStr(autre("Code")) <> "" Then
                ligne("Code") = autre("Code")
                If autre.Exists("CodeDebours") Then ligne("CodeDebours") = autre("CodeDebours")
            End If
        Else
            probleme = AjouterProblemeCegide(probleme, "Ligne TXT non confirmee par le TSV, conservee : " & ligne("TexteDetecte"))
        End If
        resultat.Add ligne
LigneTXTSuivante:
    Next ligne
    For i = 1 To spatial.Count
        If Not utilise.Exists(CStr(i)) And Not RepliContreditCegide(spatial(i), principal) Then
            resultat.Add spatial(i)
            If principal.Count > 0 Then probleme = AjouterProblemeCegide(probleme, "Ligne supplementaire TSV conservee : " & spatial(i)("TexteDetecte"))
        End If
    Next i
    If avisTSV <> "" Then probleme = AjouterProblemeCegide(probleme, "TSV : " & avisTSV)
    If avisTXT <> "" And Len(Trim$(texteTXT)) > 0 Then probleme = AjouterProblemeCegide(probleme, "TXT : " & avisTXT)
    Set AnalyserCombineCegide = resultat
End Function

Private Function CoherenceOperationCegide(ByVal ligne As Object) As Long
    If ligne.Exists("Coherence") Then CoherenceOperationCegide = CLng(ligne("Coherence"))
End Function

Private Function RepliContreditCegide(ByVal ligne As Object, ByVal autres As Collection) As Boolean
    Dim autre As Object, cle As String, cleAutre As String
    If CStr(ligne("BaseMontant")) <> "TTC" Then Exit Function
    cle = CleLibelleFusionCegide(CStr(ligne("TexteDetecte")))
    For Each autre In autres
        If CBool(autre("PresenceHT")) Then
            cleAutre = CleLibelleFusionCegide(CStr(autre("TexteDetecte")))
            If cle = cleAutre Or (Len(cle) >= 8 And Len(cleAutre) >= 8 And (InStr(cle, cleAutre) > 0 Or InStr(cleAutre, cle) > 0)) Then
                RepliContreditCegide = True
                Exit Function
            End If
        End If
    Next autre
End Function

Private Function TextePositionCegide(ByVal ligne As Object) As String
    Dim mot As Object, texte As String
    For Each mot In ligne("m")
        texte = texte & " " & mot("t")
    Next mot
    TextePositionCegide = Trim$(texte)
End Function

Private Function EstRepereReferenceCegide(ByVal texte As String) As Boolean
    Dim norme As String
    norme = NormaliserCompactCegide(texte)
    Select Case norme
        Case "REF", "REFS", "REFER", "REFERENC", "REFERENCE", "REFERENCES", "REFERANCE", "REFERANCES"
            EstRepereReferenceCegide = True
        Case "VREF", "VREFS", "VOSREF", "VOSREFS", "VOTREREF", "VOSREFERENCE", "VOSREFERENCES", "VOTREREFERENCE", "VOSREFER", "VOSREFERANCES"
            EstRepereReferenceCegide = True
        Case "NOSREF", "NOSREFS", "NOSREFERENCE", "NOSREFERENCES", "NOTREREFERENCE", "REFERENCEARAPPELER", "REFERENCESARAPPELER"
            EstRepereReferenceCegide = True
        Case "REFCLIENT", "REFSCLIENT", "REFERENCECLIENT", "REFERENCESCLIENT", "REFCLIENTS", "REFERENCESCLIENTS", "REFDOSSIER", "REFERENCEDOSSIER"
            EstRepereReferenceCegide = True
    End Select
End Function

Private Function ClientSpatialCegide(ByVal mots As Collection, ByRef probleme As String) As String
    Dim lignes As Collection, ligne As Object, voisine As Object, mot As Object, reg As Object, occurrence As Object
    Dim i As Long, j As Long, repere As Boolean, candidat As String, distance As Double, meilleure As Double, ambigu As Boolean
    Set lignes = LignesPositionsCegide(mots)
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True: reg.IgnoreCase = True
    reg.Pattern = "(^|[^A-Z0-9])L\s*/\s*([0-9]{1,7})(?![0-9A-Z])"
    meilleure = 1000
    For i = 1 To lignes.Count
        Set ligne = lignes(i): repere = False
        If ligne("p") = 1 Then
            For Each mot In ligne("m")
                If EstRepereReferenceCegide(CStr(mot("t"))) Then repere = True: Exit For
            Next mot
            If repere Then
                For j = Application.Max(1, i - 2) To Application.Min(lignes.Count, i + 2)
                    Set voisine = lignes(j)
                    If voisine("p") = ligne("p") And Abs(voisine("y") - ligne("y")) <= 8 Then
                        distance = Abs(voisine("y") - ligne("y"))
                        For Each occurrence In reg.Execute(TextePositionCegide(voisine))
                            candidat = "L/" & CStr(occurrence.SubMatches(1))
                            If distance < meilleure - 0.01 Then
                                ClientSpatialCegide = candidat: meilleure = distance: ambigu = False
                            ElseIf Abs(distance - meilleure) <= 0.01 And candidat <> ClientSpatialCegide Then
                                ambigu = True
                            End If
                        Next occurrence
                    End If
                Next j
            End If
        End If
    Next i
    If ambigu Then
        ClientSpatialCegide = ""
        probleme = AjouterProblemeCegide(probleme, "Plusieurs codes client L/ proches des references : Client laisse vide.")
    ElseIf ClientSpatialCegide = "" Then
        ClientSpatialCegide = ClientDeduitCegide(mots)
        If ClientSpatialCegide <> "" Then
            probleme = AjouterProblemeCegide(probleme, "Client deduit sans prefixe L/ reconnu, a verifier sur la facture : " & ClientSpatialCegide)
        Else
            probleme = AjouterProblemeCegide(probleme, "Client introuvable ou plusieurs numeros possibles pres du module : Client laisse vide.")
        End If
    End If
End Function

Private Function ClientDeduitCegide(ByVal mots As Collection) As String
    Dim lignes As Collection, ligne As Object, voisine As Object, mot As Object, ancre As Object
    Dim module As String, avis As String, candidats As Object, reg As Object, numero As String, cle As Variant
    Dim i As Long, j As Long, k As Long, reference As Boolean, n As String
    Dim fragment As Object, prefixeAmbigu As Boolean, regFragment As Object
    Dim regLettre As Object, correspondances As Object, lettreClient As Boolean, distanceMax As Double
    module = ModuleSpatialCegide(mots, avis)
    If Len(module) <> 10 Then Exit Function
    Set lignes = LignesPositionsCegide(mots)
    Set candidats = CreateObject("Scripting.Dictionary")
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "^[0-9]{1,7}$"
    Set regFragment = CreateObject("VBScript.RegExp")
    regFragment.Pattern = "^[UWIL|/\]]{2,4}$"
    Set regLettre = CreateObject("VBScript.RegExp")
    regLettre.Pattern = "^[: ]*L([0-9]{1,7})-?$"
    regLettre.IgnoreCase = True
    For i = 1 To lignes.Count
        Set ligne = lignes(i)
        If ligne("p") = 1 And InStr(TextePositionCegide(ligne), module) > 0 Then
            reference = False
            For k = Application.Max(1, i - 2) To Application.Min(lignes.Count, i + 2)
                Set voisine = lignes(k)
                If voisine("p") = 1 And Abs(voisine("y") - ligne("y")) <= 8 Then
                    For Each mot In voisine("m")
                        If EstRepereReferenceCegide(CStr(mot("t"))) Then reference = True
                    Next mot
                End If
            Next k
            If reference Then
                For Each ancre In ligne("m")
                    If InStr(CStr(ancre("t")), module) > 0 Then
                        For j = Application.Max(1, i - 1) To Application.Min(lignes.Count, i + 1)
                            Set voisine = lignes(j)
                            n = NormaliserCompactCegide(TextePositionCegide(voisine))
                            If voisine("p") = 1 And Abs(voisine("y") - ligne("y")) <= 4 And InStr(n, "INTRACOM") = 0 And InStr(n, "MANDAT") = 0 And InStr(n, "NOSREF") = 0 Then
                                For Each mot In voisine("m")
                                    numero = Trim$(CStr(mot("t")))
                                    lettreClient = False: distanceMax = 20
                                    Set correspondances = regLettre.Execute(numero)
                                    If correspondances.Count = 1 Then
                                        numero = CStr(correspondances(0).SubMatches(0))
                                        lettreClient = True: distanceMax = 30
                                    End If
                                    If Left$(numero, 1) = "-" Then numero = Mid$(numero, 2)
                                    If Right$(numero, 1) = "-" Then numero = Left$(numero, Len(numero) - 1)
                                    If reg.Test(numero) And (numero <> Left$(module, 4) Or lettreClient) Then
                                        If Abs((mot("x") + mot("w") / 2) - (ancre("x") + ancre("w") / 2)) <= distanceMax Then
                                            prefixeAmbigu = False
                                            For Each fragment In voisine("m")
                                                If fragment("x") < mot("x") And mot("x") - fragment("x") - fragment("w") >= -0.2 And mot("x") - fragment("x") - fragment("w") < 1.5 Then
                                                    If regFragment.Test(UCase$(Trim$(CStr(fragment("t"))))) Then prefixeAmbigu = True
                                                End If
                                            Next fragment
                                            If prefixeAmbigu Then Exit Function
                                            If Not candidats.Exists(numero) Then candidats.Add numero, True
                                        End If
                                    End If
                                Next mot
                            End If
                        Next j
                    End If
                Next ancre
            End If
        End If
    Next i
    If candidats.Count = 1 Then
        For Each cle In candidats.Keys
            ClientDeduitCegide = "L/" & CStr(cle)
        Next cle
    End If
End Function

Private Function NomProcheReferencesCegide(ByVal nom As String, ByVal mots As Collection, ByVal moduleFacture As String) As Boolean
    Dim lignes As Collection, ligne As Object, voisine As Object, mot As Object, repere As Object
    Dim i As Long, j As Long, ecart As Double, texteZone As String
    If mots Is Nothing Or Len(moduleFacture) <> 10 Then Exit Function
    Set lignes = LignesPositionsCegide(mots)
    For i = 1 To lignes.Count
        Set ligne = lignes(i)
        For Each repere In ligne("m")
            If ligne("p") = 1 And InStr(CStr(repere("t")), moduleFacture) > 0 Then
                texteZone = ""
                For j = Application.Max(1, i - 2) To Application.Min(lignes.Count, i + 2)
                    Set voisine = lignes(j)
                    If voisine("p") = ligne("p") And Abs(voisine("y") - ligne("y")) <= 6 Then
                        For Each mot In voisine("m")
                            ecart = Application.Max(0, mot("x") - repere("x") - repere("w"), repere("x") - mot("x") - mot("w"))
                            If ecart <= 25 Then texteZone = texteZone & " " & mot("t")
                        Next mot
                    End If
                Next j
                If ScoreRedacteurCegide(nom, " " & NormaliserEspacesCegide(texteZone) & " ", NormaliserCompactCegide(texteZone)) > 0 Then NomProcheReferencesCegide = True: Exit Function
            End If
        Next repere
    Next i
End Function

Private Function ModuleSpatialCegide(ByVal mots As Collection, ByRef probleme As String) As String
    Dim mot As Object, repere As Object, voisin As Object, reg As Object, occurrences As Object, occurrence As Object
    Dim candidat As String, norme As String, distance As Double, meilleure As Double, score As Double, maximum As Double
    Dim egalite As Boolean, prefixe As Boolean, reperes As New Collection
    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.Pattern = "(^|[^0-9])([0-9]{10})([^0-9]|$)"
    For Each repere In mots
        If EstRepereReferenceCegide(CStr(repere("t"))) Then reperes.Add repere
    Next repere
    For Each mot In mots
        Set occurrences = reg.Execute(CStr(mot("t")))
        For Each occurrence In occurrences
            candidat = occurrence.SubMatches(1)
            meilleure = 1000
            prefixe = False
            For Each repere In reperes
                If repere("p") = mot("p") Then
                    distance = Abs(mot("y") - repere("y")) * 2 + Abs(mot("x") - repere("x")) * 0.25
                    If Abs(mot("y") - repere("y")) < 8 And Abs(mot("x") - repere("x")) < 65 Then
                        If distance < meilleure Then meilleure = distance
                    End If
                End If
            Next repere
            reg.Pattern = "(^|[^0-9])" & Left$(candidat, 4) & "([^0-9]|$)"
            For Each repere In mots
                If repere("p") = mot("p") Then
                    If Abs(mot("y") - repere("y")) < 4 And Abs(mot("x") - repere("x")) < 18 Then
                        If reg.Test(CStr(repere("t"))) Then prefixe = True
                    End If
                End If
            Next repere
            reg.Pattern = "(^|[^0-9])([0-9]{10})([^0-9]|$)"
            If meilleure < 1000 Then
                score = 100 - meilleure
                If prefixe Then score = score + 10
                If score > maximum Then
                    maximum = score
                    ModuleSpatialCegide = candidat
                    egalite = False
                ElseIf Abs(score - maximum) < 0.01 And candidat <> ModuleSpatialCegide Then
                    egalite = True
                End If
            End If
        Next occurrence
    Next mot
    If egalite Then
        ModuleSpatialCegide = ""
        probleme = AjouterProblemeCegide(probleme, "Plusieurs modules proches des references : verification necessaire.")
    End If
End Function

Private Function TypeColonneSpatialCegide(ByVal texte As String) As String
    Dim norme As String, prefixe As Variant, suffixe As Variant, avant As String, passage As Long
    norme = NormaliserCompactCegide(texte)
    For Each suffixe In Array("EUROS", "EURO", "EUR")
        If Right$(norme, Len(suffixe)) = suffixe Then norme = Left$(norme, Len(norme) - Len(suffixe)): Exit For
    Next suffixe
    For passage = 1 To 4
        avant = norme
        For Each prefixe In Array("MONTANTS", "MONTANT", "BASES", "BASE", "PRIX", "MTS", "MT", "NETS", "NET")
            If Left$(norme, Len(prefixe)) = prefixe Then norme = Mid$(norme, Len(prefixe) + 1): Exit For
        Next prefixe
        If norme = avant Then Exit For
    Next passage
    Select Case norme
        Case "HT", "HTVA", "HORSTAXE", "HORSTAXES", "HORSTVA", "HORSTOUTETAXE", "HORSTOUTESTAXES": TypeColonneSpatialCegide = "HT"
        Case "HTNET", "HTNETS", "HTTOTAL", "HORSTAXENET", "HORSTAXESNET", "HORSDETAXE", "HORSDESTAXES": TypeColonneSpatialCegide = "HT"
        Case "TVA": TypeColonneSpatialCegide = "TVA"
        Case "TTC", "TOUTESTAXESCOMPRISES", "TOUTETAXECOMPRISE": TypeColonneSpatialCegide = "TTC"
        Case "T", "TAXE", "CREDIT", "DEBIT", "DEBITS", "TAUX", "TAUXTVA", "TAUXDETVA": TypeColonneSpatialCegide = "IGNORE"
        Case "DEBOURS", "DEBOUR", "EXO", "EXOS", "NS", "NONSOUMIS", "NONSOUMISE", "NONSOUMISES", "EXONERE", "EXONERES", "EXONEREE", "EXONEREES": TypeColonneSpatialCegide = "NS"
        Case "NONSOUMISATVA", "NONSOUMISALATVA", "NONSOUMISEATVA", "NONSOUMISEALATVA", "NONSOUMISESATVA", "NONSOUMISESALATVA": TypeColonneSpatialCegide = "NS"
        Case "FRAISNONSOUMIS", "FRAISEXONERES", "EXONERATION", "EXONERATIONS": TypeColonneSpatialCegide = "NS"
        Case "DEBOURSNONSOUMIS", "DEBOURSEXONERES", "NONASSUJETTI", "NONASSUJETTIS", "NONASSUJETTIE", "NONASSUJETTIES": TypeColonneSpatialCegide = "NS"
        Case "DEBOURSNS", "DEBOURSEXO", "FRAISDEBOURS", "NONSOUMISTVA", "NONSOUMISETVA", "NONSOUMISESTVA", "EXONEREDETVA", "EXONERESDETVA", "EXONEREEDETVA", "EXONEREESDETVA": TypeColonneSpatialCegide = "NS"
    End Select
End Function

Private Function ColonnesPositionsCegide(ByVal ligne As Object, Optional ByVal motsPage As Collection = Nothing) As Collection
    Dim resultat As New Collection, mots As Collection, mot As Object, colonne As Object
    Dim i As Long, genre As String, suivant As String, k As Long, dernier As Long, essai As String
    Dim voisin As Object, centre As Double, ecart As Double, distance As Double, norme As String
    Dim v As Long, z As Long, suite As Object, fragment As String, typeEssai As String, bord As Double
    Dim bas As Double, verticale As Object
    Set mots = ligne("m")
    For i = 1 To mots.Count
        Set mot = mots(i)
        bas = mot("y") + mot("h")
        genre = "": essai = "": dernier = i
        ' Chercher le plus long en-tete contigu, jamais des mots de colonnes eloignees.
        For k = i To Application.Min(i + 5, mots.Count)
            If k > i Then
                If mots(k)("x") - (mots(k - 1)("x") + mots(k - 1)("w")) > 2 Then Exit For
            End If
            essai = essai & CStr(mots(k)("t"))
            typeEssai = TypeColonneSpatialCegide(essai)
            If typeEssai <> "" Then
                genre = typeEssai
                dernier = k
            End If
        Next k
        centre = (mot("x") + mots(dernier)("x") + mots(dernier)("w")) / 2
        norme = NormaliserCompactCegide(CStr(mot("t")))
        If (genre = "" Or genre = "IGNORE") And Not motsPage Is Nothing Then
            If norme = "H" Or norme = "HORS" Or norme = "NON" Or norme = "N" Or norme = "BASE" Or norme = "BASES" Or norme = "MONTANT" Or norme = "MONTANTS" Or norme = "MT" Or norme = "MTS" Or norme = "PRIX" Then
                ecart = 1000
                For v = 1 To motsPage.Count
                    Set voisin = motsPage(v)
                    If voisin("p") = mot("p") Then
                        distance = Abs((voisin("x") + voisin("w") / 2) - (mot("x") + mot("w") / 2))
                        If voisin("y") > mot("y") + mot("h") * 0.4 And voisin("y") < mot("y") + mot("h") * 2.2 Then
                            If distance < Application.Max(mot("w"), voisin("w")) / 2 + 0.6 And distance < ecart Then
                                fragment = "": bord = voisin("x")
                                ' Completer uniquement des mots contigus sur la meme ligne inferieure.
                                For z = v To Application.Min(v + 3, motsPage.Count)
                                    Set suite = motsPage(z)
                                    If suite("p") <> voisin("p") Then Exit For
                                    If Abs(suite("y") - voisin("y")) > Application.Min(suite("h"), voisin("h")) * 0.5 Then Exit For
                                    If suite("x") < bord - 0.2 Or suite("x") - bord > 2 Then Exit For
                                    fragment = fragment & CStr(suite("t"))
                                    bord = suite("x") + suite("w")
                                    essai = CStr(mot("t")) & fragment
                                    typeEssai = TypeColonneSpatialCegide(essai)
                                    If typeEssai = "HT" Or typeEssai = "NS" Then
                                        genre = typeEssai
                                        ecart = distance
                                        centre = (voisin("x") + bord) / 2
                                        bas = Application.Max(bas, suite("y") + suite("h"))
                                    End If
                                Next z
                            End If
                        End If
                    End If
                Next v
            End If
        End If
        If (genre = "" Or genre = "IGNORE") And Not motsPage Is Nothing Then
            Set verticale = EnteteVerticaleCegide(mot, motsPage)
            If Not verticale Is Nothing Then
                genre = verticale("type"): centre = verticale("x"): bas = verticale("bas")
            End If
        End If
        If genre <> "" Then
            Set colonne = CreateObject("Scripting.Dictionary")
            colonne.Add "type", genre
            colonne.Add "x", centre
            colonne.Add "bas", bas
            i = dernier
            resultat.Add colonne
        End If
    Next i
    Set ColonnesPositionsCegide = resultat
End Function

Private Function EnteteVerticaleCegide(ByVal ancre As Object, ByVal mots As Collection) As Object
    Dim dernier As Object, suivant As Object, mot As Object, resultat As Object
    Dim texte As String, norme As String, genre As String, distance As Double, meilleure As Double
    Dim passage As Long, dx As Double, dy As Double
    norme = NormaliserCompactCegide(CStr(ancre("t")))
    Select Case norme
        Case "BASE", "BASES", "MONTANT", "MONTANTS", "MT", "HORS", "NON", "H", "N"
        Case Else: Exit Function
    End Select
    Set dernier = ancre: texte = CStr(ancre("t"))
    For passage = 1 To 2
        Set suivant = Nothing: meilleure = 1000
        For Each mot In mots
            If mot("p") = ancre("p") Then
                dy = mot("y") - dernier("y")
                dx = Abs(mot("x") + mot("w") / 2 - ancre("x") - ancre("w") / 2)
                If dy > dernier("h") * 0.5 And dy <= dernier("h") * 2.2 And dx <= Application.Max(ancre("w"), mot("w")) / 2 Then
                    norme = NormaliserCompactCegide(CStr(mot("t")))
                    Select Case norme
                        Case "HORS", "TAXE", "TAXES", "HT", "TVA", "T", "SOUMIS", "SOUMISE", "SOUMISES", "S"
                            distance = dy + dx
                            If distance < meilleure Then
                                meilleure = distance: Set suivant = mot
                            End If
                    End Select
                End If
            End If
        Next mot
        If suivant Is Nothing Then Exit Function
        texte = texte & " " & suivant("t")
        genre = TypeColonneSpatialCegide(texte)
        If genre = "HT" Or genre = "NS" Then
            Set resultat = CreateObject("Scripting.Dictionary")
            resultat.Add "type", genre
            resultat.Add "x", ancre("x") + ancre("w") / 2
            resultat.Add "bas", suivant("y") + suivant("h")
            Set EnteteVerticaleCegide = resultat
            Exit Function
        End If
        Set dernier = suivant
    Next passage
End Function

Private Function ControleTotalLigneCegide(ByVal colonnes As Collection, ByRef cellules() As String, ByVal libelle As String, Optional ByRef coherence As Long = 0) As String
    Dim j As Long, genre As String, trouve As Boolean, valeur As Double, montants As Object
    Dim ht As Double, tva As Double, ns As Double, ttc As Double, ecart As Double
    coherence = 0
    Set montants = CreateObject("Scripting.Dictionary")
    For j = 1 To colonnes.Count
        genre = CStr(colonnes(j)("type"))
        If genre = "HT" Or genre = "TVA" Or genre = "NS" Or genre = "TTC" Then
            If montants.Exists(genre) Then Exit Function
            If Trim$(cellules(j)) <> "" Then
                valeur = MontantColonneCegide(cellules(j), trouve, True)
                ' Controle uniquement sur les cellules completes, jamais sur un montant ambigu.
                If Not trouve Then Exit Function
                montants.Add genre, valeur
            ElseIf genre = "NS" Then
                montants.Add genre, 0#
            End If
        End If
    Next j
    If Not montants.Exists("HT") Or Not montants.Exists("TVA") Or Not montants.Exists("TTC") Then Exit Function
    ht = montants("HT"): tva = montants("TVA"): ttc = montants("TTC")
    If montants.Exists("NS") Then ns = montants("NS")
    ecart = ht + tva + ns - ttc
    If Abs(ecart) > 0.021 Then
        coherence = -1
        ControleTotalLigneCegide = "Coherence a verifier : HT " & ht & " + TVA " & tva & " + debours " & ns & " <> TTC " & ttc & " (ecart " & Format$(ecart, "0.00") & "). Montants lus conserves : " & libelle
    Else
        coherence = 1
    End If
End Function

Private Function NormaliserTexteMontantCegide(ByVal valeur As String) As String
    Dim texte As String
    texte = Replace(Replace(valeur, ChrW(160), " "), ChrW(8239), " ")
    texte = Trim$(Replace(texte, ChrW(8722), "-"))
    If Right$(texte, 1) = ChrW(8364) Then
        texte = Trim$(Left$(texte, Len(texte) - 1))
    ElseIf UCase$(Right$(texte, 3)) = "EUR" Then
        texte = Trim$(Left$(texte, Len(texte) - 3))
    End If
    Do While InStr(texte, "  ") > 0
        texte = Replace(texte, "  ", " ")
    Loop
    NormaliserTexteMontantCegide = texte
End Function

Private Function MontantColonneCegide(ByVal valeur As String, ByRef trouve As Boolean, Optional ByVal zeroExplicite As Boolean = False) As Double
    Dim reg As Object, normalise As String
    trouve = False
    normalise = NormaliserTexteMontantCegide(valeur)
    Set reg = CreateObject("VBScript.RegExp")
    If zeroExplicite Then
        reg.Pattern = "^[+-]? *0+( *[.,] *0+)?$"
        If reg.Test(normalise) Then
            trouve = True
            Exit Function
        End If
    End If
    ' Valider toute la cellule avant de retirer les espaces : ne pas fusionner deux montants.
    reg.Pattern = "^[+-]? *([0-9]{1,3}( [0-9]{3})+|[0-9]+)( *[.,] *[0-9]{1,2})?$"
    If Not reg.Test(normalise) Then Exit Function
    normalise = Replace(normalise, " ", "")
    MontantColonneCegide = Val(Replace(normalise, ",", "."))
    trouve = True
End Function

Private Function MontantAmbiguCegide(ByVal valeur As String, ByRef trouve As Boolean) As Double
    Dim reg As Object, candidats As Object, normalise As String
    trouve = False
    normalise = NormaliserTexteMontantCegide(valeur)
    Set reg = CreateObject("VBScript.RegExp")
    ' Priorite au premier montant decimal complet de la cellule, sans addition des candidats.
    reg.Pattern = "(^| )([+-]?([0-9]{1,3}( [0-9]{3})+|[0-9]+) *[.,] *[0-9]{1,2})($| )"
    Set candidats = reg.Execute(normalise)
    If candidats.Count = 0 Then
        reg.Pattern = "(^| )([+-]?[0-9]+)($| )"
        Set candidats = reg.Execute(normalise)
    End If
    If candidats.Count > 0 Then MontantAmbiguCegide = MontantColonneCegide(CStr(candidats(0).SubMatches(1)), trouve)
End Function

Private Sub CompleterColonneNSCegide(ByVal colonnes As Collection, ByVal lignes As Collection, ByVal entete As Object, ByRef probleme As String)
    Dim col As Object, ligne As Object, mot As Object, groupes As New Collection, groupe As Object, retenu As Object
    Dim xTVA As Double, xTTC As Double, x As Double, connu As Boolean, trouve As Boolean, montant As Double
    Dim texte As String, i As Long, candidats As Long, selection As Object
    For Each col In colonnes
        If col("type") = "NS" Then Exit Sub
        If col("type") = "TVA" Then xTVA = col("x")
        If col("type") = "TTC" Then xTTC = col("x")
    Next col
    If xTVA = 0 Or xTTC <= xTVA Then Exit Sub
    ' Deux lignes explicitement libellees "debours" doivent designer la meme colonne.
    ' Aucun montant n'est calcule par difference avec le TTC.
    For Each ligne In lignes
        If ligne("p") = entete("p") And ligne("y") > entete("y") + entete("h") Then
            texte = NormaliserCompactCegide(TextePositionCegide(ligne))
            If EstFinTableauFlexibleCegide(texte) Or Left$(texte, 5) = "SOLDE" Then Exit For
            If EstEnteteTableauFlexibleCegide(TextePositionCegide(ligne)) Then Exit For
            If InStr(texte, "DEBOUR") > 0 Then
                For Each mot In ligne("m")
                    x = mot("x") + mot("w") / 2
                    If x > xTVA + 2 And x < xTTC - 2 Then
                        montant = MontantColonneCegide(CStr(mot("t")), trouve)
                        connu = False
                        For Each col In colonnes
                            If Abs(x - col("x")) < 2 Then connu = True
                        Next col
                        If trouve And Not connu And montant > 0 Then
                            Set retenu = Nothing
                            For Each groupe In groupes
                                If Abs(x - groupe("x")) < 1.2 Then Set retenu = groupe
                            Next groupe
                            If retenu Is Nothing Then
                                Set retenu = CreateObject("Scripting.Dictionary")
                                retenu.Add "x", x
                                retenu.Add "n", 0
                                retenu.Add "y", -1000#
                                groupes.Add retenu
                            End If
                            If Abs(ligne("y") - retenu("y")) > ligne("h") Then
                                retenu("n") = retenu("n") + 1
                                retenu("y") = ligne("y")
                            End If
                        End If
                    End If
                Next mot
            End If
        End If
    Next ligne
    For Each groupe In groupes
        If groupe("n") >= 2 Then
            candidats = candidats + 1
            Set selection = groupe
        End If
    Next groupe
    If candidats <> 1 Then Exit Sub
    Set col = CreateObject("Scripting.Dictionary")
    col.Add "type", "NS"
    col.Add "x", selection("x")
    For i = 1 To colonnes.Count
        If colonnes(i)("x") > col("x") Then Exit For
    Next i
    colonnes.Add col, Before:=i
    probleme = AjouterProblemeCegide(probleme, "En-tete NS absent : colonne reperee par " & selection("n") & " lignes explicites de debours alignees, page " & entete("p") & ". A verifier.")
End Sub

Private Function ArticleSpatialCegide(ByVal libelle As String, ByVal articles As Collection, ByRef ambigu As Boolean) As Object
    Dim normaux As New Collection, article As Object, choisi As Object, mots As Variant, mot As Variant
    Dim cible As String, score As Long, meilleur As Long, operation As String
    For Each article In articles
        If Right$(CStr(article("Code")), 1) <> "2" Then normaux.Add article
    Next article
    ' Le moteur existant fournit les rapprochements par abreviations ; toute egalite est refusee.
    Set choisi = TrouverArticleCegide(libelle, libelle, normaux, ambigu)
    If Not choisi Is Nothing Then
        If Right$(CStr(choisi("Code")), 1) <> "2" And Not ambigu Then Set ArticleSpatialCegide = choisi
        Exit Function
    End If
    cible = " " & NormaliserEspacesCegide(libelle) & " "
    For Each article In normaux
        score = 0
        mots = Split(NormaliserEspacesCegide(CStr(article("Operation"))), " ")
        For Each mot In mots
            If Len(CStr(mot)) >= 4 Then
                If InStr(1, cible, " " & CStr(mot) & " ", vbTextCompare) > 0 Then score = score + 1
            End If
        Next mot
        If score > meilleur Then
            meilleur = score
            Set choisi = article
            ambigu = False
        ElseIf score > 0 And score = meilleur Then
            If CStr(choisi("Code")) <> CStr(article("Code")) Then ambigu = True
        End If
    Next article
    If meilleur > 0 And Not ambigu Then Set ArticleSpatialCegide = choisi
End Function

Private Sub AjouterOperationSpatialeCegide(ByVal sorties As Collection, ByVal articles As Collection, ByVal libelle As String, ByVal ht As Double, ByVal ns As Double, ByVal trouveHT As Boolean, ByVal trouveNS As Boolean, ByRef probleme As String, Optional ByVal ambiguHT As Boolean = False, Optional ByVal ambiguNS As Boolean = False, Optional ByVal repliTTC As Boolean = False, Optional ByVal coherence As Long = 0)
    Dim article As Object, info As Object, ambigu As Boolean, code As String, genre As Variant
    If Left$(libelle, 19) <> "Libelle non reconnu" Then Set article = ArticleSpatialCegide(libelle, articles, ambigu)
    If article Is Nothing And Not trouveHT And trouveNS And Left$(libelle, 19) <> "Libelle non reconnu" Then
        ambigu = False
        Set article = TrouverArticleCegide(libelle, libelle, articles, ambigu)
        If ambigu Then Set article = Nothing
    End If
    If article Is Nothing Then
        probleme = AjouterProblemeCegide(probleme, "Article absent ou ambigu, ligne conservee sans code : " & libelle)
        Set article = CreateObject("Scripting.Dictionary")
        article.Add "Operation", libelle
        article.Add "TypeFournisseur", ""
        article.Add "Article", ""
        article.Add "Code", ""
    End If
    If Not trouveHT And Not trouveNS And Not ambiguNS Then ambiguHT = True
    For Each genre In Array("HT", "NS")
        If (genre = "HT" And ((trouveHT And (ht <> 0# Or repliTTC)) Or ambiguHT)) Or (genre = "NS" And ((trouveNS And ns <> 0#) Or ambiguNS)) Then
            Set info = CreateObject("Scripting.Dictionary")
            info.Add "OperationReference", CStr(article("Operation"))
            info.Add "TexteDetecte", libelle
            info.Add "TypeFournisseur", CStr(article("TypeFournisseur"))
            info.Add "ArticleReference", CStr(article("Article"))
            info.Add "Code", CStr(article("Code"))
            info.Add "MontantHT", ht
            info.Add "Debours", 0#
            info.Add "DeboursSeulement", False
            info.Add "MontantTrouve", trouveHT
            info.Add "BaseMontant", "HT"
            info.Add "PresenceHT", Not repliTTC
            info.Add "Coherence", coherence
            If repliTTC Then info("BaseMontant") = "TTC"
            If genre = "NS" Then
                code = TrouverCodeDeboursCegide(articles, info)
                If code = "" Then
                    probleme = AjouterProblemeCegide(probleme, "Code terminant par 2 absent : " & libelle)
                    ' Conserver le montant et le libelle, sans inventer de code.
                End If
                info("Code") = code
                info.Add "CodeDebours", code
                info("MontantHT") = 0#
                info("Debours") = ns
                info("DeboursSeulement") = True
                info("BaseMontant") = "NS"
                info("MontantTrouve") = trouveNS
            End If
            sorties.Add info
        End If
    Next genre
    If Not trouveHT And Not trouveNS Then probleme = AjouterProblemeCegide(probleme, "Montant absent ou ambigu, ligne conservee : " & libelle)
End Sub

Private Function AnalyserPositionsCegide(ByVal mots As Collection, ByVal articles As Collection, ByRef probleme As String) As Collection
    Dim resultat As New Collection, lignes As Collection, ligne As Object, colonnes As Collection, entete As Collection
    Dim mot As Object, texte As String, norme As String, libelle As String, precedent As String
    Dim ht As Double, ns As Double, aHT As Boolean, aNS As Boolean, nombre As Double, trouve As Boolean
    Dim j As Long, indexCol As Long, distance As Double, minimum As Double, centre As Double
    Dim limite As Double, debut As Double, page As Long, actif As Boolean, candidates As Long, traitees As Long, ancienY As Double
    Dim ligneHT As Boolean, ligneNS As Boolean, ligneMontant As Boolean, sommeHT As Double, sommeNS As Double
    Dim cellules() As String, regNombre As Object
    Dim colonneHT As Boolean, erreurHT As Boolean, erreurNS As Boolean, aErreurHT As Boolean, aErreurNS As Boolean, colonneNS As Boolean, colonneTTC As Boolean
    Dim basEntete As Double, controle As String
    Dim coherence As Long, coherenceCourante As Long
    Dim montantsPrecedents As Boolean, montantsLigne As Boolean
    Set regNombre = CreateObject("VBScript.RegExp")
    regNombre.Pattern = "^[+0-9., -]+$"
    Set lignes = LignesPositionsCegide(mots)
    For Each ligne In lignes
        texte = TextePositionCegide(ligne)
        norme = NormaliserCompactCegide(texte)
        Set entete = ColonnesPositionsCegide(ligne, mots)
        If (EstEnteteTableauFlexibleCegide(texte) Or InStr(1, norme, "NATURE", vbTextCompare) > 0 Or InStr(1, norme, "LIBELLE", vbTextCompare) > 0) And entete.Count >= 2 Then
            If precedent <> "" Then
                AjouterOperationSpatialeCegide resultat, articles, precedent, ht, ns, aHT, aNS, probleme, aErreurHT, aErreurNS, Not colonneHT, coherence
                traitees = traitees + 1
                precedent = ""
            End If
            Set colonnes = entete
            basEntete = ligne("y") + ligne("h")
            For j = 1 To colonnes.Count
                If colonnes(j).Exists("bas") Then basEntete = Application.Max(basEntete, colonnes(j)("bas"))
            Next j
            CompleterColonneNSCegide colonnes, lignes, ligne, probleme
            actif = True
            page = ligne("p")
            limite = colonnes(1)("x") - 5
            If colonnes(1)("type") = "HT" Then limite = colonnes(1)("x") - 2
            debut = 0
            For Each mot In ligne("m")
                Select Case NormaliserCompactCegide(CStr(mot("t")))
                    Case "NATURE": debut = mot("x") - 0.5
                    Case "LIBELLE", "DESCRIPTION", "DESIGNATION": If debut = 0 Then debut = mot("x") - 0.5
                    Case "DATE": debut = mot("x") + mot("w") + 1
                End Select
            Next mot
            trouve = False
            colonneNS = False: colonneTTC = False
            For j = 1 To colonnes.Count
                If colonnes(j)("type") = "HT" Then trouve = True
                If colonnes(j)("type") = "NS" Then colonneNS = True
                If colonnes(j)("type") = "TTC" Then colonneTTC = True
            Next j
            If Not colonneNS And Not colonneTTC Then probleme = AjouterProblemeCegide(probleme, "Tableau partiel page " & page & " : colonnes debours/NS et TTC non reconnues. Verifier les zones masquees ou illisibles.")
            colonneHT = trouve
            If Not colonneHT Then probleme = AjouterProblemeCegide(probleme, "En-tete HT absent page " & page & " : repli sur TTC si disponible, a verifier (Debit toujours exclu).")
        ElseIf actif And ligne("y") >= basEntete - ligne("h") * 0.3 Then
            If ligne("p") <> page Or EstFinTableauFlexibleCegide(norme) Or Left$(norme, 5) = "SOLDE" Or InStr(norme, "MONTANTHTMONTANTTVA") > 0 Or InStr(norme, "TAUX") > 0 Or (Left$(norme, 3) = "TVA" And InStr(norme, "HT") > 0) Then
                If precedent <> "" Then
                    AjouterOperationSpatialeCegide resultat, articles, precedent, ht, ns, aHT, aNS, probleme, aErreurHT, aErreurNS, Not colonneHT, coherence
                    traitees = traitees + 1
                    precedent = ""
                End If
                actif = False
            Else
                libelle = "": ligneHT = False: ligneNS = False: ligneMontant = False: sommeHT = 0: sommeNS = 0
                erreurHT = False: erreurNS = False
                ReDim cellules(1 To colonnes.Count)
                For Each mot In ligne("m")
                    centre = mot("x") + mot("w") / 2
                    If centre < limite Then
                        If mot("x") >= debut Then
                            If colonneHT Or (Not regNombre.Test(Trim$(CStr(mot("t")))) And Trim$(CStr(mot("t"))) <> "%") Then libelle = libelle & " " & mot("t")
                        End If
                    ElseIf centre <= colonnes(colonnes.Count)("x") + 3 Then
                        minimum = 1000: indexCol = 0
                        For j = 1 To colonnes.Count
                            distance = Abs(centre - colonnes(j)("x"))
                            If distance < minimum Then minimum = distance: indexCol = j
                        Next j
                        If indexCol > 0 And regNombre.Test(NormaliserTexteMontantCegide(CStr(mot("t")))) Then
                            cellules(indexCol) = Trim$(cellules(indexCol) & " " & NormaliserTexteMontantCegide(CStr(mot("t"))))
                        End If
                    End If
                Next mot
                For j = 1 To colonnes.Count
                    If cellules(j) <> "" And cellules(j) <> "-" Then
                        nombre = MontantColonneCegide(cellules(j), trouve, colonnes(j)("type") = "NS" Or colonnes(j)("type") = "HT")
                        If Not trouve Then
                            If colonnes(j)("type") = "HT" Or colonnes(j)("type") = "NS" Or (colonnes(j)("type") = "TTC" And Not colonneHT) Then
                                nombre = MontantAmbiguCegide(cellules(j), trouve)
                                If trouve Then probleme = AjouterProblemeCegide(probleme, "Montant ambigu retenu en colonne " & colonnes(j)("type") & " : " & CStr(nombre) & " ; cellule OCR : " & cellules(j))
                            End If
                        End If
                        If trouve Then
                            If colonnes(j)("type") <> "IGNORE" Then ligneMontant = True
                            Select Case colonnes(j)("type")
                                Case "HT": sommeHT = sommeHT + nombre: ligneHT = True
                                Case "NS": sommeNS = sommeNS + nombre: ligneNS = True
                                Case "TTC"
                                    If Not colonneHT Then sommeHT = sommeHT + nombre: ligneHT = True
                                ' Le TTC ne remplace jamais une cellule HT vide ou ambigue.
                            End Select
                        ElseIf colonnes(j)("type") = "HT" Or colonnes(j)("type") = "NS" Or (colonnes(j)("type") = "TTC" And Not colonneHT) Then
                            ligneMontant = True
                            If colonnes(j)("type") = "NS" Then erreurNS = True Else erreurHT = True
                            probleme = AjouterProblemeCegide(probleme, "Montant fragmente ou ambigu en colonne " & colonnes(j)("type") & ", page " & page & " : " & cellules(j))
                        End If
                    End If
                Next j
                libelle = Trim$(SupprimerDatesCegide(Trim$(libelle)))
                controle = ControleTotalLigneCegide(colonnes, cellules, libelle, coherenceCourante)
                If controle <> "" Then probleme = AjouterProblemeCegide(probleme, controle)
                montantsLigne = ligneMontant
                If Not ligneMontant And libelle <> "" Then
                    If PositionMotifCegide(texte, "[0-9]{2}[./-][0-9]{2}[./-][0-9]{4}") > 0 Then
                        ligneMontant = True: erreurHT = True
                    ElseIf precedent = "" And Left$(NormaliserCompactCegide(libelle), 6) <> "DETAIL" Then
                        ligneMontant = True: erreurHT = True
                    End If
                End If
                If ligneMontant Then
                  ' Une ligne purement numerique proche peut completer le libelle encore sans montants.
                  If libelle = "" And precedent <> "" And Not montantsPrecedents _
                      And ligne("y") > ancienY And ligne("y") - ancienY < ligne("h") * 2.5 _
                      And PositionMotifCegide(texte, "[0-9]{2}[./-][0-9]{2}[./-][0-9]{4}") = 0 Then
                    ht = sommeHT: ns = sommeNS: aHT = ligneHT: aNS = ligneNS
                    aErreurHT = erreurHT: aErreurNS = erreurNS
                    coherence = coherenceCourante: montantsPrecedents = True
                    ancienY = ligne("y")
                    probleme = AjouterProblemeCegide(probleme, "Montants decales rattaches au libelle precedent, a verifier : " & precedent)
                  Else
                    If precedent <> "" Then
                        AjouterOperationSpatialeCegide resultat, articles, precedent, ht, ns, aHT, aNS, probleme, aErreurHT, aErreurNS, Not colonneHT, coherence
                        traitees = traitees + 1
                    End If
                    precedent = libelle: ht = sommeHT: ns = sommeNS: aHT = ligneHT: aNS = ligneNS
                    aErreurHT = erreurHT: aErreurNS = erreurNS
                    coherence = coherenceCourante
                    montantsPrecedents = montantsLigne
                    candidates = candidates + 1
                    ancienY = ligne("y")
                    If precedent = "" Then
                        precedent = "Libelle non reconnu (page " & page & ", position " & Format$(ligne("y"), "0.0") & ")"
                        probleme = AjouterProblemeCegide(probleme, "Ligne avec montant sans libelle conservee, page " & page & " position " & Format$(ligne("y"), "0.0"))
                    End If
                  End If
                ElseIf libelle <> "" And precedent <> "" Then
                    If ligne("y") - ancienY < ligne("h") * 2.5 Then
                        precedent = precedent & " " & libelle
                        ancienY = ligne("y")
                    End If
                End If
            End If
        End If
    Next ligne
    If precedent <> "" Then
        AjouterOperationSpatialeCegide resultat, articles, precedent, ht, ns, aHT, aNS, probleme, aErreurHT, aErreurNS, Not colonneHT, coherence
        traitees = traitees + 1
    End If
    If candidates = 0 Then probleme = AjouterProblemeCegide(probleme, "Aucune ligne de tableau spatial reconnue : verifier le PDF et le TSV.")
    If candidates <> traitees Then probleme = AjouterProblemeCegide(probleme, CStr(candidates - traitees) & " ligne(s) spatiale(s) sans libelle exploitable.")
    Set AnalyserPositionsCegide = resultat
End Function



' Interface appelee par CegidAutomatique.ExecuterDerniereFactureGloire.
' Ces noms sont conserves pour que les deux modules correspondent.
Public Sub Gloire_Configurer_Automatique(ByVal actif As Boolean)
    If actif Then messagesRemplissageAutomatique = ""
End Sub

Public Function Gloire_Lire_Message_Boucle() As String
    Gloire_Lire_Message_Boucle = messagesRemplissageAutomatique
End Function

Private Sub Gloire_MessageAutomatique(ByVal texte As String, Optional ByVal style As VbMsgBoxStyle = vbOKOnly, Optional ByVal titre As String = "Remplissage Cegide")
    If Len(messagesRemplissageAutomatique) > 0 Then messagesRemplissageAutomatique = messagesRemplissageAutomatique & vbCrLf & vbCrLf
    messagesRemplissageAutomatique = messagesRemplissageAutomatique & texte
    Application.StatusBar = Left$(texte, 220)
End Sub
