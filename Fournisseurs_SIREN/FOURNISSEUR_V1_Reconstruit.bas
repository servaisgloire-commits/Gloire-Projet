Option Explicit

' ==========================================================
' FOURNISSEUR_V1 - version reconstruite depuis les archives
' ==========================================================
' Entree  : texte de recherche en colonne H
' Sorties : B = SIREN
'           C = date de cessation
'           D = raison sociale
'           E = statut
' Source  : API Recherche Entreprises (api.gouv.fr)
' ==========================================================

Public Sub FOURNISSEUR_V1()
    RechercherInfosEntreprise
End Sub

Public Sub RechercherInfosEntreprise()

    Dim ws As Worksheet
    Dim zone As Range
    Dim ligne As Range
    Dim texteRecherche As String
    Dim siren As String
    Dim dateCessation As String
    Dim raisonSociale As String
    Dim statut As String
    Dim trouve As Boolean
    Dim nbTraitees As Long

    Set ws = ActiveSheet
    Application.ScreenUpdating = False
    On Error GoTo FinAvecErreur

    For Each zone In Selection.Areas
        For Each ligne In zone.Rows

            texteRecherche = Trim$(CStr(ws.Cells(ligne.Row, "H").Value))
            ws.Range("B" & ligne.Row & ":E" & ligne.Row).ClearContents

            If texteRecherche <> "" Then

                ws.Cells(ligne.Row, "E").Value = "Recherche."

                siren = ""
                dateCessation = ""
                raisonSociale = ""
                statut = ""

                trouve = ChercherInfosDataGouv(texteRecherche, siren, dateCessation, raisonSociale, statut)

                If trouve Then
                    ws.Cells(ligne.Row, "B").Value = ValeurOuAucune(siren)
                    ws.Cells(ligne.Row, "C").Value = ValeurOuAucune(dateCessation)
                    ws.Cells(ligne.Row, "D").Value = ValeurOuAucune(raisonSociale)
                    ws.Cells(ligne.Row, "E").Value = ValeurOuAucune(statut)
                Else
                    ws.Range("B" & ligne.Row & ":E" & ligne.Row).ClearContents
                End If

                nbTraitees = nbTraitees + 1
                DoEvents
                PauseCourte 0.2

            End If

        Next ligne
    Next zone

FinPropre:
    Application.ScreenUpdating = True
    MsgBox nbTraitees & " ligne(s) traitee(s).", vbInformation
    Exit Sub

FinAvecErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbExclamation

End Sub

Private Function ChercherInfosDataGouv( _
    ByVal recherche As String, _
    ByRef siren As String, _
    ByRef dateCessation As String, _
    ByRef raisonSociale As String, _
    ByRef statut As String) As Boolean

    Dim http As Object
    Dim url As String
    Dim reponse As String
    Dim etat As String

    On Error GoTo ErreurMacro

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

    If http.Status <> 200 Then
        siren = "Aucune"
        dateCessation = "Aucune"
        raisonSociale = "Aucune"
        statut = "Erreur HTTP " & http.Status
        ChercherInfosDataGouv = True
        Exit Function
    End If

    reponse = http.responseText

    If AucunResultat(reponse) Then
        ChercherInfosDataGouv = False
        Exit Function
    End If

    siren = ExtraireValeurJson(reponse, "siren")
    dateCessation = ExtraireValeurJson(reponse, "date_fermeture")
    raisonSociale = ExtraireValeurJson(reponse, "nom_raison_sociale")

    If raisonSociale = "" Then
        raisonSociale = ExtraireValeurJson(reponse, "nom_complet")
    End If

    etat = ExtraireValeurJson(reponse, "etat_administratif")

    Select Case UCase$(etat)
        Case "A"
            statut = "Active"
        Case "C", "F"
            statut = "Fermee / cessee"
        Case Else
            statut = ""
    End Select

    ChercherInfosDataGouv = True
    Exit Function

ErreurMacro:
    siren = "Aucune"
    dateCessation = "Aucune"
    raisonSociale = "Aucune"
    statut = "Erreur macro"
    ChercherInfosDataGouv = True

End Function

Private Function AucunResultat(ByVal json As String) As Boolean

    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = """results""\s*:\s*\[\s*\]"

    AucunResultat = regex.Test(json)

End Function

Private Function ExtraireValeurJson(ByVal json As String, ByVal cle As String) As String

    Dim regex As Object
    Dim matches As Object
    Dim valeur As String
    Dim motif As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True

    motif = """" & cle & """\s*:\s*(""([^""]*)""|true|false|null|-?\d+)"
    regex.Pattern = motif

    If regex.Test(json) Then
        Set matches = regex.Execute(json)
        valeur = matches(0).SubMatches(0)
        valeur = Replace(valeur, """", "")

        If LCase$(valeur) = "null" Then
            valeur = ""
        End If

        ExtraireValeurJson = valeur
    Else
        ExtraireValeurJson = ""
    End If

End Function

Private Function ValeurOuAucune(ByVal valeur As String) As String

    valeur = Trim$(valeur)

    If valeur = "" Then
        ValeurOuAucune = "Aucune"
    Else
        ValeurOuAucune = valeur
    End If

End Function

Private Function UrlEncode(ByVal texte As String) As String

    On Error GoTo Secours

    UrlEncode = Application.WorksheetFunction.EncodeURL(texte)
    Exit Function

Secours:
    UrlEncode = Replace(texte, " ", "%20")
    UrlEncode = Replace(UrlEncode, "&", "%26")
    UrlEncode = Replace(UrlEncode, "+", "%2B")
    UrlEncode = Replace(UrlEncode, "#", "%23")

End Function

Private Sub PauseCourte(ByVal secondes As Double)

    Dim depart As Double

    depart = Timer

    Do While Timer < depart + secondes
        DoEvents
    Loop

End Sub
