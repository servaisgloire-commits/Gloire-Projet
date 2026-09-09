Option Explicit

' ==========================================================
' GLOIRE V2
' Blocs principaux de la logique de traitement
' ==========================================================

Private Const NOM_FEUILLE_STATS As String = "STATISTIQUES"
Private Const NOM_TABLE_STATS As String = "TableauStatistiquesGloireV2"
Private Const NOM_FEUILLE_FOURNISSEURS As String = "FOURNISSEURS"
Private Const NOM_FEUILLE_MEMOIRE_FORMATS As String = "MEMOIRE_FORMATS"
Private Const NOM_FEUILLE_ARRET As String = "ARRETER_TRAITEMENT"
Private Const NOM_FEUILLE_PONT_OCR As String = "PON OCR"
Private Const NOM_FICHIER_PONT_OCR As String = "WinOcrBridge.exe"
Private Const SOUS_DOSSIER_TEXTES As String = "Textes OCR"
Private Const OCR_MAX_PAGES As Long = 0
Private Const RAYON_LIGNES_OCR As Long = 5
Private Const DECISION_VALIDER As String = "VALIDER"
Private Const DECISION_CORRIGER As String = "CORRIGER"
Private Const DECISION_REPORTER As String = "REPORTER"

Private DerniereErreurOCR As String
Private ArretTraitementDemande As Boolean
Private FenetresEdgeEnumerees As Object

Public Sub Gloire_V2_Tester_OCR()
    ExecuterTestOCRGloireV2
End Sub

Public Sub Gloire_V2_Configurer_Pont_OCR()
    Dim fso As Object
    Dim chemin As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    chemin = RecupererCheminPontOCR(fso, True)

    If chemin <> "" Then
        MsgBox "Chemin du pont OCR enregistré dans la feuille 'PON OCR' :" & _
               vbCrLf & vbCrLf & chemin, vbInformation
    End If
End Sub

Public Sub Gloire_V2_Classement_Factures()
    ExecuterClassementFacturesGloireV2
End Sub

Public Sub Gloire_V2_Initialiser_Memoire()
    Dim ws As Worksheet

    Set ws = PreparerFeuilleMemoireFormats(ThisWorkbook)
    ws.Visible = xlSheetVisible
    ws.Activate

    MsgBox "Feuille MEMOIRE_FORMATS préparée." & vbCrLf & vbCrLf & _
           "Une seule ligne est conservée par fournisseur." & vbCrLf & _
           "Une ligne existante est modifiée après une correction manuelle validée.", _
           vbInformation
End Sub

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

' ==========================================================
' CHAÎNE DE TRAITEMENT
' ==========================================================
'
' 1. Choix du dossier contenant les factures PDF.
' 2. Choix du dossier d'arrivée.
' 3. Chargement du référentiel fournisseurs.
' 4. Chargement de la mémoire des formats.
' 5. Vérification du pont OCR Windows.
' 6. Création de la liste des PDF à traiter.
' 7. Lecture OCR de chaque facture.
' 8. Détection du fournisseur.
' 9. Recherche du contrat ou de la référence client.
' 10. Recherche du numéro de facture.
' 11. Recherche de la date d'échéance.
' 12. Utilisation des formats fournisseur déjà validés.
' 13. Utilisation des règles textuelles si nécessaire.
' 14. Validation humaine en cas de doute.
' 15. Correction possible avec mise à jour de la mémoire.
' 16. Renommage du PDF.
' 17. Classement dans le dossier fournisseur.
' 18. Gestion des doublons.
' 19. Journalisation dans la feuille STATISTIQUES.
' 20. Les documents non validés restent dans le dossier de départ.
'
' Les autres fonctions du projet gèrent notamment :
' - l'OCR et l'extraction du texte ;
' - les expressions régulières ;
' - la détection des fournisseurs ;
' - les connecteurs textuels ;
' - la mémoire de formats ;
' - la validation utilisateur ;
' - la gestion des fichiers et des doublons ;
' - les statistiques de traitement.