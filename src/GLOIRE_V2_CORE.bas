Option Explicit

' ==========================================================
' GLOIRE_V2_CORE
' Extrait vérifié depuis l'archive Gloire V2 MAJ 7 / PON OCR.
' Ce fichier présente les points d'entrée et l'architecture
' principale du projet. Il n'est pas autonome sans le module complet.
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
        MsgBox "Chemin du pont OCR enregistre dans la feuille 'PON OCR' :" & _
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

    MsgBox "Feuille MEMOIRE_FORMATS preparee." & vbCrLf & vbCrLf & _
           "Une seule ligne est conservee par fournisseur." & vbCrLf & _
           "Une ligne existante est modifiee uniquement apres une correction manuelle validee.", _
           vbInformation
End Sub

Public Sub Gloire_V2_Arreter_Le_Traitement()
    ArretTraitementDemande = True
    Application.StatusBar = "Arret demande : la facture en cours restera dans le dossier de depart."
    MettreAJourEtatFeuilleArret ThisWorkbook, "ARRET DEMANDE", RGB(192, 0, 0)
    Beep
End Sub

' ==========================================================
' CHAINE DE TRAITEMENT ISSUE DU MODULE COMPLET
' ==========================================================
'
' 1. Choix du dossier de départ contenant les PDF.
' 2. Choix du dossier d'arrivée.
' 3. Chargement du référentiel fournisseurs et de la mémoire.
' 4. Vérification du pont OCR Windows.
' 5. Liste figée des PDF à traiter.
' 6. OCR local de chaque facture.
' 7. Détection du fournisseur.
' 8. Recherche du contrat / référence client.
' 9. Recherche du numéro de facture.
' 10. Recherche de la date d'échéance.
' 11. Utilisation prioritaire de la mémoire de formats fournisseur.
' 12. Utilisation des connecteurs textuels si la mémoire ne suffit pas.
' 13. Ouverture du PDF et validation humaine en cas de doute.
' 14. Correction manuelle possible avec mise à jour de la mémoire.
' 15. Renommage du PDF selon la nomenclature standardisée.
' 16. Création / sélection du dossier fournisseur.
' 17. Déplacement du fichier sans écrasement d'un doublon existant.
' 18. Journalisation du résultat dans la feuille STATISTIQUES.
' 19. Les documents non validés restent dans le dossier de départ.
'
' L'archive complète contient aussi les fonctions OCR, les expressions
' régulières, la gestion Edge/Win32, les dictionnaires fournisseurs,
' les connecteurs, la mémoire de formats, les statistiques, la gestion
' des doublons, le renommage et le déplacement des fichiers.
