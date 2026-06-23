Attribute VB_Name = "BOOTSTRAP_IMPORT"
Option Compare Database
Option Explicit

' =============================================================================
' RESERV+ - Module Bootstrap
' =============================================================================
' 1. Importer CE fichier seul dans Access (ALT+F11, CTRL+M)
' 2. CTRL+G puis taper :  LancerInstallationComplete
' 3. Entrer le chemin du dossier vba\ quand demande
' =============================================================================

Public Sub LancerInstallationComplete()
    Dim chemin As String
    chemin = InputBox( _
        "Entrez le chemin complet du dossier vba\" & vbCrLf & vbCrLf & _
        "Exemple :" & vbCrLf & _
        "C:\Users\abdourahamane.diallo\Documents\gestion-reservation-vba\RESERV_PLUS_ACCESS\vba", _
        "RESERV+ - Chemin des modules VBA")

    If Trim(chemin) = "" Then
        MsgBox "Installation annulee.", vbInformation
        Exit Sub
    End If

    ' Supprimer antislash final si present
    If Right(chemin, 1) = "\" Then chemin = Left(chemin, Len(chemin) - 1)

    ' Verifier que le dossier existe
    If Dir(chemin, vbDirectory) = "" Then
        MsgBox "Dossier introuvable :" & vbCrLf & chemin & vbCrLf & vbCrLf & _
               "Verifiez le chemin et recommencez.", vbCritical, "RESERV+ Erreur"
        Exit Sub
    End If

    ' Importer tous les modules
    ImporterModules chemin

    ' Lancer l'installation via Application.Run (evite l'erreur de compilation)
    Dim rep As Integer
    rep = MsgBox("30 modules importes avec succes !" & vbCrLf & vbCrLf & _
                 "Lancer maintenant Installer_RESERV_PLUS ?", _
                 vbYesNo + vbQuestion, "RESERV+")

    If rep = vbYes Then
        Application.Run "Installer_RESERV_PLUS"
    End If
End Sub

Private Sub ImporterModules(dossier As String)
    Dim vbe As Object
    Dim proj As Object
    Dim comp As Object
    Dim i As Integer
    Dim fichier As String
    Dim erreurs As String

    ' Liste dans l ordre obligatoire
    Dim modules(1 To 30) As String
    modules(1)  = "core\modConfig.bas"
    modules(2)  = "core\modDatabase.bas"
    modules(3)  = "core\modSession.bas"
    modules(4)  = "core\modSecurite.bas"
    modules(5)  = "core\modJournal.bas"
    modules(6)  = "core\modUtils.bas"
    modules(7)  = "metier\modClients.bas"
    modules(8)  = "metier\modPrestataires.bas"
    modules(9)  = "metier\modRessources.bas"
    modules(10) = "metier\modDisponibilite.bas"
    modules(11) = "metier\modReservations.bas"
    modules(12) = "metier\modPaiements.bas"
    modules(13) = "metier\modBillets.bas"
    modules(14) = "metier\modRapports.bas"
    modules(15) = "ui\modThemeAccess.bas"
    modules(16) = "ui\modNavigation.bas"
    modules(17) = "ui\modFormFactory.bas"
    modules(18) = "ui\modFormsConnexion.bas"
    modules(19) = "ui\modFormsDashboard.bas"
    modules(20) = "ui\modFormsClients.bas"
    modules(21) = "ui\modFormsPrestataires.bas"
    modules(22) = "ui\modFormsRessources.bas"
    modules(23) = "ui\modFormsReservations.bas"
    modules(24) = "ui\modFormsPaiements.bas"
    modules(25) = "ui\modFormsAdministration.bas"
    modules(26) = "ui\modFormsRapports.bas"
    modules(27) = "reports\modReportsFactory.bas"
    modules(28) = "reports\modEtatsImprimables.bas"
    modules(29) = "install\modCreationRequetes.bas"
    modules(30) = "install\modInstall_RESERV_PLUS.bas"

    Set vbe  = Application.VBE
    Set proj = vbe.ActiveVBProject

    ' Supprimer les modules standards existants sauf ce module
    For Each comp In proj.VBComponents
        If comp.Name <> "BOOTSTRAP_IMPORT" And comp.Type = 1 Then
            On Error Resume Next
            proj.VBComponents.Remove comp
            On Error GoTo 0
        End If
    Next comp

    erreurs = ""

    For i = 1 To 30
        fichier = dossier & "\" & modules(i)

        If Dir(fichier) = "" Then
            erreurs = erreurs & vbCrLf & "Manquant : " & modules(i)
        Else
            On Error Resume Next
            proj.VBComponents.Import fichier
            If Err.Number <> 0 Then
                erreurs = erreurs & vbCrLf & "Erreur : " & modules(i) & " (" & Err.Description & ")"
                Err.Clear
            End If
            On Error GoTo 0
        End If

        ' Progression dans la fenetre Execution
        If i Mod 5 = 0 Then Debug.Print "  Import " & i & "/30..."
    Next i

    If erreurs <> "" Then
        MsgBox "Avertissements lors de l'import :" & erreurs & vbCrLf & vbCrLf & _
               "Verifiez les fichiers manquants.", vbExclamation, "RESERV+"
    End If
End Sub
