Attribute VB_Name = "BOOTSTRAP_IMPORT"
' =============================================================================
' RESERV+ — Module Bootstrap (import unique)
' =============================================================================
' UTILISATION :
'   1. Importez UNIQUEMENT ce fichier dans Access (ALT+F11, CTRL+M)
'   2. Dans la fenêtre Exécution (CTRL+G), tapez :
'        ImporterTousLesModules "C:\chemin\vers\RESERV_PLUS_ACCESS\vba"
'   3. Puis : Installer_RESERV_PLUS
' =============================================================================

Public Sub ImporterTousLesModules(Optional DossierVBA As String = "")
    Dim vbe As Object
    Dim proj As Object
    Dim comp As Object
    Dim chemin As String
    Dim i As Integer

    ' --- Demander le chemin si non fourni ---
    If DossierVBA = "" Then
        DossierVBA = InputBox( _
            "Entrez le chemin complet du dossier vba\" & vbCrLf & _
            "Exemple : C:\RESERV_PLUS_ACCESS\vba", _
            "RESERV+ — Chemin des modules", _
            "C:\RESERV_PLUS_ACCESS\vba")
        If DossierVBA = "" Then
            MsgBox "Installation annulée.", vbInformation
            Exit Sub
        End If
    End If

    ' S'assurer que le chemin se termine sans antislash
    If Right(DossierVBA, 1) = "\" Then
        DossierVBA = Left(DossierVBA, Len(DossierVBA) - 1)
    End If

    ' --- Vérifier que le dossier existe ---
    If Dir(DossierVBA, vbDirectory) = "" Then
        MsgBox "Dossier introuvable :" & vbCrLf & DossierVBA & vbCrLf & _
               "Vérifiez le chemin et relancez.", vbCritical, "RESERV+ Erreur"
        Exit Sub
    End If

    Set vbe = Application.VBE
    Set proj = vbe.ActiveVBProject

    ' --- Liste des modules dans l'ordre obligatoire ---
    Dim modules(1 To 30) As String
    ' CORE
    modules(1)  = "core\modConfig.bas"
    modules(2)  = "core\modDatabase.bas"
    modules(3)  = "core\modSession.bas"
    modules(4)  = "core\modSecurite.bas"
    modules(5)  = "core\modJournal.bas"
    modules(6)  = "core\modUtils.bas"
    ' METIER
    modules(7)  = "metier\modClients.bas"
    modules(8)  = "metier\modPrestataires.bas"
    modules(9)  = "metier\modRessources.bas"
    modules(10) = "metier\modDisponibilite.bas"
    modules(11) = "metier\modReservations.bas"
    modules(12) = "metier\modPaiements.bas"
    modules(13) = "metier\modBillets.bas"
    modules(14) = "metier\modRapports.bas"
    ' UI
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
    ' REPORTS
    modules(27) = "reports\modReportsFactory.bas"
    modules(28) = "reports\modEtatsImprimables.bas"
    ' INSTALL
    modules(29) = "install\modCreationRequetes.bas"
    modules(30) = "install\modInstall_RESERV_PLUS.bas"

    ' --- Supprimer les modules déjà existants (évite les doublons) ---
    Dim nomModule As String
    For Each comp In proj.VBComponents
        If comp.Name <> "BOOTSTRAP_IMPORT" And _
           comp.Type = 1 Then  ' vbext_ct_StdModule = 1
            On Error Resume Next
            proj.VBComponents.Remove comp
            On Error GoTo 0
        End If
    Next comp

    ' --- Importer chaque module ---
    Dim erreurs As String
    erreurs = ""

    For i = 1 To 30
        chemin = DossierVBA & "\" & modules(i)

        If Dir(chemin) = "" Then
            erreurs = erreurs & vbCrLf & "Manquant : " & modules(i)
        Else
            On Error Resume Next
            proj.VBComponents.Import chemin
            If Err.Number <> 0 Then
                erreurs = erreurs & vbCrLf & "Erreur " & modules(i) & " : " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next i

    ' --- Résultat ---
    If erreurs = "" Then
        MsgBox "✔  30 modules importés avec succès !" & vbCrLf & vbCrLf & _
               "Tapez maintenant dans la fenêtre Exécution (CTRL+G) :" & vbCrLf & _
               "   Installer_RESERV_PLUS" & vbCrLf & vbCrLf & _
               "Puis appuyez sur ENTRÉE.", _
               vbInformation, "RESERV+ — Import terminé"
    Else
        MsgBox "Import terminé avec des avertissements :" & vbCrLf & erreurs & vbCrLf & vbCrLf & _
               "Vérifiez les fichiers manquants puis relancez.", _
               vbExclamation, "RESERV+ — Avertissements"
    End If
End Sub


Public Sub LancerInstallationComplete()
' Raccourci : importe + installe en une seule commande
' Usage dans la fenêtre Exécution : LancerInstallationComplete
    Dim chemin As String
    chemin = InputBox( _
        "Chemin complet du dossier vba\" & vbCrLf & _
        "Exemple : C:\RESERV_PLUS_ACCESS\vba", _
        "RESERV+ — Installation complète", _
        "C:\RESERV_PLUS_ACCESS\vba")

    If chemin = "" Then Exit Sub

    ImporterTousLesModules chemin

    Dim rep As Integer
    rep = MsgBox("Lancer maintenant Installer_RESERV_PLUS ?", _
                 vbYesNo + vbQuestion, "RESERV+")
    If rep = vbYes Then
        Installer_RESERV_PLUS
    End If
End Sub
