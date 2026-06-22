Attribute VB_Name = "modNavigation"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Navigation entre les formulaires
' Module : modNavigation.bas
' Toutes les fonctions appelées par les boutons de navigation
' ============================================================

' ============================================================
' Ouvrir un formulaire et fermer l'actuel
' ============================================================
Private Sub NaviguerVers(nomFormulaire As String, Optional fermerActuel As String = "")
    If Not EstConnecte() Then
        DoCmd.OpenForm FORM_CONNEXION
        Exit Sub
    End If

    If Not FormulaireExiste(nomFormulaire) Then
        MsgBox "Le formulaire '" & nomFormulaire & "' n'existe pas encore." & vbCrLf & _
               "Relancez l'installation pour créer tous les formulaires.", _
               vbExclamation, APP_NOM
        Exit Sub
    End If

    DoCmd.OpenForm nomFormulaire

    If Len(fermerActuel) > 0 Then
        FermerFormulaire fermerActuel
    End If
End Sub

' ============================================================
' Navigation - Tableau de bord
' ============================================================
Public Function Navigation_Dashboard() As Integer
    NaviguerVers FORM_DASHBOARD
    Navigation_Dashboard = 0
End Function

' ============================================================
' Navigation - Réservations
' ============================================================
Public Function Navigation_Reservations() As Integer
    If Not VerifierAccesRole(ROLE_UTILISATEUR) Then
        Navigation_Reservations = 0
        Exit Function
    End If
    NaviguerVers FORM_RESERVATIONS
    Navigation_Reservations = 0
End Function

' ============================================================
' Navigation - Nouvelle réservation
' ============================================================
Public Function Navigation_NouvelleReservation() As Integer
    If Not VerifierAccesRole(ROLE_AGENT) Then
        Navigation_NouvelleReservation = 0
        Exit Function
    End If
    NaviguerVers FORM_NOUVELLE_RESERVATION
    Navigation_NouvelleReservation = 0
End Function

' ============================================================
' Navigation - Clients
' ============================================================
Public Function Navigation_Clients() As Integer
    If Not VerifierAccesRole(ROLE_UTILISATEUR) Then
        Navigation_Clients = 0
        Exit Function
    End If
    NaviguerVers FORM_CLIENTS
    Navigation_Clients = 0
End Function

' ============================================================
' Navigation - Prestataires
' ============================================================
Public Function Navigation_Prestataires() As Integer
    If Not VerifierAccesRole(ROLE_UTILISATEUR) Then
        Navigation_Prestataires = 0
        Exit Function
    End If
    NaviguerVers FORM_PRESTATAIRES
    Navigation_Prestataires = 0
End Function

' ============================================================
' Navigation - Ressources
' ============================================================
Public Function Navigation_Ressources() As Integer
    If Not VerifierAccesRole(ROLE_UTILISATEUR) Then
        Navigation_Ressources = 0
        Exit Function
    End If
    NaviguerVers FORM_RESSOURCES
    Navigation_Ressources = 0
End Function

' ============================================================
' Navigation - Paiements
' ============================================================
Public Function Navigation_Paiements() As Integer
    If Not VerifierAccesRole(ROLE_AGENT) Then
        Navigation_Paiements = 0
        Exit Function
    End If
    NaviguerVers FORM_PAIEMENTS
    Navigation_Paiements = 0
End Function

' ============================================================
' Navigation - Rapports et états
' ============================================================
Public Function Navigation_Rapports() As Integer
    If Not VerifierAccesRole(ROLE_AGENT) Then
        Navigation_Rapports = 0
        Exit Function
    End If
    NaviguerVers FORM_RAPPORTS
    Navigation_Rapports = 0
End Function

' ============================================================
' Navigation - Paramètres (admin uniquement)
' ============================================================
Public Function Navigation_Parametres() As Integer
    If Not VerifierAccesRole(ROLE_ADMINISTRATEUR) Then
        Navigation_Parametres = 0
        Exit Function
    End If
    NaviguerVers FORM_PARAMETRES
    Navigation_Parametres = 0
End Function

' ============================================================
' Navigation - Utilisateurs (admin uniquement)
' ============================================================
Public Function Navigation_Utilisateurs() As Integer
    If Not VerifierAccesRole(ROLE_ADMINISTRATEUR) Then
        Navigation_Utilisateurs = 0
        Exit Function
    End If
    NaviguerVers FORM_UTILISATEURS
    Navigation_Utilisateurs = 0
End Function

' ============================================================
' Déconnecter l'utilisateur et revenir à l'écran de connexion
' ============================================================
Public Function Navigation_Deconnecter() As Integer
    If DemanderConfirmation("Voulez-vous vous déconnecter de " & APP_NOM & " ?") Then
        ' Fermer tous les formulaires sauf la connexion
        Dim frm As AccessObject
        For Each frm In CurrentProject.AllForms
            If frm.IsLoaded And frm.Name <> FORM_CONNEXION Then
                DoCmd.Close acForm, frm.Name, acSaveNo
            End If
        Next frm

        Deconnecter
        DoCmd.OpenForm FORM_CONNEXION
    End If
    Navigation_Deconnecter = 0
End Function

' ============================================================
' Ouvrir le formulaire de connexion au démarrage
' ============================================================
Public Sub DemarrerApplication()
    ' Fermer le volet de navigation Access
    DoCmd.NavigateTo "acNavigationCategoryNothingX"
    On Error Resume Next
    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    On Error GoTo 0

    ' Ouvrir l'écran de connexion
    If FormulaireExiste(FORM_CONNEXION) Then
        DoCmd.OpenForm FORM_CONNEXION
    Else
        MsgBox "RESERV+ n'est pas encore installé." & vbCrLf & _
               "Exécutez la procédure Installer_RESERV_PLUS() dans le module modInstall_RESERV_PLUS.", _
               vbExclamation, APP_NOM
    End If
End Sub

' ============================================================
' Quitter l'application proprement
' ============================================================
Public Function Quitter_Application() As Integer
    If DemanderConfirmation("Voulez-vous quitter " & APP_NOM & " ?") Then
        If g_EstConnecte Then
            Journaliser "FERMETURE_APP", "Fermeture par " & g_Login
            Deconnecter
        End If
        DoCmd.Quit acQuitSaveAll
    End If
    Quitter_Application = 0
End Function
