Attribute VB_Name = "modSession"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Gestion de session utilisateur
' Module : modSession.bas
' Variables globales de session stockant l'utilisateur connecté
' ============================================================

' ---- Variables globales de session ----
Public g_IdUtilisateur As Long
Public g_Login As String
Public g_Role As String
Public g_NomComplet As String
Public g_EstConnecte As Boolean

' ============================================================
' Ouvrir une session pour un utilisateur authentifié
' ============================================================
Public Sub OuvrirSession(idUtilisateur As Long, login As String, role As String, nomComplet As String)
    g_IdUtilisateur = idUtilisateur
    g_Login = login
    g_Role = role
    g_NomComplet = nomComplet
    g_EstConnecte = True
End Sub

' ============================================================
' Fermer la session en cours
' ============================================================
Public Sub FermerSession()
    g_IdUtilisateur = 0
    g_Login = ""
    g_Role = ""
    g_NomComplet = ""
    g_EstConnecte = False
End Sub

' ============================================================
' Vérifier si un utilisateur est connecté
' ============================================================
Public Function EstConnecte() As Boolean
    EstConnecte = g_EstConnecte
End Function

' ============================================================
' Vérifier si l'utilisateur connecté a un rôle donné ou supérieur
' Hiérarchie : Administrateur > Agent > Utilisateur
' ============================================================
Public Function ALeRole(roleRequis As String) As Boolean
    If Not g_EstConnecte Then
        ALeRole = False
        Exit Function
    End If

    Select Case roleRequis
        Case ROLE_UTILISATEUR
            ' Tout utilisateur connecté a au moins ce niveau
            ALeRole = True

        Case ROLE_AGENT
            ALeRole = (g_Role = ROLE_AGENT Or g_Role = ROLE_ADMINISTRATEUR)

        Case ROLE_ADMINISTRATEUR
            ALeRole = (g_Role = ROLE_ADMINISTRATEUR)

        Case Else
            ALeRole = False
    End Select
End Function

' ============================================================
' Obtenir les initiales de l'utilisateur connecté (pour avatar)
' ============================================================
Public Function ObtenirInitiales() As String
    Dim parties() As String
    Dim initiales As String

    If Len(g_NomComplet) = 0 Then
        ObtenirInitiales = "??"
        Exit Function
    End If

    parties = Split(g_NomComplet, " ")

    If UBound(parties) >= 1 Then
        initiales = Left(parties(0), 1) & Left(parties(1), 1)
    Else
        initiales = Left(parties(0), 2)
    End If

    ObtenirInitiales = UCase(initiales)
End Function

' ============================================================
' Vérifier l'accès et rediriger si non connecté
' ============================================================
Public Function VerifierAcces() As Boolean
    If Not g_EstConnecte Then
        MsgBox "Vous devez être connecté pour accéder à cette fonctionnalité.", _
               vbExclamation, APP_NOM
        DoCmd.OpenForm FORM_CONNEXION
        VerifierAcces = False
    Else
        VerifierAcces = True
    End If
End Function

' ============================================================
' Vérifier l'accès avec rôle minimum requis
' ============================================================
Public Function VerifierAccesRole(roleMinimum As String) As Boolean
    If Not VerifierAcces() Then
        VerifierAccesRole = False
        Exit Function
    End If

    If Not ALeRole(roleMinimum) Then
        MsgBox "Votre rôle (" & g_Role & ") ne permet pas d'accéder à cette fonctionnalité." & vbCrLf & _
               "Rôle requis : " & roleMinimum, vbExclamation, APP_NOM
        VerifierAccesRole = False
    Else
        VerifierAccesRole = True
    End If
End Function
