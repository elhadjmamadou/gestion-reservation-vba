Attribute VB_Name = "modSecurite"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Sécurité, authentification et contrôle des rôles
' Module : modSecurite.bas
' ============================================================

' ============================================================
' Connecter un utilisateur : vérifie login + mdp + actif
' Retourne True si connexion réussie
' ============================================================
Public Function Connecter(login As String, motDePasse As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo GestionErreur

    If Len(Trim(login)) = 0 Or Len(Trim(motDePasse)) = 0 Then
        MsgBox "Veuillez saisir votre identifiant et votre mot de passe.", _
               vbExclamation, APP_NOM
        Connecter = False
        Exit Function
    End If

    Set db = CurrentDb()
    sql = "SELECT id_utilisateur, login, nom_complet, role, actif, mot_de_passe " & _
          "FROM UTILISATEUR WHERE login = '" & EchapperChaine(login) & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        MsgBox "Identifiant incorrect. Aucun compte trouvé pour : " & login, _
               vbExclamation, APP_NOM
        Connecter = False
        GoTo Nettoyage
    End If

    ' Vérifier si le compte est actif
    If Not rs!actif Then
        MsgBox "Ce compte est désactivé. Contactez l'administrateur.", _
               vbCritical, APP_NOM
        Connecter = False
        GoTo Nettoyage
    End If

    ' Vérifier le mot de passe (comparaison directe - améliorer avec hash en prod)
    If rs!mot_de_passe <> motDePasse Then
        MsgBox "Mot de passe incorrect. Veuillez réessayer.", _
               vbExclamation, APP_NOM
        Connecter = False
        GoTo Nettoyage
    End If

    ' Authentification réussie : ouvrir la session
    OuvrirSession CLng(rs!id_utilisateur), CStr(rs!login), CStr(rs!role), CStr(Nz(rs!nom_complet, login))

    ' Journaliser la connexion
    Journaliser "CONNEXION", "Connexion réussie - Rôle : " & g_Role

    Connecter = True

Nettoyage:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

GestionErreur:
    MsgBox "Erreur lors de la connexion : " & Err.Description, vbCritical, APP_NOM
    Connecter = False
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Déconnecter l'utilisateur en cours
' ============================================================
Public Sub Deconnecter()
    If g_EstConnecte Then
        Journaliser "DECONNEXION", "Déconnexion de " & g_Login
    End If
    FermerSession
End Sub

' ============================================================
' Échapper les apostrophes dans une chaîne SQL
' ============================================================
Public Function EchapperChaine(chaine As String) As String
    EchapperChaine = Replace(chaine, "'", "''")
End Function

' ============================================================
' Vérifier si l'utilisateur peut effectuer une action selon son rôle
' ============================================================
Public Function PeutEffectuer(action As String) As Boolean
    If Not g_EstConnecte Then
        PeutEffectuer = False
        Exit Function
    End If

    Select Case UCase(action)
        ' Actions réservées à l'administrateur
        Case "SUPPRIMER_UTILISATEUR", "MODIFIER_PARAMETRES", "VOIR_JOURNAL", _
             "CREER_UTILISATEUR", "MODIFIER_UTILISATEUR", "DESACTIVER_UTILISATEUR"
            PeutEffectuer = (g_Role = ROLE_ADMINISTRATEUR)

        ' Actions réservées à l'agent et l'administrateur
        Case "CREER_CLIENT", "MODIFIER_CLIENT", "SUPPRIMER_CLIENT", _
             "CREER_PRESTATAIRE", "MODIFIER_PRESTATAIRE", "SUPPRIMER_PRESTATAIRE", _
             "CREER_RESSOURCE", "MODIFIER_RESSOURCE", "SUPPRIMER_RESSOURCE", _
             "CREER_RESERVATION", "MODIFIER_RESERVATION", "ANNULER_RESERVATION", _
             "CONFIRMER_RESERVATION", "ENREGISTRER_PAIEMENT", "GENERER_BILLET", _
             "VOIR_RAPPORTS"
            PeutEffectuer = (g_Role = ROLE_AGENT Or g_Role = ROLE_ADMINISTRATEUR)

        ' Actions accessibles à tous les utilisateurs connectés
        Case "VOIR_CLIENTS", "VOIR_PRESTATAIRES", "VOIR_RESSOURCES", _
             "VOIR_RESERVATIONS", "VOIR_PAIEMENTS", "VOIR_BILLETS", _
             "VOIR_DASHBOARD"
            PeutEffectuer = True

        Case Else
            ' Par défaut, seul l'admin peut faire des actions inconnues
            PeutEffectuer = (g_Role = ROLE_ADMINISTRATEUR)
    End Select
End Function

' ============================================================
' Appliquer les droits d'accès à un formulaire selon le rôle
' Masque ou désactive les contrôles non autorisés
' ============================================================
Public Sub AppliquerDroitsFormulaire(frm As Form)
    Dim ctrl As Control

    On Error Resume Next

    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
            Case "CommandButton"
                AppliquerDroitsBouton ctrl
        End Select
    Next ctrl

    On Error GoTo 0
End Sub

' ============================================================
' Appliquer les droits à un bouton selon son nom et le rôle
' Convention de nommage : btnSupprimer, btnAdmin_xxx
' ============================================================
Private Sub AppliquerDroitsBouton(btn As CommandButton)
    Dim nomBouton As String
    nomBouton = LCase(btn.Name)

    ' Boutons réservés administrateur
    If InStr(nomBouton, "supprimer") > 0 Or _
       InStr(nomBouton, "admin") > 0 Or _
       InStr(nomBouton, "utilisateurs") > 0 Or _
       InStr(nomBouton, "parametres") > 0 Then
        btn.Enabled = (g_Role = ROLE_ADMINISTRATEUR)
        btn.Visible = (g_Role = ROLE_ADMINISTRATEUR)
    End If

    ' Boutons réservés agent+
    If InStr(nomBouton, "ajouter") > 0 Or _
       InStr(nomBouton, "modifier") > 0 Or _
       InStr(nomBouton, "enregistrer") > 0 Or _
       InStr(nomBouton, "confirmer") > 0 Or _
       InStr(nomBouton, "annuler_res") > 0 Then
        btn.Enabled = (g_Role = ROLE_AGENT Or g_Role = ROLE_ADMINISTRATEUR)
    End If
End Sub

' ============================================================
' Changer le mot de passe d'un utilisateur
' ============================================================
Public Function ChangerMotDePasse(idUtilisateur As Long, ancienMdp As String, _
                                   nouveauMdp As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo GestionErreur

    If Len(Trim(nouveauMdp)) < 4 Then
        MsgBox "Le nouveau mot de passe doit contenir au moins 4 caractères.", _
               vbExclamation, APP_NOM
        ChangerMotDePasse = False
        Exit Function
    End If

    Set db = CurrentDb()
    sql = "SELECT * FROM UTILISATEUR WHERE id_utilisateur = " & idUtilisateur
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        MsgBox "Utilisateur introuvable.", vbExclamation, APP_NOM
        ChangerMotDePasse = False
        GoTo Nettoyage
    End If

    If rs!mot_de_passe <> ancienMdp Then
        MsgBox "L'ancien mot de passe est incorrect.", vbExclamation, APP_NOM
        ChangerMotDePasse = False
        GoTo Nettoyage
    End If

    rs.Edit
    rs!mot_de_passe = nouveauMdp
    rs.Update

    Journaliser "CHANGEMENT_MDP", "Changement de mot de passe pour ID : " & idUtilisateur
    ChangerMotDePasse = True

Nettoyage:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

GestionErreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, APP_NOM
    ChangerMotDePasse = False
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Réinitialiser le mot de passe (admin seulement)
' ============================================================
Public Sub ReinitialiserMotDePasse(idUtilisateur As Long, nouveauMdp As String)
    If Not ALeRole(ROLE_ADMINISTRATEUR) Then
        MsgBox "Action réservée à l'administrateur.", vbCritical, APP_NOM
        Exit Sub
    End If

    Dim sql As String
    sql = "UPDATE UTILISATEUR SET mot_de_passe = '" & EchapperChaine(nouveauMdp) & _
          "' WHERE id_utilisateur = " & idUtilisateur

    ExecuterSQL sql
    Journaliser "REINIT_MDP", "Réinitialisation mot de passe pour ID : " & idUtilisateur
    MsgBox "Mot de passe réinitialisé avec succès.", vbInformation, APP_NOM
End Sub
