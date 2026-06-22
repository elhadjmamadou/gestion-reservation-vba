Attribute VB_Name = "modFormsConnexion"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Création du formulaire F_CONNEXION
' Module : modFormsConnexion.bas
' ============================================================

' ============================================================
' Créer le formulaire de connexion
' ============================================================
Public Sub Creer_F_CONNEXION()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_CONNEXION

    ' Supprimer si déjà existant
    SupprimerObjetAccess acForm, nomFrm

    ' Créer le formulaire
    Set frm = CreateForm()
    frm.Caption = "Connexion — " & APP_NOM

    ' Dimensions : 14000 x 9000 twips
    frm.Width = 14000
    frm.InsideHeight = 9000
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.BorderStyle = 1
    frm.AutoCenter = True
    frm.Detail.BackColor = COULEUR_SIDEBAR

    ' ---- PANNEAU GAUCHE (fond sombre, titre et bullet points) ----
    AjouterRectangle frm, 0, 0, 6400, 9000, COULEUR_SIDEBAR

    ' Titre RESERV+
    AjouterLabel frm, "lblTitreLogo", APP_NOM, _
                 400, 1200, 5600, 900, 24, COULEUR_BLANC, True

    ' Sous-titre
    AjouterLabel frm, "lblSousTitreLogo", "Plateforme de réservation", _
                 400, 2200, 5600, 480, 12, CouleurAccess(176, 190, 197), False

    AjouterLabel frm, "lblSousTitreLogo2", "multi-prestataires", _
                 400, 2700, 5600, 480, 12, CouleurAccess(176, 190, 197), False

    ' Bullet points
    Dim yCourant As Long
    yCourant = 3600

    AjouterLabel frm, "lblBullet1", Chr(10003) & "  Hôtels • Taxis • Vols", _
                 400, yCourant, 5600, 400, 10, CouleurAccess(149, 165, 166), False
    yCourant = yCourant + 500

    AjouterLabel frm, "lblBullet2", Chr(10003) & "  Gestion centralisée des réservations", _
                 400, yCourant, 5600, 400, 10, CouleurAccess(149, 165, 166), False
    yCourant = yCourant + 500

    AjouterLabel frm, "lblBullet3", Chr(10003) & "  Suivi des paiements et billetterie", _
                 400, yCourant, 5600, 400, 10, CouleurAccess(149, 165, 166), False

    ' ---- PANNEAU DROIT (formulaire de connexion) ----
    AjouterRectangle frm, 6400, 0, 7600, 9000, CouleurAccess(245, 246, 250)

    ' Carte de connexion (fond blanc centré)
    AjouterRectangle frm, 6800, 1400, 6800, 6200, COULEUR_BLANC

    ' Titre "Connexion"
    AjouterLabel frm, "lblTitreConnexion", "Connexion", _
                 7100, 1700, 6200, 700, 20, COULEUR_TEXTE, True

    AjouterLabel frm, "lblSousTitreConn", "Accédez à votre espace de gestion", _
                 7100, 2450, 6200, 380, 9, COULEUR_TEXTE_SECONDAIRE, False

    ' ---- Champ Identifiant ----
    AjouterLabel frm, "lblLogin", "Identifiant", _
                 7100, 3000, 6200, 320, 9, COULEUR_TEXTE, False

    Dim txtLogin As Access.TextBox
    Set txtLogin = AjouterZoneTexte(frm, "txtLogin", 7100, 3350, 6200, 480)

    ' ---- Champ Mot de passe ----
    AjouterLabel frm, "lblMdp", "Mot de passe", _
                 7100, 4000, 6200, 320, 9, COULEUR_TEXTE, False

    Dim txtMdp As Access.TextBox
    Set txtMdp = AjouterZoneTexte(frm, "txtMotDePasse", 7100, 4350, 6200, 480, 0, True)

    ' ---- Combo Rôle ----
    AjouterLabel frm, "lblRole", "Rôle", _
                 7100, 5000, 6200, 320, 9, COULEUR_TEXTE, False

    Dim cboRole As Access.ComboBox
    Set cboRole = AjouterCombo(frm, "cboRole", 7100, 5350, 6200, 480)
    cboRole.RowSource = ROLE_ADMINISTRATEUR & ";" & ROLE_AGENT & ";" & ROLE_UTILISATEUR
    cboRole.Value = ROLE_ADMINISTRATEUR

    ' ---- Bouton SE CONNECTER ----
    Dim btnConn As Access.CommandButton
    Set btnConn = AjouterBouton(frm, "btnConnexion", "SE CONNECTER", _
                                 7100, 6100, 6200, 600, COULEUR_PRIMAIRE, COULEUR_BLANC)
    btnConn.FontSize = 12
    btnConn.OnClick = "=Connexion_Connecter()"

    ' ---- Label message d'erreur (caché par défaut) ----
    Dim lblErr As Access.Label
    Set lblErr = AjouterLabel(frm, "lblErreur", "", _
                               7100, 6850, 6200, 400, 9, COULEUR_DANGER, False)
    lblErr.Visible = False

    ' ---- Bouton Quitter (discret) ----
    Dim btnQuitter As Access.CommandButton
    Set btnQuitter = AjouterBouton(frm, "btnQuitter", "Quitter", _
                                    7100, 7350, 2800, 480, _
                                    CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnQuitter.FontSize = 9
    btnQuitter.FontBold = False
    btnQuitter.OnClick = "=Quitter_Application()"

    ' ---- Version en bas ----
    AjouterLabel frm, "lblVersion", APP_NOM & " v" & APP_VERSION, _
                 6800, 8600, 6800, 280, 8, COULEUR_TEXTE_SECONDAIRE, False

    ' Sauvegarder
    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes

    Debug.Print "Formulaire " & nomFrm & " créé avec succès."
End Sub

' ============================================================
' Procédure appelée par le bouton Connexion (depuis le formulaire)
' ============================================================
Public Function Connexion_Connecter() As Integer
    Dim login As String
    Dim mdp As String

    On Error GoTo GestionErreur

    ' Récupérer les valeurs saisies
    login = Trim(Nz(Forms(FORM_CONNEXION)!txtLogin.Value, ""))
    mdp = Trim(Nz(Forms(FORM_CONNEXION)!txtMotDePasse.Value, ""))

    ' Masquer le message d'erreur précédent
    Forms(FORM_CONNEXION)!lblErreur.Visible = False

    If Len(login) = 0 Or Len(mdp) = 0 Then
        Forms(FORM_CONNEXION)!lblErreur.Caption = "Veuillez renseigner l'identifiant et le mot de passe."
        Forms(FORM_CONNEXION)!lblErreur.Visible = True
        Connexion_Connecter = 0
        Exit Function
    End If

    ' Tenter la connexion
    If Connecter(login, mdp) Then
        ' Connexion réussie : ouvrir le tableau de bord
        DoCmd.Close acForm, FORM_CONNEXION, acSaveNo
        DoCmd.OpenForm FORM_DASHBOARD
    Else
        Forms(FORM_CONNEXION)!lblErreur.Caption = "Identifiant ou mot de passe incorrect."
        Forms(FORM_CONNEXION)!lblErreur.Visible = True
        Forms(FORM_CONNEXION)!txtMotDePasse.Value = ""
        Forms(FORM_CONNEXION)!txtMotDePasse.SetFocus
    End If

    Connexion_Connecter = 0
    Exit Function

GestionErreur:
    Forms(FORM_CONNEXION)!lblErreur.Caption = "Erreur : " & Err.Description
    Forms(FORM_CONNEXION)!lblErreur.Visible = True
    Connexion_Connecter = 0
End Function
