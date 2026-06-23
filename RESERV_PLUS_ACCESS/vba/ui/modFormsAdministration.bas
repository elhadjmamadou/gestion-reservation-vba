Attribute VB_Name = "modFormsAdministration"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Formulaires Administration et Paramètres
' Module : modFormsAdministration.bas
' ============================================================

' ============================================================
' Créer le formulaire F_UTILISATEURS (admin seulement)
' ============================================================
Public Sub Creer_F_UTILISATEURS()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_UTILISATEURS

    SupprimerObjetAccess acForm, nomFrm
    Set frm = CreateForm()

    frm.Width = 14000
    frm.Detail.Height = 8000
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.BorderStyle = 1
    frm.AutoCenter = True
    frm.Detail.BackColor = COULEUR_FOND

    ' Bandeau
    AjouterRectangle frm, 0, 0, 14000, 900, CouleurAccess(44, 62, 80)
    AjouterLabel frm, "lblTitrePage", "Gestion des utilisateurs", 200, 100, 10000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Réservé à l'administrateur", 200, 680, 6000, 280, 8, CouleurAccess(200, 200, 220), False

    ' Bouton Fermer
    Dim btnFermer As Access.CommandButton
    Set btnFermer = AjouterBouton(frm, "btnFermer", "X Fermer", _
                                   13000, 200, 800, 480, COULEUR_DANGER, COULEUR_BLANC)
    btnFermer.FontSize = 9
    btnFermer.OnClick = "=DoCmd.Close()"

    ' Liste utilisateurs
    Dim lstUsers As Access.ListBox
    Set lstUsers = CreateControl(frm.Name, acListBox, 0, , , 200, 1100, 13600, 3600)
    lstUsers.Name = "lstUtilisateurs"
    lstUsers.FontName = POLICE
    lstUsers.FontSize = 9
    lstUsers.BackColor = COULEUR_BLANC
    lstUsers.BorderStyle = 0
    lstUsers.RowSourceType = "Table/Query"
    lstUsers.RowSource = "SELECT id_utilisateur, login, nom_complet, role, " & _
                          "IIf(actif, 'Actif', 'Inactif') AS statut, date_creation " & _
                          "FROM UTILISATEUR ORDER BY login"
    lstUsers.BoundColumn = 1
    lstUsers.ColumnCount = 6
    lstUsers.ColumnWidths = "0;2000;3500;2000;1200;2000"
    lstUsers.OnDblClick = "=Admin_OuvrirUtilisateur()"

    ' Fiche utilisateur
    Dim yFiche As Long
    yFiche = 4900

    AjouterRectangle frm, 200, yFiche, 13600, 2400, COULEUR_BLANC
    AjouterLabel frm, "lblTitreFiche", "Fiche utilisateur", 400, yFiche + 120, 5000, 380, 11, COULEUR_TEXTE, True

    Dim largCol As Long
    largCol = 3000

    AjouterLabel frm, "lblLogin", "Login *", 400, yFiche + 620, largCol, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtLogin", 400, yFiche + 920, largCol, 480

    AjouterLabel frm, "lblNomComp", "Nom complet", 3600, yFiche + 620, largCol, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtNomComplet", 3600, yFiche + 920, largCol, 480

    AjouterLabel frm, "lblRoleUser", "Rôle", 6800, yFiche + 620, 2000, 280, 8, COULEUR_TEXTE, False
    Dim cboRoleUser As Access.ComboBox
    Set cboRoleUser = AjouterCombo(frm, "cboRoleUser", 6800, yFiche + 920, 2400, 480)
    cboRoleUser.RowSource = ROLE_ADMINISTRATEUR & ";" & ROLE_AGENT & ";" & ROLE_UTILISATEUR

    AjouterLabel frm, "lblMdpNew", "Nouveau MDP", 9400, yFiche + 620, largCol, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtMdpNew", 9400, yFiche + 920, largCol, 480

    Dim chkActif As Access.CheckBox
    Set chkActif = CreateControl(frm.Name, acCheckBox, 0, , , 400, yFiche + 1600, 300, 360)
    chkActif.Name = "chkActifUser"
    chkActif.Value = True
    AjouterLabel frm, "lblActifUser", "Compte actif", 800, yFiche + 1650, 1500, 280, 8, COULEUR_TEXTE, False

    Dim txtIdU As Access.TextBox
    Set txtIdU = AjouterZoneTexte(frm, "txtIdUtilisateur", 0, 0, 100, 100)
    txtIdU.Visible = False

    Dim btnNouv As Access.CommandButton
    Set btnNouv = AjouterBouton(frm, "btnNouvelUser", "+ Nouveau", 200, yFiche + 1550, 1600, 480, COULEUR_PRIMAIRE, COULEUR_BLANC)
    btnNouv.FontSize = 9
    btnNouv.OnClick = "=Admin_NouvelUtilisateur()"

    Dim btnSaveU As Access.CommandButton
    Set btnSaveU = AjouterBouton(frm, "btnSaveUser", "Enregistrer", 10000, yFiche + 1550, 1800, 480, COULEUR_SUCCESS, COULEUR_BLANC)
    btnSaveU.OnClick = "=Admin_EnregistrerUtilisateur()"

    Dim btnSuppU As Access.CommandButton
    Set btnSuppU = AjouterBouton(frm, "btnSuppUser", "Supprimer", 12000, yFiche + 1550, 1800, 480, COULEUR_DANGER, COULEUR_BLANC)
    btnSuppU.OnClick = "=Admin_SupprimerUtilisateur()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ============================================================
' Créer le formulaire F_PARAMETRES
' ============================================================
Public Sub Creer_F_PARAMETRES()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_PARAMETRES

    SupprimerObjetAccess acForm, nomFrm
    Set frm = CreateForm()

    frm.Width = 12000
    frm.Detail.Height = 7000
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.BorderStyle = 1
    frm.AutoCenter = True
    frm.Detail.BackColor = COULEUR_FOND

    AjouterRectangle frm, 0, 0, 12000, 900, CouleurAccess(44, 62, 80)
    AjouterLabel frm, "lblTitrePage", "Paramètres de l'application", 200, 100, 9000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Configuration générale - Réservé à l'administrateur", _
                 200, 680, 9000, 280, 8, CouleurAccess(200, 200, 220), False

    ' Bouton fermer
    Dim btnFerm As Access.CommandButton
    Set btnFerm = AjouterBouton(frm, "btnFermer", "X Fermer", 11100, 200, 800, 480, COULEUR_DANGER, COULEUR_BLANC)
    btnFerm.FontSize = 9
    btnFerm.OnClick = "=DoCmd.Close()"

    ' Paramètres
    Dim params() As String
    Dim labelsParams() As String
    Dim descParams() As String
    params = Split("NOM_ETABLISSEMENT|DEVISE|PREFIXE_BILLET|NOM_APPLICATION|FREQUENCE_SAUVEGARDE", "|")
    labelsParams = Split("Nom de l'établissement|Devise|Préfixe billet|Nom application|Fréq. sauvegarde (jours)", "|")

    Dim yParam As Long
    yParam = 1100

    AjouterRectangle frm, 200, yParam, 11600, 4800, COULEUR_BLANC

    Dim p As Integer
    For p = 0 To UBound(params)
        Dim valeurActuelle As String
        valeurActuelle = ObtenirParametre(params(p), "")

        AjouterLabel frm, "lblParam" & p, labelsParams(p), _
                     400, yParam + 200 + p * 800, 4000, 300, 9, COULEUR_TEXTE, False
        Dim txtP As Access.TextBox
        Set txtP = AjouterZoneTexte(frm, "txtParam" & p, 4600, yParam + 160 + p * 800, 4000, 480)
        txtP.Value = valeurActuelle
        txtP.Tag = params(p)
    Next p

    AjouterLabel frm, "lblNote", "Les modifications sont appliquées immédiatement.", _
                 400, yParam + 5050, 11200, 280, 8, COULEUR_TEXTE_SECONDAIRE, False

    Dim btnSaveP As Access.CommandButton
    Set btnSaveP = AjouterBouton(frm, "btnSaveParams", "Enregistrer les paramètres", _
                                  200, yParam + 5500, 3500, 600, COULEUR_SUCCESS, COULEUR_BLANC)
    btnSaveP.FontSize = 12
    btnSaveP.OnClick = "=Parametres_Enregistrer()"

    Dim btnJournal As Access.CommandButton
    Set btnJournal = AjouterBouton(frm, "btnViderJournal", "Vider le journal", _
                                    4000, yParam + 5500, 2400, 600, COULEUR_DANGER, COULEUR_BLANC)
    btnJournal.OnClick = "=Admin_ViderJournal()"

    Dim btnUsers As Access.CommandButton
    Set btnUsers = AjouterBouton(frm, "btnGererUsers", "Gérer utilisateurs", _
                                   6700, yParam + 5500, 2400, 600, COULEUR_PRIMAIRE, COULEUR_BLANC)
    btnUsers.OnClick = "=Navigation_Utilisateurs()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ---- Événements Administration ----
Public Function Admin_OuvrirUtilisateur() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_UTILISATEURS)
    Dim idU As Long
    idU = CLng(Nz(frm!lstUtilisateurs.Value, 0))
    If idU = 0 Then Exit Function
    Dim rs As DAO.Recordset
    Set rs = OuvrirRecordset("SELECT * FROM UTILISATEUR WHERE id_utilisateur = " & idU)
    If Not rs.EOF Then
        frm!txtIdUtilisateur.Value = rs!id_utilisateur
        frm!txtLogin.Value = Nz(rs!login, "")
        frm!txtNomComplet.Value = Nz(rs!nom_complet, "")
        frm!cboRoleUser.Value = Nz(rs!role, ROLE_UTILISATEUR)
        frm!chkActifUser.Value = CBool(Nz(rs!actif, True))
        frm!txtMdpNew.Value = ""
    End If
    rs.Close
    Admin_OuvrirUtilisateur = 0
End Function

Public Function Admin_NouvelUtilisateur() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_UTILISATEURS)
    frm!txtIdUtilisateur.Value = 0
    frm!txtLogin.Value = ""
    frm!txtNomComplet.Value = ""
    frm!cboRoleUser.Value = ROLE_UTILISATEUR
    frm!chkActifUser.Value = True
    frm!txtMdpNew.Value = ""
    Admin_NouvelUtilisateur = 0
End Function

Public Function Admin_EnregistrerUtilisateur() As Integer
    On Error Resume Next
    If Not ALeRole(ROLE_ADMINISTRATEUR) Then
        MsgBox "Accès refusé.", vbCritical, APP_NOM
        Admin_EnregistrerUtilisateur = 0
        Exit Function
    End If

    Dim frm As Form
    Set frm = Forms(FORM_UTILISATEURS)
    Dim idU As Long
    idU = CLng(Nz(frm!txtIdUtilisateur.Value, 0))
    Dim login As String
    login = Trim(Nz(frm!txtLogin.Value, ""))

    If Len(login) = 0 Then
        MsgBox "Le login est obligatoire.", vbExclamation, APP_NOM
        Admin_EnregistrerUtilisateur = 0
        Exit Function
    End If

    If idU = 0 Then
        Dim mdp As String
        mdp = Trim(Nz(frm!txtMdpNew.Value, ""))
        If Len(mdp) < 4 Then
            MsgBox "Le mot de passe doit faire au moins 4 caractères.", vbExclamation, APP_NOM
            Admin_EnregistrerUtilisateur = 0
            Exit Function
        End If
        Dim sqlIns As String
        sqlIns = "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) VALUES ('" & _
                 EchapperChaine(login) & "', '" & EchapperChaine(mdp) & "', '" & _
                 EchapperChaine(Trim(Nz(frm!txtNomComplet.Value, ""))) & "', '" & _
                 EchapperChaine(ChaineSure(frm!cboRoleUser.Value)) & "', " & _
                 IIf(CBool(Nz(frm!chkActifUser.Value, True)), "True", "False") & ", Now())"
        ExecuterSQL sqlIns
        JournaliserCreation "Utilisateur", DernierID("UTILISATEUR", "id_utilisateur")
        AfficherSucces "Utilisateur créé."
    Else
        Dim sqlUpd As String
        sqlUpd = "UPDATE UTILISATEUR SET nom_complet = '" & EchapperChaine(Trim(Nz(frm!txtNomComplet.Value, ""))) & "', " & _
                 "role = '" & EchapperChaine(ChaineSure(frm!cboRoleUser.Value)) & "', " & _
                 "actif = " & IIf(CBool(Nz(frm!chkActifUser.Value, True)), "True", "False") & " " & _
                 "WHERE id_utilisateur = " & idU
        ExecuterSQL sqlUpd
        Dim mdpNew As String
        mdpNew = Trim(Nz(frm!txtMdpNew.Value, ""))
        If Len(mdpNew) >= 4 Then
            ReinitialiserMotDePasse idU, mdpNew
        End If
        JournaliserModification "Utilisateur", idU
        AfficherSucces "Utilisateur modifié."
    End If

    frm!lstUtilisateurs.Requery
    Admin_EnregistrerUtilisateur = 0
End Function

Public Function Admin_SupprimerUtilisateur() As Integer
    On Error Resume Next
    If Not ALeRole(ROLE_ADMINISTRATEUR) Then
        Admin_SupprimerUtilisateur = 0
        Exit Function
    End If
    Dim idU As Long
    idU = CLng(Nz(Forms(FORM_UTILISATEURS)!lstUtilisateurs.Value, 0))
    If idU = 0 Or idU = g_IdUtilisateur Then
        MsgBox "Impossible de supprimer le compte en cours ou aucun compte sélectionné.", vbExclamation, APP_NOM
        Admin_SupprimerUtilisateur = 0
        Exit Function
    End If
    If DemanderConfirmation("Supprimer cet utilisateur ?") Then
        ExecuterSQL "DELETE FROM UTILISATEUR WHERE id_utilisateur = " & idU
        JournaliserSuppression "Utilisateur", idU
        Forms(FORM_UTILISATEURS)!lstUtilisateurs.Requery
        AfficherSucces "Utilisateur supprimé."
    End If
    Admin_SupprimerUtilisateur = 0
End Function

Public Function Parametres_Enregistrer() As Integer
    On Error Resume Next
    If Not ALeRole(ROLE_ADMINISTRATEUR) Then
        Parametres_Enregistrer = 0
        Exit Function
    End If
    Dim frm As Form
    Set frm = Forms(FORM_PARAMETRES)
    Dim params() As String
    params = Split("NOM_ETABLISSEMENT|DEVISE|PREFIXE_BILLET|NOM_APPLICATION|FREQUENCE_SAUVEGARDE", "|")
    Dim p As Integer
    For p = 0 To UBound(params)
        Dim valeur As String
        valeur = Trim(Nz(frm.Controls("txtParam" & p).Value, ""))
        If Len(valeur) > 0 Then
            EnregistrerParametre params(p), valeur
        End If
    Next p
    Journaliser "MODIF_PARAMETRES", "Paramètres modifiés par " & g_Login
    AfficherSucces "Paramètres enregistrés avec succès."
    Parametres_Enregistrer = 0
End Function

Public Function Admin_ViderJournal() As Integer
    ViderJournal
    Admin_ViderJournal = 0
End Function
