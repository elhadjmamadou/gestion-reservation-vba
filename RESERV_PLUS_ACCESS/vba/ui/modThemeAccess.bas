Attribute VB_Name = "modThemeAccess"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Thème et constantes visuelles
' Module : modThemeAccess.bas
' Toutes les couleurs et fonctions de création de contrôles
' ============================================================

' ---- Couleurs (format Access : R + G*256 + B*65536) ----
' #1E2D5A - Bleu marine sombre (sidebar)
Public Const COULEUR_SIDEBAR As Long = 5909790
' #F5F6FA - Gris clair (fond général)
Public Const COULEUR_FOND As Long = 16381170
' #5B6FBC - Bleu primaire
Public Const COULEUR_PRIMAIRE As Long = 12349275
' #27AE60 - Vert succès
Public Const COULEUR_SUCCESS As Long = 6336039
' #E67E22 - Orange avertissement
Public Const COULEUR_WARNING As Long = 2260710
' #E74C3C - Rouge danger
Public Const COULEUR_DANGER As Long = 3951847
' #2C3E50 - Texte sombre
Public Const COULEUR_TEXTE As Long = 5258796
' #FFFFFF - Blanc
Public Const COULEUR_BLANC As Long = 16777215
' #ECF0F1 - Gris très clair (cartes)
Public Const COULEUR_CARTE As Long = 15526369
' #BDC3C7 - Bordure grise
Public Const COULEUR_BORDURE As Long = 12436935
' #7F8C8D - Texte secondaire
Public Const COULEUR_TEXTE_SECONDAIRE As Long = 9245837
' #3498DB - Bleu info
Public Const COULEUR_INFO As Long = 3584763

' ---- Dimensions standards (en twips : 1 cm ≈ 567 twips) ----
Public Const LARGEUR_SIDEBAR As Long = 2400    ' ~4.2 cm
Public Const HAUTEUR_HEADER As Long = 720      ' ~1.3 cm
Public Const MARGE_STD As Long = 120           ' ~2 mm
Public Const HAUTEUR_BTN As Long = 420         ' ~7 mm
Public Const HAUTEUR_CTRL As Long = 360        ' ~6 mm
Public Const HAUTEUR_TITRE As Long = 600       ' ~1 cm
Public Const LARGEUR_BTN_STD As Long = 1440   ' ~2.5 cm

' ---- Police ----
Public Const POLICE As String = "Segoe UI"
Public Const TAILLE_TITRE As Integer = 18
Public Const TAILLE_SOUS_TITRE As Integer = 11
Public Const TAILLE_LABEL As Integer = 10
Public Const TAILLE_BOUTON As Integer = 10

' ============================================================
' Obtenir une couleur RGB Access
' ============================================================
Public Function CouleurAccess(r As Integer, g As Integer, b As Integer) As Long
    CouleurAccess = r + (CLng(g) * 256) + (CLng(b) * 65536)
End Function

' ============================================================
' Couleur de statut de réservation
' ============================================================
Public Function CouleurStatutReservation(statut As String) As Long
    Select Case statut
        Case STATUT_CONFIRMEE:  CouleurStatutReservation = COULEUR_SUCCESS
        Case STATUT_EN_ATTENTE: CouleurStatutReservation = COULEUR_WARNING
        Case STATUT_ANNULEE:    CouleurStatutReservation = COULEUR_DANGER
        Case Else:              CouleurStatutReservation = COULEUR_INFO
    End Select
End Function

' ============================================================
' Couleur de statut de paiement
' ============================================================
Public Function CouleurStatutPaiement(etat As String) As Long
    Select Case etat
        Case PAIEMENT_SOLDE:   CouleurStatutPaiement = COULEUR_SUCCESS
        Case PAIEMENT_PARTIEL: CouleurStatutPaiement = COULEUR_WARNING
        Case PAIEMENT_IMPAYE:  CouleurStatutPaiement = COULEUR_DANGER
        Case Else:             CouleurStatutPaiement = COULEUR_INFO
    End Select
End Function

' ============================================================
' Appliquer le style de base à un formulaire
' ============================================================
Public Sub StylerFormulaire(frm As Form, titre As String)
    With frm
        .Caption = titre & " - " & APP_NOM
        .RecordSelectors = False
        .NavigationButtons = False
        .DividingLines = False
        .ScrollBars = 0
        .BorderStyle = 1
        .AutoCenter = True
        .Detail.BackColor = COULEUR_FOND
    End With
End Sub

' ============================================================
' Créer un rectangle de fond (simuler une carte ou sidebar)
' ============================================================
Public Function AjouterRectangle(frm As Form, gauche As Long, haut As Long, _
                                  largeur As Long, hauteur As Long, _
                                  couleurFond As Long, _
                                  Optional section As Integer = 0) As Access.Rectangle
    Dim ctrl As Access.Rectangle
    Set ctrl = CreateControl(frm.Name, acRectangle, section, , , gauche, haut, largeur, hauteur)
    ctrl.BackColor = couleurFond
    ctrl.BackStyle = 1
    ctrl.BorderStyle = 0
    ctrl.SpecialEffect = 0
    Set AjouterRectangle = ctrl
End Function

' ============================================================
' Créer un label (étiquette)
' ============================================================
Public Function AjouterLabel(frm As Form, nom As String, texte As String, _
                              gauche As Long, haut As Long, largeur As Long, hauteur As Long, _
                              Optional taille As Integer = 10, _
                              Optional couleurTexte As Long = -1, _
                              Optional gras As Boolean = False, _
                              Optional section As Integer = 0) As Access.Label
    Dim lbl As Access.Label

    If couleurTexte = -1 Then couleurTexte = COULEUR_TEXTE

    Set lbl = CreateControl(frm.Name, acLabel, section, , , gauche, haut, largeur, hauteur)
    lbl.Name = nom
    lbl.Caption = texte
    lbl.FontName = POLICE
    lbl.FontSize = taille
    lbl.FontBold = gras
    lbl.ForeColor = couleurTexte
    lbl.BackStyle = 0
    lbl.BorderStyle = 0

    Set AjouterLabel = lbl
End Function

' ============================================================
' Créer une zone de texte
' ============================================================
Public Function AjouterZoneTexte(frm As Form, nom As String, _
                                  gauche As Long, haut As Long, _
                                  largeur As Long, hauteur As Long, _
                                  Optional section As Integer = 0, _
                                  Optional motDePasse As Boolean = False) As Access.TextBox
    Dim txt As Access.TextBox
    Set txt = CreateControl(frm.Name, acTextBox, section, , , gauche, haut, largeur, hauteur)
    txt.Name = nom
    txt.FontName = POLICE
    txt.FontSize = TAILLE_LABEL
    txt.BackColor = COULEUR_BLANC
    txt.ForeColor = COULEUR_TEXTE
    txt.BorderStyle = 1
    txt.BorderColor = COULEUR_BORDURE
    txt.SpecialEffect = 0
    If motDePasse Then txt.InputMask = "Password"

    Set AjouterZoneTexte = txt
End Function

' ============================================================
' Créer une liste déroulante (combo)
' ============================================================
Public Function AjouterCombo(frm As Form, nom As String, _
                              gauche As Long, haut As Long, _
                              largeur As Long, hauteur As Long, _
                              Optional section As Integer = 0) As Access.ComboBox
    Dim cbo As Access.ComboBox
    Set cbo = CreateControl(frm.Name, acComboBox, section, , , gauche, haut, largeur, hauteur)
    cbo.Name = nom
    cbo.FontName = POLICE
    cbo.FontSize = TAILLE_LABEL
    cbo.BackColor = COULEUR_BLANC
    cbo.ForeColor = COULEUR_TEXTE
    cbo.BorderStyle = 1
    cbo.BorderColor = COULEUR_BORDURE
    cbo.SpecialEffect = 0
    cbo.RowSourceType = "Value List"

    Set AjouterCombo = cbo
End Function

' ============================================================
' Créer un bouton de commande stylisé
' ============================================================
Public Function AjouterBouton(frm As Form, nom As String, texte As String, _
                               gauche As Long, haut As Long, _
                               largeur As Long, hauteur As Long, _
                               Optional couleurFond As Long = -1, _
                               Optional couleurTexte As Long = -1, _
                               Optional section As Integer = 0) As Access.CommandButton
    Dim btn As Access.CommandButton

    If couleurFond = -1 Then couleurFond = COULEUR_PRIMAIRE
    If couleurTexte = -1 Then couleurTexte = COULEUR_BLANC

    Set btn = CreateControl(frm.Name, acCommandButton, section, , , gauche, haut, largeur, hauteur)
    btn.Name = nom
    btn.Caption = texte
    btn.FontName = POLICE
    btn.FontSize = TAILLE_BOUTON
    btn.FontBold = True
    btn.BackColor = couleurFond
    btn.ForeColor = couleurTexte
    btn.BackStyle = 1
    btn.BorderStyle = 0
    btn.SpecialEffect = 0

    Set AjouterBouton = btn
End Function

' ============================================================
' Créer un bouton de navigation dans la sidebar
' ============================================================
Public Function AjouterBoutonNavigation(frm As Form, nom As String, texte As String, _
                                         haut As Long, _
                                         Optional estActif As Boolean = False) As Access.CommandButton
    Dim btn As Access.CommandButton
    Dim coulFond As Long
    Dim coulTexte As Long

    If estActif Then
        coulFond = COULEUR_PRIMAIRE
    Else
        coulFond = COULEUR_SIDEBAR
    End If
    coulTexte = COULEUR_BLANC

    Set btn = CreateControl(frm.Name, acCommandButton, 0, , , _
              MARGE_STD, haut, LARGEUR_SIDEBAR - (MARGE_STD * 2), HAUTEUR_BTN)
    btn.Name = nom
    btn.Caption = texte
    btn.FontName = POLICE
    btn.FontSize = 10
    btn.FontBold = estActif
    btn.BackColor = coulFond
    btn.ForeColor = coulTexte
    btn.BackStyle = 1
    btn.BorderStyle = 0
    btn.SpecialEffect = 0

    Set AjouterBoutonNavigation = btn
End Function

' ============================================================
' Créer une carte statistique (rectangle + valeur + label)
' ============================================================
Public Sub AjouterCarteStat(frm As Form, nomBase As String, _
                             valeur As String, libelle As String, _
                             gauche As Long, haut As Long, _
                             largeur As Long, hauteur As Long, _
                             couleurAccent As Long)
    ' Fond de la carte
    AjouterRectangle frm, gauche, haut, largeur, hauteur, COULEUR_BLANC

    ' Barre de couleur en haut à gauche (4 twips)
    AjouterRectangle frm, gauche, haut, 60, hauteur, couleurAccent

    ' Valeur principale
    AjouterLabel frm, "lbl" & nomBase & "_Val", valeur, _
                 gauche + 120, haut + 80, largeur - 200, 600, _
                 20, couleurAccent, True

    ' Libellé descriptif
    AjouterLabel frm, "lbl" & nomBase & "_Lib", libelle, _
                 gauche + 120, haut + 720, largeur - 200, 360, _
                 9, COULEUR_TEXTE_SECONDAIRE, False
End Sub

' ============================================================
' Créer un titre de module (bandeau coloré + titre)
' ============================================================
Public Sub AjouterTitreModule(frm As Form, titre As String, sousTitre As String, _
                               gauche As Long, haut As Long, largeur As Long, _
                               couleur As Long)
    ' Bandeau de fond
    AjouterRectangle frm, gauche, haut, largeur, 900, couleur

    ' Titre
    AjouterLabel frm, "lblTitreModule", titre, _
                 gauche + 200, haut + 80, largeur - 400, 480, _
                 16, COULEUR_BLANC, True

    ' Sous-titre
    If Len(sousTitre) > 0 Then
        AjouterLabel frm, "lblSousTitreModule", sousTitre, _
                     gauche + 200, haut + 580, largeur - 400, 280, _
                     9, COULEUR_BLANC, False
    End If
End Sub

' ============================================================
' Créer la sidebar de navigation complète
' ============================================================
Public Sub AjouterSidebar(frm As Form, moduleActif As String)
    Dim hauteurFrm As Long
    hauteurFrm = 9000 ' hauteur standard formulaire

    ' Fond sidebar
    AjouterRectangle frm, 0, 0, LARGEUR_SIDEBAR, hauteurFrm, COULEUR_SIDEBAR

    ' Logo RESERV+
    AjouterLabel frm, "lblLogoSidebar", APP_NOM, _
                 MARGE_STD, 200, LARGEUR_SIDEBAR - (MARGE_STD * 2), 600, _
                 16, COULEUR_BLANC, True

    ' Sous-titre logo
    AjouterLabel frm, "lblSousLogoSidebar", "Plateforme de réservation", _
                 MARGE_STD, 820, LARGEUR_SIDEBAR - (MARGE_STD * 2), 280, _
                 8, CouleurAccess(176, 190, 197), False

    ' Séparateur
    AjouterRectangle frm, MARGE_STD, 1160, LARGEUR_SIDEBAR - (MARGE_STD * 2), 30, _
                     CouleurAccess(52, 73, 94)

    ' Boutons de navigation
    Dim haut As Long
    haut = 1260

    Dim btnDB As Access.CommandButton
    Set btnDB = AjouterBoutonNavigation(frm, "btnNavDashboard", Chr(9632) & "  Tableau de bord", _
                                         haut, (moduleActif = "Dashboard"))
    btnDB.OnClick = "=Navigation_Dashboard()"
    haut = haut + HAUTEUR_BTN + 80

    Dim btnRes As Access.CommandButton
    Set btnRes = AjouterBoutonNavigation(frm, "btnNavReservations", Chr(9675) & "  Réservations", _
                                          haut, (moduleActif = "Reservations"))
    btnRes.OnClick = "=Navigation_Reservations()"
    haut = haut + HAUTEUR_BTN + 80

    Dim btnRess As Access.CommandButton
    Set btnRess = AjouterBoutonNavigation(frm, "btnNavRessources", Chr(9632) & "  Ressources", _
                                           haut, (moduleActif = "Ressources"))
    btnRess.OnClick = "=Navigation_Ressources()"
    haut = haut + HAUTEUR_BTN + 80

    Dim btnCli As Access.CommandButton
    Set btnCli = AjouterBoutonNavigation(frm, "btnNavClients", Chr(9632) & "  Clients", _
                                          haut, (moduleActif = "Clients"))
    btnCli.OnClick = "=Navigation_Clients()"
    haut = haut + HAUTEUR_BTN + 80

    Dim btnPre As Access.CommandButton
    Set btnPre = AjouterBoutonNavigation(frm, "btnNavPrestataires", Chr(9675) & "  Prestataires", _
                                          haut, (moduleActif = "Prestataires"))
    btnPre.OnClick = "=Navigation_Prestataires()"
    haut = haut + HAUTEUR_BTN + 80

    Dim btnPai As Access.CommandButton
    Set btnPai = AjouterBoutonNavigation(frm, "btnNavPaiements", "$  Paiements", _
                                          haut, (moduleActif = "Paiements"))
    btnPai.OnClick = "=Navigation_Paiements()"
    haut = haut + HAUTEUR_BTN + 80

    Dim btnEtat As Access.CommandButton
    Set btnEtat = AjouterBoutonNavigation(frm, "btnNavEtats", Chr(9632) & "  États & rapports", _
                                           haut, (moduleActif = "Etats"))
    btnEtat.OnClick = "=Navigation_Rapports()"
    haut = haut + HAUTEUR_BTN + 80

    Dim btnParam As Access.CommandButton
    Set btnParam = AjouterBoutonNavigation(frm, "btnNavParametres", Chr(9675) & "  Paramètres", _
                                            haut, (moduleActif = "Parametres"))
    btnParam.OnClick = "=Navigation_Parametres()"

    ' Informations utilisateur en bas
    AjouterRectangle frm, 0, hauteurFrm - 700, LARGEUR_SIDEBAR, 700, _
                     CouleurAccess(16, 22, 46)

    AjouterLabel frm, "lblUserNom", g_NomComplet, _
                 MARGE_STD + 300, hauteurFrm - 620, LARGEUR_SIDEBAR - 500, 320, _
                 9, COULEUR_BLANC, True

    AjouterLabel frm, "lblUserRole", g_Role, _
                 MARGE_STD + 300, hauteurFrm - 300, LARGEUR_SIDEBAR - 500, 240, _
                 8, CouleurAccess(127, 140, 141), False

    ' Bouton déconnexion
    Dim btnDeco As Access.CommandButton
    Set btnDeco = CreateControl(frm.Name, acCommandButton, 0, , , _
                  LARGEUR_SIDEBAR - 460, hauteurFrm - 520, 400, 360)
    btnDeco.Name = "btnDeconnexion"
    btnDeco.Caption = "X"
    btnDeco.FontName = POLICE
    btnDeco.FontSize = 9
    btnDeco.FontBold = True
    btnDeco.BackColor = COULEUR_DANGER
    btnDeco.ForeColor = COULEUR_BLANC
    btnDeco.BackStyle = 1
    btnDeco.BorderStyle = 0
    btnDeco.SpecialEffect = 0
    btnDeco.OnClick = "=Navigation_Deconnecter()"
End Sub

' ============================================================
' Créer un champ avec son label associé
' ============================================================
Public Function AjouterChampAvecLabel(frm As Form, nomCtrl As String, _
                                       texteLabel As String, gauche As Long, haut As Long, _
                                       largeurLabel As Long, largeurCtrl As Long, _
                                       Optional section As Integer = 0) As Access.TextBox
    ' Label
    AjouterLabel frm, "lbl" & nomCtrl, texteLabel, _
                 gauche, haut, largeurLabel, HAUTEUR_CTRL, _
                 TAILLE_LABEL, COULEUR_TEXTE, False, section

    ' TextBox
    Set AjouterChampAvecLabel = AjouterZoneTexte(frm, nomCtrl, _
                                 gauche, haut + HAUTEUR_CTRL + 60, _
                                 largeurCtrl, HAUTEUR_CTRL + 60, section)
End Function
