Attribute VB_Name = "modFormsReservations"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Formulaires de gestion des réservations
' Module : modFormsReservations.bas
' ============================================================

' ============================================================
' Créer le formulaire liste des réservations F_RESERVATIONS
' ============================================================
Public Sub Creer_F_RESERVATIONS()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_RESERVATIONS

    SupprimerObjetAccess acForm, nomFrm
    Set frm = CreateForm()

    frm.Width = 16000
    frm.Detail.Height = 9200
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.BorderStyle = 1
    frm.AutoCenter = True
    frm.Detail.BackColor = COULEUR_FOND

    AjouterSidebar frm, "Reservations"

    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200
    Dim largeurContenu As Long
    largeurContenu = 16000 - xDebut - 200

    ' Bandeau
    AjouterRectangle frm, xDebut, 0, largeurContenu, 1000, COULEUR_PRIMAIRE
    AjouterLabel frm, "lblTitrePage", "Réservations", xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Gérer et suivre toutes les réservations", _
                 xDebut + 200, 680, 8000, 300, 9, CouleurAccess(200, 210, 230), False

    ' Barre recherche + filtres
    Dim yRecherche As Long
    yRecherche = 1200

    AjouterZoneTexte frm, "txtRecherche", xDebut + 200, yRecherche, 4000, 480
    Forms(FORM_RESERVATIONS)!txtRecherche.SpecialEffect = 0  ' sera ignoré pendant création

    Dim btnRechRes As Access.CommandButton
    Set btnRechRes = AjouterBouton(frm, "btnRechercher", "Rechercher", _
                                    xDebut + 4300, yRecherche, 1600, 480, _
                                    CouleurAccess(100, 100, 130), COULEUR_BLANC)
    btnRechRes.FontSize = 9
    btnRechRes.FontBold = False
    btnRechRes.OnClick = "=Reservations_Filtrer()"

    ' Filtres statut (boutons radio-like)
    Dim filtres() As String
    filtres = Split("Toutes|Confirmées|En attente|Annulées", "|")
    Dim xFiltre As Long
    xFiltre = xDebut + 6200
    Dim f As Integer
    For f = 0 To 3
        Dim btnFiltre As Access.CommandButton
        Set btnFiltre = AjouterBouton(frm, "btnFiltre" & f, filtres(f), _
                                       xFiltre, yRecherche, 1800, 480, _
                                       IIf(f = 0, COULEUR_PRIMAIRE, CouleurAccess(200, 200, 210)), _
                                       IIf(f = 0, COULEUR_BLANC, COULEUR_TEXTE))
        btnFiltre.FontSize = 9
        btnFiltre.FontBold = False
        btnFiltre.Tag = filtres(f)
        btnFiltre.OnClick = "=Reservations_Filtrer()"
        xFiltre = xFiltre + 1900
    Next f

    ' Bouton Nouvelle réservation
    Dim btnNouv As Access.CommandButton
    Set btnNouv = AjouterBouton(frm, "btnNouvelleReservation", "+ Nouvelle réservation", _
                                  xDebut + largeurContenu - 2800, yRecherche, 2600, 480, _
                                  COULEUR_SUCCESS, COULEUR_BLANC)
    btnNouv.OnClick = "=Navigation_NouvelleReservation()"

    ' En-têtes liste
    Dim yHeader As Long
    yHeader = yRecherche + 700
    AjouterRectangle frm, xDebut + 200, yHeader, largeurContenu - 400, 480, _
                     CouleurAccess(236, 240, 241)
    AjouterLabel frm, "lblColN", "N°", xDebut + 300, yHeader + 80, 700, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColCli", "Client", xDebut + 1100, yHeader + 80, 2800, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColPer", "Période", xDebut + 4000, yHeader + 80, 2400, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColMont", "Montant", xDebut + 6500, yHeader + 80, 2000, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColStat", "Statut", xDebut + 8600, yHeader + 80, 2000, 320, 8, COULEUR_TEXTE, True

    ' Liste réservations
    Dim lstRes As Access.ListBox
    Set lstRes = CreateControl(frm.Name, acListBox, 0, , , _
                 xDebut + 200, yHeader + 480, largeurContenu - 400, 4800)
    lstRes.Name = "lstReservations"
    lstRes.FontName = POLICE
    lstRes.FontSize = 9
    lstRes.BackColor = COULEUR_BLANC
    lstRes.ForeColor = COULEUR_TEXTE
    lstRes.BorderStyle = 0
    lstRes.RowSourceType = "Table/Query"
    lstRes.RowSource = RQ_RESERVATIONS_COMPLETES
    lstRes.BoundColumn = 1
    lstRes.ColumnCount = 8
    lstRes.ColumnWidths = "700;2800;2400;2000;2000;2000;0;0"
    lstRes.OnDblClick = "=Reservations_OuvrirDetail()"

    ' Boutons action
    Dim yBtns As Long
    yBtns = yHeader + 5450

    Dim btnOuvrir As Access.CommandButton
    Set btnOuvrir = AjouterBouton(frm, "btnOuvrirDetail", "Détail", _
                                   xDebut + 200, yBtns, 1600, 480, COULEUR_PRIMAIRE, COULEUR_BLANC)
    btnOuvrir.OnClick = "=Reservations_OuvrirDetail()"

    Dim btnConf As Access.CommandButton
    Set btnConf = AjouterBouton(frm, "btnConfirmer", "Confirmer", _
                                 xDebut + 2000, yBtns, 1600, 480, COULEUR_SUCCESS, COULEUR_BLANC)
    btnConf.OnClick = "=Reservations_Confirmer()"

    Dim btnAnn As Access.CommandButton
    Set btnAnn = AjouterBouton(frm, "btnAnnulerRes", "Annuler", _
                                xDebut + 3800, yBtns, 1600, 480, COULEUR_WARNING, COULEUR_BLANC)
    btnAnn.OnClick = "=Reservations_Annuler()"

    Dim btnPai As Access.CommandButton
    Set btnPai = AjouterBouton(frm, "btnPaiement", "Paiement", _
                                xDebut + 5600, yBtns, 1600, 480, COULEUR_SUCCESS, COULEUR_BLANC)
    btnPai.OnClick = "=Navigation_Paiements()"

    Dim btnBillet As Access.CommandButton
    Set btnBillet = AjouterBouton(frm, "btnBillet", "Billet", _
                                   xDebut + 7400, yBtns, 1600, 480, _
                                   CouleurAccess(142, 68, 173), COULEUR_BLANC)
    btnBillet.OnClick = "=Reservations_GenererBillet()"

    Dim btnActu As Access.CommandButton
    Set btnActu = AjouterBouton(frm, "btnActualiser", "Actualiser", _
                                  xDebut + 9200, yBtns, 1600, 480, _
                                  CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnActu.FontBold = False
    btnActu.OnClick = "=Reservations_Actualiser()"

    ' Compteur
    AjouterLabel frm, "lblCompteur", "0 réservation(s)", _
                 xDebut + largeurContenu - 2600, yBtns + 80, 2400, 320, 8, COULEUR_TEXTE_SECONDAIRE, False

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ============================================================
' Créer le formulaire Nouvelle Réservation F_NOUVELLE_RESERVATION
' ============================================================
Public Sub Creer_F_NOUVELLE_RESERVATION()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_NOUVELLE_RESERVATION

    SupprimerObjetAccess acForm, nomFrm
    Set frm = CreateForm()

    frm.Width = 16000
    frm.Detail.Height = 9400
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.BorderStyle = 1
    frm.AutoCenter = True
    frm.Detail.BackColor = COULEUR_FOND

    AjouterSidebar frm, "Reservations"

    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200
    Dim largeurContenu As Long
    largeurContenu = 16000 - xDebut - 200

    ' Bandeau
    AjouterRectangle frm, xDebut, 0, largeurContenu, 1000, COULEUR_PRIMAIRE
    AjouterLabel frm, "lblTitrePage", "Nouvelle réservation", xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Créer une réservation avec contrôle de disponibilité", _
                 xDebut + 200, 680, 10000, 300, 9, CouleurAccess(200, 210, 230), False

    ' Zone formulaire
    Dim yFrm As Long
    yFrm = 1200

    AjouterRectangle frm, xDebut + 200, yFrm, largeurContenu - 400, 7800, COULEUR_BLANC

    ' Section informations générales
    AjouterLabel frm, "lblSectionInfo", "Informations générales", _
                 xDebut + 400, yFrm + 200, 6000, 380, 11, COULEUR_TEXTE, True

    ' Grille 2 colonnes
    Dim largCol As Long
    largCol = (largeurContenu - 1200) \ 2

    ' Client
    AjouterLabel frm, "lblClient", "Client *", xDebut + 400, yFrm + 700, largCol, 280, 8, COULEUR_TEXTE, False
    Dim cboClient As Access.ComboBox
    Set cboClient = AjouterCombo(frm, "cboClient", xDebut + 400, yFrm + 1000, largCol, 480)
    cboClient.RowSourceType = "Table/Query"
    cboClient.RowSource = "SELECT id_client, nom & ' ' & prenom & ' (' & telephone & ')' FROM CLIENT ORDER BY nom"
    cboClient.BoundColumn = 1
    cboClient.ColumnCount = 2
    cboClient.ColumnWidths = "0;5000"
    cboClient.OnChange = "=NouvelleRes_RecalculerMontant()"

    ' Catégorie
    AjouterLabel frm, "lblCateg", "Catégorie *", xDebut + 400 + largCol + 400, yFrm + 700, largCol, 280, 8, COULEUR_TEXTE, False
    Dim cboCateg As Access.ComboBox
    Set cboCateg = AjouterCombo(frm, "cboCategorie", xDebut + 400 + largCol + 400, yFrm + 1000, largCol, 480)
    cboCateg.RowSourceType = "Table/Query"
    cboCateg.RowSource = "SELECT id_categorie, libelle_categorie FROM CATEGORIE ORDER BY libelle_categorie"
    cboCateg.BoundColumn = 1
    cboCateg.ColumnCount = 2
    cboCateg.ColumnWidths = "0;3000"
    cboCateg.OnChange = "=NouvelleRes_ChargerRessources()"

    ' Ressource
    AjouterLabel frm, "lblRess", "Ressource *", xDebut + 400, yFrm + 1600, largCol, 280, 8, COULEUR_TEXTE, False
    Dim cboRess As Access.ComboBox
    Set cboRess = AjouterCombo(frm, "cboRessource", xDebut + 400, yFrm + 1900, largCol, 480)
    cboRess.RowSourceType = "Table/Query"
    cboRess.RowSource = ""
    cboRess.BoundColumn = 1
    cboRess.ColumnCount = 3
    cboRess.ColumnWidths = "0;4000;1500"
    cboRess.OnChange = "=NouvelleRes_ChoisirRessource()"

    ' Nb personnes
    AjouterLabel frm, "lblNbPers", "Nombre de personnes", xDebut + 400 + largCol + 400, yFrm + 1600, largCol, 280, 8, COULEUR_TEXTE, False
    Dim txtNbPers As Access.TextBox
    Set txtNbPers = AjouterZoneTexte(frm, "txtNbPersonnes", xDebut + 400 + largCol + 400, yFrm + 1900, largCol \ 2, 480)
    txtNbPers.Value = 1
    txtNbPers.OnChange = "=NouvelleRes_RecalculerMontant()"

    ' Date début
    AjouterLabel frm, "lblDateDeb", "Date début *", xDebut + 400, yFrm + 2500, largCol, 280, 8, COULEUR_TEXTE, False
    Dim txtDateDeb As Access.TextBox
    Set txtDateDeb = AjouterZoneTexte(frm, "txtDateDebut", xDebut + 400, yFrm + 2800, largCol, 480)
    txtDateDeb.Format = "dd/mm/yyyy"
    txtDateDeb.OnChange = "=NouvelleRes_RecalculerMontant()"

    ' Date fin
    AjouterLabel frm, "lblDateFin", "Date fin *", xDebut + 400 + largCol + 400, yFrm + 2500, largCol, 280, 8, COULEUR_TEXTE, False
    Dim txtDateFin As Access.TextBox
    Set txtDateFin = AjouterZoneTexte(frm, "txtDateFin", xDebut + 400 + largCol + 400, yFrm + 2800, largCol, 480)
    txtDateFin.Format = "dd/mm/yyyy"
    txtDateFin.OnChange = "=NouvelleRes_RecalculerMontant()"

    ' Prix unitaire
    AjouterLabel frm, "lblPrix", "Prix unitaire (GNF)", xDebut + 400, yFrm + 3400, largCol, 280, 8, COULEUR_TEXTE, False
    Dim txtPrix As Access.TextBox
    Set txtPrix = AjouterZoneTexte(frm, "txtPrixUnitaire", xDebut + 400, yFrm + 3700, largCol, 480)
    txtPrix.Format = "#,##0"

    ' Bouton vérifier disponibilité
    Dim btnDispo As Access.CommandButton
    Set btnDispo = AjouterBouton(frm, "btnVerifierDispo", "Vérifier disponibilité", _
                                  xDebut + 400 + largCol + 400, yFrm + 3700, largCol, 480, _
                                  CouleurAccess(52, 73, 94), COULEUR_BLANC)
    btnDispo.OnClick = "=NouvelleRes_VerifierDisponibilite()"

    ' Zone résultat disponibilité
    Dim lblDispo As Access.Label
    Set lblDispo = AjouterLabel(frm, "lblResultatDispo", "", _
                                 xDebut + 400, yFrm + 4350, largeurContenu - 800, 600, _
                                 11, COULEUR_SUCCESS, True)

    ' Zone montant calculé
    AjouterRectangle frm, xDebut + 400 + largCol + 400, yFrm + 4350, largCol, 760, COULEUR_PRIMAIRE
    AjouterLabel frm, "lblTitreMontant", "Montant total calculé", _
                 xDebut + 600 + largCol + 400, yFrm + 4430, largCol - 400, 300, 8, COULEUR_BLANC, False
    AjouterLabel frm, "lblMontantTotal", "0 GNF", _
                 xDebut + 600 + largCol + 400, yFrm + 4730, largCol - 400, 400, 14, COULEUR_BLANC, True

    ' Champ caché ID réservation
    Dim txtIdRes As Access.TextBox
    Set txtIdRes = AjouterZoneTexte(frm, "txtIdReservation", 0, 0, 100, 100)
    txtIdRes.Visible = False

    ' Observation
    AjouterLabel frm, "lblObs", "Observation", xDebut + 400, yFrm + 5250, largeurContenu - 800, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtObservation", xDebut + 400, yFrm + 5550, largeurContenu - 800, 600

    ' Boutons action
    Dim yBtns As Long
    yBtns = yFrm + 6400

    Dim btnEnreg As Access.CommandButton
    Set btnEnreg = AjouterBouton(frm, "btnEnregistrer", Chr(10003) & " Enregistrer", _
                                  xDebut + 400, yBtns, 2400, 600, COULEUR_SUCCESS, COULEUR_BLANC)
    btnEnreg.FontSize = 12
    btnEnreg.OnClick = "=NouvelleRes_Enregistrer()"

    Dim btnReinit As Access.CommandButton
    Set btnReinit = AjouterBouton(frm, "btnReinitialiser", "Réinitialiser", _
                                   xDebut + 3200, yBtns, 2000, 600, _
                                   CouleurAccess(149, 165, 166), COULEUR_BLANC)
    btnReinit.OnClick = "=NouvelleRes_Reinitialiser()"

    Dim btnAnnuler As Access.CommandButton
    Set btnAnnuler = AjouterBouton(frm, "btnAnnuler", "Annuler", _
                                    xDebut + largeurContenu - 2400, yBtns, 2000, 600, _
                                    COULEUR_DANGER, COULEUR_BLANC)
    btnAnnuler.OnClick = "=Navigation_Reservations()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ============================================================
' Événements du formulaire Nouvelle Réservation
' ============================================================
Public Function NouvelleRes_ChargerRessources() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_NOUVELLE_RESERVATION)

    Dim idCateg As Long
    idCateg = CLng(Nz(frm!cboCategorie.Value, 0))

    If idCateg > 0 Then
        frm!cboRessource.RowSource = "SELECT id_ressource, libelle, prix_unitaire FROM RESSOURCE " & _
                                      "WHERE id_categorie = " & idCateg & " AND disponible = True ORDER BY libelle"
        frm!cboRessource.Requery
    End If
    NouvelleRes_ChargerRessources = 0
End Function

Public Function NouvelleRes_ChoisirRessource() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_NOUVELLE_RESERVATION)

    Dim idRess As Long
    idRess = CLng(Nz(frm!cboRessource.Value, 0))

    If idRess > 0 Then
        Dim prix As Currency
        prix = CCur(Nz(ObtenirValeur("SELECT prix_unitaire FROM RESSOURCE WHERE id_ressource = " & idRess), 0))
        frm!txtPrixUnitaire.Value = prix
        NouvelleRes_RecalculerMontant
    End If
    NouvelleRes_ChoisirRessource = 0
End Function

Public Function NouvelleRes_RecalculerMontant() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_NOUVELLE_RESERVATION)

    Dim dateDebut As Variant
    Dim dateFin As Variant
    Dim prixUnitaire As Currency
    Dim nbPersonnes As Integer

    dateDebut = frm!txtDateDebut.Value
    dateFin = frm!txtDateFin.Value
    prixUnitaire = CCur(Nz(frm!txtPrixUnitaire.Value, 0))
    nbPersonnes = CInt(Nz(frm!txtNbPersonnes.Value, 1))

    If Not IsNull(dateDebut) And Not IsNull(dateFin) And IsDate(dateDebut) And IsDate(dateFin) Then
        If CDate(dateFin) > CDate(dateDebut) Then
            Dim montant As Currency
            montant = CalculerSousTotal(CDate(dateDebut), CDate(dateFin), prixUnitaire, nbPersonnes)
            frm!lblMontantTotal.Caption = FormatMontant(montant)
        End If
    End If
    NouvelleRes_RecalculerMontant = 0
End Function

Public Function NouvelleRes_VerifierDisponibilite() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_NOUVELLE_RESERVATION)

    Dim idRess As Long
    Dim dateDebut As Date
    Dim dateFin As Date

    idRess = CLng(Nz(frm!cboRessource.Value, 0))

    If idRess = 0 Or IsNull(frm!txtDateDebut.Value) Or IsNull(frm!txtDateFin.Value) Then
        MsgBox "Sélectionnez une ressource et des dates pour vérifier la disponibilité.", vbExclamation, APP_NOM
        NouvelleRes_VerifierDisponibilite = 0
        Exit Function
    End If

    dateDebut = CDate(frm!txtDateDebut.Value)
    dateFin = CDate(frm!txtDateFin.Value)

    Dim idResIgnore As Long
    idResIgnore = CLng(Nz(frm!txtIdReservation.Value, 0))

    If EstRessourceDisponible(idRess, dateDebut, dateFin, idResIgnore) Then
        frm!lblResultatDispo.Caption = Chr(10003) & "  Disponible - Aucun conflit sur cette période"
        frm!lblResultatDispo.ForeColor = COULEUR_SUCCESS
        frm!lblResultatDispo.BackColor = CouleurAccess(39, 174, 96) - 10000000
    Else
        frm!lblResultatDispo.Caption = Chr(10007) & "  Indisponible - Conflit détecté sur cette période"
        frm!lblResultatDispo.ForeColor = COULEUR_DANGER
    End If

    NouvelleRes_VerifierDisponibilite = 0
End Function

Public Function NouvelleRes_Enregistrer() As Integer
    On Error GoTo GestionErreur
    Dim frm As Form
    Set frm = Forms(FORM_NOUVELLE_RESERVATION)

    ' Valider les champs
    If IsNull(frm!cboClient.Value) Or CLng(Nz(frm!cboClient.Value, 0)) = 0 Then
        MsgBox "Sélectionnez un client.", vbExclamation, APP_NOM
        NouvelleRes_Enregistrer = 0
        Exit Function
    End If

    If IsNull(frm!cboRessource.Value) Or CLng(Nz(frm!cboRessource.Value, 0)) = 0 Then
        MsgBox "Sélectionnez une ressource.", vbExclamation, APP_NOM
        NouvelleRes_Enregistrer = 0
        Exit Function
    End If

    If IsNull(frm!txtDateDebut.Value) Or IsNull(frm!txtDateFin.Value) Then
        MsgBox "Saisissez les dates de début et de fin.", vbExclamation, APP_NOM
        NouvelleRes_Enregistrer = 0
        Exit Function
    End If

    Dim idClient As Long
    Dim idRessource As Long
    Dim dateDebut As Date
    Dim dateFin As Date
    Dim nbPers As Integer
    Dim obs As String

    idClient = CLng(frm!cboClient.Value)
    idRessource = CLng(frm!cboRessource.Value)
    dateDebut = CDate(frm!txtDateDebut.Value)
    dateFin = CDate(frm!txtDateFin.Value)
    nbPers = CInt(Nz(frm!txtNbPersonnes.Value, 1))
    obs = Trim(Nz(frm!txtObservation.Value, ""))

    ' Vérifier disponibilité avant enregistrement
    If Not EstRessourceDisponible(idRessource, dateDebut, dateFin) Then
        MsgBox "Cette ressource est indisponible pour la période sélectionnée.", vbCritical, APP_NOM
        NouvelleRes_Enregistrer = 0
        Exit Function
    End If

    ' Créer la réservation
    Dim idReservation As Long
    idReservation = CreerReservation(idClient, dateDebut, dateFin, nbPers, obs)

    If idReservation > 0 Then
        ' Ajouter le détail
        AjouterDetailReservation idReservation, idRessource, 1

        ' Générer le billet automatiquement
        GenererBillet idReservation

        AfficherSucces "Réservation #" & idReservation & " créée avec succès !" & vbCrLf & _
                       "Un billet a été généré automatiquement."

        ' Revenir à la liste
        Navigation_Reservations
    End If

    NouvelleRes_Enregistrer = 0
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de l'enregistrement : " & Err.Description
    NouvelleRes_Enregistrer = 0
End Function

Public Function NouvelleRes_Reinitialiser() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_NOUVELLE_RESERVATION)

    frm!cboClient.Value = Null
    frm!cboCategorie.Value = Null
    frm!cboRessource.Value = Null
    frm!txtDateDebut.Value = Null
    frm!txtDateFin.Value = Null
    frm!txtNbPersonnes.Value = 1
    frm!txtPrixUnitaire.Value = 0
    frm!txtObservation.Value = ""
    frm!lblResultatDispo.Caption = ""
    frm!lblMontantTotal.Caption = "0 GNF"
    frm!txtIdReservation.Value = 0
    NouvelleRes_Reinitialiser = 0
End Function

' ============================================================
' Événements liste réservations
' ============================================================
Public Function Reservations_Filtrer() As Integer
    On Error Resume Next
    Reservations_Actualiser
    Reservations_Filtrer = 0
End Function

Public Function Reservations_Actualiser() As Integer
    On Error Resume Next
    Forms(FORM_RESERVATIONS)!lstReservations.RowSource = RQ_RESERVATIONS_COMPLETES
    Forms(FORM_RESERVATIONS)!lstReservations.Requery
    Dim nb As Long
    nb = CompterEnregistrements("RESERVATION", "")
    Forms(FORM_RESERVATIONS)!lblCompteur.Caption = nb & " réservation(s)"
    Reservations_Actualiser = 0
End Function

Public Function Reservations_OuvrirDetail() As Integer
    On Error Resume Next
    Dim idRes As Long
    idRes = CLng(Nz(Forms(FORM_RESERVATIONS)!lstReservations.Value, 0))
    If idRes > 0 Then
        MsgBox ObtenirResumeReservation(idRes), vbInformation, "Réservation #" & idRes
    End If
    Reservations_OuvrirDetail = 0
End Function

Public Function Reservations_Confirmer() As Integer
    On Error Resume Next
    Dim idRes As Long
    idRes = CLng(Nz(Forms(FORM_RESERVATIONS)!lstReservations.Value, 0))
    If idRes > 0 Then
        If ConfirmerReservation(idRes) Then
            AfficherSucces "Réservation #" & idRes & " confirmée."
            Reservations_Actualiser
        End If
    Else
        MsgBox "Sélectionnez une réservation.", vbExclamation, APP_NOM
    End If
    Reservations_Confirmer = 0
End Function

Public Function Reservations_Annuler() As Integer
    On Error Resume Next
    Dim idRes As Long
    idRes = CLng(Nz(Forms(FORM_RESERVATIONS)!lstReservations.Value, 0))
    If idRes > 0 Then
        If AnnulerReservation(idRes) Then
            Reservations_Actualiser
        End If
    Else
        MsgBox "Sélectionnez une réservation.", vbExclamation, APP_NOM
    End If
    Reservations_Annuler = 0
End Function

Public Function Reservations_GenererBillet() As Integer
    On Error Resume Next
    Dim idRes As Long
    idRes = CLng(Nz(Forms(FORM_RESERVATIONS)!lstReservations.Value, 0))
    If idRes > 0 Then
        GenererBillet idRes
    Else
        MsgBox "Sélectionnez une réservation.", vbExclamation, APP_NOM
    End If
    Reservations_GenererBillet = 0
End Function
