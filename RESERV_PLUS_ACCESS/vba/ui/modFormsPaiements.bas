Attribute VB_Name = "modFormsPaiements"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Formulaire F_PAIEMENTS
' Module : modFormsPaiements.bas
' ============================================================

Public Sub Creer_F_PAIEMENTS()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_PAIEMENTS

    SupprimerObjetAccess acForm, nomFrm
    Set frm = CreateForm()

    frm.Width = 16000
    frm.InsideHeight = 9200
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.BorderStyle = 1
    frm.AutoCenter = True
    frm.Detail.BackColor = COULEUR_FOND

    AjouterSidebar frm, "Paiements"

    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200
    Dim largeurContenu As Long
    largeurContenu = 16000 - xDebut - 200

    ' Bandeau vert teal (paiements)
    AjouterRectangle frm, xDebut, 0, largeurContenu, 1000, CouleurAccess(22, 160, 133)
    AjouterLabel frm, "lblTitrePage", "Paiements", xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Encaissements et suivi des soldes", _
                 xDebut + 200, 680, 8000, 300, 9, CouleurAccess(180, 230, 220), False

    ' Cartes statistiques
    Dim yCartes As Long
    yCartes = 1200
    Dim largeurC As Long
    largeurC = (largeurContenu - 800) \ 3

    AjouterCarteStat frm, "TotalEnc", "0 GNF", "Total encaissé (GNF)", _
                     xDebut + 200, yCartes, largeurC, 1200, CouleurAccess(22, 160, 133)
    AjouterCarteStat frm, "ResteEnc", "0 GNF", "Reste à encaisser", _
                     xDebut + 200 + largeurC + 200, yCartes, largeurC, 1200, COULEUR_WARNING
    AjouterCarteStat frm, "NbPaiMois", "0", "Paiements ce mois", _
                     xDebut + 200 + (largeurC + 200) * 2, yCartes, largeurC, 1200, COULEUR_PRIMAIRE

    ' Zone enregistrer paiement
    Dim yEnreg As Long
    yEnreg = yCartes + 1400

    AjouterRectangle frm, xDebut + 200, yEnreg, largeurContenu - 400, 1200, COULEUR_BLANC
    AjouterLabel frm, "lblSectionEnreg", "Enregistrer un paiement", _
                 xDebut + 400, yEnreg + 120, 5000, 380, 11, COULEUR_TEXTE, True

    ' Réservation
    AjouterLabel frm, "lblResChoix", "Réservation *", xDebut + 400, yEnreg + 600, 2000, 280, 8, COULEUR_TEXTE, False
    Dim cboRes As Access.ComboBox
    Set cboRes = AjouterCombo(frm, "cboReservationPai", xDebut + 400, yEnreg + 900, 3000, 480)
    cboRes.RowSourceType = "Table/Query"
    cboRes.RowSource = "SELECT R.id_reservation, '#' & R.id_reservation & ' - ' & C.nom & ' ' & C.prenom " & _
                        "FROM RESERVATION AS R INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
                        "WHERE R.statut <> '" & STATUT_ANNULEE & "' ORDER BY R.id_reservation DESC"
    cboRes.BoundColumn = 1
    cboRes.ColumnCount = 2
    cboRes.ColumnWidths = "0;5000"
    cboRes.OnChange = "=Paiements_AfficherSolde()"

    ' Montant
    AjouterLabel frm, "lblMontPai", "Montant *", xDebut + 3600, yEnreg + 600, 2000, 280, 8, COULEUR_TEXTE, False
    Dim txtMont As Access.TextBox
    Set txtMont = AjouterZoneTexte(frm, "txtMontantPai", xDebut + 3600, yEnreg + 900, 2400, 480)
    txtMont.Format = "#,##0"

    ' Mode paiement
    AjouterLabel frm, "lblModePai", "Mode *", xDebut + 6200, yEnreg + 600, 2000, 280, 8, COULEUR_TEXTE, False
    Dim cboMode As Access.ComboBox
    Set cboMode = AjouterCombo(frm, "cboModePaiement", xDebut + 6200, yEnreg + 900, 2400, 480)
    cboMode.RowSource = MODE_ESPECES & ";" & MODE_CARTE & ";" & MODE_VIREMENT & ";" & MODE_MOBILE

    ' Solde restant
    AjouterLabel frm, "lblSoldeRes", "Solde restant :", xDebut + 9000, yEnreg + 600, 2000, 280, 8, COULEUR_TEXTE, False
    AjouterLabel frm, "lblSoldeValeur", "---", xDebut + 9000, yEnreg + 900, 2000, 380, 11, COULEUR_DANGER, True

    ' Bouton encaisser
    Dim btnEncaisser As Access.CommandButton
    Set btnEncaisser = AjouterBouton(frm, "btnEncaisser", "$ Encaisser", _
                                      xDebut + largeurContenu - 2400, yEnreg + 850, 2200, 600, _
                                      CouleurAccess(22, 160, 133), COULEUR_BLANC)
    btnEncaisser.FontSize = 12
    btnEncaisser.OnClick = "=Paiements_Enregistrer()"

    ' Liste des paiements
    Dim yListe As Long
    yListe = yEnreg + 1400

    AjouterRectangle frm, xDebut + 200, yListe, largeurContenu - 400, 480, CouleurAccess(236, 240, 241)
    AjouterLabel frm, "lblColN2", "N°", xDebut + 300, yListe + 80, 700, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColRes2", "Réservation", xDebut + 1100, yListe + 80, 1600, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColCli2", "Client", xDebut + 2800, yListe + 80, 2800, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColMont2", "Montant", xDebut + 5700, yListe + 80, 2000, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColMode2", "Mode", xDebut + 7800, yListe + 80, 1800, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColSolde2", "Solde", xDebut + 9700, yListe + 80, 1800, 320, 8, COULEUR_TEXTE, True
    AjouterLabel frm, "lblColEtat2", "État", xDebut + 11600, yListe + 80, 1600, 320, 8, COULEUR_TEXTE, True

    Dim lstPai As Access.ListBox
    Set lstPai = CreateControl(frm.Name, acListBox, 0, , , _
                  xDebut + 200, yListe + 480, largeurContenu - 400, 5000)
    lstPai.Name = "lstPaiements"
    lstPai.FontName = POLICE
    lstPai.FontSize = 9
    lstPai.BackColor = COULEUR_BLANC
    lstPai.ForeColor = COULEUR_TEXTE
    lstPai.BorderStyle = 0
    lstPai.RowSourceType = "Table/Query"
    lstPai.RowSource = RQ_PAIEMENTS_COMPLETS
    lstPai.BoundColumn = 1
    lstPai.ColumnCount = 9
    lstPai.ColumnWidths = "700;1600;2800;2000;1800;1800;1600;0;0"

    ' Bouton générer reçu
    Dim yBtnsPai As Long
    yBtnsPai = yListe + 5660

    Dim btnRecu As Access.CommandButton
    Set btnRecu = AjouterBouton(frm, "btnGenererRecu", "Générer reçu", _
                                  xDebut + 200, yBtnsPai, 2000, 480, CouleurAccess(142, 68, 173), COULEUR_BLANC)
    btnRecu.OnClick = "=Paiements_GenererRecu()"

    Dim btnActuPai As Access.CommandButton
    Set btnActuPai = AjouterBouton(frm, "btnActualiser", "Actualiser", _
                                    xDebut + 2400, yBtnsPai, 1600, 480, CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnActuPai.FontBold = False
    btnActuPai.OnClick = "=Paiements_Actualiser()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ---- Événements ----
Public Function Paiements_Actualiser() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_PAIEMENTS)
    frm!lstPaiements.RowSource = RQ_PAIEMENTS_COMPLETS
    frm!lstPaiements.Requery
    frm!lblTotalEnc_Val.Caption = FormatMontant(TotalEncaisseGlobal())
    frm!lblResteEnc_Val.Caption = FormatMontant(TotalResteAEncaisser())
    frm!lblNbPaiMois_Val.Caption = CStr(NbPaiementsMois())
    Paiements_Actualiser = 0
End Function

Public Function Paiements_AfficherSolde() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_PAIEMENTS)
    Dim idRes As Long
    idRes = CLng(Nz(frm!cboReservationPai.Value, 0))
    If idRes > 0 Then
        Dim solde As Currency
        solde = SoldeReservation(idRes)
        frm!lblSoldeValeur.Caption = FormatMontant(solde)
        frm!lblSoldeValeur.ForeColor = IIf(solde <= 0, COULEUR_SUCCESS, COULEUR_DANGER)
    Else
        frm!lblSoldeValeur.Caption = "---"
    End If
    Paiements_AfficherSolde = 0
End Function

Public Function Paiements_Enregistrer() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_PAIEMENTS)

    Dim idRes As Long
    idRes = CLng(Nz(frm!cboReservationPai.Value, 0))
    Dim montant As Currency
    montant = CCur(Nz(frm!txtMontantPai.Value, 0))
    Dim mode As String
    mode = ChaineSure(frm!cboModePaiement.Value)

    If idRes = 0 Then
        MsgBox "Sélectionnez une réservation.", vbExclamation, APP_NOM
        Paiements_Enregistrer = 0
        Exit Function
    End If

    If montant <= 0 Then
        MsgBox "Le montant doit être supérieur à zéro.", vbExclamation, APP_NOM
        Paiements_Enregistrer = 0
        Exit Function
    End If

    Dim idPai As Long
    idPai = EnregistrerPaiement(idRes, montant, mode)
    If idPai > 0 Then
        AfficherSucces "Paiement enregistré (réf. auto)."
        frm!cboReservationPai.Value = Null
        frm!txtMontantPai.Value = Null
        frm!cboModePaiement.Value = Null
        frm!lblSoldeValeur.Caption = "---"
        Paiements_Actualiser
    End If
    Paiements_Enregistrer = 0
End Function

Public Function Paiements_GenererRecu() As Integer
    On Error Resume Next
    Dim idPai As Long
    idPai = CLng(Nz(Forms(FORM_PAIEMENTS)!lstPaiements.Value, 0))

    If idPai = 0 Then
        MsgBox "Sélectionnez un paiement dans la liste.", vbExclamation, APP_NOM
        Paiements_GenererRecu = 0
        Exit Function
    End If

    ' Afficher un aperçu texte du reçu
    Dim rs As DAO.Recordset
    Dim sql As String
    sql = "SELECT P.id_paiement, P.date_paiement, P.montant, P.mode_paiement, " & _
          "P.reference_paiement, R.id_reservation, C.nom & ' ' & C.prenom AS client, " & _
          "C.telephone, R.montant_total " & _
          "FROM (PAIEMENT AS P INNER JOIN RESERVATION AS R ON P.id_reservation = R.id_reservation) " & _
          "INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
          "WHERE P.id_paiement = " & idPai

    Set rs = OuvrirRecordset(sql)
    If Not rs.EOF Then
        Dim nomEtab As String
        nomEtab = ObtenirParametre("NOM_ETABLISSEMENT", "RESERV+")
        Dim texte As String
        texte = String(50, "=") & vbCrLf
        texte = texte & "          " & nomEtab & vbCrLf
        texte = texte & "         REÇU DE PAIEMENT" & vbCrLf
        texte = texte & String(50, "=") & vbCrLf & vbCrLf
        texte = texte & "Réf. Paiement : " & ChaineSure(rs!reference_paiement) & vbCrLf
        texte = texte & "Date : " & FormatDateHeure(rs!date_paiement) & vbCrLf & vbCrLf
        texte = texte & "CLIENT : " & ChaineSure(rs!client) & vbCrLf
        texte = texte & "Tél : " & ChaineSure(rs!telephone) & vbCrLf & vbCrLf
        texte = texte & "Réservation #" & ChaineSure(rs!id_reservation) & vbCrLf
        texte = texte & "Montant total rés. : " & FormatMontant(CCur(Nz(rs!montant_total, 0))) & vbCrLf
        texte = texte & "MONTANT VERSÉ : " & FormatMontant(CCur(Nz(rs!montant, 0))) & vbCrLf
        texte = texte & "Mode : " & ChaineSure(rs!mode_paiement) & vbCrLf & vbCrLf
        texte = texte & "Solde : " & FormatMontant(SoldeReservation(CLng(rs!id_reservation))) & vbCrLf
        texte = texte & String(50, "=") & vbCrLf
        texte = texte & "Merci pour votre confiance !"
        MsgBox texte, vbInformation, "Reçu de paiement"
        JournaliserConsultation "Reçu paiement #" & idPai
    End If
    rs.Close
    Paiements_GenererRecu = 0
End Function
