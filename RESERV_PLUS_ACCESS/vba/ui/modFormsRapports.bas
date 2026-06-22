Attribute VB_Name = "modFormsRapports"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Formulaire F_ETATS_RAPPORTS
' Module : modFormsRapports.bas
' ============================================================

Public Sub Creer_F_ETATS_RAPPORTS()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_RAPPORTS

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

    AjouterSidebar frm, "Etats"

    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200
    Dim largeurContenu As Long
    largeurContenu = 16000 - xDebut - 200

    ' Bandeau rouge-orangé (états)
    AjouterRectangle frm, xDebut, 0, largeurContenu, 1000, CouleurAccess(192, 57, 43)
    AjouterLabel frm, "lblTitrePage", "États et rapports", xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Analyse et exports de données", _
                 xDebut + 200, 680, 8000, 300, 9, CouleurAccess(230, 200, 200), False

    ' Filtres de date
    Dim yFiltres As Long
    yFiltres = 1200

    AjouterRectangle frm, xDebut + 200, yFiltres, largeurContenu - 400, 1000, COULEUR_BLANC

    AjouterLabel frm, "lblDateDeb", "Du :", xDebut + 400, yFiltres + 200, 600, 280, 8, COULEUR_TEXTE, False
    Dim txtDeb As Access.TextBox
    Set txtDeb = AjouterZoneTexte(frm, "txtDateDebRapport", xDebut + 1000, yFiltres + 160, 2000, 480)
    txtDeb.Format = "dd/mm/yyyy"
    txtDeb.Value = Format(DateSerial(Year(Now()), 1, 1), "dd/mm/yyyy")

    AjouterLabel frm, "lblDateFinR", "Au :", xDebut + 3200, yFiltres + 200, 600, 280, 8, COULEUR_TEXTE, False
    Dim txtFin As Access.TextBox
    Set txtFin = AjouterZoneTexte(frm, "txtDateFinRapport", xDebut + 3800, yFiltres + 160, 2000, 480)
    txtFin.Format = "dd/mm/yyyy"
    txtFin.Value = Format(Now(), "dd/mm/yyyy")

    AjouterLabel frm, "lblStatutFiltre", "Statut :", xDebut + 6000, yFiltres + 200, 800, 280, 8, COULEUR_TEXTE, False
    Dim cboStatut As Access.ComboBox
    Set cboStatut = AjouterCombo(frm, "cboStatutRapport", xDebut + 6900, yFiltres + 160, 2400, 480)
    cboStatut.RowSource = "Toutes;" & STATUT_CONFIRMEE & ";" & STATUT_EN_ATTENTE & ";" & STATUT_ANNULEE
    cboStatut.Value = "Toutes"

    Dim btnAppliquer As Access.CommandButton
    Set btnAppliquer = AjouterBouton(frm, "btnAppliquerFiltre", "Appliquer", _
                                      xDebut + 9600, yFiltres + 160, 1800, 480, COULEUR_PRIMAIRE, COULEUR_BLANC)
    btnAppliquer.FontSize = 9
    btnAppliquer.OnClick = "=Rapports_Actualiser()"

    ' Boutons rapports disponibles
    Dim yBtnsRapports As Long
    yBtnsRapports = yFiltres + 1200

    AjouterLabel frm, "lblTitreBtnRap", "Rapports disponibles :", _
                 xDebut + 200, yBtnsRapports, 5000, 380, 11, COULEUR_TEXTE, True

    Dim yBR As Long
    yBR = yBtnsRapports + 560

    Dim btnListeRes As Access.CommandButton
    Set btnListeRes = AjouterBouton(frm, "btnListeReservations", Chr(9632) & " Liste réservations", _
                                     xDebut + 200, yBR, 3200, 480, COULEUR_PRIMAIRE, COULEUR_BLANC)
    btnListeRes.FontSize = 9
    btnListeRes.OnClick = "=Rapports_ListeReservations()"

    Dim btnCA As Access.CommandButton
    Set btnCA = AjouterBouton(frm, "btnChiffreAffaires", "$ Chiffre d'affaires", _
                               xDebut + 3600, yBR, 3200, 480, COULEUR_SUCCESS, COULEUR_BLANC)
    btnCA.FontSize = 9
    btnCA.OnClick = "=Rapports_ChiffreAffaires()"

    Dim btnTaux As Access.CommandButton
    Set btnTaux = AjouterBouton(frm, "btnTauxOccupation", Chr(9632) & " Taux d'occupation", _
                                  xDebut + 7000, yBR, 3200, 480, COULEUR_WARNING, COULEUR_BLANC)
    btnTaux.FontSize = 9
    btnTaux.OnClick = "=Rapports_TauxOccupation()"

    Dim btnJournal As Access.CommandButton
    Set btnJournal = AjouterBouton(frm, "btnJournal", Chr(9632) & " Journal actions", _
                                    xDebut + 10400, yBR, 3200, 480, CouleurAccess(44, 62, 80), COULEUR_BLANC)
    btnJournal.FontSize = 9
    btnJournal.OnClick = "=Rapports_Journal()"

    ' Zone résultats (liste)
    Dim yResultats As Long
    yResultats = yBR + 700

    AjouterRectangle frm, xDebut + 200, yResultats, largeurContenu - 400, 480, _
                     CouleurAccess(236, 240, 241)
    AjouterLabel frm, "lblColRes1", "Résultats du rapport sélectionné", _
                 xDebut + 400, yResultats + 100, largeurContenu - 800, 280, 10, COULEUR_TEXTE, True

    Dim lstResultats As Access.ListBox
    Set lstResultats = CreateControl(frm.Name, acListBox, 0, , , _
                        xDebut + 200, yResultats + 480, largeurContenu - 400, 5000)
    lstResultats.Name = "lstResultats"
    lstResultats.FontName = POLICE
    lstResultats.FontSize = 9
    lstResultats.BackColor = COULEUR_BLANC
    lstResultats.BorderStyle = 0
    lstResultats.RowSourceType = "Table/Query"
    lstResultats.RowSource = RQ_RESERVATIONS_COMPLETES
    lstResultats.ColumnCount = 8
    lstResultats.ColumnWidths = "700;2800;2400;2000;2000;2000;0;0"

    ' Zone synthèse
    Dim yStats As Long
    yStats = yResultats + 5660

    AjouterLabel frm, "lblSynthese", "", xDebut + 200, yStats, largeurContenu - 400, 400, _
                 9, COULEUR_TEXTE_SECONDAIRE, False

    ' Bouton exporter
    Dim btnExport As Access.CommandButton
    Set btnExport = AjouterBouton(frm, "btnExporter", "Exporter (Excel)", _
                                   xDebut + largeurContenu - 3200, yStats - 100, 3000, 480, _
                                   CouleurAccess(39, 174, 96), COULEUR_BLANC)
    btnExport.OnClick = "=Rapports_Exporter()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ---- Événements rapports ----
Public Function Rapports_Actualiser() As Integer
    On Error Resume Next
    Forms(FORM_RAPPORTS)!lstResultats.Requery
    Rapports_Actualiser = 0
End Function

Public Function Rapports_ListeReservations() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_RAPPORTS)

    Dim statut As String
    statut = ChaineSure(frm!cboStatutRapport.Value)
    Dim sql As String

    If statut = "Toutes" Or Len(statut) = 0 Then
        sql = RQ_RESERVATIONS_COMPLETES
    Else
        sql = "SELECT R.id_reservation, C.nom & ' ' & C.prenom AS client_nom, " & _
              "R.date_debut, R.date_fin, R.nb_personnes, R.montant_total, R.statut, R.date_creation " & _
              "FROM RESERVATION AS R INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
              "WHERE R.statut = '" & EchapperChaine(statut) & "' ORDER BY R.date_creation DESC"
    End If

    frm!lstResultats.ColumnCount = 8
    frm!lstResultats.ColumnWidths = "700;3000;2000;2000;1000;2000;2000;2000"
    frm!lstResultats.RowSource = sql
    frm!lstResultats.Requery

    Dim nb As Long
    nb = CompterEnregistrements("RESERVATION", IIf(statut = "Toutes", "", "statut = '" & EchapperChaine(statut) & "'"))
    frm!lblSynthese.Caption = "Total : " & nb & " réservation(s) | Statut : " & IIf(Len(statut) > 0, statut, "Toutes")
    JournaliserConsultation "Liste réservations (statut=" & statut & ")"
    Rapports_ListeReservations = 0
End Function

Public Function Rapports_ChiffreAffaires() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_RAPPORTS)

    frm!lstResultats.ColumnCount = 3
    frm!lstResultats.ColumnWidths = "3000;2000;3000"
    frm!lstResultats.RowSource = RQ_CA_PAR_CATEGORIE
    frm!lstResultats.Requery

    Dim totalCA As Currency
    totalCA = CCur(Nz(ObtenirValeur("SELECT Sum(sous_total) FROM DETAIL_RESERVATION DR " & _
              "INNER JOIN RESERVATION R ON DR.id_reservation = R.id_reservation " & _
              "WHERE R.statut <> '" & STATUT_ANNULEE & "'"), 0))
    frm!lblSynthese.Caption = "Chiffre d'affaires total : " & FormatMontant(totalCA)
    JournaliserConsultation "Chiffre d'affaires par catégorie"
    Rapports_ChiffreAffaires = 0
End Function

Public Function Rapports_TauxOccupation() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_RAPPORTS)

    Dim dateDebut As Date
    Dim dateFin As Date
    dateDebut = CDate(Nz(frm!txtDateDebRapport.Value, Date - 30))
    dateFin = CDate(Nz(frm!txtDateFinRapport.Value, Date))

    Dim rs As DAO.Recordset
    Set rs = RapportTauxOccupation(dateDebut, dateFin)
    frm!lstResultats.RowSourceType = "Table/Query"
    frm!lstResultats.ColumnCount = 5
    frm!lstResultats.ColumnWidths = "0;3500;2500;2500;2000"

    Dim sqlTaux As String
    sqlTaux = "SELECT RS.id_ressource, RS.libelle, CAT.libelle_categorie, " & _
              "P.nom_entreprise, Count(DR.id_detail) AS nb_sejours " & _
              "FROM ((DETAIL_RESERVATION AS DR INNER JOIN RESSOURCE AS RS ON DR.id_ressource = RS.id_ressource) " & _
              "INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie) " & _
              "INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire " & _
              "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
              "WHERE R.statut <> '" & STATUT_ANNULEE & "' " & _
              "GROUP BY RS.id_ressource, RS.libelle, CAT.libelle_categorie, P.nom_entreprise " & _
              "ORDER BY nb_sejours DESC"
    frm!lstResultats.RowSource = sqlTaux
    frm!lstResultats.Requery
    frm!lblSynthese.Caption = "Taux d'occupation — Période : " & FormatDate(dateDebut) & " au " & FormatDate(dateFin)
    JournaliserConsultation "Taux occupation"
    Rapports_TauxOccupation = 0
End Function

Public Function Rapports_Journal() As Integer
    On Error Resume Next
    If Not ALeRole(ROLE_ADMINISTRATEUR) Then
        MsgBox "Le journal est réservé à l'administrateur.", vbExclamation, APP_NOM
        Rapports_Journal = 0
        Exit Function
    End If
    Dim frm As Form
    Set frm = Forms(FORM_RAPPORTS)
    frm!lstResultats.ColumnCount = 5
    frm!lstResultats.ColumnWidths = "0;3000;2000;3000;2000"
    frm!lstResultats.RowSource = RQ_JOURNAL_ACTIONS
    frm!lstResultats.Requery
    Dim nb As Long
    nb = CompterEnregistrements("JOURNAL_ACTION", "")
    frm!lblSynthese.Caption = "Journal des actions — " & nb & " entrée(s)"
    Rapports_Journal = 0
End Function

Public Function Rapports_Exporter() As Integer
    MsgBox "Pour exporter : Clic-droit sur la liste → Exporter → Excel." & vbCrLf & _
           "Ou utilisez Données Externes → Exporter dans le ruban Access.", _
           vbInformation, APP_NOM & " — Export"
    Rapports_Exporter = 0
End Function
