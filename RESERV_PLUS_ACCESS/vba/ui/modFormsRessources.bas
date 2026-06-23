Attribute VB_Name = "modFormsRessources"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Formulaires de gestion des ressources
' Module : modFormsRessources.bas
' ============================================================

Public Sub Creer_F_RESSOURCES()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_RESSOURCES

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

    AjouterSidebar frm, "Ressources"

    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200
    Dim largeurContenu As Long
    largeurContenu = 16000 - xDebut - 200

    ' Bandeau vert (couleur ressources)
    AjouterRectangle frm, xDebut, 0, largeurContenu, 1000, CouleurAccess(39, 174, 96)
    AjouterLabel frm, "lblTitrePage", "Ressources", xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Hôtels, taxis et vols des prestataires", _
                 xDebut + 200, 680, 8000, 300, 9, CouleurAccess(180, 230, 180), False

    ' Filtres par catégorie (onglets)
    Dim yFiltres As Long
    yFiltres = 1200

    Dim labelsFiltres() As String
    labelsFiltres = Split("Tous|Hôtels|Taxis|Vols", "|")
    Dim xFiltre As Long
    xFiltre = xDebut + 200

    Dim fi As Integer
    For fi = 0 To 3
        Dim btnFi As Access.CommandButton
        Set btnFi = AjouterBouton(frm, "btnCat" & fi, labelsFiltres(fi), _
                                   xFiltre, yFiltres, 1400, 480, _
                                   IIf(fi = 0, CouleurAccess(39, 174, 96), CouleurAccess(236, 240, 241)), _
                                   IIf(fi = 0, COULEUR_BLANC, COULEUR_TEXTE))
        btnFi.FontSize = 9
        btnFi.FontBold = (fi = 0)
        btnFi.OnClick = "=Ressources_FiltrerCategorie(" & fi & ")"
        xFiltre = xFiltre + 1500
    Next fi

    Dim btnNouvelleRess As Access.CommandButton
    Set btnNouvelleRess = AjouterBouton(frm, "btnAjouterRessource", "+ Nouvelle ressource", _
                                         xDebut + largeurContenu - 2800, yFiltres, 2600, 480, _
                                         CouleurAccess(39, 174, 96), COULEUR_BLANC)
    btnNouvelleRess.OnClick = "=Ressources_OuvrirFicheNouveau()"

    ' Liste ressources
    Dim lstRess As Access.ListBox
    Set lstRess = CreateControl(frm.Name, acListBox, 0, , , _
                   xDebut + 200, yFiltres + 680, largeurContenu - 400, 4800)
    lstRess.Name = "lstRessources"
    lstRess.FontName = POLICE
    lstRess.FontSize = 9
    lstRess.BackColor = COULEUR_BLANC
    lstRess.ForeColor = COULEUR_TEXTE
    lstRess.BorderStyle = 0
    lstRess.RowSourceType = "Table/Query"
    lstRess.RowSource = RQ_RESSOURCES_PRESTATAIRE
    lstRess.BoundColumn = 1
    lstRess.ColumnCount = 8
    lstRess.ColumnWidths = "0;2000;3000;3500;1500;1500;800;1200"
    lstRess.OnDblClick = "=Ressources_OuvrirFicheExistant()"

    ' Boutons actions
    Dim yBtns As Long
    yBtns = yFiltres + 5660

    Dim btnMod2 As Access.CommandButton
    Set btnMod2 = AjouterBouton(frm, "btnModifier", "Modifier", _
                                  xDebut + 200, yBtns, 1600, 480, CouleurAccess(52, 152, 219), COULEUR_BLANC)
    btnMod2.OnClick = "=Ressources_OuvrirFicheExistant()"

    Dim btnDispo As Access.CommandButton
    Set btnDispo = AjouterBouton(frm, "btnBascDispo", "Disponibilité", _
                                  xDebut + 2000, yBtns, 1800, 480, COULEUR_WARNING, COULEUR_BLANC)
    btnDispo.OnClick = "=Ressources_BasculerDispo()"

    Dim btnSupp3 As Access.CommandButton
    Set btnSupp3 = AjouterBouton(frm, "btnSupprimer", "Supprimer", _
                                   xDebut + 4000, yBtns, 1600, 480, COULEUR_DANGER, COULEUR_BLANC)
    btnSupp3.OnClick = "=Ressources_Supprimer()"

    Dim btnActu3 As Access.CommandButton
    Set btnActu3 = AjouterBouton(frm, "btnActualiser", "Actualiser", _
                                   xDebut + 5800, yBtns, 1600, 480, CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnActu3.FontBold = False
    btnActu3.OnClick = "=Ressources_Actualiser()"

    ' Fiche ressource
    Dim yFiche As Long
    yFiche = yBtns + 700

    AjouterRectangle frm, xDebut + 200, yFiche, largeurContenu - 400, 2400, COULEUR_BLANC
    AjouterLabel frm, "lblTitreFiche", "Fiche ressource", xDebut + 400, yFiche + 120, 5000, 380, 11, COULEUR_TEXTE, True

    Dim largC3 As Long
    largC3 = (largeurContenu - 1400) \ 3

    ' Ligne 1 : Prestataire | Catégorie | Libellé
    AjouterLabel frm, "lblPrest", "Prestataire *", xDebut + 400, yFiche + 600, largC3, 280, 8, COULEUR_TEXTE, False
    Dim cboP As Access.ComboBox
    Set cboP = AjouterCombo(frm, "cboPrestataireRes", xDebut + 400, yFiche + 900, largC3, 480)
    cboP.RowSourceType = "Table/Query"
    cboP.RowSource = "SELECT id_prestataire, nom_entreprise FROM PRESTATAIRE ORDER BY nom_entreprise"
    cboP.BoundColumn = 1
    cboP.ColumnCount = 2
    cboP.ColumnWidths = "0;4000"

    AjouterLabel frm, "lblCategRes", "Catégorie *", xDebut + 400 + largC3 + 200, yFiche + 600, largC3, 280, 8, COULEUR_TEXTE, False
    Dim cboC2 As Access.ComboBox
    Set cboC2 = AjouterCombo(frm, "cboCategorieRes", xDebut + 400 + largC3 + 200, yFiche + 900, largC3, 480)
    cboC2.RowSourceType = "Table/Query"
    cboC2.RowSource = "SELECT id_categorie, libelle_categorie FROM CATEGORIE ORDER BY libelle_categorie"
    cboC2.BoundColumn = 1
    cboC2.ColumnCount = 2
    cboC2.ColumnWidths = "0;3000"

    AjouterLabel frm, "lblLib", "Libellé *", xDebut + 400 + (largC3 + 200) * 2, yFiche + 600, largC3, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtLibelle", xDebut + 400 + (largC3 + 200) * 2, yFiche + 900, largC3, 480

    ' Ligne 2 : Prix | Capacité | Disponible
    AjouterLabel frm, "lblPrixRes", "Prix unitaire", xDebut + 400, yFiche + 1500, largC3, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtPrixRes", xDebut + 400, yFiche + 1800, largC3, 480

    AjouterLabel frm, "lblCapRes", "Capacité", xDebut + 400 + largC3 + 200, yFiche + 1500, largC3, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtCapaciteRes", xDebut + 400 + largC3 + 200, yFiche + 1800, largC3 \ 2, 480

    Dim chkDispo As Access.CheckBox
    Set chkDispo = CreateControl(frm.Name, acCheckBox, 0, , , xDebut + 400 + (largC3 + 200) * 2, yFiche + 1800, 300, 360)
    chkDispo.Name = "chkDisponible"
    chkDispo.Value = True
    AjouterLabel frm, "lblDispoCk", "Disponible", xDebut + 800 + (largC3 + 200) * 2, yFiche + 1850, 1200, 280, 8, COULEUR_TEXTE, False

    Dim txtId3 As Access.TextBox
    Set txtId3 = AjouterZoneTexte(frm, "txtIdRessource", 0, 0, 100, 100)
    txtId3.Visible = False

    Dim btnSave3 As Access.CommandButton
    Set btnSave3 = AjouterBouton(frm, "btnEnregistrer", "Enregistrer", _
                                   xDebut + largeurContenu - 4200, yFiche + 1800, 1800, 480, COULEUR_SUCCESS, COULEUR_BLANC)
    btnSave3.OnClick = "=Ressources_Enregistrer()"

    Dim btnEff3 As Access.CommandButton
    Set btnEff3 = AjouterBouton(frm, "btnEffacer", "Effacer", _
                                  xDebut + largeurContenu - 2200, yFiche + 1800, 1800, 480, CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnEff3.FontBold = False
    btnEff3.OnClick = "=Ressources_EffacerFiche()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ---- Événements ----
Public Function Ressources_Actualiser() As Integer
    On Error Resume Next
    Forms(FORM_RESSOURCES)!lstRessources.RowSource = RQ_RESSOURCES_PRESTATAIRE
    Forms(FORM_RESSOURCES)!lstRessources.Requery
    Ressources_Actualiser = 0
End Function

Public Function Ressources_FiltrerCategorie(indexCat As Integer) As Integer
    On Error Resume Next
    Dim sql As String
    If indexCat = 0 Then
        sql = RQ_RESSOURCES_PRESTATAIRE
    Else
        Dim nomCat As String
        Select Case indexCat
            Case 1: nomCat = CAT_HOTEL
            Case 2: nomCat = CAT_TAXI
            Case 3: nomCat = CAT_VOL
        End Select
        sql = "SELECT RS.id_ressource, CAT.libelle_categorie, P.nom_entreprise, RS.libelle, " & _
              "RS.prix_unitaire, RS.capacite, RS.disponible, RS.statut_ressource " & _
              "FROM (RESSOURCE AS RS INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie) " & _
              "INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire " & _
              "WHERE CAT.libelle_categorie = '" & nomCat & "' ORDER BY RS.libelle"
    End If
    Forms(FORM_RESSOURCES)!lstRessources.RowSource = sql
    Forms(FORM_RESSOURCES)!lstRessources.Requery
    Ressources_FiltrerCategorie = 0
End Function

Public Function Ressources_OuvrirFicheNouveau() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_RESSOURCES)
    frm!txtIdRessource.Value = 0
    frm!cboPrestataireRes.Value = Null
    frm!cboCategorieRes.Value = Null
    frm!txtLibelle.Value = ""
    frm!txtPrixRes.Value = 0
    frm!txtCapaciteRes.Value = 1
    frm!chkDisponible.Value = True
    Ressources_OuvrirFicheNouveau = 0
End Function

Public Function Ressources_OuvrirFicheExistant() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_RESSOURCES)
    Dim idR As Long
    idR = CLng(Nz(frm!lstRessources.Value, 0))
    If idR = 0 Then Exit Function
    Dim rs As DAO.Recordset
    Set rs = OuvrirRecordset("SELECT * FROM RESSOURCE WHERE id_ressource = " & idR)
    If Not rs.EOF Then
        frm!txtIdRessource.Value = rs!id_ressource
        frm!cboPrestataireRes.Value = rs!id_prestataire
        frm!cboCategorieRes.Value = rs!id_categorie
        frm!txtLibelle.Value = Nz(rs!libelle, "")
        frm!txtPrixRes.Value = Nz(rs!prix_unitaire, 0)
        frm!txtCapaciteRes.Value = Nz(rs!capacite, 1)
        frm!chkDisponible.Value = Nz(rs!disponible, True)
    End If
    rs.Close
    Ressources_OuvrirFicheExistant = 0
End Function

Public Function Ressources_Enregistrer() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_RESSOURCES)
    Dim idR As Long
    idR = CLng(Nz(frm!txtIdRessource.Value, 0))

    If Len(Trim(Nz(frm!txtLibelle.Value, ""))) = 0 Then
        MsgBox "Le libellé est obligatoire.", vbExclamation, APP_NOM
        Ressources_Enregistrer = 0
        Exit Function
    End If

    If idR = 0 Then
        Dim newId As Long
        newId = CreerRessource(CLng(Nz(frm!cboPrestataireRes.Value, 0)), _
                               CLng(Nz(frm!cboCategorieRes.Value, 0)), _
                               Trim(Nz(frm!txtLibelle.Value, "")), "", _
                               CCur(Nz(frm!txtPrixRes.Value, 0)), _
                               CInt(Nz(frm!txtCapaciteRes.Value, 1)))
        If newId > 0 Then AfficherSucces "Ressource créée (ID=" & newId & ")."
    Else
        Dim sql As String
        sql = "UPDATE RESSOURCE SET libelle = '" & EchapperChaine(Trim(Nz(frm!txtLibelle.Value, ""))) & "', " & _
              "id_prestataire = " & CLng(Nz(frm!cboPrestataireRes.Value, 0)) & ", " & _
              "id_categorie = " & CLng(Nz(frm!cboCategorieRes.Value, 0)) & ", " & _
              "prix_unitaire = " & CCur(Nz(frm!txtPrixRes.Value, 0)) & ", " & _
              "capacite = " & CInt(Nz(frm!txtCapaciteRes.Value, 1)) & ", " & _
              "disponible = " & IIf(CBool(Nz(frm!chkDisponible.Value, True)), "True", "False") & ", " & _
              "statut_ressource = '" & IIf(CBool(Nz(frm!chkDisponible.Value, True)), "Disponible", "Indisponible") & "' " & _
              "WHERE id_ressource = " & idR
        ExecuterSQL sql
        JournaliserModification "Ressource", idR
        AfficherSucces "Ressource modifiée."
    End If
    Ressources_Actualiser
    Ressources_Enregistrer = 0
End Function

Public Function Ressources_BasculerDispo() As Integer
    On Error Resume Next
    Dim idR As Long
    idR = CLng(Nz(Forms(FORM_RESSOURCES)!lstRessources.Value, 0))
    If idR > 0 Then
        BasculerDisponibilite idR
        Ressources_Actualiser
    End If
    Ressources_BasculerDispo = 0
End Function

Public Function Ressources_Supprimer() As Integer
    On Error Resume Next
    Dim idR As Long
    idR = CLng(Nz(Forms(FORM_RESSOURCES)!lstRessources.Value, 0))
    If idR > 0 Then
        SupprimerRessource idR
        Ressources_Actualiser
    End If
    Ressources_Supprimer = 0
End Function

Public Function Ressources_EffacerFiche() As Integer
    Ressources_OuvrirFicheNouveau
    Ressources_EffacerFiche = 0
End Function
