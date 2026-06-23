Attribute VB_Name = "modFormsPrestataires"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Formulaire F_PRESTATAIRES
' Module : modFormsPrestataires.bas
' ============================================================

Public Sub Creer_F_PRESTATAIRES()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_PRESTATAIRES

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

    AjouterSidebar frm, "Prestataires"

    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200
    Dim largeurContenu As Long
    largeurContenu = 16000 - xDebut - 200

    ' Bandeau
    AjouterRectangle frm, xDebut, 0, largeurContenu, 1000, CouleurAccess(52, 73, 94)
    AjouterLabel frm, "lblTitrePage", "Prestataires", xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitre", "Propriétaires des ressources gérées", _
                 xDebut + 200, 680, 8000, 300, 9, CouleurAccess(200, 210, 220), False

    ' Recherche + Ajouter
    Dim yRech As Long
    yRech = 1200

    Dim txtRech As Access.TextBox
    Set txtRech = AjouterZoneTexte(frm, "txtRecherche", xDebut + 200, yRech, 5000, 480)
    txtRech.SpecialEffect = 0

    Dim btnRech As Access.CommandButton
    Set btnRech = AjouterBouton(frm, "btnRechercher", "Rechercher", _
                                 xDebut + 5400, yRech, 1600, 480, CouleurAccess(100, 100, 130), COULEUR_BLANC)
    btnRech.FontSize = 9
    btnRech.FontBold = False
    btnRech.OnClick = "=Prestataires_Rechercher()"

    Dim btnAjout As Access.CommandButton
    Set btnAjout = AjouterBouton(frm, "btnAjouter", "+ Nouveau prestataire", _
                                  xDebut + largeurContenu - 2800, yRech, 2600, 480, _
                                  CouleurAccess(52, 73, 94), COULEUR_BLANC)
    btnAjout.OnClick = "=Prestataires_OuvrirFicheNouveau()"

    ' Liste prestataires (cards simulées via ListBox)
    Dim lstPrest As Access.ListBox
    Set lstPrest = CreateControl(frm.Name, acListBox, 0, , , _
                   xDebut + 200, yRech + 680, largeurContenu - 400, 4400)
    lstPrest.Name = "lstPrestataires"
    lstPrest.FontName = POLICE
    lstPrest.FontSize = 9
    lstPrest.BackColor = COULEUR_BLANC
    lstPrest.ForeColor = COULEUR_TEXTE
    lstPrest.BorderStyle = 0
    lstPrest.RowSourceType = "Table/Query"
    lstPrest.RowSource = "SELECT P.id_prestataire, P.nom_entreprise, P.type_activite, " & _
                          "P.nom_responsable, P.telephone, Count(R.id_ressource) AS nb_ressources " & _
                          "FROM PRESTATAIRE AS P LEFT JOIN RESSOURCE AS R ON P.id_prestataire = R.id_prestataire " & _
                          "GROUP BY P.id_prestataire, P.nom_entreprise, P.type_activite, P.nom_responsable, P.telephone " & _
                          "ORDER BY P.nom_entreprise"
    lstPrest.BoundColumn = 1
    lstPrest.ColumnCount = 6
    lstPrest.ColumnWidths = "0;3500;2000;2500;2000;1200"
    lstPrest.OnDblClick = "=Prestataires_OuvrirFicheExistant()"

    ' Boutons action
    Dim yBtns As Long
    yBtns = yRech + 5260

    Dim btnMod As Access.CommandButton
    Set btnMod = AjouterBouton(frm, "btnModifier", "Modifier", _
                                xDebut + 200, yBtns, 1600, 480, CouleurAccess(52, 152, 219), COULEUR_BLANC)
    btnMod.OnClick = "=Prestataires_OuvrirFicheExistant()"

    Dim btnSupp As Access.CommandButton
    Set btnSupp = AjouterBouton(frm, "btnSupprimer", "Supprimer", _
                                 xDebut + 2000, yBtns, 1600, 480, COULEUR_DANGER, COULEUR_BLANC)
    btnSupp.OnClick = "=Prestataires_Supprimer()"

    Dim btnActu As Access.CommandButton
    Set btnActu = AjouterBouton(frm, "btnActualiser", "Actualiser", _
                                  xDebut + 3800, yBtns, 1600, 480, CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnActu.FontBold = False
    btnActu.OnClick = "=Prestataires_Actualiser()"

    ' Fiche prestataire (en dessous)
    Dim yFiche As Long
    yFiche = yBtns + 700

    AjouterRectangle frm, xDebut + 200, yFiche, largeurContenu - 400, 2000, COULEUR_BLANC
    AjouterLabel frm, "lblTitreFiche", "Fiche prestataire", xDebut + 400, yFiche + 120, 5000, 380, 11, COULEUR_TEXTE, True

    Dim largC As Long
    largC = (largeurContenu - 1400) \ 3

    AjouterLabel frm, "lblNomEntr", "Nom entreprise *", xDebut + 400, yFiche + 620, largC, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtNomEntreprise", xDebut + 400, yFiche + 920, largC, 420

    AjouterLabel frm, "lblResponsable", "Responsable", xDebut + 400 + largC + 200, yFiche + 620, largC, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtResponsable", xDebut + 400 + largC + 200, yFiche + 920, largC, 420

    AjouterLabel frm, "lblTypeAct", "Type d'activité", xDebut + 400 + (largC + 200) * 2, yFiche + 620, largC, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtTypeActivite", xDebut + 400 + (largC + 200) * 2, yFiche + 920, largC, 420

    AjouterLabel frm, "lblTel2", "Téléphone", xDebut + 400, yFiche + 1460, largC, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtTelephone", xDebut + 400, yFiche + 1760, largC, 420

    AjouterLabel frm, "lblEmail2", "Email", xDebut + 400 + largC + 200, yFiche + 1460, largC, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtEmail", xDebut + 400 + largC + 200, yFiche + 1760, largC, 420

    Dim txtId2 As Access.TextBox
    Set txtId2 = AjouterZoneTexte(frm, "txtIdPrestataire", 0, 0, 100, 100)
    txtId2.Visible = False

    Dim btnSave2 As Access.CommandButton
    Set btnSave2 = AjouterBouton(frm, "btnEnregistrer", "Enregistrer", _
                                  xDebut + largeurContenu - 4200, yFiche + 1760, 1800, 480, COULEUR_SUCCESS, COULEUR_BLANC)
    btnSave2.OnClick = "=Prestataires_Enregistrer()"

    Dim btnEff2 As Access.CommandButton
    Set btnEff2 = AjouterBouton(frm, "btnEffacer", "Effacer", _
                                  xDebut + largeurContenu - 2200, yFiche + 1760, 1800, 480, CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnEff2.FontBold = False
    btnEff2.OnClick = "=Prestataires_EffacerFiche()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ---- Événements ----
Public Function Prestataires_Actualiser() As Integer
    On Error Resume Next
    Forms(FORM_PRESTATAIRES)!lstPrestataires.Requery
    Prestataires_Actualiser = 0
End Function

Public Function Prestataires_Rechercher() As Integer
    On Error Resume Next
    Dim terme As String
    terme = EchapperChaine(Trim(Nz(Forms(FORM_PRESTATAIRES)!txtRecherche.Value, "")))
    Dim sql As String
    sql = "SELECT P.id_prestataire, P.nom_entreprise, P.type_activite, " & _
          "P.nom_responsable, P.telephone, Count(R.id_ressource) AS nb_ressources " & _
          "FROM PRESTATAIRE AS P LEFT JOIN RESSOURCE AS R ON P.id_prestataire = R.id_prestataire "
    If Len(terme) > 0 Then
        sql = sql & "WHERE P.nom_entreprise LIKE '*" & terme & "*' OR P.type_activite LIKE '*" & terme & "*' "
    End If
    sql = sql & "GROUP BY P.id_prestataire, P.nom_entreprise, P.type_activite, P.nom_responsable, P.telephone ORDER BY P.nom_entreprise"
    Forms(FORM_PRESTATAIRES)!lstPrestataires.RowSource = sql
    Forms(FORM_PRESTATAIRES)!lstPrestataires.Requery
    Prestataires_Rechercher = 0
End Function

Public Function Prestataires_OuvrirFicheNouveau() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_PRESTATAIRES)
    frm!txtIdPrestataire.Value = 0
    frm!txtNomEntreprise.Value = ""
    frm!txtResponsable.Value = ""
    frm!txtTypeActivite.Value = ""
    frm!txtTelephone.Value = ""
    frm!txtEmail.Value = ""
    Prestataires_OuvrirFicheNouveau = 0
End Function

Public Function Prestataires_OuvrirFicheExistant() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_PRESTATAIRES)
    Dim idP As Long
    idP = CLng(Nz(frm!lstPrestataires.Value, 0))
    If idP = 0 Then Exit Function
    Dim rs As DAO.Recordset
    Set rs = OuvrirRecordset("SELECT * FROM PRESTATAIRE WHERE id_prestataire = " & idP)
    If Not rs.EOF Then
        frm!txtIdPrestataire.Value = rs!id_prestataire
        frm!txtNomEntreprise.Value = Nz(rs!nom_entreprise, "")
        frm!txtResponsable.Value = Nz(rs!nom_responsable, "")
        frm!txtTypeActivite.Value = Nz(rs!type_activite, "")
        frm!txtTelephone.Value = Nz(rs!telephone, "")
        frm!txtEmail.Value = Nz(rs!email, "")
    End If
    rs.Close
    Prestataires_OuvrirFicheExistant = 0
End Function

Public Function Prestataires_Enregistrer() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_PRESTATAIRES)
    Dim idP As Long
    idP = CLng(Nz(frm!txtIdPrestataire.Value, 0))
    Dim nomE As String
    nomE = Trim(Nz(frm!txtNomEntreprise.Value, ""))
    If Len(nomE) = 0 Then
        MsgBox "Le nom de l'entreprise est obligatoire.", vbExclamation, APP_NOM
        Prestataires_Enregistrer = 0
        Exit Function
    End If
    If idP = 0 Then
        CreerPrestataire nomE, Trim(Nz(frm!txtResponsable.Value, "")), _
                          Trim(Nz(frm!txtEmail.Value, "")), Trim(Nz(frm!txtTelephone.Value, "")), _
                          "", Trim(Nz(frm!txtTypeActivite.Value, ""))
        AfficherSucces "Prestataire créé."
    Else
        ModifierPrestataire idP, nomE, Trim(Nz(frm!txtResponsable.Value, "")), _
                             Trim(Nz(frm!txtEmail.Value, "")), Trim(Nz(frm!txtTelephone.Value, "")), _
                             "", Trim(Nz(frm!txtTypeActivite.Value, ""))
        AfficherSucces "Prestataire modifié."
    End If
    Prestataires_Actualiser
    Prestataires_Enregistrer = 0
End Function

Public Function Prestataires_Supprimer() As Integer
    On Error Resume Next
    Dim idP As Long
    idP = CLng(Nz(Forms(FORM_PRESTATAIRES)!lstPrestataires.Value, 0))
    If idP > 0 Then
        SupprimerPrestataire idP
        Prestataires_Actualiser
    End If
    Prestataires_Supprimer = 0
End Function

Public Function Prestataires_EffacerFiche() As Integer
    Prestataires_OuvrirFicheNouveau
    Prestataires_EffacerFiche = 0
End Function
