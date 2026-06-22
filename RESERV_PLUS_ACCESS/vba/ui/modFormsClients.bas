Attribute VB_Name = "modFormsClients"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Création du formulaire F_CLIENTS
' Module : modFormsClients.bas
' ============================================================

Public Sub Creer_F_CLIENTS()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_CLIENTS

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

    ' ---- SIDEBAR ----
    AjouterSidebar frm, "Clients"

    ' ---- ZONE CONTENU ----
    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200
    Dim largeurContenu As Long
    largeurContenu = 16000 - xDebut - 200

    ' Bandeau titre
    AjouterRectangle frm, xDebut, 0, largeurContenu, 1000, CouleurAccess(192, 57, 43)
    AjouterLabel frm, "lblTitrePage", "Clients", xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitrePage", "Fiches clients et historique des réservations", _
                 xDebut + 200, 680, 8000, 300, 9, CouleurAccess(230, 200, 200), False

    ' ---- BARRE DE RECHERCHE ET BOUTON AJOUTER ----
    Dim yRecherche As Long
    yRecherche = 1200

    AjouterLabel frm, "lblRecherche", Chr(9829) & " Rechercher :", _
                 xDebut + 200, yRecherche + 40, 1200, 320, 9, COULEUR_TEXTE_SECONDAIRE, False

    Dim txtRecherche As Access.TextBox
    Set txtRecherche = AjouterZoneTexte(frm, "txtRecherche", xDebut + 1400, yRecherche, 4400, 480)
    txtRecherche.OnChange = "=Clients_Rechercher()"

    Dim btnRechercher As Access.CommandButton
    Set btnRechercher = AjouterBouton(frm, "btnRechercher", "Rechercher", _
                                       xDebut + 5900, yRecherche, 1800, 480, _
                                       CouleurAccess(100, 100, 130), COULEUR_BLANC)
    btnRechercher.FontSize = 9
    btnRechercher.FontBold = False
    btnRechercher.OnClick = "=Clients_Rechercher()"

    Dim btnAjouter As Access.CommandButton
    Set btnAjouter = AjouterBouton(frm, "btnAjouter", "+ Nouveau client", _
                                    xDebut + largeurContenu - 2600, yRecherche, 2400, 480, _
                                    CouleurAccess(192, 57, 43), COULEUR_BLANC)
    btnAjouter.OnClick = "=Clients_OuvrirFicheNouveau()"

    ' ---- EN-TÊTES COLONNES LISTE ----
    Dim yHeader As Long
    yHeader = yRecherche + 700
    AjouterRectangle frm, xDebut + 200, yHeader, largeurContenu - 400, 480, _
                     CouleurAccess(236, 240, 241)

    Dim cols() As String
    Dim largsCols() As Long
    Dim xCol As Long
    cols = Split("Nom|Prénom|Téléphone|Email|Réserv.|Inscrit le", "|")
    largsCols = Array(2000, 2000, 2000, 3000, 1200, 1800)
    xCol = xDebut + 300

    Dim c As Integer
    For c = 0 To UBound(cols)
        AjouterLabel frm, "lblCol" & c, cols(c), xCol, yHeader + 80, largsCols(c), 320, _
                     8, COULEUR_TEXTE, True
        xCol = xCol + largsCols(c) + 100
    Next c

    ' ---- ZONE LISTE (ListBox liée à une requête) ----
    Dim lstClients As Access.ListBox
    Set lstClients = CreateControl(frm.Name, acListBox, 0, , , _
                     xDebut + 200, yHeader + 480, largeurContenu - 400, 5200)
    lstClients.Name = "lstClients"
    lstClients.FontName = POLICE
    lstClients.FontSize = 9
    lstClients.BackColor = COULEUR_BLANC
    lstClients.ForeColor = COULEUR_TEXTE
    lstClients.BorderStyle = 0
    lstClients.RowSourceType = "Table/Query"
    lstClients.RowSource = RQ_CLIENTS_NB_RESERVATIONS
    lstClients.BoundColumn = 1
    lstClients.ColumnCount = 7
    lstClients.ColumnWidths = "0;2000;2000;2000;3000;1200;1800"
    lstClients.OnDblClick = "=Clients_OuvrirFicheExistant()"

    ' ---- BOUTONS D'ACTION ----
    Dim yBtns As Long
    yBtns = yHeader + 5850

    Dim btnModifier As Access.CommandButton
    Set btnModifier = AjouterBouton(frm, "btnModifier", "Modifier", _
                                     xDebut + 200, yBtns, 1800, 480, _
                                     CouleurAccess(52, 152, 219), COULEUR_BLANC)
    btnModifier.OnClick = "=Clients_OuvrirFicheExistant()"

    Dim btnSupprimer As Access.CommandButton
    Set btnSupprimer = AjouterBouton(frm, "btnSupprimer", "Supprimer", _
                                      xDebut + 2200, yBtns, 1800, 480, COULEUR_DANGER, COULEUR_BLANC)
    btnSupprimer.OnClick = "=Clients_Supprimer()"

    Dim btnActualiser As Access.CommandButton
    Set btnActualiser = AjouterBouton(frm, "btnActualiser", "Actualiser", _
                                       xDebut + 4200, yBtns, 1800, 480, _
                                       CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnActualiser.FontBold = False
    btnActualiser.OnClick = "=Clients_Actualiser()"

    ' Compteur
    AjouterLabel frm, "lblCompteur", "0 client(s)", _
                 xDebut + largeurContenu - 2200, yBtns + 80, 2000, 320, _
                 8, COULEUR_TEXTE_SECONDAIRE, False

    ' ---- FORMULAIRE FICHE CLIENT (zone de saisie) ----
    Dim yFiche As Long
    yFiche = yBtns + 700

    AjouterRectangle frm, xDebut + 200, yFiche, largeurContenu - 400, 2200, COULEUR_BLANC
    AjouterLabel frm, "lblTitreFiche", "Fiche client", _
                 xDebut + 400, yFiche + 120, 3000, 380, 11, COULEUR_TEXTE, True

    ' Champs en grille 3 colonnes
    Dim largChamp As Long
    largChamp = (largeurContenu - 800) \ 3 - 200

    ' Ligne 1 : Nom | Prénom
    AjouterLabel frm, "lblNom", "Nom *", xDebut + 400, yFiche + 620, largChamp, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtNom", xDebut + 400, yFiche + 920, largChamp, 420
    AjouterLabel frm, "lblPrenom", "Prénom", xDebut + 400 + largChamp + 200, yFiche + 620, largChamp, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtPrenom", xDebut + 400 + largChamp + 200, yFiche + 920, largChamp, 420

    ' Ligne 2 : Email | Téléphone
    AjouterLabel frm, "lblEmail", "Email", xDebut + 400, yFiche + 1480, largChamp, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtEmail", xDebut + 400, yFiche + 1780, largChamp, 420
    AjouterLabel frm, "lblTel", "Téléphone", xDebut + 400 + largChamp + 200, yFiche + 1480, largChamp, 280, 8, COULEUR_TEXTE, False
    AjouterZoneTexte frm, "txtTelephone", xDebut + 400 + largChamp + 200, yFiche + 1780, largChamp, 420

    ' Champ caché ID
    Dim txtID As Access.TextBox
    Set txtID = AjouterZoneTexte(frm, "txtIdClient", 0, 0, 100, 100)
    txtID.Visible = False

    ' Boutons sauvegarder / annuler fiche
    Dim btnSave As Access.CommandButton
    Set btnSave = AjouterBouton(frm, "btnEnregistrer", "Enregistrer", _
                                 xDebut + largeurContenu - 4200, yFiche + 1780, 1800, 480, _
                                 COULEUR_SUCCESS, COULEUR_BLANC)
    btnSave.OnClick = "=Clients_Enregistrer()"

    Dim btnAnnFiche As Access.CommandButton
    Set btnAnnFiche = AjouterBouton(frm, "btnAnnulerFiche", "Effacer", _
                                     xDebut + largeurContenu - 2200, yFiche + 1780, 1800, 480, _
                                     CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnAnnFiche.FontBold = False
    btnAnnFiche.OnClick = "=Clients_EffacerFiche()"

    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes
    Debug.Print "Formulaire " & nomFrm & " créé."
End Sub

' ============================================================
' Événements du formulaire clients
' ============================================================
Public Function Clients_Rechercher() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_CLIENTS)

    Dim terme As String
    terme = Trim(Nz(frm!txtRecherche.Value, ""))

    Dim sql As String
    If Len(terme) > 0 Then
        sql = "SELECT C.id_client, C.nom, C.prenom, C.email, C.telephone, " & _
              "C.date_inscription, Count(R.id_reservation) AS nb_reservations " & _
              "FROM CLIENT AS C LEFT JOIN RESERVATION AS R ON C.id_client = R.id_client " & _
              "WHERE C.nom LIKE '*" & EchapperChaine(terme) & "*' " & _
              "OR C.prenom LIKE '*" & EchapperChaine(terme) & "*' " & _
              "OR C.email LIKE '*" & EchapperChaine(terme) & "*' " & _
              "OR C.telephone LIKE '*" & EchapperChaine(terme) & "*' " & _
              "GROUP BY C.id_client, C.nom, C.prenom, C.email, C.telephone, C.date_inscription " & _
              "ORDER BY C.nom"
    Else
        sql = "SELECT C.id_client, C.nom, C.prenom, C.email, C.telephone, " & _
              "C.date_inscription, Count(R.id_reservation) AS nb_reservations " & _
              "FROM CLIENT AS C LEFT JOIN RESERVATION AS R ON C.id_client = R.id_client " & _
              "GROUP BY C.id_client, C.nom, C.prenom, C.email, C.telephone, C.date_inscription " & _
              "ORDER BY C.nom"
    End If

    frm!lstClients.RowSource = sql
    frm!lstClients.Requery

    Dim nb As Long
    nb = CompterEnregistrements("CLIENT", IIf(Len(terme) > 0, _
         "nom LIKE '*" & EchapperChaine(terme) & "*' OR prenom LIKE '*" & EchapperChaine(terme) & "*'", ""))
    frm!lblCompteur.Caption = nb & " client(s)"
    Clients_Rechercher = 0
End Function

Public Function Clients_Actualiser() As Integer
    On Error Resume Next
    Forms(FORM_CLIENTS)!lstClients.RowSource = RQ_CLIENTS_NB_RESERVATIONS
    Forms(FORM_CLIENTS)!lstClients.Requery
    Forms(FORM_CLIENTS)!txtRecherche.Value = ""
    Dim nb As Long
    nb = CompterEnregistrements("CLIENT", "")
    Forms(FORM_CLIENTS)!lblCompteur.Caption = nb & " client(s)"
    Clients_Actualiser = 0
End Function

Public Function Clients_OuvrirFicheNouveau() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_CLIENTS)
    frm!txtIdClient.Value = 0
    frm!txtNom.Value = ""
    frm!txtPrenom.Value = ""
    frm!txtEmail.Value = ""
    frm!txtTelephone.Value = ""
    frm!txtNom.SetFocus
    Clients_OuvrirFicheNouveau = 0
End Function

Public Function Clients_OuvrirFicheExistant() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_CLIENTS)

    Dim idClient As Long
    idClient = CLng(Nz(frm!lstClients.Value, 0))
    If idClient = 0 Then
        MsgBox "Sélectionnez un client dans la liste.", vbExclamation, APP_NOM
        Clients_OuvrirFicheExistant = 0
        Exit Function
    End If

    Dim rs As DAO.Recordset
    Set rs = OuvrirRecordset("SELECT * FROM CLIENT WHERE id_client = " & idClient)
    If Not rs.EOF Then
        frm!txtIdClient.Value = rs!id_client
        frm!txtNom.Value = Nz(rs!nom, "")
        frm!txtPrenom.Value = Nz(rs!prenom, "")
        frm!txtEmail.Value = Nz(rs!email, "")
        frm!txtTelephone.Value = Nz(rs!telephone, "")
    End If
    rs.Close
    Clients_OuvrirFicheExistant = 0
End Function

Public Function Clients_Enregistrer() As Integer
    On Error Resume Next
    Dim frm As Form
    Set frm = Forms(FORM_CLIENTS)

    Dim idClient As Long
    idClient = CLng(Nz(frm!txtIdClient.Value, 0))
    Dim nom As String
    nom = Trim(Nz(frm!txtNom.Value, ""))

    If Len(nom) = 0 Then
        MsgBox "Le nom est obligatoire.", vbExclamation, APP_NOM
        Clients_Enregistrer = 0
        Exit Function
    End If

    If idClient = 0 Then
        CreerClient nom, Trim(Nz(frm!txtPrenom.Value, "")), _
                    Trim(Nz(frm!txtEmail.Value, "")), _
                    Trim(Nz(frm!txtTelephone.Value, "")), ""
        AfficherSucces "Client créé avec succès."
    Else
        ModifierClient idClient, nom, Trim(Nz(frm!txtPrenom.Value, "")), _
                       Trim(Nz(frm!txtEmail.Value, "")), _
                       Trim(Nz(frm!txtTelephone.Value, "")), ""
        AfficherSucces "Client modifié."
    End If

    Clients_Actualiser
    Clients_Enregistrer = 0
End Function

Public Function Clients_Supprimer() As Integer
    On Error Resume Next
    Dim idClient As Long
    idClient = CLng(Nz(Forms(FORM_CLIENTS)!lstClients.Value, 0))
    If idClient > 0 Then
        SupprimerClient idClient
        Clients_Actualiser
    End If
    Clients_Supprimer = 0
End Function

Public Function Clients_EffacerFiche() As Integer
    On Error Resume Next
    Clients_OuvrirFicheNouveau
    Clients_EffacerFiche = 0
End Function
