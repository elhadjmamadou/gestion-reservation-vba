Attribute VB_Name = "modFormsDashboard"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Création du formulaire F_TABLEAU_BORD
' Module : modFormsDashboard.bas
' ============================================================

Public Sub Creer_F_TABLEAU_BORD()
    Dim frm As Form
    Dim nomFrm As String
    nomFrm = FORM_DASHBOARD

    SupprimerObjetAccess acForm, nomFrm

    Set frm = CreateForm()
    frm.Caption = "Tableau de bord — " & APP_NOM
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
    AjouterSidebar frm, "Dashboard"

    ' ---- ZONE CONTENU (à droite de la sidebar) ----
    Dim xDebut As Long
    xDebut = LARGEUR_SIDEBAR + 200

    ' Bandeau titre module
    AjouterRectangle frm, xDebut, 0, 16000 - xDebut, 1000, COULEUR_PRIMAIRE
    AjouterLabel frm, "lblTitrePage", "Tableau de bord", _
                 xDebut + 200, 100, 8000, 560, 16, COULEUR_BLANC, True
    AjouterLabel frm, "lblSousTitrePage", "Vue d'ensemble de l'activité", _
                 xDebut + 200, 680, 8000, 300, 9, CouleurAccess(200, 210, 230), False

    ' ---- CARTES STATISTIQUES (4 cartes) ----
    Dim yCartes As Long
    yCartes = 1200
    Dim largeurCarte As Long
    Dim espacement As Long
    Dim largeurZone As Long
    largeurZone = 16000 - xDebut - 400
    espacement = 200
    largeurCarte = (largeurZone - (espacement * 3)) \ 4

    ' Carte 1 : Total réservations
    AjouterCarteStat frm, "TotalRes", "0", "Réservations", _
                     xDebut + 200, yCartes, largeurCarte, 1400, COULEUR_PRIMAIRE

    ' Carte 2 : Aujourd'hui
    AjouterCarteStat frm, "ResJour", "0", "Aujourd'hui", _
                     xDebut + 200 + largeurCarte + espacement, yCartes, _
                     largeurCarte, 1400, COULEUR_SUCCESS

    ' Carte 3 : En attente
    AjouterCarteStat frm, "ResAttente", "0", "En attente", _
                     xDebut + 200 + (largeurCarte + espacement) * 2, yCartes, _
                     largeurCarte, 1400, COULEUR_WARNING

    ' Carte 4 : GNF encaissés
    AjouterCarteStat frm, "Encaisse", "0 GNF", "GNF encaissés", _
                     xDebut + 200 + (largeurCarte + espacement) * 3, yCartes, _
                     largeurCarte, 1400, COULEUR_DANGER

    ' ---- SECTION RÉSERVATIONS PAR CATÉGORIE ----
    Dim ySection As Long
    ySection = yCartes + 1600

    AjouterRectangle frm, xDebut + 200, ySection, 5600, 2600, COULEUR_BLANC
    AjouterLabel frm, "lblTitreCateg", "Réservations par catégorie", _
                 xDebut + 400, ySection + 200, 5200, 380, 11, COULEUR_TEXTE, True

    ' Barres simplifiées (rectangles colorés)
    AjouterRectangle frm, xDebut + 400, ySection + 800, 1400, 1600, COULEUR_PRIMAIRE
    AjouterLabel frm, "lblBarHotel", "Hôtels", xDebut + 400, ySection + 2450, 1400, 280, 8, COULEUR_TEXTE, False

    AjouterRectangle frm, xDebut + 2000, ySection + 1200, 1400, 1200, COULEUR_SUCCESS
    AjouterLabel frm, "lblBarTaxi", "Taxis", xDebut + 2000, ySection + 2450, 1400, 280, 8, COULEUR_TEXTE, False

    AjouterRectangle frm, xDebut + 3600, ySection + 1400, 1400, 1000, COULEUR_WARNING
    AjouterLabel frm, "lblBarVols", "Vols", xDebut + 3600, ySection + 2450, 1400, 280, 8, COULEUR_TEXTE, False

    ' ---- SECTION DERNIÈRES RÉSERVATIONS ----
    AjouterRectangle frm, xDebut + 6000, ySection, 9600, 2600, COULEUR_BLANC
    AjouterLabel frm, "lblTitreDernieres", "Dernières réservations", _
                 xDebut + 6200, ySection + 200, 9200, 380, 11, COULEUR_TEXTE, True

    ' En-têtes colonnes
    AjouterLabel frm, "lblColClient", "Client", xDebut + 6200, ySection + 720, 3600, 280, 8, COULEUR_TEXTE_SECONDAIRE, True
    AjouterLabel frm, "lblColStatut", "Statut", xDebut + 9900, ySection + 720, 1800, 280, 8, COULEUR_TEXTE_SECONDAIRE, True

    ' Lignes de données (5 lignes)
    Dim i As Integer
    For i = 1 To 5
        Dim yLigne As Long
        yLigne = ySection + 1000 + (i - 1) * 360
        AjouterLabel frm, "lblRes" & i, "---", xDebut + 6200, yLigne, 3600, 280, 8, COULEUR_TEXTE, False
        AjouterLabel frm, "lblStat" & i, "---", xDebut + 9900, yLigne, 1800, 280, 8, COULEUR_WARNING, False
    Next i

    ' ---- ACTIONS RAPIDES ----
    Dim yActions As Long
    yActions = ySection + 2900

    AjouterLabel frm, "lblTitreActions", "Actions rapides", _
                 xDebut + 200, yActions, 12000, 380, 11, COULEUR_TEXTE, True

    Dim yBtns As Long
    yBtns = yActions + 560

    ' Bouton Nouvelle réservation
    Dim btnNouvelleRes As Access.CommandButton
    Set btnNouvelleRes = AjouterBouton(frm, "btnNouvelleReservation", "+ Nouvelle réservation", _
                                        xDebut + 200, yBtns, 2800, 600, COULEUR_PRIMAIRE, COULEUR_BLANC)
    btnNouvelleRes.OnClick = "=Navigation_NouvelleReservation()"

    ' Bouton Nouveau client
    Dim btnNouveauCli As Access.CommandButton
    Set btnNouveauCli = AjouterBouton(frm, "btnNouveauClient", "+ Nouveau client", _
                                       xDebut + 3200, yBtns, 2800, 600, COULEUR_SUCCESS, COULEUR_BLANC)
    btnNouveauCli.OnClick = "=Navigation_Clients()"

    ' Bouton Enregistrer paiement
    Dim btnPaiement As Access.CommandButton
    Set btnPaiement = AjouterBouton(frm, "btnEnregistrerPaiement", "$ Enregistrer paiement", _
                                     xDebut + 6200, yBtns, 2800, 600, COULEUR_WARNING, COULEUR_BLANC)
    btnPaiement.OnClick = "=Navigation_Paiements()"

    ' Bouton Voir rapports
    Dim btnRapports As Access.CommandButton
    Set btnRapports = AjouterBouton(frm, "btnVoirRapports", Chr(9632) & " Voir rapports", _
                                     xDebut + 9200, yBtns, 2800, 600, _
                                     CouleurAccess(100, 100, 130), COULEUR_BLANC)
    btnRapports.OnClick = "=Navigation_Rapports()"

    ' ---- Bouton Rafraîchir ----
    Dim btnRafraichir As Access.CommandButton
    Set btnRafraichir = AjouterBouton(frm, "btnRafraichir", "Actualiser", _
                                       xDebut + 200, yBtns + 800, 1800, 480, _
                                       CouleurAccess(189, 195, 199), COULEUR_TEXTE)
    btnRafraichir.FontSize = 9
    btnRafraichir.FontBold = False
    btnRafraichir.OnClick = "=Dashboard_Rafraichir()"

    ' Sauvegarder
    DoCmd.Save acForm, nomFrm
    DoCmd.Close acForm, nomFrm, acSaveYes

    Debug.Print "Formulaire " & nomFrm & " créé avec succès."
End Sub

' ============================================================
' Rafraîchir les statistiques du tableau de bord
' ============================================================
Public Function Dashboard_Rafraichir() As Integer
    On Error Resume Next

    Dim frm As Form
    Set frm = Forms(FORM_DASHBOARD)

    ' Mettre à jour les cartes statistiques
    frm!lblTotalRes_Val.Caption = CStr(NbTotalReservations())
    frm!lblResJour_Val.Caption = CStr(NbReservationsDuJour())
    frm!lblResAttente_Val.Caption = CStr(NbReservationsEnAttente())
    frm!lblEncaisse_Val.Caption = FormatMontant(MontantEncaisseTotal())

    ' Mettre à jour les dernières réservations
    Dim rs As DAO.Recordset
    Set rs = DernieresReservations(5)

    Dim i As Integer
    i = 1
    Do While Not rs.EOF And i <= 5
        frm.Controls("lblRes" & i).Caption = Tronquer(ChaineSure(rs!client_nom), 30)
        frm.Controls("lblStat" & i).Caption = ChaineSure(rs!statut)
        frm.Controls("lblStat" & i).ForeColor = CouleurStatutReservation(ChaineSure(rs!statut))
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close

    Dashboard_Rafraichir = 0
End Function

' ============================================================
' Événement d'ouverture du tableau de bord
' ============================================================
Public Function Dashboard_OnOpen() As Integer
    Dashboard_Rafraichir
    Dashboard_OnOpen = 0
End Function
