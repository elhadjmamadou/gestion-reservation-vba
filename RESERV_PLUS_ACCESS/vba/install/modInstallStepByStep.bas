Attribute VB_Name = "modInstallStepByStep"
Option Compare Database
Option Explicit

' =============================================================================
' RESERV+ - Installation pas a pas
' Tapez chaque commande dans la fenetre Execution (CTRL+G) et appuyez sur ENTREE
' =============================================================================

' -----------------------------------------------------------------------------
' ETAPE 1a - Table PRESTATAIRE
' -----------------------------------------------------------------------------
Public Sub Etape1a_PRESTATAIRE()
    If TableExiste("PRESTATAIRE") Then
        MsgBox "PRESTATAIRE existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE PRESTATAIRE (" & _
        "id_prestataire AUTOINCREMENT PRIMARY KEY, " & _
        "nom_entreprise TEXT(150) NOT NULL, " & _
        "nom_responsable TEXT(120), " & _
        "email TEXT(120), " & _
        "telephone TEXT(30), " & _
        "adresse MEMO, " & _
        "type_activite TEXT(80), " & _
        "date_adhesion DATETIME)"
    MsgBox "PRESTATAIRE creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1b - Table CATEGORIE
' -----------------------------------------------------------------------------
Public Sub Etape1b_CATEGORIE()
    If TableExiste("CATEGORIE") Then
        MsgBox "CATEGORIE existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE CATEGORIE (" & _
        "id_categorie AUTOINCREMENT PRIMARY KEY, " & _
        "libelle_categorie TEXT(80) NOT NULL, " & _
        "description MEMO)"
    MsgBox "CATEGORIE creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1c - Table RESSOURCE
' -----------------------------------------------------------------------------
Public Sub Etape1c_RESSOURCE()
    If TableExiste("RESSOURCE") Then
        MsgBox "RESSOURCE existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE RESSOURCE (" & _
        "id_ressource AUTOINCREMENT PRIMARY KEY, " & _
        "id_prestataire LONG NOT NULL, " & _
        "id_categorie LONG NOT NULL, " & _
        "libelle TEXT(150) NOT NULL, " & _
        "description MEMO, " & _
        "prix_unitaire CURRENCY, " & _
        "capacite INTEGER, " & _
        "disponible YESNO, " & _
        "statut_ressource TEXT(50), " & _
        "date_creation DATETIME)"
    MsgBox "RESSOURCE creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1d - Table CHAMBRE_HOTEL
' -----------------------------------------------------------------------------
Public Sub Etape1d_CHAMBRE_HOTEL()
    If TableExiste("CHAMBRE_HOTEL") Then
        MsgBox "CHAMBRE_HOTEL existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE CHAMBRE_HOTEL (" & _
        "id_ressource LONG PRIMARY KEY, " & _
        "nb_etoiles INTEGER, " & _
        "type_chambre TEXT(80), " & _
        "nb_lits INTEGER, " & _
        "petit_dejeuner YESNO, " & _
        "vue TEXT(80))"
    MsgBox "CHAMBRE_HOTEL creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1e - Table TAXI
' -----------------------------------------------------------------------------
Public Sub Etape1e_TAXI()
    If TableExiste("TAXI") Then
        MsgBox "TAXI existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE TAXI (" & _
        "id_ressource LONG PRIMARY KEY, " & _
        "marque TEXT(80), " & _
        "modele TEXT(80), " & _
        "immatriculation TEXT(50), " & _
        "nb_places INTEGER, " & _
        "climatisation YESNO)"
    MsgBox "TAXI creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1f - Table VOL
' -----------------------------------------------------------------------------
Public Sub Etape1f_VOL()
    If TableExiste("VOL") Then
        MsgBox "VOL existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE VOL (" & _
        "id_ressource LONG PRIMARY KEY, " & _
        "num_vol TEXT(50), " & _
        "ville_depart TEXT(100), " & _
        "ville_arrivee TEXT(100), " & _
        "date_depart DATETIME, " & _
        "heure_depart DATETIME, " & _
        "date_arrivee DATETIME, " & _
        "heure_arrivee DATETIME, " & _
        "classe TEXT(80))"
    MsgBox "VOL creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1g - Table CLIENT
' -----------------------------------------------------------------------------
Public Sub Etape1g_CLIENT()
    If TableExiste("CLIENT") Then
        MsgBox "CLIENT existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE CLIENT (" & _
        "id_client AUTOINCREMENT PRIMARY KEY, " & _
        "nom TEXT(100) NOT NULL, " & _
        "prenom TEXT(100), " & _
        "email TEXT(120), " & _
        "telephone TEXT(30), " & _
        "adresse MEMO, " & _
        "date_inscription DATETIME)"
    MsgBox "CLIENT creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1h - Table UTILISATEUR
' -----------------------------------------------------------------------------
Public Sub Etape1h_UTILISATEUR()
    If TableExiste("UTILISATEUR") Then
        MsgBox "UTILISATEUR existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE UTILISATEUR (" & _
        "id_utilisateur AUTOINCREMENT PRIMARY KEY, " & _
        "login TEXT(80) NOT NULL, " & _
        "mot_de_passe TEXT(255), " & _
        "nom_complet TEXT(150), " & _
        "role TEXT(50), " & _
        "actif YESNO, " & _
        "date_creation DATETIME)"
    MsgBox "UTILISATEUR creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1i - Table RESERVATION
' -----------------------------------------------------------------------------
Public Sub Etape1i_RESERVATION()
    If TableExiste("RESERVATION") Then
        MsgBox "RESERVATION existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE RESERVATION (" & _
        "id_reservation AUTOINCREMENT PRIMARY KEY, " & _
        "id_client LONG NOT NULL, " & _
        "id_utilisateur LONG, " & _
        "date_debut DATETIME, " & _
        "date_fin DATETIME, " & _
        "nb_personnes INTEGER, " & _
        "montant_total CURRENCY, " & _
        "statut TEXT(50), " & _
        "date_creation DATETIME, " & _
        "observation MEMO)"
    MsgBox "RESERVATION creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1j - Table DETAIL_RESERVATION
' -----------------------------------------------------------------------------
Public Sub Etape1j_DETAIL_RESERVATION()
    If TableExiste("DETAIL_RESERVATION") Then
        MsgBox "DETAIL_RESERVATION existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE DETAIL_RESERVATION (" & _
        "id_detail AUTOINCREMENT PRIMARY KEY, " & _
        "id_reservation LONG NOT NULL, " & _
        "id_ressource LONG NOT NULL, " & _
        "quantite INTEGER, " & _
        "prix_unitaire_applique CURRENCY, " & _
        "sous_total CURRENCY)"
    MsgBox "DETAIL_RESERVATION creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1k - Table PAIEMENT
' -----------------------------------------------------------------------------
Public Sub Etape1k_PAIEMENT()
    If TableExiste("PAIEMENT") Then
        MsgBox "PAIEMENT existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE PAIEMENT (" & _
        "id_paiement AUTOINCREMENT PRIMARY KEY, " & _
        "id_reservation LONG NOT NULL, " & _
        "date_paiement DATETIME, " & _
        "montant CURRENCY, " & _
        "mode_paiement TEXT(50), " & _
        "reference_paiement TEXT(100), " & _
        "commentaire MEMO)"
    MsgBox "PAIEMENT creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1l - Table BILLET
' -----------------------------------------------------------------------------
Public Sub Etape1l_BILLET()
    If TableExiste("BILLET") Then
        MsgBox "BILLET existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE BILLET (" & _
        "id_billet AUTOINCREMENT PRIMARY KEY, " & _
        "id_reservation LONG NOT NULL, " & _
        "numero_billet TEXT(80) NOT NULL, " & _
        "date_emission DATETIME, " & _
        "statut_billet TEXT(50))"
    MsgBox "BILLET creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1m - Table JOURNAL_ACTION
' -----------------------------------------------------------------------------
Public Sub Etape1m_JOURNAL_ACTION()
    If TableExiste("JOURNAL_ACTION") Then
        MsgBox "JOURNAL_ACTION existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE JOURNAL_ACTION (" & _
        "id_action AUTOINCREMENT PRIMARY KEY, " & _
        "id_utilisateur LONG, " & _
        "action TEXT(150), " & _
        "details MEMO, " & _
        "date_action DATETIME)"
    MsgBox "JOURNAL_ACTION creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 1n - Table PARAMETRE
' -----------------------------------------------------------------------------
Public Sub Etape1n_PARAMETRE()
    If TableExiste("PARAMETRE") Then
        MsgBox "PARAMETRE existe deja - OK", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "CREATE TABLE PARAMETRE (" & _
        "id_parametre AUTOINCREMENT PRIMARY KEY, " & _
        "cle_parametre TEXT(100), " & _
        "valeur_parametre TEXT(255), " & _
        "description MEMO)"
    MsgBox "PARAMETRE creee", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 2 - Index
' -----------------------------------------------------------------------------
Public Sub Etape2_Index()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim sqls(15) As String
    sqls(0)  = "CREATE INDEX IDX_RES_PRESTATAIRE ON RESSOURCE (id_prestataire)"
    sqls(1)  = "CREATE INDEX IDX_RES_CATEGORIE ON RESSOURCE (id_categorie)"
    sqls(2)  = "CREATE INDEX IDX_RES_DISPONIBLE ON RESSOURCE (disponible)"
    sqls(3)  = "CREATE INDEX IDX_CLI_NOM ON CLIENT (nom, prenom)"
    sqls(4)  = "CREATE UNIQUE INDEX IDX_USR_LOGIN ON UTILISATEUR (login)"
    sqls(5)  = "CREATE INDEX IDX_RESV_CLIENT ON RESERVATION (id_client)"
    sqls(6)  = "CREATE INDEX IDX_RESV_STATUT ON RESERVATION (statut)"
    sqls(7)  = "CREATE INDEX IDX_RESV_DATES ON RESERVATION (date_debut, date_fin)"
    sqls(8)  = "CREATE INDEX IDX_DET_RESERVATION ON DETAIL_RESERVATION (id_reservation)"
    sqls(9)  = "CREATE INDEX IDX_DET_RESSOURCE ON DETAIL_RESERVATION (id_ressource)"
    sqls(10) = "CREATE INDEX IDX_PAI_RESERVATION ON PAIEMENT (id_reservation)"
    sqls(11) = "CREATE INDEX IDX_PAI_DATE ON PAIEMENT (date_paiement)"
    sqls(12) = "CREATE UNIQUE INDEX IDX_BIL_NUMERO ON BILLET (numero_billet)"
    sqls(13) = "CREATE INDEX IDX_BIL_RESERVATION ON BILLET (id_reservation)"
    sqls(14) = "CREATE INDEX IDX_JRN_DATE ON JOURNAL_ACTION (date_action)"
    sqls(15) = "CREATE UNIQUE INDEX IDX_PAR_CLE ON PARAMETRE (cle_parametre)"

    Dim i As Integer
    Dim ok As Integer : ok = 0
    For i = 0 To 15
        On Error Resume Next
        db.Execute sqls(i)
        If Err.Number = 0 Then ok = ok + 1
        On Error GoTo 0
    Next i

    MsgBox "Index : " & ok & "/16 crees (les autres existaient deja)", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 3a - Donnees : Categories
' -----------------------------------------------------------------------------
Public Sub Etape3a_Categories()
    If CompterEnregistrements("CATEGORIE", "") > 0 Then
        MsgBox "Categories deja presentes (" & CompterEnregistrements("CATEGORIE", "") & ")", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Hotel', 'Chambres et hebergements hoteliers')"
    CurrentDb.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Taxi', 'Vehicules de transport terrestre')"
    CurrentDb.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Vol', 'Vols et transport aerien')"
    MsgBox "3 categories inserees", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 3b - Donnees : Utilisateurs
' -----------------------------------------------------------------------------
Public Sub Etape3b_Utilisateurs()
    If CompterEnregistrements("UTILISATEUR", "") > 0 Then
        MsgBox "Utilisateurs deja presents (" & CompterEnregistrements("UTILISATEUR", "") & ")", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) VALUES ('admin', 'admin123', 'Administrateur', 'Administrateur', True, Now())"
    CurrentDb.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) VALUES ('agent', 'agent123', 'Agent de Reservation', 'Agent', True, Now())"
    CurrentDb.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) VALUES ('utilisateur', 'user123', 'Utilisateur Standard', 'Utilisateur', True, Now())"
    MsgBox "3 utilisateurs inseres (admin/agent/utilisateur)", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 3c - Donnees : Parametres
' -----------------------------------------------------------------------------
Public Sub Etape3c_Parametres()
    If CompterEnregistrements("PARAMETRE", "") > 0 Then
        MsgBox "Parametres deja presents (" & CompterEnregistrements("PARAMETRE", "") & ")", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('NOM_APPLICATION', 'RESERV+', 'Nom application')"
    CurrentDb.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('DEVISE', 'GNF', 'Devise')"
    CurrentDb.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('PREFIXE_BILLET', 'BIL', 'Prefixe billet')"
    CurrentDb.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('NOM_ETABLISSEMENT', 'RESERV+', 'Nom etablissement')"
    CurrentDb.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('VERSION_APP', '1.0.0', 'Version')"
    MsgBox "5 parametres inseres", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 3d - Donnees : Prestataires
' -----------------------------------------------------------------------------
Public Sub Etape3d_Prestataires()
    If CompterEnregistrements("PRESTATAIRE", "") > 0 Then
        MsgBox "Prestataires deja presents (" & CompterEnregistrements("PRESTATAIRE", "") & ")", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion) VALUES ('Hotel Kaloum SARL', 'M. Sylla', 'contact@hotelkaloum.gn', '620 11 22 33', 'Avenue de la Republique, Kaloum, Conakry', 'Hotellerie', Now())"
    CurrentDb.Execute "INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion) VALUES ('Express Taxi Conakry', 'Mme Barry', 'info@expresstaxi.gn', '622 33 44 55', 'Quartier Madina, Conakry', 'Transport terrestre', Now())"
    CurrentDb.Execute "INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion) VALUES ('Air Guinee Express', 'M. Camara', 'reservations@airgn.gn', '628 55 66 77', 'Aeroport International Gbessia, Conakry', 'Aerien', Now())"
    CurrentDb.Execute "INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, adresse, type_activite, date_adhesion) VALUES ('Niger Hotels Group', 'M. Diakite', 'contact@nigerhotels.gn', '664 77 88 99', 'Avenue du Niger, Conakry', 'Hotellerie', Now())"
    MsgBox "4 prestataires inseres", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 3e - Donnees : Ressources + details
' -----------------------------------------------------------------------------
Public Sub Etape3e_Ressources()
    If CompterEnregistrements("RESSOURCE", "") > 0 Then
        MsgBox "Ressources deja presentes (" & CompterEnregistrements("RESSOURCE", "") & ")", vbInformation
        Exit Sub
    End If

    Dim idHotel As Long, idTaxi As Long, idVol As Long
    idHotel = CLng(Nz(ObtenirValeur("SELECT id_categorie FROM CATEGORIE WHERE libelle_categorie='Hotel'"), 1))
    idTaxi  = CLng(Nz(ObtenirValeur("SELECT id_categorie FROM CATEGORIE WHERE libelle_categorie='Taxi'"), 2))
    idVol   = CLng(Nz(ObtenirValeur("SELECT id_categorie FROM CATEGORIE WHERE libelle_categorie='Vol'"), 3))

    ' Ressources
    CurrentDb.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation) VALUES (1," & idHotel & ",'Hotel Kaloum - Chambre 204','Chambre double vue mer, climatisee',150000,2,True,'Disponible',Now())"
    CurrentDb.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation) VALUES (4," & idHotel & ",'Hotel Niger - Chambre 11','Chambre standard confortable',100000,2,True,'Disponible',Now())"
    CurrentDb.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation) VALUES (2," & idTaxi & ",'Taxi Express 01','Berline climatisee 4 places',30000,4,True,'Disponible',Now())"
    CurrentDb.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation) VALUES (2," & idTaxi & ",'Taxi VIP 07','SUV climatise 4 places',60000,4,True,'Disponible',Now())"
    CurrentDb.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation) VALUES (3," & idVol & ",'Vol CKY-LBE','Conakry - Labe, classe economique',620000,120,True,'Disponible',Now())"
    CurrentDb.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, description, prix_unitaire, capacite, disponible, statut_ressource, date_creation) VALUES (3," & idVol & ",'Vol CKY-NZE','Conakry - N''Zerekofe, classe economique',580000,120,True,'Disponible',Now())"

    ' Details hotel
    CurrentDb.Execute "INSERT INTO CHAMBRE_HOTEL (id_ressource, nb_etoiles, type_chambre, nb_lits, petit_dejeuner, vue) VALUES (1,4,'Double',2,True,'Vue mer')"
    CurrentDb.Execute "INSERT INTO CHAMBRE_HOTEL (id_ressource, nb_etoiles, type_chambre, nb_lits, petit_dejeuner, vue) VALUES (2,3,'Standard',1,False,'Vue jardin')"

    ' Details taxi
    CurrentDb.Execute "INSERT INTO TAXI (id_ressource, marque, modele, immatriculation, nb_places, climatisation) VALUES (3,'Toyota','Corolla','RC-1234-CKY',4,True)"
    CurrentDb.Execute "INSERT INTO TAXI (id_ressource, marque, modele, immatriculation, nb_places, climatisation) VALUES (4,'Toyota','Land Cruiser','RC-5678-CKY',4,True)"

    ' Details vol
    CurrentDb.Execute "INSERT INTO VOL (id_ressource, num_vol, ville_depart, ville_arrivee, date_depart, date_arrivee, classe) VALUES (5,'AGX-001','Conakry','Labe',#2026/06/25#,#2026/06/25#,'Economique')"
    CurrentDb.Execute "INSERT INTO VOL (id_ressource, num_vol, ville_depart, ville_arrivee, date_depart, date_arrivee, classe) VALUES (6,'AGX-002','Conakry','N-Zerekore',#2026/06/26#,#2026/06/26#,'Economique')"

    MsgBox "6 ressources + details inseres", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 3f - Donnees : Clients
' -----------------------------------------------------------------------------
Public Sub Etape3f_Clients()
    If CompterEnregistrements("CLIENT", "") > 0 Then
        MsgBox "Clients deja presents (" & CompterEnregistrements("CLIENT", "") & ")", vbInformation
        Exit Sub
    End If
    CurrentDb.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription) VALUES ('Diallo','Mamadou','m.diallo@mail.gn','620 00 00 01','Kaloum, Conakry',#2026/01/02#)"
    CurrentDb.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription) VALUES ('Camara','Aissatou','a.camara@mail.gn','622 11 22 33','Ratoma, Conakry',#2026/02/15#)"
    CurrentDb.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription) VALUES ('Bah','Saikou','s.bah@mail.gn','628 44 55 66','Matoto, Conakry',#2026/03/20#)"
    CurrentDb.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription) VALUES ('Toure','Fatou','f.toure@mail.gn','655 77 88 99','Dixinn, Conakry',#2026/04/10#)"
    CurrentDb.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, adresse, date_inscription) VALUES ('Soumah','Karim','k.soumah@mail.gn','664 12 34 56','Matam, Conakry',#2026/05/05#)"
    MsgBox "5 clients inseres", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 4 - Requetes (toutes les 12)
' -----------------------------------------------------------------------------
Public Sub Etape4_Requetes()
    CreerToutesLesRequetes
    MsgBox "12 requetes creees", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5a - Formulaire F_CONNEXION
' -----------------------------------------------------------------------------
Public Sub Etape5a_F_CONNEXION()
    Creer_F_CONNEXION
    MsgBox "F_CONNEXION cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5b - Formulaire F_TABLEAU_BORD
' -----------------------------------------------------------------------------
Public Sub Etape5b_F_TABLEAU_BORD()
    Creer_F_TABLEAU_BORD
    MsgBox "F_TABLEAU_BORD cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5c - Formulaire F_CLIENTS
' -----------------------------------------------------------------------------
Public Sub Etape5c_F_CLIENTS()
    Creer_F_CLIENTS
    MsgBox "F_CLIENTS cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5d - Formulaire F_PRESTATAIRES
' -----------------------------------------------------------------------------
Public Sub Etape5d_F_PRESTATAIRES()
    Creer_F_PRESTATAIRES
    MsgBox "F_PRESTATAIRES cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5e - Formulaire F_RESSOURCES
' -----------------------------------------------------------------------------
Public Sub Etape5e_F_RESSOURCES()
    Creer_F_RESSOURCES
    MsgBox "F_RESSOURCES cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5f - Formulaire F_RESERVATIONS
' -----------------------------------------------------------------------------
Public Sub Etape5f_F_RESERVATIONS()
    Creer_F_RESERVATIONS
    MsgBox "F_RESERVATIONS cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5g - Formulaire F_NOUVELLE_RESERVATION
' -----------------------------------------------------------------------------
Public Sub Etape5g_F_NOUVELLE_RESERVATION()
    Creer_F_NOUVELLE_RESERVATION
    MsgBox "F_NOUVELLE_RESERVATION cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5h - Formulaire F_PAIEMENTS
' -----------------------------------------------------------------------------
Public Sub Etape5h_F_PAIEMENTS()
    Creer_F_PAIEMENTS
    MsgBox "F_PAIEMENTS cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5i - Formulaire F_UTILISATEURS
' -----------------------------------------------------------------------------
Public Sub Etape5i_F_UTILISATEURS()
    Creer_F_UTILISATEURS
    MsgBox "F_UTILISATEURS cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5j - Formulaire F_PARAMETRES
' -----------------------------------------------------------------------------
Public Sub Etape5j_F_PARAMETRES()
    Creer_F_PARAMETRES
    MsgBox "F_PARAMETRES cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 5k - Formulaire F_ETATS_RAPPORTS
' -----------------------------------------------------------------------------
Public Sub Etape5k_F_ETATS_RAPPORTS()
    Creer_F_ETATS_RAPPORTS
    MsgBox "F_ETATS_RAPPORTS cree - OK", vbInformation
End Sub

' -----------------------------------------------------------------------------
' ETAPE 6 - Configuration demarrage
' -----------------------------------------------------------------------------
Public Sub Etape6_Demarrage()
    On Error Resume Next
    Application.SetOption "StartUpForm", "F_CONNEXION"
    Application.SetOption "StartUpShowDBWindow", False
    Application.SetOption "StartUpShowStatusBar", True
    On Error GoTo 0
    MsgBox "Demarrage configure - F_CONNEXION s'ouvrira au lancement", vbInformation
End Sub

' -----------------------------------------------------------------------------
' BILAN FINAL
' -----------------------------------------------------------------------------
Public Sub BilanFinal()
    DiagnosticComplet
End Sub
