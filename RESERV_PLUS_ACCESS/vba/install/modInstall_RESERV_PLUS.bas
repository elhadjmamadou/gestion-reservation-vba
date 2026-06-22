Attribute VB_Name = "modInstall_RESERV_PLUS"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - MODULE D'INSTALLATION PRINCIPAL
' Module : modInstall_RESERV_PLUS.bas
'
' UTILISATION :
'   1. Ouvrir Access, créer une base vide RESERV_PLUS.accdb
'   2. ALT+F11 → Importer TOUS les fichiers .bas
'   3. Dans la fenêtre Exécution (Ctrl+G) : Installer_RESERV_PLUS
'   4. Patienter (~30 secondes)
'   5. Ouvrir F_CONNEXION avec admin / admin123
' ============================================================

' ============================================================
' PROCÉDURE PRINCIPALE D'INSTALLATION
' ============================================================
Public Sub Installer_RESERV_PLUS()
    Dim db As DAO.Database
    Dim debut As Date
    Dim duree As Long

    debut = Now()

    On Error GoTo GestionErreurInstall

    MsgBox "Démarrage de l'installation de RESERV+ v" & APP_VERSION & vbCrLf & vbCrLf & _
           "Cette opération va :" & vbCrLf & _
           "  1. Créer les tables" & vbCrLf & _
           "  2. Créer les index" & vbCrLf & _
           "  3. Insérer les données initiales" & vbCrLf & _
           "  4. Créer les requêtes" & vbCrLf & _
           "  5. Créer les formulaires" & vbCrLf & vbCrLf & _
           "Durée estimée : 30 à 60 secondes.", _
           vbInformation, "Installation RESERV+"

    Set db = CurrentDb()

    ' ---- ÉTAPE 1 : Créer les tables ----
    Debug.Print vbCrLf & "=== ÉTAPE 1 : CRÉATION DES TABLES ==="
    CreerTables db

    ' ---- ÉTAPE 2 : Créer les index ----
    Debug.Print vbCrLf & "=== ÉTAPE 2 : CRÉATION DES INDEX ==="
    CreerIndex db

    ' ---- ÉTAPE 3 : Insérer les données initiales ----
    Debug.Print vbCrLf & "=== ÉTAPE 3 : DONNÉES INITIALES ==="
    InsererDonneesInitiales db

    ' ---- ÉTAPE 4 : Créer les requêtes ----
    Debug.Print vbCrLf & "=== ÉTAPE 4 : CRÉATION DES REQUÊTES ==="
    CreerToutesLesRequetes

    ' ---- ÉTAPE 5 : Créer les formulaires ----
    Debug.Print vbCrLf & "=== ÉTAPE 5 : CRÉATION DES FORMULAIRES ==="
    CreerTousLesFormulaires

    ' ---- ÉTAPE 6 : Configurer le démarrage ----
    Debug.Print vbCrLf & "=== ÉTAPE 6 : CONFIGURATION DÉMARRAGE ==="
    ConfigurerDemarrage

    duree = DateDiff("s", debut, Now())

    MsgBox "╔══════════════════════════════════════╗" & vbCrLf & _
           "║    INSTALLATION RÉUSSIE !            ║" & vbCrLf & _
           "║    RESERV+ v" & APP_VERSION & "                   ║" & vbCrLf & _
           "╠══════════════════════════════════════╣" & vbCrLf & _
           "║  Durée : " & duree & " secondes                  " & vbCrLf & _
           "╠══════════════════════════════════════╣" & vbCrLf & _
           "║  CONNEXION :                         ║" & vbCrLf & _
           "║  Administrateur : admin / admin123   ║" & vbCrLf & _
           "║  Agent          : agent / agent123   ║" & vbCrLf & _
           "╠══════════════════════════════════════╣" & vbCrLf & _
           "║  → Ouvrez le formulaire F_CONNEXION  ║" & vbCrLf & _
           "╚══════════════════════════════════════╝", _
           vbInformation, "RESERV+ — Installation terminée"

    ' Ouvrir le formulaire de connexion
    If FormulaireExiste(FORM_CONNEXION) Then
        DoCmd.OpenForm FORM_CONNEXION
    End If

    Exit Sub

GestionErreurInstall:
    MsgBox "ERREUR lors de l'installation :" & vbCrLf & vbCrLf & _
           "Étape en cours : " & Err.Source & vbCrLf & _
           "Erreur : " & Err.Description & vbCrLf & vbCrLf & _
           "Consultez la fenêtre Exécution (Ctrl+G) pour les détails.", _
           vbCritical, "Erreur d'installation"
    Resume Next
End Sub

' ============================================================
' ÉTAPE 1 : Créer les tables (si elles n'existent pas)
' ============================================================
Private Sub CreerTables(db As DAO.Database)
    ' ---- PRESTATAIRE ----
    If Not TableExiste("PRESTATAIRE") Then
        db.Execute "CREATE TABLE PRESTATAIRE (" & _
                   "id_prestataire AUTOINCREMENT PRIMARY KEY, " & _
                   "nom_entreprise TEXT(150) NOT NULL, " & _
                   "nom_responsable TEXT(120), " & _
                   "email TEXT(120), " & _
                   "telephone TEXT(30), " & _
                   "adresse LONGTEXT, " & _
                   "type_activite TEXT(80), " & _
                   "date_adhesion DATETIME)"
        Debug.Print "Table PRESTATAIRE créée."
    Else
        Debug.Print "Table PRESTATAIRE déjà existante."
    End If

    ' ---- CATEGORIE ----
    If Not TableExiste("CATEGORIE") Then
        db.Execute "CREATE TABLE CATEGORIE (" & _
                   "id_categorie AUTOINCREMENT PRIMARY KEY, " & _
                   "libelle_categorie TEXT(80) NOT NULL, " & _
                   "description LONGTEXT)"
        Debug.Print "Table CATEGORIE créée."
    Else
        Debug.Print "Table CATEGORIE déjà existante."
    End If

    ' ---- RESSOURCE ----
    If Not TableExiste("RESSOURCE") Then
        db.Execute "CREATE TABLE RESSOURCE (" & _
                   "id_ressource AUTOINCREMENT PRIMARY KEY, " & _
                   "id_prestataire LONG NOT NULL, " & _
                   "id_categorie LONG NOT NULL, " & _
                   "libelle TEXT(150) NOT NULL, " & _
                   "description LONGTEXT, " & _
                   "prix_unitaire CURRENCY, " & _
                   "capacite INTEGER, " & _
                   "disponible YESNO, " & _
                   "statut_ressource TEXT(50), " & _
                   "date_creation DATETIME)"
        Debug.Print "Table RESSOURCE créée."
    Else
        Debug.Print "Table RESSOURCE déjà existante."
    End If

    ' ---- CHAMBRE_HOTEL ----
    If Not TableExiste("CHAMBRE_HOTEL") Then
        db.Execute "CREATE TABLE CHAMBRE_HOTEL (" & _
                   "id_ressource LONG PRIMARY KEY, " & _
                   "nb_etoiles INTEGER, " & _
                   "type_chambre TEXT(80), " & _
                   "nb_lits INTEGER, " & _
                   "petit_dejeuner YESNO, " & _
                   "vue TEXT(80))"
        Debug.Print "Table CHAMBRE_HOTEL créée."
    Else
        Debug.Print "Table CHAMBRE_HOTEL déjà existante."
    End If

    ' ---- TAXI ----
    If Not TableExiste("TAXI") Then
        db.Execute "CREATE TABLE TAXI (" & _
                   "id_ressource LONG PRIMARY KEY, " & _
                   "marque TEXT(80), " & _
                   "modele TEXT(80), " & _
                   "immatriculation TEXT(50), " & _
                   "nb_places INTEGER, " & _
                   "climatisation YESNO)"
        Debug.Print "Table TAXI créée."
    Else
        Debug.Print "Table TAXI déjà existante."
    End If

    ' ---- VOL ----
    If Not TableExiste("VOL") Then
        db.Execute "CREATE TABLE VOL (" & _
                   "id_ressource LONG PRIMARY KEY, " & _
                   "num_vol TEXT(50), " & _
                   "ville_depart TEXT(100), " & _
                   "ville_arrivee TEXT(100), " & _
                   "date_depart DATETIME, " & _
                   "heure_depart DATETIME, " & _
                   "date_arrivee DATETIME, " & _
                   "heure_arrivee DATETIME, " & _
                   "classe TEXT(80))"
        Debug.Print "Table VOL créée."
    Else
        Debug.Print "Table VOL déjà existante."
    End If

    ' ---- CLIENT ----
    If Not TableExiste("CLIENT") Then
        db.Execute "CREATE TABLE CLIENT (" & _
                   "id_client AUTOINCREMENT PRIMARY KEY, " & _
                   "nom TEXT(100) NOT NULL, " & _
                   "prenom TEXT(100), " & _
                   "email TEXT(120), " & _
                   "telephone TEXT(30), " & _
                   "adresse LONGTEXT, " & _
                   "date_inscription DATETIME)"
        Debug.Print "Table CLIENT créée."
    Else
        Debug.Print "Table CLIENT déjà existante."
    End If

    ' ---- UTILISATEUR ----
    If Not TableExiste("UTILISATEUR") Then
        db.Execute "CREATE TABLE UTILISATEUR (" & _
                   "id_utilisateur AUTOINCREMENT PRIMARY KEY, " & _
                   "login TEXT(80) NOT NULL, " & _
                   "mot_de_passe TEXT(255), " & _
                   "nom_complet TEXT(150), " & _
                   "role TEXT(50), " & _
                   "actif YESNO, " & _
                   "date_creation DATETIME)"
        Debug.Print "Table UTILISATEUR créée."
    Else
        Debug.Print "Table UTILISATEUR déjà existante."
    End If

    ' ---- RESERVATION ----
    If Not TableExiste("RESERVATION") Then
        db.Execute "CREATE TABLE RESERVATION (" & _
                   "id_reservation AUTOINCREMENT PRIMARY KEY, " & _
                   "id_client LONG NOT NULL, " & _
                   "id_utilisateur LONG, " & _
                   "date_debut DATETIME, " & _
                   "date_fin DATETIME, " & _
                   "nb_personnes INTEGER, " & _
                   "montant_total CURRENCY, " & _
                   "statut TEXT(50), " & _
                   "date_creation DATETIME, " & _
                   "observation LONGTEXT)"
        Debug.Print "Table RESERVATION créée."
    Else
        Debug.Print "Table RESERVATION déjà existante."
    End If

    ' ---- DETAIL_RESERVATION ----
    If Not TableExiste("DETAIL_RESERVATION") Then
        db.Execute "CREATE TABLE DETAIL_RESERVATION (" & _
                   "id_detail AUTOINCREMENT PRIMARY KEY, " & _
                   "id_reservation LONG NOT NULL, " & _
                   "id_ressource LONG NOT NULL, " & _
                   "quantite INTEGER, " & _
                   "prix_unitaire_applique CURRENCY, " & _
                   "sous_total CURRENCY)"
        Debug.Print "Table DETAIL_RESERVATION créée."
    Else
        Debug.Print "Table DETAIL_RESERVATION déjà existante."
    End If

    ' ---- PAIEMENT ----
    If Not TableExiste("PAIEMENT") Then
        db.Execute "CREATE TABLE PAIEMENT (" & _
                   "id_paiement AUTOINCREMENT PRIMARY KEY, " & _
                   "id_reservation LONG NOT NULL, " & _
                   "date_paiement DATETIME, " & _
                   "montant CURRENCY, " & _
                   "mode_paiement TEXT(50), " & _
                   "reference_paiement TEXT(100), " & _
                   "commentaire LONGTEXT)"
        Debug.Print "Table PAIEMENT créée."
    Else
        Debug.Print "Table PAIEMENT déjà existante."
    End If

    ' ---- BILLET ----
    If Not TableExiste("BILLET") Then
        db.Execute "CREATE TABLE BILLET (" & _
                   "id_billet AUTOINCREMENT PRIMARY KEY, " & _
                   "id_reservation LONG NOT NULL, " & _
                   "numero_billet TEXT(80) NOT NULL, " & _
                   "date_emission DATETIME, " & _
                   "statut_billet TEXT(50))"
        Debug.Print "Table BILLET créée."
    Else
        Debug.Print "Table BILLET déjà existante."
    End If

    ' ---- JOURNAL_ACTION ----
    If Not TableExiste("JOURNAL_ACTION") Then
        db.Execute "CREATE TABLE JOURNAL_ACTION (" & _
                   "id_action AUTOINCREMENT PRIMARY KEY, " & _
                   "id_utilisateur LONG, " & _
                   "action TEXT(150), " & _
                   "details LONGTEXT, " & _
                   "date_action DATETIME)"
        Debug.Print "Table JOURNAL_ACTION créée."
    Else
        Debug.Print "Table JOURNAL_ACTION déjà existante."
    End If

    ' ---- PARAMETRE ----
    If Not TableExiste("PARAMETRE") Then
        db.Execute "CREATE TABLE PARAMETRE (" & _
                   "id_parametre AUTOINCREMENT PRIMARY KEY, " & _
                   "cle_parametre TEXT(100), " & _
                   "valeur_parametre TEXT(255), " & _
                   "description LONGTEXT)"
        Debug.Print "Table PARAMETRE créée."
    Else
        Debug.Print "Table PARAMETRE déjà existante."
    End If

    Debug.Print "→ Toutes les tables vérifiées/créées."
End Sub

' ============================================================
' ÉTAPE 2 : Créer les index (ignorer les erreurs si déjà existants)
' ============================================================
Private Sub CreerIndex(db As DAO.Database)
    Dim indexSQL() As String

    indexSQL = Array( _
        "CREATE INDEX IDX_RES_PRESTATAIRE ON RESSOURCE (id_prestataire)", _
        "CREATE INDEX IDX_RES_CATEGORIE ON RESSOURCE (id_categorie)", _
        "CREATE INDEX IDX_RES_DISPONIBLE ON RESSOURCE (disponible)", _
        "CREATE INDEX IDX_CLI_NOM ON CLIENT (nom, prenom)", _
        "CREATE UNIQUE INDEX IDX_USR_LOGIN ON UTILISATEUR (login)", _
        "CREATE INDEX IDX_RES_CLIENT ON RESERVATION (id_client)", _
        "CREATE INDEX IDX_RES_STATUT ON RESERVATION (statut)", _
        "CREATE INDEX IDX_RES_DATES ON RESERVATION (date_debut, date_fin)", _
        "CREATE INDEX IDX_DET_RESERVATION ON DETAIL_RESERVATION (id_reservation)", _
        "CREATE INDEX IDX_DET_RESSOURCE ON DETAIL_RESERVATION (id_ressource)", _
        "CREATE INDEX IDX_PAI_RESERVATION ON PAIEMENT (id_reservation)", _
        "CREATE INDEX IDX_PAI_DATE ON PAIEMENT (date_paiement)", _
        "CREATE UNIQUE INDEX IDX_BIL_NUMERO ON BILLET (numero_billet)", _
        "CREATE INDEX IDX_BIL_RESERVATION ON BILLET (id_reservation)", _
        "CREATE INDEX IDX_JRN_DATE ON JOURNAL_ACTION (date_action)", _
        "CREATE UNIQUE INDEX IDX_PAR_CLE ON PARAMETRE (cle_parametre)" _
    )

    Dim i As Integer
    For i = 0 To UBound(indexSQL)
        On Error Resume Next
        db.Execute indexSQL(i)
        If Err.Number <> 0 Then
            Debug.Print "Index existant ou ignoré : " & Err.Description
        Else
            Debug.Print "Index créé : " & indexSQL(i)
        End If
        On Error GoTo 0
    Next i

    Debug.Print "→ Index vérifiés/créés."
End Sub

' ============================================================
' ÉTAPE 3 : Insérer les données initiales
' ============================================================
Private Sub InsererDonneesInitiales(db As DAO.Database)
    ' Catégories
    If CompterEnregistrements("CATEGORIE", "") = 0 Then
        db.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Hôtel', 'Chambres et hébergements hôteliers')"
        db.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Taxi', 'Véhicules de transport terrestre')"
        db.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Vol', 'Vols et transport aérien')"
        Debug.Print "Catégories insérées."
    Else
        Debug.Print "Catégories déjà présentes."
    End If

    ' Utilisateurs
    If CompterEnregistrements("UTILISATEUR", "") = 0 Then
        db.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) " & _
                   "VALUES ('admin', 'admin123', 'Administrateur Système', 'Administrateur', True, Now())"
        db.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) " & _
                   "VALUES ('agent', 'agent123', 'Agent de Réservation', 'Agent', True, Now())"
        db.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) " & _
                   "VALUES ('utilisateur', 'user123', 'Utilisateur Standard', 'Utilisateur', True, Now())"
        Debug.Print "Utilisateurs insérés."
    Else
        Debug.Print "Utilisateurs déjà présents."
    End If

    ' Paramètres
    If CompterEnregistrements("PARAMETRE", "") = 0 Then
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('NOM_APPLICATION', 'RESERV+', 'Nom de l application')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('DEVISE', 'GNF', 'Devise utilisée')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('PREFIXE_BILLET', 'BIL', 'Préfixe numéro billet')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('NOM_ETABLISSEMENT', 'RESERV+', 'Nom sur les documents')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('VERSION_APP', '1.0.0', 'Version application')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre, description) VALUES ('FREQUENCE_SAUVEGARDE', '7', 'Sauvegarde en jours')"
        Debug.Print "Paramètres insérés."
    Else
        Debug.Print "Paramètres déjà présents."
    End If

    ' Prestataires de démonstration
    If CompterEnregistrements("PRESTATAIRE", "") = 0 Then
        db.Execute "INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, type_activite, date_adhesion) " & _
                   "VALUES ('Hôtel Kaloum SARL', 'M. Sylla', 'contact@hotelkaloum.gn', '620 11 22 33', 'Hôtellerie', Now())"
        db.Execute "INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, type_activite, date_adhesion) " & _
                   "VALUES ('Express Taxi Conakry', 'Mme Barry', 'info@expresstaxi.gn', '622 33 44 55', 'Transport terrestre', Now())"
        db.Execute "INSERT INTO PRESTATAIRE (nom_entreprise, nom_responsable, email, telephone, type_activite, date_adhesion) " & _
                   "VALUES ('Air Guinée Express', 'M. Camara', 'reservations@airgn.gn', '628 55 66 77', 'Aérien', Now())"
        Debug.Print "Prestataires démo insérés."
    Else
        Debug.Print "Prestataires déjà présents."
    End If

    ' Ressources de démonstration
    If CompterEnregistrements("RESSOURCE", "") = 0 Then
        Dim idCatHotel As Long
        Dim idCatTaxi As Long
        Dim idCatVol As Long
        idCatHotel = CLng(Nz(ObtenirValeur("SELECT id_categorie FROM CATEGORIE WHERE libelle_categorie = 'Hôtel'"), 1))
        idCatTaxi = CLng(Nz(ObtenirValeur("SELECT id_categorie FROM CATEGORIE WHERE libelle_categorie = 'Taxi'"), 2))
        idCatVol = CLng(Nz(ObtenirValeur("SELECT id_categorie FROM CATEGORIE WHERE libelle_categorie = 'Vol'"), 3))

        db.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, prix_unitaire, capacite, disponible, statut_ressource, date_creation) " & _
                   "VALUES (1, " & idCatHotel & ", 'Hôtel Kaloum - Chambre 204', 150000, 2, True, 'Disponible', Now())"
        db.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, prix_unitaire, capacite, disponible, statut_ressource, date_creation) " & _
                   "VALUES (2, " & idCatTaxi & ", 'Taxi Express 01', 30000, 4, True, 'Disponible', Now())"
        db.Execute "INSERT INTO RESSOURCE (id_prestataire, id_categorie, libelle, prix_unitaire, capacite, disponible, statut_ressource, date_creation) " & _
                   "VALUES (3, " & idCatVol & ", 'Vol CKY-LBE', 620000, 120, True, 'Disponible', Now())"

        ' Détails
        db.Execute "INSERT INTO CHAMBRE_HOTEL (id_ressource, nb_etoiles, type_chambre, nb_lits, petit_dejeuner, vue) VALUES (1, 4, 'Double', 2, True, 'Vue mer')"
        db.Execute "INSERT INTO TAXI (id_ressource, marque, modele, immatriculation, nb_places, climatisation) VALUES (2, 'Toyota', 'Corolla', 'RC-1234-CKY', 4, True)"
        db.Execute "INSERT INTO VOL (id_ressource, num_vol, ville_depart, ville_arrivee, date_depart, heure_depart, date_arrivee, heure_arrivee, classe) " & _
                   "VALUES (3, 'AGX-001', 'Conakry', 'Labé', #2026/06/25#, #2026/06/25 08:00:00#, #2026/06/25#, #2026/06/25 10:30:00#, 'Économique')"

        Debug.Print "Ressources démo insérées."
    Else
        Debug.Print "Ressources déjà présentes."
    End If

    ' Clients de démonstration
    If CompterEnregistrements("CLIENT", "") = 0 Then
        db.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, date_inscription) VALUES ('Diallo', 'Mamadou', 'm.diallo@mail.gn', '620 00 00 01', #2026/01/02#)"
        db.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, date_inscription) VALUES ('Camara', 'Aïssatou', 'a.camara@mail.gn', '622 11 22 33', #2026/02/15#)"
        db.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, date_inscription) VALUES ('Bah', 'Saikou', 's.bah@mail.gn', '628 44 55 66', #2026/03/20#)"
        db.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, date_inscription) VALUES ('Touré', 'Fatou', 'f.toure@mail.gn', '655 77 88 99', #2026/04/10#)"
        db.Execute "INSERT INTO CLIENT (nom, prenom, email, telephone, date_inscription) VALUES ('Soumah', 'Karim', 'k.soumah@mail.gn', '664 12 34 56', #2026/05/05#)"
        Debug.Print "Clients démo insérés."
    Else
        Debug.Print "Clients déjà présents."
    End If

    Debug.Print "→ Données initiales insérées."
End Sub

' ============================================================
' ÉTAPE 6 : Configurer le formulaire de démarrage
' ============================================================
Private Sub ConfigurerDemarrage()
    On Error Resume Next
    ' Configurer Access pour ouvrir F_CONNEXION au démarrage
    With Application
        .SetOption "StartUpForm", FORM_CONNEXION
        .SetOption "StartUpShowDBWindow", False
        .SetOption "StartUpShowStatusBar", True
        .SetOption "AllowFullMenus", False
        .SetOption "AllowBuiltinToolbars", False
        .SetOption "AllowToolbarChanges", False
        .SetOption "AllowShortcutMenus", False
    End With
    On Error GoTo 0
    Debug.Print "Configuration démarrage effectuée."
End Sub

' ============================================================
' RÉINITIALISATION COMPLÈTE (supprime tout et réinstalle)
' ============================================================
Public Sub SupprimerObjets_RESERV_PLUS()
    If Not DemanderConfirmation("Supprimer TOUS les formulaires et requêtes RESERV+ ?" & vbCrLf & _
                                 "Les TABLES ET DONNÉES seront conservées.") Then Exit Sub

    Debug.Print "=== SUPPRESSION DES OBJETS RESERV+ ==="
    SupprimerTousLesFormulaires
    SupprimerToutesLesRequetes
    Debug.Print "=== OBJETS SUPPRIMÉS - Relancez Installer_RESERV_PLUS() ==="
    MsgBox "Formulaires et requêtes supprimés." & vbCrLf & _
           "Relancez Installer_RESERV_PLUS() pour recréer.", vbInformation, APP_NOM
End Sub

' ============================================================
' SUPPRIMER TOUTES LES TABLES (ATTENTION : EFFACE TOUTES LES DONNÉES)
' ============================================================
Public Sub SupprimerToutesLesTablesDANGER()
    If Not ALeRole(ROLE_ADMINISTRATEUR) Then
        MsgBox "Accès refusé.", vbCritical, APP_NOM
        Exit Sub
    End If

    If Not DemanderConfirmation("DANGER : Supprimer TOUTES les tables et TOUTES les données ?" & vbCrLf & _
                                 "Cette opération est IRRÉVERSIBLE !") Then Exit Sub

    If Not DemanderConfirmation("Dernière confirmation : voulez-vous vraiment TOUT supprimer ?") Then Exit Sub

    Dim tables() As String
    tables = Array("JOURNAL_ACTION", "BILLET", "PAIEMENT", "DETAIL_RESERVATION", _
                   "RESERVATION", "CHAMBRE_HOTEL", "TAXI", "VOL", "RESSOURCE", _
                   "CLIENT", "UTILISATEUR", "PARAMETRE", "CATEGORIE", "PRESTATAIRE")

    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim i As Integer
    For i = 0 To UBound(tables)
        On Error Resume Next
        db.Execute "DROP TABLE " & tables(i)
        Debug.Print "Table supprimée : " & tables(i)
        On Error GoTo 0
    Next i

    SupprimerTousLesFormulaires
    SupprimerToutesLesRequetes
    MsgBox "Toutes les tables, formulaires et requêtes ont été supprimés.", vbInformation, APP_NOM
End Sub

' ============================================================
' TEST RAPIDE DE CONNECTIVITÉ
' ============================================================
Public Sub TestInstallation()
    Dim nbTables As Integer
    Dim nbRequetes As Integer
    Dim nbForms As Integer

    Dim tables() As String
    tables = Array("PRESTATAIRE", "CATEGORIE", "RESSOURCE", "CLIENT", "UTILISATEUR", _
                   "RESERVATION", "DETAIL_RESERVATION", "PAIEMENT", "BILLET", _
                   "JOURNAL_ACTION", "PARAMETRE", "CHAMBRE_HOTEL", "TAXI", "VOL")

    Dim i As Integer
    For i = 0 To UBound(tables)
        If TableExiste(tables(i)) Then nbTables = nbTables + 1
    Next i

    Dim formsList() As String
    formsList = Array(FORM_CONNEXION, FORM_DASHBOARD, FORM_CLIENTS, FORM_PRESTATAIRES, _
                      FORM_RESSOURCES, FORM_RESERVATIONS, FORM_NOUVELLE_RESERVATION, _
                      FORM_PAIEMENTS, FORM_UTILISATEURS, FORM_PARAMETRES, FORM_RAPPORTS)

    For i = 0 To UBound(formsList)
        If FormulaireExiste(formsList(i)) Then nbForms = nbForms + 1
    Next i

    If RequeteExiste(RQ_RESERVATIONS_COMPLETES) Then nbRequetes = nbRequetes + 1

    MsgBox "=== BILAN D'INSTALLATION RESERV+ ===" & vbCrLf & vbCrLf & _
           "Tables : " & nbTables & "/14" & IIf(nbTables = 14, " ✓", " ⚠ Manquantes") & vbCrLf & _
           "Formulaires : " & nbForms & "/11" & IIf(nbForms = 11, " ✓", " ⚠ Manquants") & vbCrLf & _
           "Requêtes : " & IIf(nbRequetes > 0, "OK ✓", "⚠ Manquantes") & vbCrLf & vbCrLf & _
           "Utilisateurs : " & CompterEnregistrements("UTILISATEUR", "") & vbCrLf & _
           "Clients démo : " & CompterEnregistrements("CLIENT", "") & vbCrLf & _
           "Ressources : " & CompterEnregistrements("RESSOURCE", "") & vbCrLf & vbCrLf & _
           IIf(nbTables = 14 And nbForms = 11, _
               "→ Installation complète ! Ouvrez F_CONNEXION.", _
               "→ Installation incomplète. Relancez Installer_RESERV_PLUS()."), _
           vbInformation, "Test Installation RESERV+"
End Sub
