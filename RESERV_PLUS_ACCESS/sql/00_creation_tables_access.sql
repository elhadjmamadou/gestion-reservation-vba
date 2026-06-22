/* ============================================================
   RESERV+ - Création des tables Access
   Fichier : 00_creation_tables_access.sql
   IMPORTANT : Dans Access, exécuter chaque bloc séparément
               via le module modInstall_RESERV_PLUS.bas
   ============================================================ */

/* ---- TABLE PRESTATAIRE ---- */
CREATE TABLE PRESTATAIRE (
    id_prestataire AUTOINCREMENT PRIMARY KEY,
    nom_entreprise TEXT(150) NOT NULL,
    nom_responsable TEXT(120),
    email TEXT(120),
    telephone TEXT(30),
    adresse LONGTEXT,
    type_activite TEXT(80),
    date_adhesion DATETIME
);

/* ---- TABLE CATEGORIE ---- */
CREATE TABLE CATEGORIE (
    id_categorie AUTOINCREMENT PRIMARY KEY,
    libelle_categorie TEXT(80) NOT NULL,
    description LONGTEXT
);

/* ---- TABLE RESSOURCE ---- */
CREATE TABLE RESSOURCE (
    id_ressource AUTOINCREMENT PRIMARY KEY,
    id_prestataire LONG NOT NULL,
    id_categorie LONG NOT NULL,
    libelle TEXT(150) NOT NULL,
    description LONGTEXT,
    prix_unitaire CURRENCY,
    capacite INTEGER,
    disponible YESNO,
    statut_ressource TEXT(50),
    date_creation DATETIME
);

/* ---- TABLE CHAMBRE_HOTEL (spécialisation hôtel) ---- */
CREATE TABLE CHAMBRE_HOTEL (
    id_ressource LONG PRIMARY KEY,
    nb_etoiles INTEGER,
    type_chambre TEXT(80),
    nb_lits INTEGER,
    petit_dejeuner YESNO,
    vue TEXT(80)
);

/* ---- TABLE TAXI (spécialisation taxi) ---- */
CREATE TABLE TAXI (
    id_ressource LONG PRIMARY KEY,
    marque TEXT(80),
    modele TEXT(80),
    immatriculation TEXT(50),
    nb_places INTEGER,
    climatisation YESNO
);

/* ---- TABLE VOL (spécialisation vol) ---- */
CREATE TABLE VOL (
    id_ressource LONG PRIMARY KEY,
    num_vol TEXT(50),
    ville_depart TEXT(100),
    ville_arrivee TEXT(100),
    date_depart DATETIME,
    heure_depart DATETIME,
    date_arrivee DATETIME,
    heure_arrivee DATETIME,
    classe TEXT(80)
);

/* ---- TABLE CLIENT ---- */
CREATE TABLE CLIENT (
    id_client AUTOINCREMENT PRIMARY KEY,
    nom TEXT(100) NOT NULL,
    prenom TEXT(100),
    email TEXT(120),
    telephone TEXT(30),
    adresse LONGTEXT,
    date_inscription DATETIME
);

/* ---- TABLE UTILISATEUR ---- */
CREATE TABLE UTILISATEUR (
    id_utilisateur AUTOINCREMENT PRIMARY KEY,
    login TEXT(80) NOT NULL,
    mot_de_passe TEXT(255),
    nom_complet TEXT(150),
    role TEXT(50),
    actif YESNO,
    date_creation DATETIME
);

/* ---- TABLE RESERVATION ---- */
CREATE TABLE RESERVATION (
    id_reservation AUTOINCREMENT PRIMARY KEY,
    id_client LONG NOT NULL,
    id_utilisateur LONG,
    date_debut DATETIME,
    date_fin DATETIME,
    nb_personnes INTEGER,
    montant_total CURRENCY,
    statut TEXT(50),
    date_creation DATETIME,
    observation LONGTEXT
);

/* ---- TABLE DETAIL_RESERVATION (liaison n,n RESERVATION <-> RESSOURCE) ---- */
CREATE TABLE DETAIL_RESERVATION (
    id_detail AUTOINCREMENT PRIMARY KEY,
    id_reservation LONG NOT NULL,
    id_ressource LONG NOT NULL,
    quantite INTEGER,
    prix_unitaire_applique CURRENCY,
    sous_total CURRENCY
);

/* ---- TABLE PAIEMENT ---- */
CREATE TABLE PAIEMENT (
    id_paiement AUTOINCREMENT PRIMARY KEY,
    id_reservation LONG NOT NULL,
    date_paiement DATETIME,
    montant CURRENCY,
    mode_paiement TEXT(50),
    reference_paiement TEXT(100),
    commentaire LONGTEXT
);

/* ---- TABLE BILLET ---- */
CREATE TABLE BILLET (
    id_billet AUTOINCREMENT PRIMARY KEY,
    id_reservation LONG NOT NULL,
    numero_billet TEXT(80) NOT NULL,
    date_emission DATETIME,
    statut_billet TEXT(50)
);

/* ---- TABLE JOURNAL_ACTION ---- */
CREATE TABLE JOURNAL_ACTION (
    id_action AUTOINCREMENT PRIMARY KEY,
    id_utilisateur LONG,
    action TEXT(150),
    details LONGTEXT,
    date_action DATETIME
);

/* ---- TABLE PARAMETRE ---- */
CREATE TABLE PARAMETRE (
    id_parametre AUTOINCREMENT PRIMARY KEY,
    cle_parametre TEXT(100),
    valeur_parametre TEXT(255),
    description LONGTEXT
);
