/* ============================================================
   RESERV+ - Index et contraintes
   Fichier : 01_index_et_contraintes.sql
   IMPORTANT : Exécuter chaque instruction séparément dans Access
               via le module modInstall_RESERV_PLUS.bas
   ============================================================ */

/* ---- Index CATEGORIE ---- */
CREATE UNIQUE INDEX IDX_CAT_LIBELLE
    ON CATEGORIE (libelle_categorie);

/* ---- Index RESSOURCE ---- */
CREATE INDEX IDX_RES_PRESTATAIRE
    ON RESSOURCE (id_prestataire);

CREATE INDEX IDX_RES_CATEGORIE
    ON RESSOURCE (id_categorie);

CREATE INDEX IDX_RES_DISPONIBLE
    ON RESSOURCE (disponible);

/* ---- Index TAXI ---- */
CREATE UNIQUE INDEX IDX_TAXI_IMMAT
    ON TAXI (immatriculation);

/* ---- Index VOL ---- */
CREATE UNIQUE INDEX IDX_VOL_NUM
    ON VOL (num_vol);

/* ---- Index CLIENT ---- */
CREATE INDEX IDX_CLI_NOM
    ON CLIENT (nom, prenom);

CREATE INDEX IDX_CLI_EMAIL
    ON CLIENT (email);

/* ---- Index UTILISATEUR ---- */
CREATE UNIQUE INDEX IDX_USR_LOGIN
    ON UTILISATEUR (login);

/* ---- Index RESERVATION ---- */
CREATE INDEX IDX_RES_CLIENT
    ON RESERVATION (id_client);

CREATE INDEX IDX_RES_UTILISATEUR
    ON RESERVATION (id_utilisateur);

CREATE INDEX IDX_RES_STATUT
    ON RESERVATION (statut);

CREATE INDEX IDX_RES_DATES
    ON RESERVATION (date_debut, date_fin);

CREATE INDEX IDX_RES_DATE_CREATION
    ON RESERVATION (date_creation);

/* ---- Index DETAIL_RESERVATION ---- */
CREATE INDEX IDX_DET_RESERVATION
    ON DETAIL_RESERVATION (id_reservation);

CREATE INDEX IDX_DET_RESSOURCE
    ON DETAIL_RESERVATION (id_ressource);

/* ---- Index PAIEMENT ---- */
CREATE INDEX IDX_PAI_RESERVATION
    ON PAIEMENT (id_reservation);

CREATE INDEX IDX_PAI_DATE
    ON PAIEMENT (date_paiement);

/* ---- Index BILLET ---- */
CREATE INDEX IDX_BIL_RESERVATION
    ON BILLET (id_reservation);

CREATE UNIQUE INDEX IDX_BIL_NUMERO
    ON BILLET (numero_billet);

/* ---- Index JOURNAL_ACTION ---- */
CREATE INDEX IDX_JRN_UTILISATEUR
    ON JOURNAL_ACTION (id_utilisateur);

CREATE INDEX IDX_JRN_DATE
    ON JOURNAL_ACTION (date_action);

/* ---- Index PARAMETRE ---- */
CREATE UNIQUE INDEX IDX_PAR_CLE
    ON PARAMETRE (cle_parametre);
