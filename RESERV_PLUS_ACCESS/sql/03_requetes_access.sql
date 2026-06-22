/* ============================================================
   RESERV+ - Requêtes Access enregistrées
   Fichier : 03_requetes_access.sql
   Ces requêtes sont créées via modCreationRequetes.bas
   ============================================================ */

/* ---- RQ_RESERVATIONS_COMPLETES ---- */
/* Toutes les réservations avec nom client et utilisateur */
SELECT
    R.id_reservation,
    C.nom & " " & C.prenom AS client_nom,
    C.telephone AS client_tel,
    U.nom_complet AS utilisateur,
    R.date_debut,
    R.date_fin,
    R.nb_personnes,
    R.montant_total,
    R.statut,
    R.date_creation,
    R.observation
FROM (RESERVATION AS R
    INNER JOIN CLIENT AS C ON R.id_client = C.id_client)
    LEFT JOIN UTILISATEUR AS U ON R.id_utilisateur = U.id_utilisateur
ORDER BY R.date_creation DESC;

/* ---- RQ_RESERVATIONS_DU_JOUR ---- */
SELECT
    R.id_reservation,
    C.nom & " " & C.prenom AS client_nom,
    R.date_debut,
    R.date_fin,
    R.montant_total,
    R.statut
FROM RESERVATION AS R
    INNER JOIN CLIENT AS C ON R.id_client = C.id_client
WHERE DateValue(R.date_creation) = Date()
ORDER BY R.date_creation DESC;

/* ---- RQ_RESERVATIONS_EN_ATTENTE ---- */
SELECT
    R.id_reservation,
    C.nom & " " & C.prenom AS client_nom,
    R.date_debut,
    R.date_fin,
    R.montant_total,
    R.statut
FROM RESERVATION AS R
    INNER JOIN CLIENT AS C ON R.id_client = C.id_client
WHERE R.statut = 'En attente'
ORDER BY R.date_debut ASC;

/* ---- RQ_TOTAL_ENCAISSE ---- */
SELECT
    Sum(P.montant) AS total_encaisse,
    Count(P.id_paiement) AS nb_paiements
FROM PAIEMENT AS P;

/* ---- RQ_SOLDE_PAR_RESERVATION ---- */
SELECT
    R.id_reservation,
    C.nom & " " & C.prenom AS client_nom,
    R.montant_total,
    Nz(Sum(P.montant), 0) AS total_paye,
    R.montant_total - Nz(Sum(P.montant), 0) AS solde,
    IIf(Nz(Sum(P.montant), 0) = 0, 'Impayé',
        IIf(Nz(Sum(P.montant), 0) >= R.montant_total, 'Soldé', 'Partiel')) AS etat_paiement
FROM (RESERVATION AS R
    INNER JOIN CLIENT AS C ON R.id_client = C.id_client)
    LEFT JOIN PAIEMENT AS P ON R.id_reservation = P.id_reservation
GROUP BY R.id_reservation, C.nom, C.prenom, R.montant_total
ORDER BY R.id_reservation DESC;

/* ---- RQ_CLIENTS_AVEC_NB_RESERVATIONS ---- */
SELECT
    C.id_client,
    C.nom,
    C.prenom,
    C.email,
    C.telephone,
    C.date_inscription,
    Count(R.id_reservation) AS nb_reservations
FROM CLIENT AS C
    LEFT JOIN RESERVATION AS R ON C.id_client = R.id_client
GROUP BY C.id_client, C.nom, C.prenom, C.email, C.telephone, C.date_inscription
ORDER BY C.nom, C.prenom;

/* ---- RQ_RESSOURCES_AVEC_PRESTATAIRE_CATEGORIE ---- */
SELECT
    RS.id_ressource,
    P.nom_entreprise AS prestataire,
    CAT.libelle_categorie AS categorie,
    RS.libelle,
    RS.description,
    RS.prix_unitaire,
    RS.capacite,
    RS.disponible,
    RS.statut_ressource
FROM (RESSOURCE AS RS
    INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire)
    INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie
ORDER BY CAT.libelle_categorie, RS.libelle;

/* ---- RQ_PAIEMENTS_COMPLETS ---- */
SELECT
    PA.id_paiement,
    PA.id_reservation,
    C.nom & " " & C.prenom AS client_nom,
    PA.date_paiement,
    PA.montant,
    PA.mode_paiement,
    PA.reference_paiement,
    R.montant_total,
    Nz(SQ.total_paye, 0) AS total_paye,
    R.montant_total - Nz(SQ.total_paye, 0) AS solde
FROM ((PAIEMENT AS PA
    INNER JOIN RESERVATION AS R ON PA.id_reservation = R.id_reservation)
    INNER JOIN CLIENT AS C ON R.id_client = C.id_client)
    LEFT JOIN (
        SELECT id_reservation, Sum(montant) AS total_paye
        FROM PAIEMENT
        GROUP BY id_reservation
    ) AS SQ ON PA.id_reservation = SQ.id_reservation
ORDER BY PA.date_paiement DESC;

/* ---- RQ_CHIFFRE_AFFAIRES_PAR_CATEGORIE ---- */
SELECT
    CAT.libelle_categorie AS categorie,
    Count(DISTINCT R.id_reservation) AS nb_reservations,
    Sum(DR.sous_total) AS chiffre_affaires
FROM ((DETAIL_RESERVATION AS DR
    INNER JOIN RESSOURCE AS RS ON DR.id_ressource = RS.id_ressource)
    INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie)
    INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation
WHERE R.statut <> 'Annulée'
GROUP BY CAT.libelle_categorie
ORDER BY chiffre_affaires DESC;

/* ---- RQ_CHIFFRE_AFFAIRES_PAR_PERIODE ---- */
SELECT
    Year(R.date_creation) AS annee,
    Month(R.date_creation) AS mois,
    Count(R.id_reservation) AS nb_reservations,
    Sum(R.montant_total) AS montant_brut,
    Nz(Sum(P.montant), 0) AS montant_encaisse
FROM RESERVATION AS R
    LEFT JOIN PAIEMENT AS P ON R.id_reservation = P.id_reservation
WHERE R.statut <> 'Annulée'
GROUP BY Year(R.date_creation), Month(R.date_creation)
ORDER BY Year(R.date_creation) DESC, Month(R.date_creation) DESC;

/* ---- RQ_JOURNAL_ACTIONS ---- */
SELECT
    JA.id_action,
    Nz(U.nom_complet, 'Système') AS utilisateur,
    U.role,
    JA.action,
    JA.details,
    JA.date_action
FROM JOURNAL_ACTION AS JA
    LEFT JOIN UTILISATEUR AS U ON JA.id_utilisateur = U.id_utilisateur
ORDER BY JA.date_action DESC;

/* ---- RQ_BILLETS_COMPLETS ---- */
SELECT
    B.id_billet,
    B.numero_billet,
    B.date_emission,
    B.statut_billet,
    R.id_reservation,
    C.nom & " " & C.prenom AS client_nom,
    C.telephone AS client_tel,
    R.date_debut,
    R.date_fin,
    R.montant_total,
    R.statut AS statut_reservation
FROM ((BILLET AS B
    INNER JOIN RESERVATION AS R ON B.id_reservation = R.id_reservation)
    INNER JOIN CLIENT AS C ON R.id_client = C.id_client)
ORDER BY B.date_emission DESC;

/* ---- RQ_RESSOURCES_DISPONIBLES ---- */
SELECT
    RS.id_ressource,
    CAT.libelle_categorie AS categorie,
    P.nom_entreprise AS prestataire,
    RS.libelle,
    RS.prix_unitaire,
    RS.capacite
FROM (RESSOURCE AS RS
    INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie)
    INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire
WHERE RS.disponible = True
ORDER BY CAT.libelle_categorie, RS.libelle;

/* ---- RQ_DETAIL_RESERVATION_COMPLET ---- */
SELECT
    DR.id_detail,
    DR.id_reservation,
    RS.libelle AS ressource,
    CAT.libelle_categorie AS categorie,
    P.nom_entreprise AS prestataire,
    DR.quantite,
    DR.prix_unitaire_applique,
    DR.sous_total
FROM ((DETAIL_RESERVATION AS DR
    INNER JOIN RESSOURCE AS RS ON DR.id_ressource = RS.id_ressource)
    INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie)
    INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire
ORDER BY DR.id_reservation, DR.id_detail;
