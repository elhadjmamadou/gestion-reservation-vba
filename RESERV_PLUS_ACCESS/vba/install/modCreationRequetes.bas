Attribute VB_Name = "modCreationRequetes"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Création des requêtes enregistrées
' Module : modCreationRequetes.bas
' ============================================================

' ============================================================
' Créer ou remplacer une requête enregistrée
' ============================================================
Private Sub CreerRequete(nomRequete As String, sqlRequete As String)
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    Set db = CurrentDb()

    ' Supprimer si existante
    On Error Resume Next
    db.QueryDefs.Delete nomRequete
    On Error GoTo 0

    ' Créer la nouvelle requête
    Set qdf = db.CreateQueryDef(nomRequete, sqlRequete)
    qdf.Close
    Set qdf = Nothing

    Debug.Print "Requête créée : " & nomRequete
End Sub

' ============================================================
' Créer toutes les requêtes enregistrées de RESERV+
' ============================================================
Public Sub CreerToutesLesRequetes()
    Debug.Print "=== CRÉATION DES REQUÊTES RESERV+ ==="

    ' --- RQ_RESERVATIONS_COMPLETES ---
    CreerRequete RQ_RESERVATIONS_COMPLETES, _
        "SELECT R.id_reservation, C.nom & ' ' & C.prenom AS client_nom, " & _
        "C.telephone AS client_tel, U.nom_complet AS utilisateur, " & _
        "R.date_debut, R.date_fin, R.nb_personnes, R.montant_total, " & _
        "R.statut, R.date_creation, R.observation " & _
        "FROM (RESERVATION AS R " & _
        "INNER JOIN CLIENT AS C ON R.id_client = C.id_client) " & _
        "LEFT JOIN UTILISATEUR AS U ON R.id_utilisateur = U.id_utilisateur " & _
        "ORDER BY R.date_creation DESC"

    ' --- RQ_RESERVATIONS_DU_JOUR ---
    CreerRequete RQ_RESERVATIONS_DU_JOUR, _
        "SELECT R.id_reservation, C.nom & ' ' & C.prenom AS client_nom, " & _
        "R.date_debut, R.date_fin, R.montant_total, R.statut " & _
        "FROM RESERVATION AS R " & _
        "INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
        "WHERE DateValue(R.date_creation) = Date() " & _
        "ORDER BY R.date_creation DESC"

    ' --- RQ_RESERVATIONS_EN_ATTENTE ---
    CreerRequete RQ_RESERVATIONS_EN_ATTENTE, _
        "SELECT R.id_reservation, C.nom & ' ' & C.prenom AS client_nom, " & _
        "R.date_debut, R.date_fin, R.montant_total, R.statut " & _
        "FROM RESERVATION AS R " & _
        "INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
        "WHERE R.statut = 'En attente' " & _
        "ORDER BY R.date_debut ASC"

    ' --- RQ_TOTAL_ENCAISSE ---
    CreerRequete RQ_TOTAL_ENCAISSE, _
        "SELECT Sum(P.montant) AS total_encaisse, Count(P.id_paiement) AS nb_paiements " & _
        "FROM PAIEMENT AS P"

    ' --- RQ_SOLDE_PAR_RESERVATION ---
    CreerRequete RQ_SOLDE_PAR_RESERVATION, _
        "SELECT R.id_reservation, C.nom & ' ' & C.prenom AS client_nom, " & _
        "R.montant_total, Nz(Sum(P.montant),0) AS total_paye, " & _
        "R.montant_total - Nz(Sum(P.montant),0) AS solde, " & _
        "IIf(Nz(Sum(P.montant),0)=0,'Impayé'," & _
        "IIf(Nz(Sum(P.montant),0)>=R.montant_total,'Soldé','Partiel')) AS etat_paiement " & _
        "FROM (RESERVATION AS R " & _
        "INNER JOIN CLIENT AS C ON R.id_client = C.id_client) " & _
        "LEFT JOIN PAIEMENT AS P ON R.id_reservation = P.id_reservation " & _
        "GROUP BY R.id_reservation, C.nom, C.prenom, R.montant_total " & _
        "ORDER BY R.id_reservation DESC"

    ' --- RQ_CLIENTS_AVEC_NB_RESERVATIONS ---
    CreerRequete RQ_CLIENTS_NB_RESERVATIONS, _
        "SELECT C.id_client, C.nom, C.prenom, C.email, C.telephone, " & _
        "C.date_inscription, Count(R.id_reservation) AS nb_reservations " & _
        "FROM CLIENT AS C " & _
        "LEFT JOIN RESERVATION AS R ON C.id_client = R.id_client " & _
        "GROUP BY C.id_client, C.nom, C.prenom, C.email, C.telephone, C.date_inscription " & _
        "ORDER BY C.nom, C.prenom"

    ' --- RQ_RESSOURCES_AVEC_PRESTATAIRE_CATEGORIE ---
    CreerRequete RQ_RESSOURCES_PRESTATAIRE, _
        "SELECT RS.id_ressource, CAT.libelle_categorie AS categorie, " & _
        "P.nom_entreprise AS prestataire, RS.libelle, RS.description, " & _
        "RS.prix_unitaire, RS.capacite, RS.disponible, RS.statut_ressource " & _
        "FROM (RESSOURCE AS RS " & _
        "INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire) " & _
        "INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie " & _
        "ORDER BY CAT.libelle_categorie, RS.libelle"

    ' --- RQ_PAIEMENTS_COMPLETS ---
    CreerRequete RQ_PAIEMENTS_COMPLETS, _
        "SELECT PA.id_paiement, PA.id_reservation, " & _
        "C.nom & ' ' & C.prenom AS client_nom, " & _
        "PA.date_paiement, PA.montant, PA.mode_paiement, PA.reference_paiement, " & _
        "R.montant_total, " & _
        "Nz(SQ.total_paye,0) AS total_paye, " & _
        "R.montant_total - Nz(SQ.total_paye,0) AS solde " & _
        "FROM ((PAIEMENT AS PA " & _
        "INNER JOIN RESERVATION AS R ON PA.id_reservation = R.id_reservation) " & _
        "INNER JOIN CLIENT AS C ON R.id_client = C.id_client) " & _
        "LEFT JOIN (SELECT id_reservation, Sum(montant) AS total_paye " & _
        "FROM PAIEMENT GROUP BY id_reservation) AS SQ " & _
        "ON PA.id_reservation = SQ.id_reservation " & _
        "ORDER BY PA.date_paiement DESC"

    ' --- RQ_CHIFFRE_AFFAIRES_PAR_CATEGORIE ---
    CreerRequete RQ_CA_PAR_CATEGORIE, _
        "SELECT CAT.libelle_categorie AS categorie, " & _
        "Count(DR.id_detail) AS nb_lignes, " & _
        "Sum(DR.sous_total) AS chiffre_affaires " & _
        "FROM (DETAIL_RESERVATION AS DR " & _
        "INNER JOIN RESSOURCE AS RS ON DR.id_ressource = RS.id_ressource) " & _
        "INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie " & _
        "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
        "WHERE R.statut <> 'Annulée' " & _
        "GROUP BY CAT.libelle_categorie " & _
        "ORDER BY chiffre_affaires DESC"

    ' --- RQ_CHIFFRE_AFFAIRES_PAR_PERIODE ---
    CreerRequete RQ_CA_PAR_PERIODE, _
        "SELECT Year(R.date_creation) AS annee, Month(R.date_creation) AS mois, " & _
        "Count(R.id_reservation) AS nb_reservations, " & _
        "Sum(R.montant_total) AS montant_brut, " & _
        "Nz(Sum(P.montant),0) AS montant_encaisse " & _
        "FROM RESERVATION AS R " & _
        "LEFT JOIN PAIEMENT AS P ON R.id_reservation = P.id_reservation " & _
        "WHERE R.statut <> 'Annulée' " & _
        "GROUP BY Year(R.date_creation), Month(R.date_creation) " & _
        "ORDER BY Year(R.date_creation) DESC, Month(R.date_creation) DESC"

    ' --- RQ_JOURNAL_ACTIONS ---
    CreerRequete RQ_JOURNAL_ACTIONS, _
        "SELECT JA.id_action, Nz(U.nom_complet,'Système') AS utilisateur, " & _
        "U.role, JA.action, JA.details, JA.date_action " & _
        "FROM JOURNAL_ACTION AS JA " & _
        "LEFT JOIN UTILISATEUR AS U ON JA.id_utilisateur = U.id_utilisateur " & _
        "ORDER BY JA.date_action DESC"

    ' --- RQ_BILLETS_COMPLETS ---
    CreerRequete RQ_BILLETS_COMPLETS, _
        "SELECT B.id_billet, B.numero_billet, B.date_emission, B.statut_billet, " & _
        "R.id_reservation, C.nom & ' ' & C.prenom AS client_nom, " & _
        "C.telephone AS client_tel, R.date_debut, R.date_fin, " & _
        "R.montant_total, R.statut AS statut_reservation " & _
        "FROM ((BILLET AS B " & _
        "INNER JOIN RESERVATION AS R ON B.id_reservation = R.id_reservation) " & _
        "INNER JOIN CLIENT AS C ON R.id_client = C.id_client) " & _
        "ORDER BY B.date_emission DESC"

    Debug.Print "=== TOUTES LES REQUÊTES CRÉÉES ==="
End Sub

' ============================================================
' Supprimer toutes les requêtes RESERV+
' ============================================================
Public Sub SupprimerToutesLesRequetes()
    Dim noms() As String
    noms = Array(RQ_RESERVATIONS_COMPLETES, RQ_RESERVATIONS_DU_JOUR, _
                 RQ_RESERVATIONS_EN_ATTENTE, RQ_TOTAL_ENCAISSE, _
                 RQ_SOLDE_PAR_RESERVATION, RQ_CLIENTS_NB_RESERVATIONS, _
                 RQ_RESSOURCES_PRESTATAIRE, RQ_PAIEMENTS_COMPLETS, _
                 RQ_CA_PAR_CATEGORIE, RQ_CA_PAR_PERIODE, _
                 RQ_JOURNAL_ACTIONS, RQ_BILLETS_COMPLETS)

    Dim i As Integer
    For i = 0 To UBound(noms)
        SupprimerRequete noms(i)
        Debug.Print "Requête supprimée : " & noms(i)
    Next i
End Sub
