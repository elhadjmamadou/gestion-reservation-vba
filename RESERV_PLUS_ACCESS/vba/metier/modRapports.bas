Attribute VB_Name = "modRapports"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Fonctions de données pour les rapports
' Module : modRapports.bas
' ============================================================

' ============================================================
' Statistiques pour le tableau de bord
' ============================================================
Public Function NbTotalReservations() As Long
    NbTotalReservations = CompterEnregistrements("RESERVATION", "")
End Function

Public Function NbReservationsDuJour() As Long
    NbReservationsDuJour = CLng(Nz(ObtenirValeur( _
        "SELECT Count(*) FROM RESERVATION WHERE DateValue(date_creation) = Date()"), 0))
End Function

Public Function NbReservationsEnAttente() As Long
    NbReservationsEnAttente = CompterEnregistrements("RESERVATION", "statut = '" & STATUT_EN_ATTENTE & "'")
End Function

Public Function MontantEncaisseTotal() As Currency
    MontantEncaisseTotal = TotalEncaisseGlobal()
End Function

' ============================================================
' Données pour le graphique de répartition par catégorie
' ============================================================
Public Function ChiffreAffairesParCategorie() As DAO.Recordset
    Dim sql As String
    sql = "SELECT CAT.libelle_categorie, Count(DR.id_detail) AS nb_details, " & _
          "Sum(DR.sous_total) AS total FROM (DETAIL_RESERVATION AS DR " & _
          "INNER JOIN RESSOURCE AS RS ON DR.id_ressource = RS.id_ressource) " & _
          "INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie " & _
          "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
          "WHERE R.statut <> '" & STATUT_ANNULEE & "' " & _
          "GROUP BY CAT.libelle_categorie ORDER BY total DESC"
    Set ChiffreAffairesParCategorie = OuvrirRecordset(sql)
End Function

' ============================================================
' Dernières réservations pour le tableau de bord (top N)
' ============================================================
Public Function DernieresReservations(nbLignes As Integer) As DAO.Recordset
    Dim sql As String
    sql = "SELECT TOP " & nbLignes & " R.id_reservation, " & _
          "C.nom & ' ' & C.prenom AS client_nom, " & _
          "R.montant_total, R.statut, R.date_creation " & _
          "FROM RESERVATION AS R " & _
          "INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
          "ORDER BY R.date_creation DESC"
    Set DernieresReservations = OuvrirRecordset(sql)
End Function

' ============================================================
' Rapport chiffre d'affaires par période
' ============================================================
Public Function RapportCAParPeriode(dateDebut As Date, dateFin As Date) As DAO.Recordset
    Dim sql As String
    sql = "SELECT Year(R.date_creation) AS annee, Month(R.date_creation) AS mois, " & _
          "Count(R.id_reservation) AS nb_reservations, " & _
          "Sum(R.montant_total) AS montant_brut, " & _
          "Nz(Sum(P.montant), 0) AS montant_encaisse " & _
          "FROM RESERVATION AS R " & _
          "LEFT JOIN PAIEMENT AS P ON R.id_reservation = P.id_reservation " & _
          "WHERE R.statut <> '" & STATUT_ANNULEE & "' " & _
          "AND R.date_creation >= #" & Format(dateDebut, "yyyy/mm/dd") & "# " & _
          "AND R.date_creation <= #" & Format(dateFin, "yyyy/mm/dd") & "# " & _
          "GROUP BY Year(R.date_creation), Month(R.date_creation) " & _
          "ORDER BY annee DESC, mois DESC"
    Set RapportCAParPeriode = OuvrirRecordset(sql)
End Function

' ============================================================
' Rapport liste des réservations filtrée
' ============================================================
Public Function RapportReservations(Optional statut As String = "", _
                                     Optional dateDebut As Variant = Null, _
                                     Optional dateFin As Variant = Null) As DAO.Recordset
    Dim sql As String
    sql = "SELECT R.id_reservation, C.nom & ' ' & C.prenom AS client_nom, " & _
          "R.date_debut, R.date_fin, R.nb_personnes, R.montant_total, R.statut, " & _
          "R.date_creation, Nz(SQ.total_paye, 0) AS paye, " & _
          "R.montant_total - Nz(SQ.total_paye, 0) AS solde " & _
          "FROM (RESERVATION AS R " & _
          "INNER JOIN CLIENT AS C ON R.id_client = C.id_client) " & _
          "LEFT JOIN (SELECT id_reservation, Sum(montant) AS total_paye FROM PAIEMENT GROUP BY id_reservation) AS SQ " & _
          "ON R.id_reservation = SQ.id_reservation WHERE 1=1"

    If Len(Trim(statut)) > 0 And statut <> "Toutes" Then
        sql = sql & " AND R.statut = '" & EchapperChaine(statut) & "'"
    End If

    If Not IsNull(dateDebut) Then
        sql = sql & " AND R.date_creation >= #" & Format(dateDebut, "yyyy/mm/dd") & "#"
    End If

    If Not IsNull(dateFin) Then
        sql = sql & " AND R.date_creation <= #" & Format(dateFin, "yyyy/mm/dd") & "#"
    End If

    sql = sql & " ORDER BY R.date_creation DESC"
    Set RapportReservations = OuvrirRecordset(sql)
End Function

' ============================================================
' Rapport taux d'occupation par ressource
' ============================================================
Public Function RapportTauxOccupation(dateDebut As Date, dateFin As Date) As DAO.Recordset
    Dim sql As String
    sql = "SELECT RS.id_ressource, RS.libelle, CAT.libelle_categorie, " & _
          "P.nom_entreprise AS prestataire, " & _
          "Count(DR.id_detail) AS nb_sejours, " & _
          "Sum(DR.sous_total) AS chiffre_affaires " & _
          "FROM ((DETAIL_RESERVATION AS DR " & _
          "INNER JOIN RESSOURCE AS RS ON DR.id_ressource = RS.id_ressource) " & _
          "INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie) " & _
          "INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire " & _
          "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
          "WHERE R.statut <> '" & STATUT_ANNULEE & "' " & _
          "AND R.date_debut >= #" & Format(dateDebut, "yyyy/mm/dd") & "# " & _
          "AND R.date_fin <= #" & Format(dateFin, "yyyy/mm/dd") & "# " & _
          "GROUP BY RS.id_ressource, RS.libelle, CAT.libelle_categorie, P.nom_entreprise " & _
          "ORDER BY chiffre_affaires DESC"
    Set RapportTauxOccupation = OuvrirRecordset(sql)
End Function

' ============================================================
' Générer un texte de synthèse pour affichage rapide
' ============================================================
Public Function SyntheseDashboard() As String
    Dim texte As String
    texte = "=== TABLEAU DE BORD RESERV+ ===" & vbCrLf & vbCrLf
    texte = texte & "Réservations totales : " & NbTotalReservations() & vbCrLf
    texte = texte & "Réservations du jour : " & NbReservationsDuJour() & vbCrLf
    texte = texte & "En attente : " & NbReservationsEnAttente() & vbCrLf
    texte = texte & "Total encaissé : " & FormatMontant(MontantEncaisseTotal()) & vbCrLf
    texte = texte & "Reste à encaisser : " & FormatMontant(TotalResteAEncaisser()) & vbCrLf
    SyntheseDashboard = texte
End Function
