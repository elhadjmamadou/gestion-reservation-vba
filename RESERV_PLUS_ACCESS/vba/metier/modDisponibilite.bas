Attribute VB_Name = "modDisponibilite"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Gestion de la disponibilité des ressources
' Module : modDisponibilite.bas
' ============================================================

' ============================================================
' Vérifier si une ressource est disponible pour une période donnée
'
' Règle de chevauchement (conflit détecté si) :
'   date_debut_existante < nouvelle_date_fin
'   ET date_fin_existante > nouvelle_date_debut
'
' Les réservations "Annulée" sont ignorées.
' idReservationIgnore : ignorer une réservation lors d'une modification
' ============================================================
Public Function EstRessourceDisponible(idRessource As Long, _
                                        dateDebut As Date, _
                                        dateFin As Date, _
                                        Optional idReservationIgnore As Long = 0) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim nbConflits As Long

    On Error GoTo GestionErreur

    ' Vérifier d'abord que la ressource n'est pas marquée indisponible
    Dim ressourceActive As Variant
    ressourceActive = ObtenirValeur("SELECT disponible FROM RESSOURCE WHERE id_ressource = " & idRessource)
    If Not IsNull(ressourceActive) And ressourceActive = False Then
        EstRessourceDisponible = False
        Exit Function
    End If

    ' Requête de détection des conflits de réservation
    sql = "SELECT Count(*) FROM DETAIL_RESERVATION AS DR " & _
          "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
          "WHERE DR.id_ressource = " & idRessource & _
          " AND R.statut <> '" & STATUT_ANNULEE & "'" & _
          " AND R.date_debut < #" & Format(dateFin, "yyyy/mm/dd") & "#" & _
          " AND R.date_fin > #" & Format(dateDebut, "yyyy/mm/dd") & "#"

    ' Exclure la réservation en cours de modification
    If idReservationIgnore > 0 Then
        sql = sql & " AND R.id_reservation <> " & idReservationIgnore
    End If

    nbConflits = CLng(Nz(ObtenirValeur(sql, 0), 0))
    EstRessourceDisponible = (nbConflits = 0)
    Exit Function

GestionErreur:
    JournaliserErreur "EstRessourceDisponible", Err.Description
    EstRessourceDisponible = False
End Function

' ============================================================
' Obtenir la liste des périodes d'occupation d'une ressource
' Retourne un Recordset avec les dates de début et fin occupées
' ============================================================
Public Function ObtenirPeriodesOccupation(idRessource As Long) As DAO.Recordset
    Dim sql As String
    sql = "SELECT R.id_reservation, R.date_debut, R.date_fin, R.statut, " & _
          "C.nom & ' ' & C.prenom AS client_nom " & _
          "FROM (DETAIL_RESERVATION AS DR " & _
          "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation) " & _
          "INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
          "WHERE DR.id_ressource = " & idRessource & _
          " AND R.statut <> '" & STATUT_ANNULEE & "'" & _
          " AND R.date_fin >= Date() " & _
          "ORDER BY R.date_debut"

    Set ObtenirPeriodesOccupation = OuvrirRecordset(sql)
End Function

' ============================================================
' Obtenir le statut de disponibilité sous forme de texte
' ============================================================
Public Function StatutDisponibiliteTexte(idRessource As Long, dateDebut As Date, dateFin As Date) As String
    If EstRessourceDisponible(idRessource, dateDebut, dateFin) Then
        StatutDisponibiliteTexte = "Disponible"
    Else
        StatutDisponibiliteTexte = "Indisponible"
    End If
End Function

' ============================================================
' Vérifier la disponibilité et afficher un message à l'utilisateur
' ============================================================
Public Sub VerifierEtAfficherDisponibilite(idRessource As Long, dateDebut As Date, _
                                            dateFin As Date, _
                                            Optional idReservationIgnore As Long = 0)
    Dim disponible As Boolean
    disponible = EstRessourceDisponible(idRessource, dateDebut, dateFin, idReservationIgnore)

    If disponible Then
        MsgBox "Cette ressource est DISPONIBLE pour la période du " & _
               FormatDate(dateDebut) & " au " & FormatDate(dateFin) & ".", _
               vbInformation, APP_NOM & " - Disponibilité"
    Else
        MsgBox "ATTENTION : Cette ressource est INDISPONIBLE pour la période du " & _
               FormatDate(dateDebut) & " au " & FormatDate(dateFin) & "." & vbCrLf & _
               "Veuillez choisir une autre période ou une autre ressource.", _
               vbExclamation, APP_NOM & " - Conflit de réservation"
    End If
End Sub

' ============================================================
' Obtenir les ressources disponibles pour une période et catégorie
' ============================================================
Public Function ObtenirRessourcesDisponibles(idCategorie As Long, _
                                              dateDebut As Date, _
                                              dateFin As Date) As DAO.Recordset
    Dim sql As String

    ' Ressources qui n'ont aucun conflit sur la période
    sql = "SELECT RS.id_ressource, RS.libelle, RS.prix_unitaire, RS.capacite, " & _
          "P.nom_entreprise AS prestataire " & _
          "FROM RESSOURCE AS RS " & _
          "INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire " & _
          "WHERE RS.id_categorie = " & idCategorie & _
          " AND RS.disponible = True " & _
          " AND RS.id_ressource NOT IN (" & _
          "  SELECT DR.id_ressource FROM DETAIL_RESERVATION AS DR " & _
          "  INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
          "  WHERE R.statut <> '" & STATUT_ANNULEE & "'" & _
          "  AND R.date_debut < #" & Format(dateFin, "yyyy/mm/dd") & "#" & _
          "  AND R.date_fin > #" & Format(dateDebut, "yyyy/mm/dd") & "#" & _
          ") ORDER BY RS.libelle"

    Set ObtenirRessourcesDisponibles = OuvrirRecordset(sql)
End Function

' ============================================================
' Calculer le taux d'occupation d'une ressource sur une période
' ============================================================
Public Function TauxOccupation(idRessource As Long, dateDebutPeriode As Date, dateFinPeriode As Date) As Double
    Dim sql As String
    Dim totalJours As Long
    Dim joursOccupes As Long
    Dim rs As DAO.Recordset

    totalJours = DateDiff("d", dateDebutPeriode, dateFinPeriode)
    If totalJours <= 0 Then
        TauxOccupation = 0
        Exit Function
    End If

    joursOccupes = 0
    sql = "SELECT R.date_debut, R.date_fin FROM DETAIL_RESERVATION AS DR " & _
          "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
          "WHERE DR.id_ressource = " & idRessource & _
          " AND R.statut <> '" & STATUT_ANNULEE & "'" & _
          " AND R.date_debut < #" & Format(dateFinPeriode, "yyyy/mm/dd") & "#" & _
          " AND R.date_fin > #" & Format(dateDebutPeriode, "yyyy/mm/dd") & "#"

    Set rs = OuvrirRecordset(sql)
    Do While Not rs.EOF
        Dim debut As Date, fin As Date
        debut = IIf(rs!date_debut < dateDebutPeriode, dateDebutPeriode, rs!date_debut)
        fin = IIf(rs!date_fin > dateFinPeriode, dateFinPeriode, rs!date_fin)
        joursOccupes = joursOccupes + DateDiff("d", debut, fin)
        rs.MoveNext
    Loop
    rs.Close

    TauxOccupation = (joursOccupes / totalJours) * 100
End Function
