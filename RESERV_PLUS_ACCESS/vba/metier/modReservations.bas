Attribute VB_Name = "modReservations"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Logique métier des réservations
' Module : modReservations.bas
' ============================================================

' ============================================================
' Créer une nouvelle réservation
' Retourne l'ID de la réservation créée (0 si échec)
' ============================================================
Public Function CreerReservation(idClient As Long, dateDebut As Date, dateFin As Date, _
                                  nbPersonnes As Integer, _
                                  Optional observation As String = "") As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idReservation As Long

    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        CreerReservation = 0
        Exit Function
    End If

    ' Vérifications
    If idClient <= 0 Then
        AfficherErreur "Veuillez sélectionner un client."
        CreerReservation = 0
        Exit Function
    End If

    If dateDebut >= dateFin Then
        AfficherErreur "La date de début doit être antérieure à la date de fin."
        CreerReservation = 0
        Exit Function
    End If

    If nbPersonnes <= 0 Then nbPersonnes = 1

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("RESERVATION", dbOpenDynaset)

    rs.AddNew
    rs!id_client = idClient
    rs!id_utilisateur = g_IdUtilisateur
    rs!date_debut = dateDebut
    rs!date_fin = dateFin
    rs!nb_personnes = nbPersonnes
    rs!montant_total = 0
    rs!statut = STATUT_EN_ATTENTE
    rs!date_creation = Now()
    rs!observation = observation
    rs.Update

    idReservation = DernierID("RESERVATION", "id_reservation")

    rs.Close
    Set rs = Nothing

    JournaliserCreation "Réservation", idReservation
    CreerReservation = idReservation
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de la création de la réservation : " & Err.Description
    CreerReservation = 0
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Ajouter un détail (ligne) à une réservation
' Retourne True si succès
' ============================================================
Public Function AjouterDetailReservation(idReservation As Long, idRessource As Long, _
                                          quantite As Integer, Optional idResIgnore As Long = 0) As Boolean
    Dim db As DAO.Database
    Dim rsRes As DAO.Recordset
    Dim rsDet As DAO.Recordset
    Dim prixUnitaire As Currency
    Dim duree As Integer
    Dim sousTotal As Currency
    Dim dateDebut As Date, dateFin As Date

    On Error GoTo GestionErreur

    ' Récupérer les dates de la réservation
    Dim sqlRes As String
    sqlRes = "SELECT date_debut, date_fin FROM RESERVATION WHERE id_reservation = " & idReservation
    Set rsRes = CurrentDb.OpenRecordset(sqlRes, dbOpenSnapshot)
    If rsRes.EOF Then
        AfficherErreur "Réservation introuvable."
        AjouterDetailReservation = False
        GoTo Nettoyage
    End If
    dateDebut = rsRes!date_debut
    dateFin = rsRes!date_fin
    rsRes.Close

    ' Vérifier la disponibilité
    If Not EstRessourceDisponible(idRessource, dateDebut, dateFin, idResIgnore) Then
        AfficherErreur "Cette ressource est indisponible pour la période sélectionnée."
        AjouterDetailReservation = False
        GoTo Nettoyage
    End If

    ' Récupérer le prix unitaire de la ressource
    prixUnitaire = CCur(Nz(ObtenirValeur("SELECT prix_unitaire FROM RESSOURCE WHERE id_ressource = " & idRessource), 0))

    ' Calculer la durée et le sous-total
    duree = DureeEnJours(dateDebut, dateFin)
    sousTotal = duree * prixUnitaire * quantite

    ' Insérer le détail
    Set rsDet = CurrentDb.OpenRecordset("DETAIL_RESERVATION", dbOpenDynaset)
    rsDet.AddNew
    rsDet!id_reservation = idReservation
    rsDet!id_ressource = idRessource
    rsDet!quantite = quantite
    rsDet!prix_unitaire_applique = prixUnitaire
    rsDet!sous_total = sousTotal
    rsDet.Update
    rsDet.Close

    ' Mettre à jour le montant total de la réservation
    RecalculerMontantTotal idReservation

    AjouterDetailReservation = True

Nettoyage:
    If Not rsRes Is Nothing Then On Error Resume Next: rsRes.Close
    If Not rsDet Is Nothing Then On Error Resume Next: rsDet.Close
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de l'ajout du détail : " & Err.Description
    AjouterDetailReservation = False
    Resume Nettoyage
End Function

' ============================================================
' Recalculer et mettre à jour le montant total d'une réservation
' ============================================================
Public Sub RecalculerMontantTotal(idReservation As Long)
    Dim sql As String
    Dim montantTotal As Currency

    sql = "SELECT Sum(sous_total) FROM DETAIL_RESERVATION WHERE id_reservation = " & idReservation
    montantTotal = CCur(Nz(ObtenirValeur(sql, 0), 0))

    ExecuterSQL "UPDATE RESERVATION SET montant_total = " & montantTotal & _
                " WHERE id_reservation = " & idReservation
End Sub

' ============================================================
' Calculer le montant d'un détail (durée × prix × quantité)
' ============================================================
Public Function CalculerSousTotal(dateDebut As Date, dateFin As Date, _
                                   prixUnitaire As Currency, quantite As Integer) As Currency
    Dim duree As Integer
    duree = DureeEnJours(dateDebut, dateFin)
    CalculerSousTotal = duree * prixUnitaire * quantite
End Function

' ============================================================
' Confirmer une réservation (passer de "En attente" à "Confirmée")
' ============================================================
Public Function ConfirmerReservation(idReservation As Long) As Boolean
    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        ConfirmerReservation = False
        Exit Function
    End If

    Dim statut As String
    statut = ChaineSure(ObtenirValeur("SELECT statut FROM RESERVATION WHERE id_reservation = " & idReservation))

    If statut = STATUT_ANNULEE Then
        AfficherErreur "Impossible de confirmer une réservation annulée."
        ConfirmerReservation = False
        Exit Function
    End If

    ExecuterSQL "UPDATE RESERVATION SET statut = '" & STATUT_CONFIRMEE & "' WHERE id_reservation = " & idReservation
    Journaliser "CONFIRMATION_RESERVATION", "Réservation #" & idReservation & " confirmée par " & g_Login
    ConfirmerReservation = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de la confirmation : " & Err.Description
    ConfirmerReservation = False
End Function

' ============================================================
' Annuler une réservation
' ============================================================
Public Function AnnulerReservation(idReservation As Long, Optional motif As String = "") As Boolean
    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        AnnulerReservation = False
        Exit Function
    End If

    If Not DemanderConfirmation("Voulez-vous vraiment annuler la réservation #" & idReservation & " ?") Then
        AnnulerReservation = False
        Exit Function
    End If

    Dim obs As String
    obs = ChaineSure(ObtenirValeur("SELECT observation FROM RESERVATION WHERE id_reservation = " & idReservation))
    If Len(motif) > 0 Then
        obs = obs & IIf(Len(obs) > 0, " | ", "") & "Annulation : " & motif
    End If

    ExecuterSQL "UPDATE RESERVATION SET statut = '" & STATUT_ANNULEE & "', " & _
                "observation = '" & EchapperChaine(obs) & "' " & _
                "WHERE id_reservation = " & idReservation

    Journaliser "ANNULATION_RESERVATION", "Réservation #" & idReservation & " annulée. Motif : " & motif
    AnnulerReservation = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de l'annulation : " & Err.Description
    AnnulerReservation = False
End Function

' ============================================================
' Supprimer une réservation et ses détails (admin seulement)
' ============================================================
Public Sub SupprimerReservation(idReservation As Long)
    If Not VerifierAccesRole(ROLE_ADMINISTRATEUR) Then Exit Sub

    If Not DemanderConfirmation("Supprimer définitivement la réservation #" & idReservation & _
                                 " et tous ses détails, paiements et billets ?") Then Exit Sub

    On Error GoTo GestionErreur

    ExecuterSQL "DELETE FROM BILLET WHERE id_reservation = " & idReservation
    ExecuterSQL "DELETE FROM PAIEMENT WHERE id_reservation = " & idReservation
    ExecuterSQL "DELETE FROM DETAIL_RESERVATION WHERE id_reservation = " & idReservation
    ExecuterSQL "DELETE FROM RESERVATION WHERE id_reservation = " & idReservation

    JournaliserSuppression "Réservation", idReservation
    AfficherSucces "Réservation #" & idReservation & " supprimée avec succès."
    Exit Sub

GestionErreur:
    AfficherErreur "Erreur lors de la suppression : " & Err.Description
End Sub

' ============================================================
' Obtenir le résumé d'une réservation (pour affichage)
' ============================================================
Public Function ObtenirResumeReservation(idReservation As Long) As String
    Dim sql As String
    Dim rs As DAO.Recordset

    sql = "SELECT C.nom & ' ' & C.prenom AS client, R.date_debut, R.date_fin, " & _
          "R.montant_total, R.statut FROM RESERVATION AS R " & _
          "INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
          "WHERE R.id_reservation = " & idReservation

    Set rs = OuvrirRecordset(sql)

    If rs.EOF Then
        ObtenirResumeReservation = "Réservation introuvable."
    Else
        ObtenirResumeReservation = "Client : " & ChaineSure(rs!client) & vbCrLf & _
                                    "Période : " & FormatDate(rs!date_debut) & " - " & FormatDate(rs!date_fin) & vbCrLf & _
                                    "Montant : " & FormatMontant(CCur(Nz(rs!montant_total, 0))) & vbCrLf & _
                                    "Statut : " & ChaineSure(rs!statut)
    End If

    rs.Close
End Function
