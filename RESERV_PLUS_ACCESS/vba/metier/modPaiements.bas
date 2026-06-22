Attribute VB_Name = "modPaiements"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Gestion des paiements
' Module : modPaiements.bas
' ============================================================

' ============================================================
' Calculer le total des paiements reçus pour une réservation
' ============================================================
Public Function TotalPaye(idReservation As Long) As Currency
    Dim sql As String
    sql = "SELECT Sum(montant) FROM PAIEMENT WHERE id_reservation = " & idReservation
    TotalPaye = CCur(Nz(ObtenirValeur(sql, 0), 0))
End Function

' ============================================================
' Calculer le solde restant à payer pour une réservation
' ============================================================
Public Function SoldeReservation(idReservation As Long) As Currency
    Dim montantTotal As Currency
    Dim totalPaye As Currency

    montantTotal = CCur(Nz(ObtenirValeur("SELECT montant_total FROM RESERVATION WHERE id_reservation = " & idReservation), 0))
    totalPaye = TotalPaye(idReservation)
    SoldeReservation = montantTotal - totalPaye
End Function

' ============================================================
' Obtenir l'état de paiement d'une réservation
' Retourne : "Impayé", "Partiel" ou "Soldé"
' ============================================================
Public Function EtatPaiementReservation(idReservation As Long) As String
    Dim montantTotal As Currency
    Dim paye As Currency

    montantTotal = CCur(Nz(ObtenirValeur("SELECT montant_total FROM RESERVATION WHERE id_reservation = " & idReservation), 0))
    paye = TotalPaye(idReservation)

    If paye <= 0 Then
        EtatPaiementReservation = PAIEMENT_IMPAYE
    ElseIf paye >= montantTotal Then
        EtatPaiementReservation = PAIEMENT_SOLDE
    Else
        EtatPaiementReservation = PAIEMENT_PARTIEL
    End If
End Function

' ============================================================
' Enregistrer un paiement pour une réservation
' Retourne l'ID du paiement créé (0 si échec)
' ============================================================
Public Function EnregistrerPaiement(idReservation As Long, montant As Currency, _
                                     modePaiement As String, _
                                     Optional reference As String = "", _
                                     Optional commentaire As String = "") As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idPaiement As Long
    Dim solde As Currency

    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        EnregistrerPaiement = 0
        Exit Function
    End If

    ' Vérifications
    If idReservation <= 0 Then
        AfficherErreur "Réservation invalide."
        EnregistrerPaiement = 0
        Exit Function
    End If

    If montant <= 0 Then
        AfficherErreur "Le montant du paiement doit être supérieur à zéro."
        EnregistrerPaiement = 0
        Exit Function
    End If

    ' Vérifier que la réservation n'est pas annulée
    Dim statut As String
    statut = ChaineSure(ObtenirValeur("SELECT statut FROM RESERVATION WHERE id_reservation = " & idReservation))
    If statut = STATUT_ANNULEE Then
        AfficherErreur "Impossible d'enregistrer un paiement pour une réservation annulée."
        EnregistrerPaiement = 0
        Exit Function
    End If

    ' Avertir si le montant dépasse le solde
    solde = SoldeReservation(idReservation)
    If montant > solde And solde > 0 Then
        If Not DemanderConfirmation("Le montant saisi (" & FormatMontant(montant) & ") " & _
                                     "dépasse le solde restant (" & FormatMontant(solde) & ")." & vbCrLf & _
                                     "Continuer quand même ?") Then
            EnregistrerPaiement = 0
            Exit Function
        End If
    End If

    ' Générer une référence automatique si vide
    If Len(Trim(reference)) = 0 Then
        reference = "PAI-" & Format(Now(), "yyyymmdd") & "-" & GenererReference(6)
    End If

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("PAIEMENT", dbOpenDynaset)

    rs.AddNew
    rs!id_reservation = idReservation
    rs!date_paiement = Now()
    rs!montant = montant
    rs!mode_paiement = modePaiement
    rs!reference_paiement = reference
    rs!commentaire = commentaire
    rs.Update

    idPaiement = DernierID("PAIEMENT", "id_paiement")
    rs.Close
    Set rs = Nothing

    Journaliser "PAIEMENT", "Paiement de " & FormatMontant(montant) & _
                " enregistré pour réservation #" & idReservation & _
                " - Mode : " & modePaiement & " - Réf : " & reference

    EnregistrerPaiement = idPaiement
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de l'enregistrement du paiement : " & Err.Description
    EnregistrerPaiement = 0
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Supprimer un paiement (admin seulement)
' ============================================================
Public Sub SupprimerPaiement(idPaiement As Long)
    If Not VerifierAccesRole(ROLE_ADMINISTRATEUR) Then Exit Sub

    If Not DemanderConfirmation("Supprimer ce paiement ? Cette action est irréversible.") Then Exit Sub

    On Error GoTo GestionErreur
    ExecuterSQL "DELETE FROM PAIEMENT WHERE id_paiement = " & idPaiement
    JournaliserSuppression "Paiement", idPaiement
    AfficherSucces "Paiement supprimé."
    Exit Sub

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
End Sub

' ============================================================
' Obtenir le récapitulatif financier du jour
' ============================================================
Public Function RecapFinancierJour() As String
    Dim sqlEncaisse As String
    Dim sqlPaiements As String
    Dim totalJour As Currency
    Dim nbPaiements As Long

    sqlEncaisse = "SELECT Sum(montant) FROM PAIEMENT WHERE DateValue(date_paiement) = Date()"
    sqlPaiements = "SELECT Count(*) FROM PAIEMENT WHERE DateValue(date_paiement) = Date()"

    totalJour = CCur(Nz(ObtenirValeur(sqlEncaisse, 0), 0))
    nbPaiements = CLng(Nz(ObtenirValeur(sqlPaiements, 0), 0))

    RecapFinancierJour = "Encaissé aujourd'hui : " & FormatMontant(totalJour) & _
                         " (" & nbPaiements & " paiement(s))"
End Function

' ============================================================
' Obtenir le total encaissé global
' ============================================================
Public Function TotalEncaisseGlobal() As Currency
    TotalEncaisseGlobal = CCur(Nz(ObtenirValeur("SELECT Sum(montant) FROM PAIEMENT"), 0))
End Function

' ============================================================
' Obtenir le total restant à encaisser (somme des soldes)
' ============================================================
Public Function TotalResteAEncaisser() As Currency
    Dim sql As String
    sql = "SELECT Sum(R.montant_total) - Nz(Sum(P.montant), 0) " & _
          "FROM RESERVATION AS R " & _
          "LEFT JOIN PAIEMENT AS P ON R.id_reservation = P.id_reservation " & _
          "WHERE R.statut <> '" & STATUT_ANNULEE & "'"

    TotalResteAEncaisser = CCur(Nz(ObtenirValeur(sql, 0), 0))
End Function

' ============================================================
' Obtenir le nombre de paiements du mois en cours
' ============================================================
Public Function NbPaiementsMois() As Long
    Dim sql As String
    sql = "SELECT Count(*) FROM PAIEMENT " & _
          "WHERE Year(date_paiement) = Year(Now()) AND Month(date_paiement) = Month(Now())"
    NbPaiementsMois = CLng(Nz(ObtenirValeur(sql, 0), 0))
End Function
