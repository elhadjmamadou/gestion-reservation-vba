Attribute VB_Name = "modEtatsImprimables"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - États et rapports imprimables
' Module : modEtatsImprimables.bas
' ============================================================

' ============================================================
' Imprimer le billet d'une réservation via MsgBox (aperçu texte)
' Pour une vraie impression, utiliser un état Access
' ============================================================
Public Sub ImprimerBillet(idBillet As Long)
    JournaliserConsultation "Impression billet ID=" & idBillet
    ApercuBillet idBillet
End Sub

' ============================================================
' Imprimer le reçu de paiement (aperçu texte)
' ============================================================
Public Sub ImprimerRecuPaiement(idPaiement As Long)
    Dim rs As DAO.Recordset
    Dim sql As String
    sql = "SELECT P.id_paiement, P.date_paiement, P.montant, P.mode_paiement, " & _
          "P.reference_paiement, R.id_reservation, C.nom & ' ' & C.prenom AS client, " & _
          "C.telephone, C.email, R.montant_total, R.date_debut, R.date_fin " & _
          "FROM (PAIEMENT AS P INNER JOIN RESERVATION AS R ON P.id_reservation = R.id_reservation) " & _
          "INNER JOIN CLIENT AS C ON R.id_client = C.id_client " & _
          "WHERE P.id_paiement = " & idPaiement

    Set rs = OuvrirRecordset(sql)

    If rs.EOF Then
        AfficherErreur "Paiement introuvable."
        rs.Close
        Exit Sub
    End If

    Dim nomEtab As String
    nomEtab = ObtenirParametre("NOM_ETABLISSEMENT", "RESERV+")
    Dim devise As String
    devise = ObtenirParametre("DEVISE", "GNF")

    Dim t As String
    t = String(55, "═") & vbCrLf
    t = t & "  " & nomEtab & vbCrLf
    t = t & "  REÇU DE PAIEMENT OFFICIEL" & vbCrLf
    t = t & String(55, "═") & vbCrLf & vbCrLf
    t = t & "Référence   : " & ChaineSure(rs!reference_paiement) & vbCrLf
    t = t & "Date        : " & FormatDateHeure(rs!date_paiement) & vbCrLf & vbCrLf
    t = t & String(55, "─") & vbCrLf
    t = t & "CLIENT      : " & ChaineSure(rs!client) & vbCrLf
    t = t & "Téléphone   : " & ChaineSure(rs!telephone) & vbCrLf
    t = t & "Email       : " & ChaineSure(rs!email) & vbCrLf
    t = t & String(55, "─") & vbCrLf
    t = t & "Réservation : #" & ChaineSure(rs!id_reservation) & vbCrLf
    t = t & "Période     : " & FormatDate(rs!date_debut) & " → " & FormatDate(rs!date_fin) & vbCrLf
    t = t & "Total rés.  : " & Format(CCur(Nz(rs!montant_total, 0)), "#,##0") & " " & devise & vbCrLf
    t = t & String(55, "─") & vbCrLf
    t = t & "VERSÉ       : " & Format(CCur(Nz(rs!montant, 0)), "#,##0") & " " & devise & vbCrLf
    t = t & "Mode        : " & ChaineSure(rs!mode_paiement) & vbCrLf
    t = t & "Solde rest. : " & Format(SoldeReservation(CLng(rs!id_reservation)), "#,##0") & " " & devise & vbCrLf
    t = t & String(55, "═") & vbCrLf
    t = t & "Merci de votre confiance - " & nomEtab

    rs.Close
    MsgBox t, vbInformation, "Reçu de paiement - " & APP_NOM
    JournaliserConsultation "Reçu paiement #" & idPaiement
End Sub

' ============================================================
' Rapport de synthèse journalier (affichage MsgBox)
' ============================================================
Public Sub RapportJournalier()
    If Not ALeRole(ROLE_AGENT) Then Exit Sub

    Dim t As String
    Dim dateStr As String
    dateStr = FormatDate(Now())

    t = "═══ RAPPORT JOURNALIER - " & dateStr & " ═══" & vbCrLf & vbCrLf
    t = t & APP_NOM & " - " & ObtenirParametre("NOM_ETABLISSEMENT", "RESERV+") & vbCrLf
    t = t & String(45, "─") & vbCrLf & vbCrLf

    t = t & "RÉSERVATIONS" & vbCrLf
    t = t & "  Total      : " & NbTotalReservations() & vbCrLf
    t = t & "  Aujourd'hui: " & NbReservationsDuJour() & vbCrLf
    t = t & "  En attente : " & NbReservationsEnAttente() & vbCrLf & vbCrLf

    t = t & "FINANCIER" & vbCrLf
    t = t & "  Encaissé   : " & FormatMontant(TotalEncaisseGlobal()) & vbCrLf
    t = t & "  À encaisser: " & FormatMontant(TotalResteAEncaisser()) & vbCrLf
    t = t & "  Paiements/mois: " & NbPaiementsMois() & vbCrLf & vbCrLf

    ' CA par catégorie
    t = t & "CHIFFRE D'AFFAIRES PAR CATÉGORIE" & vbCrLf
    Dim rs As DAO.Recordset
    Set rs = ChiffreAffairesParCategorie()
    Do While Not rs.EOF
        t = t & "  " & ChaineSure(rs.Fields(0)) & " : " & _
            FormatMontant(CCur(Nz(rs.Fields(2), 0))) & vbCrLf
        rs.MoveNext
    Loop
    rs.Close

    t = t & String(45, "─") & vbCrLf
    t = t & "Généré par : " & g_NomComplet & " (" & g_Role & ")" & vbCrLf
    t = t & FormatDateHeure(Now())

    MsgBox t, vbInformation, "Rapport journalier - " & APP_NOM
    JournaliserConsultation "Rapport journalier du " & dateStr
End Sub

' ============================================================
' Exporter les réservations vers Excel via Access
' (utilise TransferSpreadsheet d'Access)
' ============================================================
Public Sub ExporterReservationsExcel()
    If Not ALeRole(ROLE_AGENT) Then Exit Sub

    On Error GoTo GestionErreur

    Dim cheminFichier As String
    cheminFichier = CurrentProject.Path & "\Export_Reservations_" & Format(Now(), "yyyymmdd_hhnn") & ".xlsx"

    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
                               RQ_RESERVATIONS_COMPLETES, cheminFichier, True

    JournaliserConsultation "Export Excel réservations → " & cheminFichier
    AfficherSucces "Fichier exporté avec succès !" & vbCrLf & "Emplacement : " & cheminFichier
    Exit Sub

GestionErreur:
    AfficherErreur "Erreur lors de l'export : " & Err.Description
End Sub
