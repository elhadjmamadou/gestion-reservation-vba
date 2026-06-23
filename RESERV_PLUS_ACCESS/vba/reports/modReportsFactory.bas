Attribute VB_Name = "modReportsFactory"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Fabrique d'états imprimables
' Module : modReportsFactory.bas
' ============================================================

' ============================================================
' Créer l'état de rapport des réservations
' ============================================================
Public Sub Creer_RPT_RESERVATIONS()
    Dim rpt As Report
    Dim nomRpt As String
    nomRpt = "RPT_RESERVATIONS"

    On Error Resume Next
    DoCmd.DeleteObject acReport, nomRpt
    On Error GoTo 0

    Set rpt = CreateReport()

    rpt.RecordSource = RQ_RESERVATIONS_COMPLETES
    rpt.Caption = "Liste des réservations - " & APP_NOM

    ' En-tête rapport
    rpt.PageHeader.BackColor = COULEUR_PRIMAIRE
    rpt.PageHeader.Height = 900

    ' Titre
    Dim lblTitre As Access.Label
    Set lblTitre = CreateReportControl(nomRpt, acLabel, acPageHeader, , , 200, 100, 8000, 600)
    lblTitre.Caption = APP_NOM & " - Liste des réservations"
    lblTitre.FontName = POLICE
    lblTitre.FontSize = 14
    lblTitre.FontBold = True
    lblTitre.ForeColor = COULEUR_BLANC
    lblTitre.BackStyle = 0
    lblTitre.BorderStyle = 0

    ' Date d'impression
    Dim lblDate As Access.Label
    Set lblDate = CreateReportControl(nomRpt, acLabel, acPageHeader, , , 10000, 200, 4000, 380)
    lblDate.Caption = "Imprimé le : " & FormatDateHeure(Now())
    lblDate.FontName = POLICE
    lblDate.FontSize = 8
    lblDate.ForeColor = COULEUR_BLANC
    lblDate.BackStyle = 0
    lblDate.BorderStyle = 0

    ' Colonnes en-tête
    Dim colsRpt() As String
    Dim largColsRpt() As Long
    Dim xColRpt As Long

    colsRpt = Split("N°|Client|Début|Fin|Montant|Statut", "|")
    largColsRpt = Array(800, 3600, 2000, 2000, 2400, 2000)
    xColRpt = 200

    Dim cr As Integer
    For cr = 0 To UBound(colsRpt)
        Dim lblCol As Access.Label
        Set lblCol = CreateReportControl(nomRpt, acLabel, acPageHeader, , , xColRpt, 550, largColsRpt(cr), 320)
        lblCol.Caption = colsRpt(cr)
        lblCol.FontName = POLICE
        lblCol.FontSize = 9
        lblCol.FontBold = True
        lblCol.ForeColor = COULEUR_BLANC
        lblCol.BackStyle = 0
        lblCol.BorderStyle = 0
        xColRpt = xColRpt + largColsRpt(cr) + 100
    Next cr

    ' Zone détail
    rpt.Detail.Height = 480
    rpt.Detail.BackColor = COULEUR_BLANC

    Dim champs() As String
    champs = Split("id_reservation|client_nom|date_debut|date_fin|montant_total|statut", "|")
    Dim xDetail As Long
    xDetail = 200

    Dim cd As Integer
    For cd = 0 To UBound(champs)
        Dim txtDetail As Access.TextBox
        Set txtDetail = CreateReportControl(nomRpt, acTextBox, acDetail, , champs(cd), xDetail, 60, largColsRpt(cd), 360)
        txtDetail.FontName = POLICE
        txtDetail.FontSize = 9
        txtDetail.ForeColor = COULEUR_TEXTE
        txtDetail.BackStyle = 0
        txtDetail.BorderStyle = 0
        If champs(cd) = "montant_total" Then txtDetail.Format = "#,##0"
        If champs(cd) = "date_debut" Or champs(cd) = "date_fin" Then txtDetail.Format = "dd/mm/yyyy"
        xDetail = xDetail + largColsRpt(cd) + 100
    Next cd

    ' Pied de page
    rpt.PageFooter.Height = 480
    Dim lblPage As Access.Label
    Set lblPage = CreateReportControl(nomRpt, acLabel, acPageFooter, , , 200, 60, 6000, 320)
    lblPage.Caption = APP_NOM & " - Confidentiel"
    lblPage.FontName = POLICE
    lblPage.FontSize = 8
    lblPage.ForeColor = COULEUR_TEXTE_SECONDAIRE
    lblPage.BackStyle = 0
    lblPage.BorderStyle = 0

    Dim txtNumPage As Access.TextBox
    Set txtNumPage = CreateReportControl(nomRpt, acTextBox, acPageFooter, , , 12000, 60, 2000, 320)
    txtNumPage.ControlSource = "='Page ' & [Page] & '/' & [Pages]"
    txtNumPage.FontName = POLICE
    txtNumPage.FontSize = 8
    txtNumPage.ForeColor = COULEUR_TEXTE_SECONDAIRE
    txtNumPage.BackStyle = 0
    txtNumPage.BorderStyle = 0

    DoCmd.Save acReport, nomRpt
    DoCmd.Close acReport, nomRpt, acSaveYes
    Debug.Print "État " & nomRpt & " créé."
End Sub

' ============================================================
' Ouvrir un aperçu de l'état des réservations
' ============================================================
Public Function OuvrirRapportReservations() As Integer
    If Not ALeRole(ROLE_AGENT) Then
        OuvrirRapportReservations = 0
        Exit Function
    End If

    If EtatExiste("RPT_RESERVATIONS") Then
        DoCmd.OpenReport "RPT_RESERVATIONS", acViewPreview
        JournaliserConsultation "Rapport RPT_RESERVATIONS"
    Else
        MsgBox "L'état RPT_RESERVATIONS n'existe pas." & vbCrLf & _
               "Exécutez Creer_RPT_RESERVATIONS() pour le créer.", _
               vbExclamation, APP_NOM
    End If
    OuvrirRapportReservations = 0
End Function

' ============================================================
' Créer tous les états imprimables
' ============================================================
Public Sub CreerTousLesEtats()
    On Error Resume Next
    Creer_RPT_RESERVATIONS
    Debug.Print "=== États créés ==="
End Sub
