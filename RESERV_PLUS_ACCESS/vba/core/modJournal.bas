Attribute VB_Name = "modJournal"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Journal des actions utilisateurs
' Module : modJournal.bas
' ============================================================

' ============================================================
' Enregistrer une action dans le journal JOURNAL_ACTION
' Procédure principale de journalisation
' ============================================================
Public Sub Journaliser(action As String, details As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo GestionErreur

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("JOURNAL_ACTION", dbOpenDynaset)

    rs.AddNew
    If g_IdUtilisateur > 0 Then
        rs!id_utilisateur = g_IdUtilisateur
    End If
    rs!action = Left(action, 150)
    rs!details = details
    rs!date_action = Now()
    rs.Update

    rs.Close
    Set rs = Nothing
    Exit Sub

GestionErreur:
    ' Le journal ne doit jamais bloquer l'application
    ' On ignore silencieusement les erreurs de journalisation
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
End Sub

' ============================================================
' Journaliser une création d'enregistrement
' ============================================================
Public Sub JournaliserCreation(entite As String, idEnregistrement As Long)
    Journaliser "CREATION_" & UCase(entite), _
                "Création " & entite & " ID=" & idEnregistrement & " par " & g_Login
End Sub

' ============================================================
' Journaliser une modification
' ============================================================
Public Sub JournaliserModification(entite As String, idEnregistrement As Long)
    Journaliser "MODIFICATION_" & UCase(entite), _
                "Modification " & entite & " ID=" & idEnregistrement & " par " & g_Login
End Sub

' ============================================================
' Journaliser une suppression
' ============================================================
Public Sub JournaliserSuppression(entite As String, idEnregistrement As Long)
    Journaliser "SUPPRESSION_" & UCase(entite), _
                "Suppression " & entite & " ID=" & idEnregistrement & " par " & g_Login
End Sub

' ============================================================
' Journaliser une erreur applicative
' ============================================================
Public Sub JournaliserErreur(source As String, messageErreur As String)
    Journaliser "ERREUR", "Source : " & source & " | Erreur : " & messageErreur
End Sub

' ============================================================
' Journaliser une consultation (rapport ou état)
' ============================================================
Public Sub JournaliserConsultation(rapport As String)
    Journaliser "CONSULTATION", "Consultation : " & rapport & " par " & g_Login
End Sub

' ============================================================
' Obtenir les dernières actions du journal (pour affichage)
' ============================================================
Public Function ObtenirDernieresActions(nbLignes As Integer) As DAO.Recordset
    Dim sql As String
    sql = "SELECT TOP " & nbLignes & " JA.id_action, " & _
          "Nz(U.nom_complet, 'Système') AS utilisateur, " & _
          "JA.action, JA.details, JA.date_action " & _
          "FROM JOURNAL_ACTION AS JA " & _
          "LEFT JOIN UTILISATEUR AS U ON JA.id_utilisateur = U.id_utilisateur " & _
          "ORDER BY JA.date_action DESC"

    Set ObtenirDernieresActions = OuvrirRecordset(sql)
End Function

' ============================================================
' Vider le journal (admin seulement, avec confirmation)
' ============================================================
Public Sub ViderJournal()
    If Not ALeRole(ROLE_ADMINISTRATEUR) Then
        MsgBox "Action réservée à l'administrateur.", vbCritical, APP_NOM
        Exit Sub
    End If

    Dim rep As Integer
    rep = MsgBox("Voulez-vous vraiment vider l'intégralité du journal des actions ?" & vbCrLf & _
                 "Cette opération est irréversible.", _
                 vbCritical + vbYesNo, APP_NOM)

    If rep = vbYes Then
        ExecuterSQL "DELETE FROM JOURNAL_ACTION"
        ' Journaliser la suppression du journal elle-même
        Journaliser "VIDAGE_JOURNAL", "Journal des actions vidé par " & g_Login
        MsgBox "Journal vidé avec succès.", vbInformation, APP_NOM
    End If
End Sub
