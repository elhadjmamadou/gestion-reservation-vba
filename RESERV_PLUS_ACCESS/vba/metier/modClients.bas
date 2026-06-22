Attribute VB_Name = "modClients"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Gestion des clients
' Module : modClients.bas
' ============================================================

' ============================================================
' Créer un nouveau client
' Retourne l'ID créé (0 si échec)
' ============================================================
Public Function CreerClient(nom As String, prenom As String, email As String, _
                             telephone As String, adresse As String) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idClient As Long

    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        CreerClient = 0
        Exit Function
    End If

    If Not ChampObligatoire(nom, "Nom") Then
        CreerClient = 0
        Exit Function
    End If

    If Len(Trim(email)) > 0 And Not EstEmailValide(email) Then
        AfficherErreur "L'adresse email '" & email & "' n'est pas valide."
        CreerClient = 0
        Exit Function
    End If

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("CLIENT", dbOpenDynaset)

    rs.AddNew
    rs!nom = Trim(nom)
    rs!prenom = Trim(prenom)
    rs!email = Trim(email)
    rs!telephone = Trim(telephone)
    rs!adresse = Trim(adresse)
    rs!date_inscription = Now()
    rs.Update

    idClient = DernierID("CLIENT", "id_client")
    rs.Close
    Set rs = Nothing

    JournaliserCreation "Client", idClient
    CreerClient = idClient
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de la création du client : " & Err.Description
    CreerClient = 0
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Modifier un client existant
' ============================================================
Public Function ModifierClient(idClient As Long, nom As String, prenom As String, _
                                email As String, telephone As String, adresse As String) As Boolean
    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        ModifierClient = False
        Exit Function
    End If

    If Not ChampObligatoire(nom, "Nom") Then
        ModifierClient = False
        Exit Function
    End If

    Dim sql As String
    sql = "UPDATE CLIENT SET " & _
          "nom = '" & EchapperChaine(Trim(nom)) & "', " & _
          "prenom = '" & EchapperChaine(Trim(prenom)) & "', " & _
          "email = '" & EchapperChaine(Trim(email)) & "', " & _
          "telephone = '" & EchapperChaine(Trim(telephone)) & "', " & _
          "adresse = '" & EchapperChaine(Trim(adresse)) & "' " & _
          "WHERE id_client = " & idClient

    ExecuterSQL sql
    JournaliserModification "Client", idClient
    ModifierClient = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
    ModifierClient = False
End Function

' ============================================================
' Supprimer un client (vérifie absence de réservations)
' ============================================================
Public Sub SupprimerClient(idClient As Long)
    If Not VerifierAccesRole(ROLE_ADMINISTRATEUR) Then Exit Sub

    ' Vérifier les réservations liées
    Dim nbRes As Long
    nbRes = CompterEnregistrements("RESERVATION", "id_client = " & idClient)

    If nbRes > 0 Then
        AfficherErreur "Impossible de supprimer ce client : " & nbRes & _
                       " réservation(s) lui sont rattachée(s)." & vbCrLf & _
                       "Annulez d'abord toutes ses réservations."
        Exit Sub
    End If

    If Not DemanderConfirmation("Supprimer définitivement ce client ?") Then Exit Sub

    On Error GoTo GestionErreur
    ExecuterSQL "DELETE FROM CLIENT WHERE id_client = " & idClient
    JournaliserSuppression "Client", idClient
    AfficherSucces "Client supprimé avec succès."
    Exit Sub

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
End Sub

' ============================================================
' Rechercher des clients (retourne un Recordset)
' ============================================================
Public Function RechercherClients(termeRecherche As String) As DAO.Recordset
    Dim sql As String
    Dim terme As String
    terme = EchapperChaine(termeRecherche)

    sql = "SELECT C.id_client, C.nom, C.prenom, C.email, C.telephone, " & _
          "C.date_inscription, Count(R.id_reservation) AS nb_reservations " & _
          "FROM CLIENT AS C " & _
          "LEFT JOIN RESERVATION AS R ON C.id_client = R.id_client "

    If Len(Trim(terme)) > 0 Then
        sql = sql & "WHERE C.nom LIKE '*" & terme & "*' " & _
                    "OR C.prenom LIKE '*" & terme & "*' " & _
                    "OR C.email LIKE '*" & terme & "*' " & _
                    "OR C.telephone LIKE '*" & terme & "*' "
    End If

    sql = sql & "GROUP BY C.id_client, C.nom, C.prenom, C.email, C.telephone, C.date_inscription " & _
                "ORDER BY C.nom, C.prenom"

    Set RechercherClients = OuvrirRecordset(sql)
End Function

' ============================================================
' Obtenir l'historique des réservations d'un client
' ============================================================
Public Function HistoriqueClient(idClient As Long) As DAO.Recordset
    Dim sql As String
    sql = "SELECT R.id_reservation, R.date_debut, R.date_fin, " & _
          "R.montant_total, R.statut, R.date_creation " & _
          "FROM RESERVATION AS R " & _
          "WHERE R.id_client = " & idClient & _
          " ORDER BY R.date_creation DESC"

    Set HistoriqueClient = OuvrirRecordset(sql)
End Function

' ============================================================
' Obtenir le nombre de réservations d'un client
' ============================================================
Public Function NbReservationsClient(idClient As Long) As Long
    NbReservationsClient = CompterEnregistrements("RESERVATION", "id_client = " & idClient)
End Function
