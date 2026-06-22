Attribute VB_Name = "modPrestataires"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Gestion des prestataires
' Module : modPrestataires.bas
' ============================================================

Public Function CreerPrestataire(nomEntreprise As String, nomResponsable As String, _
                                  email As String, telephone As String, _
                                  adresse As String, typeActivite As String) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idPrestataire As Long

    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        CreerPrestataire = 0
        Exit Function
    End If

    If Not ChampObligatoire(nomEntreprise, "Nom de l'entreprise") Then
        CreerPrestataire = 0
        Exit Function
    End If

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("PRESTATAIRE", dbOpenDynaset)

    rs.AddNew
    rs!nom_entreprise = Trim(nomEntreprise)
    rs!nom_responsable = Trim(nomResponsable)
    rs!email = Trim(email)
    rs!telephone = Trim(telephone)
    rs!adresse = Trim(adresse)
    rs!type_activite = Trim(typeActivite)
    rs!date_adhesion = Now()
    rs.Update

    idPrestataire = DernierID("PRESTATAIRE", "id_prestataire")
    rs.Close
    Set rs = Nothing

    JournaliserCreation "Prestataire", idPrestataire
    CreerPrestataire = idPrestataire
    Exit Function

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
    CreerPrestataire = 0
    If Not rs Is Nothing Then rs.Close
End Function

Public Function ModifierPrestataire(idPrestataire As Long, nomEntreprise As String, _
                                     nomResponsable As String, email As String, _
                                     telephone As String, adresse As String, _
                                     typeActivite As String) As Boolean
    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        ModifierPrestataire = False
        Exit Function
    End If

    If Not ChampObligatoire(nomEntreprise, "Nom de l'entreprise") Then
        ModifierPrestataire = False
        Exit Function
    End If

    Dim sql As String
    sql = "UPDATE PRESTATAIRE SET " & _
          "nom_entreprise = '" & EchapperChaine(nomEntreprise) & "', " & _
          "nom_responsable = '" & EchapperChaine(nomResponsable) & "', " & _
          "email = '" & EchapperChaine(email) & "', " & _
          "telephone = '" & EchapperChaine(telephone) & "', " & _
          "adresse = '" & EchapperChaine(adresse) & "', " & _
          "type_activite = '" & EchapperChaine(typeActivite) & "' " & _
          "WHERE id_prestataire = " & idPrestataire

    ExecuterSQL sql
    JournaliserModification "Prestataire", idPrestataire
    ModifierPrestataire = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
    ModifierPrestataire = False
End Function

Public Sub SupprimerPrestataire(idPrestataire As Long)
    If Not VerifierAccesRole(ROLE_ADMINISTRATEUR) Then Exit Sub

    Dim nbRes As Long
    nbRes = CompterEnregistrements("RESSOURCE", "id_prestataire = " & idPrestataire)

    If nbRes > 0 Then
        AfficherErreur "Impossible de supprimer : " & nbRes & " ressource(s) rattachée(s) à ce prestataire."
        Exit Sub
    End If

    If Not DemanderConfirmation("Supprimer ce prestataire ?") Then Exit Sub

    On Error GoTo GestionErreur
    ExecuterSQL "DELETE FROM PRESTATAIRE WHERE id_prestataire = " & idPrestataire
    JournaliserSuppression "Prestataire", idPrestataire
    AfficherSucces "Prestataire supprimé."
    Exit Sub

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
End Sub

Public Function NbRessourcesPrestataire(idPrestataire As Long) As Long
    NbRessourcesPrestataire = CompterEnregistrements("RESSOURCE", "id_prestataire = " & idPrestataire)
End Function

Public Function RechercherPrestataires(terme As String) As DAO.Recordset
    Dim sql As String
    Dim t As String
    t = EchapperChaine(terme)

    sql = "SELECT P.id_prestataire, P.nom_entreprise, P.nom_responsable, " & _
          "P.email, P.telephone, P.type_activite, P.date_adhesion, " & _
          "Count(R.id_ressource) AS nb_ressources " & _
          "FROM PRESTATAIRE AS P " & _
          "LEFT JOIN RESSOURCE AS R ON P.id_prestataire = R.id_prestataire "

    If Len(Trim(t)) > 0 Then
        sql = sql & "WHERE P.nom_entreprise LIKE '*" & t & "*' " & _
                    "OR P.type_activite LIKE '*" & t & "*' "
    End If

    sql = sql & "GROUP BY P.id_prestataire, P.nom_entreprise, P.nom_responsable, " & _
                "P.email, P.telephone, P.type_activite, P.date_adhesion " & _
                "ORDER BY P.nom_entreprise"

    Set RechercherPrestataires = OuvrirRecordset(sql)
End Function
