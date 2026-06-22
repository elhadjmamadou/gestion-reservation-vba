Attribute VB_Name = "modRessources"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Gestion des ressources (hôtels, taxis, vols)
' Module : modRessources.bas
' ============================================================

' ============================================================
' Créer une ressource de base
' Retourne l'ID créé (0 si échec)
' ============================================================
Public Function CreerRessource(idPrestataire As Long, idCategorie As Long, _
                                libelle As String, description As String, _
                                prixUnitaire As Currency, capacite As Integer) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idRessource As Long

    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        CreerRessource = 0
        Exit Function
    End If

    If Not ChampObligatoire(libelle, "Libellé de la ressource") Then
        CreerRessource = 0
        Exit Function
    End If

    If idPrestataire <= 0 Or idCategorie <= 0 Then
        AfficherErreur "Veuillez sélectionner un prestataire et une catégorie."
        CreerRessource = 0
        Exit Function
    End If

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("RESSOURCE", dbOpenDynaset)

    rs.AddNew
    rs!id_prestataire = idPrestataire
    rs!id_categorie = idCategorie
    rs!libelle = Trim(libelle)
    rs!description = Trim(description)
    rs!prix_unitaire = prixUnitaire
    rs!capacite = IIf(capacite <= 0, 1, capacite)
    rs!disponible = True
    rs!statut_ressource = "Disponible"
    rs!date_creation = Now()
    rs.Update

    idRessource = DernierID("RESSOURCE", "id_ressource")
    rs.Close
    Set rs = Nothing

    JournaliserCreation "Ressource", idRessource
    CreerRessource = idRessource
    Exit Function

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
    CreerRessource = 0
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Ajouter les détails spécifiques d'une chambre d'hôtel
' ============================================================
Public Function CreerDetailsHotel(idRessource As Long, nbEtoiles As Integer, _
                                   typeChambre As String, nbLits As Integer, _
                                   petitDejeuner As Boolean, vue As String) As Boolean
    On Error GoTo GestionErreur

    ' Supprimer si déjà existant
    ExecuterSQL "DELETE FROM CHAMBRE_HOTEL WHERE id_ressource = " & idRessource

    Dim sql As String
    sql = "INSERT INTO CHAMBRE_HOTEL (id_ressource, nb_etoiles, type_chambre, nb_lits, petit_dejeuner, vue) " & _
          "VALUES (" & idRessource & ", " & nbEtoiles & ", '" & EchapperChaine(typeChambre) & "', " & _
          nbLits & ", " & IIf(petitDejeuner, "True", "False") & ", '" & EchapperChaine(vue) & "')"

    ExecuterSQL sql
    CreerDetailsHotel = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur détails hôtel : " & Err.Description
    CreerDetailsHotel = False
End Function

' ============================================================
' Ajouter les détails spécifiques d'un taxi
' ============================================================
Public Function CreerDetailsTaxi(idRessource As Long, marque As String, modele As String, _
                                  immatriculation As String, nbPlaces As Integer, _
                                  climatisation As Boolean) As Boolean
    On Error GoTo GestionErreur

    If Len(Trim(immatriculation)) = 0 Then
        AfficherErreur "L'immatriculation du taxi est obligatoire."
        CreerDetailsTaxi = False
        Exit Function
    End If

    ' Vérifier unicité immatriculation
    If EnregistrementExiste("TAXI", "immatriculation = '" & EchapperChaine(immatriculation) & "' AND id_ressource <> " & idRessource) Then
        AfficherErreur "Cette immatriculation est déjà utilisée par un autre taxi."
        CreerDetailsTaxi = False
        Exit Function
    End If

    ExecuterSQL "DELETE FROM TAXI WHERE id_ressource = " & idRessource

    Dim sql As String
    sql = "INSERT INTO TAXI (id_ressource, marque, modele, immatriculation, nb_places, climatisation) " & _
          "VALUES (" & idRessource & ", '" & EchapperChaine(marque) & "', '" & EchapperChaine(modele) & "', " & _
          "'" & EchapperChaine(immatriculation) & "', " & nbPlaces & ", " & IIf(climatisation, "True", "False") & ")"

    ExecuterSQL sql
    CreerDetailsTaxi = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur détails taxi : " & Err.Description
    CreerDetailsTaxi = False
End Function

' ============================================================
' Ajouter les détails spécifiques d'un vol
' ============================================================
Public Function CreerDetailsVol(idRessource As Long, numVol As String, _
                                 villeDepart As String, villeArrivee As String, _
                                 dateDepart As Date, heureDepart As Date, _
                                 dateArrivee As Date, heureArrivee As Date, _
                                 classe As String) As Boolean
    On Error GoTo GestionErreur

    If Len(Trim(numVol)) = 0 Then
        AfficherErreur "Le numéro de vol est obligatoire."
        CreerDetailsVol = False
        Exit Function
    End If

    If EnregistrementExiste("VOL", "num_vol = '" & EchapperChaine(numVol) & "' AND id_ressource <> " & idRessource) Then
        AfficherErreur "Ce numéro de vol existe déjà."
        CreerDetailsVol = False
        Exit Function
    End If

    ExecuterSQL "DELETE FROM VOL WHERE id_ressource = " & idRessource

    Dim sql As String
    sql = "INSERT INTO VOL (id_ressource, num_vol, ville_depart, ville_arrivee, " & _
          "date_depart, heure_depart, date_arrivee, heure_arrivee, classe) VALUES (" & _
          idRessource & ", '" & EchapperChaine(numVol) & "', " & _
          "'" & EchapperChaine(villeDepart) & "', '" & EchapperChaine(villeArrivee) & "', " & _
          "#" & Format(dateDepart, "yyyy/mm/dd") & "#, " & _
          "#" & Format(heureDepart, "yyyy/mm/dd HH:nn:ss") & "#, " & _
          "#" & Format(dateArrivee, "yyyy/mm/dd") & "#, " & _
          "#" & Format(heureArrivee, "yyyy/mm/dd HH:nn:ss") & "#, " & _
          "'" & EchapperChaine(classe) & "')"

    ExecuterSQL sql
    CreerDetailsVol = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur détails vol : " & Err.Description
    CreerDetailsVol = False
End Function

' ============================================================
' Basculer la disponibilité d'une ressource
' ============================================================
Public Sub BasculerDisponibilite(idRessource As Long)
    If Not VerifierAccesRole(ROLE_AGENT) Then Exit Sub

    Dim disponible As Boolean
    disponible = CBool(Nz(ObtenirValeur("SELECT disponible FROM RESSOURCE WHERE id_ressource = " & idRessource), False))

    Dim nouveauStatut As Boolean
    Dim nouveauTexte As String
    nouveauStatut = Not disponible
    nouveauTexte = IIf(nouveauStatut, "Disponible", "Indisponible")

    ExecuterSQL "UPDATE RESSOURCE SET disponible = " & IIf(nouveauStatut, "True", "False") & _
                ", statut_ressource = '" & nouveauTexte & "' " & _
                "WHERE id_ressource = " & idRessource

    Journaliser "STATUT_RESSOURCE", "Ressource #" & idRessource & " → " & nouveauTexte
    AfficherSucces "Ressource marquée comme : " & nouveauTexte
End Sub

' ============================================================
' Supprimer une ressource (vérifie absence de réservations actives)
' ============================================================
Public Sub SupprimerRessource(idRessource As Long)
    If Not VerifierAccesRole(ROLE_ADMINISTRATEUR) Then Exit Sub

    ' Vérifier les réservations actives
    Dim nbResActives As Long
    nbResActives = CLng(Nz(ObtenirValeur( _
        "SELECT Count(*) FROM DETAIL_RESERVATION AS DR " & _
        "INNER JOIN RESERVATION AS R ON DR.id_reservation = R.id_reservation " & _
        "WHERE DR.id_ressource = " & idRessource & _
        " AND R.statut <> '" & STATUT_ANNULEE & "'"), 0))

    If nbResActives > 0 Then
        AfficherErreur "Impossible de supprimer : " & nbResActives & " réservation(s) active(s) concernent cette ressource."
        Exit Sub
    End If

    If Not DemanderConfirmation("Supprimer définitivement cette ressource et ses détails spécifiques ?") Then Exit Sub

    On Error GoTo GestionErreur
    ExecuterSQL "DELETE FROM CHAMBRE_HOTEL WHERE id_ressource = " & idRessource
    ExecuterSQL "DELETE FROM TAXI WHERE id_ressource = " & idRessource
    ExecuterSQL "DELETE FROM VOL WHERE id_ressource = " & idRessource
    ExecuterSQL "DELETE FROM RESSOURCE WHERE id_ressource = " & idRessource
    JournaliserSuppression "Ressource", idRessource
    AfficherSucces "Ressource supprimée avec succès."
    Exit Sub

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
End Sub

' ============================================================
' Obtenir la catégorie d'une ressource (libellé)
' ============================================================
Public Function CategorieRessource(idRessource As Long) As String
    Dim sql As String
    sql = "SELECT CAT.libelle_categorie FROM RESSOURCE AS RS " & _
          "INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie " & _
          "WHERE RS.id_ressource = " & idRessource

    CategorieRessource = ChaineSure(ObtenirValeur(sql))
End Function

' ============================================================
' Rechercher des ressources avec filtres
' ============================================================
Public Function RechercherRessources(idCategorie As Long, _
                                      termeRecherche As String, _
                                      seulementDisponibles As Boolean) As DAO.Recordset
    Dim sql As String
    Dim conditions() As String
    Dim nbCond As Integer
    Dim cond As String
    nbCond = 0

    sql = "SELECT RS.id_ressource, CAT.libelle_categorie AS categorie, " & _
          "P.nom_entreprise AS prestataire, RS.libelle, " & _
          "RS.prix_unitaire, RS.capacite, RS.disponible, RS.statut_ressource " & _
          "FROM (RESSOURCE AS RS " & _
          "INNER JOIN CATEGORIE AS CAT ON RS.id_categorie = CAT.id_categorie) " & _
          "INNER JOIN PRESTATAIRE AS P ON RS.id_prestataire = P.id_prestataire WHERE 1=1"

    If idCategorie > 0 Then
        sql = sql & " AND RS.id_categorie = " & idCategorie
    End If

    If Len(Trim(termeRecherche)) > 0 Then
        Dim t As String
        t = EchapperChaine(termeRecherche)
        sql = sql & " AND (RS.libelle LIKE '*" & t & "*' OR P.nom_entreprise LIKE '*" & t & "*')"
    End If

    If seulementDisponibles Then
        sql = sql & " AND RS.disponible = True"
    End If

    sql = sql & " ORDER BY CAT.libelle_categorie, RS.libelle"
    Set RechercherRessources = OuvrirRecordset(sql)
End Function
