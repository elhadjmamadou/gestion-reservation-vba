Attribute VB_Name = "modConfig"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Configuration générale de l'application
' Module : modConfig.bas
' ============================================================

' ---- Informations de l'application ----
Public Const APP_NOM As String = "RESERV+"
Public Const APP_VERSION As String = "1.0.0"
Public Const APP_DESCRIPTION As String = "Plateforme de réservation multi-prestataires"

' ---- Rôles utilisateurs ----
Public Const ROLE_ADMINISTRATEUR As String = "Administrateur"
Public Const ROLE_AGENT As String = "Agent"
Public Const ROLE_UTILISATEUR As String = "Utilisateur"

' ---- Statuts de réservation ----
Public Const STATUT_EN_ATTENTE As String = "En attente"
Public Const STATUT_CONFIRMEE As String = "Confirmée"
Public Const STATUT_ANNULEE As String = "Annulée"

' ---- Statuts de billet ----
Public Const STATUT_BILLET_EMIS As String = "Émis"
Public Const STATUT_BILLET_ANNULE As String = "Annulé"
Public Const STATUT_BILLET_UTILISE As String = "Utilisé"

' ---- États de paiement ----
Public Const PAIEMENT_IMPAYE As String = "Impayé"
Public Const PAIEMENT_PARTIEL As String = "Partiel"
Public Const PAIEMENT_SOLDE As String = "Soldé"

' ---- Modes de paiement ----
Public Const MODE_ESPECES As String = "Espèces"
Public Const MODE_CARTE As String = "Carte"
Public Const MODE_VIREMENT As String = "Virement"
Public Const MODE_MOBILE As String = "Mobile Money"

' ---- Catégories de ressources ----
Public Const CAT_HOTEL As String = "Hôtel"
Public Const CAT_TAXI As String = "Taxi"
Public Const CAT_VOL As String = "Vol"

' ---- Noms des formulaires ----
Public Const FORM_CONNEXION As String = "F_CONNEXION"
Public Const FORM_DASHBOARD As String = "F_TABLEAU_BORD"
Public Const FORM_CLIENTS As String = "F_CLIENTS"
Public Const FORM_PRESTATAIRES As String = "F_PRESTATAIRES"
Public Const FORM_RESSOURCES As String = "F_RESSOURCES"
Public Const FORM_CHAMBRE As String = "F_CHAMBRE_HOTEL"
Public Const FORM_TAXI As String = "F_TAXI"
Public Const FORM_VOL As String = "F_VOL"
Public Const FORM_RESERVATIONS As String = "F_RESERVATIONS"
Public Const FORM_NOUVELLE_RESERVATION As String = "F_NOUVELLE_RESERVATION"
Public Const FORM_DETAIL_RESERVATION As String = "F_DETAIL_RESERVATION"
Public Const FORM_PAIEMENTS As String = "F_PAIEMENTS"
Public Const FORM_BILLETS As String = "F_BILLETS"
Public Const FORM_UTILISATEURS As String = "F_UTILISATEURS"
Public Const FORM_PARAMETRES As String = "F_PARAMETRES"
Public Const FORM_RAPPORTS As String = "F_ETATS_RAPPORTS"

' ---- Noms des requêtes enregistrées ----
Public Const RQ_RESERVATIONS_COMPLETES As String = "RQ_RESERVATIONS_COMPLETES"
Public Const RQ_RESERVATIONS_DU_JOUR As String = "RQ_RESERVATIONS_DU_JOUR"
Public Const RQ_RESERVATIONS_EN_ATTENTE As String = "RQ_RESERVATIONS_EN_ATTENTE"
Public Const RQ_TOTAL_ENCAISSE As String = "RQ_TOTAL_ENCAISSE"
Public Const RQ_SOLDE_PAR_RESERVATION As String = "RQ_SOLDE_PAR_RESERVATION"
Public Const RQ_CLIENTS_NB_RESERVATIONS As String = "RQ_CLIENTS_AVEC_NB_RESERVATIONS"
Public Const RQ_RESSOURCES_PRESTATAIRE As String = "RQ_RESSOURCES_AVEC_PRESTATAIRE_CATEGORIE"
Public Const RQ_PAIEMENTS_COMPLETS As String = "RQ_PAIEMENTS_COMPLETS"
Public Const RQ_CA_PAR_CATEGORIE As String = "RQ_CHIFFRE_AFFAIRES_PAR_CATEGORIE"
Public Const RQ_CA_PAR_PERIODE As String = "RQ_CHIFFRE_AFFAIRES_PAR_PERIODE"
Public Const RQ_JOURNAL_ACTIONS As String = "RQ_JOURNAL_ACTIONS"
Public Const RQ_BILLETS_COMPLETS As String = "RQ_BILLETS_COMPLETS"

' ============================================================
' Obtenir la valeur d'un paramètre depuis la table PARAMETRE
' ============================================================
Public Function ObtenirParametre(cleParam As String, Optional valeurDefaut As String = "") As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo GestionErreur

    Set db = CurrentDb()
    sql = "SELECT valeur_parametre FROM PARAMETRE WHERE cle_parametre = '" & cleParam & "'"
    Set rs = db.OpenRecordset(sql)

    If Not rs.EOF Then
        ObtenirParametre = Nz(rs!valeur_parametre, valeurDefaut)
    Else
        ObtenirParametre = valeurDefaut
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

GestionErreur:
    ObtenirParametre = valeurDefaut
End Function

' ============================================================
' Enregistrer ou mettre à jour un paramètre
' ============================================================
Public Sub EnregistrerParametre(cleParam As String, valeur As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo GestionErreur

    Set db = CurrentDb()
    sql = "SELECT * FROM PARAMETRE WHERE cle_parametre = '" & cleParam & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!cle_parametre = cleParam
        rs!valeur_parametre = valeur
        rs.Update
    Else
        rs.Edit
        rs!valeur_parametre = valeur
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub

GestionErreur:
    If Not rs Is Nothing Then rs.Close
    MsgBox "Erreur lors de l'enregistrement du paramètre : " & Err.Description, vbCritical, APP_NOM
End Sub
