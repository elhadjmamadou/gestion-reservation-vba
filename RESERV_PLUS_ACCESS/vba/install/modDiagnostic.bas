Attribute VB_Name = "modDiagnostic"
Option Compare Database
Option Explicit

' =============================================================================
' RESERV+ - Diagnostic et reparation
' Usage dans la fenetre Execution (CTRL+G) :
'   DiagnosticComplet        -> voir ce qui marche et ce qui echoue
'   ReparerFormulaires       -> retenter la creation des formulaires
'   CreerFormulaireSeul "F_CONNEXION"  -> creer un seul formulaire
' =============================================================================

Public Sub DiagnosticComplet()
    Dim msg As String
    msg = "=== DIAGNOSTIC RESERV+ ===" & vbCrLf & vbCrLf

    ' Tables
    Dim tables(13) As String
    tables(0) = "PRESTATAIRE" : tables(1) = "CATEGORIE"
    tables(2) = "RESSOURCE"   : tables(3) = "CHAMBRE_HOTEL"
    tables(4) = "TAXI"        : tables(5) = "VOL"
    tables(6) = "CLIENT"      : tables(7) = "UTILISATEUR"
    tables(8) = "RESERVATION" : tables(9) = "DETAIL_RESERVATION"
    tables(10) = "PAIEMENT"   : tables(11) = "BILLET"
    tables(12) = "JOURNAL_ACTION" : tables(13) = "PARAMETRE"

    Dim nbT As Integer : nbT = 0
    Dim i As Integer
    For i = 0 To 13
        If TableExiste(tables(i)) Then nbT = nbT + 1
    Next i
    msg = msg & "TABLES    : " & nbT & "/14 " & IIf(nbT = 14, "[OK]", "[INCOMPLET]") & vbCrLf

    ' Utilisateurs
    Dim nbU As Long
    nbU = CompterEnregistrements("UTILISATEUR", "")
    msg = msg & "UTILISATEURS : " & nbU & " " & IIf(nbU >= 3, "[OK]", "[VIDE - relancer InsererDonneesInitiales]") & vbCrLf

    ' Requetes
    Dim nbQ As Integer : nbQ = 0
    Dim requetes(11) As String
    requetes(0) = "RQ_RESERVATIONS_COMPLETES"
    requetes(1) = "RQ_RESERVATIONS_DU_JOUR"
    requetes(2) = "RQ_RESERVATIONS_EN_ATTENTE"
    requetes(3) = "RQ_TOTAL_ENCAISSE"
    requetes(4) = "RQ_SOLDE_PAR_RESERVATION"
    requetes(5) = "RQ_CLIENTS_AVEC_NB_RESERVATIONS"
    requetes(6) = "RQ_RESSOURCES_AVEC_PRESTATAIRE_CATEGORIE"
    requetes(7) = "RQ_PAIEMENTS_COMPLETS"
    requetes(8) = "RQ_CHIFFRE_AFFAIRES_PAR_CATEGORIE"
    requetes(9) = "RQ_CHIFFRE_AFFAIRES_PAR_PERIODE"
    requetes(10) = "RQ_JOURNAL_ACTIONS"
    requetes(11) = "RQ_BILLETS_COMPLETS"
    For i = 0 To 11
        If RequeteExiste(requetes(i)) Then nbQ = nbQ + 1
    Next i
    msg = msg & "REQUETES  : " & nbQ & "/12 " & IIf(nbQ = 12, "[OK]", "[INCOMPLET]") & vbCrLf

    ' Formulaires
    Dim formsListe(10) As String
    formsListe(0) = "F_CONNEXION"
    formsListe(1) = "F_TABLEAU_BORD"
    formsListe(2) = "F_CLIENTS"
    formsListe(3) = "F_PRESTATAIRES"
    formsListe(4) = "F_RESSOURCES"
    formsListe(5) = "F_RESERVATIONS"
    formsListe(6) = "F_NOUVELLE_RESERVATION"
    formsListe(7) = "F_PAIEMENTS"
    formsListe(8) = "F_UTILISATEURS"
    formsListe(9) = "F_PARAMETRES"
    formsListe(10) = "F_ETATS_RAPPORTS"

    Dim nbF As Integer : nbF = 0
    Dim manquants As String : manquants = ""
    For i = 0 To 10
        If FormulaireExiste(formsListe(i)) Then
            nbF = nbF + 1
        Else
            manquants = manquants & "  - " & formsListe(i) & vbCrLf
        End If
    Next i
    msg = msg & "FORMULAIRES : " & nbF & "/11 " & IIf(nbF = 11, "[OK]", "[INCOMPLET]") & vbCrLf
    If Len(manquants) > 0 Then
        msg = msg & "Manquants :" & vbCrLf & manquants
    End If

    msg = msg & vbCrLf
    If nbF < 11 Then
        msg = msg & "ACTION : Tapez  ReparerFormulaires  dans la fenetre Execution"
    Else
        msg = msg & "Tout est complet ! Ouvrez F_CONNEXION."
    End If

    MsgBox msg, vbInformation, "RESERV+ Diagnostic"
    Debug.Print msg
End Sub

' =============================================================================
' Retenter la creation de tous les formulaires manquants
' =============================================================================
Public Sub ReparerFormulaires()
    ' Fermer tous les formulaires ouverts d'abord
    FermerTousLesFormulaires

    ' Compiler le projet avant de creer les formulaires
    On Error Resume Next
    Application.VBE.ActiveVBProject.VBE.CommandBars("Visual Basic").Controls("Debug").Controls("Compile").Execute
    On Error GoTo 0

    MsgBox "Creation des formulaires en cours..." & vbCrLf & _
           "Des fenetres vont s'ouvrir puis se fermer automatiquement." & vbCrLf & vbCrLf & _
           "Ne cliquez sur rien pendant l'operation.", vbInformation, "RESERV+"

    CreerFormulaireSeul "F_CONNEXION"
    CreerFormulaireSeul "F_TABLEAU_BORD"
    CreerFormulaireSeul "F_CLIENTS"
    CreerFormulaireSeul "F_PRESTATAIRES"
    CreerFormulaireSeul "F_RESSOURCES"
    CreerFormulaireSeul "F_RESERVATIONS"
    CreerFormulaireSeul "F_NOUVELLE_RESERVATION"
    CreerFormulaireSeul "F_PAIEMENTS"
    CreerFormulaireSeul "F_UTILISATEURS"
    CreerFormulaireSeul "F_PARAMETRES"
    CreerFormulaireSeul "F_ETATS_RAPPORTS"

    DiagnosticComplet
End Sub

' =============================================================================
' Creer un seul formulaire avec rapport d'erreur detaille
' =============================================================================
Public Sub CreerFormulaireSeul(nomForme As String)
    On Error GoTo ErrCreation

    Select Case nomForme
        Case "F_CONNEXION"          : Creer_F_CONNEXION
        Case "F_TABLEAU_BORD"       : Creer_F_TABLEAU_BORD
        Case "F_CLIENTS"            : Creer_F_CLIENTS
        Case "F_PRESTATAIRES"       : Creer_F_PRESTATAIRES
        Case "F_RESSOURCES"         : Creer_F_RESSOURCES
        Case "F_RESERVATIONS"       : Creer_F_RESERVATIONS
        Case "F_NOUVELLE_RESERVATION" : Creer_F_NOUVELLE_RESERVATION
        Case "F_PAIEMENTS"          : Creer_F_PAIEMENTS
        Case "F_UTILISATEURS"       : Creer_F_UTILISATEURS
        Case "F_PARAMETRES"         : Creer_F_PARAMETRES
        Case "F_ETATS_RAPPORTS"     : Creer_F_ETATS_RAPPORTS
        Case Else
            Debug.Print "Formulaire inconnu : " & nomForme
    End Select

    Debug.Print "[OK] " & nomForme & " cree."
    Exit Sub

ErrCreation:
    Debug.Print "[ERREUR] " & nomForme & " -> " & Err.Number & " : " & Err.Description
    MsgBox "Erreur creation " & nomForme & vbCrLf & vbCrLf & _
           "Numero : " & Err.Number & vbCrLf & _
           "Detail : " & Err.Description & vbCrLf & vbCrLf & _
           "Notez ce message et continuez.", vbExclamation, "RESERV+"
    Resume Next
End Sub

' =============================================================================
' Fermer tous les formulaires ouverts
' =============================================================================
Public Sub FermerTousLesFormulaires()
    Dim frm As AccessObject
    For Each frm In CurrentProject.AllForms
        If frm.IsLoaded Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    Next frm
    Debug.Print "Tous les formulaires fermes."
End Sub

' =============================================================================
' Compiler le projet VBA (force la detection des erreurs de syntaxe)
' =============================================================================
Public Sub CompilerProjet()
    On Error GoTo ErrCompile
    Application.VBE.ActiveVBProject.VBE.MainWindow.Visible = True
    ' La compilation se fait via le menu Debug > Compile dans l'editeur VBA
    ' Ouvrir ALT+F11 puis : Debug > Compile RESERV_PLUS
    MsgBox "Dans l'editeur VBA (ALT+F11) :" & vbCrLf & _
           "Menu  Debug  >  Compile RESERV_PLUS" & vbCrLf & vbCrLf & _
           "Si aucune erreur n'apparait, revenez ici et tapez :" & vbCrLf & _
           "ReparerFormulaires", vbInformation, "RESERV+"
    Exit Sub
ErrCompile:
    Debug.Print "Info compilation : " & Err.Description
End Sub

' =============================================================================
' Reinsertion des donnees initiales (si tables vides)
' =============================================================================
Public Sub ReinsererDonnees()
    Dim db As DAO.Database
    Set db = CurrentDb()

    If CompterEnregistrements("CATEGORIE", "") = 0 Then
        db.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Hotel', 'Chambres et hebergements hoteliers')"
        db.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Taxi', 'Vehicules de transport terrestre')"
        db.Execute "INSERT INTO CATEGORIE (libelle_categorie, description) VALUES ('Vol', 'Vols et transport aerien')"
        Debug.Print "Categories inserees."
    End If

    If CompterEnregistrements("UTILISATEUR", "") = 0 Then
        db.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) VALUES ('admin', 'admin123', 'Administrateur', 'Administrateur', True, Now())"
        db.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) VALUES ('agent', 'agent123', 'Agent de Reservation', 'Agent', True, Now())"
        db.Execute "INSERT INTO UTILISATEUR (login, mot_de_passe, nom_complet, role, actif, date_creation) VALUES ('utilisateur', 'user123', 'Utilisateur Standard', 'Utilisateur', True, Now())"
        Debug.Print "Utilisateurs inseres."
    End If

    If CompterEnregistrements("PARAMETRE", "") = 0 Then
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre) VALUES ('NOM_APPLICATION', 'RESERV+')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre) VALUES ('DEVISE', 'GNF')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre) VALUES ('PREFIXE_BILLET', 'BIL')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre) VALUES ('NOM_ETABLISSEMENT', 'RESERV+')"
        db.Execute "INSERT INTO PARAMETRE (cle_parametre, valeur_parametre) VALUES ('VERSION_APP', '1.0.0')"
        Debug.Print "Parametres inseres."
    End If

    MsgBox "Donnees initiales verifiees/inserees." & vbCrLf & _
           "Utilisateurs : " & CompterEnregistrements("UTILISATEUR", "") & vbCrLf & _
           "Categories : " & CompterEnregistrements("CATEGORIE", ""), vbInformation, "RESERV+"
End Sub
