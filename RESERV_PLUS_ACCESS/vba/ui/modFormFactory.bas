Attribute VB_Name = "modFormFactory"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Fabrique de formulaires (orchestrateur)
' Module : modFormFactory.bas
' ============================================================

' ============================================================
' Créer TOUS les formulaires de l'application
' ============================================================
Public Sub CreerTousLesFormulaires()
    Dim total As Integer
    Dim crees As Integer

    total = 9
    crees = 0

    Debug.Print "=== CRÉATION DES FORMULAIRES RESERV+ ==="

    On Error GoTo ErreurFormulaire

    ' 1. Connexion
    Debug.Print "1/" & total & " F_CONNEXION..."
    Creer_F_CONNEXION
    crees = crees + 1

    ' 2. Tableau de bord
    Debug.Print "2/" & total & " F_TABLEAU_BORD..."
    Creer_F_TABLEAU_BORD
    crees = crees + 1

    ' 3. Clients
    Debug.Print "3/" & total & " F_CLIENTS..."
    Creer_F_CLIENTS
    crees = crees + 1

    ' 4. Prestataires
    Debug.Print "4/" & total & " F_PRESTATAIRES..."
    Creer_F_PRESTATAIRES
    crees = crees + 1

    ' 5. Ressources
    Debug.Print "5/" & total & " F_RESSOURCES..."
    Creer_F_RESSOURCES
    crees = crees + 1

    ' 6. Réservations
    Debug.Print "6/" & total & " F_RESERVATIONS..."
    Creer_F_RESERVATIONS
    crees = crees + 1

    ' 7. Nouvelle réservation
    Debug.Print "7/" & total & " F_NOUVELLE_RESERVATION..."
    Creer_F_NOUVELLE_RESERVATION
    crees = crees + 1

    ' 8. Paiements
    Debug.Print "8/" & total & " F_PAIEMENTS..."
    Creer_F_PAIEMENTS
    crees = crees + 1

    ' 9. Administration
    Debug.Print "9/" & total & " F_UTILISATEURS..."
    Creer_F_UTILISATEURS
    crees = crees + 1

    Debug.Print "10/10 F_PARAMETRES..."
    Creer_F_PARAMETRES

    Debug.Print "11/11 F_ETATS_RAPPORTS..."
    Creer_F_ETATS_RAPPORTS

    Debug.Print "=== " & crees & " formulaires créés avec succès ==="
    Exit Sub

ErreurFormulaire:
    Debug.Print "ERREUR création formulaire : " & Err.Description
    MsgBox "Erreur lors de la création d'un formulaire :" & vbCrLf & Err.Description & vbCrLf & _
           "Continuez avec les formulaires suivants.", vbExclamation, APP_NOM
    Resume Next
End Sub

' ============================================================
' Supprimer tous les formulaires générés
' ============================================================
Public Sub SupprimerTousLesFormulaires()
    Dim formulaires() As String
    formulaires = Array(FORM_CONNEXION, FORM_DASHBOARD, FORM_CLIENTS, _
                        FORM_PRESTATAIRES, FORM_RESSOURCES, _
                        FORM_RESERVATIONS, FORM_NOUVELLE_RESERVATION, _
                        FORM_DETAIL_RESERVATION, FORM_PAIEMENTS, _
                        FORM_BILLETS, FORM_UTILISATEURS, _
                        FORM_PARAMETRES, FORM_RAPPORTS, _
                        FORM_CHAMBRE, FORM_TAXI, FORM_VOL)

    Dim i As Integer
    For i = 0 To UBound(formulaires)
        SupprimerObjetAccess acForm, formulaires(i)
        Debug.Print "Formulaire supprimé : " & formulaires(i)
    Next i
End Sub
