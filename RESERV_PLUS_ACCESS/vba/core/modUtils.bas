Attribute VB_Name = "modUtils"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Fonctions utilitaires générales
' Module : modUtils.bas
' ============================================================

' ============================================================
' Formater un montant avec la devise de l'application
' ============================================================
Public Function FormatMontant(montant As Currency, Optional devise As String = "") As String
    If devise = "" Then
        devise = ObtenirParametre("DEVISE", "GNF")
    End If
    FormatMontant = Format(montant, "#,##0") & " " & devise
End Function

' ============================================================
' Formater une date en format lisible français
' ============================================================
Public Function FormatDate(dt As Variant) As String
    If IsNull(dt) Or IsEmpty(dt) Then
        FormatDate = ""
        Exit Function
    End If
    FormatDate = Format(dt, "dd/mm/yyyy")
End Function

' ============================================================
' Formater une date + heure
' ============================================================
Public Function FormatDateHeure(dt As Variant) As String
    If IsNull(dt) Or IsEmpty(dt) Then
        FormatDateHeure = ""
        Exit Function
    End If
    FormatDateHeure = Format(dt, "dd/mm/yyyy HH:mm")
End Function

' ============================================================
' Afficher un message d'erreur standard
' ============================================================
Public Sub AfficherErreur(message As String, Optional titre As String = "")
    If titre = "" Then titre = APP_NOM & " - Erreur"
    MsgBox message, vbCritical, titre
End Sub

' ============================================================
' Afficher un message de succès
' ============================================================
Public Sub AfficherSucces(message As String, Optional titre As String = "")
    If titre = "" Then titre = APP_NOM
    MsgBox message, vbInformation, titre
End Sub

' ============================================================
' Demander confirmation à l'utilisateur
' ============================================================
Public Function DemanderConfirmation(message As String, Optional titre As String = "") As Boolean
    If titre = "" Then titre = APP_NOM & " - Confirmation"
    DemanderConfirmation = (MsgBox(message, vbQuestion + vbYesNo, titre) = vbYes)
End Function

' ============================================================
' Générer une chaîne aléatoire pour les références
' ============================================================
Public Function GenererReference(longueur As Integer) As String
    Dim caracteres As String
    Dim reference As String
    Dim i As Integer

    caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    reference = ""

    Randomize
    For i = 1 To longueur
        reference = reference & Mid(caracteres, Int(Rnd() * Len(caracteres)) + 1, 1)
    Next i

    GenererReference = reference
End Function

' ============================================================
' Valider une adresse email (vérification basique)
' ============================================================
Public Function EstEmailValide(email As String) As Boolean
    EstEmailValide = (InStr(email, "@") > 1 And InStr(email, ".") > 2 And Len(email) > 5)
End Function

' ============================================================
' Convertir une valeur en chaîne sécurisée (évite Null)
' ============================================================
Public Function ChaineSure(valeur As Variant) As String
    ChaineSure = CStr(Nz(valeur, ""))
End Function

' ============================================================
' Convertir une valeur en entier sécurisé
' ============================================================
Public Function EntierSur(valeur As Variant, Optional defaut As Long = 0) As Long
    If IsNull(valeur) Or IsEmpty(valeur) Then
        EntierSur = defaut
    ElseIf IsNumeric(valeur) Then
        EntierSur = CLng(valeur)
    Else
        EntierSur = defaut
    End If
End Function

' ============================================================
' Vérifier qu'un champ obligatoire est rempli
' ============================================================
Public Function ChampObligatoire(valeur As Variant, nomChamp As String) As Boolean
    If IsNull(valeur) Or IsEmpty(valeur) Or Len(Trim(CStr(Nz(valeur, "")))) = 0 Then
        MsgBox "Le champ '" & nomChamp & "' est obligatoire.", vbExclamation, APP_NOM
        ChampObligatoire = False
    Else
        ChampObligatoire = True
    End If
End Function

' ============================================================
' Calculer la durée entre deux dates en jours (minimum 1)
' ============================================================
Public Function DureeEnJours(dateDebut As Date, dateFin As Date) As Integer
    Dim duree As Integer
    duree = DateDiff("d", dateDebut, dateFin)
    If duree <= 0 Then duree = 1
    DureeEnJours = duree
End Function

' ============================================================
' Obtenir la couleur RGB Access à partir de codes R,G,B
' ============================================================
Public Function CouleurRGB(r As Integer, g As Integer, b As Integer) As Long
    CouleurRGB = r + (g * 256) + (b * 65536)
End Function

' ============================================================
' Capitaliser la première lettre d'une chaîne
' ============================================================
Public Function Capitaliser(texte As String) As String
    If Len(texte) = 0 Then
        Capitaliser = ""
    Else
        Capitaliser = UCase(Left(texte, 1)) & LCase(Mid(texte, 2))
    End If
End Function

' ============================================================
' Tronquer un texte avec ellipsis si trop long
' ============================================================
Public Function Tronquer(texte As String, longueurMax As Integer) As String
    If Len(texte) > longueurMax Then
        Tronquer = Left(texte, longueurMax - 3) & "..."
    Else
        Tronquer = texte
    End If
End Function

' ============================================================
' Rafraîchir tous les formulaires ouverts
' ============================================================
Public Sub RafraichirFormulaireOuvert(nomFormulaire As String)
    On Error Resume Next
    If CurrentProject.AllForms(nomFormulaire).IsLoaded Then
        Forms(nomFormulaire).Requery
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Fermer proprement un formulaire s'il est ouvert
' ============================================================
Public Sub FermerFormulaire(nomFormulaire As String)
    On Error Resume Next
    If CurrentProject.AllForms(nomFormulaire).IsLoaded Then
        DoCmd.Close acForm, nomFormulaire, acSaveNo
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Convertir un numéro de mois en libellé français
' ============================================================
Public Function NomMois(numeroMois As Integer) As String
    Dim mois() As String
    mois = Array("", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", _
                 "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre")
    If numeroMois >= 1 And numeroMois <= 12 Then
        NomMois = mois(numeroMois)
    Else
        NomMois = ""
    End If
End Function
