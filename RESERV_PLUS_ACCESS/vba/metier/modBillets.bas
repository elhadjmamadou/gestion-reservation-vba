Attribute VB_Name = "modBillets"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Gestion des billets de réservation
' Module : modBillets.bas
' ============================================================

' ============================================================
' Générer un numéro de billet unique
' Format : PREFIXE-YYYYMMDD-000001
' Le préfixe est lu depuis la table PARAMETRE
' ============================================================
Public Function GenererNumeroBillet() As String
    Dim prefixe As String
    Dim dateStr As String
    Dim sequence As Long
    Dim numeroBillet As String

    ' Récupérer le préfixe depuis les paramètres
    prefixe = ObtenirParametre("PREFIXE_BILLET", "BIL")
    dateStr = Format(Now(), "yyyymmdd")

    ' Calculer le prochain numéro séquentiel du jour
    Dim sql As String
    sql = "SELECT Count(*) FROM BILLET WHERE Left(numero_billet, " & _
          (Len(prefixe) + 1 + 8) & ") = '" & prefixe & "-" & dateStr & "'"
    sequence = CLng(Nz(ObtenirValeur(sql, 0), 0)) + 1

    ' Formater : BIL-20260622-000001
    numeroBillet = prefixe & "-" & dateStr & "-" & Format(sequence, "000000")

    ' Vérifier l'unicité (sécurité)
    Do While EnregistrementExiste("BILLET", "numero_billet = '" & numeroBillet & "'")
        sequence = sequence + 1
        numeroBillet = prefixe & "-" & dateStr & "-" & Format(sequence, "000000")
    Loop

    GenererNumeroBillet = numeroBillet
End Function

' ============================================================
' Générer et enregistrer un billet pour une réservation
' Retourne l'ID du billet créé (0 si échec)
' ============================================================
Public Function GenererBillet(idReservation As Long) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim numeroBillet As String
    Dim idBillet As Long

    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        GenererBillet = 0
        Exit Function
    End If

    ' Vérifier que la réservation existe et n'est pas annulée
    Dim statut As String
    statut = ChaineSure(ObtenirValeur("SELECT statut FROM RESERVATION WHERE id_reservation = " & idReservation))

    If Len(statut) = 0 Then
        AfficherErreur "Réservation #" & idReservation & " introuvable."
        GenererBillet = 0
        Exit Function
    End If

    If statut = STATUT_ANNULEE Then
        AfficherErreur "Impossible de générer un billet pour une réservation annulée."
        GenererBillet = 0
        Exit Function
    End If

    ' Vérifier si un billet actif existe déjà
    Dim billetExistant As Long
    billetExistant = CLng(Nz(ObtenirValeur("SELECT id_billet FROM BILLET WHERE id_reservation = " & _
                                            idReservation & " AND statut_billet = '" & STATUT_BILLET_EMIS & "'"), 0))

    If billetExistant > 0 Then
        If Not DemanderConfirmation("Un billet (ID=" & billetExistant & ") existe déjà pour cette réservation." & vbCrLf & _
                                     "Générer un nouveau billet quand même ?") Then
            GenererBillet = billetExistant
            Exit Function
        End If
        ' Annuler l'ancien billet
        ExecuterSQL "UPDATE BILLET SET statut_billet = '" & STATUT_BILLET_ANNULE & "' " & _
                    "WHERE id_billet = " & billetExistant
    End If

    ' Générer le numéro
    numeroBillet = GenererNumeroBillet()

    ' Insérer le billet
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("BILLET", dbOpenDynaset)

    rs.AddNew
    rs!id_reservation = idReservation
    rs!numero_billet = numeroBillet
    rs!date_emission = Now()
    rs!statut_billet = STATUT_BILLET_EMIS
    rs.Update

    idBillet = DernierID("BILLET", "id_billet")
    rs.Close
    Set rs = Nothing

    Journaliser "GENERATION_BILLET", "Billet " & numeroBillet & " généré pour réservation #" & idReservation

    AfficherSucces "Billet généré avec succès !" & vbCrLf & _
                   "Numéro : " & numeroBillet, APP_NOM & " - Billet"

    GenererBillet = idBillet
    Exit Function

GestionErreur:
    AfficherErreur "Erreur lors de la génération du billet : " & Err.Description
    GenererBillet = 0
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Annuler un billet (le marquer comme annulé)
' ============================================================
Public Function AnnulerBillet(idBillet As Long) As Boolean
    On Error GoTo GestionErreur

    If Not VerifierAccesRole(ROLE_AGENT) Then
        AnnulerBillet = False
        Exit Function
    End If

    If Not DemanderConfirmation("Annuler ce billet ?") Then
        AnnulerBillet = False
        Exit Function
    End If

    ExecuterSQL "UPDATE BILLET SET statut_billet = '" & STATUT_BILLET_ANNULE & "' " & _
                "WHERE id_billet = " & idBillet

    Journaliser "ANNULATION_BILLET", "Billet ID=" & idBillet & " annulé par " & g_Login
    AnnulerBillet = True
    Exit Function

GestionErreur:
    AfficherErreur "Erreur : " & Err.Description
    AnnulerBillet = False
End Function

' ============================================================
' Obtenir les informations complètes d'un billet pour impression
' ============================================================
Public Function ObtenirInfosBillet(idBillet As Long) As DAO.Recordset
    Dim sql As String
    sql = "SELECT B.numero_billet, B.date_emission, B.statut_billet, " & _
          "R.id_reservation, R.date_debut, R.date_fin, R.nb_personnes, R.montant_total, R.statut, " & _
          "C.nom & ' ' & C.prenom AS client_nom, C.email, C.telephone, " & _
          "U.nom_complet AS agent " & _
          "FROM ((BILLET AS B " & _
          "INNER JOIN RESERVATION AS R ON B.id_reservation = R.id_reservation) " & _
          "INNER JOIN CLIENT AS C ON R.id_client = C.id_client) " & _
          "LEFT JOIN UTILISATEUR AS U ON R.id_utilisateur = U.id_utilisateur " & _
          "WHERE B.id_billet = " & idBillet

    Set ObtenirInfosBillet = OuvrirRecordset(sql)
End Function

' ============================================================
' Afficher un aperçu texte du billet dans une MessageBox
' ============================================================
Public Sub ApercuBillet(idBillet As Long)
    Dim rs As DAO.Recordset
    Dim nomEtablissement As String

    nomEtablissement = ObtenirParametre("NOM_ETABLISSEMENT", "RESERV+")

    Set rs = ObtenirInfosBillet(idBillet)

    If rs.EOF Then
        AfficherErreur "Billet introuvable."
        rs.Close
        Exit Sub
    End If

    Dim texte As String
    texte = String(50, "=") & vbCrLf
    texte = texte & "        " & nomEtablissement & vbCrLf
    texte = texte & "     Plateforme de réservation" & vbCrLf
    texte = texte & String(50, "=") & vbCrLf & vbCrLf
    texte = texte & "BILLET DE RÉSERVATION" & vbCrLf & vbCrLf
    texte = texte & "N° Billet : " & ChaineSure(rs!numero_billet) & vbCrLf
    texte = texte & "Date d'émission : " & FormatDateHeure(rs!date_emission) & vbCrLf & vbCrLf
    texte = texte & "CLIENT : " & ChaineSure(rs!client_nom) & vbCrLf
    texte = texte & "Tél : " & ChaineSure(rs!telephone) & vbCrLf & vbCrLf
    texte = texte & "RÉSERVATION #" & ChaineSure(rs!id_reservation) & vbCrLf
    texte = texte & "Période : " & FormatDate(rs!date_debut) & " → " & FormatDate(rs!date_fin) & vbCrLf
    texte = texte & "Personnes : " & ChaineSure(rs!nb_personnes) & vbCrLf
    texte = texte & "Montant total : " & FormatMontant(CCur(Nz(rs!montant_total, 0))) & vbCrLf & vbCrLf
    texte = texte & "Statut réservation : " & ChaineSure(rs!statut) & vbCrLf
    texte = texte & "Statut billet : " & ChaineSure(rs!statut_billet) & vbCrLf
    texte = texte & String(50, "=")

    rs.Close
    MsgBox texte, vbInformation, "Aperçu billet - " & APP_NOM
End Sub
