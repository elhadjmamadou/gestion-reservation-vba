Attribute VB_Name = "modDatabase"
Option Compare Database
Option Explicit

' ============================================================
' RESERV+ - Utilitaires base de données
' Module : modDatabase.bas
' ============================================================

' ============================================================
' Exécuter une instruction SQL sans retour de données
' ============================================================
Public Sub ExecuterSQL(sql As String)
    On Error GoTo GestionErreur

    CurrentDb.Execute sql, dbFailOnError
    Exit Sub

GestionErreur:
    Err.Raise Err.Number, "ExecuterSQL", "Erreur SQL : " & Err.Description & vbCrLf & "Requête : " & sql
End Sub

' ============================================================
' Obtenir une valeur scalaire depuis une requête SQL
' ============================================================
Public Function ObtenirValeur(sql As String, Optional valeurDefaut As Variant = Null) As Variant
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo GestionErreur

    Set db = CurrentDb()
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        If IsNull(rs.Fields(0).Value) Then
            ObtenirValeur = valeurDefaut
        Else
            ObtenirValeur = rs.Fields(0).Value
        End If
    Else
        ObtenirValeur = valeurDefaut
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

GestionErreur:
    ObtenirValeur = valeurDefaut
    If Not rs Is Nothing Then rs.Close
End Function

' ============================================================
' Ouvrir un Recordset en lecture
' ============================================================
Public Function OuvrirRecordset(sql As String) As DAO.Recordset
    Dim db As DAO.Database
    Set db = CurrentDb()
    Set OuvrirRecordset = db.OpenRecordset(sql, dbOpenSnapshot)
End Function

' ============================================================
' Ouvrir un Recordset modifiable (dynaset)
' ============================================================
Public Function OuvrirRecordsetEditable(nomTable As String) As DAO.Recordset
    Dim db As DAO.Database
    Set db = CurrentDb()
    Set OuvrirRecordsetEditable = db.OpenRecordset(nomTable, dbOpenDynaset)
End Function

' ============================================================
' Compter le nombre d'enregistrements correspondant à une condition
' ============================================================
Public Function CompterEnregistrements(nomTable As String, condition As String) As Long
    Dim sql As String
    Dim valeur As Variant

    If Len(condition) > 0 Then
        sql = "SELECT Count(*) FROM " & nomTable & " WHERE " & condition
    Else
        sql = "SELECT Count(*) FROM " & nomTable
    End If

    valeur = ObtenirValeur(sql, 0)
    CompterEnregistrements = CLng(Nz(valeur, 0))
End Function

' ============================================================
' Vérifier si un enregistrement existe
' ============================================================
Public Function EnregistrementExiste(nomTable As String, condition As String) As Boolean
    EnregistrementExiste = (CompterEnregistrements(nomTable, condition) > 0)
End Function

' ============================================================
' Obtenir le dernier ID inséré dans une table
' ============================================================
Public Function DernierID(nomTable As String, champID As String) As Long
    Dim sql As String
    sql = "SELECT Max(" & champID & ") FROM " & nomTable
    DernierID = CLng(Nz(ObtenirValeur(sql, 0), 0))
End Function

' ============================================================
' Vérifier si une table existe dans la base
' ============================================================
Public Function TableExiste(nomTable As String) As Boolean
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef

    On Error GoTo GestionErreur

    Set db = CurrentDb()
    For Each tbl In db.TableDefs
        If tbl.Name = nomTable Then
            TableExiste = True
            Exit Function
        End If
    Next tbl

    TableExiste = False
    Exit Function

GestionErreur:
    TableExiste = False
End Function

' ============================================================
' Vérifier si une requête enregistrée existe
' ============================================================
Public Function RequeteExiste(nomRequete As String) As Boolean
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    On Error GoTo GestionErreur

    Set db = CurrentDb()
    For Each qdf In db.QueryDefs
        If qdf.Name = nomRequete Then
            RequeteExiste = True
            Exit Function
        End If
    Next qdf

    RequeteExiste = False
    Exit Function

GestionErreur:
    RequeteExiste = False
End Function

' ============================================================
' Vérifier si un formulaire existe
' ============================================================
Public Function FormulaireExiste(nomFormulaire As String) As Boolean
    Dim obj As AccessObject

    On Error GoTo GestionErreur

    For Each obj In CurrentProject.AllForms
        If obj.Name = nomFormulaire Then
            FormulaireExiste = True
            Exit Function
        End If
    Next obj

    FormulaireExiste = False
    Exit Function

GestionErreur:
    FormulaireExiste = False
End Function

' ============================================================
' Vérifier si un état (report) existe
' ============================================================
Public Function EtatExiste(nomEtat As String) As Boolean
    Dim obj As AccessObject

    On Error GoTo GestionErreur

    For Each obj In CurrentProject.AllReports
        If obj.Name = nomEtat Then
            EtatExiste = True
            Exit Function
        End If
    Next obj

    EtatExiste = False
    Exit Function

GestionErreur:
    EtatExiste = False
End Function

' ============================================================
' Supprimer un objet Access (formulaire ou état)
' ============================================================
Public Sub SupprimerObjetAccess(typeObjet As AcObjectType, nomObjet As String)
    On Error Resume Next
    DoCmd.DeleteObject typeObjet, nomObjet
    On Error GoTo 0
End Sub

' ============================================================
' Supprimer une requête enregistrée
' ============================================================
Public Sub SupprimerRequete(nomRequete As String)
    Dim db As DAO.Database

    On Error Resume Next
    Set db = CurrentDb()
    db.QueryDefs.Delete nomRequete
    On Error GoTo 0
End Sub
