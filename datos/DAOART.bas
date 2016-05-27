Attribute VB_Name = "DAOART"
Option Explicit

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As ART
    Dim a As ART
    Dim id As Long
    id = GetValue(rs, indice, tabla, "id")
    If id > 0 Then
        Set a = New ART
        a.id = id
        a.nombre = GetValue(rs, indice, tabla, "nombre")
    End If
    Set Map = a
End Function

Public Function FindAll(Optional filtro As String = Empty) As Collection

    On Error GoTo err1
    Dim col As New Collection
    Dim a As ART
    Dim indice As Dictionary
    Dim q As String
    Dim rs As Recordset

    q = "select * from art a where 1=1"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    Set rs = conectar.RSFactory(q)

    conectar.BuildFieldsIndex rs, indice
    While Not rs.EOF
        Set a = New ART
        Set a = Map(rs, indice, "a")
        col.Add a, CStr(a.id)
        rs.MoveNext
    Wend

    Set FindAll = col
    Exit Function
err1:
    Set FindAll = Nothing
End Function
