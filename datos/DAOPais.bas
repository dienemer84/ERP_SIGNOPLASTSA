Attribute VB_Name = "DAOPais"
Option Explicit

Public Function FindById(id As Long) As pais
    Set FindById = FindAll(" And pa.id=" & id)(1)
End Function

Public Function FindAll(Optional filtro As String) As Collection
    On Error GoTo err1
    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    strsql = "Select * from Pais pa where 1=1"

    If LenB(filtro) > 0 Then strsql = strsql & filtro
    Dim col As New Collection
    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, idx

    While Not rs.EOF And Not rs.BOF
        col.Add Map(rs, idx, "pa")

        rs.MoveNext
    Wend

    Set FindAll = col
    Exit Function
err1:
    Set FindAll = Nothing

End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As pais

    Dim pais As pais
    Dim id As Long: id = GetValue(rs, indice, tabla, "ID")

    If id > 0 Then
        Set pais = New pais
        pais.id = id
        pais.nombre = GetValue(rs, indice, tabla, "Nombre")

    End If

    Set Map = pais
End Function

Public Function LlenarCombo(cbo As Xtremesuitecontrols.ComboBox)
    Dim P As pais
    Dim col As New Collection
    Set col = FindAll()
    For Each P In col
        If IsSomething(P) Then
            cbo.AddItem P.nombre
            cbo.ItemData(cbo.NewIndex) = P.id
        End If
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Function


Public Function Save(pais As pais) As Boolean
    Dim q As String
    Dim n As Boolean
    n = False
    On Error GoTo err1
    If pais.id > 0 Then
        q = "UPDATE sp.Pais  SET   Nombre = '" & UCase(pais.nombre) & "'   WHERE   ID = '" & pais.id & "' "
    Else
        q = "INSERT INTO sp.Pais (Nombre)VALUES('" & UCase(pais.nombre) & "')"
        n = True
    End If

    If Not conectar.execute(q) Then GoTo err1
    If n Then pais.id = conectar.UltimoId2
    Exit Function
err1:
    Save = False
End Function
