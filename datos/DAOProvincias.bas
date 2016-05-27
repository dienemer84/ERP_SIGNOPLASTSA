Attribute VB_Name = "DAOProvincias"
Option Explicit



Public Function FindAll(Optional filtro As String) As Collection
    On Error GoTo err1
    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    strsql = "Select * from Provincia p inner join Pais pa on p.idPais=pa.id where 1=1 "
    If LenB(filtro) > 0 Then strsql = strsql & filtro
    Dim col As New Collection
    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, idx

    While Not rs.EOF And Not rs.BOF
        col.Add Map(rs, idx, "p", "pa")

        rs.MoveNext
    Wend

    Set FindAll = col
    Exit Function
err1:
    Set FindAll = Nothing

End Function
Public Function FindAllByPais(idpais As Long) As Collection
    Set FindAllByPais = FindAll("and pa.id=" & idpais)
End Function

Public Function FindById(idprovincia As Long) As provincia

    Dim col As New Collection
    Set col = FindAll("and p.id=" & idprovincia)


    Set FindById = col(1)
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaPais As String) As provincia

    Dim prov As provincia
    Dim id As Long: id = GetValue(rs, indice, tabla, "ID")

    If id > 0 Then
        Set prov = New provincia
        prov.id = id
        prov.nombre = GetValue(rs, indice, tabla, "Nombre")
        If LenB(tablaPais) > 0 Then Set prov.pais = DAOPais.Map(rs, indice, tablaPais)
    End If

    Set Map = prov
End Function

Public Function LlenarCombo(cbo As Xtremesuitecontrols.ComboBox, id As Long)
    Dim P As provincia
    cbo.Clear
    Dim col As New Collection
    Set col = FindAllByPais(id)
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


Public Function Save(P As provincia) As Boolean
    Dim q As String
    On Error GoTo err1
    Dim n As Boolean
    If P.id > 0 Then

        q = "UPDATE sp.Provincia  SET  idPais=" & P.pais.id & ", Nombre = '" & UCase(P.nombre) & "'   WHERE   ID = '" & P.id & "' "
        n = False
    Else
        q = "INSERT INTO sp.Provincia (Nombre,idPais)VALUES('" & UCase(P.nombre) & "'," & P.pais.id & ")"
        n = True
    End If

    If Not conectar.execute(q) Then GoTo err1
    If n Then P.id = conectar.UltimoId2

    Exit Function
err1:
    Save = False
End Function
