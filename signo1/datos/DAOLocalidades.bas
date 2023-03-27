Attribute VB_Name = "DAOLocalidades"
Option Explicit



Public Function FindAll(Optional filtro As String) As Collection
    On Error GoTo err1
    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    strsql = "Select * from Localidades l inner join Provincia p on l.idProvincia=p.id inner join Pais pa on p.idPais=pa.id where 1=1  "
    If LenB(filtro) > 0 Then strsql = strsql & filtro

    strsql = strsql & " order by l.Nombre"
    Dim col As New Collection
    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, idx


    While Not rs.EOF And Not rs.BOF
        col.Add Map(rs, idx, "l", "p", "pa")

        rs.MoveNext
    Wend

    Set FindAll = col
    Exit Function
err1:
    Set FindAll = Nothing

End Function

Public Function FindById(Id As Long) As localidad
    Dim col As New Collection
    Set col = FindAll("And l.id=" & Id)
    Set FindById = col(1)
End Function
Public Function FindAllByProvincia(idDto As Long) As Collection
    Dim c As New Collection
    Set FindAllByProvincia = FindAll("and p.id=" & idDto)
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaProv As String = vbNullString, _
                    Optional tablaPais As String = vbNullString _
                  ) As localidad

    Dim loc As localidad
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "ID")

    If Id > 0 Then
        Set loc = New localidad
        loc.Id = Id
        loc.nombre = GetValue(rs, indice, tabla, "Nombre")
        loc.cp = GetValue(rs, indice, tabla, "CP")
        If LenB(tablaProv) > 0 Then Set loc.provincia = DAOProvincias.Map(rs, indice, tablaProv, tablaPais)

    End If

    Set Map = loc
End Function



Public Function LlenarCombo(cbo As Xtremesuitecontrols.ComboBox, Id As Long)
    Dim P As localidad
    cbo.Clear
    Dim col As New Collection
    Set col = FindAllByProvincia(Id)
    For Each P In col
        If IsSomething(P) Then
            cbo.AddItem P.nombre
            cbo.ItemData(cbo.NewIndex) = P.Id
        End If
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Function

Public Function Save(l As localidad) As Boolean
    Dim q As String
    On Error GoTo err1
    If l.Id > 0 Then

        q = "UPDATE sp.Localidades  SET  CP='" & l.cp & "', idProvincia=" & l.provincia.Id & ", Nombre = '" & UCase(l.nombre) & "'   WHERE   ID = '" & l.Id & "' "
    Else
        q = "INSERT INTO sp.Localidades (Nombre,idProvincia,CP)VALUES('" & UCase(l.nombre) & "'," & l.provincia.Id & ",'" & l.cp & "')"
    End If

    If Not conectar.execute(q) Then GoTo err1
    Exit Function
err1:
    Save = False
End Function
