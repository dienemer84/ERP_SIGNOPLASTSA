Attribute VB_Name = "DAODepartamentos"
Option Explicit


Public Function FindAll(Optional filtro As String) As Collection
    On Error GoTo err1
    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    strsql = "Select * from Departamento d inner join Provincia p on d.idProvincia=p.id inner join Pais pa on p.idPais=pa.id where 1=1 "

    If LenB(filtro) > 0 Then strsql = strsql & filtro

    Dim col As New Collection
    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, idx

    While Not rs.EOF And Not rs.BOF
        col.Add Map(rs, idx, "d", "p")

        rs.MoveNext
    Wend

    Set FindAll = col
    Exit Function
err1:
    Set FindAll = Nothing

End Function
Public Function FindById(Id As Long) As Departamento
    Set FindById = FindAll("And d.id=" & Id)(0)
End Function
Public Function FindAllByProvincia(idprovincia As Long) As Collection
    Dim c As New Collection
    Set FindAllByProvincia = FindAll("and p.id=" & idprovincia)
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaProv As String = vbNullString, _
                    Optional tablaPais As String = vbNullString _
                  ) As Departamento

    Dim dep As Departamento
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "ID")

    If Id > 0 Then
        Set dep = New Departamento
        dep.Id = Id
        dep.nombre = GetValue(rs, indice, tabla, "Nombre")
        If LenB(tablaProv) > 0 Then Set dep.provincia = DAOProvincias.Map(rs, indice, tablaProv, tablaPais)

    End If

    Set Map = dep
End Function

Public Function LlenarCombo(cbo As Xtremesuitecontrols.ComboBox, idprovincia As Long)
    Dim P As Departamento
    cbo.Clear
    Dim col As New Collection
    Set col = FindAllByProvincia(idprovincia)
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

