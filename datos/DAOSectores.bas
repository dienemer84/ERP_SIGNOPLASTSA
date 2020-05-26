Attribute VB_Name = "DAOSectores"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "sector"
Public Const TABLA_SECTOR As String = "sec"

Dim rs As ADODB.Recordset

Public Function GetById(id) As clsSector
    On Error GoTo er1
    Dim a As clsSector
    Set rs = conectar.RSFactory("select * from sectores where id=" & id)
    If Not rs.EOF And Not rs.BOF Then
        Set a = New clsSector
        a.Sector = rs!Sector
        a.id = rs!id
    End If
    Set GetById = a
    Exit Function
er1:
    Set GetById = Nothing
End Function


Public Function GetAll() As Collection
    Dim col As New Collection
    Dim a As clsSector
    Dim fieldsIndex As Dictionary
    Set rs = conectar.RSFactory("select sec.* from sectores sec order by sec.sector asc")
    conectar.BuildFieldsIndex rs, fieldsIndex
    While Not rs.EOF
        'Set a = map(rs, fieldsindex, TABLA_SECTOR)
        'a.Sector = rs!Sector
        'a.Id = rs!Id


        col.Add Map(rs, fieldsIndex, TABLA_SECTOR), CStr(rs.Fields(fieldsIndex("sec.id")).value)
        rs.MoveNext
    Wend
    Set a = Nothing
    Set GetAll = col
End Function
Public Function GetByIdEmpleado(id As Long) As Collection
    Dim col As New Collection
    Dim a As clsSector
    Set rs = conectar.RSFactory("select s.id as id_sector, e.idEmpleado, s.sector from sectorizacion e inner join sectores s on e.idSector=s.id  where idEmpleado=" & id)
    While Not rs.EOF
        Set a = New clsSector
        a.Sector = rs!Sector
        a.id = rs!id_sector
        col.Add a, CStr(a.id)
        rs.MoveNext
    Wend
    Set a = Nothing
    Set GetByIdEmpleado = col
End Function
Public Sub LlenarCombo(cbo As ComboBox, Optional sectores As Collection = Nothing)
    Dim col As Collection
    Dim sec As clsSector
    cbo.Clear
    If sectores Is Nothing Then
        Set col = DAOSectores.GetAll
    Else
        Set col = sectores
    End If

    For i = 1 To col.count
        Set sec = col(i)
        cbo.AddItem sec.Sector
        cbo.ItemData(cbo.NewIndex) = sec.id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Sub LlenarComboXtreme(cbo As Xtremesuitecontrols.ComboBox, Optional sectores As Collection = Nothing)
    Dim col As Collection
    Dim sec As clsSector
    cbo.Clear
    If sectores Is Nothing Then
        Set col = DAOSectores.GetAll
    Else
        Set col = sectores
    End If

    For i = 1 To col.count
        Set sec = col(i)
        cbo.AddItem sec.Sector
        cbo.ItemData(cbo.NewIndex) = sec.id
    Next i
End Sub


Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef table As String) As clsSector
    Dim s As clsSector
    id = GetValue(rs, fieldsIndex, table, CAMPO_ID)
    If id > 0 Then
        Set s = New clsSector
        s.id = id
        s.Sector = GetValue(rs, fieldsIndex, table, CAMPO_NOMBRE)
        Set Map = s

    End If

End Function




