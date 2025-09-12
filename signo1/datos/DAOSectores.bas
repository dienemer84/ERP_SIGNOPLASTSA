Attribute VB_Name = "DAOSectores"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "sector"
Public Const CAMPO_SECTORIZACION As String = "sectorizacion"
Public Const CAMPO_MODULO As String = "modulo"
Public Const TABLA_SECTOR As String = "sec"

Dim rs As ADODB.Recordset

Public Function GetById(Id) As clsSector
    On Error GoTo er1
    Dim A As clsSector
    Set rs = conectar.RSFactory("select * from sectores where id=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        Set A = New clsSector
        A.Sector = rs!Sector
        A.Id = rs!Id
    End If
    Set GetById = A
    Exit Function
er1:
    Set GetById = Nothing
End Function


Public Function GetByIdModulo(Id) As clsSector
    On Error GoTo er1
    Dim A As clsSector
    Set rs = conectar.RSFactory("select * from sectores where sectorizacion=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        Set A = New clsSector
        A.Modulo = rs!Modulo
        A.Sectorizacion = rs!Id
    End If
    Set GetByIdModulo = A
    Exit Function
er1:
    Set GetByIdModulo = Nothing
End Function


Public Function GetAll() As Collection
    Dim col As New Collection
    Dim A As clsSector
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
    Set A = Nothing
    Set GetAll = col
End Function


Public Function GetAllModulos() As Collection
    Dim col As New Collection
    Dim A As clsSector
    Dim fieldsIndex As Dictionary
    Set rs = conectar.RSFactory("select DISTINCT sec.sectorizacion, sec.modulo from sectores sec WHERE sec.sectorizacion IS NOT NULL order by sec.sectorizacion asc")
    conectar.BuildFieldsIndex rs, fieldsIndex
    While Not rs.EOF
        col.Add MapModulos(rs, fieldsIndex, TABLA_SECTOR), CStr(rs.Fields(fieldsIndex("sec.sectorizacion")).value)
        rs.MoveNext
    Wend
    Set A = Nothing
    Set GetAllModulos = col
End Function


Public Function GetByIdEmpleado(Id As Long) As Collection
    Dim col As New Collection
    Dim A As clsSector
    Set rs = conectar.RSFactory("select s.id as id_sector, e.idEmpleado, s.sector from sectorizacion e inner join sectores s on e.idSector=s.id  where idEmpleado=" & Id)
    While Not rs.EOF
        Set A = New clsSector
        A.Sector = rs!Sector
        A.Id = rs!id_sector
        col.Add A, CStr(A.Id)
        rs.MoveNext
    Wend
    Set A = Nothing
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
        cbo.ItemData(cbo.NewIndex) = sec.Id
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
        cbo.ItemData(cbo.NewIndex) = sec.Id
    Next i
End Sub


Public Sub LlenarComboXtremeModulos(cbo As Xtremesuitecontrols.ComboBox, Optional sectores As Collection = Nothing)
    Dim col As Collection
    Dim sec As clsSector
    cbo.Clear
    If sectores Is Nothing Then
        Set col = DAOSectores.GetAllModulos
    Else
        Set col = sectores
    End If

    For i = 1 To col.count
        Set sec = col(i)
        cbo.AddItem sec.Modulo
        cbo.ItemData(cbo.NewIndex) = sec.Sectorizacion
    Next i
End Sub

Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef table As String) As clsSector
    Dim s As clsSector
    Id = GetValue(rs, fieldsIndex, table, CAMPO_ID)
    If Id > 0 Then
        Set s = New clsSector
        s.Id = Id
        s.Sector = GetValue(rs, fieldsIndex, table, CAMPO_NOMBRE)
        Set Map = s
    End If
End Function


Public Function MapModulos(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef table As String) As clsSector
    Dim s As clsSector
    Id = GetValue(rs, fieldsIndex, table, CAMPO_SECTORIZACION)
    If Id > 0 Then
        Set s = New clsSector
        
        s.Sectorizacion = Id
        s.Modulo = GetValue(rs, fieldsIndex, table, CAMPO_MODULO)
        
        Set MapModulos = s
    End If
End Function



