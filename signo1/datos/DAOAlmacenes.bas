Attribute VB_Name = "DAOAlmacenes"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_DETALLE As String = "detalle"

Public Function GetAll() As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim almacen As clsAlmacen
    Set rs = conectar.RSFactory("select * from materialesAlmacenes")
    While Not rs.EOF
        Set almacen = New clsAlmacen
        almacen.almacen = rs!detalle
        almacen.Id = rs!Id
        col.Add almacen
        rs.MoveNext
    Wend

    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function

Public Function GetById(Id As Long) As clsAlmacen
    On Error GoTo err1
    Dim Grupo As New clsGrupo
    Dim almacen As clsAlmacen
    Set rs = conectar.RSFactory("select * from materialesAlmacenes where id=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        Set almacen = New clsAlmacen
        almacen.almacen = rs!detalle
        almacen.Id = rs!Id
        Set GetById = almacen

    Else
        GoTo err1
    End If
    Exit Function

err1:
    Set GetById = Nothing
End Function


Public Sub LlenarCombo(cbo As ComboBox)

    Dim col As New Collection
    Set col = DAOAlmacenes.GetAll
    Dim alma As clsAlmacen
    cbo.Clear
    For i = 1 To col.count
        Set alma = col(i)
        cbo.AddItem alma.almacen
        cbo.ItemData(cbo.NewIndex) = alma.Id

    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)

    Dim col As New Collection
    Set col = DAOAlmacenes.GetAll
    Dim alma As clsAlmacen
    cbo.Clear
    For i = 1 To col.count
        Set alma = col(i)
        cbo.AddItem alma.almacen
        cbo.ItemData(cbo.NewIndex) = alma.Id

    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub



Public Function Save(T As clsAlmacen) As Boolean
    On Error GoTo err1
    Save = True
    Dim strsql As String
    If T.Id = 0 Then
        strsql = "insert into materialesAlmacenes (detalle) values ('" & T.almacen & "')"
    Else
        strsql = "update materialesAlmacenes set detalle='" & T.almacen & "' where id=" & T.Id
    End If
    Save = conectar.execute(strsql)
    Exit Function
err1:
    Save = False
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsAlmacen
    Dim A As clsAlmacen
    Dim Id As Long

    Id = GetValue(rs, indice, tabla, CAMPO_ID)

    If Id > 0 Then
        Set A = New clsAlmacen
        A.Id = Id
        A.almacen = GetValue(rs, indice, tabla, CAMPO_DETALLE)
    End If

    Set Map = A
End Function
