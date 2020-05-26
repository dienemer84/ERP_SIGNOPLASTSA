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
        almacen.id = rs!id
        col.Add almacen
        rs.MoveNext
    Wend

    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function

Public Function GetById(id As Long) As clsAlmacen
    On Error GoTo err1
    Dim Grupo As New clsGrupo
    Dim almacen As clsAlmacen
    Set rs = conectar.RSFactory("select * from materialesAlmacenes where id=" & id)
    If Not rs.EOF And Not rs.BOF Then
        Set almacen = New clsAlmacen
        almacen.almacen = rs!detalle
        almacen.id = rs!id
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
        cbo.ItemData(cbo.NewIndex) = alma.id

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
        cbo.ItemData(cbo.NewIndex) = alma.id

    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub



Public Function Save(T As clsAlmacen) As Boolean
    On Error GoTo err1
    Save = True
    Dim strsql As String
    If T.id = 0 Then
        strsql = "insert into materialesAlmacenes (detalle) values ('" & T.almacen & "')"
    Else
        strsql = "update materialesAlmacenes set detalle='" & T.almacen & "' where id=" & T.id
    End If
    Save = conectar.execute(strsql)
    Exit Function
err1:
    Save = False
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsAlmacen
    Dim a As clsAlmacen
    Dim id As Long

    id = GetValue(rs, indice, tabla, CAMPO_ID)

    If id > 0 Then
        Set a = New clsAlmacen
        a.id = id
        a.almacen = GetValue(rs, indice, tabla, CAMPO_DETALLE)
    End If

    Set Map = a
End Function
