Attribute VB_Name = "DAOGrupos"
Dim rs As ADODB.Recordset
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_GRUPO As String = "grupo"
Public Const CAMPO_ID_RUBRO As String = "id_rubro"

Public Function GetAll() As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim Grupo As clsGrupo
    Set rs = conectar.RSFactory("select * from grupos")
    While Not rs.EOF
        Set Grupo = New clsGrupo
        Grupo.Grupo = rs!Grupo
        Grupo.Id = rs!Id
        Grupo.rubros = DAORubros.FindById(rs!id_rubro)
        col.Add Grupo
        rs.MoveNext
    Wend
    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function

Public Function GetAllByRubro(id_rubro As Long) As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim Grupo As clsGrupo
    Set rs = conectar.RSFactory("select * from grupos where id_rubro=" & id_rubro)
    While Not rs.EOF
        Set Grupo = New clsGrupo
        Grupo.Grupo = rs!Grupo
        Grupo.Id = rs!Id
        Grupo.rubros = DAORubros.FindById(rs!id_rubro)
        col.Add Grupo
        rs.MoveNext
    Wend
    Set GetAllByRubro = col
    Exit Function
err1:
    Set GetAllByRubro = Nothing
End Function


Public Function GetById(Id As Long) As clsGrupo
    On Error GoTo err1
    Dim Grupo As New clsGrupo
    Set rs = conectar.RSFactory("select * from grupos where id=" & Id)
    If Not rs.EOF And Not rs.BOF Then
        Grupo.Grupo = rs!Grupo
        Grupo.Id = rs!Id
        Grupo.rubros = DAORubros.FindById(rs!id_rubro)
        Set GetById = Grupo
    Else
        GoTo err1
    End If
    Exit Function

err1:
    Set GetById = Nothing
End Function

Public Sub LlenarCombo(cbo As ComboBox, Optional rubros As clsRubros)
    Dim Grupo As clsGrupo
    Dim col As Collection
    If rubros Is Nothing Then
        Set col = DAOGrupos.GetAll
    Else
        Set col = DAOGrupos.GetAllByRubro(rubros.Id)
    End If

    Dim rub As clsRubros
    cbo.Clear
    For i = 1 To col.count
        Set Grupo = col(i)
        cbo.AddItem Grupo.Grupo
        cbo.ItemData(cbo.NewIndex) = Grupo.Id

    Next i

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

End Sub
Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox, Optional rubros As clsRubros)
    Dim Grupo As clsGrupo
    Dim col As Collection
    If rubros Is Nothing Then
        Set col = DAOGrupos.GetAll
    Else
        Set col = DAOGrupos.GetAllByRubro(rubros.Id)
    End If
    Dim rub As clsRubros
    cbo.Clear
    For i = 1 To col.count
        Set Grupo = col(i)
        cbo.AddItem Grupo.Grupo
        cbo.ItemData(cbo.NewIndex) = Grupo.Id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

End Sub
Public Function Save(T As clsGrupo) As Boolean
    Save = True
    Dim strsql As String
    On Error GoTo er1
    Dim EVENTO As New clsEventoObserver
    If T.Id = 0 Then
        EVENTO.EVENTO = agregar_
        strsql = "insert into grupos (grupo,id_rubro) VALUES ('" & T.Grupo & "',  " & T.rubros.Id & ")"
    Else
        strsql = "update grupos set grupo='" & T.Grupo & "',id_rubro=" & T.rubros.Id & " where id=" & T.rubros.Id
        EVENTO.EVENTO = modificar_
    End If
    Save = conectar.execute(strsql)
    Exit Function

er1:
    Save = False
End Function
Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, tablaRubro As String) As clsGrupo
    Dim g As clsGrupo
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, DAOGrupos.CAMPO_ID)
    If Id > 0 Then
        Set g = New clsGrupo
        g.Id = Id
        g.Grupo = GetValue(rs, indice, tabla, DAOGrupos.CAMPO_GRUPO)
        If LenB(tablaRubro) Then g.rubros = DAORubros.Map(rs, indice, tablaRubro)    'poner el set
    End If
    Set Map = g
End Function

