Attribute VB_Name = "DAORubros"
Dim rs As ADODB.Recordset

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_RUBRO As String = "rubro"
Public Const CAMPO_INICIALES As String = "iniciales"
Public Const CAMPO_CONTADOR As String = "contador"

Public Function Save(Rubro As clsRubros) As Boolean
    On Error GoTo er1
    Dim strsql As String
    Dim a As Long
    Save = True
    Dim n As Boolean

    If Rubro.id = 0 Then
        n = True
        strsql = "insert into rubros (rubro,iniciales) VALUES ('" & Rubro.Rubro & "','" & Rubro.iniciales & "')"
    Else
        strsql = "update rubros set rubro='" & Rubro.Rubro & "',iniciales='" & Rubro.iniciales & "' where id=" & Rubro.id
        n = False
    End If
    Save = conectar.execute(strsql)

    Rubro.id = conectar.UltimoId2
    Dim EVENTO As New clsEventoObserver
    Set vento.Elemento = Rubro
    If n Then
        EVENTO.EVENTO = agregar_
    Else
        EVENTO.EVENTO = modificar_
    End If
    EVENTO.Tipo = RubrosGrupos

    Channel.Notificar EVENTO, RubrosGrupos_

    Exit Function
er1:
    Save = False

End Function

Public Function FindById(id As Long) As clsRubros
    Dim col As Collection
    Set col = FindAll("id = " & id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function

Public Function FindAll(Optional filter As String = "1=1") As Collection
    Dim rs As ADODB.Recordset
    Dim q As String
    Dim rubros As New Collection

    q = "select * from rubros where " & filter

    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim r As clsRubros

    While Not rs.EOF
        Set r = Map(rs, fieldsIndex, "rubros")
        rubros.Add r, CStr(r.id)
        rs.MoveNext
    Wend

    Set FindAll = rubros
End Function

Public Function FindAllByProveedor(idProv As Long) As Collection
    Set FindAllByProveedor = FindAll("id IN (SELECT id_rubro FROM asignacion where id_proveedor = " & idProv & ")")
End Function


Public Sub LlenarCombo(cbo As ComboBox)
    Dim col As Collection
    Set col = DAORubros.FindAll
    Dim rub As clsRubros
    cbo.Clear



    For i = 1 To col.count
        Set rub = col(i)
        cbo.AddItem rub.iniciales & " - " & rub.Rubro
        cbo.ItemData(cbo.NewIndex) = rub.id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Sub LlenarComboExtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As Collection
    Set col = DAORubros.FindAll
    Dim rub As clsRubros
    cbo.Clear


    For i = 1 To col.count
        Set rub = col(i)
        cbo.AddItem rub.iniciales & " - " & rub.Rubro
        cbo.ItemData(cbo.NewIndex) = rub.id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsRubros
    Dim r As clsRubros
    Dim id As Long

    id = GetValue(rs, indice, tabla, DAORubros.CAMPO_ID)

    If id > 0 Then
        Set r = New clsRubros
        r.id = id
        r.Rubro = GetValue(rs, indice, tabla, DAORubros.CAMPO_RUBRO)
        r.Contador = GetValue(rs, indice, tabla, DAORubros.CAMPO_CONTADOR)
        r.iniciales = GetValue(rs, indice, tabla, DAORubros.CAMPO_INICIALES)
    End If

    Set Map = r
End Function
