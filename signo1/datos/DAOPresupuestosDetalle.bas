Attribute VB_Name = "DAOPresupuestosDetalle"
Option Explicit
Dim rs As Recordset
Dim tmpPresupuestoDetalle As clsPresupuestoDetalle
Dim strsql As String
Public Const TABLA_DETALLE = "dp"
Public Const TABLA_PIEZA = "s"
Public Const CAMPO_CANTIDAD = "cantidad"
Public Const CAMPO_AMORT = "amort"
Public Const CAMPO_PRECIO_SISTEMA = "ValorUnitario"
Public Const CAMPO_PRECIO_MANUAL = "valorUnitarioManual"
Public Const CAMPO_FORMA_COTIZAR = "forma_cotizar"
Public Const CAMPO_ENTREGA_ITEM = "entregaItem"
Public Const CAMPO_MAS_DETALLE = "masDetalles"
'Public Const CAMPO_PIEZA_ID = "Pieza"
Public Const CAMPO_ITEM = "item"
Public Const CAMPO_ID = "id"


Dim col As Collection


Public Function GetAllById(Id As Long) As clsPresupuestoDetalle

    strsql = "SELECT dp.*,s.*, p.* FROM detalle_presupuesto dp LEFT JOIN stock s ON dp.idPieza=s.id  left join presupuestos p on dp.idPresupuesto=p.id WHERE dp.id = " & Id & " ORDER BY dp.item"


    Set col = New Collection
    Set rs = conectar.RSFactory(strsql)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim tmp As clsPresupuestoDetalle

    While Not rs.EOF
        Set tmp = Map(rs, fieldsIndex, TABLA_DETALLE, TABLA_PIEZA)
        'Set tmp.presupuesto = T



        col.Add tmp, CStr(tmp.Id)
        rs.MoveNext
    Wend

    Set GetAllById = tmp
End Function

Public Function GetAllByPieza(idPieza As Long) As Collection

    strsql = "SELECT dp.*,s.*,p.* FROM detalle_presupuesto dp LEFT JOIN stock s ON dp.idPieza=s.id left join presupuestos p on dp.idPresupuesto=p.id WHERE dp.idPieza = " & idPieza & " ORDER BY dp.item"


    Set col = New Collection
    Set rs = conectar.RSFactory(strsql)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim tmp As clsPresupuestoDetalle

    While Not rs.EOF
        Set tmp = Map(rs, fieldsIndex, TABLA_DETALLE, TABLA_PIEZA)




        col.Add tmp, CStr(tmp.Id)
        rs.MoveNext
    Wend

    Set GetAllByPieza = col
End Function



Public Function GetAllByPresupuesto(T As clsPresupuesto) As Collection

    strsql = "SELECT dp.*,s.*,p.*  FROM detalle_presupuesto dp LEFT JOIN stock s ON dp.idPieza=s.id left join presupuestos p on dp.idPresupuesto=p.id WHERE dp.idPresupuesto = " & T.Id & " ORDER BY dp.item"


    Set col = New Collection
    Set rs = conectar.RSFactory(strsql)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim tmp As clsPresupuestoDetalle

    While Not rs.EOF
        Set tmp = Map(rs, fieldsIndex, TABLA_DETALLE, TABLA_PIEZA)
        Set tmp.presupuesto = T



        col.Add tmp, CStr(tmp.Id)
        rs.MoveNext
    Wend

    Set GetAllByPresupuesto = col
End Function

Public Function Save(T As clsPresupuesto) As Boolean
    On Error GoTo er1
    Dim i As Long
    Save = True
    'elimino los datos viejos
    Dim tmp As clsPresupuestoDetalle
    conectar.execute "delete from detalle_presupuesto where idPresupuesto=" & T.Id

    For i = 1 To T.DetallePresupuesto.count
        Set tmp = T.DetallePresupuesto.item(i)
        If tmp.entrega = 0 Then tmp.entrega = T.FechaEntrega
        If T.FechaEntrega = 0 Then tmp.entrega = 0
        strsql = "insert into detalle_presupuesto (indice_ajuste,idpresupuesto,item,idpieza,cantidad,valorunitario,valorUnitarioManual,masDetalles,entregaItem,amort,forma_cotizar) VALUES (" & tmp.indiceAjuste & " ," & T.Id & ",'" & tmp.item & "'," & tmp.Pieza.Id & "," & tmp.Cantidad & "," & tmp.ValorManual & "," & tmp.ValorSistema & ",'" & tmp.Detalles & "'," & tmp.entrega & "," & tmp.Amortizacion & "," & tmp.FormaCotizar & ")"
        conectar.execute strsql
    Next
    Exit Function
er1:
    Save = False
End Function


Public Function Map(ByRef rs As Recordset, Index As Dictionary, ByRef tableNameOrAlias As String, Optional tablePieza As String = vbNullString) As clsPresupuestoDetalle
    Dim Id As Long
    Id = GetValue(rs, Index, tableNameOrAlias, CAMPO_ID)
    If Id > 0 Then
        Set tmpPresupuestoDetalle = New clsPresupuestoDetalle
        tmpPresupuestoDetalle.Cantidad = GetValue(rs, Index, tableNameOrAlias, CAMPO_CANTIDAD)
        tmpPresupuestoDetalle.Amortizacion = GetValue(rs, Index, tableNameOrAlias, CAMPO_AMORT)
        tmpPresupuestoDetalle.entrega = GetValue(rs, Index, tableNameOrAlias, CAMPO_ENTREGA_ITEM)
        tmpPresupuestoDetalle.FormaCotizar = GetValue(rs, Index, tableNameOrAlias, CAMPO_FORMA_COTIZAR)
        tmpPresupuestoDetalle.Id = Id
        tmpPresupuestoDetalle.FechaPresupuesto = GetValue(rs, Index, "p", "fecha")
        tmpPresupuestoDetalle.idPreuspuesto = GetValue(rs, Index, tableNameOrAlias, "idPresupuesto")
        tmpPresupuestoDetalle.item = GetValue(rs, Index, tableNameOrAlias, CAMPO_ITEM)
        'tmpPresupuestoDetalle.pieza = GetValue(rs, Index, tableNameOrAlias, CAMPO_PIEZA_ID)
        tmpPresupuestoDetalle.ValorManual = GetValue(rs, Index, tableNameOrAlias, CAMPO_PRECIO_SISTEMA)
        tmpPresupuestoDetalle.ValorSistema = GetValue(rs, Index, tableNameOrAlias, CAMPO_PRECIO_MANUAL)
        tmpPresupuestoDetalle.Detalles = GetValue(rs, Index, tableNameOrAlias, CAMPO_MAS_DETALLE)
        tmpPresupuestoDetalle.indiceAjuste = GetValue(rs, Index, tableNameOrAlias, "indice_ajuste")
        Set tmpPresupuestoDetalle.Pieza = DAOPieza.Map(rs, Index, TABLA_PIEZA)

        Set Map = tmpPresupuestoDetalle
    End If
End Function
