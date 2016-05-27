VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim claseS As New classStock
Dim cn As ADODB.Connection
Dim rs As ADODB.recordset
Dim rs2 As ADODB.recordset
Dim estados(1 To 8) As String
Dim strEstados(1 To 8) As String
Dim strsql As String
'Public Function recotizar(idpresu As Long) As Long
'    On Error GoTo err4
'    If MsgBox("¿Está seguro de recotizar el presupuesto?", vbYesNo, "Confirmación") = vbYes Then
'        cn.BeginTrans
'        'copio los datos del encabezado del presupuesto
'        Set rs = conectar.RSFactory("Select * from presupuestos where id=" & idpresu)
'        If Not rs.EOF And Not rs.BOF Then
'            fecha_creado = funciones.datetimeFormateada(Now)
'Dim ultimo As Long
'            fecha_entrega = rs!FechaEntrega
'            idcli = rs!idCliente
'            idVendedor = rs!idVendedor
'            esta = 6    'rs!estado
'            deta = rs!detalle
'            Descu = rs!descuento
'            PorcMDO = rs!PorcMDO
'            PorcMen10 = rs!PorcMen10
'
'            PorcMen15 = rs!PorcMen15
'            PorcMas15 = rs!PorcMas15
'            gastoss = rs!Gastos
'            ManteOferta = rs!ManteOferta
'            venci = rs!VencimientoPresupuesto
'            idMon = rs!idMoneda
'
'            strsql = "insert into presupuestos (fecha, fechaEntrega, idCliente, idVendedor, estado, detalle, descuento, PorcMDO, PorcMen10, PorcMen15, PorcMas15, gastos,manteOferta,vencimientoPresupuesto,idmoneda) "
'            strsql = strsql & " Values ('" & fecha_creado & "','" & fecha_entrega & "'," & idcli & "," & idVendedor & "," & esta & ",'" & deta & "'," & Descu & "," & PorcMDO & "," & PorcMen10 & "," & PorcMen15 & "," & PorcMas15 & "," & gastoss & "," & ManteOferta & ",'" & Format(venci, "yyyy/mm/dd") & "'," & idMon & ")"
'            cn.execute strsql
'
'            conectar.UltimoId "presupuestos", ultimo
'
'
'            'copio los datls del detalle
'            Set rs = conectar.RSFactory("select * from detalle_presupuesto where idPresupuesto=" & idpresu)
'            While Not rs.EOF
'                forma_cotizar = rs!forma_cotizar
'                amort = rs!amort
'                Cantidad = rs!Cantidad
'                Item = rs!Item
'                valorUnitario = rs!valorUnitario
'                valorUnitarioManual = rs!valorUnitarioManual
'                idpresupuesto = ultimo
'                entregaItem = rs!entregaItem
'                masInfo = rs!masDetalles
'                idPieza = rs!idPieza
'                strsql2 = "insert into detalle_presupuesto (entregaitem,amort,idpresupuesto,item,idpieza,cantidad,valorunitario,valorUnitarioManual,masDetalles,forma_cotizar) VALUES (" & entregaItem & "," & amort & "," & idpresupuesto & ",'" & Item & "'," & idPieza & "," & Cantidad & "," & valorUnitario & "," & valorUnitarioManual & ",'" & masInfo & "'," & forma_cotizar & ")"
'                cn.execute strsql2
'                rs.MoveNext
'            Wend
'
'        Else
'            Exit Function
'        End If
'
'
'
'
'
'        cn.CommitTrans
'        MsgBox "Recotización exitosa!" & Chr(10) & "Número: " & ultimo, vbInformation, "Información"
'        recotizar = ultimo
'    End If
'    Exit Function
'err4:
'    cn.RollbackTrans
'
'
'End Function

Public Property Get count() As Long
    count = rs!count
End Property

Public Property Get estadoPresu() As Integer
    estadoPresu = rs!estado
End Property
Public Property Get Precio() As Double
    Precio = rs!Precio
End Property

Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion
    estados(1) = "Pendiente"
    estados(2) = "Enviado"
    estados(3) = "Procesado"
    estados(4) = "Modificado"
    estados(5) = "Rechazado"
    estados(6) = "A cotizar"
    estados(7) = "No cotizado"
    estados(8) = "Desactivado"
    strEstados(1) = "Aprobación del presupuesto"
    strEstados(2) = "Presupuesto enviado al cliente"
    strEstados(3) = "Orden de trabajo creada desde presupuesto"
    strEstados(4) = "Presupuesto modificado"
    strEstados(5) = "Presupuesto rechazado totalmente"
    strEstados(6) = "Presupuesto pendiente de finalización"
    strEstados(7) = "Presupuesto no cotizado"
    strEstados(8) = "Presupuesto desactivado"
    Exit Sub
err22:
    MsgBox Err.Description
End Sub



Public Function buscar_presupuesto(ByVal id As Integer, ByVal lst As ListView, ByRef detalle, ByRef Cliente, ByRef entrega, ByRef idcli, Optional ByRef dtos, Optional ByRef idMoneda) As Boolean
    strsql = "select p.idMoneda,p.descuento,p.estado,p.detalle, p.fechaEntrega, c.id as idcli, c.razon from presupuestos p inner join clientes c on p.idcliente=c.id and p.id=" & id
    Set rs = conectar.RSFactory(strsql)

    buscar_presupuesto = False
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend

    If c = 1 Then
        rs.MoveFirst
        If rs!estado = 2 Then    'si el estadado es enviado.
            buscar_presupuesto = True
            dtos = rs!descuento
            detalle = rs!detalle
            Cliente = rs!Razon
            entrega = rs!FechaEntrega
            idcli = rs!idcli
            idMoneda = rs!idMoneda
            lst.ListItems.Clear
            strsql = "select dp.entregaItem,dp.masDetalles,dp.valorUnitario,dp.item,dp.idpieza,dp.cantidad, s.detalle, s.cantidad as cStock from detalle_presupuesto dp inner join stock s on dp.idpieza=s.id and dp.idpresupuesto=" & id
            Set rs = conectar.RSFactory(strsql)
            Dim X As ListItem
            While Not rs.EOF
                Set X = lst.ListItems.Add(, , rs!Item)
                X.SubItems(1) = rs!Cantidad
                If Trim(rs!masDetalles) = Empty Then
                    X.SubItems(2) = rs!detalle & rs!masDetalles
                Else
                    X.SubItems(2) = rs!detalle & ", " & rs!masDetalles
                End If

                X.SubItems(3) = rs!idPieza
                X.SubItems(4) = funciones.RedondearDecimales(rs!valorUnitario, 2)
                X.SubItems(5) = funciones.RedondearDecimales(rs!valorUnitario * rs!Cantidad, 2)
                Dim a As Date
                a = DateAdd("d", rs!entregaItem, Now)
                X.SubItems(6) = Format(a, "dd/mm/yyyy")
                If rs!cstock >= rs!Cantidad Then
                    X.SubItems(9) = rs!Cantidad
                Else
                    X.SubItems(9) = rs!cstock
                End If
                X.SubItems(7) = rs!detalle
                X.SubItems(8) = rs!masDetalles







                If rs!cstock >= rs!Cantidad Then
                    X.ForeColor = vbRed
                    X.ListSubItems(1).ForeColor = vbRed
                    X.ListSubItems(2).ForeColor = vbRed
                    X.ListSubItems(3).ForeColor = vbRed
                    X.ListSubItems(4).ForeColor = vbRed
                    X.ListSubItems(5).ForeColor = vbRed
                    X.ListSubItems(6).ForeColor = vbRed
                End If

                rs.MoveNext
            Wend
        ElseIf rs!estado = 1 Then
            MsgBox "Primero debería mandar el presupuesto al cliente.", vbCritical, "Error"
        ElseIf rs!estado = 5 Then

        ElseIf rs!estado = 3 Then
            MsgBox "No puede procesar un prespuesto ya procesado por planeamiento.", vbCritical, "Error"
        ElseIf rs!estado = 6 Then
            MsgBox "No puede procesar un prespuesto no terminado.", vbCritical, "Error"
        ElseIf rs!estado = 8 Then
            MsgBox "No puede procesar un prespuesto desactivado.", vbCritical, "Error"
        End If
    Else
        MsgBox "Número de presupuesto no encontrado", vbCritical, "Error"
    End If
End Function
Public Function buscarDetalle(ref As String) As Boolean
    Set rs = conectar.RSFactory("select count(id) as cant from presupuestos where detalle='" & ref & "'")
    If rs!Cant > 0 Then
        buscarDetalle = True
    Else
        buscarDetalle = False
    End If
End Function

