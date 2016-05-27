Attribute VB_Name = "DAORecibo"
Option Explicit

Public Function FindById(id As Long, _
                         Optional includeRetenciones As Boolean = False, _
                         Optional includeCheques As Boolean = False, _
                         Optional includeBanco As Boolean = False, _
                         Optional includeCaja As Boolean = False, _
                         Optional includeFacturas As Boolean = False _
                         ) As recibo

    Dim col As Collection
    Set col = FindAll("rec.id = " & id, includeRetenciones, includeCheques, includeBanco, includeCaja, includeFacturas)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function


Public Function proximo() As Long
On Error GoTo err1
Dim q As String
Dim rs As Recordset
q = "select max(id)+1 as ultimo from AdminRecibos"
Set rs = conectar.RSFactory(q)
If Not rs.EOF And Not rs.BOF Then
proximo = rs!ultimo
End If
Exit Function
err1:
proximo = -1

End Function

Public Function Anular(recibo As recibo) As Boolean
    conectar.BeginTransaction

    If recibo.estado = EstadoRecibo.Aprobado Then
        'cambio el estado del recibo
        recibo.estado = EstadoRecibo.Reciboanulado



        'borro los cheques
        If Not conectar.execute("DELETE FROM Cheques WHERE id IN (SELECT idCheque FROM AdminRecibosCheques a WHERE a.`idRecibo`=" & recibo.id & ")") Then GoTo err101

        'borro los cheques x recibo
        If Not conectar.execute("DELETE FROM AdminRecibosCheques WHERE idRecibo=" & recibo.id) Then GoTo err101


        'borro las operaciones
        If Not conectar.execute("DELETE FROM `AdminRecibosDepositos` WHERE idRecibo=" & recibo.id) Then GoTo err101
        'DELETE FROM `AdminRecibosDepositos` WHERE idRecibo=5331


        'borro las facturas
        'if Not conectar.execute("DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo= " & recibo.Id) Then GoTo err101
        'DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo=5311


        Dim q As String
        q = "select * from AdminRecibosDetalleFacturas where idRecibo=" & recibo.id
        Dim rs As Recordset
        Set rs = conectar.RSFactory(q)
        Dim F As Factura
        Dim rs2 As Recordset
        While Not rs.EOF And Not rs.BOF

            q = "SELECT * FROM `AdminRecibosDetalleFacturas` f WHERE f.`idFactura`= " & rs!idFactura & "  AND f.`idRecibo`<>" & recibo.id

            Set rs2 = conectar.RSFactory(q)
            Dim pagoParcial As Boolean
            pagoParcial = False

            While Not rs2.EOF And Not rs2.BOF

                'si hay facturas aca es porq estan pagas en otro recibo
                pagoParcial = True

                rs2.MoveNext

            Wend

            Set F = DAOFactura.FindById(rs!idFactura)
            If IsSomething(F) Then
                If pagoParcial Then
                    F.Saldado = SaldadoParcial
                Else
                    F.Saldado = NoSaldada
                End If
                DAOFactura.Guardar F
            Else
                GoTo err101
            End If
            rs.MoveNext
        Wend
        If Not conectar.execute("DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo= " & recibo.id) Then GoTo err101



        'borro retencione
        If Not conectar.execute("DELETE FROM `AdminRecibosDetalleRetenciones` WHERE idRecibo=" & recibo.id) Then GoTo err101
        'DELETE FROM `AdminRecibosDetalleRetenciones` WHERE idRecibo=5331



        'libero los comprobasntes





        DAORecibo.Guardar recibo

        conectar.CommitTransaction

    Else
        GoTo err100


    End If
    Exit Function
err100:
    Err.Raise 100, , "El recibo debería estar aprobado para poder anularlo"
    conectar.RollBackTransaction
err101:
    Err.Raise 101, , "Error al anular el recibo." & Chr(10) & Err.Description
    conectar.RollBackTransaction


End Function


Public Function aprobar(recibo As recibo) As Boolean
    On Error GoTo err5
    Dim estAnt As EstadoRecibo
    Dim fechaAnt As Variant
    Dim Factura As Factura
    conectar.BeginTransaction


    estAnt = recibo.estado
    recibo.FechaAprobacion = Now
    Set recibo.UsuarioAprobador = funciones.GetUserObj
    recibo.estado = EstadoRecibo.Aprobado


    If recibo.IsValid Then
        'totalizo recibo
        Dim totEst As New TotalEstaticoRecibo
        totEst.TotalChequesEstatico = recibo.TotalCheques
        totEst.TotalDepositosEstatico = recibo.TotalOperacionesBanco
        totEst.TotalEfectivoEstatico = recibo.TotalOperacionesCaja
        totEst.TotalReciboEstatico = recibo.Total
        Set recibo.TotalEstatico = totEst

        If Not DAORecibo.Guardar(recibo) Then GoTo err5

        Dim q As String
        Dim montoSaldado As Double
        Dim r2 As Recordset
        Dim newEstadoSaldadoFactura As TipoSaldadoFactura

        For Each Factura In recibo.facturas
            montoSaldado = DAOFactura.PagosRealizados(Factura.id)

            If montoSaldado = 0 Then
                newEstadoSaldadoFactura = NoSaldada
            ElseIf montoSaldado >= Factura.Total Then
                newEstadoSaldadoFactura = SaldadoTotal
            Else
                newEstadoSaldadoFactura = SaldadoParcial
            End If

            If Not conectar.execute("update AdminFacturas set saldada=" & newEstadoSaldadoFactura & " where id=" & Factura.id) Then
                GoTo err5
            End If

        Next Factura

    Else
        GoTo err5
    End If





    aprobar = True
    conectar.CommitTransaction
    Exit Function
err5:
    aprobar = False
    recibo.estado = estAnt
    Set recibo.UsuarioAprobador = Nothing
    recibo.FechaAprobacion = fechaAnt
    conectar.RollBackTransaction

End Function


Public Function FindAll(Optional filter As String = "1 = 1", _
                        Optional includeRetenciones As Boolean = False, _
                        Optional includeCheques As Boolean = False, _
                        Optional includeBanco As Boolean = False, _
                        Optional includeCaja As Boolean = False, _
                        Optional includeFacturas As Boolean = False _
                        ) As Collection

    Dim q As String
    q = "SELECT *" _
        & " FROM AdminRecibos rec" _
        & " LEFT JOIN clientes cli ON cli.id = rec.idCliente" _
        & " LEFT JOIN usuarios ucre ON ucre.id = rec.idUsuarioCreador" _
        & " LEFT JOIN usuarios uapro ON uapro.id = rec.idUsuarioAprobador" _
        & " LEFT JOIN AdminConfigMonedas mon ON mon.id = rec.idMoneda" _
        & " LEFT JOIN AdminRecibosDetalleRetenciones detaret ON detaret.idRecibo = rec.id" _
        & " LEFT JOIN retenciones ret ON ret.id = detaret.idRetencion" _
        & " WHERE " & filter & " ORDER BY rec.id DESC" _



Dim col As New Collection
    Dim rec As recibo

    Dim idx As Dictionary
    Dim rs As Recordset

    Dim ret As retencionRecibo

    Set rs = conectar.RSFactory(q)
    BuildFieldsIndex rs, idx

    While Not rs.EOF
        Set rec = Map(rs, idx, "rec", "cli", "mon", "ucre", "uapro")

        'las retenciones vienen por defecto ya
        'If includeRetenciones Then
        '    Set rec.Retenciones = DAOReciboRetencion.FindAllByRecibo(rec.id)
        'End If

        If funciones.BuscarEnColeccion(col, CStr(rec.id)) Then
            Set rec = col.item(CStr(rec.id))
        End If

        Set ret = DAOReciboRetencion.Map(rs, idx, "detaret", "ret")
        If IsSomething(ret) Then
            rec.retenciones.Add ret, CStr(ret.id)
        End If


        If includeCheques Then
            Set rec.Cheques = DAOCheques.FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID & " IN (SELECT idCheque FROM AdminRecibosCheques WHERE idRecibo = " & rec.id & ")")
        End If

        If includeBanco Then
            Set rec.OperacionesBanco = DAOOperacion.FindAll(Banco, "op.id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.id & ")")
        End If

        If includeCaja Then
            Set rec.OperacionesCaja = DAOOperacion.FindAll(caja, "op.id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.id & ")")
        End If

        If includeFacturas Then
            Set rec.facturas = DAOFactura.FindAll("AdminFacturas.id IN (SELECT idFactura FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.id & ")", True, True)

            'traigo los montos pagados de cada factura
            Dim q2 As String
            q2 = "SELECT monto_pagado, idFactura FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.id
            Dim r2 As Recordset
            Set r2 = conectar.RSFactory(q2)
            Set rec.PagosDeFacturas = New Dictionary
            While Not r2.EOF
                If Not rec.PagosDeFacturas.Exists(CStr(r2!idFactura)) Then
                    rec.PagosDeFacturas.Add CStr(r2!idFactura), CDbl(r2!monto_pagado)
                End If

                r2.MoveNext
            Wend
        End If

        If Not funciones.BuscarEnColeccion(col, CStr(rec.id)) Then
            col.Add rec, CStr(rec.id)
        End If

        rs.MoveNext
    Wend

    Set FindAll = col
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaCliente As String = vbNullString, _
                    Optional tablaMoneda As String = vbNullString, _
                    Optional tablaUsuarioCreador As String = vbNullString, _
                    Optional tablaUsuarioAprobador As String = vbNullString _
                    ) As recibo

    Dim r As recibo
    Dim id As Long

    id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set r = New recibo
        r.id = id
        r.estado = GetValue(rs, indice, tabla, "estado")
        r.FechaAprobacion = GetValue(rs, indice, tabla, "fechaAprobacion")
        r.FechaCreacion = GetValue(rs, indice, tabla, "fechaCreacion")
        r.FechaModificacion = GetValue(rs, indice, tabla, "fechaModificacion")
        'r.PagoACuenta = GetValue(rs, indice, tabla, "pagoACuenta")
        r.Redondeo = GetValue(rs, indice, tabla, "redondeo")
        r.ACuenta = GetValue(rs, indice, tabla, "a_cuenta")
        r.ACuentaUsado = GetValue(rs, indice, tabla, "a_cuenta_usado")
        r.FEcha = GetValue(rs, indice, tabla, "fecha")


        Dim totEstatico As New TotalEstaticoRecibo
        totEstatico.TotalChequesEstatico = GetValue(rs, indice, tabla, "tot_estatico_cheques")
        totEstatico.TotalDepositosEstatico = GetValue(rs, indice, tabla, "tot_estatico_depositos")
        totEstatico.TotalEfectivoEstatico = GetValue(rs, indice, tabla, "tot_estatico_efectivo")
        totEstatico.TotalReciboEstatico = GetValue(rs, indice, tabla, "tot_estatico_recibo")
        Set r.TotalEstatico = totEstatico


        If LenB(tablaMoneda) > 0 Then Set r.Moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaCliente) > 0 Then Set r.Cliente = DAOCliente.Map(rs, indice, tablaCliente)
        If LenB(tablaUsuarioCreador) > 0 Then Set r.UsuarioCreador = DAOUsuarios.Map(rs, indice, tablaUsuarioCreador)
        If LenB(tablaUsuarioAprobador) > 0 Then Set r.UsuarioAprobador = DAOUsuarios.Map(rs, indice, tablaUsuarioAprobador)
    End If

    Set Map = r
End Function

Public Function Save(rec As recibo) As Boolean
    On Error GoTo E
    conectar.BeginTransaction

    If Not Guardar(rec) Then GoTo E

    Save = True
    conectar.CommitTransaction
    Exit Function
E:
    Save = False
    conectar.RollBackTransaction

End Function

Public Function Guardar(rec As recibo) As Boolean
    On Error GoTo E


    Dim q As String
    Dim reciboId As Long

    If rec.id = 0 Then

        q = "INSERT INTO AdminRecibos" _
            & "            (idCliente," _
            & "             fechaCreacion," _
            & "             idUsuarioCreador," _
            & "             fechaModificacion," _
            & "             idUsuarioModificador," _
            & "             fechaAprobacion," _
            & "             idUsuarioAprobador," _
            & "             estado," _
            & "             idMoneda," _
            & "             redondeo," _
            & "             pagoACuenta," _
            & "             totalAplicadoCuenta," _
            & "             todo_aplicado," _
    & "             fecha, a_cuenta)"
        q = q _
            & " VALUES ('idCliente'," _
            & "        'fechaCreacion'," _
            & "        'idUsuarioCreador'," _
            & "        'fechaModificacion'," _
            & "        'idUsuarioModificador'," _
            & "        'fechaAprobacion'," _
            & "        'idUsuarioAprobador'," _
            & "        'estado'," _
            & "        'idMoneda'," _
            & "        'redondeo'," _
            & "        'pagoACuenta'," _
            & "        'totalAplicadoCuenta'," _
            & "        'todo_aplicado'," _
            & "        'fecha', 'a_cuenta')"

        rec.FechaCreacion = Now
    Else

        q = "Update AdminRecibos" _
            & " SET " _
            & " idCliente = 'idCliente' ," _
            & " fechaCreacion = 'fechaCreacion' ," _
            & " idUsuarioCreador = 'idUsuarioCreador' ," _
            & " fechaModificacion = 'fechaModificacion' ," _
            & " idUsuarioModificador = 'idUsuarioModificador' ," _
            & " idUsuarioAprobador = 'idUsuarioAprobador' ," _
            & " fechaAprobacion = 'fechaAprobacion' ," _
            & " estado = 'estado' ," _
            & " idMoneda = 'idMoneda' ," _
            & " redondeo = 'redondeo' ," _
            & " pagoACuenta = 'pagoACuenta' ," _
            & " fecha = 'fecha', a_cuenta = 'a_cuenta'," _
            & " tot_estatico_cheques='tot_estatico_cheques', tot_estatico_efectivo = 'tot_estatico_efectivo'," _
            & " tot_estatico_depositos='tot_estatico_depositos', tot_estatico_recibo = 'tot_estatico_recibo'" _
            & " Where id = 'id'"

        '& " totalAplicadoCuenta = 'totalAplicadoCuenta' ," _
         '& " todo_aplicado = 'todo_aplicado' ," _

         rec.FechaModificacion = Now

        q = Replace(q, "'idUsuarioAprobador'", conectar.GetEntityId(rec.UsuarioAprobador))
        q = Replace(q, "'id'", conectar.GetEntityId(rec))
        q = Replace(q, "'idUsuarioModificador'", funciones.getUser)
        q = Replace(q, "'fechaAprobacion'", conectar.Escape(rec.FechaAprobacion))

        If IsSomething(rec.TotalEstatico) Then
            q = Replace(q, "'tot_estatico_cheques'", conectar.Escape(rec.TotalEstatico.TotalChequesEstatico))
            q = Replace(q, "'tot_estatico_efectivo'", conectar.Escape(rec.TotalEstatico.TotalEfectivoEstatico))
            q = Replace(q, "'tot_estatico_depositos'", conectar.Escape(rec.TotalEstatico.TotalDepositosEstatico))
            q = Replace(q, "'tot_estatico_recibo'", conectar.Escape(rec.TotalEstatico.TotalReciboEstatico))
        Else
            q = Replace(q, "'tot_estatico_cheques'", 0)
            q = Replace(q, "'tot_estatico_efectivo'", 0)
            q = Replace(q, "'tot_estatico_depositos'", 0)
            q = Replace(q, "'tot_estatico_recibo'", 0)
        End If


    End If

    q = Replace(q, "'idCliente'", conectar.GetEntityId(rec.Cliente))
    q = Replace(q, "'fechaCreacion'", conectar.Escape(rec.FechaCreacion))
    q = Replace(q, "'idUsuarioCreador'", conectar.GetEntityId(rec.UsuarioCreador))
    q = Replace(q, "'fechaModificacion'", conectar.Escape(rec.FechaModificacion))
    q = Replace(q, "'estado'", conectar.Escape(rec.estado))
    q = Replace(q, "'idMoneda'", conectar.GetEntityId(rec.Moneda))
    q = Replace(q, "'redondeo'", conectar.Escape(rec.Redondeo))
    'q = Replace(q, "'pagoACuenta'", conectar.Escape(rec.PagoACuenta))
    q = Replace(q, "'fecha'", conectar.Escape(rec.FEcha))
    q = Replace(q, "'a_cuenta'", conectar.Escape(rec.ACuenta))

    Dim esNuevo As Boolean
    esNuevo = False
    If Not conectar.execute(q) Then GoTo E
    If rec.id = 0 Then esNuevo = True
    If rec.id <> 0 And rec.estado = pendiente Then  'en el insert no tiene nada de agregacion

        'retenciones----------------------------------------------------------
        q = "idRecibo = " & rec.id
        If Not DAOReciboRetencion.Delete(q) Then GoTo E
        Dim ret As retencionRecibo
        For Each ret In rec.retenciones
            If Not DAOReciboRetencion.Save(ret, rec) Then GoTo E
        Next ret

        'cheques----------------------------------------------------------
        q = "DELETE FROM AdminRecibosCheques WHERE idRecibo = " & rec.id
        If Not conectar.execute(q) Then GoTo E
        Dim cheq As cheque
        For Each cheq In rec.Cheques

            If cheq.id = 0 Then
                cheq.EnCartera = True
                cheq.Propio = False
                cheq.OrigenDestino = UCase(rec.Cliente.razon)
            Else
                'If IsSomething(DAOCheques.FindById(cheq.id)) Then
                '    q = "DELETE FROM Cheques WHERE id = " & cheq.id
                '    If Not conectar.execute(q) Then GoTo E
                'End If
            End If

            If Not DAOCheques.Guardar(cheq) Then GoTo E

            q = "INSERT INTO AdminRecibosCheques (idRecibo, idCheque) VALUES (" & rec.id & ", " & cheq.id & ")"
            If Not conectar.execute(q) Then GoTo E
        Next cheq


        'facturas''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        q = "DELETE FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.id
        If Not conectar.execute(q) Then GoTo E
        Dim fac As Factura
        Dim montoPagado As Double
        For Each fac In rec.facturas
            If rec.PagosDeFacturas.Exists(CStr(fac.id)) Then
                montoPagado = rec.PagosDeFacturas.item(CStr(fac.id))
            Else
                montoPagado = 0
            End If

            q = "INSERT INTO AdminRecibosDetalleFacturas (idRecibo, idFactura, monto_pagado) VALUES (" & rec.id & ", " & fac.id & ", " & Escape(montoPagado) & ")"
            If Not conectar.execute(q) Then GoTo E
        Next fac



        If Not DAOOperacion.Delete("id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.id & ")") Then GoTo E
        If Not conectar.execute("DELETE FROM operaciones_recibos WHERE reciboId = " & rec.id) Then GoTo E

        '''''''''''''''''''''''''''''CAJA
        Dim op As operacion
        Dim recId As Long
        For Each op In rec.OperacionesCaja
            If Not DAOOperacion.Save(op) Then GoTo E
            conectar.UltimoId "operaciones", recId
            If recId = 0 Then GoTo E
            If Not conectar.execute("INSERT INTO operaciones_recibos VALUES (" & recId & "," & rec.id & ")") Then GoTo E
        Next op

        '''''''''''''''''''''''''''''BANCO
        For Each op In rec.OperacionesBanco
            If Not DAOOperacion.Save(op) Then GoTo E
            conectar.UltimoId "operaciones", recId
            If recId = 0 Then GoTo E
            If Not conectar.execute("INSERT INTO operaciones_recibos VALUES (" & recId & "," & rec.id & ")") Then GoTo E
        Next op

    End If


    Dim EVENTO As New clsEventoObserver
    Set EVENTO.Elemento = rec

    If esNuevo Then
        EVENTO.EVENTO = agregar_
    Else
        EVENTO.EVENTO = modificar_
    End If


    Set EVENTO.Originador = Nothing
    EVENTO.Tipo = Recibos_

    Channel.Notificar EVENTO, Recibos_

    Guardar = True

    Exit Function
E:
    Guardar = False


End Function


Public Sub Imprimir(idRecibo As Long)

    Dim recibo As recibo
    Set recibo = DAORecibo.FindById(idRecibo, True, True, True, True, True)

    If IsSomething(recibo) Then
        '        Printer.CurrentY = 300
        '        Printer.CurrentX = 6800
        '        Printer.Font.Size = 12
        '        Printer.Font.Size = 14
        '        Printer.Line (8800, 1400)-(10100, 1400)
        '        Printer.Print Format(Day(objFac.FechaEmision), "00") & "/" & Format(Month(objFac.FechaEmision), "00") & "/" & Format(Year(objFac.FechaEmision) - 2000, "00")
        '        Printer.Print Tab(14);
        Dim origin As Integer

        Printer.FontBold = True
        origin = Printer.FontSize
        Printer.FontSize = origin + 5

        Dim cx As Integer
        Printer.Print "SIGNO PLAST S.A."
        Printer.FontSize = origin + 3
        Printer.Print "Número: " & recibo.id
        Printer.Print "Estado: " & enums.EnumEstadoRecibo(recibo.estado)
        Printer.Print "Fecha: " & Format(Day(recibo.FEcha), "00") & "/" & Format(Month(recibo.FEcha), "00") & "/" & Format(Year(recibo.FEcha), "0000")
        Printer.Print "Cliente: " & recibo.Cliente.razon
        Printer.FontSize = origin
        Printer.FontBold = False
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
        Printer.Print Chr(10)


        Printer.FontBold = True
        Printer.Print "Facturas "
        Printer.FontBold = False
        Dim F As Factura
        For Each F In recibo.facturas
            Printer.Print F.FechaEmision, F.GetShortDescription(False, True), F.Moneda.NombreCorto & " " & recibo.PagosDeFacturas(CStr(F.id))
        Next F

        If recibo.facturas.count > 0 Then
            Printer.Print "Total Facturas: " & recibo.TotalFacturas
        End If
        Printer.Print Chr(10)

        Printer.FontBold = True

        Printer.Print "Retenciones "
        Printer.FontBold = False
        Dim r As retencionRecibo
        For Each r In recibo.retenciones
            Printer.Print r.FEcha, r.Retencion.nombre, r.NroRetencion, r.Valor
        Next r
        If recibo.retenciones.count > 0 Then
            Printer.Print "Total Retenciones: " & recibo.TotalRetenciones
        End If
        Printer.Print Chr(10)
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
        Printer.Print Chr(10)

        Printer.FontBold = True
        Printer.Print "Valores recibidos "
        Printer.FontBold = False
        If recibo.OperacionesBanco.count > 0 Then
            Printer.FontBold = True
            Printer.Print "Banco"
            Printer.FontBold = False
        Else
            Printer.Print "Sin operaciones de banco"
        End If

        Dim o As operacion
        For Each o In recibo.OperacionesBanco
            Printer.Print o.FechaOperacion, o.CuentaBancaria.DescripcionFormateada, o.Monto
        Next o

        If recibo.OperacionesBanco.count > 0 Then
            Printer.Print "Total Banco: " & recibo.TotalOperacionesBanco
        End If

        If recibo.OperacionesCaja.count > 0 Then
            Printer.FontBold = True
            Printer.Print "Caja"
            Printer.FontBold = False
        Else

            Printer.Print "Sin operaciones de caja"

        End If

        For Each o In recibo.OperacionesCaja
            Printer.Print o.FechaOperacion, o.CuentaBancaria.DescripcionFormateada, o.Monto
        Next o
        Printer.FontBold = True
        If recibo.OperacionesCaja.count > 0 Then
            Printer.Print "Total Caja: " & recibo.TotalOperacionesCaja
        End If
        Printer.Print Chr(10)
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)

        Printer.Print " Total Recibo:  " & recibo.Total
        Printer.Print " Total Recibido:  " & recibo.TotalRecibido
        Printer.FontBold = False
        Printer.EndDoc
    End If

End Sub


