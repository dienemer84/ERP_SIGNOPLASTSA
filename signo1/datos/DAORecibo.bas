Attribute VB_Name = "DAORecibo"
Option Explicit

Public Function FindById(Id As Long, _
                         Optional includeRetenciones As Boolean = False, _
                         Optional includeCheques As Boolean = False, _
                         Optional includeBanco As Boolean = False, _
                         Optional includeCaja As Boolean = False, _
                         Optional includeFacturas As Boolean = False _
                       ) As recibo

    Dim col As Collection
    Set col = FindAll("rec.id = " & Id, includeRetenciones, includeCheques, includeBanco, includeCaja, includeFacturas)
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

    Err.Raise 9999, , "Funcionalidad en desarrollo"

    conectar.BeginTransaction

    If recibo.estado = EstadoRecibo.Aprobado Then
        'cambio el estado del recibo
        recibo.estado = EstadoRecibo.Reciboanulado



        'borro los cheques
        If Not conectar.execute("DELETE FROM Cheques WHERE id IN (SELECT idCheque FROM AdminRecibosCheques a WHERE a.`idRecibo`=" & recibo.Id & ")") Then GoTo err101

        'borro los cheques x recibo
        If Not conectar.execute("DELETE FROM AdminRecibosCheques WHERE idRecibo=" & recibo.Id) Then GoTo err101


        'borro las operaciones
        If Not conectar.execute("DELETE FROM `AdminRecibosDepositos` WHERE idRecibo=" & recibo.Id) Then GoTo err101
        'DELETE FROM `AdminRecibosDepositos` WHERE idRecibo=5331


        'borro las facturas
        'if Not conectar.execute("DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo= " & recibo.Id) Then GoTo err101
        'DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo=5311


        Dim q As String
        q = "select * from AdminRecibosDetalleFacturas where idRecibo=" & recibo.Id
        Dim rs As Recordset
        Set rs = conectar.RSFactory(q)
        Dim F As Factura
        Dim rs2 As Recordset
        While Not rs.EOF And Not rs.BOF

            q = "SELECT * FROM `AdminRecibosDetalleFacturas` f WHERE f.`idFactura`= " & rs!idFactura & "  AND f.`idRecibo`<>" & recibo.Id

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
        If Not conectar.execute("DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo= " & recibo.Id) Then GoTo err101



        'borro retencione
        If Not conectar.execute("DELETE FROM `AdminRecibosDetalleRetenciones` WHERE idRecibo=" & recibo.Id) Then GoTo err101
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
    Set recibo.usuarioAprobador = funciones.GetUserObj
    recibo.estado = EstadoRecibo.Aprobado


    If recibo.IsValid Then
        'totalizo recibo
        Dim totEst As New TotalEstaticoRecibo
        totEst.TotalChequesEstatico = recibo.TotalCheques
        totEst.TotalDepositosEstatico = recibo.TotalOperacionesBanco
        totEst.TotalEfectivoEstatico = recibo.TotalOperacionesCaja
        totEst.TotalReciboEstatico = recibo.total
        Set recibo.TotalEstatico = totEst

        If Not DAORecibo.Guardar(recibo) Then GoTo err5

        Dim q As String
        Dim montoSaldado As Double
        Dim r2 As Recordset
        Dim newEstadoSaldadoFactura As TipoSaldadoFactura

        For Each Factura In recibo.facturas
            montoSaldado = DAOFactura.PagosRealizados(Factura.Id)

            If montoSaldado = 0 Then
                newEstadoSaldadoFactura = NoSaldada
            ElseIf montoSaldado >= Factura.total Then
                newEstadoSaldadoFactura = saldadoTotal
            Else
                newEstadoSaldadoFactura = SaldadoParcial
            End If

            If Not conectar.execute("update AdminFacturas set saldada=" & newEstadoSaldadoFactura & " where id=" & Factura.Id) Then
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
    Set recibo.usuarioAprobador = Nothing
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

        If funciones.BuscarEnColeccion(col, CStr(rec.Id)) Then
            Set rec = col.item(CStr(rec.Id))
        End If

        Set ret = DAOReciboRetencion.Map(rs, idx, "detaret", "ret")
        If IsSomething(ret) Then
            rec.retenciones.Add ret, CStr(ret.Id)
        End If


        If includeCheques Then
            Set rec.Cheques = DAOCheques.FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID & " IN (SELECT idCheque FROM AdminRecibosCheques WHERE idRecibo = " & rec.Id & ")")
        End If

        If includeBanco Then
            Set rec.OperacionesBanco = DAOOperacion.FindAll(Banco, "op.id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.Id & ")")
        End If

        If includeCaja Then
            Set rec.OperacionesCaja = DAOOperacion.FindAll(caja, "op.id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.Id & ")")
        End If

        If includeFacturas Then
            Set rec.facturas = DAOFactura.FindAll("AdminFacturas.id IN (SELECT idFactura FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.Id & ")", True, True)

            'traigo los montos pagados de cada factura
            Dim q2 As String
            q2 = "SELECT monto_pagado, idFactura FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.Id
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

        If Not funciones.BuscarEnColeccion(col, CStr(rec.Id)) Then
            col.Add rec, CStr(rec.Id)
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
    Dim Id As Long

    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set r = New recibo
        r.Id = Id
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


        If LenB(tablaMoneda) > 0 Then Set r.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaCliente) > 0 Then Set r.cliente = DAOCliente.Map(rs, indice, tablaCliente)
        If LenB(tablaUsuarioCreador) > 0 Then Set r.usuarioCreador = DAOUsuarios.Map(rs, indice, tablaUsuarioCreador)
        If LenB(tablaUsuarioAprobador) > 0 Then Set r.usuarioAprobador = DAOUsuarios.Map(rs, indice, tablaUsuarioAprobador)
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

    If rec.Id = 0 Then

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

        q = Replace(q, "'idUsuarioAprobador'", conectar.GetEntityId(rec.usuarioAprobador))
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

    q = Replace(q, "'idCliente'", conectar.GetEntityId(rec.cliente))
    q = Replace(q, "'fechaCreacion'", conectar.Escape(rec.FechaCreacion))
    q = Replace(q, "'idUsuarioCreador'", conectar.GetEntityId(rec.usuarioCreador))
    q = Replace(q, "'fechaModificacion'", conectar.Escape(rec.FechaModificacion))
    q = Replace(q, "'estado'", conectar.Escape(rec.estado))
    q = Replace(q, "'idMoneda'", conectar.GetEntityId(rec.moneda))
    q = Replace(q, "'redondeo'", conectar.Escape(rec.Redondeo))
    'q = Replace(q, "'pagoACuenta'", conectar.Escape(rec.PagoACuenta))
    q = Replace(q, "'fecha'", conectar.Escape(rec.FEcha))
    q = Replace(q, "'a_cuenta'", conectar.Escape(rec.ACuenta))

    Dim esNuevo As Boolean
    esNuevo = False
    If Not conectar.execute(q) Then GoTo E

    'ACA ES DONDE REVISA SI ES UN RECIBO NUEVO O NO. Y EN QUE ESTADO ESTÁ
    If rec.Id = 0 Then esNuevo = True

    'If rec.id <> 0 And rec.estado = EstadoRecibo.Pendiente Then  'en el insert no tiene nada de agregacion

    If rec.Id <> 0 Then   'en el insert no tiene nada de agregacion

        'retenciones----------------------------------------------------------
        q = "idRecibo = " & rec.Id
        If Not DAOReciboRetencion.Delete(q) Then GoTo E
        Dim ret As retencionRecibo
        For Each ret In rec.retenciones
            If Not DAOReciboRetencion.Save(ret, rec) Then GoTo E
        Next ret

        'cheques----------------------------------------------------------
        q = "DELETE FROM AdminRecibosCheques WHERE idRecibo = " & rec.Id
        If Not conectar.execute(q) Then GoTo E
        Dim cheq As cheque
        For Each cheq In rec.Cheques

            If cheq.Id = 0 Then
                cheq.EnCartera = True
                cheq.Propio = False
                cheq.OrigenDestino = UCase(rec.cliente.razon)
            Else
                'If IsSomething(DAOCheques.FindById(cheq.id)) Then
                '    q = "DELETE FROM Cheques WHERE id = " & cheq.id
                '    If Not conectar.execute(q) Then GoTo E
                'End If
            End If

            If Not DAOCheques.Guardar(cheq) Then GoTo E

            q = "INSERT INTO AdminRecibosCheques (idRecibo, idCheque) VALUES (" & rec.Id & ", " & cheq.Id & ")"
            If Not conectar.execute(q) Then GoTo E
        Next cheq


        'facturas''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        q = "DELETE FROM AdminRecibosDetalleFacturas WHERE idRecibo = " & rec.Id
        If Not conectar.execute(q) Then GoTo E
        Dim fac As Factura
        Dim montoPagado As Double
        For Each fac In rec.facturas
            If rec.PagosDeFacturas.Exists(CStr(fac.Id)) Then
                montoPagado = rec.PagosDeFacturas.item(CStr(fac.Id))
            Else
                montoPagado = 0
            End If

            q = "INSERT INTO AdminRecibosDetalleFacturas (idRecibo, idFactura, monto_pagado) VALUES (" & rec.Id & ", " & fac.Id & ", " & Escape(montoPagado) & ")"
            If Not conectar.execute(q) Then GoTo E
        Next fac



        If Not DAOOperacion.Delete("id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.Id & ")") Then GoTo E
        If Not conectar.execute("DELETE FROM operaciones_recibos WHERE reciboId = " & rec.Id) Then GoTo E

        '''''''''''''''''''''''''''''CAJA
        Dim op As operacion
        Dim recId As Long
        For Each op In rec.OperacionesCaja
            If Not DAOOperacion.Save(op) Then GoTo E
            conectar.UltimoId "operaciones", recId
            If recId = 0 Then GoTo E
            If Not conectar.execute("INSERT INTO operaciones_recibos VALUES (" & recId & "," & rec.Id & ")") Then GoTo E
        Next op

        '''''''''''''''''''''''''''''BANCO
        For Each op In rec.OperacionesBanco
            If Not DAOOperacion.Save(op) Then GoTo E
            conectar.UltimoId "operaciones", recId
            If recId = 0 Then GoTo E
            If Not conectar.execute("INSERT INTO operaciones_recibos VALUES (" & recId & "," & rec.Id & ")") Then GoTo E
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
        
        Dim origin As Integer

        Printer.FontBold = True
        origin = Printer.FontSize
        Printer.FontSize = origin + 5

        Dim cx As Integer
        Printer.Print "SIGNO PLAST S.A."
        Printer.FontSize = origin + 3
        Printer.Print "Número: " & recibo.Id
        Printer.Print "Estado: " & enums.EnumEstadoRecibo(recibo.estado)
        Printer.Print "Fecha: " & Format(Day(recibo.FEcha), "00") & "/" & Format(Month(recibo.FEcha), "00") & "/" & Format(Year(recibo.FEcha), "0000")
        Printer.Print "Cliente: " & recibo.cliente.razon
        Printer.FontSize = origin
        Printer.FontBold = False
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
        Printer.Print Chr(10)


        Printer.FontBold = True
        Printer.Print "Facturas "
        Printer.FontBold = False
        Dim F As Factura
        For Each F In recibo.facturas
            Printer.Print F.FechaEmision, F.GetShortDescription(False, True), F.moneda.NombreCorto & " " & recibo.PagosDeFacturas(CStr(F.Id))
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
            Printer.FontBold = True
            Printer.Print "Total Banco: " & recibo.TotalOperacionesBanco
        End If

        Printer.Print Chr(10)

        If recibo.OperacionesCaja.count > 0 Then
            Printer.FontBold = True
            Printer.Print "Caja"
            Printer.FontBold = False
        Else

            Printer.Print "Sin operaciones de caja"

        End If



        For Each o In recibo.OperacionesCaja
            'Printer.Print o.FechaOperacion, o.CuentaBancaria.DescripcionFormateada, o.Monto
            Printer.Print o.FechaOperacion, o.Monto
        Next o

        Printer.FontBold = True
        If recibo.OperacionesCaja.count > 0 Then
            Printer.Print "Total Caja: " & recibo.TotalOperacionesCaja
        End If

        Printer.Print Chr(10)

        If recibo.Cheques.count > 0 Then
            Printer.FontBold = True
            Printer.Print "Cheques Recibidos"
            Printer.FontBold = False
        Else

            Printer.Print "Sin cheques recibidos."

        End If

        Printer.Print "Numero,"; vbTab; "Monto"; vbTab; vbTab; "Fecha Vto."; vbTab; "Banco"

        Dim che As cheque
        For Each che In recibo.Cheques

            'Printer.Print o.FechaOperacion, o.CuentaBancaria.DescripcionFormateada, o.Monto

            Printer.Print che.numero, che.Monto, che.FechaVencimiento, che.Banco.nombre
        Next che

        Printer.FontBold = True
        If recibo.Cheques.count > 0 Then
            Printer.Print "Total Cheques Recibidos: " & recibo.TotalCheques
        End If
        Printer.Print Chr(10)
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)



        Printer.Print " Total Recibo:  " & recibo.total
        Printer.Print " Total Recibido:  " & recibo.TotalRecibido
        Printer.FontBold = False
        Printer.EndDoc
    End If

End Sub

Public Function ExportarColeccion(col As Collection, Optional ProgressBar As Object) As Boolean
    On Error GoTo err1

    ExportarColeccion = True

    '    Dim detalle As DetalleOrdenTrabajo
    '    Dim Entregas As Collection
    '    Dim remitoDetalle As remitoDetalle

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlWorkbook As Object
    Set xlWorkbook = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xlWorksheet As Object
    Set xlWorksheet = CreateObject("Excel.Application")

    'Dim xlApplication As New Excel.Application
    Dim xlApplication As Object
    Set xlApplication = CreateObject("Excel.Application")

    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    'fila, columna

    Dim offset As Long
    offset = 3
    xlWorksheet.Cells(offset, 1).value = "Número"
    xlWorksheet.Cells(offset, 2).value = "Fecha Emisión"
    xlWorksheet.Cells(offset, 3).value = "Cliente"
    xlWorksheet.Cells(offset, 4).value = "Fecha Creación"
    xlWorksheet.Cells(offset, 5).value = "Moneda"
    xlWorksheet.Cells(offset, 6).value = "Total Recibo"
    xlWorksheet.Cells(offset, 7).value = "Total Recibido"
    xlWorksheet.Cells(offset, 8).value = "Saldo a Cuenta"
    xlWorksheet.Cells(offset, 9).value = "Estado"


    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 9)).Font.Bold = True
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 9)).Interior.Color = &HC0C0C0

    Dim rec As recibo
    Dim fac As clsFacturaProveedor

    Dim initoffset As Long
    initoffset = offset

    ProgressBar.min = 0
    ProgressBar.max = col.count


    Dim d As Long
    d = 0

    For Each rec In col

        Dim i As Integer

        i = 1

        d = d + 1
        ProgressBar.value = d

        offset = offset + 1

        xlWorksheet.Cells(offset, 1).value = rec.Id
        xlWorksheet.Cells(offset, 2).value = Format(rec.FEcha, "yyyy/mm/dd", vbSunday)
        xlWorksheet.Cells(offset, 3).value = rec.cliente.razon
        xlWorksheet.Cells(offset, 4).value = rec.FechaCreacion
        xlWorksheet.Cells(offset, 5).value = rec.moneda.NombreCorto
        xlWorksheet.Cells(offset, 6).value = rec.TotalEstatico.TotalReciboEstatico
        xlWorksheet.Cells(offset, 7).value = rec.TotalEstatico.TotalRecibidoEstatico
        xlWorksheet.Cells(offset, 8).value = rec.ACuentaDisponible

        Select Case rec.estado
        Case 1
            xlWorksheet.Cells(offset, 9).value = "Pendiente"
        Case 2
            xlWorksheet.Cells(offset, 9).value = "Aprobado"
        Case 3
            xlWorksheet.Cells(offset, 9).value = "Anulado"
        End Select


        '        xlWorksheet.Cells(offset, 9).value = rec.EstadoRecibo
        '    Pendiente = 1
        '    Aprobado = 2
        '    Reciboanulado = 3
        'End Enum


    Next

    xlWorksheet.Range(xlWorksheet.Cells(initoffset, 1), xlWorksheet.Cells(offset, 9)).Borders.LineStyle = xlContinuous

    'autosize
    xlApplication.ScreenUpdating = False
    Dim wkSt As String
    wkSt = xlWorksheet.Name
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlWorkbook.Sheets(wkSt).Select
    xlApplication.ScreenUpdating = True
    ''

    Dim ruta As String
    ruta = Environ$("TEMP")
    If LenB(ruta) = 0 Then ruta = Environ$("TMP")
    If LenB(ruta) = 0 Then ruta = App.path
    ruta = ruta & "\" & funciones.CreateGUID() & ".xls"

    xlWorkbook.SaveAs ruta

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit

    ShellExecute -1, "open", ruta, "", "", 4

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    ProgressBar.value = 0

    Exit Function
err1:
    ExportarColeccion = False
End Function

Public Function ResumenPagos(ByRef Cheques As Collection, ByRef caja As Collection, ByRef bancos As Collection, ByRef comp As Collection, ByRef retenciones As Collection, ByRef cheques3 As Collection, Optional filtro As String, Optional idProveedor As Long = -1) As Boolean
    On Error GoTo err1
    ResumenPagos = True
    Dim q As String
    Dim rs As Recordset

    '#'CHEQUES'
    q = "SELECT b.Nombre,SUM(monto * acm.cambio) as monto FROM AdminRecibos rec " _
      & " INNER JOIN AdminRecibosCheques arc ON arc.idRecibo = rec.id " _
      & " LEFT JOIN Cheques c ON c.id = arc.idCheque " _
      & " LEFT JOIN AdminConfigBancos b ON c.id_banco=b.id " _
      & " LEFT JOIN AdminConfigMonedas acm ON c.id_moneda=acm.id WHERE c.propio=1 AND rec.estado = 2 and 1=1 " _

If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    q = q & " GROUP BY b.id "

    Dim d As DTONombreMonto

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        d.nombre = rs!nombre
        Cheques.Add d
        rs.MoveNext
    Wend

   '#OPERACIONES CAJA
    q = " SELECT ca.nombre,SUM(monto * acm.cambio ) as monto FROM AdminRecibos rec " _
      & " INNER JOIN operaciones_recibos opr ON opr.reciboId=rec.id " _
      & " LEFT JOIN operaciones o ON o.id=opr.operacionId " _
      & " LEFT JOIN cajas ca ON ca.id=o.cuentabanc_o_caja_id " _
      & " LEFT JOIN AdminConfigMonedas acm ON o.moneda_id=acm.id " _
      & " WHERE o.pertenencia='caja' AND entrada_salida=1 AND rec.estado = 2 AND 1=1 "
      
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    q = q & " GROUP BY o.cuentabanc_o_caja_id"

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        
'        d.nombre = rs!nombre
        
        If Not IsNull(rs!nombre) Then
            d.nombre = rs!nombre
        Else
            d.nombre = ""
        End If
        
        caja.Add d
        rs.MoveNext
    Wend


    '
    '#OPERACIONES BANCO
    q = "SELECT  ba.nombre, SUM(monto * acm.cambio ) AS monto " _
      & " FROM AdminRecibos rec INNER JOIN operaciones_recibos opr ON opr.reciboId=rec.id " _
      & " LEFT JOIN operaciones o ON o.id=opr.operacionId " _
      & " LEFT JOIN AdminConfigCuentas cba ON cba.id = o.cuentabanc_o_caja_id " _
      & " INNER JOIN AdminConfigBancos ba ON cba.idBanco=ba.id LEFT JOIN AdminConfigMonedas acm ON o.moneda_id = acm.id " _
      & " WHERE o.pertenencia='banco' AND entrada_salida= 1 AND 1=1 AND rec.estado = 2"
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    q = q & "GROUP BY o.cuentabanc_o_caja_id"
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto

        If Not IsNull(rs!nombre) Then d.nombre = rs!nombre Else d.nombre = vbNullString

        bancos.Add d
        rs.MoveNext
    Wend

''#COMPENSATORIOS
'
'    q = "SELECT fp.numero_factura, (IF (com.tipo=1,(com.importe * acm.cambio),(com.importe * acm.cambio*-1))) AS monto  FROM ordenes_pago op " _
'      & " INNER JOIN ordenes_pago_compensatorios com ON com.id_orden_pago=op.id " _
'      & " INNER JOIN AdminComprasFacturasProveedores fp ON com.id_comprobante=fp.id " _
'      & " INNER JOIN AdminConfigMonedas acm ON fp.id_moneda=acm.id  where 1=1 "
'
'    If LenB(filtro) > 0 Then
'        q = q & " and " & filtro
'    End If
'
'
'    Set rs = conectar.RSFactory(q)
'    While Not rs.EOF And Not rs.BOF
'        Set d = New DTONombreMonto
'        d.Monto = rs!Monto
'        If Not IsNull(rs!numero_factura) Then d.nombre = rs!numero_factura Else d.nombre = vbNullString
'
'        comp.Add d
'        rs.MoveNext
'    Wend

'#RETENCIONES

'    q = "SELECT 'IIBB' AS nombre, SUM(static_total_a_retener) AS monto FROM ordenes_pago op WHERE 1=1"
    
    q = "SELECT 'Retenciones' AS nombre, SUM(recr.valor) AS monto" _
        & " FROM AdminRecibosDetalleRetenciones recr" _
        & " INNER JOIN AdminRecibos rec ON rec.id = recr.idRecibo" _
        & " WHERE 1=1 AND rec.estado = 2"

    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If

    'q = q & " GROUP BY id "

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        
        If Not IsNull(rs!Monto) Then d.Monto = rs!Monto Else d.Monto = 0
                
'        d.Monto = rs!Monto
        d.nombre = rs!nombre
        retenciones.Add d
        rs.MoveNext
    Wend


    '#'CHEQUES 3ROS'
    q = "SELECT b.Nombre, SUM(monto * acm.cambio) AS monto FROM AdminRecibos rec " _
      & " INNER JOIN AdminRecibosCheques arc ON arc.idRecibo = rec.id " _
      & " LEFT JOIN Cheques c ON c.id = arc.idCheque " _
      & " LEFT JOIN AdminConfigBancos b ON c.id_banco=b.id " _
      & " LEFT JOIN AdminConfigMonedas acm ON c.id_moneda=acm.id WHERE c.propio=0 and 1=1 AND rec.estado = 2 " _

If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    q = q & " GROUP BY b.id "



    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        Set d = New DTONombreMonto
        d.Monto = rs!Monto
        If Not IsNull(rs!nombre) Then d.nombre = rs!nombre Else d.nombre = vbNullString
        cheques3.Add d
        rs.MoveNext
    Wend

    Exit Function
err1:
    ResumenPagos = False
    MsgBox ("Consulte al administrador del sistema porque ha ocurrido un error en la búsqueda de datos")
End Function


