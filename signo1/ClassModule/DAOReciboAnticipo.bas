Attribute VB_Name = "DAOReciboAnticipo"
Option Explicit

Public Function FindById(Id As Long, _
                         Optional includeRetenciones As Boolean = False, _
                         Optional includeCheques As Boolean = False, _
                         Optional includeBanco As Boolean = False, _
                         Optional includeCaja As Boolean = False, _
                         Optional includeFacturas As Boolean = False _
                         ) As Recibo

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
q = "select max(id)+1 as ultimo from AdminRecibosAnticipo"
Set rs = conectar.RSFactory(q)
If Not rs.EOF And Not rs.BOF Then
proximo = rs!ultimo
End If
Exit Function
err1:
proximo = -1

End Function

Public Function Anular(Recibo As Recibo) As Boolean
    
    Dim estadoAnterior As EstadoRecibo
    Dim numeroError As Long
    Dim descripcionError As String
    
    conectar.BeginTransaction

    If Recibo.estado = EstadoRecibo.Aprobado Then
        'cambio el estado del recibo
        Recibo.estado = EstadoRecibo.Reciboanulado



        'borro los cheques
        If Not conectar.execute("DELETE FROM Cheques WHERE id IN (SELECT idCheque FROM AdminRecibosCheques a WHERE a.`idRecibo`=" & Recibo.Id & ")") Then GoTo err101

        'borro los cheques x recibo
        If Not conectar.execute("DELETE FROM AdminRecibosCheques WHERE idRecibo=" & Recibo.Id) Then GoTo err101


        'borro las operaciones
        If Not conectar.execute("DELETE FROM `AdminRecibosDepositos` WHERE idRecibo=" & Recibo.Id) Then GoTo err101
        'DELETE FROM `AdminRecibosDepositos` WHERE idRecibo=5331


        'borro las facturas
        'if Not conectar.execute("DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo= " & recibo.Id) Then GoTo err101
        'DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo=5311


        Dim q As String
        q = "select * from AdminRecibosDetalleFacturas where idRecibo=" & Recibo.Id
        Dim rs As Recordset
        Set rs = conectar.RSFactory(q)
        Dim F As Factura
        Dim rs2 As Recordset
        While Not rs.EOF And Not rs.BOF

            q = "SELECT * FROM `AdminRecibosDetalleFacturas` f WHERE f.`idFactura`= " & rs!idFactura & "  AND f.`idRecibo`<>" & Recibo.Id

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
        If Not conectar.execute("DELETE FROM `AdminRecibosDetalleFacturas` WHERE idRecibo= " & Recibo.Id) Then GoTo err101



        'borro retencione
        If Not conectar.execute("DELETE FROM `AdminRecibosDetalleRetenciones` WHERE idRecibo=" & Recibo.Id) Then GoTo err101
        'DELETE FROM `AdminRecibosDetalleRetenciones` WHERE idRecibo=5331



        'libero los comprobasntes
        
    If Not DAORecibo.Guardar(Recibo) Then
        GoTo err101
    End If
    
    conectar.CommitTransaction
    
    Anular = True
    Exit Function

    Else
        GoTo err101


    End If
    Exit Function
err101:

    On Error Resume Next

    Recibo.estado = estadoAnterior

    conectar.RollBackTransaction

    Anular = False

    MsgBox _
        "Error al anular el recibo." & _
        vbCrLf & vbCrLf & _
        Err.Number & " - " & Err.Description, _
        vbCritical, _
        "Anular recibo"

End Function


Public Function aprobar(Recibo As Recibo) As Boolean
    On Error GoTo err5
    
    Dim estAnt As EstadoRecibo
    Dim fechaAnt As Variant
    Dim Factura As Factura
    
    estAnt = Recibo.estado
    fechaAnt = Recibo.FechaAprobacion
    
    conectar.BeginTransaction
    
    Recibo.FechaAprobacion = Now
    Set Recibo.usuarioAprobador = funciones.GetUserObj
    Recibo.estado = EstadoRecibo.Aprobado

    If Recibo.IsValid Then
        'totalizo recibo
        Dim totEst As New TotalEstaticoRecibo
        totEst.TotalChequesEstatico = Recibo.TotalCheques
        totEst.TotalDepositosEstatico = Recibo.TotalOperacionesBanco
        totEst.TotalEfectivoEstatico = Recibo.TotalOperacionesCaja
        totEst.TotalReciboEstatico = Recibo.total
        Set Recibo.TotalEstatico = totEst

        If Not DAOReciboAnticipo.Guardar(Recibo) Then GoTo err5

        Dim q As String
        Dim montoSaldado As Double
        Dim r2 As Recordset
        Dim newEstadoSaldadoFactura As TipoSaldadoFactura

        For Each Factura In Recibo.facturas
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
    Recibo.estado = estAnt
    Set Recibo.usuarioAprobador = Nothing
    Recibo.FechaAprobacion = fechaAnt
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
        & " FROM AdminRecibosAnticipo rec" _
        & " LEFT JOIN clientes cli ON cli.id = rec.idCliente" _
        & " LEFT JOIN usuarios ucre ON ucre.id = rec.idUsuarioCreador" _
        & " LEFT JOIN usuarios uapro ON uapro.id = rec.idUsuarioAprobador" _
        & " LEFT JOIN AdminConfigMonedas mon ON mon.id = rec.idMoneda" _
        & " LEFT JOIN AdminRecibosDetalleRetenciones detaret ON detaret.idRecibo = rec.id" _
        & " LEFT JOIN retenciones ret ON ret.id = detaret.idRetencion" _
        & " WHERE " & filter & " ORDER BY rec.id DESC" _



Dim col As New Collection
    Dim rec As Recibo

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




        If includeCheques Then
            Set rec.Cheques = DAOCheques.FindAll(DAOCheques.TABLA_CHEQUE & "." & DAOCheques.CAMPO_ID & " IN (SELECT idCheque FROM AdminRecibosCheques WHERE idRecibo = " & rec.Id & ")")
        End If

        If includeBanco Then
            Set rec.operacionesBanco = DAOOperacion.FindAll(Banco, "op.id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.Id & ")")
        End If

        If includeCaja Then
            Set rec.OperacionesCaja = DAOOperacion.FindAll(caja, "op.id IN (SELECT operacionId FROM operaciones_recibos WHERE reciboId = " & rec.Id & ")")
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
                    ) As Recibo

    Dim r As Recibo
    Dim Id As Long

    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set r = New Recibo
        r.Id = Id
        r.estado = GetValue(rs, indice, tabla, "estado")
        r.FechaAprobacion = GetValue(rs, indice, tabla, "fechaAprobacion")
        r.fechaCreacion = GetValue(rs, indice, tabla, "fechaCreacion")
        r.fechaModificacion = GetValue(rs, indice, tabla, "fechaModificacion")
        'r.PagoACuenta = GetValue(rs, indice, tabla, "pagoACuenta")
        r.redondeo = GetValue(rs, indice, tabla, "redondeo")
        r.aCuenta = GetValue(rs, indice, tabla, "a_cuenta")
        r.aCuentaUsado = GetValue(rs, indice, tabla, "a_cuenta_usado")
        r.FEcha = GetValue(rs, indice, tabla, "fecha")


        Dim totEstatico As New TotalEstaticoRecibo
        totEstatico.TotalChequesEstatico = GetValue(rs, indice, tabla, "tot_estatico_cheques")
        totEstatico.TotalDepositosEstatico = GetValue(rs, indice, tabla, "tot_estatico_depositos")
        totEstatico.TotalEfectivoEstatico = GetValue(rs, indice, tabla, "tot_estatico_efectivo")
        totEstatico.TotalReciboEstatico = GetValue(rs, indice, tabla, "tot_estatico_recibo")
        Set r.TotalEstatico = totEstatico


        If LenB(tablaMoneda) > 0 Then Set r.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaCliente) > 0 Then Set r.Cliente = DAOCliente.Map(rs, indice, tablaCliente)
        If LenB(tablaUsuarioCreador) > 0 Then Set r.usuarioCreador = DAOUsuarios.Map(rs, indice, tablaUsuarioCreador)
        If LenB(tablaUsuarioAprobador) > 0 Then Set r.usuarioAprobador = DAOUsuarios.Map(rs, indice, tablaUsuarioAprobador)
    End If

    Set Map = r
End Function

Public Function Save(rec As Recibo) As Boolean
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

Public Function Guardar(rec As Recibo) As Boolean
    On Error GoTo E


    Dim q As String
    Dim ReciboID As Long

    If rec.Id = 0 Then

        q = "INSERT INTO AdminRecibosAnticipo" _
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

        rec.fechaCreacion = Now
    Else

        q = "Update AdminRecibosAnticipo" _
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

         rec.fechaModificacion = Now

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

    q = Replace(q, "'idCliente'", conectar.GetEntityId(rec.Cliente))
    q = Replace(q, "'fechaCreacion'", conectar.Escape(rec.fechaCreacion))
    q = Replace(q, "'idUsuarioCreador'", conectar.GetEntityId(rec.usuarioCreador))
    q = Replace(q, "'fechaModificacion'", conectar.Escape(rec.fechaModificacion))
    q = Replace(q, "'estado'", conectar.Escape(rec.estado))
    q = Replace(q, "'idMoneda'", conectar.GetEntityId(rec.moneda))
    q = Replace(q, "'redondeo'", conectar.Escape(rec.redondeo))
    q = Replace(q, "'fecha'", conectar.Escape(rec.FEcha))
    q = Replace(q, "'a_cuenta'", conectar.Escape(rec.aCuenta))

    Dim esNuevo As Boolean
    esNuevo = False
    If Not conectar.execute(q) Then GoTo E
    If rec.Id = 0 Then esNuevo = True
    If rec.Id <> 0 And rec.estado = EstadoRecibo.Pendiente Then  'en el insert no tiene nada de agregacion


        'CHEQUES----------------------------------------------------------
        q = "DELETE FROM AdminRecibosCheques WHERE idRecibo = " & rec.Id
        If Not conectar.execute(q) Then GoTo E
        Dim cheq As cheque
        For Each cheq In rec.Cheques

            If cheq.Id = 0 Then
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

            q = "INSERT INTO AdminRecibosCheques (idRecibo, idCheque) VALUES (" & rec.Id & ", " & cheq.Id & ")"
            If Not conectar.execute(q) Then GoTo E
        Next cheq


        'CAJA----------------------------------------------------------
        Dim op As operacion
        Dim recId As Long
        For Each op In rec.OperacionesCaja
            If Not DAOOperacion.Save(op) Then GoTo E
            conectar.UltimoId "operaciones", recId
            If recId = 0 Then GoTo E
            If Not conectar.execute("INSERT INTO operaciones_recibos VALUES (" & recId & "," & rec.Id & ")") Then GoTo E
        Next op

        'BANCO----------------------------------------------------------
        For Each op In rec.operacionesBanco
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

    Dim Recibo As Recibo
    Set Recibo = DAOReciboAnticipo.FindById(idRecibo, True, True, True, True, True)
    
    Dim Espacio As Integer
    Espacio = 300
    
    If IsSomething(Recibo) Then

        Dim origin As Integer
        Printer.CurrentY = Espacio
        Printer.CurrentX = Espacio
        Printer.FontBold = True
        origin = Printer.FontSize
        Printer.FontSize = origin + 5

        Dim cx As Integer
                                Printer.Print Chr(10)
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
              Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
                                Printer.Print Chr(10)
        Printer.CurrentX = Espacio
        Printer.Print "SIGNO PLAST S.A."
        Printer.CurrentX = Espacio
        Printer.Print "RECIBO DE ANTICIPO CLIENTE"
        Printer.FontSize = origin + 3
                Printer.CurrentX = Espacio
                Printer.Print "Número: " & Recibo.Id
                Printer.CurrentX = Espacio
                Printer.Print "Estado: " & enums.EnumEstadoRecibo(Recibo.estado)
                Printer.CurrentX = Espacio
                Printer.Print "Fecha: " & Format(Day(Recibo.FEcha), "00") & "/" & Format(Month(Recibo.FEcha), "00") & "/" & Format(Year(Recibo.FEcha), "0000")
                Printer.CurrentX = Espacio
                Printer.Print "Cliente: " & Recibo.Cliente.razon
        Printer.FontSize = origin
        Printer.FontBold = False
                        Printer.Print Chr(10)
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
        Printer.Print Chr(10)

        Printer.FontBold = True
        Printer.CurrentX = Espacio
        Printer.Print "Valores recibidos "
        Printer.FontBold = False
        If Recibo.operacionesBanco.count > 0 Then
            Printer.FontBold = True
 Printer.CurrentX = Espacio
            Printer.Print "Banco"
            Printer.FontBold = False
        Else
        Printer.CurrentX = Espacio
            Printer.Print "Sin operaciones de banco"
        End If

        Dim o As operacion
        For Each o In Recibo.operacionesBanco
           Printer.CurrentX = Espacio
           Printer.Print o.FechaOperacion, o.CuentaBancaria.DescripcionFormateada, (FormatCurrency(funciones.FormatearDecimales(o.Monto)))
        Next o

        If Recibo.operacionesBanco.count > 0 Then
            Printer.FontBold = True
            Printer.CurrentX = Espacio
            Printer.Print "Total Banco: " & (FormatCurrency(funciones.FormatearDecimales(Recibo.TotalOperacionesBanco)))
        End If

        Printer.Print Chr(10)
        
        If Recibo.OperacionesCaja.count > 0 Then
            Printer.FontBold = True
            Printer.CurrentX = Espacio
            Printer.Print "Caja"
            Printer.FontBold = False
        Else
Printer.CurrentX = Espacio
            Printer.Print "Sin operaciones de caja"

        End If

        For Each o In Recibo.OperacionesCaja
            'Printer.Print o.FechaOperacion, o.CuentaBancaria.DescripcionFormateada, o.Monto
            
            Printer.CurrentX = Espacio
            Printer.Print o.FechaOperacion, (FormatCurrency(funciones.FormatearDecimales(o.Monto)))
        Next o
        
        Printer.FontBold = True
        If Recibo.OperacionesCaja.count > 0 Then
            Printer.CurrentX = Espacio
            Printer.Print "Total Caja: " & (FormatCurrency(funciones.FormatearDecimales(Recibo.TotalOperacionesCaja)))
        End If
     
        Printer.Print Chr(10)
        
        If Recibo.Cheques.count > 0 Then
            Printer.FontBold = True
            Printer.CurrentX = Espacio
            Printer.Print "Cheques Recibidos"
            Printer.FontBold = False
            Printer.CurrentX = Espacio
            Printer.Print "Numero,"; vbTab; "Monto"; vbTab; vbTab; "Fecha Vto."; vbTab; "Banco"
        Else
Printer.CurrentX = Espacio
            Printer.Print "Sin cheques recibidos."

        End If
        

        
        Dim che As cheque
        For Each che In Recibo.Cheques
            
            'Printer.Print o.FechaOperacion, o.CuentaBancaria.DescripcionFormateada, o.Monto
Printer.CurrentX = Espacio
            Printer.Print che.numero, (FormatCurrency(funciones.FormatearDecimales(che.Monto))), che.FechaVencimiento, che.Banco.nombre
        Next che
        
        Printer.FontBold = True
        If Recibo.Cheques.count > 0 Then
        Printer.CurrentX = Espacio
            Printer.Print "Total Cheques Recibidos: " & (FormatCurrency(funciones.FormatearDecimales(Recibo.TotalCheques)))
        End If
        Printer.Print Chr(10)
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
'        Printer.CurrentX = Espacio
'        Printer.Print " Total Recibo:  " & (FormatCurrency(funciones.FormatearDecimales(recibo.Total)))
        Printer.CurrentX = Espacio
        Printer.Print " Total Recibido:  " & (FormatCurrency(funciones.FormatearDecimales(Recibo.TotalRecibido)))
        Printer.FontBold = False
                Printer.Print Chr(10)
                Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
                        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
                                                Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
        Printer.EndDoc
    End If

End Sub




