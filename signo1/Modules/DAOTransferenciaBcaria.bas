Attribute VB_Name = "DAOTransferenciaBcaria"
Option Explicit

Public Function FindAll( _
    Origen As OrigenOperacion, _
    Optional ByVal extraFilter As String = "1 = 1", _
    Optional orderBy As String = "1", _
    Optional IncluirMovimientosCajaBanco As Boolean = False _
) As Collection

    Dim q As String

    q = "SELECT *, " _
      & " (op.pertenencia + 0) AS pertenencia2 " _
      & " FROM operaciones op" _

    'Cuenta bancaria solamente cuando la operación
    'pertenece realmente a Banco.
    q = q _
      & " LEFT JOIN AdminConfigCuentas cu" _
      & " ON cu.id = op.cuentabanc_o_caja_id" _
      & " AND (op.pertenencia + 0) = " & Banco _

    'Caja solamente cuando la operación pertenece a Caja.
    q = q _
      & " LEFT JOIN cajas caj" _
      & " ON caj.id = op.cuentabanc_o_caja_id" _
      & " AND (op.pertenencia + 0) = " & caja _

    q = q _
      & " LEFT JOIN AdminConfigMonedas mon" _
      & " ON op.moneda_id = mon.id" _
      & " LEFT JOIN AdminConfigBancos ban" _
      & " ON ban.id = cu.idBanco" _
      & " LEFT JOIN ordenes_pago_operaciones opope" _
      & " ON opope.id_operacion = op.id" _
      & " LEFT JOIN pagos_a_cuenta_operaciones pagosctaope" _
      & " ON pagosctaope.id_operacion = op.id" _
      & " LEFT JOIN pagos_a_cuenta pagoscta" _
      & " ON pagoscta.id = pagosctaope.id_pago_a_cuenta" _
      & " LEFT JOIN ordenes_pago opp" _
      & " ON opp.id = opope.id_orden_pago LEFT JOIN ordenes_pago_facturas opfac" _
      & " ON opfac.id_orden_pago = opp.id LEFT JOIN ordenes_pago_pagos_a_cuenta" _
      & " ON ordenes_pago_pagos_a_cuenta.id_pago_a_cuenta = pagoscta.id" _
      & " LEFT JOIN liquidaciones_caja_operaciones lco ON lco.id_operacion = op.id" _
      & " LEFT JOIN liquidaciones_caja liqc ON liqc.id = lco.id_liquidacion_caja" _
      & " LEFT JOIN liquidaciones_caja_facturas liqf ON liqf.id_liquidacion_caja = liqc.id" _
      & " LEFT JOIN AdminComprasFacturasProveedores facprov ON facprov.id = opfac.id_factura_proveedor" _
      & " LEFT JOIN proveedores prov ON prov.id = facprov.id_proveedor" _
      & " LEFT JOIN proveedores prov1 ON prov1.id = pagoscta.id_proveedor" _
      & " LEFT JOIN movimientos_caja_bancos_operaciones mcbop" _
      & " ON mcbop.id_operacion = op.id" _
      & " LEFT JOIN movimientos_caja_bancos mcb" _
      & " ON mcb.id = mcbop.id_movimiento_caja_bancos"

    '-------------------------------------------------
    ' ORIGEN
    '-------------------------------------------------
    If IncluirMovimientosCajaBanco Then

        'Mantiene todas las operaciones bancarias
        'del listado anterior.
        '
        'Las operaciones de caja solamente se incluyen
        'cuando pertenecen a movimientos_caja_bancos.
        q = q _
          & " WHERE (" _
          & " (op.pertenencia + 0) = " & Banco _
          & " OR (" _
          & "     (op.pertenencia + 0) = " & caja _
          & "     AND mcb.id IS NOT NULL" _
          & "    )" _
          & " )"

    Else

        q = q _
          & " WHERE (op.pertenencia + 0) = " & Origen

    End If

    'Solo pagos, egresos o el lado de salida
    'de una transferencia.
    q = q _
      & " AND op.entrada_salida = '-1'" _
      & " AND " & extraFilter

    q = q & " ORDER BY " & orderBy

    Dim col As New Collection
    Dim op As clsTransferenciaBcaria
    Dim idx As Dictionary
    Dim rs As Recordset

    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx

    While Not rs.EOF

        Set op = Map( _
            rs, _
            idx, _
            "op", _
            "cu", _
            "mon", _
            "ban", _
            "opope", _
            "pagosctaope", _
            "pagoscta", _
            "opp", _
            "opfac", _
            "liqc", _
            "liqf", _
            "facprov", _
            "prov", _
            "prov1", _
            "ordenes_pago_pagos_a_cuenta", _
            "caj", _
            "mcb")

        If IsSomething(op) Then

            If Not funciones.BuscarEnColeccion( _
                col, CStr(op.Id)) Then

                col.Add op, CStr(op.Id)

            End If

        End If

        rs.MoveNext

    Wend

    Set FindAll = col

End Function


Public Function Map( _
        rs As Recordset, _
        indice As Dictionary, _
        tabla As String, _
        Optional tablaCuentaBanc As String = vbNullString, _
        Optional tablaMoneda As String = vbNullString, _
        Optional tablaConfigBancos As String = vbNullString, _
        Optional tablaOrdenesPagoOperaciones As String = vbNullString, _
        Optional tablaPagosACuentaOperaciones As String = vbNullString, _
        Optional tablaPagosACuenta As String = vbNullString, _
        Optional tablaOrdenesPago As String = vbNullString, _
        Optional tablaOrdenesPagoFacturas As String = vbNullString, _
        Optional tablaLiquidacionesCaja As String = vbNullString, _
        Optional tablaLiquidacionesCajaFacturas As String = vbNullString, _
        Optional tablaFacturasProveedores As String = vbNullString, _
        Optional tablaProveedores As String = vbNullString, _
        Optional tablaProveedores1 As String = vbNullString, _
        Optional tablaOrdenesPagoACuenta As String = vbNullString, _
        Optional tablaCaja As String = vbNullString, _
        Optional tablaMovimientoCajaBanco As String = vbNullString _
    ) As clsTransferenciaBcaria
   
   
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    Dim op As clsTransferenciaBcaria


    If Id > 0 Then
        Set op = New clsTransferenciaBcaria
        op.Id = Id
        op.FechaCarga = GetValue(rs, indice, tabla, "fecha_carga")
        op.FechaOperacion = GetValue(rs, indice, tabla, "fecha_operacion")
        op.Pertenencia = GetValue(rs, indice, vbNullString, "pertenencia2")
        op.Monto = GetValue(rs, indice, tabla, "monto")
        op.EntradaSalida = GetValue(rs, indice, tabla, "entrada_salida")
        op.Comprobante = GetValue(rs, indice, tabla, "comprobante")

        If LenB(tablaOrdenesPago) > 0 Then Set op.OrdenPago = DAOOrdenPago.Map(rs, indice, tablaOrdenesPago)
        
        If LenB(tablaLiquidacionesCaja) > 0 Then Set op.LiquidacionCaja = DAOLiquidacionCaja.Map(rs, indice, tablaLiquidacionesCaja)
     
        If LenB(tablaMoneda) > 0 Then Set op.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
     
        op.IdCtaBancaria = GetValue(rs, indice, tablaCuentaBanc, "id")
        op.CuentaBancaria = GetValue(rs, indice, tablaCuentaBanc, "cuenta")
        op.NombreBanco = GetValue(rs, indice, tablaConfigBancos, "Nombre")

        op.ProveedorRazon = GetValue(rs, indice, tablaProveedores, "razon")
        
        op.PagoACuentaID = GetValue(rs, indice, tablaPagosACuenta, "id")
        
        op.PagoACuentaProveedor = GetValue(rs, indice, tablaProveedores1, "razon")
        
        op.OPAplicada = GetValue(rs, indice, tablaOrdenesPagoACuenta, "id_orden_pago")

        
    End If
    
        '---------------------------------------------
        ' Caja
        '---------------------------------------------
        If LenB(tablaCaja) > 0 Then
            op.NombreCaja = GetValue( _
                                rs, _
                                indice, _
                                tablaCaja, _
                                "nombre")
        End If

        '---------------------------------------------
        ' Movimiento manual de Caja y Bancos
        '---------------------------------------------
        If LenB(tablaMovimientoCajaBanco) > 0 Then

            op.MovimientoCajaBancoID = GetValue( _
                                                rs, _
                                                indice, _
                                                tablaMovimientoCajaBanco, _
                                                "id")

            op.TipoMovimientoCajaBanco = GetValue( _
                                                  rs, _
                                                  indice, _
                                                  tablaMovimientoCajaBanco, _
                                                  "tipo_movimiento")

            op.ObservacionesMovimiento = GetValue( _
                                                  rs, _
                                                  indice, _
                                                  tablaMovimientoCajaBanco, _
                                                  "observaciones")

        End If
        
                op.ProveedorRazon = GetValue( _
                                rs, _
                                indice, _
                                tablaProveedores, _
                                "razon")

        op.PagoACuentaID = GetValue( _
                                rs, _
                                indice, _
                                tablaPagosACuenta, _
                                "id")

        op.PagoACuentaProveedor = GetValue( _
                                        rs, _
                                        indice, _
                                        tablaProveedores1, _
                                        "razon")

        op.OPAplicada = GetValue( _
                                rs, _
                                indice, _
                                tablaOrdenesPagoACuenta, _
                                "id_orden_pago")

        If LenB(tablaCaja) > 0 Then

            op.NombreCaja = GetValue( _
                                    rs, _
                                    indice, _
                                    tablaCaja, _
                                    "nombre")

        End If

        If LenB(tablaMovimientoCajaBanco) > 0 Then

            op.MovimientoCajaBancoID = GetValue( _
                                                rs, _
                                                indice, _
                                                tablaMovimientoCajaBanco, _
                                                "id")

            op.TipoMovimientoCajaBanco = GetValue( _
                                                  rs, _
                                                  indice, _
                                                  tablaMovimientoCajaBanco, _
                                                  "tipo_movimiento")

            op.ObservacionesMovimiento = GetValue( _
                                                  rs, _
                                                  indice, _
                                                  tablaMovimientoCajaBanco, _
                                                  "observaciones")

        End If

    Set Map = op
    
End Function


Public Function ExportarColeccion(col As Collection, Optional ProgressBar As Object) As Boolean
    On Error GoTo err1

    ExportarColeccion = True


    
    Dim xlWorkbook As Object
    Set xlWorkbook = CreateObject("Excel.Application")

    Dim xlWorksheet As Object
    Set xlWorksheet = CreateObject("Excel.Application")

    Dim xlApplication As Object
    Set xlApplication = CreateObject("Excel.Application")

    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate
    
    Dim titulo As String
    titulo = "Reporte de Transferencias y Movimientos de Caja y Bancos"
    
    With xlWorksheet.Range("A1:J1")
        .Merge
        .Font.Bold = True
        .value = titulo
        .HorizontalAlignment = -4108 ' xlCenter
    End With


    'fila, columna

    Dim offset As Long
    offset = 3
    xlWorksheet.Cells(offset, 1).value = "ID"
    xlWorksheet.Cells(offset, 2).value = "Proveedor Destino"
    xlWorksheet.Cells(offset, 3).value = "N° Cta | Banco"
    xlWorksheet.Cells(offset, 4).value = "Fecha Operación"
    xlWorksheet.Cells(offset, 5).value = "Moneda"
    xlWorksheet.Cells(offset, 6).value = "Monto"
    xlWorksheet.Cells(offset, 7).value = "Comprobante"
    xlWorksheet.Cells(offset, 8).value = "OP/LIQ"
    xlWorksheet.Cells(offset, 9).value = "Estado PCTA"
    xlWorksheet.Cells(offset, 10).value = "OP Aplicado el PCTA"
    
        
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 10)).Font.Bold = True
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 10)).Interior.Color = &HC0C0C0

    Dim transf As clsTransferenciaBcaria

    Dim initoffset As Long
    
    initoffset = offset

    frmLoading.ProgressBar.min = 0
    
    frmLoading.ProgressBar.max = col.count
    Dim i As Integer
    i = 0
    
    For Each transf In col

        i = i + 1
        
        offset = offset + 1
       
        xlWorksheet.Cells(offset, 1).value = transf.Id
        
        If transf.LiquidacionCaja Is Nothing Then
             xlWorksheet.Cells(offset, 2).value = UCase(transf.ProveedorRazon)
        Else
            xlWorksheet.Cells(offset, 2).value = "VARIOS"
        End If
        
        If transf.Pertenencia = Banco Then

        xlWorksheet.Cells(offset, 3).value = _
                "N° " & transf.CuentaBancaria & _
                " | " & transf.NombreBanco
        
        ElseIf transf.Pertenencia = caja Then
        
            xlWorksheet.Cells(offset, 3).value = _
                "CAJA | " & transf.NombreCaja
        
        End If
        xlWorksheet.Cells(offset, 4).value = transf.FechaOperacion
        xlWorksheet.Cells(offset, 5).value = transf.moneda.NombreCorto
        xlWorksheet.Cells(offset, 6).value = transf.Monto
        xlWorksheet.Cells(offset, 7).value = transf.Comprobante
        
      If transf.EsMovimientoCajaBanco Then

    xlWorksheet.Cells(offset, 2).value = _
        "MOVIMIENTO MANUAL"

    xlWorksheet.Cells(offset, 8).value = _
        "MOV: " & transf.MovimientoCajaBancoID

    If transf.Pertenencia = Banco Then

        xlWorksheet.Cells(offset, 9).value = _
            UCase$(transf.TipoMovimientoCajaBanco) & _
            " - BANCO"

    ElseIf transf.Pertenencia = caja Then

        xlWorksheet.Cells(offset, 9).value = _
            UCase$(transf.TipoMovimientoCajaBanco) & _
            " - CAJA"

    End If

    xlWorksheet.Cells(offset, 10).value = ""

Else

    If transf.LiquidacionCaja Is Nothing Then

        If transf.OrdenPago Is Nothing Then

            xlWorksheet.Cells(offset, 8).value = _
                "PCTA: " & transf.PagoACuentaID

            xlWorksheet.Cells(offset, 2).value = _
                UCase$(transf.PagoACuentaProveedor)

            If transf.OPAplicada = 0 Then
                xlWorksheet.Cells(offset, 9).value = _
                    "Disponible"
            Else
                xlWorksheet.Cells(offset, 9).value = _
                    "Procesada"
            End If

            xlWorksheet.Cells(offset, 10).value = _
                transf.OPAplicada

        Else

            xlWorksheet.Cells(offset, 8).value = _
                "OP: " & transf.OrdenPago.Id

        End If

    Else

        xlWorksheet.Cells(offset, 8).value = _
            "LIQ: " & _
            transf.LiquidacionCaja.NumeroLiq

    End If

End If

        If transf.LiquidacionCaja Is Nothing Then
                If transf.OrdenPago Is Nothing Then
                    xlWorksheet.Cells(offset, 8).value = "PCTA: " & transf.PagoACuentaID
                    xlWorksheet.Cells(offset, 2).value = UCase(transf.PagoACuentaProveedor)
                    If transf.OPAplicada = "0" Then xlWorksheet.Cells(offset, 9).value = "Disponible" Else xlWorksheet.Cells(offset, 9).value = "Procesada"
                    xlWorksheet.Cells(offset, 10).value = transf.OPAplicada
                Else
                     xlWorksheet.Cells(offset, 8).value = "OP: " & transf.OrdenPago.Id
                End If
        Else
            xlWorksheet.Cells(offset, 8).value = "LIQ: " & transf.LiquidacionCaja.NumeroLiq
        End If
        
        
       
        
        
        frmLoading.ProgressBar.value = i
        
    Next

        xlWorksheet.Range(xlWorksheet.Cells(initoffset, 1), xlWorksheet.Cells(offset, 10)).Borders.LineStyle = xlContinuous

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

    If i = frmLoading.ProgressBar.max Then Unload frmLoading

    Exit Function
    
err1:
    ExportarColeccion = False
End Function


Public Function FindById(Id As Long) As clsTransferenciaBcaria
    Dim col As Collection
    Dim filtro As String
   
    filtro = "op.id = " & Id
    
    Set col = FindAll(Banco, filtro)
    
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
    
End Function

Public Function ActualizarDetallesComprobante(T As clsTransferenciaBcaria) As Boolean
    On Error GoTo err1

    Dim q As String
    q = "UPDATE sp.operaciones SET comprobante='comprobante', fecha_operacion = 'fecha_operacion', cuentabanc_o_caja_id='cuentabanc_o_caja_id' where id='id'"

    q = Replace$(q, "'id'", conectar.Escape(T.Id))
    q = Replace$(q, "'comprobante'", conectar.Escape(T.Comprobante))
    q = Replace$(q, "'fecha_operacion'", conectar.Escape(T.FechaOperacion))
    q = Replace$(q, "'cuentabanc_o_caja_id'", conectar.Escape(T.IdCtaBancaria))

    If Not conectar.execute(q) Then
        Err.Raise 112233, "No se pudieron actualizar los datos de la transferencia."
    End If
    ActualizarDetallesComprobante = True
    Exit Function
err1:
    Err.Raise Err.Number, Err.Description
    
End Function


