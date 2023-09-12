Attribute VB_Name = "DAOSubdiarios"
Option Explicit
Public Enum FcGroupMethod
    GroupNone = -1
    GroupByMonth = 1
    GroupByDate = 2
    GroupByYear = 3
End Enum

Public Function ExisteComprobanteEnLiquidacion(Id As Long) As Boolean
    On Error GoTo err1
    Dim qry As String
    Dim rs As Recordset
    qry = "select count(id) as c from liquidacion_subdiario_compras_detalles"
    Set rs = conectar.RSFactory(qry)
    Dim Cantidad As Integer
    While Not rs.EOF
        Cantidad = rs!c
    Wend

    ExisteComprobanteEnLiquidacion = Cantidad > 0

err1:
    ExisteComprobanteEnLiquidacion = False

End Function

Public Function ComprobanteComprasLiquidado(Id As Long) As Boolean
    On Error GoTo err1
    Dim qry As String
    Dim rs As Recordset
    qry = "select count(id) as c from liquidacion_subdiario_compras_detalles where id_factura=" & Id
    Set rs = conectar.RSFactory(qry)
    Dim Cantidad As Integer
    While Not rs.EOF
        Cantidad = rs!c
        rs.MoveNext
    Wend

    ComprobanteComprasLiquidado = Cantidad > 0
    Exit Function
err1:
    ComprobanteComprasLiquidado = True

End Function

Public Function SubDiarioCompras(FechaDesde As Date, FechaHasta As Date, Optional Orden As String = vbNullString, Optional idCliente As Long = -1, Optional idContratoMarco As Long = -1) As Collection
    Dim col_facturas As New Collection
    Dim newcol As New Collection
    Dim fc As clsFacturaProveedor
    Dim q As String
    Dim negativo As Integer
    Dim sv As SubdiarioVentasDetalle
    q = "AdminComprasFacturasProveedores.fecha between '" & funciones.dateFormateada(FechaDesde) & "' and '" & funciones.dateFormateada(FechaHasta) & "'"  ' and (AdminFacturas.estado IN (" & EstadoFacturaCliente.Aprobada & ", " & EstadoFacturaCliente.Anulada & ", " & EstadoFacturaCliente.CanceladaNC & "))"
    Set col_facturas = DAOFacturaProveedor.FindAll(q)

    Dim alicuotas As Collection
    Set alicuotas = DAOFacturaProveedor.FindAllAlicuotasIVA()
    Dim ali As Variant

    'Dim sumImpInt As Double
    'Dim sumRedondeo As Double
    Dim tipo_cambio As Double
    For Each fc In col_facturas

        ''If fc.Id = 16194 Then Stop

        'esto de abajo es porque puedo estar en pesos y a la ves tener un tipo de cambio (por la convertibilidad)
        '01-7-13
        If fc.moneda.Id = DAOMoneda.FindFirstByPatronOrDefault().Id Then
            tipo_cambio = 1
        Else
            tipo_cambio = fc.TipoCambio


        End If



        If fc.tipoDocumentoContable = tipoDocumentoContable.NotaCredito Then
            negativo = -1
        Else
            negativo = 1
        End If


        Set sv = New SubdiarioVentasDetalle
        sv.Comprobante = fc.NumeroFormateado
        sv.CondicionIva = fc.configFactura.TipoIvaProveedor.detalle
        sv.Cuit = fc.Proveedor.Cuit
        sv.FEcha = fc.FEcha


        'If sv.ComprobanteNro = 9225 Then Stop
        sv.Iva = RedondearDecimales(fc.TotalIVA) * negativo * tipo_cambio

        For Each ali In alicuotas
            sv.AlicuotasIva.Add fc.TotalIVADiscriminado(CDbl(ali)) * negativo * tipo_cambio, CStr(ali)
            sv.NetosGravado.Add fc.TotalNetoGravadoDiscriminado(CDbl(ali)) * negativo * tipo_cambio, CStr(ali)
        Next ali

        'sv.Percepciones = RedondearDecimales(fc.TotalPercepcionesDiscriminado(1)) * negativo


        Dim per As clsPercepcionesAplicadas

        Set sv.ListaPercepciones = New Collection


        'ver aca el indice de la coleccion  per.id o per.percepciones.id

        For Each per In fc.percepciones


            If Not BuscarEnColeccion(sv.ListaPercepciones, CStr(per.Percepcion.Id)) Then sv.ListaPercepciones.Add per, CStr(per.Percepcion.Id)


            per.Monto = fc.TotalPercepcionesDiscriminado(per.Percepcion.Id) * negativo * tipo_cambio

        Next

        Dim perTodas As Double

        perTodas = 0

        For Each per In sv.ListaPercepciones

            perTodas = perTodas + per.Monto

            'Debug.Print perTodas

        Next


        'If fc.totalPercepciones > 0 Then Stop


        'sv.percepciones = RedondearDecimales(fc.totalPercepciones) * negativo * tipo_cambio

        sv.percepciones = RedondearDecimales(perTodas)    '* negativo * tipo_cambio





        'sv.PercepcionesIVA = RedondearDecimales(fc.TotalPercepcionesDiscriminado(2)) * negativo

        sv.Exento = RedondearDecimales(fc.TotalIVADiscriminado(0)) * negativo * tipo_cambio
        sv.NetoGravado = RedondearDecimales(fc.NetoGravado) * negativo * tipo_cambio
        sv.ImpuestoInterno = RedondearDecimales(fc.ImpuestoInterno) * negativo * tipo_cambio
        sv.Redondeo = fc.Redondeo * negativo * tipo_cambio

        ' If fc.Id = 16325 Then Stop

        'sv.Total = funciones.RedondearDecimales(fc.Total) * negativo * tipo_cambio

        sv.total = funciones.RedondearDecimales(sv.NetoGravado + sv.ImpuestoInterno + sv.Redondeo + sv.Iva + (sv.percepciones / tipo_cambio))    '* negativo * tipo_cambio



        sv.RazonSocial = fc.Proveedor.RazonSocial
        sv.estado = fc.estado
        sv.FacturaId = fc.Id

        'sumImpInt = sumImpInt + fc.ImpuestoInterno
        'sumRedondeo = sumRedondeo + fc.redondeo

        newcol.Add sv
    Next

    'Debug.Print "impuestointerno: ", sumImpInt
    'Debug.Print "redonde: ", sumRedondeo
    Set SubDiarioCompras = newcol
End Function

Public Function SubDiarioVentas(FechaDesde As Date, FechaHasta As Date, Optional Orden As String = vbNullString, Optional idCliente As Long = -1, Optional idContratoMarco As Long = -1, Optional idprovincia As Long = -1) As Collection
    Dim col_facturas As New Collection
    Dim fc As Factura
    Dim sv As SubdiarioVentasDetalle
    Dim newcol As New Collection

    Dim q As String


    If Orden = vbNullString Then
        Orden = " nroFactura  ASC, AdminFacturas.FechaEmision  asc"
    End If



    q = "AdminFacturas.FechaEmision between '" & funciones.dateFormateada(FechaDesde) & "' and '" & funciones.dateFormateada(FechaHasta) & "' and (AdminFacturas.estado IN (" & EstadoFacturaCliente.Aprobada & ", " & EstadoFacturaCliente.Anulada & ", " & EstadoFacturaCliente.CanceladaNC & ", " & EstadoFacturaCliente.CanceladaNCParcial & ", " & EstadoFacturaCliente.AplicadaACbte & ", " & EstadoFacturaCliente.AplicadaND & ") AND aprobacion_afip = 1)"

    If idCliente > 0 Then
        q = q & " and AdminFacturas.idCliente=" & idCliente
    End If

    If idContratoMarco > -1 Then
        q = q & " and AdminFacturas.id IN (SELECT DISTINCT  fd.idFactura FROM pedidos p  INNER JOIN detalles_pedidos dp    ON dp.idPedido = p.id  INNER JOIN entregas e    ON e.idDetallePedido = dp.id  INNER JOIN AdminFacturasDetalleNueva fd    ON fd.idEntrega = e.id WHERE id_ot_padre = " & idContratoMarco & "))"
    End If


    If idprovincia > -1 Then
        q = q & " and Provincia.ID  = " & idprovincia
    End If


    Set col_facturas = DAOFactura.FindAll(q, True, True, Orden)

    Dim negativo As Integer
    For Each fc In col_facturas
        If fc.Tipo.TipoDoc = tipoDocumentoContable.NotaCredito Then
            negativo = -1
        Else
            negativo = 1
        End If


        Set sv = New SubdiarioVentasDetalle
        sv.Comprobante = fc.GetShortDescription(False, False)
        sv.CondicionIva = fc.cliente.TipoIVA.detalle
        sv.Cuit = fc.cliente.Cuit
        'sv.Exento = fc.TotalEstatico.TotalExento
        sv.FEcha = fc.FechaEmision


        'If fc.numero = 10125 Then Stop
        'If sv.ComprobanteNro = 9225 Then Stop
        sv.Iva = RedondearDecimales(fc.TotalEstatico.TotalIVADiscrimandoONo * fc.CambioAPatron) * negativo



        sv.percepciones = RedondearDecimales(fc.TotalEstatico.TotalPercepcionesIB * fc.CambioAPatron) * negativo

        'If sv.PercepcionesIB > 0.01 And sv.PercepcionesIB < 0.1 Then Stop


        sv.Exento = RedondearDecimales(fc.TotalEstatico.TotalExento * fc.CambioAPatron) * negativo
        ' If sv.Exento = 0 Then
        sv.NetoGravado = RedondearDecimales(fc.TotalEstatico.TotalNetoGravado * fc.CambioAPatron) * negativo
        'End If
        sv.total = funciones.RedondearDecimales(sv.percepciones + sv.NetoGravado + sv.Iva)  '* negativo '+ excento quitado 19-11-14



        sv.RazonSocial = fc.cliente.razon
        sv.estado = fc.estado
        sv.FacturaId = fc.Id

        newcol.Add sv
    Next

    Set SubDiarioVentas = newcol
End Function


Public Function FindAllLiquidacionesVenta(Optional venta As Boolean = True) As Collection
    Dim q As String
    q = "SELECT * FROM " & IIf(venta, "liquidacion_subdiario", "liquidacion_subdiario_compras") & " l ORDER BY l.desde DESC"
    Dim rs As Recordset
    Dim liquis As New Collection
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim liq As LiquidacionSubdiarioVenta

    While Not rs.EOF
        Set liq = New LiquidacionSubdiarioVenta

        liq.Id = GetValue(rs, fieldsIndex, "l", "id")
        liq.nombre = GetValue(rs, fieldsIndex, "l", "nombre")
        liq.desde = GetValue(rs, fieldsIndex, "l", "desde")
        liq.hasta = GetValue(rs, fieldsIndex, "l", "hasta")
        liq.EsDeVenta = venta

        liquis.Add liq, CStr(liq.Id)
        rs.MoveNext
    Wend

    Set FindAllLiquidacionesVenta = liquis
End Function


Public Function FindAllDetallesLiquiVentaByLiquiVenta(Id As Long, Optional venta As Boolean = True) As Collection
    Dim q As String

    If venta Then
        q = "SELECT * FROM liquidacion_subdiario_detalles ld WHERE ld.id_liquidacion = " & Id
    Else
        q = "SELECT * FROM liquidacion_subdiario_compras_detalles ld LEFT JOIN liquidacion_subdiario_compras_detalles_iva i ON i.subd_compra_detalle_id = ld.id LEFT JOIN liquidacion_subdiario_compras_detalles_ng ng ON ng.subd_compra_detalle_id = ld.id"
        q = q & " WHERE ld.id_liquidacion = " & Id
    End If

    Dim rs As Recordset
    Dim Detalles As New Collection
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim det As SubdiarioVentasDetalle

    Dim colAlicuotas As Collection
    Set colAlicuotas = DAOFacturaProveedor.FindAllAlicuotasIVA()
    Dim tmpAli As Variant

    While Not rs.EOF
        If Not funciones.BuscarEnColeccion(Detalles, CStr(GetValue(rs, fieldsIndex, "ld", "id"))) Then

            Set det = New SubdiarioVentasDetalle

            det.Id = GetValue(rs, fieldsIndex, "ld", "id")
            det.Comprobante = GetValue(rs, fieldsIndex, "ld", "comprobante")
            det.CondicionIva = GetValue(rs, fieldsIndex, "ld", "condicion_iva")
            det.Cuit = GetValue(rs, fieldsIndex, "ld", "cuit")
            det.estado = GetValue(rs, fieldsIndex, "ld", "estado_factura")
            det.Exento = GetValue(rs, fieldsIndex, "ld", "exento")
            det.FEcha = GetValue(rs, fieldsIndex, "ld", "fecha")
            det.Iva = GetValue(rs, fieldsIndex, "ld", "iva")
            det.NetoGravado = GetValue(rs, fieldsIndex, "ld", "neto_gravado")
            det.percepciones = GetValue(rs, fieldsIndex, "ld", "percepciones_iibb")
            det.RazonSocial = GetValue(rs, fieldsIndex, "ld", "razon_social")
            det.total = GetValue(rs, fieldsIndex, "ld", "total")
            det.FacturaId = GetValue(rs, fieldsIndex, "ld", "id_factura")
            det.LiquidacionId = GetValue(rs, fieldsIndex, "ld", "id_liquidacion")

            If Not venta Then
                det.ImpuestoInterno = GetValue(rs, fieldsIndex, "ld", "impuesto_interno")
                det.Redondeo = GetValue(rs, fieldsIndex, "ld", "redondeo")
                'det.PercepcionesIVA = GetValue(rs, fieldsIndex, "ld", "percepciones_iva")
                'falta percepciones iva



            End If

            Detalles.Add det, CStr(det.Id)
        Else
            Set det = Detalles.item(CStr(GetValue(rs, fieldsIndex, "ld", "id")))
        End If
        Dim aa As String
        'agergar alicuotaiva
        If Not venta Then
            aa = CStr(GetValue(rs, fieldsIndex, "i", "alicuota_iva_id"))
            If aa = "11" Then aa = "10.5"

            If Not funciones.BuscarEnColeccion(det.AlicuotasIva, aa) Then    ' CStr(GetValue(rs, fieldsIndex, "i", "alicuota_iva_id"))) Then






                det.AlicuotasIva.Add GetValue(rs, fieldsIndex, "i", "monto"), aa    'CStr(GetValue(rs, fieldsIndex, "i", "alicuota_iva_id"))
            End If

            aa = CStr(GetValue(rs, fieldsIndex, "ng", "alicuota_iva_id"))
            If aa = "11" Then aa = "10.5"

            If Not funciones.BuscarEnColeccion(det.NetosGravado, aa) Then    ' CStr(GetValue(rs, fieldsIndex, "ng", "alicuota_iva_id"))) Then


                det.NetosGravado.Add GetValue(rs, fieldsIndex, "ng", "monto"), aa    'CStr(GetValue(rs, fieldsIndex, "ng", "alicuota_iva_id"))
            End If

        End If

        rs.MoveNext
    Wend


    For Each det In Detalles
        For Each tmpAli In colAlicuotas

            If tmpAli = 10.5 Then tmpAli = 11

            If Not funciones.BuscarEnColeccion(det.AlicuotasIva, CStr(tmpAli)) Then
                det.AlicuotasIva.Add 0, CStr(tmpAli)
            End If
            If Not funciones.BuscarEnColeccion(det.NetosGravado, CStr(tmpAli)) Then
                det.NetosGravado.Add 0, CStr(tmpAli)
            End If
        Next tmpAli
    Next det


    Set FindAllDetallesLiquiVentaByLiquiVenta = Detalles

End Function

Public Function MaxFechaLiqui(Optional venta As Boolean = True) As Date
    Dim rs As New Recordset
    Dim q As String
    q = "SELECT MAX(hasta) as maxhasta FROM " & IIf(venta, "liquidacion_subdiario", "liquidacion_subdiario_compras")
    Set rs = conectar.RSFactory(q)

    Dim ret As Date

    If Not rs.EOF Then
        If Not IsNull(rs!maxhasta) Then
            ret = DateAdd("d", 1, rs!maxhasta)
        End If
    End If

    MaxFechaLiqui = ret
End Function


Public Function Guardar(liq As LiquidacionSubdiarioVenta) As Boolean
    On Error GoTo E

    If liq.Id <> 0 Then
        Guardar = False
        Exit Function
    End If

    Dim colAlicuotas As Collection
    Set colAlicuotas = DAOFacturaProveedor.FindAllAlicuotasIVA()
    Dim tmpAli As Variant

    Dim det As SubdiarioVentasDetalle
    conectar.BeginTransaction








    If liq.EsDeVenta Then
        'hay que comprobar correlatividad
        Dim minTipoLetraComprobante As New Collection
        Dim maxTipoLetraComprobante As New Collection
        Dim Valor As Long

        For Each det In liq.Detalles
            If funciones.BuscarEnColeccion(minTipoLetraComprobante, det.ComprobanteTipoLetra) Then
                Valor = minTipoLetraComprobante.item(det.ComprobanteTipoLetra)
                If CLng(det.ComprobanteNro) < Valor Then
                    minTipoLetraComprobante.remove det.ComprobanteTipoLetra
                    minTipoLetraComprobante.Add CLng(det.ComprobanteNro), det.ComprobanteTipoLetra
                End If
            Else
                minTipoLetraComprobante.Add CLng(det.ComprobanteNro), det.ComprobanteTipoLetra
            End If

            If funciones.BuscarEnColeccion(maxTipoLetraComprobante, det.ComprobanteTipoLetra) Then
                Valor = maxTipoLetraComprobante.item(det.ComprobanteTipoLetra)
                If CLng(det.ComprobanteNro) > Valor Then
                    maxTipoLetraComprobante.remove det.ComprobanteTipoLetra
                    maxTipoLetraComprobante.Add CLng(det.ComprobanteNro), det.ComprobanteTipoLetra
                End If
            Else
                maxTipoLetraComprobante.Add CLng(det.ComprobanteNro), det.ComprobanteTipoLetra
            End If
        Next det
        'tengo los maximos y minimos por tipo comprobante, ahora tengo que ver si son correlativos

        Dim min As Long
        Dim max As Long
        Dim letrasFacturas As Variant
        Dim letra As Variant
        Dim i As Long
        letrasFacturas = Array("A", "B", "C", "E")

        Dim estanCorrelativas As Boolean
        Dim existeNumero As Boolean
        Dim letrasNoCorrelativas As New Collection

        For Each letra In letrasFacturas
            If funciones.BuscarEnColeccion(minTipoLetraComprobante, CStr(letra)) And _
               funciones.BuscarEnColeccion(maxTipoLetraComprobante, CStr(letra)) Then
                'tengo los maximos y minimos para esa letra de factura

                min = minTipoLetraComprobante.item(CStr(letra))
                max = maxTipoLetraComprobante.item(CStr(letra))

                estanCorrelativas = True

                For i = min To max
                    existeNumero = False
                    For Each det In liq.Detalles
                        If det.ComprobanteTipoLetra = CStr(letra) And CLng(det.ComprobanteNro) = i Then
                            existeNumero = True
                            Exit For
                        End If
                    Next det

                    estanCorrelativas = estanCorrelativas And existeNumero
                Next i

                If Not estanCorrelativas Then
                    letrasNoCorrelativas.Add CStr(letra)
                End If

            End If

        Next letra
        'para sascar el filtro
        Set letrasNoCorrelativas = New Collection
        If letrasNoCorrelativas.count > 0 Then
            MsgBox "Los siguientes tipo de comprobantes no estan correlativos: " & vbNewLine & funciones.JoinCollectionValues(letrasNoCorrelativas, ", "), vbExclamation + vbOKOnly
            GoTo E
        End If
    End If

    'Dim facturasEnProceso As New Collection

    Dim q As String
    q = "INSERT INTO " & IIf(liq.EsDeVenta, "liquidacion_subdiario", "liquidacion_subdiario_compras") & "  (nombre, desde, hasta) VALUES ('nombre', 'desde', 'hasta')"
    q = Replace$(q, "'nombre'", conectar.Escape(liq.nombre))
    q = Replace$(q, "'desde'", conectar.Escape(liq.desde))
    q = Replace$(q, "'hasta'", conectar.Escape(liq.hasta))

    Dim Id As Long

    If conectar.execute(q) Then
        Id = conectar.UltimoId2()
        If Id = 0 Then GoTo E

        Dim iddet As Long

        For Each det In liq.Detalles

            '        If det.estado = EstadoFacturaCliente.EnProceso Then
            '            facturasEnProceso.Add det.Comprobante
            '        End If

            If liq.EsDeVenta Then
                q = "INSERT INTO liquidacion_subdiario_detalles (id_liquidacion, fecha, comprobante, razon_social, cuit, condicion_iva, neto_gravado, iva, percepciones_iibb, exento, total, estado_factura, id_factura) " _
                  & " VALUES ('id_liquidacion', 'fecha', 'comprobante', 'razon_social', 'cuit', 'condicion_iva', 'neto_gravado', 'iva', 'percepciones_iibb', 'exento', 'total', 'estado_factura', 'id_factura')"
            Else
                q = "INSERT INTO liquidacion_subdiario_compras_detalles (id_liquidacion, fecha, comprobante, razon_social, cuit, condicion_iva, neto_gravado, iva, percepciones_iibb, exento, total, estado_factura, id_factura, percepciones_iva, impuesto_interno, redondeo) " _
                  & " VALUES ('id_liquidacion', 'fecha', 'comprobante', 'razon_social', 'cuit', 'condicion_iva', 'neto_gravado', 'iva', 'percepciones_iibb', 'exento', 'total', 'estado_factura', 'id_factura', 'percepciones_iva', 'impuesto_interno', 'redondeo')"
            End If

            det.LiquidacionId = Id

            q = Replace$(q, "'id_liquidacion'", conectar.Escape(det.LiquidacionId))
            q = Replace$(q, "'fecha'", conectar.Escape(CDate(det.FEcha)))
            q = Replace$(q, "'comprobante'", conectar.Escape(det.Comprobante))
            q = Replace$(q, "'razon_social'", conectar.Escape(det.RazonSocial))
            q = Replace$(q, "'cuit'", conectar.Escape(det.Cuit))
            q = Replace$(q, "'condicion_iva'", conectar.Escape(det.CondicionIva))
            q = Replace$(q, "'neto_gravado'", conectar.Escape(det.NetoGravado))
            q = Replace$(q, "'iva'", conectar.Escape(det.Iva))
            q = Replace$(q, "'percepciones_iibb'", conectar.Escape(det.percepciones))
            q = Replace$(q, "'exento'", conectar.Escape(det.Exento))
            q = Replace$(q, "'total'", conectar.Escape(det.total))
            q = Replace$(q, "'estado_factura'", conectar.Escape(det.estado))
            q = Replace$(q, "'id_factura'", conectar.Escape(det.FacturaId))

            If Not liq.EsDeVenta Then
                q = Replace$(q, "'percepciones_iva'", conectar.Escape(det.percepciones))
                q = Replace$(q, "'impuesto_interno'", conectar.Escape(det.ImpuestoInterno))
                q = Replace$(q, "'redondeo'", conectar.Escape(det.Redondeo))
            End If

            If conectar.execute(q) Then
                iddet = 0
                iddet = conectar.UltimoId2()
                If iddet = 0 Then GoTo E
                det.Id = iddet

                For Each tmpAli In colAlicuotas
                    If funciones.BuscarEnColeccion(det.AlicuotasIva, CStr(tmpAli)) Then
                        q = "INSERT INTO liquidacion_subdiario_compras_detalles_iva (monto, alicuota_iva_id, subd_compra_detalle_id)" _
                          & " VALUES ('monto', 'alicuota_iva_id', 'subd_compra_detalle_id')"

                        q = Replace$(q, "'monto'", conectar.Escape(det.AlicuotasIva.item(CStr(tmpAli))))
                        q = Replace$(q, "'alicuota_iva_id'", conectar.Escape(tmpAli))
                        q = Replace$(q, "'subd_compra_detalle_id'", conectar.Escape(det.Id))

                        If Not conectar.execute(q) Then GoTo E
                    End If

                    If funciones.BuscarEnColeccion(det.NetosGravado, CStr(tmpAli)) Then
                        q = "INSERT INTO liquidacion_subdiario_compras_detalles_ng (monto, alicuota_iva_id, subd_compra_detalle_id)" _
                          & " VALUES ('monto', 'alicuota_iva_id', 'subd_compra_detalle_id')"

                        q = Replace$(q, "'monto'", conectar.Escape(det.NetosGravado.item(CStr(tmpAli))))
                        q = Replace$(q, "'alicuota_iva_id'", conectar.Escape(tmpAli))
                        q = Replace$(q, "'subd_compra_detalle_id'", conectar.Escape(det.Id))

                        If Not conectar.execute(q) Then GoTo E
                    End If

                Next tmpAli
            Else
                GoTo E
            End If

        Next det
    Else
        GoTo E
    End If

    'If facturasEnProceso.count > 0 Then
    '    MsgBox "Los siguientes comprobantes no estan aprobados ni anulados: " & vbNewLine & funciones.JoinCollectionValues(facturasEnProceso, ", "), vbExclamation + vbOKOnly
    '    GoTo E
    'End If


    Guardar = True
    conectar.CommitTransaction
    liq.Id = Id

    Exit Function
E:
    Guardar = False
    conectar.RollBackTransaction
End Function

Public Function UpdateDetalle(deta As SubdiarioVentasDetalle) As Boolean
    Dim q As String
    q = "UPDATE liquidacion_subdiario_detalles SET neto_gravado = 'neto_gravado', iva = 'iva' , percepciones_iibb = 'percepciones_iibb' , exento = 'exento' , total = 'total' WHERE id = 'id'"
    q = Replace$(q, "'neto_gravado'", conectar.Escape(deta.NetoGravado))
    q = Replace$(q, "'iva'", conectar.Escape(deta.Iva))
    q = Replace$(q, "'percepciones_iibb'", conectar.Escape(deta.percepciones))
    q = Replace$(q, "'exento'", conectar.Escape(deta.Exento))
    q = Replace$(q, "'total'", conectar.Escape(deta.total))
    q = Replace$(q, "'id'", deta.Id)
    UpdateDetalle = conectar.execute(q)
End Function

Public Sub PosicionIvaMensual()
    Dim mes As Integer
    Dim MesNombre As String
    Dim anio As Integer

    Dim d As New frmDateSelector
    d.Show 1
    If d.DateSelected Then
        mes = d.cboMes.ListIndex + 1
        MesNombre = d.cboMes.text
        anio = d.cboAnio.text
        Unload d
    Else
        Unload d
        Exit Sub
    End If

    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblTitulo").caption = "Posición IVA Mensual (" & MesNombre & " " & anio & ")"

    Dim Items As Collection
    Dim itsub As SubdiarioVentasDetalle
    Dim sumaDebitoFiscal As Double
    Set Items = DAOSubdiarios.SubDiarioVentas(DateSerial(anio, mes, 1), DateAdd("d", -1, DateSerial(anio, mes + 1, 1)))
    For Each itsub In Items
        sumaDebitoFiscal = sumaDebitoFiscal + itsub.Iva
    Next itsub
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblDebitoFiscalIVA").caption = FormatCurrency(sumaDebitoFiscal)


    Dim sumaCreditoFiscal5 As Double
    Dim sumaCreditoFiscal10 As Double
    Dim sumaCreditoFiscal21 As Double
    Dim sumaCreditoFiscal27 As Double
    Dim sumaPercepcionesIva As Double
    Dim per As clsPercepcionesAplicadas
    Set Items = DAOSubdiarios.SubDiarioCompras(DateSerial(anio, mes, 1), DateAdd("d", -1, DateSerial(anio, mes + 1, 1)))

    For Each itsub In Items
        sumaCreditoFiscal5 = sumaCreditoFiscal5 + itsub.AlicuotasIva(CStr(5))
        sumaCreditoFiscal10 = sumaCreditoFiscal10 + itsub.AlicuotasIva(CStr(10.5))
        sumaCreditoFiscal21 = sumaCreditoFiscal21 + itsub.AlicuotasIva(CStr(21))
        sumaCreditoFiscal27 = sumaCreditoFiscal27 + itsub.AlicuotasIva(CStr(27))
        '''' ver aca!!!!! ver ver ver

        For Each per In itsub.ListaPercepciones
            If per.Percepcion.Id = 2 Then    'percepciones de iva hardcodedddd
                sumaPercepcionesIva = sumaPercepcionesIva + per.Monto
            End If
        Next


    Next itsub
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblCreditoFiscalIVA5").caption = FormatCurrency(sumaCreditoFiscal5)
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblCreditoFiscalIVA10").caption = FormatCurrency(sumaCreditoFiscal10)
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblCreditoFiscalIVA21").caption = FormatCurrency(sumaCreditoFiscal21)
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblCreditoFiscalIVA27").caption = FormatCurrency(sumaCreditoFiscal27)
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblCreditoFiscalIVASUMA").caption = "(" & FormatCurrency(sumaCreditoFiscal5 + sumaCreditoFiscal10 + sumaCreditoFiscal21 + sumaCreditoFiscal27) & ")"
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblFiscalIVADif").caption = FormatCurrency(sumaDebitoFiscal - (sumaCreditoFiscal5 + sumaCreditoFiscal10 + sumaCreditoFiscal21 + sumaCreditoFiscal27))
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblPercepcionesIVA").caption = FormatCurrency(sumaPercepcionesIva)



    Dim sumRetIva As Double
    Dim F As New frmAdminSubdiarioRetenciones
    F.Visible = False
    F.DTDesde.value = DateSerial(anio, mes, 1)
    F.DTHasta.value = DateAdd("d", -1, DateSerial(anio, mes + 1, 1))
    F.Command2_Click
    sumRetIva = Val(F.lstSubdiarioRetenciones.ListItems.item(F.lstSubdiarioRetenciones.ListItems.count).SubItems(5))
    Unload F
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblRetencionesIVA").caption = FormatCurrency(sumRetIva)
    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblIVASUM").caption = "(" & FormatCurrency(sumaPercepcionesIva + sumRetIva) & ")"

    drpPosicionIvaMensual.Sections("seccion").Controls.item("lblTotal").caption = FormatCurrency((sumaDebitoFiscal - (sumaCreditoFiscal5 + sumaCreditoFiscal10 + sumaCreditoFiscal21 + sumaCreditoFiscal27) - (sumaPercepcionesIva + sumRetIva)))


    Dim r As Recordset
    Set r = conectar.RSFactory("SELECT 1")
    Set drpPosicionIvaMensual.DataSource = r
    'drpPosicionIvaMensual.Show
End Sub



