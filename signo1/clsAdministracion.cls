VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classAdministracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strsql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim claseP As New classPlaneamiento
Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion
    Set rs = New ADODB.Recordset
    Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function ejecutarComando(comando As String) As Boolean
    On Error GoTo err54
    ejecutarComando = True
    cn.execute comando
    Exit Function
err54:
    ejecutarComando = False
End Function
Public Function ejecutarConsulta(consulta As String) As Boolean
    On Error GoTo err1
    ejecutarConsulta = True
    Set rs = conectar.RSFactory(consulta)
    Exit Function
err1:
    MsgBox "Se produjo un error: " & Err.Description
    ejecutarConsulta = False
End Function


Public Function facturaImpresa(idFactura) As Long
    On Error GoTo err441
    Me.ejecutarConsulta "select impresa from AdminFacturas where id=" & idFactura
    Dim Im As Long
    Im = rs!impresa
    If Im > 0 Then
        facturaImpresa = CLng(rs!impresa)
    Else
        facturaImpresa = 0
    End If
    Exit Function
err441:
    facturaImpresa = -1    'no se pudo comprobar
End Function
Public Function facturasEntregadas(Origen As Integer, id As Long) As Recordset



    If Origen = 1 Or Origen = 2 Then    'facturas aplicadas a una OT then
        strsql = "select f.id,f.nroFactura as factura from AdminFacturasDetalleNueva df inner join entregas e on df.idEntrega=e.id inner join AdminFacturas f on f.id=df.idFactura where e.idPedido=" & id & " and e.origen=" & Origen & " group by df.idFactura "
    ElseIf Origen = 3 Then    'factura aplicadas a un remito
        strsql = "select f.id,f.nroFactura as factura from AdminFacturasDetalleNueva df inner join entregas e on df.idEntrega=e.id inner join AdminFacturas f on f.id=df.idFactura where e.remito=" & id & " group by df.idFactura"
    ElseIf Origen = -1 Then
        strsql = "select f.id,f.nroFactura as factura from AdminFacturas f  inner join AdminFacturasDetalleNueva df on df.idFactura=f.id inner join entregas e on e.id=df.idEntrega where idDetallePedido=" & id
    End If



    Me.ejecutarConsulta strsql
    Set facturasEntregadas = rs
End Function

Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub



Public Function ActualizarPadronIB(ruta As String, Tipo As TipoPadron) As Boolean
    On Error GoTo err4
    ActualizarPadronIB = True
    cn.BeginTrans

    If Tipo = TipoPadronPercepcion Then
        cn.execute "TRUNCATE sp_permisos.IIBB2_PercepcionAnt"
        cn.execute "INSERT INTO sp_permisos.IIBB2_PercepcionAnt SELECT * FROM sp_permisos.IIBB2_Percepcion"
        cn.execute "TRUNCATE sp_permisos.IIBB2_Percepcion"
        cn.execute "LOAD DATA LOCAL INFILE '" & ruta & "' INTO TABLE sp_permisos.IIBB2_Percepcion FIELDS TERMINATED BY ';'"
        
       'cn.execute "DELETE FROM sp_permisos.Padron_Detalles WHERE Padron='1'"
       'cn.execute "INSERT INTO sp_permisos.Padron_Detalles (Cuit, AlicuotaPercepcion, AlicuotaRetencion, Padron) SELECT Cuit, Alicuota, '0,00', '1' FROM sp_permisos.IIBB2_Percepcion"


    ElseIf Tipo = TipoPadronRetencion Then
        cn.execute "TRUNCATE sp_permisos.IIBB2_RetencionAnt"
        cn.execute "INSERT INTO sp_permisos.IIBB2_RetencionAnt SELECT * FROM sp_permisos.IIBB2_Retencion"
        cn.execute "TRUNCATE sp_permisos.IIBB2_Retencion"
        cn.execute "LOAD DATA LOCAL INFILE '" & ruta & "' INTO TABLE sp_permisos.IIBB2_Retencion FIELDS TERMINATED BY ';'"
 
        'cn.execute "UPDATE sp_permisos.Padron_Detalles f SET f.AlicuotaRetencion = (SELECT c.Alicuota from sp_permisos.IIBB2_Retencion c WHERE f.Cuit = c.Cuit)"
        'cn.execute "INSERT INTO sp_permisos.Padron_Detalles (Cuit, AlicuotaPercepcion, AlicuotaRetencion, Padron) SELECT Cuit, '0,00', Alicuota, '1' FROM sp_permisos.IIBB2_Retencion WHERE Cuit NOT IN (SELECT Cuit FROM sp_permisos.IIBB2_Percepcion)"
        'cn.execute "UPDATE sp_permisos.Padron_Detalles f SET f.AlicuotaRetencion = '0,00' WHERE f.AlicuotaRetencion IS NULL"
        
                    
    ElseIf Tipo = TipoPadronUnificadoCABA Then
        cn.execute "TRUNCATE sp_permisos.IIBB2_Padron_CABA_Ant"
        cn.execute "INSERT INTO sp_permisos.IIBB2_Padron_CABA_Ant SELECT * FROM sp_permisos.IIBB2_Padron_CABA"
        cn.execute "TRUNCATE sp_permisos.IIBB2_Padron_CABA"
        cn.execute "LOAD DATA LOCAL INFILE '" & ruta & "' INTO TABLE sp_permisos.IIBB2_Padron_CABA FIELDS TERMINATED BY ';'"
        
        'cn.execute "DELETE FROM sp_permisos.Padron_Detalles WHERE Padron='2'"
        'cn.execute "INSERT INTO sp_permisos.Padron_Detalles (Cuit, AlicuotaPercepcion, AlicuotaRetencion, Padron) SELECT Cuit, AlicuotaPercepcion, AlicuotaRetencion, '2' FROM sp_permisos.IIBB2_Padron_CABA"
        
    End If
    
    cn.CommitTrans
    Exit Function
err4:
    ActualizarPadronIB = False
    cn.RollbackTrans
End Function


Public Function ProcesarPadronIB() As Boolean
    On Error GoTo err4
    ProcesarPadronIB = True
    cn.BeginTrans

        cn.execute "TRUNCATE sp_permisos.Padron_Detalles_Ant"
        cn.execute "INSERT INTO sp_permisos.Padron_Detalles_Ant SELECT * FROM sp_permisos.Padron_Detalles"
      
        cn.execute "TRUNCATE sp_permisos.Padron_Detalles"
        
        
        Dim q  As String
        q = "INSERT INTO sp_permisos.Padron_Detalles (Cuit, AlicuotaPercepcion, AlicuotaRetencion, Padron, FechaDesdePercepcion, FechaHastaPercepcion, FechaDesdeRetencion, FechaHastaRetencion)" _
                     & "select * from (SELECT p.Cuit, p.Alicuota AS AlicuotaPercepcion, IFNULL(r.Alicuota,'0,00') AS AlicuotaRetencion, '1', p.FechaDesde AS pd, p.FechaHasta AS ph,r.FechaDesde,r.FechaHasta FROM sp_permisos.IIBB2_Percepcion p LEFT JOIN sp_permisos.IIBB2_Retencion r ON r.Cuit=p.Cuit" _
                     & " UNION SELECT r.Cuit, r.Alicuota AS AlicuotaRetencion, 0, '1','' AS c,'' AS d, r.FechaDesde AS a, r.FechaHasta AS b FROM sp_permisos.IIBB2_Retencion r WHERE NOT EXISTS (SELECT * FROM sp_permisos.IIBB2_Percepcion p WHERE p.Cuit=r.Cuit)" _
                     & " UNION SELECT c.Cuit, c.AlicuotaPercepcion, c.AlicuotaRetencion, '2' ,c.FechaDesde, c.FechaHasta, c.FechaDesde, c.FechaHasta FROM sp_permisos.IIBB2_Padron_CABA c) q"

        
        cn.execute q
        
        cn.CommitTrans
        
      
    Exit Function
    
err4:
        ProcesarPadronIB = False
    cn.RollbackTrans
End Function



Public Function realizaCambio(Origen As Double, idMonedaOrigen, idMonedaDestino) As Double
    Dim Total As Double
    Me.ejecutarConsulta "select id,cambio from AdminConfigMonedas where patron=1"
    If Not rs.EOF And Not rs.BOF Then
        CambioPatron = rs!Cambio
        idMonedaPatron = rs!id
        If CambioPatron <> 1 Then
            MsgBox "El patron debe ser valor 1", vbCritical, "Error"
            Exit Function
        End If
    End If

    Me.ejecutarConsulta "select cambio from AdminConfigMonedas where id=" & idMonedaOrigen
    If Not rs.EOF And Not rs.BOF Then
        cambioOrigen = rs!Cambio
    End If
    Me.ejecutarConsulta "select idMonedaCambio,cambio from AdminConfigMonedas where id=" & idMonedaDestino
    If Not rs.EOF And Not rs.BOF Then
        cambioDestino = rs!Cambio
        idMonedaCambio = rs!idMonedaCambio
    End If
    Me.ejecutarConsulta "select cambio from AdminConfigMonedas where id=" & idMonedaCambio
    If Not rs.EOF And Not rs.BOF Then
        Cambio = rs!Cambio
    End If
    If idMonedaCambio = idMonedaDestino Then Cambio = 1
    If idMonedaPatron = idMonedaDestino Then Cambio = 1
    If idMonedaOrigen = idMonedaDestino Then Cambio = 1
    Total = ((Origen / cambioOrigen) * CambioPatron * Cambio) / cambioDestino
    total2 = ((Origen * cambioOrigen) / cambioDestino * Cambio) / CambioPatron
    realizaCambio = total2    'funciones.RedondearDecimales(total, 3)

End Function



Public Function subdiario_cobros(lst As ListView, desde, hasta) As Recordset
    Dim totalfacturs As Double
    Dim tot_Ef As Double
    Dim tot_bco As Double
    Dim tot_ch As Double
    Dim tot_re As Double
    Dim rs_p As Recordset
    'ARMO EL ENCABEZADO
    lst.ColumnHeaders.Clear
    lst.ColumnHeaders.Add(1).text = "Fecha"
    lst.ColumnHeaders.Add(2).text = "Recibo"
    lst.ColumnHeaders.Add(3).text = "Razón Social"
    lst.ColumnHeaders(3).Width = 2200
    lst.ColumnHeaders.Add(4).text = "CUIT"
    lst.ColumnHeaders.Add(5).text = "Total facturas"


    lst.ColumnHeaders.Add(6).text = "Efectivo"
    lst.ColumnHeaders(6).Alignment = lvwColumnRight
    lst.ColumnHeaders(6).Width = 1000
    lst.ColumnHeaders.Add(7).text = "Cheques"
    lst.ColumnHeaders(7).Alignment = lvwColumnRight
    lst.ColumnHeaders(7).Width = 1440
    lst.ColumnHeaders.Add(8).text = "Banco"
    lst.ColumnHeaders(8).Alignment = lvwColumnRight
    lst.ColumnHeaders(8).Width = 1440
    lst.ColumnHeaders.Add(9).text = "Redondeo"
    lst.ColumnHeaders(9).Alignment = lvwColumnRight
    lst.ColumnHeaders(9).Width = 1440


    'encabezado de lst result




    II = 10
    i = 10
    Set rs_p = conectar.RSFactory("select * from retenciones order by id asc")

    While Not rs_p.EOF
        lst.ColumnHeaders.Add(i).text = rs_p!codigo
        lst.ColumnHeaders(i).Tag = rs_p!id
        lst.ColumnHeaders(i).Width = 1000

        lst.ColumnHeaders(i).Alignment = lvwColumnRight

        i = i + 1
        rs_p.MoveNext
    Wend
    Dim vec() As Double

    TOTAL_DE_RET = i - II
    ReDim Preserve vec(TOTAL_DE_RET) As Double
    Dim ret As retencionRecibo
    Dim rs_lista As Recordset
    Dim x As ListItem
    'TRAIGO LOS DATOS
    strsql = "SELECT c.razon,c.cuit,r.fechaCreacion,r.id,r.efectivo,r.cheques,r.idMoneda, redondeo from AdminRecibos r inner join clientes c on r.idCliente=c.id where r.fechaCreacion >= '" & Format(desde, "yyyy-mm-dd") & "' and r.fechaCreacion <=  '" & Format(hasta, "yyyy-mm-dd") & "'"
    Dim di As New Dictionary
    Set rs_lista = conectar.RSFactory(strsql)
    tot_Ef = 0
    tot_ch = 0
    tot_bco = 0
    tot_re = 0
    tot_fc = 0
    Dim r As recibo
    Dim y As ListItem



    While Not rs_lista.EOF
        '' Set ret = Me.totalizarRecibo(rs_lista!Id, totalfacturs) '' -->> VER
        Set r = DAORecibo.FindById(rs_lista!id, True, False, False, False, True)
        Set x = lst.ListItems.Add(, , Format(rs_lista!FechaCreacion, "dd-mm-yyyy"))
        x.SubItems(1) = Format(rs_lista!id, "0000-00000000")
        x.SubItems(2) = rs_lista!razon
        x.SubItems(3) = rs_lista!Cuit
        x.SubItems(4) = funciones.FormatearDecimales(r.TotalFacturas, 2)
        x.SubItems(5) = funciones.FormatearDecimales(r.TotalEstatico.TotalEfectivoEstatico, 2)
        x.SubItems(6) = funciones.FormatearDecimales(r.TotalEstatico.TotalChequesEstatico, 2)
        x.SubItems(7) = funciones.FormatearDecimales(r.TotalEstatico.TotalDepositosEstatico, 2)
        x.SubItems(8) = funciones.FormatearDecimales(r.Redondeo, 2)

        tot_Ef = tot_Ef + funciones.FormatearDecimales(r.TotalEstatico.TotalEfectivoEstatico, 2)
        tot_ch = tot_ch + funciones.FormatearDecimales(r.TotalEstatico.TotalChequesEstatico, 2)
        tot_re = tot_re + funciones.FormatearDecimales(r.Redondeo, 2)
        tot_fc = tot_fc + funciones.FormatearDecimales(r.TotalFacturas, 2)
        tot_bco = tot_bco + funciones.FormatearDecimales(r.TotalEstatico.TotalDepositosEstatico, 2)
        ia = 9
        ia_orig = ia
        Dim d As Double
        For o = 10 To 10 + TOTAL_DE_RET - 1    '
            For Each ret In r.retenciones


                If ret.Retencion.id = lst.ColumnHeaders(o).Tag Then
                    x.SubItems(o - 1) = ret.Valor
                Else
                    If (Not IsNumeric(x.SubItems(o - 1))) Then x.SubItems(o - 1) = "0.00"
                End If
            Next ret

        Next o

        For Each ret In r.retenciones

            If di.Exists(ret.Retencion.id) Then

                d = di.item(ret.Retencion.id)
                d = d + ret.Valor
                di.item(ret.Retencion.id) = d
            Else
                di.Add ret.Retencion.id, ret.Valor
            End If
        Next ret


        rs_lista.MoveNext
    Wend


    Set y = lst.ListItems.Add(, , Empty)
    y.SubItems(1) = Empty
    y.SubItems(2) = Empty
    y.SubItems(3) = Empty
    y.SubItems(4) = tot_fc
    y.SubItems(5) = funciones.FormatearDecimales(tot_Ef)
    y.SubItems(6) = funciones.FormatearDecimales(tot_ch)
    y.SubItems(7) = funciones.FormatearDecimales(tot_bco)
    y.SubItems(8) = funciones.FormatearDecimales(tot_re)
    y.ListSubItems(4).Bold = True
    y.ListSubItems(5).Bold = True
    y.ListSubItems(6).Bold = True
    y.ListSubItems(8).Bold = True
    y.ListSubItems(7).Bold = True



    For o = 10 To 10 + TOTAL_DE_RET - 1
        y.SubItems(o - 1) = di.item(lst.ColumnHeaders.item(o).Tag)
        y.ListSubItems(o - 1).Bold = True
    Next o

End Function

Public Function exportaSubDiarioCobros(lst As ListView, desde, hasta) As Boolean
    On Error GoTo errEXCEL
    Dim xlb As New Excel.Workbook
    Dim xla As New Excel.Worksheet
    Dim xls As New Excel.Application



    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate


    With xla



        canti_col = lst.ColumnHeaders.count

        .Range("A1:j3").HorizontalAlignment = xlHAlignCenter


        .Range(.Cells(1, 1), .Cells(1, canti_col)).Merge
        .Range(.Cells(2, 1), .Cells(2, canti_col)).Merge
        .Cells(1, 1).value = "SIGNOPLAST S.A. Subdiario cobranzas"
        .Cells(2, 1).value = "Periodo " & Format(desde, "dd/mm/yyyy") & " - " & Format(hasta, "dd/mm/yyyy")
        .Range(.Cells(1, 1), .Cells(3, canti_col)).Font.Bold = True

        .Range(.Cells(3, 1), .Cells(3, canti_col)).Interior.Color = &HC0C0C0
        .Range(.Cells(3, 1), .Cells(3, canti_col)).Font.Bold = True    '







        For x = 1 To canti_col
            .Cells(3, x).value = lst.ColumnHeaders(x).text

        Next
        For i = 1 To canti_col
            .Columns(i + 6).HorizontalAlignment = xlHAlignRight
            .Columns(i + 6).ColumnWidth = 13
        Next i
        .Columns("a").HorizontalAlignment = xlHAlignCenter
        .Columns("b").HorizontalAlignment = xlHAlignCenter
        .Columns("d").HorizontalAlignment = xlHAlignCenter
        .Columns("e").HorizontalAlignment = xlHAlignCenter

        .Columns("j").HorizontalAlignment = xlHAlignRight

        .Columns("a").ColumnWidth = 10
        .Columns("b").ColumnWidth = 15
        .Columns("c").ColumnWidth = 35
        .Columns("d").ColumnWidth = 13
        .Columns("e").ColumnWidth = 15
        .Columns("f").ColumnWidth = 13



        Dim vec() As Variant
        ReDim Preserve vec(canti_col - 1 - 3)

        For x = 1 To lst.ListItems.count - 1

            .Cells(x + 3, 1).value = lst.ListItems(x)
            .Cells(x + 3, 2).value = lst.ListItems(x).ListSubItems(1)
            .Cells(x + 3, 3).value = lst.ListItems(x).ListSubItems(2)
            .Cells(x + 3, 4).value = lst.ListItems(x).ListSubItems(3)

            'campos numericos
            For y = 4 To canti_col - 1
                .Cells(x + 3, y + 1).value = CDbl(lst.ListItems(x).ListSubItems(y))
                vec(y - 4) = vec(y - 4) + CDbl(lst.ListItems(x).ListSubItems(y))
            Next y


        Next x
        A = 1

        offset = x + 3


        b = .Cells(canti_col, x)    '"j" & Offset
        .Range("f1", .Cells(x + 2, canti_col)).NumberFormat = "0.00"
        .Range(.Cells(1, 1), .Cells(x + 2, canti_col)).Borders.LineStyle = xlContinuous

        .Range(.Cells(x + 3, 5), .Cells(x + 3, canti_col)).Interior.Color = &HC0C0C0
        .Range(.Cells(x + 3, 5), .Cells(x + 3, canti_col)).Font.Bold = True

        For o = 0 To UBound(vec)
            .Cells(x + 3, 5 + o) = vec(o)
        Next o
        .Range("e" & x + 3, .Cells(x, canti_col)).Borders.LineStyle = xlContinuous
        '.Range("f" & X + 3, .Cells(X, canti_col)).Font.Bold = True

        '.Cells(Offset, 10).value = total
        '.Cells(Offset, 9).value = totexen
        '.Cells(Offset, 8).value = totperi
        '.Cells(Offset, 7).value = totiv
        '.Cells(Offset, 6).value = totnetog
        '.Cells(Offset, 5).value = "Totales"





        'xls.Visible = True 'NO MUESTRO LA HOJA XLS
        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        Set CDLGMAIN = frmPrincipal.cd



        '    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter
        Dim refe As String
        refe = ref
        Periodo = 1
        Periodo = Format(desde, "ddmmyyyy") & "-" & Format(hasta, "ddmmyyyy")

        archi = "SUB_COBRANZAS_" & Periodo & ".xls"
        frmPrincipal.cd.CancelError = True
        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbExclamation, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            exportaSubDiarioCobros = False
        End If
        xlb.Saved = True
        xlb.Close

        xls.Quit
        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        '    End If
        exportaSubDiarioCobros = True



    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        exportaSubDiarioCobros = False
    Else
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        exportaSubDiarioCobros = False
    End If
    xlb.Saved = True
    xlb.Close
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing

End Function



Public Function subdiario_retenciones(lst As ListView, desde, hasta) As Recordset
    Dim rs_p As Recordset
    'ARMO EL ENCABEZADO
    lst.ColumnHeaders.Clear
    lst.ColumnHeaders.Add(1).text = "Fecha Recibo"
    lst.ColumnHeaders.Add(2).text = "Razón Social"
    lst.ColumnHeaders(2).Width = 3460

    lst.ColumnHeaders.Add(3).text = "CUIT"
    lst.ColumnHeaders.Add(4).text = "Recibo"


    lst.ColumnHeaders.Add(5).text = "Nro.Retencion"
    lst.ColumnHeaders(5).Width = 2500





    i = 5
    I_Orig = i
    Set rs_p = conectar.RSFactory("select * from retenciones order by id asc")

    While Not rs_p.EOF
        i = i + 1
        lst.ColumnHeaders.Add(i).text = rs_p!codigo
        lst.ColumnHeaders(i).Tag = rs_p!id
        lst.ColumnHeaders(i).Width = 1000
        lst.ColumnHeaders(i).Tag = rs_p!id


        rs_p.MoveNext
    Wend

    tot_columnas = i

    TOTAL_DE_RET = i


    Dim x As ListItem
    Dim y As ListItem
    'TRAIGO LOS DATOS
    'strsql = "SELECT c.razon,c.cuit,r.fechaCreacion,r.id,r.idMoneda from AdminRecibos r inner join clientes c on r.idCliente=c.id where r.fechaCreacion >= '" & Format(desde, "yyyy-mm-dd") & "' and r.fechaCreacion <=  '" & Format(hasta, "yyyy-mm-dd") & "'"

    strsql = "SELECT c.razon,c.cuit,r.fecha as fecha, r.id,r.idMoneda from AdminRecibos r inner join clientes c on r.idCliente=c.id right join AdminRecibosDetalleRetenciones dr on dr.idRecibo=r.id where dr.fecha >= '" & Format(desde, "yyyy-mm-dd") & "' and dr.fecha <=  '" & Format(hasta, "yyyy-mm-dd") & "' group by r.id"
    Dim ret As New Recordset
    Set rs_lista = conectar.RSFactory(strsql)

    While Not rs_lista.EOF
        '    Set ret = Me.totalizarRecibo(rs_lista!id, totalfacturs)
        '
        Set x = lst.ListItems.Add(, , Format(rs_lista!FEcha, "dd-mm-yyyy"))

        x.SubItems(1) = rs_lista!razon
        x.SubItems(2) = rs_lista!Cuit
        x.SubItems(3) = Format(rs_lista!id, "0000-00000000")

        Set ret = conectar.RSFactory("select *, fecha as fechaRet from AdminRecibosDetalleRetenciones where idRecibo=" & rs_lista!id & " and fecha >= '" & Format(desde, "yyyy-mm-dd") & "' and fecha <=  '" & Format(hasta, "yyyy-mm-dd") & "'")

        If ret.EOF = False And ret.BOF = False Then
            x.SubItems(4) = ret!NroRetencion & " (" & ret!fechaRet & ")"
            For h = 6 To tot_columnas
                If ret!idretencion = lst.ColumnHeaders(h).Tag Then
                    Valor = ret!Valor

                Else
                    Valor = 0

                End If
                x.SubItems(h - 1) = Valor
            Next h
            ret.MoveNext
        End If

        While Not ret.EOF
            Set y = lst.ListItems.Add(, , Empty)
            y.SubItems(1) = Empty
            y.SubItems(2) = Empty
            y.SubItems(3) = Empty
            y.SubItems(4) = ret!NroRetencion & " (" & ret!fechaRet & ")"
            For h = 6 To tot_columnas
                If ret!idretencion = lst.ColumnHeaders(h).Tag Then
                    Valor = ret!Valor
                Else
                    Valor = 0
                End If
                y.SubItems(h - 1) = Valor
            Next h
            '
            ret.MoveNext
            '
            '
            '
        Wend

        rs_lista.MoveNext
    Wend

    'calculo los totales
    Dim vec1() As Double
    tot_ret = tot_columnas - 6
    ReDim Preserve vec1(tot_ret) As Double

    For YY = 1 To lst.ListItems.count
        For xx = 0 To tot_ret
            vec1(xx) = CDbl(lst.ListItems(YY).ListSubItems(5 + xx)) + vec1(xx)
        Next xx
    Next YY

    Set y = lst.ListItems.Add(, , Empty)
    y.SubItems(1) = Empty
    y.SubItems(2) = Empty
    y.SubItems(3) = Empty
    y.SubItems(4) = Empty
    For i = 0 To tot_ret
        y.SubItems(i + 5) = vec1(i)
        y.Bold = True
        y.ListSubItems(i + 5).Bold = True
    Next i


End Function






Public Function exportaSubDiarioRetenciones(lst As ListView, desde, hasta) As Boolean
    On Error GoTo errEXCEL
    Dim xlb As New Excel.Workbook
    Dim xla As New Excel.Worksheet
    Dim xls As New Excel.Application



    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate


    With xla
        canti_col = lst.ColumnHeaders.count




        .Range("A1:j3").HorizontalAlignment = xlHAlignCenter


        .Range(.Cells(1, 1), .Cells(1, canti_col)).Merge
        .Range(.Cells(2, 1), .Cells(2, canti_col)).Merge
        .Cells(1, 1).value = "SIGNOPLAST S.A. Subdiario retenciones"
        .Cells(2, 1).value = "Periodo " & Format(desde, "dd/mm/yyyy") & " - " & Format(hasta, "dd/mm/yyyy")

        .Range(.Cells(1, 1), .Cells(3, canti_col)).Font.Bold = True

        .Range(.Cells(3, 1), .Cells(3, canti_col)).Interior.Color = &HC0C0C0
        .Range(.Cells(3, 1), .Cells(3, canti_col)).Font.Bold = True    '







        For x = 1 To canti_col
            .Cells(3, x).value = lst.ColumnHeaders(x).text

        Next
        For i = 1 To canti_col
            .Columns(i + 6).HorizontalAlignment = xlHAlignRight
            .Columns(i + 6).ColumnWidth = 13
        Next i
        .Columns("a").HorizontalAlignment = xlHAlignCenter
        .Columns("b").HorizontalAlignment = xlHAlignCenter
        .Columns("d").HorizontalAlignment = xlHAlignCenter
        .Columns("e").HorizontalAlignment = xlHAlignCenter

        .Columns("j").HorizontalAlignment = xlHAlignRight

        .Columns("a").ColumnWidth = 10
        .Columns("b").ColumnWidth = 35
        .Columns("c").ColumnWidth = 15
        .Columns("d").ColumnWidth = 13
        .Columns("e").ColumnWidth = 15
        .Columns("f").ColumnWidth = 13




        Dim vec() As Variant
        ReDim Preserve vec(canti_col - 1 - 3)
        '
        For x = 1 To lst.ListItems.count - 1
            '
            .Cells(x + 3, 1).value = lst.ListItems(x)
            .Cells(x + 3, 2).value = lst.ListItems(x).ListSubItems(1)
            .Cells(x + 3, 3).value = lst.ListItems(x).ListSubItems(2)
            .Cells(x + 3, 4).value = lst.ListItems(x).ListSubItems(3)
            .Cells(x + 3, 5).value = lst.ListItems(x).ListSubItems(4)
            '
            '    'campos numericos
            For y = 5 To canti_col - 1
                .Cells(x + 3, y + 1).value = CDbl(lst.ListItems(x).ListSubItems(y))
                vec(y - 5) = vec(y - 5) + CDbl(lst.ListItems(x).ListSubItems(y))
            Next y
            '
            '
        Next x
        A = 1

        offset = x + 3


        b = .Cells(canti_col, x)    '"j" & Offset
        .Range("f1", .Cells(x + 2, canti_col)).NumberFormat = "0.00"
        .Range(.Cells(1, 1), .Cells(x + 2, canti_col)).Borders.LineStyle = xlContinuous

        .Range(.Cells(x + 3, 6), .Cells(x + 3, canti_col)).Interior.Color = &HC0C0C0
        .Range(.Cells(x + 3, 6), .Cells(x + 3, canti_col)).Font.Bold = True

        For o = 0 To UBound(vec)
            .Cells(x + 3, 6 + o) = vec(o)
        Next o
        .Range("f" & x + 3, .Cells(x, canti_col)).Borders.LineStyle = xlContinuous


        'xls.Visible = True 'NO MUESTRO LA HOJA XLS
        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        Set CDLGMAIN = frmPrincipal.cd



        '    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter
        Dim refe As String
        refe = ref
        Periodo = 1
        Periodo = Format(desde, "ddmmyyyy") & "-" & Format(hasta, "ddmmyyyy")

        archi = "SUB_RETENCIONES_" & Periodo & ".xls"
        frmPrincipal.cd.CancelError = True
        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbExclamation, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            exportaSubDiarioRetenciones = False
        End If
        xlb.Saved = True
        xlb.Close

        xls.Quit
        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        '    End If
        exportaSubDiarioRetenciones = True



    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        exportaSubDiarioRetenciones = False
    Else
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        exportaSubDiarioRetenciones = False
    End If
    xlb.Saved = True
    xlb.Close
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing

End Function



Public Function escanearFactura(idFactura) As Boolean
    newdoc = sTwain.TransferWithoutUI(300, BW, 1, 1, 2, 5, "noui_mono.bmp")
End Function
