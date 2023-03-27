Attribute VB_Name = "DAOCuentaContable"
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Private Const PERIVA As Long = 4
Private Const PERIIBB As Long = 5
Private Const IVACREDITO As Long = 71

Public Function PutSaldos(col As Collection, rango As String) As Collection
    On Error GoTo err1
    Dim col1 As New Collection
    Dim q As String
    Dim rs As Recordset




    'TENGO Q MAPEAR ALGUNAS CUENTAS CON VALORES
    'IVA @3  ---> CON PERCEPCIONES IVA
    'IIBB @6---> CON PERCEPCIONES IIBB
    'IVA CREDITO FISCAL  IC ---> IVA COMPRAS





    q = "select  cc.id, SUM( IF(tipo_doc_contable=1, -1* cf.monto,cf.monto )) AS gastado,  IF (fp.id_moneda=0,1,fp.`tipo_cambio`) AS cambio FROM AdminComprasCuentasContables cc" _
      & " LEFT JOIN AdminComprasCuentasFacturas cf " _
      & " ON cf.id_cuenta = cc.id   LEFT JOIN AdminComprasFacturasProveedores fp " _
      & " ON cf.id_factura = fp.id  " _
      & " Where 1 = 1   " & rango

    q = q & "GROUP BY cc.id ORDER BY cc.codigo ASC "




    Dim c As clsCuentaContable
    Set rs = conectar.RSFactory(q)

    Dim A As Integer
    While Not rs.EOF And Not rs.BOF


        If BuscarEnColeccion(col, CStr(rs!Id)) Then
            Set c = col(CStr(rs!Id))
            If IsNumeric(rs!gastado) Then


                c.TotalAcumulado = (rs!gastado * rs!Cambio) + c.TotalAcumulado
            Else
                c.TotalAcumulado = 0
            End If
        End If

        rs.MoveNext
    Wend
    Configurar.LoadConfiguration



    'q = "SELECT   SUM(   ( IF(tipo_doc_contable=1, -1* fpi.valor * cm.cambio,fpi.valor * cm.cambio))*(acia.alicuota/100)     ) AS gastado FROM AdminComprasFacturasProveedores fp " _
     & " LEFT JOIN AdminConfigMonedas cm    ON fp.id_moneda = cm.id " _
     & " LEFT JOIN AdminComprasFacturasProveedoresIva fpi     ON fpi.id_factura_proveedor = fp.id " _
     & " LEFT JOIN AdminConfigIvaAlicuotas acia ON fpi.id_iva=acia.id " _
     & " Where 1 = 1 " & rango


    q = "SELECT   SUM(   ( IF(tipo_doc_contable=1, -1* fpi.valor ,fpi.valor ))*(acia.alicuota/100)     ) AS gastado ,  IF (fp.id_moneda=0,1,fp.`tipo_cambio`) AS cambio FROM AdminComprasFacturasProveedores fp " _
      & " LEFT JOIN AdminComprasFacturasProveedoresIva fpi     ON fpi.id_factura_proveedor = fp.id " _
      & " LEFT JOIN AdminConfigIvaAlicuotas acia ON fpi.id_iva=acia.id " _
      & " Where 1 = 1 " & rango _
      & " group by fp.id"

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF

        If BuscarEnColeccion(col, CStr(Configurar.IdCtaIVACredito)) Then
            Set c = col(CStr(Configurar.IdCtaIVACredito))
            If IsNumeric(rs!gastado) Then


                c.TotalAcumulado = (rs!gastado * rs!Cambio) + c.TotalAcumulado
            Else
                c.TotalAcumulado = 0
            End If
        End If

        rs.MoveNext
    Wend





    'q = " SELECT acp.id, SUM( IF(tipo_doc_contable=1, -1* fpp.valor * cm.cambio,fpp.valor * cm.cambio)) AS gastado  FROM AdminComprasFacturasProveedores fp " _
     & "  LEFT JOIN AdminConfigMonedas cm  ON fp.id_moneda = cm.id " _
     & " LEFT JOIN AdminComprasFacturasProveedoresPercepciones fpp ON fpp.id_factura_proveedor=fp.id " _
     & " LEFT JOIN AdminConfigPercepciones acp ON fpp.id_percepcion=acp.id " _
     & " Where 1 = 1 " & rango _
     & " GROUP BY fpp.id_percepcion "

    q = " SELECT acp.id, SUM( IF(tipo_doc_contable=1, -1* fpp.valor * (IF (    fp.id_moneda = 0,    1,    fp.`tipo_cambio`  )), fpp.valor * (IF (    fp.id_moneda = 0,    1,    fp.`tipo_cambio`  )) )) AS gastado,  IF (fp.id_moneda=0,1,fp.`tipo_cambio`) AS cambio   FROM AdminComprasFacturasProveedores fp " _
      & "  LEFT JOIN AdminConfigMonedas cm  ON fp.id_moneda = cm.id " _
      & " LEFT JOIN AdminComprasFacturasProveedoresPercepciones fpp ON fpp.id_factura_proveedor=fp.id " _
      & " LEFT JOIN AdminConfigPercepciones acp ON fpp.id_percepcion=acp.id " _
      & " Where 1 = 1 " & rango _
      & " GROUP BY fpp.id_percepcion "

    Dim idp As Long



    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF
        If rs!Id = 1 Then idp = Configurar.IdCtaPercepcionesIIBB
        If rs!Id = 2 Then idp = Configurar.IdCtaPercepcionesIVA
        If rs!Id > 2 Then idp = Configurar.IdCtaPercepcionesIIBBResto



        If BuscarEnColeccion(col, CStr(idp)) Then
            Set c = col(CStr(idp))
            If IsNumeric(rs!gastado) Then


                'c.TotalAcumulado = (rs!gastado * rs!Cambio) + c.TotalAcumulado
                c.TotalAcumulado = (rs!gastado) + c.TotalAcumulado
            Else
                c.TotalAcumulado = 0
            End If
        End If

        rs.MoveNext
    Wend


    q = " SELECT SUM( IF(tipo_doc_contable=1, -1* fp.impuesto_interno ,fp.impuesto_interno )) AS gastado ,  IF (fp.id_moneda=0,1,fp.`tipo_cambio`) AS cambio FROM AdminComprasFacturasProveedores fp " _
      & " LEFT JOIN AdminConfigMonedas cm  ON fp.id_moneda = cm.id " _
      & " Where 1 = 1 " & rango & " group by fp.id"

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF

        If BuscarEnColeccion(col, CStr(Configurar.IdCtaCombustible)) Then
            Set c = col(CStr(Configurar.IdCtaCombustible))
            If IsNumeric(rs!gastado) Then


                c.TotalAcumulado = (rs!gastado * rs!Cambio) + c.TotalAcumulado
            Else
                c.TotalAcumulado = 0
            End If
        End If

        rs.MoveNext
    Wend

    'q = " SELECT SUM( IF(tipo_doc_contable=1, -1* fp.redondeo_iva * cm.cambio,fp.redondeo_iva * cm.cambio)) AS gastado  FROM AdminComprasFacturasProveedores fp " _
     & " LEFT JOIN AdminConfigMonedas cm  ON fp.id_moneda = cm.id " _
     & " Where 1 = 1 " & rango

    q = " SELECT SUM( IF(tipo_doc_contable=1, -1* fp.redondeo_iva ,fp.redondeo_iva )) AS gastado , IF (fp.id_moneda=0,1,fp.`tipo_cambio`) AS cambio FROM AdminComprasFacturasProveedores fp " _
      & " LEFT JOIN AdminConfigMonedas cm  ON fp.id_moneda = cm.id " _
      & " Where 1 = 1 " & rango & " group by fp.id"

    Set rs = conectar.RSFactory(q)
    While Not rs.EOF And Not rs.BOF

        If BuscarEnColeccion(col, CStr(Configurar.IdCtaRedondeo)) Then
            Set c = col(CStr(Configurar.IdCtaRedondeo))
            If IsNumeric(rs!gastado) Then


                c.TotalAcumulado = (rs!gastado * rs!Cambio) + c.TotalAcumulado

            Else
                c.TotalAcumulado = 0
            End If
        End If

        rs.MoveNext
    Wend




    Set PutSaldos = col
    Exit Function
err1:
    Set PutSaldos = col

End Function

Public Function GetAll(Optional orderByCodigo As Boolean = False, Optional filtro As String = vbNullString) As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim cta As clsCuentaContable
    Dim q As String


    q = "select * from AdminComprasCuentasContables cc WHERE 1=1"




    If LenB(filtro) > 0 Then

        q = q & " And " & filtro

    End If


    If orderByCodigo Then
        q = q & " order by cc.codigo asc"
    Else
        q = q & " order by cc.nombre asc"
    End If





    Set rs = conectar.RSFactory(q)
    While Not rs.EOF
        Set cta = New clsCuentaContable
        cta.codigo = rs!codigo
        cta.Id = rs!Id
        cta.nombre = rs!nombre


        col.Add cta, CStr(cta.Id)

        rs.MoveNext
    Wend
    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function
Public Function GetById(id_cuenta As Long) As clsCuentaContable
    On Error GoTo err1
    Dim col As Collection
    Dim rs As Recordset
    Dim cta As clsCuentaContable
    Set rs = conectar.RSFactory("select * from AdminComprasCuentasContables where id=" & id_cuenta)
    If Not rs.EOF And Not rs.BOF Then
        Set cta = New clsCuentaContable
        cta.codigo = rs!codigo
        cta.Id = rs!Id
        cta.nombre = rs!nombre
        Set GetById = cta

    Else
        Set GetById = Nothing
    End If
    Exit Function
err1:
    Set GetById = Nothing
End Function
Public Function Save(cuenta As clsCuentaContable) As Boolean
    On Error GoTo err1
    Set cn = conectar.obternerConexion
    Save = True
    cn.execute "insert into AdminComprasCuentasContables (codigo, nombre) values  ('" & cuenta.codigo & "','" & cuenta.nombre & "')"
    Exit Function
err1:
    Save = False
End Function
Public Function Update(cuenta As clsCuentaContable) As Boolean
    Set cn = conectar.obternerConexion
    On Error GoTo err1
    Update = True
    cn.execute "update AdminComprasCuentasContables set nombre='" & cuenta.nombre & "',codigo='" & cuenta.codigo & "' where id=" & cuenta.Id
    Exit Function
err1:
    Update = False
End Function



Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As clsCuentaContable

    Dim cc As clsCuentaContable
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set cc = New clsCuentaContable
        cc.Id = Id
        cc.codigo = GetValue(rs, indice, tabla, "codigo")
        cc.nombre = GetValue(rs, indice, tabla, "nombre")
    End If

    Set Map = cc
End Function


Public Function ImprimirColeccion(col As Collection, rango As String, valuados As Boolean) As Boolean
    On Error GoTo err1
    ImprimirColeccion = True
    Dim cta As clsCuentaContable
    Dim yPos As Long
    Dim rowcount As Long
    Dim colcod As Long
    Dim colmonto As Long
    Dim colcount As Long
    Printer.Print "RESUMEN DE CUENTAS CONTABLES"
    Printer.Print "Rango de Fechas: " & rango
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    Printer.FontSize = 8
    yPos = Printer.CurrentY
    rowcount = 0
    colcount = 0
    colcod = colcount + 1
    colmonto = colcod + 53
    Dim totacu As Double
    totacu = 0
    For Each cta In col

        If (valuados And cta.TotalAcumulado) Or Not valuados Then

            Printer.Print Tab(colcod + 4);
            Printer.Print cta.codigo & " | " & cta.nombre;
            Printer.Print Tab(colmonto + 4);
            Printer.Print funciones.FormatearDecimales(cta.TotalAcumulado)
            totacu = cta.TotalAcumulado + totacu
            rowcount = rowcount + 1

            If rowcount = 75 Then
                If colcount > 65 Then colcount = 0
                rowcount = 0
                Printer.CurrentY = yPos + 150
                colcount = colcount + 75
                colcod = colcount + 1
                colmonto = colcod + 53
            End If
            xpos = Printer.CurrentX
        End If

    Next cta
    Printer.Print
    Printer.Print Tab(colcount + 4);
    Printer.Print "TOTAL ACUMULADO: " & funciones.FormatearDecimales(totacu)
    Printer.EndDoc



    Exit Function
err1:
    ImprimirColeccion = False
End Function



Public Function ExportarColeccion(col As Collection, rango As String) As Boolean
    On Error GoTo err1
    ExportarColeccion = True
    Dim detalle As DetalleOrdenTrabajo
    Dim Entregas As Collection
    Dim remitoDetalle As remitoDetalle

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

    xlWorksheet.Cells(1, 1).value = "Rango Fecha "
    xlWorksheet.Cells(1, 2).value = rango
    xlWorksheet.Cells(2, 1).value = "Cód"
    xlWorksheet.Cells(2, 2).value = "Cuenta"
    xlWorksheet.Cells(2, 3).value = "Importe"
    Dim cta As clsCuentaContable

    Dim row As Long
    row = 3
    For Each cta In col
        xlWorksheet.Cells(row, 1) = cta.codigo
        xlWorksheet.Cells(row, 2) = cta.nombre
        'xlWorksheet.Range(xlWorksheet.Cells(row, 1), xlWorksheet.Cells(row, 1)).HorizontalAlignment = xlLeft
        xlWorksheet.Cells(row, 3) = cta.TotalAcumulado

        row = row + 1
    Next cta
    xlWorksheet.Cells(row + 1, 2).value = "Total"

    xlWorksheet.Cells(row + 1, 3).Formula = "=sum(c3:c" & row & ")"



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

    Exit Function
err1:
    ExportarColeccion = False
End Function
