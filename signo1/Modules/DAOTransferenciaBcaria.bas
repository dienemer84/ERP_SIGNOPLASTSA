Attribute VB_Name = "DAOTransferenciaBcaria"
Option Explicit

Public Function FindAll(Origen As OrigenOperacion, Optional ByVal extraFilter As String = "1 = 1") As Collection
    Dim q As String
     
    q = "SELECT *, (op.pertenencia + 0) as pertenencia2 FROM" _
      & " operaciones op" _
      & " LEFT JOIN AdminConfigCuentas cu ON cu.id = op.cuentabanc_o_caja_id" _
      & " LEFT JOIN AdminConfigMonedas mon ON op.moneda_id = mon.id" _
      & " LEFT JOIN AdminConfigBancos ban ON ban.id = cu.idBanco" _
      & " LEFT JOIN ordenes_pago_operaciones opope ON opope.id_operacion = op.id" _
      & " LEFT JOIN ordenes_pago opp ON opp.id = opope.id_orden_pago" _
      & " LEFT JOIN ordenes_pago_facturas opfac ON opfac.id_orden_pago = opp.id" _
      & " LEFT JOIN liquidaciones_caja liqc ON liqc.id = opope.id_orden_pago" _
      & " LEFT JOIN liquidaciones_caja_facturas liqf ON liqf.id_liquidacion_caja = liqc.id" _
      & " LEFT JOIN AdminComprasFacturasProveedores facprov ON facprov.id = opfac.id_factura_proveedor" _
      & " LEFT JOIN proveedores prov ON prov.id = facprov.id_proveedor" _
      & " WHERE op.pertenencia = " & Origen & " AND op.entrada_salida = '-1' AND " & extraFilter
      
    Dim col As New Collection

    Dim op As clsTransferenciaBcaria
    
    Dim idx As Dictionary
    Dim rs As Recordset

    Set rs = conectar.RSFactory(q)
    BuildFieldsIndex rs, idx

    While Not rs.EOF

        Set op = Map(rs, idx, "op", "cu", "mon", "ban", "opope", "opp", "opfac", "liqc", "liqf", "facprov", "prov")
        
        If Not funciones.BuscarEnColeccion(col, CStr(op.Id)) Then col.Add op, CStr(op.Id)
        rs.MoveNext

    Wend

    Set FindAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaCuentaBanc As String = vbNullString, _
                    Optional tablaMoneda As String = vbNullString, _
                    Optional tablaConfigBancos As String = vbNullString, _
                    Optional tablaOrdenesPagoOperaciones As String = vbNullString, _
                    Optional tablaOrdenesPago As String = vbNullString, _
                    Optional tablaOrdenesPagoFacturas As String = vbNullString, _
                    Optional tablaLiquidacionesCaja As String = vbNullString, _
                    Optional tablaLiquidacionesCajaFacturas As String = vbNullString, _
                    Optional tablaFacturasProveedores As String = vbNullString, _
                    Optional tablaProveedores As String = vbNullString _
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
     
        op.CuentaBancaria = GetValue(rs, indice, tablaCuentaBanc, "cuenta")
        op.NombreBanco = GetValue(rs, indice, tablaConfigBancos, "Nombre")

        op.ProveedorRazon = GetValue(rs, indice, tablaProveedores, "razon")

        
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
        
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 8)).Font.Bold = True
    xlWorksheet.Range(xlWorksheet.Cells(offset, 1), xlWorksheet.Cells(offset, 8)).Interior.Color = &HC0C0C0

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
        
        xlWorksheet.Cells(offset, 3).value = "N° " & transf.CuentaBancaria & " | " & transf.NombreBanco
        xlWorksheet.Cells(offset, 4).value = transf.FechaOperacion
        xlWorksheet.Cells(offset, 5).value = transf.moneda.NombreCorto
        xlWorksheet.Cells(offset, 6).value = transf.Monto
        xlWorksheet.Cells(offset, 7).value = transf.Comprobante
        
        If transf.LiquidacionCaja Is Nothing Then
             xlWorksheet.Cells(offset, 8).value = transf.OrdenPago.Id
        Else
            xlWorksheet.Cells(offset, 8).value = transf.LiquidacionCaja.NumeroLiq
        End If
        
        frmLoading.ProgressBar.value = i
        
    Next

        xlWorksheet.Range(xlWorksheet.Cells(initoffset, 1), xlWorksheet.Cells(offset, 8)).Borders.LineStyle = xlContinuous

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
