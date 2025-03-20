VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasAdeudadas2 
   Caption         =   "Cashflow"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17670
   Icon            =   "frmAdminFacturasAdeudadas2.frx":0000
   LinkTopic       =   "Facturas Adeudadas"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   17670
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.PushButton Copiar 
      Height          =   375
      Left            =   15000
      TabIndex        =   19
      Top             =   285
      Width           =   1215
      _Version        =   786432
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin MSChart20Lib.MSChart grafico 
      Height          =   6510
      Left            =   30
      OleObjectBlob   =   "frmAdminFacturasAdeudadas2.frx":000C
      TabIndex        =   0
      Top             =   1005
      Width           =   17595
   End
   Begin XtremeSuiteControls.PushButton btnGenerar 
      Height          =   390
      Left            =   16260
      TabIndex        =   9
      Top             =   285
      Width           =   1260
      _Version        =   786432
      _ExtentX        =   2222
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "Generar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox grpResumen 
      Height          =   795
      Left            =   4680
      TabIndex        =   6
      Top             =   60
      Width           =   1995
      _Version        =   786432
      _ExtentX        =   3519
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Resumen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton rdoDiario 
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Diario"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdoMensual 
         Height          =   225
         Left            =   990
         TabIndex        =   8
         Top             =   330
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Mensual"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox grpRango 
      Height          =   795
      Left            =   135
      TabIndex        =   1
      Top             =   60
      Width           =   4455
      _Version        =   786432
      _ExtentX        =   7858
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Rango de Fechas"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   300
         Left            =   765
         TabIndex        =   4
         Top             =   300
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   529
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40280.4952083333
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   300
         Left            =   2925
         TabIndex        =   5
         Top             =   300
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   529
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40280.4952083333
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   2355
         TabIndex        =   3
         Top             =   330
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   330
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox grpLineas 
      Height          =   795
      Left            =   6765
      TabIndex        =   10
      Top             =   60
      Width           =   4230
      _Version        =   786432
      _ExtentX        =   7461
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Resumen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkCobros 
         Height          =   195
         Left            =   165
         TabIndex        =   11
         Top             =   240
         Width           =   750
         _Version        =   786432
         _ExtentX        =   1323
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cobros"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkPagos 
         Height          =   195
         Left            =   1545
         TabIndex        =   12
         Top             =   240
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Pagos"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkRestaOT 
         Height          =   195
         Left            =   2835
         TabIndex        =   13
         Top             =   240
         Width           =   1260
         _Version        =   786432
         _ExtentX        =   2222
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "A Facturar OT"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkRegresionCobros 
         Height          =   195
         Left            =   345
         TabIndex        =   14
         Top             =   480
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Regresión"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkRegresionPagos 
         Height          =   195
         Left            =   1725
         TabIndex        =   15
         Top             =   480
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Regresión"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox ckhRegresionFacturarOT 
         Height          =   195
         Left            =   3015
         TabIndex        =   16
         Top             =   480
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Regresión"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   795
      Left            =   11085
      TabIndex        =   17
      Top             =   60
      Width           =   3390
      _Version        =   786432
      _ExtentX        =   5980
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Cliente"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   290
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnSacarCliente 
         Height          =   300
         Left            =   3060
         TabIndex        =   18
         Top             =   285
         Width           =   240
         _Version        =   786432
         _ExtentX        =   423
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAdminFacturasAdeudadas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TipoResumen
    TR_Diario
    TR_Mensual
End Enum

Private Sub btnGenerar_Click()
    Graficar
End Sub

Private Sub btnSacarCliente_Click()
    Me.cboClientes.ListIndex = -1
End Sub

Private Sub cboClientes_Click()
    If Me.cboClientes.ListIndex = -1 Then
        Me.chkPagos.Enabled = True
        Me.chkPagos.value = xtpChecked
        'Me.chkRegresionPagos.value = xtpChecked
        Me.chkRegresionPagos.Enabled = True
    Else
        Me.chkPagos.Enabled = False
        Me.chkPagos.value = xtpUnchecked
        Me.chkRegresionPagos.value = xtpUnchecked
        Me.chkRegresionPagos.Enabled = False
    End If
End Sub

Private Sub chkCobros_Click()
    On Error Resume Next
    Me.grafico.Plot.SeriesCollection(1).Position.Hidden = IIf(Me.chkCobros.value = xtpChecked, False, True)
End Sub

Private Sub chkPagos_Click()
    On Error Resume Next
    Me.grafico.Plot.SeriesCollection(2).Position.Hidden = IIf(Me.chkPagos.value = xtpChecked, False, True)
End Sub


Private Sub chkRegresionCobros_Click()
    On Error Resume Next
    With Me.grafico.Plot.SeriesCollection(1).StatLine
        .Flag = VtChStatsRegression    ' set stats lines to draw
        If Me.chkRegresionCobros.value = xtpChecked Then
            .Style(VtChStatsRegression) = VtPenStyleDitted
        Else
            .Style(VtChStatsRegression) = VtPenStyleNull
        End If
        .VtColor.Set 255, 0, 0
    End With
End Sub

Private Sub chkRegresionPagos_Click()
    On Error Resume Next
    With Me.grafico.Plot.SeriesCollection(2).StatLine
        .Flag = VtChStatsRegression    ' set stats lines to draw
        If Me.chkRegresionPagos.value = xtpChecked Then
            .Style(VtChStatsRegression) = VtPenStyleDitted
        Else
            .Style(VtChStatsRegression) = VtPenStyleNull
        End If
        .VtColor.Set 0, 255, 0
    End With
End Sub

Private Sub chkRestaOT_Click()
    On Error Resume Next
    Me.grafico.Plot.SeriesCollection(3).Position.Hidden = IIf(Me.chkRestaOT.value = xtpChecked, False, True)
End Sub

Private Sub ckhRegresionFacturarOT_Click()
    On Error Resume Next
    With Me.grafico.Plot.SeriesCollection(3).StatLine
        .Flag = VtChStatsRegression    ' set stats lines to draw
        If Me.ckhRegresionFacturarOT.value = xtpChecked Then
            .Style(VtChStatsRegression) = VtPenStyleDitted
        Else
            .Style(VtChStatsRegression) = VtPenStyleNull
        End If
        .VtColor.Set 0, 0, 255
    End With
End Sub

Private Sub Copiar_Click()
    On Error GoTo err1
    frmPrincipal.CD.ShowPrinter

    Printer.Orientation = 2
    Me.grafico.EditCopy
    Clipboard.GetData vbCFMetafile
    DoEvents


    Me.grafico.EditCopy
    DoEvents
    Dim Size As Integer
    Size = Printer.FontSize
    Printer.FontBold = True
    Printer.FontSize = Size + 3
    Printer.Print "CASHFLOW"
    Printer.FontSize = Size
    Printer.FontBold = False
    Dim Tipo As String
    If Me.rdoDiario.value = True Then
        Tipo = rdoDiario.caption
    Else
        Tipo = Me.rdoMensual.caption
    End If
    Printer.FontUnderline = True
    Printer.Print "Período:";
    Printer.FontUnderline = False
    Printer.Print "  Desde: " & Me.dtpDesde.value & "  Hasta:  " & Me.dtpHasta.value
    Printer.FontUnderline = True
    Printer.Print "Tipo de Resúmen:";
    Printer.FontUnderline = False
    Printer.Print "  " & Tipo
    If Me.cboClientes.ListIndex <> -1 Then
        Printer.FontUnderline = True
        Printer.Print "Cliente:";
        Printer.FontUnderline = False
        Printer.Print "  " & Me.cboClientes
    End If

    Tipo = vbNullString
    If Me.chkCobros.value = xtpChecked Then Tipo = Tipo & vbTab & Me.chkCobros.caption & Chr(10)
    If Me.chkPagos.value = xtpChecked Then Tipo = Tipo & vbTab & Me.chkPagos.caption & Chr(10)
    If Me.chkRestaOT.value = xtpChecked Then Tipo = Tipo & vbTab & Me.chkRestaOT.caption & Chr(10)


    Printer.FontUnderline = True
    Printer.Print "Mostrar: "

    Printer.FontUnderline = False
    Printer.Print Tipo





    Printer.PaintPicture Clipboard.GetData(), 0, Printer.CurrentY + 100, Printer.Width - 200, Printer.Height - 2000
    Printer.EndDoc
    Exit Sub
err1:
End Sub

Private Sub Form_Initialize()
    Me.dtpDesde.value = Date
    Me.dtpHasta.value = DateAdd("m", 6, Date)
End Sub


Private Sub Form_Load()
    Customize Me
    Me.grafico.ColumnCount = 0
    DAOCliente.llenarComboXtremeSuite Me.cboClientes
    Me.cboClientes.ListIndex = -1
End Sub

Private Sub Graficar()

    Dim cliente_id As Long
    If Me.cboClientes.ListIndex <> -1 Then
        cliente_id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    End If


    Dim facturas As Collection
    Dim Factura As Factura
    Dim valores As New Dictionary

    Dim tr As TipoResumen

    If Me.rdoDiario.value Then
        tr = TR_Diario
    Else
        tr = TR_Mensual
    End If

    Dim datos()
    Dim flowItems As New Collection
    Dim dateIndex As String
    Dim total As Double
    Dim item As CashFlowItem

    Dim fechaParaTrabajar As Date

    '------------------------------------------------------------------------

    Set facturas = DAOFactura.FindAllNoSaldadasNiVencidas(Me.dtpDesde.value, Me.dtpHasta.value, cliente_id)

    For Each Factura In facturas

        If Factura.EstaAtrasada Then
            fechaParaTrabajar = Factura.FechaPropuestaPago
        Else
            fechaParaTrabajar = Factura.Vencimiento
        End If

        Select Case tr
        Case TipoResumen.TR_Diario
            dateIndex = FormatDateTime(fechaParaTrabajar, vbShortDate)
        Case TipoResumen.TR_Mensual
            dateIndex = Year(fechaParaTrabajar) & "-" & IIf(Month(fechaParaTrabajar) < 10, "0" & Month(fechaParaTrabajar), Month(fechaParaTrabajar))
        End Select

        If Factura.Saldado = SaldadoParcial Then
            total = Factura.TotalEstatico.total - DAOFactura.PagosRealizados(Factura.Id)
        Else
            total = Factura.TotalEstatico.total
        End If
        'por los dolares
        total = MonedaConverter.Convertir(total, Factura.moneda.Id, MonedaConverter.Patron.Id)

        If Factura.TipoDocumento = tipoDocumentoContable.notaCredito Then total = total * -1

        If funciones.BuscarEnColeccion(flowItems, CStr(dateIndex)) Then
            Set item = flowItems.item(CStr(dateIndex))
            item.ValorCobro = item.ValorCobro + total
        Else
            Set item = New CashFlowItem
            item.FechaIndex = dateIndex
            item.ValorCobro = total
            flowItems.Add item, dateIndex
        End If

    Next Factura


    '------------------------------------------------------------------------


    Dim Orden As OrdenPago
    Dim ordenes As Collection
    'talvez como se hace una explosion de los origenes, habria qeu traer todas las ordenes y despues filtrar por el rango de fecha pero de los origenes
    Set ordenes = DAOOrdenPago.FindAll("ordenes_pago.fecha >= " & conectar.Escape(Me.dtpDesde.value) & " AND ordenes_pago.fecha <= " & conectar.Escape(Me.dtpHasta.value))

    Dim it As Collection

    For Each Orden In ordenes
        For Each it In Orden.TotalOrigenesDiscriminado

            Select Case tr
            Case TipoResumen.TR_Diario
                dateIndex = FormatDateTime(it.item(1), vbShortDate)
            Case TipoResumen.TR_Mensual
                dateIndex = Year(it.item(1)) & "-" & IIf(Month(it.item(1)) < 10, "0" & Month(it.item(1)), Month(it.item(1)))
            End Select

            If funciones.BuscarEnColeccion(flowItems, CStr(dateIndex)) Then
                Set item = flowItems.item(CStr(dateIndex))
                item.ValorPago = item.ValorPago + MonedaConverter.Convertir(it.item(2), Orden.moneda.Id, MonedaConverter.Patron.Id)
            Else
                Set item = New CashFlowItem
                item.FechaIndex = dateIndex
                item.ValorPago = MonedaConverter.Convertir(it.item(2), Orden.moneda.Id, MonedaConverter.Patron.Id)
                flowItems.Add item, dateIndex
            End If

        Next it
    Next Orden


    '------------------------------------------------------------------------


    'descartar marcos
    Dim detallesOrdenes As Collection
    Dim deta As DetalleOrdenTrabajo
    Dim F As String

    If cliente_id = 0 Then

        Set ordenes = New Collection
        Dim monedaTmp As clsMoneda

        F = DAODetalleOrdenTrabajo.TABLA_DETALLE_PEDIDO & "." & DAODetalleOrdenTrabajo.CAMPO_FECHA_ENTREGA & " >= " & conectar.Escape(Me.dtpDesde.value)
        F = F & " AND " & DAODetalleOrdenTrabajo.TABLA_DETALLE_PEDIDO & "." & DAODetalleOrdenTrabajo.CAMPO_FECHA_ENTREGA & " <= " & conectar.Escape(Me.dtpHasta.value)
        F = F & " AND " & DAODetalleOrdenTrabajo.TABLA_DETALLE_PEDIDO & ".IdDetalleOtPadre > -1"

        Set detallesOrdenes = DAODetalleOrdenTrabajo.FindAll(F)
        For Each deta In detallesOrdenes

            Select Case tr
            Case TipoResumen.TR_Diario
                dateIndex = FormatDateTime(deta.FechaEntrega, vbShortDate)
            Case TipoResumen.TR_Mensual
                dateIndex = Year(deta.FechaEntrega) & "-" & IIf(Month(deta.FechaEntrega) < 10, "0" & Month(deta.FechaEntrega), Month(deta.FechaEntrega))
            End Select

            'falta moneda

            If Not funciones.BuscarEnColeccion(ordenes, CStr(deta.OrdenTrabajo.Id)) Then
                ordenes.Add DAOOrdenTrabajo.FindById(deta.OrdenTrabajo.Id), CStr(deta.OrdenTrabajo.Id)
            End If

            Set monedaTmp = ordenes.item(CStr(deta.OrdenTrabajo.Id)).moneda

            If funciones.BuscarEnColeccion(flowItems, CStr(dateIndex)) Then
                Set item = flowItems.item(CStr(dateIndex))
                item.ValorAFacturarOT = item.ValorAFacturarOT + MonedaConverter.Convertir(deta.TotalConDescuento, monedaTmp.Id, MonedaConverter.Patron.Id)
            Else
                Set item = New CashFlowItem
                item.FechaIndex = dateIndex
                item.ValorAFacturarOT = MonedaConverter.Convertir(deta.TotalConDescuento, monedaTmp.Id, MonedaConverter.Patron.Id)
                flowItems.Add item, dateIndex
            End If
        Next deta

    Else
        'Dim ordenes As Collection
        Dim Ot As OrdenTrabajo

        Set ordenes = DAOOrdenTrabajo.FindAll(DAOOrdenTrabajo.TABLA_PEDIDO & "." & DAOOrdenTrabajo.CAMPO_CLIENTE_ID & " = " & cliente_id & " AND " & DAOOrdenTrabajo.TABLA_PEDIDO & ".id_ot_padre <> -1", , , , True)
        For Each Ot In ordenes
            For Each deta In Ot.detalles
                If deta.FechaEntrega >= Me.dtpDesde.value And deta.FechaEntrega <= Me.dtpHasta.value Then
                    Select Case tr
                    Case TipoResumen.TR_Diario
                        dateIndex = FormatDateTime(deta.FechaEntrega, vbShortDate)
                    Case TipoResumen.TR_Mensual
                        dateIndex = Year(deta.FechaEntrega) & "-" & IIf(Month(deta.FechaEntrega) < 10, "0" & Month(deta.FechaEntrega), Month(deta.FechaEntrega))
                    End Select

                    If funciones.BuscarEnColeccion(flowItems, CStr(dateIndex)) Then
                        Set item = flowItems.item(CStr(dateIndex))
                        item.ValorAFacturarOT = item.ValorAFacturarOT + MonedaConverter.Convertir(deta.TotalConDescuento, Ot.moneda.Id, MonedaConverter.Patron.Id)
                    Else
                        Set item = New CashFlowItem
                        item.FechaIndex = dateIndex
                        item.ValorAFacturarOT = MonedaConverter.Convertir(deta.TotalConDescuento, Ot.moneda.Id, MonedaConverter.Patron.Id)
                        flowItems.Add item, dateIndex
                    End If

                End If
            Next deta
        Next Ot

    End If

    '------------------------------------------------------------------------


    '----------------- begin sort
    Dim sortedFlowItems As New Collection
    Dim Item2 As CashFlowItem
    Dim idxToPut As String
    For Each item In flowItems

        idxToPut = vbNullString

        For Each Item2 In sortedFlowItems
            If DateIndexToLong(item.FechaIndex) <= DateIndexToLong(Item2.FechaIndex) Then
                idxToPut = Item2.FechaIndex
                Exit For
            End If
        Next Item2

        If LenB(idxToPut) = 0 Then
            sortedFlowItems.Add item, item.FechaIndex
        Else
            sortedFlowItems.Add item, item.FechaIndex, idxToPut
        End If

    Next item
    '----------------- end sort


    If sortedFlowItems.count = 0 Then
        Me.grafico.ColumnCount = 0
        Me.grafico.rowcount = 0
    Else
        ReDim datos(1 To sortedFlowItems.count, 1 To 4)

        Dim i As Long: i = 0
        For Each item In sortedFlowItems
            i = i + 1
            datos(i, 1) = item.FechaIndex
            datos(i, 2) = item.ValorCobro
            datos(i, 3) = item.ValorPago
            datos(i, 4) = item.ValorAFacturarOT
        Next item

        Me.grafico.ChartData = datos
        Me.grafico.ColumnCount = 3

        Me.grafico.ColumnLabelCount = 1
        Me.grafico.Column = 1
        Me.grafico.ColumnLabel = "Cobros"
        Me.grafico.Plot.SeriesCollection(1).SeriesMarker.Show = True
        Me.grafico.Plot.SeriesCollection(1).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint

        Me.grafico.Column = 2
        Me.grafico.ColumnLabel = "Pagos"
        Me.grafico.Plot.SeriesCollection(2).SeriesMarker.Show = True
        Me.grafico.Plot.SeriesCollection(2).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint

        Me.grafico.Column = 3
        Me.grafico.ColumnLabel = "A Facturar OT"
        Me.grafico.Plot.SeriesCollection(3).SeriesMarker.Show = True
        Me.grafico.Plot.SeriesCollection(3).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint

        Me.grafico.Refresh
    End If


End Sub

Private Function DateIndexToLong(dateIdx As String) As Long
    Dim valor1 As Long

    If Len(dateIdx) <= 7 Then    'aaaa-mm
        valor1 = Replace(dateIdx, "-", vbNullString)
    Else
        valor1 = Mid(dateIdx, 7, 4) & Mid(dateIdx, 4, 2) & Mid(dateIdx, 1, 2)
    End If

    DateIndexToLong = valor1
End Function


Private Sub Form_Resize()
    Me.grafico.Width = Me.Width - 175
    Me.grafico.Height = Me.Height - 1550

    Me.btnGenerar.Left = Me.Width - 300 - Me.btnGenerar.Width
    Me.Copiar.Left = (Me.GroupBox1.Width + Me.grpLineas.Width + Me.grpRango.Width + Me.grpResumen.Width + 600)
    Me.btnGenerar.Left = (Me.GroupBox1.Width + Me.grpLineas.Width + Me.grpRango.Width + Me.grpResumen.Width + 800 + Me.Copiar.Width)
End Sub



