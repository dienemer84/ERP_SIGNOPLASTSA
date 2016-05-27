VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAdminSubdiarioCompras 
   Caption         =   "Subdiario IVA Compras"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15090
   Icon            =   "frmAdminSubdiarioCompras.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   15090
   Begin XtremeSuiteControls.GroupBox grpTotales 
      Height          =   1695
      Left            =   8745
      TabIndex        =   0
      Top             =   7005
      Width           =   6255
      _Version        =   786432
      _ExtentX        =   11033
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Totales"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label lblRedondeo 
         Height          =   195
         Left            =   4905
         TabIndex        =   32
         Top             =   765
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   3315
         TabIndex        =   31
         Top             =   765
         Width           =   795
         _Version        =   786432
         _ExtentX        =   1402
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Redondeo:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   3315
         TabIndex        =   30
         Top             =   495
         Width           =   1230
         _Version        =   786432
         _ExtentX        =   2170
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Impuesto Interno:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblImpuestoInterno 
         Height          =   195
         Left            =   4905
         TabIndex        =   29
         Top             =   495
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblExento 
         Height          =   195
         Left            =   4905
         TabIndex        =   28
         Top             =   225
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   3315
         TabIndex        =   27
         Top             =   225
         Width           =   540
         _Version        =   786432
         _ExtentX        =   953
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Exento:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIVA 
         Height          =   195
         Left            =   195
         TabIndex        =   26
         Top             =   1035
         Width           =   315
         _Version        =   786432
         _ExtentX        =   556
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "IVA:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIVATotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   25
         Top             =   1035
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblNetoGravado 
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   1065
         _Version        =   786432
         _ExtentX        =   1879
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Neto Gravado:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPercepcionesIIBB 
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   1425
         _Version        =   786432
         _ExtentX        =   2514
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Percepciones Total:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblTotal 
         Height          =   195
         Left            =   3315
         TabIndex        =   4
         Top             =   1410
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total:"
         AutoSize        =   -1  'True
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   6090
         X2              =   165
         Y1              =   1335
         Y2              =   1335
      End
      Begin XtremeSuiteControls.Label lblNetoGravadoTotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   3
         Top             =   225
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblPercepcionesTotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   2
         Top             =   480
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblTotalTotal 
         Height          =   195
         Left            =   4920
         TabIndex        =   1
         Top             =   1395
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4350
      Left            =   0
      TabIndex        =   7
      Top             =   1755
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   7673
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmAdminSubdiarioCompras.frx":000C
      Column(2)       =   "frmAdminSubdiarioCompras.frx":0164
      Column(3)       =   "frmAdminSubdiarioCompras.frx":025C
      Column(4)       =   "frmAdminSubdiarioCompras.frx":0358
      Column(5)       =   "frmAdminSubdiarioCompras.frx":0444
      Column(6)       =   "frmAdminSubdiarioCompras.frx":053C
      Column(7)       =   "frmAdminSubdiarioCompras.frx":0728
      GroupCount      =   1
      Group(1)        =   "frmAdminSubdiarioCompras.frx":0854
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdminSubdiarioCompras.frx":08BC
      FormatStyle(2)  =   "frmAdminSubdiarioCompras.frx":09F4
      FormatStyle(3)  =   "frmAdminSubdiarioCompras.frx":0AA4
      FormatStyle(4)  =   "frmAdminSubdiarioCompras.frx":0B58
      FormatStyle(5)  =   "frmAdminSubdiarioCompras.frx":0C30
      FormatStyle(6)  =   "frmAdminSubdiarioCompras.frx":0CE8
      FormatStyle(7)  =   "frmAdminSubdiarioCompras.frx":0DC8
      ImageCount      =   0
      PrinterProperties=   "frmAdminSubdiarioCompras.frx":0EA8
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1680
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   14940
      _Version        =   786432
      _ExtentX        =   26352
      _ExtentY        =   2963
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkOcultarIVACero 
         Height          =   255
         Left            =   8865
         TabIndex        =   33
         Top             =   270
         Width           =   2730
         _Version        =   786432
         _ExtentX        =   4815
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Listado Percepciones de IVA"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdoRangoFechas 
         Height          =   255
         Left            =   420
         TabIndex        =   9
         Top             =   285
         Width           =   1725
         _Version        =   786432
         _ExtentX        =   3043
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Por rango de fechas"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnMostrar 
         Height          =   360
         Left            =   8850
         TabIndex        =   10
         Top             =   840
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Mostrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   360
         Left            =   11100
         TabIndex        =   11
         Top             =   825
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   360
         Left            =   11100
         TabIndex        =   12
         Top             =   1200
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdoLiquidacion 
         Height          =   255
         Left            =   3225
         TabIndex        =   13
         Top             =   285
         Width           =   1290
         _Version        =   786432
         _ExtentX        =   2275
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Por liquidacion"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox grpRangoFechas 
         Height          =   1140
         Left            =   255
         TabIndex        =   14
         Top             =   330
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4577
         _ExtentY        =   2011
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   915
            TabIndex        =   15
            Top             =   300
            Width           =   1440
            _Version        =   786432
            _ExtentX        =   2540
            _ExtentY        =   556
            _StockProps     =   68
            Format          =   1
            CurrentDate     =   40241
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   300
            Left            =   915
            TabIndex        =   16
            Top             =   660
            Width           =   1440
            _Version        =   786432
            _ExtentX        =   2540
            _ExtentY        =   529
            _StockProps     =   68
            Format          =   1
            CurrentDate     =   40241
         End
         Begin XtremeSuiteControls.Label lblDesde 
            Height          =   195
            Left            =   270
            TabIndex        =   18
            Top             =   345
            Width           =   510
            _Version        =   786432
            _ExtentX        =   900
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde:"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   675
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta:"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox grpLiquidacion 
         Height          =   1140
         Left            =   3045
         TabIndex        =   19
         Top             =   330
         Width           =   5565
         _Version        =   786432
         _ExtentX        =   9816
         _ExtentY        =   2011
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cboLiquidaciones 
            Height          =   315
            Left            =   1170
            TabIndex        =   20
            Top             =   450
            Width           =   4185
            _Version        =   786432
            _ExtentX        =   7382
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "cboLiquidaciones"
         End
         Begin XtremeSuiteControls.Label lblLiquidacion 
            Height          =   195
            Left            =   225
            TabIndex        =   21
            Top             =   495
            Width           =   840
            _Version        =   786432
            _ExtentX        =   1482
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Liquidacion:"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton btnGuardarLiquidacion 
         Height          =   360
         Left            =   8850
         TabIndex        =   22
         Top             =   1200
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Generar liquidación"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkOcultarIIBBCero 
         Height          =   255
         Left            =   8865
         TabIndex        =   34
         Top             =   525
         Width           =   2730
         _Version        =   786432
         _ExtentX        =   4815
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Listado Percepciones de IIBB"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox grpTotalesIVA 
      Height          =   1695
      Left            =   5460
      TabIndex        =   23
      Top             =   7005
      Width           =   3180
      _Version        =   786432
      _ExtentX        =   5609
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Totales Alicuotas IVA"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridTotalesIVA 
         Height          =   1380
         Left            =   105
         TabIndex        =   24
         Top             =   225
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   2434
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   3
         Column(1)       =   "frmAdminSubdiarioCompras.frx":1080
         Column(2)       =   "frmAdminSubdiarioCompras.frx":11BC
         Column(3)       =   "frmAdminSubdiarioCompras.frx":12E0
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminSubdiarioCompras.frx":13F4
         FormatStyle(2)  =   "frmAdminSubdiarioCompras.frx":152C
         FormatStyle(3)  =   "frmAdminSubdiarioCompras.frx":15DC
         FormatStyle(4)  =   "frmAdminSubdiarioCompras.frx":1690
         FormatStyle(5)  =   "frmAdminSubdiarioCompras.frx":1768
         FormatStyle(6)  =   "frmAdminSubdiarioCompras.frx":1820
         ImageCount      =   0
         PrinterProperties=   "frmAdminSubdiarioCompras.frx":1900
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1695
      Left            =   2055
      TabIndex        =   35
      Top             =   6975
      Width           =   3225
      _Version        =   786432
      _ExtentX        =   5689
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Totales Percepciones"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX GridEX2 
         Height          =   1380
         Left            =   105
         TabIndex        =   36
         Top             =   225
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   2434
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAdminSubdiarioCompras.frx":1AD8
         Column(2)       =   "frmAdminSubdiarioCompras.frx":1C18
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminSubdiarioCompras.frx":1D34
         FormatStyle(2)  =   "frmAdminSubdiarioCompras.frx":1E6C
         FormatStyle(3)  =   "frmAdminSubdiarioCompras.frx":1F1C
         FormatStyle(4)  =   "frmAdminSubdiarioCompras.frx":1FD0
         FormatStyle(5)  =   "frmAdminSubdiarioCompras.frx":20A8
         FormatStyle(6)  =   "frmAdminSubdiarioCompras.frx":2160
         ImageCount      =   0
         PrinterProperties=   "frmAdminSubdiarioCompras.frx":2240
      End
   End
End
Attribute VB_Name = "frmAdminSubdiarioCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col As New Collection
Dim item As SubdiarioVentasDetalle
Private liqui As LiquidacionSubdiarioVenta
Private liquidaciones As Collection
Private desdeAbsoluto As Date
Private totales As New Dictionary
Private totalesIva As Collection
Dim Factura As clsFacturaProveedor
Private dataFiltered As Boolean
Private alicuotas As Collection
Dim percepciones As New Collection
Dim totalesper As New Collection
Private Enum PosicionTotales

    TotNetoGravado = 0
    totIva = 1
    totPercep = 2
    TotExento = 3
    TotTot = 4
    TotImpuestoInterno = 6
    TotRedondeo = 7
End Enum
Private Sub Totalizar()
    Dim sumNeto As Double: sumNeto = 0
    Dim sumIVA As Double: sumIVA = 0
    Dim sumPercep As Double: sumPercep = 0
    Dim sumPercepIVA As Double: sumPercepIVA = 0
    Dim sumImpuestoInterno As Double: sumImpuestoInterno = 0
    Dim sumExento As Double: sumExento = 0
    Dim sumTot As Double: sumTot = 0
    Dim sumRedondeo As Double: sumRedondeo = 0
    Dim tmpValue As Double
    Dim i As SubdiarioVentasDetalle
    Dim alis21 As New Collection
    Dim c As Double
    Dim ali As Variant
    Dim per As clsPercepciones

    Set totales = New Dictionary
    Set totalesIva = New Collection

    For Each ali In alicuotas


        'If ali = 10.5 Then ali = 11
        totalesIva.Add 0, CStr(ali)
    Next ali

    Set percepciones = DAOPercepciones.GetAll

    Dim dtop As DTOPercepcionImporte
    Set totalesper = New Collection
    For Each per In percepciones
        Set dtop = New DTOPercepcionImporte
        dtop.Importe = 0
        Set dtop.Percepcion = per
        totalesper.Add dtop, CStr(per.id)
    Next

    Dim pera As clsPercepcionesAplicadas
    For Each i In col
        sumNeto = sumNeto + i.NetoGravado
        sumIVA = sumIVA + (i.Iva)
        sumPercep = sumPercep + i.percepciones
        ' sumPercepIVA = sumPercepIVA + i.PercepcionesIVA

        sumExento = sumExento + i.AlicuotasIva(CStr(0))

        sumTot = sumTot + i.Total
        sumImpuestoInterno = sumImpuestoInterno + i.ImpuestoInterno
        sumRedondeo = sumRedondeo + i.Redondeo

        For Each pera In i.ListaPercepciones

            Set dtop = totalesper(CStr(pera.Percepcion.id))
            'tmpValue = funciones.RedondearDecimales(totalesper(CStr(pera.Percepcion.Id)).importe)
            totalesper.remove CStr(pera.Percepcion.id)


            dtop.Importe = funciones.RedondearDecimales(dtop.Importe + pera.Monto)
            totalesper.Add dtop, CStr(pera.Percepcion.id)
        Next





        For Each ali In alicuotas
            'If ali = 10.5 Then ali = 11
            tmpValue = funciones.RedondearDecimales(totalesIva(CStr(ali)))
            totalesIva.remove CStr(ali)
            totalesIva.Add funciones.RedondearDecimales(i.AlicuotasIva(CStr(ali))) + tmpValue, CStr(ali)
        Next ali
    Next i

    totales.Add PosicionTotales.TotNetoGravado, sumNeto
    totales.Add PosicionTotales.totIva, sumIVA
    totales.Add PosicionTotales.totPercep, sumPercep
    totales.Add PosicionTotales.TotExento, sumExento
    totales.Add PosicionTotales.TotTot, sumTot
    totales.Add PosicionTotales.TotImpuestoInterno, sumImpuestoInterno
    totales.Add PosicionTotales.TotRedondeo, sumRedondeo

    Me.lblNetoGravadoTotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotNetoGravado))
    Me.lblIVATotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.totIva))
    Me.lblPercepcionesTotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.totPercep))
    'Me.lblPercepcionesIVA.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotPercepIVA))
    Me.lblImpuestoInterno.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotImpuestoInterno))
    Me.lblExento.caption = funciones.FormatearDecimales(totalesIva(CStr(0)))
    Me.lblTotalTotal.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotTot))
    Me.lblRedondeo.caption = funciones.FormatearDecimales(totales.item(PosicionTotales.TotRedondeo))

End Sub


Private Sub btnExportar_Click()
    MsgBox "en construcción"
    '    ExportaSubDiarioVentas
End Sub

Private Sub btnGuardarLiquidacion_Click()
    If col.count = 0 Then
        MsgBox "No hay detalles para poder guardar la liquidacion", vbExclamation + vbOKOnly
        Exit Sub
    End If


    If dataFiltered Then
        MsgBox "Al mostrar el listado fueron activados alguno de los filtros de ocultamiento de percepciones." & vbNewLine & "Vuelva a mostrar el listado sin los filtros activados.", vbExclamation
        Exit Sub
    End If


    Dim l As New LiquidacionSubdiarioVenta
    Dim nombre As String
    nombre = InputBox("Ingrese una descripcion para la liquidacion", "Descripcion de liquidacion")
    If LenB(nombre) = 0 Then
        MsgBox "Debe ingresar un nombre para la liquidacion", vbExclamation
    Else
        l.nombre = nombre
        l.desde = Me.dtpDesde.value
        l.hasta = Me.dtpHasta.value
        l.EsDeVenta = False
        Set l.Detalles = col
        If DAOSubdiarios.Guardar(l) Then
            SetearMaxDesde
            MsgBox "La liquidacion se guardó con éxito", vbInformation + vbOKOnly
            CargarLiquidaciones
            If Me.cboLiquidaciones.ListCount > 0 Then
                Me.cboLiquidaciones.ListIndex = Me.cboLiquidaciones.ListCount - 1
            End If
            rdoLiquidacion.value = True
            llenarLista
        Else
            MsgBox "Error al guardar la liquidacion", vbOKOnly + vbCritical
        End If
    End If
End Sub




Private Sub Form_Load()

    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False

    GridEXHelper.AutoSizeColumns Me.GridEX1, True

    GridEXHelper.CustomizeGrid Me.gridTotalesIVA
    GridEXHelper.CustomizeGrid Me.GridEX2
    SetearMaxDesde
    CargarLiquidaciones


    Set alicuotas = DAOFacturaProveedor.FindAllAlicuotasIVA()
    Dim ali As Variant
    Dim col As JSColumn
    For Each ali In alicuotas
        If ali = 0 Then
            Set col = Me.GridEX1.Columns.Add("Exento", jgexText, jgexEditNone, "IVA_" & ali)
            col.TextAlignment = jgexAlignRight
            col.AggregateFunction = jgexSum
            col.GroupFormat = "0.00"
            col.TotalRowFormat = "0.00"
        Else

            '            If ali = 10.5 Then ali = 11



            'Set col = Me.GridEX1.Columns.Add("NG " & ali & "%", jgexText, jgexEditNone, "NG_" & ali )
            Set col = Me.GridEX1.Columns.Add("NG", jgexText, jgexEditNone, "NG_" & ali)
            col.TextAlignment = jgexAlignRight
            col.AggregateFunction = jgexSum
            col.GroupFormat = "0.00"
            col.TotalRowFormat = "0.00"


            Set col = Me.GridEX1.Columns.Add("IVA " & ali & "%", jgexText, jgexEditNone, "IVA_" & ali)
            col.TextAlignment = jgexAlignRight
            col.AggregateFunction = jgexSum
            col.GroupFormat = "0.00"
            col.TotalRowFormat = "0.00"
        End If

    Next

    '    Set col = Me.GridEX1.Columns.Add("Per IIBB", jgexText, jgexEditNone, "percepcionesiibb")
    '    col.TextAlignment = jgexAlignRight
    '    col.AggregateFunction = jgexSum
    '    col.GroupFormat = "0.00"
    '    col.TotalRowFormat = "0.00"
    '
    '    Set col = Me.GridEX1.Columns.Add("Percep IVA", jgexText, jgexEditNone, "percepcionesiva")
    '    col.TextAlignment = jgexAlignRight
    '    col.AggregateFunction = jgexSum
    '    col.GroupFormat = "0.00"
    '    col.TotalRowFormat = "0.00"


    Dim cole As New Collection
    Set cole = DAOPercepciones.GetAll
    Dim per As clsPercepciones
    For Each per In cole

        Set col = Me.GridEX1.Columns.Add(per.Percepcion, jgexText, jgexEditNone, "PER_" & per.id)
        col.TextAlignment = jgexAlignRight
        col.AggregateFunction = jgexSum
        col.GroupFormat = "0.00"
        col.TotalRowFormat = "0.00"
        col.Tag = per.id

    Next per





    Set col = Me.GridEX1.Columns.Add("Imp. Int.", jgexText, jgexEditNone, "impuestointerno")
    col.TextAlignment = jgexAlignRight
    col.AggregateFunction = jgexSum
    col.GroupFormat = "0.00"
    col.TotalRowFormat = "0.00"


    Set col = Me.GridEX1.Columns.Add("Total", jgexText, jgexEditNone, "total")
    col.TextAlignment = jgexAlignRight
    col.AggregateFunction = jgexSum
    col.GroupFormat = "0.00"
    col.TotalRowFormat = "0.00"


    Me.rdoRangoFechas.value = True
    Me.GridEX1.ItemCount = 0
    Me.gridTotalesIVA.ItemCount = 0
End Sub

Private Sub SetearMaxDesde()
    desdeAbsoluto = DAOSubdiarios.MaxFechaLiqui(False)

    If CDbl(desdeAbsoluto) <> 0 Then
        Me.dtpDesde.value = desdeAbsoluto
        Me.dtpDesde.MinDate = desdeAbsoluto
        'Me.dtpDesde.MaxDate = desdeAbsoluto
        Me.dtpHasta.MinDate = desdeAbsoluto
    Else
        Me.dtpDesde.value = DateSerial(Year(Date), Month(Date), 1)
    End If

    If CLng(desdeAbsoluto) >= CLng(Date) Then
        Me.dtpHasta.value = DateAdd("d", 1, desdeAbsoluto)
    Else
        Me.dtpHasta.value = Date
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.GroupBox1.Width = Me.ScaleWidth - 100

    Me.GridEX1.Width = Me.ScaleWidth - 100
    Me.GridEX1.Height = Me.ScaleHeight - 3600

    Me.grpTotales.Left = Me.ScaleWidth - Me.grpTotales.Width - 150
    Me.grpTotales.Top = Me.ScaleHeight - Me.grpTotales.Height - 100



    Me.grpTotalesIVA.Left = Me.ScaleWidth - Me.grpTotalesIVA.Width - 6500
    Me.grpTotalesIVA.Top = Me.ScaleHeight - Me.grpTotalesIVA.Height - 100
    Me.GroupBox2.Left = Me.ScaleWidth - Me.GroupBox2.Width - 9800
    Me.GroupBox2.Top = Me.grpTotalesIVA.Top

End Sub

Private Sub llenarLista()
    Dim i As Long

    If Me.rdoRangoFechas.value Then
        Set col = DAOSubdiarios.SubDiarioCompras(Me.dtpDesde.value, Me.dtpHasta.value)
    Else
        If Me.cboLiquidaciones.ListIndex <> -1 Then
            Set liqui = liquidaciones.item(CStr(Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.ListIndex)))
            Set col = liqui.Detalles
        Else
            Set col = New Collection
        End If
    End If

    dataFiltered = False

    If chkOcultarIVACero.value = xtpChecked Or chkOcultarIIBBCero.value = xtpChecked Then
        For i = col.count To 1 Step -1
            If (Not col.item(i).TienePercepcionesIVA And chkOcultarIVACero.value = xtpChecked) Or _
               (Not col.item(i).TienePercepcionesIIBB And chkOcultarIIBBCero.value = xtpChecked) _
               Then
                col.remove i
                dataFiltered = True
            End If
        Next i
    End If

    Me.GridEX1.ItemCount = 0
    If col.count > 0 Then Me.GridEX1.ItemCount = col.count
    GridEXHelper.AutoSizeColumns Me.GridEX1


    Totalizar
    Me.gridTotalesIVA.ItemCount = 0
    Me.gridTotalesIVA.ItemCount = alicuotas.count - 1    'para que no me agregue la alicuota 0%

    Me.GridEX2.ItemCount = 0
    Me.GridEX2.ItemCount = percepciones.count
    GridEXHelper.AutoSizeColumns Me.gridTotalesIVA
    Me.GridEX1.Columns(3).Width = Me.GridEX1.Columns(3).Width / 2

    Me.caption = "Subdiario IVA Compras (" & col.count & " comprobantes encontrados)"
End Sub

Private Sub GridEX1_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    Me.GridEX1.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & " de " & nPages
End Sub

Private Sub GridEX1_DblClick()
    If col.count > 0 Then

        Set Factura = DAOFacturaProveedor.FindById(col(Me.GridEX1.RowIndex(Me.GridEX1.row)).FacturaId)

        Dim frm As frmAdminComprasNuevaFCProveedor
        Set frm = New frmAdminComprasNuevaFCProveedor

        frm.ver = True
        frm.Factura = Factura
        frm.Show
    End If
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.GridEX1
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If
End Sub



Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
    '    If RowBuffer.RowIndex > 0 And col.count > 0 Then
    '        Set Item = col.Item(RowBuffer.RowIndex)
    '        If Item.estado = Anulada Then
    '            RowBuffer.RowStyle = "anulada"
    '        End If
    '    End If
End Sub

Private Sub GridEX1_SelectionChange()
    If Me.GridEX1.row <> -1 Then
        If Me.GridEX1.RowIndex(Me.GridEX1.row) <> 0 Then
            Set item = col.item(Me.GridEX1.RowIndex(Me.GridEX1.row))
        End If
    End If
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    'On Error Resume Next
    If col.count > 0 Then
        Set item = col.item(RowIndex)

        Values(1) = item.FEcha
        Values(2) = item.Comprobante
        Values(3) = funciones.RazonSocialFormateada(item.RazonSocial)
        Values(4) = item.Cuit
        Values(5) = item.CondicionIva
        Values(6) = funciones.FormatearDecimales(item.NetoGravado)

        Values(7) = " "

        Dim ali As Variant
        For Each ali In alicuotas
            If ali <> 0 Then

                'If ali = 10.5 Then ali = 11

                Values(Me.GridEX1.Columns.item("NG_" & ali).index) = funciones.FormatearDecimales(item.NetosGravado.item(CStr(ali)))
            End If
            Values(Me.GridEX1.Columns.item("IVA_" & ali).index) = funciones.FormatearDecimales(funciones.RedondearDecimales(item.AlicuotasIva.item(CStr(ali))))
        Next

        '        Values(Me.GridEX1.Columns.Item("percepcionesiibb").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, Item.PercepcionesIB))
        '        Values(Me.GridEX1.Columns.Item("percepcionesiva").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, Item.PercepcionesIVA))
        '        Values(Me.GridEX1.Columns.Item("impuestointerno").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, Item.ImpuestoInterno))
        '        Values(Me.GridEX1.Columns.Item("total").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, (Item.Total)))



        Dim colper As New Collection
        Dim per As clsPercepciones
        Set colper = DAOPercepciones.GetAll

        For Each per In colper
            '  Debug.Print per.Id


            If BuscarEnColeccion(item.ListaPercepciones, CStr(per.id)) Then
                Values(Me.GridEX1.Columns.item("per_" & per.id).index) = funciones.FormatearDecimales(item.ListaPercepciones(CStr(per.id)).Monto)
            Else
                Values(Me.GridEX1.Columns.item("per_" & per.id).index) = 0
            End If



            'Values(Me.GridEX1.Columns.item("percepcionesiibb").index) = funciones.FormatearDecimales(item.Percepciones)
        Next

        'Values(Me.GridEX1.Columns.item("percepcionesiva").index) = funciones.FormatearDecimales(item.PercepcionesIVA)


        Values(Me.GridEX1.Columns.item("impuestointerno").index) = funciones.FormatearDecimales(item.ImpuestoInterno)
        Values(Me.GridEX1.Columns.item("total").index) = funciones.FormatearDecimales(item.Total)
    End If

End Sub
Private Sub btnMostrar_Click()
    llenarLista
End Sub


Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 Then
        '   If Item.estado = Anulada Then
        If MsgBox("¿Desea realmente actualizar los valores del item?", vbYesNo + vbQuestion) = vbYes Then
            Set item = col.item(RowIndex)
            item.NetoGravado = Values(6)
            item.Iva = Values(7)
            item.percepciones = Values(8)
            item.Exento = Values(9)
            item.Total = Values(10)

            DAOSubdiarios.UpdateDetalle item
            Totalizar
        End If
        '    End If
    End If
End Sub



Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim dtop As New DTOPercepcionImporte




    Values(2) = funciones.FormatearDecimales(totalesper(RowIndex).Importe)    '/ (va / 100))
    If IsSomething(totalesper(RowIndex).Percepcion) Then
        Values(1) = totalesper(RowIndex).Percepcion.Percepcion
    Else
        Values(1) = funciones.FormatearDecimales(0)
    End If

End Sub

Private Sub gridTotalesIVA_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim va As Variant
    va = alicuotas(RowIndex)

    Values(1) = funciones.FormatearDecimales(totalesIva(RowIndex) / (va / 100))
    Values(2) = va & "%"

    Values(3) = totalesIva(RowIndex)

End Sub

Private Sub PushButton2_Click()


    With Me.GridEX1.PrinterProperties
        .HeaderDistance = 300
        .FooterDistance = 500

        .TopMargin = 600
        .BottomMargin = 500
        .LeftMargin = 300
        .RightMargin = 300

        .FitColumns = True
        .DocumentName = "Subdiario IVA Compras"
        .PrintProgressDialog = True

        .RepeatHeaders = True
        .Orientation = jgexPPLandscape

        Dim header As String
        If Me.rdoRangoFechas.value Then
            header = Me.dtpDesde.value & " a " & Me.dtpHasta.value
        Else
            header = "Liquidación " & Me.cboLiquidaciones.text
        End If


        Dim tipoListado As String
        If Me.chkOcultarIIBBCero.value = xtpChecked Or _
           Me.chkOcultarIVACero.value = xtpChecked Then
            tipoListado = " [Listado Percepciones de"

            If Me.chkOcultarIVACero.value = xtpChecked Then
                tipoListado = tipoListado & " IVA"
            End If

            If Me.chkOcultarIIBBCero.value = xtpChecked Then
                If Me.chkOcultarIVACero.value = xtpChecked Then
                    tipoListado = tipoListado & " e"
                End If
                tipoListado = tipoListado & " IIBB"
            End If

            tipoListado = tipoListado & "]"
        End If


        .HeaderString(jgexHFCenter) = "Subdiario IVA Compras (" & header & ")" & tipoListado
        '  .FooterString(jgexHFLeft) = "Impreso el " & Now

    End With

    Dim F As New frmPrintPreview
    Me.GridEX1.PrintPreview F.GEXPreview1, Me.GridEX1.SelectedItems.count > 1
    F.WindowState = 2
    F.Show

End Sub

Private Sub rdoLiquidacion_Click()
    ActualizarFrames
    If Me.rdoLiquidacion.value And Me.cboLiquidaciones.ListIndex = -1 And Me.cboLiquidaciones.ListCount > 0 Then
        Me.cboLiquidaciones.ListIndex = 0
    End If
End Sub

Private Sub rdoRangoFechas_Click()
    ActualizarFrames
End Sub

Private Sub ActualizarFrames()
    Me.GridEX1.ItemCount = 0
    Me.gridTotalesIVA.ItemCount = 0
    Me.GridEX2.ItemCount = 0
    Set col = New Collection
    Me.grpRangoFechas.Enabled = Me.rdoRangoFechas.value
    Me.grpLiquidacion.Enabled = Me.rdoLiquidacion.value
    Me.btnGuardarLiquidacion.Enabled = Me.rdoRangoFechas.value


    Me.lblIVATotal.caption = ".-"
    Me.lblNetoGravadoTotal.caption = ".-"
    Me.lblTotalTotal.caption = ".-"
    Me.lblPercepcionesTotal.caption = ".-"
    Me.lblExento.caption = ".-"
    'Me.lblPercepcionesIVA.caption = ".-"
    Me.lblImpuestoInterno.caption = ".-"
    Me.lblRedondeo.caption = ".-"

    Me.GridEX1.AllowEdit = Me.rdoLiquidacion.value
    Dim Column As JSColumn
    For Each Column In Me.GridEX1.Columns
        Column.EditType = jgexEditNone
    Next Column

    Me.GridEX1.Columns("neto_gravado").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    Dim ali As Variant
    For Each ali In alicuotas
        'If ali = 10.5 Then ali = 11
        Me.GridEX1.Columns.item("IVA_" & ali).EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    Next
    '
    'Me.GridEX1.Columns("percepcionesiibb").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    'Me.GridEX1.Columns("percepcionesiva").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)

    Me.GridEX1.Columns("impuestointerno").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
    Me.GridEX1.Columns("total").EditType = IIf(Me.rdoRangoFechas.value, jgexEditTypeConstants.jgexEditNone, jgexEditTypeConstants.jgexEditTextBox)
End Sub

Private Sub CargarLiquidaciones()
    Me.cboLiquidaciones.Clear

    Set liquidaciones = DAOSubdiarios.FindAllLiquidacionesVenta(False)
    For Each liqui In liquidaciones
        Me.cboLiquidaciones.AddItem liqui.nombre & " (" & liqui.desde & " a " & liqui.hasta & ")"
        Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.NewIndex) = liqui.id
    Next liqui

End Sub


Public Function ExportaSubDiarioVentas() As Boolean
    If Not rdoLiquidacion.value Then Exit Function

    On Error GoTo errEXCEL
    Dim xlb As New Excel.Workbook
    Dim xla As New Excel.Worksheet
    Dim xls As New Excel.Application

    Dim a As String
    Dim b As String
    Dim offset As Long
    Dim strMsg As String
    Dim CDLGMAIN As CommonDialog
    Dim sFilter As String


    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate


    With xla

        .Range("A1:Q1").Merge
        .Range("A2:Q2").Merge
        .Range("A1:Q3").HorizontalAlignment = xlHAlignCenter
        .Range("A1:Q2").Font.Bold = True
        .Range("A3:Q2").Font.Bold = True


        .cells(1, 1).value = "SIGNOPLAST S.A. Subdiario compras" & IIf(Me.rdoRangoFechas.value, " (NO LIQUIDADO)", vbNullString)

        Dim desde As Date
        Dim hasta As Date
        If Me.rdoRangoFechas.value Then
            desde = Me.dtpDesde.value
            hasta = Me.dtpHasta.value
        Else
            Dim liq As LiquidacionSubdiarioVenta
            Set liq = liquidaciones.item(CStr(Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.ListIndex)))
            desde = liq.desde
            hasta = liq.hasta
        End If

        .cells(2, 1).value = "Periodo " & Format(desde, "dd/mm/yyyy") & " - " & Format(hasta, "dd/mm/yyyy")
        .Range("A3:Q3").Interior.Color = &HC0C0C0


        Dim Column As JSColumn
        Dim x As Integer

        For Each Column In Me.GridEX1.Columns
            x = x + 1
            .cells(3, x).value = Column.caption
        Next Column


        x = 1
        For Each item In liq.Detalles
            .cells(x + 3, 1).value = item.FEcha
            .cells(x + 3, 2).value = item.Comprobante
            .cells(x + 3, 3).value = item.RazonSocial
            .cells(x + 3, 4).value = item.Cuit
            .cells(x + 3, 5).value = item.CondicionIva
            .cells(x + 3, 6).value = item.NetoGravado


            .cells(x + 3, 7).value = item.NetosGravado(CStr(27))
            .cells(x + 3, 8).value = item.AlicuotasIva(CStr(27))

            .cells(x + 3, 9).value = item.NetosGravado(CStr(21))
            .cells(x + 3, 10).value = item.AlicuotasIva(CStr(21))

            .cells(x + 3, 11).value = item.NetosGravado(CStr(11))
            .cells(x + 3, 12).value = item.AlicuotasIva(CStr(11))

            .cells(x + 3, 13).value = item.Exento
            .cells(x + 3, 14).value = item.percepciones

            '.Cells(x + 3, 15).value = item.PercepcionesIVA
            .cells(x + 3, 16).value = item.ImpuestoInterno
            .cells(x + 3, 17).value = item.Total

            x = x + 1
        Next item

        a = "Q" & x + 2
        offset = x + 3
        b = "Q" & offset
        .Range("f1", b).NumberFormat = "0.00"
        .Range("a1", a).Borders.LineStyle = xlContinuous

        .Range("f" & x + 3, b).Interior.Color = &HC0C0C0
        .Range("f" & x + 3, b).Borders.LineStyle = xlContinuous
        .Range("f" & x + 3, b).Font.Bold = True


        .cells(offset, 5).value = "Totales"
        .Range(.cells(offset, 6), .cells(offset, 6)).Formula = "=SUM(F3:F" & x + 2 & ")"    'totales.Item(PosicionTotales.TotNetoGravado)
        .Range(.cells(offset, 7), .cells(offset, 7)).Formula = "=SUM(G3:G" & x + 2 & ")"    'totalesIva.Item(CStr(27))
        .Range(.cells(offset, 8), .cells(offset, 8)).Formula = "=SUM(H3:H" & x + 2 & ")"    'totalesIva.Item(CStr(21))
        .Range(.cells(offset, 9), .cells(offset, 9)).Formula = "=SUM(I3:I" & x + 2 & ")"    'totalesIva.Item(CStr(10.5))
        .Range(.cells(offset, 10), .cells(offset, 10)).Formula = "=SUM(J3:J" & x + 2 & ")"    'totales.Item(PosicionTotales.TotExento)
        .Range(.cells(offset, 11), .cells(offset, 11)).Formula = "=SUM(K3:K" & x + 2 & ")"    'totales.Item(PosicionTotales.TotPercepIB)
        .Range(.cells(offset, 12), .cells(offset, 12)).Formula = "=SUM(L3:L" & x + 2 & ")"    'totales.Item(PosicionTotales.TotPercepIVA)
        .Range(.cells(offset, 13), .cells(offset, 13)).Formula = "=SUM(M3:M" & x + 2 & ")"    'totales.Item(PosicionTotales.TotImpuestoInterno)
        .Range(.cells(offset, 14), .cells(offset, 14)).Formula = "=SUM(N3:N" & x + 2 & ")"    'totales.Item(PosicionTotales.TotTot)
        .Range(.cells(offset, 15), .cells(offset, 15)).Formula = "=SUM(O3:O" & x + 2 & ")"    'totales.Item(PosicionTotales.TotTot)
        .Range(.cells(offset, 16), .cells(offset, 16)).Formula = "=SUM(P3:P" & x + 2 & ")"    'totales.Item(PosicionTotales.TotTot)
        .Range(.cells(offset, 17), .cells(offset, 17)).Formula = "=SUM(Q3:Q" & x + 2 & ")"    'totales.Item(PosicionTotales.TotTot)

        Set CDLGMAIN = frmPrincipal.cd

        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter

        Dim Periodo As String
        Periodo = 1
        Periodo = Format(desde, "ddmmyyyy") & "-" & Format(hasta, "ddmmyyyy")

        Dim archi As String
        archi = "SUBDIARIO_COMPRAS_" & Periodo & ".xls"
        frmPrincipal.cd.CancelError = True
        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbInformation + vbOKOnly, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            ExportaSubDiarioVentas = False
        End If
        xlb.Saved = True
        xlb.Close

        xls.Quit
        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        ExportaSubDiarioVentas = True

    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        ExportaSubDiarioVentas = False
    Else
        'Resume
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        ExportaSubDiarioVentas = False
    End If
    xlb.Saved = True
    xlb.Close
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing

End Function
