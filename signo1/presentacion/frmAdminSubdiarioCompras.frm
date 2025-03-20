VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminSubdiarioCompras 
   Caption         =   "Subdiario IVA Compras"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14385
   Icon            =   "frmAdminSubdiarioCompras.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   14385
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
      ColumnsCount    =   6
      Column(1)       =   "frmAdminSubdiarioCompras.frx":000C
      Column(2)       =   "frmAdminSubdiarioCompras.frx":0164
      Column(3)       =   "frmAdminSubdiarioCompras.frx":025C
      Column(4)       =   "frmAdminSubdiarioCompras.frx":0358
      Column(5)       =   "frmAdminSubdiarioCompras.frx":0444
      Column(6)       =   "frmAdminSubdiarioCompras.frx":053C
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmAdminSubdiarioCompras.frx":0728
      FormatStyle(2)  =   "frmAdminSubdiarioCompras.frx":0860
      FormatStyle(3)  =   "frmAdminSubdiarioCompras.frx":0910
      FormatStyle(4)  =   "frmAdminSubdiarioCompras.frx":09C4
      FormatStyle(5)  =   "frmAdminSubdiarioCompras.frx":0A9C
      FormatStyle(6)  =   "frmAdminSubdiarioCompras.frx":0B54
      FormatStyle(7)  =   "frmAdminSubdiarioCompras.frx":0C34
      ImageCount      =   0
      PrinterProperties=   "frmAdminSubdiarioCompras.frx":0D14
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1680
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   18060
      _Version        =   786432
      _ExtentX        =   31856
      _ExtentY        =   2963
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ProgressBar progreso 
         Height          =   420
         Left            =   13560
         TabIndex        =   37
         Top             =   1160
         Visible         =   0   'False
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   741
         _StockProps     =   93
         Appearance      =   6
         BarColor        =   65280
      End
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
         Column(1)       =   "frmAdminSubdiarioCompras.frx":0EEC
         Column(2)       =   "frmAdminSubdiarioCompras.frx":1028
         Column(3)       =   "frmAdminSubdiarioCompras.frx":114C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminSubdiarioCompras.frx":1260
         FormatStyle(2)  =   "frmAdminSubdiarioCompras.frx":1398
         FormatStyle(3)  =   "frmAdminSubdiarioCompras.frx":1448
         FormatStyle(4)  =   "frmAdminSubdiarioCompras.frx":14FC
         FormatStyle(5)  =   "frmAdminSubdiarioCompras.frx":15D4
         FormatStyle(6)  =   "frmAdminSubdiarioCompras.frx":168C
         ImageCount      =   0
         PrinterProperties=   "frmAdminSubdiarioCompras.frx":176C
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
         Column(1)       =   "frmAdminSubdiarioCompras.frx":1944
         Column(2)       =   "frmAdminSubdiarioCompras.frx":1A84
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminSubdiarioCompras.frx":1BA0
         FormatStyle(2)  =   "frmAdminSubdiarioCompras.frx":1CD8
         FormatStyle(3)  =   "frmAdminSubdiarioCompras.frx":1D88
         FormatStyle(4)  =   "frmAdminSubdiarioCompras.frx":1E3C
         FormatStyle(5)  =   "frmAdminSubdiarioCompras.frx":1F14
         FormatStyle(6)  =   "frmAdminSubdiarioCompras.frx":1FCC
         ImageCount      =   0
         PrinterProperties=   "frmAdminSubdiarioCompras.frx":20AC
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
    Dim ali As Variant
    Dim per As clsPercepciones

    Set totales = New Dictionary
    Set totalesIva = New Collection

    For Each ali In alicuotas


        'If ali = 10.5 Then ali = 11
        totalesIva.Add 0, CStr(ali)
    Next ali

    Set percepciones = DAOPercepciones.GetAll

    Dim dtOP As DTOPercepcionImporte
    Set totalesper = New Collection
    For Each per In percepciones
        Set dtOP = New DTOPercepcionImporte
        dtOP.importe = 0
        Set dtOP.Percepcion = per
        totalesper.Add dtOP, CStr(per.Id)
    Next

    Dim pera As clsPercepcionesAplicadas
    For Each i In col
        sumNeto = sumNeto + i.NetoGravado
        sumIVA = sumIVA + (i.Iva)
        sumPercep = sumPercep + i.percepciones
        ' sumPercepIVA = sumPercepIVA + i.PercepcionesIVA

        sumExento = sumExento + i.AlicuotasIva(CStr(0))

        sumTot = sumTot + i.total
        sumImpuestoInterno = sumImpuestoInterno + i.ImpuestoInterno
        sumRedondeo = sumRedondeo + i.Redondeo

        'NB: error al mostrarl iquidaciones, revisar. Este fix es par aq no se cierre, pero no se si est? bien
        '18.1.2021
        If Not IsSomething(i.ListaPercepciones) Then Set i.ListaPercepciones = New Collection



        For Each pera In i.ListaPercepciones

            Set dtOP = totalesper(CStr(pera.Percepcion.Id))
            'tmpValue = funciones.RedondearDecimales(totalesper(CStr(pera.Percepcion.Id)).importe)
            totalesper.remove CStr(pera.Percepcion.Id)


            dtOP.importe = funciones.RedondearDecimales(dtOP.importe + pera.Monto)
            totalesper.Add dtOP, CStr(pera.Percepcion.Id)
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
    If Me.rdoRangoFechas.value Then
        ExportaSubDiarioComprasFechas
    Else
        ExportaSubDiarioComprasLiquidacion
    End If
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
        Set l.detalles = col
        If DAOSubdiarios.Guardar(l) Then
            SetearMaxDesde
            MsgBox "La liquidacion se guard? con ?xito", vbInformation + vbOKOnly
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

    '    Me.dtpDesde = "01/01/2022"
    '    Me.dtpHasta = "01/01/2022"


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

        Set col = Me.GridEX1.Columns.Add(per.Percepcion, jgexText, jgexEditNone, "PER_" & per.Id)
        col.TextAlignment = jgexAlignRight
        col.AggregateFunction = jgexSum
        col.GroupFormat = "0.00"
        col.TotalRowFormat = "0.00"
        col.Tag = per.Id

    Next per





    Set col = Me.GridEX1.Columns.Add("Imp. Int.", jgexText, jgexEditNone, "impuestointerno")
    col.TextAlignment = jgexAlignRight
    col.AggregateFunction = jgexSum
    col.GroupFormat = "0.00"
    col.TotalRowFormat = "0.00"

    Set col = Me.GridEX1.Columns.Add("Redondeo", jgexText, jgexEditNone, "redondeo")
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
            Set col = liqui.detalles
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
    Me.GridEX1.PrinterProperties.FooterString(jgexHFRight) = "P?gina " & PageNumber & " de " & nPages
End Sub

Private Sub GridEX1_DblClick()
    If col.count > 0 Then

        Set Factura = DAOFacturaProveedor.FindById(col(Me.GridEX1.rowIndex(Me.GridEX1.row)).FacturaId)

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
        If Me.GridEX1.rowIndex(Me.GridEX1.row) <> 0 Then
            Set item = col.item(Me.GridEX1.rowIndex(Me.GridEX1.row))
        End If
    End If
End Sub
Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
'On Error Resume Next
    If col.count > 0 Then
        Set item = col.item(rowIndex)

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

                Values(Me.GridEX1.Columns.item("NG_" & ali).Index) = funciones.FormatearDecimales(item.NetosGravado.item(CStr(ali)))
            End If
            Values(Me.GridEX1.Columns.item("IVA_" & ali).Index) = funciones.FormatearDecimales(funciones.RedondearDecimales(item.AlicuotasIva.item(CStr(ali))))

        Next

        '        Values(Me.GridEX1.Columns.Item("percepcionesiibb").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, Item.PercepcionesIB))
        '        Values(Me.GridEX1.Columns.Item("percepcionesiva").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, Item.PercepcionesIVA))
        '        Values(Me.GridEX1.Columns.Item("impuestointerno").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, Item.ImpuestoInterno))
        '        Values(Me.GridEX1.Columns.Item("total").index) = funciones.FormatearDecimales(IIf(Item.estado = Anulada, 0, (Item.Total)))


        ' PERCEPCIONES POR CADA COMPROBANTE

        Dim colper As New Collection
        Dim per As clsPercepciones

        Set colper = DAOPercepciones.GetAll

        For Each per In colper
            Dim PercepcionAcumulada As Double

            If BuscarEnColeccion(item.ListaPercepciones, CStr(per.Id)) Then

                PercepcionAcumulada = funciones.FormatearDecimales(item.ListaPercepciones(CStr(per.Id)).Monto)

                Values(Me.GridEX1.Columns.item("per_" & per.Id).Index) = PercepcionAcumulada

            Else
                Values(Me.GridEX1.Columns.item("per_" & per.Id).Index) = 0
            End If

        Next

        ' PERCEPCIONES POR CADA COMPROBANTE


        Values(Me.GridEX1.Columns.item("impuestointerno").Index) = funciones.FormatearDecimales(item.ImpuestoInterno)
        Values(Me.GridEX1.Columns.item("redondeo").Index) = funciones.FormatearDecimales(item.Redondeo)
        Values(Me.GridEX1.Columns.item("total").Index) = funciones.FormatearDecimales(item.total)
    End If

End Sub
Private Sub btnMostrar_Click()
    llenarLista
End Sub


Private Sub GridEX1_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 Then
        '   If Item.estado = Anulada Then
        If MsgBox("?Desea realmente actualizar los valores del item?", vbYesNo + vbQuestion) = vbYes Then
            Set item = col.item(rowIndex)
            item.NetoGravado = Values(6)
            item.Iva = Values(7)
            item.percepciones = Values(8)
            item.Exento = Values(9)
            item.total = Values(10)

            DAOSubdiarios.UpdateDetalle item
            Totalizar
        End If
        '    End If
    End If
End Sub



Private Sub GridEX2_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Values(2) = funciones.FormatearDecimales(totalesper(rowIndex).importe)    '/ (va / 100))
    If IsSomething(totalesper(rowIndex).Percepcion) Then
        Values(1) = totalesper(rowIndex).Percepcion.Percepcion
    Else
        Values(1) = funciones.FormatearDecimales(0)
    End If

End Sub

Private Sub gridTotalesIVA_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim va As Variant
    va = alicuotas(rowIndex)

    Values(1) = funciones.FormatearDecimales(totalesIva(rowIndex) / (va / 100))
    Values(2) = va & "%"

    Values(3) = totalesIva(rowIndex)

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
            header = "Liquidaci?n " & Me.cboLiquidaciones.Text
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
        Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.NewIndex) = liqui.Id
    Next liqui

End Sub

'#236

Public Function ExportaSubDiarioComprasFechas() As Boolean

    On Error GoTo errEXCEL

    'INICIA EL PROGRESSBAR Y LO MUESTRA
    Me.progreso.Visible = True

    'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
    progreso.min = 0
    progreso.max = col.count


    '    Dim xlb As New Excel.Workbook
    '    Dim xla As New Excel.Worksheet
    '    Dim xls As New Excel.Application

    'Dim xlApplication As New Excel.Application
    Dim xls As Object
    Set xls = CreateObject("Excel.Application")

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlb As Object
    Set xlb = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xla As Object
    Set xla = CreateObject("Excel.Application")

    Dim A As String
    Dim B As String
    Dim offset As Long
    Dim strMsg As String
    Dim CDLGMAIN As CommonDialog
    Dim sFilter As String


    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate


    With xla

        .Range("A1:AQ1").Merge
        .Range("A2:AQ2").Merge
        .Range("A1:AQ3").HorizontalAlignment = xlHAlignCenter
        .Range("A1:AQ2").Font.Bold = True
        .Range("A3:AQ2").Font.Bold = True


        .Cells(1, 1).value = "SIGNOPLAST S.A. Subdiario compras" & IIf(Me.rdoRangoFechas.value, " (NO LIQUIDADO)", vbNullString)

        Dim desde As Date
        Dim hasta As Date

        desde = Me.dtpDesde.value
        hasta = Me.dtpHasta.value

        .Cells(2, 1).value = "Periodo " & Format(desde, "dd/mm/yyyy") & " - " & Format(hasta, "dd/mm/yyyy")

        .Range("A3:AQ3").Interior.Color = &HC0C0C0

        Dim Column As JSColumn
        Dim x As Integer
        
        For Each Column In Me.GridEX1.Columns
            x = x + 1
            .Cells(3, x).value = Column.caption

        Next Column

        .Columns("f").HorizontalAlignment = xlHAlignRight
        .Columns("g").HorizontalAlignment = xlHAlignRight
        .Columns("h").HorizontalAlignment = xlHAlignRight
        .Columns("i").HorizontalAlignment = xlHAlignRight

        .Columns("a").HorizontalAlignment = xlHAlignCenter
        .Columns("b").HorizontalAlignment = xlHAlignCenter
        .Columns("d").HorizontalAlignment = xlHAlignCenter
        .Columns("e").HorizontalAlignment = xlHAlignCenter

        .Columns("j").HorizontalAlignment = xlHAlignRight

        .Columns("a").ColumnWidth = 10
        .Columns("b").ColumnWidth = 8
        .Columns("c").ColumnWidth = 35
        .Columns("d").ColumnWidth = 13
        .Columns("e").ColumnWidth = 15
        .Columns("f").ColumnWidth = 13
        .Columns("g").ColumnWidth = 13
        .Columns("h").ColumnWidth = 13
        .Columns("i").ColumnWidth = 13
        .Columns("j").ColumnWidth = 15
        .Columns("k").ColumnWidth = 15
        .Columns("l").ColumnWidth = 15
        .Columns("m").ColumnWidth = 15
        .Columns("n").ColumnWidth = 15
        .Columns("o").ColumnWidth = 15
        .Columns("p").ColumnWidth = 15
        .Columns("q").ColumnWidth = 15
        .Columns("r").ColumnWidth = 15
        .Columns("s").ColumnWidth = 15
        .Columns("t").ColumnWidth = 15
        .Columns("u").ColumnWidth = 15
        .Columns("v").ColumnWidth = 15
        .Columns("w").ColumnWidth = 15
        .Columns("x").ColumnWidth = 15
        .Columns("y").ColumnWidth = 15
        .Columns("z").ColumnWidth = 15
        .Columns("aa").ColumnWidth = 15
        .Columns("ab").ColumnWidth = 15
        .Columns("ac").ColumnWidth = 15
        .Columns("ad").ColumnWidth = 15
        .Columns("ae").ColumnWidth = 15
        .Columns("af").ColumnWidth = 15
        .Columns("ag").ColumnWidth = 15
        .Columns("ah").ColumnWidth = 15
        .Columns("ai").ColumnWidth = 15
        .Columns("aj").ColumnWidth = 15
        .Columns("ak").ColumnWidth = 15
        .Columns("al").ColumnWidth = 15
        .Columns("am").ColumnWidth = 15
        .Columns("an").ColumnWidth = 15
        .Columns("ao").ColumnWidth = 15

        Dim total As Double
        Dim totnetog As Double
        Dim totIV As Double
        Dim totperi As Double
        Dim totexen As Double
        total = 0
        totnetog = 0
        totIV = 0
        totperi = 0
        totexen = 0

        x = 1

        'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
        Dim d As Long
        d = 0

        For Each item In col

            '.Cells(x + 3, 1).value = item.FEcha

            .Cells(x + 3, 1).value = Format(item.FEcha, "mm/dd/yyyy")

            .Cells(x + 3, 2).value = item.Comprobante
            .Cells(x + 3, 3).value = item.RazonSocial
            .Cells(x + 3, 4).value = item.Cuit
            .Cells(x + 3, 5).value = item.CondicionIva
            .Cells(x + 3, 6).value = item.NetoGravado


            If item.NetosGravado.item(1) Then
                .Cells(x + 3, 7).value = item.NetosGravado.item(1)
            End If

            If item.NetosGravado.item(1) Then
                .Cells(x + 3, 8).value = FormatearDecimales(item.NetosGravado.item(1) * 27 / 100)
            End If


            If item.NetosGravado.item(2) Then
                .Cells(x + 3, 9).value = item.NetosGravado.item(2)
            End If

            If item.NetosGravado.item(2) Then
                .Cells(x + 3, 10).value = FormatearDecimales(item.NetosGravado.item(2) * 21 / 100)
            End If


            If item.NetosGravado.item(3) Then
                .Cells(x + 3, 11).value = item.NetosGravado.item(3)
            End If

            If item.NetosGravado.item(3) Then
                .Cells(x + 3, 12).value = FormatearDecimales(item.NetosGravado.item(3) * 10.5 / 100)
            End If


            If item.NetosGravado.item(4) Then
                .Cells(x + 3, 13).value = item.NetosGravado.item(4)
            End If



            If item.NetosGravado.item(4) Then
                .Cells(x + 3, 14).value = FormatearDecimales(item.NetosGravado.item(4) * 5 / 100)
            End If

            If item.NetosGravado.item(5) Then
                .Cells(x + 3, 15).value = item.NetosGravado.item(5)
            End If



            If item.ListaPercepciones.count <> 0 Then

                Dim i
                For i = 1 To item.ListaPercepciones.count Step 1

                    Select Case item.ListaPercepciones.item(i).Percepcion.Percepcion
                    Case "IIBB CABA"
                        .Cells(x + 3, 16).value = item.ListaPercepciones.item(i).Monto
                    Case "IVA"
                        .Cells(x + 3, 17).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SANTA FE"
                        .Cells(x + 3, 18).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SALTA"
                        .Cells(x + 3, 19).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB BUENOS AIRES"
                        .Cells(x + 3, 20).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB MISIONES"
                        .Cells(x + 3, 21).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB TUCUMAN"
                        .Cells(x + 3, 22).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SAN LUIS"
                        .Cells(x + 3, 23).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CORRIENTES"
                        .Cells(x + 3, 24).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB RIO NEGRO"
                        .Cells(x + 3, 25).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB ENTRE RIOS"
                        .Cells(x + 3, 26).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CORDOBA"
                        .Cells(x + 3, 27).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CATAMARCA"
                        .Cells(x + 3, 28).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB NEUQUEN"
                        .Cells(x + 3, 29).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB LA PAMPA"
                        .Cells(x + 3, 30).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB MENDOZA"
                        .Cells(x + 3, 31).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SAN JUAN"
                        .Cells(x + 3, 32).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SANTA CRUZ"
                        .Cells(x + 3, 33).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CHUBUT"
                        .Cells(x + 3, 34).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB LA RIOJA"
                        .Cells(x + 3, 35).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SANTIAGO DEL ESTERO"
                        .Cells(x + 3, 36).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CHACO"
                        .Cells(x + 3, 37).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB FORMOSA"
                        .Cells(x + 3, 38).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB JUJUY"
                        .Cells(x + 3, 39).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB TIERRA DEL FUEGO"
                        .Cells(x + 3, 40).value = item.ListaPercepciones.item(i).Monto


                    End Select

                Next i

            End If

            .Cells(x + 3, 41).value = item.ImpuestoInterno
            .Cells(x + 3, 42).value = item.Redondeo
            .Cells(x + 3, 43).value = item.total

            x = x + 1

            'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
            d = d + 1
            progreso.value = d


        Next item


        A = "aq" & x + 2
        offset = x + 3
        B = "aq" & offset
        .Range("f1", B).NumberFormat = "0.00"
        .Range("a1", A).Borders.LineStyle = xlContinuous

        .Range("f" & x + 3, B).Interior.Color = &HC0C0C0
        .Range("f" & x + 3, B).Borders.LineStyle = xlContinuous
        .Range("f" & x + 3, B).Font.Bold = True


        .Cells(offset, 5).value = "Totales"
        .Cells(offset, 6).value = xls.WorksheetFunction.SUM(.Range("f4", "f" & x + 3))
        .Cells(offset, 7).value = xls.WorksheetFunction.SUM(.Range("g4", "g" & x + 3))
        .Cells(offset, 8).value = xls.WorksheetFunction.SUM(.Range("h4", "h" & x + 3))
        .Cells(offset, 9).value = xls.WorksheetFunction.SUM(.Range("i4", "i" & x + 3))
        .Cells(offset, 10).value = xls.WorksheetFunction.SUM(.Range("j4", "j" & x + 3))
        .Cells(offset, 11).value = xls.WorksheetFunction.SUM(.Range("k4", "k" & x + 3))
        .Cells(offset, 12).value = xls.WorksheetFunction.SUM(.Range("l4", "l" & x + 3))
        .Cells(offset, 13).value = xls.WorksheetFunction.SUM(.Range("m4", "m" & x + 3))
        .Cells(offset, 14).value = xls.WorksheetFunction.SUM(.Range("n4", "n" & x + 3))
        .Cells(offset, 15).value = xls.WorksheetFunction.SUM(.Range("o4", "o" & x + 3))
        .Cells(offset, 16).value = xls.WorksheetFunction.SUM(.Range("p4", "p" & x + 3))
        .Cells(offset, 17).value = xls.WorksheetFunction.SUM(.Range("q4", "q" & x + 3))
        .Cells(offset, 18).value = xls.WorksheetFunction.SUM(.Range("r4", "r" & x + 3))
        .Cells(offset, 19).value = xls.WorksheetFunction.SUM(.Range("s4", "s" & x + 3))
        .Cells(offset, 20).value = xls.WorksheetFunction.SUM(.Range("t4", "t" & x + 3))
        .Cells(offset, 21).value = xls.WorksheetFunction.SUM(.Range("u4", "u" & x + 3))
        .Cells(offset, 22).value = xls.WorksheetFunction.SUM(.Range("v4", "v" & x + 3))
        .Cells(offset, 23).value = xls.WorksheetFunction.SUM(.Range("w4", "w" & x + 3))
        .Cells(offset, 24).value = xls.WorksheetFunction.SUM(.Range("x4", "x" & x + 3))
        .Cells(offset, 25).value = xls.WorksheetFunction.SUM(.Range("y4", "y" & x + 3))
        .Cells(offset, 26).value = xls.WorksheetFunction.SUM(.Range("z4", "z" & x + 3))
        .Cells(offset, 27).value = xls.WorksheetFunction.SUM(.Range("aa4", "aa" & x + 3))
        .Cells(offset, 28).value = xls.WorksheetFunction.SUM(.Range("ab4", "ab" & x + 3))
        .Cells(offset, 29).value = xls.WorksheetFunction.SUM(.Range("ac4", "ac" & x + 3))
        .Cells(offset, 30).value = xls.WorksheetFunction.SUM(.Range("ad4", "ad" & x + 3))
        .Cells(offset, 31).value = xls.WorksheetFunction.SUM(.Range("ae4", "ae" & x + 3))
        .Cells(offset, 32).value = xls.WorksheetFunction.SUM(.Range("af4", "af" & x + 3))
        .Cells(offset, 33).value = xls.WorksheetFunction.SUM(.Range("ag4", "ag" & x + 3))
        .Cells(offset, 34).value = xls.WorksheetFunction.SUM(.Range("ah4", "ah" & x + 3))
        .Cells(offset, 35).value = xls.WorksheetFunction.SUM(.Range("ai4", "ai" & x + 3))
        .Cells(offset, 36).value = xls.WorksheetFunction.SUM(.Range("aj4", "aj" & x + 3))
        .Cells(offset, 37).value = xls.WorksheetFunction.SUM(.Range("ak4", "ak" & x + 3))
        .Cells(offset, 38).value = xls.WorksheetFunction.SUM(.Range("al4", "al" & x + 3))
        .Cells(offset, 39).value = xls.WorksheetFunction.SUM(.Range("am4", "am" & x + 3))
        .Cells(offset, 40).value = xls.WorksheetFunction.SUM(.Range("an4", "an" & x + 3))
        .Cells(offset, 41).value = xls.WorksheetFunction.SUM(.Range("ao4", "ao" & x + 3))
        .Cells(offset, 42).value = xls.WorksheetFunction.SUM(.Range("ap4", "ap" & x + 3))
        .Cells(offset, 43).value = xls.WorksheetFunction.SUM(.Range("aq4", "aq" & x + 3))

        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        Set CDLGMAIN = frmPrincipal.CD



        '    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter

        Dim Periodo As String
        Periodo = 1
        Periodo = Format(desde, "ddmmyyyy") & "-" & Format(hasta, "ddmmyyyy")

        Dim archi As String
        archi = "SUBDIARIO_COMPRAS_" & Periodo & ".xls"
        frmPrincipal.CD.CancelError = True
        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbInformation + vbOKOnly, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            ExportaSubDiarioComprasFechas = False
        End If
        xlb.Saved = True

        'xlb.Close

        xls.Visible = True    'NO MUESTRO LA HOJA XLS

        'xls.Quit

        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        '    End If
        ExportaSubDiarioComprasFechas = True

        'REINICIA EL PROGRESSBAR Y LO OCULTA
        progreso.value = 0
        Me.progreso.Visible = False


    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        ExportaSubDiarioComprasFechas = False
    Else
        ' Resume
        MsgBox "Se produjo un error o se canceló la exportación del archivo." & vbNewLine & "No se graban los cambios", vbCritical, "Error"
        ExportaSubDiarioComprasFechas = False
        Me.progreso.Visible = False

    End If
    xlb.Saved = True
    xlb.Close
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing

End Function

Public Function ExportaSubDiarioComprasLiquidacion() As Boolean

    On Error GoTo errEXCEL

    'INICIA EL PROGRESSBAR Y LO MUESTRA
    Me.progreso.Visible = True


    'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
    progreso.min = 0
    progreso.max = col.count

    'INICIA EL PROGRESSBAR Y LO MUESTRA
    Me.progreso.Visible = True


    'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
    progreso.min = 0
    progreso.max = col.count

    '    Dim xlb As New Excel.Workbook
    '    Dim xla As New Excel.Worksheet
    '    Dim xls As New Excel.Application

    'Dim xlApplication As New Excel.Application
    Dim xls As Object
    Set xls = CreateObject("Excel.Application")

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlb As Object
    Set xlb = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xla As Object
    Set xla = CreateObject("Excel.Application")

    Dim A As String
    Dim B As String
    Dim offset As Long
    Dim strMsg As String
    Dim CDLGMAIN As CommonDialog
    Dim sFilter As String


    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate


    With xla

        .Range("A1:a1").Merge
        .Range("A2:an2").Merge
        .Range("A1:an3").HorizontalAlignment = xlHAlignCenter
        .Range("A1:an2").Font.Bold = True
        .Range("A3:an2").Font.Bold = True


        .Cells(1, 1).value = "SIGNOPLAST S.A. Subdiario compras" & " (LIQUIDADO)"

        .Cells(2, 1).value = "Periodo " & Format(liqui.desde, "dd/mm/yyyy") & " - " & Format(liqui.hasta, "dd/mm/yyyy")


        .Range("A3:an3").Interior.Color = &HC0C0C0

        Dim Column As JSColumn
        Dim x As Integer

        For Each Column In Me.GridEX1.Columns
            x = x + 1
            .Cells(3, x).value = Column.caption

        Next Column

        .Columns("f").HorizontalAlignment = xlHAlignRight
        .Columns("g").HorizontalAlignment = xlHAlignRight
        .Columns("h").HorizontalAlignment = xlHAlignRight
        .Columns("i").HorizontalAlignment = xlHAlignRight

        .Columns("a").HorizontalAlignment = xlHAlignCenter
        .Columns("b").HorizontalAlignment = xlHAlignCenter
        .Columns("d").HorizontalAlignment = xlHAlignCenter
        .Columns("e").HorizontalAlignment = xlHAlignCenter

        .Columns("j").HorizontalAlignment = xlHAlignRight

        .Columns("a").ColumnWidth = 10
        .Columns("b").ColumnWidth = 8
        .Columns("c").ColumnWidth = 35
        .Columns("d").ColumnWidth = 13
        .Columns("e").ColumnWidth = 15
        .Columns("f").ColumnWidth = 13
        .Columns("g").ColumnWidth = 13
        .Columns("h").ColumnWidth = 13
        .Columns("i").ColumnWidth = 13
        .Columns("j").ColumnWidth = 15
        .Columns("k").ColumnWidth = 15
        .Columns("l").ColumnWidth = 15
        .Columns("m").ColumnWidth = 15
        .Columns("n").ColumnWidth = 15
        .Columns("o").ColumnWidth = 15
        .Columns("p").ColumnWidth = 15
        .Columns("q").ColumnWidth = 15
        .Columns("r").ColumnWidth = 15
        .Columns("s").ColumnWidth = 15
        .Columns("t").ColumnWidth = 15
        .Columns("u").ColumnWidth = 15
        .Columns("v").ColumnWidth = 15
        .Columns("w").ColumnWidth = 15
        .Columns("x").ColumnWidth = 15
        .Columns("y").ColumnWidth = 15
        .Columns("z").ColumnWidth = 15
        .Columns("aa").ColumnWidth = 15
        .Columns("ab").ColumnWidth = 15
        .Columns("ac").ColumnWidth = 15
        .Columns("ad").ColumnWidth = 15
        .Columns("ae").ColumnWidth = 15
        .Columns("af").ColumnWidth = 15
        .Columns("ag").ColumnWidth = 15
        .Columns("ah").ColumnWidth = 15
        .Columns("ai").ColumnWidth = 15
        .Columns("aj").ColumnWidth = 15
        .Columns("ak").ColumnWidth = 15
        .Columns("al").ColumnWidth = 15
        .Columns("am").ColumnWidth = 15
        .Columns("an").ColumnWidth = 15


        Dim total As Double
        Dim totnetog As Double
        Dim totIV As Double
        Dim totperi As Double
        Dim totexen As Double
        total = 0
        totnetog = 0
        totIV = 0
        totperi = 0
        totexen = 0

        x = 1

        'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
        Dim d As Long
        d = 0

        For Each item In col

            '.Cells(x + 3, 1).value = item.FEcha

            .Cells(x + 3, 1).value = Format(item.FEcha, "mm/dd/yyyy")

            .Cells(x + 3, 2).value = item.Comprobante
            .Cells(x + 3, 3).value = item.RazonSocial
            .Cells(x + 3, 4).value = item.Cuit
            .Cells(x + 3, 5).value = item.CondicionIva
            .Cells(x + 3, 6).value = item.NetoGravado

            'IVA

            If item.NetosGravado.item(4) Then
                .Cells(x + 3, 7).value = item.NetosGravado.item(4)
            End If

            If item.NetosGravado.item(4) Then
                .Cells(x + 3, 8).value = FormatearDecimales(item.NetosGravado.item(4) * 27 / 100)
                ''''''''
            End If

            If item.NetosGravado.item(3) Then
                .Cells(x + 3, 9).value = item.NetosGravado.item(3)
            End If

            If item.NetosGravado.item(3) Then
                .Cells(x + 3, 10).value = FormatearDecimales(item.NetosGravado.item(3) * 21 / 100)
                ''''''''
            End If

            If item.NetosGravado.item(2) Then
                .Cells(x + 3, 11).value = item.NetosGravado.item(2)
            End If

            If item.NetosGravado.item(2) Then
                .Cells(x + 3, 12).value = FormatearDecimales(item.NetosGravado.item(2) * 10.5 / 100)
                ''''''''
            End If

            If item.NetosGravado.item(1) Then
                .Cells(x + 3, 13).value = item.NetosGravado.item(1)
            End If

            'PERCEPCIONES

            If item.ListaPercepciones.count <> 0 Then

                Dim i
                For i = 1 To item.ListaPercepciones.count Step 1

                    Select Case item.ListaPercepciones.item(i).Percepcion.Percepcion
                    Case "IIBB CABA"
                        .Cells(x + 3, 14).value = item.ListaPercepciones.item(i).Monto
                    Case "IVA"
                        .Cells(x + 3, 15).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SANTA FE"
                        .Cells(x + 3, 16).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SALTA"
                        .Cells(x + 3, 17).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB BUENOS AIRES"
                        .Cells(x + 3, 18).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB MISIONES"
                        .Cells(x + 3, 19).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB TUCUMAN"
                        .Cells(x + 3, 20).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SAN LUIS"
                        .Cells(x + 3, 21).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CORRIENTES"
                        .Cells(x + 3, 22).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB RIO NEGRO"
                        .Cells(x + 3, 23).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB ENTRE RIOS"
                        .Cells(x + 3, 24).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CORDOBA"
                        .Cells(x + 3, 25).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CATAMARCA"
                        .Cells(x + 3, 26).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB NEUQUEN"
                        .Cells(x + 3, 27).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB LA PAMPA"
                        .Cells(x + 3, 28).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB MENDOZA"
                        .Cells(x + 3, 29).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SAN JUAN"
                        .Cells(x + 3, 30).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SANTA CRUZ"
                        .Cells(x + 3, 31).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CHUBUT"
                        .Cells(x + 3, 32).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB LA RIOJA"
                        .Cells(x + 3, 33).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB SANTIAGO DEL ESTERO"
                        .Cells(x + 3, 34).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB CHACO"
                        .Cells(x + 3, 35).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB FORMOSA"
                        .Cells(x + 3, 36).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB JUJUY"
                        .Cells(x + 3, 37).value = item.ListaPercepciones.item(i).Monto
                    Case "IIBB TIERRA DEL FUEGO"
                        .Cells(x + 3, 38).value = item.ListaPercepciones.item(i).Monto


                    End Select

                Next i

            End If

            .Cells(x + 3, 39).value = item.ImpuestoInterno
            .Cells(x + 3, 40).value = item.total

            x = x + 1

            'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
            d = d + 1
            progreso.value = d


        Next item


        A = "an" & x + 2
        offset = x + 3
        B = "an" & offset
        .Range("f1", B).NumberFormat = "0.00"
        .Range("a1", A).Borders.LineStyle = xlContinuous

        .Range("f" & x + 3, B).Interior.Color = &HC0C0C0
        .Range("f" & x + 3, B).Borders.LineStyle = xlContinuous
        .Range("f" & x + 3, B).Font.Bold = True


        .Cells(offset, 5).value = "Totales"
        .Cells(offset, 6).value = xls.WorksheetFunction.SUM(.Range("f4", "f" & x + 3))
        .Cells(offset, 7).value = xls.WorksheetFunction.SUM(.Range("g4", "g" & x + 3))
        .Cells(offset, 8).value = xls.WorksheetFunction.SUM(.Range("h4", "h" & x + 3))
        .Cells(offset, 9).value = xls.WorksheetFunction.SUM(.Range("i4", "i" & x + 3))
        .Cells(offset, 10).value = xls.WorksheetFunction.SUM(.Range("j4", "j" & x + 3))
        .Cells(offset, 11).value = xls.WorksheetFunction.SUM(.Range("k4", "k" & x + 3))
        .Cells(offset, 12).value = xls.WorksheetFunction.SUM(.Range("l4", "l" & x + 3))
        .Cells(offset, 13).value = xls.WorksheetFunction.SUM(.Range("m4", "m" & x + 3))
        .Cells(offset, 14).value = xls.WorksheetFunction.SUM(.Range("n4", "n" & x + 3))
        .Cells(offset, 15).value = xls.WorksheetFunction.SUM(.Range("o4", "o" & x + 3))
        .Cells(offset, 16).value = xls.WorksheetFunction.SUM(.Range("p4", "p" & x + 3))
        .Cells(offset, 17).value = xls.WorksheetFunction.SUM(.Range("q4", "q" & x + 3))
        .Cells(offset, 18).value = xls.WorksheetFunction.SUM(.Range("r4", "r" & x + 3))
        .Cells(offset, 19).value = xls.WorksheetFunction.SUM(.Range("s4", "s" & x + 3))
        .Cells(offset, 20).value = xls.WorksheetFunction.SUM(.Range("t4", "t" & x + 3))
        .Cells(offset, 21).value = xls.WorksheetFunction.SUM(.Range("u4", "u" & x + 3))
        .Cells(offset, 22).value = xls.WorksheetFunction.SUM(.Range("v4", "v" & x + 3))
        .Cells(offset, 23).value = xls.WorksheetFunction.SUM(.Range("w4", "w" & x + 3))
        .Cells(offset, 24).value = xls.WorksheetFunction.SUM(.Range("x4", "x" & x + 3))
        .Cells(offset, 25).value = xls.WorksheetFunction.SUM(.Range("y4", "y" & x + 3))
        .Cells(offset, 26).value = xls.WorksheetFunction.SUM(.Range("z4", "z" & x + 3))
        .Cells(offset, 27).value = xls.WorksheetFunction.SUM(.Range("aa4", "aa" & x + 3))
        .Cells(offset, 28).value = xls.WorksheetFunction.SUM(.Range("ab4", "ab" & x + 3))
        .Cells(offset, 29).value = xls.WorksheetFunction.SUM(.Range("ac4", "ac" & x + 3))
        .Cells(offset, 30).value = xls.WorksheetFunction.SUM(.Range("ad4", "ad" & x + 3))
        .Cells(offset, 31).value = xls.WorksheetFunction.SUM(.Range("ae4", "ae" & x + 3))
        .Cells(offset, 32).value = xls.WorksheetFunction.SUM(.Range("af4", "af" & x + 3))
        .Cells(offset, 33).value = xls.WorksheetFunction.SUM(.Range("ag4", "ag" & x + 3))
        .Cells(offset, 34).value = xls.WorksheetFunction.SUM(.Range("ah4", "ah" & x + 3))
        .Cells(offset, 35).value = xls.WorksheetFunction.SUM(.Range("ai4", "ai" & x + 3))
        .Cells(offset, 36).value = xls.WorksheetFunction.SUM(.Range("aj4", "aj" & x + 3))
        .Cells(offset, 37).value = xls.WorksheetFunction.SUM(.Range("ak4", "ak" & x + 3))
        .Cells(offset, 38).value = xls.WorksheetFunction.SUM(.Range("al4", "al" & x + 3))
        .Cells(offset, 39).value = xls.WorksheetFunction.SUM(.Range("am4", "am" & x + 3))
        .Cells(offset, 40).value = xls.WorksheetFunction.SUM(.Range("an4", "an" & x + 3))



        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        Set CDLGMAIN = frmPrincipal.CD



        '    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter

        Dim Periodo As String
        Periodo = 1
        Periodo = Format(liqui.desde, "ddmmyyyy") & "-" & Format(liqui.hasta, "ddmmyyyy")

        Dim archi As String
        archi = "SUBDIARIO_COMPRAS_" & Periodo & ".xls"
        frmPrincipal.CD.CancelError = True
        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbInformation + vbOKOnly, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            ExportaSubDiarioComprasLiquidacion = False
        End If
        xlb.Saved = True

        'xlb.Close

        xls.Visible = True    'NO MUESTRO LA HOJA XLS

        'xls.Quit

        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        'REINICIA EL PROGRESSBAR Y LO OCULTA
        progreso.value = 0
        Me.progreso.Visible = False

        '    End If
        ExportaSubDiarioComprasLiquidacion = True



    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        ExportaSubDiarioComprasLiquidacion = False
    Else
        ' Resume
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        ExportaSubDiarioComprasLiquidacion = False
    End If
    xlb.Saved = True
    xlb.Close
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing

End Function




