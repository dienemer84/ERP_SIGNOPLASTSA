VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminResumenesFacturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumenes de Facturacion"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmAdminResumenesFacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   6735
   Begin XtremeSuiteControls.CheckBox chkGrafico 
      Height          =   255
      Left            =   2115
      TabIndex        =   21
      Top             =   6615
      Width           =   1035
      _Version        =   786432
      _ExtentX        =   1826
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ver Gráfico"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   6630
      Width           =   3135
      _Version        =   786432
      _ExtentX        =   5530
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Resúmen"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agrupado Por Fecha"
         Appearance      =   6
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadioButton2 
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agrupado Por Mes"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton RadioButton3 
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   960
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agrupado Por Año"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   1320
         TabIndex        =   19
         Top             =   1800
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Generar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   315
         Left            =   390
         TabIndex        =   20
         Top             =   1260
         Width           =   2325
         _Version        =   786432
         _ExtentX        =   4101
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Generar Gráfico"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin MSChart20Lib.MSChart grafica 
      Height          =   8415
      Left            =   6720
      OleObjectBlob   =   "frmAdminResumenesFacturacion.frx":000C
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9015
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3720
      Left            =   135
      TabIndex        =   1
      Top             =   2790
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6562
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmAdminResumenesFacturacion.frx":1DB8
      Column(2)       =   "frmAdminResumenesFacturacion.frx":1F3C
      Column(3)       =   "frmAdminResumenesFacturacion.frx":20B8
      Column(4)       =   "frmAdminResumenesFacturacion.frx":221C
      Column(5)       =   "frmAdminResumenesFacturacion.frx":238C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminResumenesFacturacion.frx":24CC
      FormatStyle(2)  =   "frmAdminResumenesFacturacion.frx":2604
      FormatStyle(3)  =   "frmAdminResumenesFacturacion.frx":26B4
      FormatStyle(4)  =   "frmAdminResumenesFacturacion.frx":2768
      FormatStyle(5)  =   "frmAdminResumenesFacturacion.frx":2840
      FormatStyle(6)  =   "frmAdminResumenesFacturacion.frx":28F8
      ImageCount      =   0
      PrinterProperties=   "frmAdminResumenesFacturacion.frx":29D8
   End
   Begin XtremeSuiteControls.GroupBox grpTotales 
      Height          =   1680
      Left            =   3480
      TabIndex        =   2
      Top             =   6630
      Width           =   3090
      _Version        =   786432
      _ExtentX        =   5450
      _ExtentY        =   2963
      _StockProps     =   79
      Caption         =   "Totales"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label lblTotalTotal 
         Height          =   195
         Left            =   825
         TabIndex        =   12
         Top             =   1380
         Width           =   2115
         _Version        =   786432
         _ExtentX        =   3731
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         BackColor       =   -2147483633
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblExentoTotal 
         Height          =   195
         Left            =   825
         TabIndex        =   11
         Top             =   1020
         Width           =   2115
         _Version        =   786432
         _ExtentX        =   3731
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         BackColor       =   -2147483633
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblPercepcionesTotal 
         Height          =   195
         Left            =   1545
         TabIndex        =   10
         Top             =   750
         Width           =   1395
         _Version        =   786432
         _ExtentX        =   2461
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         BackColor       =   -2147483633
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblIVATotal 
         Height          =   195
         Left            =   705
         TabIndex        =   9
         Top             =   495
         Width           =   2235
         _Version        =   786432
         _ExtentX        =   3942
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         BackColor       =   -2147483633
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblNetoGravadoTotal 
         Height          =   195
         Left            =   1305
         TabIndex        =   8
         Top             =   225
         Width           =   1635
         _Version        =   786432
         _ExtentX        =   2884
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         BackColor       =   -2147483633
         Alignment       =   1
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFDBBF&
         DrawMode        =   9  'Not Mask Pen
         X1              =   2955
         X2              =   135
         Y1              =   1305
         Y2              =   1305
      End
      Begin XtremeSuiteControls.Label lblTotal 
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1380
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total:"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblExento 
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1020
         Width           =   570
         _Version        =   786432
         _ExtentX        =   1005
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Exento:"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPercepcionesIIBB 
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   750
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Percepciones IIBB:"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIVA 
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   495
         Width           =   315
         _Version        =   786432
         _ExtentX        =   556
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "IVA:"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNetoGravado 
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   225
         Width           =   1065
         _Version        =   786432
         _ExtentX        =   1879
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Neto Gravado:"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2565
      Left            =   120
      TabIndex        =   13
      Top             =   60
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   4524
      _StockProps     =   79
      Caption         =   "Filtro"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   5115
         TabIndex        =   14
         Top             =   2100
         Width           =   1185
         _Version        =   786432
         _ExtentX        =   2090
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Generar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   285
         Left            =   4620
         TabIndex        =   23
         Top             =   240
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1050
         Left            =   255
         TabIndex        =   24
         Top             =   1410
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Fecha Entrega"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   825
            TabIndex        =   27
            Top             =   225
            Width           =   3645
            _Version        =   786432
            _ExtentX        =   6429
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   825
            TabIndex        =   25
            Top             =   615
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            Format          =   1
            CurrentDate     =   40274.7427083333
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   3000
            TabIndex        =   26
            Top             =   600
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            Format          =   1
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   195
            Left            =   255
            TabIndex        =   30
            Top             =   270
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   270
            TabIndex        =   29
            Top             =   645
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2430
            TabIndex        =   28
            Top             =   660
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton cmdSinContrato 
         Height          =   285
         Left            =   4620
         TabIndex        =   31
         Top             =   615
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboContratos 
         Height          =   315
         Left            =   810
         TabIndex        =   32
         Top             =   600
         Width           =   3765
         _Version        =   786432
         _ExtentX        =   6641
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   4605
         TabIndex        =   34
         Top             =   1020
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboProvincias 
         Height          =   315
         Left            =   810
         TabIndex        =   35
         Top             =   1005
         Width           =   3765
         _Version        =   786432
         _ExtentX        =   6641
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   810
         TabIndex        =   37
         Top             =   225
         Width           =   3765
         _Version        =   786432
         _ExtentX        =   6641
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "cboClientes"
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   195
         Left            =   30
         TabIndex        =   36
         Top             =   1065
         Width           =   660
         _Version        =   786432
         _ExtentX        =   1164
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Provincia"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   660
         Width           =   450
         _Version        =   786432
         _ExtentX        =   794
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Marco"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   285
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cliente"
         BackColor       =   -2147483633
         AutoSize        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAdminResumenesFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sv As SubdiarioVentasDetalle
Dim col As New Collection
Dim newcol As New Collection
Private Const SIN_GRAFICO As Long = 6825
Private Const CON_GRAFICO As Long = 15945
Dim idCliente As Long
Dim idContrato As Long
Dim idprovincia As Long
Dim otsMarco As Collection
Dim OTMarco As OrdenTrabajo



Private Sub cboClientes_Click()
    LlenarComboContratos
End Sub

Private Sub cboContratos_Click()
    If Me.cboContratos.ListIndex <> -1 Then
        idContrato = Me.cboContratos.ItemData(Me.cboContratos.ListIndex)
        Set OTMarco = DAOOrdenTrabajo.FindById(idContrato)
        If IsSomething(OTMarco) Then
            Me.dtpDesde.value = OTMarco.FechaInicioMarco
            Me.dtpHasta.value = OTMarco.FechaFinMarco

        End If
    End If
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub

Private Sub chkGrafico_Click()
    cambiar_tamaño
End Sub

Private Sub cambiar_tamaño()
    If Me.chkGrafico.value = xtpChecked Then
        Me.Width = CON_GRAFICO
    Else
        Me.Width = SIN_GRAFICO
    End If
End Sub

Private Sub CMDsINCliente_Click()
    Me.cboClientes.ListIndex = -1
    Me.cboContratos.Clear
End Sub

Private Sub cmdSinContrato_Click()
    Me.cboContratos.ListIndex = -1
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, False
    Me.GridEX1.ItemCount = 0
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1

    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_MonthCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i

    DAOProvincias.LlenarCombo Me.cboProvincias, 1
    Me.cboProvincias.ListIndex = -1
End Sub
Private Sub llenarLista()
    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = newcol.count
    Totalizar
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set sv = newcol.item(RowIndex)
    Values(1) = sv.FEcha
    Values(2) = Replace(FormatCurrency(funciones.FormatearDecimales(sv.NetoGravado)), "$", "")
    Values(3) = Replace(FormatCurrency(funciones.FormatearDecimales(sv.Iva)), "$", "")
    Values(4) = Replace(FormatCurrency(funciones.FormatearDecimales(sv.percepciones)), "$", "")
    Values(5) = Replace(FormatCurrency(funciones.FormatearDecimales(sv.Total)), "$", "")
    Exit Sub
err1:
End Sub

Private Sub LlenarComboContratos()
    If Me.cboClientes.ListIndex <> -1 Then
        idCliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
        Set otsMarco = DAOOrdenTrabajo.otsMarco(idCliente)
        If IsSomething(otsMarco) Then
            Dim ot1 As OrdenTrabajo
            Me.cboContratos.Clear
            For Each ot1 In otsMarco
                Me.cboContratos.AddItem ot1.IdFormateado & " - " & ot1.descripcion & " (" & ot1.FechaInicioMarco & " - " & ot1.FechaFinMarco & ")"
                Me.cboContratos.ItemData(Me.cboContratos.NewIndex) = ot1.Id
            Next ot1
            Me.cboContratos.ListIndex = -1
        End If
    End If
End Sub

Private Sub PushButton1_Click()
    Dim x As ListItem

    Dim idCliente As Long
    If Me.cboClientes.ListIndex <> -1 Then
        idCliente = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    Else
        idCliente = -1
    End If

    If Me.cboContratos.ListIndex <> -1 Then
        idContrato = Me.cboContratos.ItemData(Me.cboContratos.ListIndex)
    Else
        idContrato = -1
    End If



    If Me.cboProvincias.ListIndex <> -1 Then
        idprovincia = Me.cboProvincias.ItemData(Me.cboProvincias.ListIndex)
    Else
        idprovincia = -1
    End If

    Set col = DAOSubdiarios.SubDiarioVentas(Me.dtpDesde.value, Me.dtpHasta.value, "AdminFacturas.FechaEmision  asc", idCliente, idContrato, idprovincia)
    PushButton3_Click
    llenarLista

End Sub

Private Sub Totalizar()
    Dim sv As SubdiarioVentasDetalle

    Dim tot_iva As Double
    Dim tot_exento As Double
    Dim tot_neto As Double
    Dim tot_ib As Double
    Dim tot As Double

    For Each sv In newcol
        tot_iva = tot_iva + sv.Iva
        'If sv.Exento > 0 Then 'if borrado 24-02-15 ya que no acumulaba neto gravado pra el resumen
        tot_exento = tot_exento + sv.Exento
        '    Else
        tot_neto = tot_neto + sv.NetoGravado
        '   End If
        tot_ib = tot_ib + sv.percepciones
    Next sv

    'Values(2) = Replace(FormatCurrency(funciones.FormatearDecimales(sv.NetoGravado)), "$", "")

    Me.lblExentoTotal.caption = FormatCurrency(funciones.FormatearDecimales(tot_exento))
    Me.lblIVATotal.caption = FormatCurrency(funciones.FormatearDecimales(tot_iva))
    Me.lblNetoGravadoTotal.caption = FormatCurrency(funciones.FormatearDecimales(tot_neto))
    Me.lblPercepcionesTotal.caption = FormatCurrency(funciones.FormatearDecimales(tot_ib))


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Antes del 11.09.20 esto estaba asi:

    'Me.lblTotalTotal.caption = FormatearDecimales(tot_neto + tot_ib + tot_iva)   'tot_exento +  borrado 24-02-15

    'Despues del 11.09.20 queda asi:
    'Se reincopora a la sumatoria del total total el valor de total exento indicado por pedido de Karin el 27.07.20
    Me.lblTotalTotal.caption = FormatCurrency(funciones.FormatearDecimales(tot_neto + tot_ib + tot_iva + tot_exento))

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub
Private Function AgruparColeccion(ByVal col As Collection, groupmethod As FcGroupMethod) As Collection
    Dim newcol2 As New Collection
    Dim sv2 As SubdiarioVentasDetalle
    If groupmethod = GroupByDate Then
        For Each sv In col

            If Not BuscarEnColeccion(newcol2, CStr(sv.FEcha)) Then
                Set sv2 = New SubdiarioVentasDetalle
                sv2.FEcha = CStr(sv.FEcha)
                sv2.Exento = sv.Exento
                sv2.Iva = sv.Iva
                sv2.NetoGravado = sv.NetoGravado
                sv2.percepciones = sv.percepciones
                sv2.Total = sv.Total
                newcol2.Add sv2, CStr(sv.FEcha)
            Else
                Set sv2 = newcol2.item(CStr(sv.FEcha))
                sv2.Exento = sv2.Exento + sv.Exento
                sv2.Iva = sv2.Iva + sv.Iva
                sv2.NetoGravado = sv2.NetoGravado + sv.NetoGravado
                sv2.percepciones = sv2.percepciones + sv.percepciones
                sv2.Total = sv2.Total + sv.Total
                sv2.FEcha = CStr(sv2.FEcha)
            End If
        Next
    ElseIf groupmethod = GroupByMonth Then
        For Each sv In col
            If Not BuscarEnColeccion(newcol2, CStr(ArmarSTR_YYYY_MM(sv.FEcha))) Then
                Set sv2 = New SubdiarioVentasDetalle
                sv2.FEcha = CStr(ArmarSTR_YYYY_MM(sv.FEcha))
                sv2.Exento = sv.Exento
                sv2.Iva = sv.Iva
                sv2.NetoGravado = sv.NetoGravado
                sv2.percepciones = sv.percepciones
                sv2.Total = sv.Total
                newcol2.Add sv2, CStr(ArmarSTR_YYYY_MM(sv2.FEcha))
            Else
                Set sv2 = newcol2.item(CStr(ArmarSTR_YYYY_MM(sv.FEcha)))
                sv2.Exento = sv2.Exento + sv.Exento
                sv2.Iva = sv2.Iva + sv.Iva
                sv2.NetoGravado = sv2.NetoGravado + sv.NetoGravado
                sv2.percepciones = sv2.percepciones + sv.percepciones
                sv2.Total = sv2.Total + sv.Total
                sv2.FEcha = CStr(ArmarSTR_YYYY_MM(sv.FEcha))
            End If
        Next
    ElseIf groupmethod = GroupByYear Then
        For Each sv In col
            If Not BuscarEnColeccion(newcol2, CStr(ArmarSTR_YYYY(sv.FEcha))) Then

                Set sv2 = New SubdiarioVentasDetalle
                sv2.FEcha = CStr(ArmarSTR_YYYY(sv.FEcha))
                sv2.Exento = sv.Exento
                sv2.Iva = sv.Iva
                sv2.NetoGravado = sv.NetoGravado
                sv2.percepciones = sv.percepciones
                sv2.Total = sv.Total
                newcol2.Add sv2, CStr(ArmarSTR_YYYY(sv.FEcha))
            Else
                Set sv2 = newcol2.item(CStr(ArmarSTR_YYYY(sv.FEcha)))
                sv2.FEcha = CStr(ArmarSTR_YYYY(sv.FEcha))
                sv2.Exento = sv2.Exento + sv.Exento
                sv2.Iva = sv2.Iva + sv.Iva
                sv2.NetoGravado = sv2.NetoGravado + sv.NetoGravado
                sv2.percepciones = sv2.percepciones + sv.percepciones
                sv2.Total = sv2.Total + sv.Total

            End If
        Next
    End If
    Set AgruparColeccion = newcol2
End Function
Private Function ArmarSTR_YYYY_MM(FEcha As Date) As String
    ArmarSTR_YYYY_MM = CStr(Year(FEcha)) & "-" & CStr(Month(FEcha))
End Function
Private Function ArmarSTR_YYYY(FEcha As Date) As String
    ArmarSTR_YYYY = CStr(Year(FEcha))
End Function
Private Sub PushButton3_Click()
    Dim groupmethod As FcGroupMethod
    If Me.RadioButton1.value Then
        groupmethod = GroupByDate
    ElseIf Me.RadioButton2.value Then
        groupmethod = GroupByMonth
    ElseIf Me.RadioButton3.value Then
        groupmethod = GroupByYear
    End If
    Set newcol = AgruparColeccion(col, groupmethod)
    Dim i As Integer
    Dim dto As SubdiarioVentasDetalle
    i = 0
    If newcol.count > 0 Then
        Me.grafica.Visible = True
    Else
        Exit Sub
    End If
    grafica.ColumnCount = 2
    grafica.rowcount = newcol.count
    For Each dto In newcol
        i = i + 1
        grafica.Column = 1
        grafica.row = i
        grafica.RowLabel = dto.FEcha
        grafica.data = dto.NetoGravado
        grafica.Column = 2
        grafica.data = dto.Iva
    Next dto
    Me.chkGrafico.value = xtpChecked
    chkGrafico_Click
    llenarLista
    grafica.Visible = True
    grafica.Refresh
End Sub

Private Sub PushButton4_Click()
    Me.cboProvincias.ListIndex = -1
End Sub
