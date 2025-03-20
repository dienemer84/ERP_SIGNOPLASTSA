VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminMasInfoCliente2 
   Caption         =   "Resúmen de Facturación por Cliente / OT"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "frmAdminMasInfoCliente2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11550
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   375
      Left            =   150
      TabIndex        =   20
      Top             =   6315
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   7858
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      ColumnAutoResize=   -1  'True
      MultiSelect     =   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   11
      Column(1)       =   "frmAdminMasInfoCliente2.frx":000C
      Column(2)       =   "frmAdminMasInfoCliente2.frx":011C
      Column(3)       =   "frmAdminMasInfoCliente2.frx":0210
      Column(4)       =   "frmAdminMasInfoCliente2.frx":031C
      Column(5)       =   "frmAdminMasInfoCliente2.frx":0420
      Column(6)       =   "frmAdminMasInfoCliente2.frx":0540
      Column(7)       =   "frmAdminMasInfoCliente2.frx":0634
      Column(8)       =   "frmAdminMasInfoCliente2.frx":077C
      Column(9)       =   "frmAdminMasInfoCliente2.frx":08D8
      Column(10)      =   "frmAdminMasInfoCliente2.frx":0A30
      Column(11)      =   "frmAdminMasInfoCliente2.frx":0B50
      GroupCount      =   1
      Group(1)        =   "frmAdminMasInfoCliente2.frx":0C74
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminMasInfoCliente2.frx":0CDC
      FormatStyle(2)  =   "frmAdminMasInfoCliente2.frx":0E14
      FormatStyle(3)  =   "frmAdminMasInfoCliente2.frx":0EC4
      FormatStyle(4)  =   "frmAdminMasInfoCliente2.frx":0F78
      FormatStyle(5)  =   "frmAdminMasInfoCliente2.frx":1050
      FormatStyle(6)  =   "frmAdminMasInfoCliente2.frx":1108
      ImageCount      =   0
      PrinterProperties=   "frmAdminMasInfoCliente2.frx":11E8
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1530
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   11340
      _Version        =   786432
      _ExtentX        =   20002
      _ExtentY        =   2699
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1020
         TabIndex        =   29
         Top             =   285
         Width           =   2835
         _Version        =   786432
         _ExtentX        =   5001
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.TextBox txtPorcentajeFacturado 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1755
         TabIndex        =   27
         Top             =   1050
         Width           =   585
      End
      Begin XtremeSuiteControls.CheckBox chkMarcos 
         Height          =   255
         Left            =   9270
         TabIndex        =   21
         Top             =   1005
         Width           =   1950
         _Version        =   786432
         _ExtentX        =   3440
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Excluír Contratos Marco"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1050
         Left            =   4470
         TabIndex        =   4
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Fecha Entrega"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   825
            TabIndex        =   5
            Top             =   615
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   3000
            TabIndex        =   6
            Top             =   600
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   825
            TabIndex        =   7
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2430
            TabIndex        =   10
            Top             =   660
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   270
            TabIndex        =   9
            Top             =   645
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   195
            Left            =   255
            TabIndex        =   8
            Top             =   270
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton btnListar 
         Height          =   375
         Left            =   9480
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Listar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   285
         Left            =   3945
         TabIndex        =   11
         Top             =   255
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnClearEstadoOT 
         Height          =   285
         Left            =   3945
         TabIndex        =   24
         Top             =   645
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoOT 
         Height          =   315
         Left            =   1020
         TabIndex        =   23
         Top             =   645
         Width           =   2850
         _Version        =   786432
         _ExtentX        =   5027
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboMayorMenor 
         Height          =   315
         Left            =   1020
         TabIndex        =   26
         Top             =   1050
         Width           =   675
         _Version        =   786432
         _ExtentX        =   1191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnClearFacturado 
         Height          =   285
         Left            =   2445
         TabIndex        =   28
         Top             =   1065
         Width           =   330
         _Version        =   786432
         _ExtentX        =   582
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "% Fact"
         Height          =   195
         Left            =   405
         TabIndex        =   25
         Top             =   1110
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado OT"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   675
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   285
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.GroupBox grpTotales 
      Height          =   1200
      Left            =   8325
      TabIndex        =   12
      Top             =   6255
      Width           =   3090
      _Version        =   786432
      _ExtentX        =   5450
      _ExtentY        =   2117
      _StockProps     =   79
      Caption         =   "Totales"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label lblExento 
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   780
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total Pendiente:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNetoGravado 
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   225
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total a Facturar:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIVA 
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   495
         Width           =   1170
         _Version        =   786432
         _ExtentX        =   2064
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total Facturado:"
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblTotal 
         Height          =   195
         Left            =   1425
         TabIndex        =   16
         Top             =   225
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblTotalFacturado 
         Height          =   195
         Left            =   1425
         TabIndex        =   15
         Top             =   495
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblTotalPendiente 
         Height          =   195
         Left            =   1425
         TabIndex        =   14
         Top             =   780
         Width           =   1515
         _Version        =   786432
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblTotalTotal 
         Height          =   195
         Left            =   1785
         TabIndex        =   13
         Top             =   1380
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   ".-"
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmAdminMasInfoCliente2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private ordenes As Collection
Dim cliente As clsCliente
Dim Ot As OrdenTrabajo
Dim tot As Double
Dim totPend As Double
Dim totFact As Double


Private Sub btnClearEstadoOT_Click()
    Me.cboEstadoOT.ListIndex = -1
End Sub

Private Sub btnClearFacturado_Click()
    Me.cboMayorMenor.ListIndex = -1
    Me.txtPorcentajeFacturado.Text = vbNullString
End Sub

Private Sub btnListar_Click()
    llenarLista
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub

Private Sub cmdImprimir_Click()
    Dim headercenter As String
    Dim footerright As String
    Dim elegidos As Boolean
    elegidos = (Me.GridEX1.SelectedItems.count > 1)

    headercenter = "RESUMEN FACTURACION POR CLIENTE"


    If Me.cboClientes.ListIndex <> -1 Then
        headercenter = headercenter & Chr(10) & " Cliente: " & Me.cboClientes.Text
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        headercenter = headercenter & Chr(10) & " Desde: " & Me.dtpDesde.value
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        headercenter = headercenter & "  Hasta: " & Me.dtpHasta.value
    End If

    If Me.cboEstadoOT.ListIndex <> -1 Then
        headercenter = headercenter & Chr(10) & " Estado: " & funciones.estado_pedido(Me.cboEstadoOT.ItemData(Me.cboEstadoOT.ListIndex))
    End If

    If Me.cboMayorMenor.ListIndex <> -1 Then
        headercenter = headercenter & Chr(10) & " Facturado:   " & Me.cboMayorMenor.Text & " " & Me.txtPorcentajeFacturado & " %"
    End If

    footerright = "Total A Facturar: " & MonedaConverter.Patron.NombreCorto & " " & tot & Chr(10) _
                & "Total Facturado: " & MonedaConverter.Patron.NombreCorto & " " & totFact & Chr(10) _
                & "Total Pendiente: " & MonedaConverter.Patron.NombreCorto & " " & tot - totFact





    With Me.GridEX1.PrinterProperties
        .HeaderDistance = 500
        .FooterDistance = 1550
        .TopMargin = 2000
        .BottomMargin = 2000


        .FitColumns = True
        .DocumentName = "RESUMEN DE FACTURACIÓN POR OT"

        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = headercenter


        .FooterString(jgexHFLeft) = footerright
        .FooterString(jgexHFCenter) = Now



    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    GridEX1.PrintPreview frmPrintPreview.GEXPreview1, elegidos
    frmPrintPreview.Show 1
End Sub


Private Sub CMDsINCliente_Click()
    Me.cboClientes.ListIndex = -1
End Sub

Private Sub Form_Activate()
    Me.GridEX1.Refresh
End Sub

Private Sub Form_Load()
    Customize Me
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False
    GridEXHelper.AutoSizeColumns Me.GridEX1
    Me.GridEX1.ItemCount = 0

    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i


    Me.cboEstadoOT.Clear
    For i = LBound(funciones.estados_pedidos) To UBound(funciones.estados_pedidos)
        Me.cboEstadoOT.AddItem estados_pedidos(i)
        Me.cboEstadoOT.ItemData(Me.cboEstadoOT.NewIndex) = i
    Next i
    Me.cboEstadoOT.ListIndex = -1

    Me.cboMayorMenor.Clear
    Me.cboMayorMenor.AddItem ">="
    Me.cboMayorMenor.AddItem "<="
    Me.cboMayorMenor.ListIndex = -1

End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Me.GridEX1.Width = Me.ScaleWidth - 250
    Me.GridEX1.Height = Me.ScaleHeight - Me.GroupBox2.Height - 2200
    Me.grpTotales.Top = Me.GridEX1.Height + Me.GroupBox1.Height + 400
    Me.grpTotales.Left = Me.Width - Me.grpTotales.Width - 280
    Me.cmdImprimir.Top = Me.GridEX1.Height + 1800
End Sub

Private Sub llenarLista()
    Dim q As String
    q = "1 = 1 "
    tot = 0
    totFact = 0
    totPend = 0

    If Me.cboClientes.ListIndex <> -1 Then
        q = q & " AND idClienteFacturar = " & Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        q = q & " AND " & DAOOrdenTrabajo.TABLA_PEDIDO & "." & DAOOrdenTrabajo.CAMPO_FECHA_ENTREGA & " >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        q = q & " AND " & DAOOrdenTrabajo.TABLA_PEDIDO & "." & DAOOrdenTrabajo.CAMPO_FECHA_ENTREGA & " <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If Me.chkMarcos.value = xtpChecked Then
        q = q & " AND " & DAOOrdenTrabajo.TABLA_PEDIDO & ".id_ot_padre<>-1"
    End If

    If Me.cboEstadoOT.ListIndex <> -1 Then
        q = q & " AND " & DAOOrdenTrabajo.TABLA_PEDIDO & "." & DAOOrdenTrabajo.CAMPO_ESTADO & " = " & Me.cboEstadoOT.ItemData(Me.cboEstadoOT.ListIndex)
    End If





    Set ordenes = DAOOrdenTrabajo.FindAll(q, , , , True, True, True, False)

    Dim Ot As OrdenTrabajo


    Dim i As Long
    Dim value2Compare As Double
    Dim remove As Boolean
    If Me.cboMayorMenor.ListIndex <> -1 Then
        value2Compare = Val(Me.txtPorcentajeFacturado.Text)

        For i = ordenes.count To 1 Step -1
            remove = False
            Set Ot = ordenes.item(i)

            If Me.cboMayorMenor.ListIndex = 0 Then    '>=
                remove = (Ot.PorcentajeFacturado < value2Compare)
            ElseIf Me.cboMayorMenor.ListIndex = 1 Then    '<=
                remove = (Ot.PorcentajeFacturado > value2Compare)
            End If

            If remove Then ordenes.remove i
        Next i

    End If


    For Each Ot In ordenes

        tot = tot + Ot.total



        totFact = Ot.TotalFacturado + totFact
    Next Ot

    Me.lblTotal.caption = MonedaConverter.Patron.NombreCorto & " " & funciones.FormatearDecimales(tot, 2)
    Me.lblTotalFacturado.caption = MonedaConverter.Patron.NombreCorto & " " & funciones.FormatearDecimales(totFact, 2)
    Me.lblTotalPendiente.caption = MonedaConverter.Patron.NombreCorto & " " & funciones.FormatearDecimales(tot - totFact, 2)


    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = ordenes.count
    Me.GridEX1.Refresh

End Sub


Private Sub GridEX1_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    GridEX1.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & " de " & nPages
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.GridEX1
        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If

End Sub

Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set Ot = ordenes.item(rowIndex)
    'If OT.id = 1197 Then Stop

    Values(1) = Ot.IdFormateado
    Values(2) = Ot.ClienteFacturar.razon
    Values(3) = Ot.cliente.provincia.nombre

    Values(4) = Ot.descripcion
    Values(5) = Ot.FechaEntrega
    Values(6) = funciones.estado_pedido(Ot.estado)
    Values(7) = funciones.FormatearDecimales((Ot.total))    ', ot.Moneda.Id, MonedaConverter.Patron.Id))
    Values(8) = funciones.FormatearDecimales(Ot.TotalFacturado)
    Values(9) = funciones.FormatearDecimales(Values(7) - Values(8))
    Values(10) = funciones.FormatearDecimales((Values(8) * 100) / Values(7))
    'Values(10) = funciones.FormatearDecimales(MonedaConverter.Convertir(ot.PorcentajeEntregas, ot.Moneda.Id, MonedaConverter.Patron.Id))
    Values(11) = funciones.FormatearDecimales(Ot.PorcentajeEntregas)
    Exit Sub
err1:
End Sub




