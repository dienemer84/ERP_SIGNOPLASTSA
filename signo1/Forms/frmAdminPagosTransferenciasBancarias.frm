VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosTransferenciasBancarias 
   Caption         =   "Transferencias Bancarias en OP"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   17475
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   17175
      _Version        =   786432
      _ExtentX        =   30295
      _ExtentY        =   3201
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.ComboBox cboCuentaBancaria 
         Height          =   315
         Left            =   1155
         TabIndex        =   17
         Top             =   620
         Width           =   3885
      End
      Begin VB.TextBox txtComprobante 
         Height          =   315
         Left            =   1155
         TabIndex        =   16
         Top             =   1000
         Width           =   3885
      End
      Begin VB.TextBox txtOP 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         TabIndex        =   11
         Top             =   1365
         Width           =   2205
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   495
         Left            =   14640
         TabIndex        =   10
         Top             =   240
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnTraerDatos 
         Height          =   495
         Left            =   14640
         TabIndex        =   2
         Top             =   960
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1215
         Index           =   1
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Fecha de Operación"
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   720
            TabIndex        =   4
            Top             =   720
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
            Left            =   2925
            TabIndex        =   5
            Top             =   720
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
            Left            =   720
            TabIndex        =   6
            Top             =   300
            Width           =   3675
            _Version        =   786432
            _ExtentX        =   6482
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   165
            TabIndex        =   8
            Top             =   780
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2400
            TabIndex        =   7
            Top             =   780
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            BackColor       =   12632256
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   1155
         TabIndex        =   12
         Top             =   240
         Width           =   3885
         _Version        =   786432
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton CMDsINCliente 
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   13
         Top             =   270
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CMDsINCtaBancaria 
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   20
         Top             =   650
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "X"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   680
         Width           =   960
         _Version        =   786432
         _ExtentX        =   1693
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cta. Bancaria"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   18
         Top             =   1060
         Width           =   945
         _Version        =   786432
         _ExtentX        =   1667
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Comprobante"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   315
         TabIndex        =   15
         Top             =   300
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   14
         Top             =   1440
         Width           =   450
         _Version        =   786432
         _ExtentX        =   794
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nº OP"
         BackColor       =   12632256
         Enabled         =   0   'False
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX gridTransferencias 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   11668
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowCardSizing =   0   'False
      AllowEdit       =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmAdminPagosTransferenciasBancarias.frx":0000
      Column(2)       =   "frmAdminPagosTransferenciasBancarias.frx":0180
      Column(3)       =   "frmAdminPagosTransferenciasBancarias.frx":02D8
      Column(4)       =   "frmAdminPagosTransferenciasBancarias.frx":0434
      Column(5)       =   "frmAdminPagosTransferenciasBancarias.frx":0598
      Column(6)       =   "frmAdminPagosTransferenciasBancarias.frx":06E8
      Column(7)       =   "frmAdminPagosTransferenciasBancarias.frx":0838
      Column(8)       =   "frmAdminPagosTransferenciasBancarias.frx":0998
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosTransferenciasBancarias.frx":0AE0
      FormatStyle(2)  =   "frmAdminPagosTransferenciasBancarias.frx":0C18
      FormatStyle(3)  =   "frmAdminPagosTransferenciasBancarias.frx":0CC8
      FormatStyle(4)  =   "frmAdminPagosTransferenciasBancarias.frx":0D7C
      FormatStyle(5)  =   "frmAdminPagosTransferenciasBancarias.frx":0E54
      FormatStyle(6)  =   "frmAdminPagosTransferenciasBancarias.frx":0F0C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosTransferenciasBancarias.frx":0FEC
   End
   Begin XtremeSuiteControls.Label Label 
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   2040
      Width           =   17175
      _Version        =   786432
      _ExtentX        =   30295
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   $"frmAdminPagosTransferenciasBancarias.frx":11C4
   End
End
Attribute VB_Name = "frmAdminPagosTransferenciasBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private transferencias As New Collection
Private TransfBancaria As clsTransferenciaBcaria
Private desde
Private colProveedores As New Collection
Private colCuentasBancarias As New Collection
Private prov As clsProveedor
Private ctabancaria As CuentaBancaria


Private Sub btnExportar_Click()

    If IsSomething(transferencias) Then
        If Not DAOTransferenciaBcaria.ExportarColeccion(transferencias) Then GoTo err1
    End If

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"

    
End Sub

Private Sub btnTraerDatos_Click()
    CompletarGridEx

End Sub

Private Sub CompletarGridEx()

    Me.gridTransferencias.ItemCount = 0

    Dim condition As String
    condition = " 1 = 1 "

    If Not IsNull(Me.dtpDesde.value) Then
        condition = condition & " AND op.fecha_operacion >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        condition = condition & " AND op.fecha_operacion <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If cboProveedores.ListIndex > -1 Then
        condition = condition & " AND prov.id = " & cboProveedores.ItemData(Me.cboProveedores.ListIndex)
    End If
    
    If Me.cboCuentaBancaria.ListIndex > -1 Then
        condition = condition & " AND cu.id = " & Me.cboCuentaBancaria.ItemData(Me.cboCuentaBancaria.ListIndex)
    End If
    
    If LenB(Me.txtOP) > 0 Then
        condition = condition & " AND opope.id_orden_pago like '%" & Trim(Me.txtOP.text) & "%'"
    End If

    If LenB(Me.txtComprobante) > 0 Then
        condition = condition & " AND op.comprobante like '%" & Trim(Me.txtComprobante.text) & "%'"
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set transferencias = DAOTransferenciaBcaria.FindAll(Banco, condition)
    
    Me.gridTransferencias.ItemCount = transferencias.count

    GridEXHelper.AutoSizeColumns Me.gridTransferencias, True

    Me.caption = "Transferencias [Cantidad: " & transferencias.count & "]"
  
End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta

End Sub


Private Sub CMDsINCliente_Click(Index As Integer)
    Me.cboProveedores.ListIndex = -1
End Sub


Private Sub CMDsINCtaBancaria_Click(Index As Integer)
    Me.cboCuentaBancaria.ListIndex = -1
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.gridTransferencias, True, True
    
    Me.Height = 9855
    Me.Width = 17595
    
    'INICIO- GroupBox de Fecha de Operación
    Dim i As Integer
    
    desde = DateSerial(Year(Date), Month(Date), 1)
    funciones.FillComboBoxDateRanges Me.cboRangos
    
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
    
    'FIN- GroupBox de Fecha de Operación
    
    'INICIO- Llenado de Combo Proveedores
    Set colProveedores = DAOProveedor.FindAll
    For Each prov In colProveedores
        cboProveedores.AddItem UCase(prov.RazonSocial)
        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.Id
    Next
    'FIN- Llenado de Combo Proveedores
    
    'INICIO- Llenado de Combo Proveedores
    Set colCuentasBancarias = DAOCuentaBancaria.FindAll
    For Each ctabancaria In colCuentasBancarias
        cboCuentaBancaria.AddItem "N° " & ctabancaria.numero & " | " & ctabancaria.Banco.nombre
        cboCuentaBancaria.ItemData(cboCuentaBancaria.NewIndex) = ctabancaria.Id
    Next
    'FIN- Llenado de Combo Proveedores
    
End Sub


Private Sub gridTransferencias_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
If rowIndex > 0 And transferencias.count > 0 Then
    Set TransfBancaria = transferencias.item(rowIndex)
        Values(1) = TransfBancaria.Id
        

        Values(3) = "N° " & TransfBancaria.CuentaBancaria & " | " & TransfBancaria.NombreBanco
        Values(4) = TransfBancaria.FechaOperacion
        Values(5) = TransfBancaria.moneda.NombreCorto
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(TransfBancaria.Monto)), "$", "")
        Values(7) = TransfBancaria.Comprobante
        
        If TransfBancaria.LiquidacionCaja Is Nothing Then
                Values(8) = "OP: " & TransfBancaria.OrdenPago.Id
                Values(2) = UCase(TransfBancaria.ProveedorRazon)
        Else
                Values(8) = "LIQ: " & TransfBancaria.LiquidacionCaja.NumeroLiq
                Values(2) = "VARIOS"
        End If


End If
 
End Sub

Private Sub gridTransferencias_SelectionChange()
    On Error Resume Next
    Set TransfBancaria = transferencias.item(gridTransferencias.rowIndex(gridTransferencias.row))
End Sub


Private Sub gridTransferencias_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridTransferencias, Column
End Sub

Private Sub gridTransferencias_DblClick()
    gridTransferencias_SelectionChange
    mnuVer_Click
End Sub

Private Sub mnuVer_Click()
    
    If TransfBancaria.LiquidacionCaja Is Nothing Then
        Dim f22 As New frmAdminPagosCrearOrdenPago
        f22.Show
        f22.ReadOnly = True
        f22.Cargar TransfBancaria.OrdenPago
    Else
        Dim f25 As New frmAdminPagosLiqCajaListaDG
        f25.Show
        f25.ReadOnly = True
        f25.Cargar TransfBancaria.LiquidacionCaja
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' 13 es el código ASCII de la tecla Enter
        ' Realizar la acción de búsqueda aquí
        CompletarGridEx
    End If
End Sub
