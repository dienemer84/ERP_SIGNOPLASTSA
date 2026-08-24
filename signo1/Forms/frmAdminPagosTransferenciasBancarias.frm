VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosTransferenciasBancarias 
   Caption         =   "Transferencias de pagos"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   -210
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleMode       =   0  'User
   ScaleWidth      =   18082.37
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   17175
      _Version        =   786432
      _ExtentX        =   30295
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
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
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1335
         Index           =   2
         Left            =   5760
         TabIndex        =   21
         Top             =   240
         Width           =   3855
         _Version        =   786432
         _ExtentX        =   6800
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Importes"
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
         Begin VB.TextBox textbMayor 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox textbMenor 
            Height          =   315
            Left            =   2040
            TabIndex        =   22
            Top             =   720
            Width           =   1215
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   24
            Top             =   750
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   25
            Top             =   750
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "X"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label 
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   27
            Top             =   480
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Menor que:"
         End
         Begin XtremeSuiteControls.Label Label 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Mayor que:"
         End
      End
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
         Visible         =   0   'False
         Width           =   2205
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   495
         Left            =   14760
         TabIndex        =   10
         Top             =   1080
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
         Index           =   0
         Left            =   14760
         TabIndex        =   2
         Top             =   240
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
         Height          =   1335
         Index           =   1
         Left            =   9720
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Fecha de Operación"
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Visible         =   0   'False
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
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   11456
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
      ColumnsCount    =   10
      Column(1)       =   "frmAdminPagosTransferenciasBancarias.frx":0000
      Column(2)       =   "frmAdminPagosTransferenciasBancarias.frx":0180
      Column(3)       =   "frmAdminPagosTransferenciasBancarias.frx":02D8
      Column(4)       =   "frmAdminPagosTransferenciasBancarias.frx":0434
      Column(5)       =   "frmAdminPagosTransferenciasBancarias.frx":0610
      Column(6)       =   "frmAdminPagosTransferenciasBancarias.frx":077C
      Column(7)       =   "frmAdminPagosTransferenciasBancarias.frx":0940
      Column(8)       =   "frmAdminPagosTransferenciasBancarias.frx":0ABC
      Column(9)       =   "frmAdminPagosTransferenciasBancarias.frx":0C30
      Column(10)      =   "frmAdminPagosTransferenciasBancarias.frx":0DA4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosTransferenciasBancarias.frx":0F20
      FormatStyle(2)  =   "frmAdminPagosTransferenciasBancarias.frx":1058
      FormatStyle(3)  =   "frmAdminPagosTransferenciasBancarias.frx":1108
      FormatStyle(4)  =   "frmAdminPagosTransferenciasBancarias.frx":11BC
      FormatStyle(5)  =   "frmAdminPagosTransferenciasBancarias.frx":1294
      FormatStyle(6)  =   "frmAdminPagosTransferenciasBancarias.frx":134C
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosTransferenciasBancarias.frx":142C
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Begin VB.Menu mnuVer 
         Caption         =   "Ver Documento de Pago"
      End
      Begin VB.Menu mnuModificar 
         Caption         =   "Modificar detalles"
      End
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
        condition = condition & " AND (prov.id = " & cboProveedores.ItemData(Me.cboProveedores.ListIndex) & " OR prov1.Id = " & cboProveedores.ItemData(Me.cboProveedores.ListIndex) & ")"
    End If
    
    If Me.cboCuentaBancaria.ListIndex > -1 Then
        condition = condition & " AND cu.id = " & Me.cboCuentaBancaria.ItemData(Me.cboCuentaBancaria.ListIndex)
    End If
    
    If LenB(Me.txtOP) > 0 Then
        condition = condition & " AND opope.id_orden_pago like '%" & Trim(Me.txtOP.Text) & "%'"
    End If

    If LenB(Me.txtComprobante) > 0 Then
        condition = condition & " AND op.comprobante like '%" & Trim(Me.txtComprobante.Text) & "%'"
    End If
    
    If LenB(Me.textbMayor) > 0 Then
        condition = condition & " AND op.monto >= " & Me.textbMayor.Text
    End If
    
    If LenB(Me.textbMenor) > 0 Then
        condition = condition & " AND op.monto <= " & Me.textbMenor.Text
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set transferencias = _
        DAOTransferenciaBcaria.FindAll( _
            Banco, _
            condition, _
            "op.id DESC", _
            True)
    
    Me.gridTransferencias.ItemCount = transferencias.count

    GridEXHelper.AutoSizeColumns Me.gridTransferencias, True

    Me.caption = "Transferencias [Cantidad: " & transferencias.count & "]"
    
    Me.Label3.caption = "Transferencias mostradas [" & transferencias.count & "]"
  
End Sub

Public Sub btnTraerDatos_Click(Index As Integer)
    CompletarGridEx

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
    
'''    Set colProveedores = DAOProveedor.FindAll
'''    For Each prov In colProveedores
'''        cboProveedores.AddItem UCase(prov.RazonSocial)
'''        cboProveedores.ItemData(cboProveedores.NewIndex) = prov.Id
'''    Next

    Call DAOProveedor.llenarComboProveedores(cboProveedores)
    Me.cboProveedores.ListIndex = -1

    'FIN- Llenado de Combo Proveedores
    
    'INICIO- Llenado de Combo Proveedores
    

    
    Set colCuentasBancarias = DAOCuentaBancaria.FindAll
    For Each ctabancaria In colCuentasBancarias
        cboCuentaBancaria.AddItem ctabancaria.DescripcionFormateada
        cboCuentaBancaria.ItemData(cboCuentaBancaria.NewIndex) = ctabancaria.Id
    Next
    'FIN- Llenado de Combo Proveedores
    
    Me.gridTransferencias.ItemCount = 0
    
    Me.Label3.caption = "Transferencias mostradas [" & transferencias.count & "]"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.gridTransferencias.Width = Me.ScaleWidth - 400
    Me.gridTransferencias.Height = Me.ScaleHeight - 3200

    GridEXHelper.AutoSizeColumns Me.gridTransferencias
End Sub

Private Sub gridTransferencias_SelectionChange()
    On Error Resume Next
    Set TransfBancaria = BuscarTransferenciaSeleccionada()
End Sub


Private Sub gridTransferencias_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If transferencias.count = 0 Then Exit Sub

    If Button = 2 Then
        ' Acá idealmente seleccionar la fila bajo el mouse
        ' según el método que soporte tu GridEX
        
        gridTransferencias_SelectionChange
        
        Me.mnuVer.Enabled = Not (TransfBancaria Is Nothing)
        Me.mnuModificar.Enabled = Not (TransfBancaria Is Nothing)
        
        Me.PopupMenu menu
    End If
End Sub


Private Sub gridTransferencias_UnboundReadData( _
    ByVal RowIndex As Long, _
    ByVal Bookmark As Variant, _
    ByVal Values As GridEX20.JSRowData)

    On Error GoTo err1

    If RowIndex <= 0 Then Exit Sub
    If transferencias.count = 0 Then Exit Sub
    If RowIndex > transferencias.count Then Exit Sub

    Dim T As clsTransferenciaBcaria

    Set T = transferencias.item(RowIndex)

    Values(1) = T.Id
    Values(3) = DescripcionOrigen(T)
    Values(4) = T.FechaOperacion

    If IsSomething(T.moneda) Then
        Values(5) = T.moneda.NombreCorto
    Else
        Values(5) = ""
    End If

    Values(6) = Replace( _
                    FormatCurrency( _
                        funciones.FormatearDecimales(T.Monto)), _
                    "$", "")

    Values(7) = T.Comprobante

    '================================================
    ' MOVIMIENTO MANUAL DE CAJA Y BANCOS
    '================================================
    If T.EsMovimientoCajaBanco Then

        Values(2) = "MOVIMIENTO MANUAL"

        Values(8) = _
            "MOV: " & T.MovimientoCajaBancoID

        If T.Pertenencia = Banco Then

            Values(9) = _
                UCase$(T.TipoMovimientoCajaBanco) & _
                " - BANCO"

        ElseIf T.Pertenencia = caja Then

            Values(9) = _
                UCase$(T.TipoMovimientoCajaBanco) & _
                " - CAJA"

        Else

            Values(9) = _
                UCase$(T.TipoMovimientoCajaBanco)

        End If

        Values(10) = ""

        Exit Sub

    End If

    '================================================
    ' PAGO A CUENTA / ORDEN DE PAGO / LIQUIDACIÓN
    '================================================
    If T.LiquidacionCaja Is Nothing Then

        If T.OrdenPago Is Nothing Then

            Values(8) = "PCTA: " & T.PagoACuentaID
            Values(2) = UCase$(T.PagoACuentaProveedor)

            If T.OPAplicada = 0 Then
                Values(9) = "Disponible"
            Else
                Values(9) = "Procesada"
            End If

            Values(10) = T.OPAplicada

        Else

            Values(8) = "OP: " & T.OrdenPago.Id
            Values(2) = UCase$(T.ProveedorRazon)
            Values(9) = ""
            Values(10) = ""

        End If

    Else

        Values(8) = _
            "LIQ: " & T.LiquidacionCaja.NumeroLiq

        Values(2) = "VARIOS"
        Values(9) = ""
        Values(10) = ""

    End If

    Exit Sub

err1:
    Debug.Print _
        "gridTransferencias_UnboundReadData: " & _
        Err.Number & " - " & Err.Description

End Sub

'Private Sub gridTransferencias_SelectionChange()
'    On Error Resume Next
'    Set TransfBancaria = transferencias.item(gridTransferencias.rowIndex(gridTransferencias.row))
'End Sub



Private Sub gridTransferencias_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridTransferencias, Column
End Sub


Private Sub mnuModificar_Click()

    Set TransfBancaria = BuscarTransferenciaSeleccionada()

    If TransfBancaria Is Nothing Then Exit Sub

    '================================================
    ' MOVIMIENTO DE CAJA Y BANCOS
    '================================================
    If TransfBancaria.EsMovimientoCajaBanco Then

        Dim mov As clsAsientoContable

        Set mov = DAOAsientoContable.FindById( _
                    TransfBancaria.MovimientoCajaBancoID)

        If mov Is Nothing Then

            MsgBox _
                "No se encontró el movimiento de Caja y Bancos.", _
                vbExclamation

            Exit Sub

        End If

        Dim fMov As New frmAdminCajaBancosCrearAsientoBancario

        Load fMov

        fMov.ReadOnly = False
        fMov.Cargar mov
        fMov.Show

        Exit Sub

    End If

    '================================================
    ' TRANSFERENCIA TRADICIONAL
    '================================================
    Dim f_ADFE As New _
        frmAdminPagosTransferenciasBancariasEditar

    f_ADFE.idTransfBancaria = TransfBancaria.Id
    f_ADFE.Show

End Sub

Private Sub mnuVer_Click()

    Set TransfBancaria = BuscarTransferenciaSeleccionada()

    If TransfBancaria Is Nothing Then Exit Sub

    '================================================
    ' MOVIMIENTO DE CAJA Y BANCOS
    '================================================
    If TransfBancaria.EsMovimientoCajaBanco Then

        Dim mov As clsAsientoContable

        Set mov = DAOAsientoContable.FindById( _
                    TransfBancaria.MovimientoCajaBancoID)

        If mov Is Nothing Then

            MsgBox _
                "No se encontró el movimiento de Caja y Bancos.", _
                vbExclamation

            Exit Sub

        End If

        Dim fMov As New frmAdminCajaBancosCrearAsientoBancario

        Load fMov

        fMov.ReadOnly = True
        fMov.Cargar mov
        fMov.Show

        Exit Sub

    End If

    '================================================
    ' LIQUIDACIÓN
    '================================================
    If Not TransfBancaria.LiquidacionCaja Is Nothing Then

        Dim f25 As New frmAdminPagosLiqCajaListaDG

        MsgBox _
            "Abriendo Liquidación de Caja: " & _
            TransfBancaria.LiquidacionCaja.NumeroLiq

        Load f25

        f25.ReadOnly = True
        f25.Cargar TransfBancaria.LiquidacionCaja
        f25.Show

        Exit Sub

    End If

    '================================================
    ' ORDEN DE PAGO
    '================================================
    If Not TransfBancaria.OrdenPago Is Nothing Then

        Dim f22 As New frmAdminPagosCrearOrdenPago

        MsgBox _
            "Abriendo OP: " & _
            TransfBancaria.OrdenPago.Id

        Load f22

        f22.ReadOnly = True
        f22.Cargar TransfBancaria.OrdenPago
        f22.Show

        Exit Sub

    End If

    '================================================
    ' PAGO A CUENTA
    '================================================
    If TransfBancaria.PagoACuentaID > 0 Then

        MsgBox _
            "La transferencia seleccionada corresponde " & _
            "a un Pago a Cuenta.", _
            vbInformation

        Exit Sub

    End If

End Sub
    



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' 13 es el código ASCII de la tecla Enter
        ' Realizar la acción de búsqueda aquí
        CompletarGridEx
    End If
End Sub


Private Sub PushButton2_Click(Index As Integer)
    Me.textbMenor.Text = ""
End Sub

Private Sub PushButton1_Click(Index As Integer)
    Me.textbMayor.Text = ""
End Sub


Private Function BuscarTransferenciaSeleccionada() As clsTransferenciaBcaria
    Dim idTransf As Long
    Dim T As clsTransferenciaBcaria
    
    On Error GoTo err1
    
    Set BuscarTransferenciaSeleccionada = Nothing
    
    If Me.gridTransferencias.row <= 0 Then Exit Function
    
    idTransf = CLng(Me.gridTransferencias.value(1))
    
    For Each T In transferencias
        If T.Id = idTransf Then
            Set BuscarTransferenciaSeleccionada = T
            Exit Function
        End If
    Next
    
    Exit Function
err1:
    Set BuscarTransferenciaSeleccionada = Nothing
    
End Function


Private Sub gridTransferencias_DblClick()
'''    On Error Resume Next
'''
'''    MsgBox "Row: " & Me.gridTransferencias.row & vbCrLf & _
'''           "RowIndex: " & Me.gridTransferencias.RowIndex(Me.gridTransferencias.row) & vbCrLf & _
'''           "Col1 value: " & Me.gridTransferencias.value(Me.gridTransferencias.Columns(1).Index)
'''
'''    gridTransferencias_SelectionChange
'''
'''    mnuVer_Click
    
End Sub


Private Function DescripcionOrigen( _
    ByVal T As clsTransferenciaBcaria _
) As String

    If T Is Nothing Then
        DescripcionOrigen = vbNullString
        Exit Function
    End If

    If T.Pertenencia = Banco Then

        DescripcionOrigen = _
            "N° " & T.CuentaBancaria & _
            " | " & T.NombreBanco

    ElseIf T.Pertenencia = caja Then

        DescripcionOrigen = _
            "CAJA | " & T.NombreCaja

    Else

        DescripcionOrigen = "SIN ORIGEN"

    End If

End Function

