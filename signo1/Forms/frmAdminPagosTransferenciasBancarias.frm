VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosTransferenciasBancarias 
   Caption         =   "Modificar"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   -210
   ClientWidth     =   17475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
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
         TabIndex        =   25
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
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox textbMenor 
            Height          =   315
            Left            =   2040
            TabIndex        =   26
            Top             =   720
            Width           =   1215
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   31
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
            TabIndex        =   30
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
         Left            =   14640
         TabIndex        =   10
         Top             =   960
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
         Left            =   14640
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
      Top             =   3120
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
      ColumnsCount    =   8
      Column(1)       =   "frmAdminPagosTransferenciasBancarias.frx":0000
      Column(2)       =   "frmAdminPagosTransferenciasBancarias.frx":0180
      Column(3)       =   "frmAdminPagosTransferenciasBancarias.frx":02D8
      Column(4)       =   "frmAdminPagosTransferenciasBancarias.frx":0434
      Column(5)       =   "frmAdminPagosTransferenciasBancarias.frx":0610
      Column(6)       =   "frmAdminPagosTransferenciasBancarias.frx":0760
      Column(7)       =   "frmAdminPagosTransferenciasBancarias.frx":0908
      Column(8)       =   "frmAdminPagosTransferenciasBancarias.frx":0A68
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosTransferenciasBancarias.frx":0BBC
      FormatStyle(2)  =   "frmAdminPagosTransferenciasBancarias.frx":0CF4
      FormatStyle(3)  =   "frmAdminPagosTransferenciasBancarias.frx":0DA4
      FormatStyle(4)  =   "frmAdminPagosTransferenciasBancarias.frx":0E58
      FormatStyle(5)  =   "frmAdminPagosTransferenciasBancarias.frx":0F30
      FormatStyle(6)  =   "frmAdminPagosTransferenciasBancarias.frx":0FE8
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosTransferenciasBancarias.frx":10C8
   End
   Begin XtremeSuiteControls.Label Label 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Width           =   13575
      _Version        =   786432
      _ExtentX        =   23945
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "* Modificar el número de transferencia ingresado."
   End
   Begin XtremeSuiteControls.Label Label 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   2520
      Width           =   13575
      _Version        =   786432
      _ExtentX        =   23945
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "* Ver los Documentos de Pago (OP/LIQ)."
   End
   Begin XtremeSuiteControls.Label Label 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   2280
      Width           =   17175
      _Version        =   786432
      _ExtentX        =   30295
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Menú del boton derecho sobre cada transferencia:"
   End
   Begin XtremeSuiteControls.Label Label 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   2040
      Width           =   17175
      _Version        =   786432
      _ExtentX        =   30295
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Se muestran las transferencias que están aplicadas a cada OP o Liquidación."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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

    Set transferencias = DAOTransferenciaBcaria.FindAll(Banco, condition, "op.id DESC")
    
    Me.gridTransferencias.ItemCount = transferencias.count

    GridEXHelper.AutoSizeColumns Me.gridTransferencias, True

    Me.caption = "Transferencias [Cantidad: " & transferencias.count & "]"
  
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

    Call DAOProveedor.LlenarComboProveedores(cboProveedores)
    Me.cboProveedores.ListIndex = -1

    'FIN- Llenado de Combo Proveedores
    
    'INICIO- Llenado de Combo Proveedores
    Set colCuentasBancarias = DAOCuentaBancaria.FindAll
    For Each ctabancaria In colCuentasBancarias
        cboCuentaBancaria.AddItem "N° " & ctabancaria.numero & " | " & ctabancaria.Banco.nombre
        cboCuentaBancaria.ItemData(cboCuentaBancaria.NewIndex) = ctabancaria.Id
    Next
    'FIN- Llenado de Combo Proveedores
    
    Me.gridTransferencias.ItemCount = 0
    
    
End Sub

Private Sub gridTransferencias_SelectionChange()
    On Error Resume Next
    Set TransfBancaria = transferencias.item(gridTransferencias.RowIndex(gridTransferencias.row))
End Sub


Private Sub gridTransferencias_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If transferencias.count > 0 Then
        gridTransferencias_SelectionChange
        If Button = 2 Then
'            Me.mnuAprobar.Enabled = (LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente)
'            Me.mnuEditar.Enabled = (LiquidacionCaja.estado = EstadoLiquidacionCaja_pendiente)
            Me.mnuVer.Enabled = True
            Me.mnuModificar.Enabled = True
            'OCULTO LA OPCION DE ANULAR QUE NO ESTÁ DESARROLLADA (DNEMER 30.05.2023)
            'Me.mnuAnular.Enabled = Not (LiquidacionCaja.estado = EstadoLiquidacionCaja_Anulada)

            Me.PopupMenu menu

        End If
    End If
End Sub


Private Sub gridTransferencias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
If RowIndex > 0 And transferencias.count > 0 Then
    Set TransfBancaria = transferencias.item(RowIndex)
        Values(1) = TransfBancaria.Id
        

        Values(3) = "N° " & TransfBancaria.CuentaBancaria & " | " & TransfBancaria.NombreBanco
        Values(4) = TransfBancaria.FechaOperacion
        Values(5) = TransfBancaria.moneda.NombreCorto
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(TransfBancaria.Monto)), "$", "")
        Values(7) = TransfBancaria.Comprobante
        
        If TransfBancaria.LiquidacionCaja Is Nothing Then
            If TransfBancaria.OrdenPago Is Nothing Then
                    Values(8) = "PCTA: " & TransfBancaria.PagoACuentaID
                    Values(2) = UCase(TransfBancaria.PagoACuentaProveedor)
            Else
                    Values(8) = "OP: " & TransfBancaria.OrdenPago.Id
                    Values(2) = UCase(TransfBancaria.ProveedorRazon)
            End If
        Else
                Values(8) = "LIQ: " & TransfBancaria.LiquidacionCaja.NumeroLiq
                Values(2) = "VARIOS"
        End If



End If
 
End Sub

'Private Sub gridTransferencias_SelectionChange()
'    On Error Resume Next
'    Set TransfBancaria = transferencias.item(gridTransferencias.rowIndex(gridTransferencias.row))
'End Sub



Private Sub gridTransferencias_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridTransferencias, Column
End Sub

Private Sub gridTransferencias_DblClick()
    gridTransferencias_SelectionChange
    mnuVer_Click
End Sub

Private Sub mnuModificar_Click()
    Dim f_ADFE As New frmAdminPagosTransferenciasBancariasEditar
    f_ADFE.idTransfBancaria = TransfBancaria.Id
    f_ADFE.Show
    
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


Private Sub PushButton2_Click(Index As Integer)
    Me.textbMenor.Text = ""
End Sub

Private Sub PushButton1_Click(Index As Integer)
    Me.textbMayor.Text = ""
End Sub
