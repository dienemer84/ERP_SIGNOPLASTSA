VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosTransferenciasBancariasCobro 
   Caption         =   "Transferencias de cobro"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16440
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   16440
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
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
      Begin VB.ComboBox cboCuentaBancaria 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   620
         Width           =   3885
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1335
         Index           =   2
         Left            =   5760
         TabIndex        =   1
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
         Begin VB.TextBox textbMenor 
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox textbMayor 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   4
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
            TabIndex        =   5
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
            Index           =   4
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Mayor que:"
         End
         Begin XtremeSuiteControls.Label Label 
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   6
            Top             =   480
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Menor que:"
         End
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   495
         Left            =   14760
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2400
            TabIndex        =   17
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   165
            TabIndex        =   16
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   195
            Left            =   120
            TabIndex        =   15
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
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1155
         TabIndex        =   18
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
         TabIndex        =   19
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   495
         TabIndex        =   22
         Top             =   300
         Width           =   480
         _Version        =   786432
         _ExtentX        =   847
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cliente"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Index           =   2
         Left            =   15
         TabIndex        =   21
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
   End
   Begin GridEX20.GridEX gridTransferencias 
      Height          =   6495
      Left            =   120
      TabIndex        =   23
      Top             =   2040
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
      ColumnsCount    =   7
      Column(1)       =   "frmAdminPagosTransferenciasBancariasCobro.frx":0000
      Column(2)       =   "frmAdminPagosTransferenciasBancariasCobro.frx":0180
      Column(3)       =   "frmAdminPagosTransferenciasBancariasCobro.frx":02D4
      Column(4)       =   "frmAdminPagosTransferenciasBancariasCobro.frx":0430
      Column(5)       =   "frmAdminPagosTransferenciasBancariasCobro.frx":060C
      Column(6)       =   "frmAdminPagosTransferenciasBancariasCobro.frx":0778
      Column(7)       =   "frmAdminPagosTransferenciasBancariasCobro.frx":093C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminPagosTransferenciasBancariasCobro.frx":0AA8
      FormatStyle(2)  =   "frmAdminPagosTransferenciasBancariasCobro.frx":0BE0
      FormatStyle(3)  =   "frmAdminPagosTransferenciasBancariasCobro.frx":0C90
      FormatStyle(4)  =   "frmAdminPagosTransferenciasBancariasCobro.frx":0D44
      FormatStyle(5)  =   "frmAdminPagosTransferenciasBancariasCobro.frx":0E1C
      FormatStyle(6)  =   "frmAdminPagosTransferenciasBancariasCobro.frx":0ED4
      ImageCount      =   0
      PrinterProperties=   "frmAdminPagosTransferenciasBancariasCobro.frx":0FB4
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   5775
   End
End
Attribute VB_Name = "frmAdminPagosTransferenciasBancariasCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private transferencias As New Collection
Private TransfBancaria As clsTransferenciaBcariaCobro
Private desde
Private colClientes As New Collection
Private colCuentasBancarias As New Collection
Private cli As clsCliente
Private ctabancaria As CuentaBancaria


Private Sub btnExportar_Click()

    If IsSomething(transferencias) Then
        If Not DAOTransferenciaBcariaCobro.ExportarColeccion(transferencias) Then GoTo err1
    End If

    Exit Sub
err1:
    MsgBox "Se produjo un error al exportar!", vbCritical, "Error"

    
End Sub

Private Sub CompletarGridEx()

    Dim condition As String
    condition = " 1 = 1 "

    If Not IsNull(Me.dtpDesde.value) Then
        condition = condition & " AND op.fecha_operacion >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        condition = condition & " AND op.fecha_operacion <= " & conectar.Escape(Me.dtpHasta.value)
    End If
    
    If cboClientes.ListIndex > -1 Then
        condition = condition & " AND (cli.id = " & cboClientes.ItemData(Me.cboClientes.ListIndex) & " OR cli.Id = " & cboClientes.ItemData(Me.cboClientes.ListIndex) & ")"
    End If
    
    If Me.cboCuentaBancaria.ListIndex > -1 Then
        condition = condition & " AND cu.id = " & Me.cboCuentaBancaria.ItemData(Me.cboCuentaBancaria.ListIndex)
    End If
    
    If LenB(Me.textbMayor) > 0 Then
        condition = condition & " AND op.monto >= " & Me.textbMayor.Text
    End If
    
    If LenB(Me.textbMenor) > 0 Then
        condition = condition & " AND op.monto <= " & Me.textbMenor.Text
    End If
    
    Set transferencias = DAOTransferenciaBcariaCobro.FindAll(Banco, condition, "op.id DESC")
    
    Me.gridTransferencias.ItemCount = transferencias.count

    GridEXHelper.AutoSizeColumns Me.gridTransferencias, True
    

    Me.Label3.caption = "Transferencias mostradas [ " & transferencias.count & " ]"
  
End Sub

Public Sub btnTraerDatos_Click(Index As Integer)
    CompletarGridEx

End Sub

Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta

End Sub


Private Sub CMDsINCliente_Click(Index As Integer)
    Me.cboClientes.ListIndex = -1
End Sub


Private Sub CMDsINCtaBancaria_Click(Index As Integer)
    Me.cboCuentaBancaria.ListIndex = -1
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.gridTransferencias, True, True
    
    Me.Height = 9855
    Me.Width = 17595
    
    Dim i As Integer
    
    desde = DateSerial(Year(Date), Month(Date), 1)
    funciones.FillComboBoxDateRanges Me.cboRangos
    
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
    
    Call DAOCliente.llenarComboXtremeSuite(cboClientes)
    Me.cboClientes.ListIndex = -1


    
    DAOCuentaBancaria.LlenarCombo Me.cboCuentaBancaria
    Me.cboCuentaBancaria.ListIndex = -1
    
    Me.gridTransferencias.ItemCount = 0
    
    Me.Label3.caption = "Transferencias mostradas [ 0 ]"
    
    
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


Private Sub gridTransferencias_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If RowIndex > 0 And transferencias.count > 0 Then
        Dim T As clsTransferenciaBcariaCobro
    
        Set T = transferencias.item(RowIndex)
        
        Values(1) = T.Id
        Values(2) = UCase(T.ClienteRazon)
        Values(3) = "N° " & T.CuentaBancaria & " | " & T.NombreBanco
        Values(4) = T.FechaOperacion
        Values(5) = T.moneda.NombreCorto
        Values(6) = Replace(FormatCurrency(funciones.FormatearDecimales(T.Monto)), "$", "")
        
        If Not T.Recibo Is Nothing Then
            Values(7) = T.Recibo.Id
         Else
            Values(7) = "-"
        End If
    End If

End Sub

Private Sub gridTransferencias_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.gridTransferencias, Column
End Sub


Private Sub mnuModificar_Click()
    Dim f_ADFE As New frmAdminPagosTransferenciasBancariasEditar
    f_ADFE.idTransfBancaria = TransfBancaria.Id
    f_ADFE.Show
    
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


Private Function BuscarTransferenciaSeleccionada() As clsTransferenciaBcariaCobro
    Dim idTransf As Long
    Dim T As clsTransferenciaBcariaCobro
    
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


