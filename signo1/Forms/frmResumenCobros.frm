VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmResumenCobros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmen de Cobros por Período"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1530
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   18435
      _Version        =   786432
      _ExtentX        =   32517
      _ExtentY        =   2699
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   450
         Left            =   5160
         TabIndex        =   1
         Top             =   960
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   960
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
         Index           =   0
         Left            =   3000
         TabIndex        =   3
         Top             =   960
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Default         =   -1  'True
         Height          =   450
         Left            =   6840
         TabIndex        =   4
         Top             =   960
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox 
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   2143
         _StockProps     =   79
         BackColor       =   16744576
         Appearance      =   4
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   720
            TabIndex        =   19
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
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   120
            TabIndex        =   22
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
            Index           =   1
            Left            =   165
            TabIndex        =   21
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
            Index           =   1
            Left            =   2400
            TabIndex        =   20
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
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1755
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   3096
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmResumenCobros.frx":0000
      Column(2)       =   "frmResumenCobros.frx":0118
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenCobros.frx":0278
      FormatStyle(2)  =   "frmResumenCobros.frx":03B0
      FormatStyle(3)  =   "frmResumenCobros.frx":0460
      FormatStyle(4)  =   "frmResumenCobros.frx":0514
      FormatStyle(5)  =   "frmResumenCobros.frx":05EC
      FormatStyle(6)  =   "frmResumenCobros.frx":06A4
      ImageCount      =   0
      PrinterProperties=   "frmResumenCobros.frx":0784
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1755
      Left            =   240
      TabIndex        =   6
      Top             =   4335
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   3096
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmResumenCobros.frx":095C
      Column(2)       =   "frmResumenCobros.frx":0A74
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenCobros.frx":0BD4
      FormatStyle(2)  =   "frmResumenCobros.frx":0D0C
      FormatStyle(3)  =   "frmResumenCobros.frx":0DBC
      FormatStyle(4)  =   "frmResumenCobros.frx":0E70
      FormatStyle(5)  =   "frmResumenCobros.frx":0F48
      FormatStyle(6)  =   "frmResumenCobros.frx":1000
      ImageCount      =   0
      PrinterProperties=   "frmResumenCobros.frx":10E0
   End
   Begin GridEX20.GridEX gridRetenciones 
      Height          =   1755
      Left            =   6480
      TabIndex        =   7
      Top             =   2040
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   3096
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmResumenCobros.frx":12B8
      Column(2)       =   "frmResumenCobros.frx":13D0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenCobros.frx":1530
      FormatStyle(2)  =   "frmResumenCobros.frx":1668
      FormatStyle(3)  =   "frmResumenCobros.frx":1718
      FormatStyle(4)  =   "frmResumenCobros.frx":17CC
      FormatStyle(5)  =   "frmResumenCobros.frx":18A4
      FormatStyle(6)  =   "frmResumenCobros.frx":195C
      ImageCount      =   0
      PrinterProperties=   "frmResumenCobros.frx":1A3C
   End
   Begin GridEX20.GridEX gridChequesTerceros 
      Height          =   1755
      Left            =   6480
      TabIndex        =   8
      Top             =   4320
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   3096
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmResumenCobros.frx":1C14
      Column(2)       =   "frmResumenCobros.frx":1D2C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenCobros.frx":1E8C
      FormatStyle(2)  =   "frmResumenCobros.frx":1FC4
      FormatStyle(3)  =   "frmResumenCobros.frx":2074
      FormatStyle(4)  =   "frmResumenCobros.frx":2128
      FormatStyle(5)  =   "frmResumenCobros.frx":2200
      FormatStyle(6)  =   "frmResumenCobros.frx":22B8
      ImageCount      =   0
      PrinterProperties=   "frmResumenCobros.frx":2398
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "* Este total corresponde a la suma de Caja + Banco + Cheques ** / (sin tener en cuenta el total de las Retenciones)."
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   7080
      Width           =   12255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "CAJA"
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   1830
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "BANCO"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   4125
      Width           =   6015
   End
   Begin VB.Label lblTotalCaja 
      Height          =   240
      Left            =   240
      TabIndex        =   15
      Top             =   3825
      Width           =   6015
   End
   Begin VB.Label lblTotalBancos 
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   6135
      Width           =   6015
   End
   Begin VB.Label lblTotalGeneral 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   13
      Top             =   6600
      Width           =   12240
   End
   Begin VB.Label lblTotalRetenciones 
      Height          =   240
      Left            =   6495
      TabIndex        =   12
      Top             =   3810
      Width           =   3735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "RETENCIONES"
      Height          =   240
      Left            =   6480
      TabIndex        =   11
      Tag             =   "RETENCIONES"
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Label lblTotalChequesTerceros 
      Height          =   240
      Left            =   6420
      TabIndex        =   10
      Top             =   6105
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CHEQUES **"
      Height          =   240
      Index           =   1
      Left            =   6480
      TabIndex        =   9
      Top             =   4080
      Width           =   5895
   End
End
Attribute VB_Name = "frmResumenCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dto As DTONombreMonto
Dim condition As String
Dim cheques As New Collection
Dim Cajas As New Collection
Dim compe As New Collection
Dim retenciones As New Collection
Dim cheques3 As New Collection
Dim bancos As New Collection
Private desde






Private Sub cboRangos_Click()
        funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde(0), Me.dtpHasta(0)
End Sub


Private Sub cmdBuscar_Click()
'    Me.gridCheques.ItemCount = 0
    Me.gridCajas.ItemCount = 0
    Me.gridBancos.ItemCount = 0
'    Me.gridCompensatorios.ItemCount = 0
    Me.gridChequesTerceros.ItemCount = 0

    Set cheques = New Collection
    Set Cajas = New Collection
    Set compe = New Collection
    Set bancos = New Collection
    Set cheques3 = New Collection
    Set retenciones = New Collection

    condition = " 1=1 "
    If Not IsNull(Me.dtpDesde(0).value) Then
        condition = condition & " AND rec.fecha >= " & conectar.Escape(Format(Me.dtpDesde(0).value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHasta(0).value) Then
        condition = condition & " AND rec.fecha <= " & conectar.Escape(Format(Me.dtpHasta(0).value, "yyyy-mm-dd"))
    End If

    DAORecibo.ResumenPagos cheques, Cajas, bancos, compe, retenciones, cheques3, condition

'    Me.gridCheques.ItemCount = 0
    Me.gridCajas.ItemCount = 0
    Me.gridBancos.ItemCount = 0
'    Me.gridCompensatorios.ItemCount = 0
    Me.gridRetenciones.ItemCount = 0
    Me.gridChequesTerceros.ItemCount = 0

'    Me.gridCheques.ItemCount = Cheques.count
    Me.gridCajas.ItemCount = Cajas.count
    Me.gridBancos.ItemCount = bancos.count
'    Me.gridCompensatorios.ItemCount = compe.count
    Me.gridRetenciones.ItemCount = retenciones.count
    Me.gridChequesTerceros.ItemCount = cheques3.count

    GridEXHelper.AutoSizeColumns gridCajas
'    GridEXHelper.AutoSizeColumns gridCheques
    GridEXHelper.AutoSizeColumns gridBancos
'    GridEXHelper.AutoSizeColumns Me.gridCompensatorios
    GridEXHelper.AutoSizeColumns Me.gridRetenciones
    GridEXHelper.AutoSizeColumns Me.gridChequesTerceros

    Dim T As Double
    Dim tt As Double
    T = 0
    tt = 0
    For Each dto In cheques
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    tt = tt + T
'    Me.lblTotalCheques.caption = "Total: " & FormatCurrency(funciones.FormatearDecimales(T))

    T = 0
    For Each dto In Cajas
        T = T + funciones.FormatearDecimales(dto.Monto)

    Next
    tt = tt + T
    Me.lblTotalCaja.caption = "Total: " & FormatCurrency(funciones.FormatearDecimales(T))

    T = 0
    For Each dto In bancos
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    tt = tt + T
    Me.lblTotalBancos.caption = "Total: " & FormatCurrency(funciones.FormatearDecimales(T))

    T = 0
    For Each dto In compe
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    'NO SUMA LOS COMPENSATORIOS AL TOTALIZADOR
    'tt = tt + T
'    Me.lblTotalCompe.caption = "Total: " & FormatCurrency(funciones.FormatearDecimales(T))

    T = 0
    For Each dto In retenciones
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    'NO SUMA LAS RETENCIONES AL TOTALIZADOR
    'tt = tt + T
    Me.lblTotalRetenciones.caption = "Total: " & FormatCurrency(funciones.FormatearDecimales(T))

    T = 0
    For Each dto In cheques3
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    tt = tt + T
    Me.lblTotalChequesTerceros.caption = "Total: " & FormatCurrency(funciones.FormatearDecimales(T))


'    Me.lblTotalGeneral.caption = "Total General AR$ " & tt
    Me.lblTotalGeneral.caption = FormatCurrency(funciones.FormatearDecimales(tt))

End Sub

Private Sub cmdImprimir_Click()

    If MsgBox("¿Desea imprimir?", vbYesNo, "Consulta") = vbNo Then Exit Sub


    Dim s As String
    Dim i As Long

    Printer.Print Tab(2);
    i = Printer.FontSize
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.Print "RESUMEN DE COBROS "

    Printer.FontSize = 12
    If Not IsNull(Me.dtpDesde(0).value) Then Printer.Print " Desde: " & Format(Me.dtpDesde(0).value, "dd-mm-yyyy");
    If Not IsNull(Me.dtpHasta(0).value) Then Printer.Print " Hasta: " & Format(Me.dtpHasta(0).value, "dd-mm-yyyy");

    Printer.Print
    Printer.Print

'    dtoHeader Cheques, "CHEQUES *", Me.lblTotalCheques.caption
    dtoHeader cheques3, "CHEQUES **", Me.lblTotalChequesTerceros.caption
    dtoHeader Cajas, "CAJAS", Me.lblTotalCaja.caption
    dtoHeader bancos, "BANCOS", Me.lblTotalBancos.caption
    dtoHeader retenciones, "RETENCIONES", Me.lblTotalRetenciones.caption
'    dtoHeader compe, "COMPENSATORIOS", Me.lblTotalCompe.caption


    Printer.Print
    Printer.FontSize = 16
    Printer.Print Tab(2);
    Printer.FontBold = True
    Printer.Print Me.lblTotalGeneral
    Printer.FontBold = False
    Printer.FontSize = i
    Printer.EndDoc



End Sub

Private Function dtoHeader(col As Collection, titulo As String, total_titulo As String)
    Dim C As DTONombreMonto
    Printer.FontBold = True

    Printer.Print Tab(2);
    Printer.Print titulo;
    Printer.Print Tab(30);
    Printer.Print total_titulo

    Printer.FontBold = False
    For Each C In col
        printDto C
    Next C
    Printer.Print
    Printer.Print
End Function
Private Function printDto(C As DTONombreMonto)
    Dim x As Long
    Dim xval As Long
    Printer.Print Tab(5);
    Printer.Print C.nombre;
    Printer.Print Tab(50);

    x = Printer.CurrentX
    xval = x - Printer.TextWidth(FormatCurrency(funciones.FormatearDecimales(C.Monto)))
    Printer.CurrentX = xval
    Printer.Print FormatCurrency(funciones.FormatearDecimales(C.Monto));




End Function

Private Sub Form_Load()
    Customize Me

'    GridEXHelper.CustomizeGrid Me.gridCheques, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
'    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, False
    GridEXHelper.CustomizeGrid Me.gridChequesTerceros, False, False
    GridEXHelper.CustomizeGrid Me.gridRetenciones, False, False

    Me.gridRetenciones.ItemCount = 0
    Me.gridChequesTerceros.ItemCount = 0
'    Me.gridCheques.ItemCount = 0
    Me.gridCajas.ItemCount = 0
    Me.gridBancos.ItemCount = 0
'    Me.gridCompensatorios.ItemCount = 0
    
        desde = DateSerial(Year(Date), Month(Date), 1)   ' CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    funciones.FillComboBoxDateRanges Me.cboRangos
    
    Dim i As Integer
    
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i


    
    
End Sub

Private Sub Form_Resize()
    Me.GroupBox1.Width = Me.ScaleWidth - 100
End Sub
Private Sub gridBancos_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = bancos(rowIndex)
    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub

Private Sub gridCajas_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = Cajas(rowIndex)
    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub
Private Sub gridCheques_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = cheques(rowIndex)
    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub
Private Sub gridChequesTerceros_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = cheques3(rowIndex)
    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub
Private Sub gridCompensatorios_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = compe(rowIndex)
    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub
Private Sub gridRetenciones_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = retenciones(rowIndex)
    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub

Private Sub PushButton1_Click()

End Sub

