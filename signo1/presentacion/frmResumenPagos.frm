VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmResumenPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmen de Pagos por Período"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX gridCheques 
      Height          =   1695
      Left            =   90
      TabIndex        =   6
      Top             =   1665
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2990
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
      Column(1)       =   "frmResumenPagos.frx":0000
      Column(2)       =   "frmResumenPagos.frx":0118
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenPagos.frx":0278
      FormatStyle(2)  =   "frmResumenPagos.frx":03B0
      FormatStyle(3)  =   "frmResumenPagos.frx":0460
      FormatStyle(4)  =   "frmResumenPagos.frx":0514
      FormatStyle(5)  =   "frmResumenPagos.frx":05EC
      FormatStyle(6)  =   "frmResumenPagos.frx":06A4
      ImageCount      =   0
      PrinterProperties=   "frmResumenPagos.frx":0784
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1290
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   9675
      _Version        =   786432
      _ExtentX        =   17066
      _ExtentY        =   2275
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   450
         Left            =   5130
         TabIndex        =   1
         Top             =   600
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
         Left            =   960
         TabIndex        =   2
         Top             =   690
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
         Left            =   3150
         TabIndex        =   3
         Top             =   690
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
         Left            =   6555
         TabIndex        =   25
         Top             =   585
         Width           =   1350
         _Version        =   786432
         _ExtentX        =   2381
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   2685
         TabIndex        =   5
         Top             =   750
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
         Left            =   450
         TabIndex        =   4
         Top             =   750
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
   End
   Begin GridEX20.GridEX gridCajas 
      Height          =   1755
      Left            =   3375
      TabIndex        =   8
      Top             =   1650
      Width           =   3135
      _ExtentX        =   5530
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
      Column(1)       =   "frmResumenPagos.frx":095C
      Column(2)       =   "frmResumenPagos.frx":0A74
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenPagos.frx":0BD4
      FormatStyle(2)  =   "frmResumenPagos.frx":0D0C
      FormatStyle(3)  =   "frmResumenPagos.frx":0DBC
      FormatStyle(4)  =   "frmResumenPagos.frx":0E70
      FormatStyle(5)  =   "frmResumenPagos.frx":0F48
      FormatStyle(6)  =   "frmResumenPagos.frx":1000
      ImageCount      =   0
      PrinterProperties=   "frmResumenPagos.frx":10E0
   End
   Begin GridEX20.GridEX gridBancos 
      Height          =   1755
      Left            =   3390
      TabIndex        =   10
      Top             =   3945
      Width           =   3135
      _ExtentX        =   5530
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
      Column(1)       =   "frmResumenPagos.frx":12B8
      Column(2)       =   "frmResumenPagos.frx":13D0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenPagos.frx":1530
      FormatStyle(2)  =   "frmResumenPagos.frx":1668
      FormatStyle(3)  =   "frmResumenPagos.frx":1718
      FormatStyle(4)  =   "frmResumenPagos.frx":17CC
      FormatStyle(5)  =   "frmResumenPagos.frx":18A4
      FormatStyle(6)  =   "frmResumenPagos.frx":195C
      ImageCount      =   0
      PrinterProperties=   "frmResumenPagos.frx":1A3C
   End
   Begin GridEX20.GridEX gridCompensatorios 
      Height          =   1755
      Left            =   105
      TabIndex        =   12
      Top             =   3960
      Width           =   3135
      _ExtentX        =   5530
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
      Column(1)       =   "frmResumenPagos.frx":1C14
      Column(2)       =   "frmResumenPagos.frx":1D2C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenPagos.frx":1E8C
      FormatStyle(2)  =   "frmResumenPagos.frx":1FC4
      FormatStyle(3)  =   "frmResumenPagos.frx":2074
      FormatStyle(4)  =   "frmResumenPagos.frx":2128
      FormatStyle(5)  =   "frmResumenPagos.frx":2200
      FormatStyle(6)  =   "frmResumenPagos.frx":22B8
      ImageCount      =   0
      PrinterProperties=   "frmResumenPagos.frx":2398
   End
   Begin GridEX20.GridEX gridRetenciones 
      Height          =   1755
      Left            =   6660
      TabIndex        =   19
      Top             =   1635
      Width           =   3135
      _ExtentX        =   5530
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
      Column(1)       =   "frmResumenPagos.frx":2570
      Column(2)       =   "frmResumenPagos.frx":2688
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenPagos.frx":27E8
      FormatStyle(2)  =   "frmResumenPagos.frx":2920
      FormatStyle(3)  =   "frmResumenPagos.frx":29D0
      FormatStyle(4)  =   "frmResumenPagos.frx":2A84
      FormatStyle(5)  =   "frmResumenPagos.frx":2B5C
      FormatStyle(6)  =   "frmResumenPagos.frx":2C14
      ImageCount      =   0
      PrinterProperties=   "frmResumenPagos.frx":2CF4
   End
   Begin GridEX20.GridEX gridChequesTerceros 
      Height          =   1755
      Left            =   6690
      TabIndex        =   22
      Top             =   3915
      Width           =   3135
      _ExtentX        =   5530
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
      Column(1)       =   "frmResumenPagos.frx":2ECC
      Column(2)       =   "frmResumenPagos.frx":2FE4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenPagos.frx":3144
      FormatStyle(2)  =   "frmResumenPagos.frx":327C
      FormatStyle(3)  =   "frmResumenPagos.frx":332C
      FormatStyle(4)  =   "frmResumenPagos.frx":33E0
      FormatStyle(5)  =   "frmResumenPagos.frx":34B8
      FormatStyle(6)  =   "frmResumenPagos.frx":3570
      ImageCount      =   0
      PrinterProperties=   "frmResumenPagos.frx":3650
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CHEQUES DE TERCEROS"
      Height          =   240
      Index           =   1
      Left            =   6660
      TabIndex        =   24
      Top             =   3705
      Width           =   3135
   End
   Begin VB.Label lblTotalChequesTerceros 
      Height          =   240
      Left            =   6705
      TabIndex        =   23
      Top             =   5715
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "RETENCIONES"
      Height          =   240
      Left            =   6600
      TabIndex        =   21
      Tag             =   "RETENCIONES"
      Top             =   1410
      Width           =   3135
   End
   Begin VB.Label lblTotalRetenciones 
      Height          =   240
      Left            =   6660
      TabIndex        =   20
      Top             =   3420
      Width           =   3135
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
      Height          =   330
      Left            =   105
      TabIndex        =   18
      Top             =   6255
      Width           =   9720
   End
   Begin VB.Label lblTotalBancos 
      Height          =   240
      Left            =   3435
      TabIndex        =   17
      Top             =   5745
      Width           =   3135
   End
   Begin VB.Label lblTotalCompe 
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label lblTotalCaja 
      Height          =   240
      Left            =   3375
      TabIndex        =   15
      Top             =   3435
      Width           =   3135
   End
   Begin VB.Label lblTotalCheques 
      Height          =   240
      Left            =   75
      TabIndex        =   14
      Top             =   3405
      Width           =   3165
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "COMPENSATORIOS"
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   3765
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "BANCO"
      Height          =   240
      Index           =   0
      Left            =   3360
      TabIndex        =   11
      Top             =   3735
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "CAJA"
      Height          =   240
      Left            =   3330
      TabIndex        =   9
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CHEQUES"
      Height          =   240
      Left            =   45
      TabIndex        =   7
      Top             =   1455
      Width           =   3180
   End
End
Attribute VB_Name = "frmResumenPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dto As DTONombreMonto
Dim condition As String
Dim Cheques As New Collection
Dim Cajas As New Collection
Dim compe As New Collection
Dim retenciones As New Collection
Dim cheques3 As New Collection
Dim bancos As New Collection


Private Sub cmdBuscar_Click()
    Me.gridCheques.ItemCount = 0
    Me.gridCajas.ItemCount = 0
    Me.gridBancos.ItemCount = 0
    Me.gridCompensatorios.ItemCount = 0

    Set Cheques = New Collection
    Set Cajas = New Collection
    Set compe = New Collection
    Set bancos = New Collection
    Set cheques3 = New Collection
    Set retenciones = New Collection

    condition = " 1=1 "
    If Not IsNull(Me.dtpDesde.value) Then
        condition = condition & " AND op.fecha >= " & conectar.Escape(Format(Me.dtpDesde.value, "yyyy-mm-dd"))
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        condition = condition & " AND op.fecha <= " & conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd"))
    End If



    DAOOrdenPago.ResumenPagos Cheques, Cajas, bancos, compe, retenciones, cheques3, condition

    Me.gridCheques.ItemCount = 0
    Me.gridCajas.ItemCount = 0
    Me.gridBancos.ItemCount = 0
    Me.gridCompensatorios.ItemCount = 0
    Me.gridRetenciones.ItemCount = 0
    Me.gridChequesTerceros.ItemCount = 0

    Me.gridCheques.ItemCount = Cheques.count
    Me.gridCajas.ItemCount = Cajas.count
    Me.gridBancos.ItemCount = bancos.count
    Me.gridCompensatorios.ItemCount = compe.count
    Me.gridRetenciones.ItemCount = retenciones.count
    Me.gridChequesTerceros.ItemCount = cheques3.count

    GridEXHelper.AutoSizeColumns gridCajas
    GridEXHelper.AutoSizeColumns gridCheques
    GridEXHelper.AutoSizeColumns gridBancos
    GridEXHelper.AutoSizeColumns Me.gridCompensatorios
    GridEXHelper.AutoSizeColumns Me.gridRetenciones
    GridEXHelper.AutoSizeColumns Me.gridChequesTerceros

    Dim T As Double
    Dim tt As Double
    T = 0
    tt = 0
    For Each dto In Cheques
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    tt = tt + T
    Me.lblTotalCheques.caption = "Total AR$ " & funciones.FormatearDecimales(T)

    T = 0
    For Each dto In Cajas
        T = T + funciones.FormatearDecimales(dto.Monto)

    Next
    tt = tt + T
    Me.lblTotalCaja.caption = "Total AR$ " & funciones.FormatearDecimales(T)

    T = 0
    For Each dto In bancos
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    tt = tt + T
    Me.lblTotalBancos.caption = "Total AR$ " & funciones.FormatearDecimales(T)

    T = 0
    For Each dto In compe
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    'tt = tt + T
    Me.lblTotalCompe.caption = "Total AR$ " & funciones.FormatearDecimales(T)

    T = 0
    For Each dto In retenciones
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    tt = tt + T
    Me.lblTotalRetenciones.caption = "Total AR$ " & funciones.FormatearDecimales(T)

    T = 0
    For Each dto In cheques3
        T = T + funciones.FormatearDecimales(dto.Monto)
    Next
    tt = tt + T
    Me.lblTotalChequesTerceros.caption = "Total AR$ " & funciones.FormatearDecimales(T)


    Me.lblTotalGeneral.caption = "Total General AR$ " & tt


End Sub

Private Sub cmdImprimir_Click()

    If MsgBox("¿Desea imprimir?", vbYesNo, "Consulta") = vbNo Then Exit Sub


    Dim s As String
    Dim i As Long

    Printer.Print Tab(2);
    i = Printer.FontSize
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.Print "RESUMEN DE PAGOS "

    Printer.FontSize = 12
    If Not IsNull(Me.dtpDesde.value) Then Printer.Print " Desde: " & Format(Me.dtpDesde.value, "dd-mm-yyyy");
    If Not IsNull(Me.dtpHasta.value) Then Printer.Print " Hasta: " & Format(Me.dtpHasta.value, "dd-mm-yyyy");

    Printer.Print
    Printer.Print

    dtoHeader Cheques, "CHEQUES PROPIOS", Me.lblTotalCheques.caption
    dtoHeader cheques3, "CHEQUES DE TERCEROS", Me.lblTotalChequesTerceros.caption
    dtoHeader Cajas, "CAJAS", Me.lblTotalCaja.caption
    dtoHeader bancos, "BANCOS", Me.lblTotalBancos.caption
    dtoHeader retenciones, "RETENCIONES", Me.lblTotalRetenciones.caption
    dtoHeader compe, "COMPENSATORIOS", Me.lblTotalCompe.caption


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
    Dim c As DTONombreMonto
    Printer.FontBold = True

    Printer.Print Tab(2);
    Printer.Print titulo;
    Printer.Print Tab(30);
    Printer.Print total_titulo

    Printer.FontBold = False
    For Each c In col
        printDto c
    Next c
    Printer.Print
    Printer.Print
End Function
Private Function printDto(c As DTONombreMonto)
    Dim x As Long
    Dim xval As Long
    Printer.Print Tab(5);
    Printer.Print c.nombre;
    Printer.Print Tab(50);

    x = Printer.CurrentX
    xval = x - Printer.TextWidth(funciones.FormatearDecimales(c.Monto))
    Printer.CurrentX = xval
    Printer.Print funciones.FormatearDecimales(c.Monto);




End Function
Private Sub Form_Load()
    Customize Me

    GridEXHelper.CustomizeGrid Me.gridCheques, False, False
    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, False
    GridEXHelper.CustomizeGrid Me.gridChequesTerceros, False, False
    GridEXHelper.CustomizeGrid Me.gridRetenciones, False, False

    Me.gridRetenciones.ItemCount = 0
    Me.gridChequesTerceros.ItemCount = 0
    Me.gridCheques.ItemCount = 0
    Me.gridCajas.ItemCount = 0
    Me.gridBancos.ItemCount = 0
    Me.gridCompensatorios.ItemCount = 0
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
    Set dto = Cheques(rowIndex)
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

