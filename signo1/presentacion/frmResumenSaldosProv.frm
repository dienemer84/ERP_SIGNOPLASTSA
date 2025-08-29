VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmResumenSaldosProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmen de Saldos de Proveedores"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   9975
      _Version        =   786432
      _ExtentX        =   17595
      _ExtentY        =   1296
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   480
         Left            =   6120
         TabIndex        =   10
         Top             =   180
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   847
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnExportarXLS 
         Height          =   480
         Left            =   8040
         TabIndex        =   11
         Top             =   180
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   847
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton cmdParar 
      Height          =   420
      Left            =   9480
      TabIndex        =   8
      Top             =   160
      Width           =   525
      _Version        =   786432
      _ExtentX        =   926
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "X"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5445
      Left            =   30
      TabIndex        =   0
      Top             =   645
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   9604
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
      Column(1)       =   "frmResumenSaldosProv.frx":0000
      Column(2)       =   "frmResumenSaldosProv.frx":0120
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenSaldosProv.frx":020C
      FormatStyle(2)  =   "frmResumenSaldosProv.frx":0344
      FormatStyle(3)  =   "frmResumenSaldosProv.frx":03F4
      FormatStyle(4)  =   "frmResumenSaldosProv.frx":04A8
      FormatStyle(5)  =   "frmResumenSaldosProv.frx":0580
      FormatStyle(6)  =   "frmResumenSaldosProv.frx":0638
      ImageCount      =   0
      PrinterProperties=   "frmResumenSaldosProv.frx":0718
   End
   Begin XtremeSuiteControls.PushButton Obtener 
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   1305
      _Version        =   786432
      _ExtentX        =   2302
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Obtener"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   315
      Left            =   2325
      TabIndex        =   4
      Top             =   225
      Width           =   1470
      _Version        =   786432
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   68
      CheckBox        =   -1  'True
      Format          =   1
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   3855
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   4455
      _Version        =   786432
      _ExtentX        =   7858
      _ExtentY        =   529
      _StockProps     =   93
      Appearance      =   6
   End
   Begin VB.Label lblCant 
      Height          =   195
      Left            =   8400
      TabIndex        =   7
      Top             =   280
      Width           =   990
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   7800
      TabIndex        =   6
      Top             =   6360
      Width           =   2205
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   285
      Width           =   420
      _Version        =   786432
      _ExtentX        =   741
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Hasta"
      BackColor       =   12632256
      AutoSize        =   -1  'True
   End
   Begin VB.Label lblproceso 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   7470
   End
End
Attribute VB_Name = "frmResumenSaldosProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dto As DTONombreMonto
Dim col As New Collection
Dim col2 As New Collection
Dim condition As String
Dim enable As Boolean
Public TipoPersonaCta As TipoPersona


Private Sub btnExportarXLS_Click()

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim i As Long
    Dim ultimaFila As Long
    Dim sumaTotal As Double

    ' Crear Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    ' Configurar título
    xlSheet.Range("A1:B1").Merge
    xlSheet.Range("A1:B1").value = "Reporte de Saldos de Clientes al " & Format(Me.dtpHasta.value, "dd/mm/yyyy")
    xlSheet.Range("A1:B1").HorizontalAlignment = xlLeft
    xlSheet.Range("A1:B1").VerticalAlignment = xlCenter
    xlSheet.Range("A1:B1").Font.Bold = True

    ' Escribir encabezados
    xlSheet.Cells(3, 1).value = "Cliente / Proveedor"
    xlSheet.Cells(3, 2).value = "Saldo"
    xlSheet.rows(3).Font.Bold = True
    
    ' Calcular suma total mientras escribimos los datos
    sumaTotal = 0
    For i = 1 To col2.count
        xlSheet.Cells(i + 3, 1).value = col2(i).nombre
        xlSheet.Cells(i + 3, 2).value = funciones.FormatearDecimales(col2(i).Monto)
        sumaTotal = sumaTotal + col2(i).Monto
    Next i
    
    ' Determinar la última fila de datos
    ultimaFila = 3 + col2.count
    
    ' Aplicar formato a TODA la columna B (desde fila 4 hasta el final)
    ultimaFila = 3 + col2.count
    xlSheet.Range("B4:B" & ultimaFila).NumberFormat = "#,##0.00"
    
    ' Agregar fila de totales
    xlSheet.Cells(ultimaFila + 2, 1).value = "TOTAL:"
    xlSheet.Cells(ultimaFila + 2, 2).value = funciones.FormatearDecimales(sumaTotal)
    
    ' Agregar esta línea después de poner el valor del total
    xlSheet.Cells(ultimaFila + 2, 2).NumberFormat = "#,##0.00"
    
    ' Formatear la fila de totales
    With xlSheet.Range("A" & ultimaFila + 1 & ":B" & ultimaFila + 1)
        .Font.Bold = True
    End With
    
    ' Opcional: Agregar línea separadora antes del total
    With xlSheet.Range("A" & ultimaFila & ":B" & ultimaFila)
    End With

    ' Ajustar columnas
    xlSheet.Columns("A:B").AutoFit

    ' Mostrar Excel
    xlApp.Visible = True
    
    ' Liberar objetos
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

End Sub

Private Sub cmdParar_Click()
    enable = False
    cmdParar.Enabled = enable
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, False
    Me.GridEX1.ItemCount = 0




End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = col2(RowIndex)


    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub

Private Sub Obtener_Click()
    enable = True
    cmdParar.Enabled = enable
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount
    Me.lblCant.Visible = True
    Me.lblproceso.Visible = True
    Me.ProgressBar1.Visible = True
    Me.GridEX1.ItemCount = 0
    Dim detalles As Collection
    Set detalles = New Collection
    Set col2 = New Collection
    Dim c As Long
    Dim rs As Recordset

    If TipoPersonaCta = TipoPersona.proveedor_ Then
        Set rs = conectar.RSFactory("SELECT * FROM proveedores  order  by razon asc ")
    Else
        Set rs = conectar.RSFactory("SELECT * FROM clientes  order by razon asc ")
    End If
    c = 0
    While Not rs.EOF And Not rs.BOF
        c = c + 1
        rs.MoveNext
    Wend


    Dim dto As DTONombreMonto

    If c >= 1 Then rs.MoveFirst


    Me.ProgressBar1.max = c
    Dim d As Long
    d = 0
    While Not rs.EOF And Not rs.BOF
        d = d + 1



        If Not IsNull(Me.dtpHasta.value) Then
            condition = conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd"))
        End If

        If TipoPersonaCta = TipoPersona.proveedor_ Then

            Set detalles = DAOCuentaCorriente.FindAllDetallesProveedor(rs!Id, , condition, True, False)

        Else
            If Not IsNull(Me.dtpHasta.value) Then
                condition = Format(Me.dtpHasta.value, "yyyy-mm-dd")
            End If
            Set detalles = DAOCuentaCorriente.FindAllDetalles(rs!Id, , condition)
        End If




        Set dto = New DTONombreMonto
        dto.Monto = DAOCuentaCorriente.GetSaldo(detalles)
        dto.nombre = rs!razon
        If (dto.Monto >= 0.01 Or dto.Monto < -0.01) Then
            col2.Add dto
        End If
        Me.lblCant = CStr(d) & "/" & CStr(c)
        Me.lblproceso = "Procesando " & rs!razon
        Me.ProgressBar1.value = d
        DoEvents
        rs.MoveNext
        Me.GridEX1.ItemCount = col2.count
        If Not enable Then Exit Sub
    Wend

    Me.GridEX1.ItemCount = col2.count
    Me.ProgressBar1.Visible = False
    Dim T As Double

    For Each dto In col2
        T = T + dto.Monto
    Next
    Me.lblTotal = "Total: " & funciones.FormatearDecimales(T)
    Me.lblCant.Visible = False
    tickend = GetTickCount
    'Debug.Print "Tiempo total  ", tickend - tickStart
End Sub

Private Sub PushButton1_Click()


    With Me.GridEX1.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPPortrait
        .HeaderString(jgexHFCenter) = "Resumen de saldos"
        If Not IsNull(dtpHasta.value) Then
            .HeaderString(jgexHFLeft) = "Hasta  " & Format(Me.dtpHasta, "dd-mm-yyyy")
        End If
        .FooterString(jgexHFCenter) = Now
        .FooterString(jgexHFRight) = Me.lblTotal
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.GridEX1.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub
