VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmResumenSaldosProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmen de Saldos de Proveedores"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   9855
      _Version        =   786432
      _ExtentX        =   17383
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Período de movimientos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   60
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
         CurrentDate     =   46282.6832175926
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   60
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
         CurrentDate     =   46282.6835763889
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   0
         Left            =   5520
         TabIndex        =   12
         Top             =   150
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblDesde 
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   11
         Top             =   150
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   7440
      Width           =   9975
      _Version        =   786432
      _ExtentX        =   17595
      _ExtentY        =   1296
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   480
         Left            =   6120
         TabIndex        =   8
         Top             =   240
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
         TabIndex        =   9
         Top             =   240
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
      TabIndex        =   6
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
      Top             =   1440
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
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   130
      Width           =   1545
      _Version        =   786432
      _ExtentX        =   2725
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Obtener"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   2775
      TabIndex        =   3
      Top             =   220
      Width           =   6480
      _Version        =   786432
      _ExtentX        =   11430
      _ExtentY        =   529
      _StockProps     =   93
      Appearance      =   6
   End
   Begin VB.Label lblCant 
      Height          =   195
      Left            =   8400
      TabIndex        =   5
      Top             =   280
      Width           =   990
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   7800
      TabIndex        =   4
      Top             =   7080
      Width           =   2205
   End
   Begin VB.Label lblproceso 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   7080
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

    Dim xlApp As Object, xlBook As Object, xlsheet As Object
    Dim i As Long
    Dim ultimaFila As Long
    Dim sumaTotal As Double

    ' Crear Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlsheet = xlBook.Worksheets(1)
    
    ' Configurar título
    xlsheet.Range("A1:B1").Merge
    xlsheet.Range("A1:B1").value = "Reporte de Saldos de Proveedores al " & Format(Me.dtpHasta.value, "dd/mm/yyyy")
    xlsheet.Range("A1:B1").HorizontalAlignment = xlLeft
    xlsheet.Range("A1:B1").VerticalAlignment = xlCenter
    xlsheet.Range("A1:B1").Font.Bold = True

    ' Escribir encabezados
    xlsheet.Cells(3, 1).value = "Cliente / Proveedor"
    xlsheet.Cells(3, 2).value = "Saldo"
    xlsheet.rows(3).Font.Bold = True
    
    ' Calcular suma total mientras escribimos los datos
    sumaTotal = 0
    For i = 1 To col2.count
        xlsheet.Cells(i + 3, 1).value = col2(i).nombre
        xlsheet.Cells(i + 3, 2).value = funciones.FormatearDecimales(col2(i).Monto)
        sumaTotal = sumaTotal + col2(i).Monto
    Next i
    
    ' Determinar la última fila de datos
    ultimaFila = 3 + col2.count
    
    ' Aplicar formato a TODA la columna B (desde fila 4 hasta el final)
    ultimaFila = 3 + col2.count
    xlsheet.Range("B4:B" & ultimaFila).NumberFormat = "#,##0.00"
    
    ' Agregar fila de totales
    xlsheet.Cells(ultimaFila + 2, 1).value = "TOTAL:"
    xlsheet.Cells(ultimaFila + 2, 2).value = funciones.FormatearDecimales(sumaTotal)
    
    ' Agregar esta línea después de poner el valor del total
    xlsheet.Cells(ultimaFila + 2, 2).NumberFormat = "#,##0.00"
    
    ' Formatear la fila de totales
    With xlsheet.Range("A" & ultimaFila + 1 & ":B" & ultimaFila + 1)
        .Font.Bold = True
    End With
    
    ' Opcional: Agregar línea separadora antes del total
    With xlsheet.Range("A" & ultimaFila & ":B" & ultimaFila)
    End With

    ' Ajustar columnas
    xlsheet.Columns("A:B").AutoFit

    ' Mostrar Excel
    xlApp.Visible = True
    
    ' Liberar objetos
    Set xlsheet = Nothing
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

    'La fecha Desde queda desactivada inicialmente.
    Me.dtpDesde.value = Null

    'La fecha Hasta queda seleccionada con la fecha actual.
    If IsNull(Me.dtpHasta.value) Then
        Me.dtpHasta.value = Date
    End If

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

    On Error GoTo ErrorHandler

    Dim tickStart As Double
    Dim tickend As Double

    Dim Detalles As Collection
    Dim rs As ADODB.Recordset

    Dim itemResumen As DTONombreMonto

    Dim fechaDesdeResumen As String
    Dim fechaHastaResumen As String
    
    Dim tipoResultado As String
    Dim mensajeError As String

    Dim c As Long
    Dim d As Long
    Dim totalResumen As Double

    tickStart = GetTickCount

    enable = True
    
    condition = vbNullString
    
    fechaDesdeResumen = vbNullString
    fechaHastaResumen = vbNullString
    
    Me.cmdParar.Enabled = True
    Me.lblCant.Visible = True
    Me.lblproceso.Visible = True
    Me.ProgressBar1.Visible = True

    Me.lblCant = vbNullString
    Me.lblproceso = "Preparando reporte..."
    Me.lblTotal = vbNullString

    Me.GridEX1.ItemCount = 0

    Set col2 = New Collection

    '----------------------------------------------------------
    ' OBTENER FECHAS DEL BALANCE
    '----------------------------------------------------------
    
    If Not IsNull(Me.dtpDesde.value) Then
        fechaDesdeResumen = Format$(Me.dtpDesde.value, "yyyy-mm-dd")
    End If
    
    If Not IsNull(Me.dtpHasta.value) Then
        fechaHastaResumen = Format$(Me.dtpHasta.value, "yyyy-mm-dd")
    
        'Se mantiene para el procedimiento anterior de clientes.
        condition = fechaHastaResumen
    End If
    
    'Si se seleccionó Desde, también debe existir Hasta.
    If LenB(fechaDesdeResumen) > 0 And LenB(fechaHastaResumen) = 0 Then
        MsgBox "Debe seleccionar una fecha Hasta.", _
               vbExclamation, _
               "Resumen de saldos"
    GoTo SalirPorValidacion
    End If
    
    'Validar el orden de las fechas.
    If LenB(fechaDesdeResumen) > 0 And LenB(fechaHastaResumen) > 0 Then
    
        If CDate(Me.dtpDesde.value) > CDate(Me.dtpHasta.value) Then
            MsgBox "La fecha Desde no puede ser posterior a la fecha Hasta.", _
                   vbExclamation, _
                   "Resumen de saldos"
    GoTo SalirPorValidacion
        End If
    
    End If

    '=========================================================
    ' REPORTE RÁPIDO DE PROVEEDORES
    '=========================================================

    If TipoPersonaCta = TipoPersona.proveedor_ Then

        tipoResultado = "proveedores"

        Me.lblproceso = _
            "Calculando saldos de proveedores..."

        Me.lblCant = vbNullString
        Me.ProgressBar1.Visible = False
        Me.cmdParar.Enabled = False

        Screen.MousePointer = vbHourglass

        DoEvents

    Set col2 = _
        DAOCuentaCorriente.FindResumenSaldosProveedoresRapido( _
            fechaHastaResumen, _
            fechaDesdeResumen)

    Else

        '=====================================================
        ' REPORTE DE CLIENTES
        ' Se conserva el funcionamiento anterior.
        '=====================================================

        tipoResultado = "clientes"

        Set rs = conectar.RSFactory( _
            "SELECT id, razon " & _
            "FROM clientes " & _
            "ORDER BY razon ASC")

        c = 0

        While Not rs.EOF And Not rs.BOF

            c = c + 1
            rs.MoveNext

        Wend

        If c > 0 Then

            rs.MoveFirst
            Me.ProgressBar1.max = c
            Me.ProgressBar1.value = 0

        End If

        d = 0

        While Not rs.EOF And Not rs.BOF

            DoEvents

            If Not enable Then
                GoTo ProcesoCancelado
            End If

            d = d + 1

            Me.lblCant = CStr(d) & "/" & CStr(c)

            Me.lblproceso = _
                "Procesando " & CStr(rs!razon)

            Set Detalles = _
                DAOCuentaCorriente.FindAllDetalles( _
                    CLng(rs!Id), False, condition)

            Set itemResumen = New DTONombreMonto

            itemResumen.Monto = _
                DAOCuentaCorriente.GetSaldo(Detalles)

            itemResumen.nombre = CStr(rs!razon)

            If itemResumen.Monto >= 0.01 Or _
               itemResumen.Monto < -0.01 Then

                col2.Add itemResumen

            End If

            Me.ProgressBar1.value = d

            rs.MoveNext

        Wend

    End If

    '=========================================================
    ' MOSTRAR RESULTADOS
    '=========================================================

    Me.GridEX1.ItemCount = col2.count
    Me.GridEX1.Refresh

    totalResumen = 0

    For Each itemResumen In col2

        totalResumen = _
            totalResumen + itemResumen.Monto

    Next itemResumen

    Me.lblTotal = _
        "Total: " & _
        funciones.FormatearDecimales(totalResumen)

    Me.lblproceso = _
        "Proceso finalizado: " & _
        CStr(col2.count) & " " & _
        tipoResultado & " con saldo."

    Me.lblCant.Visible = False
    Me.ProgressBar1.Visible = False
    Me.cmdParar.Enabled = False

    Screen.MousePointer = vbDefault

    tickend = GetTickCount

    Debug.Print _
        "Tiempo total frmResumenSaldosProv: " & _
        Format$((tickend - tickStart) / 1000, "0.00") & _
        " segundos"

    Set Detalles = Nothing
    Set rs = Nothing


SalirPorValidacion:

    Me.lblCant.Visible = True
    Me.ProgressBar1.Visible = True
    Me.cmdParar.Enabled = True

    Screen.MousePointer = vbDefault

    Me.lblproceso = "Revise el período seleccionado."

    Set Detalles = Nothing
    Set rs = Nothing

    Exit Sub


    Exit Sub

ProcesoCancelado:

    Me.lblproceso = "Proceso cancelado por el usuario."
    Me.lblCant.Visible = False
    Me.ProgressBar1.Visible = False
    Me.cmdParar.Enabled = False

    Screen.MousePointer = vbDefault

    Set Detalles = Nothing
    Set rs = Nothing

    Exit Sub

ErrorHandler:

    mensajeError = Err.Description

    Screen.MousePointer = vbDefault

    Me.lblCant.Visible = False
    Me.ProgressBar1.Visible = False
    Me.cmdParar.Enabled = False

    Me.lblproceso = "No se pudo generar el reporte."

    MsgBox "No se pudo generar el resumen de saldos." & _
           vbCrLf & vbCrLf & _
           mensajeError, _
           vbCritical, _
           "Resumen de saldos"

    Set Detalles = Nothing
    Set rs = Nothing

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
