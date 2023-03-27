VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminExtrasReporteCMC 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comparación de Comprobantes desde AFIP"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15660
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13.97
   ScaleMode       =   0  'User
   ScaleWidth      =   27.623
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ProgressBar progreso 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   15135
      _Version        =   786432
      _ExtentX        =   26696
      _ExtentY        =   450
      _StockProps     =   93
      Appearance      =   6
   End
   Begin VB.Frame FramePaso3 
      Caption         =   "Paso 3"
      Enabled         =   0   'False
      Height          =   2535
      Left            =   12480
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      Begin XtremeSuiteControls.PushButton PushButtonRestaurar 
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   1920
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reestablecer"
         BackColor       =   -2147483639
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButtonExportarResultados 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar Resultados"
         BackColor       =   -2147483643
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButtonMostrarResultados 
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2175
         _Version        =   786432
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Mostrar Resultados"
         BackColor       =   -2147483643
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000013&
         X1              =   300
         X2              =   2700
         Y1              =   1750
         Y2              =   1750
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comprobantes encontrados: "
      Height          =   4695
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   15135
      Begin GridEX20.GridEX GridEXComprobantes 
         Height          =   3975
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   7011
         Version         =   "2.0"
         AllowRowSizing  =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   10
         Column(1)       =   "frmAdminExtrasReporteCMC.frx":0000
         Column(2)       =   "frmAdminExtrasReporteCMC.frx":0188
         Column(3)       =   "frmAdminExtrasReporteCMC.frx":0274
         Column(4)       =   "frmAdminExtrasReporteCMC.frx":03CC
         Column(5)       =   "frmAdminExtrasReporteCMC.frx":050C
         Column(6)       =   "frmAdminExtrasReporteCMC.frx":064C
         Column(7)       =   "frmAdminExtrasReporteCMC.frx":0754
         Column(8)       =   "frmAdminExtrasReporteCMC.frx":08AC
         Column(9)       =   "frmAdminExtrasReporteCMC.frx":09F4
         Column(10)      =   "frmAdminExtrasReporteCMC.frx":0B2C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAdminExtrasReporteCMC.frx":0C74
         FormatStyle(2)  =   "frmAdminExtrasReporteCMC.frx":0DAC
         FormatStyle(3)  =   "frmAdminExtrasReporteCMC.frx":0E5C
         FormatStyle(4)  =   "frmAdminExtrasReporteCMC.frx":0F10
         FormatStyle(5)  =   "frmAdminExtrasReporteCMC.frx":0FE8
         FormatStyle(6)  =   "frmAdminExtrasReporteCMC.frx":10A0
         ImageCount      =   0
         PrinterProperties=   "frmAdminExtrasReporteCMC.frx":1180
      End
   End
   Begin VB.Frame FramePaso2 
      Caption         =   "Paso 2"
      Enabled         =   0   'False
      Height          =   2535
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin XtremeSuiteControls.PushButton PushButtonProcesar 
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Procesar"
         BackColor       =   -2147483643
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   2025
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   68
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2025
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   68
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.ComboBox cboRangos 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1380
         Width           =   3975
         _Version        =   786432
         _ExtentX        =   7011
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label LabelHasta 
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   1740
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Hasta"
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.Label LabelDesde 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1740
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Desde"
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.Label LabelRango 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rango"
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   855
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6255
         _Version        =   786432
         _ExtentX        =   11033
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Label2"
         Enabled         =   0   'False
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paso 1"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin XtremeSuiteControls.PushButton PushButtonImportarArchivoAFIP 
         Height          =   495
         Left            =   3600
         TabIndex        =   1
         Top             =   1920
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Importar"
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   2175
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5295
         _Version        =   786432
         _ExtentX        =   9340
         _ExtentY        =   3836
         _StockProps     =   79
         Caption         =   "Label1"
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAdminExtrasReporteCMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim c As New classAdministracion
Dim comprobantes As Collection
Dim tmpComprobantes As ComprobantesRecibidos


Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me

    GridEXHelper.CustomizeGrid Me.GridEXComprobantes, True
    Me.GridEXComprobantes.ItemCount = 0

    Dim i As Integer

    funciones.FillComboBoxDateRanges Me.cboRangos

    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i

    Me.cboRangos.ListIndex = 5

    Me.Label1.caption = "Importar el archivo CSV que se reporta desde el portal de AFIP." & vbCrLf & "" & vbCrLf & "" _
                      & "> AFIP " & vbCrLf & "> Mis Comprobantes " & vbCrLf & "> Recibidos " & vbCrLf & "> Consulta " & vbCrLf & "> Buscar (según rango de fechas) " & vbCrLf & "> Exportar resultados en formato CSV."

    Me.Label2.caption = "Seleccionar el rango de fechas con el cual se va a comparar el reporte importado anteriormente." & vbCrLf & "" _
                      & "Es importante que el rango sea exactamente el mismo para que el resultado de la comparación sea consistente."

    ''Me.caption = caption & " (" & Name & ")"

End Sub


Private Sub PushButtonImportarArchivoAFIP_Click()
    On Error GoTo err4

    Dim filename As String

    Me.CommonDialog1.filter = "CSV File (*.csv)|*.csv"
    Me.CommonDialog1.DialogTitle = "Open CSV"
    Me.CommonDialog1.ShowOpen
    filename = CommonDialog1.filename
    'filename = "\\192.168.0.1\temporal\TEXTOS\1.csv"
    filename = Replace(filename, "\", "/")

    If c.ImportarComprobantesAFIP(filename) Then
        MsgBox "Importación del archivo realizada!", vbInformation, "Información"

        Me.FramePaso2.Enabled = True
        Me.cboRangos.Enabled = True
        Me.dtpDesde.Enabled = True
        Me.dtpHasta.Enabled = True
        Me.LabelDesde.Enabled = True
        Me.LabelHasta.Enabled = True
        Me.LabelRango.Enabled = True
        Me.Label2.Enabled = True
        Me.PushButtonProcesar.Enabled = True

        Me.Frame1.Enabled = False
        Me.Label1.Enabled = False
        Me.PushButtonImportarArchivoAFIP.Enabled = False

    Else
        MsgBox "Error, la importación del archivo no se efectuó!", vbInformation, "Información"

    End If

    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"

End Sub


Private Sub PushButtonProcesar_Click()
    On Error GoTo err4

    Dim condition As String
    condition = " 1 = 1 "

    If Not IsNull(Me.dtpDesde.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        condition = condition & " AND AdminComprasFacturasProveedores.fecha <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    Set comprobantes = DAOFacturaProveedor.FindAll(condition, , "AdminComprasFacturasProveedores.id DESC")

    If DAOFacturaProveedor.CrearTablaTempComprobantes(comprobantes) Then
        MsgBox "Procesamiento completo!", vbInformation, "Información"

        Me.FramePaso3.Enabled = True
        Me.PushButtonExportarResultados.Enabled = True
        Me.PushButtonMostrarResultados.Enabled = True
        Me.PushButtonRestaurar.Enabled = True

        Me.FramePaso2.Enabled = False
        Me.cboRangos.Enabled = False
        Me.dtpDesde.Enabled = False
        Me.dtpHasta.Enabled = False
        Me.LabelDesde.Enabled = False
        Me.LabelHasta.Enabled = False
        Me.LabelRango.Enabled = False
        Me.Label2.Enabled = False
        Me.PushButtonProcesar.Enabled = False
    Else
        MsgBox "Error en SQL!, no se pudo procesar la información.", vbInformation, "Información"
    End If

    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"

End Sub


Private Sub PushButtonMostrarResultados_Click()
    llenar_Grilla

End Sub


Private Sub llenar_Grilla()
    On Error GoTo err4

    GridEXComprobantes.ItemCount = 0

    Set comprobantes = DAOComprobantesRecibidos.FindAll()

    GridEXComprobantes.ItemCount = comprobantes.count

    GridEXHelper.AutoSizeColumns Me.GridEXComprobantes, True

    Me.Frame3.caption = "Comprobantes encontrados: " & comprobantes.count

    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"
End Sub


Private Sub GridEXComprobantes_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)

    If comprobantes.count < 0 Then
        MsgBox "No hay comprobantes que mostrar.", vbCritical, "Error"

    Else
        Set tmpComprobantes = comprobantes.item(RowIndex)

        Values(1) = tmpComprobantes.Fecha_
        Values(2) = tmpComprobantes.Tipo_
        Values(3) = tmpComprobantes.PuntoDeVenta_
        Values(4) = tmpComprobantes.NumeroDesde_
        Values(5) = tmpComprobantes.NroDocEmisor_
        Values(6) = tmpComprobantes.DenominacionEmisor_
        Values(7) = tmpComprobantes.TipoCambio_
        Values(8) = tmpComprobantes.Moneda_
        Values(9) = tmpComprobantes.Iva_
        Values(10) = tmpComprobantes.ImpTotal_

    End If
End Sub


Private Sub PushButtonExportarResultados_Click()
    ExportarResultado

End Sub


Public Function ExportarResultado() As Boolean
    On Error GoTo errEXCEL

    'INICIA EL PROGRESSBAR Y LO MUESTRA
    Me.progreso.Visible = True

    'DEFINE EL VALOR MINIMO Y EL MAXIMO DEL PROGRESSBAR (CANTIDAD DE DATOS EN LA COLECCIÓN COL)
    progreso.min = 0
    progreso.max = comprobantes.count

    'Dim xlApplication As New Excel.Application
    Dim xls As Object
    Set xls = CreateObject("Excel.Application")

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlb As Object
    Set xlb = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xla As Object
    Set xla = CreateObject("Excel.Application")

    Dim A As String
    Dim B As String
    Dim offset As Long
    Dim strMsg As String
    Dim CDLGMAIN As CommonDialog
    Dim sFilter As String

    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add

    xla.Activate

    With xla
        .Range("A1:M1").Font.Bold = True
        .Range("A1:M1").Interior.Color = &HC0C0C0

        'Dim Column As JSColumn

        Dim x As Integer

        .Cells(x + 1, 1).value = "Fecha"
        .Cells(x + 1, 2).value = "Tipo"
        .Cells(x + 1, 3).value = "Punto de Venta"
        .Cells(x + 1, 4).value = "Numero "
        .Cells(x + 1, 5).value = "CUIT"
        .Cells(x + 1, 6).value = "Razón Social"
        .Cells(x + 1, 7).value = "Tipo de Cambio"
        .Cells(x + 1, 8).value = "Moneda"
        .Cells(x + 1, 9).value = "Imp Neto Gravado"
        .Cells(x + 1, 10).value = "Imp Neto No Gravado"
        .Cells(x + 1, 11).value = "Imp Op Exentas"
        .Cells(x + 1, 12).value = "Iva"
        .Cells(x + 1, 13).value = "Imp Total"
        x = 1

        'DEFINE EL CONTADOR DEL PROGRESSBAR Y LO INICIA EN 0
        Dim d As Long
        d = 0

        .Columns("a:m").HorizontalAlignment = xlHAlignCenter

        .Columns("a").ColumnWidth = 15
        .Columns("b").ColumnWidth = 23
        .Columns("c").ColumnWidth = 15
        .Columns("d").ColumnWidth = 15
        .Columns("e").ColumnWidth = 15
        .Columns("f").ColumnWidth = 65
        .Columns("g").ColumnWidth = 15
        .Columns("h").ColumnWidth = 15
        .Columns("i").ColumnWidth = 15
        .Columns("j").ColumnWidth = 15
        .Columns("k").ColumnWidth = 15
        .Columns("l").ColumnWidth = 15
        .Columns("m").ColumnWidth = 15

        For Each tmpComprobantes In comprobantes
            .Cells(x + 1, 1).value = tmpComprobantes.Fecha_
            .Cells(x + 1, 2).value = tmpComprobantes.Tipo_
            .Cells(x + 1, 3).value = tmpComprobantes.PuntoDeVenta_
            .Cells(x + 1, 4).value = tmpComprobantes.NumeroDesde_
            .Cells(x + 1, 5).value = tmpComprobantes.NroDocEmisor_
            .Cells(x + 1, 6).value = tmpComprobantes.DenominacionEmisor_
            .Cells(x + 1, 7).value = tmpComprobantes.TipoCambio_
            .Cells(x + 1, 8).value = tmpComprobantes.Moneda_
            .Cells(x + 1, 9).value = tmpComprobantes.ImpNetoGravado_
            .Cells(x + 1, 10).value = tmpComprobantes.ImpNetoNoGravado_
            .Cells(x + 1, 11).value = tmpComprobantes.ImpOpExentas_
            .Cells(x + 1, 12).value = tmpComprobantes.Iva_
            .Cells(x + 1, 13).value = tmpComprobantes.ImpTotal_

            x = x + 1

            'POR CADA ITERACION SUMA UN VALOR A LA VARIABLE D DEL PROGRESSBAR
            d = d + 1
            progreso.value = d

        Next tmpComprobantes


        A = "m" & x
        offset = x + 2
        B = "m" & offset

        .Range("A1:M1").NumberFormat = "0.00"
        .Range("a1", A).Borders.LineStyle = xlContinuous


        'xls.Visible = True NO MUESTRO LA HOJA XLS
        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"

        Set CDLGMAIN = frmPrincipal.cd

        '    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter

        Dim Periodo As String
        Periodo = 1
        Periodo = Format(dtpDesde, "ddmmyyyy") & "-" & Format(dtpHasta, "ddmmyyyy")

        Dim archi As String
        archi = "COMPARACIÓN CTES COMPRAS " & Periodo & ".xlsx"

        frmPrincipal.cd.CancelError = True

        CDLGMAIN.filename = archi
        CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbInformation + vbOKOnly, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            ExportarResultado = False
        End If
        xlb.Saved = True
        xlb.Close
        xls.Quit

        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        'REINICIA EL PROGRESSBAR Y LO OCULTA
        progreso.value = 0
        Me.progreso.Visible = False

        ExportarResultado = True

    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        ExportarResultado = False
    Else
        ' Resume
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        ExportarResultado = False
    End If
    xlb.Saved = True
    xlb.Close

    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing

End Function



Private Sub PushButtonRestaurar_Click()


    Me.GridEXComprobantes.ItemCount = 0

    Dim i As Integer

    funciones.FillComboBoxDateRanges Me.cboRangos

    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i

    Me.cboRangos.ListIndex = 5

    Me.Frame1.Enabled = True
    Me.Label1.Enabled = True
    Me.PushButtonImportarArchivoAFIP.Enabled = True

    Me.FramePaso3.Enabled = False
    Me.PushButtonExportarResultados.Enabled = False
    Me.PushButtonMostrarResultados.Enabled = False
    Me.PushButtonRestaurar.Enabled = False



End Sub
