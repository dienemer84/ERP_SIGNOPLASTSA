VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmCtaCte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCtaCte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   9975
   Begin VB.TextBox txtImporteMinimo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      Text            =   "0"
      Top             =   525
      Width           =   1695
   End
   Begin XtremeSuiteControls.PushButton button_ExportToXls 
      Height          =   435
      Left            =   2040
      TabIndex        =   8
      Top             =   6240
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Exportar a XLS"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   1680
      _Version        =   786432
      _ExtentX        =   2963
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboClientes 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   75
      Width           =   6840
      _Version        =   786432
      _ExtentX        =   12065
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin GridEX20.GridEX gridDetalles 
      Height          =   5040
      Left            =   90
      TabIndex        =   1
      Top             =   1005
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   8890
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmCtaCte.frx":000C
      Column(2)       =   "frmCtaCte.frx":01B0
      Column(3)       =   "frmCtaCte.frx":0300
      Column(4)       =   "frmCtaCte.frx":04A4
      Column(5)       =   "frmCtaCte.frx":0648
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCtaCte.frx":07EC
      FormatStyle(2)  =   "frmCtaCte.frx":0914
      FormatStyle(3)  =   "frmCtaCte.frx":09C4
      FormatStyle(4)  =   "frmCtaCte.frx":0A78
      FormatStyle(5)  =   "frmCtaCte.frx":0B50
      FormatStyle(6)  =   "frmCtaCte.frx":0C08
      FormatStyle(7)  =   "frmCtaCte.frx":0CE8
      FormatStyle(8)  =   "frmCtaCte.frx":0DA0
      ImageCount      =   0
      PrinterProperties=   "frmCtaCte.frx":0E34
   End
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   510
      Width           =   1470
      _Version        =   786432
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   68
      CheckBox        =   -1  'True
      Format          =   1
   End
   Begin XtremeSuiteControls.PushButton cmdVerCtaCte 
      Default         =   -1  'True
      Height          =   420
      Left            =   8400
      TabIndex        =   4
      Top             =   480
      Width           =   1440
      _Version        =   786432
      _ExtentX        =   2540
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Ver"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblImporteMinimo 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   480
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Importe mínimo:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   6360
      Width           =   5625
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   570
      Width           =   540
      _Version        =   786432
      _ExtentX        =   953
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Hasta:"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   630
      _Version        =   786432
      _ExtentX        =   1111
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Cliente:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      AutoSize        =   -1  'True
   End
End
Attribute VB_Name = "frmCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Detalles As Collection
Private deta As DTODetalleCuentaCorriente
Private saldo As Double
Private saldos As New Dictionary


Private Sub button_ExportToXls_Click()

    ExportToXls


End Sub

Private Sub cmdVerCtaCte_Click()

    If Me.cboClientes.ListIndex = -1 Then
        MsgBox "Debe seleccionar un cliente.", _
               vbExclamation + vbOKOnly, _
               "Cuenta corriente"
        Exit Sub
    End If

    Dim importeMinimo As Double

    If Not ObtenerImporteMinimo(importeMinimo) Then
        Exit Sub
    End If

    Dim fecha_hasta As String

    If Not IsNull(Me.dtpHasta.value) Then
        fecha_hasta = Format$(Me.dtpHasta.value, "yyyy-mm-dd")
    End If

    Dim detallesCompletos As Collection

    Set detallesCompletos = _
        DAOCuentaCorriente.FindAllDetalles( _
            Me.cboClientes.ItemData(Me.cboClientes.ListIndex), _
            , _
            fecha_hasta)

    Set Detalles = FiltrarDetallesPorImporte( _
        detallesCompletos, _
        importeMinimo)

    If IsSomething(detallesCompletos) Then

        saldo = DAOCuentaCorriente.GetSaldo(detallesCompletos)

        If importeMinimo > 0 Then
            Me.lblSaldo.caption = _
                "Saldo real: " & _
                Replace(FormatCurrency( _
                    funciones.FormatearDecimales(saldo)), "$", "")
        Else
            Me.lblSaldo.caption = _
                "Saldo: " & _
                Replace(FormatCurrency( _
                    funciones.FormatearDecimales(saldo)), "$", "")
        End If

    Else
        saldo = 0
        Me.lblSaldo.caption = "Saldo: 0,00"
    End If

    Set saldos = New Dictionary

    Me.gridDetalles.ItemCount = 0
    Me.gridDetalles.ItemCount = Detalles.count

End Sub



Private Sub Form_Load()

    Customize Me
    GridEXHelper.CustomizeGrid Me.gridDetalles

    DAOCliente.llenarComboXtremeSuite Me.cboClientes
    Me.cboClientes.ListIndex = -1

    Set Detalles = New Collection

    Me.gridDetalles.ItemCount = 0

End Sub


Private Sub gridDetalles_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
'GridEXHelper.ColumnHeaderClick Me.gridDetalles, Column
End Sub

Private Sub gridDetalles_DblClick()
    Set deta = Detalles.item(Me.gridDetalles.RowIndex(Me.gridDetalles.row))

    If (deta.tipoComprobante = TipoComprobanteUsado.Factura_) Then
        Dim frm As New frmAdminFacturasEdicion
        frm.idFactura = deta.IdComprobante
        frm.ReadOnly = True
        frm.Show
    End If

    If (deta.tipoComprobante = TipoComprobanteUsado.Recibo_ Or deta.tipoComprobante = TipoComprobanteUsado.Retencion_) Then
        Dim frm1 As New frmAdminCobranzasNuevoRecibo
        frm1.editar = False
        frm1.ReciboID = deta.IdComprobante
        frm1.Show
    End If


End Sub

Private Sub gridDetalles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = 2 Then    'CTRL + C
        GridEXHelper.Grid2Clipboard Me.gridDetalles

        DoEvents
        MsgBox "La lista ha sido copiada al portapapeles.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub gridDetalles_RowFormat(RowBuffer As GridEX20.JSRowData)

    Set deta = Detalles.item(RowBuffer.RowIndex)
    If deta.AtributoExtra Then
        RowBuffer.RowStyle = "saldado"
    End If
End Sub

Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Detalles.count > 0 Then
        Set deta = Detalles.item(RowIndex)
        Values(1) = deta.FEcha
        Values(2) = deta.Comprobante
        'Values(3) = deta.Debe
        'Values(4) = deta.Haber

        Values(3) = Replace(FormatCurrency(funciones.FormatearDecimales(deta.Debe)), "$", "")
        Values(4) = Replace(FormatCurrency(funciones.FormatearDecimales(deta.Haber)), "$", "")

        '   If saldos.Exists(CStr(RowIndex)) Then
        '        Values(5) = saldos.item(CStr(RowIndex))
        '   Else
        '        saldo = saldo + deta.Debe - deta.Haber
        '     saldos.Add CStr(RowIndex), saldo
        '        Values(5) = funciones.RedondearDecimales(saldo)
        '  End If

        'Values(5) = deta.saldo
        Values(5) = Replace(FormatCurrency(funciones.FormatearDecimales(deta.saldo)), "$", "")

    End If
End Sub

Private Sub PushButton1_Click()

    With Me.gridDetalles.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPPortrait
        .HeaderString(jgexHFCenter) = "Cuenta Corriente de " & Me.cboClientes.Text
        If Not IsNull(dtpHasta.value) Then
            .HeaderString(jgexHFLeft) = "Hasta  " & Format(Me.dtpHasta, "dd-mm-yyyy")
        End If
        .FooterString(jgexHFCenter) = Now
        .FooterString(jgexHFRight) = Me.lblSaldo
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.gridDetalles.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub


Public Function ExportToXls() As Boolean

    On Error GoTo ManejarError

    ExportToXls = False

    'Primero se controla que la colección exista
    If Not IsSomething(Detalles) Then

        MsgBox "No hay datos para exportar." & vbCrLf & _
               "Seleccione un cliente y presione Ver.", _
               vbInformation + vbOKOnly, _
               "Exportar cuenta corriente"

        Exit Function

    End If

    'Después se controla que tenga elementos
    If Detalles.count = 0 Then

        MsgBox "No hay movimientos para exportar." & vbCrLf & _
               "Seleccione otro cliente o reduzca el importe mínimo.", _
               vbInformation + vbOKOnly, _
               "Exportar cuenta corriente"

        Exit Function

    End If

    Dim xlApplication As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object

    Set xlApplication = CreateObject("Excel.Application")
    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlApplication.DisplayAlerts = False
    xlApplication.ScreenUpdating = False

    With xlWorksheet

        .Range("A1:E1").Merge
        .Range("A2:E2").Merge
        .Range("A1:E3").Font.Bold = True

        .Cells(1, 1).value = "Resumen de Cuenta Corriente"
        .Cells(2, 1).value = "Cliente: " & Me.cboClientes.Text

        .Cells(3, 1).value = "Fecha"
        .Cells(3, 2).value = "Comprobante"
        .Cells(3, 3).value = "Debe"
        .Cells(3, 4).value = "Haber"
        .Cells(3, 5).value = "Saldo"

    End With

    Dim idx As Long
    Dim detalleActual As DTODetalleCuentaCorriente

    idx = 4

    For Each detalleActual In Detalles

        xlWorksheet.Cells(idx, 1).value = detalleActual.FEcha
        xlWorksheet.Cells(idx, 2).value = detalleActual.Comprobante
        xlWorksheet.Cells(idx, 3).value = detalleActual.Debe
        xlWorksheet.Cells(idx, 4).value = detalleActual.Haber
        xlWorksheet.Cells(idx, 5).value = detalleActual.saldo

        idx = idx + 1

    Next detalleActual

    xlWorksheet.Cells.EntireColumn.AutoFit

    With xlWorksheet.PageSetup

        .Orientation = xlLandscape
        .BottomMargin = xlApplication.CentimetersToPoints(1)
        .TopMargin = xlApplication.CentimetersToPoints(1)
        .LeftMargin = xlApplication.CentimetersToPoints(1)
        .RightMargin = xlApplication.CentimetersToPoints(1)

    End With

    Dim filename As String

    filename = funciones.GetTmpPath() & _
               "CuentaCorriente_" & _
               Format$(Now, "yyyymmdd_hhnnss") & _
               ".xlsx"

    If Dir$(filename) <> vbNullString Then
        Kill filename
    End If

    xlWorkbook.SaveAs filename

    xlWorkbook.Saved = True
    xlWorkbook.Close False

    xlApplication.ScreenUpdating = True
    xlApplication.DisplayAlerts = True
    xlApplication.Quit

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    funciones.ShellExecute 0, "open", filename, "", "", 0

    ExportToXls = True
    Exit Function

ManejarError:

    Dim descripcionError As String
    descripcionError = Err.Description

    On Error Resume Next

    If Not xlWorkbook Is Nothing Then
        xlWorkbook.Close False
    End If

    If Not xlApplication Is Nothing Then
        xlApplication.DisplayAlerts = True
        xlApplication.Quit
    End If

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

    On Error GoTo 0

    MsgBox "No se pudo exportar la cuenta corriente." & vbCrLf & _
           descripcionError, _
           vbCritical + vbOKOnly, _
           "Error de exportación"

    ExportToXls = False

End Function


Private Function ObtenerImporteMinimo( _
    ByRef importeMinimo As Double) As Boolean

    Dim texto As String

    texto = Trim$(Me.txtImporteMinimo.Text)

    If LenB(texto) = 0 Then
        texto = "0"
    End If

    If Not IsNumeric(texto) Then
        MsgBox "El importe mínimo debe ser un número válido.", _
               vbExclamation + vbOKOnly, _
               "Cuenta corriente"

        Me.txtImporteMinimo.SetFocus
        ObtenerImporteMinimo = False
        Exit Function
    End If

    importeMinimo = CDbl(texto)

    If importeMinimo < 0 Then
        MsgBox "El importe mínimo no puede ser negativo.", _
               vbExclamation + vbOKOnly, _
               "Cuenta corriente"

        Me.txtImporteMinimo.SetFocus
        ObtenerImporteMinimo = False
        Exit Function
    End If

    Me.txtImporteMinimo.Text = Format$(importeMinimo, "0.00")

    ObtenerImporteMinimo = True

End Function

Private Function FiltrarDetallesPorImporte( _
    ByVal detallesOrigen As Collection, _
    ByVal importeMinimo As Double) As Collection

    Dim resultado As New Collection
    Dim detalleActual As DTODetalleCuentaCorriente

    If Not IsSomething(detallesOrigen) Then
        Set FiltrarDetallesPorImporte = resultado
        Exit Function
    End If

    For Each detalleActual In detallesOrigen

        If importeMinimo <= 0 _
           Or Abs(CDbl(detalleActual.Debe)) >= importeMinimo _
           Or Abs(CDbl(detalleActual.Haber)) >= importeMinimo Then

            resultado.Add detalleActual

        End If

    Next detalleActual

    Set FiltrarDetallesPorImporte = resultado

End Function


