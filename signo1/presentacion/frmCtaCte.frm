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
      Left            =   7440
      TabIndex        =   7
      Top             =   6360
      Width           =   2385
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
    If Me.cboClientes.ListIndex <> -1 Then
        Dim fecha_hasta As String
        If Not IsNull(Me.dtpHasta.value) Then
            fecha_hasta = Format(Me.dtpHasta.value, "yyyy-mm-dd")
        End If



        Set Detalles = DAOCuentaCorriente.FindAllDetalles(Me.cboClientes.ItemData(Me.cboClientes.ListIndex), , fecha_hasta)
        saldo = 0
        
        
       
        If IsSomething(Detalles) Then
            Me.lblSaldo = "Saldo: " & Replace(FormatCurrency(funciones.FormatearDecimales(DAOCuentaCorriente.GetSaldo(Detalles))), "$", "")
        End If
        Set saldos = New Dictionary
        saldo = 0
        Me.gridDetalles.ItemCount = 0
        Me.gridDetalles.ItemCount = Detalles.count
    End If


End Sub



Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridDetalles

    DAOCliente.llenarComboXtremeSuite Me.cboClientes
    Me.cboClientes.ListIndex = -1

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
        frm1.reciboId = deta.IdComprobante
        frm1.Show
    End If
    
    
    If (deta.tipoComprobante = TipoComprobanteUsado.ReciboAnticipo_) Then
        Dim frm2 As New frmAdminCobranzasNuevoReciboAnticipo
        frm2.editar = False
        frm2.reciboId = deta.IdComprobante
        frm2.Show
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
        .HeaderString(jgexHFCenter) = "Cuenta Corriente de " & Me.cboClientes.text
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
    
    'Dim xlApplication As New Excel.Application
    Dim xlApplication As Object
    Set xlApplication = CreateObject("Excel.Application")

    'Dim xlWorkbook As New Excel.Workbook
    Dim xlWorkbook As Object
    Set xlWorkbook = CreateObject("Excel.Application")

    'Dim xlWorksheet As New Excel.Worksheet
    Dim xlWorksheet As Object
    Set xlWorksheet = CreateObject("Excel.Application")

    Set xlWorkbook = xlApplication.Workbooks.Add

    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Range("A1:E1").Merge
    xlWorksheet.Range("A2:E2").Merge
    xlWorksheet.Range("A1:E3").Font.Bold = True
    xlWorksheet.Cells(1, 1).value = "Resumen de Cuenta Corriente"
    xlWorksheet.Cells(2, 1).value = "Cliente: " & Me.cboClientes.text
    xlWorksheet.Cells(3, 1).value = "Fecha"
    xlWorksheet.Cells(3, 2).value = "Comprobante"
    xlWorksheet.Cells(3, 3).value = "Debe"
    xlWorksheet.Cells(3, 4).value = "Haber"
    xlWorksheet.Cells(3, 5).value = "Saldo"

    Dim idx As Integer
    idx = 4
    
   For Each deta In Detalles
  

         xlWorksheet.Cells(idx, 1).value = deta.FEcha
         xlWorksheet.Cells(idx, 2).value = deta.Comprobante
         xlWorksheet.Cells(idx, 3).value = deta.Debe
         xlWorksheet.Cells(idx, 4).value = deta.Haber
         xlWorksheet.Cells(idx, 5).value = deta.saldo
         
    idx = idx + 1
    
     Next

        xlApplication.ScreenUpdating = False
        
        Dim wkSt As String
        
        wkSt = xlWorksheet.Name
        
        xlWorksheet.Cells.EntireColumn.AutoFit
        
        xlWorkbook.Sheets(wkSt).Select
        
        xlApplication.ScreenUpdating = True
        
        xlWorksheet.PageSetup.Orientation = xlLandscape
        xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
        xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)
    
        Dim filename As String
        filename = funciones.GetTmpPath() & "tmp_info " & Hour(Now) & Minute(Now) & Second(Now) & " .xlsx"
    
        If Dir(filename) <> vbNullString Then Kill filename
       
        xlWorkbook.SaveAs filename
    
        xlWorkbook.Saved = True
        xlWorkbook.Close
        xlApplication.Quit
        
        funciones.ShellExecute 0, "open", filename, "", "", 0
    
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApplication = Nothing
        

End Function
