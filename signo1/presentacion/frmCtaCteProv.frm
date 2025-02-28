VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmCtaCteProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente Proveedores"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCtaCteProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9435
   Begin XtremeSuiteControls.PushButton button_ExportToXlsProv 
      Height          =   435
      Left            =   2160
      TabIndex        =   18
      Top             =   6960
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Exportar a XLS"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX gridDetalles 
      Height          =   5040
      Left            =   105
      TabIndex        =   0
      Top             =   1695
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   8890
      Version         =   "2.0"
      HoldSortSettings=   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmCtaCteProv.frx":000C
      Column(2)       =   "frmCtaCteProv.frx":0184
      Column(3)       =   "frmCtaCteProv.frx":02A8
      Column(4)       =   "frmCtaCteProv.frx":0400
      Column(5)       =   "frmCtaCteProv.frx":0558
      FormatStylesCount=   7
      FormatStyle(1)  =   "frmCtaCteProv.frx":06B0
      FormatStyle(2)  =   "frmCtaCteProv.frx":07D8
      FormatStyle(3)  =   "frmCtaCteProv.frx":0888
      FormatStyle(4)  =   "frmCtaCteProv.frx":093C
      FormatStyle(5)  =   "frmCtaCteProv.frx":0A14
      FormatStyle(6)  =   "frmCtaCteProv.frx":0ACC
      FormatStyle(7)  =   "frmCtaCteProv.frx":0BAC
      ImageCount      =   0
      PrinterProperties=   "frmCtaCteProv.frx":0C38
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1515
      Left            =   7290
      TabIndex        =   1
      Top             =   75
      Width           =   2085
      _Version        =   786432
      _ExtentX        =   3678
      _ExtentY        =   2672
      _StockProps     =   79
      Caption         =   "Estado Proveedor"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkContado 
         Height          =   195
         Left            =   405
         TabIndex        =   2
         Top             =   225
         Width           =   1635
         _Version        =   786432
         _ExtentX        =   2884
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Contado"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkCtaCte 
         Height          =   315
         Left            =   405
         TabIndex        =   3
         Top             =   465
         Width           =   1605
         _Version        =   786432
         _ExtentX        =   2831
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Cuenta Corriente"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkEliminado 
         Height          =   315
         Left            =   405
         TabIndex        =   4
         Top             =   750
         Width           =   1410
         _Version        =   786432
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Inactivos"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1530
      Left            =   60
      TabIndex        =   6
      Top             =   75
      Width           =   7185
      _Version        =   786432
      _ExtentX        =   12674
      _ExtentY        =   2699
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkAnteriores 
         Height          =   240
         Left            =   2685
         TabIndex        =   16
         Top             =   1095
         Width           =   2025
         _Version        =   786432
         _ExtentX        =   3572
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Ver detalles anteriores"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   360
         Left            =   5985
         TabIndex        =   7
         Top             =   1050
         Width           =   1065
         _Version        =   786432
         _ExtentX        =   1879
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Generar Liq"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdVerCtaCte 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   270
         Width           =   1080
         _Version        =   786432
         _ExtentX        =   1905
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Ver"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   270
         Width           =   4830
         _Version        =   786432
         _ExtentX        =   8520
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpHasta 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   1020
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.ComboBox cboLiquidaciones 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   645
         Width           =   4830
         _Version        =   786432
         _ExtentX        =   8520
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Default         =   -1  'True
         Height          =   315
         Left            =   5985
         TabIndex        =   15
         Top             =   630
         Width           =   1080
         _Version        =   786432
         _ExtentX        =   1905
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Ver"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   555
         TabIndex        =   14
         Top             =   1080
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   255
         TabIndex        =   13
         Top             =   315
         Width           =   750
         _Version        =   786432
         _ExtentX        =   1323
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   690
         Width           =   945
         _Version        =   786432
         _ExtentX        =   1667
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Liquidaciones"
         Alignment       =   1
         AutoSize        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   435
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   1680
      _Version        =   786432
      _ExtentX        =   2963
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   150
      TabIndex        =   5
      Top             =   7005
      Width           =   8625
   End
End
Attribute VB_Name = "frmCtaCteProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Detalles As Collection
Private deta As DTODetalleCuentaCorriente
Private saldo As Double
Private saldos As New Dictionary



Private Sub LlenarLiquidaciones(Id As Long)

    Dim liq As CuentaCorrienteHistoric
    Dim col As Collection
    Dim i As Long
    Set col = DAOCuentaCorrienteHistoric.GetAll(proveedor_, Id, False)
    cboLiquidaciones.Clear
    For i = 1 To col.count
        Set liq = col(i)
        cboLiquidaciones.AddItem liq.Periodo
        cboLiquidaciones.ItemData(cboLiquidaciones.NewIndex) = liq.Id
    Next i
    If cboLiquidaciones.ListCount > 0 Then
        cboLiquidaciones.ListIndex = 0
    End If

End Sub


Private Sub ver()
    Dim Id As Long
    Dim condition As String

    If Me.cboClientes.ListIndex <> -1 Then
        Id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
        LlenarLiquidaciones (Id)
        If Not IsNull(Me.dtpHasta.value) Then
            condition = conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd"))
            Set Detalles = DAOCuentaCorriente.FindAllDetallesProveedor2(Id, , condition, True)
            saldo = 0
        Else
            Set Detalles = DAOCuentaCorriente.FindAllDetallesProveedor2(Id, , , True)
            saldo = 0
        End If


        If IsSomething(Detalles) Then
            Me.lblSaldo = "Saldo: " & funciones.FormatearDecimales(DAOCuentaCorriente.GetSaldo(Detalles))
        End If

        saldo = 0
        Set saldos = New Dictionary
        'trello  #179
        Me.gridDetalles.Refetch

        Me.gridDetalles.ItemCount = 0
        If Detalles.count > 0 Then
            Me.gridDetalles.ItemCount = Detalles.count
            GridEXHelper.AutoSizeColumns Me.gridDetalles
        End If
    End If
End Sub

Private Sub ver2()

    Dim Id As Long

    If Me.cboLiquidaciones.ListIndex <> -1 Then
        Id = Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.ListIndex)

        Set Detalles = DAOCuentaCorrienteHistoric.GetById(proveedor_, Id).Detalles
        If IsSomething(Detalles) Then
            Me.lblSaldo = "Saldo: " & funciones.FormatearDecimales(DAOCuentaCorriente.GetSaldo(Detalles))
        End If

        saldo = 0
        Set saldos = New Dictionary
        Me.gridDetalles.ItemCount = 0
        Me.gridDetalles.ItemCount = Detalles.count
        GridEXHelper.AutoSizeColumns Me.gridDetalles
    End If
End Sub

Private Sub LlenarProveedores()
    DAOProveedor.llenarComboXtremeSuite Me.cboClientes, (Me.chkCtaCte.value = xtpChecked), (Me.chkContado.value = xtpChecked), (Me.chkEliminado.value = xtpChecked)
    Me.cboClientes.ListIndex = -1
    Me.gridDetalles.ItemCount = 0
End Sub


Private Sub button_ExportToXlsProv_Click()
    ExportToXlsProv

End Sub

Private Sub cboClientes_Click()
    Dim Id As Long
    If Me.cboClientes.ListIndex > 0 Then
        Id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
        LlenarLiquidaciones Id
        setmaxdesde Id
    End If
End Sub



Private Sub chkContado_Click()
    LlenarProveedores
End Sub
Private Sub chkCtaCte_Click()
    LlenarProveedores
End Sub
Private Sub chkEliminado_Click()
    LlenarProveedores
End Sub

Private Sub setmaxdesde(Id As Long)
'Me.dtpHasta.MinDate = DAOCuentaCorriente.getMaxDesdeProveedor(id)

End Sub

Private Sub cmdVerCtaCte_Click()
    ver
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.gridDetalles
    LlenarProveedores
End Sub

Private Sub gridDetalles_DblClick()
    Set deta = Detalles.item(gridDetalles.RowIndex(gridDetalles.row))


    If (deta.tipoComprobante = TipoComprobanteUsado.FacturaProveedor_) Then

        Dim frm As frmAdminComprasNuevaFCProveedor
        Set frm = New frmAdminComprasNuevaFCProveedor
        frm.ver = True
        frm.Factura = DAOFacturaProveedor.FindById(deta.IdComprobante)
        frm.Show
    End If


    If deta.tipoComprobante = TipoComprobanteUsado.OrdenPago_ Then
        Dim f22 As New frmAdminPagosCrearOrdenPago
        f22.Show
        f22.ReadOnly = True
        f22.Cargar DAOOrdenPago.FindById(deta.IdComprobante)
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
    If Detalles.count = 0 Then Exit Sub

    If RowBuffer.RowIndex > 0 Then
        Set deta = Detalles.item(RowBuffer.RowIndex)
        If Not deta.AtributoExtra And deta.Debe > 0 And deta.Haber = 0 Then    'no esta en ninguna orden
            RowBuffer.RowStyle = "Impaga"
        End If
    End If

    ''debug.print (deta.IdComprobante)

End Sub

Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Detalles.count > 0 Then
        Set deta = Detalles.item(RowIndex)
        Values(1) = deta.FEcha
        Values(2) = deta.Comprobante
        Values(3) = deta.Debe
        Values(4) = deta.Haber

        '        If saldos.Exists(CStr(RowIndex)) Then
        '            Values(5) = saldos.item(CStr(RowIndex))
        '        Else
        '            saldo = saldo + deta.Debe - deta.Haber
        '            saldos.Add CStr(RowIndex), saldo
        '            Values(5) = funciones.RedondearDecimales(saldo)
        '        End If
        '
        Values(5) = deta.saldo

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

Private Sub PushButton2_Click()
    Dim Id As Long
    If IsNull(Me.dtpHasta.value) Then
        MsgBox "Debe definir una fecha tope para cerrar!", vbCritical, "Advertencia"
    Else

        If MsgBox("¿Desea cerrar el periodo hasta " & Format(Me.dtpHasta, "dd-mm-yyyy") & " del cliente " & Me.cboClientes.Text & " ?", vbYesNo, "Consulta") = vbYes Then
            Id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
            If Not DAOCuentaCorriente.CerrarPeriodoCtaCteProveedor(Id, Me.dtpHasta) Then
                MsgBox "No puede cerrar el periodo seleccionado!", vbCritical
            Else
                MsgBox "Período cerrado correctamente!", vbInformation
                cboClientes_Click

            End If
        End If
    End If
End Sub

Private Sub PushButton3_Click()
    ver2
End Sub

Public Function ExportToXlsProv() As Boolean

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
    xlWorksheet.Cells(2, 1).value = "Proveedor: " & Me.cboClientes.Text
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
