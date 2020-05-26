VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmCtaCteProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente Proveedores"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
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
   ScaleHeight     =   6810
   ScaleWidth      =   9495
   Begin GridEX20.GridEX gridDetalles 
      Height          =   4605
      Left            =   105
      TabIndex        =   0
      Top             =   1695
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   8123
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
         TabIndex        =   17
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
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   330
         Left            =   4860
         TabIndex        =   8
         Top             =   1065
         Width           =   1065
         _Version        =   786432
         _ExtentX        =   1879
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdVerCtaCte 
         Height          =   315
         Left            =   6000
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   150
      TabIndex        =   5
      Top             =   6435
      Width           =   9225
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



Private Sub LlenarLiquidaciones(id As Long)

    Dim liq As CuentaCorrienteHistoric
    Dim col As Collection
    Dim i As Long
    Set col = DAOCuentaCorrienteHistoric.GetAll(proveedor_, id, False)
    cboLiquidaciones.Clear
    For i = 1 To col.count
        Set liq = col(i)
        cboLiquidaciones.AddItem liq.Periodo
        cboLiquidaciones.ItemData(cboLiquidaciones.NewIndex) = liq.id
    Next i
    If cboLiquidaciones.ListCount > 0 Then
        cboLiquidaciones.ListIndex = 0
    End If

End Sub


Private Sub ver()
    Dim id As Long
    Dim condition As String



    If Me.cboClientes.ListIndex <> -1 Then
        id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
        LlenarLiquidaciones (id)
        If Not IsNull(Me.dtpHasta.value) Then
            condition = conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd"))
            Set Detalles = DAOCuentaCorriente.FindAllDetallesProveedor(id, , condition, True)
        Else
            Set Detalles = DAOCuentaCorriente.FindAllDetallesProveedor(id, , , True)
        End If


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
Private Sub ver2()

    Dim id As Long
    Dim condition As String
    Dim cuenta As CuentaCorrienteHistoric


    If Me.cboLiquidaciones.ListIndex <> -1 Then
        id = Me.cboLiquidaciones.ItemData(Me.cboLiquidaciones.ListIndex)

        Set Detalles = DAOCuentaCorrienteHistoric.GetById(proveedor_, id).Detalles
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


Private Sub cboClientes_Click()
    Dim id As Long
    If Me.cboClientes.ListIndex > 0 Then
        id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
        LlenarLiquidaciones id
        setmaxdesde id
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

Private Sub setmaxdesde(id As Long)
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
        Dim f22 As New frmCrearOrdenPago
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
    If RowBuffer.RowIndex > 0 Then
        Set deta = Detalles.item(RowBuffer.RowIndex)
        If Not deta.AtributoExtra And deta.Debe > 0 And deta.Haber = 0 Then    'no esta en ninguna orden
            RowBuffer.RowStyle = "Impaga"
        End If
    End If
End Sub
Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Detalles.count > 0 Then
        Set deta = Detalles.item(RowIndex)
        Values(1) = deta.FEcha
        Values(2) = deta.Comprobante
        Values(3) = deta.Debe
        Values(4) = deta.Haber

        If saldos.Exists(CStr(RowIndex)) Then
            Values(5) = saldos.item(CStr(RowIndex))
        Else
            saldo = saldo + deta.Debe - deta.Haber
            saldos.Add CStr(RowIndex), saldo
            Values(5) = funciones.RedondearDecimales(saldo)
        End If

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

Private Sub PushButton2_Click()
    Dim id As Long
    If IsNull(Me.dtpHasta.value) Then
        MsgBox "Debe definir una fecha tope para cerrar!", vbCritical, "Advertencia"
    Else

        If MsgBox("¿Desea cerrar el periodo hasta " & Format(Me.dtpHasta, "dd-mm-yyyy") & " del cliente " & Me.cboClientes.text & " ?", vbYesNo, "Consulta") = vbYes Then
            id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
            If Not DAOCuentaCorriente.CerrarPeriodoCtaCteProveedor(id, Me.dtpHasta) Then
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
