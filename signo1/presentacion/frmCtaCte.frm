VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmCtaCte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente"
   ClientHeight    =   6450
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
   ScaleHeight     =   6450
   ScaleWidth      =   9975
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   315
      Left            =   8805
      TabIndex        =   6
      Top             =   495
      Width           =   1080
      _Version        =   786432
      _ExtentX        =   1905
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboClientes 
      Height          =   315
      Left            =   780
      TabIndex        =   5
      Top             =   105
      Width           =   7800
      _Version        =   786432
      _ExtentX        =   13758
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin GridEX20.GridEX gridDetalles 
      Height          =   4920
      Left            =   90
      TabIndex        =   1
      Top             =   1005
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   8678
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
      Column(2)       =   "frmCtaCte.frx":0184
      Column(3)       =   "frmCtaCte.frx":02A8
      Column(4)       =   "frmCtaCte.frx":0420
      Column(5)       =   "frmCtaCte.frx":0598
      FormatStylesCount=   8
      FormatStyle(1)  =   "frmCtaCte.frx":0710
      FormatStyle(2)  =   "frmCtaCte.frx":0838
      FormatStyle(3)  =   "frmCtaCte.frx":08E8
      FormatStyle(4)  =   "frmCtaCte.frx":099C
      FormatStyle(5)  =   "frmCtaCte.frx":0A74
      FormatStyle(6)  =   "frmCtaCte.frx":0B2C
      FormatStyle(7)  =   "frmCtaCte.frx":0C0C
      FormatStyle(8)  =   "frmCtaCte.frx":0CC4
      ImageCount      =   0
      PrinterProperties=   "frmCtaCte.frx":0D58
   End
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   315
      Left            =   795
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
      Height          =   315
      Left            =   8805
      TabIndex        =   4
      Top             =   75
      Width           =   1080
      _Version        =   786432
      _ExtentX        =   1905
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Ver"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8235
      TabIndex        =   7
      Top             =   6060
      Width           =   1425
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   570
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
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   495
      _Version        =   786432
      _ExtentX        =   873
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Cliente"
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

Private Sub cmdVerCtaCte_Click()
    If Me.cboClientes.ListIndex <> -1 Then
        Dim fecha_hasta As String
        If Not IsNull(Me.dtpHasta.value) Then
            fecha_hasta = Format(Me.dtpHasta.value, "yyyy-mm-dd")
        End If



        Set Detalles = DAOCuentaCorriente.FindAllDetalles(Me.cboClientes.ItemData(Me.cboClientes.ListIndex), , fecha_hasta)
        saldo = 0


        If IsSomething(Detalles) Then
            Me.lblSaldo = "Saldo: " & funciones.FormatearDecimales(DAOCuentaCorriente.GetSaldo(Detalles))
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
        Dim frm As New frmFacturaEdicion
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
        Values(3) = deta.Debe
        Values(4) = deta.Haber

        '   If saldos.Exists(CStr(RowIndex)) Then
        '        Values(5) = saldos.item(CStr(RowIndex))
        '   Else
        '        saldo = saldo + deta.Debe - deta.Haber
        '     saldos.Add CStr(RowIndex), saldo
        '        Values(5) = funciones.RedondearDecimales(saldo)
        '  End If

        Values(5) = deta.saldo

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
