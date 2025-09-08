VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmSistemaTests 
   Caption         =   "Tests"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   Icon            =   "frmSistemasTests.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   14985
   Begin XtremeSuiteControls.PushButton btnSeguimientoAvanzado 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   4935
      _Version        =   786432
      _ExtentX        =   8705
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "PushButton1"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin GridEX20.GridEX grilla_monedas 
      Height          =   3495
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
      Version         =   "2.0"
      BoundColumnIndex=   "numero"
      ReplaceColumnIndex=   "moneda"
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmSistemasTests.frx":000C
      Column(2)       =   "frmSistemasTests.frx":00F8
      Column(3)       =   "frmSistemasTests.frx":020C
      Column(4)       =   "frmSistemasTests.frx":0318
      Column(5)       =   "frmSistemasTests.frx":0404
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmSistemasTests.frx":0508
      FormatStyle(2)  =   "frmSistemasTests.frx":0640
      FormatStyle(3)  =   "frmSistemasTests.frx":06F0
      FormatStyle(4)  =   "frmSistemasTests.frx":07A4
      FormatStyle(5)  =   "frmSistemasTests.frx":087C
      FormatStyle(6)  =   "frmSistemasTests.frx":0934
      ImageCount      =   0
      PrinterProperties=   "frmSistemasTests.frx":0A14
   End
   Begin GridEX20.GridEX grilla_moneda 
      Height          =   3495
      Left            =   13200
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   6165
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "moneda"
      ActAsDropDown   =   -1  'True
      ColumnAutoResize=   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      ColumnHeaders   =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmSistemasTests.frx":0BEC
      Column(2)       =   "frmSistemasTests.frx":0D10
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmSistemasTests.frx":0E04
      FormatStyle(2)  =   "frmSistemasTests.frx":0F3C
      FormatStyle(3)  =   "frmSistemasTests.frx":0FEC
      FormatStyle(4)  =   "frmSistemasTests.frx":10A0
      FormatStyle(5)  =   "frmSistemasTests.frx":1178
      FormatStyle(6)  =   "frmSistemasTests.frx":1230
      ImageCount      =   0
      PrinterProperties=   "frmSistemasTests.frx":1310
   End
   Begin XtremeSuiteControls.PushButton PushButton 
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4935
      _Version        =   786432
      _ExtentX        =   8705
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Agenda nueva"
      Appearance      =   6
   End
   Begin VB.CommandButton Command 
      Caption         =   "frmAdminExtrasReporteIVACompras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton btnPrueba_04_Click 
      Caption         =   "frmAdminPagosCrearOrdenPagoNew"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4935
   End
End
Attribute VB_Name = "frmSistemaTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private moneda As clsMoneda
Private colMonedas As New Collection
Dim vOrdenPago As OrdenPago
Dim monedaplicada As clsMonedaAplicada
Private operacion As operacion


Private Sub btnSeguimientoAvanzado_Click()
    Dim f228 As New frmPlaneamientoSeguimientoAvanzado
    f228.Show
End Sub

Private Sub Command1_Click()
    
    Set monedaplicada = New clsMonedaAplicada

   If colMonedas.count > 0 Then
   
       Me.grilla_monedas.ItemCount = 0
       Me.grilla_monedas.Refresh
       
       monedaplicada.moneda = DAOMoneda.GetById(colMonedas(1).Id)

       Me.grilla_monedas.ItemCount = 1
       
    End If
    
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    
    GridEXHelper.CustomizeGrid Me.grilla_moneda, False, True
    GridEXHelper.CustomizeGrid Me.grilla_monedas, False, True
    
    Set colMonedas = DAOMoneda.GetAll()
    Me.grilla_moneda.ItemCount = colMonedas.count
  
    Set Me.grilla_monedas.Columns("moneda").DropDownControl = Me.grilla_moneda
  
End Sub


Private Sub grilla_moneda_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set moneda = colMonedas.item(RowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
End Sub


Private Sub grilla_monedas_GotFocus()
    grilla_monedas.SelStart = 0
    grilla_monedas.SelLength = -1
End Sub


Private Sub grilla_monedas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set monedaplicada = vFactura.IvaAplicado.item(RowIndex)
    Values(1) = ""
    Values(2) = monedaplicada.moneda.NombreCorto
    Values(3) = ""
    Values(4) = ""
    Values(5) = ""
End Sub



Private Sub PushButton_Click()
    Dim f227 As New frmAgendaNueva
    f227.Show
End Sub

