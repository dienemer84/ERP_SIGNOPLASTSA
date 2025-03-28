VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmSistemaTests 
   Caption         =   "Tests"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   Icon            =   "frmSistemasTests.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   12315
   Begin GridEX20.GridEX grilla_monedas 
      Height          =   3495
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6165
      Version         =   "2.0"
      BoundColumnIndex=   "numero"
      ReplaceColumnIndex=   "moneda"
      HideSelection   =   2
      UseEvenOddColor =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmSistemasTests.frx":000C
      Column(2)       =   "frmSistemasTests.frx":0174
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmSistemasTests.frx":03E8
      FormatStyle(2)  =   "frmSistemasTests.frx":0520
      FormatStyle(3)  =   "frmSistemasTests.frx":05D0
      FormatStyle(4)  =   "frmSistemasTests.frx":0684
      FormatStyle(5)  =   "frmSistemasTests.frx":075C
      FormatStyle(6)  =   "frmSistemasTests.frx":0814
      ImageCount      =   0
      PrinterProperties=   "frmSistemasTests.frx":08F4
   End
   Begin GridEX20.GridEX grilla_moneda 
      Height          =   3495
      Left            =   9240
      TabIndex        =   3
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   6165
      Version         =   "2.0"
      BoundColumnIndex=   "id"
      ReplaceColumnIndex=   "moneda"
      ActAsDropDown   =   -1  'True
      HideSelection   =   2
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmSistemasTests.frx":0ACC
      Column(2)       =   "frmSistemasTests.frx":0BF0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmSistemasTests.frx":0CE4
      FormatStyle(2)  =   "frmSistemasTests.frx":0E1C
      FormatStyle(3)  =   "frmSistemasTests.frx":0ECC
      FormatStyle(4)  =   "frmSistemasTests.frx":0F80
      FormatStyle(5)  =   "frmSistemasTests.frx":1058
      FormatStyle(6)  =   "frmSistemasTests.frx":1110
      ImageCount      =   0
      PrinterProperties=   "frmSistemasTests.frx":11F0
   End
   Begin XtremeSuiteControls.PushButton PushButton 
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2280
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
Private Monedas As New Collection

Private Sub Form_Load()
    FormHelper.Customize Me
    
    GridEXHelper.CustomizeGrid Me.grilla_moneda, False, True
    
    GridEXHelper.CustomizeGrid Me.grilla_monedas, False, True
    
    Set Monedas = DAOMoneda.GetAll()
    Me.grilla_moneda.ItemCount = Monedas.count
  
    Set Me.grilla_monedas.Columns("moneda").DropDownControl = Me.grilla_moneda
  
End Sub



Private Sub grilla_moneda_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Monedas.count > 0 Then
        Set moneda = Monedas.item(rowIndex)
        Values(1) = moneda.Id
        Values(2) = moneda.NombreCorto
    End If
End Sub


Private Sub PushButton_Click()
    Dim f227 As New frmAgendaNueva
    f227.Show
End Sub

