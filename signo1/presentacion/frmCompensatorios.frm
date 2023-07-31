VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmCompensatorios 
   Caption         =   "Documentos Compensatorios"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   8400
   Begin GridEX20.GridEX GridEX1 
      Height          =   2805
      Left            =   15
      TabIndex        =   2
      Top             =   1305
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4948
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "observacion"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   8
      Column(1)       =   "frmCompensatorios.frx":0000
      Column(2)       =   "frmCompensatorios.frx":0108
      Column(3)       =   "frmCompensatorios.frx":0200
      Column(4)       =   "frmCompensatorios.frx":0318
      Column(5)       =   "frmCompensatorios.frx":0480
      Column(6)       =   "frmCompensatorios.frx":056C
      Column(7)       =   "frmCompensatorios.frx":0670
      Column(8)       =   "frmCompensatorios.frx":0790
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmCompensatorios.frx":088C
      FormatStyle(2)  =   "frmCompensatorios.frx":09C4
      FormatStyle(3)  =   "frmCompensatorios.frx":0A74
      FormatStyle(4)  =   "frmCompensatorios.frx":0B28
      FormatStyle(5)  =   "frmCompensatorios.frx":0C00
      FormatStyle(6)  =   "frmCompensatorios.frx":0CB8
      ImageCount      =   0
      PrinterProperties=   "frmCompensatorios.frx":0D98
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1200
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8205
      _Version        =   786432
      _ExtentX        =   14473
      _ExtentY        =   2117
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   405
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   1005
         _Version        =   786432
         _ExtentX        =   1773
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCompensatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col As New Collection
Dim comp As New Compensatorio

Private Sub cmdBuscar_Click()
    llenarLista
    Me.GridEX1.ItemCount = col.count
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False
    Me.GridEX1.ItemCount = 0
End Sub
Private Sub llenarLista()
    Set col = DAOCompensatorios.FindAll()
End Sub

Private Sub Form_Resize()
    Me.GroupBox1.Width = Me.ScaleWidth - 100
    Me.GridEX1.Width = Me.GroupBox1.Width
    Me.GridEX1.Height = Me.ScaleHeight - 500
End Sub
Private Sub GridEX1_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set comp = col.item(rowIndex)
    Values(1) = comp.Id
    Values(2) = comp.IdOrdenPago
    Values(3) = comp.FechaCancelacion
    Values(4) = comp.Monto
    Values(5) = TiposCompensatorio.item(CStr(comp.Tipo))
    Values(6) = comp.Comprobante.NumeroFormateado
    Values(7) = comp.Observacion
    Values(8) = comp.Comprobante.Proveedor.RazonSocial
End Sub

Private Sub PushButton1_Click()
End Sub
