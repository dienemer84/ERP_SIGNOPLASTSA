VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmTip 
   AutoRedraw      =   -1  'True
   Caption         =   "Log de actualización"
   ClientHeight    =   4080
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   945
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdNuevoDetalle 
      Caption         =   "Cargar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin GridEX20.GridEX grid 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   5953
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ShowEmptyFields =   0   'False
      GroupFooterStyle=   1
      RowHeight       =   33
      TabKeyBehavior  =   1
      ColumnAutoResize=   -1  'True
      DetectRowDrag   =   -1  'True
      HeaderStyle     =   3
      UseEvenOddColor =   -1  'True
      ReadOnly        =   -1  'True
      MethodHoldFields=   -1  'True
      AllowCardSizing =   0   'False
      Options         =   -1
      RecordsetType   =   1
      AllowEdit       =   0   'False
      BorderStyle     =   2
      GroupByBoxVisible=   0   'False
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   4
      Column(1)       =   "frmTip.frx":0000
      Column(2)       =   "frmTip.frx":017C
      Column(3)       =   "frmTip.frx":02F0
      Column(4)       =   "frmTip.frx":0410
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmTip.frx":0558
      FormatStyle(2)  =   "frmTip.frx":0690
      FormatStyle(3)  =   "frmTip.frx":0740
      FormatStyle(4)  =   "frmTip.frx":07F4
      FormatStyle(5)  =   "frmTip.frx":08CC
      FormatStyle(6)  =   "frmTip.frx":0984
      ImageCount      =   0
      PrinterProperties=   "frmTip.frx":0A64
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   12840
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim actualizaciones As Collection
Dim sin As Actualizacion
Dim CantArchivos As Dictionary

Private Sub cmdNuevoDetalle_Click()
    frmSistemaAgregarNotasActualizacion.Show
End Sub

'Private Sub cmdOK_Click()
'    Unload Me
'
'End Sub

Private Sub Form_Load()
    Cargar
    
        Me.caption = caption & " (" & Name & ")"
    
End Sub

Private Sub Cargar()
    Set actualizaciones = DAOActualizar.FindAll()
    Me.grid.ItemCount = 0
    Me.grid.ItemCount = actualizaciones.count

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.grid.Width = Me.ScaleWidth - 20
    Me.grid.Height = Me.ScaleHeight - 50
    
    'Me.cmdOK.Left = Me.ScaleWidth - 180
    'Me.cmdOK.Top = Me.grid.Height + 100
    
    'Me.Height = 4600
    'Me.grid.ColumnAutoResize = True

End Sub

Private Sub grid_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    
    If RowIndex > 0 And actualizaciones.count > 0 Then
        
        Set sin = actualizaciones(RowIndex)
     
        Values(1) = sin.Id_
        Values(2) = sin.Fecha_
        Values(3) = sin.Detalle_
        Values(4) = sin.Modulo_


    End If
End Sub


