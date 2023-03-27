VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmHistoricoMateriales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historico Materiales"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin GridEX20.GridEX GridEX1 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   6271
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "Mensaje"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmHistoricoMateriales.frx":0000
      Column(2)       =   "frmHistoricoMateriales.frx":0164
      Column(3)       =   "frmHistoricoMateriales.frx":0258
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmHistoricoMateriales.frx":033C
      FormatStyle(2)  =   "frmHistoricoMateriales.frx":0474
      FormatStyle(3)  =   "frmHistoricoMateriales.frx":0524
      FormatStyle(4)  =   "frmHistoricoMateriales.frx":05D8
      FormatStyle(5)  =   "frmHistoricoMateriales.frx":06B0
      FormatStyle(6)  =   "frmHistoricoMateriales.frx":0804
      ImageCount      =   0
      PrinterProperties=   "frmHistoricoMateriales.frx":08E4
   End
End
Attribute VB_Name = "frmHistoricoMateriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp As Historial
Dim vLista As Collection

Public Property Let IdMaterial(Id As Long)
    Set vLista = DaoHistorico.GetAll("materiales_historial", Id)
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub llenarGrilla()
    Me.GridEX1.ItemCount = vLista.count
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    llenarGrilla
    GridEXHelper.CustomizeGrid Me.GridEX1
End Sub

Private Sub Form_Resize()
    Me.GridEX1.Width = Me.ScaleWidth
    Me.GridEX1.Height = Me.ScaleHeight - 150

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set tmp = vLista.item(RowIndex)
    Values(1) = tmp.FEcha
    Values(2) = tmp.Autor
    Values(3) = tmp.mensaje
End Sub

