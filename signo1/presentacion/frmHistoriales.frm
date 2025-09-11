VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmHistoriales 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Historial"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6135
   Icon            =   "frmHistoriales.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   6135
   Begin GridEX20.GridEX GridEX1 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5318
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
      Column(1)       =   "frmHistoriales.frx":000C
      Column(2)       =   "frmHistoriales.frx":0170
      Column(3)       =   "frmHistoriales.frx":0264
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmHistoriales.frx":0348
      FormatStyle(2)  =   "frmHistoriales.frx":0480
      FormatStyle(3)  =   "frmHistoriales.frx":0530
      FormatStyle(4)  =   "frmHistoriales.frx":05E4
      FormatStyle(5)  =   "frmHistoriales.frx":06BC
      FormatStyle(6)  =   "frmHistoriales.frx":0810
      ImageCount      =   0
      PrinterProperties=   "frmHistoriales.frx":08F0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmHistoriales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp As clsHistorial
Dim vLista As Collection

Public Property Let lista(nValue As Collection)
    Set vLista = nValue
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
    Me.GridEX1.Height = Me.ScaleHeight - 650
    Me.Command1.Top = Me.GridEX1.Height + 150
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set tmp = vLista.item(RowIndex)
    Values(1) = tmp.FEcha
    Values(2) = tmp.Usuario.Usuario
    Values(3) = tmp.mensaje
End Sub
