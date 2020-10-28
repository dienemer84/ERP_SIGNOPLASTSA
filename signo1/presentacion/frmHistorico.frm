VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Begin VB.Form frmHistorico 
   Caption         =   "Historico"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
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
      Column(1)       =   "frmHistorico.frx":0000
      Column(2)       =   "frmHistorico.frx":0164
      Column(3)       =   "frmHistorico.frx":0258
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmHistorico.frx":0368
      FormatStyle(2)  =   "frmHistorico.frx":04A0
      FormatStyle(3)  =   "frmHistorico.frx":0550
      FormatStyle(4)  =   "frmHistorico.frx":0604
      FormatStyle(5)  =   "frmHistorico.frx":06DC
      FormatStyle(6)  =   "frmHistorico.frx":0830
      ImageCount      =   0
      PrinterProperties=   "frmHistorico.frx":0910
   End
End
Attribute VB_Name = "frmHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp As Historial
Dim vLista As New Collection


Public Function Configurar(tabla As String, id_source As Long, titulo As String)

    Me.caption = "Histórico de " & titulo
    Set vLista = DaoHistorico.GetAll(tabla, id_source)
    llenarGrilla
    If Not IsSomething(vLista) Then Set vLista = New Collection
End Function


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
    Values(2) = tmp.Autor
    Values(3) = tmp.mensaje
End Sub

