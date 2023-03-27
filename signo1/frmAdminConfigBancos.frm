VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAdminConfigBancos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bancos..."
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10935
   Begin GridEX20.GridEX GridEX1 
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   9075
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      RowHeaders      =   -1  'True
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmAdminConfigBancos.frx":0000
      Column(2)       =   "frmAdminConfigBancos.frx":0134
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminConfigBancos.frx":0228
      FormatStyle(2)  =   "frmAdminConfigBancos.frx":0360
      FormatStyle(3)  =   "frmAdminConfigBancos.frx":0410
      FormatStyle(4)  =   "frmAdminConfigBancos.frx":04C4
      FormatStyle(5)  =   "frmAdminConfigBancos.frx":059C
      FormatStyle(6)  =   "frmAdminConfigBancos.frx":0654
      ImageCount      =   0
      PrinterProperties=   "frmAdminConfigBancos.frx":0734
   End
End
Attribute VB_Name = "frmAdminConfigBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bancos As New Collection
Private Banco As Banco
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, True
    llenarLista
End Sub
Private Sub Form_Resize()
    Me.GridEX1.Width = Me.ScaleWidth
    Me.GridEX1.Height = Me.ScaleHeight
End Sub
Private Sub llenarLista()
    Me.GridEX1.ItemCount = 0
    Set bancos = DAOBancos.GetAll
    Me.GridEX1.ItemCount = bancos.count
End Sub


Private Sub GridEX1_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Not (MsgBox("¿Está seguro de actualizar los datos?", vbYesNo, "Consulta") = vbYes)
End Sub
Private Sub GridEX1_SelectionChange()
'On Error Resume Next'
'   Set banco = bancos.Item(Me.GridEX1.RowIndex(Me.GridEX1.Row))
End Sub


Private Sub GridEX1_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set Banco = New Banco
    Banco.Id = Values(1)
    Banco.nombre = Values(2)
    If DAOBancos.Save(Banco) Then bancos.Add Banco, CStr(Banco.Id)

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set Banco = bancos.item(RowIndex)
    Values(1) = Banco.Id
    Values(2) = "(ID " & Banco.Id & ")- " & Banco.nombre
End Sub
Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set Banco = bancos.item(RowIndex)
    Banco.Id = Values(1)
    Banco.nombre = Values(2)
    If Not DAOBancos.Save(Banco) Then GoTo err1
    llenarLista
    Exit Sub
err1:
End Sub
