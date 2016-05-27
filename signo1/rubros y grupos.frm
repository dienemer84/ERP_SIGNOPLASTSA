VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmRubrosGrupos 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros y Grupos..."
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   12165
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   12165
   Begin GridEX20.GridEX GridEX2 
      Height          =   4815
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8493
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   1
      Column(1)       =   "rubros y grupos.frx":0000
      FormatStylesCount=   6
      FormatStyle(1)  =   "rubros y grupos.frx":00F0
      FormatStyle(2)  =   "rubros y grupos.frx":0228
      FormatStyle(3)  =   "rubros y grupos.frx":02D8
      FormatStyle(4)  =   "rubros y grupos.frx":038C
      FormatStyle(5)  =   "rubros y grupos.frx":0464
      FormatStyle(6)  =   "rubros y grupos.frx":051C
      ImageCount      =   0
      PrinterProperties=   "rubros y grupos.frx":05FC
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8493
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "rubros y grupos.frx":07D4
      Column(2)       =   "rubros y grupos.frx":08CC
      FormatStylesCount=   6
      FormatStyle(1)  =   "rubros y grupos.frx":0998
      FormatStyle(2)  =   "rubros y grupos.frx":0AD0
      FormatStyle(3)  =   "rubros y grupos.frx":0B80
      FormatStyle(4)  =   "rubros y grupos.frx":0C34
      FormatStyle(5)  =   "rubros y grupos.frx":0D0C
      FormatStyle(6)  =   "rubros y grupos.frx":0DC4
      ImageCount      =   0
      PrinterProperties=   "rubros y grupos.frx":0EA4
   End
End
Attribute VB_Name = "frmRubrosGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rubros As Collection
Dim grupos As Collection
Dim tmpRubro As clsRubros
Dim tmpGrupo As clsGrupo
Dim rubroElegido As clsRubros

Private Sub Command6_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, False
    GridEXHelper.CustomizeGrid Me.GridEX2, False, True
    Set rubros = DAORubros.FindAll
    Me.GridEX1.ItemCount = rubros.count
    Set rubroElegido = rubros.item(Me.GridEX1.row)
    Set grupos = DAOGrupos.GetAllByRubro(rubroElegido.id)
    Me.GridEX2.ItemCount = grupos.count
End Sub


Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_SelectionChange()
    Me.GridEX2.ItemCount = 0
    Set rubroElegido = rubros.item(Me.GridEX1.RowIndex(Me.GridEX1.row))
    Set grupos = DAOGrupos.GetAllByRubro(rubroElegido.id)
    Me.GridEX2.ItemCount = grupos.count
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpRubro = rubros.item(RowIndex)
    Values(1) = tmpRubro.iniciales
    Values(2) = tmpRubro.Rubro
End Sub

Private Sub GridEX2_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    If MsgBox("¿Está seguro de modificar?", vbYesNo, "Confirmación") = vbNo Then Cancel = True


End Sub

Private Sub GridEX2_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX2, Column
End Sub

Private Sub GridEX2_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set tmpGrupo = New clsGrupo
    tmpGrupo.Grupo = Values(1)
    tmpGrupo.id = 0
    tmpGrupo.rubros = rubroElegido

    If DAOGrupos.Save(tmpGrupo) Then
        MsgBox "Alta exitosa!", vbInformation, "Información"
        grupos.Add tmpGrupo
    End If
End Sub
Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set tmpGrupo = grupos.item(RowIndex)
    Values(1) = tmpGrupo.Grupo
    Exit Sub
err1:
End Sub



Private Sub GridEX2_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpGrupo = grupos.item(RowIndex)
    tmpGrupo.Grupo = Values(1)
    If DAOGrupos.Save(tmpGrupo) Then
        MsgBox "Modificación Exitosa!", vbInformation, "Información"
    End If
End Sub
