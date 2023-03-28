VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmComprasProveedoresRubros 
   BackColor       =   &H00FF8080&
   Caption         =   "Rubros"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   6390
   ClipControls    =   0   'False
   Icon            =   "frmRubrosProveedores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   6390
   Begin GridEX20.GridEX GridEX1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16761024
      DataMode        =   99
      HeaderFontBold  =   -1  'True
      HeaderFontWeight=   700
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmRubrosProveedores.frx":000C
      Column(2)       =   "frmRubrosProveedores.frx":00E0
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmRubrosProveedores.frx":01AC
      FormatStyle(2)  =   "frmRubrosProveedores.frx":02E4
      FormatStyle(3)  =   "frmRubrosProveedores.frx":0394
      FormatStyle(4)  =   "frmRubrosProveedores.frx":0448
      FormatStyle(5)  =   "frmRubrosProveedores.frx":0520
      FormatStyle(6)  =   "frmRubrosProveedores.frx":05D8
      ImageCount      =   0
      PrinterProperties=   "frmRubrosProveedores.frx":06B8
   End
End
Attribute VB_Name = "frmComprasProveedoresRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rubros As Collection
Dim tmpRubro As clsRubros
'Private Sub Command2_Click()
'    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
'        Unload Me
'    End If
'End Sub

Private Sub Form_Activate()
    Me.GridEX1.Refresh
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, True
    Me.GridEX1.ItemCount = 0
    Set rubros = DAORubros.FindAll
    Me.GridEX1.ItemCount = rubros.count

    ''Me.caption = caption & " (" & Name & ")"


End Sub

Private Sub Form_Resize()
    Me.GridEX1.Width = Me.ScaleWidth
    Me.GridEX1.Height = Me.ScaleHeight
End Sub

Private Sub GridEX1_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Len(Me.GridEX1.value(1)) < 2 Or Len(Me.GridEX1.value(1)) > 5
    If Cancel Then MsgBox "La inicial debe contener de 2 a 5 caracteres.", vbExclamation + vbOKOnly
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    tmpRubro.Id = 0
    tmpRubro.iniciales = Values(1)
    tmpRubro.rubro = Values(2)
    rubros.Add tmpRubro
    If DAORubros.Save(tmpRubro) Then
        MsgBox "Alta exitosa!", vbInformation, "Información"
    End If
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpRubro = rubros.item(RowIndex)
    Values(1) = tmpRubro.iniciales
    Values(2) = tmpRubro.rubro
End Sub

Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpRubro = rubros.item(RowIndex)
    tmpRubro.iniciales = Values(1)
    tmpRubro.rubro = Values(2)
    If DAORubros.Save(tmpRubro) Then
        MsgBox "Actualización exitosa!", vbInformation, "Información"
    End If
End Sub
