VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMaterialesAlmacenes 
   BackColor       =   &H00FF8080&
   Caption         =   "Almacenes"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   5805
   Begin GridEX20.GridEX GridEX1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5741
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
      Column(1)       =   "frmMaterialesAlmacenes.frx":0000
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmMaterialesAlmacenes.frx":00D0
      FormatStyle(2)  =   "frmMaterialesAlmacenes.frx":0208
      FormatStyle(3)  =   "frmMaterialesAlmacenes.frx":02B8
      FormatStyle(4)  =   "frmMaterialesAlmacenes.frx":036C
      FormatStyle(5)  =   "frmMaterialesAlmacenes.frx":0444
      FormatStyle(6)  =   "frmMaterialesAlmacenes.frx":04FC
      ImageCount      =   0
      PrinterProperties=   "frmMaterialesAlmacenes.frx":05DC
   End
End
Attribute VB_Name = "frmMaterialesAlmacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim almacenes As Collection
Dim tmpAlmacen As clsAlmacen
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, , True
    Set almacenes = DAOAlmacenes.GetAll
    Me.GridEX1.ItemCount = almacenes.count
End Sub
Private Sub Form_Resize()
    Me.GridEX1.Height = Me.ScaleHeight
    Me.GridEX1.Width = Me.ScaleWidth
End Sub


Private Sub GridEX1_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    If MsgBox("¿Está seguro de actualizar?", vbYesNo, "Confirmación") = vbNo Then Cancel = True
End Sub

Private Sub GridEX1_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    Set tmpAlmacen = New clsAlmacen
    tmpAlmacen.Id = 0
    tmpAlmacen.almacen = Values(1)
    almacenes.Add tmpAlmacen
    If DAOAlmacenes.Save(tmpAlmacen) Then MsgBox "Alta exitosa!", vbInformation, "Información"
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpAlmacen = almacenes.item(RowIndex)
    Values(1) = tmpAlmacen.almacen
End Sub


Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpAlmacen = almacenes.item(RowIndex)
    tmpAlmacen.almacen = Values(1)
    If DAOAlmacenes.Save(tmpAlmacen) Then MsgBox "Actualización exitosa!", vbInformation, "Información"
End Sub
