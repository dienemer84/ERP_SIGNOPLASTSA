VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmComprasPreciosHistorico 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Historicos de materiales..."
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11085
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   Begin GridEX20.GridEX GridEX1 
      Height          =   5295
      Left            =   7560
      TabIndex        =   2
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9340
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowColumnDrag =   0   'False
      GroupByBoxInfoText=   ""
      BackColorGBBox  =   16744576
      BackColorHeader =   16761024
      DataMode        =   99
      BackColorBkg    =   16777215
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmHistorico.frx":0000
      Column(2)       =   "frmHistorico.frx":00F0
      Column(3)       =   "frmHistorico.frx":01F4
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmHistorico.frx":02C4
      FormatStyle(2)  =   "frmHistorico.frx":03FC
      FormatStyle(3)  =   "frmHistorico.frx":04AC
      FormatStyle(4)  =   "frmHistorico.frx":0560
      FormatStyle(5)  =   "frmHistorico.frx":0638
      FormatStyle(6)  =   "frmHistorico.frx":06F0
      ImageCount      =   0
      PrinterProperties=   "frmHistorico.frx":07D0
   End
   Begin GridEX20.GridEX grilla_material 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9340
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      GroupByBoxInfoText=   "Arrastrar una columna para agrupar"
      AllowEdit       =   0   'False
      BackColorGBBox  =   16744576
      BackColorHeader =   16761024
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   5
      Column(1)       =   "frmHistorico.frx":09A8
      Column(2)       =   "frmHistorico.frx":0A9C
      Column(3)       =   "frmHistorico.frx":0B68
      Column(4)       =   "frmHistorico.frx":0C34
      Column(5)       =   "frmHistorico.frx":0D0C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmHistorico.frx":0DDC
      FormatStyle(2)  =   "frmHistorico.frx":0F14
      FormatStyle(3)  =   "frmHistorico.frx":0FC4
      FormatStyle(4)  =   "frmHistorico.frx":1078
      FormatStyle(5)  =   "frmHistorico.frx":1150
      FormatStyle(6)  =   "frmHistorico.frx":1208
      ImageCount      =   0
      PrinterProperties=   "frmHistorico.frx":12E8
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   975
   End
End
Attribute VB_Name = "frmComprasPreciosHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Materiales As Collection
Dim tmpMaterial As clsMaterial
Dim historicos As Collection
Dim tmpHistorico As clsMaterialHistorico
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.GridEX1.Refresh
    Me.grilla_material.Refresh
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1
    GridEXHelper.CustomizeGrid Me.grilla_material, True

    Me.grilla_material.ItemCount = 0
    Set Materiales = DAOMateriales.FindAll()
    Me.grilla_material.ItemCount = Materiales.count
    aRow = Me.grilla_material.RowIndex(Me.grilla_material.row)
    Set historicos = Materiales.item(aRow).historico
    
    ''Me.caption = caption & " (" & Name & ")"
        
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If Not historicos Is Nothing Then
        Set tmpHistorico = historicos.item(RowIndex)
        Values(1) = tmpHistorico.FEcha
        Values(2) = funciones.FormatearDecimales(tmpHistorico.Valor, 2)
        Values(3) = tmpHistorico.moneda.NombreCorto
    End If
End Sub

Private Sub grilla_material_SelectionChange()
    Me.GridEX1.ItemCount = 0
    Set historicos = DAOMaterialHistorico.getAllByMaterial(tmpMaterial.id)
    Me.GridEX1.ItemCount = historicos.count
End Sub

Private Sub grilla_material_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set tmpMaterial = Materiales.item(RowIndex)
    With tmpMaterial
        Values(1) = .codigo
        Values(2) = .Grupo.rubros.rubro
        Values(3) = .Grupo.Grupo
        Values(4) = .descripcion
        Values(5) = .Espesor
    End With

End Sub


