VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprasElegirProveedor2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elegir proveedor..."
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdUsar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usar"
      Height          =   375
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   4320
      Width           =   4335
   End
   Begin VB.ComboBox cboRubro 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstProveedores 
      Height          =   4215
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label P 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rubros"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   495
   End
End
Attribute VB_Name = "frmComprasElegirProveedor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim marca
Dim id_rubro As Integer
Dim baseP As New classCompras
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdUsar_Click()
funciones.idPRoveedorElegido = CLng(Me.lstProveedores.SelectedItem)
Unload Me
End Sub
Private Sub Command1_Click()
id_rubro = Me.cboRubro.ItemData(cboRubro.ListIndex)
'baseP.llenar_lista_proveedores Me.lstProveedores, id_rubro, Me.Text1, marca
End Sub
Private Sub con_tacto_Click()
If Me.lstProveedores.ListItems.count > 0 Then
    idproveedor = CLng(Me.lstProveedores.SelectedItem)
    frmComprasProveedoresNuevoContacto.idproveedor = idproveedor
    frmComprasProveedoresNuevoContacto.Show
End If
End Sub
Private Sub estado_Click()
marca = Me.lstProveedores.SelectedItem
g = MsgBox("¿Seguro que desea cambiar el estado del proveedor seleccionado?", vbYesNo, "Confirmacion")
If g = 6 Then
baseP.cambiar_estado CInt(Me.lstProveedores.SelectedItem), CInt(Me.lstProveedores.SelectedItem.ListSubItems(14))
End If
id_rubro = Me.cboRubro.ItemData(cboRubro.ListIndex)
'baseP.llenar_lista_proveedores Me.lstProveedores, id_rubro, Me.Text1, marca
End Sub
Private Sub Command2_Click()
funciones.idPRoveedorElegido = -1
Unload Me
End Sub
Private Sub Form_Load()
marca = 0
DAORubros.llenarCombo Me.cboRubro
id_rubro = Me.cboRubro.ItemData(cboRubro.ListIndex)
'baseP.llenar_lista_proveedores Me.lstProveedores, id_rubro, Me.Text1, marca
End Sub
Function LstOrdenar(columna As Integer)
Me.lstProveedores.SortKey = columna - 1
If Me.lstProveedores.SortOrder = lvwAscending Then
    Me.lstProveedores.SortOrder = lvwDescending
        Else
    Me.lstProveedores.SortOrder = lvwAscending
End If
End Function
Private Sub lstProveedores_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.lstProveedores.Sorted = True
LstOrdenar (CInt(ColumnHeader.Index))
End Sub
Private Sub lstProveedores_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim li As ListItem
Set li = Me.lstProveedores.HitTest(x, Y)
If li Is Nothing Then
    Me.lstProveedores.ToolTipText = ""
Else
    Me.lstProveedores.ToolTipText = li.Tag
End If

End Sub
Private Sub verDetalles_Click()
If Me.lstProveedores.ListItems.count > 0 Then
    marca = Me.lstProveedores.SelectedItem
    frmComprasProveedoresModifica.idproveedor = Me.lstProveedores.SelectedItem
    frmComprasProveedoresModifica.Show
End If

End Sub

Private Sub verRubros_Click()
frmVerRubros.lblid = Me.lstProveedores.SelectedItem
marca = Me.lstProveedores.SelectedItem
frmVerRubros.Caption = "[ " & truncar(Me.lstProveedores.SelectedItem.ListSubItems(2), 20) & " ]"
frmVerRubros.Show
End Sub

