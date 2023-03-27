VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmListarStock_seleccion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elegir stock..."
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10560
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCliente 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   4320
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Filtrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstStock 
      Height          =   4215
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Cliente"
         Object.Width           =   11024
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Cantidad"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "estado"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label P 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label marcado 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblVisible 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   5400
      Width           =   735
   End
End
Attribute VB_Name = "frmListarStock_seleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As classStock
Private Sub Command1_Click()
    base.llenar_lista_stock Me.lstStock, Me.cboCliente.ItemData(Me.cboCliente.ListIndex), Trim(Text1), marcado, , True
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Text1 = funciones.quePiezaElegidabusqueda
    base.llenar_lista_stock Me.lstStock, Me.cboCliente.ItemData(Me.cboCliente.ListIndex), Trim(Text1), marcadom, , True

End Sub
Private Sub Form_Load()
    FormHelper.Customize Me
    Me.lblVisible = 0
    Set base = New classStock
    marcado = -1

    DAOCliente.LlenarCombo Me.cboCliente, True
    'base.llenar_lista_stock Me.lstStock, -1, ""

    ''Me.caption = caption & " (" & Name & ")"


End Sub
Private Sub lstStock_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstStock.Sorted = True
    LstOrdenar (CInt(ColumnHeader.Index))
End Sub
Function LstOrdenar(columna As Integer)
    lstStock.SortKey = columna - 1
    If lstStock.SortOrder = lvwAscending Then
        lstStock.SortOrder = lvwDescending
    Else
        lstStock.SortOrder = lvwAscending
    End If
End Function
Private Sub lstStock_DblClick()
    If Me.lstStock.ListItems.count > 0 Then
        funciones.quePiezaElegidabusqueda = Me.Text1
        funciones.quePiezaElegida = CLng(Me.lstStock.selectedItem)
        funciones.quePiezaElegidaDetalle = Me.lstStock.selectedItem.ListSubItems(2)
        Unload Me
        'aca va el evento1
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub
