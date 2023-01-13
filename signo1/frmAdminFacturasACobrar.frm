VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAdminFacturasACobrar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Facturas a Cobrar"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   4530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Usar"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstFacturas 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Factura"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cliente"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmAdminFacturasACobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New Recordset
Dim vIdCliente As Long
Dim clasea As New classAdministracion
Public Property Let idCliente(nIdCliente As Long)
    vIdCliente = nIdCliente
End Property

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    funciones.idFactura = CLng(Me.lstFacturas.selectedItem.Tag)
    Unload Me

End Sub

Private Sub Form_Activate()

    Me.lstFacturas.SetFocus
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    llenarLST
    
    'Me.caption = caption & "(" & Name & ")"
        

End Sub

Private Sub llenarLST()
    Me.lstFacturas.ListItems.Clear
    Dim x As ListItem
    Set rs = conectar.RSFactory("select f.saldada,f.id,f.nroFactura,c.razon,f.fechaEmision from AdminFacturas f inner join clientes c on f.idCliente=c.id where f.estado=2 and (f.saldada=0 or f.saldada=2 or f.saldada=3 or f.saldada=4)  and  f.idCliente=" & vIdCliente)
    While Not rs.EOF
        Set x = Me.lstFacturas.ListItems.Add(, , Format(rs!nroFactura, "0000"))
        x.Tag = rs!id
        x.SubItems(1) = rs!razon
        x.SubItems(2) = Format(rs!FechaEmision, "dd-mm-yyyy")
        rs.MoveNext
    Wend

End Sub

Private Sub lstFacturas_DblClick()
    'funciones.idFactura = CLng(Me.lstFacturas.SelectedItem)
    'Unload Me
    Command2_Click
End Sub
