VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdminCobranzasElegirRecibo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elegir reciibo..."
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin MSComctlLib.ListView lstRecibos 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminCobranzasElegirRecibo.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdminCobranzasElegirRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
'Dim clase As New classAdministracion
Private Sub Command1_Click()
    funciones.idReciboElegido = -1
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    'solo pago a cuenta
    Dim x As ListItem
    Set rs = conectar.RSFactory("select id,fechaCreacion as fecha from AdminRecibos where pagoACuenta=1 and (todo_aplicado=0 or todo_aplicado=1)")
    While Not rs.EOF
        Set x = Me.lstRecibos.ListItems.Add(, , Format(rs!Id, "0000"), 1)
        x.SubItems(1) = Format(rs!FEcha, "dd-mm-yyyy")

        rs.MoveNext
    Wend
    Set rs = Nothing

End Sub

Private Sub Form_Terminate()
    funciones.idReciboElegido = -1
End Sub


Private Sub lstRecibos_DblClick()
    If Me.lstRecibos.ListItems.count > 0 Then
        funciones.idReciboElegido = CLng(Me.lstRecibos.selectedItem)
        Unload Me
    End If

End Sub
