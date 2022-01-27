VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmAdminFacturasAplicadas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas aplicadas..."
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4920
   Begin MSComctlLib.ListView lstFacturas 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7858
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
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4800
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
            Picture         =   "frmAdminFacturasAplicadas.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   7080
      Width           =   735
   End
End
Attribute VB_Name = "frmAdminFacturasAplicadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim vorigen As Integer
Dim vId As Long
Dim clasea As New classAdministracion
Public Property Let Origen(nOrigen As Integer)
    vorigen = nOrigen
End Property
Public Property Let idOrigen(nId As Long)
    vId = nId
End Property
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.lstFacturas.Refresh
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Set rs = clasea.facturasEntregadas(vorigen, vId)
    While Not rs.EOF
        Dim x As ListItem
        Set x = Me.lstFacturas.ListItems.Add(, , rs!Factura, 1)
        x.Tag = rs!Id

        rs.MoveNext
    Wend
End Sub

Private Sub lstFacturas_DblClick()
    If Me.lstFacturas.ListItems.count > 0 Then
        idf = CLng(Me.lstFacturas.selectedItem.Tag)
        Dim f_c3h3 As New frmAdminFacturasEdicion
        f_c3h3.ReadOnly = True
        f_c3h3.idFactura = idf
        f_c3h3.Show

    End If
End Sub

