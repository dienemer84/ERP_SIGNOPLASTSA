VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComprasRequesProveedoresElegidos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proveedores"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quitar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Agregar"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin MSComctlLib.ListView lstProveedoresElegidos 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proveedor"
         Object.Width           =   11774
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "frmComprasRequesProveedoresElegidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As recordset
Dim compras As New classCompras
Dim vConcepto As Boolean
Dim vIdRequeDetalle As Long
Public Property Let concepto(nConcepto As Boolean)
vConcepto = nConcepto
End Property


Public Property Let idRequeDetalle(nIdRequeDetalle As Long)
vIdRequeDetalle = nIdRequeDetalle
End Property

Private Sub Command1_Click()
If MsgBox("¿Desea actualizar los datos?", vbYesNo, "Confirmación") = vbYes Then
    'actualizar datos de proveedores
    If compras.ActualizarProveedoresReque(vIdRequeDetalle, Me.lstProveedoresElegidos) Then
        MsgBox "Actualizacion exitosa!", vbExclamation, "Información"
        Unload Me
    Else
        MsgBox "Se produjo algun error!", vbCritical, "Error"
    End If
    
    
Else
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Dim RS As recordset
'id_p = compras.Elegir_proveedor()


'busco que no este
esta = False
For x = 1 To Me.lstProveedoresElegidos.ListItems.count
If id_p = Me.lstProveedoresElegidos.ListItems(x).Tag Then
  esta = True
End If
Next x

If Not esta Then
If id_p > 0 Then
    Set RS = conectar.RSFactory("SELECT razon FROM proveedores where id=" & id_p)
    If Not RS.EOF And Not RS.BOF Then
    Set x = Me.lstProveedoresElegidos.ListItems.Add(, , UCase(RS!Razon))
        x.Tag = id_p
    End If
End If
Else
MsgBox "El proveedor elegido ya fue seleccionado!", vbInformation, "Información"
End If
End Sub

Private Sub Command3_Click()

For I = Me.lstProveedoresElegidos.ListItems.count To 1 Step -1
If Me.lstProveedoresElegidos.ListItems(I).Checked = True Then
 Me.lstProveedoresElegidos.ListItems.Remove (I)
End If
Next I


End Sub

Private Sub Form_Load()

Set RS = conectar.RSFactory("select p.id,p.razon from ComprasRequerimientosProveedores c inner join proveedores p on c.idProveedor=p.id where idReque=" & vIdRequeDetalle)
While Not RS.EOF
    Set x = Me.lstProveedoresElegidos.ListItems.Add(, , RS!Razon)
        x.Tag = RS!id
RS.MoveNext
Wend
End Sub
