VERSION 5.00
Begin VB.Form frmAdminFacturaRechazoAfip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rechazo de factura"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CheckBox chkRechazo 
      Caption         =   "¿Comprobante rechazado?"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Motivos"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmAdminFacturaRechazoAfip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Factura As Factura

Private Sub cmdGuardar_Click()
On Error GoTo errCae

If MsgBox("¿Está seguro de actualizar el estado de rechazo del comprobante?", vbYesNo, "Confirmación") = vbYes Then

    Factura.MotivosAnulacionAFIP = Me.Text1.text
    If Me.chkRechazo.value = 1 Then
    ' 154541541
    ' 21/10/2022
    ' SE CAMBIA EL VALOR DE Y POR S (EL VALOR DE Y ESTABA DANDO ERROR ULTIMAMENTE
    ' Factura.AnulacionAFIP = "Y"
        Factura.AnulacionAFIP = "S"
    Else
        Factura.AnulacionAFIP = "N"
   End If
    DAOFactura.RechazoAfip Factura
    MsgBox "Datos de estado de rechazo actualizados correctamente", vbInformation, "Proceso correcto"
   Unload Me
End If
Exit Sub

errCae:
 MsgBox Err.Description, vbCritical, "Se produjo un error"
End Sub

Private Sub Form_Load()
  FormHelper.Customize Me
  
  If IsSomething(Factura) Then
         Set Factura = DAOFactura.FindById(Factura.Id)
       If Factura.AnulacionAFIP = "Y" Then
          
            Me.chkRechazo.value = 1
        Else
        Me.chkRechazo.value = 0
        End If
        Me.Text1.text = Factura.MotivosAnulacionAFIP
    End If
End Sub

