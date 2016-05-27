VERSION 5.00
Begin VB.Form frmAdminFacturasAplicarRecibo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplicar recibo por pago a cuenta..."
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      Default         =   -1  'True
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtRecibo 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtTotalAplicado 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblTotalRecibo 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblTotalFactura 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblFactura 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total aplicado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total recibo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Recibo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdminFacturasAplicarRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totRecibo As Double
Dim totFactura As Double
Dim idRecibo As Long
Dim idMonedaFactura As Integer
Dim idMonedaRecibo As Integer
Dim RS As recordset
Dim claseA As New classAdministracion
Dim vIdFactura As Long

Public Property Let idFactura(nIdFactura As Long)
    vIdFactura = nIdFactura
End Property

Private Sub Command1_Click()
    totalAplicado = CDbl(Me.txtTotalAplicado)

    If totalAplicado > totRecibo Then
        MsgBox "No puede aplicar un valor mayor al total del recibo!", vbCritical, "Error"
    Else
        If totalAplicado > totFactura Then
            MsgBox "No puede aplicar un valor mayor que el total de la factura!", vbCritical, "Error"
        Else
            'si es apto, se aplica
            If MsgBox("¿Confirma la aplicación del recibo a esta factura?", vbYesNo, "Confirmacion") = vbYes Then
                If claseA.aplicarRecibo(idRecibo, vIdFactura, totalAplicado) Then
                    MsgBox "Aplicación de recibo exitosa!", vbInformation, "Información"
                Else
                    MsgBox "Se produjo algún error, no se aplicará el recibo!", vbCritical, "Error"
                End If
            End If
        End If
    End If



End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    Dim Total As Double
    Dim idMon As Integer
    Dim rz As String
    Dim idcli As Long
    'nroFC = claseA.queFactura(vIdFactura)
    Me.lblFactura = nroFC
    claseA.TotalFactura vIdFactura, Total, idMonedaFactura, rz, idcli
    totFactura = funciones.FormatearDecimales(Total, 2)
    Me.lblTotalFactura = funciones.queMoneda(idMon) & " " & totFactura

End Sub


Private Sub txtRecibo_DblClick()
    Dim a As Long
    Dim cta As Integer

    frmAdminCobranzasElegirRecibo.Show 1
    a = funciones.idReciboElegido
    Set RS = conectar.RSFactory("select idMoneda from AdminRecibos where id=" & a)
    If Not RS.EOF And Not RS.BOF Then
        idMonedaRecibo = RS!IdMoneda
    Else
        Exit Sub
    End If
    If a = -1 Then
        Me.txtRecibo = Empty
        Me.lblTotalRecibo = 0
        idRecibo = 0
    Else
        Me.txtRecibo = a
        idRecibo = a
        totRecibo = funciones.FormatearDecimales(claseA.totalRecibo2(a, 1), 2)
        Me.lblTotalRecibo = funciones.queMoneda(idMonedaRecibo) & " " & funciones.FormatearDecimales(totRecibo, 2)
    End If
    If idMonedaRecibo <> idMonedaFactura Then
        MsgBox "No puede aplicar un recibo a una factura que sea en otra moneda!", vbCritical, "Error"
        Me.Command1.Enabled = False
    Else
        If totFactura >= totRecibo Then
            Me.txtTotalAplicado = funciones.FormatearDecimales(totRecibo, 2)
        Else
            Me.txtTotalAplicado = funciones.FormatearDecimales(totFactura, 2)
        End If
        Me.Command1.Enabled = True
    End If


    Set RS = Nothing
End Sub
