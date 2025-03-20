VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasEditarDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Datos Cargados en Comprobante"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   8775
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   495
         Left            =   7080
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   495
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reestablecer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cancelar"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin XtremeSuiteControls.PushButton PushButton7 
         Height          =   375
         Left            =   8040
         TabIndex        =   17
         Top             =   3600
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   375
         Left            =   8040
         TabIndex        =   16
         Top             =   2760
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   8040
         TabIndex        =   15
         Top             =   1920
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   375
         Left            =   8040
         TabIndex        =   14
         Top             =   1080
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox Text6 
         Height          =   2295
         Left            =   240
         TabIndex        =   4
         Text            =   "Text6"
         Top             =   3600
         Width           =   7695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   7695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1920
         Width           =   7695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   2760
         Width           =   7695
      End
      Begin VB.Label lblTextoAdicional 
         Caption         =   "Orden de Compra / Referencia:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label lblTextoAdicional 
         Caption         =   "Observaciones 2 / Aplicación / Cancelación:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label lblTextoAdicional 
         Caption         =   "Observaciones 1 / Condición:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2775
      End
      Begin VB.Line Line1 
         X1              =   7920
         X2              =   240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblTextoAdicional 
         Caption         =   "Texto Adicional:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   2535
      End
      Begin XtremeSuiteControls.Label lblNumeroCbte 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   7695
         _Version        =   786432
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Datos de CBTE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAdminFacturasEditarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Factura As Factura


Public Property Let idFactura(value As Long)
    Set Factura = DAOFactura.FindById(value, True, True)

End Property

Private Sub Form_Load()

    Customize Me

    Me.lblNumeroCbte.caption = "N° de Cbte: " & Factura.NumeroFormateado & "- " & Factura.cliente.razon

    Me.Text1.Text = Factura.observaciones
    Me.Text2.Text = Factura.observaciones_cancela
    Me.Text3.Text = Factura.OrdenCompra
    Me.Text6.Text = Factura.TextoAdicional


End Sub

'BOTÓN CERRAR
Private Sub PushButton1_Click()
    Unload Me

End Sub

'BOTÓN REESTABLECER
Private Sub PushButton2_Click()
    Me.Text1.Text = Factura.observaciones
    Me.Text2.Text = Factura.observaciones_cancela
    Me.Text3.Text = Factura.OrdenCompra
    Me.Text6.Text = Factura.TextoAdicional

End Sub

'BOTÓN GUARDAR CAMBIOS
Private Sub PushButton3_Click()
    If MsgBox("Está segur@ de los cambios realizados?", vbYesNo, "Confirmación") = vbYes Then

        Factura.observaciones = Me.Text1.Text
        Factura.observaciones_cancela = Me.Text2.Text
        Factura.OrdenCompra = Me.Text3.Text
        Factura.TextoAdicional = Me.Text6.Text

        If DAOFactura.Save(Factura, True) Then
            MsgBox "Los datos del comprobante han sido actualizados.", vbOKOnly + vbInformation
            Unload Me
        Else
            Err.Raise "9999", "Guardando factura", Err.Description
        End If
    End If

End Sub


Private Sub PushButton4_Click()
    Me.Text2.Text = ""
End Sub

Private Sub PushButton5_Click()
    Me.Text1.Text = ""
End Sub

Private Sub PushButton6_Click()
    Me.Text3.Text = ""
End Sub

Private Sub PushButton7_Click()
    Me.Text6.Text = ""
End Sub
