VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasProformasEditarDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Datos Cargados en Proforma"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8775
      Begin VB.TextBox Text 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Text            =   "Text"
         Top             =   2760
         Width           =   7695
      End
      Begin VB.TextBox Text 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Text            =   "Text"
         Top             =   1920
         Width           =   7695
      End
      Begin VB.TextBox Text 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Text            =   "Text"
         Top             =   1080
         Width           =   7695
      End
      Begin VB.TextBox Text 
         Height          =   2295
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Text            =   "Text"
         Top             =   3600
         Width           =   7695
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   3
         Left            =   8040
         TabIndex        =   4
         Top             =   3600
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   4
         Left            =   8040
         TabIndex        =   5
         Top             =   2760
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   5
         Left            =   8040
         TabIndex        =   6
         Top             =   1920
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   6
         Left            =   8040
         TabIndex        =   7
         Top             =   1080
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNumeroCbte 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   7695
         _Version        =   786432
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Datos de PROFORMA"
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
         TabIndex        =   15
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Line Line 
         X1              =   7920
         X2              =   240
         Y1              =   600
         Y2              =   600
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
         TabIndex        =   14
         Top             =   840
         Width           =   2775
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
         TabIndex        =   13
         Top             =   1680
         Width           =   4215
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
         TabIndex        =   12
         Top             =   2520
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   8775
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   495
         Index           =   0
         Left            =   7080
         TabIndex        =   1
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
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   495
         Index           =   1
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reestablecer"
         UseVisualStyle  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAdminFacturasProformasEditarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FacturaProforma As clsFacturaProforma


Public Property Let idFactura(value As Long)
    Set FacturaProforma = DAOFacturaProforma.FindById(value, True, True)

End Property

Private Sub Form_Load()

    Customize Me

    Me.lblNumeroCbte.caption = "N° de PROFORMA: " & FacturaProforma.NumeroFormateado & "- " & FacturaProforma.cliente.razon

    Me.Text(1).Text = FacturaProforma.observaciones
    Me.Text(2).Text = FacturaProforma.observaciones_cancela
    Me.Text(3).Text = FacturaProforma.OrdenCompra
    Me.Text(0).Text = FacturaProforma.TextoAdicional


End Sub


Private Sub PushButton_Click(Index As Integer)
If Index = 6 Then
    Me.Text(1) = ""
    
ElseIf Index = 0 Then
        If MsgBox("Está segur@ de los cambios realizados?", vbYesNo, "Confirmación") = vbYes Then
    
            FacturaProforma.observaciones = Me.Text(1).Text
            FacturaProforma.observaciones_cancela = Me.Text(2).Text
            FacturaProforma.OrdenCompra = Me.Text(3).Text
            FacturaProforma.TextoAdicional = Me.Text(0).Text
    
            If DAOFacturaProforma.Save(FacturaProforma, True) Then
                MsgBox "Los datos del comprobante han sido actualizados.", vbOKOnly + vbInformation
                Unload Me
            Else
                Err.Raise "9999", "Guardando factura", Err.Description
            End If
        End If
    
ElseIf Index = 1 Then
        Me.Text(1).Text = FacturaProforma.observaciones
        Me.Text(2).Text = FacturaProforma.observaciones_cancela
        Me.Text(3).Text = FacturaProforma.OrdenCompra
        Me.Text(0).Text = FacturaProforma.TextoAdicional

ElseIf Index = 2 Then
        Unload Me
    
ElseIf Index = 3 Then
        Me.Text(0) = ""
        
ElseIf Index = 4 Then
        Me.Text(3) = ""
        
ElseIf Index = 5 Then
        Me.Text(2) = ""
        
ElseIf Index = 6 Then
        Me.Text(1) = ""
End If

End Sub
