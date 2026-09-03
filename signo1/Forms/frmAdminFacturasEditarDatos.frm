VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasEditarDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar Datos Cargados en Comprobante"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   6840
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
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin XtremeSuiteControls.ComboBox cboProvincias 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   6240
         Width           =   3495
         _Version        =   786432
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
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
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmAdminFacturasEditarDatos.frx":0000
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
         Caption         =   "Provincia/Jurisdicción:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   6000
         Width           =   2535
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

Public Property Let idFactura(ByVal value As Long)

    Set Factura = DAOFactura.FindById(value, True, True)

End Property

Private Sub Form_Load()

    Customize Me

    If Factura Is Nothing Then
        MsgBox "No se pudo cargar el comprobante.", vbCritical, "Error"
        Unload Me
        Exit Sub
    End If

    Me.lblNumeroCbte.caption = _
        "N° de Cbte: " & Factura.NumeroFormateado & " - " & Factura.Cliente.razon

    CargarDatos

End Sub

Private Sub CargarDatos()

    Me.Text1.Text = Factura.Observaciones
    Me.Text2.Text = Factura.observaciones_cancela
    Me.Text3.Text = Factura.OrdenCompra
    Me.Text6.Text = Factura.TextoAdicional

    DAOProvincias.LlenarComboNoDefinido Me.cboProvincias, 1, True

    Me.cboProvincias.ListIndex = _
        funciones.PosIndexCbo(Factura.idProvincia, Me.cboProvincias)

    If Me.cboProvincias.ListIndex < 0 Then
        Me.cboProvincias.ListIndex = _
            funciones.PosIndexCbo(0, Me.cboProvincias)
    End If

End Sub

'BOTÓN CERRAR
Private Sub PushButton1_Click()

    Unload Me

End Sub

'BOTÓN REESTABLECER
Private Sub PushButton2_Click()

    Me.Text1.Text = Factura.Observaciones
    Me.Text2.Text = Factura.observaciones_cancela
    Me.Text3.Text = Factura.OrdenCompra
    Me.Text6.Text = Factura.TextoAdicional

    Me.cboProvincias.ListIndex = _
        funciones.PosIndexCbo(Factura.idProvincia, Me.cboProvincias)

    If Me.cboProvincias.ListIndex < 0 Then
        Me.cboProvincias.ListIndex = _
            funciones.PosIndexCbo(0, Me.cboProvincias)
    End If

End Sub

'BOTÓN GUARDAR CAMBIOS
Private Sub PushButton3_Click()

    On Error GoTo ManejarError

    If MsgBox( _
        "¿Está seguro de guardar los cambios realizados?", _
        vbYesNo + vbQuestion, _
        "Confirmación") <> vbYes Then

        Exit Sub
    End If

    Factura.Observaciones = Me.Text1.Text
    Factura.observaciones_cancela = Me.Text2.Text
    Factura.OrdenCompra = Me.Text3.Text
    Factura.TextoAdicional = Me.Text6.Text

    If Me.cboProvincias.ListIndex >= 0 Then
        Factura.idProvincia = _
            Me.cboProvincias.ItemData(Me.cboProvincias.ListIndex)
    Else
        Factura.idProvincia = 0
    End If

    If Not DAOFactura.Save(Factura, True) Then
        Err.Raise vbObjectError + 1000, _
                  "Guardando factura", _
                  "No se pudo guardar el comprobante."
    End If

    MsgBox "Los datos del comprobante han sido actualizados.", _
           vbInformation, _
           "Comprobante actualizado"

    Unload Me
    Exit Sub

ManejarError:
    MsgBox "Ocurrió un error al guardar el comprobante:" & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Error"

End Sub

Private Sub PushButton4_Click()
    Me.Text2.Text = vbNullString
End Sub

Private Sub PushButton5_Click()
    Me.Text1.Text = vbNullString
End Sub

Private Sub PushButton6_Click()
    Me.Text3.Text = vbNullString
End Sub

Private Sub PushButton7_Click()
    Me.Text6.Text = vbNullString
End Sub

