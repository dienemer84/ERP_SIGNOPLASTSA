VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAgendaNueva 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directorio de contactos"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   6015
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   12855
      _Version        =   786432
      _ExtentX        =   22675
      _ExtentY        =   10610
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX dgDatos 
         Height          =   5415
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   9551
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         ColumnAutoResize=   -1  'True
         ReadOnly        =   -1  'True
         MethodHoldFields=   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   5
         Column(1)       =   "frmAgendaNueva.frx":0000
         Column(2)       =   "frmAgendaNueva.frx":0110
         Column(3)       =   "frmAgendaNueva.frx":020C
         Column(4)       =   "frmAgendaNueva.frx":0308
         Column(5)       =   "frmAgendaNueva.frx":0404
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmAgendaNueva.frx":04F0
         FormatStyle(2)  =   "frmAgendaNueva.frx":0628
         FormatStyle(3)  =   "frmAgendaNueva.frx":06D8
         FormatStyle(4)  =   "frmAgendaNueva.frx":078C
         FormatStyle(5)  =   "frmAgendaNueva.frx":0864
         FormatStyle(6)  =   "frmAgendaNueva.frx":091C
         ImageCount      =   0
         PrinterProperties=   "frmAgendaNueva.frx":09FC
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   135
         Index           =   4
         Left            =   7440
         TabIndex        =   20
         Top             =   240
         Width           =   5175
         _Version        =   786432
         _ExtentX        =   9128
         _ExtentY        =   238
         _StockProps     =   79
         Caption         =   "Doble-click sobre el contacto para abrir el detalle"
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _Version        =   786432
      _ExtentX        =   22675
      _ExtentY        =   3413
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButtonX 
         Height          =   375
         Left            =   9360
         TabIndex        =   18
         Top             =   1200
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   2
         Left            =   9360
         TabIndex        =   17
         Top             =   720
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   16
         Top             =   1200
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton 
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   15
         Top             =   720
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnNuevo 
         Height          =   495
         Left            =   10680
         TabIndex        =   9
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Nuevo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCod 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   1200
         Width           =   3135
         _Version        =   786432
         _ExtentX        =   5530
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtLocalidad 
         Height          =   375
         Left            =   6120
         TabIndex        =   4
         Top             =   720
         Width           =   3135
         _Version        =   786432
         _ExtentX        =   5530
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtDomicilio 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   1200
         Width           =   3255
         _Version        =   786432
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   3255
         _Version        =   786432
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   495
         Left            =   10680
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
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
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   3
         Left            =   4920
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Email"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   13
         Top             =   720
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Localidad"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Domicilio"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nombre"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   10
         Top             =   240
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cod"
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmAgendaNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim contactos As Collection
Dim rectemp As clsContactoPpal

Private Sub btnBuscar_Click()
        llenarGrilla
        
End Sub


Private Sub btnNuevo_Click()
        Dim f12 As New frmAgendaNuevaDetalles
        f12.Show
End Sub


Private Sub dgDatos_DblClick()
    verDeta
End Sub


Private Sub verDeta()
    If Me.dgDatos.rowcount Then
        Set rectemp = contactos(Me.dgDatos.RowIndex(Me.dgDatos.row))
            frmAgendaNuevaDetalles.contacto = rectemp
            frmAgendaNuevaDetalles.Show
    End If
End Sub


Private Sub dgDatos_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    Set rectemp = contactos.item(RowIndex)
    With rectemp
        Values(1) = .id
        Values(2) = .Empresa
        Values(3) = .direccion
        Values(4) = .localidad
        Values(5) = .email
        

    End With
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    
    llenarGrilla
    
    End Sub


Public Function llenarGrilla()

    Dim filtro As String

    If LenB(Me.txtCod.Text) > 0 Then
        filtro = filtro & " AND a.id LIKE '%" & Trim(Me.txtCod.Text) & "%'"
    End If

    If LenB(Me.txtNombre.Text) > 0 Then
        filtro = filtro & " AND a.empresa LIKE '%" & Trim(Me.txtNombre.Text) & "%'"
    End If

    If LenB(Me.txtDomicilio.Text) > 0 Then
        filtro = filtro & " AND a.direccion LIKE '%" & Trim(Me.txtDomicilio.Text) & "%'"
    End If
    
    If LenB(Me.txtLocalidad.Text) > 0 Then
        filtro = filtro & " AND a.localidad LIKE '%" & Trim(Me.txtLocalidad.Text) & "%'"
    End If
    
    If LenB(Me.txtEmail.Text) > 0 Then
        filtro = filtro & " AND a.email LIKE '%" & Trim(Me.txtEmail.Text) & "%'"
    End If
        
    Set contactos = DAOContactoPpal.FindAll(filtro, "a.empresa DESC")

    Me.dgDatos.ItemCount = 0
    
    Me.dgDatos.ItemCount = contactos.count
    
    Me.dgDatos.ReBind

End Function

Private Sub PushButton_Click(Index As Integer)
If Index = 3 Then
    Me.txtCod.Text = ""
ElseIf Index = 0 Then
    Me.txtNombre.Text = ""
ElseIf Index = 1 Then
    Me.txtDomicilio.Text = ""
ElseIf Index = 2 Then
    Me.txtLocalidad.Text = ""
End If
End Sub

Private Sub PushButtonX_Click()
    Me.txtEmail.Text = ""
End Sub
