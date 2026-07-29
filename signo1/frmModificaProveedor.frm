VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmComprasProveedoresModifica 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar proveedor"
   ClientHeight    =   10635
   ClientLeft      =   210
   ClientTop       =   195
   ClientWidth     =   8595
   ClipControls    =   0   'False
   Icon            =   "frmModificaProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1455
      Left            =   120
      TabIndex        =   42
      Top             =   5640
      Width           =   8055
      _Version        =   786432
      _ExtentX        =   14208
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Datos Bancarios"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtCBU 
         Height          =   285
         Left            =   1560
         TabIndex        =   45
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox txtTitularCta 
         Height          =   285
         Left            =   1560
         TabIndex        =   44
         Top             =   960
         Width           =   6375
      End
      Begin VB.TextBox txtAlias 
         Height          =   285
         Left            =   1560
         TabIndex        =   43
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label LabelTitular 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Titular"
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
         TabIndex        =   48
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label ALIAS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Alias"
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
         Left            =   240
         TabIndex        =   47
         Top             =   615
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "CBU"
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
         Left            =   480
         TabIndex        =   46
         Top             =   255
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.PushButton btnVerificarCUIT 
      Height          =   375
      Left            =   4440
      TabIndex        =   41
      Top             =   120
      Width           =   2415
      _Version        =   786432
      _ExtentX        =   4260
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cargar desde ARCA"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboIva 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   6495
      _Version        =   786432
      _ExtentX        =   11456
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin VB.ComboBox cboEstadoProveedor 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   7320
      Width           =   2775
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1830
      Left            =   600
      TabIndex        =   35
      Top             =   8160
      Width           =   7935
      _Version        =   786432
      _ExtentX        =   13996
      _ExtentY        =   3228
      _StockProps     =   79
      Caption         =   "Rubros"
      UseVisualStyle  =   -1  'True
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin MSComctlLib.ListView lstRubros 
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2566
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
            Object.Width           =   5733
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1455
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2566
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
            Object.Width           =   5733
         EndProperty
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF8080&
      Caption         =   "Dólares"
      Height          =   300
      Left            =   5160
      TabIndex        =   15
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Pago contra entrega"
      Height          =   300
      Left            =   6525
      TabIndex        =   16
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1695
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3120
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   3840
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4200
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4560
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1680
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   5280
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4920
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "El cuit va sin guiones!"
      Top             =   135
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2400
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   6495
   End
   Begin XtremeSuiteControls.PushButton cmdPlanCuentas 
      Height          =   375
      Left            =   615
      TabIndex        =   37
      Top             =   10110
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Definir plan de cuentas"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton btnCrearNew 
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   38
      Top             =   10080
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboMonedas 
      Height          =   315
      Left            =   1800
      TabIndex        =   39
      Top             =   7710
      Width           =   2760
      _Version        =   786432
      _ExtentX        =   4868
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Appearance      =   6
      Text            =   "cboMoneda"
      DropDownItemCount=   3
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Moneda"
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
      Left            =   885
      TabIndex        =   40
      Top             =   7740
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1095
      TabIndex        =   36
      Top             =   7410
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Razón Social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   34
      Top             =   1725
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Domicilio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   33
      Top             =   2400
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Ciudad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   32
      Top             =   2760
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "CP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   31
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Teléfonos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   30
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   29
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   28
      Top             =   4200
      Width           =   540
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Contacto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   27
      Top             =   4560
      Width           =   780
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   26
      Top             =   4920
      Width           =   450
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Bonificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   25
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "CUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   24
      Top             =   180
      Width           =   450
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "IIBB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   23
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Nombre Fantasía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   22
      Top             =   1005
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   21
      Top             =   1320
      Width           =   315
   End
End
Attribute VB_Name = "frmComprasProveedoresModifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Id As Long
Dim vTipo As TipoOperacionProveedor
Dim proveedor_ As clsProveedor
'Dim baseP As New classCompras

Public Property Let Proveedor(nValue As clsProveedor)
    Set proveedor_ = DAOProveedor.FindById(nValue.Id)
End Property

Public Property Let tipoOperacion(Tipo As TipoOperacionProveedor)
    vTipo = Tipo
End Property

Public Property Let idProveedor(nId As Long)
    Id = nId
End Property

Private Sub btnCrearNew_Click(Index As Integer)
    Dim cleanedText As String

    cleanedText = Replace(Trim$(Me.Text1(10).Text), " ", "")
    cleanedText = Replace(cleanedText, "-", "")
    Me.Text1(10).Text = cleanedText

    If Trim$(Me.Text1(9).Text) = "" Then Me.Text1(9).Text = "0"

    If LenB(Trim$(Me.Text1(0).Text)) = 0 Or LenB(Trim$(Me.Text1(12).Text)) = 0 Then
        MsgBox "Debe especificar una razón social y nombre fantasia.", vbExclamation
        Exit Sub
    End If

    Call Accion
End Sub

Private Sub btnVerificarCUIT_Click()

    On Error GoTo ManejarError

    Dim consulta As clsConsultaARCA
    Dim cuitIngresado As String
    Dim mensajeError As String
    Dim condicionIVAEncontrada As Boolean
    Dim respuesta As VbMsgBoxResult

    Dim numeroError As Long
    Dim descripcionError As String

    cuitIngresado = SoloNumerosProveedor( _
                        Me.Text1(10).Text _
                    )

    If Len(cuitIngresado) = 0 Then

        MsgBox _
            "Ingrese primero el CUIT del proveedor.", _
            vbExclamation, _
            "Consulta ARCA"

        Me.Text1(10).SetFocus
        Exit Sub

    End If

    If Len(cuitIngresado) <> 11 Then

        MsgBox _
            "El CUIT debe contener 11 números.", _
            vbExclamation, _
            "Consulta ARCA"

        Me.Text1(10).SetFocus
        Exit Sub

    End If

    'Si ya hay datos identificatorios, advertimos que se reemplazarán.
    If Len(Trim$(Me.Text1(0).Text)) > 0 Or _
       Len(Trim$(Me.Text1(11).Text)) > 0 Or _
       Len(Trim$(Me.Text1(12).Text)) > 0 Then

        respuesta = MsgBox( _
            "La consulta reemplazará los datos identificatorios " & _
            "del proveedor:" & vbCrLf & vbCrLf & _
            "• Razón social" & vbCrLf & _
            "• Nombre de fantasía" & vbCrLf & _
            "• Domicilio" & vbCrLf & _
            "• Ciudad" & vbCrLf & _
            "• Código postal" & vbCrLf & _
            "• Condición frente al IVA" & vbCrLf & vbCrLf & _
            "¿Desea continuar?", _
            vbYesNo + vbQuestion, _
            "Consulta ARCA" _
        )

        If respuesta <> vbYes Then
            Exit Sub
        End If

    End If

    Me.btnVerificarCUIT.Enabled = False
    Me.btnVerificarCUIT.caption = "Consultando..."

    Screen.MousePointer = vbHourglass
    DoEvents

    Set consulta = New clsConsultaARCA

    If Not consulta.Consultar(cuitIngresado) Then

        mensajeError = consulta.UltimoError

        LimpiarDatosIdentificatoriosARCA

        MsgBox _
            mensajeError & _
            vbCrLf & vbCrLf & _
            "No se obtuvieron datos desde ARCA." & _
            vbCrLf & _
            "Los datos identificatorios fueron limpiados.", _
            vbExclamation, _
            "Consulta ARCA"

        GoTo salir

    End If

    '=========================================
    ' Completar datos obtenidos desde ARCA
    '=========================================

    Me.Text1(10).Text = consulta.cuit

    Me.Text1(0).Text = _
        UCase$(Trim$(consulta.RazonSocial))

    'ARCA no devuelve nombre de fantasía.
    'Usamos la razón social como valor inicial.
    Me.Text1(12).Text = _
        UCase$(Trim$(consulta.RazonSocial))

    Me.Text1(11).Text = _
        UCase$(Trim$(consulta.direccion))

    'Si no vino la dirección separada, usamos el domicilio completo.
    If Len(Trim$(Me.Text1(11).Text)) = 0 Then

        Me.Text1(11).Text = _
            UCase$(Trim$(consulta.Domicilio))

    End If

    Me.Text1(2).Text = _
        UCase$(Trim$(consulta.localidad))

    Me.Text1(3).Text = _
        Trim$(consulta.CodigoPostal)

    'La consulta de constancia no devuelve el número de IIBB.
    Me.Text1(1).Text = vbNullString

    condicionIVAEncontrada = _
        SeleccionarTipoIVAProveedorARCA( _
            consulta.CondicionIVA _
        )

    If Not condicionIVAEncontrada Then

        MsgBox _
            "ARCA devolvió correctamente los datos del proveedor," & _
            vbCrLf & _
            "pero no se encontró una condición de IVA equivalente " & _
            "en el combo:" & _
            vbCrLf & vbCrLf & _
            consulta.CondicionIVA & _
            vbCrLf & vbCrLf & _
            "Seleccione la condición frente al IVA manualmente." & _
            vbCrLf & _
            "El número de IIBB también debe completarse manualmente.", _
            vbExclamation, _
            "Consulta ARCA"

    Else

        MsgBox _
            "Los datos del proveedor se obtuvieron correctamente " & _
            "desde ARCA." & _
            vbCrLf & vbCrLf & _
            "Razón social: " & consulta.RazonSocial & vbCrLf & _
            "Condición IVA: " & consulta.CondicionIVA & vbCrLf & _
            "Domicilio: " & consulta.Domicilio & vbCrLf & vbCrLf & _
            "Revise el nombre de fantasía y complete IIBB " & _
            "si corresponde.", _
            vbInformation, _
            "Consulta ARCA"

    End If

    Me.Text1(12).SetFocus

salir:

    Screen.MousePointer = vbDefault

    Me.btnVerificarCUIT.Enabled = True
    Me.btnVerificarCUIT.caption = "Consultar ARCA"

    Set consulta = Nothing
    Exit Sub

ManejarError:

    numeroError = Err.Number
    descripcionError = Err.Description

    Screen.MousePointer = vbDefault

    LimpiarDatosIdentificatoriosARCA

    Me.btnVerificarCUIT.Enabled = True
    Me.btnVerificarCUIT.caption = "Consultar ARCA"

    MsgBox _
        "Error al consultar ARCA: " & _
        CStr(numeroError) & " - " & descripcionError & _
        vbCrLf & vbCrLf & _
        "Los datos identificatorios fueron limpiados.", _
        vbCritical, _
        "Consulta ARCA"

    Set consulta = Nothing

End Sub

Private Sub cmdPlanCuentas_Click()
    Dim frm As New frmAdminComprasCuentasDefinir
    Set frm.vProveedor = proveedor_

    frm.Show
End Sub

Private Sub Command2_Click()
    Buscar
End Sub

Private Sub Command3_Click()
    Dim i As Long
    For i = Me.ListView1.ListItems.count To 1 Step -1
        If Me.ListView1.ListItems(i).Checked = True Then
            Me.ListView1.ListItems.remove (i)
        End If
    Next i
End Sub

Private Sub Buscar()
    Dim x As Long
    Dim esta As Boolean
    Dim i As Long
    Dim h As ListItem

    For x = 1 To Me.lstRubros.ListItems.count
        If Me.lstRubros.ListItems(x).Checked = True Then
            esta = False

            For i = 1 To Me.ListView1.ListItems.count
                If Me.ListView1.ListItems(i) = Me.lstRubros.ListItems(x) Then esta = True
            Next i

            If Not esta Then
                Set h = Me.ListView1.ListItems.Add(, , Me.lstRubros.ListItems(x))
                Set h.Tag = Me.lstRubros.ListItems(x).Tag
            End If
        End If
    Next x
End Sub

Private Function Accion() As Boolean
    On Error GoTo err123

    Dim a1 As clsRubros
    Dim colRubros As New Collection
    Dim l As Long
    Dim esNuevo As Boolean

    Accion = False

    esNuevo = Not IsSomething(proveedor_)
    If esNuevo Then Set proveedor_ = New clsProveedor

    proveedor_.RazonSocial = UCase$(Trim$(Me.Text1(0).Text))
    proveedor_.direccion = Trim$(Me.Text1(11).Text)
    proveedor_.Ciudad = Trim$(Me.Text1(2).Text)
    proveedor_.cp = Trim$(Me.Text1(3).Text)
    proveedor_.tel = Trim$(Me.Text1(4).Text)
    proveedor_.Fax = Trim$(Me.Text1(5).Text)
    proveedor_.email = Trim$(Me.Text1(6).Text)
    proveedor_.Contacto = Trim$(Me.Text1(7).Text)
    proveedor_.FormaPago = Trim$(Me.Text1(8).Text)
    proveedor_.bonificacion = CDbl(val(Me.Text1(9).Text))

    proveedor_.CBU = Trim$(Me.txtCBU.Text)
    proveedor_.ALIAS = Trim$(Me.txtAlias.Text)
    proveedor_.TitularCta = Trim$(Me.txtTitularCta.Text)

    proveedor_.estado = Me.cboEstadoProveedor.ListIndex

    If Not IsNumeric(Me.Text1(1).Text) Then
        proveedor_.IIBB = 0
    Else
        proveedor_.IIBB = Me.Text1(1).Text
    End If

    proveedor_.razonFantasia = UCase$(Trim$(Me.Text1(12).Text))
    proveedor_.pagoDolares = Abs(Me.Check2.value)
    proveedor_.pagocontraEntrega = Abs(Me.Check1.value)
    proveedor_.cuit = Replace(Replace(Trim$(Me.Text1(10).Text), " ", ""), "-", "")

    Set proveedor_.moneda = DAOMoneda.GetById(CLng(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)))
    Set proveedor_.TipoIVA = DAOTipoIvaProveedor.GetById(CLng(Me.cboIVA.ItemData(Me.cboIVA.ListIndex)))

    Set colRubros = Nothing
    For l = 1 To Me.ListView1.ListItems.count
        Set a1 = New clsRubros
        Set a1 = Me.ListView1.ListItems(l).Tag
        colRubros.Add a1
    Next l

    proveedor_.rubros = colRubros

    If proveedor_.estado <> EstadoProveedorEliminado Then
        If LenB(proveedor_.cuit) > 0 And Not IsNumeric(proveedor_.cuit) Then
            Err.Raise 400, "Proveedor", "El CUIT debe ser numérico."
        End If

        If Not EsProveedorExterior Then
            Dim F As String

            F = "proveedores.cuit = " & Escape(proveedor_.cuit)

            If proveedor_.Id > 0 Then
                F = F & " AND proveedores.id <> " & proveedor_.Id
            End If

            If DAOProveedor.FindAll(F).count > 0 Then
                Err.Raise 400, "Proveedor", "El CUIT ya se encuentra asignado a otro proveedor."
            End If
        End If
    End If

    If Not DAOProveedor.Save(proveedor_) Then
        MsgBox "Se produjo un error, no se guardarán los cambios.", vbCritical
        Exit Function
    End If

    Accion = True

    If esNuevo Then
        MsgBox "Proveedor guardado correctamente.", vbInformation
    Else
        MsgBox "Proveedor actualizado correctamente.", vbInformation
    End If

    Exit Function

err123:
    MsgBox Err.Description, vbCritical, "·Error·"
End Function

Private Sub mostrarCampos()
    'Set vProveedor = DAOProveedor.BuscarPorID(id)
    Check1.value = Abs(proveedor_.pagocontraEntrega)
    Check2.value = Abs(proveedor_.pagoDolares)
    Text1(0) = proveedor_.RazonSocial
    Text1(11) = proveedor_.direccion
    Text1(2) = proveedor_.Ciudad
    Text1(3) = proveedor_.cp
    Text1(4) = proveedor_.tel
    Text1(5) = proveedor_.Fax
    Text1(6) = proveedor_.email
    Text1(7) = proveedor_.Contacto
    Text1(8) = proveedor_.FormaPago
    Text1(9) = proveedor_.bonificacion
    Text1(10) = proveedor_.cuit
    Text1(1) = proveedor_.IIBB
    Text1(12) = proveedor_.razonFantasia
    cboMonedas.ListIndex = funciones.PosIndexCbo(proveedor_.moneda.Id, cboMonedas)
    cboIVA.ListIndex = funciones.PosIndexCbo(proveedor_.TipoIVA.Id, cboIVA)
    Me.cboEstadoProveedor.ListIndex = funciones.PosIndexCbo(proveedor_.estado, Me.cboEstadoProveedor)

    Me.txtCBU.Text = proveedor_.CBU
    Me.txtAlias.Text = proveedor_.ALIAS
    Me.txtTitularCta.Text = proveedor_.TitularCta
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    If proveedor_ Is Nothing Then
        Me.caption = "Crear Proveedor..."
    
    Else
        Me.caption = "Crear Modificar Proveedor..."
    End If

    If vTipo = ver Then
        Me.caption = "Consultar Proveedor..."
    End If

    LlenarEstadosProveedor
    llenarIva
    
    Me.btnVerificarCUIT.caption = "Cargar desde ARCA"
    
    
    llenarListarubros
    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

    If Not proveedor_ Is Nothing Then
        mostrarCampos
        llenarListaRubrosProveedor
    Else
        limpiar
    End If

    ''Me.caption = caption & " (" & Name & ")"
End Sub

Private Sub LlenarEstadosProveedor()
    Dim i As Long

    For i = 0 To 2
        Me.cboEstadoProveedor.AddItem EnumEstadoProveedor(i)
        Me.cboEstadoProveedor.ItemData(Me.cboEstadoProveedor.NewIndex) = i
    Next i

    Me.cboEstadoProveedor.ListIndex = 1
End Sub

Private Sub llenarListarubros()
    Dim ListaRubros As Collection
    Dim rubro As clsRubros
    Dim u As Long
    Dim x As ListItem

    Set ListaRubros = DAORubros.FindAll
    lstRubros.ListItems.Clear

    For u = 1 To ListaRubros.count
        Set rubro = ListaRubros(u)
        Set x = Me.lstRubros.ListItems.Add(, , rubro.rubro)
        Set x.Tag = rubro
    Next
End Sub

Private Sub llenarListaRubrosProveedor()
    Dim ListaRubros As New Collection
    Dim rubro As clsRubros
    Dim x As ListItem
    Dim u As Long

    Set ListaRubros = DAORubros.FindAllByProveedor(proveedor_.Id)
    Me.ListView1.ListItems.Clear

    For u = 1 To ListaRubros.count
        Set rubro = ListaRubros(u)
        Set x = Me.ListView1.ListItems.Add(, , rubro.rubro)
        Set x.Tag = rubro
    Next
End Sub

Private Sub limpiar()

    Dim x As Integer

    For x = 0 To 12
        Text1(x).Text = vbNullString
    Next x

    Text1(9).Text = "0"

    Me.txtCBU.Text = vbNullString
    Me.txtAlias.Text = vbNullString
    Me.txtTitularCta.Text = vbNullString

    Me.ListView1.ListItems.Clear

    If Me.cboIVA.ListCount > 0 Then
        Me.cboIVA.ListIndex = -1
    End If

End Sub

'Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
'    If EVENTO.EVENTO = agregar_ Then
'
'    Else
'
'    End If
'End Function

Private Sub lstRubros_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    funciones.LstOrdenar Me.lstRubros, ColumnHeader.Index
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    foco Me.Text1(Index)
End Sub

Public Sub llenarIva()
    DAOTipoIvaProveedor.llenarComboXtremeSuite Me.cboIVA
End Sub

Private Function EsProveedorExterior() As Boolean
    If Me.cboIVA.ListIndex < 0 Then
        EsProveedorExterior = False
        Exit Function
    End If

    EsProveedorExterior = (UCase$(Trim$(Me.cboIVA.Text)) = "EXTERIOR")
End Function


Private Sub LimpiarDatosIdentificatoriosARCA()

    On Error Resume Next

    Me.Text1(10).Text = vbNullString   'CUIT
    Me.Text1(0).Text = vbNullString    'Razón social
    Me.Text1(12).Text = vbNullString   'Nombre fantasía
    Me.Text1(1).Text = vbNullString    'IIBB
    Me.Text1(11).Text = vbNullString   'Domicilio
    Me.Text1(2).Text = vbNullString    'Ciudad
    Me.Text1(3).Text = vbNullString    'Código postal

    If Me.cboIVA.ListCount > 0 Then
        Me.cboIVA.ListIndex = -1
    End If

End Sub

Private Function SoloNumerosProveedor( _
    ByVal valor As String _
) As String

    Dim i As Long
    Dim caracter As String
    Dim resultado As String

    resultado = vbNullString

    For i = 1 To Len(valor)

        caracter = Mid$(valor, i, 1)

        If caracter >= "0" And _
           caracter <= "9" Then

            resultado = resultado & caracter

        End If

    Next i

    SoloNumerosProveedor = resultado

End Function

Private Function BuscarTextoExactoComboProveedor( _
    ByVal combo As Object, _
    ByVal textoBuscado As String _
) As Long

    Dim i As Long
    Dim textoBuscadoNormalizado As String
    Dim textoComboNormalizado As String

    BuscarTextoExactoComboProveedor = -1

    textoBuscadoNormalizado = _
        NormalizarTextoProveedor(textoBuscado)

    For i = 0 To combo.ListCount - 1

        textoComboNormalizado = _
            NormalizarTextoProveedor( _
                CStr(combo.list(i)) _
            )

        If textoComboNormalizado = _
           textoBuscadoNormalizado Then

            BuscarTextoExactoComboProveedor = i
            Exit Function

        End If

    Next i

End Function

Private Function SeleccionarTipoIVAProveedorARCA( _
    ByVal condicionARCA As String _
) As Boolean

    Dim condicionNormalizada As String
    Dim candidatos As Variant
    Dim posicion As Long
    Dim i As Long

    SeleccionarTipoIVAProveedorARCA = False

    condicionNormalizada = _
        NormalizarTextoProveedor(condicionARCA)

    Select Case condicionNormalizada

        Case "MONOTRIBUTO"

            candidatos = Array( _
                "Monotributo" _
            )

        Case "EXENTO", _
             "IVA EXENTO"

            candidatos = Array( _
                "Exento", _
                "IVA Exento" _
            )

        Case "RESP INSCRIPTO", _
             "RESPONSABLE INSCRIPTO", _
             "IVA RESPONSABLE INSCRIPTO"

            candidatos = Array( _
                "Resp. Inscripto", _
                "Responsable Inscripto" _
            )

        Case "SIN DATOS", ""

            candidatos = Array( _
                "Sin Datos" _
            )

        Case Else

            candidatos = Array( _
                condicionARCA _
            )

    End Select

    For i = LBound(candidatos) To UBound(candidatos)

        posicion = BuscarTextoExactoComboProveedor( _
                        Me.cboIVA, _
                        CStr(candidatos(i)) _
                    )

        If posicion >= 0 Then

            Me.cboIVA.ListIndex = posicion

            SeleccionarTipoIVAProveedorARCA = True
            Exit Function

        End If

    Next i

End Function


Private Function NormalizarTextoProveedor( _
    ByVal valor As String _
) As String

    Dim resultado As String

    resultado = UCase$(Trim$(valor))

    resultado = Replace(resultado, "Á", "A")
    resultado = Replace(resultado, "É", "E")
    resultado = Replace(resultado, "Í", "I")
    resultado = Replace(resultado, "Ó", "O")
    resultado = Replace(resultado, "Ú", "U")
    resultado = Replace(resultado, "Ü", "U")

    resultado = Replace(resultado, ".", " ")
    resultado = Replace(resultado, ",", " ")
    resultado = Replace(resultado, "-", " ")
    resultado = Replace(resultado, "_", " ")

    Do While InStr(resultado, "  ") > 0
        resultado = Replace(resultado, "  ", " ")
    Loop

    NormalizarTextoProveedor = Trim$(resultado)

End Function
