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
      TabIndex        =   43
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
         TabIndex        =   46
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox txtTitularCta 
         Height          =   285
         Left            =   1560
         TabIndex        =   45
         Top             =   960
         Width           =   6375
      End
      Begin VB.TextBox txtAlias 
         Height          =   285
         Left            =   1560
         TabIndex        =   44
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
         Top             =   255
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.PushButton btnVerificarCUIT 
      Height          =   375
      Left            =   6240
      TabIndex        =   42
      Top             =   480
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Verificar CUIT"
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
      Left            =   5205
      TabIndex        =   15
      Top             =   7350
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Pago contra entrega"
      Height          =   300
      Left            =   6525
      TabIndex        =   16
      Top             =   7350
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
   Begin XtremeSuiteControls.Label Label17 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   41
      Top             =   480
      Width           =   4215
      _Version        =   786432
      _ExtentX        =   7435
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Verifique si el CUIT ingresado es correcto >>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
    Dim Ie As New InternetExplorer
    Ie.Visible = True
    Ie.Navigate "https://seti.afip.gob.ar/padron-puc-constancia-internet/ConsultaConstanciaAction.do"
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
    proveedor_.Cuit = Replace(Replace(Trim$(Me.Text1(10).Text), " ", ""), "-", "")

    Set proveedor_.moneda = DAOMoneda.GetById(CLng(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)))
    Set proveedor_.TipoIVA = DAOTipoIvaProveedor.GetById(CLng(Me.cboIva.ItemData(Me.cboIva.ListIndex)))

    Set colRubros = Nothing
    For l = 1 To Me.ListView1.ListItems.count
        Set a1 = New clsRubros
        Set a1 = Me.ListView1.ListItems(l).Tag
        colRubros.Add a1
    Next l

    proveedor_.rubros = colRubros

    If proveedor_.estado <> EstadoProveedorEliminado Then
        If LenB(proveedor_.Cuit) > 0 And Not IsNumeric(proveedor_.Cuit) Then
            Err.Raise 400, "Proveedor", "El CUIT debe ser numérico."
        End If

        If Not EsProveedorExterior Then
            Dim F As String

            F = "proveedores.cuit = " & Escape(proveedor_.Cuit)

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
    Text1(10) = proveedor_.Cuit
    Text1(1) = proveedor_.IIBB
    Text1(12) = proveedor_.razonFantasia
    cboMonedas.ListIndex = funciones.PosIndexCbo(proveedor_.moneda.Id, cboMonedas)
    cboIva.ListIndex = funciones.PosIndexCbo(proveedor_.TipoIVA.Id, cboIva)
    Me.cboEstadoProveedor.ListIndex = funciones.PosIndexCbo(proveedor_.estado, Me.cboEstadoProveedor)

    Me.txtCBU.Text = proveedor_.CBU
    Me.txtAlias.Text = proveedor_.ALIAS
    Me.txtTitularCta.Text = proveedor_.TitularCta
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    If proveedor_ Is Nothing Then
        Me.caption = "Crear Proveedor..."
        Me.limpiar
    Else
        Me.caption = "Crear Modificar Proveedor..."
    End If

    If vTipo = ver Then
        Me.caption = "Consultar Proveedor..."
    End If

    LlenarEstadosProveedor
    llenarIva
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

Function limpiar()
    Dim x As Integer

    For x = 0 To 12
        Text1(x) = Empty
    Next x

    Text1(9) = 0
    Me.ListView1.ListItems.Clear
End Function

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
    DAOTipoIvaProveedor.llenarComboXtremeSuite Me.cboIva
End Sub

Private Function EsProveedorExterior() As Boolean
    If Me.cboIva.ListIndex < 0 Then
        EsProveedorExterior = False
        Exit Function
    End If

    EsProveedorExterior = (UCase$(Trim$(Me.cboIva.Text)) = "EXTERIOR")
End Function
