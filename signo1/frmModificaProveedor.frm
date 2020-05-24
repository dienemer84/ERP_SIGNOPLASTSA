VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmComprasProveedoresModifica 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar proveedor"
   ClientHeight    =   8535
   ClientLeft      =   210
   ClientTop       =   195
   ClientWidth     =   8325
   ClipControls    =   0   'False
   Icon            =   "frmModificaProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ComboBox cboIva 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   840
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5280
      Width           =   2775
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1830
      Left            =   240
      TabIndex        =   35
      Top             =   6120
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
      Left            =   4845
      TabIndex        =   15
      Top             =   5310
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Pago contra entrega"
      Height          =   300
      Left            =   6165
      TabIndex        =   16
      Top             =   5310
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1215
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2280
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2640
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3000
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   3360
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3720
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1440
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4080
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4800
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1440
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4440
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "El cuit va sin guiones!"
      Top             =   135
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1920
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   6495
   End
   Begin XtremeSuiteControls.PushButton cmdPlanCuentas 
      Height          =   375
      Left            =   255
      TabIndex        =   37
      Top             =   8070
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Definir plan de cuentas"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton btnCrear 
      Height          =   375
      Left            =   6870
      TabIndex        =   38
      Top             =   8055
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
      Left            =   1440
      TabIndex        =   39
      Top             =   5670
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
      Left            =   525
      TabIndex        =   40
      Top             =   5700
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
      Left            =   735
      TabIndex        =   36
      Top             =   5370
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
      Left            =   240
      TabIndex        =   34
      Top             =   1245
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
      Left            =   600
      TabIndex        =   33
      Top             =   1920
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
      Left            =   720
      TabIndex        =   32
      Top             =   2280
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
      Left            =   1080
      TabIndex        =   31
      Top             =   2640
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
      Left            =   480
      TabIndex        =   30
      Top             =   3000
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
      Left            =   960
      TabIndex        =   29
      Top             =   3360
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
      Left            =   720
      TabIndex        =   28
      Top             =   3720
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
      Left            =   480
      TabIndex        =   27
      Top             =   4080
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
      Left            =   840
      TabIndex        =   26
      Top             =   4440
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
      Left            =   240
      TabIndex        =   25
      Top             =   4800
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
      Left            =   960
      TabIndex        =   24
      Top             =   135
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
      Left            =   960
      TabIndex        =   23
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Fantasía"
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
      Left            =   600
      TabIndex        =   22
      Top             =   480
      Width           =   765
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
      Left            =   1080
      TabIndex        =   21
      Top             =   840
      Width           =   315
   End
End
Attribute VB_Name = "frmComprasProveedoresModifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim id As Long
Dim vTipo As TipoOperacionProveedor
Dim proveedor_ As clsProveedor
Dim baseP As New classCompras

Public Property Let Proveedor(nvalue As clsProveedor)
    Set proveedor_ = DAOProveedor.FindById(nvalue.id)
End Property

Public Property Let tipoOperacion(Tipo As TipoOperacionProveedor)
    vTipo = Tipo
End Property
Public Property Let idProveedor(nId As Long)
    id = nId
End Property
Private Sub btnCrear_Click()


    If Trim(Text1(9)) = Empty Then Text1(9) = 0
    If LenB(Text1(10)) = 0 Then Text1(10) = 0
    If LenB(Text1(0)) = 0 Or LenB(Text1(12)) = 0 Then
        MsgBox "Debe especificar una razon social y nombre fantasia.", vbExclamation
        Exit Sub
    End If
    accion

End Sub
Private Sub Command1_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
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
Private Function accion() As Boolean
    On Error GoTo err123
    accion = True
    Dim a1 As clsRubros
    Dim colRubros As New Collection


    If Not IsSomething(proveedor_) Then Set proveedor_ = New clsProveedor

    proveedor_.RazonSocial = Me.Text1(0)
    proveedor_.direccion = Me.Text1(11)
    proveedor_.Ciudad = Me.Text1(2)
    proveedor_.cp = Me.Text1(3)
    proveedor_.tel = Me.Text1(4)
    proveedor_.Fax = Me.Text1(5)
    proveedor_.email = Me.Text1(6)
    proveedor_.contacto = Me.Text1(7)
    proveedor_.FormaPago = Me.Text1(8)
    proveedor_.bonificacion = CDbl(Me.Text1(9))
    proveedor_.estado = Me.cboEstadoProveedor.ListIndex
    If Not IsNumeric(Me.Text1(1)) Then
        proveedor_.IIBB = 0
    Else
        proveedor_.IIBB = Me.Text1(1)
    End If
    proveedor_.razonFantasia = Me.Text1(12)
    proveedor_.pagoDolares = Abs(Me.Check2.value)
    proveedor_.pagocontraEntrega = Abs(Me.Check1.value)
    proveedor_.Cuit = Me.Text1(10)
    Set proveedor_.Moneda = DAOMoneda.GetById(CLng(Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)))
    Set proveedor_.TipoIVA = DAOTipoIvaProveedor.GetById(CLng(Me.cboIVA.ItemData(Me.cboIVA.ListIndex)))

    'busco rubros

    Set colRubros = Nothing
    Dim l As Long
    For l = 1 To Me.ListView1.ListItems.count
        Set a1 = New clsRubros
        Set a1 = Me.ListView1.ListItems(l).Tag
        colRubros.Add a1
    Next l


    proveedor_.rubros = colRubros
    If proveedor_.estado <> EstadoProveedorEliminado Then
        If Not DAOProveedor.ValidarCuit(proveedor_) Then
            Err.Raise 400, "Proveedor", "El CUIT ya se encuentra asignado o no tiene el formato correcto."
        End If
    End If

    If Not DAOProveedor.Save(proveedor_) Then
        MsgBox "Se produjo un error, no se guardarán los cambios!", vbCritical
    Else
        MsgBox "Actualización exitosa!", vbInformation
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
    Text1(1) = proveedor_.direccion
    Text1(2) = proveedor_.Ciudad
    Text1(3) = proveedor_.cp
    Text1(4) = proveedor_.tel
    Text1(5) = proveedor_.Fax
    Text1(6) = proveedor_.email
    Text1(7) = proveedor_.contacto
    Text1(8) = proveedor_.FormaPago
    Text1(9) = proveedor_.bonificacion
    Text1(10) = proveedor_.Cuit
    Text1(11) = proveedor_.IIBB
    Text1(12) = proveedor_.razonFantasia
    cboMonedas.ListIndex = funciones.PosIndexCbo(proveedor_.Moneda.id, cboMonedas)
    cboIVA.ListIndex = funciones.PosIndexCbo(proveedor_.TipoIVA.id, cboIVA)
    Me.cboEstadoProveedor.ListIndex = funciones.PosIndexCbo(proveedor_.estado, Me.cboEstadoProveedor)
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
    Set ListaRubros = DAORubros.FindAll
    Dim Rubro As clsRubros
    lstRubros.ListItems.Clear
    Dim u As Long
    Dim x As ListItem
    For u = 1 To ListaRubros.count
        Set Rubro = ListaRubros(u)
        Set x = Me.lstRubros.ListItems.Add(, , Rubro.Rubro)
        Set x.Tag = Rubro
    Next
End Sub

Private Sub llenarListaRubrosProveedor()
    Dim ListaRubros As New Collection
    Set ListaRubros = DAORubros.FindAllByProveedor(proveedor_.id)
    Dim Rubro As clsRubros
    Me.ListView1.ListItems.Clear
    Dim x As ListItem
    Dim u As Long
    For u = 1 To ListaRubros.count
        Set Rubro = ListaRubros(u)
        Set x = Me.ListView1.ListItems.Add(, , Rubro.Rubro)
        Set x.Tag = Rubro
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

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    If EVENTO.EVENTO = agregar_ Then

    Else

    End If
End Function

Private Sub lstRubros_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    funciones.LstOrdenar Me.lstRubros, ColumnHeader.index
End Sub

Private Sub Text1_GotFocus(index As Integer)
    foco Me.Text1(index)
End Sub
Public Sub llenarIva()
    DAOTipoIvaProveedor.llenarComboXtremeSuite Me.cboIVA
End Sub

Private Sub Text1_Validate(index As Integer, Cancel As Boolean)
    If index = 10 Then    '10=cuit
        Cancel = Not IsNumeric(Me.Text1(10)) And LenB(Me.Text1(10)) > 0

        If Not Cancel Then
            Dim F As String
            F = "proveedores.cuit = " & Escape(Me.Text1(10))
            If IsSomething(proveedor_) Then
                F = F & " AND proveedores.id <> " & proveedor_.id
            End If

            Cancel = DAOProveedor.FindAll(F).count > 0
            If Cancel Then MsgBox "Ya existe un proveedor con ese Nº de CUIT.", vbExclamation
        End If

    End If
End Sub
