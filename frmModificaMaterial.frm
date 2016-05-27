VERSION 5.00
Begin VB.Form frmModificaMaterial 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificar material..."
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmModificaTarea 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Verificar datos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   6240
         TabIndex        =   30
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3000
         Width           =   6015
      End
      Begin VB.ComboBox cboRubros 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   6015
      End
      Begin VB.ComboBox cboGrupos 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   6015
      End
      Begin VB.ComboBox cboUnidades 
         Height          =   315
         ItemData        =   "frmModificaMaterial.frx":0000
         Left            =   1320
         List            =   "frmModificaMaterial.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2520
         Width           =   6015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   3360
         Width           =   6015
      End
      Begin VB.ComboBox cboMonedas 
         Height          =   315
         ItemData        =   "frmModificaMaterial.frx":0024
         Left            =   1320
         List            =   "frmModificaMaterial.frx":002E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3720
         Width           =   6015
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "txtFecha"
         Top             =   4080
         Width           =   6015
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label13"
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblUnidades 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label13"
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblGrupos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label12"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblRubros 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label12"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   5040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblValor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label12"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   4920
         Visible         =   0   'False
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label11"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label10"
         Height          =   375
         Left            =   960
         TabIndex        =   22
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidad "
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
         TabIndex        =   21
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kg x M2/Ml "
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
         TabIndex        =   20
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Espesor "
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
         TabIndex        =   19
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción "
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grupo "
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rubro "
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
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código "
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
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor Unitario  "
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
         TabIndex        =   14
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label FEcha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha "
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
         TabIndex        =   13
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Moneda "
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
         TabIndex        =   12
         Top             =   3720
         Width           =   1215
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   495
      Left            =   3120
      TabIndex        =   27
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmModificaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As New classConfigurar
Private Sub cboRubros_Click()
Set base = New classModificaMaterial
 If cboRubros.ListIndex <> -1 Then
    IntIdRubro = cboRubros.ItemData(cboRubros.ListIndex)
    base.llenar_combo_grupos CInt(IntIdRubro)
End If
End Sub

Private Sub cboUnidades_Click()
'If cboUnidades.ItemData(cboUnidades.ListIndex) = 4 Or cboUnidades.ItemData(cboUnidades.ListIndex) = 1 Then
' Text1(3) = 1
' Text1(3).Locaed = True
'Else
' Text1(3) = Empty
' Text1(3).Locaaked = False
'End If
If Text1(0) = Empty Or Text1(3) = Empty Or Text1(1) = Empty Or Text1(2) = Empty Or txtValor = Emptyy Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
Private Sub Command1_Click()
Dim ver As ClassListaMateriales
Set ver = New ClassListaMateriales
id_a_buscar = CInt(Label11)
On Error GoTo err1
ErrorCode = 0
If cboRubros.ListIndex = -1 Or cboGrupos.ListIndex = -1 Or cboUnidades.ListIndex = -1 Then
    ErrorCode = 1
End If
If cboMonedas.ListIndex = -1 Then
    ErrorCode = 3
End If
If Not IsNumeric(Text1(2)) Or Not IsNumeric(Text1(3)) Or Not IsNumeric(txtValor) Then
    ErrorCode = 2
End If
If ver.verifica_codigo(Text1(0)) = 1 And Text1(0).Text <> Label10 Then   '1=codigo existente.
    ErrorCode = 4
End If
Select Case ErrorCode
 Case 1: MsgBox "Debe seleccionar Rubros/Grupos", vbCritical, "Error"
 Case 2: MsgBox "Debe introducir datos válidos para espesor/Peso/Valor", vbCritical, "Error"
 Case 3: MsgBox "Debe seleccionar Moneda", vbCritical, "Error"
 Case 4: MsgBox "El código existe", vbCritical, "Error"
End Select
If ErrorCode = 0 Then
h = MsgBox("¿Está seguro de actualizar datos?", vbYesNo, "Advertencia")
If h = 6 Then
 id_rubro = Me.cboRubros.ItemData(cboRubros.ListIndex)
 id_grupo = Me.cboGrupos.ItemData(cboGrupos.ListIndex)
 id_unidad = Me.cboUnidades.ItemData(cboUnidades.ListIndex)
 codigo = Text1(0)
 descripcion = Text1(1)
 espesor = CDbl(Text1(2))
 pesoxUn = CDbl(Text1(3))
 valor = CDbl(txtValor)
 id_moneda = Me.cboMonedas.ItemData(Me.cboMonedas.ListIndex)
 Fech = Format(Date, "yyyy/mm/dd")
 strsql = "update materiales set id_rubro=" & id_rubro & ",id_grupo=" & id_grupo & ",id_unidad=" & id_unidad & ",codigo='" & codigo & "',descripcion='" & descripcion & "',espesor=" & espesor & ",pesoxunidad=" & pesoxUn & " WHERE id=" & id_a_buscar
 base.ejecutar_consulta (strsql)
  strsql = "update valores_MATERIALES set valor_unitario=" & valor & ",Fecha_actualizacion='" & Fech & "',id_moneda=" & id_moneda & " where id_material = " & id_a_buscar
  base.ejecutar_consulta (strsql)
If CDbl(lblValor <> txtValor) Then
    MsgBox "Hubo una variacion. Se almacena en históricos", vbInformation, "Información"
    Set base2 = New ClassListaMateriales
    d = base2.crear_historico(id_a_buscar, valor, id_moneda, 1)
End If
frmListaMateriales.Label1 = Me.Label11
Unload Me

End If
End If


Exit Sub
err1:
MsgBox Err.Description
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
base.llenar_combo_rubros
If cboRubros.ListCount > 0 Then
    IntIdRubro = cboRubros.ItemData(cboRubros.ListIndex)
    base.llenar_combo_grupos CInt(IntIdRubro)
    base.llenar_form_mod (CInt(Label11))
    Me.cboGrupos.ListIndex = PosIndexCbo(Me.lblGrupos, Me.cboGrupos)
    Me.cboMonedas.ListIndex = PosIndexCbo(Me.lblMoneda, Me.cboMonedas)
    Me.cboRubros.ListIndex = PosIndexCbo(Me.lblRubros, Me.cboRubros)
   Me.cboUnidades.ListIndex = PosIndexCbo(Me.lblUnidades, Me.cboUnidades)
End If
End Sub
Private Sub Form_Load()
For X = 0 To 3
Text1(X) = Empty
Next
txtValor = Empty
txtFecha = Date
Set base = New classModificaMaterial
base.llenar_combo_rubros
Me.cboUnidades.ListIndex = 0
Me.cboMonedas.ListIndex = 0
End Sub

Private Sub Text1_Change(Index As Integer)
If Text1(0) = Empty Or Text1(3) = Empty Or Text1(1) = Empty Or Text1(2) = Empty Or txtValor = Emptyy Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))
End Sub
Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(txtFecha)
End Sub
Private Sub txtValor_Change()
If Text1(0) = Empty Or Text1(3) = Empty Or Text1(1) = Empty Or Text1(2) = Empty Or txtValor = Empty Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
Private Sub txtValor_GotFocus()
txtValor.SelStart = 0
txtValor.SelLength = Len(txtValor)
End Sub
