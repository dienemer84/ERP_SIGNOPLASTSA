VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAltaEmpleados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleado"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   7755
   ClipControls    =   0   'False
   Icon            =   "frmAltaEmpleados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7755
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   255
      Left            =   6120
      TabIndex        =   33
      Top             =   1800
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Imagen"
      UseVisualStyle  =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpFechaNac 
      Height          =   300
      Left            =   3585
      TabIndex        =   11
      Top             =   4410
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Format          =   55836673
      CurrentDate     =   40119
   End
   Begin VB.TextBox txtGrupoSanguineo 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Text            =   "txtGrupoSanguineo"
      Top             =   4395
      Width           =   900
   End
   Begin VB.TextBox txtNroLegajo 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtApellido 
      Height          =   285
      Left            =   1335
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1140
      Width           =   4695
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1500
      Width           =   4695
   End
   Begin VB.TextBox TxtDireccion 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2580
      Width           =   6375
   End
   Begin VB.TextBox TxtLocalidad 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2940
      Width           =   6375
   End
   Begin VB.TextBox txtTel1 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3315
      Width           =   6375
   End
   Begin VB.TextBox txtTel2 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3660
      Width           =   6375
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5415
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4845
      Width           =   1095
   End
   Begin VB.TextBox txtNroDocumento 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2220
      Width           =   6375
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Si"
      Height          =   255
      Left            =   2310
      TabIndex        =   13
      Top             =   4860
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "No"
      Height          =   255
      Left            =   2910
      TabIndex        =   14
      Top             =   4860
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   300
      Width           =   1695
   End
   Begin VB.TextBox txtNombres 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1860
      Width           =   4695
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4020
      Width           =   6375
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4845
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFechaIng 
      Height          =   300
      Left            =   6405
      TabIndex        =   12
      Top             =   4410
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Format          =   55836673
      CurrentDate     =   40119
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Ingreso"
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
      Left            =   5100
      TabIndex        =   32
      Top             =   4455
      Width           =   1230
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Nac"
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
      Left            =   2550
      TabIndex        =   31
      Top             =   4440
      Width           =   945
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grupo Sang"
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
      TabIndex        =   30
      Top             =   4425
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Legajo"
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
      TabIndex        =   29
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apellido"
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
      TabIndex        =   28
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
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
      TabIndex        =   27
      Top             =   1500
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dirección"
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
      TabIndex        =   26
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Localidad"
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
      TabIndex        =   25
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Teléfono 1"
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
      TabIndex        =   24
      Top             =   3300
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Teléfono 2"
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
      TabIndex        =   23
      Top             =   3660
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documento"
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
      TabIndex        =   22
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "¿Apto para sistema?"
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
      Left            =   510
      TabIndex        =   21
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario"
      Height          =   255
      Left            =   660
      TabIndex        =   20
      Top             =   330
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombres"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   1875
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   4020
      Width           =   975
   End
End
Attribute VB_Name = "frmAltaEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strsql As String
'Dim basePE As New ClassPersonal

Private m_empleado As clsEmpleado

Public Property Set Empleado(value As clsEmpleado)
    On Error GoTo err1
    Set m_empleado = value

    txtDireccion = m_empleado.direccion
    txtApellido = m_empleado.Apellido
    txtNombre = m_empleado.nombre
    txtLocalidad = m_empleado.localidad
    txtTel1 = m_empleado.Telefono1
    txtNroDocumento = m_empleado.documento
    txtTel2 = m_empleado.Telefono2
    txtNroLegajo = m_empleado.legajo
    txtNombres = m_empleado.Nombres
    Me.Option2.value = m_empleado.estado
    Me.Option1.value = Not m_empleado.estado
    'Me.txtUsuario = m_empleado.
    Me.txtEmail = m_empleado.email
    Me.txtGrupoSanguineo.text = m_empleado.GrupoSanguineo
    Me.dtpFechaIng.value = m_empleado.FechaIngreso
    Me.dtpFechaNac.value = m_empleado.FechaNacimiento

    If IsSomething(m_empleado) Then

        Dim tmppath As String
        Dim clasea As New classArchivos
        '  Dim col As New Collection
        Dim Foto As archivo

        Set Foto = DAOArchivo.FindAll(OA_FotoEmpleado, "idPieza=" & m_empleado.id)(1)

        If IsSomething(Foto) Then

            tmppath = clasea.exportarArchivo(Foto.id)
            If LenB(tmppath) > 0 Then
                Set Me.Image1.Picture = LoadPicture(tmppath)
                Kill tmppath
            End If
        End If

        'DIC.Item (m_empleado.Id)

    End If


    Exit Property
err1:

End Property

Private Sub cmdGuardar_Click()
    On Error GoTo err44



    If Not IsSomething(m_empleado) Then
        Set m_empleado = New clsEmpleado
        Dim b As Boolean
        b = True




    End If

    Dim tmpEmp As clsEmpleado
    Set tmpEmp = DAOEmpleados.GetByLegajo(CLng(Me.txtNroLegajo))

    If IsSomething(tmpEmp) Then
        If tmpEmp.id <> m_empleado.id Then
            MsgBox "El número de legajo ya existe", vbCritical, "Error"
            Exit Sub
        End If
    End If

    If IsValidEmail(Trim(Me.txtEmail)) Then Exit Sub


    m_empleado.email = Trim(Me.txtEmail)
    m_empleado.legajo = CLng(Me.txtNroLegajo)
    m_empleado.documento = CLng(Me.txtNroDocumento)
    m_empleado.Apellido = Me.txtApellido
    m_empleado.nombre = Me.txtNombre
    m_empleado.direccion = Me.txtDireccion
    m_empleado.localidad = Me.txtLocalidad
    m_empleado.Telefono1 = Me.txtTel1
    m_empleado.Telefono2 = Me.txtTel2
    m_empleado.Nombres = Me.txtNombres
    m_empleado.GrupoSanguineo = Me.txtGrupoSanguineo.text
    m_empleado.FechaIngreso = Me.dtpFechaIng.value
    m_empleado.FechaNacimiento = Me.dtpFechaNac.value
    m_empleado.estado = EstadoUsuario.activo




    If DAOEmpleados.Save(m_empleado) Then
        MsgBox "Empleado guardado.", vbInformation

        If LenB(Me.Image1.Tag) > 0 Then

            Dim clasea As New classArchivos
            If Not clasea.grabarArchivo(m_empleado.id, funciones.GetFileName(Me.Image1.Tag), CStr(Me.Image1.Tag), "Empleado", 812, False) Then GoTo err44


        End If


        Unload Me
    Else
        MsgBox "Se produjo un error", vbCritical, "Error"
    End If
    'End If

    Exit Sub
err44:
    MsgBox "Error: " & Err.Description

End Sub
Private Sub limpiar()
    txtDireccion = Empty
    txtApellido = Empty
    txtNombre = Empty
    txtLocalidad = Empty
    txtTel1 = Empty
    txtNroDocumento = Empty
    txtTel2 = Empty
    txtNroLegajo = Empty
    txtNombres = Empty
    Me.Option2.value = True
    Me.txtUsuario = Empty
    Me.txtEmail = Empty
    Me.txtGrupoSanguineo.text = vbNullString
    Me.dtpFechaIng.value = Now
    Me.dtpFechaNac.value = Now
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    verificar
End Sub
Private Sub Form_Load()
    FormHelper.Customize Me

    limpiar



End Sub



Private Sub Option1_Click()
    If Option1.value Then

    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_GotFocus()

End Sub

Private Sub PushButton1_Click()

    On Error GoTo err1
    frmPrincipal.cd.ShowOpen
    Me.Image1.Tag = frmPrincipal.cd.filename
    Set Me.Image1.Picture = LoadPicture(Me.Image1.Tag)
    Exit Sub
err1:


End Sub

Private Sub txtApellido_Change()
    verificar
    If Trim(Me.txtApellido) <> Empty And Trim(Me.txtNombre) <> Empty Then
        If m_empleado Is Nothing Then
            Me.txtUsuario = crearUsuario(Trim(Me.txtNombre), Trim(Me.txtApellido))
        End If
    Else
        Me.txtUsuario = Empty
    End If
End Sub
Private Sub txtApellido_GotFocus()
    foco txtApellido
End Sub
Private Sub TxtDireccion_Change()
    verificar
End Sub
Private Sub TxtDireccion_GotFocus()
    foco txtDireccion
End Sub
Private Sub TxtLocalidad_Change()
    verificar
End Sub
Private Sub TxtLocalidad_GotFocus()
    foco txtLocalidad
End Sub
Private Sub TxtNombre_Change()
    verificar
    If Trim(Me.txtApellido) <> Empty And Trim(Me.txtNombre) <> Empty Then
        If m_empleado Is Nothing Then
            Me.txtUsuario = crearUsuario(Trim(Me.txtNombre), Trim(Me.txtApellido))
        End If
    Else
        Me.txtUsuario = Empty
    End If

End Sub
Private Sub TxtNombre_GotFocus()
    foco txtNombre
End Sub

Private Sub txtNombres_GotFocus()
    foco Me.txtNombres
End Sub

Private Sub txtNroDocumento_Change()
    verificar
End Sub

Private Sub txtNroDocumento_GotFocus()
    foco txtNroDocumento
End Sub

Private Sub txtNroDocumento_Validate(Cancel As Boolean)
    If Not IsNumeric(txtNroDocumento) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub txtNroLegajo_Change()
    verificar
End Sub
Private Sub txtNroLegajo_GotFocus()
    foco txtNroLegajo
End Sub

Private Sub txtNroLegajo_Validate(Cancel As Boolean)
    If Not IsNumeric(txtNroLegajo) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub txtTel1_Change()
    verificar
End Sub
Private Sub txtTel1_GotFocus()
    foco txtTel1
End Sub
Private Sub txtTel2_Change()
    verificar
End Sub
Private Sub txtTel2_GotFocus()
    foco txtTel2
End Sub
Public Sub verificar()
    If Trim(Me.txtNroDocumento) = Empty Or Trim(Me.txtApellido) = Empty Or Trim(Me.txtDireccion) = Empty Or Trim(Me.txtLocalidad) = Empty Or Trim(Me.txtLocalidad) = Empty Or Trim(Me.txtNombre) = Empty Or Trim(Me.txtNroLegajo) = Empty Or Trim(Me.txtTel1) = Empty Or Trim(Me.txtTel2) = Empty Then
        cmdGuardar.Enabled = False
    Else
        cmdGuardar.Enabled = True
    End If
End Sub


