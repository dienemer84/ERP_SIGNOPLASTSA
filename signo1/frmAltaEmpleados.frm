VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAltaEmpleados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleado"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   8235
   ClipControls    =   0   'False
   Icon            =   "frmAltaEmpleados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8235
   Begin XtremeSuiteControls.ComboBox cboOS 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   5400
      Width           =   5175
      _Version        =   786432
      _ExtentX        =   9128
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin VB.TextBox txtCuil 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   2475
      Width           =   3135
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   255
      Left            =   6480
      TabIndex        =   35
      Top             =   2280
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
      Left            =   1545
      TabIndex        =   7
      Top             =   2850
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   40119
   End
   Begin VB.TextBox txtGrupoSanguineo 
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Text            =   "txtGrupoSanguineo"
      Top             =   5835
      Width           =   900
   End
   Begin VB.TextBox txtNroLegajo 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtApellido 
      Height          =   285
      Left            =   1575
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1020
      Width           =   4695
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1380
      Width           =   4695
   End
   Begin VB.TextBox TxtDireccion 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3420
      Width           =   6375
   End
   Begin VB.TextBox TxtLocalidad 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3780
      Width           =   6375
   End
   Begin VB.TextBox txtTel1 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4155
      Width           =   6375
   End
   Begin VB.TextBox txtTel2 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4500
      Width           =   6375
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6285
      Width           =   1095
   End
   Begin VB.TextBox txtNroDocumento 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2100
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Si"
      Height          =   255
      Left            =   2070
      TabIndex        =   15
      Top             =   6420
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "No"
      Height          =   255
      Left            =   2670
      TabIndex        =   16
      Top             =   6420
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   180
      Width           =   1695
   End
   Begin VB.TextBox txtNombres 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1740
      Width           =   4695
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4860
      Width           =   6375
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6285
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFechaIng 
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16777217
      CurrentDate     =   40119
   End
   Begin VB.Label lblDatoActualizacion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "00/00/0000 00:00:00"
      Height          =   255
      Left            =   6360
      TabIndex        =   39
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label lblActualizacion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ultima actualización:"
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
      Left            =   4320
      TabIndex        =   38
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Obra Social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   37
      Top             =   5460
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuil"
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
      Left            =   540
      TabIndex        =   36
      Top             =   2490
      Width           =   855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   600
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
      Left            =   3600
      TabIndex        =   34
      Top             =   645
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
      Left            =   510
      TabIndex        =   33
      Top             =   2910
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
      Left            =   480
      TabIndex        =   32
      Top             =   5865
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
      Left            =   480
      TabIndex        =   31
      Top             =   645
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
      Left            =   480
      TabIndex        =   30
      Top             =   1035
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
      Left            =   480
      TabIndex        =   29
      Top             =   1395
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
      Left            =   480
      TabIndex        =   28
      Top             =   3435
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
      Left            =   480
      TabIndex        =   27
      Top             =   3795
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
      Left            =   480
      TabIndex        =   26
      Top             =   4170
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
      Left            =   480
      TabIndex        =   25
      Top             =   4515
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
      Left            =   480
      TabIndex        =   24
      Top             =   2115
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
      Left            =   270
      TabIndex        =   23
      Top             =   6420
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario"
      Height          =   255
      Left            =   900
      TabIndex        =   22
      Top             =   210
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros Nombres"
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
      Top             =   1755
      Width           =   1335
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
      Height          =   240
      Left            =   480
      TabIndex        =   20
      Top             =   4875
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



    Set m_empleado = DAOEmpleados.GetById(value.Id)

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

    Me.txtCuil = m_empleado.Cuil
    Me.lblDatoActualizacion = m_empleado.UltimaActualizacion


    If IsSomething(m_empleado.ObraSocial) Then

        Me.cboOS.ListIndex = PosIndexCbo(m_empleado.ObraSocial.Id, Me.cboOS)
    Else



        Dim tmp As ObraSocial
        Set tmp = DAOObraSocial.GetDefault()

        If IsSomething(tmp) Then
            Me.cboOS.ListIndex = PosIndexCbo(tmp.Id, Me.cboOS)
        End If

        MsgBox "El empleado no tiene asignada una obra social, se cargará una por default", vbCritical
    End If


    If IsSomething(m_empleado) Then

        Dim tmppath As String
        Dim clasea As New classArchivos
        '  Dim col As New Collection
        Dim Foto As archivo

        Set Foto = DAOArchivo.FindAll(OA_FotoEmpleado, "idPieza=" & m_empleado.Id)(1)

        If IsSomething(Foto) Then

            tmppath = clasea.exportarArchivo(Foto.Id)
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
        Dim B As Boolean
        B = True



    End If

    Dim tmpEmp As clsEmpleado
    Set tmpEmp = DAOEmpleados.GetByLegajo(CLng(Me.txtNroLegajo))

    If IsSomething(tmpEmp) Then
        If tmpEmp.Id <> m_empleado.Id Then
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

    m_empleado.Cuil = Me.txtCuil

    Set m_empleado.ObraSocial = DAOObraSocial.GetById(Me.cboOS.ItemData(Me.cboOS.ListIndex))


    m_empleado.UltimaActualizacion = Me.lblDatoActualizacion




    If DAOEmpleados.Save(m_empleado) Then
        MsgBox "Empleado guardado.", vbInformation

        If LenB(Me.Image1.Tag) > 0 Then

            Dim clasea As New classArchivos
            If Not clasea.grabarArchivo(m_empleado.Id, funciones.GetFileName(Me.Image1.Tag), CStr(Me.Image1.Tag), "Empleado", 812, False) Then GoTo err44


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


    Me.txtCuil = Empty
    Me.cboOS = Empty
    Me.lblDatoActualizacion = Now

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

    DAOObraSocial.llenarComboXtremeSuite Me.cboOS




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


