VERSION 5.00
Begin VB.Form frmConfigurarSistema 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurar Sistema"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5070
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Configuracion ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grabar"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Servidor Remoto "
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
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "SMTP Externo"
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
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmConfigurarSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clasea As New classArchivos
Public bbdd As String
Public bbdd2 As String
Dim SALIR As Boolean
Private Sub Command1_Click()
    On Error GoTo err32
    bbdd2 = LCase(Trim(Me.Text1))
    smtpE2 = LCase(Trim(Me.Text2))

    If Trim(bbdd2) <> Trim(bbdd) Or Trim(smtpE2) <> Trim(smtpE) Then
        'si se produjo algún cambio
        If MsgBox("¿Desea actualizar los cambios?", vbYesNo, "Confiramción") = vbYes Then

            'si hubo cambios entonces si, grabo y salgo.
            If bbdd2 <> bbdd Then
                If MsgBox("Se cerrará el sistema para aplicar cambios", vbYesNo, "Continuar?") = vbYes Then
                    GuardarIni App.path & "\config.ini", "Configurar", "ServidorBBDD", bbdd2
                    conectar.SetServidorBBDD bbdd2
                    SALIR = True
                End If
            End If

            'verifico el smtp de correo externo, si hubo cambios y ademas no esta vacio
            'grabo
            'If Trim(Me.Text2) > 0 And smtpE2 <> smtpE Then
            ' GuardarIni App.Path & "\config.ini", "Configurar", "ServidorSmtpExterno", smtpE2
            ' funciones.serverSMTPe = smtpE2
            'End If




            leerDatos
            If SALIR Then End
        End If
    End If

    Exit Sub
err32:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"

End Sub

Public Sub leerDatos()
    On Error Resume Next
    bbdd = conectar.GetServidorBBDD
    smtpE = "saf"    'funciones.getServerSMTPe
    Me.Text1 = bbdd

    If smtpE2 <> -1 Then
        Me.Text2 = smtpE
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    On Error GoTo err1
    Dim vruta As String
    Dim ve_Rev As Long
    Dim ve_max As Long
    Dim ve_min As Long
    Dim nombre As String
    Dim notas As String
    Dim version_information As VersionInformationType
    frmPrincipal.cd.ShowOpen
    vruta = frmPrincipal.cd.filename
    nombre = funciones.GetFileName(vruta)

    version_information = VersionInformation1(vruta)


    ' Display the version information.
    versionado = version_information.ProductVersion

    versio = Split(versionado, ".")

    ve_max = versio(0)
    ve_min = versio(1)
    ve_Rev = versio(3)


    If Not clasea.CompararConVersionActual(ve_max, ve_min, ve_Rev) Then
        If MsgBox("¿Seguro de subir la actualizacion?", vbYesNo, "Confirmación") = vbYes Then
            If clasea.cargarActualizacion(nombre, vruta, ve_max, ve_min, ve_Rev, notas) Then






            Else
                MsgBox "No se cargo la actualización!", vbCritical, "Error"
            End If
        End If
    Else
        MsgBox "La versión que actualmente está cargada, es mayor o igual a la que la que intenta cargar!", vbInformation, "Información"
    End If
    Exit Sub
err1:
    'MsgBox "Se produjo algun error al cargar la actualizacion!", vbCritical, "Error"
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    leerDatos
End Sub
Private Sub Text1_GotFocus()
    foco Me.Text1
End Sub
Private Sub Text2_GotFocus()
    foco Me.Text2
End Sub
