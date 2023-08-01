VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmSiniestro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Siniestro"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSiniestro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   4695
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   570
      Left            =   1260
      TabIndex        =   20
      Top             =   6270
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   1005
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtGestor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      TabIndex        =   19
      Top             =   5820
      Width           =   2820
   End
   Begin VB.TextBox txtPrestadorMedico 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      TabIndex        =   9
      Top             =   3585
      Width           =   2820
   End
   Begin VB.TextBox txtDiagnostico 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1650
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2670
      Width           =   2820
   End
   Begin MSComCtl2.DTPicker dtpFechaOcurrido 
      Height          =   315
      Left            =   1650
      TabIndex        =   3
      Top             =   1005
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   62455811
      CurrentDate     =   40414.6993055556
   End
   Begin VB.TextBox txtNroSiniestro 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      TabIndex        =   1
      Top             =   585
      Width           =   1410
   End
   Begin XtremeSuiteControls.ComboBox cboAsegurado 
      Height          =   315
      Left            =   1650
      TabIndex        =   5
      Top             =   1455
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipoAccidente 
      Height          =   315
      Left            =   1650
      TabIndex        =   11
      Top             =   4005
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTratamiento 
      Height          =   315
      Left            =   1650
      TabIndex        =   13
      Top             =   4455
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboGravedad 
      Height          =   315
      Left            =   1650
      TabIndex        =   15
      Top             =   4905
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin MSComCtl2.DTPicker dtpRenaudaTareas 
      Height          =   315
      Left            =   1650
      TabIndex        =   17
      Top             =   5355
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   62455809
      CurrentDate     =   40414
   End
   Begin XtremeSuiteControls.ComboBox cboART 
      Height          =   315
      Left            =   1650
      TabIndex        =   21
      Top             =   135
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboSupervisor 
      Height          =   315
      Left            =   1650
      TabIndex        =   23
      Top             =   1890
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   2925
      TabIndex        =   25
      Top             =   6270
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   1005
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboSector 
      Height          =   315
      Left            =   1650
      TabIndex        =   26
      Top             =   2280
      Width           =   2835
      _Version        =   786432
      _ExtentX        =   5001
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   -1  'True
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label Label14 
      Height          =   195
      Left            =   1020
      TabIndex        =   27
      Top             =   2325
      Width           =   465
      _Version        =   786432
      _ExtentX        =   820
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Sector"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label13 
      Height          =   195
      Left            =   750
      TabIndex        =   24
      Top             =   1935
      Width           =   750
      _Version        =   786432
      _ExtentX        =   1323
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Supervisor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label12 
      Height          =   195
      Left            =   1215
      TabIndex        =   22
      Top             =   180
      Width           =   330
      _Version        =   786432
      _ExtentX        =   582
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "ART"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label11 
      Height          =   195
      Left            =   1050
      TabIndex        =   18
      Top             =   5835
      Width           =   465
      _Version        =   786432
      _ExtentX        =   820
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Gestor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label9 
      Height          =   195
      Left            =   315
      TabIndex        =   16
      Top             =   5400
      Width           =   1200
      _Version        =   786432
      _ExtentX        =   2117
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Renauda Tareas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label8 
      Height          =   195
      Left            =   810
      TabIndex        =   14
      Top             =   4950
      Width           =   705
      _Version        =   786432
      _ExtentX        =   1244
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Gravedad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   195
      Left            =   675
      TabIndex        =   12
      Top             =   4500
      Width           =   840
      _Version        =   786432
      _ExtentX        =   1482
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Tratamiento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   435
      TabIndex        =   10
      Top             =   4035
      Width           =   1080
      _Version        =   786432
      _ExtentX        =   1905
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Tipo Accidente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   3615
      Width           =   1245
      _Version        =   786432
      _ExtentX        =   2196
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Prestador Medico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   195
      Left            =   675
      TabIndex        =   6
      Top             =   2670
      Width           =   840
      _Version        =   786432
      _ExtentX        =   1482
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Diagnostico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   195
      Left            =   750
      TabIndex        =   4
      Top             =   1500
      Width           =   765
      _Version        =   786432
      _ExtentX        =   1349
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Asegurado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   1035
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Fecha Ocurrido"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   195
      Left            =   630
      TabIndex        =   0
      Top             =   615
      Width           =   900
      _Version        =   786432
      _ExtentX        =   1588
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Nro Siniestro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
   End
End
Attribute VB_Name = "frmSiniestro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sin As SiniestroPersonal

Public Sub Cargar()
    Me.caption = "Siniestro " & sin.NroSiniestro

    Me.txtNroSiniestro.text = sin.NroSiniestro
    Me.dtpFechaOcurrido.value = sin.FechaHoraOcurrido
    If IsSomething(sin.ART) Then
        Me.cboART.ListIndex = funciones.PosIndexCbo(sin.ART.Id, Me.cboART)
    Else
        Me.cboART.ListIndex = -1
    End If
    Me.cboAsegurado.ListIndex = funciones.PosIndexCbo(sin.Asegurado.Id, Me.cboAsegurado)
    Me.cboSupervisor.ListIndex = funciones.PosIndexCbo(sin.Supervisor.Id, Me.cboSupervisor)
    Me.txtDiagnostico.text = sin.Diagnostico
    Me.txtPrestadorMedico.text = sin.PrestadorMedico
    Me.cboTipoAccidente.ListIndex = funciones.PosIndexCbo(sin.TipoAccidente, Me.cboTipoAccidente)
    Me.cboGravedad.ListIndex = funciones.PosIndexCbo(sin.TipoGravedad, Me.cboGravedad)
    Me.cboTratamiento.ListIndex = funciones.PosIndexCbo(sin.TipoTratamiento, Me.cboTratamiento)

    Me.txtGestor.text = sin.Gestor
    If IsSomething(sin.Sector) Then
        Me.cboSector.ListIndex = funciones.PosIndexCbo(sin.Sector.Id, Me.cboSector)
    Else
        Me.cboSector.ListIndex = -1
    End If

    If CDbl(sin.RenaudaTareas) = 0 Then
        Me.dtpRenaudaTareas.value = Null
    Else
        Me.dtpRenaudaTareas.value = sin.RenaudaTareas
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnGuardar_Click()

    If Me.cboAsegurado.ListIndex = -1 Or _
       Me.cboSupervisor.ListIndex = -1 Or _
       LenB(Me.txtNroSiniestro.text) = 0 Or _
       Me.cboART.ListIndex = -1 _
       Then
        MsgBox "Falta completar datos obligatorios (asegurado, supervisor, nº siniestro, art).", vbExclamation + vbOKOnly
        Exit Sub
    End If

    If Not IsSomething(sin) Then Set sin = New SiniestroPersonal

    Set sin.Asegurado = DAOEmpleados.GetById(Me.cboAsegurado.ItemData(Me.cboAsegurado.ListIndex))
    Set sin.Supervisor = DAOEmpleados.GetById(Me.cboAsegurado.ItemData(Me.cboSupervisor.ListIndex))
    Set sin.ART = DAOART.FindAll("id = " & Me.cboART.ItemData(Me.cboART.ListIndex))(1)
    sin.Diagnostico = Me.txtDiagnostico.text
    sin.FechaHoraOcurrido = Me.dtpFechaOcurrido.value
    sin.Gestor = Me.txtGestor.text
    sin.NroSiniestro = Me.txtNroSiniestro.text
    sin.PrestadorMedico = Me.txtPrestadorMedico.text

    If IsNull(Me.dtpRenaudaTareas) Then
        sin.RenaudaTareas = 0
    Else
        sin.RenaudaTareas = Me.dtpRenaudaTareas.value
    End If


    If Me.cboTipoAccidente.ListIndex <> -1 Then
        sin.TipoAccidente = Me.cboTipoAccidente.ItemData(Me.cboTipoAccidente.ListIndex)
    Else
        sin.TipoAccidente = -1
    End If

    If Me.cboGravedad.ListIndex <> -1 Then
        sin.TipoGravedad = Me.cboGravedad.ItemData(Me.cboGravedad.ListIndex)
    Else
        sin.TipoGravedad = -1
    End If

    If Me.cboTratamiento.ListIndex <> -1 Then
        sin.TipoTratamiento = Me.cboTratamiento.ItemData(Me.cboTratamiento.ListIndex)
    Else
        sin.TipoTratamiento = -1
    End If

    If Me.cboSector.ListIndex <> -1 Then
        Set sin.Sector = DAOSectores.GetById(Me.cboSector.ItemData(Me.cboSector.ListIndex))
    Else
        Set sin.Sector = Nothing
    End If


    If DAOSiniestroPersonal.Save(sin) Then

        Dim ev As New clsEventoObserver
        ev.EVENTO = agregar_
        Set ev.Elemento = sin
        Set ev.Originador = Me
        ev.Tipo = TS_InformeAccidente
        Channel.Notificar ev, TS_InformeAccidente

        MsgBox "Siniestro guardado.", vbInformation + vbOKOnly
        Unload Me
    Else
        MsgBox "Error al guardar el siniestro.", vbCritical + vbOKOnly
    End If
End Sub

Private Sub Form_Load()
    Customize Me

    Dim emps As Collection
    Dim emp As clsEmpleado
    Set emps = DAOEmpleados.GetAll()

    Me.dtpFechaOcurrido.value = Now
    Me.dtpRenaudaTareas.value = Null

    Me.cboAsegurado.Clear
    For Each emp In emps
        Me.cboAsegurado.AddItem emp.NombreCompleto & " (Leg " & emp.legajo & ")"
        Me.cboAsegurado.ItemData(Me.cboAsegurado.NewIndex) = emp.Id
    Next emp

    Me.cboSupervisor.Clear
    For Each emp In emps
        Me.cboSupervisor.AddItem emp.NombreCompleto & " (Leg " & emp.legajo & ")"
        Me.cboSupervisor.ItemData(Me.cboSupervisor.NewIndex) = emp.Id
    Next emp

    Dim K As Variant
    Me.cboTipoAccidente.Clear
    For Each K In enums.TiposAccidente.Keys
        Me.cboTipoAccidente.AddItem enums.TiposAccidente.item(K)
        Me.cboTipoAccidente.ItemData(Me.cboTipoAccidente.NewIndex) = K
    Next K

    Me.cboGravedad.Clear
    For Each K In enums.TiposGravedad.Keys
        Me.cboGravedad.AddItem enums.TiposGravedad.item(K)
        Me.cboGravedad.ItemData(Me.cboGravedad.NewIndex) = K
    Next K

    Me.cboTratamiento.Clear
    For Each K In enums.TiposTratamiento.Keys
        Me.cboTratamiento.AddItem enums.TiposTratamiento.item(K)
        Me.cboTratamiento.ItemData(Me.cboTratamiento.NewIndex) = K
    Next K

    Dim A As ART
    Me.cboART.Clear
    For Each A In DAOART.FindAll
        Me.cboART.AddItem A.nombre
        Me.cboART.ItemData(Me.cboART.NewIndex) = A.Id
    Next A

    DAOSectores.LlenarComboXtreme Me.cboSector


End Sub

