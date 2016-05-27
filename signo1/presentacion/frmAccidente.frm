VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAccidente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Accidente"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAccidente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   10305
   Begin VB.CheckBox chkHorasExtras 
      Alignment       =   1  'Right Justify
      Caption         =   "Horas Extras"
      Height          =   225
      Left            =   4275
      TabIndex        =   33
      Top             =   1485
      Width           =   1260
   End
   Begin XtremeSuiteControls.GroupBox grpResolucion 
      Height          =   1995
      Left            =   105
      TabIndex        =   22
      Top             =   3990
      Width           =   10050
      _Version        =   786432
      _ExtentX        =   17727
      _ExtentY        =   3519
      _StockProps     =   79
      Caption         =   "Resolución por Responsable de higiene y seguridad o Ingeniero"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtRecomendaciones 
         Height          =   1200
         Left            =   5745
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   570
         Width           =   4125
      End
      Begin VB.TextBox txtAgenteMaterial 
         Height          =   300
         Left            =   1995
         TabIndex        =   29
         Top             =   1470
         Width           =   3360
      End
      Begin VB.TextBox txtFormaAccidente 
         Height          =   300
         Left            =   1995
         TabIndex        =   27
         Top             =   1080
         Width           =   3360
      End
      Begin VB.TextBox txtUbicacionLesion 
         Height          =   300
         Left            =   1995
         TabIndex        =   25
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox txtNaturalezaLesion 
         Height          =   300
         Left            =   1995
         TabIndex        =   23
         Top             =   300
         Width           =   3360
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Recomendaciones para evitar repeticiones"
         Height          =   195
         Left            =   5760
         TabIndex        =   31
         Top             =   285
         Width           =   3045
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agente Material"
         Height          =   195
         Left            =   735
         TabIndex        =   30
         Top             =   1485
         Width           =   1140
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Forma del accidente"
         Height          =   195
         Left            =   435
         TabIndex        =   28
         Top             =   1095
         Width           =   1440
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ubicación de la lesión"
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   705
         Width           =   1515
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Naturaleza de la lesión"
         Height          =   195
         Left            =   255
         TabIndex        =   24
         Top             =   315
         Width           =   1620
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2070
      Left            =   4260
      TabIndex        =   8
      Top             =   1815
      Width           =   5895
      _Version        =   786432
      _ExtentX        =   10398
      _ExtentY        =   3651
      _StockProps     =   79
      Caption         =   "Factores Contribuyentes"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtOtros 
         Height          =   300
         Left            =   1845
         TabIndex        =   15
         Top             =   1650
         Width           =   3780
      End
      Begin VB.TextBox txtActoInseguro 
         Height          =   300
         Left            =   1845
         TabIndex        =   13
         Top             =   1200
         Width           =   3780
      End
      Begin VB.TextBox txtFaltaElementos 
         Height          =   300
         Left            =   1845
         TabIndex        =   11
         Top             =   750
         Width           =   3780
      End
      Begin VB.TextBox txtFallaMaquinas 
         Height          =   300
         Left            =   1845
         TabIndex        =   9
         Top             =   285
         Width           =   3780
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Otros"
         Height          =   195
         Left            =   1350
         TabIndex        =   16
         Top             =   1665
         Width           =   405
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Acto inseguro"
         Height          =   195
         Left            =   750
         TabIndex        =   14
         Top             =   1230
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Falta de elementos de protección personal"
         Height          =   405
         Left            =   180
         TabIndex        =   12
         Top             =   690
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Falla de máquinas o equipos"
         Height          =   390
         Left            =   465
         TabIndex        =   10
         Top             =   225
         Width           =   1275
      End
   End
   Begin XtremeSuiteControls.GroupBox grpInfo 
      Height          =   1260
      Left            =   105
      TabIndex        =   7
      Top             =   60
      Width           =   10065
      _Version        =   786432
      _ExtentX        =   17754
      _ExtentY        =   2222
      _StockProps     =   79
      Caption         =   "Información del Siniestro"
      UseVisualStyle  =   -1  'True
      Begin VB.Label lblSupervisor 
         AutoSize        =   -1  'True
         Caption         =   "Supervisor:"
         Height          =   195
         Left            =   4335
         TabIndex        =   21
         Tag             =   "Supervisor: "
         Top             =   585
         Width           =   825
      End
      Begin VB.Label lblAsegurado 
         AutoSize        =   -1  'True
         Caption         =   "Asegurado:"
         Height          =   195
         Left            =   4335
         TabIndex        =   20
         Tag             =   "Asegurado: "
         Top             =   285
         Width           =   840
      End
      Begin VB.Label lblFechaOcurrencia 
         AutoSize        =   -1  'True
         Caption         =   "Fecha ocurrencia:"
         Height          =   195
         Left            =   255
         TabIndex        =   19
         Tag             =   "Fecha ocurrencia: "
         Top             =   900
         Width           =   1290
      End
      Begin VB.Label lblArt 
         AutoSize        =   -1  'True
         Caption         =   "ART:"
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Tag             =   "ART: "
         Top             =   585
         Width           =   360
      End
      Begin VB.Label lblNroSiniestro 
         AutoSize        =   -1  'True
         Caption         =   "Nro Siniestro:"
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Tag             =   "Nro Siniestro: "
         Top             =   285
         Width           =   975
      End
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   1395
      Left            =   1005
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2505
      Width           =   3090
   End
   Begin VB.TextBox txtTestigos 
      Height          =   495
      Left            =   1005
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1875
      Width           =   3075
   End
   Begin VB.TextBox txtPuesto 
      Height          =   300
      Left            =   990
      TabIndex        =   3
      Top             =   1470
      Width           =   3090
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   540
      Left            =   6900
      TabIndex        =   6
      Top             =   6150
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   952
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   540
      Left            =   8580
      TabIndex        =   34
      Top             =   6150
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   952
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripción del hecho"
      Height          =   480
      Left            =   75
      TabIndex        =   2
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Testigos"
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   1875
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Puesto"
      Height          =   195
      Left            =   390
      TabIndex        =   0
      Top             =   1500
      Width           =   495
   End
End
Attribute VB_Name = "frmAccidente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_sin As SiniestroPersonal


Public Sub Cargar(sin As SiniestroPersonal)
    Set m_sin = sin
    Me.lblNroSiniestro.caption = Me.lblNroSiniestro.Tag & m_sin.NroSiniestro
    Me.lblArt.caption = Me.lblArt.Tag & m_sin.ART.nombre
    Me.lblFechaOcurrencia.caption = Me.lblFechaOcurrencia.Tag & m_sin.FechaHoraOcurrido
    Me.lblAsegurado.caption = Me.lblAsegurado.Tag & m_sin.Asegurado.NombreCompleto
    Me.lblSupervisor.caption = Me.lblSupervisor.Tag & m_sin.Supervisor.NombreCompleto

    If IsSomething(m_sin.InformeAccidente) Then
        With m_sin.InformeAccidente
            Me.txtPuesto.text = .Puesto
            Me.chkHorasExtras.value = CInt(.HsExtras) * -1
            Me.txtTestigos.text = .NombreTestigos
            Me.txtDescripcion.text = .DescripcionHecho
            Me.txtFallaMaquinas.text = .FallaMaquinasEquipos
            Me.txtFaltaElementos.text = .FaltaElementosProteccionPersonal
            Me.txtActoInseguro.text = .ActoInseguro
            Me.txtOtros.text = .Otros

            Me.txtNaturalezaLesion.text = .NaturalezaLesion
            Me.txtUbicacionLesion.text = .UbicacionLesion
            Me.txtFormaAccidente.text = .FormaAccidente
            Me.txtAgenteMaterial.text = .AgenteMaterial
            Me.txtRecomendaciones.text = .RecomendacionParaEvitarRepeticion
        End With
    End If

End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnGuardar_Click()
    Dim acc As InformeAccidente
    Dim ev As New clsEventoObserver
    If IsSomething(m_sin.InformeAccidente) Then
        Set acc = m_sin.InformeAccidente
        ev.EVENTO = modificar_
    Else
        Set acc = New InformeAccidente
        ev.EVENTO = agregar_
    End If

    acc.Puesto = Me.txtPuesto.text
    acc.HsExtras = Me.chkHorasExtras.value
    acc.NombreTestigos = Me.txtTestigos.text
    acc.DescripcionHecho = Me.txtDescripcion.text
    acc.FallaMaquinasEquipos = Me.txtFallaMaquinas.text
    acc.FaltaElementosProteccionPersonal = Me.txtFaltaElementos.text
    acc.ActoInseguro = Me.txtActoInseguro.text
    acc.Otros = Me.txtOtros.text

    acc.NaturalezaLesion = Me.txtNaturalezaLesion.text
    acc.UbicacionLesion = Me.txtUbicacionLesion.text
    acc.FormaAccidente = Me.txtFormaAccidente.text
    acc.AgenteMaterial = Me.txtAgenteMaterial.text
    acc.RecomendacionParaEvitarRepeticion = Me.txtRecomendaciones.text

    Set m_sin.InformeAccidente = acc

    If DAOSiniestroPersonal.Save(m_sin) Then
        Set ev.Elemento = m_sin
        Set ev.Originador = Me
        ev.Tipo = TS_InformeAccidente
        Channel.Notificar ev, TS_InformeAccidente

        MsgBox "Informe de accidente guardado.", vbInformation + vbOKOnly

        Unload Me
    Else
        If m_sin.InformeAccidente.id = 0 Then Set m_sin.InformeAccidente = Nothing
        MsgBox "Hubo un error al guardar el informe de accidente.", vbCritical + vbOK
    End If

End Sub

Private Sub Form_Load()
    Customize Me
    Me.grpResolucion.Enabled = Permisos.RRHHSiniestros
End Sub

