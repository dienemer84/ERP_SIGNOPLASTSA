VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminIIBB 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Padrón IIBB"
   ClientHeight    =   7080
   ClientLeft      =   5700
   ClientTop       =   3015
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Resultado ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   120
      TabIndex        =   35
      Top             =   2160
      Width           =   6975
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta:"
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
         Left            =   4080
         TabIndex        =   50
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vigencia Per desde:"
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
         TabIndex        =   49
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblAlicuotaR 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   48
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alicuota Retención:"
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
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta:"
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
         Left            =   4080
         TabIndex        =   46
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vigencia Ret desde:"
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
         TabIndex        =   45
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblAlicuotaP 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   44
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alicuota Percepción:"
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
         TabIndex        =   43
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblVigenciaDesdeP 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   42
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblVigenciaDesdeR 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblVigenciaHastaP 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4800
         TabIndex        =   40
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblVigenciaHastaR 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4800
         TabIndex        =   39
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000000&
         X1              =   360
         X2              =   6600
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   360
         X2              =   6600
         Y1              =   1100
         Y2              =   1100
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "LA CONSULTA GENERÓ UN RESULTADO CON FECHA VENCIDA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2280
         Visible         =   0   'False
         Width           =   6495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Paso 2 - Procesar Padrones ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      TabIndex        =   30
      Top             =   4920
      Width           =   3615
      Begin VB.CommandButton btnProcesar 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Procesar Padrones"
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   480
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3600
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Resultado ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   7320
      TabIndex        =   4
      Top             =   2280
      Width           =   6975
      Begin MSComCtl2.DTPicker Fpublicacion 
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483639
         CalendarTrailingForeColor=   -2147483639
         Format          =   60096512
         CurrentDate     =   39421
      End
      Begin MSComCtl2.DTPicker Fdesde 
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60096513
         CurrentDate     =   39421
      End
      Begin MSComCtl2.DTPicker Fhasta 
         Height          =   255
         Left            =   4440
         TabIndex        =   20
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60096513
         CurrentDate     =   39421
      End
      Begin VB.Label lblVencida 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "LA CONSULTA GENERÓ UN RESULTADO CON FECHA VENCIDA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2955
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vigencia desde"
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
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Publicacion"
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
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblGrupo 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label lblAlicuota 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label lblCambio 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         ToolTipText     =   "'S' - Cambió 'N' - No Cambió"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label lblAltaBaja 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         ToolTipText     =   "'S' - Alta  'N' - Baja"
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label lblTipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   " "
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         ToolTipText     =   "'C' - Convenio Multilateral 'D' Directo PCIA Bs As"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nro Grupo"
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
         TabIndex        =   9
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alicuota"
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
         TabIndex        =   8
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cambio Alícuota"
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
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alta - Baja Sujeto"
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
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Cont Inscr"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Búsqueda ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar P en Bs.As."
         Default         =   -1  'True
         Height          =   375
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   960
         Width           =   1515
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar R en CABA"
         Height          =   375
         Index           =   2
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1680
         Width           =   1515
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Left            =   5160
         TabIndex        =   34
         Top             =   1320
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar P en CABA"
         Height          =   375
         Index           =   1
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1515
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar R en Bs.As."
         Height          =   375
         Index           =   0
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txtCuit 
         Height          =   285
         Left            =   1965
         TabIndex        =   2
         Top             =   360
         Width           =   2865
      End
      Begin XtremeSuiteControls.ComboBox cboPadron 
         Height          =   315
         Left            =   1965
         TabIndex        =   24
         Top             =   1320
         Width           =   2880
         _Version        =   786432
         _ExtentX        =   5080
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboRegion 
         Height          =   315
         Left            =   1965
         TabIndex        =   33
         Top             =   840
         Width           =   2880
         _Version        =   786432
         _ExtentX        =   5080
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Appearance      =   6
         Text            =   "Seleccione..."
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Padrón:"
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
         Left            =   1155
         TabIndex        =   32
         Top             =   900
         Width           =   660
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Versión a utilizar:"
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
         Left            =   345
         TabIndex        =   23
         Top             =   1380
         Width           =   1470
      End
      Begin VB.Label In 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ingrese Nro CUIT:"
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
         TabIndex        =   1
         Top             =   375
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[Paso 1 - Importar Padrones ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   3255
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Unificado (CABA)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Retenciones (BS.AS.)"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Percepciones (BS.AS.)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   480
      X2              =   6720
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "frmAdminIIBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New classAdministracion
Dim rs As Recordset




Private Sub MostrarResultado(tabla As String, Cuit As String)

    If IsNumeric(Cuit) Then
        Set rs = conectar.RSFactory("select * from sp_permisos.padron_detalles where cuit=" & CDbl(Me.txtCuit))

        If Not rs.EOF And Not rs.BOF Then

            If rs!Discriminador = "R" Then
                Me.frame2.caption = "[ Resultado RETENCIONES Padrón Buenos Aires]"
            ElseIf rs!Discriminador = "P" Then
                Me.frame2.caption = "[ Resultado PERCEPCIONES Padrón Buenos Aires]"
            Else
                Me.frame2.caption = "[ Sin resultado ]"
            End If
            Me.lblAltaBaja = rs!AltaBaja
            Me.lblCambio = rs!Cambio
            Me.lblGrupo = rs!Grupo
            Me.lblAlicuota = rs!alicuota
            Me.lblTipo = rs!Tipo

            FechaDesde = rs!FechaDesde
            f_desde_anio = Right(FechaDesde, 4)
            f_desde_mes = Mid(FechaDesde, 3, 2)
            f_desde_dia = Mid(FechaDesde, 1, 2)
            Me.Fdesde = f_desde_dia & "/" & f_desde_mes & "/" & f_desde_anio

            FechaHasta = rs!FechaHasta
            f_hasta_anio = Right(FechaHasta, 4)
            f_hasta_mes = Mid(FechaHasta, 3, 2)
            f_hasta_dia = Mid(FechaHasta, 1, 2)
            Me.Fhasta = f_hasta_dia & "/" & f_hasta_mes & "/" & f_hasta_anio


            fechapub = rs!FechaPublicacion
            f_pub_anio = Right(fechapub, 4)
            f_pub_mes = Mid(fechapub, 3, 2)
            f_pub_dia = Mid(fechapub, 1, 2)
            Me.Fpublicacion = f_pub_dia & "/" & f_pub_mes & "/" & f_pub_anio


            If Now() > Fhasta Then
                Me.lblVencida.Visible = True
            Else
                Me.lblVencida.Visible = False
            End If



        Else
            MsgBox "sin coincidencias!"
        End If
    End If





End Sub



Private Sub MostrarResultado2(Cuit As String, IdPadron As String, tabla As String)

    If IsNumeric(Cuit) And IdPadron <> "-1" Then
        Set rs = conectar.RSFactory("SELECT * FROM sp_permisos." & tabla & " pd INNER JOIN sp_permisos.Padron_Config pc ON pd.Padron = pc.id  WHERE pc.id=" & IdPadron & " and cuit=" & CDbl(Me.txtCuit))

        If Not rs.EOF And Not rs.BOF Then

'
              Me.frame2.caption = "[ Resultado Padrón  " & rs!detalle & "]"
              

          '   Me.lblGrupo = rs!GrupoPercepcion
            Me.lblAlicuotaP = rs!alicuotaPercepcion

              
              
           
            Me.lblAlicuotaR = rs!alicuotaRetencion

                       If Not IsNull(rs.Fields("FechaDesdePercepcion").value) Then
                        Me.lblVigenciaDesdeP = ConvertirAFecha(rs!FechaDesdePercepcion)
                       End If
                       
                       If Not IsNull(rs.Fields("FechaDesdeRetencion").value) Then
                        Me.lblVigenciaDesdeR = ConvertirAFecha(rs!FechaDesdeRetencion)
                       End If
                       
                       If Not IsNull(rs.Fields("FechaHastaPercepcion").value) Then
                      Me.lblVigenciaHastaP = ConvertirAFecha(rs!FechaHastaPercepcion)
                       End If
                       
                        If Not IsNull(rs.Fields("FechaHastaRetencion").value) Then
                        Me.lblVigenciaHastaR = ConvertirAFecha(rs!FechaHastaRetencion)
                       End If
            
'             Me.lblVigenciaDesdeR = rs!FechaDesdeRetencion
'             Me.lblVigenciaHastaP = rs!FechaHastaPercepcion
'             Me.lblVigenciaHastaR = rs!FechaHastaRetencion
             
'            f_desde_anio = Right(FechaDesde, 4)
'            f_desde_mes = Mid(FechaDesde, 3, 2)
'            f_desde_dia = Mid(FechaDesde, 1, 2)
'            Me.Fdesde = f_desde_dia & "/" & f_desde_mes & "/" & f_desde_anio

             '  Me.Fhasta = rs!FechaHasta
'            f_hasta_anio = Right(FechaHasta, 4)
'            f_hasta_mes = Mid(FechaHasta, 3, 2)
'            f_hasta_dia = Mid(FechaHasta, 1, 2)
'            Me.Fhasta = f_hasta_dia & "/" & f_hasta_mes & "/" & f_hasta_anio


'            fechapub = rs!FechaPublicacion
'            f_pub_anio = Right(fechapub, 4)
'            f_pub_mes = Mid(fechapub, 3, 2)
'            f_pub_dia = Mid(fechapub, 1, 2)
'            Me.Fpublicacion = f_pub_dia & "/" & f_pub_mes & "/" & f_pub_anio


'            If Now() > Fhasta Then
'                Me.lblVencida.Visible = True
'            Else
'                Me.lblVencida.Visible = False
'            End If



        Else
            MsgBox "sin coincidencias!"
        End If
    End If





End Sub
Public Function ConvertirAFecha(entrada As String) As String
Dim FEcha As String
Dim f_anio As String, f_mes As String, f_dia As String
f_anio = Right(entrada, 4)
f_mes = Mid(entrada, 3, 2)
 f_dia = Mid(entrada, 1, 2)
ConvertirAFecha = f_dia & "/" & f_mes & "/" & f_anio
End Function



Private Sub Command1s_Click()
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Retencion"
    Else
        tabla = "IIBB2_RetencionAnt"
    End If
    MostrarResultado tabla, Me.txtCuit

End Sub
Private Sub Command1ss_Click()
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Retencion"
    Else
        tabla = "IIBB2_RetencionAnt"
    End If
    MostrarResultado tabla, Me.txtCuit

End Sub


Private Sub Command3ss_Click()
    On Error GoTo err4
    Dim strsql As String
    Dim filename As String
    Me.cd.ShowOpen
    filename = cd.filename
    filename = Replace(filename, "\", "/")
    If MsgBox("¿Está seguro de continuar?", vbYesNo, "Confirmación") = vbYes Then
        If c.ActualizarPadronIB(filename, TipoPadronRetencion) Then
            MsgBox "Actualización exitosa!", vbInformation, "Información"
        Else
            MsgBox "Error, la actualización no se efectuó!", vbInformation, "Información"
        End If
    End If
    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"

End Sub



Private Sub btnBuscar_Click()

    'Dim F As New frmLoading
    

    Dim tabla As String


    If Me.cboPadron.ListIndex = 1 Then
    
        'PADRON ACTUAL
        tabla = "Padron_Detalles_Ant"
            
        Else
        'PADRON ANTERIOR'
            tabla = "Padron_Detalles"
    
    End If

    If Me.cboRegion.ListIndex = -1 Then
    
    MsgBox ("Debe seleccionar el padrón correspondiente")
    
    Else
    
   
    MostrarResultado2 Me.txtCuit, Me.cboRegion.ItemData(Me.cboRegion.ListIndex), tabla
    
    Me.Frame5.Enabled = True

    
    End If
    

    
    
End Sub

Private Sub btnProcesar_Click()
    
    On Error GoTo err4

    If MsgBox("¿Está seguro de continuar con el procesamiento?", vbYesNo, "Confirmación") = vbYes Then
        If c.ProcesarPadronIB() Then
            MsgBox "Procesamiento de Padrones éxitosa!", vbInformation, "Información"
        Else
            MsgBox "Error, el procesamiento no se efectuó!", vbInformation, "Información"
        End If
    End If
    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"

End Sub



Private Sub Command1_Click(index As Integer)
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Retencion"
    Else
        tabla = "IIBB2_RetencionAnt"
    End If

    MostrarResultado tabla, Me.txtCuit

End Sub

Private Sub Command2_Click(index As Integer)
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Padron_CABA"
    Else
        tabla = "IIBB2_Padron_CABA_Ant"
    End If

    ' MostrarResultadoCABA tabla, Me.txtCuit, "P"

End Sub

Private Sub Command3_Click(index As Integer)
    On Error GoTo err4
    Dim strsql As String
    Dim filename As String
    Me.cd.ShowOpen
    filename = cd.filename
    filename = Replace(filename, "\", "/")
    If MsgBox("¿Está seguro de continuar?", vbYesNo, "Confirmación") = vbYes Then
        If c.ActualizarPadronIB(filename, TipoPadronRetencion) Then
            MsgBox "Actualización exitosa!", vbInformation, "Información"
        Else
            MsgBox "Error, la actualización no se efectuó!", vbInformation, "Información"
        End If
    End If
    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    On Error GoTo err4
    Dim strsql As String
    Dim filename As String
    Me.cd.ShowOpen
    filename = cd.filename
    filename = Replace(filename, "\", "/")
    If MsgBox("¿Está seguro de continuar?", vbYesNo, "Confirmación") = vbYes Then
        If c.ActualizarPadronIB(filename, TipoPadronPercepcion) Then
            MsgBox "Actualización exitosa!", vbInformation, "Información"
        Else
            MsgBox "Error, la actualización no se efectuó!", vbInformation, "Información"
        End If
    End If
    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"
End Sub

Private Sub Command6_Click()
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Percepcion"
    Else
        tabla = "IIBB2_PercepcionAnt"
    End If

    MostrarResultado tabla, Me.txtCuit


End Sub

'BotonNuevoNemer-Actualizar Padron (CABA)
Private Sub Command7_Click()
    On Error GoTo err4
    Dim strsql As String
    Dim filename As String
    Me.cd.ShowOpen
    filename = cd.filename
    filename = Replace(filename, "\", "/")
    If MsgBox("¿Está seguro de continuar?", vbYesNo, "Confirmación") = vbYes Then
        If c.ActualizarPadronIB(filename, TipoPadronUnificadoCABA) Then
            MsgBox "Actualización exitosa del padrón CABA!", vbInformation, "Información"
        Else
            MsgBox "Error, la actualización no se efectuó!", vbInformation, "Información"
        End If
    End If
    Exit Sub
err4:
    If Err.Number <> 32755 Then MsgBox "Se produjo algun error!", vbCritical, "Error"
End Sub



Private Sub Command9_Click(index As Integer)
    Dim tabla As String
    If Me.cboPadron.ListIndex = 0 Then
        tabla = "IIBB2_Padron_CABA"
    Else
        tabla = "IIBB2_Padron_CABA_Ant"
    End If

   ' MostrarResultadoCABA tabla, Me.txtCuit, "R"
End Sub

Private Sub Form_Load()

   'Me.Left = (Screen.Width - Me.Width) / 2
   'Me.Top = (Screen.Height - Me.Height) / 5
    
    Me.cboPadron.Clear
    cboPadron.AddItem "Actual"
    Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 0
    cboPadron.AddItem "Anterior"
    Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 1
    Me.cboPadron.ListIndex = 0
    
'    Me.cboTipo.Clear
'    cboTipo.AddItem "Percepciones"
'    Me.cboTipo.ItemData(Me.cboTipo.NewIndex) = 0
'    cboTipo.AddItem "Retenciones"
'    Me.cboTipo.ItemData(Me.cboTipo.NewIndex) = 1
'    'Me.cboTipo.ListIndex = 0
    
    Me.cboRegion.Clear
    cboRegion.AddItem "CABA"
    Me.cboRegion.ItemData(Me.cboRegion.NewIndex) = 2
    cboRegion.AddItem "BUENOS AIRES"
    Me.cboRegion.ItemData(Me.cboRegion.NewIndex) = 1
    'Me.cboRegion.ListIndex = 0

    FormHelper.Customize Me
    If Permisos.AdminIIBB Then
        'Me.Command3.Enabled = True
    Else
        'Me.Command3.Enabled = False
    End If
    
   
End Sub

Private Sub TaskDialog1_ButtonClicked(ByVal id As Long, CloseDialog As Variant)

End Sub
