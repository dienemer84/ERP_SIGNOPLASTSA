VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAdminIIBB 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Padrón IIBB"
   ClientHeight    =   6240
   ClientLeft      =   7995
   ClientTop       =   1575
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar R en CABA"
      Height          =   375
      Index           =   2
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizar Padrón Unificado (CABA)"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5280
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar P en Bs.As."
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizar Padrón P (Bs.As.)"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5280
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3240
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6225
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizar Padrón R (Bs.As.)"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   2655
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
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   6975
      Begin MSComCtl2.DTPicker Fpublicacion 
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483639
         CalendarTrailingForeColor=   -2147483639
         Format          =   58392576
         CurrentDate     =   39421
      End
      Begin MSComCtl2.DTPicker Fdesde 
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   58392577
         CurrentDate     =   39421
      End
      Begin MSComCtl2.DTPicker Fhasta 
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   58392577
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
         TabIndex        =   24
         Top             =   2955
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3840
         TabIndex        =   21
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblGrupo 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label lblAlicuota 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label lblCambio 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         ToolTipText     =   "'S' - Cambió 'N' - No Cambió"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label lblAltaBaja 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         ToolTipText     =   "'S' - Alta  'N' - Baja"
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label lblTipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   " "
         Height          =   255
         Left            =   2400
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar P en CABA"
         Height          =   375
         Index           =   1
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1080
         Width           =   1515
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar R en Bs.As."
         Height          =   375
         Index           =   0
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtCuit 
         Height          =   285
         Left            =   1725
         TabIndex        =   2
         Top             =   570
         Width           =   2865
      End
      Begin XtremeSuiteControls.ComboBox cboPadron 
         Height          =   315
         Left            =   1725
         TabIndex        =   26
         Top             =   210
         Width           =   2880
         _Version        =   786432
         _ExtentX        =   5080
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   -1  'True
         Text            =   "cboPadron"
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Padrón a utilizar"
         Height          =   195
         Left            =   480
         TabIndex        =   25
         Top             =   255
         Width           =   1125
      End
      Begin VB.Label In 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ingrese Nro CUIT"
         Height          =   255
         Left            =   375
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3840
      X2              =   3480
      Y1              =   1680
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3480
      X2              =   3480
      Y1              =   5280
      Y2              =   6240
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
        Set rs = conectar.RSFactory("select * from sp_permisos." & tabla & " where cuit=" & CDbl(Me.txtCuit))

        If Not rs.EOF And Not rs.BOF Then

            If rs!Discriminador = "R" Then
                Me.frame2.caption = "[ Resultado RETENCIONES ]"
            ElseIf rs!Discriminador = "P" Then
                Me.frame2.caption = "[ Resultado PERCEPCIONES ]"
            Else
                Me.frame2.caption = "[ Sin resultado ]"
            End If
            Me.lblAltaBaja = rs!AltaBaja
            Me.lblCambio = rs!Cambio
            Me.lblGrupo = rs!Grupo
            Me.lblAlicuota = rs!Alicuota
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
    Me.CD.ShowOpen
    filename = CD.filename
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

Private Sub Command3_Click(index As Integer)
    On Error GoTo err4
    Dim strsql As String
    Dim filename As String
    Me.CD.ShowOpen
    filename = CD.filename
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
    Me.CD.ShowOpen
    filename = CD.filename
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
    Me.CD.ShowOpen
    filename = CD.filename
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



Private Sub Form_Load()
    Me.cboPadron.Clear

    cboPadron.AddItem "Actual"
    Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 0
    cboPadron.AddItem "Anterior"
    Me.cboPadron.ItemData(Me.cboPadron.NewIndex) = 1

    Me.cboPadron.ListIndex = 0

    FormHelper.Customize Me
    If Permisos.AdminIIBB Then
        'Me.Command3.Enabled = True
    Else
        'Me.Command3.Enabled = False
    End If
End Sub

