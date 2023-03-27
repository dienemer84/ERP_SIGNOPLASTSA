VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasReservarReciboAnticipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservar Recibo de Anticipo..."
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2449.2
   ScaleMode       =   0  'User
   ScaleWidth      =   6270
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6225
      _Version        =   786432
      _ExtentX        =   10980
      _ExtentY        =   3810
      _StockProps     =   79
      Caption         =   "Datos"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cancelar"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtReciboNro 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   390
         Width           =   1455
      End
      Begin XtremeSuiteControls.PushButton cmdCrear 
         Height          =   375
         Left            =   4755
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Crear"
         BackColor       =   12632256
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   900
         Width           =   5130
         _Version        =   786432
         _ExtentX        =   9049
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número:"
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
         Left            =   135
         TabIndex        =   5
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente:"
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
         Top             =   930
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAdminCobranzasReservarReciboAnticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseS As New classStock
Dim strsql As String
Dim clasea As New classAdministracion

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdCrear_Click()
    If IsNumeric(Me.txtReciboNro) Then
        If MsgBox("¿Desea crear el recibo?", vbYesNo, "Confirmación") = vbYes Then
            If Not IsSomething(DAOReciboAnticipo.FindById(CLng(Me.txtReciboNro))) Then   '   claseA.existeRecibo(CLng(Me.txtReciboNro)) Then
                vIdRecibo = CLng(Me.txtReciboNro)
                vIdCliente = CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
                'FechaCreacion =funciones.datetimeFormateada(Now)' Format(Me.DTPicker1, "yyyy-mm-dd")    'funciones.datetimeFormateada(Now)
                idUsuarioCreador = funciones.getUser



                strsql = "insert into AdminRecibosAnticipo (id,idCliente,fechaCreacion,idUsuarioCreador, fecha) values (" & vIdRecibo & "," & vIdCliente & ",NOW()," & idUsuarioCreador & ", NOW())"
                If Not clasea.ejecutarComando(strsql) Then
                    MsgBox "Se produjo algún error!, no se creará el recibo!", vbCritical, "Error"
                Else
                    MsgBox "Recibo creado con éxito!", vbInformation, "Información"
                    Me.txtReciboNro = DAOReciboAnticipo.proximo
                    'Me.txtReciboNro = ""
                    DAOCliente.llenarComboXtremeSuite Me.cboClientes
                End If
            Else
                MsgBox "El recibo indicado ya existe en la BBDD!", vbCritical, "Error"
                Exit Sub
            End If
        End If
    Else
        MsgBox "Ingrese datos válidos!", vbCritical, "Error"
    End If

End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    Me.txtReciboNro = DAOReciboAnticipo.proximo
    'Me.txtReciboNro = ""
    DAOCliente.llenarComboXtremeSuite Me.cboClientes

End Sub

