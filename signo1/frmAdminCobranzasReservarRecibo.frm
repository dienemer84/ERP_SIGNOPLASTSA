VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminCobranzasReservarRecibo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservar Recibo..."
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   6330
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1620
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6225
      _Version        =   786432
      _ExtentX        =   10980
      _ExtentY        =   2857
      _StockProps     =   79
      Caption         =   "Datos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdCrear 
         Height          =   375
         Left            =   4710
         TabIndex        =   5
         Top             =   1125
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Crear"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtReciboNro 
         Height          =   285
         Left            =   915
         TabIndex        =   1
         Top             =   270
         Width           =   1455
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   915
         TabIndex        =   4
         Top             =   660
         Width           =   5130
         _Version        =   786432
         _ExtentX        =   9049
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente "
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
         TabIndex        =   3
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número "
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
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAdminCobranzasReservarRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseS As New classStock
Dim strsql As String
Dim clasea As New classAdministracion
Private Sub cmdCrear_Click()
    If IsNumeric(Me.txtReciboNro) Then
        If MsgBox("¿Desea crear el recibo?", vbYesNo, "Confirmación") = vbYes Then
            If Not IsSomething(DAORecibo.FindById(CLng(Me.txtReciboNro))) Then   '   claseA.existeRecibo(CLng(Me.txtReciboNro)) Then
                vIdRecibo = CLng(Me.txtReciboNro)
                vIdCliente = CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
                'FechaCreacion =funciones.datetimeFormateada(Now)' Format(Me.DTPicker1, "yyyy-mm-dd")    'funciones.datetimeFormateada(Now)
                idUsuarioCreador = funciones.getUser



                strsql = "insert into AdminRecibos (id,idCliente,fechaCreacion,idUsuarioCreador, fecha) values (" & vIdRecibo & "," & vIdCliente & ",NOW()," & idUsuarioCreador & ", NOW())"
                If Not clasea.ejecutarComando(strsql) Then
                    MsgBox "Se produjo algún error!, no se creará el recibo!", vbCritical, "Error"
                Else
                    MsgBox "Recibo creado con éxito!", vbInformation, "Información"
                    'Me.txtReciboNro = DAORecibo.proximo
                    Me.txtReciboNro = ""
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

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.txtReciboNro = DAORecibo.proximo
    Me.txtReciboNro = ""
    DAOCliente.llenarComboXtremeSuite Me.cboClientes

End Sub

