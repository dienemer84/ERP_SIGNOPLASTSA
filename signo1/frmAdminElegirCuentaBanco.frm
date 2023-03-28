VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdminElegirCuentaBanco 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ingrese monto depositado"
   ClientHeight    =   2715
   ClientLeft      =   2220
   ClientTop       =   3450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos ]"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtMonto 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.ComboBox cboCuentas 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.ComboBox cboBancos 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   39220
      End
      Begin VB.Label Label4 
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
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Monto "
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
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta "
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
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco "
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
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAdminElegirCuentaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clasea As New classAdministracion


Private Sub cboBancos_Click()
    llenarCuenta
    'Comentario
End Sub

Private Sub llenarCuenta()
    idb = CLng(Me.cboBancos.ItemData(Me.cboBancos.ListIndex))

    DAOCuentaBancaria.LlenarCombo Me.cboCuentas
End Sub

Private Sub Command1_Click()
    If IsNumeric(Trim(Me.txtMonto)) Then
        If MsgBox("¿Está seguro de agregar este depósito?", vbYesNo, "Confirmación") = vbYes Then



            Valor = CDbl(Trim(Me.txtMonto))
            fech = Me.DTPicker1
            bco = CLng(Me.cboBancos.ItemData(Me.cboBancos.ListIndex))
            If cboCuentas.ListCount > 0 Then
                idCuenta = CLng(Me.cboCuentas.ItemData(Me.cboCuentas.ListIndex))
            Else
                idCuenta = -1
            End If
            funciones.depositoIdCuenta = idCuenta
            funciones.depositoFecha = fech
            funciones.depositoMonto = Valor

            Unload Me
        End If
    Else
        MsgBox "Ingrese datos válidos!!", vbCritical, "Error"
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.DTPicker1 = Now
    DAOBancos.llenarComboXtremeSuite Me.cboBancos

End Sub



