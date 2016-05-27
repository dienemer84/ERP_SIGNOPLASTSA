VERSION 5.00
Begin VB.Form frmCambiarPassword 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar Password"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox confirma 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox nuevo 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox actual 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label mensaje 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   7
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nuevo Password"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Password Actual"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCambiarPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sp As New classSignoplast

Private Sub actual_Change()
    valida
End Sub


Private Sub valida()
    If Trim(nuevo) = Empty Or Trim(Me.actual) = Empty Or Trim(Me.confirma) = Empty Then
        Me.Command1.Enabled = False
    Else
        Me.Command1.Enabled = True
    End If

End Sub

Private Sub actual_GotFocus()
    foco Me.actual
End Sub

Private Sub Command1_Click()
    Dim md As New classMD5
    usu = funciones.getUser
    Dim r As Recordset
    Set r = conectar.RSFactory("select password from usuarios where id=" & usu)
    c = 0
    While Not r.EOF
        c = c + 1
        r.MoveNext
    Wend

    If c = 1 Then
        r.MoveFirst
        pass = r!PassWord
        newpass = md.DigestStrToHexStr(Trim(Me.nuevo))
        actua = md.DigestStrToHexStr(Trim(actual))
        If pass = actua Then
            If Trim(Me.nuevo) = Trim(Me.confirma) Then
                'pass actual y confirmación correcta
                'cambiar password
                If sp.cambiarPass(usu, newpass) Then
                    MsgBox "Cambio exitoso!" & Chr(10) & " Debe reiniciar el sistema para que se efectuen los cambios", vbInformation, "Confirmación"
                Else
                    MsgBox "Se produjo un error. No se actualizo password!", vbCritical, "Error"
                End If
            Else
                Me.mensaje = "** Confirmación incorrecta **"
            End If
        Else
            Me.mensaje = "** Password incorrecto **"
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub confirma_Change()
    valida
End Sub

Private Sub confirma_GotFocus()
    foco Me.confirma
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    valida
End Sub

Private Sub nuevo_Change()
    valida
End Sub

Private Sub nuevo_GotFocus()
    foco Me.nuevo
End Sub
