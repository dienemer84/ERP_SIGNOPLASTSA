VERSION 5.00
Begin VB.Form frmModificarStock 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modifiacar stock..."
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtUbicacion 
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Default         =   -1  'True
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Accion"
         Height          =   975
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Egreso"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ingreso"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox txtCantaIngresar 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ubicación "
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad a Modificar "
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
         Width           =   1935
      End
      Begin VB.Label lblStockActual 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label3"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stock Actual "
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
         Width           =   1935
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label2"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
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
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label lblIStock 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
End
Attribute VB_Name = "frmModificarStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vIdPieza As Long
Dim baseS As New classStock
Public Property Let idPieza(nIdPieza As Long)
    vIdPieza = nIdPieza
End Property

Private Sub Command1_Click()
    Set baseS = New classStock


    g = MsgBox("¿Confirma actualización de stock?", vbYesNo, "Confirmación")
    If g = vbYes Then

        If Option1.value Then
            baseS.modifica_Stock vIdPieza, 1, -2, CInt(Me.txtCantaIngresar), Trim(Me.txtUbicacion)
            Unload Me
        Else
            If CInt(Me.lblStockActual) >= CInt(Me.txtCantaIngresar) Then
                baseS.modifica_Stock vIdPieza, 2, -1, CInt(Me.txtCantaIngresar), Trim(Me.txtUbicacion)
                Unload Me
            Else
                MsgBox "No puede restar mas de los que tiene", vbCritical, "Error"

            End If
        End If
    End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim r As Recordset
    Set r = conectar.RSFactory("select s.detalle_stock,s.cantidad,s.detalle,c.razon from stock s inner join clientes c on s.id_cliente=c.id  where s.id=" & vIdPieza)    '
    If Not r.EOF And Not r.BOF Then
        Me.txtUbicacion = UCase(r!detalle_stock)
        lblStockActual = r!Cantidad
        Frame1.caption = "[ " & r!detalle & " ]"
        lblCliente = r!razon

    End If


    Set r = Nothing
    
    ''Me.caption = caption & " (" & Name & ")"
    
End Sub

Private Sub txtCantaIngresar_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtCantaIngresar) Then
        Cancel = True
    Else
        If CInt(Me.txtCantaIngresar) < 0 Then
            Cancel = True
        Else
            Cancel = False
        End If
    End If
End Sub
