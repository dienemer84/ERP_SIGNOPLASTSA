VERSION 5.00
Begin VB.Form frmConfirmacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmar..."
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "[ Detalles ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   5415
         Begin VB.TextBox txtTotalFijo 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   2280
            Width           =   3855
         End
         Begin VB.TextBox txtTotalCambio 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1920
            Width           =   3855
         End
         Begin VB.TextBox txtTotalMDO 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1560
            Width           =   3855
         End
         Begin VB.TextBox txtUSSTotales 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox txtM2Totales 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtKgTotales 
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label Label8 
            Caption         =   "Total Fijos"
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
            TabIndex        =   11
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Total Cambio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Total Mdo"
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
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "u$S Totales"
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
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "M2. Totales"
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
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Kg. Totales"
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
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Volver"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label lblIdCliente 
         Caption         =   "Label9"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblCliente 
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblDetalle 
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
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
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Detalle"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim base As classConfigurar
Private Sub Command1_Click()
g = MsgBox("¿Confirma la nueva pieza?", vbYesNo, "Confirmación")
If g = 6 Then
strSQL = "insert into stock (detalle,id_cliente,KgTotales,M2Totales,UssTotales,MdoTotales,FijoTotales,CambioTotales,cantidad) VALUES ('" & Me.lblDetalle & "'," & CInt(Me.lblIdCliente) & "," & CDbl(Me.txtKgTotales) & "," & CDbl(Me.txtM2Totales) & "," & CDbl(Me.txtUSSTotales) & "," & CDbl(Me.txtTotalMDO) & "," & CDbl(Me.txtTotalFijo) & "," & CDbl(Me.txtTotalCambio) & ",1)"
Set base = New classConfigurar
base.ejecutar_consulta (strSQL)
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
