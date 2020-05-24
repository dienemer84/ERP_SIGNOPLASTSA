VERSION 5.00
Begin VB.Form frmComprasOrdenesNueva 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva OC..."
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8430
   ClipControls    =   0   'False
   Icon            =   "frmComprasNuevaOrden.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   8430
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Datos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear"
         Default         =   -1  'True
         Height          =   375
         Left            =   7185
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1065
         Width           =   1095
      End
      Begin VB.ComboBox cboProveedores 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   7095
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   7440
         TabIndex        =   6
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
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
         Left            =   6840
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor "
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
         Top             =   600
         Width           =   1095
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmComprasOrdenesNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grabado As Boolean
Dim vAccion As Long
Dim claseC As New classCompras
Public Property Let accion(naccion As Long)
    vAccion = naccion
End Property

Private Sub btnSalir_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    If Me.cboProveedores.ListIndex = -1 Then Exit Sub
    
    Dim idProveedor As Long
    If MsgBox("¿Desea crear la orden?", vbYesNo, "Confirmación") = vbYes Then
        idProveedor = Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)
        If claseC.CrearOrden(idProveedor, Me.lblFecha) Then
            MsgBox "Creación exitosa!", vbInformation, "Información"
            Me.txtNumero = Format(proximaOC, "0000")
        End If
    End If
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    Me.lblFecha = Format(Now, "dd-mm-yyyy")
    'claseC.llenarComboProveedores Me.cboProveedores, 0
    Me.txtNumero = Format(proximaOC, "0000")
End Sub
Private Function proximaOC() As Long
    Dim r As Recordset
    Set r = conectar.RSFactory("select max(id)+1 as proxima from ComprasOrdenes")
    If Not r.EOF And Not r.BOF Then
        If Not IsNumeric(r!proxima) Then
            proxima = 1
        Else
            proxima = r!proxima
        End If
    End If
    Set r = Nothing
    proximaOC = proxima
End Function
