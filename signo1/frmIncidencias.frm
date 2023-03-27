VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmVerIncidencias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Incidencias"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8445
   FillStyle       =   3  'Vertical Line
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstIncidencia 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nota..."
         Object.Width           =   8644
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2893
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Nueva Incidencia ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton Command3 
         Caption         =   "Nueva..."
         Height          =   495
         Left            =   5040
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   4
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtIncidencias 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   8055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detallar la incidencia"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmVerIncidencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vReferencia As Long
Dim baseSP As New classSignoplast
Dim vorigen As OrigenIncidencias
Public Property Let Origen(nOrigen As OrigenIncidencias)
    vorigen = nOrigen
End Property
Public Property Let referencia(nReferencia As Long)
    vReferencia = nReferencia
End Property
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de agregar la incidencia?", vbYesNo, "Confirmación") = vbYes Then
        If baseSP.agregarIncidencia(vReferencia, vorigen, UCase(Me.txtIncidencias)) Then

            Dim tipoEv As TipoEventoBroadcast: tipoEv = -1
            Select Case vorigen
            Case OrigenIncidencias.OI_OrdenesTrabajo
                tipoEv = TEB_IncidenciaOrdenTrabajo
            Case OrigenIncidencias.OI_Piezas
                tipoEv = TEB_IncidenciaPieza
            Case OrigenIncidencias.OI_OrdenesTrabajoDetalles
                tipoEv = TEB_IncidenciaDetalleOrdenTrabajo
            End Select

            If tipoEv <> -1 Then
                DAOEvento.Publish vReferencia, tipoEv
            End If

            MsgBox "Incidencia agregada con éxito!", vbInformation, "Información"
            Me.txtIncidencias = Empty

            llenarLST
        Else
            MsgBox "Se produjo un error. No se agrego incidencia!", vbCritical, "Error"
        End If
    End If
End Sub

Private Sub Command3_Click()
    Me.txtIncidencias = Empty
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    llenarLST
End Sub
Public Sub llenarLST()
    On Error GoTo err211
    Me.lstIncidencia.ListItems.Clear
    Dim x As ListItem
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select i.nota,i.fecha,u.usuario from Incidencias i inner join usuarios u on i.usuario=u.id where origen=" & vorigen & " and idReferencia=" & vReferencia)
    While Not rs.EOF
        Set x = Me.lstIncidencia.ListItems.Add(, , rs!Nota)
        x.Bold = True
        x.SubItems(1) = rs!usuario
        x.SubItems(2) = Format(rs!FEcha, "DD-MM-YYYY hh:mm")
        rs.MoveNext
    Wend
    Exit Sub
err211:
    MsgBox Err.Description
End Sub



Private Sub lstIncidencia_ItemClick(ByVal item As MSComctlLib.ListItem)
    On Error Resume Next
    If Me.lstIncidencia.ListItems.count > 0 Then
        Me.txtIncidencias = Me.lstIncidencia.selectedItem
    End If
End Sub

Private Sub txtIncidencias_GotFocus()
    foco Me.txtIncidencias
End Sub
