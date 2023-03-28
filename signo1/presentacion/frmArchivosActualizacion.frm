VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchivosActualizacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Archivos actualización..."
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Quitar Seleccionados"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Cerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstArchivos 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Archivo"
         Object.Width           =   10054
      EndProperty
   End
   Begin VB.Label id_Version 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label2"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmArchivosActualizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vId_version As Long
Dim x As ListItem


Public Property Let idVersion(nvalue As Long)
    vId_version = nvalue
End Property


Private Sub Cerrar_Click()
    If MsgBox("¿Seguro de salir?", vbYesNo, "Consulta") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Command1_Click()
    On Error GoTo er1
    frmPrincipal.CD.ShowOpen
    vruta = frmPrincipal.CD.filename
    nombre = funciones.GetFileName(vruta)
    Set x = Me.lstArchivos.ListItems.Add(, , nombre)
    x.Tag = vruta

    Exit Sub
er1:

End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de procesar?", vbYesNo, "Confirmar") = vbYes Then
        If grabar Then
            MsgBox "Carga Exitosa!", vbExclamation, "Información"
            Unload Me
        Else
            MsgBox "Se produjo algún error!", vbCritical, "Error"
        End If
    End If


End Sub

Private Sub Command3_Click()
    funciones.quitar_de_lista Me.lstArchivos
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Me.id_Version = vId_version
End Sub

Private Function grabar() As Boolean
    On Error GoTo err22
    grabar = True
    Dim My As ADODB.Stream
    Dim rs As Recordset
    For i = 1 To Me.lstArchivos.ListItems.count
        ruta = Me.lstArchivos.ListItems(i).Tag
        nombre = Me.lstArchivos.ListItems(i).text
        Set rs = conectar.RSFactory("select * from ActualizacionSistema_anexos")
        rs.AddNew
        Set My = New ADODB.Stream
        My.Type = adTypeBinary
        My.Open

        rs!id_Version = vId_version
        rs!nombre = UCase(nombre)

        My.LoadFromFile ruta
        rs!archivo = My.Read
        rs!Tamano = My.Size
        My.Close
        rs.Update
        rs.MoveNext

    Next i
    rs.Close

    Exit Function
err22:
    grabar = False
End Function
