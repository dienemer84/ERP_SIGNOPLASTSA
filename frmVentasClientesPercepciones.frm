VERSION 5.00
Begin VB.Form frmVentasClientesPercepciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Percepciones"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Aplicar percepciones ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   6015
      Begin VB.CommandButton Command4 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   8
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.ListBox lstAplicadas 
         Height          =   2205
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.ListBox lstDisponibles 
         Height          =   2205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Cliente ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label lblCliente 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmVentasClientesPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grabado As Boolean
Dim clsConf As New classConfigurar
Dim vIdCliente As Long
Public Property Let idCliente(nIdCliente As Long)
    vIdCliente = nIdCliente
End Property

Private Sub Command1_Click()


    errora = False
    If Not grabado Then
        If MsgBox("¿Está seguro de actualizar?", vbYesNo, "Confirmación") = vbYes Then
            'elimino todos los datos de ese cliente
            If Not clsConf.ejecutar_comando("delete from clientesPercepciones where idCliente=" & vIdCliente) Then
                errora = True
            Else
                'grabo nuevamente los datos
                For i = 0 To Me.lstAplicadas.ListCount - 1
                    Me.lstAplicadas.ListIndex = i
                    idPercepcion = Me.lstAplicadas.ItemData(Me.lstAplicadas.ListIndex)
                    If Not clsConf.ejecutar_comando("insert into clientesPercepciones (idCliente,idPercepcion) values (" & vIdCliente & "," & idPercepcion & ")") Then
                        errora = True
                    End If
                Next



            End If



        End If
    End If
End Sub

Private Sub Command2_Click()
    If Not grabado Then
        If MsgBox("¿Está seguro de perder los cambios?", vbYesNo, "Confirmación") = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If

End Sub

Private Sub Command3_Click()
    On Error GoTo err4
    'saco de una
    id = Me.lstDisponibles.ItemData(Me.lstDisponibles.ListIndex)
    Percepcion = Me.lstDisponibles
    Me.lstDisponibles.RemoveItem Me.lstDisponibles.ListIndex

    'pongo en la otra
    Me.lstAplicadas.AddItem Percepcion
    Me.lstAplicadas.ItemData(Me.lstAplicadas.NewIndex) = id
    grabado = False
    Exit Sub
err4:
End Sub

Private Sub Command4_Click()
    On Error GoTo err4
    'saco de la otra
    id = Me.lstAplicadas.ItemData(Me.lstAplicadas.ListIndex)
    Percepcion = Me.lstAplicadas
    Me.lstAplicadas.RemoveItem Me.lstAplicadas.ListIndex
    'pongo en una
    Me.lstDisponibles.AddItem Percepcion
    Me.lstDisponibles.ItemData(Me.lstDisponibles.NewIndex) = id
    grabado = False
    Exit Sub
err4:
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    Dim rs As Recordset
    Set rs = conectar.RSFactory("select razon from clientes where id=" & vIdCliente)
    If Not rs.EOF And Not rs.BOF Then
        Me.lblCliente = Format(vIdCliente, "0000") & " - " & rs!razon
        grabado = True
    Else
        Exit Sub
    End If

    verListas
End Sub





Private Sub verListas()
    Dim rs As Recordset
    Dim rs1 As Recordset
    Set rs = conectar.RSFactory("select * from AdminConfigPercepciones")
    While Not rs.EOF
        idPercepcion = rs!id
        Set rs1 = conectar.RSFactory("select count(id) as cant from clientesPercepciones where idCliente=" & vIdCliente & " and idPercepcion=" & idPercepcion)
        If rs1!Cant = 1 Then
            Me.lstAplicadas.AddItem rs!codigo & " - " & rs!Percepcion
            Me.lstAplicadas.ItemData(Me.lstAplicadas.NewIndex) = rs!id
        Else
            Me.lstDisponibles.AddItem rs!codigo & " - " & rs!Percepcion
            Me.lstDisponibles.ItemData(Me.lstDisponibles.NewIndex) = rs!id
        End If


        rs.MoveNext
    Wend


End Sub

Private Sub lstDisponibles_Click()
    a = Me.lstDisponibles.ListIndex
End Sub
