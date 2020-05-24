VERSION 5.00
Begin VB.Form frmPlaneamientoOrganizacionProcesosModificar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificar datos  de planeamiento"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDuracion 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblTarea 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Duración"
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
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tarea"
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
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmPlaneamientoOrganizacionProcesosModificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rss As recordset
Dim grabado As Boolean
Dim plan As New classPlaneamiento
Dim vidDetalleTiempos As Long
Dim vDuracion As Double
Public Property Let idDetalleTiempo(nIdDetalleTiempo)
    vidDetalleTiempos = nIdDetalleTiempo
End Property

Public Property Let duracion(nDuracion)
    vDuracion = nDuracion
End Property


Private Sub Command1_Click()
    If MsgBox("¿Está seguro de actualiza?", vbYesNo, "Confirmación") = vbYes Then
        frmPlaneamientoOrganizacionProcesos.lstTareasAplicacion.SelectedItem.ListSubItems(2) = CDbl(Me.txtDuracion)
        Unload Me
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

Private Sub Form_Load()
FormHelper.Customize Me
    Set rss = conectar.RSFactory("select concat(ptc.codigoTarea,' - ' , t.tarea) as tarea  from PlaneamientoTiemposProcesos as ptc inner join tareas t on ptc.codigoTarea=t.id where ptc.id=" & vidDetalleTiempos)

    If Not rss.EOF And Not rss.BOF Then
        Me.lblTarea = rss!Tarea
        Me.txtDuracion = vDuracion
    End If
    grabado = True
End Sub

Private Sub txtDuracion_Change()
    grabado = False
End Sub

Private Sub txtDuracion_GotFocus()
    foco Me.txtDuracion
End Sub
