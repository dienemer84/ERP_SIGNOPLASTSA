VERSION 5.00
Begin VB.Form frmEntregaTotal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega total"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   ClipControls    =   0   'False
   Icon            =   "frmEntregaTotal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   510
      Width           =   1095
   End
   Begin VB.TextBox txtNroRto 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Doble click para seleccionar remito"
      Top             =   90
      Width           =   2985
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A Stock..."
      Height          =   375
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Remito Nro"
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
      Left            =   225
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmEntregaTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseP As New classPlaneamiento
Public Pedido As Long
Private remitoId As Long
Private Sub Command1_Click()
    error2 = False


    If claseP.RealizarEntrega(2, remitoId, , , Pedido, 1) Then
        Dim Ot As OrdenTrabajo
        Set Ot = DAOOrdenTrabajo.FindById(Pedido)
        Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, True, True, True)

        If DAOOrdenTrabajo.Cerrar(Ot) Then
            Dim EVENTO As New clsEventoObserver
            Set EVENTO.Elemento = Ot
            Set EVENTO.Originador = Me
            EVENTO.EVENTO = modificar_
            Channel.Notificar EVENTO, ordenesTrabajo

            MsgBox "El pedido " & Pedido & " se cerro correctamente.", vbInformation, "Información"
            Unload Me
        End If
    End If


End Sub

Private Sub Command2_Click()
    error2 = False
    rto = -1

    Dim Ot As OrdenTrabajo
    Set Ot = DAOOrdenTrabajo.FindById(Pedido)
    Set Ot.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Ot.id, True, True, True)

    If DAOOrdenTrabajo.Cerrar(Ot, True) Then
        Dim EVENTO As New clsEventoObserver
        Set EVENTO.Elemento = Ot
        Set EVENTO.Originador = Me
        EVENTO.EVENTO = modificar_
        Channel.Notificar EVENTO, ordenesTrabajo


        MsgBox "El pedido " & pedidod & " se cerro correctamente.", vbInformation, "Información"
        Unload Me
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    ver
End Sub

Private Sub txtNroRto_Change()
    ver
End Sub
Public Sub ver()
    If Trim(txtNroRto) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Private Sub txtNroRto_DblClick()
    frmPlaneamientoRemitosListaProceso.mostrar = 0
    frmPlaneamientoRemitosListaProceso.Show 1
    If IsSomething(Selecciones.RemitoElegido) Then
        Me.txtNroRto = Selecciones.RemitoElegido.numero
        remitoId = Selecciones.RemitoElegido.id
    Else
        Me.txtNroRto = Empty
    End If
End Sub
