VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlaneamientoOrganizacionProcesos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Organización de procesos"
   ClientHeight    =   8850
   ClientLeft      =   10980
   ClientTop       =   6450
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asignar"
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Finalizar"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Finalizar"
      Height          =   375
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
      Height          =   315
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Agregar nueva tarea"
      Top             =   8400
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-"
      Height          =   315
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Quitar tarea elegida"
      Top             =   8400
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   1095
   End
   Begin MSComctlLib.ListView lstDetallePedido 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   11553
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Procesos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Entrega"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstTareasAplicacion 
      Height          =   3735
      Left            =   6360
      TabIndex        =   1
      Top             =   4560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tarea"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Durac"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "PTO"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ListView lstDetalleConjunto 
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pieza"
         Object.Width           =   6615
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cant"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Definido"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle de ORDEN de TRABAJO"
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
      TabIndex        =   11
      Top             =   120
      Width           =   12135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Despiece Conjunto"
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
      Top             =   4320
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tareas definidas"
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
      Left            =   6360
      TabIndex        =   0
      Top             =   4320
      Width           =   5895
   End
   Begin VB.Menu ott 
      Caption         =   "ot"
      Visible         =   0   'False
      Begin VB.Menu noDefinir 
         Caption         =   "No Definir procesos..."
      End
   End
End
Attribute VB_Name = "frmPlaneamientoOrganizacionProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conj As Integer
Dim vIdPieza As Long
Dim vIdDetallePedido As Long
Dim vidDetallePedidoConjunto As Long
Dim plan As New classPlaneamiento
Dim vot As Long
Dim estadoElegido As Integer
Public Property Let ot(NuevoOT As Long)

    vot = NuevoOT
End Property

Private Sub Command1_Click()
    estado = CInt(Me.lstDetallePedido.SelectedItem.ListSubItems(3).Tag)

    If estado = 1 Then
        MsgBox "No puede actualizar si el proceso ya está definido!", vbInformation, "Información"
        Exit Sub
    End If

    If estado = 2 Then
        MsgBox "No puede actualizar en estado no definido!", vbInformation, "Información"
        Exit Sub
    End If

    If MsgBox("¿Está seguro de actualizar los datos?", vbYesNo, "Confimación") = vbYes Then
        update_procesos
    End If
End Sub

Private Sub update_procesos()

    If conj = 0 Then
        IdDetallePedido = vidDetallePedidoConjunto
    Else
        IdDetallePedido = vIdDetallePedido
    End If
    plan.ejecutarComando "delete from  PlaneamientoTiemposProcesos where idDetallePedido=" & IdDetallePedido
    For X = 1 To Me.lstTareasAplicacion.ListItems.count

        codigoTarea = CLng(Me.lstTareasAplicacion.ListItems(X))
        dura = CDbl(Me.lstTareasAplicacion.ListItems(X).ListSubItems(2))
        IdPedido = vot
        idPieza = vIdPieza
        If conj = 0 Then
            IdDetallePedido = vidDetallePedidoConjunto
        Else
            IdDetallePedido = vIdDetallePedido
        End If
        plan.actualizarTiemposProceso IdPedido, idPieza, IdDetallePedido, codigoTarea, dura


    Next X

End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de volver?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Command3_Click()
    For X = 1 To Me.lstTareasAplicacion.ListItems.count
        If Me.lstTareasAplicacion.ListItems(X).Selected Then

            If X - 1 >= 1 Then
                tmp_tarea = Me.lstTareasAplicacion.ListItems(X)
                tmp_duracion = Me.lstTareasAplicacion.ListItems(X).SubItems(1)
                tmp_prioridad = Me.lstTareasAplicacion.ListItems(X).SubItems(2)
                tmp_tag = Me.lstTareasAplicacion.ListItems(X).Tag
                Me.lstTareasAplicacion.ListItems(X).text = Me.lstTareasAplicacion.ListItems(X - 1).text
                Me.lstTareasAplicacion.ListItems(X).ListSubItems(1).text = Me.lstTareasAplicacion.ListItems(X - 1).SubItems(1)
                Me.lstTareasAplicacion.ListItems(X).ListSubItems(2).text = Me.lstTareasAplicacion.ListItems(X - 1).SubItems(2)
                Me.lstTareasAplicacion.ListItems(X).Tag = Me.lstTareasAplicacion.ListItems(X - 1).Tag
                Me.lstTareasAplicacion.ListItems(X - 1).text = tmp_tarea
                Me.lstTareasAplicacion.ListItems(X - 1).ListSubItems(1) = tmp_duracion
                Me.lstTareasAplicacion.ListItems(X - 1).ListSubItems(2) = tmp_prioridad
                Me.lstTareasAplicacion.ListItems(X - 1).Tag = tmp_tag
                Me.lstTareasAplicacion.ListItems(X - 1).Selected = True
                Me.lstTareasAplicacion.ListItems(X - 1).EnsureVisible

            End If

        End If
    Next X

End Sub

Private Sub Command4_Click()



    For X = Me.lstTareasAplicacion.ListItems.count To 1 Step -1
        If Me.lstTareasAplicacion.ListItems(X).Selected Then
            If X + 1 <= Me.lstTareasAplicacion.ListItems.count Then
                tmp_tarea = Me.lstTareasAplicacion.ListItems(X)
                tmp_duracion = Me.lstTareasAplicacion.ListItems(X).SubItems(1)
                tmp_prioridad = Me.lstTareasAplicacion.ListItems(X).SubItems(2)
                tmp_tag = Me.lstTareasAplicacion.ListItems(X).Tag
                Me.lstTareasAplicacion.ListItems(X).text = Me.lstTareasAplicacion.ListItems(X + 1).text
                Me.lstTareasAplicacion.ListItems(X).ListSubItems(1).text = Me.lstTareasAplicacion.ListItems(X + 1).SubItems(1)
                Me.lstTareasAplicacion.ListItems(X).ListSubItems(2).text = Me.lstTareasAplicacion.ListItems(X + 1).SubItems(2)
                Me.lstTareasAplicacion.ListItems(X).Tag = Me.lstTareasAplicacion.ListItems(X + 1).Tag
                Me.lstTareasAplicacion.ListItems(X + 1).text = tmp_tarea
                Me.lstTareasAplicacion.ListItems(X + 1).ListSubItems(1) = tmp_duracion
                Me.lstTareasAplicacion.ListItems(X + 1).ListSubItems(2) = tmp_prioridad
                Me.lstTareasAplicacion.ListItems(X + 1).Tag = tmp_tag
                Me.lstTareasAplicacion.ListItems(X + 1).Selected = True
            End If
        End If
    Next X


End Sub

Private Sub Command5_Click()
    Dim es As Boolean
    Dim hayCero As Boolean

    estado = CInt(Me.lstDetallePedido.SelectedItem.ListSubItems(3).Tag)

    If estado = 1 Then
        MsgBox "No puede finalizar si el proceso ya está definido!", vbInformation, "Información"
        Exit Sub
    End If

    If estado = 2 Then
        MsgBox "No puede finaliar en estado no definido!", vbInformation, "Información"
        Exit Sub
    End If

    If MsgBox("¿Desea finalizar la definición de procesos del elemento seleccionado?", vbYesNo, "Confirmación") = vbYes Then


        'no puede haber tareas definidas de duración 0

        hayCero = False
        For X = 1 To Me.lstTareasAplicacion.ListItems.count
            If CDbl(Me.lstTareasAplicacion.ListItems(X).ListSubItems(2)) = 0 Then hayCero = True
        Next X


        a = plan.duracionProcesos(vot)


        'la duracion de las tareas no puden superar el plazo de ejecución de las tareas
        'sumo todas las tareas


        a = Math.Round(a, 0)
        fechaEstimadafin = Format(Now + a, "dd-mm-yyyy")



        Dim rs As recordset
        Set rs = conectar.RSFactory("select fechaEntrega from pedidos where id=" & vot)
        If Not rs.EOF And Not rs.BOF Then
            FechaFin = rs!FechaEntrega
        Else
            Exit Sub
        End If

        If CDate(fechaEstimadafin) > FechaFin Then
            MsgBox "Es imposible cursar una OT con procesos que superen la fecha tope de entrega", vbCritical, "Error"
            Exit Sub
        End If


        If hayCero Then
            MsgBox "No puede haber tareas definidas de duración 0 (cero)!", vbCritical, "Error"
        Else

            update_procesos



            If conj = 0 Then
                es = plan.ejecutarComando("update detalles_pedidos_conjuntos set procesos_definidos =1 where id=" & vidDetallePedidoConjunto)

            Else
                es = plan.ejecutarComando("update detalles_pedidos set procesos_definidos =1 where id=" & vIdDetallePedido)
            End If

            If es Then
                MsgBox "Finalización exitosa!", vbInformation, "Información"
                If conj = 0 Then
                    Me.lstDetalleConjunto.SelectedItem.ListSubItems(2).text = "SI"
                    Me.lstDetalleConjunto.SelectedItem.ListSubItems(2).Tag = 1
                End If

            Else
                MsgBox "Se produjo algún error, se abortan los cambios!", vbCritical, "Error"
            End If
        End If
    End If
End Sub

Private Sub Command6_Click()
'para finalizar los procesos del conjunto, deben estar todos los procesos finalizados de sus partes

    If conj = 0 Then
        If plan.procesos_definidos_conjuntos(vIdDetallePedido) Then

            If MsgBox("¿Desea finalizar la definición de procesos para el conjunto seleccionado?", vbYesNo, "Confirmación") = vbYes Then



                If plan.ejecutarComando("update detalles_pedidos set procesos_definidos =1 where id=" & vIdDetallePedido) Then
                    MsgBox "Finalización exitosa!", vbInformation, "Información"

                    Me.lstDetallePedido.SelectedItem.ListSubItems(3).text = "Definido"
                    Me.lstDetallePedido.SelectedItem.ListSubItems(3).Tag = 1

                Else
                    MsgBox "Se produjo algún error, se abortan los cambios!", vbCritical, "Error"
                End If
            End If

        Else

            MsgBox "Debe tener todos los procesos definidos para el conjunto seleccionado!", vbCritical, "Error"
        End If
    End If
End Sub

Private Sub Command7_Click()
    For o = 1 To Me.lstTareasAplicacion.ListItems.count
        Me.lstTareasAplicacion.ListItems(o).ListSubItems(2).text = Me.lstTareasAplicacion.ListItems(o).ListSubItems(3).text

    Next o
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
   Me.Label4 = "Detalle de ORDEN de TRABAJO Nro. " & Format(vot, "0000")
    llenarDetallePedido
End Sub


Private Sub llenarDetallePedido()
    Dim X As ListItem
    Me.lstDetallePedido.ListItems.Clear
    Dim rs_pedido As recordset
    Set rs_pedido = conectar.RSFactory("select dp.idPieza,s.conjunto,dp.fechaEntrega,dp.item,dp.id,dp.idPieza,dp.cantidad,dp.procesos_definidos,s.detalle from detalles_pedidos dp inner join stock s on dp.idPieza=s.id  where idPedido=" & vot)

    While Not rs_pedido.EOF And Not rs_pedido.BOF
      Set X = Me.lstDetallePedido.ListItems.Add(, , "1111")
        X.SubItems(1) = rs_pedido!detalle
        X.SubItems(2) = rs_pedido!Cantidad
        X.ListSubItems(1).Tag = rs_pedido!conjunto

        X.ListSubItems(2).Tag = rs_pedido!idPieza


        aa = CInt(rs_pedido!procesos_definidos)
        If aa = 0 Then
            proceso = "Pendiente"
        ElseIf aa = 1 Then
            proceso = "Definido"
        ElseIf aa = 2 Then
            proceso = "No Definido"
        End If
        X.SubItems(3) = proceso
        X.ListSubItems(3).Tag = rs_pedido!procesos_definidos

        X.SubItems(4) = rs_pedido!FechaEntrega

        X.Tag = rs_pedido!Id

        rs_pedido.MoveNext
    Wend
    Set rs_pedido = Nothing
End Sub



Private Sub lstDetalleConjunto_ItemClick(ByVal Item As MSComctlLib.ListItem)
    vidDetallePedidoConjunto = CLng(Me.lstDetalleConjunto.SelectedItem.Tag)  'conjunto
    llenarLstProcesos vidDetallePedidoConjunto
End Sub

Private Sub lstDetallePedido_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    vIdDetallePedido = CLng(Me.lstDetallePedido.SelectedItem.Tag)
    conj = CInt(Me.lstDetallePedido.SelectedItem.ListSubItems(1).Tag)
    vIdPieza = CLng(Me.lstDetallePedido.SelectedItem.ListSubItems(2).Tag)


    Me.lstDetalleConjunto.ListItems.Clear
    Me.lstTareasAplicacion.ListItems.Clear
    pro = CInt(Me.lstDetallePedido.SelectedItem.ListSubItems(3).Tag)

    If pro = 1 Or pro = 2 Then
        Me.lstDetalleConjunto.Enabled = False
        Me.lstTareasAplicacion.Enabled = False
    Else
        Me.lstTareasAplicacion.Enabled = True
    End If

    If conj = -1 Then
        llenarLstProcesos vIdDetallePedido
    Else
        llenarLstDetalleConjunto
    End If
End Sub


Private Sub llenarLstProcesos(idDetaPedi)
    Dim X As ListItem
    Dim rs_t As recordset
    Me.lstTareasAplicacion.ListItems.Clear

    cant_ = CDbl(Me.lstDetallePedido.SelectedItem.ListSubItems(2))

    If conj = 0 Then
        cant_conj = CDbl(Me.lstDetalleConjunto.SelectedItem.ListSubItems(1))
    Else
        cant_conj = 1
    End If


    Set rs_t = conectar.RSFactory("select ptp.id,ptp.codigoTarea, t.tarea, ptp.duracion, t.cantxProc,dmdo.codigo,sum(dmdo.tiempo*dmdo.cantidad) as tiempo from  PlaneamientoTiemposProcesos as ptp inner join tareas t on ptp.codigoTarea=t.id inner join desarrollo_mdo dmdo on dmdo.id_Pieza=ptp.idPieza and dmdo.codigo=ptp.codigoTarea where idDetallePedido=" & idDetaPedi & " group by t.id")
    While Not rs_t.EOF And Not rs_t.BOF
        Set X = Me.lstTareasAplicacion.ListItems.Add(, , rs_t!codigoTarea)
        X.SubItems(1) = rs_t!Tarea
        X.Tag = rs_t!codigoTarea

        X.SubItems(2) = Format(Math.Round(rs_t!duracion, 4), "0.0000")
        X.Tag = rs_t!Id

        If rs_t!cantxproc > 0 Then
            taim = rs_t!Tiempo * cant_
        Else
            taim = rs_t!Tiempo
        End If
        X.SubItems(3) = Format(Math.Round((taim * cant_conj) / 1440, 4), "0.0000")    'saco el total de días!!
        rs_t.MoveNext
    Wend

    Set rs_t = Nothing


End Sub


Private Sub llenarLstDetalleConjunto()
    Dim rs As recordset
    Dim X As ListItem
    Me.lstDetalleConjunto.ListItems.Clear

    Set rs = conectar.RSFactory("select sc.procesos_definidos,sc.id,sc.idPieza,sc.cantidad,s.detalle,s.cantidad as stock from detalles_pedidos_conjuntos sc inner join stock s on sc.idPieza=s.id  where iddetalle_Pedido=" & vIdDetallePedido)
    While Not rs.EOF
        Set X = Me.lstDetalleConjunto.ListItems.Add(, , rs!detalle)
        X.SubItems(1) = rs!Cantidad

        If CInt(rs!procesos_definidos) = 0 Then
            defi = "NO"
        Else
            defi = "SI"
        End If

        X.SubItems(2) = defi
        X.ListSubItems(2).Tag = rs!procesos_definidos
        X.Tag = rs!Id
        rs.MoveNext
    Wend

End Sub

Private Sub lstDetallePedido_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        esta = CInt(Me.lstDetallePedido.SelectedItem.ListSubItems(3).Tag)
        If esta = 1 Or esta = 2 Then
            'si el estado es 1 no muestro NO DEFINIR PROCEOS
            Me.noDefinir.Enabled = False
        ElseIf esta = 0 Then
            Me.noDefinir.Enabled = True
        End If


        Me.PopupMenu ott



    End If
End Sub

Private Sub lstTareasAplicacion_DblClick()
    If Me.lstTareasAplicacion.ListItems.count > 0 Then
        idP = CLng(Me.lstTareasAplicacion.SelectedItem.Tag)
        dura = CDbl(Me.lstTareasAplicacion.SelectedItem.ListSubItems(2))

        frmPlaneamientoOrganizacionProcesosModificar.idDetalleTiempo = idP
        frmPlaneamientoOrganizacionProcesosModificar.duracion = dura

        frmPlaneamientoOrganizacionProcesosModificar.Show 1
    End If
End Sub

Private Sub noDefinir_Click()
    On Error GoTo err1
    Dim col As New Collection
    For X = 1 To 1
        If Me.lstDetallePedido.ListItems(X) Then col.Add CLng(Me.lstDetallePedido.SelectedItem.Tag)
    Next



    If col.count = 1 Then
        idDetaPedi = CLng(Me.lstDetallePedido.SelectedItem.Tag)
        If MsgBox("¿Está seguro de no definir procesos para este item?", vbYesNo, "Confirmación") = vbYes Then
            plan.ejecutarComando "update detalles_pedidos set procesos_definidos=2 where id=" & idDetaPedi
            Me.lstDetallePedido.SelectedItem.ListSubItems(3).Tag = 2
            Me.lstDetallePedido.SelectedItem.ListSubItems(3).text = "No Definir Proceso"
            Me.lstDetalleConjunto.Enabled = False
            Me.lstTareasAplicacion.Enabled = False
        End If
    ElseIf col.count > 1 Then
        If MsgBox("¿Está seguro de no definir procesos para estos items?", vbYesNo, "Confirmación") = vbYes Then
            For av = 1 To col.count
                plan.ejecutarComando "update detalles_pedidos set procesos_definidos=2 where id=" & col.Item(av)
                Debug.Print col.Item(av)
                Me.lstDetallePedido.ListItems(av).ListSubItems(3).Tag = 2
                Me.lstDetallePedido.ListItems(av).ListSubItems(3).text = "No Definir Proceso"
                Me.lstDetalleConjunto.Enabled = False
                Me.lstTareasAplicacion.Enabled = False
            Next av
        End If

    End If


    Exit Sub
err1:

End Sub
