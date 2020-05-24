VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlaneamientoSeguimientoRutas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seguimiento rutas..."
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Tareas aplicables ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   13515
      Begin VB.CommandButton cmdAgregarTarea 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar Tarea..."
         Height          =   330
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3915
         Width           =   2865
      End
      Begin VB.Frame framePieza 
         BackColor       =   &H00C0C0C0&
         Caption         =   "[ Detalle tarea ]"
         Height          =   3615
         Left            =   9000
         TabIndex        =   8
         Top             =   240
         Width           =   4470
         Begin VB.CommandButton btnNuevo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nuevo"
            Height          =   360
            Left            =   3585
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   3195
            Width           =   825
         End
         Begin VB.TextBox txtCantProc 
            Height          =   285
            Left            =   1575
            TabIndex        =   19
            Top             =   2595
            Width           =   855
         End
         Begin VB.TextBox txtHoras 
            Height          =   285
            Left            =   1575
            TabIndex        =   18
            Top             =   2205
            Width           =   855
         End
         Begin VB.CommandButton btnFinalizarTarea 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Finalizar Tarea"
            Height          =   375
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3180
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker dtpInicio 
            Height          =   300
            Left            =   1560
            TabIndex        =   17
            Top             =   1425
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy HH:mm"
            Format          =   65011715
            CurrentDate     =   39231
         End
         Begin VB.CommandButton cmdAgregar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Agregar"
            Height          =   360
            Left            =   2475
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2340
            Width           =   825
         End
         Begin VB.TextBox txtLegajo 
            Height          =   285
            Left            =   1560
            TabIndex        =   16
            Top             =   1080
            Width           =   1755
         End
         Begin MSComCtl2.DTPicker dtpFin 
            Height          =   300
            Left            =   1560
            TabIndex        =   27
            Top             =   1800
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy HH:mm"
            Format          =   65011715
            CurrentDate     =   39231
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad"
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
            Left            =   405
            TabIndex        =   24
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Horas"
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
            Left            =   405
            TabIndex        =   22
            Top             =   2250
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hora Fin"
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
            Left            =   375
            TabIndex        =   15
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hora Inicio"
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
            Left            =   390
            TabIndex        =   14
            Top             =   1470
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Legajo"
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
            Left            =   390
            TabIndex        =   13
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   4335
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label lblTarea 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   840
            TabIndex        =   12
            Top             =   615
            Width           =   45
         End
         Begin VB.Label lblPieza 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   840
            TabIndex        =   11
            Top             =   315
            Width           =   45
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tarea "
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
            Top             =   645
            Width           =   615
         End
         Begin VB.Label Pie 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pieza "
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
            TabIndex        =   9
            Top             =   300
            Width           =   615
         End
      End
      Begin MSComctlLib.ListView lstDetalleItem 
         Height          =   3495
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   6165
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
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tarea"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tarea"
            Object.Width           =   4233
         EndProperty
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   330
         Left            =   9015
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3930
         Width           =   4455
      End
      Begin MSComctlLib.ListView lstHorasPorTarea 
         Height          =   3495
         Left            =   6240
         TabIndex        =   23
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   6165
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Leg"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empleado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Horas"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cant"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Inicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fin"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstDetalleConjunto 
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   6165
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
            Text            =   "Pieza"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cant"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cant Total"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Seleccione O/T ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   13515
      Begin MSComctlLib.ListView lstDetalleOT 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   13350
         _ExtentX        =   23548
         _ExtentY        =   3836
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pieza"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nota"
            Object.Width           =   5627
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cant"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fabricado"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ver"
         Default         =   -1  'True
         Height          =   300
         Left            =   2340
         TabIndex        =   3
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtOTNumero 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
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
         Left            =   135
         TabIndex        =   1
         Top             =   330
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPlaneamientoSeguimientoRutas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim claseSP As New ClassPersonal
Dim claseP As New classPlaneamiento
Dim claseS As New classStock
Dim idTiempoSeleccionado As Long
Dim idDetalleSeleccionado As Long
Dim idDetalleConjuntoSeleccionado As Long
Dim esConjuntoSeleccionado As Long
Dim codTareaSeleccionado As Long

Dim last_detalle_pedido_id As Long
Dim pieza_principal_conjunto As Boolean
Dim nuevo_detalle_proceso As Boolean

Private Sub cmdAgregar_Click()

    If Me.lstDetalleItem.ListItems.count = 0 Then
        MsgBox "no hay tarea para agregar el detalle", vbOKOnly + vbCritical
        Exit Sub
    End If

    If Not IsNumeric(txtCantProc.text) Then
        MsgBox "Debe especificar una cantidad mayor a cero", vbOKOnly + vbCritical
        Exit Sub
    End If

    If (IsNumeric(Me.txtCantProc.text) And CLng(Me.txtCantProc.text) = 0) Then
        MsgBox "Debe especificar una cantidad mayor a cero", vbOKOnly + vbCritical
        Exit Sub
    End If


    On Error GoTo en33
    Dim horas As Double
    Dim legajo As Long
    Dim fec As Date
    Dim inicio As Date, fin As Date

    If validarHoras(Me.dtpInicio.value, Me.dtpFin.value, horas) Then
        horas = Me.txtHoras

        If MsgBox("¿Está seguro de actualizar los datos?", vbYesNo, "Confirmación") = vbYes Then
            legajo = CLng(Me.txtLegajo)

            Dim idtiempoprocesodetalle As Long
            If nuevo_detalle_proceso Then
                idtiempoprocesodetalle = 0
            Else
                If Me.lstHorasPorTarea.ListItems.count > 0 Then
                    idtiempoprocesodetalle = Me.lstHorasPorTarea.selectedItem.ListSubItems(6)
                End If
            End If

            claseP.cargarTiemposOT idTiempoSeleccionado, idtiempoprocesodetalle, legajo, horas, Me.dtpInicio.value, Me.dtpFin.value, CDbl(Me.txtCantProc)

            nuevo_detalle_proceso = False

            verElementoSeleccionado
            verTiemposProcesos
        End If
    Else
        MsgBox "Error en la carga de datos!", vbCritical, "Error"
        Exit Sub
    End If

    Exit Sub
en33:
    MsgBox "Ingrese datos válidos!" & vbNewLine & Err.Description
End Sub

Private Sub cmdAgregarTarea_Click()
    If Me.lstDetalleOT.ListItems.count = 0 Then
        MsgBox "No hay detalles para la OT", vbOKOnly + vbInformation
        Exit Sub
    End If

    Dim id_pieza As Long
    Dim id_detalle_pedido As Long
    If Int(Me.lstDetalleOT.selectedItem.ListSubItems(6).Tag) = -1 Then    'si no es conjunto tomo id_pieza de la lista del detalle_pedido
        id_pieza = CLng(Me.lstDetalleOT.selectedItem.Tag)
        id_detalle_pedido = Me.lstDetalleOT.selectedItem.Tag
    Else    'si es conjunto, tomo el id_pieza de la lista de piezas del conjunto
        If pieza_principal_conjunto Then
            id_pieza = CLng(Me.lstDetalleOT.selectedItem.Tag)
            id_detalle_pedido = Me.lstDetalleOT.selectedItem.Tag
        Else
            id_pieza = CLng(Me.lstDetalleConjunto.selectedItem.ListSubItems(2))
            id_detalle_pedido = Me.lstDetalleConjunto.selectedItem.Tag
        End If

    End If

    frmPlaneamientoAgregarTareaAProceso.PIEZA_ID = id_pieza
    frmPlaneamientoAgregarTareaAProceso.idDetallePedido = id_detalle_pedido
    frmPlaneamientoAgregarTareaAProceso.Show 1

    If Int(Me.lstDetalleOT.selectedItem.ListSubItems(6).Tag) = -1 Or (Int(Me.lstDetalleOT.selectedItem.ListSubItems(6).Tag) <> -1 And pieza_principal_conjunto) Then
        lstDetalleOT_ItemClick Me.lstDetalleOT.selectedItem
    Else
        lstDetalleConjunto_ItemClick Me.lstDetalleConjunto.selectedItem
    End If


End Sub

Private Sub Command1_Click()
    On Error GoTo err411
    Dim idpedido As Long
    If IsNumeric(Trim(Me.txtOTNumero)) Then
        idpedido = CLng(Trim(Me.txtOTNumero))

        If claseP.ExistePedido(idpedido) = 1 Then
            If claseP.estadoPedido(idpedido) = 2 Or claseP.estadoPedido(idpedido) = 3 Then
                'pedido en proceso, se podrá visualizar los datos
                'para hacer el seguimiento.
                llenarLST idpedido
                If Me.lstDetalleOT.ListItems.count > 0 Then lstDetalleOT_ItemClick Me.lstDetalleOT.ListItems(1)


                'idDetalleSeleccionado = CLng(Me.lstDetalleOT.SelectedItem)
                'llenarLstDetalle (idDetalleSeleccionado)

                '  verElementoSeleccionado
                ' verTiemposProcesos
            Else
                MsgBox "El pedido no está en proceso.", vbCritical, "Error"
                Exit Sub
            End If
        Else
            MsgBox "El pedido no existe!", vbCritical, "Error"
        End If
    End If

    Exit Sub
err411:

    MsgBox Err.Description
'    Resume
Resume Next
End Sub


Private Sub llenarLST(idpedido As Long)
    Dim rs As recordset
    Dim XU As ListItem
    Me.lstDetalleOT.ListItems.Clear
    Dim c As String
    Set rs = conectar.RSFactory("select s.conjunto,dp.id,dp.item,dp.nota,dp.idPieza,s.detalle,dp.cantidad,dp.cantidad_fabricados,dp.fechaEntrega from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where dp.idPedido=" & idpedido)
    While Not rs.EOF
        Set XU = Me.lstDetalleOT.ListItems.Add(, , rs!item)
        XU.SubItems(1) = rs!detalle
        XU.SubItems(2) = rs!Nota
        XU.SubItems(3) = rs!cantidad
        XU.SubItems(4) = rs!cantidad_fabricados
        XU.SubItems(5) = rs!FechaEntrega
        c = "Unidad"


        XU.SubItems(6) = funciones.EsConjunto(rs!conjunto)
        XU.ListSubItems(6).Tag = rs!conjunto
        XU.Tag = rs!id
        rs.MoveNext
    Wend



End Sub

Private Sub Command2_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then Unload Me
End Sub
Private Sub llenarLstDetalle(idDetallePedido As Long)
    Dim rs As recordset
    Dim x As ListItem
    Me.lstDetalleItem.ListItems.Clear
    Me.lstDetalleConjunto.ListItems.Clear
    Me.lstDetalleConjunto.Enabled = (esConjuntoSeleccionado = 1)

    If esConjuntoSeleccionado = 1 Then
        'si es conjunto el seleccionado tengo que ver el detalle del conjunto
        llenarLstDetalleConjunto idDetallePedido
        If Me.lstDetalleConjunto.ListItems.count > 0 Then
            lstDetalleConjunto_ItemClick Me.lstDetalleConjunto.ListItems(1)
        End If
    Else
        'llena lista si no es conjunto
        llenarListaMDO idDetallePedido
        If Me.lstDetalleItem.ListItems.count > 0 Then lstDetalleItem_ItemClick Me.lstDetalleItem.ListItems(1)
    End If

End Sub
Private Sub btnFinalizarTarea_Click()
    If Me.lstDetalleItem.ListItems.count = 0 Then
        MsgBox "no hay tarea para finalizar", vbOKOnly + vbCritical
        Exit Sub
    End If

    'VERIFICAR QUE SE HAYAN REALIZADO TODAS LAS PIEZAS PARA PODER FINALIZAR LA TAREA
    'tengo que ver si la saco directamente del detalle o del conjunto del detalle

    Dim cantidad_pedida As Long


    If Int(Me.lstDetalleOT.selectedItem.ListSubItems(6).Tag) = -1 Then    'si no es conjunto tomo la cantidad del detalle pedido
        cantidad_pedida = Me.lstDetalleOT.selectedItem.ListSubItems(3)
    Else    'si es conjunto, tomo el id_pieza de la lista de piezas del conjunto
        If pieza_principal_conjunto Then
            cantidad_pedida = Me.lstDetalleOT.selectedItem.ListSubItems(3)
        Else
            cantidad_pedida = Me.lstDetalleConjunto.selectedItem.ListSubItems(3)
        End If
    End If

    Dim q As String
    Dim totalProcesado As Long
    totalProcesado = 0
    q = "SELECT SUM(cantidad_procesada) as total FROM PlaneamientoTiemposProcesosDetalle WHERE idTiemposProcesos = " & Me.lstDetalleItem.selectedItem.text
    Dim rs As New recordset
    Set rs = conectar.RSFactory(q)
    If Not rs.EOF Then
        If Not IsNull(rs!Total) Then totalProcesado = rs!Total
    End If

    If totalProcesado < cantidad_pedida Then
        MsgBox "No se puede finalizar la tarea [" & Me.lblTarea.caption & "] para la pieza [" & Me.lblPieza.caption & "]" & vbNewLine & "Se requieren procesar por lo menos " & cantidad_pedida & " y se procesaron " & totalProcesado & ".", vbOKOnly + vbInformation
        Exit Sub
    End If

    If MsgBox("¿Está seguro de finalizar la tarea seleccionada?", vbYesNo, "Confirmación") = vbYes Then
        If claseP.finalizarTarea(idTiempoSeleccionado) Then
            Me.framePieza.Enabled = False
            verElementoSeleccionado
            verTiemposProcesos
        End If
    End If
End Sub

Private Sub btnNuevo_Click()
'Me.txtLegajo.Text = vbNullString


    Me.dtpInicio.value = Now
    Me.dtpFin.value = Now
    Me.txtHoras.text = vbNullString
    Me.txtCantProc.text = vbNullString
    If Me.txtLegajo.Visible And Me.txtLegajo.Enabled Then Me.txtLegajo.SetFocus

    nuevo_detalle_proceso = True

End Sub



Private Sub dtpFin_Change()
    ActualizarHoras
End Sub

Private Sub dtpInicio_Change()
    ActualizarHoras
End Sub

Private Sub ActualizarHoras()
    Me.txtHoras.text = Round(DateDiff("n", Me.dtpInicio.value, Me.dtpFin.value) / 60, 2)
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    Me.dtpInicio.value = Now
    Me.dtpFin.value = Now
End Sub

Private Sub lstDetalleConjunto_ItemClick(ByVal item As MSComctlLib.ListItem)
    If Me.lstDetalleConjunto.ListItems.count > 0 Then
        pieza_principal_conjunto = False
        idDetalleConjuntoSeleccionado = Me.lstDetalleConjunto.selectedItem.Tag
        llenarListaMDO idDetalleConjuntoSeleccionado, True
        If Me.lstDetalleItem.ListItems.count > 0 Then
            codTareaSeleccionado = CLng(Me.lstDetalleItem.selectedItem.SubItems(1))
        End If
        verElementoSeleccionado
    End If
End Sub

Private Sub lstDetalleItem_ItemClick(ByVal item As MSComctlLib.ListItem)
    codTareaSeleccionado = CLng(Me.lstDetalleItem.selectedItem.SubItems(1))
    verElementoSeleccionado
End Sub
Private Sub lstDetalleOT_ItemClick(ByVal item As MSComctlLib.ListItem)
    Dim rs As recordset
    Dim idDetalle As Long
    If Me.lstDetalleOT.ListItems.count > 0 Then

        esConjuntoSeleccionado = CInt(Me.lstDetalleOT.selectedItem.ListSubItems(6).Tag) + 1

        If last_detalle_pedido_id = Me.lstDetalleOT.selectedItem.Tag And esConjuntoSeleccionado = 1 Then
            'es el mismo que volvio a clickear, si es conjunto, cargo la pieza
            idDetalle = Me.lstDetalleOT.selectedItem.Tag
            pieza_principal_conjunto = True
            llenarListaMDO idDetalle
            If Me.lstDetalleItem.ListItems.count > 0 Then
                lstDetalleItem_ItemClick Me.lstDetalleItem.ListItems(1)
            Else
                Me.lstHorasPorTarea.ListItems.Clear
                btnNuevo_Click
            End If
        Else
            pieza_principal_conjunto = False

            idDetalle = Me.lstDetalleOT.selectedItem.Tag    'CLng(Me.lstDetalleOT.SelectedItem)
            Me.lstDetalleConjunto.ListItems.Clear
            Me.lstDetalleItem.ListItems.Clear
            Me.lstHorasPorTarea.ListItems.Clear

            llenarLstDetalle idDetalle

            idDetalleSeleccionado = idDetalle

            If Me.lstDetalleItem.ListItems.count > 0 Then
                codTareaSeleccionado = CLng(Me.lstDetalleItem.selectedItem.SubItems(1))
            End If

            verElementoSeleccionado

        End If


        last_detalle_pedido_id = Me.lstDetalleOT.selectedItem.Tag
    End If
End Sub
Private Sub verElementoSeleccionado()
    Dim rs As recordset

    'If esConjuntoSeleccionado = 1 Then
    '    'si no es conjunto el seleccionado entonces cargo directamente las tareas
    '    If pieza_principal_conjunto Then
    '        Set rs = conectar.RSFactory("select s.detalle,t.tarea from detalles_pedidos dp inner join stock s on dp.idPieza=s.id inner join desarrollo_mdo dmdo on dmdo.id_pieza=s.id inner join tareas t on dmdo.codigo=t.id where dp.id=" & idDetalleSeleccionado & " and t.id=" & codTareaSeleccionado)
    '    Else
    '        Set rs = conectar.RSFactory("select s.detalle,t.tarea from detalles_pedidos_conjuntos dp inner join stock s on dp.idPieza=s.id inner join desarrollo_mdo dmdo on dmdo.id_pieza=s.id inner join tareas t on dmdo.codigo=t.id where dp.id=" & idDetalleConjuntoSeleccionado & " and t.id=" & codTareaSeleccionado)
    '    End If
    ''idDetalleSeleccionado = idDetalleConjuntoSeleccionado
    'Else
    '    Set rs = conectar.RSFactory("select s.detalle,t.tarea from detalles_pedidos dp inner join stock s on dp.idPieza=s.id inner join desarrollo_mdo dmdo on dmdo.id_pieza=s.id inner join tareas t on dmdo.codigo=t.id where dp.id=" & idDetalleSeleccionado & " and t.id=" & codTareaSeleccionado)
    'End If

    Dim q As String
    q = "SELECT s.detalle, t.tarea" _
        & " FROM PlaneamientoTiemposProcesos ptp" _
        & " INNER JOIN stock s ON s.id = ptp.idpieza" _
        & " INNER JOIN tareas t ON t.id = ptp.codigotarea" _
        & " Where ptp.id = " & Me.lstDetalleItem.selectedItem

    Set rs = conectar.RSFactory(q)

    If Not rs.EOF And Not rs.BOF Then
        Me.lblPieza = rs!detalle
        Me.lblTarea = codTareaSeleccionado & " - " & rs!Tarea
    End If

    If Me.lstDetalleItem.ListItems.count = 0 Then
        'MsgBox "No hay tareas asignadas!", vbCritical, "Error"
        '        Me.lblPieza = Empty
        Me.lblTarea = Empty
        idTiempoSeleccionado = 0  'no marco nada
        Exit Sub
    End If

    idTiempoSeleccionado = CLng(Me.lstDetalleItem.selectedItem)
    Set rs = conectar.RSFactory("select estado from PlaneamientoTiemposProcesos where id=" & idTiempoSeleccionado)
    Dim estado As Long
    estado = 0

    If Not rs.EOF And Not rs.BOF Then estado = rs!estado

    'SI LA TAREA ESTÁ COMPLETA ANULO EL FRAME PARA CARGAR TIEMPOS
    Me.framePieza.Enabled = (estado <> 1)

    verTiemposProcesos

    If Me.Frame2.Enabled Then btnNuevo_Click    'limpio
    If Me.lstHorasPorTarea.ListItems.count > 0 Then lstHorasPorTarea_ItemClick Me.lstHorasPorTarea.ListItems(1)

End Sub
Private Sub verTiemposProcesos()
    Dim rs As recordset
    Me.lstHorasPorTarea.ListItems.Clear
    Dim x As ListItem
    Set rs = conectar.RSFactory("select * from PlaneamientoTiemposProcesosDetalle where idTiemposProcesos=" & idTiempoSeleccionado)
    While Not rs.EOF
        Set x = Me.lstHorasPorTarea.ListItems.Add(, , rs!legajo)
        x.SubItems(1) = DAOEmpleados.GetByLegajo(rs!legajo).NombreCompleto
        x.SubItems(2) = rs!Tiempo
        x.SubItems(3) = rs!cantidad_procesada
        x.SubItems(4) = IIf(IsNull(rs!inico), vbNullString, rs!inico)
        x.SubItems(5) = IIf(IsNull(rs!fin), vbNullString, rs!fin)
        x.SubItems(6) = rs!id
        rs.MoveNext
    Wend
End Sub
Public Function cuantasHoras(entrada As Date, salida As Date) As Double
    Dim segundos, minutos, horas As Double
    segundos = DateDiff("s", entrada, salida)
    minutos = (segundos / 60)
    horas = Math.Round(minutos / 60, 2)
    cuantasHoras = horas
End Function
Private Function validarHoras(Optional inicio, Optional fin, Optional ByRef horas) As Boolean
    If IsDate(inicio) And IsDate(fin) Then

        If fin < inicio Then
            validarHoras = False
        Else
            horas = funciones.cuantasHoras(CDate(Me.dtpInicio.value), CDate(Me.dtpFin.value))
            Me.txtHoras = horas
            validarHoras = True
        End If

    Else
        validarHoras = False
    End If

End Function



Private Sub lstHorasPorTarea_ItemClick(ByVal item As MSComctlLib.ListItem)
    Me.txtLegajo.text = item.text
    Me.txtHoras.text = item.ListSubItems(2)
    Me.txtCantProc.text = item.ListSubItems(3)
    Me.dtpInicio.value = IIf(item.ListSubItems(4) = vbNullString, Now, item.ListSubItems(4))
    Me.dtpFin.value = IIf(item.ListSubItems(5) = vbNullString, Now, item.ListSubItems(5))

    nuevo_detalle_proceso = False
End Sub

Private Sub txtFin_Change()
    On Error GoTo en33
    validarHoras CDate(Me.dtpInicio), CDate(Me.dtpFin)

    Exit Sub
en33:
End Sub

Private Sub txtInicio_Change()
    On Error GoTo en33
    validarHoras CDate(Me.dtpInicio), CDate(Me.dtpFin)

    Exit Sub
en33:

End Sub



Private Sub llenarLstDetalleConjunto(idDetallePedido As Long)
    Dim rs As recordset
    Dim x As ListItem


    Set rs = conectar.RSFactory("select sc.id,sc.idPieza,sc.cantidad,s.detalle,s.cantidad as stock from detalles_pedidos_conjuntos sc inner join stock s on sc.idPieza=s.id  where iddetalle_Pedido=" & idDetallePedido)
    While Not rs.EOF
        Set x = Me.lstDetalleConjunto.ListItems.Add(, , rs!detalle)
        x.SubItems(1) = rs!cantidad
        x.SubItems(2) = rs!idPieza
        x.SubItems(3) = rs!cantidad * Me.lstDetalleOT.selectedItem.ListSubItems(3)
        x.Tag = rs!id
        rs.MoveNext
    Wend

End Sub

Private Sub llenarListaMDO(idDetallePedido, Optional esConj As Boolean = False)
    Me.lstDetalleItem.ListItems.Clear
    Dim rs As recordset
    'If Not esConj Then
    'si no es conjunto
    Set rs = conectar.RSFactory("select p.id, p.codigoTarea, t.tarea, p.agregado, p.observacion_agregado from PlaneamientoTiemposProcesos p inner join tareas t on t.id = p.codigoTarea where p.idDetallePedido=" & idDetallePedido & IIf(esConj, " AND p.conjunto = 1", vbNullString))
    'Else
    '    Set rs = conectar.RSFactory("select p.id,p.codigoTarea, t.tarea from PlaneamientoTiemposProcesos p inner join tareas t on t.id=p.codigoTarea where p.idDetallePedido=" & IdDetallePedido)
    'End If


    Dim x As ListItem
    While Not rs.EOF
        Set x = Me.lstDetalleItem.ListItems.Add(, , rs!id)
        x.SubItems(1) = rs!codigoTarea
        x.SubItems(2) = rs!Tarea

        If (rs!agregado = 1) Then
            x.ToolTipText = rs!observacion_agregado
            x.ForeColor = vbRed
            x.ListSubItems(1).ForeColor = vbRed
            x.ListSubItems(1).ToolTipText = rs!observacion_agregado
            x.ListSubItems(2).ForeColor = vbRed
            x.ListSubItems(2).ToolTipText = rs!observacion_agregado
        End If
        rs.MoveNext
    Wend
End Sub
