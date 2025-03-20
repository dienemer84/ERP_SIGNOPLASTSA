VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmPlaneamientoRemitoVer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   Icon            =   "frmPlaneamientoRemitoVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   8790
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   8535
      _Version        =   786432
      _ExtentX        =   15055
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtLugarEntrega 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   8175
      End
      Begin VB.Label lblEntrega 
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar de Entrega:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8535
      _Version        =   786432
      _ExtentX        =   15055
      _ExtentY        =   3413
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtObservaciones 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   6495
      End
      Begin VB.TextBox lblDetalle 
         Height          =   300
         Left            =   1560
         TabIndex        =   6
         Top             =   975
         Width           =   6525
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   5220
         _Version        =   786432
         _ExtentX        =   9208
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   -1  'True
         Text            =   "ComboBox1"
         AutoComplete    =   -1  'True
      End
      Begin VB.Label lblObservaciones 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   870
      End
      Begin VB.Label dsfsadf 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   240
         Left            =   600
         TabIndex        =   9
         Top             =   225
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
         Height          =   255
         Left            =   630
         TabIndex        =   8
         Top             =   998
         Width           =   825
      End
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   8281
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "observaciones"
      PreviewRowLines =   2
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowDelete     =   -1  'True
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   7
      Column(1)       =   "frmPlaneamientoRemitoVer.frx":000C
      Column(2)       =   "frmPlaneamientoRemitoVer.frx":0118
      Column(3)       =   "frmPlaneamientoRemitoVer.frx":020C
      Column(4)       =   "frmPlaneamientoRemitoVer.frx":033C
      Column(5)       =   "frmPlaneamientoRemitoVer.frx":0490
      Column(6)       =   "frmPlaneamientoRemitoVer.frx":05B0
      Column(7)       =   "frmPlaneamientoRemitoVer.frx":06CC
      FormatStylesCount=   10
      FormatStyle(1)  =   "frmPlaneamientoRemitoVer.frx":07F4
      FormatStyle(2)  =   "frmPlaneamientoRemitoVer.frx":092C
      FormatStyle(3)  =   "frmPlaneamientoRemitoVer.frx":09DC
      FormatStyle(4)  =   "frmPlaneamientoRemitoVer.frx":0A90
      FormatStyle(5)  =   "frmPlaneamientoRemitoVer.frx":0B68
      FormatStyle(6)  =   "frmPlaneamientoRemitoVer.frx":0C20
      FormatStyle(7)  =   "frmPlaneamientoRemitoVer.frx":0D00
      FormatStyle(8)  =   "frmPlaneamientoRemitoVer.frx":0DDC
      FormatStyle(9)  =   "frmPlaneamientoRemitoVer.frx":0E94
      FormatStyle(10) =   "frmPlaneamientoRemitoVer.frx":0F20
      ImageCount      =   0
      PrinterProperties=   "frmPlaneamientoRemitoVer.frx":0FD4
   End
   Begin XtremeSuiteControls.GroupBox GroupBox 
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   8535
      _Version        =   786432
      _ExtentX        =   15055
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdValorizar 
         Height          =   495
         Left            =   6000
         TabIndex        =   3
         Top             =   240
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnFacturar 
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   2295
         _Version        =   786432
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Facturar"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.Label lblIndicaciones 
      BackStyle       =   0  'Transparent
      Caption         =   "Para cambiar el estado a Item NO FACTURABLE. Abra el menu con click derecho sobre el item seleccionado."
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.Menu mnuDetalleRemito 
      Caption         =   "mnuDetalleRemito"
      Visible         =   0   'False
      Begin VB.Menu mnuNoFacturable 
         Caption         =   "Facturable"
      End
   End
End
Attribute VB_Name = "frmPlaneamientoRemitoVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber
'Dim deta_pedi As DetalleOrdenTrabajo
Dim noadd As Boolean
Dim detaEntrega As remitoDetalle
Dim scannedbuffer As String
Dim cli_viejo As clsCliente
Public MostrarInfoAdministracion As Boolean
Public valorizable As Boolean
Public editar As Boolean
Public conceptuable As Boolean
Public Usable As Boolean
Dim it As Long
Dim item

Public Remito As Remito
Dim tmp As remitoDetalle

Dim vId As String
Public ParaFacturar As Boolean
Public IdFormSuscriber As String

Private Sub btnFacturar_Click()
    Dim returnCol As New Collection
    Dim js As JSSelectedItem

    If Me.grilla.MultiSelect Then
        For Each js In Me.grilla.SelectedItems
            Set tmp = Remito.detalles.item(js.rowIndex)
            If tmp.Facturado Then
                MsgBox "No puede facturar un item que ya fue facturado.", vbExclamation
                Exit Sub
            End If
            returnCol.Add tmp, CStr(tmp.Id)
        Next js
    Else
        Set tmp = Remito.detalles.item(Me.grilla.rowIndex(Me.grilla.row))


        'cuando elremito esta facturado (de concepto) no deja aplicarlo a la OT
        'ver como distingo si esta ventana es desde apliacr remito o desde aplicar factura
        If tmp.Facturado And ParaFacturar Then
            MsgBox "No puede facturar un item que ya fue facturado.", vbExclamation
            Exit Sub
        End If
        returnCol.Add tmp, CStr(tmp.Id)
    End If

    Dim ev As New clsEventoObserver
    If ParaFacturar Then
        ev.Tipo = FacturarRemitosDetalle_
    Else
        ev.Tipo = FacturarRemitosDetalle_

    End If

    Set ev.Originador = Me
    Set ev.Elemento = returnCol

    ''If ParaFacturar Then
    Channel.Notificar ev, FacturarRemitosDetalle_
    ''  Else
    ' Channel.Notificar ev, RemitosDetalle_

    ' End If
    Unload Me
End Sub


Private Sub cmdValorizar_Click()
    On Error GoTo err1
    If MsgBox("¿Seguro de guardar los cambios?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then

        Set cli_viejo = Remito.cliente

        Remito.detalle = UCase(Me.lblDetalle)
        Remito.observaciones = UCase(Me.txtObservaciones)
        Remito.lugarEntrega = UCase(Me.txtLugarEntrega)
        
        Set Remito.cliente = DAOCliente.BuscarPorID(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
        
        If Not DAORemitoS.Save(Remito, True, False) Then
            MsgBox "Se produjo algun error al guardar!", vbCritical, "Error"
            Set Remito.cliente = cli_viejo
        Else

            MsgBox "Guardado correctamente!", vbInformation, "Información"

            Me.grilla.ReBind
            '    evento.Tipo = Remitos_
            '    Channel.Notificar evento, Remitos_
        End If
    End If
    Exit Sub
err1:
    Set Remito.cliente = cli_viejo

End Sub


Private Function CrearDetalleDeOT() As Boolean
    Set detaEntrega = New remitoDetalle
    Dim detapedido As DetalleOrdenTrabajo
    Dim idDetallePedido As Long
    idDetallePedido = Val(Right(scannedbuffer, Len(scannedbuffer) - 1))
    Set detapedido = DAODetalleOrdenTrabajo.FindById(idDetallePedido)

    If IsSomething(detapedido) Then
        detaEntrega.Cantidad = detapedido.CantidadPedida
        detaEntrega.facturable = True
        detaEntrega.Facturado = False
        detaEntrega.FEcha = Now
        detaEntrega.idDetallePedido = detapedido.Id
        detaEntrega.idpedido = detapedido.OrdenTrabajo.Id
        detaEntrega.Origen = OrigenRemitoOt
        detaEntrega.Remito = Me.Remito.Id
        detaEntrega.Valor = detapedido.Precio
        detaEntrega.ValorModificado = False
        Set detaEntrega.DetallePedido = detapedido


        If Not funciones.BuscarEnColeccion(Remito.detalles, CStr(detaEntrega.idDetallePedido)) Then
            Me.Remito.detalles.Add detaEntrega, CStr(detaEntrega.idDetallePedido)
            CrearDetalleDeOT = True
        Else
            CrearDetalleDeOT = False
        End If

    End If


End Function


Private Sub Form_Load()

    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    
    FormHelper.Customize Me

    Me.btnFacturar.Visible = Me.ParaFacturar Or Me.Usable

    GridEXHelper.CustomizeGrid grilla, False, True
    
    Me.grilla.AllowDelete = (editar And Remito.estado = RemitoPendiente)

    Me.cboClientes.Locked = Not editar  'Or Not valorizable And Not Usable
    Me.lblDetalle.Locked = Not editar
    Me.txtLugarEntrega.Locked = Not editar
    Me.txtObservaciones.Locked = Not editar

    If Me.Usable Then Me.btnFacturar.caption = "Usar Item"

    vId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, RemitosDetalle_
    
    mostrarRemito

End Sub


Private Sub mostrarRemito()

    If IsSomething(Remito) Then
        Me.caption = "Remito " & Remito.numero
        Set Remito.detalles = DAORemitoSDetalle.FindAllByRemito(Remito.Id, False, True)

        Me.lblFecha.caption = Remito.FEcha
        Me.lblDetalle = Remito.detalle
        Me.cboClientes.ListIndex = funciones.PosIndexCbo(Remito.cliente.Id, Me.cboClientes)
        Me.txtObservaciones = Remito.observaciones
        Me.txtLugarEntrega = Remito.lugarEntrega

        grilla.Columns(6).Visible = MostrarInfoAdministracion

        Me.cmdValorizar.Enabled = conceptuable Or valorizable Or editar And Not Usable
        grilla.AllowEdit = editar Or valorizable And Not Usable
        grilla.AllowAddNew = editar And Not Usable

        llenarLista

    End If

End Sub


Private Sub llenarLista()
    Me.grilla.ItemCount = 0
    Me.grilla.ItemCount = Remito.detalles.count
    
End Sub


Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub


Private Sub grilla_AfterUpdate()
    If Not noadd Then
        grilla.ItemCount = Remito.detalles.count
    End If
End Sub


Private Sub grilla_BeforeDelete(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = MsgBox("¿Está seguro de eliminar el detalle?", vbYesNo + vbInformation, "Confirmación") = vbNo     'Or tmp.Origen <> OrigenRemitoConcepto Or remito.estado <> RemitoPendiente

End Sub

Private Sub grilla_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (grilla.value(5) < 0 Or Not IsNumeric(grilla.value(5)))

End Sub


Private Sub grilla_DblClick()
    On Error Resume Next
    grilla_SelectionChange
    Dim pos As Long
    If Usable Then
        Set Selecciones.RemitoElegido = Remito
        Unload Me
    End If

    If editar Then
        pos = grilla.rowIndex(grilla.row)
        If Remito.CantidadDeLineasActuales > funciones.itemsPorRemito Then
            MsgBox "La cantidad de líneas superan a lo permitido"

            tmp.observaciones = vbNullString
        Else
            tmp.observaciones = UCase(InputBox("Observación", "Observacion", tmp.observaciones))
        End If

    End If

    grilla.RefreshRowIndex pos


End Sub


Private Sub grilla_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And IsSomething(tmp) Then
        Me.mnuNoFacturable.Enabled = Not tmp.Facturado And Remito.estado = EstadoRemito.RemitoAprobado And (Remito.EstadoFacturado = RemitoNoFacturado Or Remito.EstadoFacturado = RemitoFacturadoParcial)

        If tmp.facturable Then
            Me.mnuNoFacturable.caption = "Hacer No Facturable"
        Else
            Me.mnuNoFacturable.caption = "Hacer Facturable"
        End If
        Me.PopupMenu Me.mnuDetalleRemito
    End If
End Sub


Private Sub grilla_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error Resume Next
    'xxxx
    If RowBuffer.rowIndex > 0 And Remito.detalles.count > 0 Then
        Set tmp = Remito.detalles(RowBuffer.rowIndex)
        If tmp.facturable Then
            If Not tmp.Facturado Then
                RowBuffer.CellStyle(6) = "NoFacturado"
            Else
                RowBuffer.CellStyle(6) = "Facturado"
            End If
        Else
            RowBuffer.CellStyle(6) = "NoFacturable"
        End If
    End If
End Sub


Private Sub grilla_SelectionChange()
    Dim it As Long
    it = grilla.rowIndex(grilla.row)
    If it > 0 And Remito.detalles.count > 0 Then
        Set tmp = Remito.detalles.item(it)

        If tmp.Origen = OrigenRemitoConcepto Then
            grilla.Columns(2).EditType = jgexEditTextBox
            grilla.Columns(4).EditType = jgexEditTextBox
        Else
            grilla.Columns(2).EditType = jgexEditNone
            grilla.Columns(4).EditType = jgexEditTextBox
        End If

        If (Not tmp.facturable Or tmp.Facturado) Or (tmp.Origen <> OrigenRemitoConcepto And Not valorizable) Then
            grilla.Columns(5).EditType = jgexEditNone
        Else
            grilla.Columns(5).EditType = jgexEditTextBox
        End If
    Else
        grilla.Columns(2).EditType = jgexEditTextBox
        grilla.Columns(4).EditType = jgexEditTextBox
        grilla.Columns(5).EditType = jgexEditTextBox
    End If
End Sub


Private Sub grilla_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)

    Dim cod As String

    cod = StrConv(Left(Values(2), 1), vbUpperCase)

    If cod = "R" And Len(Values(2)) = 9 Then
        scannedbuffer = Values(2)
        noadd = CrearDetalleDeOT
        editar = True
        scannedbuffer = vbNullString

        cod = vbNullString

    Else

        Set tmp = New remitoDetalle
        tmp.Origen = OrigenRemitoConcepto

        tmp.Concepto = UCase(Values(2))
        tmp.Cantidad = CDbl(Values(4))
        If grilla.Columns(5).Visible Then tmp.Valor = Values(5)


        tmp.facturable = True
        tmp.Facturado = False
        tmp.FEcha = Now

        Remito.detalles.Add tmp
    End If
End Sub


Private Sub grilla_UnboundDelete(ByVal rowIndex As Long, ByVal Bookmark As Variant)
    If rowIndex > 0 And Remito.detalles.count > 0 Then

        'eliminar de la entrega....!!!

        Set tmp = Remito.detalles(rowIndex)
        If tmp.Origen = OrigenRemitoOt Or tmp.Origen = OrigenRemitoAplicado Then

            If DAORemitoSDetalle.Delete(tmp) Then
                Remito.detalles.remove rowIndex
            Else
                MsgBox "Se produjo algún error!", vbCritical
            End If

        Else
            Remito.detalles.remove rowIndex
        End If

    End If
End Sub


Private Sub grilla_UnboundReadData(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error Resume Next
    If rowIndex > 0 And Remito.detalles.count > 0 Then
    Debug.Print (Remito.detalles.count)
        Set tmp = Remito.detalles(rowIndex)

        With Values
            .value(1) = rowIndex
            .value(2) = tmp.VerElemento

            If Not IsSomething(tmp.DetallePedido) Then
                .value(7) = tmp.VerOrigen & Chr(10) & tmp.observaciones
            Else
                .value(7) = tmp.VerOrigen & " | " & tmp.DetallePedido.item & Chr(10) & tmp.observaciones
            End If
            .value(4) = funciones.FormatearDecimales(tmp.Cantidad, 2)
            .value(5) = funciones.FormatearDecimales(tmp.Valor, 2)
            .value(6) = tmp.VerFacturado
        End With
    End If

End Sub


Private Sub grilla_UnboundUpdate(ByVal rowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If rowIndex > 0 And Remito.detalles.count > 0 Then
        Set tmp = Remito.detalles.item(rowIndex)
        tmp.Concepto = UCase(Values(2))
        tmp.Cantidad = CDbl(Values(4))


        If grilla.Columns(5).Visible Then tmp.Valor = Values(5)

    End If
End Sub


Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property


Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    mostrarRemito
End Function

Private Sub mnuNoFacturable_Click()
    Dim A As Long
    A = grilla.rowIndex(grilla.row)
    If A > 0 Then
        If DAORemitoSDetalle.CambiarEstadoFacturable(Not tmp.facturable, tmp) Then
            MsgBox "Cambio exitoso!", vbInformation, "Información"
        End If
        grilla.RefreshRowIndex A
    End If
End Sub

