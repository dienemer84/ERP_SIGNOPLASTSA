VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntregas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cerrar OT"
   ClientHeight    =   6480
   ClientLeft      =   3675
   ClientTop       =   5520
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   13335
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remitar"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar remito..."
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Estimados ]"
      Height          =   1575
      Left            =   5760
      TabIndex        =   1
      Top             =   3000
      Width           =   3375
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porc de fabricación"
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
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porc de entregas"
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
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblPorcFab 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPorcEnt 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblAvance 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Avance "
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
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12720
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView lstDetallePedido 
      Height          =   2775
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstEntregas 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Cant"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Rto"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Fecha Remito"
         Object.Width           =   5644
      EndProperty
   End
   Begin VB.Label lblIdOT 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmEntregas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim claseP As New classPlaneamiento
Public idOt As Long
Dim idcli As Long
Dim iditem

Private Sub Command1_Click()
    Dim error1 As Boolean
    
    'verifico que esten todos los ítems fabricados
    error1 = False
    'verifico q no haya alguno entregado completamente
    error2 = False
    'idp = CLng(Me.lblIdOT)
    If Not claseP.estaTodoEntregado(idOt) Then

        If Not claseP.estaCerrado(idOt) Then
            For nn = 1 To Me.lstDetallePedido.ListItems.count
                cantidad_pedida = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(2))
                Cantidad_Fabricada = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(4))
                cantidad_deStock = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(5))
                resto = Cantidad_Fabricada + cantidad_deStock
                claseP.ejecutar_consulta "select estado from pedidos where id=" & idOt
                estado = claseP.estadoOT
                If cantidad_pedida > resto Or estado = 4 Then
                    error1 = True
                End If
            Next nn

            If Not error1 Then
                'si todo lo pedido esta fabricado o proveniente de stock, proceso a realizar la entrega.
                frmEntregas.lblIdOT = idOt
                frmEntregaTotal.Pedido = idOt
                frmEntregaTotal.Show 1
            End If

        Else
            MsgBox "El pedido se encuentra cerrado", vbInformation, "Información"
        End If
    Else
        'el pedido ya se entrego, falta cerrar.
        
        Dim ot As OrdenTrabajo
        Set ot = DAOOrdenTrabajo.FindById(idOt)
        Set ot.detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(ot.Id, True, True, True)
    
        If DAOOrdenTrabajo.Cerrar(ot) Then
            MsgBox "El pedido " & idOt & " se cerro correctamente.", vbInformation, "Información"
            'Unload Me
        End If
    End If



    If error1 Then
        MsgBox "Para cerrar el pedido debe tener todo fabricado o proveniente de stock.", vbCritical, "Error"
    End If
    verPorcentajes
End Sub

Private Sub Command2_Click()
    Me.realizaEntrega
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    On Error GoTo err4
    Me.CommonDialog1.ShowPrinter
    For x = 1 To Me.CommonDialog1.Copies
        ImprimirEntregas
    Next
    Exit Sub
err4:
End Sub

Private Function ImprimirEntregas()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Printer.Font.size = 10
    espacio = 0
    Printer.Font.Bold = True
    Printer.Orientation = 1
    Printer.Print "DETALLE DE ENTREGAS O/T Nro " & Format(idOt, "0000") & " al día " & Format(Now, "dd-mm-yyyy")

    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    Set rs = conectar.RSFactory("select p.*,c.razon from pedidos p inner join clientes c on p.idcliente=c.id where p.id=" & idOt)
    If Not rs.EOF And Not rs.BOF Then
        cli = rs!idCliente
        clie = rs!Razon
        referencia = rs!Descripcion
        entrega = rs!FechaEntrega
    Else
        Exit Function
    End If

    Printer.Print "Cliente: " & cli & " - " & clie
    Printer.Print "Referencia: " & UCase(referencia)
    Printer.Print "Entrega: " & Format(entrega, "dd-mm-yyyy")
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)


    'Acá se imprimen los encabezados de la Lista
    Printer.Print Tab(1);
    Printer.Print "Item";
    Printer.Print Tab(10);
    Printer.Print "Detalle";
    Printer.Print Tab(80);
    Printer.Print "Cant";
    Printer.Print Tab(90);
    Printer.Print "Entregados"
    Printer.Font.Bold = False
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)


    'aca se imprime la lista de elementos con sus entregas


    Set rs = conectar.RSFactory("select dp.id,dp.item,dp.cantidad as cant,dp.cantidad_entregada as entregados,s.detalle from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where idPedido=" & idOt)
    While Not rs.EOF
        Printer.Print Tab(1);
        Printer.Print Format(rs!Item, "000");
        Printer.Print Tab(12);
        Printer.Print UCase(rs!detalle);
        Printer.Print Tab(90);
        Printer.Print rs!Cant;
        Printer.Print Tab(100);
        Printer.Print rs!Entregados
        Set rs2 = conectar.RSFactory("select e.cantidad,e.remito,r.fecha from entregas e inner join remitos r on e.remito=r.id where idDetallePedido=" & rs!id & " and r.estado <> 3")
        c = 0

        While Not rs2.EOF
            c = c + 1
            rs2.MoveNext

        Wend
        If c > 0 Then
            Printer.Print
            Printer.FontBold = True
            Printer.Print Tab(65);
            Printer.Print "Cant";
            Printer.Print Tab(75);
            Printer.Print "Remito";
            Printer.Print Tab(85);
            Printer.Print "Fecha";
            Printer.FontBold = False
            rs2.MoveFirst
            While Not rs2.EOF
                Printer.Print Tab(75);
                Printer.Print rs2!cantidad;
                Printer.Print Tab(85);
                Printer.Print rs2!remito;
                Printer.Print Tab(95);
                Printer.Print Format(rs2!FEcha, "dd-mm-yyyy")

                rs2.MoveNext
            Wend
        End If

        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print
        rs.MoveNext
    Wend





    'Otro espacio en blanco



    Printer.Print

    ''Imprime la línea de final de impresión

    'Texto del pie>

    Printer.Print "Fecha emisión " & Format(Date, "dd-mm-yyyy")


    'Comenzamos la impresión
    Printer.EndDoc



End Function

Private Sub Command5_Click()
    Dim rs1 As Recordset
    Dim idDetallePedido As Long
    Dim ide As Long
    Dim rs2 As Recordset
    idDetallePedido = Me.lstDetallePedido.selectedItem.Tag
    Set rs1 = conectar.RSFactory("select reserva_stock as deStock, cantidad as cantPedida,cantidad_fabricados as fabricados, cantidad_entregada as entregados from detalles_pedidos where id=" & idDetallePedido)
    If Not rs1.EOF And Not rs1.BOF Then
        Pedido = rs1!cantpedida
        Fabricados = rs1!Fabricados
        Entregados = rs1!Entregados
        deStock = rs1!deStock
        disponibles = Fabricados + deStock
        paraEntregar = disponibles - Entregados
        faltantes = Pedido - Entregados


        If paraEntregar > 0 Then
            '<= faltantes Then
            'si hay elementos disponibles, procedo con elegir el item del remito que voy a aplicar
            'a esta OT

            'frmPlaneamientoRemitosListaProceso.idCliMostrar = idcli
            frmPlaneamientoRemitosListaProceso.Mostrar = -1
            frmPlaneamientoRemitosListaProceso.Show 1
            idRem = Selecciones.RemitoElegido.id

            If idRem = -1 Then Exit Sub

            frmPlaneamientoRemitosDetalle.rtoNro = idRem
            frmPlaneamientoRemitosDetalle.usarItem = True
            frmPlaneamientoRemitosDetalle.Show 1

            ide = funciones.itemRemito

            If ide < 0 Then Exit Sub
            'End If
            Set rs2 = conectar.RSFactory("select cantidad from entregas where id=" & ide)
            If Not rs2.EOF And Not rs2.BOF Then
                If rs2!cantidad <= faltantes Then

                    If MsgBox("¿Está seguro de aplicar este remito a este item de la OT?", vbYesNo, "Confirmación") = vbYes Then
                        If claseP.aplicarRemitoAOT(idOt, ide, idDetallePedido, rs2!cantidad) Then
                            MsgBox "Remito aplicado correctamente!", vbInformation, "Información"
                            funciones.itemRemito = -1
                        Else
                            MsgBox "Se produjo algún error. No se graban los cambios!", vbError, "Información"
                        End If
                    End If
                End If
            Else
                MsgBox "Se produjo un error. No se puede continuar!", vbCritical, "Error"
                Exit Sub
            End If


            'si esta aca es pq es factible aplicar el rto






        Else

            MsgBox "No hay elementos disponibles para entregar!", vbInformation, "Error"
        End If



    Else
        MsgBox "Se produjo un error. No se puede continuar!", vbCritical, "Error"
    End If


    Set rs1 = Nothing
    Set rs2 = Nothing

    Form_Activate
End Sub

Private Sub Form_Activate()
    Me.Refresh
    Dim iditem
    idOt = CLng(Me.lblIdOT)
    claseP.ejecutar_consulta "select idcliente from pedidos where id=" & idOt
    idcli = claseP.idCliente

    If claseP.ExistePedido(idOt) Then
        claseP.llenar_lista_detalle Me.lstDetallePedido, idOt, 1
        iditem = Me.lstDetallePedido.selectedItem.Tag
        llenarLstEntregas iditem
        If claseP.estaCerrado(idOt) Then
            Me.Command1.Enabled = False
        Else
            Me.Command1.Enabled = True
        End If
    End If
    verPorcentajes

End Sub

Private Sub Form_Load()
FormHelper.Customize Me
End Sub

Private Sub lstDetallePedido_Click()
    iditem = Me.lstDetallePedido.selectedItem.Tag
    llenarLstEntregas iditem

End Sub
Function realizaEntrega()
    c = 0
    erro = 0
    For P = 1 To Me.lstDetallePedido.ListItems.count
        If Me.lstDetallePedido.ListItems(P).Selected Then
            c = c + 1
            Fabricados = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(4))
            Entregados = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(5))
            pedidos = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(2))
            deStock = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(6))

            If Fabricados + deStock = 0 Then
                erro = 1
                'MsgBox "Para entregar este item debería tenerlo fabricado", vbCritical, "Error"
            End If
            If pedidos = Entregados Then
                erro = 2
                'MsgBox "El ítem está completamente entregado. No se permiten más entregas", vbCritical, "Error"
            End If
        End If

    Next P


    If erro = 1 Then
        MsgBox "Para hacer la entrega marcada, debería tener los itemes seleccionados" & Chr(10) & "Totalmente fabricados.", vbCritical, "Error"
        Exit Function
    End If


    If c = 1 Then    'entrega uno
        frmPlaneamientoRealizarEntrega.lblIdPieza = Me.lstDetallePedido.selectedItem.Tag
        frmPlaneamientoRealizarEntrega.lblPieza = Me.lstDetallePedido.selectedItem.ListSubItems(1)
        frmPlaneamientoRealizarEntrega.lblPedidos = Me.lstDetallePedido.selectedItem.ListSubItems(2)
        frmPlaneamientoRealizarEntrega.Text1 = pedidos - Fabricados
        frmPlaneamientoRealizarEntrega.lblFabricados = Me.lstDetallePedido.selectedItem.ListSubItems(4)
        frmPlaneamientoRealizarEntrega.lblEntregados = Me.lstDetallePedido.selectedItem.ListSubItems(5)
        frmPlaneamientoRealizarEntrega.lblDeStock = Me.lstDetallePedido.selectedItem.ListSubItems(6)
        frmPlaneamientoRealizarEntrega.lblOT = Me.lblIdOT
        frmPlaneamientoRealizarEntrega.lblItem = Me.lstDetallePedido.selectedItem
        frmPlaneamientoRealizarEntrega.Show 1


    Else    'entrega muchos
        Dim V() As Long
        ReDim Preserve V(c) As Long
        c = 0
        For o = 1 To Me.lstDetallePedido.ListItems.count

            If Me.lstDetallePedido.ListItems(o).Selected Then
                V(c) = Me.lstDetallePedido.ListItems(o).Tag
                c = c + 1
            End If
        Next o
        frmPlaneamientoRealizarEntregaMultiple.IdP = idOt
        frmPlaneamientoRealizarEntregaMultiple.vector V
        frmPlaneamientoRealizarEntregaMultiple.Show 1

    End If
    verPorcentajes
End Function


Public Sub llenarLstEntregas(iditem)
    On Error GoTo err44
    Dim r As Recordset
    Dim estado
    Dim rto
    'Set r = claseP.listaRS("select id,cantidad,remito,fecha from entregas where idDetallePedido=" & iditem & " and origen=1") 'origen 1 es de OT
    Set r = conectar.RSFactory("select d.id,d.cantidad,d.remito,d.fecha,r.estado, r.numero from entregas d inner join remitos r on d.remito=r.id where idDetallePedido=" & iditem)     ' & " and origen=1")

    Dim x As ListItem
    Me.lstEntregas.ListItems.Clear
    While Not r.EOF

        Set x = Me.lstEntregas.ListItems.Add(, , r!id)
        estado = r!estado
        x.SubItems(1) = r!cantidad
        rto = r!Numero
        If estado = 3 Then rto = rto & "*"
        x.SubItems(2) = rto

        x.SubItems(3) = r!FEcha
        If estado = 3 Then
            x.ListSubItems(2).ForeColor = vbRed
        End If
        r.MoveNext
    Wend
    Exit Sub
err44:
    MsgBox Err.Description

End Sub

Private Sub lstDetallePedido_DblClick()
'Me.realizaEntrega
End Sub

Private Sub lstDetallePedido_ItemClick(ByVal Item As MSComctlLib.ListItem)
'funciones.ca
End Sub

Private Sub verPorcentajes()
    Dim fab As Double
    Dim ent As Double
    Dim avance As Double
    claseP.porcentajesOT idOt, fab, ent, avance
    Me.lblPorcEnt = ent & "%"
    Me.lblPorcFab = fab & "%"
    Me.lblAvance = avance & "%"

End Sub
