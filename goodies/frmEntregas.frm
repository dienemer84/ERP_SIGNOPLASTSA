VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmEntregas 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cerrar OT"
   ClientHeight    =   8445
   ClientLeft      =   3675
   ClientTop       =   5520
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   13335
   Begin VB.CommandButton Command6 
      Caption         =   "FIXER"
      Height          =   285
      Left            =   6060
      TabIndex        =   15
      Top             =   7965
      Width           =   1260
   End
   Begin XtremeSuiteControls.TabControl tab1 
      Height          =   3390
      Left            =   90
      TabIndex        =   13
      Top             =   4905
      Width           =   5685
      _Version        =   786432
      _ExtentX        =   10028
      _ExtentY        =   5980
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Ver Entregas"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "grilla_entregas"
      Item(1).Caption =   "Ver Facturas"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gridFacturas"
      Item(2).Caption =   "Ver Fabricación"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "gridFabricados"
      Begin GridEX20.GridEX grilla_entregas 
         Height          =   2775
         Left            =   150
         TabIndex        =   14
         Top             =   450
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   4895
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16744576
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   3
         Column(1)       =   "frmEntregas.frx":0000
         Column(2)       =   "frmEntregas.frx":0110
         Column(3)       =   "frmEntregas.frx":020C
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmEntregas.frx":0300
         FormatStyle(2)  =   "frmEntregas.frx":0438
         FormatStyle(3)  =   "frmEntregas.frx":04E8
         FormatStyle(4)  =   "frmEntregas.frx":059C
         FormatStyle(5)  =   "frmEntregas.frx":0674
         FormatStyle(6)  =   "frmEntregas.frx":072C
         FormatStyle(7)  =   "frmEntregas.frx":080C
         ImageCount      =   0
         PrinterProperties=   "frmEntregas.frx":08F0
      End
      Begin GridEX20.GridEX gridFacturas 
         Height          =   2775
         Left            =   -69850
         TabIndex        =   16
         Top             =   450
         Visible         =   0   'False
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   4895
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16744576
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   4
         Column(1)       =   "frmEntregas.frx":0AC8
         Column(2)       =   "frmEntregas.frx":0BCC
         Column(3)       =   "frmEntregas.frx":0CB8
         Column(4)       =   "frmEntregas.frx":0DA4
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmEntregas.frx":0EA0
         FormatStyle(2)  =   "frmEntregas.frx":0FD8
         FormatStyle(3)  =   "frmEntregas.frx":1088
         FormatStyle(4)  =   "frmEntregas.frx":113C
         FormatStyle(5)  =   "frmEntregas.frx":1214
         FormatStyle(6)  =   "frmEntregas.frx":12CC
         FormatStyle(7)  =   "frmEntregas.frx":13AC
         ImageCount      =   0
         PrinterProperties=   "frmEntregas.frx":1490
      End
      Begin GridEX20.GridEX gridFabricados 
         Height          =   2775
         Left            =   -69850
         TabIndex        =   17
         Top             =   450
         Visible         =   0   'False
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   4895
         Version         =   "2.0"
         AutomaticSort   =   -1  'True
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewRowLines =   1
         ColumnAutoResize=   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         BackColorHeader =   16744576
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmEntregas.frx":1668
         Column(2)       =   "frmEntregas.frx":1778
         FormatStylesCount=   7
         FormatStyle(1)  =   "frmEntregas.frx":1874
         FormatStyle(2)  =   "frmEntregas.frx":19AC
         FormatStyle(3)  =   "frmEntregas.frx":1A5C
         FormatStyle(4)  =   "frmEntregas.frx":1B10
         FormatStyle(5)  =   "frmEntregas.frx":1BE8
         FormatStyle(6)  =   "frmEntregas.frx":1CA0
         FormatStyle(7)  =   "frmEntregas.frx":1D80
         ImageCount      =   0
         PrinterProperties=   "frmEntregas.frx":1E64
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   12030
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7950
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Remitar"
      Height          =   375
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar remito..."
      Height          =   375
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6675
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Estimados ]"
      Height          =   1575
      Left            =   9855
      TabIndex        =   0
      Top             =   4830
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblPorcFab 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPorcEnt 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblAvance 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2280
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12075
      Top             =   1215
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin GridEX20.GridEX grilla 
      Height          =   4695
      Left            =   15
      TabIndex        =   12
      Top             =   -15
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   8281
      Version         =   "2.0"
      AutomaticSort   =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      PreviewColumn   =   "pieza"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      AllowEdit       =   0   'False
      GroupByBoxVisible=   0   'False
      BackColorHeader =   16744576
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   10
      Column(1)       =   "frmEntregas.frx":203C
      Column(2)       =   "frmEntregas.frx":2178
      Column(3)       =   "frmEntregas.frx":2264
      Column(4)       =   "frmEntregas.frx":2360
      Column(5)       =   "frmEntregas.frx":246C
      Column(6)       =   "frmEntregas.frx":2568
      Column(7)       =   "frmEntregas.frx":2670
      Column(8)       =   "frmEntregas.frx":2778
      Column(9)       =   "frmEntregas.frx":2880
      Column(10)      =   "frmEntregas.frx":297C
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmEntregas.frx":2A84
      FormatStyle(2)  =   "frmEntregas.frx":2BBC
      FormatStyle(3)  =   "frmEntregas.frx":2C6C
      FormatStyle(4)  =   "frmEntregas.frx":2D20
      FormatStyle(5)  =   "frmEntregas.frx":2DF8
      FormatStyle(6)  =   "frmEntregas.frx":2EB0
      ImageCount      =   0
      PrinterProperties=   "frmEntregas.frx":2F90
   End
End
Attribute VB_Name = "frmEntregas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim remito As remito
Public Pedido As OrdenTrabajo
Dim detalle As DetalleOrdenTrabajo
Dim Entregas As New Collection
Dim entrega As RemitoDetalle

Private facturas As New Collection
Private factura As factura

Private cantidadesFabricadas As New Collection
Private cantidad As clsDetalleOrdenTrabajoCantidades

Dim rs As recordset

Dim claseP As New classPlaneamiento
Public idOt As Long
Dim idcli As Long
Dim iditem


Private Sub LlenarListaDetalles()
    Me.grilla.ItemCount = 0
     Set Pedido.Detalles = DAODetalleOrdenTrabajo.FindAllByOrdenTrabajo(Pedido.id, True, True, True)
    Me.grilla.ItemCount = Pedido.Detalles.count
End Sub


Private Sub Command1_Click()
''que solo cambie de estado


'
'
'
'    Dim error1 As Boolean
'
'    'verifico que esten todos los ítems fabricados
'    error1 = False
'    'verifico q no haya alguno entregado completamente
'    error2 = False
'    'idp = CLng(Me.lblIdOT)
'
'
'    If Not claseP.estaTodoEntregado(idOt) And Not Pedido.EsMarco Then
'
'        If Not claseP.estaCerrado(idOt) Then
'            For nn = 1 To Me.lstDetallePedido.ListItems.count
'                cantidad_pedida = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(2))
'                Cantidad_Fabricada = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(4))
'                cantidad_deStock = CLng(Me.lstDetallePedido.ListItems(nn).ListSubItems(5))
'                resto = Cantidad_Fabricada + cantidad_deStock
'                claseP.ejecutar_consulta "select estado from pedidos where id=" & idOt
'                estado = claseP.estadoOT
'                If cantidad_pedida > resto Or estado = 4 Then
'                    error1 = True
'                End If
'            Next nn
'
'            If Not error1 Then
'                'si todo lo pedido esta fabricado o proveniente de stock, proceso a realizar la entrega.
'
'                frmEntregaTotal.Show 1
'            End If
'
'        Else
'            MsgBox "El pedido se encuentra cerrado", vbInformation, "Información"
'        End If
'    Else
'
'
'        'el pedido ya se entrego, falta cerrar.
'
'        If MsgBox("¿Desea cerrar el pedido?", vbYesNo, "Confirmación") = vbYes Then
'            If claseP.CerrarPedido(idOt) Then
'                MsgBox "El pedido " & idOt & " se cerro correctamente.", vbInformation, "Información"
'                'Unload Me
'            End If
'        End If
'    End If
'
'
'
'    If error1 Then
'        MsgBox "Para cerrar el pedido debe tener todo fabricado o proveniente de stock.", vbCritical, "Error"
'    End If
'    verPorcentajes
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
        imprimirEntregas
    Next
    Exit Sub
err4:
End Sub

Private Function imprimirEntregas()
    Dim rs As recordset
    Dim rs2 As recordset
    Printer.Font.Size = 10
    Espacio = 0
    Printer.Font.Bold = True
    Printer.Orientation = 1
    Printer.Print "DETALLE DE ENTREGAS O/T Nro " & Format(idOt, "0000")

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
        Printer.Print rs!entregados
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


    If Pedido.EsMarco Then
        MsgBox "No puede remitar contrato marco!", vbCritical, "Error"
        Exit Sub
    End If

    Dim rs1 As recordset
    Dim idDetallePedido As Long
    Dim ide As Long
    Dim rs2 As recordset
    idDetallePedido = detalle.id
    Dim cantpedida

    Set rs1 = conectar.RSFactory("select reserva_stock as deStock, cantidad as cantPedida,cantidad_fabricados as fabricados, cantidad_entregada as entregados from detalles_pedidos where id=" & idDetallePedido)
    If Not rs1.EOF And Not rs1.BOF Then
        
        cantpedida = rs1!cantpedida
        fabricados = rs1!fabricados
        entregados = rs1!entregados
        deStock = rs1!deStock
        disponibles = fabricados + deStock
        paraEntregar = disponibles - entregados
        faltantes = cantpedida - entregados


        If paraEntregar > 0 Then
            '<= faltantes Then
            'si hay elementos disponibles, procedo con elegir el item del remito que voy a aplicar
            'a esta OT


            'frmPlaneamientoRemitosListaProceso.idCliMostrar = idcli
            frmPlaneamientoRemitosListaProceso.Mostrar = -1
            frmPlaneamientoRemitosListaProceso.Show 1
            idRem = funciones.queRemitoElegido

            If idRem = -1 Then Exit Sub

            frmPlaneamientoRemitosDetalle.rtoNro = idRem
            frmPlaneamientoRemitosDetalle.usarItem = True
            frmPlaneamientoRemitosDetalle.Show 1

            ide = funciones.itemRemito

            If ide < 0 Then Exit Sub
            'End If
            Set rs2 = conectar.RSFactory("select cantidad from entregas where id=" & ide)
            If Not rs2.EOF And Not rs2.BOF Then
                
            
                If rs2!cantidad <= faltantes And rs2!cantidad <= paraEntregar Then

                    If MsgBox("¿Está seguro de aplicar este remito a este item de la OT?", vbYesNo, "Confirmación") = vbYes Then
                        If claseP.aplicarRemitoAOT(idOt, ide, idDetallePedido, rs2!cantidad) Then
                            DAODetalleOrdenTrabajo.SaveCantidad idDetallePedido, rs2!cantidad, CantidadEntregada_, 0
                            MsgBox "Remito aplicado correctamente!", vbInformation, "Información"
                            funciones.itemRemito = -1
                        Else
                            MsgBox "Se produjo algún error. No se graban los cambios!", vbError, "Información"
                        End If
                    End If
                Else
                    MsgBox "La cant del remito es mayor a la entregar o mayor a los faltantes", vbInformation
                    Exit Sub
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


End Sub

Private Sub Command6_Click()

'
'Dim rs As recordset
'
Dim rs1 As recordset
'

Set rs1 = conectar.RSFactory("select * from detalles_pedidos  where idPedido=1175")
While Not rs1.EOF And Not rs1.BOF

'If rs1!id = 29059 Then
'    Stop
    Set detalle = DAODetalleOrdenTrabajo.FindById(rs1!id, True, True, True)
'End If


Dim a


If detalle.Cantidad_Entregada <> rs1!Cantidad_Entregada Then
    DAODetalleOrdenTrabajo.SaveCantidad detalle.id, rs1!Cantidad_Entregada - detalle.Cantidad_Entregada, CantidadEntregada_, 0

    Debug.Print rs1!idpedido
End If


If detalle.Cantidad_Fabricada <> rs1!cantidad_fabricados Then
DAODetalleOrdenTrabajo.SaveCantidad detalle.id, rs1!cantidad_fabricados - detalle.Cantidad_Fabricada, CantidadFabricada_, 0
    Debug.Print rs1!idpedido
End If

If detalle.Cantidad_Facturada <> rs1!Cantidad_Facturada Then
    Debug.Print rs1!idpedido
    DAODetalleOrdenTrabajo.SaveCantidad detalle.id, rs1!Cantidad_Facturada - detalle.Cantidad_Facturada, CantidadFacturada_, rs1!Precio

End If
rs1.MoveNext
Wend






End Sub

Private Sub Form_Load()
    
    idOt = Pedido.id
    
    If claseP.ExistePedido(idOt) Then
        
        
        
        If claseP.estaCerrado(idOt) Then
            Me.Command1.Enabled = False
        Else
            Me.Command1.Enabled = True
        End If
    End If
    Set Pedido = DAOOrdenTrabajo.FindById(idOt)
    verPorcentajes
    
    
    
    
    FormHelper.Customize Me
    GridEXHelper.customizeGrid Me.grilla, False, False
    GridEXHelper.customizeGrid Me.grilla_entregas, False, False
    GridEXHelper.customizeGrid Me.gridFacturas, False, False
    GridEXHelper.customizeGrid Me.gridFabricados, False, False
    
    Me.caption = "Cerrar Pedido Nro. " & Format(Pedido.id, "0000")
 
    LlenarListaDetalles
 
End Sub
Private Sub LlenarEntregas()
    Set Entregas = DAORemitoSDetalle.FindAllByDetallePedido(detalle.id)
    Me.grilla_entregas.ItemCount = 0
    Me.grilla_entregas.ItemCount = Entregas.count
End Sub

Private Sub gridFabricados_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
If cantidadesFabricadas.count > 0 Then
    Set cantidad = cantidadesFabricadas.Item(RowIndex)
    Values(1) = cantidad.FEcha
    Values(2) = cantidad.cantidad
End If
End Sub

Private Sub gridFacturas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
If facturas.count > 0 Then

    Dim facDeta As FacturaDetalle
    
    Set factura = facturas.Item(RowIndex)
    Values(1) = factura.NumeroFormateado
    Values(3) = factura.FechaEmision
    
    Dim sumaCant As Double
    Dim sumaTotal As Double
    For Each facDeta In factura.Detalles
        If facDeta.DetalleRemito.idDetallePedido = detalle.id Then
            sumaCant = sumaCant + facDeta.DetalleRemito.cantidad
            sumaTotal = sumaTotal + facDeta.Total
        End If
    Next
    Values(2) = sumaTotal
    Values(4) = sumaCant
End If
End Sub

Private Sub grilla_entregas_RowFormat(RowBuffer As GridEX20.JSRowData)
    If Entregas.count > 0 Then
        Set entrega = Entregas.Item(RowBuffer.RowIndex)
        Set rs = conectar.RSFactory("select estado from remitos where id=" & entrega.remito)
        If Not rs.EOF And Not rs.BOF Then
    
            If rs!estado = 3 Then
                RowBuffer.RowStyle = "anulado"
            End If
        End If
    End If
End Sub

Private Sub grilla_entregas_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
If Entregas.count > 0 Then
    Set entrega = Entregas.Item(RowIndex)
    Values(1) = entrega.FEcha
    Values(2) = entrega.cantidad
    Values(3) = entrega.remito
End If
    
    
End Sub

Private Sub grilla_SelectionChange()
Dim cont As Double
    
    cont = GetTickCount
    LlenarEntregas
    Debug.Print "entregas:", GetTickCount - cont
    
    cont = GetTickCount
    LlenarFacturas
    Debug.Print "facturas:", GetTickCount - cont
    
    LlenarFabricados
End Sub

Private Sub LlenarFacturas()
Dim q As String
q = "AdminFacturas.id IN (SELECT af.id FROM AdminFacturas af LEFT JOIN AdminFacturasDetalleNueva afdn ON afdn.idFactura = af.id LEFT JOIN entregas e ON e.id = afdn.idEntrega WHERE e.idDetallePedido = " & detalle.id & ")"

 Set facturas = DAOFactura.FindAll(q, True, True)
 Me.gridFacturas.ItemCount = 0
 Me.gridFacturas.ItemCount = facturas.count
End Sub

Private Sub LlenarFabricados()
Set cantidadesFabricadas = DAODetalleOrdenTrabajo.MapCantidad(detalle.id, CantidadFabricada_)
 Me.gridFabricados.ItemCount = 0
 Me.gridFabricados.ItemCount = cantidadesFabricadas.count

End Sub

Private Sub grilla_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If Pedido.Detalles.count > 0 Then
    Set detalle = Pedido.Detalles(RowIndex)
        Values(1) = detalle.Item
        Values(2) = detalle.Nota
        Values(3) = detalle.CantidadPedida
        Values(4) = detalle.FechaEntrega
        Values(5) = detalle.ReservaStock
        Values(6) = detalle.Cantidad_Fabricada
        Values(7) = detalle.Cantidad_Entregada
        Values(8) = detalle.Cantidad_Facturada
        If IsSomething(detalle.pieza) Then
            Values(9) = detalle.pieza.UnidadMedida
            Values(10) = detalle.pieza.nombre
        End If
    End If
End Sub

Function realizaEntrega()
    If Pedido.EsMarco Then
        MsgBox "No puede remitar contrato marco!", vbCritical, "Error"
        Exit Function
    End If

    c = 0
    erro = 0
'    For P = 1 To Me.lstDetallePedido.ListItems.count
'        If Me.lstDetallePedido.ListItems(P).Selected Then
'            c = c + 1
'            fabricados = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(4))
'            entregados = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(5))
'            pedidos = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(2))
'            deStock = CLng(Me.lstDetallePedido.ListItems(P).ListSubItems(6))
'
'            If fabricados + deStock = 0 Then
'                erro = 1
'                'MsgBox "Para entregar este item debería tenerlo fabricado", vbCritical, "Error"
'            End If
'            If pedidos = entregados Then
'                erro = 2
'                'MsgBox "El ítem está completamente entregado. No se permiten más entregas", vbCritical, "Error"
'            End If
'        End If
'
'    Next P
'
'
'    If erro = 1 Then
'        MsgBox "Para hacer la entrega marcada, debería tener los itemes seleccionados" & Chr(10) & "Totalmente fabricados.", vbCritical, "Error"
'        Exit Function
'    End If
'
'
'    If c = 1 Then    'entrega uno
'        frmPlaneamientoRealizarEntrega.lblIdPieza = Me.lstDetallePedido.SelectedItem.Tag
'        frmPlaneamientoRealizarEntrega.lblPieza = Me.lstDetallePedido.SelectedItem.ListSubItems(1)
'        frmPlaneamientoRealizarEntrega.lblPedidos = Me.lstDetallePedido.SelectedItem.ListSubItems(2)
'        frmPlaneamientoRealizarEntrega.Text1 = pedidos - fabricados
'        frmPlaneamientoRealizarEntrega.lblFabricados = Me.lstDetallePedido.SelectedItem.ListSubItems(4)
'        frmPlaneamientoRealizarEntrega.lblEntregados = Me.lstDetallePedido.SelectedItem.ListSubItems(5)
'        frmPlaneamientoRealizarEntrega.lblDeStock = Me.lstDetallePedido.SelectedItem.ListSubItems(6)
'        frmPlaneamientoRealizarEntrega.lblOT = Pedido.id
'        frmPlaneamientoRealizarEntrega.lblItem = Me.lstDetallePedido.SelectedItem
'        frmPlaneamientoRealizarEntrega.Show 1
'
'
'    Else    'entrega muchos
'        Dim V() As Long
'        ReDim Preserve V(c) As Long
'        c = 0
'        For o = 1 To Me.lstDetallePedido.ListItems.count
'
'            If Me.lstDetallePedido.ListItems(o).Selected Then
'                V(c) = Me.lstDetallePedido.ListItems(o).Tag
'                c = c + 1
'            End If
'        Next o
'        frmPlaneamientoRealizarEntregaMultiple.idP = idOt
'        frmPlaneamientoRealizarEntregaMultiple.vector V
'        frmPlaneamientoRealizarEntregaMultiple.Show 1
'
'    End If
'    verPorcentajes
End Function



Private Sub verPorcentajes()
    Dim fab As Double
    Dim ent As Double
    Dim avance As Double
    claseP.porcentajesOT idOt, fab, ent, avance
    Me.lblPorcEnt = ent & "%"
    Me.lblPorcFab = fab & "%"
    Me.lblAvance = avance & "%"

End Sub

