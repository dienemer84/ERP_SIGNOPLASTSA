VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminPagosCrearOrdenPagoNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   14535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   14535
   ScaleWidth      =   13920
   Begin VB.Frame Frame 
      Caption         =   "Totales"
      Height          =   2295
      Index           =   6
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   7815
   End
   Begin VB.Frame Frame 
      Caption         =   "Detalle de retenciones e Ingresos Brutos"
      Height          =   2175
      Index           =   5
      Left            =   5880
      TabIndex        =   6
      Top             =   2400
      Width           =   7815
   End
   Begin VB.Frame Frame 
      Caption         =   "Cargar pago por comprobante"
      Height          =   4215
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   5535
   End
   Begin VB.Frame Frame 
      Caption         =   "Detalles de la OP"
      Height          =   4455
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      Begin XtremeSuiteControls.PushButton btnConfirmarProveedor 
         Height          =   315
         Left            =   4200
         TabIndex        =   9
         Top             =   360
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboProveedores 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3975
         _Version        =   786432
         _ExtentX        =   7011
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "cboProveedores"
      End
   End
   Begin VB.Frame Frame 
      Height          =   12135
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.Frame Frame 
         Caption         =   "Listado de comprobantes"
         Height          =   4215
         Index           =   2
         Left            =   5760
         TabIndex        =   3
         Top             =   4560
         Width           =   7815
         Begin XtremeSuiteControls.PushButton btnExportarListBox 
            Height          =   375
            Left            =   6120
            TabIndex        =   11
            Top             =   3720
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin VB.ListBox lstFacturas 
            Height          =   3180
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   7575
         End
         Begin XtremeSuiteControls.CommonDialog CommonDialog1 
            Left            =   120
            Top             =   3600
            _Version        =   786432
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   4
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Detalle de pagos realizados"
         Height          =   3135
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   8760
         Width           =   13455
         Begin GridEX20.GridEX datagrid_PagosRealizados 
            Height          =   2655
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   4683
            Version         =   "2.0"
            BoundColumnIndex=   ""
            ReplaceColumnIndex=   ""
            ScrollToolTipColumn=   ""
            ColumnHeaderHeight=   285
            IntProp1        =   0
            IntProp2        =   0
            IntProp7        =   0
            ColumnsCount    =   2
            Column(1)       =   "frmAdminPagosCrearOrdenPagoNew.frx":0000
            Column(2)       =   "frmAdminPagosCrearOrdenPagoNew.frx":00C8
            FormatStylesCount=   6
            FormatStyle(1)  =   "frmAdminPagosCrearOrdenPagoNew.frx":016C
            FormatStyle(2)  =   "frmAdminPagosCrearOrdenPagoNew.frx":02A4
            FormatStyle(3)  =   "frmAdminPagosCrearOrdenPagoNew.frx":0354
            FormatStyle(4)  =   "frmAdminPagosCrearOrdenPagoNew.frx":0408
            FormatStyle(5)  =   "frmAdminPagosCrearOrdenPagoNew.frx":04E0
            FormatStyle(6)  =   "frmAdminPagosCrearOrdenPagoNew.frx":0598
            ImageCount      =   0
            PrinterProperties=   "frmAdminPagosCrearOrdenPagoNew.frx":0678
         End
      End
   End
End
Attribute VB_Name = "frmAdminPagosCrearOrdenPagoNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private id_susc As String
Dim formLoading As Boolean
Dim formLoaded As Boolean
Dim alicuotas As New Collection

Dim total_por_factura As New Dictionary
Dim vFactElegida As clsFacturaProveedor
Dim vCompeElegido As Compensatorio
Dim vFacturaProveedor As clsFacturaProveedor
Dim colProveedores As New Collection
Dim colFacturas As New Collection
Dim colDeudaCompensatorios As New Collection
Dim prov As clsProveedor
Dim Factura As clsFacturaProveedor

Private Banco As Banco
Private caja As caja
Private CuentaBancaria As CuentaBancaria
Private moneda As clsMoneda
Private alicuotaRetencion As DTORetencionAlicuota
Private cuentasBancarias As New Collection
Private retenciones As New Collection
Private Monedas As New Collection
Private Cajas As New Collection
Private bancos As New Collection
Private chequesDisponibles As New Collection
Private chequeras As New Collection

Dim compe As Compensatorio

Private OrdenPago As New OrdenPago
Private operacion As operacion
Private cheque As cheque
Private tmpChequera As chequera

Private chequesChequeraSeleccionada As New Collection

Public ReadOnly As Boolean


Private Sub btnConfirmarProveedor_Click()
    If Me.cboProveedores.ListIndex <> -1 Then

        Set prov = colProveedores.item(CStr(Me.cboProveedores.ItemData(Me.cboProveedores.ListIndex)))
        
        Dim d As clsDTOPadronIIBB

        Set d = DTOPadronIIBB.FindByCUIT(prov.Cuit, TipoPadronRetencion)

''        If IsSomething(d) Then
''            Me.txtRetenciones = str(d.alicuota)   ' Val(d.Retencion )
''        Else
''            Me.txtRetenciones = 0
''        End If

    Else
        Set prov = Nothing
    End If

    MostrarFacturas
    
'''    MostrarDeudaCompensatorios
'''
'''    btnCargar_Click
    
End Sub

Private Sub btnExportarListBox_Click()
    ' Abre el cuadro de diálogo para seleccionar la ubicación y el nombre del archivo
    CommonDialog1.filter = "Archivos de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.ShowSave

    ' Construye el nombre del archivo con el formato deseado
    Dim nombreArchivo As String
    nombreArchivo = "COMPROBANTES_" & Format(Now, "hhmmss") & ".TXT"

    ' Asigna el nombre de archivo personalizado al cuadro de diálogo
    CommonDialog1.filename = nombreArchivo

    If CommonDialog1.filename = "" Then
        Exit Sub ' El usuario canceló la selección
    End If

    ' Abre el archivo para escritura
    Open CommonDialog1.filename For Output As #1

    Dim i As Integer

    ' Recorre los elementos del ListBox y escribe cada elemento en una nueva línea
    For i = 0 To Me.lstFacturas.ListCount - 1
        Print #1, Me.lstFacturas.list(i)
    Next i

    ' Cierra el archivo
    Close #1

    MsgBox "Contenido exportado exitosamente al archivo " & CommonDialog1.filename, vbInformation
End Sub

Private Sub Form_Load()
    
'''    formLoading = True
    
'''    Me.Left = frmPrincipal.ScaleWidth / 6
'''    Me.Top = frmPrincipal.ScaleHeight / 22
    
'''    Me.gridChequeras.Visible = False
'''    Me.gridChequesChequera.Visible = False
    
'''    Me.gridCompensatorios.ItemCount = 0
    
'''    id_susc = funciones.CreateGUID
'''    Channel.AgregarSuscriptor Me, PasajeChequePropioCartera
    FormHelper.Customize Me
    
'''    GridEXHelper.CustomizeGrid Me.gridCajaOperaciones, False, True
'''    GridEXHelper.CustomizeGrid Me.gridDepositosOperaciones, False, True
'''    GridEXHelper.CustomizeGrid Me.gridCheques, False, True
'''    GridEXHelper.CustomizeGrid Me.gridChequesDisponibles, False, False
'''    GridEXHelper.CustomizeGrid Me.gridBancos, False, False
'''    GridEXHelper.CustomizeGrid Me.gridCuentasBancarias, False, False
'''    GridEXHelper.CustomizeGrid Me.gridMonedas, False, False
'''    GridEXHelper.CustomizeGrid Me.gridCajas, False, False
'''    GridEXHelper.CustomizeGrid Me.gridChequeras, False, False
'''    GridEXHelper.CustomizeGrid Me.gridChequesPropios, False, True
'''    GridEXHelper.CustomizeGrid Me.gridCompensatorios, False, True
'''    GridEXHelper.CustomizeGrid Me.gridChequesChequera
'''    GridEXHelper.CustomizeGrid Me.gridRetenciones, False, True
'''
'''
'''
'''    Set Cajas = DAOCaja.FindAll()
'''    Me.gridCajas.ItemCount = Cajas.count
'''
'''    Set Monedas = DAOMoneda.GetAll()
'''    Me.gridMonedas.ItemCount = Monedas.count
'''
'''    Set cuentasBancarias = DAOCuentaBancaria.FindAll()
'''    Me.gridCuentasBancarias.ItemCount = cuentasBancarias.count
'''
'''    Set bancos = DAOBancos.GetAll()
'''    Me.gridBancos.ItemCount = bancos.count
'''
'''    Set chequeras = DAOChequeras.FindAllWithChequesDisponibles()
'''    Me.gridChequeras.ItemCount = chequeras.count
'''
'''
'''    CargarChequesDisponibles


    Set colProveedores = DAOProveedor.FindAllProveedoresWithFacturasImpagas
    
    For Each prov In colProveedores
        Me.cboProveedores.AddItem prov.RazonSocial
        Me.cboProveedores.ItemData(Me.cboProveedores.NewIndex) = prov.id
    Next

'''    Dim cuentasContables As Collection
'''    Set cuentasContables = DAOCuentaContable.GetAll()
'''    Dim cc As clsCuentaContable
'''    Me.cboCuentas.Clear
'''    For Each cc In cuentasContables
'''        cboCuentas.AddItem cc.nombre & " - " & cc.codigo
'''        cboCuentas.ItemData(cboCuentas.NewIndex) = cc.Id
'''    Next cc


'''    radioFacturaProveedor_Click
'''
'''    Me.gridCajaOperaciones.ItemCount = OrdenPago.OperacionesCaja.count
'''
'''    Me.gridDepositosOperaciones.ItemCount = OrdenPago.OperacionesBanco.count
'''
'''    Me.gridCheques.ItemCount = OrdenPago.ChequesTerceros.count
'''    Me.gridChequesPropios.ItemCount = OrdenPago.ChequesPropios.count


'''    Set Me.gridCheques.Columns("numero").DropDownControl = Me.gridChequesDisponibles
'''
'''    Set Me.gridDepositosOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
'''    Set Me.gridDepositosOperaciones.Columns("cuenta").DropDownControl = Me.gridCuentasBancarias
'''
'''    Set Me.gridCajaOperaciones.Columns("moneda").DropDownControl = Me.gridMonedas
'''    Set Me.gridCajaOperaciones.Columns("caja").DropDownControl = Me.gridCajas
'''
'''    Set Me.gridChequesPropios.Columns("chequera").DropDownControl = Me.gridChequeras
'''    Set Me.gridChequesPropios.Columns("numero").DropDownControl = Me.gridChequesChequera
'''    gridChequesChequera.ItemCount = 0
'''    GridEXHelper.AutoSizeColumns Me.gridChequeras


'''    DAOMoneda.llenarComboXtremeSuite Me.cboMonedas

'''    Me.dtpFecha.value = OrdenPago.FEcha

'''    Totalizar

    formLoaded = True
    formLoading = False
End Sub

Private Sub MostrarFacturas()

    Me.lstFacturas.Clear

    If IsSomething(prov) Then
        Set colFacturas = DAOFacturaProveedor.FindAll("AdminComprasFacturasProveedores.id_proveedor=" & prov.id & " and (AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.Aprobada & " or AdminComprasFacturasProveedores.estado=" & EstadoFacturaProveedor.pagoParcial & ")", False, "", False, True)

        If OrdenPago.id <> 0 And OrdenPago.EsParaFacturaProveedor Then
            If prov.id = OrdenPago.FacturasProveedor.item(1).Proveedor.id Then
                For Each Factura In OrdenPago.FacturasProveedor
                    If Not funciones.BuscarEnColeccion(colFacturas, CStr(Factura.id)) Then

                        colFacturas.Add DAOFacturaProveedor.FindById(Factura.id), CStr(Factura.id)
                    End If
                Next
            End If
        End If

        Dim T As String

        For Each Factura In colFacturas    'en ese for traigo los pendientes a abonar que estan asociados a ops sin aprobar

            Dim c As Collection
            Set c = DAOOrdenPago.FindAbonadoPendiente(Factura.id, OrdenPago.id)

            Factura.TotalAbonadoGlobalPendiente = 0    ' c(1) 'que esta en ops sin aprobar
            Factura.NetoGravadoAbonadoGlobalPendiente = 0    ' c(2)
            Factura.OtrosAbonadoGlobalPendiente = 0    'c(3)

            T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.total & ")" & " (" & Factura.FEcha & ")" & " TC: (" & Factura.TipoCambioPago & ")"
            If Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente > 0 Then
                T = Factura.NumeroFormateado & " (" & Factura.moneda.NombreCorto & " " & Factura.total & " - Abonado: " & Factura.TotalAbonadoGlobal + Factura.TotalAbonadoGlobalPendiente & ")" & " (" & Factura.FEcha & ")" & " TC: (" & Factura.TipoCambioPago & ")"

                'MsgBox (c.count)

            End If

            Me.lstFacturas.AddItem T
            Me.lstFacturas.ItemData(Me.lstFacturas.NewIndex) = Factura.id


        Next

        ' 22/08/2022
        'AGREGO UN LABEL QUE MUESTRA LA CANTIDAD DE COMPROBANTES MOSTRADOS EN EL LIST

'''        Me.lblCantidadComprobantes.caption = "Cbtes. Mostrados: " & colFacturas.count

    Else

        Set colFacturas = New Collection

        'MsgBox (colFacturas.count)

    End If

End Sub

Private Sub lstFacturas_Click()
'    'debug.print (Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex))

    Set vFactElegida = colFacturas.item(CStr(Me.lstFacturas.ItemData(Me.lstFacturas.ListIndex)))

    If IsSomething(vFactElegida) Then

        Dim c As Collection

        If OrdenPago.estado = EstadoOrdenPago_pendiente And vFactElegida.NetoGravadoAbonado = 0 And vFactElegida.OtrosAbonado = 0 Then
            Set c = DAOOrdenPago.FindAbonadoFactura(vFactElegida.id, OrdenPago.id)

            vFactElegida.NetoGravadoAbonado = c(2)
            vFactElegida.OtrosAbonado = c(3)
        End If

'    MostrarPago vFactElegida
    
    MostrarPagos vFactElegida
    
'    RecalcularFacturaElegida
        
    End If
'    Totalizar
End Sub

Private Sub MostrarPagos(F As clsFacturaProveedor)

        Dim PagosRealizados As clsFacturaProveedor
        
        Set PagosRealizados = DAOFacturaProveedor.FindById(F.id)
        
'        Me.gridRetenciones.ItemCount = .RetencionesAlicuota.count
'        Set alicuotas = .RetencionesAlicuota
        
End Sub


