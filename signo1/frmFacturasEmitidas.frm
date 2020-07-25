VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmAdminFacturasEmitidas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Comprobantes Emitidos"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   Icon            =   "frmFacturasEmitidas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11955
   Begin XtremeSuiteControls.GroupBox grp 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   -15
      Width           =   18990
      _Version        =   786432
      _ExtentX        =   33496
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Filtros"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkCredito 
         Height          =   255
         Left            =   11400
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "DE CRÉDITO"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   9450
         TabIndex        =   23
         Top             =   1350
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtReferencia 
         Height          =   300
         Left            =   11400
         TabIndex        =   20
         Top             =   690
         Width           =   2835
      End
      Begin XtremeSuiteControls.ComboBox cboClientes 
         Height          =   315
         Left            =   1620
         TabIndex        =   19
         Top             =   720
         Width           =   3510
         _Version        =   786432
         _ExtentX        =   6191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
      End
      Begin VB.TextBox txtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1620
         TabIndex        =   4
         Top             =   330
         Width           =   1290
      End
      Begin VB.TextBox txtOrdenCompra 
         Height          =   300
         Left            =   1605
         TabIndex        =   3
         Top             =   1110
         Width           =   1740
      End
      Begin VB.TextBox txtRemitoAplicado 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   3960
         TabIndex        =   2
         Top             =   1110
         Width           =   1170
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   5190
         TabIndex        =   5
         Top             =   735
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Default         =   -1  'True
         Height          =   420
         Left            =   14640
         TabIndex        =   6
         Top             =   960
         Width           =   1650
         _Version        =   786432
         _ExtentX        =   2910
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1050
         Left            =   5640
         TabIndex        =   7
         Top             =   270
         Width           =   4695
         _Version        =   786432
         _ExtentX        =   8281
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Fecha Emision"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.DateTimePicker dtpDesde 
            Height          =   315
            Left            =   825
            TabIndex        =   8
            Top             =   615
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.DateTimePicker dtpHasta 
            Height          =   315
            Left            =   3000
            TabIndex        =   9
            Top             =   630
            Width           =   1470
            _Version        =   786432
            _ExtentX        =   2593
            _ExtentY        =   556
            _StockProps     =   68
            CheckBox        =   -1  'True
            Format          =   1
         End
         Begin XtremeSuiteControls.ComboBox cboRangos 
            Height          =   315
            Left            =   825
            TabIndex        =   10
            Top             =   225
            Width           =   3645
            _Version        =   786432
            _ExtentX        =   6429
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   195
            Left            =   2430
            TabIndex        =   13
            Top             =   675
            Width           =   420
            _Version        =   786432
            _ExtentX        =   741
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Hasta"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   195
            Left            =   255
            TabIndex        =   12
            Top             =   660
            Width           =   465
            _Version        =   786432
            _ExtentX        =   820
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Desde"
            AutoSize        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   285
            Width           =   480
            _Version        =   786432
            _ExtentX        =   847
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Rango"
            AutoSize        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   420
         Left            =   16560
         TabIndex        =   18
         Top             =   360
         Width           =   810
         _Version        =   786432
         _ExtentX        =   1429
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   420
         Left            =   16560
         TabIndex        =   22
         ToolTipText     =   "Exporta sólo pendientes"
         Top             =   960
         Width           =   810
         _Version        =   786432
         _ExtentX        =   1429
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboPuntosVenta 
         Height          =   360
         Left            =   3585
         TabIndex        =   25
         Top             =   300
         Width           =   1530
         _Version        =   786432
         _ExtentX        =   2699
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   285
         Left            =   5190
         TabIndex        =   26
         Top             =   360
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEstados 
         Height          =   360
         Left            =   11400
         TabIndex        =   28
         Top             =   217
         Width           =   2355
         _Version        =   786432
         _ExtentX        =   4154
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         Text            =   "cboMoneda"
         DropDownItemCount=   3
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   13875
         TabIndex        =   30
         Top             =   255
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "X"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblTotalNeto 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   17640
         TabIndex        =   36
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblTotalIVA 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   17640
         TabIndex        =   35
         Top             =   550
         Width           =   1095
      End
      Begin VB.Label lblTotalPercepciones 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   17640
         TabIndex        =   34
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Filtrado $:"
         Height          =   195
         Left            =   17640
         TabIndex        =   33
         Top             =   1200
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   285
         Left            =   10650
         TabIndex        =   32
         Top             =   1065
         Width           =   705
         _Version        =   786432
         _ExtentX        =   1244
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "MiPyMES"
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   285
         Left            =   10800
         TabIndex        =   29
         Top             =   255
         Width           =   555
         _Version        =   786432
         _ExtentX        =   979
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Estado"
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   285
         Left            =   3240
         TabIndex        =   27
         Top             =   330
         Width           =   585
         _Version        =   786432
         _ExtentX        =   1032
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "PV"
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         Height          =   195
         Left            =   10440
         TabIndex        =   21
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Comprobrante"
         Height          =   270
         Left            =   30
         TabIndex        =   17
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden Compra"
         Height          =   270
         Left            =   270
         TabIndex        =   16
         Top             =   1125
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   270
         Left            =   270
         TabIndex        =   15
         Top             =   735
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rto"
         Height          =   270
         Left            =   2655
         TabIndex        =   14
         Top             =   1125
         Width           =   1260
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4395
      Left            =   15
      TabIndex        =   0
      Top             =   1635
      Width           =   19125
      _ExtentX        =   33734
      _ExtentY        =   7752
      Version         =   "2.0"
      PreviewRowIndent=   100
      DefaultGroupMode=   1
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      GroupFooterStyle=   2
      PreviewColumn   =   "preview"
      PreviewRowLines =   1
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      ImageCount      =   1
      ImagePicture1   =   "frmFacturasEmitidas.frx":000C
      RowHeaders      =   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   22
      Column(1)       =   "frmFacturasEmitidas.frx":0326
      Column(2)       =   "frmFacturasEmitidas.frx":04C6
      Column(3)       =   "frmFacturasEmitidas.frx":05B2
      Column(4)       =   "frmFacturasEmitidas.frx":069A
      Column(5)       =   "frmFacturasEmitidas.frx":0786
      Column(6)       =   "frmFacturasEmitidas.frx":0886
      Column(7)       =   "frmFacturasEmitidas.frx":09EA
      Column(8)       =   "frmFacturasEmitidas.frx":0BE2
      Column(9)       =   "frmFacturasEmitidas.frx":0D2A
      Column(10)      =   "frmFacturasEmitidas.frx":0E2A
      Column(11)      =   "frmFacturasEmitidas.frx":0F26
      Column(12)      =   "frmFacturasEmitidas.frx":1026
      Column(13)      =   "frmFacturasEmitidas.frx":117A
      Column(14)      =   "frmFacturasEmitidas.frx":12C2
      Column(15)      =   "frmFacturasEmitidas.frx":141A
      Column(16)      =   "frmFacturasEmitidas.frx":1562
      Column(17)      =   "frmFacturasEmitidas.frx":1656
      Column(18)      =   "frmFacturasEmitidas.frx":173A
      Column(19)      =   "frmFacturasEmitidas.frx":1836
      Column(20)      =   "frmFacturasEmitidas.frx":195A
      Column(21)      =   "frmFacturasEmitidas.frx":1A7E
      Column(22)      =   "frmFacturasEmitidas.frx":1BC2
      FormatStylesCount=   14
      FormatStyle(1)  =   "frmFacturasEmitidas.frx":1CCE
      FormatStyle(2)  =   "frmFacturasEmitidas.frx":1E06
      FormatStyle(3)  =   "frmFacturasEmitidas.frx":1EB6
      FormatStyle(4)  =   "frmFacturasEmitidas.frx":1F6A
      FormatStyle(5)  =   "frmFacturasEmitidas.frx":2042
      FormatStyle(6)  =   "frmFacturasEmitidas.frx":20FA
      FormatStyle(7)  =   "frmFacturasEmitidas.frx":21DA
      FormatStyle(8)  =   "frmFacturasEmitidas.frx":2266
      FormatStyle(9)  =   "frmFacturasEmitidas.frx":2346
      FormatStyle(10) =   "frmFacturasEmitidas.frx":23F6
      FormatStyle(11) =   "frmFacturasEmitidas.frx":24AA
      FormatStyle(12) =   "frmFacturasEmitidas.frx":255A
      FormatStyle(13) =   "frmFacturasEmitidas.frx":262E
      FormatStyle(14) =   "frmFacturasEmitidas.frx":26E2
      ImageCount      =   1
      ImagePicture(1) =   "frmFacturasEmitidas.frx":27BA
      PrinterProperties=   "frmFacturasEmitidas.frx":2AD4
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   15630
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkVerObservaciones 
      Height          =   225
      Left            =   45
      TabIndex        =   24
      Top             =   6105
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Ver Observaciones"
      Appearance      =   6
      Value           =   1
   End
   Begin XtremeSuiteControls.TaskDialog taskDialog 
      Left            =   14955
      Top             =   750
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      WindowTitle     =   "TaskDialog1"
   End
   Begin VB.Menu mnuFacturas 
      Caption         =   "armnuFacturas"
      Visible         =   0   'False
      Begin VB.Menu NRO 
         Caption         =   "nro"
         Enabled         =   0   'False
      End
      Begin VB.Menu editar 
         Caption         =   "Editar"
      End
      Begin VB.Menu aprobarFactura 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu mnuAprobarSinEnvio 
         Caption         =   "Aprobar sin envío a AFIP"
      End
      Begin VB.Menu mnuEditarCAE 
         Caption         =   "Editar datos de CAE"
      End
      Begin VB.Menu mnuDesaprobarFactura 
         Caption         =   "Desaprobar..."
      End
      Begin VB.Menu ImprimirFactura 
         Caption         =   "Imprimir..."
      End
      Begin VB.Menu AnularFactura 
         Caption         =   "Anular"
      End
      Begin VB.Menu desAnular 
         Caption         =   "Quitar Anulación"
         Visible         =   0   'False
      End
      Begin VB.Menu aplicar 
         Caption         =   "Aplicar Recibo..."
      End
      Begin VB.Menu aplicarNCaFC 
         Caption         =   "Aplicar a Factura..."
      End
      Begin VB.Menu mnuAplicarANC 
         Caption         =   "Aplicar a NC..."
      End
      Begin VB.Menu mnuCrearCopiaFactura 
         Caption         =   "Crear copia a partir de comprobante"
      End
      Begin VB.Menu mnuFechaPagoPropuesta 
         Caption         =   "Establecer Fecha Pago Propuesta"
      End
      Begin VB.Menu mnuFechaEntrega 
         Caption         =   "Establecer Fecha Entrega..."
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
      Begin VB.Menu verHistorialFactura 
         Caption         =   "Ver Historial..."
      End
      Begin VB.Menu mnuArchivos 
         Caption         =   "Archivos Asociados..."
      End
      Begin VB.Menu verFactura 
         Caption         =   "Ver Detalle..."
      End
      Begin VB.Menu archivos 
         Caption         =   "Archivos Asociados..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu scanear 
         Caption         =   "Adquirir..."
      End
   End
End
Attribute VB_Name = "frmAdminFacturasEmitidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISuscriber

Dim vId As String
Dim facturas As Collection
Dim Factura As Factura
Dim m_Archivos As Dictionary


Private Sub AnularFactura_Click()
    Dim r As Long
    r = Me.GridEX1.RowIndex(Me.GridEX1.row)
    If MsgBox("¿Desea anular el comprobante?", vbYesNo, "Confirmacion") = vbYes Then

        If DAOFactura.Anular(Factura) Then
            MsgBox "Comprobante anulado con éxito!", vbInformation, "Información"
            Me.GridEX1.RefreshRowIndex r
        Else
            MsgBox "Hubo un error. No se anulo el comprobante!", vbCritical, "Error"
        End If

    End If
End Sub



Private Sub aplicarNCaFC_Click()
    If MsgBox("¿Seguro de aplicar NC?", vbYesNo, "Confirmación") = vbYes Then
        'seleccionar factura para aplicar
        Set Selecciones.Factura = Nothing
          Dim F As New frmAdminFacturasNCElegirFC
        
        F.idCliente = Factura.cliente.id
            F.TiposDocs.Add tipoDocumentoContable.Factura
            F.TiposDocs.Add tipoDocumentoContable.notaDebito
            F.EstadosDocs.Add EstadoFacturaCliente.Aprobada
            F.Show 1

        If IsSomething(Selecciones.Factura) Then
            If DAOFactura.aplicarNCaFC(Selecciones.Factura.id, Factura.id) Then
                MsgBox "Aplicación existosa!", vbInformation, "Información"
            Else
                MsgBox "Se produjo un error, se abortan los cambios!", vbCritical, "Error"
            End If
        End If
    End If
End Sub

Private Sub aprobarFactura_Click()
    On Error GoTo err1
    Dim g As Long

    If MsgBox("¿Desea aprobar el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
        If DAOFactura.aprobar(Factura) Then
            
            
            Dim msg As String
            msg = "Comprobante aprobado con éxito!"
            If IsSomething(Factura.CaeSolicitarResponse) Then
             If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then
            
              msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
            End If
            End If
            MsgBox msg, vbInformation, "Información"
            
            Me.GridEX1.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    'MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical

    MsgBox Err.Description, vbCritical, Err.Source
    Me.GridEX1.RefreshRowIndex g
End Sub

Private Sub archivos_Click()
    Dim F As New frmArchivos2
    F.Origen = 101
    F.ObjetoId = Factura.id
    F.caption = "Comprobante " & Factura.GetShortDescription(False, True)
    F.Show
End Sub

'Private Sub Command1_Click()
'On Error GoTo err1
'Dim col As Collection
'Set col = DAOFactura.FindAll(, True)
'
'Dim F As Factura
'Dim q As String
'conectar.BeginTransaction
'Dim c
'c = 0
'For Each F In col
'    c = c + 1
'    If F.estado = EstadoFacturaCliente.Aprobada Then
'
'    F.TotalEstatico.Total = F.Total
'    F.TotalEstatico.TotalExento = F.TotalExento
'    F.TotalEstatico.TotalIVA = F.TotalIVA
'    F.TotalEstatico.TotalIVADiscrimandoONo = F.TotalIVADiscrimandoONo
'    F.TotalEstatico.TotalNetoGravado = F.TotalNetoGravado
'    F.TotalEstatico.TotalPercepcionesIB = F.totalPercepciones
'
'    'a = DAOFactura.Guardar(f) 'Then GoTo err1
'
'    q = "UPDATE sp.AdminFacturas " _
     '                     & " SET total_estatico = '" & F.Total & "', " _
     '                    & "  total_iva_estatico = ' " & F.TotalIVA & "  ', " _
     '                & "total_perIB_estatico = '" & F.totalPercepciones & "', " _
     '                & " total_neto_estatico = '" & F.TotalNetoGravado & "'," _
     '                & " total_exento_estatico = '" & F.TotalExento & "', " _
     '                & " total_iva_discono_estatico = '" & F.TotalIVADiscrimandoONo & "' " _
     '                & "Where id = " & F.id
'
'                conectar.execute q
'    End If
'Debug.Print c
'Next F
'
'conectar.CommitTransaction
'
'Exit Sub
'err1:
'
'conectar.RollBackTransaction
'
'End Sub





Private Sub cboRangos_Click()
    funciones.CalculateDateRange Me.cboRangos, Me.dtpDesde, Me.dtpHasta
End Sub





Private Sub chkVerObservaciones_Click()
    verObservaciones
End Sub
Private Sub verObservaciones()
    If Me.chkVerObservaciones Then
        Me.GridEX1.PreviewRowLines = 1
    Else
        Me.GridEX1.PreviewRowLines = 0
    End If
End Sub


Private Sub cmdBuscar_Click()
    llenarGrilla
End Sub

Private Sub cmdImprimir_Click()


    With Me.GridEX1.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPLandscape
        .HeaderString(jgexHFCenter) = "Emitidos"
        .FooterString(jgexHFCenter) = Now
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    GridEX1.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub



Private Sub Command1_Click()

    DAODetalleOrdenTrabajo.arreglarCagada
End Sub

Private Sub editar_Click()

    Dim f_c3h3 As New frmFacturaEdicion
    f_c3h3.idFactura = Factura.id
    f_c3h3.Show

End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, True, False
    DAOCliente.llenarComboXtremeSuite Me.cboClientes, False, True, False
    Me.cboClientes.ListIndex = -1

    vId = funciones.CreateGUID
    Channel.AgregarSuscriptor Me, FacturaCliente_
    
'Modificación 15/05/20 (Se muestran todos los comprobanes sin filtrar por punto de venta)
    DAOPuntoVenta.llenarComboXtremeSuite Me.cboPuntosVenta, False
    
    cboEstados.Clear
    cboEstados.AddItem "Pendientes"
    cboEstados.ItemData(cboEstados.NewIndex) = 1
    cboEstados.AddItem "Aprobadas"
    cboEstados.ItemData(cboEstados.NewIndex) = 2
    cboEstados.AddItem "Anuladas"
    cboEstados.ItemData(cboEstados.NewIndex) = 3
    

    Dim i As Integer
    funciones.FillComboBoxDateRanges Me.cboRangos
    For i = 0 To Me.cboRangos.ListCount - 1
        If Me.cboRangos.ItemData(i) = DateRangeValue.DRV_YearCurrent Then Exit For
    Next i
    Me.cboRangos.ListIndex = i
    llenarGrilla
    verObservaciones
End Sub

Private Sub llenarGrilla()
    Dim cliente As clsCliente
    Dim filtro As String
    Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_factura)

    Me.GridEX1.ItemCount = 0
    filtro = "1=1"
    If Me.cboClientes.ListIndex >= 0 Then
        filtro = filtro & " and idCliente=" & cboClientes.ItemData(Me.cboClientes.ListIndex)
    End If

    If Me.cboPuntosVenta.ListIndex >= 0 Then
        filtro = filtro & " and pv.id=" & cboPuntosVenta.ItemData(Me.cboPuntosVenta.ListIndex)
    End If


    If Me.cboEstados.ListIndex >= 0 Then
        filtro = filtro & " and AdminFacturas.estado=" & cboEstados.ItemData(Me.cboEstados.ListIndex)
    End If


    If Me.chkCredito.value > 0 Then
    filtro = filtro & " and AdminFacturas.EsCredito=" & Me.chkCredito.value
   End If
    

    If LenB(Me.txtOrdenCompra) > 0 Then
        filtro = filtro & " and OrdenCompra like '%" & Trim(Me.txtOrdenCompra) & "%'"
    End If
    If LenB(Me.txtNroFactura) > 0 And IsNumeric(Me.txtNroFactura) Then
        filtro = filtro & " and nroFactura=" & Me.txtNroFactura
    End If

    If Not IsNull(Me.dtpDesde.value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision >= " & conectar.Escape(Me.dtpDesde.value)
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        filtro = filtro & " AND AdminFacturas.FechaEmision <= " & conectar.Escape(Me.dtpHasta.value)
    End If

    If LenB(Me.txtRemitoAplicado.text) > 0 Then
        filtro = filtro & " and AdminFacturas.id IN (SELECT fd.idFactura FROM AdminFacturasDetalleNueva fd INNER JOIN entregas e ON e.id = fd.idEntrega INNER JOIN remitos r ON r.id = e.Remito WHERE r.numero = " & Me.txtRemitoAplicado.text & ")"
    End If

    If LenB(Me.txtReferencia.text) > 0 Then
        filtro = filtro & " and AdminFacturas.OrdenCompra like '%" & Trim(Me.txtReferencia.text) & "%'"
    End If

    Set facturas = DAOFactura.FindAll(filtro)
    Dim F As Factura
    Dim c As Integer
    For Each F In facturas


        Dim Total As Double
        Dim totalNG As Double
        Dim TotalIVA As Double
        Dim totalPercepcionesIIBB As Double


        If F.TipoDocumento = tipoDocumentoContable.notaCredito Then c = -1 Else c = 1



        Total = Total + MonedaConverter.ConvertirForzado2(F.TotalEstatico.Total * c, MonedaConverter.Patron.id, F.moneda.id, F.CambioAPatron)

        '    Total = Total + MonedaConverter.ConvertirForzado2(F.TotalEstatico.Total, F.Moneda.Id, MonedaConverter.Patron.Id, F.CambioAPatron)

        TotalIVA = TotalIVA + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalIVA * c, MonedaConverter.Patron.id, F.moneda.id, F.CambioAPatron)
        'TotalIVA = TotalIVA + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalIVA, F.Moneda.Id, MonedaConverter.Patron.Id, F.CambioAPatron)
        totalNG = totalNG + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalNetoGravado * c, MonedaConverter.Patron.id, F.moneda.id, F.CambioAPatron)
        '    totalNG = totalNG + MonedaConverter.ConvertirForzado2(F.TotalEstatico.TotalNetoGravado, F.Moneda.Id, MonedaConverter.Patron.Id, F.CambioAPatron)
        totalPercepcionesIIBB = totalPercepcionesIIBB + MonedaConverter.Convertir(F.TotalEstatico.TotalPercepcionesIB * c, F.moneda.id, MonedaConverter.Patron.id)

    Next


    Me.lblTotal = "Total: $ " & funciones.FormatearDecimales(Total)
    Me.lblTotalPercepciones = "Total Percepciones: $ " & funciones.FormatearDecimales(totalPercepcionesIIBB)
    Me.lblTotalIVA = "Total IVA: $ " & funciones.FormatearDecimales(TotalIVA)
    Me.lblTotalNeto = "Total NG: $ " & funciones.FormatearDecimales(totalNG)





    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = facturas.count
    Me.caption = "Emitidos [Cantidad: " & facturas.count & "]"


' Desabilito la apertura directa de la Factura al encontrar exacto
    'If facturas.count = 1 Then
    '   Dim f_c3h3 As New frmFacturaEdicion
    '    f_c3h3.idFactura = facturas(1).id
    '    f_c3h3.Show
    'End If

    'GridEXHelper.AutoSizeColumns Me.GridEX1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.GridEX1.Width = Me.ScaleWidth
    Me.GridEX1.Height = Me.ScaleHeight - 1900
    Me.grp.Width = Me.GridEX1.Width - 180
End Sub

Private Sub Form_Terminate()
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Channel.RemoverSuscripcionTotal Me
End Sub

Private Sub GridEX1_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)
    GridEX1.PrinterProperties.FooterString(jgexHFRight) = "Página" & PageNumber & " de " & nPages
End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_DblClick()
    verFactura_Click
End Sub

Private Sub GridEX1_FetchIcon(ByVal RowIndex As Long, ByVal ColIndex As Integer, ByVal RowBookmark As Variant, ByVal IconIndex As GridEX20.JSRetInteger)
    If ColIndex = 20 And m_Archivos.item(Factura.id) > 0 Then IconIndex = 1
End Sub

Private Sub GridEX1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If facturas.count > 0 Then
        SeleccionarFactura
        If Button = 2 Then
            Me.NRO.caption = "[ Nro. " & Format(Factura.numero, "0000") & " ]"


            Me.mnuFechaPagoPropuesta.Enabled = False

            If Factura.estado = EstadoFacturaCliente.EnProceso Then
                Me.aplicarNCaFC.Enabled = False
                Me.mnuAplicarANC = False
                Me.editar.Enabled = True
                Me.desAnular.Visible = False
                Me.AnularFactura.Visible = True
                Me.AnularFactura.Enabled = False
                Me.aprobarFactura.Enabled = True
                
                
                Me.aprobarFactura.Visible = True 'Not Factura.EsCredito
                
                Me.mnuAprobarSinEnvio.Enabled = False
                Me.mnuAprobarSinEnvio.Visible = False
                
                Me.mnuEditarCAE.Enabled = False
                Me.mnuEditarCAE.Visible = False
                
                Me.ImprimirFactura.Enabled = False
                Me.mnuDesaprobarFactura.Visible = False
                Me.aplicar.Enabled = False
                Me.mnuFechaPagoPropuesta.Enabled = True
                Me.mnuFechaEntrega.Enabled = False
                Me.aprobarFactura.Enabled = Permisos.AdminFacturasAprobaciones
            
            ElseIf Factura.estado = EstadoFacturaCliente.Aprobada Then    'estado = 2 Then
                Me.editar.Enabled = False
                Me.mnuFechaEntrega.Enabled = True
                Me.desAnular.Visible = False
                Me.mnuDesaprobarFactura.Visible = True
                Me.AnularFactura.Visible = True
                Me.AnularFactura.Enabled = True
                Me.aprobarFactura.Enabled = False
                Me.aprobarFactura.Visible = False
                
                Me.mnuAprobarSinEnvio.Enabled = False
                Me.mnuAprobarSinEnvio.Visible = False
                
                Me.mnuEditarCAE.Enabled = Not Factura.EstaImpresa And Factura.Tipo.PuntoVenta.CaeManual
                Me.mnuEditarCAE.Visible = Not Factura.EstaImpresa And Factura.Tipo.PuntoVenta.CaeManual
                
                Me.ImprimirFactura.Enabled = True
                Me.aplicar.Enabled = (Factura.Saldado = TipoSaldadoFactura.NoSaldada Or Factura.Saldado = TipoSaldadoFactura.saldadoTotal)
                
                
                Me.mnuFechaPagoPropuesta.Enabled = True
                Me.aplicarNCaFC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.notaCredito) And (Factura.estado = EstadoFacturaCliente.Aprobada)
                Me.mnuAplicarANC.Enabled = (Factura.TipoDocumento = tipoDocumentoContable.notaDebito Or Factura.TipoDocumento = tipoDocumentoContable.Factura) And (Factura.estado = EstadoFacturaCliente.Aprobada)

            ElseIf Factura.estado = EstadoFacturaCliente.Anulada Then
                Me.mnuFechaEntrega.Enabled = False
                Me.editar.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
                Me.mnuEditarCAE.Enabled = False
                Me.mnuEditarCAE.Visible = False
                
                Me.mnuAprobarSinEnvio.Enabled = False
                Me.mnuAprobarSinEnvio.Visible = False

                Me.ImprimirFactura.Enabled = False
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                Me.mnuAplicarANC = False
                
                
            ElseIf Factura.estado = EstadoFacturaCliente.CanceladaNC Then
                Me.editar.Enabled = False
                Me.mnuFechaEntrega.Enabled = False
                Me.AnularFactura.Enabled = False
                Me.AnularFactura.Visible = False
                Me.aprobarFactura.Enabled = False
              
              
                Me.mnuAprobarSinEnvio.Enabled = False
                Me.mnuAprobarSinEnvio.Visible = False
                
                Me.mnuEditarCAE.Enabled = False
                Me.mnuEditarCAE.Visible = False
              
                Me.ImprimirFactura.Enabled = True
                Me.aplicar.Enabled = False
                Me.aplicarNCaFC.Enabled = False
                Me.mnuAplicarANC = False
                
                
                
                
                
            End If
            Me.archivos.Enabled = Permisos.SistemaArchivosVer

            If Factura.Saldado <> NoSaldada Then
                Me.mnuFechaEntrega.Enabled = False
            End If


            If Factura.Tipo.PuntoVenta.CaeManual Then
                  
                  Me.aprobarFactura.Enabled = False
                  Me.aprobarFactura.Visible = False
                        
                  Me.mnuEditarCAE.Enabled = False
                  Me.mnuEditarCAE.Visible = False
                        
                  Me.mnuAprobarSinEnvio.Enabled = True
                  Me.mnuAprobarSinEnvio.Visible = True
                  
             End If
             
             If Factura.estado = EstadoFacturaCliente.Aprobada And Factura.Tipo.PuntoVenta.CaeManual Then
                  
                  Me.mnuEditarCAE.Enabled = True
                  Me.mnuEditarCAE.Visible = True
                  
                  Me.mnuAprobarSinEnvio.Enabled = False
                  Me.mnuAprobarSinEnvio.Visible = False
                  
             End If
           
              If Factura.estado = EstadoFacturaCliente.CanceladaNC And Factura.Tipo.PuntoVenta.CaeManual Then
                  
                  Me.mnuEditarCAE.Enabled = True
                  Me.mnuEditarCAE.Visible = True
                  
                  Me.mnuAprobarSinEnvio.Enabled = False
                  Me.mnuAprobarSinEnvio.Visible = False
                  
           End If
            

            Me.AnularFactura.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual
            Me.desAnular.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual
            Me.mnuDesaprobarFactura.Enabled = Not Factura.Tipo.PuntoVenta.EsElectronico


            Me.PopupMenu Me.mnuFacturas
        End If
    End If
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas.item(RowBuffer.RowIndex)
    If Factura.estado = EstadoFacturaCliente.Anulada Then
        RowBuffer.RowStyle = "anulada"
    Else
        If Factura.estado = EstadoFacturaCliente.EnProceso Then
            RowBuffer.CellStyle(12) = "pendiente"
        ElseIf Factura.estado = EstadoFacturaCliente.Aprobada Then
            RowBuffer.CellStyle(12) = "aprobada"
        End If

        If Factura.Saldado = TipoSaldadoFactura.NoSaldada Or Factura.Saldado = TipoSaldadoFactura.SaldadoParcial Or Factura.Saldado = TipoSaldadoFactura.notaCredito Then
            If Factura.EstaAtrasada Then
                RowBuffer.CellStyle(15) = "no_saldada"
            Else
                RowBuffer.CellStyle(15) = "no_vencida"
            End If
        ElseIf Factura.Saldado = saldadoTotal Then
            RowBuffer.CellStyle(15) = "saldada"
        End If


    End If
    Exit Sub
err1:

End Sub

Private Sub GridEX1_SelectionChange()
    SeleccionarFactura
End Sub

Private Sub SeleccionarFactura()
    On Error Resume Next
    Set Factura = facturas.item(Me.GridEX1.RowIndex(Me.GridEX1.row))

End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set Factura = facturas.item(RowIndex)
    Values(1) = Factura.GetShortDescription(True, False)    'enums.EnumTipoDocumentoContable(Factura.TipoDocumento)

    If IsSomething(Factura.Tipo) Then
        Values(2) = Factura.Tipo.TipoFactura.Tipo
    End If

    Values(3) = Factura.Tipo.PuntoVenta.PuntoVenta

    If Factura.esCredito Then
        Values(4) = "(FCE)"
    Else
        Values(4) = ""
    End If


    Values(5) = Factura.NumeroFormateado
    Values(6) = Factura.FechaEmision
    Values(7) = funciones.FormatearDecimales(Factura.TotalEstatico.Total)

    If Factura.moneda.id = 0 Then
        Values(8) = Factura.moneda.NombreCorto
    Else
        Values(8) = Factura.moneda.NombreCorto & " " & Factura.CambioAPatron
    End If

    Values(9) = funciones.FormatearDecimales(Factura.TotalEstatico.Total * Factura.CambioAPatron)

    Values(10) = Factura.OrdenCompra
    Values(11) = Factura.cliente.razon
    Values(12) = enums.EnumEstadoDocumentoContable(Factura.estado)
    Values(13) = EnumTipoSaldadoFactura(Factura.Saldado)
    

    Values(14) = Factura.Vencimiento

    
    Values(15) = Factura.StringDiasAtraso
    Values(16) = Factura.usuarioCreador.usuario
    Values(17) = Factura.observaciones

    If Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual Then
    
    If Factura.estado = EstadoFacturaCliente.EnProceso Then
         Values(17) = "Comprobante en proceso"
    Else
    
        If LenB(Factura.CAE) <= 2 Then
          Values(17) = "/ CAE no definido"
        Else
   
          Values(17) = Values(17) & "/ CAE: " & Factura.CAE
        End If
    End If
End If
    If IsSomething(Factura.UsuarioAprobacion) Then
        Values(18) = Factura.UsuarioAprobacion.usuario
    Else
        Values(18) = vbNullString
    End If
     
    
    If CDbl(Factura.FechaPropuestaPago) > 0 Then Values(19) = Factura.FechaPropuestaPago

    If Factura.DiferenciaDiasEntrega = -1 Then
        Values(19) = "Defina fecha"
    Else

        If CDbl(Factura.FechaEntrega) > 0 And Factura.estado <> EstadoFacturaCliente.Anulada Then
            If Factura.Saldado = NoSaldada Then Values(20) = Format(Factura.FechaEntrega, "dd/mm/yyyy") & " (" & Factura.DiferenciaDiasEntrega & ")"



            '       If Factura.Saldado = SaldadoTotal Then
            '      Values(18) = "Saldada"
            '    Else
            '           Values(18) = "Anulada"
            '    End If
            '        Values(18) = Factura.FechaEntrega
            '    End If

        Else
            If Factura.estado = EstadoFacturaCliente.Anulada Then
                Values(20) = "Anulada"
            Else
                Values(20) = Factura.FechaEntrega
            End If
        End If

    End If


    Values(21) = Factura.TasaAjusteMensual

    Values(22) = "(" & Val(m_Archivos.item(Factura.id)) & ")"

    Exit Sub
err1:
End Sub

Private Sub ImprimirFactura_Click()

    On Error GoTo err451:
    Dim clasea As New classAdministracion
    Dim veces As Long


    If Factura.Tipo.PuntoVenta.EsElectronico Or Factura.Tipo.PuntoVenta.CaeManual Then
        veces = clasea.facturaImpresa(Factura.id)
        If veces > 0 Then
            If MsgBox("Este comprobante ya fué generarlo" & Chr(10) & "¿Desea volver a generarlo?", vbYesNo, "Confirmación") = vbYes Then
                'DAOFactura.GenerarPdf (Factura.id)
                DAOFactura.VerFacturaElectronicaParaImpresion (Factura.id)
            End If
        Else


            DAOFactura.VerFacturaElectronicaParaImpresion (Factura.id)


        End If
    Else

        veces = clasea.facturaImpresa(Factura.id)
        If veces = 0 Or veces = -1 Then
            If MsgBox("¿Desea imprimir este comprobante?", vbYesNo, "Confirmación") = vbYes Then
               cd.Flags = cdlPDUseDevModeCopies
                cd.Copies = 3
                cd.ShowPrinter
                Dim i As Long
                For i = 1 To cd.Copies
                    DAOFactura.Imprimir Factura.id
                Next
            End If

        ElseIf veces > 0 Then
            If MsgBox("Este comprobante ya fué impreso." & Chr(10) & "¿Desea volver a imprimirlo?", vbYesNo, "Confirmación") = vbYes Then
                cd.Flags = cdlPDUseDevModeCopies
                cd.Copies = 3
                cd.ShowPrinter

                For i = 1 To cd.Copies
                    DAOFactura.Imprimir Factura.id
                Next i
            End If

        End If
    End If
    Exit Sub
err451:

End Sub

Private Property Get ISuscriber_id() As String
    ISuscriber_id = vId
End Property

Private Function ISuscriber_Notificarse(EVENTO As clsEventoObserver) As Variant
    Dim tmp As Factura
    If EVENTO.EVENTO = agregar_ Then
        llenarGrilla
        Me.GridEX1.Refresh
    ElseIf EVENTO.EVENTO = modificar_ Then
        Set tmp = EVENTO.Elemento

        Dim i As Long
        For i = facturas.count To 1 Step -1

            If facturas(i).id = tmp.id Then

                '
                '                Set Factura = facturas(i)
                '
                '                Factura.Id = tmp.Id
                '
                '
                '                Factura.Detalles = tmp.Detalles
                '                Factura.estado = tmp.estado
                '                Factura.OrdenCompra = tmp.OrdenCompra
                '                Factura.estado = tmp.estado
                '                Factura.Observaciones = tmp.Observaciones
                '                Factura.TasaAjusteMensual = tmp.TasaAjusteMensual
                '
                '
                '                Set Factura.Cliente = tmp.Cliente

                facturas.remove i
                If facturas.count > 0 Then
                    If i = 1 Then    'ver esto cuand oes un solo item
                        facturas.Add tmp, CStr(tmp.id), 1
                    ElseIf (i - 1) = facturas.count Then
                        facturas.Add tmp, CStr(tmp.id), , i - 1
                    Else
                        facturas.Add tmp, CStr(tmp.id), i
                    End If
                Else
                    facturas.Add tmp, CStr(tmp.id)
                End If

                'DAOFactura.FindById(tmp.Id, True)



                Me.GridEX1.RefreshRowIndex i
                Exit For

            End If

        Next

    End If


End Function

Private Sub mnuAplicarANC_Click()
  If MsgBox("¿Seguro de aplicar a NC?", vbYesNo, "Confirmación") = vbYes Then
        'seleccionar factura para aplicar
        Set Selecciones.Factura = Nothing
          Dim F As New frmAdminFacturasNCElegirFC
        
        F.idCliente = Factura.cliente.id
            F.TiposDocs.Add tipoDocumentoContable.notaCredito
            F.EstadosDocs.Add EstadoFacturaCliente.Aprobada
            F.Show 1

        If IsSomething(Selecciones.Factura) Then
            If DAOFactura.aplicarNCaFC(Factura.id, Selecciones.Factura.id) Then
                MsgBox "Aplicación existosa!", vbInformation, "Información"
            Else
                MsgBox "Se produjo un error, se abortan los cambios!", vbCritical, "Error"
            End If
        End If
    End If


End Sub

Private Sub mnuAprobarSinEnvio_Click()

On Error GoTo err1
    Dim g As Long

    If MsgBox("¿Desea aprobar el comprobante SIN ENVÍAR A LA AFIP?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
        
        If DAOFactura.aprobar(Factura, False) Then
            
            
              MsgBox "Recuerde agregar al comprobante: CAE y fecha de vencimiento del CAE ", vbInformation, "Información"
            
            
'            Dim msg As String
'            msg = "Comprobante aprobado con éxito!"
'            If IsSomething(Factura.CaeSolicitarResponse) Then
'             If LenB(Factura.CaeSolicitarResponse.observaciones) > 5 Then
'
'              msg = msg & Chr(10) & Factura.CaeSolicitarResponse.observaciones
'            End If
'            End If
'            MsgBox msg, vbInformation, "Información"
            
            Me.GridEX1.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    'MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical

    MsgBox Err.Description, vbCritical, Err.Source
    Me.GridEX1.RefreshRowIndex g



End Sub

Private Sub mnuArchivos_Click()
    Dim archi As New frmArchivos2

    archi.Origen = OrigenArchivos.OA_factura
    archi.ObjetoId = Factura.id
    archi.caption = Factura.GetShortDescription(False, True)
    archi.Show

End Sub

Private Sub mnuCrearCopiaFactura_Click()
    Me.taskDialog.Reset
    Me.taskDialog.MessageBoxStyle = True
    Me.taskDialog.WindowTitle = "Copia fiel de Comprobante"
    Me.taskDialog.MainInstructionText = "¿De que tipo es el nuevo comprobante?"
    Me.taskDialog.ContentText = "Elija el tipo de comprobante para el nuevo comprobante."
    taskDialog.RelativePosition = False

    Me.taskDialog.CommonButtons = 0
    taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonOk
    taskDialog.CommonButtons = taskDialog.CommonButtons Or xtpTaskButtonCancel

    taskDialog.DefaultRadioButton = -1
    taskDialog.AddRadioButton "Factura", tipoDocumentoContable.Factura
    taskDialog.AddRadioButton "Nota de Débito", tipoDocumentoContable.notaDebito
    taskDialog.AddRadioButton "Nota de Crédito", tipoDocumentoContable.notaCredito


    taskDialog.MainIcon = xtpTaskIconInformation

    If taskDialog.ShowDialog = xtpTaskButtonOk Then
        If Me.taskDialog.DefaultRadioButton = -1 Then
            MsgBox "Debe seleccionar un tipo para el nuevo comprobante.", vbExclamation + vbOKOnly
        Else
            Dim newFact As Factura
            Set newFact = DAOFactura.CrearCopiaFiel(Factura, Me.taskDialog.DefaultRadioButton)
            If IsSomething(newFact) Then
                MsgBox "Se creó un nuevo comprobante (" & newFact.GetShortDescription(False, True) & ")", vbInformation + vbOKOnly
            Else
                MsgBox "Hubo un error al copiar la factura.", vbCritical + vbOKOnly
            End If
        End If
    End If



End Sub

Private Sub mnuDesaprobarFactura_Click()

    On Error GoTo err1
    Dim g As Long

    If MsgBox("¿Desea desaprobar el comprobante?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        g = Me.GridEX1.RowIndex(Me.GridEX1.row)
        If DAOFactura.desaprobar(Factura) Then
            MsgBox "Comprobante desaprobado con éxito!", vbInformation, "Información"
            Me.GridEX1.RefreshRowIndex g
            Me.txtNroFactura.SetFocus
        Else
            GoTo err1
        End If
    End If
    Exit Sub
err1:
    MsgBox "Factura no aprobada, compruebe:" & vbNewLine & "Si la factura es de anticipo, compruebe que el valor de la misma sea el mismo que el anticipo de la OT." & vbNewLine & "Que el detalle del remito no este ya facturado." & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub mnuEditarCAE_Click()
    Dim g As Long
    g = Me.GridEX1.RowIndex(Me.GridEX1.row)

    Dim F As New frmAdminFacturasAprobarSinAfip
    Set F.Factura = Factura
    F.Show 1

 Me.GridEX1.RefreshRowIndex g


End Sub

Private Sub mnuFechaEntrega_Click()
    Dim fechaAnterior As String
    Dim fechaPosterior As String
    Dim nuevaFecha As Date
    Dim Update As Boolean

    If CDbl(Factura.FechaEntrega) > 0 Then fechaAnterior = Factura.FechaEntrega

    fechaPosterior = InputBox("Establezca fecha de entrega", "Fecha de Entrega", fechaAnterior)

    If LenB(fechaPosterior) = 0 Then
        nuevaFecha = 1 / 1 / 2005
        Update = True
    Else
        If IsDate(fechaPosterior) Then
            nuevaFecha = CDate(fechaPosterior)
            Update = True
        Else
            MsgBox "La fecha no es válida.", vbOKOnly + vbExclamation, "Fecha"
        End If
    End If

    If Update Then
        Factura.FechaEntrega = nuevaFecha
        If DAOFactura.Guardar(Factura) Then
            Me.GridEX1.RefreshRowIndex (Me.GridEX1.row)
        Else
            MsgBox "Error al guardar la factura.", vbOKOnly + vbCritical, "Error"
        End If
    End If

End Sub

Private Sub mnuFechaPagoPropuesta_Click()
    Dim fechaAnterior As String
    Dim fechaPosterior As String
    Dim nuevaFecha As Date
    Dim Update As Boolean

    If CDbl(Factura.FechaPropuestaPago) > 0 Then fechaAnterior = Factura.FechaPropuestaPago

    fechaPosterior = InputBox("Establezca fecha de pago propuesta", "Fecha de Pago", fechaAnterior)

    If LenB(fechaPosterior) = 0 Then
        Update = (MsgBox("¿Desea dejar en blanco la fecha de pago propuesta?", vbYesNo + vbQuestion) = vbYes)
    Else
        If IsDate(fechaPosterior) Then
            nuevaFecha = CDate(fechaPosterior)
            Update = True
        Else
            MsgBox "La fecha no es válida.", vbOKOnly + vbExclamation, "Fecha"
        End If
    End If

    If Update Then
        Factura.FechaPropuestaPago = nuevaFecha
        If DAOFactura.Guardar(Factura) Then
            Me.GridEX1.ReBind
        Else
            MsgBox "Error al guardar la factura.", vbOKOnly + vbCritical, "Error"
        End If
    End If

End Sub

Private Sub PushButton1_Click()
    Me.cboClientes.ListIndex = -1
End Sub

Private Sub PushButton2_Click()
    Dim id As Long
    If (Me.cboClientes.ListIndex > 0) Then
        id = Me.cboClientes.ItemData(Me.cboClientes.ListIndex)
    Else
        id = -1
    End If

    Dim col As New Collection

    If (id > 0) Then
        Set col = DAOFactura.FindAllByEstadoSaldoAndCliente(NoSaldada, EstadoFacturaCliente.Aprobada, id)
    Else
        Set col = DAOFactura.FindAllByEstadoSaldoAndCliente(NoSaldada, EstadoFacturaCliente.Aprobada)

    End If


    Dim xlWorkbook As New Excel.Workbook
    Dim xlWorksheet As New Excel.Worksheet
    Dim xlApplication As New Excel.Application

    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Cells(1, 1).value = "Cliente"

    If (id > 0) Then
        xlWorksheet.Cells(1, 2).value = DAOCliente.BuscarPorID(id).razon
    Else
        xlWorksheet.Cells(1, 2).value = "Todos"
    End If

    xlWorksheet.Columns(4).HorizontalAlignment = xlLeft
    xlWorksheet.Columns(10).HorizontalAlignment = xlLeft
    xlWorksheet.Cells(2, 1).value = "Comprobante"
    xlWorksheet.Cells(2, 2).value = "Emision"
    xlWorksheet.Cells(2, 3).value = "Moneda"
    xlWorksheet.Cells(2, 4).value = "Detalle"
    xlWorksheet.Cells(2, 5).value = "Importe en " & DAOMoneda.FindFirstByPatronOrDefault.NombreCorto

    xlWorksheet.Cells(2, 6).value = "Vencimiento"
    xlWorksheet.Cells(2, 7).value = "Atraso"
    xlWorksheet.Cells(2, 8).value = "Entrega"
    xlWorksheet.Cells(2, 9).value = "Atraso"
    If (id < 0) Then xlWorksheet.Cells(2, 10).value = "Cliente"
    Dim idx As Integer
    idx = 3
    Dim fac As Factura
    For Each fac In col

        xlWorksheet.Cells(idx, 1).value = fac.GetShortDescription(False, True)

        xlWorksheet.Cells(idx, 2).value = fac.FechaEmision
        xlWorksheet.Cells(idx, 3).value = fac.moneda.NombreCorto
        xlWorksheet.Cells(idx, 4).value = fac.OrdenCompra


        If fac.TipoDocumento = tipoDocumentoContable.notaCredito Then
            xlWorksheet.Cells(idx, 5).value = funciones.RedondearDecimales(fac.TotalEstatico.Total * fac.CambioAPatron) * -1
        Else
            xlWorksheet.Cells(idx, 5).value = funciones.RedondearDecimales(fac.TotalEstatico.Total * fac.CambioAPatron)


        End If
        xlWorksheet.Cells(idx, 6).value = fac.Vencimiento
        xlWorksheet.Cells(idx, 7).value = fac.StringDiasAtraso
        If (fac.DiferenciaDiasEntrega <> -1) Then
            xlWorksheet.Cells(idx, 8).value = Format(fac.FechaEntrega, "dd/mm/yyyy")
            xlWorksheet.Cells(idx, 9).value = fac.DiferenciaDiasEntrega & " dias"
        Else
            xlWorksheet.Cells(idx, 8).value = "no definida"
            xlWorksheet.Cells(idx, 9).value = 0
        End If

        If (id < 0) Then xlWorksheet.Cells(idx, 10).value = fac.cliente.razon

        idx = idx + 1
    Next
    xlWorksheet.Cells(idx, 5).Formula = "=SUM(E3:E" & idx - 1 & ")"

    'autosize
    xlApplication.ScreenUpdating = False
    Dim wkSt As String
    wkSt = xlWorksheet.Name
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlWorkbook.Sheets(wkSt).Select
    xlApplication.ScreenUpdating = True

    'xlWorksheet.PageSetup.PrintTitleRows = "$1:$3" 'para que al imprimir queden las columnas fijas
    xlWorksheet.PageSetup.Orientation = xlLandscape
    xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)

    Dim filename As String
    filename = funciones.GetTmpPath() & "tmp_info " & Hour(Now) & Minute(Now) & Second(Now) & " .xls"

    If Dir(filename) <> vbNullString Then Kill filename

    xlWorkbook.SaveAs filename

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit


    funciones.ShellExecute 0, "open", filename, "", "", 0

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

End Sub

Private Sub PushButton3_Click()
    Me.cboPuntosVenta.ListIndex = -1
End Sub

Private Sub PushButton4_Click()
    Me.cboEstados.ListIndex = -1
End Sub



Private Sub scanear_Click()
    On Error Resume Next
    Dim archivos As New classArchivos
    If archivos.escanearDocumento(OrigenArchivos.OA_factura, Factura.id) Then
        Set m_Archivos = DAOArchivo.GetCantidadArchivosPorReferencia(OA_factura)
        Me.GridEX1.RefreshRowIndex (Factura.id)
    End If
End Sub

Private Sub txtOrdenCompra_GotFocus()
    foco Me.txtOrdenCompra
End Sub


Private Sub verFactura_Click()
    Dim f_c3h3 As New frmFacturaEdicion
    f_c3h3.ReadOnly = True
    f_c3h3.idFactura = Factura.id
    f_c3h3.Show

End Sub

Private Sub verHistorialFactura_Click()
    Set Factura.Historial = DAOFacturaHistorial.getAllByIdFactura(Factura.id)
    frmHistoriales.lista = Factura.Historial
    frmHistoriales.Show
End Sub
