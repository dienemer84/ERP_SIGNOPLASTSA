VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAdminFacturasVer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerFactura.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9225
   Begin VB.TextBox txtObs 
      BackColor       =   &H00F0E1D1&
      Height          =   585
      Left            =   6045
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   1560
      Width           =   2910
   End
   Begin VB.Frame fraDetalles 
      Caption         =   "Detalles"
      Height          =   3720
      Left            =   180
      TabIndex        =   22
      Top             =   2310
      Width           =   8865
      Begin GridEX20.GridEX gridDetalles 
         Height          =   3315
         Left            =   135
         TabIndex        =   23
         Top             =   270
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   5847
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         PreviewColumn   =   "obs"
         MethodHoldFields=   -1  'True
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   9
         Column(1)       =   "frmVerFactura.frx":000C
         Column(2)       =   "frmVerFactura.frx":0120
         Column(3)       =   "frmVerFactura.frx":0234
         Column(4)       =   "frmVerFactura.frx":0328
         Column(5)       =   "frmVerFactura.frx":044C
         Column(6)       =   "frmVerFactura.frx":0568
         Column(7)       =   "frmVerFactura.frx":0688
         Column(8)       =   "frmVerFactura.frx":07A8
         Column(9)       =   "frmVerFactura.frx":089C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmVerFactura.frx":0990
         FormatStyle(2)  =   "frmVerFactura.frx":0AB8
         FormatStyle(3)  =   "frmVerFactura.frx":0B68
         FormatStyle(4)  =   "frmVerFactura.frx":0C1C
         FormatStyle(5)  =   "frmVerFactura.frx":0CF4
         FormatStyle(6)  =   "frmVerFactura.frx":0DAC
         ImageCount      =   0
         PrinterProperties=   "frmVerFactura.frx":0E8C
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copiar FC a Concepto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1125
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   8880
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6405
      TabIndex        =   2
      Top             =   6105
      Width           =   2640
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "999.999.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1410
         TabIndex        =   10
         Top             =   1095
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   9
         Top             =   1095
         Width           =   405
      End
      Begin VB.Label lblIva2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   930
         TabIndex        =   8
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subtotal "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblSubTotal1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "999.999.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   6
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Percepciones "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   495
         Width           =   1020
      End
      Begin VB.Label lblPercepciones 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "999.999.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   4
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label lblIVA4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "999.999.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   3
         Top             =   750
         Width           =   1080
      End
   End
   Begin VB.Label lblFP 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Forma Pago "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6045
      TabIndex        =   11
      Top             =   735
      Width           =   900
   End
   Begin VB.Label lblMoneda 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6045
      TabIndex        =   25
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label lblCiudad 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ciudad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   21
      Top             =   1650
      Width           =   495
   End
   Begin VB.Label lblDireccion 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   20
      Top             =   1035
      Width           =   675
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   19
      Top             =   435
      Width           =   480
   End
   Begin VB.Label lblCuit 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "CUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   18
      Top             =   735
      Width           =   375
   End
   Begin VB.Label lblTipoFactura 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo - Nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   17
      Top             =   135
      Width           =   630
   End
   Begin VB.Label lblOC 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "O/C "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6045
      TabIndex        =   16
      Top             =   435
      Width           =   345
   End
   Begin VB.Label lblObserva 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6045
      TabIndex        =   15
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label lblCp 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "C.P."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   14
      Top             =   1950
      Width           =   300
   End
   Begin VB.Label lblLocalidad 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   13
      Top             =   1350
      Width           =   690
   End
   Begin VB.Label lblIVA 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6045
      TabIndex        =   12
      Top             =   135
      Width           =   255
   End
   Begin VB.Menu aplicar 
      Caption         =   "aplicar"
      Visible         =   0   'False
      Begin VB.Menu aplicarRemito 
         Caption         =   "Aplicar a remito..."
      End
   End
End
Attribute VB_Name = "frmAdminFacturasVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Factura As Factura
Private detalle As FacturaDetalle

Private nroItem As Long

Public Property Let idFactura(nIdFactura As Long)
   Set Factura = DAOFactura.FindById(nIdFactura, True, True)
End Property

Private Sub aplicarRemito_Click()
    Dim idRto As Long
    Dim idEntrega As Long
    
    frmPlaneamientoRemitosListaProceso.idCliMostrar = -1
    frmPlaneamientoRemitosListaProceso.Mostrar = 2  'ente 'CLng(Me.cboClientes.ItemData(Me.cboClientes.ListIndex))
    frmPlaneamientoRemitosListaProceso.Show 1
    If funciones.queRemitoElegido <> -1 Then
        idRto = funciones.queRemitoElegido
    Else
        Exit Sub
    End If

    frmPlaneamientoRemitosDetalle.usable = False
    frmPlaneamientoRemitosDetalle.Editable = False
    frmPlaneamientoRemitosDetalle.usarItem = False
    frmPlaneamientoRemitosDetalle.usarItemFactura = True
    frmPlaneamientoRemitosDetalle.rtoNro = idRto
    frmPlaneamientoRemitosDetalle.Show 1
    idEntrega = funciones.itemRemito
    
    If idEntrega <> -1 Then
        Dim detaRto As RemitoDetalle
        Set detaRto = DAORemitoSDetalle.FindById(idEntrega)
        
        If IsSomething(detaRto) Then
            Dim claseA As New classAdministracion 'sacar esto
        
            If claseA.aplicarEntregaAFactura(detaRto.id, detalle.id, idRto) Then
                MsgBox "Aplicación exitosa!", vbInformation, "Información"
                'Me.lstFactura.SelectedItem.ListSubItems(3).Tag = 1
            Else
                MsgBox "Error al aplicar la entrega!", vbCritical, "Error"
            End If
        End If
    End If


End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim idfo As Long

    Dim claseA As New classAdministracion

    Dim idfd As Long
    idfd = 1460
    idfo = Factura.id 'vIdFactura

    If claseA.copiarFacturaAConcepto(idfo, idfd) Then
        MsgBox "Copia Exitosa!"
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    GridEXHelper.customizeGrid Me.gridDetalles
    llenarDatos
End Sub

Private Sub llenarDatos()
    Me.caption = "Factura  Nº " & Format(Factura.Numero, "0000")

    '        Discrimina = rs!Discriminada  'discriminaIVA
    '        Descuento = rs!Descuento
    '        Alicuota = rs!AlicuotaAplicada

    Me.lblCiudad = "Ciudad: " & Factura.Cliente.Ciudad
    Me.lblDireccion = "Direccion: " & Factura.Cliente.Domicilio
    Me.lblCp = "Cod Postal: " & Factura.Cliente.CP
    Me.lblCuit = "CUIT: " & Factura.Cliente.Cuit
    Me.lblMoneda.caption = "Moneda: " & Factura.Moneda.NombreCorto

    If IsSomething(Factura.Cliente.tipoIva) Then
        Me.lblIVA = "IVA: " & Factura.Cliente.tipoIva.detalle
    Else
        Me.lblIVA = "IVA: "
    End If

    Me.lblCliente = "Cliente: " & Factura.Cliente.Razon
    Me.lblLocalidad = "Localidad: " & Factura.Cliente.localidad
    Me.lblTipoFactura = "Tipo - Nº: " & Factura.Tipo.Tipo & " - " & Factura.NumeroFormateado
    Me.lblOC = "O/C: " & Replace(Factura.OrdenCompra, vbNewLine, vbNullString)
    Me.lblFP = "Forma Pago: " & Factura.CantDiasPago & " Días FF"
    Me.txtObs = Factura.Observaciones

    nroItem = 0
    Me.gridDetalles.ItemCount = 0
    Me.gridDetalles.ItemCount = Factura.Detalles.count
    
    GridEXHelper.AutoSizeColumns Me.gridDetalles


    
    Me.lblIva2.caption = "IVA " & Factura.AlicuotaAplicada & "%"
    Me.lblIVA4.caption = funciones.formatearDecimales(Factura.TotalIva)
    Me.lblSubTotal1.caption = funciones.formatearDecimales(Factura.TotalSubTotal)
    Me.lblPercepciones.caption = funciones.formatearDecimales(Factura.totalPercepciones)

    Me.lblTotal.caption = Factura.Moneda.NombreCorto & " " & funciones.formatearDecimales(Factura.Total)

End Sub




Private Sub gridDetalles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.aplicarRemito.Enabled = Not detalle.AplicadoARemito
        If detalle.DetalleRemitoId = -1 And detalle.Bruto >= 0 Then    'si es concepto se puede aplicar y no es crédito
            Me.PopupMenu Me.aplicar
        End If
    End If
End Sub

Private Sub gridDetalles_SelectionChange()
If Me.gridDetalles.Row > 0 Then
    Set detalle = Factura.Detalles.item(Me.gridDetalles.RowIndex(Me.gridDetalles.Row))
Else
    Set detalle = Nothing
End If
End Sub

Private Sub gridDetalles_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And Factura.Detalles.count > 0 Then
    
        
        Set detalle = Factura.Detalles(RowIndex)
        Values(1) = detalle.Observacion
        Values(2) = Format(detalle.cantidad, "0.00")
        Values(3) = detalle.detalle
        Values(4) = funciones.formatearDecimales(detalle.Bruto, 2)
        Values(5) = funciones.formatearDecimales(detalle.SubTotal)  '.Bruto * detalle.Cantidad, 2)
        Values(6) = funciones.formatearDecimales(detalle.PorcentajeDescuento)
'
        Values(7) = funciones.formatearDecimales(detalle.Total)
       
        If IsSomething(detalle.DetalleRemito) Then
            Values(8) = detalle.DetalleRemito.VerOrigen
            Values(9) = detalle.DetalleRemito.remito
        End If
    End If
End Sub

