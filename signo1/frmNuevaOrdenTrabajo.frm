VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmNuevaOrdenTrabajo 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Orden de Trabajo"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNuevaOrdenTrabajo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5400
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   315
      Left            =   1545
      TabIndex        =   12
      Top             =   540
      Width           =   3735
      _Version        =   786432
      _ExtentX        =   6588
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.GroupBox grpMarco 
      Height          =   900
      Left            =   285
      TabIndex        =   7
      Top             =   2670
      Width           =   4875
      _Version        =   786432
      _ExtentX        =   8599
      _ExtentY        =   1587
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkMarco 
         Height          =   270
         Left            =   165
         TabIndex        =   8
         Top             =   -15
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   476
         _StockProps     =   79
         Caption         =   "Forma Parte de Contrato Abierto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboContratos 
         Height          =   315
         Left            =   255
         TabIndex        =   9
         Top             =   330
         Width           =   4395
         _Version        =   786432
         _ExtentX        =   7752
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
   End
   Begin VB.ComboBox cboMoneda 
      Height          =   315
      Left            =   1545
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2295
      Width           =   840
   End
   Begin VB.TextBox txtReferencia 
      Height          =   630
      Left            =   1545
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1635
      Width           =   3705
   End
   Begin MSComCtl2.DTPicker dtpFechaEntrega 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1260
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   63373313
      CurrentDate     =   40077
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   3705
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdCerrar 
      Height          =   495
      Left            =   2775
      TabIndex        =   11
      Top             =   3705
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboClienteFacturar 
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      Top             =   915
      Width           =   3735
      _Version        =   786432
      _ExtentX        =   6588
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipoOt 
      Height          =   315
      Left            =   1545
      TabIndex        =   15
      Top             =   120
      Width           =   3735
      _Version        =   786432
      _ExtentX        =   6588
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Text            =   "Seleccione"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   16
      Top             =   180
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   855
      TabIndex        =   13
      Top             =   975
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   5
      Top             =   2370
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fecha Entrega"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   1335
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   495
      TabIndex        =   1
      Top             =   1785
      Width           =   915
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Centro de Costos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   615
      Width           =   1440
   End
End
Attribute VB_Name = "frmNuevaOrdenTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clientes As New Collection
Private Monedas As New Collection
Private otsMarco As New Collection
Private marco_ot As OrdenTrabajo

Private TipoOt As TipoOt


Private Sub cboCliente_Click()
    Me.cboMoneda.ListIndex = PosIndexCbo(DAOCliente.BuscarPorID(Me.cboCliente.ItemData(Me.cboCliente.ListIndex)).idMonedaDefault, Me.cboMoneda)
End Sub


Private Sub cboContratos_Click()
    If Me.cboContratos.ListIndex <> -1 Then
        Set marco_ot = otsMarco.item(CStr(Me.cboContratos.ItemData(Me.cboContratos.ListIndex)))
        Me.cboCliente.ListIndex = funciones.PosIndexCbo(marco_ot.cliente.Id, Me.cboCliente)
        Me.cboClienteFacturar.ListIndex = funciones.PosIndexCbo(marco_ot.ClienteFacturar.Id, Me.cboCliente)
        Me.cboMoneda.ListIndex = funciones.PosIndexCbo(marco_ot.moneda.Id, Me.cboMoneda)

        Dim proxfecha As Date
        proxfecha = marco_ot.ProximaFechaActualizacionPrecios
        If CDbl(proxfecha) > 0 Then
            MsgBox "La proxima fecha de vencimiento de precios es " & proxfecha, vbOKOnly + vbInformation
        End If
    End If
End Sub

Private Sub cboTipoOt_Change()



    Me.grpMarco.Enabled = (TipoOt = OT_TRADICIONAL)


End Sub



Private Sub cboTipoOt_Click()

    Me.grpMarco.Enabled = (TipoOt = OT_TRADICIONAL)

End Sub

Private Sub chkMarco_Click()
    Me.cboCliente.Enabled = Not (Me.chkMarco.value * -1)

    'Me.cboClienteFacturar.Enabled = Not (Me.chkMarco.value * -1)

    'Me.cboCliente.ListIndex = -1
    Me.txtReferencia.text = vbNullString
    Me.cboMoneda.Enabled = Not (Me.chkMarco.value * -1)
    Me.cboMoneda.ListIndex = -1

    Me.cboContratos.ListIndex = -1
    Me.cboContratos.Enabled = Me.chkMarco.value * -1
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    If Me.cboCliente.ListIndex = -1 Or Me.cboMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar cliente y moneda.", vbInformation
        Exit Sub
    End If

    If Me.chkMarco.value = xtpChecked And Me.cboContratos.ListIndex = -1 Then
        MsgBox "Debe seleccionar un contrato marco.", vbInformation
        Exit Sub
    End If

    If Me.cboMoneda.ListIndex = 3 Then
        MsgBox "No se puede cargar una OT con Moneda U$A Administrativo. Modifiquelo por favor.", vbInformation
        Exit Sub
    End If


    Dim Ot As New OrdenTrabajo

    Ot.TipoOrden = Me.cboTipoOt.ListIndex + 1

    Set Ot.cliente = DAOCliente.BuscarPorID(Me.cboCliente.ItemData(Me.cboCliente.ListIndex))
    Set Ot.ClienteFacturar = DAOCliente.BuscarPorID(Me.cboClienteFacturar.ItemData(Me.cboClienteFacturar.ListIndex))

    Ot.FechaEntrega = Me.dtpFechaEntrega.value
    Ot.descripcion = Me.txtReferencia.text


    Set Ot.moneda = Monedas.item(CStr(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex)))


    If Me.chkMarco.value = xtpChecked And Me.cboContratos.ListIndex <> -1 Then
        Ot.OTMarcoIdPadre = Me.cboContratos.ItemData(Me.cboContratos.ListIndex)

        Dim otPadre As OrdenTrabajo
        Set otPadre = DAOOrdenTrabajo.FindById(Ot.OTMarcoIdPadre)

        Ot.Anticipo = otPadre.Anticipo
        Ot.AnticipoFacturado = otPadre.AnticipoFacturado
        Ot.CantDiasAnticipo = otPadre.CantDiasAnticipo
        Ot.CantDiasSaldo = otPadre.CantDiasSaldo
        Ot.Descuento = otPadre.Descuento
        Ot.FormaDePagoAnticipo = otPadre.FormaDePagoAnticipo
        Ot.FormaDePagoSaldo = otPadre.FormaDePagoSaldo
        Ot.MismaFechaEntregaParaDetalles = otPadre.MismaFechaEntregaParaDetalles
        Set Ot.ClienteFacturar = otPadre.ClienteFacturar
        Set Ot.cliente = otPadre.cliente

    End If


    If DAOOrdenTrabajo.Save(Ot) Then
        MsgBox "La orden de trabajo se guardo con el número " & Ot.Id, vbInformation + vbOKOnly
        Dim EVENTO As New clsEventoObserver
        Set EVENTO.Elemento = Ot
        EVENTO.EVENTO = agregar_
        Set EVENTO.Originador = Me
        Channel.Notificar EVENTO, ordenesTrabajo

        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
    DAOCliente.llenarComboXtremeSuite Me.cboCliente, True, True, True
    DAOCliente.llenarComboXtremeSuite Me.cboClienteFacturar, True, True, True
    Me.dtpFechaEntrega = Now
    TipoOt = OT_TRADICIONAL

    Set Monedas = DAOMoneda.GetAll()
    Dim mon As clsMoneda

    Me.cboTipoOt.AddItem "Tradicional"
    Me.cboTipoOt.ItemData(Me.cboTipoOt.NewIndex) = 1

    Me.cboTipoOt.AddItem "De Stock"
    Me.cboTipoOt.ItemData(Me.cboTipoOt.NewIndex) = 2
    Me.cboTipoOt.AddItem "De Entrega"
    Me.cboTipoOt.ItemData(Me.cboTipoOt.NewIndex) = 3

    Me.cboTipoOt.ListIndex = 0


    Set otsMarco = DAOOrdenTrabajo.FindAll("p.id_ot_padre = -1 and p.estado = " & EstadoOrdenTrabajo.EstadoOT_EnProceso)
    Dim ot1 As OrdenTrabajo
    Me.cboContratos.Clear
    For Each ot1 In otsMarco
        Me.cboContratos.AddItem ot1.IdFormateado & " - " & ot1.descripcion & " (" & ot1.FechaInicioMarco & " - " & ot1.FechaFinMarco & ")"
        Me.cboContratos.ItemData(Me.cboContratos.NewIndex) = ot1.Id
    Next ot1
    Me.cboContratos.ListIndex = -1

    Me.cboMoneda.Clear
    For Each mon In Monedas
        Me.cboMoneda.AddItem mon.NombreCorto
        Me.cboMoneda.ItemData(Me.cboMoneda.NewIndex) = mon.Id
    Next mon
    If Monedas.count > 0 Then cboMoneda.ListIndex = 0
End Sub
