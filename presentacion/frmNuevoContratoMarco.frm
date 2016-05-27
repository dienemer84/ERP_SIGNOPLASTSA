VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmNuevoContratoMarco 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Contrato Abierto"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNuevoContratoMarco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5385
   Begin VB.TextBox txtMontoTotal 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1425
      TabIndex        =   14
      Text            =   "0"
      Top             =   2175
      Width           =   1260
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   495
      Left            =   3450
      TabIndex        =   10
      Top             =   3435
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox grpFechas 
      Height          =   2145
      Left            =   405
      TabIndex        =   8
      Top             =   2895
      Width           =   2745
      _Version        =   786432
      _ExtentX        =   4842
      _ExtentY        =   3784
      _StockProps     =   79
      Caption         =   "Fechas de Actualizacion de Precios"
      UseVisualStyle  =   -1  'True
      Begin GridEX20.GridEX gridFechasPrecios 
         Height          =   1755
         Left            =   195
         TabIndex        =   9
         Top             =   240
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   3096
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
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
         ColumnsCount    =   1
         Column(1)       =   "frmNuevoContratoMarco.frx":000C
         FormatStylesCount=   6
         FormatStyle(1)  =   "frmNuevoContratoMarco.frx":0164
         FormatStyle(2)  =   "frmNuevoContratoMarco.frx":028C
         FormatStyle(3)  =   "frmNuevoContratoMarco.frx":033C
         FormatStyle(4)  =   "frmNuevoContratoMarco.frx":03F0
         FormatStyle(5)  =   "frmNuevoContratoMarco.frx":04C8
         FormatStyle(6)  =   "frmNuevoContratoMarco.frx":0580
         ImageCount      =   0
         PrinterProperties=   "frmNuevoContratoMarco.frx":0660
      End
   End
   Begin VB.ComboBox cboMoneda 
      Height          =   315
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2505
      Width           =   1260
   End
   Begin VB.TextBox txtReferencia 
      Height          =   630
      Left            =   1425
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1530
      Width           =   3705
   End
   Begin MSComCtl2.DTPicker dtpFechaInicio 
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   855
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   56492033
      CurrentDate     =   40077
   End
   Begin XtremeSuiteControls.PushButton cmdCerrar 
      Height          =   495
      Left            =   3450
      TabIndex        =   11
      Top             =   4050
      Width           =   1590
      _Version        =   786432
      _ExtentX        =   2805
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpFin 
      Height          =   315
      Left            =   1425
      TabIndex        =   12
      Top             =   1200
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   56492033
      CurrentDate     =   40077
   End
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      Top             =   180
      Width           =   3735
      _Version        =   786432
      _ExtentX        =   6588
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboClienteFacturar 
      Height          =   315
      Left            =   1440
      TabIndex        =   17
      Top             =   525
      Width           =   3735
      _Version        =   786432
      _ExtentX        =   6588
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
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
      Left            =   750
      TabIndex        =   15
      Top             =   570
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Monto Tope"
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
      Left            =   345
      TabIndex        =   13
      Top             =   2205
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fecha Fin"
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
      Left            =   555
      TabIndex        =   7
      Top             =   1245
      Width           =   780
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
      Left            =   615
      TabIndex        =   5
      Top             =   2535
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fecha Inicio"
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
      Left            =   315
      TabIndex        =   2
      Top             =   885
      Width           =   1005
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
      Left            =   420
      TabIndex        =   1
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Centro de Costo"
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
      Left            =   30
      TabIndex        =   0
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmNuevoContratoMarco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clientes As New Collection
Private monedas As New Collection

Private fechas As New Collection

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim Ot As New OrdenTrabajo

    'Set ot.FechasPreciosMarco = fechas

    Set Ot.Cliente = DAOCliente.BuscarPorID(Me.cboCliente.ItemData(Me.cboCliente.ListIndex))
    Set Ot.ClienteFacturar = DAOCliente.BuscarPorID(Me.cboClienteFacturar.ItemData(Me.cboClienteFacturar.ListIndex))

    Ot.FechaInicioMarco = Me.dtpFechaInicio.value
    Ot.FechaFinMarco = Me.dtpFin.value
    Ot.descripcion = Me.txtReferencia.text
    Set Ot.Moneda = monedas.item(CStr(Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex)))

    Ot.FechaEntrega = Me.dtpFin.value
    Ot.OTMarcoIdPadre = -1

    Ot.MontoTopeMarco = Val(Me.txtMontoTotal.text)

    If DAOOrdenTrabajo.Save(Ot) Then
        MsgBox "La orden de trabajo se guardo con el número " & Ot.id, vbInformation + vbOKOnly
        Dim EVENTO As New clsEventoObserver
        Set EVENTO.Elemento = Ot
        EVENTO.EVENTO = agregar_
        Set EVENTO.Originador = Me
        Channel.Notificar EVENTO, ordenesTrabajo

        Unload Me
    End If
End Sub

Private Sub dtpFechaInicio_Change()
    Me.dtpFin.value = DateAdd("m", 12, Me.dtpFechaInicio.value)
End Sub

Private Sub Form_Load()

    FormHelper.Customize Me
    CustomizeGrid Me.gridFechasPrecios, , True

    DAOCliente.llenarComboXtremeSuite Me.cboCliente, True, True, True
    DAOCliente.llenarComboXtremeSuite Me.cboClienteFacturar, True, True, True
    Me.dtpFechaInicio.value = Now
    Me.dtpFin.value = DateAdd("m", 12, Now)

    Me.gridFechasPrecios.ItemCount = 0

    Set monedas = DAOMoneda.GetAll()
    Dim mon As clsMoneda

    Me.cboMoneda.Clear
    For Each mon In monedas
        Me.cboMoneda.AddItem mon.NombreCorto
        Me.cboMoneda.ItemData(Me.cboMoneda.NewIndex) = mon.id
    Next mon
    If monedas.count > 0 Then cboMoneda.ListIndex = 0
End Sub

Private Sub gridFechasPrecios_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = Not IsDate(Me.gridFechasPrecios.value(1))
End Sub

Private Sub gridFechasPrecios_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    fechas.Add Values(1)
End Sub

Private Sub gridFechasPrecios_UnboundDelete(ByVal RowIndex As Long, ByVal Bookmark As Variant)
    If RowIndex > 0 And fechas.count > 0 Then
        fechas.remove RowIndex
    End If
End Sub

Private Sub gridFechasPrecios_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And fechas.count > 0 Then Values(1) = fechas.item(RowIndex)
End Sub

Private Sub gridFechasPrecios_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    If RowIndex > 0 And fechas.count > 0 Then
        fechas.Add Values(1), , , RowIndex
        fechas.remove RowIndex
    End If
End Sub
