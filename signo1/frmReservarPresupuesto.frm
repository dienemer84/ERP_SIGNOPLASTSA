VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmVentasPresupuestoNuevo 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservar Presupuesto"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   ClipControls    =   0   'False
   Icon            =   "frmReservarPresupuesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   7695
   Begin XtremeSuiteControls.PushButton Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   2415
      TabIndex        =   7
      Top             =   1305
      Width           =   1305
      _Version        =   786432
      _ExtentX        =   2302
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Generar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Height          =   405
      Left            =   3915
      TabIndex        =   6
      Top             =   1290
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Salir"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.DateTimePicker DTPicker1 
      Height          =   270
      Left            =   1200
      TabIndex        =   5
      Top             =   855
      Width           =   3150
      _Version        =   786432
      _ExtentX        =   5556
      _ExtentY        =   476
      _StockProps     =   68
      CurrentDate     =   40154.378125
   End
   Begin VB.TextBox txtReferencia 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   6375
   End
   Begin MSComctlLib.ListView lstPresupuesto 
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Amort"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Detalle"
         Object.Width           =   6209
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Unitario"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Venta"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Info Adicional"
         Object.Width           =   4939
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   315
      Left            =   1215
      TabIndex        =   8
      Top             =   90
      Width           =   6375
      _Version        =   786432
      _ExtentX        =   11245
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   -1  'True
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Cliente"
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
      Left            =   255
      TabIndex        =   4
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Referencia"
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
      Left            =   165
      TabIndex        =   3
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Vencimiento"
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
      Left            =   90
      TabIndex        =   2
      Top             =   855
      Width           =   1050
   End
End
Attribute VB_Name = "frmVentasPresupuestoNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim claseConf As New classConfigurar

Dim tmpPresu As clsPresupuesto
Private Sub Guardar()
    Dim idCliente As clsCliente
    If Me.cboCliente.ListIndex = -1 Then Exit Sub
    If DAOPresupuestos.ExisteDetalle(Trim(txtReferencia)) Then
        MsgBox "La referencia del presupuesto ya existe en la base de datos", vbCritical, "Error"
    Else

        Set tmpPresu = New clsPresupuesto

        tmpPresu.detalle = UCase(Me.txtReferencia)
        tmpPresu.FechaEntrega = 0
        tmpPresu.PorcentajeManoObraMuerta = Configurar.Mano_obra_muerta
        tmpPresu.fechaCreado = Now()
        tmpPresu.VencimientoPresupuesto = Format(Me.DTPicker1, "YYYY/MM/DD")
        Set tmpPresu.UsuarioCreado = funciones.GetUserObj
        tmpPresu.PorcMDO = Configurar.PorcMO
        tmpPresu.PorcMen10 = Configurar.PorMAMenos10
        tmpPresu.PorcMen15 = Configurar.PorMAMenos15
        tmpPresu.PorcMas15 = Configurar.PorMaMas15
        tmpPresu.Gastos = claseConf.Gastos
        Set tmpPresu.moneda = DAOMoneda.GetById(0)
        Set tmpPresu.cliente = DAOCliente.BuscarPorID(CLng(Me.cboCliente.ItemData(Me.cboCliente.ListIndex)))
        tmpPresu.manteOferta = 0
        tmpPresu.Descuento = 0
        tmpPresu.EstadoPresupuesto = ACotizar_

        If DAOPresupuestos.Save(tmpPresu) Then
            DAOEvento.Publish tmpPresu.Id, TipoEventoBroadcast.TEB_PresupuestoCreado

            Dim EVENTO As New clsEventoObserver
            Set EVENTO.Elemento = tmpPresu
            EVENTO.EVENTO = agregar_
            Set EVENTO.Originador = Me
            Channel.Notificar EVENTO, Presupuestos_
        End If
        nuevoPresu
    End If
End Sub

Private Sub Command1_Click()
    Guardar
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    FormHelper.Customize Me
    nuevoPresu
    Me.DTPicker1 = Now
    DAOCliente.llenarComboXtremeSuite Me.cboCliente, False, True
    If Trim(Me.txtReferencia) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
    nuevoPresu
End Sub
Private Sub txtReferencia_Change()
    If Trim(Me.txtReferencia) = Empty Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub

Public Sub nuevoPresu()
    Me.caption = "Reservar Presupuesto Nro " & Format(DAOPresupuestos.ProximoPresupuesto, "0000") & "]"
    Me.txtReferencia = Empty
End Sub

