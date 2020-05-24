VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReCotizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recotizar presupuesto..."
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "[ Datos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txtReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   720
         Width           =   10335
      End
      Begin VB.Label lblCliente 
         Caption         =   "Label14"
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   360
         Width           =   10575
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Re 
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
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Elementos ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      TabIndex        =   19
      Top             =   3720
      Width           =   11775
      Begin VB.Frame Frame4 
         Height          =   495
         Left            =   3240
         TabIndex        =   44
         Top             =   5160
         Width           =   4695
         Begin VB.Label Label19 
            Caption         =   "Materiales"
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
            TabIndex        =   48
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblPMA 
            Height          =   255
            Left            =   1200
            TabIndex        =   47
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Mano de obra"
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
            Left            =   2160
            TabIndex        =   46
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblPMD 
            Height          =   255
            Left            =   3480
            TabIndex        =   45
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ver..."
         Default         =   -1  'True
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Qui 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   5400
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generar"
         Height          =   375
         Left            =   10680
         TabIndex        =   20
         Top             =   5280
         Width           =   975
      End
      Begin MSComctlLib.ListView lstPresupuesto 
         Height          =   4335
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7646
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
         NumItems        =   8
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
            Text            =   "Total"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Info Adicional"
            Object.Width           =   4939
         EndProperty
      End
      Begin VB.Label tota 
         Caption         =   "Sub-Total"
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
         Left            =   6720
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.Label subtot 
         Height          =   255
         Left            =   7800
         TabIndex        =   36
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label tot 
         Height          =   255
         Left            =   10080
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Total"
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
         Left            =   9360
         TabIndex        =   34
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Lista de piezas disponibles"
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
         TabIndex        =   24
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblidpieza 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Configuración ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11775
      Begin VB.CommandButton Command5 
         Caption         =   "Estimar días"
         Height          =   255
         Left            =   3480
         TabIndex        =   43
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox manteOferta 
         Height          =   285
         Left            =   2400
         TabIndex        =   40
         Text            =   "0"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Left            =   1800
         TabIndex        =   38
         Text            =   "1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtDescuento 
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtMarkupMDO 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   7935
      End
      Begin VB.TextBox men10 
         Height          =   285
         Left            =   2520
         TabIndex        =   4
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox men15 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox mas15 
         Height          =   285
         Left            =   5760
         TabIndex        =   2
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Recalcular"
         Height          =   375
         Left            =   10200
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Mantenimiento de oferta"
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
         TabIndex        =   42
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "Días"
         Height          =   255
         Left            =   4800
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Días"
         Height          =   255
         Left            =   3000
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Descuento"
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
         Left            =   3360
         TabIndex        =   33
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblidpresupuesto 
         Caption         =   "lblIdPresupuesto"
         Height          =   255
         Left            =   7800
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Gastos Configurados"
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblGastos 
         Caption         =   "Label3"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "MarkUp Mano de obra"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "MarkUp Materiales"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "<10Kg"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "<15Kg"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   ">15Kg"
         Height          =   255
         Left            =   5280
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "%"
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
         Left            =   6360
         TabIndex        =   11
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "%"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "%"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblGastosReal 
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Fecha de entrega"
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
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label idcliente 
         Caption         =   "Label14"
         Height          =   375
         Left            =   7800
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmReCotizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseV As New classVentas
Dim claseC As New classConfigurar
Dim claseS As New classStock
Dim claseP As New classPlaneamiento
Private Sub Command1_Click()
    frmElegirPieza.lblGastos = Me.lblGastosReal
    frmElegirPieza.mas15 = Me.mas15
    frmElegirPieza.men10 = Me.men10
    frmElegirPieza.men15 = Me.men15
    frmElegirPieza.lblMuMDO = Me.txtMarkupMDO
    frmElegirPieza.lblOrigen = 2
    frmElegirPieza.lblCliente = Me.lblCliente
    frmElegirPieza.lblidCliente = Me.idCliente
    frmElegirPieza.Show 1
End Sub


Private Sub Command2_Click()
recalcule
End Sub
Private Sub Command3_Click()
deta = normaliza(Me.txtReferencia)
'fecha_entrega = Format(Me.DTPicker1, "yyyy/mm/dd")
fecha_creado = Format(Now(), "yyyy/mm/dd hh:mm:ss")

'idcli = Me.cboCliente.ItemData(Me.cboCliente.ListIndex)
idcli = Me.idCliente
idVendedor = -1
esta = 0 '1- generado
descuento = Me.txtDescuento  'el descuento se efectuara una ves grabado el presupuesto(recotizar)
porcMDO = CDbl(Me.txtMarkupMDO)
PorcMEN10 = CDbl(Me.men10)
porcMEN15 = CDbl(Me.men15)
PorcMas15 = CDbl(Me.mas15)
gastoss = CDbl(Me.lblGastosReal)

If claseV.buscarDetalle(Trim(Me.txtReferencia)) Then
MsgBox "La referencia del presupuesto ya existe en la base de datos", vbCritical, "Error"
Else
g = MsgBox("¿Está conforme con los datos ingresados?", vbYesNo, "Confirmación")
1 If g = 6 Then
    claseV.agregar_presupuesto deta, fecha_entrega, fecha_creado, idcli, idVendedor, porcMDO, PorcMEN10, porcMEN15, PorcMas15, gastoss, esta, descuento, Me.lstPresupuesto, CInt(Me.manteOferta), venci, CLng(Me.lblidpresupuesto)
End If
End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
Dim id As Integer, canti As Long
enCurso = claseP.TiempoPedidos(2)
Pendientes = claseP.TiempoPedidos(1)
estePresu = 0
For X = 1 To Me.lstPresupuesto.ListItems.count
    id = Me.lstPresupuesto.ListItems(X).ListSubItems(6)
    canti = Me.lstPresupuesto.ListItems(X).ListSubItems(1)
    estePresu = estePresu + claseS.TiemposPieza(id, canti)
Next X
dias = cuantosDias((enCurso + Pendientes + estePresu) / 60) 'horas diariasº
Me.txtDias = Round(dias, 0)
'Command4_Click
End Sub

Private Sub Form_Activate()
Me.txtDescuento = 0
Dim strsql As String
strsql = "Select p.idCliente,c.razon,p.detalle,p.porcMDO, p.PorcMen15,p.PorcMas15,p.PorcMen10 from presupuestos p, clientes c where p.idcliente=c.id and p.id=" & CInt(Me.lblidpresupuesto)
claseV.ejecutar strsql
Me.lblCliente = claseV.cliente
Me.txtReferencia = "Re: " & claseV.detallePresupuesto
Me.lblGastos = claseC.gastos & "%"
Me.lblGastosReal = claseC.gastos
Me.txtMarkupMDO = claseV.muMDO
Me.mas15 = claseV.MuMas15
Me.men10 = claseV.MuMen10
Me.men15 = claseV.MuMen15
claseC.ejecutar "Select manteOferta from signoplast.configuracion limit 1"
Me.manteOferta = claseC.manteOferta

Me.idCliente = claseV.idCliente
claseV.llenar_lista_presupuesto CInt(Me.lblidpresupuesto), Me.lstPresupuesto, 1
Me.recalcule
End Sub

Private Sub lstPresupuesto_DblClick()
If Me.lstPresupuesto.ListItems.count > 0 Then
    frmModificaCantidad.Text1 = Me.lstPresupuesto.SelectedItem.ListSubItems(1)
    frmModificaCantidad.lblid = Me.lstPresupuesto.SelectedItem.ListSubItems(6)
    frmModificaCantidad.origen = 2
    If Trim(Me.lstPresupuesto.SelectedItem.ListSubItems(7).Text) <> Empty Then
        frmModificaCantidad.Text2 = Me.lstPresupuesto.SelectedItem.ListSubItems(7)
    End If
       
    frmModificaCantidad.Show 1
    'Command2_Click
End If

End Sub

Private Sub manteOferta_GotFocus()
foco Me.manteOferta
End Sub

Private Sub manteOferta_Validate(Cancel As Boolean)
If Not IsNumeric(Trim(Me.manteOferta)) Then
Cancel = True
Else
Cancel = False
End If


End Sub

Private Sub mas15_GotFocus()
foco Me.mas15
End Sub
Private Sub men10_GotFocus()
foco Me.men10
End Sub
Private Sub men15_GotFocus()
foco Me.men15
End Sub

Private Sub txtDescuento_GotFocus()
foco Me.txtDescuento
End Sub
Private Sub txtDescuento_Validate(Cancel As Boolean)
If Not IsNumeric(Trim(Me.txtDescuento)) Then
Cancel = True
Else
Cancel = False
End If
End Sub

Public Sub recalcule()
On Error Resume Next
Dim precio As Double, kg As Double
Dim id As Integer
tote = 0
For nn = 1 To Me.lstPresupuesto.ListItems.count
 id = Me.lstPresupuesto.ListItems(nn).ListSubItems(6)
 claseS.calcular_valor_materiales id, kg, precio
 materiales = precio * valorPorPeso(kg, men10, men15, mas15)
 mdo = claseS.calcular_valor_mdo(id, 1)
 fijo = claseS.calcular_valor_mdo(id, 0)
 cantidad = CInt(Me.lstPresupuesto.ListItems(nn).ListSubItems(1))
 cambio = claseS.calcular_valor_mdo(id, -1)
 'amort = amortiza(cantidad)
 amort = CLng(Me.lstPresupuesto.ListItems(nn).ListSubItems(2))
 

 manodeobra = mdo + ((cambio / cantidad) + (fijo / amort))
 manodeobra = manodeobra * (CDbl((Me.txtMarkupMDO) / 100) + 1)
 pma = materiales + pma
 PMD = manodeobra + PMD
 
 unitario = manodeobra + materiales
 unitario = ((CDbl(lblGastosReal) * unitario) / 100) + unitario
 unitario = Math.Round(unitario, 2)
 total = unitario * cantidad

tote = total + tote
des = (1 - (CDbl(Me.txtDescuento) / 100))
tote2 = tote * des
 Me.lstPresupuesto.ListItems(nn).ListSubItems(5).Text = total
 Me.lstPresupuesto.ListItems(nn).ListSubItems(4).Text = unitario
Next nn
Me.subtot = Math.Round(tote, 2)
Me.tot = Math.Round(tote2, 2)
'calculo los porcentuales de MDO Y MAT
TOTILLO = PMD + pma
If TOTILLO = 0 Then TOTILLO = 1
PMD = PMD / TOTILLO
pma = pma / TOTILLO
p1 = Math.Round(PMD, 2) * 100
p2 = Math.Round(pma, 2) * 100

Me.lblPMA = p2 & "%"
Me.lblPMD = p1 & "%"
End Sub

Private Sub txtMarkupMDO_GotFocus()
foco Me.txtMarkupMDO
End Sub
