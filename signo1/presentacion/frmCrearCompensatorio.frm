VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmCrearCompensatorio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compensatorio"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.DateTimePicker DateTimePicker1 
      Height          =   300
      Left            =   1305
      TabIndex        =   11
      Top             =   1395
      Width           =   3150
      _Version        =   786432
      _ExtentX        =   5556
      _ExtentY        =   529
      _StockProps     =   68
      CurrentDate     =   40725.6314467593
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   810
      Left            =   1305
      TabIndex        =   9
      Top             =   1755
      Width           =   3150
   End
   Begin VB.TextBox txtImporte 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1050
      Width           =   3150
   End
   Begin VB.TextBox txtComprobante 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   345
      Width           =   3150
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   420
      Left            =   1170
      TabIndex        =   0
      Top             =   2775
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Aceptar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Cancel          =   -1  'True
      Height          =   420
      Left            =   2445
      TabIndex        =   1
      Top             =   2790
      Width           =   1095
      _Version        =   786432
      _ExtentX        =   1931
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboTipos 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   690
      Width           =   3150
      _Version        =   786432
      _ExtentX        =   5556
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelación"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      Height          =   195
      Left            =   645
      TabIndex        =   6
      Top             =   1125
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   765
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobante"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   405
      Width           =   945
   End
End
Attribute VB_Name = "frmCrearCompensatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vFactura As clsFacturaProveedor
Private op As OrdenPago
Dim compe As Compensatorio
Private Sub Form_Load()
    Customize Me
End Sub

Private Sub PushButton1_Click()
    Set compe = New Compensatorio
    Set compe.Comprobante = vFactura
    compe.Monto = CDbl(Me.txtImporte)
    compe.Observacion = Me.txtObservaciones
    compe.FechaCancelacion = Me.DateTimePicker1.value
    compe.Tipo = Me.cboTipos.ItemData(Me.cboTipos.ListIndex)
    compe.IdOrdenPago = op.id
    compe.alicuotaPercepcion = op.Alicuota
    compe.NetoGravadoCompensado = (funciones.FormatearDecimales(compe.Monto / (1 + funciones.FormatearDecimales(compe.Comprobante.IvaAplicado(1).Alicuota.Alicuota / 100))))
    compe.MontoAPercibir = compe.Monto - compe.NetoGravadoCompensado

    op.Compensatorios.Add compe
    Unload Me
End Sub

Private Sub PushButton2_Click()
    Unload Me

End Sub


Public Function Cargar(Optional documento As clsFacturaProveedor, Optional OrdenPago As OrdenPago)
    Load Me
    Set vFactura = documento
    Set op = OrdenPago



    Me.txtComprobante = documento.NumeroFormateado
    Me.cboTipos.AddItem TiposCompensatorio.item(CStr(TipoCompensatorio.TC_Credito))
    cboTipos.ItemData(cboTipos.NewIndex) = CStr(TipoCompensatorio.TC_Credito)

    Me.cboTipos.AddItem TiposCompensatorio.item(CStr(TipoCompensatorio.TC_Debido))
    cboTipos.ItemData(cboTipos.NewIndex) = CStr(TipoCompensatorio.TC_Debido)



    'Me.cboTipos.AddItem CStr(TiposCompensatorio.Item(CStr(TipoCompensatorio.TC_Debido))), TiposCompensatorio.Item(CStr(TipoCompensatorio.TC_Debido))

End Function
