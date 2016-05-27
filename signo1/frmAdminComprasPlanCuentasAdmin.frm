VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmAdminComprasPlanCuentasAdmin 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Plan de Cuentas"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6465
   Begin XtremeSuiteControls.PushButton cmdExport 
      Height          =   330
      Left            =   5010
      TabIndex        =   10
      Top             =   675
      Width           =   1080
      _Version        =   786432
      _ExtentX        =   1905
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Exportar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1605
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   6315
      _Version        =   786432
      _ExtentX        =   11139
      _ExtentY        =   2831
      _StockProps     =   79
      Caption         =   "Parámetros de búsqueda"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   435
         Left            =   4800
         TabIndex        =   12
         Top             =   1110
         Width           =   1155
         _Version        =   786432
         _ExtentX        =   2037
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Imprimir Sólo Valuados"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   330
         Left            =   4920
         TabIndex        =   4
         Top             =   195
         Width           =   1080
         _Version        =   786432
         _ExtentX        =   1905
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtCuenta 
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         Top             =   345
         Width           =   3630
      End
      Begin XtremeSuiteControls.DateTimePicker dtpDesde 
         Height          =   315
         Left            =   1050
         TabIndex        =   5
         Top             =   705
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
         Left            =   3225
         TabIndex        =   6
         Top             =   705
         Width           =   1470
         _Version        =   786432
         _ExtentX        =   2593
         _ExtentY        =   556
         _StockProps     =   68
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   315
         Left            =   3570
         TabIndex        =   11
         Top             =   1140
         Width           =   1125
         _Version        =   786432
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   270
         Left            =   105
         TabIndex        =   13
         Top             =   1200
         Width           =   1500
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   750
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Desde"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   2655
         TabIndex        =   7
         Top             =   765
         Width           =   420
         _Version        =   786432
         _ExtentX        =   741
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Hasta"
         BackColor       =   12632256
         AutoSize        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   240
         Left            =   255
         TabIndex        =   2
         Top             =   390
         Width           =   1305
      End
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   1755
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   8096
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      AllowDelete     =   -1  'True
      DataMode        =   99
      AllowAddNew     =   -1  'True
      ColumnHeaderHeight=   285
      IntProp1        =   0
      ColumnsCount    =   3
      Column(1)       =   "frmAdminComprasPlanCuentasAdmin.frx":0000
      Column(2)       =   "frmAdminComprasPlanCuentasAdmin.frx":0110
      Column(3)       =   "frmAdminComprasPlanCuentasAdmin.frx":01FC
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmAdminComprasPlanCuentasAdmin.frx":0360
      FormatStyle(2)  =   "frmAdminComprasPlanCuentasAdmin.frx":0498
      FormatStyle(3)  =   "frmAdminComprasPlanCuentasAdmin.frx":0548
      FormatStyle(4)  =   "frmAdminComprasPlanCuentasAdmin.frx":05FC
      FormatStyle(5)  =   "frmAdminComprasPlanCuentasAdmin.frx":06D4
      FormatStyle(6)  =   "frmAdminComprasPlanCuentasAdmin.frx":078C
      ImageCount      =   0
      PrinterProperties=   "frmAdminComprasPlanCuentasAdmin.frx":086C
   End
   Begin XtremeSuiteControls.Label lblTotal 
      Height          =   195
      Left            =   105
      TabIndex        =   9
      Top             =   6525
      Width           =   6255
      _Version        =   786432
      _ExtentX        =   11033
      _ExtentY        =   344
      _StockProps     =   79
      Alignment       =   2
      AutoSize        =   -1  'True
   End
End
Attribute VB_Name = "frmAdminComprasPlanCuentasAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vCuenta As clsCuentaContable
Dim cuentas As Collection
Dim edita As Boolean

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    If Not DAOCuentaContable.ExportarColeccion(cuentas, "") Then GoTo err1

    Exit Sub
err1:
    MsgBox "Error al exportar", vbCritical, "Error"
End Sub

Private Sub Form_Load()

    FormHelper.Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, True

    llenarPlan
End Sub
Private Sub llenarPlan()
    Dim cuentas1 As New Collection
    Dim filtro As String
    Dim rango As String
    filtro = "1 = 1 "
    If Not IsNull(Me.dtpDesde.value) Then
        rango = rango & " AND fp.fecha >= '" & Format(Me.dtpDesde.value, "yyyy-mm-dd") & "' "
    End If

    If Not IsNull(Me.dtpHasta.value) Then
        rango = rango & " AND fp.fecha <= '" & Format(Me.dtpHasta.value, "yyyy-mm-dd") & "' "
    End If

    filtro = filtro & " and     " & "codigo like '%" & Me.txtCuenta & "%' or nombre like '%" & Me.txtCuenta & "%'"


    Set cuentas = DAOCuentaContable.GetAll(True, filtro)
    DAOCuentaContable.PutSaldos cuentas, rango
    Dim T As Double
    Dim c As clsCuentaContable
    For Each c In cuentas
        T = T + c.TotalAcumulado
    Next
    Me.lblTotal.caption = "Total AR$ " & funciones.FormatearDecimales(T)

    Me.GridEX1.ItemCount = 0
    Me.GridEX1.ItemCount = cuentas.count

End Sub


Private Sub GridEX1_BeforeUpdate(ByVal Cancel As GridEX20.JSRetBoolean)
    Cancel = (GridEX1.value(1) = vbNullString) Or (GridEX1.value(1) = vbNullString)

End Sub

Private Sub GridEX1_UnboundAddNew(ByVal NewRowBookmark As GridEX20.JSRetVariant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set vCuenta = New clsCuentaContable
    vCuenta.codigo = UCase(Values(1))
    vCuenta.nombre = UCase(Values(2))
    If Not DAOCuentaContable.Save(vCuenta) Then GoTo err1
    cuentas.Add vCuenta
    Exit Sub
err1:


End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set vCuenta = cuentas.item(RowIndex)
    Values(1) = vCuenta.codigo
    Values(2) = vCuenta.nombre
    Values(3) = funciones.FormatearDecimales(vCuenta.TotalAcumulado)
    Exit Sub
err1:
End Sub

Private Sub GridEX1_UnboundUpdate(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    On Error GoTo err1
    Set vCuenta = cuentas(RowIndex)
    vCuenta.codigo = UCase(Values(1))
    vCuenta.nombre = UCase(Values(2))
    If Not DAOCuentaContable.Update(vCuenta) Then GoTo err1
    Exit Sub
err1:

End Sub


Private Sub PushButton1_Click()
    llenarPlan
End Sub

Private Sub PushButton2_Click()
    Dim rango As String
    If Not IsNull(Me.dtpDesde.value) Then rango = "Desde " & Format(Me.dtpDesde.value, "dd-mm-yyyy")
    If Not IsNull(Me.dtpHasta.value) Then rango = rango & " Hasta " & Format(Me.dtpHasta, "dd-mm-yyyy")

    If IsNull(Me.dtpDesde) And IsNull(Me.dtpHasta) Then rango = "SIN ESPECIFICAR"


    DAOCuentaContable.ImprimirColeccion cuentas, rango, Me.CheckBox1.value = xtpChecked
End Sub
