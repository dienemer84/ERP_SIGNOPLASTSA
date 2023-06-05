VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmResumenSaldosProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmen de Saldos de Proveedores"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdParar 
      Height          =   420
      Left            =   1200
      TabIndex        =   9
      Top             =   6255
      Width           =   525
      _Version        =   786432
      _ExtentX        =   926
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "X"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   360
      Left            =   60
      TabIndex        =   7
      Top             =   6255
      Width           =   1125
      _Version        =   786432
      _ExtentX        =   1984
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin GridEX20.GridEX GridEX1 
      Height          =   5445
      Left            =   30
      TabIndex        =   0
      Top             =   645
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   9604
      Version         =   "2.0"
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ColumnAutoResize=   -1  'True
      MethodHoldFields=   -1  'True
      DataMode        =   99
      ColumnHeaderHeight=   285
      IntProp1        =   0
      IntProp2        =   0
      IntProp7        =   0
      ColumnsCount    =   2
      Column(1)       =   "frmResumenSaldosProv.frx":0000
      Column(2)       =   "frmResumenSaldosProv.frx":0120
      FormatStylesCount=   6
      FormatStyle(1)  =   "frmResumenSaldosProv.frx":020C
      FormatStyle(2)  =   "frmResumenSaldosProv.frx":0344
      FormatStyle(3)  =   "frmResumenSaldosProv.frx":03F4
      FormatStyle(4)  =   "frmResumenSaldosProv.frx":04A8
      FormatStyle(5)  =   "frmResumenSaldosProv.frx":0580
      FormatStyle(6)  =   "frmResumenSaldosProv.frx":0638
      ImageCount      =   0
      PrinterProperties=   "frmResumenSaldosProv.frx":0718
   End
   Begin XtremeSuiteControls.PushButton Obtener 
      Height          =   360
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   1305
      _Version        =   786432
      _ExtentX        =   2302
      _ExtentY        =   635
      _StockProps     =   79
      Caption         =   "Obtener"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   3855
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   4455
      _Version        =   786432
      _ExtentX        =   7858
      _ExtentY        =   529
      _StockProps     =   93
      Appearance      =   6
   End
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   315
      Left            =   2325
      TabIndex        =   4
      Top             =   225
      Width           =   1470
      _Version        =   786432
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   68
      CheckBox        =   -1  'True
      Format          =   1
   End
   Begin VB.Label lblCant 
      Height          =   315
      Left            =   8355
      TabIndex        =   8
      Top             =   210
      Width           =   990
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   9255
      TabIndex        =   6
      Top             =   6315
      Width           =   45
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   1695
      TabIndex        =   5
      Top             =   285
      Width           =   420
      _Version        =   786432
      _ExtentX        =   741
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Hasta"
      BackColor       =   12632256
      AutoSize        =   -1  'True
   End
   Begin VB.Label lblproceso 
      Height          =   390
      Left            =   1770
      TabIndex        =   1
      Top             =   6255
      Width           =   7470
   End
End
Attribute VB_Name = "frmResumenSaldosProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dto As DTONombreMonto
Dim col As New Collection
Dim col2 As New Collection
Dim condition As String
Dim enable As Boolean
Public TipoPersonaCta As TipoPersona


Private Sub cmdParar_Click()
    enable = False
    cmdParar.Enabled = enable
End Sub

Private Sub Form_Load()
    Customize Me
    GridEXHelper.CustomizeGrid Me.GridEX1, False, False
    Me.GridEX1.ItemCount = 0




End Sub

Private Sub GridEX1_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    GridEXHelper.ColumnHeaderClick Me.GridEX1, Column
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Set dto = col2(RowIndex)


    Values(1) = dto.nombre
    Values(2) = funciones.FormatearDecimales(dto.Monto)
End Sub

Private Sub Obtener_Click()
    enable = True
    cmdParar.Enabled = enable
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount
    Me.lblCant.Visible = True
    Me.lblproceso.Visible = True
    Me.ProgressBar1.Visible = True
    Me.GridEX1.ItemCount = 0
    Dim Detalles As Collection
    Set Detalles = New Collection
    Set col2 = New Collection
    Dim c As Long
    Dim rs As Recordset

    If TipoPersonaCta = TipoPersona.proveedor_ Then
        Set rs = conectar.RSFactory("SELECT * FROM proveedores  order  by razon asc ")
    Else
        Set rs = conectar.RSFactory("SELECT * FROM clientes  order by razon asc ")
    End If
    c = 0
    While Not rs.EOF And Not rs.BOF
        c = c + 1
        rs.MoveNext
    Wend


    Dim dto As DTONombreMonto

    If c >= 1 Then rs.MoveFirst


    Me.ProgressBar1.max = c
    Dim d As Long
    d = 0
    While Not rs.EOF And Not rs.BOF
        d = d + 1



        If Not IsNull(Me.dtpHasta.value) Then
            condition = conectar.Escape(Format(Me.dtpHasta.value, "yyyy-mm-dd"))
        End If

        If TipoPersonaCta = TipoPersona.proveedor_ Then

            Set Detalles = DAOCuentaCorriente.FindAllDetallesProveedor(rs!Id, , condition, True, False)

        Else
            If Not IsNull(Me.dtpHasta.value) Then
                condition = Format(Me.dtpHasta.value, "yyyy-mm-dd")
            End If
            Set Detalles = DAOCuentaCorriente.FindAllDetalles(rs!Id, , condition)
        End If




        Set dto = New DTONombreMonto
        dto.Monto = DAOCuentaCorriente.GetSaldo(Detalles)
        dto.nombre = rs!razon
        If (dto.Monto >= 0.01 Or dto.Monto < -0.01) Then
            col2.Add dto
        End If
        Me.lblCant = CStr(d) & "/" & CStr(c)
        Me.lblproceso = "Procesando " & rs!razon
        Me.ProgressBar1.value = d
        DoEvents
        rs.MoveNext
        Me.GridEX1.ItemCount = col2.count
        If Not enable Then Exit Sub
    Wend

    Me.GridEX1.ItemCount = col2.count
    Me.ProgressBar1.Visible = False
    Dim T As Double

    For Each dto In col2
        T = T + dto.Monto
    Next
    Me.lblTotal = "Total: " & funciones.FormatearDecimales(T)
    Me.lblCant.Visible = False
    tickend = GetTickCount
    'Debug.Print "Tiempo total  ", tickend - tickStart
End Sub

Private Sub PushButton1_Click()


    With Me.GridEX1.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = True
        .Orientation = jgexPPPortrait
        .HeaderString(jgexHFCenter) = "Resumen de saldos"
        If Not IsNull(dtpHasta.value) Then
            .HeaderString(jgexHFLeft) = "Hasta  " & Format(Me.dtpHasta, "dd-mm-yyyy")
        End If
        .FooterString(jgexHFCenter) = Now
        .FooterString(jgexHFRight) = Me.lblTotal
    End With
    Load frmPrintPreview
    frmPrintPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    Me.GridEX1.PrintPreview frmPrintPreview.GEXPreview1
    frmPrintPreview.Show 1
End Sub
