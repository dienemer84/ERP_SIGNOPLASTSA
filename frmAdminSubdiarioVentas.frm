VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdminSubdiarioVentas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Subdiario de ventas..."
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstSubdiarioVentas 
      Height          =   4575
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Comp"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Razon Social"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CUIT"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Cond. IVA"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Neto Gravado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "I.V.A."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Percep IB"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Exento"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Total"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Seleccione período de facturación ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exportar"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Default         =   -1  'True
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   39660
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   57475073
         CurrentDate     =   39660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
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
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
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
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAdminSubdiarioVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clase As New classAdministracion
Dim rs As recordset
Dim desde As Date
Dim hasta As Date

Private Sub Command2_Click()

    Dim neto As Double, Total As Double
    Dim riva As Double, excento As Double
    Dim x As ListItem
    Dim per As Double
    Me.lstSubdiarioVentas.ListItems.Clear


    If Me.DTDesde > Me.DTHasta Then
        MsgBox "Error en la seleccion de fechas!", vbCritical, "Error"
        Exit Sub
    Else
        desde = Me.DTDesde
        hasta = Me.DTHasta
    End If

    Set rs = clase.subdiario_ventas(desde, hasta)
    Total = 0
    neto = 0
    riva = 0
    per = 0
    excento = 0

    While Not rs.EOF
        Set x = Me.lstSubdiarioVentas.ListItems.Add(, , rs!FEcha)
        x.SubItems(1) = rs!factura
        x.SubItems(2) = rs!Cliente
        x.SubItems(3) = rs!Cuit
        x.SubItems(4) = rs!Iva
        x.SubItems(5) = rs!netograbado
        x.SubItems(6) = rs!TotalIva
        x.SubItems(7) = rs!totalPerib
        x.SubItems(8) = rs!excento
        x.SubItems(9) = rs!Total
        Total = Total + rs!Total
        neto = neto + rs!netograbado
        riva = riva + rs!TotalIva
        per = per + rs!totalPerib
        excento = excento + rs!excento
        If rs!estado = 3 Then
            x.ForeColor = vbRed
        End If

        rs.MoveNext
    Wend

    Set x = Me.lstSubdiarioVentas.ListItems.Add(, , Empty)
    x.SubItems(1) = Empty
    x.SubItems(2) = Empty
    x.SubItems(3) = Empty
    x.SubItems(4) = Empty

    x.SubItems(5) = funciones.formatearDecimales(neto, 2)
    x.SubItems(6) = funciones.formatearDecimales(riva, 2)
    x.SubItems(7) = funciones.formatearDecimales(per, 2)
    x.SubItems(8) = funciones.formatearDecimales(excento, 2)
    x.SubItems(9) = funciones.formatearDecimales(Total, 2)
    x.ListSubItems(5).Bold = True
    x.ListSubItems(6).Bold = True
    x.ListSubItems(7).Bold = True
    x.ListSubItems(8).Bold = True
    x.ListSubItems(9).Bold = True
End Sub

Private Sub Command3_Click()
    clase.exportaSubDiarioVentas Me.lstSubdiarioVentas, desde, hasta
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
frmAdminSubdiariosVentasv2.Show
End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    desde = CDate(1 & "-" & Month(Now) & "-" & Year(Now))
    Me.DTDesde = desde
    Me.DTHasta = Now


End Sub

