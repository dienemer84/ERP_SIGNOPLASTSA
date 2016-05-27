VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVentasEstadisticasCotizaciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13065
   ClipControls    =   0   'False
   Icon            =   "frmEstadisticasCotizaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   13065
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "[ Por cliente ]"
      Height          =   6015
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4335
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listado"
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   4095
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Left            =   1200
            TabIndex        =   10
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   16711681
            CurrentDate     =   39069
         End
         Begin VB.CheckBox chFecha 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Rango"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chProc 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Procesados"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chEnv 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Envíados"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox ChNoCotizado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Cotizado"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox ChPendiente 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   16711681
            CurrentDate     =   39069
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar"
         Default         =   -1  'True
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
      Begin VB.ComboBox cboClientes 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstPorPeriodo 
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Período"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Monto Cotizado"
            Object.Width           =   4851
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seleccione Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblTotCot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label2"
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
         Top             =   5640
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   975
   End
   Begin MSComctlLib.ListView listados 
      Height          =   5415
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro Presu"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Detalle"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sub Total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dto"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Estado"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmVentasEstadisticasCotizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
Dim rs As Recordset
Dim clsStock As New classStock

Dim cli As Long
Dim fec1 As String
Dim fec2 As String
Dim Periodo As Boolean
Dim totcot2 As Double

Private Sub Check1_Click()
    If Me.chFecha.value Then
        MsgBox "si"
    Else
        MsgBox "no"
    End If


End Sub

Private Sub cboClientes_Change()
    cli = Me.cboClientes.ItemData(cboClientes.ListIndex)
    llenarPeriodo

End Sub

Private Sub cboClientes_Click()
    cli = Me.cboClientes.ItemData(cboClientes.ListIndex)
    llenarPeriodo
End Sub

Private Sub chFecha_Click()
    If Me.chFecha.value Then
        Me.DTPicker1.Enabled = True
        Me.DTPicker2.Enabled = True
    Else
        Me.DTPicker1.Enabled = False
        Me.DTPicker2.Enabled = False
    End If
End Sub

Private Sub Command1_Click()



    If Me.chEnv Or Me.chProc Or Me.ChNoCotizado Or Me.ChPendiente Then
        filtro1 = ""
        filtro = ""


        If Me.chProc And Me.chEnv And Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=2 or p.estado=3 or p.estado=1 or p.estado=7)"
        If Me.chProc And Not Me.chEnv And Not Me.ChNoCotizado And Not Me.ChPendiente Then filtro1 = "and p.estado=3"
        If Not Me.chProc And Me.chEnv And Not Me.ChNoCotizado And Not Me.ChPendiente Then filtro1 = "and p.estado=2"
        If Not Me.chProc And Not Me.chEnv And Me.ChNoCotizado And Not Me.ChPendiente Then filtro1 = "and p.estado=7"
        If Not Me.chProc And Not Me.chEnv And Not Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and p.estado=1"

        'procesado y enviado
        If Me.chProc And Me.chEnv And Not Me.ChNoCotizado And Not Me.ChPendiente Then filtro1 = "and (p.estado=3 or p.estado=2)"
        'procesado y no cotizado
        If Me.chProc And Not Me.chEnv And Me.ChNoCotizado And Not Me.ChPendiente Then filtro1 = "and (p.estado=3 or p.estado=7)"
        'procesado y pendiente
        If Me.chProc And Not Me.chEnv And Not Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=3 or p.estado=1)"
        'pendiente y enviado
        If Not Me.chProc And Me.chEnv And Not Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=1 or p.estado=2)"
        'pendiente y no cotizado
        If Not Me.chProc And Not Me.chEnv And Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=1 or p.estado=2)"
        'enviado y no cotizado
        If Not Me.chProc And Me.chEnv And Me.ChNoCotizado And Not Me.ChPendiente Then filtro1 = "and (p.estado=2 or p.estado=7)"


        'procesado y enviado y pendiente
        If Me.chProc And Me.chEnv And Not Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=3 or p.estado=2 or p.estado=1)"
        'procesado y enviado y no cotizado
        If Me.chProc And Me.chEnv And Me.ChNoCotizado And Not Me.ChPendiente Then filtro1 = "and (p.estado=3 or p.estado=2 or p.estado=7)"
        'procesado y no cotizado y pendiente
        If Me.chProc And Not Me.chEnv And Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=3 or p.estado=7 or p.estado=1)"
        'enviado y pendiente y no cotizado
        If Not Me.chProc And Me.chEnv And Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=7 or p.estado=2 or p.estado=1)"
        'enviado y pendiente y no cotizado y procesado
        If Me.chProc And Me.chEnv And Me.ChNoCotizado And Me.ChPendiente Then filtro1 = "and (p.estado=7 or p.estado=2 or p.estado=1 or p.estado=3)"





        'If Me.chProc And Me.chEnv Then filtro1 = "and (p.estado=3 or p.estado=2)"
        ' If Me.chProc And Not Me.chEnv Then filtro1 = "and (p.estado=3)"
        ' If Me.chEnv And Not Me.chProc Then filtro1 = "and p.estado=2"




        If Me.chFecha Then
            Periodo = True
            filtro2 = "and fecha>'" & Format(CDate(Me.DTPicker1), "yyyy-mm-dd") & "' and fecha <'" & Format(CDate(Me.DTPicker2), "yyyy-mm-dd") & "'"
            fec1 = Me.DTPicker1
            fec2 = Me.DTPicker2
        Else
            Periodo = False
        End If
        strsql = "select p.id,p.detalle,sum((dp.ValorUnitario*dp.Cantidad)) as SubtotalPedido,p.descuento, sum((dp.ValorUnitario*dp.Cantidad) * (1-(p.descuento/100))) as totalPedido, if(p.estado=2,'Enviado',if(p.estado=3,'Procesado',if(p.estado=1,'Pendiente','No Cotizado'))) as estado from presupuestos p, detalle_presupuesto dp  where p.idCliente=" & cli & " and dp.idPresupuesto=p.id " & filtro1 & " " & filtro2 & " group by p.id"
        Set rs = conectar.RSFactory(strsql)
        Me.listados.ListItems.Clear
        totcot = 0
        totcot2 = 0
        cantCot = 0
        While Not rs.EOF
            Set x = Me.listados.ListItems.Add(, , Format(rs!id, "0000"))
            x.SubItems(1) = rs!detalle
            x.SubItems(2) = Format(Math.Round(rs!subTotalPedido, 2), "0.00")
            x.SubItems(3) = rs!Descuento & "%"
            x.SubItems(4) = Format(Math.Round(rs!totalPedido, 2), "0.00")
            x.SubItems(5) = rs!estado
            totcot2 = rs!totalPedido + totcot2
            cantCot = cantCot + 1
            rs.MoveNext
        Wend
        'Me.lblInfo = "Total " & cantCot & " cotizaciones,  monto: $" & Math.Round(totcot2, 2)

    End If

End Sub

Private Sub Command2_Click()
    On Error GoTo Err554
    Me.CommonDialog1.ShowPrinter
    'If MsgBox("¿Seguro de imprimir?", vbYesNo, "Confirmación") = vbYes Then




    Set rs = conectar.RSFactory("select razon from clientes where id=" & cli)
    clie = rs!razon

    AnchoCol = 0

    For i = 1 To Me.listados.ColumnHeaders.count
        AnchoCol = AnchoCol + listados.ColumnHeaders(i).Width
    Next
    Espacio = 0

    Printer.Print "HISTORIAL COTIZACIONES"
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    Printer.Print "Cliente: " & cli & " - " & clie
    If Periodo Then
        Printer.Print "Periodo: " & fec1 & " hasta " & fec2

    End If

    Printer.Print
    'Imprime una línea
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)

    With Me.listados

        'Acá se imprimen los encabezados del ListView
        For i = 1 To .ColumnHeaders.count
            Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
            Printer.Print listados.ColumnHeaders(i).text;
            Printer.CurrentX = Espacio
        Next

        Printer.Print

        'Imprime una línea
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print

        'Este bucle recorre los items y subitems del ListView  y los imprime
        For i = 1 To .ListItems.count
            Espacio = 0

            Set lItem = .ListItems(i)
            Printer.Print lItem.text;
            'Recorremos las columnas
            For x = 1 To .ColumnHeaders.count - 1
                Espacio = Espacio + CInt(.ColumnHeaders(x).Width * Printer.ScaleWidth / AnchoCol)
                Printer.CurrentX = Espacio
                Printer.Print lItem.SubItems(x);

            Next

            'Otro espacio en blanco
            Printer.Print
        Next

    End With

    Printer.Print

    ''Imprime la línea de final de impresión
    Printer.Print
    Printer.Print "Total: $" & Format(totcot2, "0.00")
    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print

    'Texto del pie>
    Printer.Print Format(Date, "dd-mm-yyyy")


    'Comenzamos la impresión
    Printer.EndDoc

Err554:

    'End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me

    DAOCliente.LlenarCombo Me.cboClientes
    cli = Me.cboClientes.ItemData(cboClientes.ListIndex)
    If Me.chFecha.value Then
        Me.DTPicker1.Enabled = True
        Me.DTPicker2.Enabled = True
    Else
        Me.DTPicker1.Enabled = False
        Me.DTPicker2.Enabled = False
    End If
End Sub



Private Function llenarPeriodo()
    strsql = "select concat(month(p.fecha),'- ',year(p.fecha)) as periodo,sum((dp.ValorUnitario*dp.Cantidad) * (1-(p.descuento/100))) as totalCotizado from presupuestos p, detalle_presupuesto dp  where p.idCliente=" & cli & " and (p.estado=3 or p.estado=2)  and dp.idPresupuesto=p.id group by year(p.fecha),month(p.fecha) "
    totcot = 0
    Set rs = conectar.RSFactory(strsql)
    Me.lstPorPeriodo.ListItems.Clear
    While Not rs.EOF
        Set x = Me.lstPorPeriodo.ListItems.Add(, , rs!Periodo)
        x.SubItems(1) = Math.Round(rs!totalCotizado, 2)
        totcot = rs!totalCotizado + totcot
        rs.MoveNext

    Wend
    Me.lblTotCot = Math.Round(totcot, 2)
End Function

Private Function listado()
    strsql = "select p.detalle,sum((dp.ValorUnitario*dp.Cantidad)) as SubtotalPedido,p.descuento, sum((dp.ValorUnitario*dp.Cantidad) * (1-(p.descuento/100))) as totalPedido, if(p.estado=2,'Enviado','Procesado') as estado from presupuestos p, detalle_presupuesto dp  where p.idCliente=120 and (p.estado=3 or p.estado=2)  and dp.idPresupuesto=p.id group by p.id "
End Function

Private Sub listados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Me.listados.Sorted = True
    funciones.LstOrdenar Me.listados, ColumnHeader.index

End Sub
