VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdminFacturasAdeudadas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Facturas impagas..."
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Volver"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Vencidas"
      TabPicture(0)   =   "frmAdminFacturasAdeudadas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstFacturasEmitidasVencidas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "A Vencer"
      TabPicture(1)   =   "frmAdminFacturasAdeudadas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstFacturasEmitidasAVencer"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Informes"
      TabPicture(2)   =   "frmAdminFacturasAdeudadas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(3)=   "lbltotal"
      Tab(2).Control(4)=   "lstFacturasAPagar"
      Tab(2).Control(5)=   "desde"
      Tab(2).Control(6)=   "hasta"
      Tab(2).Control(7)=   "Command2"
      Tab(2).Control(8)=   "Command3"
      Tab(2).ControlCount=   9
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   -70800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Height          =   375
         Left            =   -72120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5280
         Width           =   1215
      End
      Begin MSComCtl2.MonthView hasta 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   5
         Top             =   3360
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   1
         StartOfWeek     =   52690945
         CurrentDate     =   39777
      End
      Begin MSComCtl2.MonthView desde 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   4
         Top             =   720
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   1
         StartOfWeek     =   52690945
         CurrentDate     =   39777
      End
      Begin MSComctlLib.ListView lstFacturasEmitidasVencidas 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   9128
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "O.C."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "F.P."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Vencimiento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Atraso"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Usuario "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Propuesta"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstFacturasEmitidasAVencer 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   9128
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "O.C."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "F.P."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fecha"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Vencimiento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Atraso"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Usuario "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Propuesta"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView lstFacturasAPagar 
         Height          =   4695
         Left            =   -72120
         TabIndex        =   8
         Top             =   480
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   8281
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cliente"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "O.C."
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label lbltotal 
         BackColor       =   &H00FFC0C0&
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
         Left            =   -61320
         TabIndex        =   12
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Total del período"
         Height          =   255
         Left            =   -62640
         TabIndex        =   11
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "HASTA"
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
         Left            =   -74880
         TabIndex        =   7
         Top             =   3120
         Width           =   2595
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "DESDE"
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
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   2595
      End
   End
   Begin VB.Menu facturas 
      Caption         =   "facturas"
      Visible         =   0   'False
      Begin VB.Menu numero 
         Caption         =   "numero"
         Enabled         =   0   'False
      End
      Begin VB.Menu fechaPago 
         Caption         =   "Definir nueva fecha..."
      End
      Begin VB.Menu ver 
         Caption         =   "Ver factura..."
      End
   End
End
Attribute VB_Name = "frmAdminFacturasAdeudadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_1 As Recordset
Dim rs As Recordset
Dim claseA As New classAdministracion
Dim marca
Dim claseSP As New classSignoplast
'Dim rs As Recordset
Dim strsql As String
Private Sub llenarLSTVencidas()
    On Error GoTo er1
    Me.lstFacturasEmitidasVencidas.ListItems.Clear
    strsql = "select f.propuesta,f.saldada,ft.TipoFactura,f.id,f.nroFactura,c.razon,f.tipo,f.idMoneda,f.FechaEmision,f.idUsuarioEmision,f.FormaPago,f.OrdenCompra,f.estado from AdminFacturas f inner join AdminConfigFacturas cf on f.tipoFactura=cf.id inner join clientes c on f.idCliente=c.id inner join AdminConfigFacturasTipos ft on cf.tipoFactura=ft.id WHERE f.saldada= 0 and f.estado=2 and date_add(fechaEmision,interval formaPago day)<'" & Format(Now, "yyyy-mm-dd") & "'  order by date_add(fechaEmision,interval formaPago day) asc"
    Set RS_1 = conectar.RSFactory(strsql)
    Dim esti As Integer
    Dim T As Integer
    Dim x As ListItem
    While Not RS_1.EOF
        Set x = Me.lstFacturasEmitidasVencidas.ListItems.Add(, , Format(RS_1!nroFactura, "0000") & " - " & RS_1!TipoFactura)
        x.Tag = RS_1!id
        x.SubItems(2) = RS_1!Razon
        x.SubItems(1) = funciones.queTipoFactura(RS_1!Tipo)
        x.SubItems(3) = funciones.queMoneda(RS_1!IdMoneda)

        T = RS_1!Tipo
        If T = 1 Then
            occ = "Factura " & RS_1!OrdenCompra
        Else
            occ = RS_1!OrdenCompra
        End If
        x.SubItems(4) = occ
        x.SubItems(5) = RS_1!FormaPago & " días FF"
        x.SubItems(6) = Format(RS_1!FechaEmision, "dd/mm/yyyy")
        est = funciones.estado_factura(RS_1!estado)
        esti = RS_1!estado
        estado_sald = RS_1!Saldada
        diasVenci = RS_1!FormaPago
        hoy = Format(Now, "dd/mm/yyyy")
        fechaF = RS_1!FechaEmision
        vencida = DateAdd("d", diasVenci, fechaF)
        venci = DateDiff("d", vencida, CDate(hoy))
        x.SubItems(7) = vencida
        x.SubItems(8) = venci & " días"
        x.SubItems(9) = claseSP.queUsuario(RS_1!idUsuarioEmision)

        fechaPro = Format(RS_1!Propuesta, "dd/mm/yyyy")
        If fechaPro = Empty Then
            fechaPro = "N/D"
        End If
        x.SubItems(10) = fechaPro


        x.ListSubItems(8).ForeColor = vbRed
        If marca = x Then
            x.Selected = True
            x.EnsureVisible
        End If

        If T = 0 Then
            x.ListSubItems(1).ForeColor = vbMagenta
        ElseIf T = 1 Then
            x.ListSubItems(1).ForeColor = vbBlue
        ElseIf T = 2 Then
            x.ListSubItems(1).ForeColor = ColorConstants.vbRed
        End If


        '
        RS_1.MoveNext
    Wend
    Exit Sub
er1:
    MsgBox Err.Description

End Sub


Private Sub llenarLSTAVencer()
    On Error Resume Next

    Me.lstFacturasEmitidasAVencer.ListItems.Clear
    strsql = "select f.propuesta,f.saldada,ft.TipoFactura,f.id,f.nroFactura,c.razon,f.tipo,f.idMoneda,f.FechaEmision,f.idUsuarioEmision,f.FormaPago,f.OrdenCompra,f.estado from AdminFacturas f inner join AdminConfigFacturas cf on f.tipoFactura=cf.id inner join clientes c on f.idCliente=c.id inner join AdminConfigFacturasTipos ft on cf.tipoFactura=ft.id WHERE f.saldada= 0 and f.estado=2 and date_add(fechaEmision,interval formaPago day)>'" & Format(Now, "yyyy-mm-dd") & "' group by date_add(fechaEmision,interval formaPago day) ASC"
    Set rs = conectar.RSFactory(strsql)
    Dim esti As Integer
    Dim T As Integer
    Dim x As ListItem
    While Not rs.EOF
        Set x = Me.lstFacturasEmitidasAVencer.ListItems.Add(, , Format(rs!nroFactura, "0000") & " - " & rs!TipoFactura)
        x.Tag = rs!id
        x.SubItems(2) = rs!Razon
        x.SubItems(1) = funciones.queTipoFactura(rs!Tipo)
        x.SubItems(3) = funciones.queMoneda(rs!IdMoneda)

        T = rs!Tipo
        If T = 1 Then
            occ = "Factura " & rs!OrdenCompra
        Else
            occ = rs!OrdenCompra
        End If
        x.SubItems(4) = occ
        x.SubItems(5) = rs!FormaPago & " días FF"
        x.SubItems(6) = Format(rs!FechaEmision, "dd/mm/yyyy")
        est = funciones.estado_factura(rs!estado)
        esti = rs!estado
        estado_sald = rs!Saldada
        'saldada = funciones.estado_factura_cobranza(rs!saldada)
        'x.SubItems(7) = saldada
        'x.ListSubItems(7).Tag = estado_sald
        diasVenci = rs!FormaPago
        hoy = Format(Now, "dd/mm/yyyy")
        fechaF = rs!FechaEmision
        vencida = DateAdd("d", diasVenci, fechaF)
        venci = DateDiff("d", CDate(hoy), vencida)

        fechaPro = Format(CDate(rs!Propuesta), "dd/mm/yyyy")
        If fechaPro = Empty Then
            fechaPro = "N/D"
        End If
        x.SubItems(10) = fechaPro

        x.SubItems(7) = vencida
        x.SubItems(8) = venci & " días"
        x.SubItems(9) = claseSP.queUsuario(rs!idUsuarioEmision)
        x.ListSubItems(8).ForeColor = vbBlue

        If marca = x Then
            x.Selected = True
            x.EnsureVisible
        End If

        If T = 0 Then
            x.ListSubItems(1).ForeColor = vbMagenta
        ElseIf T = 1 Then
            x.ListSubItems(1).ForeColor = vbBlue
        ElseIf T = 2 Then
            x.ListSubItems(1).ForeColor = ColorConstants.vbRed
        End If


        rs.MoveNext
    Wend

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    llenarLstRango
End Sub

Private Sub Command4_Click()
frmAdminFacturasAdeudadas2.Show
End Sub

Private Sub fechaPago_Click()
    idFactu = CLng(Numero.Tag)
    frmAdminFacturasAdeudadasRegistrar.idFactura = idFactu
    frmAdminFacturasAdeudadasRegistrar.Show 1
    llenarLSTVencidas
    llenarLSTAVencer

End Sub

Private Sub Form_Load()
FormHelper.Customize Me
    Me.desde = Now
    Me.hasta = Now
    llenarLSTVencidas
    llenarLSTAVencer
End Sub

Private Sub lstFacturasAPagar_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Me.lstFacturasAPagar.Sorted = True
    LstOrdenar Me.lstFacturasAPagar, CInt(ColumnHeader.index)
End Sub

Private Sub lstFacturasEmitidasAVencer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Me.lstFacturasEmitidasAVencer.Sorted = True
    LstOrdenar Me.lstFacturasEmitidasAVencer, CInt(ColumnHeader.index)
End Sub

Private Sub lstFacturasEmitidasAVencer_ItemClick(ByVal item As MSComctlLib.ListItem)
    ver.Tag = 1
    Me.Numero.Tag = CLng(Me.lstFacturasEmitidasAVencer.selectedItem.Tag)
End Sub

Private Sub lstFacturasEmitidasAVencer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstFacturasEmitidasAVencer.ListItems.count > 0 Then
        If Button = 2 Then
            Me.Numero.caption = "[ " & Me.lstFacturasEmitidasAVencer.selectedItem & " ]"
            Me.PopupMenu facturas
        End If
    End If

End Sub

Private Sub lstFacturasEmitidasVencidas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Me.lstFacturasEmitidasVencidas.Sorted = True
    LstOrdenar Me.lstFacturasEmitidasVencidas, CInt(ColumnHeader.index)
End Sub

Private Sub lstFacturasEmitidasVencidas_ItemClick(ByVal item As MSComctlLib.ListItem)
    ver.Tag = 0
    Me.Numero.Tag = CLng(Me.lstFacturasEmitidasVencidas.selectedItem.Tag)
End Sub

Private Sub lstFacturasEmitidasVencidas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.lstFacturasEmitidasVencidas.ListItems.count > 0 Then
        If Button = 2 Then
            Me.Numero.caption = "[ " & Me.lstFacturasEmitidasVencidas.selectedItem & " ]"
            Me.PopupMenu facturas
        End If
    End If
End Sub

Private Sub ver_Click()

    If ver.Tag = 0 Then
        idf = CLng(Me.lstFacturasEmitidasVencidas.selectedItem.Tag)
        canti = Me.lstFacturasEmitidasVencidas.ListItems.count
    ElseIf ver.Tag = 1 Then
        idf = CLng(Me.lstFacturasEmitidasAVencer.selectedItem.Tag)
        canti = Me.lstFacturasEmitidasAVencer.ListItems.count
    End If
    If canti > 0 Then
        frmAdminFacturasVer.idFactura = idf
        frmAdminFacturasVer.Show

    End If

End Sub



Private Sub llenarLstRango()
    On Error Resume Next
    Monto = 0
    Dim DESDE_ As Date
    Dim HASTA_ As Date
    Me.lstFacturasAPagar.ListItems.Clear
    DESDE_ = CDate(Me.desde)
    HASTA_ = CDate(Me.hasta)
    Dim Total As Double
    Dim IdMoneda As Integer
    Dim aa As Double
    Dim Razon As String
    Dim idCliente As Long
    Dim monto_nero As Double
    strsql = "select f.cambio_a_patron,f.propuesta,f.saldada,ft.TipoFactura,f.id,f.nroFactura,c.razon,f.tipo,f.idMoneda,f.FechaEmision,f.idUsuarioEmision,f.FormaPago,f.OrdenCompra,f.estado from AdminFacturas f inner join AdminConfigFacturas cf on f.tipoFactura=cf.id inner join clientes c on f.idCliente=c.id inner join AdminConfigFacturasTipos ft on cf.tipoFactura=ft.id WHERE f.saldada= 0 and f.estado=2 and propuesta>='" & Format(DESDE_, "yyyy-mm-dd") & "' and propuesta <='" & Format(HASTA_, "yyyy-mm-dd") & "'"
    Set rs = conectar.RSFactory(strsql)
    Dim esti As Integer
    Dim x As ListItem
    While Not rs.EOF
        Set x = Me.lstFacturasAPagar.ListItems.Add(, , Format(rs!nroFactura, "0000") & " - " & rs!TipoFactura)
        x.Tag = rs!id
        claseA.TotalFactura rs!id, Total, IdMoneda, Razon, idCliente

        x.SubItems(1) = rs!id '
        x.SubItems(2) = rs!Razon
        x.SubItems(3) = funciones.queMoneda(rs!IdMoneda)
        x.SubItems(4) = funciones.FormatearDecimales(Total, 2)
        x.SubItems(5) = rs!OrdenCompra
        x.SubItems(6) = rs!FechaEmision
        aa = claseA.realizaCambio(Total, rs!IdMoneda, 0)
        monto_nero = monto_nero + funciones.FormatearDecimales(aa, 2)
        Me.lblTotal = "AR$ " & funciones.FormatearDecimales(monto_nero, 2)



        rs.MoveNext
    Wend

End Sub

