VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesarrolloVerTiempos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Tiempos ejecutados..."
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPieza 
      Height          =   285
      Left            =   795
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   5100
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver"
      Default         =   -1  'True
      Height          =   255
      Left            =   6045
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin MSComctlLib.ListView lstDetallePieza 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6376
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod Tarea"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tarea"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cant Total"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cant Prom"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Agregada"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pieza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmDesarrolloVerTiempos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim claseS As New classStock
Dim idPieza As Long
Dim claseP As New classPlaneamiento
Private Sub Command1_Click()
    llenarLST
End Sub

Private Sub Command3_Click()
    If MsgBox("¿Está seguro de salir?", vbYesNo, "Confirmación") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    FormHelper.Customize Me
End Sub

Private Sub txtPieza_DblClick()
    frmListarStock_seleccion.Show 1
    Me.txtPieza = funciones.quePiezaElegidaDetalle
    idPieza = funciones.quePiezaElegida
    Command1_Click
End Sub


Private Sub llenarLST()
    Dim rs As Recordset
    Dim x As ListItem
    Dim q As String

    Me.lstDetallePieza.ListItems.Clear

    If claseS.EsConjunto(idPieza) = -1 Then    'no es conjunto
        'Set rs = conectar.RSFactory("select t.tarea,dmdo.codigo,dmdo.cantidad,dmdo.tiempo from desarrollo_mdo dmdo inner join tareas t on dmdo.codigo=t.id and dmdo.id_pieza=" & idPieza)
        q = "SELECT" _
            & " m.codigo AS cod_tarea," _
            & " t.tarea," _
            & " (m.tiempo * m.cantidad) AS Cant_Total," _
            & " (AVG(d.tiempo)/ AVG(d.cantidad_procesada)) AS Cant_Prom," _
            & " 0        AS Agregada" _
            & " FROM desarrollo_mdo m" _
            & " INNER JOIN stock s" _
            & " ON s.id = m.id_pieza" _
            & " LEFT JOIN PlaneamientoTiemposProcesos p" _
            & " ON p.idPieza = s.id" _
            & " AND m.codigo = p.codigoTarea" _
            & " And p.agregado = 0" _
            & " LEFT JOIN PlaneamientoTiemposProcesosDetalle d" _
            & " ON p.id = d.idTiemposProcesos" _
            & " LEFT JOIN tareas t" _
            & " ON t.id = m.codigo" _
            & " Where m.id_pieza = " & idPieza _
            & " GROUP BY m.codigo "
        q = q & " Union"
        q = q & " SELECT" _
            & " p.codigotarea," _
            & " t.tarea," _
            & " NULL," _
            & " (AVG(d.tiempo) / AVG(d.cantidad_procesada))," _
            & " 1" _
            & " FROM PlaneamientoTiemposProcesos p" _
            & " LEFT JOIN PlaneamientoTiemposProcesosDetalle d" _
            & " ON d.idTiemposProcesos = p.id" _
            & " INNER JOIN tareas t" _
            & " ON t.id = p.codigoTarea" _
            & " Where idPieza = " & idPieza _
            & " AND p.agregado = 1" _
            & " GROUP BY p.codigotarea"

        Set rs = conectar.RSFactory(q)
    Else
        'Set rs = conectar.RSFactory("select t.id as codigo, dmdo.cantidad,t.tarea,dmdo.tiempo from stockConjuntos sc inner join desarrollo_mdo dmdo on  sc.idPiezahija=dmdo.id_pieza inner join tareas t on dmdo.codigo=t.id where sc.idPiezaPadre=" & idPieza & " group by codigo")
    End If

    While Not rs.EOF
        '    Set X = Me.lstDetallePieza.ListItems.Add(, , rs!codigo)
        '        X.SubItems(1) = rs!Tarea
        '        X.SubItems(2) = rs!Cantidad
        '        X.SubItems(3) = funciones.formatearDecimales(rs!tiempo, 2)
        '        X.SubItems(4) = 0
        '        X.SubItems(5) = 0

        Set x = Me.lstDetallePieza.ListItems.Add(, , rs!cod_tarea)
        x.SubItems(1) = rs!Tarea
        x.SubItems(2) = IIf(IsNull(rs!Cant_Total), "", Format(Math.Round(rs!Cant_Total, 2), "0.00"))
        x.SubItems(3) = IIf(IsNull(rs!cant_prom), "", Format(Math.Round(rs!cant_prom, 2), "0.00"))
        x.SubItems(4) = IIf(rs!Agregada = 1, "Si", "No")
        rs.MoveNext
    Wend

    Set rs = Nothing
    'actualizarlst
End Sub
Private Sub actualizarlst()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Set rs = conectar.RSFactory("select tp.codigoTarea,avg(tpd.tiempo) as tiempo  from detalles_pedidos dp inner join PlaneamientoTiemposProcesos tp on tp.idDetallePedido=dp.id inner join PlaneamientoTiemposProcesosDetalle tpd on tpd.idTiemposProcesos=tp.id where dp.idPieza=" & idPieza & " group by tp.codigoTarea")
    While Not rs.EOF
        'recorro la lista para ver las coincidencias
        For x = 1 To Me.lstDetallePieza.ListItems.count
            If Me.lstDetallePieza.ListItems(x) = rs!codigoTarea Then
                Me.lstDetallePieza.ListItems(x).ListSubItems(5) = funciones.FormatearDecimales(rs!Tiempo, 2)
                strsql = "select tpd.legajo  from PlaneamientoTiemposProcesos tp join PlaneamientoTiemposProcesosDetalle tpd on tpd.idTiemposProcesos=tp.id where tp.idPieza=" & idPieza & "  and tp.codigoTarea=" & rs!codigoTarea & " group by legajo"
                Set rs2 = conectar.RSFactory(strsql)
                cnt = 0
                While Not rs2.EOF
                    cnt = cnt + 1
                    rs2.MoveNext
                Wend
                Me.lstDetallePieza.ListItems(x).ListSubItems(4) = cnt
                Exit For
            End If
        Next x
        rs.MoveNext
    Wend





    Set rs = Nothing
End Sub






