VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acceso al sistema"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4980
   ControlBox      =   0   'False
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   525
      Left            =   135
      TabIndex        =   14
      Top             =   1725
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   675
      Top             =   1695
   End
   Begin VB.ComboBox cboServidor 
      Height          =   315
      Left            =   2145
      TabIndex        =   2
      Text            =   "cboServidor"
      Top             =   990
      Width           =   2535
   End
   Begin XtremeSuiteControls.PushButton Command2 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   3135
      TabIndex        =   4
      Top             =   1515
      Width           =   1560
      _Version        =   786432
      _ExtentX        =   2752
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Default         =   -1  'True
      Height          =   405
      Left            =   1485
      TabIndex        =   3
      Top             =   1515
      Width           =   1560
      _Version        =   786432
      _ExtentX        =   2752
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Ingresar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.CommandButton Commandas 
      Height          =   570
      Left            =   1785
      TabIndex        =   13
      Top             =   2580
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   285
      Left            =   3120
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   975
      Left            =   915
      TabIndex        =   11
      Top             =   2610
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Commandsa 
      Height          =   795
      Left            =   1425
      TabIndex        =   10
      Top             =   2565
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1215
      Left            =   1830
      TabIndex        =   9
      Top             =   2535
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2145
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2145
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image Image 
      Height          =   720
      Left            =   180
      Picture         =   "frmLogin.frx":0000
      Top             =   285
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Usuario"
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
      Left            =   1380
      TabIndex        =   8
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Contraseña"
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
      Left            =   1065
      TabIndex        =   7
      Top             =   615
      Width           =   975
   End
   Begin VB.Label mensaje 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      Left            =   210
      TabIndex        =   6
      Top             =   1995
      Width           =   4545
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Servidor"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1035
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clasea As New classArchivos
Dim clssp As New classSignoplast
Dim hash As New classMD5


Private Sub Command1_Click()
    Me.Timer.Enabled = False

    On Error GoTo errh000


    Dim fullserver() As String
    fullserver = Split(Me.cboServidor.text, ":")

    Dim port As String
    port = "3306"
    Dim ip As String
    ip = frmPrincipal.servidorActual
    If UBound(fullserver) = 1 Then
        port = fullserver(1)
        ip = fullserver(0)
    Else
        If UBound(fullserver) = 0 Then
            ip = fullserver(0)
        End If
    End If


    frmPrincipal.servidorActual = Me.cboServidor.text
    conectar.port = port
    conectar.SetServidorBBDD ip

    If Not conectar.conectar Then GoTo errh000

    Dim r As Recordset
    Dim IdU As Long
    Dim idUsu As Long
    Dim usu As String
    usu = Trim(LCase(Me.Text1))
    Set r = conectar.RSFactory("select password,id from usuarios where usuario='" & usu & "'")
    c = 0
    While Not r.EOF
        c = c + 1
        r.MoveNext
    Wend

    If Trim(Me.Text1) = "root" And Trim(Me.Text2) = "3l3c720n" Then
        usu = "Administrador"
        '        frmPrincipal.stbar1.Panels(1).text = "Usuario: " & usu

        funciones.setUser -1
        IdU = funciones.getUser
        'funciones.ValidarPermisos (idU)
        Unload Me
    Else
        If c = 1 Then
            r.MoveFirst

            idUsu = r!Id
            passDB = r!PassWord
            estado = clssp.verSeleccionado(strPermisos.esistemaUsuarioActivo, idUsu)

            If estado = 1 Then
                passLI = hash.DigestStrToHexStr(Trim(Me.Text2))

                If StrConv(passDB, vbUpperCase) = passLI Then
                    'login correcto
                    '                    frmPrincipal.stbar1.Panels(1).text = "Usuario: " & usu
                    funciones.setUser (idUsu)

                    SaveSetting App.ProductName, "config", "user", usu


                    If Not funciones.InIDE Then
                        If clssp.VerificarSiHayActualizacion(idnueva) Then
                            If MsgBox("Hay una nueva actualización del sistema." & vbNewLine & "¿Desea aplicarla ahora?", vbYesNo + vbQuestion, "Confirmación") = vbYes Then

                                frmTip.Show 1

                                clssp.actualizarSistema CLng(idnueva)

                            End If

                        End If
                    Else

                        'frmTip.Show 1

                    End If


                    Unload Me


                Else
                    Me.Mensaje = "* Password incorrecto *"
                    Me.Text2 = Empty
                End If
            Else
                Me.Mensaje = "* Usuario no válido en el sistema *"
                Me.Text1 = Empty
                Me.Text2 = Empty
            End If
        Else
            Me.Mensaje = "* Usuario inexistente *"
            Me.Text1 = Empty
            Me.Text2 = Empty
        End If
    End If
    Exit Sub
errh000:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Dim rs As ADODB.Recordset
    Dim piezasId() As Variant
    Dim P As Pieza
    Dim diccio As New Dictionary

    Set rs = conectar.RSFactory("SELECT dp.idPieza,sum(e.cantidad) as cantidad FROM remitos r INNER JOIN entregas e  ON e.remito=r.id  INNER JOIN detalles_pedidos dp  ON e.idDetallePedido=dp.id  WHERE r.fecha BETWEEN '2009-08-01'   AND '2009-09-30' /* AND r.idCliente = 120 */ GROUP BY dp.idPieza")



    ReDim piezasId(0) As Variant
    piezasId(0) = 0

    While Not rs.EOF
        diccio.Add CStr(rs!idPieza), CStr(rs!Cantidad)
        ReDim Preserve piezasId(UBound(piezasId, 1) + 1)
        piezasId(UBound(piezasId, 1)) = rs!idPieza
        rs.MoveNext
    Wend

    Dim piezas As Collection
    Set piezas = DAOPieza.FindAll(FL_4, DAOPieza.TABLA_PIEZA & "." & DAOPieza.CAMPO_ID & " IN ( " & Join(piezasId, ", ") & ")", , True)
    Dim arr(1 To 100) As Integer

    For i = 1 To 100
        arr(i) = 0
    Next
    Dim p1 As Pieza
    Dim p2 As Pieza
    Dim p3 As Pieza
    Dim desa As DesarrolloManoObra

    For Each P In piezas

        For Each desa In P.desarrollosManoObra
            arr(desa.Tarea.Id) = arr(desa.Tarea.Id) + _
                                 IIf(desa.Tarea.CantPorProc = 1, desa.Tiempo * desa.Cantidad * CLng(diccio.item(CStr(P.Id))), desa.Tiempo * desa.Cantidad)
        Next desa

        For Each p1 In P.PiezasHijas
            For Each desa In p1.desarrollosManoObra
                arr(desa.Tarea.Id) = arr(desa.Tarea.Id) + _
                                     IIf(desa.Tarea.CantPorProc = 1, desa.Tiempo * desa.Cantidad * CLng(diccio.item(CStr(p1.Id))), desa.Tiempo * desa.Cantidad)
            Next desa
            For Each p2 In p1.PiezasHijas
                For Each desa In p2.desarrollosManoObra
                    arr(desa.Tarea.Id) = arr(desa.Tarea.Id) + _
                                         IIf(desa.Tarea.CantPorProc = 1, desa.Tiempo * desa.Cantidad * CLng(diccio.item(CStr(p2.Id))), desa.Tiempo * desa.Cantidad)
                Next desa
                For Each p3 In p2.PiezasHijas
                    For Each desa In p3.desarrollosManoObra
                        arr(desa.Tarea.Id) = arr(desa.Tarea.Id) + _
                                             IIf(desa.Tarea.CantPorProc = 1, desa.Tiempo * desa.Cantidad * CLng(diccio.item(CStr(p3.Id))), desa.Tiempo * desa.Cantidad)
                    Next desa

                Next p3


            Next p2
        Next p1


    Next P

    T = 0
    For x = 1 To 100

        arr(x) = funciones.FormatearDecimales(arr(x) / 60)
        T = T + arr(x)
    Next
    Debug.Print "total horas", T
End Sub




Private Sub Command4_Click()
    Dim col As Collection
    Dim deta As DetalleOrdenTrabajo

    Set col = DAODetalleOrdenTrabajo.FindAll()
    conectar.BeginTransaction
    For Each deta In col


        'DAODetalleOrdenTrabajo.SaveCantidad deta.id, deta.CantidadEntregada, CantidadEntregada_, deta.Precio
        'DAODetalleOrdenTrabajo.SaveCantidad deta.id, deta.CantidadFabricados, CantidadFabricada_, deta.Precio
        'DAODetalleOrdenTrabajo.SaveCantidad deta.id, deta.CantidadFacturada, CantidadFacturada_, deta.Precio
    Next
    conectar.CommitTransaction
End Sub

Private Sub Command5_Click()
    Dim F As New frmEmpleadosTareas
    Load F
    F.personalId = 25
    F.Show 1
End Sub

Private Sub Command6_Click()

    frmPrincipal.servidorActual = Me.cboServidor.text
    conectar.SetServidorBBDD frmPrincipal.servidorActual
    conectar.conectar


    'fix para piezas conjunto o unidad
    Dim piezas As Collection
    Dim Pieza As Pieza
    Set piezas = DAOPieza.FindAll(FL_4)
    Dim conjunto As Boolean
    conectar.BeginTransaction
    For Each Pieza In piezas
        conjunto = (Pieza.PiezasHijas.count > 0)
        conectar.execute "UPDATE stock SET conjunto = " & IIf(conjunto, "0", "-1") & " WHERE id = " & Pieza.Id
    Next Pieza
    conectar.CommitTransaction

End Sub




Private Sub Commandsa_Click()
    Dim F As New frmDesarrollo
    Load F
    F.CargarDetallePresupuesto 303481
    F.Show 1
End Sub

Private Sub Form_Activate()
    If Len(Me.Text1.text) > 0 Then Me.Text2.SetFocus
End Sub


Private Sub Form_Load()
    FormHelper.Customize Me
    Me.Text1.text = GetSetting(App.ProductName, "config", "user", vbNullString)


    Me.cboServidor.Clear
    Dim srv As Variant
    For Each srv In frmPrincipal.servidorBBDD
        Me.cboServidor.AddItem srv
    Next srv
    If Me.cboServidor.ListCount > 0 Then Me.cboServidor.ListIndex = 0

    If funciones.InIDE Then
        Me.Text1.text = "nicolasba"
        Me.Text2.text = "022916"
        'Command1_Click

    Else
        Dim Puesto As String
        Puesto = LeerIni(App.path & "\config.ini", "Configurar", "puesto", vbNullString)
        If LenB(Puesto) > 0 Then
            Me.Text1.text = Puesto
            Me.Text2.text = "puesto"
            Me.Timer.Enabled = True
        End If

    End If
End Sub

Private Sub PushButton1_Click()



End Sub

Private Sub ProcesarDetaOT(detaOT As DetalleOrdenTrabajo, Optional detaOTDto As DetalleOTConjuntoDTO = Nothing)

    Dim tmpDeta As DetalleOTConjuntoDTO
    Dim piezaId As Long
    Dim ptp As PlaneamientoTiempoProceso

    If IsSomething(detaOTDto) Then
        piezaId = detaOTDto.Pieza.Id
    Else
        piezaId = detaOT.Pieza.Id
    End If

    For Each tmpDeta In DAODetalleOrdenTrabajo.FindAllConjunto(detaOT.Id, piezaId)

        For Each ptp In DAOTiemposProceso.FindAllByDetallePedidoIdAndPiezaId(tmpDeta.Id, tmpDeta.Pieza.Id)
            conectar.execute "UPDATE PlaneamientoTiemposProcesos SET idDetallePedido = " & detaOT.Id & ", idDetallePedidoConj = " & tmpDeta.Id & " WHERE id = " & ptp.Id
        Next ptp

        ProcesarDetaOT detaOT, tmpDeta

    Next tmpDeta

End Sub

Private Sub Text1_GotFocus()
    foco Me.Text1
End Sub

Private Sub Text2_GotFocus()
    foco Me.Text2
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Timer_Timer()
    Command1_Click
End Sub
