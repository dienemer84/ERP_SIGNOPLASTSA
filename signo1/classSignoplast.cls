VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classSignoplast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim clasea As New classArchivos
Dim strsql As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim vtabla As Integer
Dim Cantidad_rubros As Integer
Dim Cantidad_grupos As Integer
Dim vusu As Long
Public Function ejecutarComando(comando) As Boolean
    On Error GoTo err45
    ejecutarComando = True
    cn.execute comando
    Exit Function
err45:
    ejecutarComando = False
    MsgBox Err.Description
End Function


Public Function verSeleccionado(ByVal indice, idUsuario As Long, Optional actualizar As Boolean = False, Optional Valor = Empty) As String
    Select Case indice
            'panel de control y config
        Case 100: strs = "rootPanelControl"
            strtabla = "Config"
        Case 101: strs = "Activo"
            strtabla = "Config"
        Case 102: strs = "tableros"
            strtabla = "Config"
        Case 103: strs = "AgendaVer"
            strtabla = "Config"
        Case 104: strs = "AgendaModif"
            strtabla = "Config"
        Case 107: strs = "GrupoDefault"
            strtabla = "Config"
        Case 108: strs = "manoObraConfig"
            strtabla = "Config"
        Case 109: strs = "materialesConfig"
            strtabla = "Config"
        Case 110: strs = "verPrecios"
            strtabla = "Config"
        Case 110: strs = "verPrecios"
            strtabla = "Config"
        Case 111: strs = "PanelControlGeneral"
            strtabla = "Config"
        Case 112: strs = "ArchivosVer"
            strtabla = "Config"
        Case 113: strs = "ArchivosScannear"
            strtabla = "Config"
        Case 114: strs = "ArchivosCompra"
            strtabla = "Config"

        Case 115: strs = "VerUpdates"
            strtabla = "Config"

        Case 116: strs = "VerEventos"
            strtabla = "Config"



            'ventas
        Case 200: strs = "rootVentas"
            strtabla = "Ventas"
        Case 201: strs = "CotizControl"
            strtabla = "Ventas"
        Case 202: strs = "CotizConsultas"
            strtabla = "Ventas"
        Case 203: strs = "CotizAprobaciones"
            strtabla = "Ventas"
        Case 204: strs = "PedidosControl"
            strtabla = "Ventas"
        Case 205: strs = "PedidosConsultas"
            strtabla = "Ventas"
        Case 206: strs = "ClientesControl"
            strtabla = "Ventas"
        Case 207: strs = "ClientesConsultas"
            strtabla = "Ventas"
        Case 208: strs = "InfoPantalla"
            strtabla = "Ventas"
        Case 209: strs = "CotizModificar"
            strtabla = "Ventas"

            '---------------------------------'
            'Planeamiento

        Case 300: strs = "rootPlan"
            strtabla = "Plan"
        Case 301: strs = "infoPantalla"
            strtabla = "Plan"
        Case 302: strs = "OEcontrol"
            strtabla = "Plan"
        Case 303: strs = "OEconsultas"
            strtabla = "Plan"
        Case 304: strs = "OEaprobaciones"
            strtabla = "Plan"
        Case 305: strs = "OTcontrol"
            strtabla = "Plan"
        Case 306: strs = "OTconsultas"
            strtabla = "Plan"
        Case 307: strs = "OTaprobaciones"
            strtabla = "Plan"
        Case 308: strs = "SeguimientoGlobal"
            strtabla = "Plan"
        Case 309: strs = "SeguimientoRutas"
            strtabla = "Plan"
        Case 310: strs = "RemitosControl"
            strtabla = "Plan"
        Case 311: strs = "RemitosConsultas"
            strtabla = "Plan"
        Case 312: strs = "RemitosAprobaciones"
            strtabla = "Plan"
        Case 313: strs = "OTmodificar"
            strtabla = "Plan"
        Case 314: strs = "OEmodificar"
            strtabla = "Plan"

            '---------------------------------'
            'desarrollo

        Case 400: strs = "rootDesa"
            strtabla = "Desarrollo"
        Case 401: strs = "infoPantalla"
            strtabla = "Desarrollo"
        Case 402: strs = "Control"
            strtabla = "Desarrollo"
        Case 403: strs = "Consultas"
            strtabla = "Desarrollo"
        Case 404: strs = "ConsultaTiempos"
            strtabla = "Desarrollo"
        Case 405: strs = "manejoStock"
            strtabla = "Desarrollo"


            '--------------------------------
            'ventas


            'admin
        Case 500: strs = "rootAdmin"
            strtabla = "Administracion"
        Case 501: strs = "FacturasControl"
            strtabla = "Administracion"
        Case 502: strs = "FacturasConsulta"
            strtabla = "Administracion"
        Case 503: strs = "FacturasAprobacion"
            strtabla = "Administracion"
        Case 504: strs = "CobroControl"
            strtabla = "Administracion"
        Case 505: strs = "CobroConsulta"
            strtabla = "Administracion"
        Case 506: strs = "CobroAprobacion"
            strtabla = "Administracion"
        Case 507: strs = "SubdiariosControl"
            strtabla = "Administracion"
        Case 511: strs = "CCControl"
            strtabla = "Administracion"
        Case 512: strs = "infoPantalla"
            strtabla = "Administracion"
        Case 508: strs = "IIBBAdmin"
            strtabla = "Administracion"
        Case 509: strs = "IIBBAdminActualizar"
            strtabla = "Administracion"
        Case 510: strs = "centroCambio"
            strtabla = "Administracion"
        Case 513: strs = "informesCashFlow"
            strtabla = "Administracion"
        Case 514: strs = "informesVarios"
            strtabla = "Administracion"

        Case 515: strs = "cajaybancos"
            strtabla = "Administracion"
        Case 516: strs = "opcontrol"
            strtabla = "Administracion"
        Case 517: strs = "opconsultas"
            strtabla = "Administracion"

        Case 518: strs = "fpcontrol"
            strtabla = "Administracion"
        Case 519: strs = "fpconsultas"
            strtabla = "Administracion"
        Case 520: strs = "plancuentas"
            strtabla = "Administracion"
        Case 521: strs = "fpversolopropias"
            strtabla = "Administracion"


        Case 700: strs = "rootCompras"
            strtabla = "Compras"
        Case 701: strs = "infoPantalla"
            strtabla = "Compras"
        Case 702: strs = "ProveedoresControl"
            strtabla = "Compras"
        Case 703: strs = "ProveedoresConsultas"
            strtabla = "Compras"
        Case 704: strs = "RequesProcesar"
            strtabla = "Compras"
        Case 705: strs = "RequesControl"
            strtabla = "Compras"
        Case 706: strs = "RequesConsultas"
            strtabla = "Compras"
        Case 707: strs = "requesAprobaciones"
            strtabla = "Compras"
        Case 708: strs = "requesAnular"
            strtabla = "Compras"
        Case 709: strs = "poCrear"
            strtabla = "Compras"
        Case 710: strs = "poConsultar"
            strtabla = "Compras"
        Case 711: strs = "ocControl"
            strtabla = "Compras"
        Case 712: strs = "ocConsultas"
            strtabla = "Compras"
        Case 713: strs = "adminPrecios"
            strtabla = "Compras"
        Case 714: strs = "verPrecios"
            strtabla = "Compras"

        Case 800: strs = "Siniestro"
            strtabla = "rrhh"
        Case 801: strs = "InformeAccidente"
            strtabla = "rrhh"

    End Select




    permiso = strs
    tabla = strtabla

    If Not actualizar Then
        strsql = "select " & permiso & " as resu from sp_permisos." & tabla & " where idUsuario=" & idUsuario
        Me.ejecutar strsql
        'If Not rs.EOF Then verSeleccionado = rs!resu
        verSeleccionado = rs!resu
    Else
        strsql = "update  sp_permisos." & tabla & " set " & permiso & "=" & Valor & " where idUsuario=" & idUsuario
        cn.execute strsql
    End If


End Function

Property Get valorMDO() As Double
    valorMDO = rs!Valor
End Property
Public Property Let idUsuarioLogueado(idUsuario As Long)
    vusu = idUsuario
End Property
Public Property Get idUsuarioLogueado() As Long
    idUsuarioLogueado = vusu
End Property
Public Function cambiarPass(usuario, nuevo) As Boolean
    On Error GoTo rr
    cambiarPass = True
    cn.execute "update usuarios set password='" & nuevo & "' where id=" & usuario
    Exit Function
rr:
    cambiarPass = False
End Function
Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion

    Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function ejecutar(consulta) As Boolean
    On Error GoTo errje1
    ejecutar = True
    '    If rs.State = 1 Then rs.Close
    '    rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText

    Set rs = conectar.RSFactory(consulta)
    Exit Function
errje1:
    MsgBox "Se produjo un error: " & Err.Description
    ejecutar = False
End Function

Public Function Terminar()
    cn.Close
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Set cn = Nothing
End Function
Public Sub llenarLstRubros(lst As ListView)
    lst.ListItems.Clear
    lst.HideSelection = False
    lst.CheckBoxes = False
    lst.View = lvwReport
    lst.FullRowSelect = True
    lst.Gridlines = True
    lst.ColumnHeaders.Clear
    lst.Width = 3600
    lst.ColumnHeaders.Add(1).text = "Cód"
    lst.ColumnHeaders(1).Width = 700
    lst.ColumnHeaders.Add(2).text = "Rubro"
    lst.ColumnHeaders(2).Width = 2550
    lst.ColumnHeaders(2).Alignment = lvwColumnLeft
    Me.ejecutar "select * from rubros"
    Dim x As ListItem
    While Not rs.EOF
        Set x = lst.ListItems.Add(, , rs!id)
        x.SubItems(1) = rs!rubro
        rs.MoveNext
    Wend
End Sub
Public Function llenarLstGrupos(id_rubro As Integer, lst As ListView)
    lst.ListItems.Clear
    lst.HideSelection = False
    lst.CheckBoxes = False
    lst.View = lvwReport
    lst.FullRowSelect = True
    lst.Gridlines = True
    lst.ColumnHeaders.Clear
    lst.Width = 3600
    lst.ColumnHeaders.Add(1).text = "Cód"
    lst.ColumnHeaders(1).Width = 700
    lst.ColumnHeaders.Add(2).text = "Grupo"
    lst.ColumnHeaders(2).Width = 2550
    lst.ColumnHeaders(2).Alignment = lvwColumnLeft


    If IsNumeric(id_rubro) Then
        lst.ListItems.Clear

        Me.ejecutar "select * from grupos where id_rubro=" & id_rubro
        While Not rs.EOF
            Dim x As ListItem
            Dim i As Integer
            Set x = lst.ListItems.Add(, , rs!id)
            x.SubItems(1) = rs!Grupo
            rs.MoveNext
        Wend
        rs.Close
    Else
        lst.ListItems.Clear
    End If
End Function
Public Function llenarLstmateriales(id_rubro As Integer, id_grupo As Integer, lst As ListView)
    lst.ListItems.Clear
    lst.HideSelection = False
    lst.CheckBoxes = False
    lst.View = lvwReport
    lst.FullRowSelect = True
    lst.Gridlines = True
    lst.ColumnHeaders.Clear
    lst.Width = 6450
    lst.ColumnHeaders.Add(1).text = "id"
    lst.ColumnHeaders(1).Width = 0
    lst.ColumnHeaders.Add(2).text = "Descripción"
    lst.ColumnHeaders(2).Width = 4000
    lst.ColumnHeaders(2).Alignment = lvwColumnLeft
    lst.ColumnHeaders.Add(3).text = "Moneda"
    lst.ColumnHeaders(3).Width = 500
    lst.ColumnHeaders(3).Alignment = lvwColumnRight
    lst.ColumnHeaders.Add(4).text = "Valor"
    lst.ColumnHeaders(4).Width = 900
    lst.ColumnHeaders(4).Alignment = lvwColumnRight
    lst.ColumnHeaders.Add(5).text = "Unidad"
    lst.ColumnHeaders(5).Width = 700
    lst.ColumnHeaders(5).Alignment = lvwColumnRight


    If IsNumeric(id_rubro) And IsNumeric(id_grupo) Then
        lst.ListItems.Clear
        Me.ejecutar "select m.id_moneda,m.id_unidad,m.id,m.codigo,m.descripcion,m.espesor,m.valor_unitario from materiales m where m.id_rubro=" & id_rubro & " and m.id_grupo=" & id_grupo

        While Not rs.EOF
            Dim x As ListItem
            Dim i As Integer
            Set x = lst.ListItems.Add(, , rs!id)
            x.SubItems(1) = rs!codigo & " - " & rs!descripcion & " " & rs!Espesor & "mm"
            x.SubItems(2) = queMoneda(rs!id_moneda)
            x.SubItems(3) = funciones.RedondearDecimales(rs!valor_unitario, 3)
            x.SubItems(4) = queUnidad(rs!id_Unidad)
            rs.MoveNext
        Wend
        rs.Close
    Else
        lst.ListItems.Clear
    End If
End Function
Public Function llenar_combo_rubros(cbo As ComboBox)
    Set rs = New ADODB.Recordset
    cbo.Clear
    Me.ejecutar "select * from rubros"
    While Not rs.EOF
        cbo.AddItem rs!rubro
        cbo.ItemData(cbo.NewIndex) = rs!id
        rs.MoveNext
    Wend

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Function
Public Function add_rubro(txtRubro As String)
    ejecutar "insert into rubros (rubro) VALUES ('" & txtRubro & "')"
End Function
Public Function add_grupo(txtGrupo As String, id_rubro As Integer)
    Me.ejecutar "insert into grupos (grupo,id_rubro) VALUES ('" & txtGrupo & "',  " & id_rubro & ")"
End Function
Public Function del_grupo(idGrupo)
    Me.ejecutar "DELETE FROM grupos where id = " & idGrupo
End Function
Public Function del_rubro(idRubro)
    Me.ejecutar "DELETE FROM rubros where id = " & idRubro
    Me.ejecutar "DELETE FROM grupos where id_rubro = " & idRubro
End Function

Public Function ListaSectores() As Recordset
    If Me.ejecutar("select * from sectores order by sector") Then
        Set ListaSectores = rs
    Else
        MsgBox "Se produjo un error", vbCritical, "Error"
    End If
End Function

Public Function cambiarPermisos(panel, ventas, compras, planeamiento, desa, clientes, proveedores, seguimientos, segui, idUsuario, est, agVer, AgModif) As Boolean
    cambiarPermisos = True
    On Error GoTo errt
    cn.BeginTrans
    cn.execute "update usuariosPermisos set panel=" & panel & ",ventas=" & ventas & ",compras=" & compras & ",planeamiento=" & planeamiento & ",desarrollo=" & desa & ",clientes=" & clientes & ",proveedores=" & proveedores & ",seguimientos=" & segui & ",AgendaVer=" & agVer & ",AgendaModificar=" & AgModif & " where idUsuario=" & idUsuario
    cn.execute "update usuarios set estado=" & est & "  where id=" & idUsuario
    cn.CommitTrans
    Exit Function
errt:
    cambiarPermisos = False
    cn.RollbackTrans
End Function



Public Function queUsuario(idUsuario As Long) As String
    Me.ejecutar "select usuario from usuarios where id=" & idUsuario
    If Not rs.EOF And Not rs.BOF Then queUsuario = rs!usuario
End Function


Public Function agregarIncidencia(IdReferencia As Long, Origen As OrigenIncidencias, nIncidencia As String) As Boolean
    On Error GoTo err551
    agregarIncidencia = True
    usuario = funciones.getUser
    fec = funciones.datetimeFormateada(Now)
    Nota = nIncidencia
    cn.execute "insert into Incidencias (origen,usuario,fecha,nota,idReferencia) values (" & Origen & "," & usuario & ",'" & fec & "','" & Nota & "'," & IdReferencia & ")"


    Exit Function
err551:
    agregarIncidencia = False
End Function

Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub



Public Function llenarComboRubrosGrupos(cbo As ComboBox, idUsuario) As Long
    Set rs = New ADODB.Recordset
    cbo.Clear
    Me.ejecutar "select ug.id,ug.grupo from usuariosGruposDetalle ugd  inner join usuariosGrupos ug on ugd.idGrupo=ug.id where ugd.idUsuario=" & idUsuario
    c = 0
    While Not rs.EOF
        cbo.AddItem rs!Grupo
        cbo.ItemData(cbo.NewIndex) = rs!id
        rs.MoveNext
        c = c + 1
    Wend
    rs.Close
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If


End Function

Public Function cantidadGruposUsuario(idUsuario) As Long
    Set rs = New ADODB.Recordset

    Me.ejecutar "select count(ugd.id) as canti from usuariosGruposDetalle ugd  inner join usuariosGrupos ug on ugd.idGrupo=ug.id where ugd.idUsuario=" & idUsuario
    c = 0
    If Not rs.EOF And Not rs.BOF Then
        cantidadGruposUsuario = rs!canti
    End If

End Function




Public Function actualizarGrupos(idUsuario, lst As ListView) As Boolean
    On Error GoTo fuck
    actualizarGrupos = True
    cn.BeginTrans
    'borro todos los anteriores
    cn.execute "delete from usuariosGruposDetalle where idUsuario=" & idUsuario
    'agrego la nueva lista
    For x = 1 To lst.ListItems.count
        id = lst.ListItems(x).Tag
        cn.execute "insert into usuariosGruposDetalle (idUsuario,idGrupo) values (" & idUsuario & "," & id & ")"

    Next x

    cn.CommitTrans

    Exit Function
fuck:
    cn.RollbackTrans
    actualizarGrupos = False
End Function



Public Function VerificarSiHayActualizacion(Optional ByRef idnueva) As Boolean
    On Error GoTo err23
    maxi = App.Major
    mini = App.Minor
    revi = App.Revision
    If clasea.CompararConVersionActual(maxi, mini, revi, idnueva) Then
        VerificarSiHayActualizacion = True
    Else
        VerificarSiHayActualizacion = False
    End If

    Exit Function
err23:
    Err.Clear
End Function

Public Function actualizarSistema(idNuevo As Long) As Boolean
    If clasea.exportarArchivoSistema(idNuevo) Then
        aa = App.path & "\updater.exe " & App.path & "\abb.exe|" & App.path & "\abb.tmp"
        
        Shell aa, vbNormalFocus
        

        SalirForzado
        
        End
    Else
        MsgBox "Hubo un error al exportar la nueva versión del sistema.", vbCritical
    End If
End Function
