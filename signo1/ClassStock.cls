VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim clasea As New classAdministracion
Private rs_mm As Recordset
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim vCantidad As Integer
Dim strsql As String
Dim claseP As New classPlaneamiento

Property Get descripcion() As String
    descripcion = rs!descripcion
End Property
Property Get rubro() As String
    rubro = rs!rubro
End Property
Property Get Grupo() As String
    Grupo = rs!Grupo
End Property
Property Get Espesor() As String
    Espesor = rs!Espesor
End Property
Property Get idUnidad() As Integer
    idUnidad = rs!id_Unidad
End Property
Property Get idCliente() As Long
    idCliente = rs!id_cliente
End Property

Private Sub Class_Initialize()
    On Error GoTo err22
    Set cn = conectar.obternerConexion    'New ADODB.Connection
    Set rs = conectar.RSFactory("select count(id) as canti from stock")
    vCantidad = rs!canti

    Exit Sub
err22:
    MsgBox Err.Description
End Sub


Public Function calcular_valor_materiales(ByVal Id As Integer, ByRef kgs As Double, ByRef precio2 As Double, ByVal idMonedaDestino As Long)
'Busco detalles pieza

    Dim strsql As String
    strsql = "select m.fecha_valor,c.dolar,m.valor_unitario,m.id_moneda,dm.largoTerm,dm.anchoTerm,dm.largo, dm.ancho, dm.scrap, dm.cantidad, m.pesoxunidad, m.id as idmat,m.id_unidad from desarrollo_material dm, materiales m, stock s, configuracion c where dm.id_material=m.id  and s.id=" & Id & " and dm.id_pieza=" & Id
    Me.ejecutar2 strsql
    prec = 0
    While Not rs2.EOF
        'dolar = rs!dolar
        PesoXUnidad = rs2!PesoXUnidad
        valor_unitario = rs2!valor_unitario

        moneda = rs2!id_moneda  'moneda origen material (en que moneda está costeado)

        y = rs2!Largo   'CAMBIADOSSSSS 28-3-8
        x = rs2!Ancho
        y1 = rs2!LargoTerm
        x1 = rs2!AnchoTerm
        Scrap = rs2!Scrap
        ccantidad = rs2!Cantidad
        idmat = rs2!idmat
        Dim tmpKg As Double

        tmpKg = Me.calcularM2MLKGMaterial(x1, y1, idmat, Scrap, x, y, ccantidad, Kg, m2ml, Pieza, costo, idMonedaDestino)
        pr = pr + costo

        kgs = kgs + tmpKg
        rs2.MoveNext
    Wend
    precio2 = pr
End Function

Public Sub ejecutar(consulta As String)
    On Error GoTo err6
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, CN, adOpenDynamic, adLockOptimistic, adCmdText

    Set rs = conectar.RSFactory(consulta)
    Exit Sub
err6:
    MsgBox "Se produjo un error: " & Err.Description
End Sub
Public Sub ejecutar2(consulta As String)
    On Error GoTo err62

    Set rs2 = conectar.RSFactory(consulta)
    Exit Sub
err62:
    MsgBox "Se produjo un error: " & Err.Description
End Sub

Public Sub llenar_lista_stock(lst As ListView, cliente As Long, filtro As String, Optional marcado, Optional Cantidad, Optional elegir As Boolean = False)
    On Error Resume Next
    lst.ListItems.Clear
    If Trim(cliente) = -1 Then
        'si son todos los clientes.
        If Trim(filtro) = Empty Then
            If elegir Then
                strsql = "select S.conjunto,S.estado,S.id,S.detalle, S.cantidad , c.razon From stock S inner join clientes c on S.id_cliente=c.id  WHERE S.estado=1  group by S.id order by S.detalle asc"
            Else
                strsql = "select S.conjunto,S.estado,S.id,S.detalle, S.cantidad , c.razon From stock S inner join clientes c on S.id_cliente=c.id  WHERE S.estado=1  group by S.id order by S.detalle asc"
            End If
        Else
            If elegir Then
                strsql = "select S.conjunto,S.estado,S.id,S.detalle, S.cantidad , c.razon From stock S inner join clientes c on S.id_cliente=c.id  WHERE S.detalle LIKE '%" & Trim(filtro) & "%' AND S.estado=1 group by S.id order by S.detalle asc"
            Else
                strsql = "select S.conjunto,S.estado,S.id,S.detalle, S.cantidad , c.razon From stock S inner join clientes c on S.id_cliente=c.id  WHERE S.detalle LIKE '%" & Trim(filtro) & "%' AND S.estado=1 group by S.id order by S.detalle asc"
            End If
        End If
    ElseIf Trim(cliente) = -2 Then    'si quieren ver los registros marcados para eliminar
        If elegir Then
            strsql = "SELECT S.conjunto,S.estado,S.id,S.detalle, S.cantidad , c.razon From stock S inner join clientes c on S.id_cliente=c.id WHERE S.estado=0 group by S.id order by S.detalle asc"
        Else
            strsql = "SELECT S.conjunto,S.estado,S.id,S.detalle, S.cantidad , c.razon From stock S inner join clientes c on S.id_cliente=c.id WHERE S.estado=0 group by S.id order by S.detalle asc"
        End If
    Else
        'si hay filtro de cliente
        If Trim(filtro) = Empty Then
            If elegir Then
                strsql = "select S.conjunto,S.estado,S.id,S.detalle,  S.Cantidad,  c.razon From stock S inner join clientes c on S.id_cliente=c.id AND S.id_cliente = " & cliente & " AND S.estado=1 group by S.id order by S.detalle asc"
            Else
                strsql = "select S.conjunto,S.estado,S.id,S.detalle,  S.Cantidad,  c.razon From stock S inner join clientes c on S.id_cliente=c.id AND S.id_cliente = " & cliente & " AND S.estado=1 group by S.id order by S.detalle asc"
            End If

        Else
            If elegir Then
                strsql = "select S.conjunto,S.estado,S.id,S.detalle,  S.Cantidad,  c.razon From stock S inner join clientes c on S.id_cliente=c.id AND S.id_cliente = " & cliente & "  WHERE S.detalle LIKE '%" & Trim(filtro) & "%' AND S.estado=1 group by S.id order by S.detalle asc"
            Else
                strsql = "select S.conjunto,S.estado,S.id,S.detalle,  S.Cantidad,  c.razon From stock S inner join clientes c on S.id_cliente=c.id AND S.id_cliente = " & cliente & "  WHERE S.detalle LIKE '%" & Trim(filtro) & "%' AND S.estado=1 group by S.id order by S.detalle asc"
            End If

        End If

    End If
    Dim r_inci As Recordset
    Dim r_archi As Recordset

    Me.ejecutar strsql
    Dim x As ListItem
    If rs.EOF = False Then
        lst.Visible = False

        While Not rs.EOF
            Set x = lst.ListItems.Add(, , rs!Id)
            x.Tag = rs!Id
            x.SubItems(1) = rs!razon
            x.SubItems(2) = UCase(rs!detalle)
            x.SubItems(3) = rs!Cantidad
            x.SubItems(4) = rs!estado
            x.SubItems(5) = funciones.EsConjunto(rs!conjunto)
            x.SubItems(6) = rs!conjunto

            rs.MoveNext

            If x = marcado Then
                x.Selected = True
                x.EnsureVisible
            End If
        Wend

        lst.Visible = True
    End If

    Cantidad = CLng(lst.ListItems.count) & " Registros"
End Sub



Public Sub modifica_Stock(Id As Long, operacion As Integer, Nota, Cantidad, Optional ubicacion = Empty)
    On Error GoTo err5
    'operacion
    '0- ingreso a stock
    '1- alta en stock desde ot
    '2- Baja de stock
    '3-baja en stock desde oe
    '4-baja
    'nota
    '0- ingreso existoso
    '-1 defectuoso/egreso sin causa
    '>0 nro. orden de trabajo
    'CN.BeginTrans
    If operacion = 1 Then charO = "+" Else charO = "-"
    Me.ejecutar "select * from stock where id =" & Id
    If IsEmpty(ubicacion) Then
        cn.execute "update stock set cantidad = cantidad" & charO & Cantidad & " where id=" & Id
    Else
        cn.execute "update stock set detalle_stock='" & UCase(ubicacion) & "',cantidad = cantidad" & charO & Cantidad & " where id=" & Id
    End If

    cha = Format(Date, "yyyy/mm/dd")
    cn.execute "insert into movimiento_stock (id_pieza,cantidad,operacion,fecha,nota)values (" & Id & "," & Cantidad & "," & operacion & ",'" & cha & "'," & Nota & ")"
    'CN.CommitTrans
    Exit Sub
err5:
    'CN.RollbackTrans
End Sub
Function llenar_historico_stock(Id As Integer, lst As ListView)
    Dim strsql As String
    strsql = "select * from movimiento_stock where id_pieza = " & Id
    Me.ejecutar strsql
    Dim operacion(0 To 4)

    operacion(0) = "Ingreso"
    operacion(1) = "Alta"
    operacion(2) = "Baja"
    operacion(3) = "Baja"
    operacion(4) = "Baja"
    lst.ListItems.Clear

    While Not rs.EOF
        If rs!Nota = 0 Then
            Nota = "Ingreso exitoso"
        ElseIf rs!Nota = -1 Then
            Nota = "Defectuoso/sin causa"
        ElseIf rs!Nota = -2 Then
            Nota = "Manual"
        ElseIf rs!Nota = -4 Then
            Nota = "Devolucion a stock"
        Else
            If rs!operacion = 1 Then
                Nota = "O/T " & rs!Nota
            ElseIf rs!operacion = 3 Then
                Nota = "O/E " & rs!Nota
            ElseIf rs!operacion = 4 Then
                Nota = "O/T " & rs!Nota
            End If
        End If

        Set x = lst.ListItems.Add(, , rs!Id)
        x.SubItems(1) = rs!Cantidad
        x.SubItems(2) = operacion(rs!operacion)

        x.SubItems(3) = Nota
        x.SubItems(4) = rs!FEcha

        rs.MoveNext

    Wend

    'End If
End Function

Public Function cambiar_estado(Id As Integer, accion As Integer) As Boolean
    Dim strsql As String
    strsql = "select estado from stock where id = " & Id
    Me.ejecutar strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
        estado = rs!estado
        If accion = 1 Then    'si manda a deshabilitar
            If estado = 0 Then
                MsgBox "Esta pieza ya fue inhabilitada", vbCritical, "Error"
            Else
                cn.execute "update stock set estado=0 where id=" & Id
            End If
        ElseIf accion = 0 Then
            If estado = 1 Then
                MsgBox "Esta pieza ya fue habilitada", vbCritical, "Error"
            Else
                cn.execute "update stock set estado=1 where id=" & Id
            End If

        End If


    Else
        MsgBox "Se produjo un error de redundancia de datos", vbCritical, "Error"
    End If

End Function


Function buscar_pieza(str As String) As Integer
    Dim strsql As String
    strsql = "Select id from stock where detalle='" & str & "'"
    Me.ejecutar strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
        buscar_pieza = rs!Id
    Else
        buscar_pieza = -1    'error
    End If

End Function

Public Function detalle_pieza(Id As Long) As String
    Dim strsql As String
    strsql = "Select detalle from stock where id=" & Id
    Me.ejecutar strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    Dim A As String

    If c = 1 Then
        rs.MoveFirst
        A = rs!detalle
    Else
        A = -1    'error
    End If
    detalle_pieza = A
End Function
Public Function cLiente_pieza(Id As Integer) As String
    Dim strsql As String
    strsql = "Select c.razon AS AA from stock s, clientes c where s.id_cliente=c.id and s.id=" & Id
    Me.ejecutar strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    Dim A As String

    If c = 1 Then
        rs.MoveFirst
        A = rs!aa
    Else
        A = -1    'error
    End If
    cLiente_pieza = A
End Function


Public Function TiemposPieza(idPieza As Integer, Cantidad As Long) As Double
    strsql = "select sum( if(t.cantxproc>0,dm.cantidad*dm.tiempo*" & Cantidad & ",dm.cantidad*dm.tiempo)) as tota from desarrollo_mdo dm,tareas t, stock s where id_pieza=" & idPieza & " and t.id=dm.codigo and s.id=dm.id_pieza "


    Me.ejecutar strsql
    tot = 0
    While Not rs.EOF
        tot = tot + rs!tota
        rs.MoveNext
    Wend
    If Not IsNumeric(tot) Then tot = 0
    TiemposPieza = tot

End Function

Function calcularM2MLKGMaterial(ByVal x2, ByVal y2, ByVal IdMaterial, ByVal Scrap, ByVal x, ByVal y, ByVal Cant As Double, ByRef Kg, ByRef m2ml, ByRef Pieza, ByRef costo, ByVal idMonedaDestino As Long) As Double

    Dim Precio As Double
    Dim id_moneda_origen As Long
    'strsql = "select s.detalle,dm.largoTerm,dm.AnchoTerm,dm.scrap,dm.id_pieza,dm.largo,dm.ancho,dm.cantidad,dm.id_material,m.id_unidad,m.pesoxunidad, vm.valor_unitario,s.id from stock.stock s inner join stock.desarrollo_material dm on dm.id_pieza=s.id inner join signoplast.materiales m on dm.id_material=m.id inner join signoplast.valores_MATERIALES vm on vm.id_material=m.id and m.id=" & IdMaterial

    strsql = "select m.fecha_valor, m.id_moneda as monedaOrigen,m.valor_unitario, m.pesoxunidad, m.id_unidad from materiales m where m.id=" & IdMaterial
    Me.ejecutar strsql
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c > 0 Then
        rs.MoveFirst
        Dim cantidad_h As Double
        cantidad_h = Cant
        anchop = x    'cambiadosss
        largop = y
        anchoT = x2
        largoT = y2
        unidad = rs!id_Unidad

        id_moneda_origen = rs!monedaOrigen

        valorxunidad = rs!valor_unitario
        PesoXUnidad = rs!PesoXUnidad

        Scrap = (Scrap / 100) + 1

        If unidad = 3 Then    'ml

            Kg = Math.Round(PesoXUnidad * (largoT / 1000) * cantidad_h, 3)
            m2ml = Math.Round(largop / 1000 * cantidad_h, 3)
            Precio = valorxunidad * PesoXUnidad
            'realizo el cambio a la moneda que pida el origen de costos
            'desde la moneda en que está grabado el material
            Precio = clasea.realizaCambio(Precio, id_moneda_origen, idMonedaDestino)
            prec = ((Precio * largop * Scrap * 2) / 1000) * cantidad_h
            prec = prec / 2    'verificar!!!!!!!!!!!!!!!!!!!1

            Pieza = cantidad_h & "x" & largoT
            piezaTerm2 = cantidad_h & "x" & largop
        ElseIf unidad = 1 Then    'kg
            Kg = Math.Round(PesoXUnidad * cantidad_h, 3)
            m2ml = 0

            Precio = valorxunidad * PesoXUnidad
            Precio = clasea.realizaCambio(Precio, id_moneda_origen, idMonedaDestino)
            prec = (Precio * cantidad_h)

            Pieza = cantidad_h
            piezaTerm2 = cantidad_h
        ElseIf unidad = 2 Then    'm2
            Kg = Math.Round(PesoXUnidad * ((anchoT * Largo) / 1000000) * cantidad_h, 3)
            m2ml = Math.Round((anchop * largop) / 1000000 * cantidad_h, 3)
            Precio = valorxunidad * PesoXUnidad
            Precio = clasea.realizaCambio(Precio, id_moneda_origen, idMonedaDestino)

            prec = Precio * ((cantidad_h * x * y * Scrap) / 1000000)
            Pieza = cantidad_h & "x" & largoT & "x" & anchoT
            piezaTerm2 = cantidad_h & "x" & largop & "x" & anchop
        ElseIf unidad = 4 Then    'uni

            m2ml = 0
            Precio = valorxunidad
            Precio = clasea.realizaCambio(Precio, id_moneda_origen, idMonedaDestino)
            prec = Precio * cantidad_h
            Pieza = cantidad_h
            Kg = Pieza
            piezaTerm2 = cantidad_h
        End If
        cxh = funciones.cantxhoja(x2, y2, x, y)
        If cxh = 0 Then cxh = 1
        costo = prec / cxh
    End If

    calcularM2MLKGMaterial = Kg
End Function

Public Function CopiarPieza(idOrigen, NombreDestino) As Boolean
    On Error GoTo err00
    Dim ultid As Long
    CopiarPieza = True
    'selecciono todos los datos necesarios de la pieza original

    Me.ejecutar "select id_cliente from stock where id=" & idOrigen
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend
    If c = 1 Then
        rs.MoveFirst
        idcli = rs!id_cliente
        cn.BeginTrans
        tra = True
        'alta en stock.
        cn.execute "insert into stock (detalle,id_cliente,cantidad) VALUES ('" & UCase(NombreDestino) & "'," & idcli & ",0)"


        Me.ejecutar "select last_insert_id() as idd from stock"
        ultid = rs!idd

        'alta en desarrollo materiales del elemento
        Me.ejecutar "select * from desarrollo_material where id_pieza=" & idOrigen
        While Not rs.EOF
            Scrap = rs!Scrap
            id_material = rs!id_material
            Largo = rs!Largo
            Ancho = rs!Ancho
            largoT = rs!LargoTerm
            anchoT = rs!AnchoTerm
            Cantidad = rs!Cantidad
            cn.execute "insert into desarrollo_material (id_pieza,scrap,largo,ancho,largoTerm,AnchoTerm,id_material,cantidad) values (" & CInt(ultid) & "," & CDbl(Scrap) & "," & CDbl(Largo) & "," & CDbl(Ancho) & " ," & CDbl(largoT) & "," & CDbl(anchoT) & ", " & id_material & "," & Cantidad & ")"
            rs.MoveNext
        Wend
        'alta en desarrollo de mdo del elemento
        Me.ejecutar "select * from desarrollo_mdo where id_pieza=" & idOrigen
        While Not rs.EOF
            id_mdo = rs!codigo
            Cant = rs!Cantidad
            Tiempo = rs!Tiempo
            cn.execute "insert into desarrollo_mdo (id_pieza,codigo,cantidad,tiempo) values (" & CInt(ultid) & "," & CInt(id_mdo) & "," & CDbl(Cant) & "," & CDbl(Tiempo) & ")"
            rs.MoveNext
        Wend
        tra = False

        cn.CommitTrans
        Me.modifica_Stock ultid, 0, 0, 0, ""    'doy d alta en bbdd historicos de stock
    End If
    Exit Function
err00:
    If tra = Tru Then cn.RollbackTrans
    CopiarPieza = False
End Function


Public Function definirConjunto2(detalle As String, cliente As Long, lista As ListView) As Boolean
    On Error GoTo err55
    definirConjunto = True
    Dim tra As Boolean
    cn.BeginTrans
    tra = True
    'inserto el elemento en stock
    cn.execute "insert into stock (detalle,id_cliente,conjunto) values ('" & detalle & "'," & cliente & ",0 )"
    Me.ejecutar "select last_insert_id() as ultid from stock"
    idPiezaPadre = rs!ultid
    For K = 1 To lista.ListItems.count
        Cantidad = lista.ListItems(K)
        idPiezaHija = lista.ListItems(K).ListSubItems(2)
        cn.execute "insert into stockConjuntos (idPiezaPadre, idPiezaHija, cantidad) values (" & idPiezaPadre & "," & idPiezaHija & "," & Cantidad & ")"
    Next K
    Me.crearHistoricoPieza2 idPiezaPadre, "CONJUNTO AGREGADO"
    cn.CommitTrans
    Exit Function
err55:
    If tra Then cn.RollbackTrans
    definirConjunto = False
End Function
Public Function modificarConjunto(detalle As String, cliente As Long, lista As ListView, idConjunto As Long) As Boolean
    On Error GoTo err55
    modificarConjunto = True
    Dim tra As Boolean
    cn.BeginTrans
    tra = True
    'updateo la pieza padre en tabla stock
    cn.execute "update stock  set detalle='" & detalle & "',id_cliente=" & cliente & " ,conjunto=0 where id=" & idConjunto
    'elimino la definicion anterior del conjunto
    cn.execute "delete from stockConjuntos where idPiezaPadre=" & idConjunto
    'agrego nuevos datos
    For K = 1 To lista.ListItems.count
        Cantidad = lista.ListItems(K)
        idPiezaHija = lista.ListItems(K).ListSubItems(2)
        cn.execute "insert into stockConjuntos (idPiezaPadre, idPiezaHija, cantidad) values (" & idConjunto & "," & idPiezaHija & "," & Cantidad & ")"
    Next K
    Me.crearHistoricoPieza2 idConjunto, "CONJUNTO MODIFICADO"
    cn.CommitTrans
    Exit Function
err55:
    If tra Then cn.RollbackTrans
    modificarConjunto = False
End Function


Public Function exporta(idPieza As Long, lst As ListView) As Boolean
    Me.ejecutar "select s.detalle, c.razon from stock s, clientes c where s.id=" & idPieza & " and s.id_cliente=c.id"
    cliente = rs!razon
    Pieza = rs!detalle
    On Error GoTo errEXCEL
    Dim xlb As New Excel.Workbook
    Dim xla As New Excel.Worksheet
    Dim xls As New Excel.Application



    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate


    With xla
        .Columns("A").HorizontalAlignment = xlHAlignCenter
        .Columns("B").HorizontalAlignment = xlHAlignCenter
        .Columns("C").HorizontalAlignment = xlHAlignCenter
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("E").HorizontalAlignment = xlHAlignCenter
        .Range("A1:A4").HorizontalAlignment = xlHAlignLeft

        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 18
        .Range("a1:b1").Merge
        .Range("b2:b3").HorizontalAlignment = xlHAlignLeft
        .Range("a1:b1").HorizontalAlignment = xlHAlignCenter
        .Range("C1:E1").HorizontalAlignment = xlHAlignCenter
        .Range("C2:E2").HorizontalAlignment = xlHAlignCenter
        .Range("C3:E3").HorizontalAlignment = xlHAlignCenter
        .Range("C4:E4").HorizontalAlignment = xlHAlignCenter
        .Cells(1, 1).value = "Incidencia de materiales"
        .Cells(2, 1).value = "Pieza"
        .Cells(2, 2).value = Pieza

        .Cells(3, 1).value = "Cliente"
        .Cells(3, 2).value = cliente
        .Cells(5, 1).value = "Rubro"
        .Cells(5, 2).value = "Incidencia"

        For xx = 1 To lst.ListItems.count
            .Cells(xx + 5, 1).value = lst.ListItems(xx)
            .Cells(xx + 5, 2).value = lst.ListItems(xx).ListSubItems(1)


        Next xx



        xls.Visible = True    'NO MUESTRO LA HOJA XLS
        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        'Set CDLGMAIN = frmCostosIncidencia.CD
        '
        '
        '
        ''    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        '      sFilter = "Hoja de Calculo|*.xls"
        '        CDLGMAIN.Filter = sFilter
        '        Dim refe As String
        '        refe = ref
        '
        '        archi = "PRES" & Format(presu, "00000") & ".xls"
        '        frmPrincipal.CD.CancelError = True
        '        CDLGMAIN.FileName = archi
        '        CDLGMAIN.ShowSave
        '
        '        If CDLGMAIN.FileName <> Empty Then
        '            xla.SaveAs (CDLGMAIN.FileName)
        '            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
        '            strMsg = strMsg & CDLGMAIN.FileName
        '            MsgBox strMsg, vbExclamation, "Hoja de calculo guardada"
        '        Else
        '        exporta = False
        '
        '
        '       End If

        xlb.Saved = False
        ' xlb.Close
        ' xls.Quit


        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        '    End If
        exporta = True



    End With
    Exit Function
errEXCEL:
    If Err.Number = -2147221080 Then
        exporta = False
    Else
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        exporta = False
    End If
    ' xlb.Close
    ' xls.Quit


    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing
End Function

Public Function EsConjunto(idPieza) As Long
    Dim P As New Recordset
    On Error GoTo err4
    Set P = conectar.RSFactory("select conjunto from stock where id=" & idPieza)
    If Not P.EOF And Not P.BOF Then
        EsConjunto = P!conjunto
    End If
    Exit Function
err4:
    EsConjunto = -2    'hay error!!
End Function




Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
    Set CN2 = Nothing
    Set rs2 = Nothing
End Sub

Public Function GrabarTiemposConjunto(idConjunto As Long, lista As ListView) As Boolean
    error1 = False
    On Error GoTo err1
    cn.BeginTrans
    GrabarTiemposConjunto = True
    cn.execute "DELETE FROM desarrollo_mdo where id_Pieza=" & idConjunto
    For x = 1 To lista.ListItems.count
        codigo = lista.ListItems(x)
        Cantidad = lista.ListItems(x).ListSubItems(1)
        Tiempo = lista.ListItems(x).ListSubItems(2)
        cn.execute "insert into desarrollo_mdo (id_pieza, codigo, cantidad, tiempo) values (" & idConjunto & "," & codigo & "," & Cantidad & "," & Tiempo & ")"
    Next x



    cn.CommitTrans
    Exit Function
err1:
    error1 = True
    GrabarTiemposConjunto = False
    cn.RollbackTrans
End Function


Public Function copiarConjuntoV2(idConjunto As Long, nombreNuevo As String, cliente As Long) As Boolean
'Dim rs As Recordset
    Dim rs1 As Recordset
    Dim rs2 As Recordset
    Dim rs3 As Recordset
    Dim r_mdo As Recordset
    Dim s As New classStock

    On Error GoTo err65
    copiarConjuntoV2 = True
    cn.BeginTrans
    tra = True
    'selecciono el id_cliente para el nvo conj
    Me.ejecutar "Select id_cliente from stock where id=" & idConjunto
    cliente = rs!id_cliente

    cn.execute "insert into stock (detalle,id_cliente,conjunto) values ('" & UCase(nombreNuevo) & "'," & cliente & ",0 )"

    Me.ejecutar "select last_insert_id() as ultid from stock"
    idPiezaPadre = rs!ultid

    If s.EsConjunto(idPiezaPadre) = 0 Then

        strsql = "select * from stockConjuntos where idPiezaPadre=" & idPiezaPadre
        Set rs1 = conectar.RSFactory(strsql)
        While Not rs1.EOF    'itero sobre la primer rama del arbol
            Id = rs1!idPiezaHija
            If s.EsConjunto(Id) = 0 Then
                'si es conjunto, analizo la rama
                strsql = "select * from stockConjuntos where idPiezaPadre=" & Id
                Set rs2 = conectar.RSFactory(strsql)
                While Not rs2.EOF    'itero sobre la segunda rama del arbol
                    id2 = rs2!idPiezaHija
                    If s.EsConjunto(id2) = 0 Then
                        'si es subconjunto, analizo la rama
                        strsql = "select * from stockConjuntos where idPiezaPadre=" & id2
                        Set rs3 = conectar.RSFactory(strsql)
                        While Not rs3.EOF
                            'en esta rama que es la ultima, todas las piezas deberian no ser conjunto
                            id3 = rs3!idPiezaHija

                            Me.ejecutar "select * from stockConjuntos where idPiezaPadre=" & id3
                            While Not rs.EOF
                                Cantidad = rs!Cantidad
                                idPiezaHija = rs!idPiezaHija
                                cn.execute "insert into stockConjuntos (idPiezaPadre, idPiezaHija, cantidad) values (" & idPiezaPadre & "," & idPiezaHija & "," & Cantidad & ")"
                                rs.MoveNext
                            Wend

                            rs3.MoveNext
                        Wend
                    Else

                        'si no es subconjunto
                        Me.ejecutar "select * from stockConjuntos where idPiezaPadre=" & id2
                        While Not rs.EOF
                            Cantidad = rs!Cantidad
                            idPiezaHija = rs!idPiezaHija
                            cn.execute "insert into stockConjuntos (idPiezaPadre, idPiezaHija, cantidad) values (" & idPiezaPadre & "," & idPiezaHija & "," & Cantidad & ")"
                            rs.MoveNext
                        Wend

                    End If
                    rs2.MoveNext
                Wend
            Else
                Me.ejecutar "select * from stockConjuntos where idPiezaPadre=" & Id
                While Not rs.EOF
                    Cantidad = rs!Cantidad
                    idPiezaHija = rs!idPiezaHija
                    cn.execute "insert into stockConjuntos (idPiezaPadre, idPiezaHija, cantidad) values (" & idPiezaPadre & "," & idPiezaHija & "," & Cantidad & ")"
                    rs.MoveNext
                Wend



            End If
            rs1.MoveNext
        Wend

    Else
        MsgBox "no es conjunto para copiar"
    End If

    Me.ejecutar "Select * from stockConjuntos where idPiezaPadre=" & idConjunto
    While Not rs.EOF
        Cantidad = rs!Cantidad
        idPiezaHija = rs!idPiezaHija
        cn.execute "insert into stockConjuntos (idPiezaPadre, idPiezaHija, cantidad) values (" & idPiezaPadre & "," & idPiezaHija & "," & Cantidad & ")"
        rs.MoveNext
    Wend

    Set r_mdo = conectar.RSFactory("select * from desarrollo_mdo where id_pieza=" & idConjunto)
    While Not r_mdo.EOF
        m_codigo = r_mdo!codigo
        m_cantidad = r_mdo!Cantidad
        m_tiempo = r_mdo!Tiempo
        m_detalle = r_mdo!detalle
        cn.execute "insert into desarrollo_mdo (id_Pieza,codigo,cantidad,tiempo,detalle) values (" & idPiezaPadre & "," & m_codigo & "," & m_cantidad & "," & m_tiempo & ",'" & m_detalle & "')"

        r_mdo.MoveNext
    Wend



    cn.CommitTrans
    Exit Function
err65:
    MsgBox Err.Description
    copiarConjuntoV2 = False
    cn.RollbackTrans
End Function
Public Function ArbolConjunto(idConjunto) As Recordset  'idConjunto=detalle pedido
    On Error GoTo err2

    'ArbolConjunto = True
    Dim strsql As String, strsql2 As String
    Dim r_tmp As Recordset
    Dim rs1 As Recordset, rs2 As Recordset, rs3 As Recordset
    Dim Id As Long, id2 As Long, id3 As Long
    Dim cant1 As Double, cant2 As Double, cant3 As Double
    Set r_tmp = New Recordset
    CantConj = 1    'lo pongo en uno para que no quede empty
    With r_tmp
        .Fields.Append "idPieza", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "cantt", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "conjunto", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "idPiezaPadre", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "Cantidad", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "id_pos", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable

    End With
    r_tmp.Open
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim pos3 As Integer
    strsql = "select sc.idPiezaPadre,sc.idPiezaHija,sc.cantidad,s.cantidad as stock,s.detalle from stockConjuntos sc inner join stock s on sc.idPiezaHija=s.id  where idPiezaPadre=" & idConjunto
    Set rs1 = conectar.RSFactory(strsql)
    c1 = 0
    While Not rs1.EOF    'itero sobre la primer rama del arbol
        c1 = c1 + 1
        Id = rs1!idPiezaHija

        idPiezaPadre = rs1!idPiezaPadre
        Cantidad = rs1!Cantidad
        If Me.EsConjunto(Id) = 0 Then
            'si es conjunto, analizo la rama
            'strsql = "select * from stockConjuntos where idPiezaPadre=" & id
            strsql = "select s.cantidad as stock,sc.idPiezaPadre,sc.idPiezaHija,sc.idPiezaHija,sc.cantidad,s.detalle from stockConjuntos sc inner join stock s on sc.idPiezaHija=s.id  where idPiezaPadre=" & Id
            r_tmp.AddNew
            r_tmp!idPieza = Id
            r_tmp!Cantt = Cantidad * CantConj
            r_tmp!conjunto = 1
            r_tmp!idPiezaPadre = idPiezaPadre
            r_tmp!Cantidad = Cantidad
            r_tmp!id_pos = c1
            r_tmp.Update
            Set rs2 = conectar.RSFactory(strsql)
            c2 = 0
            While Not rs2.EOF    'itero sobre la segunda rama del arbol
                c2 = c2 + 1
                id2 = rs2!idPiezaHija
                cantidad2 = rs2!Cantidad
                detalle2 = rs2!detalle
                idPiezaPadre2 = rs2!idPiezaPadre
                stock2 = rs2!stock
                If Me.EsConjunto(id2) = 0 Then
                    r_tmp.AddNew
                    r_tmp!idPieza = id2
                    r_tmp!Cantt = cantidad2 * CantConj * Cantidad
                    r_tmp!conjunto = 1
                    r_tmp!idPiezaPadre = idPiezaPadre2
                    r_tmp!Cantidad = cantidad2
                    r_tmp!id_pos = c1 & "." & c2
                    r_tmp.Update
                    'si es conjunto, analizo la rama
                    'strsql = "select * from stockConjuntos where idPiezaPadre=" & id2
                    strsql = "select sc.idPiezaPadre,sc.idPiezaHija,sc.cantidad,s.cantidad as stock,s.detalle from stockConjuntos sc inner join stock s on sc.idPiezaHija=s.id  where idPiezaPadre=" & id2
                    Set rs3 = conectar.RSFactory(strsql)
                    c3 = 0
                    While Not rs3.EOF
                        c3 = c3 + 1
                        'en esta rama que es la ultima, todas las piezas deberian no ser conjunto
                        'asique directamente sumamos los costos
                        id3 = rs3!idPiezaHija
                        idPiezaPadre3 = rs3!idPiezaPadre
                        cantidad3 = rs3!Cantidad
                        detalle3 = rs3!detalle
                        stock3 = rs3!stock
                        r_tmp.AddNew
                        r_tmp!idPieza = id3
                        r_tmp!Cantt = cantidad3 * CantConj * cantidad2 * Cantidad
                        r_tmp!conjunto = 0
                        r_tmp!idPiezaPadre = idPiezaPadre3
                        r_tmp!Cantidad = cantidad3
                        r_tmp!id_pos = c1 & "." & c2 & "." & c3
                        r_tmp.Update
                        rs3.MoveNext
                    Wend
                Else
                    'si no es conjunto, acumulo el costo
                    r_tmp.AddNew
                    r_tmp!idPieza = id2
                    r_tmp!Cantt = cantidad2 * CantConj * Cantidad
                    r_tmp!conjunto = 0
                    r_tmp!idPiezaPadre = idPiezaPadre2
                    r_tmp!Cantidad = cantidad2
                    r_tmp!id_pos = c1 & "." & c2
                    r_tmp.Update
                End If
                rs2.MoveNext
            Wend
        Else
            r_tmp.AddNew
            r_tmp!idPieza = Id
            r_tmp!Cantt = Cantidad * CantConj
            r_tmp!conjunto = 0
            r_tmp!idPiezaPadre = idPiezaPadre
            r_tmp!Cantidad = Cantidad
            r_tmp!id_pos = c1
            r_tmp.Update

        End If
        rs1.MoveNext
    Wend

    If Not r_tmp.EOF Then r_tmp.MoveFirst




    Set ArbolConjunto = r_tmp
    'r_tmp.Close

    Exit Function
err2:
    MsgBox Err.Description
    ArbolConjunto False
End Function
Public Function ListaPiezas(Id As Long, Optional Ot As Boolean = False, Optional presupuesto As Boolean = False, Optional otro As Boolean = False, Optional OTGrabada As Boolean = False, Optional ListaPieza As Dictionary) As Recordset
    On Error GoTo err5
    Dim clsC As New classStock
    Dim conju As Integer
    Dim r_l As New Recordset
    Dim r1 As New Recordset
    Dim r_conj As New Recordset
    If Ot And presupuesto And otro Then Exit Function
    With r1
        .Fields.Append "idPieza", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "cantidad", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        .Fields.Append "detalle", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
    End With
    r1.Open

    If presupuesto Then
        strsql = "select dp.idPieza,dp.cantidad,s.conjunto as cc,s.detalle from detalle_presupuesto dp inner join stock s on dp.idPieza=s.id where idPresupuesto=" & Id
    ElseIf Ot Then
        strsql = "select dp.idpieza,dp.cantidad,s.conjunto as cc,s.detalle from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where idPedido=" & Id
    ElseIf otro Then
        strsql = "select id as idpieza,1 as cantidad,conjunto as cc,s.detalle from stock s where id=" & Id
    Else
        Dim K As Variant
        Dim pp As Variant
        Dim r2 As New Recordset

        With r2
            .Fields.Append "idPieza", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "cantidad", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "cc", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
            .Fields.Append "detalle", adVarChar, 255, adFldUpdatable    ' And adFldIsNullable
        End With
        r2.Open
        If IsSomething(ListaPieza) Then
            For Each K In ListaPieza
                r2.AddNew
                r2!Cantidad = ListaPieza(K)
                r2!idPieza = K.Id
                r2!cc = IIf(K.EsConjunto, "0", "-1")
                r2!detalle = K.nombre
                r2.Update
            Next


        End If

    End If



    If Not presupuesto And Not otro And Not Ot And IsSomething(ListaPieza) Then
        Set r_l = r2
        r2.MoveFirst
    Else

        Set r_l = conectar.RSFactory(strsql)
    End If

    While Not r_l.EOF
        Cantidad = r_l!Cantidad
        idp = r_l!idPieza
        conju = r_l!cc


        If conju = 0 Then
            'si es conjunto

            Set r_conj = clsC.ArbolConjunto(idp)
            While Not r_conj.EOF

                idPieza_c = r_conj!idPieza
                Debug.Print idPieza_c
                cant_c = r_conj!Cantt
                r1.AddNew
                r1!idPieza = idPieza_c
                r1!Cantidad = cant_c * Cantidad

                r1.Update
                r_conj.MoveNext
            Wend
        End If    '

        '        Else
        'si no es conjunto
        r1.AddNew
        r1!idPieza = idp
        r1!Cantidad = Cantidad
        r1.Update
        'End If
        r_l.MoveNext
    Wend
    r1.MoveFirst


    Set ListaPiezas = r1
    Exit Function
err5:
    MsgBox "se produjo un error: " & Err.Description
End Function


Public Function materializacion(idPieza) As Recordset
    On Error GoTo er1
    'Me.ejecutar "select m.codigo,m.descripcion,r.rubro,g.grupo,m.espesor from desarrollo_material dm inner join materiales m on dm.id_material=m.id inner join rubros r on m.id_rubro=r.id inner join grupos g on m.id_grupo=g.id where dm.id_pieza=" & idPieza

    Me.ejecutar "select m.id_unidad,dm.scrap,dm.cantidad,dm.id_material,m.codigo,m.descripcion,r.rubro,g.grupo,m.espesor,dm.largo,dm.ancho,dm.largoTerm,dm.AnchoTerm from desarrollo_material dm inner join materiales m on dm.id_material=m.id inner join rubros r on m.id_rubro=r.id inner join grupos g on m.id_grupo=g.id where dm.id_pieza=" & idPieza
    Set materializacion = rs
    Exit Function
er1:

End Function



Public Function exportaMaterializacion(Id As Long, Optional vpresu As Boolean = False, Optional vot As Boolean = False, Optional votro As Boolean = False) As Boolean
    On Error GoTo errEXCEL
    Dim claseP As New classPlaneamiento
    Dim claseS As New classStock
    Dim rs As Recordset
    Dim rs_c As Recordset

    Dim idPiezaC As Long
    Dim strsql As String, strsql2 As String
    Dim xlb As New Excel.Workbook
    Dim xla As New Excel.Worksheet
    Dim xls As New Excel.Application
    Set xlb = xls.Workbooks.Add
    Set xla = xlb.Worksheets.Add
    xla.Activate
    If vot Then    'desde ot
        strsql = "select c.razon,p.descripcion, p.fechaEntrega as fecha from pedidos p inner join clientes c on p.idCliente=c.id where p.id=" & Id
        strsql2 = "select dp.precio,dp.idPieza,dp.item,s.detalle,dp.cantidad  from detalles_pedidos dp inner join stock s on dp.idPieza=s.id where idPedido=" & Id
        enc = "OT"
    ElseIf vpresu Then    'desde presu
        strsql = "select c.razon,p.detalle as descripcion, concat(p.fechaEntrega, ' días') as fecha from  presupuestos p inner join clientes c on p.idCliente=c.id where p.id=" & Id
        strsql2 = "select dp.valorUnitarioManual,dp.idPieza,dp.item,s.detalle,dp.cantidad  from detalle_presupuesto dp inner join stock s on dp.idPieza=s.id where idPresupuesto=" & Id
        enc = "PRES"
    End If


    Set rs = conectar.RSFactory(strsql)

    With xla
        .Columns("A").HorizontalAlignment = xlHAlignCenter
        .Columns("B").HorizontalAlignment = xlHAlignCenter
        .Columns("C").HorizontalAlignment = xlHAlignCenter
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("F").HorizontalAlignment = xlHAlignCenter
        .Columns("G").HorizontalAlignment = xlHAlignCenter
        .Columns("H").HorizontalAlignment = xlHAlignCenter
        .Columns("I").HorizontalAlignment = xlHAlignCenter

        .Range("A7", "I7").Font.Bold = True
        .Range("A7", "I7").Font.Size = 12
        .Range("a7:i7").Interior.Color = &HC0C0C0




        '.Range("A1:A4").HorizontalAlignment = xlHAlignRight

        .Columns("C").ColumnWidth = 40
        .Columns("e").ColumnWidth = 30
        .Columns("f").ColumnWidth = 55
        .Columns("i").ColumnWidth = 30
        .Range("b1:C1").Merge
        .Range("B2:C2").Merge
        .Range("B3:C3").Merge
        .Range("B4:C4").Merge

        .Cells(1, 1).value = enc
        .Cells(1, 2).value = str(Format(Id, "0000"))
        .Cells(2, 1).value = "Fecha"
        .Cells(2, 2).value = rs!FEcha
        .Cells(3, 1).value = "Cliente"
        .Cells(3, 2).value = rs!razon
        .Cells(4, 1).value = "Referencia"
        .Cells(4, 2).value = rs!descripcion


        .Cells(7, 1).value = "Item"
        .Cells(7, 2).value = "Cant"
        .Cells(7, 3).value = "Detalle"
        .Cells(7, 6).value = "Materiales"
        .Cells(7, 7).value = "M2"
        .Cells(7, 8).value = "Kg"
        .Cells(7, 9).value = "Pieza"






        Set rs = conectar.RSFactory(strsql2)
        c = 0
        'encambezado

        offset = 7
        While Not rs.EOF

            'precio = rs!valorUnitarioManual
            offset = offset + 1
            idPieza = rs!idPieza
            .Cells(offset, 1).value = rs!item
            .Cells(offset, 2).value = rs!Cantidad
            .Cells(offset, 3).value = rs!detalle
            '.Cells(Offset, 10).value = precio

            If claseS.EsConjunto(idPieza) = 0 Then

                Set rs_c = claseS.ArbolConjunto(idPieza)
                offsetInicial = offset
                While Not rs_c.EOF
                    offset = offset + 1
                    idPiezaC = rs_c!idPieza
                    .Cells(offset, 4).value = rs_c!Cantidad
                    .Cells(offset, 5).value = claseS.detalle_pieza(idPiezaC)


                    Set rs_m = Me.materializacion(idPiezaC)

                    While Not rs_m.EOF

                        .Cells(offset, 6).value = rs_m!rubro & " " & rs_m!Grupo & " " & rs_m!descripcion & " " & rs_m!Espesor

                        x2 = rs_m!LargoTerm
                        y2 = rs_m!AnchoTerm
                        x = rs_m!Largo
                        y = rs_m!Ancho
                        Scrap = rs_m!Scrap
                        id_material = rs_m!id_material
                        Cantidad = rs_m!Cantidad

                        Me.calcularM2MLKGMaterial x2, y2, id_material, Scrap, x, y, Cantidad, Kg, m2, Pieza, costo, 0
                        .Cells(offset, 7).value = m2
                        .Cells(offset, 8).value = Kg
                        .Cells(offset, 9).value = Pieza
                        offset = offset + 1
                        rs_m.MoveNext


                    Wend
                    offset = offset - 1

                    rs_c.MoveNext
                Wend
                A = "a" & offsetInicial
                B = "i" & offset
                .Range(A, B).BorderAround , xlMedium
            Else
                Set rs_mm = Me.materializacion(idPieza)

                c = 0
                While Not rs_mm.EOF
                    c = c + 1
                    rs_mm.MoveNext
                Wend
                If c > 0 Then rs_mm.MoveFirst

                While Not rs_mm.EOF
                    offsetInicial = offset
                    .Cells(offset, 6).value = rs_mm!rubro & " " & rs_mm!Grupo & " " & rs_mm!descripcion & " " & rs_mm!Espesor
                    x2 = rs_mm!LargoTerm
                    y2 = rs_mm!AnchoTerm
                    x = rs_mm!Largo
                    y = rs_mm!Ancho
                    Scrap = rs_mm!Scrap
                    id_material = rs_mm!id_material
                    Cantidad = rs_mm!Cantidad

                    Me.calcularM2MLKGMaterial x2, y2, id_material, Scrap, x, y, Cantidad, Kg, m2, Pieza, costo, 0
                    .Cells(offset, 7).value = m2
                    .Cells(offset, 8).value = Kg
                    .Cells(offset, 9).value = Pieza

                    offset = offset + 1
                    A = "a" & offsetInicial
                    B = "i" & offset
                    .Range(A, B).BorderAround , xlMedium
                    rs_mm.MoveNext
                Wend
                offset = offset - 1
            End If


            rs.MoveNext
        Wend










        'xls.Visible = True NO MUESTRO LA HOJA XLS
        strMsg = "Se han transportado los datos correctamente"
        strMsg = strMsg & vbCrLf & "a una hoja de calculo de Excel."
        strMsg = strMsg & vbCrLf & vbCrLf
        strMsg = strMsg & "¿Desea guardar la hoja de calculo de Excel?"
        Set CDLGMAIN = frmPrincipal.cd



        '    If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
        sFilter = "Hoja de Calculo|*.xls"
        CDLGMAIN.filter = sFilter
        Dim refe As String
        refe = ref

        archi = enc & Format(Id, "00000") & ".xls"
        '        frmPrincipal.cd.CancelError = True
        frmPrincipal.cd.filename = archi
        frmPrincipal.cd.ShowSave

        'CDLGMAIN.filename = archi
        '    CDLGMAIN.ShowSave

        If CDLGMAIN.filename <> Empty Then
            xla.SaveAs (CDLGMAIN.filename)
            strMsg = "Los datos del reporte se han guardado en un archivo: " & vbCrLf & vbCrLf
            strMsg = strMsg & CDLGMAIN.filename
            MsgBox strMsg, vbExclamation, "Hoja de calculo guardada"
            archi = CDLGMAIN.filename
        Else
            exportaMaterializacion = False
        End If
        xlb.Saved = True
        xlb.Close
        xls.Quit
        Set xls = Nothing
        Set xla = Nothing
        Set xlb = Nothing

        '    End If
        exportaMaterializacion = True



    End With
    Exit Function
errEXCEL:


    If Err.Number = -2147221080 Then
        exportaMaterializacion = False
    Else
        MsgBox "Se produjo un error. No se graban los cambios", vbCritical, "Error"
        exportaMaterializacion = False
    End If
    xlb.Saved = True
    xlb.Close
    xls.Quit
    Set xls = Nothing
    Set xla = Nothing
    Set xlb = Nothing
End Function
Public Function crearHistoricoPieza2(idPieza, Nota)
    idUsuario = funciones.getUser
    fechas_ = Now    'Format(Now, "yyyy-mm-dd")
    fechas_ = funciones.datetimeFormateada(Now)
    cn.execute "insert into stockHistorico (idPieza,fecha,nota,autor) values (" & idPieza & ",'" & fechas_ & "','" & UCase(Nota) & "'," & idUsuario & ")"

End Function



