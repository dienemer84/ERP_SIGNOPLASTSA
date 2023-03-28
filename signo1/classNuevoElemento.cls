VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classNuevoElemento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim baseS As New classStock
Dim strsql As String
Dim cn As ADODB.Connection    'materiales
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim valor_dolar As Double
Private Sub Class_Initialize()
'On Error GoTo err22
    Set cn = conectar.obternerConexion
    Exit Sub
err22:
    MsgBox Err.Description
End Sub


Public Function ejecutar_consulta1(consulta As String)
    On Error GoTo err10
    Set rs1 = conectar.RSFactory(consulta)


    Exit Function
err10:
    MsgBox "Se produjo un error: " & Err.Description
End Function


Public Function ejecutar_consulta2(consulta As String)
    On Error GoTo err11
    Set rs2 = conectar.RSFactory(consulta)
    Exit Function
err11:
    MsgBox "Se produjo un error: " & Err.Description
End Function
Public Function ejecutar_consulta3(consulta As String)
    On Error GoTo err12

    Set rs3 = conectar.RSFactory(consulta)
    Exit Function
err12:
    MsgBox "Se produjo un error: " & Err.Description
End Function
Public Sub ver_detalle_elemento(ByVal cod As String, frm As frmNuevoElemento, accion As Integer)
    strsql = "select v.valor_unitario, v.fecha_actualizacion,v.id_moneda,g.grupo,m.id, m.codigo, m.descripcion,m.espesor,m.pesoxunidad,m.id_unidad,r.rubro from materiales m inner join rubros r on m.id_rubro=r.id inner join grupos g on g.id=m.id_grupo inner join valores_MATERIALES v on v.id_material=m.id and m.codigo= '" & cod & "'"
    Me.ejecutar_consulta1 strsql
    c = 0
    While Not rs1.EOF
        c = c + 1
        rs1.MoveNext
    Wend
    If Trim(cod) <> Empty Then
        If c = 1 Then
            rs1.MoveFirst
            valor_dolar = 3    'modificar este valor por uno en tabla configuraciones
            PesoXUnidad = rs1!PesoXUnidad
            unidad = rs1!id_Unidad
            id_material = rs1!Id
            rubro = rs1!rubro
            Grupo = rs1!Grupo
            descripcion = rs1!descripcion
            Espesor = rs1!Espesor
            valorxunidad = rs1!valor_unitario
            moneda = rs1!id_moneda
            frm.lblEspesor = Espesor
            frm.lblMaterial = rubro & " " & Grupo
            If moneda = 1 Then
                Precio = Precio * valor_dolar
            End If
            IdMaterial = rs1!Id
            If Trim(frm.txtAnchoPieza) = Empty Or Trim(frm.txtLargoPieza) = Empty Then
            Else
                If IsNumeric(frm.txtAnchoPieza) And IsNumeric(frm.txtLargoPieza) Then
                    cantidad_h = CDbl(frm.txtCantidad)


                    anchop = CDbl(frm.txtAnchoPieza)
                    largop = CDbl(frm.txtLargoPieza)

                    Scrap = CDbl(frm.txtScrap)
                    scrap_viejo = Scrap
                    Scrap = (Scrap / 100) + 1

                    frm.lblEspesor = Espesor
                    frm.lblMaterial = rubro & " " & Grupo



                    If unidad = 3 Then
                        anchop = 1
                        frm.lblkg = Math.Round(PesoXUnidad * (CDbl(largop) / 1000) * cantidad_h, 3)
                        frm.lblm2 = Math.Round((CDbl(largop)) / 1000 * cantidad_h, 3)
                        Precio = valorxunidad * PesoXUnidad
                        prec = ((Precio * largop * Scrap * 2) / 1000) * cantidad_h
                        Pieza = cantidad_h & "x" & largop
                    ElseIf unidad = 1 Then
                        frm.lblkg = Math.Round(PesoXUnidad * cantidad_h, 3)
                        frm.lblm2 = 0    'Math.Round((CDbl(anchop) * CDbl(LargoP)) / 1000000 * cantidad_h, 3)
                        Precio = valorxunidad * PesoXUnidad
                        prec = (Precio * cantidad_h)
                        Pieza = cantidad_h & "x" & anchop & "x" & largop
                    ElseIf unidad = 2 Then
                        frm.lblkg = Math.Round(PesoXUnidad * ((CDbl(anchop) * CDbl(largop)) / 1000000) * cantidad_h, 3)
                        frm.lblm2 = Math.Round((CDbl(anchop) * CDbl(largop)) / 1000000 * cantidad_h, 3)
                        Precio = valorxunidad * PesoXUnidad
                        prec = ((Precio * (anchop * largop) * Scrap) / 1000000) * cantidad_h
                        Pieza = cantidad_h & "x" & anchop & "x" & largop
                    ElseIf unidad = 4 Then
                        frm.lblkg = 0
                        frm.lblm2 = 0
                        anchop = 0
                        largop = 0
                        Precio = valorxunidad
                        prec = Precio * cantidad_h
                        Pieza = cantidad_h
                    End If

                    frm.lblidMaterial = IdMaterial

                    prec = prec / valor_dolar
                    prec = Math.Round(prec, 3)
                    If accion = 1 Then    'agregar
                        Set x = frm.ListView1.ListItems.Add(, , Id)
                        x.SubItems(1) = UCase(cod)
                        x.SubItems(2) = frm.lblMaterial & " " & frm.lblEspesor & " mm"
                        x.SubItems(3) = Pieza
                        x.SubItems(4) = anchop
                        x.SubItems(5) = largop
                        x.SubItems(6) = Scrap
                        x.SubItems(7) = Math.Round(CDbl(frm.lblkg), 3)
                        x.SubItems(8) = Math.Round(CDbl(frm.lblm2), 3)
                        x.SubItems(9) = cantidad_h
                        x.SubItems(10) = id_material
                        x.SubItems(11) = ""    'agregar term
                        x.SubItems(12) = "x"    'suma?


                    End If

                Else
                    MsgBox "Ingrese datos válidos", vbCritical, "Error"
                End If
            End If
        Else
            If accion <> 2 And accion <> 3 Then
                MsgBox "El código es inexistente", vbCritical, "Error"
                frm.txtCodigoMaterial.SetFocus
            End If
        End If
    End If
End Sub

Public Function calcular_totales_material(frm As frmNuevoElemento)
    canti = frm.ListView1.ListItems.count
    TotalKG = 0
    TotalM2 = 0
    totalS = 0
    For i = 1 To canti
        If frm.ListView1.ListItems(i).ListSubItems(12) = "x" Then
            TotalKG = TotalKG + CDbl(frm.ListView1.ListItems(i).ListSubItems(7))
            TotalM2 = TotalM2 + CDbl(frm.ListView1.ListItems(i).ListSubItems(8))
        End If
    Next i
    frm.lblTotalKg = Math.Round(TotalKG, 3)
    frm.lblTotalM2 = Math.Round(TotalM2, 3)

End Function

Public Sub ver_detalle_mdo(ByVal codi As Integer, ByRef idcpp, ByRef cantxproc, ByRef mdoDescrip, ByRef Tarea, ByRef Sector, ByRef Valor)
    strsql = "select v.id_tarea,v.valor, v.fecha, v.descripcion,s.sector,t.tarea,t.id_sector,t.cantxproc from tareas t inner join sectores s on t.id_sector=s.id inner join  valores_MDO v on v.id_tarea=t.id and t.id=" & codi
    Me.ejecutar_consulta2 strsql


    'calculo la cantidad de registros
    c = 0
    While Not rs2.EOF
        c = c + 1
        rs2.MoveNext
    Wend

    If c = 1 Then
        rs2.MoveFirst
        Sector = rs2!Sector
        cantxproc = rs2!cantxproc

        If cantxproc = -1 Then cantxproc = "Cambio"
        If cantxproc = 0 Then cantxproc = "Fijo"

        idcpp = rs2!cantxproc
        Tarea = rs2!Tarea
        mdoDescrip = rs2!descripcion
        Valor = rs2!Valor
        ' frm.lblCPP = cantxproc
        ' rm.lblTarea = tarea
        ' frm.lblDescripcion = MDOdescrip

        ' frm.lblidMDO = codi
        ' frm.lblSector = sector
    Else
        MsgBox "El código es inexistente", vbCritical, "Error"
    End If
End Sub



Public Sub agregar(lstMAT As ListView, lstMdo As ListView, detalle As String, id_cliente As Integer, Complejidad As Integer)
    On Error GoTo er2
    'id_cliente = 120
    detail = UCase(detalle)
    Cantidad = 0

    cn.BeginTrans
    'alta en stock.

    Dim tipo_complejidad As TipoComplejidad
    tipo_complejidad = ComplejidadBaja
    If Complejidad = 1 Then tipo_complejidad = ComplejidadMedia
    If Complejidad = 2 Then tipo_complejidad = ComplejidadAlta

    cn.execute "insert into stock (detalle,id_cliente,cantidad, tipo_complejidad) VALUES ('" & detail & "'," & id_cliente & "," & Cantidad & "," & Complejidad & ")"

    Me.ejecutar_consulta3 "select last_insert_id() as idd from stock"
    ultid = rs3!idd
    cantL1 = lstMAT.ListItems.count
    'alta en desarrollo materiales del elemento
    For x = 1 To cantL1
        Scrap = lstMAT.ListItems(x).ListSubItems(9)
        id_material = lstMAT.ListItems(x).ListSubItems(1)
        Largo = lstMAT.ListItems(x).ListSubItems(5)
        Ancho = lstMAT.ListItems(x).ListSubItems(6)
        largoT = lstMAT.ListItems(x).ListSubItems(7)
        anchoT = lstMAT.ListItems(x).ListSubItems(8)
        Cantidad = lstMAT.ListItems(x).ListSubItems(14)
        deta = UCase(Trim(lstMAT.ListItems(x).Tag))
        cn.execute "insert into desarrollo_material (id_pieza,scrap,largo,ancho,largoTerm,AnchoTerm,id_material,cantidad,detalle) values (" & CInt(ultid) & "," & CDbl(Scrap) & "," & CDbl(Largo) & "," & CDbl(Ancho) & " ," & CDbl(largoT) & "," & CDbl(anchoT) & ", " & id_material & "," & Cantidad & ",'" & deta & "')"


    Next x

    cantL2 = lstMdo.ListItems.count
    'alta en desarrollo de mdo del elemento
    For y = 1 To cantL2


        id_mdo = lstMdo.ListItems(y).ListSubItems(1)
        Cant = lstMdo.ListItems(y).ListSubItems(2)
        Tiempo = lstMdo.ListItems(y).ListSubItems(3)
        deta1 = UCase(Trim(lstMdo.ListItems(y).Tag))
        cn.execute "insert into desarrollo_mdo (id_pieza,codigo,cantidad,tiempo,detalle) values (" & CInt(ultid) & "," & CInt(id_mdo) & "," & CDbl(Cant) & "," & CDbl(Tiempo) & ",'" & deta1 & "')"
    Next y


    baseS.crearHistoricoPieza2 ultid, "PIEZA AGREGADA"
    cn.CommitTrans
    Dim nico As Long
    nico = ultid
    Set baseS = New classStock
    baseS.modifica_Stock nico, 0, 0, 0, ""    'doy d alta en bbdd historicos de stock
    rs3.Close

    Exit Sub
er2:
    cn.RollbackTrans

End Sub



Public Sub llenarListaMDO(idPieza, lst As ListView)
    Dim x As ListItem

    'Me.ejecutar_consulta1 "select vm.valor,t.tarea,d.id,t.id as codigo,t.cantxproc,d.cantidad,d.tiempo,s.sector,t.cantxproc from sectores s,stock.desarrollo_mdo d, tareas t  where d.id_pieza=" & idPieza & " and t.id=d.codigo and s.id=t.id_sector"
    Me.ejecutar_consulta1 "select vm.valor,t.cantxproc,vm.descripcion,t.id as codigo,d.id,d.cantidad,t.tarea,d.detalle,d.tiempo,s.sector,t.cantxproc from valores_MDO vm,sectores s,desarrollo_mdo d, tareas t  where d.id_pieza=" & idPieza & " and t.id=d.codigo and s.id=t.id_sector and t.id=vm.id_tarea"
    While Not rs1.EOF
        Set x = lst.ListItems.Add(, , rs1!Id)
        Valor = rs1!Valor
        Tiempo = rs1!Tiempo
        cantop = rs1!Cantidad
        cpa = rs1!cantxproc
        cpp = rs1!cantxproc
        Tarea = rs1!Tarea
        codigo = rs1!codigo
        Sector = rs1!Sector
        Desc = rs1!descripcion
        deta = rs1!detalle
        If cpa = -1 Then cpa = "Cambio"
        If cpa = 0 Then cpa = "Fijo"
        x.SubItems(1) = codigo
        x.SubItems(2) = cantop
        x.SubItems(3) = Tiempo
        x.SubItems(4) = Sector
        x.SubItems(5) = cpa
        x.SubItems(6) = Tarea
        x.SubItems(7) = cpp
        x.SubItems(8) = Desc


        If cpp > 0 Then    '(cpp variable)
            totmin = cantop * Tiempo / cpp
            totplata = totmin * Valor
        Else
            totmin = cantop * Tiempo
            totplata = totmin * Valor

        End If
        x.SubItems(9) = Math.Round(totmin, 2)
        x.SubItems(10) = Math.Round(totplata, 2)
        x.Tag = deta
        rs1.MoveNext
    Wend
End Sub


Public Sub llenarLstmateriales(idPieza, lst As ListView)
    Dim Kg As Double, m2ml As Double
    Dim descripcion As String
    'Me.ejecutar_consulta1 "select * from stock.desarrollo_material where id_pieza=" & idPieza
    Me.ejecutar_consulta1 "select d.id,d.id_pieza,d.cantidad,d.scrap,d.largo,d.ancho,d.LargoTerm,d.AnchoTerm,d.id_material,m.codigo,d.detalle from desarrollo_material d, materiales m where d.id_pieza=" & idPieza & " and m.id=d.id_material"
    While Not rs1.EOF
        Id = rs1!id_material
        x = rs1!Ancho
        y = rs1!Largo
        x1 = rs1!AnchoTerm
        y1 = rs1!LargoTerm
        deta = rs1!detalle
        Scrap = rs1!Scrap
        Cant = rs1!Cantidad
        codigo = rs1!codigo
        'si existe el código
        If Id <> -1 Then
            cxh = funciones.cantxhoja(x, y, x1, y1)
            baseS.ejecutar "select m.id_unidad,m.espesor,m.descripcion, g.grupo, r.rubro from materiales m,grupos g, rubros r where m.id_grupo=g.id and m.id_rubro=r.id and  m.id=" & Id
            descripcion = baseS.descripcion
            Espesor = baseS.Espesor
            Grupo = baseS.Grupo
            rubro = baseS.rubro

            uni = baseS.idUnidad

            If uni = 3 Then cxh = 1
            'si son ml cantxhoja debe quedar en 1

            If uni = 4 Then
                x = 0
                x1 = 0
                y = 0
                y1 = 0
            End If
            Dim costo As Double

            baseS.calcularM2MLKGMaterial x1, y1, Id, Scrap, x, y, Cant, Kg, m2ml, Pieza, costo, 0
            'agrego datos a la lista
            Dim h As ListItem
            Set h = lst.ListItems.Add(, , codigo)
            h.SubItems(1) = Id
            h.SubItems(2) = rubro & " " & Grupo & " " & descripcion
            h.SubItems(3) = Espesor
            h.SubItems(4) = Pieza
            h.SubItems(5) = y    'ancho
            h.SubItems(6) = x
            h.SubItems(7) = y1    'ancho T
            h.SubItems(8) = x1
            h.SubItems(9) = Scrap
            h.SubItems(10) = Kg
            h.SubItems(11) = m2ml
            h.SubItems(12) = Math.Round(costo, 2)
            h.SubItems(13) = ""
            h.SubItems(14) = Cant

            If IsNull(deta) Then
                h.Tag = vbNullString
            Else
                h.Tag = deta
            End If

        End If
        rs1.MoveNext
    Wend
End Sub


Public Sub modificar(lstMAT As ListView, lstMdo As ListView, detalle As String, id_cliente As Integer, idPieza As Long, Complejidad As TipoComplejidad)
    On Error GoTo err555
    detail = UCase(detalle)
    Cantidad = 0
    cn.BeginTrans
    tra = True
    'modifico en stock.

    cn.execute "update stock set tipo_complejidad='" & Complejidad & "', detalle='" & detail & "',id_cliente= " & id_cliente & " where id=" & idPieza



    ultid = idPieza
    cantL1 = lstMAT.ListItems.count
    'elimino datos anteriores
    cn.execute "delete from desarrollo_material where id_pieza=" & idPieza
    'alta en desarrollo materiales del elemento
    For x = 1 To cantL1
        Scrap = lstMAT.ListItems(x).ListSubItems(9)
        id_material = lstMAT.ListItems(x).ListSubItems(1)
        Largo = lstMAT.ListItems(x).ListSubItems(5)
        Ancho = lstMAT.ListItems(x).ListSubItems(6)
        largoT = lstMAT.ListItems(x).ListSubItems(7)
        anchoT = lstMAT.ListItems(x).ListSubItems(8)
        Cantidad = lstMAT.ListItems(x).ListSubItems(14)
        deta = UCase(Trim(lstMAT.ListItems(x).Tag))
        cn.execute "insert into desarrollo_material (id_pieza,scrap,largo,ancho,largoTerm,AnchoTerm,id_material,cantidad,detalle) values (" & CInt(ultid) & "," & CDbl(Scrap) & "," & CDbl(Largo) & "," & CDbl(Ancho) & " ," & CDbl(largoT) & "," & CDbl(anchoT) & ", " & id_material & "," & Cantidad & ",'" & deta & "')"
        'CN3.Execute "insert into desarrollo_material (id_pieza,scrap,largo,ancho,largoTerm,AnchoTerm,id_material,cantidad,detalle) values (" & CInt(ultid) & "," & CDbl(scrap) & "," & CDbl(largo) & "," & CDbl(ancho) & " ," & CDbl(largoT) & "," & CDbl(anchoT) & ", " & id_material & "," & cantidad & ")"

    Next x

    cantL2 = lstMdo.ListItems.count
    'elimino datos anteriores
    cn.execute "delete from desarrollo_mdo where id_pieza=" & idPieza
    'alta en desarrollo de mdo del elemento
    For y = 1 To cantL2


        id_mdo = lstMdo.ListItems(y).ListSubItems(1)
        Cant = lstMdo.ListItems(y).ListSubItems(2)
        Tiempo = lstMdo.ListItems(y).ListSubItems(3)
        deta1 = UCase(Trim(lstMdo.ListItems(y).Tag))
        cn.execute "insert into desarrollo_mdo (id_pieza,codigo,cantidad,tiempo,detalle) values (" & CInt(ultid) & "," & CInt(id_mdo) & "," & CDbl(Cant) & "," & CDbl(Tiempo) & ",'" & deta1 & "')"
        'CN3.Execute "insert into desarrollo_mdo (id_pieza,codigo,cantidad,tiempo) values (" & CInt(ultid) & "," & CInt(id_mdo) & "," & CDbl(Cant) & "," & CDbl(tiempo) & ")"
    Next y
    tra = False

    cn.CommitTrans
    baseS.crearHistoricoPieza2 ultid, "PIEZA EDITADA"
    Dim nico As Long
    nico = ultid
    'Set baseS = New classStock
    'baseS.modifica_Stock nico, 0, 0, 0 'doy d alta en bbdd historicos de stock
    Exit Sub
err555:
    cn.RollbackTrans

End Sub

Private Sub Class_Terminate()
    Set rs3 = Nothing
    Set rs1 = Nothing
    Set cn = Nothing
    Set rs2 = Nothing
End Sub

