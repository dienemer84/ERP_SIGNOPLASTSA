Attribute VB_Name = "DAORequeMateriales"
Option Explicit
Dim rs As Recordset
Public Const CANT_DIAS_AVISO_VENCIMIENTO As Long = 3

Public Function FindAllRequesIdProximosAVencer() As Dictionary
    Set FindAllRequesIdProximosAVencer = FindAllRequeIdVencimiento("ent.fecha > " & conectar.Escape(Date) & " AND ent.fecha <= " & conectar.Escape(DateAdd("d", CANT_DIAS_AVISO_VENCIMIENTO, Date)))
End Function

Public Function FindAllRequesVencenHoy() As Dictionary
    Set FindAllRequesVencenHoy = FindAllRequeIdVencimiento("ent.fecha = " & conectar.Escape(Date))
End Function

Public Function FindAllRequesVencidos() As Dictionary
    Set FindAllRequesVencidos = FindAllRequeIdVencimiento("ent.fecha < " & conectar.Escape(Date))
End Function

Private Function FindAllRequeIdVencimiento(filter As String) As Dictionary
    Dim F As String
    Dim estados As New Collection
    estados.Add EstadoRequeCompra.EnEdición_    'para que no se lespase cuando estan editando
    estados.Add EstadoRequeCompra.Finalizado_    'por si se olvidan de aprobar

    estados.Add EstadoRequeCompra.Aprobado_
    estados.Add EstadoRequeCompra.EnProceso_
    estados.Add EstadoRequeCompra.Procesado_
    estados.Add EstadoRequeCompra.AprobadoParcial_
    estados.Add EstadoRequeCompra.EnProcesoParcial_
    estados.Add EstadoRequeCompra.ProcesadoParcial_

    Dim deta As clsRequeMateriales

    Dim detas As New Collection
    Set detas = Find("det.estado IN (" & funciones.JoinCollectionValues(estados, ", ") & ") AND " & filter)

    Dim RequesId As New Dictionary

    For Each deta In detas
        If Not RequesId.Exists(CStr(deta.RequeId)) Then
            RequesId.Add CStr(deta.RequeId), deta.RequeId
        End If
    Next deta


    Set FindAllRequeIdVencimiento = RequesId
End Function

Public Function FindListosParaPetOf() As Collection
    Set FindListosParaPetOf = Find("det.estado in (" & EstadoRequeCompra.Procesado_ & ", " & EstadoRequeCompra.EnPOParcial_ & ")")
End Function

Public Function Find(Optional filter As String = vbNullString) As Collection
    Dim q As String
    Dim r As Recordset
    Dim col As New Collection
    Dim MAT As clsRequeMateriales

    q = "SELECT *" _
      & " FROM ComprasRequerimientosDetalleMaterial det" _
      & " LEFT JOIN materiales mat ON mat.id = det.idMaterial" _
      & " LEFT JOIN grupos g ON g.id = mat.id_grupo" _
      & " LEFT JOIN rubros r ON r.id = g.id_rubro" _
      & " LEFT JOIN ComprasRequerimientosDetallesEntregas ent ON ent.id_detalle_material = det.id" _
      & " LEFT JOIN ComprasRequerimientosProveedores prov ON prov.idDetalleReque = det.id" _
      & " LEFT JOIN proveedores p ON p.id = prov.idProveedor" _
      & " LEFT JOIN AdminConfigMonedas ON p.id_moneda = AdminConfigMonedas.id" _
      & " WHERE 1 = 1"

    If LenB(filter) > 0 Then q = q & " AND " & filter


    Dim idx As New Dictionary

    Dim tmpEntrega As clsRequeEntregas

    Set r = conectar.RSFactory(q)
    BuildFieldsIndex r, idx

    While Not r.EOF
        Set MAT = New clsRequeMateriales
        If funciones.BuscarEnColeccion(col, CStr(conectar.GetValue(r, idx, "det", "id"))) Then
            Set MAT = col.item(CStr(conectar.GetValue(r, idx, "det", "id")))
        Else
            Set MAT = New clsRequeMateriales
            MAT.Id = conectar.GetValue(r, idx, "det", "id")
            MAT.Ancho = conectar.GetValue(r, idx, "det", "ancho")
            MAT.Largo = conectar.GetValue(r, idx, "det", "largo")
            MAT.Cantidad = conectar.GetValue(r, idx, "det", "cantidad")
            MAT.estado = conectar.GetValue(r, idx, "det", "estado")
            MAT.observaciones = conectar.GetValue(r, idx, "det", "detalle")
            MAT.RequeId = conectar.GetValue(r, idx, "det", "idReque")
            MAT.Material = DAOMateriales.Map(r, idx, "mat", , , "g", "r")


            col.Add MAT, CStr(MAT.Id)
        End If

        If Not IsEmpty(conectar.GetValue(r, idx, "ent", "id")) Then
            If Not funciones.BuscarEnColeccion(MAT.Entregas, CStr(conectar.GetValue(r, idx, "ent", "id"))) Then
                Set tmpEntrega = New clsRequeEntregas
                tmpEntrega.Id = conectar.GetValue(r, idx, "ent", "id")
                tmpEntrega.Tipo = material_
                tmpEntrega.Cantidad = conectar.GetValue(r, idx, "ent", "cantidad")
                tmpEntrega.FEcha = conectar.GetValue(r, idx, "ent", "fecha")
                MAT.Entregas.Add tmpEntrega, CStr(tmpEntrega.Id)
            End If
        End If

        If Not IsEmpty(conectar.GetValue(r, idx, "prov", "id")) Then
            If Not funciones.BuscarEnColeccion(MAT.ListaProveedores, CStr(conectar.GetValue(r, idx, "prov", "id"))) Then
                MAT.ListaProveedores.Add DAOProveedor.Map2(r, idx, "p"), CStr(conectar.GetValue(r, idx, "prov", "id"))
            End If
        End If

        r.MoveNext
    Wend

    Set Find = col
End Function

Public Function GetByRequeByProveedor(id_reque As Long, id_proveedor As Long) As Collection
'    Dim rs As Recordset
'    Dim col As New Collection
'    Set rs = conectar.RSFactory("select m.* from ComprasRequerimientosProveedores p inner join ComprasRequerimientosDetalleMaterial m on p.idDetalleReque=m.id where p.idProveedor=" & id_proveedor & " and m.idReque= " & id_reque)
'    Dim a As clsRequeMateriales
'
'    While Not rs.EOF And Not rs.BOF
'        Set a = New clsRequeMateriales
'        a.cantidad = rs!cantidad
'        a.id = rs!id
'        a.Largo = rs!Largo
'        a.ancho = rs!ancho
'        a.Entregas = DAORequeEntregas.GetEntregaById(rs!id, material_)
'        a.Material = DAOMateriales.FindById(rs!IdMaterial)
'        a.Observaciones = IIf(IsNull(rs!detalle), vbNullString, rs!detalle)
'        a.ListaProveedores = Nothing    '
'        a.id = rs!id
'        a.estado = rs!estado
'        col.Add a
'        rs.MoveNext
'    Wend
'
'    Set GetByRequeByProveedor = col
'    Set col = Nothing

    Set GetByRequeByProveedor = Find("prov.idProveedor=" & id_proveedor & " and det.idReque= " & id_reque)
End Function

Public Function GetById(Id As Long) As clsRequeMateriales
'    Dim a As clsRequeMateriales
'    Set rs = conectar.RSFactory("select * from ComprasRequerimientosDetalleMaterial where id=" & id)
'    If Not rs.EOF And Not rs.BOF Then
'        Set a = New clsRequeMateriales
'        a.cantidad = rs!cantidad
'        a.id = rs!id
'        a.Largo = rs!Largo
'        a.ancho = rs!ancho
'        a.Entregas = DAORequeEntregas.GetEntregaById(rs!id, material_)
'        a.Material = DAOMateriales.FindById(rs!IdMaterial)
'        a.Observaciones = IIf(IsNull(rs!detalle), vbNullString, rs!detalle)
'        a.estado = rs!estado
'        a.ListaProveedores = DAORequeProveedores.GetByDetalleReque(rs!id, material_)
'        a.id = rs!id
'    End If
'    Set GetById = a
    Dim col As Collection
    Set col = Find("det.id= " & Id)
    If col.count > 0 Then
        Set GetById = col.item(1)
    Else
        Set GetById = Nothing
    End If


End Function



Public Function GetByReque(id_reque As Long) As Collection
'    Dim col As New Collection
'    Dim a As clsRequeMateriales
'    Set rs = conectar.RSFactory("select * from ComprasRequerimientosDetalleMaterial where idReque=" & id_reque)
'    While Not rs.EOF
'        Set a = New clsRequeMateriales
'        a.cantidad = rs!cantidad
'        a.id = rs!id
'        a.Largo = rs!Largo
'        a.ancho = rs!ancho
'        a.Entregas = DAORequeEntregas.GetEntregaById(rs!id, material_)
'        a.Material = DAOMateriales.FindById(rs!IdMaterial)
'        a.Observaciones = IIf(IsNull(rs!detalle), vbNullString, rs!detalle)
'        a.ListaProveedores = DAORequeProveedores.GetByDetalleReque(rs!id, material_)
'        a.id = rs!id
'        a.estado = rs!estado
'        col.Add a
'        rs.MoveNext
'    Wend
'    Set a = Nothing
'    Set GetByReque = col
    Set GetByReque = Find("det.idReque= " & id_reque)
End Function
Public Function Save(T As clsRequerimiento) As Boolean
    On Error GoTo er1
    Save = True
    Dim A As clsRequeMateriales
    Dim Id_ As Long
    'traigo el ultimo ID de detalle_Reque


    Dim idsTmp As New Collection
    idsTmp.Add -1    'para que cuando uno los valores por lo menos haya uno
    For Each A In T.Materiales
        If A.Id <> 0 Then idsTmp.Add A.Id
    Next A
    Set rs = conectar.RSFactory("SELECT id FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & T.Id & " AND id NOT IN (" & funciones.JoinCollectionValues(idsTmp, ", ") & ")")
    Dim ids2Delete As New Collection
    ids2Delete.Add -1
    While Not rs.EOF And Not rs.BOF
        ids2Delete.Add rs.Fields("Id").value
        rs.MoveNext
    Wend

    If Not conectar.execute("delete from ComprasRequerimientosProveedores where idDetalleReque IN (" & funciones.JoinCollectionValues(idsTmp, ", ") & ")") Then GoTo er1
    If Not conectar.execute("delete from ComprasRequerimientosDetallesEntregas where id_detalle_material IN (" & funciones.JoinCollectionValues(idsTmp, ", ") & ")") Then GoTo er1
    If Not conectar.execute("delete from ComprasRequerimientosDetalleMaterial where id IN (" & funciones.JoinCollectionValues(ids2Delete, ", ") & ")") Then GoTo er1

    Dim q As String
    For Each A In T.Materiales

        If A.Id = 0 Then
            Save = conectar.execute("insert into ComprasRequerimientosDetalleMaterial (idMaterial,idReque, detalle, cantidad,largo,ancho, estado) Values (" & A.Material.Id & "," & T.Id & ",'" & A.observaciones & "'," & A.Cantidad & "," & A.Material.Largo & "," & A.Ancho & ", " & A.estado & ")")
            conectar.UltimoId "ComprasRequerimientosDetalleMaterial", Id_
            A.Id = Id_
        Else
            q = "UPDATE ComprasRequerimientosDetalleMaterial SET" _
              & " idReque = 'idReque'," _
              & " idMaterial = 'idMaterial' ," _
              & " detalle = 'detalle' ," _
              & " cantidad = 'cantidad' ," _
              & " ancho = 'ancho' ," _
              & " largo = 'largo' ," _
              & " estado = 'estado'" _
              & " WHERE id = 'id'"

            q = Replace$(q, "'idReque'", T.Id)
            q = Replace$(q, "'idMaterial'", A.Material.Id)
            q = Replace$(q, "'detalle'", conectar.Escape(A.observaciones))
            q = Replace$(q, "'cantidad'", conectar.Escape(A.Cantidad))
            q = Replace$(q, "'ancho'", conectar.Escape(A.Ancho))
            q = Replace$(q, "'largo'", conectar.Escape(A.Largo))
            q = Replace$(q, "'estado'", conectar.Escape(A.estado))
            q = Replace$(q, "'id'", A.Id)

            Save = conectar.execute(q)
            Id_ = A.Id
        End If

        If Save Then
            Save = DAORequeEntregas.saveAll(A.Entregas, Id_)
        Else
            GoTo er1
        End If

        If Save Then
            Save = DAORequeProveedores.Save(A.ListaProveedores, A.Id, material_)
            If Not Save Then GoTo er1
        Else
            GoTo er1
        End If

    Next A

    Exit Function
er1:
    Save = False
End Function

Public Function aPO(ByRef deta As clsRequeMateriales, ByRef RequeId As Long) As Boolean
    On Error GoTo E

    'If deta.estado <> EstadoRequeCompra.EnPO_ Then GoTo E

    'conectar.BeginTransaction

    Dim q As String
    q = "UPDATE ComprasRequerimientosDetalleMaterial SET estado = " & EstadoRequeCompra.EnPOParcial_ & " WHERE id = " & deta.Id
    If Not conectar.execute(q) Then GoTo E

    Dim r As Recordset
    Dim estadoReque As EstadoRequeCompra

    q = "SELECT (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & RequeId & ") as cantTot, " _
      & " (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & RequeId & " AND estado = " & EstadoRequeCompra.EnPOParcial_ & ") as cantPO"
    Set r = conectar.RSFactory(q)
    While Not r.EOF
        If r!cantPO = r!cantTot Then
            estadoReque = EnPO_
        Else
            estadoReque = EnPOParcial_
        End If
        r.MoveNext
    Wend

    q = "UPDATE ComprasRequerimientos SET estado = " & estadoReque & " WHERE id = " & RequeId
    If Not conectar.execute(q) Then GoTo E

    deta.estado = EnPOParcial_
    'reque.estado = estadoReque

    'conectar.CommitTransaction

    aPO = True
    Exit Function
E:
    aPO = False
    'conectar.RollBackTransaction
End Function


Public Function aprobar(ByRef deta As clsRequeMateriales, ByRef reque As clsRequerimiento) As Boolean
    On Error GoTo E

    If deta.estado <> EstadoRequeCompra.Finalizado_ Then GoTo E

    conectar.BeginTransaction

    Dim q As String
    q = "UPDATE ComprasRequerimientosDetalleMaterial SET estado = " & EstadoRequeCompra.Aprobado_ & " WHERE id = " & deta.Id
    If Not conectar.execute(q) Then GoTo E

    Dim r As Recordset
    Dim estadoReque As EstadoRequeCompra

    q = "SELECT (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & ") as cantTot, " _
      & " (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & " AND estado = " & EstadoRequeCompra.Aprobado_ & ") as cantAprob"
    Set r = conectar.RSFactory(q)
    While Not r.EOF
        If r!cantAprob = r!cantTot Then
            estadoReque = Aprobado_
        Else
            estadoReque = AprobadoParcial_
        End If
        r.MoveNext
    Wend

    q = "UPDATE ComprasRequerimientos SET estado = " & estadoReque & " WHERE id = " & reque.Id
    If Not conectar.execute(q) Then GoTo E

    deta.estado = Aprobado_
    reque.estado = estadoReque

    conectar.CommitTransaction

    aprobar = True
    Exit Function
E:
    aprobar = False
    conectar.RollBackTransaction
End Function


Public Function Anular(ByRef deta As clsRequeMateriales, ByRef reque As clsRequerimiento) As Boolean
    On Error GoTo E

    conectar.BeginTransaction

    Dim q As String
    q = "UPDATE ComprasRequerimientosDetalleMaterial SET estado = " & EstadoRequeCompra.Anulado & " WHERE id = " & deta.Id
    If Not conectar.execute(q) Then GoTo E

    Dim r As Recordset
    Dim estadoReque As EstadoRequeCompra

    q = "SELECT (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & ") as cantTot, " _
      & " (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & " AND estado = " & EstadoRequeCompra.Anulado & ") as cantAnu"
    Set r = conectar.RSFactory(q)
    While Not r.EOF
        If r!cantAnu = r!cantTot Then
            estadoReque = Anulado
            q = "UPDATE ComprasRequerimientos SET estado = " & estadoReque & " WHERE id = " & reque.Id
            If Not conectar.execute(q) Then GoTo E
            reque.estado = estadoReque
            'Else
            'estadoReque = AnuladoParcial
        End If
        r.MoveNext
    Wend


    deta.estado = Anulado

    conectar.CommitTransaction

    Anular = True
    Exit Function
E:
    Anular = False
    conectar.RollBackTransaction
End Function


Public Function procesarProveedores(ByRef deta As clsRequeMateriales, ByRef reque As clsRequerimiento) As Boolean
    On Error GoTo E
    conectar.BeginTransaction

    If deta.estado <> EstadoRequeCompra.Aprobado_ And deta.estado <> EstadoRequeCompra.EnProceso_ Then GoTo E


    Dim q As String
    q = "UPDATE ComprasRequerimientosDetalleMaterial SET estado = " & EstadoRequeCompra.EnProceso_ & " WHERE id = " & deta.Id
    If Not conectar.execute(q) Then GoTo E

    Dim r As Recordset
    Dim estadoReque As EstadoRequeCompra

    q = "SELECT (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & ") as cantTot, " _
      & " (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & " AND estado = " & EstadoRequeCompra.EnProceso_ & ") as cantEnProc"
    Set r = conectar.RSFactory(q)
    While Not r.EOF
        If r!cantEnProc = r!cantTot Then
            estadoReque = EnProceso_
        Else
            estadoReque = EnProcesoParcial_
        End If
        r.MoveNext
    Wend

    q = "UPDATE ComprasRequerimientos SET estado = " & estadoReque & " WHERE id = " & reque.Id
    If Not conectar.execute(q) Then GoTo E

    deta.estado = EnProceso_
    reque.estado = estadoReque

    conectar.CommitTransaction

    procesarProveedores = True
    Exit Function
E:
    procesarProveedores = False
    conectar.RollBackTransaction
End Function


Public Function finalizarProcesoProveedores(ByRef deta As clsRequeMateriales, ByRef reque As clsRequerimiento) As Boolean
    On Error GoTo E
    conectar.BeginTransaction

    If deta.estado <> EstadoRequeCompra.EnProceso_ Then GoTo E


    Dim q As String
    q = "UPDATE ComprasRequerimientosDetalleMaterial SET estado = " & EstadoRequeCompra.Procesado_ & " WHERE id = " & deta.Id
    If Not conectar.execute(q) Then GoTo E

    Dim r As Recordset
    Dim estadoReque As EstadoRequeCompra

    q = "SELECT (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & ") as cantTot, " _
      & " (SELECT COUNT(0) FROM ComprasRequerimientosDetalleMaterial WHERE idReque = " & reque.Id & " AND estado = " & EstadoRequeCompra.Procesado_ & ") as cantEnProc"
    Set r = conectar.RSFactory(q)
    While Not r.EOF
        If r!cantEnProc = r!cantTot Then
            estadoReque = EstadoRequeCompra.Procesado_
        Else
            estadoReque = ProcesadoParcial_
        End If
        r.MoveNext
    Wend

    q = "UPDATE ComprasRequerimientos SET estado = " & estadoReque & " WHERE id = " & reque.Id
    If Not conectar.execute(q) Then GoTo E

    deta.estado = EstadoRequeCompra.Procesado_
    reque.estado = estadoReque

    conectar.CommitTransaction

    finalizarProcesoProveedores = True
    Exit Function
E:
    finalizarProcesoProveedores = False
    conectar.RollBackTransaction
End Function



