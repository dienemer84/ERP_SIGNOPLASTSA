Attribute VB_Name = "DAOPeticionOfertaDetalle"
Option Explicit
Dim rs As ADODB.Recordset

Public Function FindAll(Optional ByVal POid As Long = 0, Optional filter As String = vbNullString, Optional withEntregas As Boolean = False) As Collection    'seguir para filtrar por proveedor
    On Error GoTo err1
    Dim col As New Collection
    Dim tmpDetalle As clsPeticionOfertaDetalle
    Dim idx As New Dictionary
    Dim strsql As String


    If withEntregas Then
        strsql = "SELECT *" _
               & " FROM ComprasPeticionOfertaDetalle pod" _
               & " LEFT JOIN ComprasPeticionOfertaDetalleEntregas pode ON pode.id_peticion_oferta_detalle_id = pod.id" _
               & " LEFT JOIN ComprasPeticionOferta po ON po.id = pod.id_peticion_oferta" _
               & " LEFT JOIN AdminConfigMonedas mon ON mon.id = po.moneda_id" _
               & " Where 1=1 "

    Else
        strsql = "SELECT *" _
               & " FROM ComprasPeticionOfertaDetalle pod" _
               & " LEFT JOIN ComprasPeticionOferta po ON po.id = pod.id_peticion_oferta" _
               & " LEFT JOIN AdminConfigMonedas mon ON mon.id = po.moneda_id" _
               & " Where 1=1 "

    End If


    'join con ComprasPeticionOferta por algun filtro


    If POid > 0 Then
        strsql = strsql & " AND pod.id_peticion_oferta = " & POid
    End If
    If LenB(filter) > 0 Then
        strsql = strsql & " AND " & filter
    End If

    strsql = strsql & " ORDER BY pod.id"


    If withEntregas Then
        Set rs = conectar.RSFactory(strsql)
        BuildFieldsIndex rs, idx

        Dim inCol As Boolean

        Dim ent As EntregaPetOfDetalle

        While Not rs.EOF

            inCol = funciones.BuscarEnColeccion(col, CStr(GetValue(rs, idx, "pod", "id")))

            If inCol Then
                Set tmpDetalle = col.item(CStr(GetValue(rs, idx, "pod", "id")))
            Else
                Set tmpDetalle = New clsPeticionOfertaDetalle
                tmpDetalle.DetalleReque = DAORequeMateriales.GetById(GetValue(rs, idx, "pod", "id_detalle_reque"))
                tmpDetalle.FechaValor = GetValue(rs, idx, "pod", "fecha")
                tmpDetalle.Terminado = GetValue(rs, idx, "pod", "finalizado")
                tmpDetalle.Valor = GetValue(rs, idx, "pod", "valor")
                tmpDetalle.Id = GetValue(rs, idx, "pod", "id")
                tmpDetalle.Cantidad = GetValue(rs, idx, "pod", "cantidad")
                tmpDetalle.ProveedorId = GetValue(rs, idx, "po", "id_proveedor")
                tmpDetalle.POid = GetValue(rs, idx, "pod", "id_peticion_oferta")
                tmpDetalle.estado = GetValue(rs, idx, "pod", "estado")
                Set tmpDetalle.moneda = DAOMoneda.Map(rs, idx, "mon")
            End If

            If withEntregas Then
                Set ent = New EntregaPetOfDetalle
                ent.Id = GetValue(rs, idx, "pode", "id")
                ent.FEcha = GetValue(rs, idx, "pode", "fecha")
                ent.Cantidad = GetValue(rs, idx, "pode", "cantidad")

                tmpDetalle.Entregas.Add ent, CStr(ent.Id)
            End If



            If Not inCol Then col.Add tmpDetalle, CStr(tmpDetalle.Id)

            rs.MoveNext
        Wend


    Else

        Set rs = conectar.RSFactory(strsql)
        BuildFieldsIndex rs, idx

        While Not rs.EOF

            Set tmpDetalle = New clsPeticionOfertaDetalle
            tmpDetalle.DetalleReque = DAORequeMateriales.GetById(GetValue(rs, idx, "pod", "id_detalle_reque"))
            tmpDetalle.FechaValor = GetValue(rs, idx, "pod", "fecha")
            tmpDetalle.Terminado = GetValue(rs, idx, "pod", "finalizado")
            tmpDetalle.Valor = GetValue(rs, idx, "pod", "valor")
            tmpDetalle.Id = GetValue(rs, idx, "pod", "id")
            tmpDetalle.Cantidad = GetValue(rs, idx, "pod", "cantidad")
            tmpDetalle.ProveedorId = GetValue(rs, idx, "po", "id_proveedor")
            tmpDetalle.POid = GetValue(rs, idx, "pod", "id_peticion_oferta")
            tmpDetalle.estado = GetValue(rs, idx, "pod", "estado")
            Set tmpDetalle.moneda = DAOMoneda.Map(rs, idx, "mon")
            col.Add tmpDetalle, CStr(tmpDetalle.Id)

            rs.MoveNext
        Wend

    End If
    Set FindAll = col

    Exit Function
err1:
    Set FindAll = Nothing

End Function
Public Function Guardar(T As clsPeticionOferta) As Boolean
    On Error GoTo err1
    Dim tmpDetalle As clsPeticionOfertaDetalle
    Dim strsql As String
    Dim q As String
    Guardar = True

    Dim id_peticion_oferta_detalle_id As Long

    Dim ent As EntregaPetOfDetalle
    Dim x As Long
    For x = 1 To T.detalle.count
        Set tmpDetalle = T.detalle.item(x)
        Dim A As Integer
        A = tmpDetalle.Terminado

        strsql = "insert into ComprasPeticionOfertaDetalle (id_detalle_reque, id_detalle_proveedor, valor, fecha, id_peticion_oferta, finalizado, cantidad)  values  (" & tmpDetalle.DetalleReque.Id & "," & T.Proveedor.Id & "," & tmpDetalle.Valor & ",'" & Format(tmpDetalle.FechaValor, "yyyy-mm-dd") & "'," & T.numero & "," & A & ", " & tmpDetalle.Cantidad & ")"

        If Not conectar.execute(strsql) Then GoTo err1

        If conectar.UltimoId("ComprasPeticionOfertaDetalle", id_peticion_oferta_detalle_id) Then

            For Each ent In tmpDetalle.Entregas
                q = "INSERT INTO ComprasPeticionOfertaDetalleEntregas" _
                  & " (id_peticion_oferta_detalle_id," _
                  & " cantidad," _
                  & " fecha)" _
                  & " VALUES (" & id_peticion_oferta_detalle_id & ", " _
                  & " " & ent.Cantidad & "," _
                  & " " & conectar.Escape(ent.FEcha) & ")"

                If Not conectar.execute(q) Then GoTo err1
            Next ent

        Else
            GoTo err1
        End If

        Guardar = True
    Next x
    Exit Function
err1:
    Guardar = False
End Function

Public Function Update(ByRef deta As clsPeticionOfertaDetalle, pet As clsPeticionOferta) As Boolean
    On Error GoTo E

    Dim q As String

    '    q = "DELETE FROM ComprasPeticionOfertaDetalle WHERE id = " & deta.Id
    '    If Not conectar.execute(q) Then GoTo E
    '
    '    q = "insert into ComprasPeticionOfertaDetalle (id_detalle_reque, id_detalle_proveedor, valor, fecha, id_peticion_oferta, finalizado)  values  (" & deta.DetalleReque.Id & "," & pet.proveedor.Id & "," & deta.Valor & ",'" & Format(deta.FechaValor, "yyyy-mm-dd") & "'," & pet.numero & "," & conectar.Escape(deta.Terminado) & ")"
    '    If Not conectar.execute(q) Then GoTo E
    '
    '    Dim l As Long
    '    If conectar.UltimoId("ComprasPeticionOfertaDetalle", l) Then
    '        deta.Id = l
    '    Else
    '        GoTo E
    '    End If

    'Update  = True

    q = "Update ComprasPeticionOfertaDetalle" _
      & " SET" _
      & " valor = " & conectar.Escape(deta.Valor) & " ," _
      & " cantidad = " & conectar.Escape(deta.Cantidad) & ", " _
      & " estado = " & conectar.Escape(deta.estado) _
      & " Where id = " & deta.Id

    Update = conectar.execute(q)
    If Not Update Then GoTo E

    q = "DELETE FROM ComprasPeticionOfertaDetalleEntregas WHERE id_peticion_oferta_detalle_id = " & deta.Id
    Update = conectar.execute(q)
    If Not Update Then GoTo E

    Dim ent As EntregaPetOfDetalle
    For Each ent In deta.Entregas
        q = "INSERT INTO ComprasPeticionOfertaDetalleEntregas " _
          & " (id_peticion_oferta_detalle_id, cantidad, fecha) VALUES (" _
          & deta.Id & ", " & ent.Cantidad & ", " & Escape(ent.FEcha) & ")"

        Update = conectar.execute(q)
        If Not Update Then GoTo E
    Next ent

    Exit Function
E:
    Update = False
End Function
