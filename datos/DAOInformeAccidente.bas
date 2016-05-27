Attribute VB_Name = "DAOInformeAccidente"
Option Explicit

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As InformeAccidente
    Dim a As InformeAccidente
    Dim id As Long
    id = GetValue(rs, indice, tabla, "id")
    If id <> 0 Then
        Set a = New InformeAccidente
        a.id = id
        a.Puesto = GetValue(rs, indice, tabla, "puesto")
        a.NombreTestigos = GetValue(rs, indice, tabla, "testigos")
        a.HsExtras = GetValue(rs, indice, tabla, "hs_extras")
        a.DescripcionHecho = GetValue(rs, indice, tabla, "descripcion_hecho")
        a.FallaMaquinasEquipos = GetValue(rs, indice, tabla, "falla_maquinas")
        a.FaltaElementosProteccionPersonal = GetValue(rs, indice, tabla, "falta_elementos_proteccion")
        a.ActoInseguro = GetValue(rs, indice, tabla, "acto_inseguro")
        a.Otros = GetValue(rs, indice, tabla, "otros")
        a.NaturalezaLesion = GetValue(rs, indice, tabla, "naturaleza_lesion")
        a.UbicacionLesion = GetValue(rs, indice, tabla, "ubicacion_lesion")
        a.FormaAccidente = GetValue(rs, indice, tabla, "forma_accidente")
        a.AgenteMaterial = GetValue(rs, indice, tabla, "agente_material")
        a.RecomendacionParaEvitarRepeticion = GetValue(rs, indice, tabla, "recomendaciones")
    End If
    Set Map = a
End Function

Public Function Save(acc As InformeAccidente) As Boolean
    On Error GoTo E

    Dim q As String


    If acc.id = 0 Then
        q = "INSERT INTO accidentes" _
            & " (puesto," _
            & " testigos," _
            & " hs_extras," _
            & " descripcion_hecho," _
            & " falla_maquinas," _
            & " falta_elementos_proteccion," _
            & " acto_inseguro," _
            & " otros," _
            & " naturaleza_lesion," _
            & " ubicacion_lesion," _
            & " forma_accidente," _
            & " agente_material," _
      & " recomendaciones) Values"
        q = q & " ('puesto'," _
            & " 'testigos'," _
            & " 'hs_extras'," _
            & " 'descripcion_hecho'," _
            & " 'falla_maquinas'," _
            & " 'falta_elementos_proteccion'," _
            & " 'acto_inseguro'," _
            & " 'otros'," _
            & " 'naturaleza_lesion'," _
            & " 'ubicacion_lesion'," _
            & " 'forma_accidente'," _
            & " 'agente_material'," _
            & " 'recomendaciones')"
    Else
        q = "Update accidentes " _
            & " SET" _
            & " puesto = 'puesto' ," _
            & " testigos = 'testigos' ," _
            & " hs_extras = 'hs_extras' ," _
            & " descripcion_hecho = 'descripcion_hecho' ," _
            & " falla_maquinas = 'falla_maquinas' ," _
            & " falta_elementos_proteccion = 'falta_elementos_proteccion' ," _
            & " acto_inseguro = 'acto_inseguro' ," _
            & " otros = 'otros' ," _
            & " naturaleza_lesion = 'naturaleza_lesion' ," _
            & " ubicacion_lesion = 'ubicacion_lesion' ," _
            & " forma_accidente = 'forma_accidente' ," _
            & " agente_material = 'agente_material' ," _
            & " recomendaciones = 'recomendaciones'" _
            & " WHERE id = 'id'"
    End If

    q = Replace$(q, "'puesto'", conectar.Escape(acc.Puesto))
    q = Replace$(q, "'testigos'", conectar.Escape(acc.NombreTestigos))
    q = Replace$(q, "'hs_extras'", conectar.Escape(acc.HsExtras))
    q = Replace$(q, "'descripcion_hecho'", conectar.Escape(acc.DescripcionHecho))
    q = Replace$(q, "'falla_maquinas'", conectar.Escape(acc.FallaMaquinasEquipos))
    q = Replace$(q, "'falta_elementos_proteccion'", conectar.Escape(acc.FaltaElementosProteccionPersonal))
    q = Replace$(q, "'acto_inseguro'", conectar.Escape(acc.ActoInseguro))
    q = Replace$(q, "'naturaleza_lesion'", conectar.Escape(acc.NaturalezaLesion))
    q = Replace$(q, "'ubicacion_lesion'", conectar.Escape(acc.UbicacionLesion))
    q = Replace$(q, "'forma_accidente'", conectar.Escape(acc.FormaAccidente))
    q = Replace$(q, "'agente_material'", conectar.Escape(acc.AgenteMaterial))
    q = Replace$(q, "'recomendaciones'", conectar.Escape(acc.RecomendacionParaEvitarRepeticion))
    q = Replace$(q, "'id'", conectar.GetEntityId(acc))

    Save = conectar.execute(q)
    If Save And acc.id = 0 Then
        acc.id = conectar.UltimoId2()
        Save = (acc.id <> 0)
    End If

    Exit Function
E:
    Save = False
End Function
