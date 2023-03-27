Attribute VB_Name = "DAODetallePresupuestoHistorico"
Option Explicit

Public Function FindAllByDetallePresupuestoId(id_detalle_presupuesto As Long) As Collection
    Set FindAllByDetallePresupuestoId = DAODetallePresupuestoHistorico.FindAll(id_detalle_presupuesto, "dph.id_detalle_presupuesto = " & id_detalle_presupuesto)
End Function

Private Function FindAll(ByVal id_detalle_presupuesto As Long, Optional ByVal filter As String = vbNullString) As Collection
    On Error GoTo E

    'Dim tickStart As Double
    'Dim tickEnd As Double
    'tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim detallesHistoricos As New Collection

    q = "SELECT dph.* FROM detalle_presupuesto_historico dph WHERE 1 = 1"

    If LenB(filter) > 0 Then
        q = q & " AND " & filter
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim findex As Dictionary

    Dim presu As clsPresupuesto
    Dim detaPresu As clsPresupuestoDetalle
    Set detaPresu = DAOPresupuestosDetalle.GetAllById(id_detalle_presupuesto)

    If Not detaPresu Is Nothing Then
        Dim q2 As String
        q2 = "SELECT idPresupuesto FROM detalle_presupuesto WHERE id = " & detaPresu.Id
        Dim r3 As Recordset
        Set r3 = RSFactory(q2)

        Set presu = DAOPresupuestos.GetById(r3!idpresupuesto)
        Set detaPresu.presupuesto = presu
    End If

    Dim pdh As clsPresupuestoDetalleHistorico
    Dim detaPadre As clsPresupuestoDetalleHistorico
    Dim pdhmdo As PresupuestoDetalleHistoricoMDO
    Dim pdhm As PresupuestoDetalleHistoricoMAT
    Dim r2 As Recordset
    Dim pdh33 As clsPresupuestoDetalleHistorico

    While Not rs.EOF
        Set pdh = Map(rs, fieldsIndex)
        Set pdh.Pieza = DAOPieza.FindById(rs!PIEZA_ID, FL_0, True, True)
        Set pdh.DetallePresupuesto = DAOPresupuestosDetalle.GetAllById(rs!id_detalle_presupuesto)

        q = "SELECT dphmo.* FROM detalle_presupuesto_historico_mdo dphmo WHERE dphmo.id_detalle_presupuesto_historico = " & pdh.Id
        Set r2 = RSFactory(q)
        Set findex = Nothing
        BuildFieldsIndex r2, findex

        While Not r2.EOF
            Set pdhmdo = New PresupuestoDetalleHistoricoMDO
            pdhmdo.Id = GetValue(r2, findex, "dphmo", "id")
            pdhmdo.CantOperarios = GetValue(r2, findex, "dphmo", "cantidad")
            pdhmdo.Tiempo = GetValue(r2, findex, "dphmo", "tiempo")
            pdhmdo.Valor = GetValue(r2, findex, "dphmo", "valor")  '* pdhmdo.CantOperarios * pdhmdo.Tiempo
            If Not IsNull(r2!tarea_id) Then
                Set pdhmdo.Tarea = DAOTareas.FindById(r2!tarea_id)
            End If
            pdh.historicoMDO.Add pdhmdo, CStr(pdhmdo.Id)

            r2.MoveNext
        Wend

        q = "SELECT dphm.*, mon.* FROM detalle_presupuesto_historico_mat dphm left join AdminConfigMonedas mon on dphm.id_moneda=mon.id WHERE dphm.id_detalle_presupuesto_historico = " & pdh.Id
        Set r2 = RSFactory(q)
        Set findex = Nothing
        BuildFieldsIndex r2, findex
        While Not r2.EOF
            Set pdhm = New PresupuestoDetalleHistoricoMAT
            pdhm.Id = GetValue(r2, findex, "dphm", "id")
            pdhm.Ancho = GetValue(r2, findex, "dphm", "ancho")
            pdhm.AnchoPieza = GetValue(r2, findex, "dphm", "ancho_pieza")
            pdhm.Cantidad = GetValue(r2, findex, "dphm", "cantidad")
            pdhm.Largo = GetValue(r2, findex, "dphm", "largo")
            pdhm.LargoPieza = GetValue(r2, findex, "dphm", "largo_pieza")
            pdhm.Scrap = GetValue(r2, findex, "dphm", "Scrap")
            pdhm.Valor = GetValue(r2, findex, "dphm", "valor")
            If Not IsNull(r2!id_moneda) Then Set pdhm.moneda = DAOMoneda.Map(r2, findex, "mon")

            If Not IsNull(r2!id_detalle_presupuesto_historico) Then
                Set pdhm.Material = DAOMateriales.FindById(r2!material_id, False)
            End If

            pdh.HistoricoMAT.Add pdhm, CStr(pdhm.Id)

            r2.MoveNext
        Wend


        '        If Not IsNull(rs!id_detalle_presupuesto_historico_padre) Then
        '            If funciones.BuscarEnColeccion(detallesHistoricos, CStr(rs!id_detalle_presupuesto_historico_padre)) Then
        '                Set detaPadre = detallesHistoricos.Item(CStr(rs!id_detalle_presupuesto_historico_padre))
        '                If Not funciones.BuscarEnColeccion(detaPadre.HistoricoHijos, CStr(pdh.id)) Then
        '                    detaPadre.HistoricoHijos.Add pdh, CStr(pdh.id)
        '                End If
        '            End If
        '        End If

        If Not detaPresu Is Nothing Then
            Set pdh.DetallePresupuesto = detaPresu
        End If

        If IsNull(rs!id_detalle_presupuesto_historico_padre) Then
            detallesHistoricos.Add pdh, CStr(pdh.Id)
        Else
            Set pdh33 = FindItemInCollection(detallesHistoricos, CStr(rs!id_detalle_presupuesto_historico_padre))
            If Not pdh33 Is Nothing Then
                pdh33.HistoricoHijos.Add pdh, CStr(pdh.Id)
            End If
        End If

        rs.MoveNext
    Wend


    'tickEnd = GetTickCount
    'Debug.Print tickEnd - tickStart, "ms elapsed"

    Set FindAll = detallesHistoricos
    Exit Function
E:
    MsgBox Err.Description
    '   Stop
    '   Resume
End Function

Public Function FindItemInCollection(col As Collection, Id As String) As clsPresupuestoDetalleHistorico
    Dim tmp As clsPresupuestoDetalleHistorico
    For Each tmp In col
        If tmp.Id = Id Then Set FindItemInCollection = tmp
        If Not FindItemInCollection Is Nothing Then Exit For
        Set FindItemInCollection = FindItemInCollection(tmp.HistoricoHijos, Id)
        If Not FindItemInCollection Is Nothing Then Exit For
    Next tmp
End Function

Private Function Map(rs As Recordset, fieldsIndex As Dictionary) As clsPresupuestoDetalleHistorico
    Dim pdh As clsPresupuestoDetalleHistorico
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, "dph", "id")

    If Id > 0 Then
        Set pdh = New clsPresupuestoDetalleHistorico
        pdh.Id = Id
        pdh.NombrePieza = GetValue(rs, fieldsIndex, "dph", "nombre_pieza")
        pdh.FEcha = GetValue(rs, fieldsIndex, "dph", "fecha")
        'If LenB(ivaTableNameOrAlias) > 0 Then Set c.tipoIva = DAOTipoIva.Map(rs, fieldsIndex, ivaTableNameOrAlias, tipoFacturaTableNameOrAlias)
    End If

    Set Map = pdh

End Function
