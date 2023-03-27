Attribute VB_Name = "DAODesarrolloManoObra"
Option Explicit

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_ID_PIEZA As String = "id_pieza"
Public Const CAMPO_CODIGO_TAREA As String = "codigo"
Public Const CAMPO_CANTIDAD As String = "cantidad"
Public Const CAMPO_TIEMPO As String = "tiempo"
Public Const CAMPO_DETALLE As String = "detalle"

Public Const TABLA_DESARROLLO_MANO_OBRA As String = "dmo"
Public Const TABLA_TAREAS As String = "tar"
Public Const TABLA_SECTOR As String = "sec"
Public Const TABLA_VALORES_MO As String = "vmo"
Public Const TABLA_MONEDA As String = "mon"

Public Function FindAllByPiezaId(piezaId As Long, Optional withPromedioHistorico As Boolean = False) As Collection
    Set FindAllByPiezaId = DAODesarrolloManoObra.FindAll(DAODesarrolloManoObra.TABLA_DESARROLLO_MANO_OBRA & "." & DAODesarrolloManoObra.CAMPO_ID_PIEZA & "=" & piezaId, withPromedioHistorico)
End Function

Public Function FindAll(Optional whereFilter As String = vbNullString, Optional withPromedioHistorico As Boolean = False) As Collection

    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String
    Dim desarrollosManoObra As New Collection

    q = "SELECT dmo.*, tar.*, sec.*, vmo.*, mon.*, cs.*" _
      & " FROM desarrollo_mdo dmo" _
      & " LEFT JOIN tareas tar ON tar.id = dmo.codigo" _
      & " LEFT JOIN sectores sec ON sec.id = tar.id_sector" _
      & " LEFT JOIN valores_MDO vmo ON vmo.id_tarea = tar.id" _
      & " LEFT JOIN categoria_sueldo cs ON cs.id = tar.categoria_sueldo_id" _
      & " LEFT JOIN AdminConfigMonedas mon ON mon.id = cs.id_moneda" _
      & " WHERE 1 = 1"

    If LenB(whereFilter) > 0 Then
        q = q & " AND " & whereFilter
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim dmo As DesarrolloManoObra

    While Not rs.EOF
        Set dmo = DAODesarrolloManoObra.Map(rs, fieldsIndex, _
                                            DAODesarrolloManoObra.TABLA_DESARROLLO_MANO_OBRA, _
                                            DAODesarrolloManoObra.TABLA_TAREAS, _
                                            DAODesarrolloManoObra.TABLA_SECTOR, _
                                            DAODesarrolloManoObra.TABLA_VALORES_MO, _
                                            DAODesarrolloManoObra.TABLA_MONEDA, "cs")

        If withPromedioHistorico Then
            dmo.TiempoPromedioHistorico = DAOTiemposProcesosDetalles.FindPromedioByTareaOfPieza(dmo.Tarea.Id, rs.Fields(fieldsIndex.item("dmo.id_pieza")))
        End If
        desarrollosManoObra.Add dmo, CStr(dmo.Id)
        rs.MoveNext
    Wend

    tickend = GetTickCount



    'Debug.Print tickEnd - tickStart, "ms elapsed"

    Set FindAll = desarrollosManoObra

End Function

Public Function Map(ByRef rs As Recordset, _
                    ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef tareasTableNameOrAlias As String = vbNullString, _
                    Optional ByRef sectorTableNameOrAlias As String = vbNullString, _
                    Optional ByRef valoresMOTableNameOrAlias As String = vbNullString, _
                    Optional ByRef monedaTableNameOrAlias As String = vbNullString, _
                    Optional ByRef CategoriaSueldoTabla As String = vbNullString _
                  ) As DesarrolloManoObra

    Dim dmo As DesarrolloManoObra
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloManoObra.CAMPO_ID)

    If Id > 0 Then
        Set dmo = New DesarrolloManoObra
        dmo.Id = Id
        dmo.Cantidad = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloManoObra.CAMPO_CANTIDAD)
        dmo.detalle = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloManoObra.CAMPO_DETALLE)
        dmo.Tiempo = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloManoObra.CAMPO_TIEMPO)

        If LenB(tareasTableNameOrAlias) > 0 Then Set dmo.Tarea = DAOTareas.Map(rs, fieldsIndex, tareasTableNameOrAlias, monedaTableNameOrAlias, sectorTableNameOrAlias, valoresMOTableNameOrAlias, CategoriaSueldoTabla)
    End If

    Set Map = dmo
End Function

Public Function Save(dmo As DesarrolloManoObra, Optional ByVal paraRevision As Boolean = False) As Boolean
    Dim q As String

    If dmo.Id = 0 Then
        q = "INSERT INTO {tabla} (id_pieza, codigo, cantidad, tiempo, detalle)" _
          & " Values" _
          & " ('id_pieza', 'codigo', 'cantidad', 'tiempo', 'detalle')"
    Else
        q = "UPDATE {tabla} SET" _
          & " id_pieza = 'id_pieza'," _
          & " codigo = 'codigo'," _
          & " cantidad = 'cantidad'," _
          & " tiempo = 'tiempo'," _
          & " detalle = 'detalle'" _
          & " Where id = 'id'"
    End If

    If paraRevision Then
        q = Replace$(q, "{tabla}", "desarrollo_mdo_rev")
    Else
        q = Replace$(q, "{tabla}", "desarrollo_mdo")
    End If


    q = Replace$(q, "'id_pieza'", dmo.Pieza.Id)
    q = Replace$(q, "'codigo'", dmo.Tarea.Id)
    q = Replace$(q, "'cantidad'", Escape(dmo.Cantidad))
    q = Replace$(q, "'tiempo'", Escape(dmo.Tiempo))
    q = Replace$(q, "'detalle'", Escape(dmo.detalle))
    q = Replace$(q, "'id'", dmo.Id)

    Save = conectar.execute(q)
End Function


