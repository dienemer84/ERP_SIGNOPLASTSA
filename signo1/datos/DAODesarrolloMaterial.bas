Attribute VB_Name = "DAODesarrolloMaterial"
Option Explicit


Public Const CAMPO_ID As String = "id"
Public Const CAMPO_ID_PIEZA As String = "id_pieza"
Public Const CAMPO_SCRAP As String = "scrap"
Public Const CAMPO_LARGO As String = "largo"
Public Const CAMPO_ANCHO As String = "ancho"
Public Const CAMPO_LARGO_TERM As String = "LargoTerm"
Public Const CAMPO_ANCHO_TERM As String = "AnchoTerm"
Public Const CAMPO_ID_MATERIAL As String = "id_material"
Public Const CAMPO_CANTIDAD As String = "cantidad"
Public Const CAMPO_DETALLE As String = "detalle"

Public Const TABLA_DESARROLLO_MATERIAL As String = "dm"
Public Const TABLA_MATERIAL As String = "m"
Public Const TABLA_GRUPO As String = "g"
Public Const TABLA_RUBRO As String = "r"
Public Const TABLA_ALMACEN As String = "a"
Public Const TABLA_MONEDA As String = "mon"

Public Function FindAllByPiezaId(piezaId As Long) As Collection
    Set FindAllByPiezaId = DAODesarrolloMaterial.FindAll(DAODesarrolloMaterial.TABLA_DESARROLLO_MATERIAL & "." & DAODesarrolloMaterial.CAMPO_ID_PIEZA & "=" & piezaId)
End Function

Public Function FindAll(Optional whereFilter As String = vbNullString) As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim DesarrollosMaterial As New Collection

    q = "SELECT " _
        & " dm.*, g.*, r.*, a.*, mon.*, m.*" _
        & " FROM desarrollo_material dm" _
        & " INNER JOIN materiales m" _
        & " ON m.id = dm.id_material" _
        & " LEFT JOIN grupos g" _
        & " ON g.id = m.id_grupo" _
        & " LEFT JOIN rubros r" _
        & " ON r.id = g.id_rubro" _
        & " LEFT JOIN materialesAlmacenes a" _
        & " ON a.id = m.idAlmacen" _
        & " LEFT JOIN AdminConfigMonedas mon" _
        & " ON mon.id = m.id_moneda" _
        & " WHERE 1 = 1"

    If LenB(whereFilter) > 0 Then
        q = q & " AND " & whereFilter
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim dmat As DesarrolloMaterial

    While Not rs.EOF
        Set dmat = DAODesarrolloMaterial.Map(rs, fieldsIndex, DAODesarrolloMaterial.TABLA_DESARROLLO_MATERIAL, DAODesarrolloMaterial.TABLA_MATERIAL, DAODesarrolloMaterial.TABLA_GRUPO, DAODesarrolloMaterial.TABLA_RUBRO, DAODesarrolloMaterial.TABLA_ALMACEN, DAODesarrolloMaterial.TABLA_MONEDA)
        DesarrollosMaterial.Add dmat, CStr(dmat.id)
        rs.MoveNext
    Wend


    tickend = GetTickCount


    'Debug.Print "DaoDesarrolloMaterial.FindAll()", tickEnd - tickStart, "ms elapsed"

    Set FindAll = DesarrollosMaterial

End Function

Public Function Map(ByRef rs As Recordset, _
                    ByRef fieldsIndex As Dictionary, _
                    ByRef tableNameOrAlias As String, _
                    Optional ByRef materialTableNameOrAlias As String = vbNullString, _
                    Optional ByRef grupoTableNameOrAlias As String = vbNullString, _
                    Optional ByRef rubroTableNameOrAlias As String = vbNullString, _
                    Optional ByRef almacenTableNameOrAlias As String = vbNullString, _
                    Optional ByRef monedaTableNameOrAlias As String = vbNullString _
                    ) As DesarrolloMaterial

    Dim dm As DesarrolloMaterial
    Dim id As Variant

    id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_ID)

    If id > 0 Then
        Set dm = New DesarrolloMaterial
        dm.id = id
        dm.Ancho = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_ANCHO)
        dm.AnchoTerm = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_ANCHO_TERM)
        dm.Cantidad = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_CANTIDAD)
        dm.detalle = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_DETALLE)
        dm.Largo = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_LARGO)
        dm.LargoTerm = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_LARGO_TERM)
        If LenB(materialTableNameOrAlias) > 0 Then Set dm.Material = DAOMateriales.Map(rs, fieldsIndex, materialTableNameOrAlias, almacenTableNameOrAlias, monedaTableNameOrAlias, grupoTableNameOrAlias, rubroTableNameOrAlias)
        dm.Scrap = GetValue(rs, fieldsIndex, tableNameOrAlias, DAODesarrolloMaterial.CAMPO_SCRAP)
    End If

    Set Map = dm
End Function


Public Function Save(dm As DesarrolloMaterial, Optional ByVal paraRevision As Boolean = False) As Boolean

    Dim q As String

    If dm.id = 0 Then
        q = "INSERT INTO {tabla} " _
            & " (id_pieza,scrap,largo,ancho,LargoTerm,AnchoTerm,id_material,cantidad,detalle)" _
            & " Values" _
            & " ('id_pieza','scrap','largo','ancho','LargoTerm','AnchoTerm','id_material','cantidad','detalle')"
    Else
        q = "Update {tabla} " _
            & " SET" _
            & " id_pieza = 'id_pieza'," _
            & " scrap = 'scrap'," _
            & " largo = 'largo'," _
            & " ancho = 'ancho'," _
            & " LargoTerm = 'LargoTerm'," _
            & " AnchoTerm = 'AnchoTerm'," _
            & " id_material = 'id_material'," _
            & " cantidad = 'cantidad'," _
            & " detalle = 'detalle'" _
            & " WHERE id = 'id'"
    End If

    If paraRevision Then
        q = Replace$(q, "{tabla}", "desarrollo_material_rev")
    Else
        q = Replace$(q, "{tabla}", "desarrollo_material")
    End If

    q = Replace$(q, "'id_pieza'", GetEntityId(dm.Pieza))
    q = Replace$(q, "'scrap'", Escape(dm.Scrap))
    q = Replace$(q, "'largo'", Escape(dm.Largo))
    q = Replace$(q, "'ancho'", Escape(dm.Ancho))
    q = Replace$(q, "'LargoTerm'", Escape(dm.LargoTerm))
    q = Replace$(q, "'AnchoTerm'", Escape(dm.AnchoTerm))
    q = Replace$(q, "'id_material'", GetEntityId(dm.Material))
    q = Replace$(q, "'cantidad'", Escape(dm.Cantidad))
    q = Replace$(q, "'detalle'", Escape(dm.detalle))
    q = Replace$(q, "'id'", GetEntityId(dm))

    Save = conectar.execute(q)
End Function

