Attribute VB_Name = "DAOTareas"
Option Explicit
Dim rs As ADODB.Recordset

Public Const CAMPO_ID As String = "id"
Public Const CAMPO_ID_SECTOR As String = "id_sector"
Public Const CAMPO_NOMBRE_TAREA As String = "tarea"
Public Const CAMPO_CANT_X_PROC As String = "cantxproc"

Public Const CAMPO_VMDO_ID As String = "id"
Public Const CAMPO_VMDO_DESCRIPCION As String = "descripcion"
Public Const CAMPO_VMDO_VALOR As String = "valor"
Public Const CAMPO_VMDO_FECHA As String = "fecha"
Public Const CAMPO_VMDO_ID_MONEDA As String = "id_moneda"

Public Const TABLA_TAREA As String = "t"
Public Const TABLA_MONEDA As String = "mon"
Public Const TABLA_SECTOR As String = "s"
Public Const TABLA_VMDO As String = "vmdo"
Public Const TABLA_CATEGORIA_SUELDO As String = "cs"


Public Function FindById(id As Long) As clsTarea
    Dim col As Collection
    Set col = DAOTareas.FindAll(DAOTareas.TABLA_TAREA & "." & DAOTareas.CAMPO_ID & " = " & id)
    If col.count > 0 Then
        Set FindById = col.item(1)
    Else
        Set FindById = Nothing
    End If
End Function

Public Function FindAll(Optional whereFilter As String) As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount
    Dim rs As ADODB.Recordset
    Dim q As String
    Dim tareas As New Collection

    q = "SELECT" _
        & " t.*," _
        & " s.*," _
        & " vmdo.*," _
        & " mon.*," _
        & " cs.*" _
        & " FROM tareas t" _
        & " LEFT JOIN sectores s" _
        & " ON s.id = t.id_sector" _
        & " LEFT JOIN valores_MDO vmdo" _
        & " ON vmdo.id_tarea = t.id" _
        & " LEFT JOIN categoria_sueldo cs" _
        & " ON cs.id = t.categoria_sueldo_id" _
        & " LEFT JOIN AdminConfigMonedas mon" _
        & " ON mon.id = cs.id_moneda" _
        & " WHERE 1 = 1"


    If LenB(whereFilter) > 0 Then
        q = q & " AND " & whereFilter
    End If

    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim tar As clsTarea

    While Not rs.EOF
        Set tar = DAOTareas.Map(rs, fieldsIndex, DAOTareas.TABLA_TAREA, DAOTareas.TABLA_MONEDA, DAOTareas.TABLA_SECTOR, DAOTareas.TABLA_VMDO, DAOTareas.TABLA_CATEGORIA_SUELDO)
        If Not BuscarEnColeccion(tareas, CStr(tar.id)) Then    'por la id_tarea 59 que esta 2 veces
            tareas.Add tar, CStr(tar.id)
        End If
        rs.MoveNext
    Wend

    Set FindAll = tareas
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional monedaTabla As String = vbNullString, Optional sectorTabla As String = vbNullString, Optional valorMDOTabla As String = vbNullString, Optional CategoriaSueldoTabla As String = vbNullString) As clsTarea
    Dim T As clsTarea
    Dim id As Long

    id = GetValue(rs, indice, tabla, DAOTareas.CAMPO_ID)

    If id > 0 Then
        Set T = New clsTarea
        T.id = id
        T.Tarea = GetValue(rs, indice, tabla, DAOTareas.CAMPO_NOMBRE_TAREA)
        T.CantPorProc = GetValue(rs, indice, tabla, DAOTareas.CAMPO_CANT_X_PROC)

        T.SectorID = GetValue(rs, indice, tabla, "id_sector")

        If LenB(valorMDOTabla) > 0 Then
            T.IdValorMDO = GetValue(rs, indice, valorMDOTabla, DAOTareas.CAMPO_VMDO_ID)
            T.Valor = GetValue(rs, indice, valorMDOTabla, DAOTareas.CAMPO_VMDO_VALOR)
            T.FEcha = GetValue(rs, indice, valorMDOTabla, DAOTareas.CAMPO_VMDO_FECHA)
            T.descripcion = GetValue(rs, indice, valorMDOTabla, DAOTareas.CAMPO_VMDO_DESCRIPCION)
        End If

        If LenB(sectorTabla) > 0 Then Set T.Sector = DAOSectores.Map(rs, indice, sectorTabla)
        If LenB(CategoriaSueldoTabla) > 0 Then Set T.CategoriaSueldo = DAOCategoriaSueldo.Map(rs, indice, CategoriaSueldoTabla, monedaTabla)

    End If

    Set Map = T
End Function

Public Function Save(Tarea As clsTarea) As Boolean
    On Error GoTo E

    Dim q As String
    conectar.BeginTransaction

    If Tarea.id = 0 Then
        Dim tareaId As Long

        q = "insert into tareas (id_sector,cantxproc,tarea, categoria_sueldo_id) VALUES (" & conectar.Escape(Tarea.Sector.id) & "," & conectar.Escape(Tarea.CantPorProc) & "," & conectar.Escape(Tarea.Tarea) & "," & GetEntityId(Tarea.CategoriaSueldo) & ")"
        If Not conectar.execute(q) Then GoTo E

        conectar.UltimoId "tareas", tareaId
        q = "insert into valores_MDO (id_tarea,descripcion,valor,fecha) VALUES (" & tareaId & "," & conectar.Escape(Tarea.descripcion) & "," & conectar.Escape(Tarea.Valor) & "," & conectar.Escape(Tarea.FEcha) & ")"
        If Not conectar.execute(q) Then GoTo E

    Else
        q = "update tareas set id_sector=" & conectar.Escape(Tarea.Sector.id) & ", cantxproc= " & conectar.Escape(Tarea.CantPorProc) & ", tarea=" & conectar.Escape(Tarea.Tarea) & ", categoria_sueldo_id = " & GetEntityId(Tarea.CategoriaSueldo) & " where id=" & Tarea.id
        If Not conectar.execute(q) Then GoTo E

        q = "update valores_MDO set descripcion= " & conectar.Escape(Tarea.descripcion) & ", valor=" & conectar.Escape(Tarea.Valor) & ",fecha=" & conectar.Escape(Tarea.FEcha) & " where id_tarea=" & Tarea.id
        If Not conectar.execute(q) Then GoTo E
    End If

    conectar.CommitTransaction
    Save = True
    Exit Function
E:
    Save = False
    conectar.RollBackTransaction
End Function

Public Function LlenarComboPorSector(cbo As ComboBox, Sector As clsSector)
    Dim col As Collection
    cbo.Clear
    Set col = DAOTareas.FindAll(DAOTareas.TABLA_TAREA & "." & DAOTareas.CAMPO_ID_SECTOR & "=" & Sector.id)

    Dim T As clsTarea

    For Each T In col
        cbo.AddItem T.Description
        cbo.ItemData(cbo.NewIndex) = T.id
    Next
End Function
