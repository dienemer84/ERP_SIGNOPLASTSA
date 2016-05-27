Attribute VB_Name = "DAOTiempoProcesoPlanificado"
Option Explicit

Dim rs As ADODB.Recordset

Public Function Save(ptpp As TiempoProcesoPlanificado) As Boolean
    On Error GoTo er1
    Dim strsql As String
    Dim a As Long
    Save = True
    Dim n As Boolean

    If ptpp.id = 0 Then
        n = True
        strsql = "INSERT INTO sp.PlaneamientoTiemposProcesosPlanificacion   (" _
                 & "id_ptp,   inicio,   Fin, color, critica,prioridad  ) " _
                 & " Values " _
                 & "('id_ptp',  'inicio',  'fin', 'color', 'critica','prioridad' );"

    Else
        strsql = "Update sp.PlaneamientoTiemposProcesosPlanificacion   SET " _
                 & "  id_ptp = 'id_ptp' ,  inicio = 'inicio' ,   fin = 'fin' , color ='color' , critica='critica' , prioridad ='prioridad' " _
                 & " Where     id = 'id'  "

        n = False
    End If

    strsql = Replace(strsql, "'id'", conectar.GetEntityId(ptpp))
    strsql = Replace(strsql, "'inicio'", conectar.Escape(ptpp.Inicio))
    strsql = Replace(strsql, "'critica'", conectar.Escape(ptpp.Critica))
    strsql = Replace(strsql, "'color'", conectar.Escape(ptpp.Color))
    strsql = Replace(strsql, "'fin'", conectar.Escape(ptpp.Fin))
    strsql = Replace(strsql, "'prioridad'", conectar.Escape(ptpp.Prioridad))
    strsql = Replace(strsql, "'id_ptp'", conectar.Escape(ptpp.idTiempoProceso))


    Save = conectar.execute(strsql)

    If n Then ptpp.id = conectar.UltimoId2


    Exit Function
er1:
    Save = False

End Function

Public Function FindByIdTiempoProceso(id As Long) As clsRubros
    Dim col As Collection
    Set col = FindAll("PlaneamientoTiemposProcesos.id = " & id)
    If col.count = 0 Then
        Set FindByIdTiempoProceso = Nothing
    Else
        Set FindByIdTiempoProceso = col.item(1)
    End If
End Function

Public Function FindById(id As Long) As clsRubros
    Dim col As Collection
    Set col = FindAll("PlaneamientoTiemposProcesosPlanificacion.id = " & id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function

Public Function FindAll(Optional filter As String = " WHERE 1=1") As Collection
    Dim rs As ADODB.Recordset
    Dim q As String
    Dim ptpp As New Collection
    Dim r As TiempoProcesoPlanificado

    q = "SELECT * " _
        & "From PlaneamientoTiemposProcesosPlanificacion " _
        & "LEFT JOIN tareas " _
        & "ON PlaneamientoTiemposProcesos.codigoTarea=tarea.id " _
        & "WHERE " & filter


    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex



    While Not rs.EOF
        Set r = Map(rs, fieldsIndex, "PlaneamientoTiemposProcesosPlanificacion", "tareas")
        ptpp.Add r, CStr(r.id)
        rs.MoveNext
    Wend

    Set FindAll = ptpp
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaTarea As String = vbNullString) As TiempoProcesoPlanificado
    Dim r As TiempoProcesoPlanificado
    Dim id As Long

    id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set r = New TiempoProcesoPlanificado
        r.id = id
        r.Fin = GetValue(rs, indice, tabla, "fin")
        r.Inicio = GetValue(rs, indice, tabla, "inicio")
        r.Color = GetValue(rs, indice, tabla, "color")
        r.Critica = GetValue(rs, indice, tabla, "critica")
        r.Prioridad = GetValue(rs, indice, tabla, "prioridad")
        r.idTiempoProceso = GetValue(rs, indice, tabla, "id_ptp")
    End If

    Set Map = r
End Function

