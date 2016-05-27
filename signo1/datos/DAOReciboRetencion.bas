Attribute VB_Name = "DAOReciboRetencion"
Option Explicit
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_RECIBO As String = "idRecibo"
Public Const CAMPO_VALOR As String = "valor"
Public Const CAMPO_NRO_RET As String = "nroRetencion"
Public Const CAMPO_FECHA As String = "fecha"
Public Const TABLA_RECXRET As String = "acdr"
Public Const TABLA_RETENCION As String = "acr"

Public Function FindAllByRecibo(reciboId As Long) As Collection
    Set FindAllByRecibo = FindAll("idRecibo = " & reciboId)
End Function


Public Function FindAll(Optional filter As String = " 1 = 1") As Collection
    Dim rs As Recordset
    Dim indice As Dictionary
    Dim col As New Collection
    Dim q As String
    q = "SELECT acr.*,acdr.* FROM AdminRecibosDetalleRetenciones acdr LEFT JOIN retenciones acr ON acdr.idRetencion=acr.id WHERE " & filter
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, indice
    Dim ret As retencionRecibo

    While Not rs.EOF
        Set ret = DAOReciboRetencion.Map(rs, indice, TABLA_RECXRET, TABLA_RETENCION)
        col.Add ret, CStr(ret.id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function



Public Function Map(rs As Recordset, indice As Dictionary, Tabla1 As String, TablaRetenciones As String) As retencionRecibo
    Dim id As Long
    Dim T As retencionRecibo
    id = GetValue(rs, indice, Tabla1, DAORetenciones.CAMPO_ID)
    If id <> 0 Then
        Set T = New retencionRecibo
        T.id = id
        T.idRecibo = GetValue(rs, indice, Tabla1, DAOReciboRetencion.CAMPO_RECIBO)
        T.NroRetencion = GetValue(rs, indice, Tabla1, DAOReciboRetencion.CAMPO_NRO_RET)
        Set T.Retencion = DAORetenciones.Map(rs, indice, TablaRetenciones)
        T.Valor = GetValue(rs, indice, Tabla1, DAOReciboRetencion.CAMPO_VALOR)
        T.FEcha = GetValue(rs, indice, Tabla1, DAOReciboRetencion.CAMPO_FECHA)
    End If

    Set Map = T
End Function

Public Function Delete(Optional filter As String = "1 = 1") As Boolean
    Delete = conectar.execute("DELETE FROM AdminRecibosDetalleRetenciones WHERE  " & filter)
End Function

Public Function Save(Retencion As retencionRecibo, recibo As recibo) As Boolean
    Dim q As String


    q = "INSERT INTO AdminRecibosDetalleRetenciones" _
        & " (idRecibo," _
        & " idRetencion," _
        & " valor," _
        & " nroRetencion,fecha)" _
        & " Values" _
        & " ('idRecibo'," _
        & " 'idRetencion'," _
        & " 'valor'," _
        & " 'nroRetencion'," _
        & " 'fecha')"

    q = Replace(q, "'idRecibo'", conectar.GetEntityId(recibo))
    q = Replace(q, "'idRetencion'", conectar.GetEntityId(Retencion.Retencion))
    q = Replace(q, "'valor'", conectar.Escape(Retencion.Valor))
    q = Replace(q, "'nroRetencion'", conectar.Escape(Retencion.NroRetencion))
    q = Replace(q, "'fecha'", conectar.Escape(Retencion.FEcha))


    Save = conectar.execute(q)
End Function


