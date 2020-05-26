Attribute VB_Name = "DAOCompensatorios"
Option Explicit


Public Function Save(c As Compensatorio) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    If Not Guardar(c) Then GoTo err1

    conectar.CommitTransaction
    Exit Function

err1:
    conectar.RollBackTransaction
End Function



Public Function Guardar(c As Compensatorio) As Boolean
    Dim q As String
    On Error GoTo err2
    Guardar = True
    If c.id = 0 Then
        q = "INSERT INTO sp.ordenes_pago_compensatorios  (id_comprobante,   fecha,   importe,   Observacion, tipo ,id_orden_pago, " _
            & "neto_gravado_compensado, alicuota_percepcion, monto_a_percibir, cancelado)" _
            & " values  " _
            & " ( 'id_comprobante',   'fecha',   'importe',   'observacion','tipo','id_orden_pago', " _
            & " 'neto_gravado_compensado', 'alicuota_percepcion', 'monto_a_percibir', 'cancelado')" _

Else
        q = " Update sp.ordenes_pago_compensatorios   SET" _
            & " id_comprobante = 'id_comprobante' , " _
            & " id_orden_pago = 'id_orden_pago', " _
            & " fecha = 'fecha' ,  " _
            & " tipo = 'tipo' ,  " _
            & " importe = 'importe' , " _
            & " observacion = 'observacion' " _
            & " neto_gravado_compensado = 'neto_gravado_compensado' " _
            & " alicuota_percepcion = 'alicuota_percepcion' " _
            & " monto_a_percibir = 'monto_a_percibir' " _
            & " cancelado = 'cancelado' " _
            & " Where  id = 'id'  "
    End If

    q = Replace$(q, "'id_comprobante'", conectar.GetEntityId(c.Comprobante))
    q = Replace$(q, "'id_orden_pago'", conectar.Escape(c.IdOrdenPago))
    q = Replace$(q, "'fecha'", conectar.Escape(c.FechaCancelacion))
    q = Replace$(q, "'importe'", conectar.Escape(c.Monto))
    q = Replace$(q, "'observacion'", conectar.Escape(c.Observacion))
    q = Replace$(q, "'tipo'", conectar.Escape(c.Tipo))
    q = Replace$(q, "'id'", conectar.GetEntityId(c))

    q = Replace$(q, "'neto_gravado_compensado'", conectar.Escape(c.NetoGravadoCompensado))
    q = Replace$(q, "'alicuota_percepcion'", conectar.Escape(c.alicuotaPercepcion))
    q = Replace$(q, "'monto_a_percibir'", conectar.Escape(c.MontoAPercibir))
    q = Replace$(q, "'cancelado'", conectar.Escape(c.Cancelado))








    If Not conectar.execute(q) Then GoTo err2

    Exit Function

err2:
    Guardar = False
End Function


Public Function FindByOP(idOP As Long) As Collection
    Set FindByOP = FindAll("id_orden_pago=" & idOP)
End Function

Public Function FindAll(Optional filtro As String) As Collection
    On Error GoTo err1
    Dim rs As Recordset
    Dim a As String
    Dim col As New Collection
    Dim index As New Dictionary
    a = "SELECT * FROM sp.ordenes_pago_compensatorios WHERE 1=1 "

    If LenB(filtro) > 0 Then a = a & " and " & filtro

    Set rs = conectar.RSFactory(a)
    conectar.BuildFieldsIndex rs, index

    While Not rs.EOF And Not rs.BOF
        col.Add Map(rs, index, "ordenes_pago_compensatorios")
        rs.MoveNext
    Wend
    Set FindAll = col

    Exit Function
err1:
    Set FindAll = Nothing
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As Compensatorio
    Dim c As Compensatorio
    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    If id > 0 Then
        Dim idf As Long
        idf = GetValue(rs, indice, tabla, "id_comprobante")
        Set c = New Compensatorio
        c.id = id
        Set c.Comprobante = DAOFacturaProveedor.FindById(idf)
        c.Tipo = GetValue(rs, indice, tabla, "tipo")
        c.FechaCancelacion = GetValue(rs, indice, tabla, "fecha")
        c.IdOrdenPago = GetValue(rs, indice, tabla, "id_orden_pago")
        c.Monto = GetValue(rs, indice, tabla, "importe")
        c.Observacion = GetValue(rs, indice, tabla, "observacion")
        c.NetoGravadoCompensado = GetValue(rs, indice, tabla, "neto_gravado_compensado")
        c.alicuotaPercepcion = GetValue(rs, indice, tabla, "alicuota_percepcion")
        c.MontoAPercibir = GetValue(rs, indice, tabla, "monto_a_percibir")
        c.Cancelado = GetValue(rs, indice, tabla, "cancelado")

        Set Map = c
    End If

End Function
