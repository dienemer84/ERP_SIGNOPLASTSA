Attribute VB_Name = "DAOCompensatorios"
Option Explicit


Public Function Save(C As Compensatorio) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    If Not Guardar(C) Then GoTo err1

    conectar.CommitTransaction
    Exit Function

err1:
    conectar.RollBackTransaction
End Function



Public Function Guardar(C As Compensatorio) As Boolean
    Dim q As String
    On Error GoTo err2
    Guardar = True
    If C.id = 0 Then
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
            & " observacion = 'observacion', " _
            & " neto_gravado_compensado = 'neto_gravado_compensado' ," _
            & " alicuota_percepcion = 'alicuota_percepcion', " _
            & " monto_a_percibir = 'monto_a_percibir', " _
            & " cancelado = 'cancelado' " _
            & " Where  id = 'id'  "
    End If

    q = Replace$(q, "'id_comprobante'", conectar.GetEntityId(C.Comprobante))
    q = Replace$(q, "'id_orden_pago'", conectar.Escape(C.IdOrdenPago))
    q = Replace$(q, "'fecha'", conectar.Escape(C.FechaCancelacion))
    q = Replace$(q, "'importe'", conectar.Escape(C.Monto))
    q = Replace$(q, "'observacion'", conectar.Escape(C.Observacion))
    q = Replace$(q, "'tipo'", conectar.Escape(C.Tipo))
    q = Replace$(q, "'id'", conectar.GetEntityId(C))

    q = Replace$(q, "'neto_gravado_compensado'", conectar.Escape(C.NetoGravadoCompensado))
    q = Replace$(q, "'alicuota_percepcion'", conectar.Escape(C.alicuotaPercepcion))
    q = Replace$(q, "'monto_a_percibir'", conectar.Escape(C.MontoAPercibir))
    q = Replace$(q, "'cancelado'", conectar.Escape(C.Cancelado))








    If Not conectar.execute(q) Then GoTo err2

    Exit Function

err2:
    Guardar = False
End Function


Public Function FindByOP(idOP As Long) As Collection
    Set FindByOP = FindAll("id_orden_pago=" & idOP)
End Function

Public Function FindAllPendientesByProveedor(idp As Integer) As Collection
    On Error GoTo err1
    Dim rs As Recordset
    Dim A As String
    Dim col As New Collection
    Dim index As New Dictionary
    A = "SELECT opc.* FROM ordenes_pago_compensatorios opc  JOIN ordenes_pago op ON opc.id_orden_pago=op.id JOIN AdminComprasFacturasProveedores acfp ON opc.id_comprobante=acfp.id WHERE acfp.id_proveedor = " & idp & " AND op.estado=1 AND opc.cancelado=0 "
    
    Set rs = conectar.RSFactory(A)
    conectar.BuildFieldsIndex rs, index
    Dim C As Compensatorio
    
    While Not rs.EOF And Not rs.BOF
       Set C = Map(rs, index, "opc")
        col.Add C, CStr(C.id)
        
        rs.MoveNext
    Wend
    Set FindAllPendientesByProveedor = col

    Exit Function
err1:
    Set FindAllPendientesByProveedor = Nothing
End Function



Public Function FindAll(Optional filtro As String) As Collection
    On Error GoTo err1
    Dim rs As Recordset
    Dim A As String
    Dim col As New Collection
    Dim index As New Dictionary
    A = "SELECT * FROM sp.ordenes_pago_compensatorios WHERE 1=1 "

    If LenB(filtro) > 0 Then A = A & " and " & filtro

    Set rs = conectar.RSFactory(A)
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
    Dim C As Compensatorio
    Dim id As Long: id = GetValue(rs, indice, tabla, "id")
    If id > 0 Then
        Dim idf As Long
        idf = GetValue(rs, indice, tabla, "id_comprobante")
        Set C = New Compensatorio
        C.id = id
        Set C.Comprobante = DAOFacturaProveedor.FindById(idf)
        C.Tipo = GetValue(rs, indice, tabla, "tipo")
        C.FechaCancelacion = GetValue(rs, indice, tabla, "fecha")
        C.IdOrdenPago = GetValue(rs, indice, tabla, "id_orden_pago")
        C.Monto = GetValue(rs, indice, tabla, "importe")
        C.Observacion = GetValue(rs, indice, tabla, "observacion")
        C.NetoGravadoCompensado = GetValue(rs, indice, tabla, "neto_gravado_compensado")
        C.alicuotaPercepcion = GetValue(rs, indice, tabla, "alicuota_percepcion")
        C.MontoAPercibir = GetValue(rs, indice, tabla, "monto_a_percibir")
        C.Cancelado = GetValue(rs, indice, tabla, "cancelado")

        Set Map = C
    End If

End Function
