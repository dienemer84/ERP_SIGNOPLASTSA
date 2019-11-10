Attribute VB_Name = "DAOOrdenDeEntrega"

Public Function GetById(id As Long) As OrdenDeEntrega
Set GetById = GetAll("oe.id=" & id)
End Function

Public Function GetAll(Optional filter As String = vbNullString) As Collection
    
    Dim strsql As String
    Dim indice As Dictionary
    Dim rs As Recordset
    Dim col As New Collection
    strsql = "SELECT * FROM  PedidosEntregas pe" _
    & " LEFT JOIN usuarios u1 on pe.usuario=u1.id " _
    & " LEFT JOIN usuarios u2 on pe.IdUsuarioAprobado=u2.id " _
    & " LEFT JOIN AdminConfigMonedas m on pe.IdMoneda=m.id " _
    & " LEFT JOIN clientes c on pe.IdCliente=c.id " _
    & "WHERE 1=1 "
    
     
    If Len(filter) > 0 Then strsql = strsql & " " & filter
    

    Set rs = conectar.RSFactory(strsql)

    conectar.BuildFieldsIndex rs, indice
    Dim oe As OrdenDeEntrega

    While Not rs.EOF
        Set oe = Map(rs, indice, "pe", "c", "u1", "u2", "m")
        col.Add oe, CStr(oe.id)
        rs.MoveNext
    Wend

    Set GetAll = col
End Function

Public Function Map(ByRef rs As Recordset, ByRef indice As Dictionary, ByRef tabla As String, Optional ByRef tablaCliente As String, Optional ByRef tablaUsuCreador As String, Optional ByRef TablaUsuAprobador As String, Optional ByRef tablaMoneda As String) As OrdenDeEntrega

    Dim oe As OrdenDeEntrega
    Dim id As Variant
    id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set oe = New OrdenDeEntrega
        oe.id = id
        oe.estado = GetValue(rs, indice, tabla, "estado")
        oe.FEcha = GetValue(rs, indice, tabla, "fecha")
        oe.fechaCreado = GetValue(rs, indice, tabla, "fechaCreado")
        oe.fechaAprobado = GetValue(rs, indice, tabla, "fechaAprobado")
        oe.referencia = GetValue(rs, indice, tabla, "referencia")
        
        
        If LenB(tablaCliente) > 0 Then Set oe.cliente = DAOCliente.Map(rs, indice, tablaCliente)
        If LenB(tablaMoneda) > 0 Then Set oe.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaUsuCreador) > 0 Then Set oe.usuarioCreador = DAOUsuarios.Map(rs, indice, tablaUsuCreador)
        If LenB(TablaUsuAprobador) > 0 Then Set oe.usuarioAprobador = DAOUsuarios.Map(rs, indice, TablaUsuAprobador)
    End If

    Set Map = oe
End Function

