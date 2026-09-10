Attribute VB_Name = "DAOOrdenDeEntrega"

Public Function GetById(Id As Long) As OrdenDeEntrega
    Set GetById = GetAll("oe.id=" & Id)
End Function


Public Function GetAll(Optional filter As String = vbNullString) As Collection

    On Error GoTo errHandler

    Dim strsql As String
    Dim indice As Dictionary
    Dim rs As Recordset
    Dim col As New Collection
    Dim oe As OrdenDeEntrega

    Dim paso As String
    Dim fila As Long

    paso = "Armando SQL"

    strsql = "SELECT * FROM PedidosEntregas pe " _
           & "LEFT JOIN usuarios u1 ON pe.usuario = u1.id " _
           & "LEFT JOIN usuarios u2 ON pe.IdUsuarioAprobado = u2.id " _
           & "LEFT JOIN AdminConfigMonedas m ON pe.IdMoneda = m.id " _
           & "LEFT JOIN clientes c ON pe.IdCliente = c.id " _
           & "WHERE 1=1 "

    If Len(filter) > 0 Then
        strsql = strsql & " " & filter
    End If

    paso = "Ejecutando SQL"

    Set rs = conectar.RSFactory(strsql)

    If rs Is Nothing Then
        Err.Raise vbObjectError + 1000, _
                  "DAOOrdenDeEntrega.GetAll", _
                  "RSFactory devolvió Nothing."
    End If

    paso = "Creando índice de campos"

    conectar.BuildFieldsIndex rs, indice

    fila = 0

    While Not rs.EOF

        fila = fila + 1

        paso = "Mapeando OE - fila " & fila

        Set oe = Map(rs, indice, "pe", "c", "u1", "u2", "m")

        paso = "Agregando OE a colección - fila " & fila

        If oe Is Nothing Then
            Err.Raise vbObjectError + 1001, _
                      "DAOOrdenDeEntrega.GetAll", _
                      "Map devolvió Nothing en la fila " & fila
        End If

        col.Add oe, CStr(oe.Id)

        rs.MoveNext

    Wend

    Set GetAll = col
    Exit Function


errHandler:

    MsgBox "DAOOrdenDeEntrega.GetAll" & vbCrLf & vbCrLf & _
           "Paso: " & paso & vbCrLf & _
           "Fila: " & fila & vbCrLf & _
           "Error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error Orden de Entrega"

    Set GetAll = Nothing

End Function

Public Function Map(ByRef rs As Recordset, ByRef indice As Dictionary, ByRef tabla As String, Optional ByRef tablaCliente As String, Optional ByRef tablaUsuCreador As String, Optional ByRef TablaUsuAprobador As String, Optional ByRef tablaMoneda As String) As OrdenDeEntrega

    Dim oe As OrdenDeEntrega
    Dim Id As Variant
    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set oe = New OrdenDeEntrega
        oe.Id = Id
        oe.estado = GetValue(rs, indice, tabla, "estado")
        oe.FEcha = GetValue(rs, indice, tabla, "fecha")
        oe.fechaCreado = GetValue(rs, indice, tabla, "fechaCreado")
        oe.fechaAprobado = GetValue(rs, indice, tabla, "fechaAprobado")
        oe.referencia = GetValue(rs, indice, tabla, "referencia")


        If LenB(tablaCliente) > 0 Then Set oe.Cliente = DAOCliente.Map(rs, indice, tablaCliente)
        If LenB(tablaMoneda) > 0 Then Set oe.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaUsuCreador) > 0 Then Set oe.usuarioCreador = DAOUsuarios.Map(rs, indice, tablaUsuCreador)
        If LenB(TablaUsuAprobador) > 0 Then Set oe.usuarioAprobador = DAOUsuarios.Map(rs, indice, TablaUsuAprobador)
    End If

    Set Map = oe
End Function

