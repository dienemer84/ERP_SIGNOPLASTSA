Attribute VB_Name = "DAOOrdenEntregaHistorial"
Option Explicit

Dim rs As ADODB.Recordset


Public Function GetAllByOE(ByVal idOE As Long) As Collection

    On Error GoTo errHandler

    Dim col As New Collection
    Dim cn As ADODB.Connection
    Dim rsHistorial As ADODB.Recordset

    Dim h As clsHistorial
    Dim u As clsUsuario

    Dim sql As String
    Dim paso As String


    Set GetAllByOE = Nothing


    '--------------------------------------------------
    ' CONEXION
    '--------------------------------------------------
    paso = "Obteniendo conexión"

    Set cn = conectar.obternerConexion


    '--------------------------------------------------
    ' CONSULTA
    '--------------------------------------------------
    sql = _
        "SELECT " & _
        "h.fecha, " & _
        "h.nota, " & _
        "h.usuario AS idUsuario, " & _
        "u.usuario AS nombreUsuario " & _
        "FROM historico_PedidoEntrega h " & _
        "LEFT JOIN usuarios u ON u.id = h.usuario " & _
        "WHERE h.idPedidoEntrega = " & CStr(idOE) & " " & _
        "ORDER BY h.fecha DESC"


    paso = "Ejecutando consulta de historial"

    Set rsHistorial = cn.execute(sql)


    '--------------------------------------------------
    ' ARMAR COLECCION
    '--------------------------------------------------
    paso = "Leyendo registros"


    Do While Not rsHistorial.EOF

        Set h = New clsHistorial


        '----------------------------------------------
        ' FECHA
        '----------------------------------------------
        If Not IsNull(rsHistorial!FEcha) Then
            h.FEcha = CDate(rsHistorial!FEcha)
        End If


        '----------------------------------------------
        ' MENSAJE
        ' La columna real se llama NOTA
        '----------------------------------------------
        If IsNull(rsHistorial!Nota) Then
            h.mensaje = vbNullString
        Else
            h.mensaje = CStr(rsHistorial!Nota)
        End If


        '----------------------------------------------
        ' USUARIO
        ' historico_PedidoEntrega.usuario es BIGINT
        '----------------------------------------------
        If Not IsNull(rsHistorial!idUsuario) Then

            Set u = New clsUsuario

            u.Id = CLng(rsHistorial!idUsuario)

            If IsNull(rsHistorial!nombreUsuario) Then
                u.Usuario = "Usuario ID " & CStr(u.Id)
            Else
                u.Usuario = CStr(rsHistorial!nombreUsuario)
            End If

            h.Usuario = u

        End If


        col.Add h

        rsHistorial.MoveNext

    Loop


    Set GetAllByOE = col

    Exit Function


errHandler:

    Set GetAllByOE = Nothing

    MsgBox "Error al obtener el historial de la Orden de Entrega." & _
           vbCrLf & vbCrLf & _
           "Paso: " & paso & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "SQL:" & vbCrLf & sql, _
           vbCritical, _
           "Historial O/E"

End Function

