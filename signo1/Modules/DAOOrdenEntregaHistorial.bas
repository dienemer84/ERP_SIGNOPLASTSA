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
        "fecha, " & _
        "mensaje, " & _
        "usuario " & _
        "FROM historico_PedidoEntrega " & _
        "WHERE id_source = " & CStr(idOE) & " " & _
        "ORDER BY fecha DESC"


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
        '----------------------------------------------
        If IsNull(rsHistorial!mensaje) Then
            h.mensaje = vbNullString
        Else
            h.mensaje = CStr(rsHistorial!mensaje)
        End If


        '----------------------------------------------
        ' USUARIO
        ' En esta tabla está guardado como TEXTO
        '----------------------------------------------
        If Not IsNull(rsHistorial!Usuario) Then

            Set u = New clsUsuario

            u.Usuario = CStr(rsHistorial!Usuario)

            Set h.Usuario = u

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

