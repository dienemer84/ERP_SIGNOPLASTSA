Attribute VB_Name = "DAOOrdenEntregaHistorial"
Option Explicit

Dim rs As ADODB.Recordset


Public Function GetAllByOE(ByVal idOE As Long) As Collection

    On Error GoTo errHandler

    Dim col As New Collection
    Dim h As clsHistorial
    Dim rsHistorial As ADODB.Recordset
    Dim idUsuario As Long
    Dim sql As String


    sql = _
        "SELECT fecha, nota, usuario " & _
        "FROM historico_PedidoEntrega " & _
        "WHERE idPedidoEntrega = " & CStr(idOE) & " " & _
        "ORDER BY fecha DESC"


    Set rsHistorial = conectar.RSFactory(sql)


    Do While Not rsHistorial.EOF

        Set h = New clsHistorial


        '---------------------------------------------
        ' FECHA
        '---------------------------------------------
        If Not IsNull(rsHistorial!FEcha) Then
            h.FEcha = CDate(rsHistorial!FEcha)
        End If


        '---------------------------------------------
        ' MENSAJE
        '---------------------------------------------
        If IsNull(rsHistorial!Nota) Then
            h.mensaje = vbNullString
        Else
            h.mensaje = CStr(rsHistorial!Nota)
        End If


        '---------------------------------------------
        ' USUARIO
        '---------------------------------------------
        idUsuario = 0

        If Not IsNull(rsHistorial!Usuario) Then
            idUsuario = CLng(rsHistorial!Usuario)
        End If


        If idUsuario > 0 Then
            Set h.Usuario = DAOUsuarios.GetById(idUsuario)
        End If


        col.Add h

        rsHistorial.MoveNext

    Loop


    Set GetAllByOE = col

    Exit Function


errHandler:

    Set GetAllByOE = Nothing

    MsgBox "Error al obtener el historial de la Orden de Entrega." & _
           vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Historial O/E"

End Function

