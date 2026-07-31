Attribute VB_Name = "conectar"
Dim serverBBDD As String
Public port As String
Dim cn As ADODB.Connection
Dim vcount As Long
Private mTransaccionActiva As Boolean

Public Const ERR_CONEXION_BD As Long = vbObjectError + 2100
Public Const ERR_TRANSACCION_BD As Long = vbObjectError + 2101
Public Const ERR_RESULTADO_INCIERTO_BD As Long = vbObjectError + 2102


Public Function conectar() As Boolean
    conectar = True
    On Error GoTo err22
    Set cn = New ADODB.Connection
    'http://dev.mysql.com/doc/refman/5.0/en/connector-odbc-configuration-connection-parameters.html
    'http://dev.mysql.com/tech-resources/articles/vb-blob-handling.html
    cn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};port=" & port & " ;server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;Option=" & (1 + 1024) & "';AllowZeroDateTime=true; ConvertZeroDateTime=True'"  ';connection=adUseClient" ' era 3 ' eraç 1 + 1024
    ' cn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;Option=" & (1 + 1024)   ';connection=adUseClient" ' era 3
    'cn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=" & serverBBDD & ";Database=sp;User=root; Password=3l3c720n;Option=" & (1 Or 2 Or 1024)
    'cn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=" & serverBBDD & ";Database=sp;User=root; Password=3l3c720n;Option=" & (1 + 2 + 64 + 1024) & ";"



    'cn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;ConvertZeroDateTime=True"


    cn.Open
    vcount = 0
    Exit Function
err22:
    MsgBox "Se produjo un error: " & Err.Description
    conectar = False
    Err.Clear
End Function

Private Sub DescartarConexion()
    On Error Resume Next

    If Not cn Is Nothing Then
        If cn.State <> adStateClosed Then cn.Close
    End If

    Set cn = Nothing
End Sub


Private Function AbrirConexionNueva( _
                    Optional ByVal MostrarMensaje As Boolean = False) As Boolean

    Dim detalleError As String

    On Error GoTo ErrorAbrir

    DescartarConexion

    Set cn = New ADODB.Connection

    cn.ConnectionTimeout = 5
    cn.CommandTimeout = 30
    cn.CursorLocation = adUseClient

    cn.ConnectionString = "driver={MySQL ODBC 3.51 Driver};port=" & port & " ;server=" & serverBBDD & ";uid=root;pwd=3l3c720n;database=sp;Option=" & (1 + 1024) & "';AllowZeroDateTime=true; ConvertZeroDateTime=True'"  ';connection=adUseClient" ' era 3 ' eraç 1 + 1024

    cn.Open

    vcount = 0
    mTransaccionActiva = False
    AbrirConexionNueva = True
    Exit Function

ErrorAbrir:
    detalleError = Err.Description

    mTransaccionActiva = False
    DescartarConexion

    If MostrarMensaje Then
        MsgBox "No se pudo conectar con el servidor." & vbCrLf & _
               detalleError, _
               vbCritical, "Conexión"
    End If

    AbrirConexionNueva = False
End Function



Public Function ConexionActiva() As Boolean
    Dim rsPing As ADODB.Recordset

    On Error GoTo ErrorConexion

    If cn Is Nothing Then GoTo salir
    If cn.State <> adStateOpen Then GoTo salir

    'No alcanza con comprobar State: se ejecuta una consulta real.
    Set rsPing = cn.execute("SELECT 1")

    ConexionActiva = True

salir:
    On Error Resume Next

    If Not rsPing Is Nothing Then
        If rsPing.State = adStateOpen Then rsPing.Close
    End If

    Set rsPing = Nothing
    Exit Function

ErrorConexion:
    ConexionActiva = False
    Resume salir
End Function


Public Function IntentarReconectar() As Boolean
    'Nunca se reconstruye una conexión en medio de una transacción.
    If mTransaccionActiva Then Exit Function

    IntentarReconectar = AbrirConexionNueva(False)
End Function


Public Function AsegurarConexion() As Boolean
    If ConexionActiva() Then
        AsegurarConexion = True
        Exit Function
    End If

    If mTransaccionActiva Then
        'La transacción que utilizaba la conexión anterior ya no es válida.
        mTransaccionActiva = False
        DescartarConexion
        Exit Function
    End If

    AsegurarConexion = AbrirConexionNueva(False)
End Function


Public Function obternerConexion() As ADODB.Connection

    If Not AsegurarConexion() Then
        Err.Raise ERR_CONEXION_BD, _
                  "Motor de base de datos", _
                  "No fue posible restablecer la conexión con el servidor."
    End If

    Set obternerConexion = cn
End Function


Public Property Get count() As Long
    count = vcount
End Property


Private Function AbrirRecordsetConReconexion( _
                    ByVal consulta As String, _
                    ByVal usarCursorCliente As Boolean) As ADODB.Recordset

    Dim rstmp As ADODB.Recordset
    Dim reintentado As Boolean
    Dim estabaEnTransaccion As Boolean

    Dim numeroError As Long
    Dim origenError As String
    Dim descripcionError As String

    estabaEnTransaccion = mTransaccionActiva

Reintentar:
    On Error GoTo ErrorConsulta

    If Not AsegurarConexion() Then

        If estabaEnTransaccion Then
            Err.Raise ERR_TRANSACCION_BD, _
                      "Motor de base de datos", _
                      "La conexión se perdió durante una transacción."
        Else
            Err.Raise ERR_CONEXION_BD, _
                      "Motor de base de datos", _
                      "No fue posible conectarse con el servidor."
        End If

    End If

    Set rstmp = New ADODB.Recordset

    If usarCursorCliente Then
        rstmp.CursorLocation = adUseClient
        rstmp.Open consulta, cn, _
                   adOpenDynamic, adLockOptimistic, adCmdText
    Else
        rstmp.Open consulta, cn, _
                   adOpenStatic, adLockOptimistic, adCmdText
    End If

    Set AbrirRecordsetConReconexion = rstmp

    vcount = vcount + 1
    Exit Function

ErrorConsulta:
    numeroError = Err.Number
    origenError = Err.Source
    descripcionError = Err.Description

    Debug.Print consulta

    On Error Resume Next

    If Not rstmp Is Nothing Then
        If rstmp.State = adStateOpen Then rstmp.Close
    End If

    Set rstmp = Nothing
    On Error GoTo 0

    'Si la conexión realmente murió, se intenta una sola vez.
    If Not ConexionActiva() Then

        If estabaEnTransaccion Then
            mTransaccionActiva = False
            DescartarConexion

            Err.Raise ERR_TRANSACCION_BD, _
                      "Motor de base de datos", _
                      "Se perdió la conexión durante una transacción." & _
                      vbCrLf & _
                      "La operación no fue confirmada."
        End If

        If Not reintentado Then
            reintentado = True

            If AbrirConexionNueva(False) Then
                GoTo Reintentar
            End If
        End If

        Err.Raise ERR_CONEXION_BD, _
                  "Motor de base de datos", _
                  "No hay conexión con el servidor." & vbCrLf & _
                  "El sistema permanecerá abierto."
    End If

    'Era un error SQL, no un problema de conexión.
    Err.Raise numeroError, origenError, descripcionError
End Function


Public Function RSFactory(ByVal consulta As String) As ADODB.Recordset

    Set RSFactory = _
        AbrirRecordsetConReconexion(consulta, False)

End Function


Public Function RSFactoryCliente(ByVal consulta As String) As ADODB.Recordset

    Set RSFactoryCliente = _
        AbrirRecordsetConReconexion(consulta, True)

End Function


Public Function SetServidorBBDD(nServidorBBDD As String)
    serverBBDD = nServidorBBDD
End Function

Public Function GetServidorBBDD() As String
    GetServidorBBDD = serverBBDD
End Function


Public Function execute(ByVal cmdText As String) As Boolean

    Dim comandoIniciado As Boolean
    Dim estabaEnTransaccion As Boolean

    Dim numeroError As Long
    Dim origenError As String
    Dim descripcionError As String

    estabaEnTransaccion = mTransaccionActiva

    On Error GoTo ErrorExecute

    If Not AsegurarConexion() Then

        If estabaEnTransaccion Then
            Err.Raise ERR_TRANSACCION_BD, _
                      "Motor de base de datos", _
                      "La conexión se perdió durante la transacción."
        Else
            Err.Raise ERR_CONEXION_BD, _
                      "Motor de base de datos", _
                      "No fue posible conectarse con el servidor."
        End If

    End If

    comandoIniciado = True
    cn.execute cmdText

    execute = True
    Exit Function

ErrorExecute:
    numeroError = Err.Number
    origenError = Err.Source
    descripcionError = Err.Description

    execute = False

    If comandoIniciado And Not ConexionActiva() Then

        mTransaccionActiva = False
        DescartarConexion

        'Se reconecta solamente para las operaciones posteriores.
        'No se vuelve a ejecutar el comando.
        If Not estabaEnTransaccion Then
            AbrirConexionNueva False
        End If

        If estabaEnTransaccion Then
            Err.Raise ERR_TRANSACCION_BD, _
                      "Motor de base de datos", _
                      "Se perdió la conexión durante una transacción." & _
                      vbCrLf & _
                      "La operación completa debe volver a realizarse."
        Else
            Err.Raise ERR_RESULTADO_INCIERTO_BD, _
                      "Motor de base de datos", _
                      "La conexión se perdió durante el guardado." & _
                      vbCrLf & _
                      "El comando no fue repetido para evitar duplicados." & _
                      vbCrLf & _
                      "Verifique el registro antes de volver a guardar."
        End If

    End If

    Err.Raise numeroError, origenError, descripcionError
End Function


Public Function ExecuteRa(ByVal sql As String, ByRef ra As Long) As Boolean
  On Error GoTo err1
  cn.execute sql, ra, adExecuteNoRecords
  ExecuteRa = True
  Exit Function
err1:
  ra = -1
  ExecuteRa = False
End Function

Public Sub BeginTransaction()

    If Not AsegurarConexion() Then
        Err.Raise ERR_CONEXION_BD, _
                  "Motor de base de datos", _
                  "No se puede iniciar la transacción porque no hay conexión."
    End If

    cn.BeginTrans
    mTransaccionActiva = True
End Sub


Public Sub CommitTransaction()
    Dim detalleError As String

    On Error GoTo ErrorCommit

    cn.CommitTrans
    mTransaccionActiva = False
    Exit Sub

ErrorCommit:
    detalleError = Err.Description

    mTransaccionActiva = False
    DescartarConexion

    'Reconexión para las próximas operaciones.
    AbrirConexionNueva False

    Err.Raise ERR_RESULTADO_INCIERTO_BD, _
              "Motor de base de datos", _
              "No fue posible confirmar la transacción." & vbCrLf & _
              "Verifique si la operación fue registrada." & vbCrLf & _
              detalleError
End Sub


Public Sub RollBackTransaction()
    On Error Resume Next

    If mTransaccionActiva Then
        If Not cn Is Nothing Then
            If cn.State = adStateOpen Then cn.RollbackTrans
        End If
    End If

    mTransaccionActiva = False
End Sub


Public Function UltimoId(tableName As String, ByRef Id As Long) As Boolean
    On Error GoTo er1
    UltimoId = True
    Dim rs As Recordset
    Set rs = RSFactory("select last_insert_id() as ultimo from " & tableName)
    If Not rs.EOF And Not rs.BOF Then
        Id = rs!ultimo
    Else
        GoTo er1
    End If
    Exit Function
er1:
    UltimoId = False
    Id = 0
End Function


Public Function UltimoId2() As Long
    On Error GoTo er1
    UltimoId2 = 0
    Dim rs As Recordset
    Set rs = RSFactory("select last_insert_id() as ultimo")
    If Not rs.EOF And Not rs.BOF Then UltimoId2 = rs!ultimo
    Exit Function
er1:
    UltimoId2 = 0
End Function


Public Sub BuildFieldsIndex( _
                ByRef rs As ADODB.Recordset, _
                ByRef fieldsIndex As Dictionary)

    Dim F As ADODB.Field
    Dim counter As Long
    Dim nombreTabla As String
    Dim clave As String

    Set fieldsIndex = New Dictionary
    counter = 0

    For Each F In rs.Fields

        nombreTabla = vbNullString

        If Not IsNull(F.Properties(3).value) Then
            nombreTabla = CStr(F.Properties(3).value)
        End If

        clave = nombreTabla & "." & CStr(F.Name)

        If Not fieldsIndex.Exists(clave) Then

            fieldsIndex.Add clave, counter

        Else

            'El SELECT con varios JOIN devolvió esta clave repetida.
            'Se conserva la primera aparición.
            Debug.Print "BuildFieldsIndex - clave repetida: " & _
                        clave & _
                        " - campo ignorado número: " & _
                        CStr(counter)

        End If

        counter = counter + 1

    Next F

End Sub


Public Function ProximoId(tableName As String) As Long
    On Error GoTo er1
    ProximoId = True
    Dim rs As Recordset
    Set rs = RSFactory("select MAX(id)+1 as proximo from " & tableName)
    If Not rs.EOF And Not rs.BOF Then
        ProximoId = rs!proximo
    Else
        GoTo er1
    End If
    Exit Function
er1:
    ProximoId = False
    Id = 0
End Function


Public Function Escape(value As Variant) As Variant
    Dim retorno As Variant
    retorno = value

    Select Case VarType(value)
    Case VbVarType.vbString
        If LenB(Trim$(value)) = 0 Then
            retorno = "NULL"
        Else
            retorno = value
            retorno = Replace$(retorno, "'", "\'")

            retorno = "'" & retorno & "'"
        End If
    Case VbVarType.vbDate
        If CDbl(value) = 0 Then
            retorno = "NULL"
        Else
            If (CDbl(value) - Int(CDbl(value))) = 0 Then    'no tiene decimales es solo fecha sin hora
                retorno = Format(value, "yyyy-MM-dd")
            Else
                retorno = Format(value, "yyyy-MM-dd HH:mm:ss")
            End If

            retorno = "'" & retorno & "'"
        End If
    Case VbVarType.vbBoolean
        retorno = CInt(value) * -1
    Case VbVarType.vbDouble
        retorno = "'" & Trim$(str$(value)) & "'"
    End Select
    'If retorno = vbEmpty Then retorno = " "
    Escape = retorno

End Function


Public Function GetValue(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableName As String, ByRef fieldName As String)
    'Debug.Print (tableName & "." & fieldName)
    GetValue = rs.Fields.item(fieldsIndex(tableName & "." & fieldName)).value
    If IsNull(GetValue) Then GetValue = Empty
End Function


Public Function GetEntityId(entity As Object) As Variant
    If entity Is Nothing Then
        GetEntityId = 0    '"NULL"
    Else
        GetEntityId = entity.Id
    End If
End Function


Public Sub MostrarErrorBaseDatos( _
                ByVal numero As Long, _
                ByVal descripcion As String)

    Select Case numero

    Case ERR_CONEXION_BD
        MsgBox "No hay conexión con el servidor." & vbCrLf & _
               "El sistema permanecerá abierto." & vbCrLf & _
               "Verifique la red y vuelva a intentar.", _
               vbCritical, "Servidor no disponible"

    Case ERR_TRANSACCION_BD
        MsgBox "La conexión se perdió durante la operación." & vbCrLf & _
               "La operación no fue confirmada." & vbCrLf & _
               "Los datos permanecerán en pantalla.", _
               vbExclamation, "Operación interrumpida"

    Case ERR_RESULTADO_INCIERTO_BD
        MsgBox descripcion, _
               vbExclamation, "Verificar operación"

    Case Else
        MsgBox "Se produjo un error:" & vbCrLf & _
               descripcion, _
               vbExclamation, "Signo Plast ERP"

    End Select
End Sub

