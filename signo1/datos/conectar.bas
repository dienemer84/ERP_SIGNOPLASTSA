Attribute VB_Name = "conectar"
Dim serverBBDD As String
Public port As String
Dim cn As ADODB.Connection
Dim vcount As Long
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
Public Function obternerConexion() As ADODB.Connection
    Set obternerConexion = cn
End Function
Public Property Get count() As Long
    count = vcount
End Property

Public Function RSFactory(consulta) As ADODB.Recordset
    Dim rstmp As New ADODB.Recordset
    On Error GoTo err10

    If rstmp.State = 1 Then rstmp.Close
    rstmp.Open consulta, cn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockOptimistic, adCmdText
    Set RSFactory = rstmp

    vcount = vcount + 1
    Exit Function
err10:
    Debug.Print (consulta)
    Err.Raise 2, "Motor de base de datos", "Imposible realizar la consulta solicitada" & Chr(10) & consulta

End Function
Public Function RSFactoryCliente(consulta) As ADODB.Recordset
    Dim rstmp As New ADODB.Recordset
    On Error GoTo err10
    rstmp.CursorLocation = adUseClient
    If rstmp.State = 1 Then rstmp.Close
    rstmp.Open consulta, cn, adOpenDynamic, adLockOptimistic, adCmdText
    Set RSFactoryCliente = rstmp
    vcount = vcount + 1
    Exit Function
err10:
    MsgBox "Se produjo un error: " & Err.Description
End Function
Public Function SetServidorBBDD(nServidorBBDD As String)
    serverBBDD = nServidorBBDD
End Function

Public Function GetServidorBBDD() As String
    GetServidorBBDD = serverBBDD
End Function
Public Function execute(cmdText As String) As Boolean
    On Error GoTo e12
    cn.execute cmdText
    execute = True
    Exit Function
e12:
    Err.Raise 1, "Motor de bases de datos", "Imposible ejecutar el comando " & Chr(10) & cmdText
    execute = False
End Function
Public Sub BeginTransaction()
    cn.BeginTrans
End Sub
Public Sub CommitTransaction()
    cn.CommitTrans
End Sub
Public Sub RollBackTransaction()
    cn.RollbackTrans
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


Public Sub BuildFieldsIndex(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary)
    Dim F As Field
    Dim counter As Long
    counter = 0
    'If fieldsIndex Is Nothing Then Set fieldsIndex = New Dictionary
    Set fieldsIndex = New Dictionary

    Dim prop As Property

    For Each F In rs.Fields
        'For Each prop In f.Properties
        '    Debug.Print prop.Name, prop.value, prop.Attributes
        'Next prop

        fieldsIndex.Add F.Properties(3).value & "." & F.Name, counter


        counter = counter + 1
        '''debug.print (F.Name)
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
'    ''debug.print (tableName & "." & fieldName)
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


