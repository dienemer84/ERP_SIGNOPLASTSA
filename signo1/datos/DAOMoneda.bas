Attribute VB_Name = "DAOMoneda"
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Public Const MONEDA_PESO_ID As Long = 0

Private last_query As Date

Private last_moneda As clsMoneda
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE_CORTO As String = "nombre_corto"
Public Const CAMPO_NOMBRE_LARGO As String = "nombre_largo"
Public Const CAMPO_CAMBIO As String = "cambio"
Public Const CAMPO_PATRON As String = "patron"
Public Const CAMPO_FECHA_ACTUAL As String = "FechaActual"

Public Function GetAll(Optional filtro As String = vbNullString) As Collection

    On Error GoTo err1
    Dim col As New Collection
    Dim idx As Dictionary
        
    Dim moneda As clsMoneda
    Dim q As String

    Dim withPatron As Boolean
    withPatron = True    'TODO: filtrar esto para optimizar mas tema de ordenes de pago

    If Not withPatron Then
        q = "select  * from AdminConfigMonedas mon where 1=1"

    Else

        q = "select * from AdminConfigMonedas mon LEFT JOIN AdminConfigMonedas mon2 on mon2.id = mon.idMonedaCambio where 1=1"
    End If
    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    Set rs = conectar.RSFactory(q)

    BuildFieldsIndex rs, idx

    While Not rs.EOF
        If Not withPatron Then
            Set moneda = Map(rs, idx, "mon")
        Else

            Set moneda = Map(rs, idx, "mon", "mon2")
        End If
        col.Add moneda, CStr(moneda.Id)
        rs.MoveNext
    Wend

    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function

Public Function GetById(Id As Long) As clsMoneda
    Set GetById = GetAll("mon.id=" & Id)(1)
    Exit Function
err1:
    Set GetById = Nothing
    
End Function

Public Sub LlenarCombo(cbo As ComboBox)
    Dim col As Collection
    Set col = DAOMoneda.GetAll()
    Dim moneda As clsMoneda
    cbo.Clear
    For i = 1 To col.count
        Set moneda = col(i)
        cbo.AddItem moneda.NombreCorto
        cbo.ItemData(cbo.NewIndex) = moneda.Id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
    
End Sub


Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox, Optional withValue As Boolean = False)
    Dim col As Collection
    Set col = DAOMoneda.GetAll
    Dim moneda As clsMoneda
    cbo.Clear
    For i = 1 To col.count
        Set moneda = col(i)
        If withValue Then
            cbo.AddItem moneda.NombreCorto & " | " & moneda.Cambio

        Else
            cbo.AddItem moneda.NombreCorto
        End If
        cbo.ItemData(cbo.NewIndex) = moneda.Id
    Next i
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub



Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String, Optional tablaMon2 As String = vbNullString) As clsMoneda
    Dim tmpMoneda As clsMoneda
    Dim Id As Variant
    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_ID)

    If Id >= 0 Then
        Set tmpMoneda = New clsMoneda
        tmpMoneda.Id = Id
        tmpMoneda.NombreCorto = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NOMBRE_CORTO)
        tmpMoneda.NombreLargo = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_NOMBRE_LARGO)
        tmpMoneda.Cambio = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_CAMBIO)
        tmpMoneda.Patron = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_PATRON)
        tmpMoneda.FechaActual = GetValue(rs, fieldsIndex, tableNameOrAlias, CAMPO_FECHA_ACTUAL)

        If LenB(tablaMon2) > 0 Then
            Set tmpMoneda.MonedaCambio = Map(rs, fieldsIndex, tablaMon2)
        End If
    End If

    Set Map = tmpMoneda
End Function

Public Function FindFirstByPatronOrDefault() As clsMoneda
    Dim comparar As Date
    comparar = DateAdd("n", 10, last_query)
    If comparar < Now Then
        last_query = Now

        Set last_moneda = GetAll("mon.patron=1")(1)

    End If
    Set FindFirstByPatronOrDefault = last_moneda



End Function
