Attribute VB_Name = "DAOCuentaBancaria"
Option Explicit

Public Sub LlenarCombo(cbo As ComboBox)
    Dim col As Collection
    Set col = FindAll()
    Dim c As CuentaBancaria

    For Each c In col
        cbo.AddItem c.DescripcionFormateada
        cbo.ItemData(cbo.NewIndex) = c.id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As Collection
    Set col = FindAll()
    Dim c As CuentaBancaria

    For Each c In col
        cbo.AddItem c.DescripcionFormateada
        cbo.ItemData(cbo.NewIndex) = c.id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Sub llenarComboCBUXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As Collection
    Set col = FindAllWithCBU()
    Dim c As CuentaBancaria

    For Each c In col
        cbo.AddItem c.DescripcionCBUFormateada
        cbo.ItemData(cbo.NewIndex) = c.id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Function FindAllWithCBU() As Collection
    Dim col As Collection
    Set col = FindAll("cbu IS NOT null")
    
    Set FindAllWithCBU = col
End Function

Public Function FindById(id As Long) As CuentaBancaria
    Dim col As Collection
    Set col = FindAll("c.id = " & id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function


Public Function FindByCBU(cbu As String) As CuentaBancaria
    Dim col As Collection
    Set col = FindAll("c.cbu = " & Escape(cbu))
    If col.count = 0 Then
        Set FindByCBU = Nothing
    Else
        Set FindByCBU = col.item(1)
    End If
End Function



Public Function FindAll(Optional ByVal filter As String = " 1 = 1 ") As Collection
    Dim q As String
    q = "SELECT *" _
        & " FROM AdminConfigCuentas c" _
        & " LEFT JOIN AdminConfigBancos b ON b.id = c.idBanco" _
        & " LEFT JOIN AdminConfigMonedas m ON m.id = c.moneda_id WHERE " & filter

    Dim col As New Collection
    Dim rs As Recordset
    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim tmp As CuentaBancaria

    While Not rs.EOF
        Set tmp = Map(rs, fieldsIndex, "c", "b", "m")
        col.Add tmp, CStr(tmp.id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function



Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaBanco As String = vbNullString, Optional tablaMoneda As String = vbNullString) As CuentaBancaria
    Dim c As CuentaBancaria
    Dim id As Long

    id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set c = New CuentaBancaria
        c.id = id
        c.numero = GetValue(rs, indice, tabla, "cuenta")
        c.TipoCuenta = GetValue(rs, indice, tabla, "tipo")
             c.cbu = GetValue(rs, indice, tabla, "cbu")
        If LenB(tablaBanco) > 0 Then Set c.Banco = DAOBancos.Map(rs, indice, tablaBanco)
        If LenB(tablaMoneda) > 0 Then Set c.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
    End If

    Set Map = c
End Function

Public Function Save(cuenta As CuentaBancaria) As Boolean
    Dim q As String
    If cuenta.id = 0 Then
        q = "INSERT INTO AdminConfigCuentas (idBanco, cuenta, tipo, moneda_id,CBU)" _
            & " VALUES (" & GetEntityId(cuenta.Banco) & ", " & Escape(cuenta.numero) & "," & cuenta.TipoCuenta & "," & GetEntityId(cuenta.moneda) & "," & Escape(cuenta.cbu) & " )"
    Else
        q = "Update AdminConfigCuentas" _
            & " SET" _
            & " idBanco = " & GetEntityId(cuenta.Banco) & " ," _
            & " cuenta = " & Escape(cuenta.numero) & " ," _
            & " tipo = " & cuenta.TipoCuenta & " ," _
            & " moneda_id = " & GetEntityId(cuenta.moneda) _
            & ", cbu = " & Escape(cuenta.cbu) _
            & " WHERE id = " & cuenta.id
    End If

    Save = conectar.execute(q)
    Dim id As Long
    If Save And cuenta.id = 0 Then
        If conectar.UltimoId("AdminConfigCuentas", id) Then
            cuenta.id = id
        End If
    End If
End Function
