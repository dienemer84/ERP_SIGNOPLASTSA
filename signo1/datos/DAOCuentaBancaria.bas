Attribute VB_Name = "DAOCuentaBancaria"
Option Explicit

Public Sub LlenarCombo(cbo As ComboBox)
    Dim col As Collection
    Set col = FindAll()
    Dim c As CuentaBancaria

    For Each c In col
        cbo.AddItem c.DescripcionFormateada
        cbo.ItemData(cbo.NewIndex) = c.Id
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
        cbo.ItemData(cbo.NewIndex) = c.Id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Sub llenarComboCBU(cbo As ComboBox)
    Dim col As Collection
    Set col = FindAllWithCBU()
    Dim c As CuentaBancaria

    For Each c In col
        cbo.AddItem c.DescripcionCBUFormateada
        cbo.ItemData(cbo.NewIndex) = c.Id
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

Public Function FindById(Id As Long) As CuentaBancaria
    Dim col As Collection
    Set col = FindAll("c.id = " & Id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function


Public Function FindByCBU(CBU As String) As CuentaBancaria
    Dim col As Collection
    Set col = FindAll("c.cbu = " & Escape(CBU))
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
        col.Add tmp, CStr(tmp.Id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function



Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaBanco As String = vbNullString, Optional tablaMoneda As String = vbNullString) As CuentaBancaria
    Dim c As CuentaBancaria
    Dim Id As Long

    Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set c = New CuentaBancaria
        c.Id = Id
        c.numero = GetValue(rs, indice, tabla, "cuenta")
        c.TipoCuenta = GetValue(rs, indice, tabla, "tipo")
        c.CBU = GetValue(rs, indice, tabla, "cbu")
        If LenB(tablaBanco) > 0 Then Set c.Banco = DAOBancos.Map(rs, indice, tablaBanco)
        If LenB(tablaMoneda) > 0 Then Set c.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
    End If

    Set Map = c
End Function

Public Function Save(cuenta As CuentaBancaria) As Boolean
    Dim q As String
    If cuenta.Id = 0 Then
        q = "INSERT INTO AdminConfigCuentas (idBanco, cuenta, tipo, moneda_id,CBU)" _
          & " VALUES (" & GetEntityId(cuenta.Banco) & ", " & Escape(cuenta.numero) & "," & cuenta.TipoCuenta & "," & GetEntityId(cuenta.moneda) & "," & Escape(cuenta.CBU) & " )"
    Else
        q = "Update AdminConfigCuentas" _
          & " SET" _
          & " idBanco = " & GetEntityId(cuenta.Banco) & " ," _
          & " cuenta = " & Escape(cuenta.numero) & " ," _
          & " tipo = " & cuenta.TipoCuenta & " ," _
          & " moneda_id = " & GetEntityId(cuenta.moneda) _
          & ", cbu = " & Escape(cuenta.CBU) _
          & " WHERE id = " & cuenta.Id
    End If

    Save = conectar.execute(q)
    Dim Id As Long
    If Save And cuenta.Id = 0 Then
        If conectar.UltimoId("AdminConfigCuentas", Id) Then
            cuenta.Id = Id
        End If
    End If
End Function
