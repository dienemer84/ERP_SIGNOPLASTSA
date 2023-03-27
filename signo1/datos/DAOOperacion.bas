Attribute VB_Name = "DAOOperacion"
Option Explicit

Public Function FindAll(Origen As OrigenOperacion, Optional ByVal extraFilter As String = "1 = 1") As Collection
    Dim q As String
    q = "SELECT *, (op.pertenencia + 0) as pertenencia2 From" _
      & " operaciones op" _
      & " LEFT JOIN AdminComprasCuentasContables cc ON op.cuenta_contable_id = cc.id" _
      & " LEFT JOIN AdminConfigMonedas mon ON op.moneda_id = mon.id" _
      & " LEFT JOIN cajas caj ON caj.id = op.cuentabanc_o_caja_id" _
      & " LEFT JOIN AdminConfigCuentas cu ON cu.id = op.cuentabanc_o_caja_id" _
      & " WHERE op.pertenencia = " & Origen & " AND " & extraFilter

    Dim col As New Collection
    Dim op As operacion

    Dim idx As Dictionary
    Dim rs As Recordset

    Set rs = conectar.RSFactory(q)
    BuildFieldsIndex rs, idx

    While Not rs.EOF
        Set op = Map(rs, idx, "op", "cc", "mon", "cu", "caj")

        col.Add op, CStr(op.Id)
        rs.MoveNext

    Wend

    Set FindAll = col
End Function


Public Function Delete(filter As String) As Boolean
    Delete = conectar.execute("DELETE FROM operaciones  WHERE " & filter)
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaCuentaContable As String = vbNullString, _
                    Optional tablaMoneda As String = vbNullString, _
                    Optional tablaCuentaBanc As String = vbNullString, _
                    Optional tablaCaja As String = vbNullString _
                  ) As operacion

    Dim op As operacion
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set op = New operacion
        op.Id = Id

        op.FechaCarga = GetValue(rs, indice, tabla, "fecha_carga")
        op.FechaOperacion = GetValue(rs, indice, tabla, "fecha_operacion")

        op.Pertenencia = GetValue(rs, indice, vbNullString, "pertenencia2")    'pertenencia2 = (pertenencia + 0) por el enum de mysql que viene con el nombre y no con valor numerico
        'Debug.Assert op.Pertenencia = 0

        op.Monto = GetValue(rs, indice, tabla, "monto")
        op.EntradaSalida = GetValue(rs, indice, tabla, "entrada_salida")
        op.Comprobante = GetValue(rs, indice, tabla, "comprobante")
        If LenB(tablaMoneda) > 0 Then Set op.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaCuentaContable) > 0 Then Set op.CuentaContable = DAOCuentaContable.Map(rs, indice, tablaCuentaContable)


        If op.Pertenencia = Banco Then    'cargo la cuenta
            Set op.CuentaBancaria = DAOCuentaBancaria.Map(rs, indice, tablaCuentaBanc)
        ElseIf op.Pertenencia = caja Then    'cargo la caja
            Set op.caja = DAOCaja.Map(rs, indice, tablaCaja)
        End If

    End If

    Set Map = op
End Function



Public Function Save(ope As operacion) As Boolean
    Dim q As String

    q = "INSERT INTO operaciones" _
      & " (monto," _
      & " moneda_id," _
      & " fecha_carga," _
      & " fecha_operacion," _
      & " cuenta_contable_id," _
      & " pertenencia," _
      & " cuentabanc_o_caja_id, entrada_salida,comprobante)" _
      & " Values" _
      & " ('monto'," _
      & " 'moneda_id'," _
      & " 'fecha_carga'," _
      & " 'fecha_operacion'," _
      & " 'cuenta_contable_id'," _
      & " 'pertenencia'," _
      & " 'cuentabanc_o_caja_id', 'entrada_salida','comprobante')"

    ope.FechaCarga = Now

    q = Replace(q, "'monto'", conectar.Escape(ope.Monto))
    q = Replace(q, "'moneda_id'", conectar.GetEntityId(ope.moneda))
    q = Replace(q, "'fecha_carga'", conectar.Escape(ope.FechaCarga))
    q = Replace(q, "'fecha_operacion'", conectar.Escape(ope.FechaOperacion))
    q = Replace(q, "'cuenta_contable_id'", conectar.GetEntityId(ope.CuentaContable))
    q = Replace(q, "'pertenencia'", ope.Pertenencia)
    q = Replace(q, "'entrada_salida'", ope.EntradaSalida)

    If LenB(ope.Comprobante) = 0 Then
        q = Replace(q, "'comprobante'", "'-'")
    Else

        q = Replace(q, "'comprobante'", conectar.Escape(ope.Comprobante))
    End If
    If ope.Pertenencia = Banco Then
        q = Replace(q, "'cuentabanc_o_caja_id'", conectar.GetEntityId(ope.CuentaBancaria))
    Else
        q = Replace(q, "'cuentabanc_o_caja_id'", conectar.GetEntityId(ope.caja))
    End If

    Save = conectar.execute(q)
End Function

