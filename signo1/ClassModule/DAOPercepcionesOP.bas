Attribute VB_Name = "DAOPercepcionesOP"
Option Explicit

Public Function Save(per As clsPercepcionesOrdenPago) As Boolean
    Dim q As String

    q = "INSERT INTO AdminPercepcionOPDetalle" _
      & " (monto," _
      & " moneda_id," _
      & " fecha," _
      & " comprobante," _
      & " tipo)" _
      & " Values" _
      & " ('monto'," _
      & " 'moneda_id'," _
      & " 'fecha'," _
      & " 'comprobante'," _
      & " 'tipo')"

    q = Replace(q, "'monto'", conectar.Escape(per.Monto))
    q = Replace(q, "'moneda_id'", conectar.GetEntityId(per.moneda))
    q = Replace(q, "'fecha'", conectar.Escape(per.FEcha))

    If LenB(per.Comprobante) = 0 Then
        q = Replace(q, "'comprobante'", "'-'")
    Else

        q = Replace(q, "'comprobante'", conectar.Escape(per.Comprobante))
    End If

    If LenB(per.Tipo) = 0 Then
        q = Replace(q, "'tipo'", "'-'")
    Else

        q = Replace(q, "'tipo'", conectar.Escape(per.Tipo))
    End If
    
    Save = conectar.execute(q)
    
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, _
                    Optional tablaMoneda As String = vbNullString _
                  ) As clsPercepcionesOrdenPago

    Dim per As clsPercepcionesOrdenPago
    
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")

    If Id > 0 Then
        Set per = New clsPercepcionesOrdenPago
        per.Id = Id

        per.FEcha = GetValue(rs, indice, tabla, "fecha")
        per.Monto = GetValue(rs, indice, tabla, "monto")
        If LenB(tablaMoneda) > 0 Then Set per.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        per.Comprobante = GetValue(rs, indice, tabla, "comprobante")
        per.Tipo = GetValue(rs, indice, tabla, "tipo")
        
        On Error GoTo 0 ' Restaurar el manejo normal de errores
        
    End If

    Set Map = per
End Function

