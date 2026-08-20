Attribute VB_Name = "DAOChequeras"
Option Explicit

Public Const TABLA_CHEQUERA As String = "chs"
Public Const TABLA_MONEDA As String = "mon"
Public Const TABLA_BANCO As String = "banco"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NUMERO As String = "numero"
Public Const CAMPO_OBSERVACIONES As String = "observaciones"
Public Const CAMPO_DESDE As String = "numero_desde"
Public Const CAMPO_HASTA As String = "numero_hasta"
Public Const CAMPO_FECHA_CREACION As String = "fecha_creacion"
Public Const CAMPO_USADA As String = "usada"

Public Sub llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As New Collection
    Set col = DAOChequeras.GetAll
    Dim bco As chequera
    cbo.Clear
    For Each bco In col
        cbo.AddItem bco.Description
        cbo.ItemData(cbo.NewIndex) = bco.Id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub


Public Function Guardar(chequera As chequera) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    Dim cheque As cheque
    Dim nuevo_id As Long
    Dim q As String
    q = "INSERT INTO Chequeras" _
      & "(numero," _
      & "id_banco," _
      & "id_cuenta_bancaria," _
      & "numero_desde," _
      & "numero_hasta," _
      & "fecha_creacion," _
      & "id_moneda," _
      & "observaciones) Values" _
      & "('numero'," _
      & "'id_banco'," _
      & "'id_cuenta_bancaria'," _
      & "'numero_desde'," _
      & "'numero_hasta'," _
      & "'fecha_creacion'," _
      & "'id_moneda'," _
      & "'observaciones')" _

    q = Replace(q, "'numero'", Escape(chequera.numero))
        q = Replace(q, "'id_banco'", Escape(chequera.Banco.Id))
        q = Replace(q, "'numero_desde'", Escape(chequera.NumeroDesde))
        q = Replace(q, "'numero_hasta'", Escape(chequera.NumeroHasta))
        q = Replace(q, "'id_moneda'", Escape(chequera.moneda.Id))
        q = Replace(q, "'fecha_creacion'", Escape(chequera.fechaCreacion))
        q = Replace(q, "'id_cuenta_bancaria'", Escape(chequera.CuentaBancaria.Id))
        q = Replace(q, "'observaciones'", Escape(chequera.Observaciones))

    Guardar = conectar.execute(q)
    If Not Guardar Then GoTo err1
    conectar.UltimoId "Chequeras", nuevo_id
    chequera.Id = nuevo_id
    If chequera.Cheques.count > 0 Then
        For Each cheque In chequera.Cheques
            cheque.IdChequera = chequera.Id
            If Not DAOCheques.Guardar(cheque) Then GoTo err1
        Next cheque
    End If
    conectar.CommitTransaction
    Exit Function
err1:
    Guardar = False
    conectar.RollBackTransaction
End Function

Public Function GetById(Id As Long) As chequera
    Dim col As Collection
    Set col = GetAll("chs.id = " & Id)
    If col.count = 0 Then
        Set GetById = Nothing
    Else
        Set GetById = col.item(1)
    End If
End Function

'''Public Function FindAllWithChequesDisponibles() As Collection
'''    Dim F As String
'''    F = "chs.id IN (SELECT DISTINCT id_chequera FROM Cheques WHERE id_chequera IS NOT NULL AND en_cartera = 0 AND fecha_vencimiento IS NULL ) ORDER BY fecha_creacion DESC"
'''    Set FindAllWithChequesDisponibles = GetAll(F)
'''End Function

Public Function FindAllWithChequesDisponibles() As Collection

    Dim F As String

    F = "IFNULL(chs.usada, 0) = 0 " _
      & "AND chs.id IN (" _
      & "   SELECT DISTINCT id_chequera " _
      & "   FROM Cheques " _
      & "   WHERE id_chequera IS NOT NULL " _
      & "   AND en_cartera = 0 " _
      & "   AND fecha_vencimiento IS NULL " _
      & ") " _
      & "ORDER BY chs.fecha_creacion DESC"

    Set FindAllWithChequesDisponibles = GetAll(F)

End Function


Public Function GetAll(Optional filtro As String = Empty) As Collection    'of chequeras
    On Error GoTo err1
    Dim rs As Recordset
    Dim col As New Collection
    Dim q As String
    Dim indice As Dictionary
    
    q = "SELECT * FROM Chequeras chs " _
      & "LEFT JOIN AdminConfigBancos banco " _
      & " ON chs.id_banco = banco.id " _
      & "LEFT JOIN AdminConfigMonedas mon " _
      & " ON chs.id_moneda = mon.id " _
      & "LEFT JOIN AdminConfigCuentas cta " _
      & " ON cta.id = chs.id_cuenta_bancaria " _
      & "LEFT JOIN AdminConfigBancos bancocta " _
      & " ON bancocta.id = cta.idBanco " _
      & "LEFT JOIN AdminConfigMonedas moncta " _
      & " ON moncta.id = cta.moneda_id " _
      & "WHERE 1=1"
      
    If LenB(filtro) > 0 Then
        q = q & " AND " & filtro
    End If
    
    Set rs = conectar.RSFactory(q)
    
    conectar.BuildFieldsIndex rs, indice
    
    While Not rs.EOF
        col.Add Map(rs, indice, _
            TABLA_CHEQUERA, _
            TABLA_MONEDA, _
            TABLA_BANCO, _
            "cta", _
            "moncta", _
            "bancocta")
        rs.MoveNext
        
    Wend

    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function


Public Function Map( _
    ByRef rs As Recordset, _
    ByRef indice As Dictionary, _
    ByRef tablaChequera As String, _
    Optional tablaMoneda As String, _
    Optional tablaBanco As String, _
    Optional tablaCuentaBancaria As String, _
    Optional tablaMonedaCuenta As String, _
    Optional tablaBancoCuenta As String) As chequera

    Dim tmp As chequera
    Dim Id As Variant
    Dim valorUsada As Variant

    Id = GetValue(rs, indice, tablaChequera, CAMPO_ID)

    If Id > 0 Then
        Set tmp = New chequera

        tmp.Id = Id
        
        tmp.numero = GetValue( _
            rs, indice, tablaChequera, CAMPO_NUMERO)

        tmp.NumeroDesde = GetValue( _
            rs, indice, tablaChequera, CAMPO_DESDE)

        tmp.fechaCreacion = GetValue( _
            rs, indice, tablaChequera, CAMPO_FECHA_CREACION)

        tmp.NumeroHasta = GetValue( _
            rs, indice, tablaChequera, CAMPO_HASTA)

        valorUsada = GetValue( _
            rs, indice, tablaChequera, CAMPO_USADA)

        If IsNull(valorUsada) Or IsEmpty(valorUsada) Then
            tmp.usada = False
        Else
            tmp.usada = (CLng(valorUsada) <> 0)
        End If

        If LenB(tablaMoneda) > 0 Then
            Set tmp.moneda = DAOMoneda.Map( _
                rs, indice, tablaMoneda)
        End If

        If LenB(tablaCuentaBancaria) > 0 Then
        
            Set tmp.CuentaBancaria = _
                DAOCuentaBancaria.Map( _
                    rs, _
                    indice, _
                    tablaCuentaBancaria, _
                    tablaBancoCuenta, _
                    tablaMonedaCuenta)
        
        End If
    End If

    Set Map = tmp

End Function

Public Function ActualizarUsada( _
    ByVal IdChequera As Long, _
    ByVal usada As Boolean) As Boolean
    
    On Error GoTo err1
    
    Dim q As String
    Dim valorUsada As Integer
    
    valorUsada = Abs(CInt(usada))
    
    q = "UPDATE Chequeras " _
      & "SET usada = " & valorUsada & "  " _
      & "WHERE id = " & IdChequera
    
    ActualizarUsada = conectar.execute(q)
    Exit Function
    
err1:
    ActualizarUsada = False


End Function

