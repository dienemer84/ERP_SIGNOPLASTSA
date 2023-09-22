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


Public Sub llenarComboXtremeSuite(cbo As XtremeSuiteControls.ComboBox)
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
      & "numero_desde," _
      & "numero_hasta," _
      & "fecha_creacion," _
      & "id_moneda," _
      & "observaciones) Values" _
      & "('numero'," _
      & "'id_banco'," _
      & "'numero_desde'," _
      & "'numero_hasta'," _
      & "'fecha_creacion'," _
      & "'id_moneda','observaciones')" _

q = Replace(q, "'numero'", Escape(chequera.numero))
    q = Replace(q, "'id_banco'", Escape(chequera.Banco.Id))
    q = Replace(q, "'numero_desde'", Escape(chequera.NumeroDesde))
    q = Replace(q, "'numero_hasta'", Escape(chequera.NumeroHasta))
    q = Replace(q, "'id_moneda'", Escape(chequera.moneda.Id))
    q = Replace(q, "'fecha_creacion'", Escape(chequera.FechaCreacion))
    q = Replace(q, "'observaciones'", Escape(chequera.observaciones))

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
        Set GetById = col.Item(1)
    End If
End Function

Public Function FindAllWithChequesDisponibles() As Collection
    Dim F As String
    F = "chs.id IN (SELECT DISTINCT id_chequera FROM Cheques WHERE id_chequera IS NOT NULL AND en_cartera = 0 AND fecha_vencimiento IS NULL ) ORDER BY fecha_creacion DESC"
    Set FindAllWithChequesDisponibles = GetAll(F)
End Function


Public Function GetAll(Optional filtro As String = Empty) As Collection    'of chequeras
    On Error GoTo err1
    Dim rs As Recordset
    Dim col As New Collection
    Dim q As String
    Dim indice As Dictionary
    q = "SELECT * FROM Chequeras chs LEFT JOIN AdminConfigBancos banco ON chs.id_banco=banco.id LEFT JOIN AdminConfigMonedas mon ON chs.id_moneda=mon.id WHERE 1=1"

    If LenB(filtro) > 0 Then
        q = q & " AND " & filtro
    End If
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, indice
    While Not rs.EOF
        col.Add Map(rs, indice, TABLA_CHEQUERA, TABLA_MONEDA, TABLA_BANCO)
        rs.MoveNext
    Wend

    Set GetAll = col
    Exit Function
err1:
    Set GetAll = Nothing
End Function

Public Function Map(ByRef rs As Recordset, ByRef indice As Dictionary, ByRef tablaChequera As String, Optional tablaMoneda As String, Optional tablaBanco As String) As chequera
    Dim tmp As chequera

    Dim Id As Variant
    Id = GetValue(rs, indice, tablaChequera, CAMPO_ID)
    If Id > 0 Then
        Set tmp = New chequera

        tmp.Id = Id
        tmp.numero = GetValue(rs, indice, tablaChequera, DAOChequeras.CAMPO_NUMERO)
        tmp.NumeroDesde = GetValue(rs, indice, tablaChequera, CAMPO_DESDE)
        tmp.FechaCreacion = GetValue(rs, indice, tablaChequera, CAMPO_FECHA_CREACION)
        tmp.NumeroHasta = GetValue(rs, indice, tablaChequera, CAMPO_HASTA)
        If LenB(tablaMoneda) > 0 Then Set tmp.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)
        If LenB(tablaBanco) > 0 Then Set tmp.Banco = DAOBancos.Map(rs, indice, tablaBanco)






    End If

    Set Map = tmp
End Function

