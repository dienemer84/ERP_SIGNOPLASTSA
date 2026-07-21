Attribute VB_Name = "DAORetenciones"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "retencion"
Public Const CAMPO_CODIGO As String = "codigo"
Public Const CAMPO_PORCENTAJE As String = "porcentaje"
Public Const CAMPO_MINIMO As String = "minimo_imponible"
Public Const TABLA_RETENCION As String = "ret"


Public Function FindAllWithAlicuotas(Cuit As String) As Collection
    Dim colPadron As Collection
    Dim retenciones As Collection
    Dim ali As Collection

    Dim rx As Retencion
    Dim c As clsDTOPadronIIBB
    Dim x As DTORetencionAlicuota

    Set colPadron = DTOPadronIIBB.FindByCUIT2(Cuit)
    Set retenciones = FindAllEsAgente()
    Set ali = New Collection

    For Each rx In retenciones

        Set c = Nothing

        If rx.idPadron <> 0 Then
            Set c = BuscarPadronPorId(colPadron, rx.idPadron)
        End If

        Set x = New DTORetencionAlicuota
        Set x.Retencion = rx
        x.Importe = 0

        If Not c Is Nothing Then
            x.alicuotaRetencion = c.alicuotaRetencion
            x.alicuotaPercepcion = c.alicuotaPercepcion
            x.dePadron = True
        Else
            x.dePadron = False
        End If

        ali.Add x

    Next rx

    Set FindAllWithAlicuotas = ali
End Function


Private Function Contains(r As Retencion, c As Collection)
    Dim c1 As Boolean
    c1 = False
    Dim i As DTORetencionAlicuota
    For Each i In c
        If i.Retencion.Id = r.Id Then
            c1 = True
        End If
    Next i
    Contains = c1
End Function


Public Function FindAllWithAlicuotasAnt(Cuit As String) As Collection
    Dim colPadron As Collection
    Dim retenciones As Collection
    Dim ali As Collection

    Dim rx As Retencion
    Dim c As clsDTOPadronIIBB
    Dim x As DTORetencionAlicuota

    Set colPadron = DTOPadronIIBB.FindByCUITAnt(Cuit)
    Set retenciones = FindAllEsAgente()
    Set ali = New Collection

    For Each rx In retenciones

        Set c = Nothing

        If rx.idPadron <> 0 Then
            Set c = BuscarPadronPorId(colPadron, rx.idPadron)
        End If

        Set x = New DTORetencionAlicuota
        Set x.Retencion = rx
        x.Importe = 0

        If Not c Is Nothing Then
            x.alicuotaRetencion = c.alicuotaRetencion
            x.alicuotaPercepcion = c.alicuotaPercepcion
            x.dePadron = True
        Else
            x.dePadron = False
        End If

        ali.Add x

    Next rx

    Set FindAllWithAlicuotasAnt = ali
End Function

Public Function FindById(Id As Long) As Retencion
    Dim col As Collection: Set col = FindAll("id = " & Id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function

Public Function FindAllEsAgente() As Collection
    Set FindAllEsAgente = FindAll("1=1 AND retiene=1", OrderByRetencionesFijo())
    
End Function

Public Function FindAll(Optional whereFilter As String = "1 = 1", Optional orderBy As String = "") As Collection
    Dim rs As ADODB.Recordset
    Dim q As String
    Dim col As New Collection

    q = "SELECT * FROM retenciones ret WHERE " & whereFilter

    If LenB(Trim$(orderBy)) > 0 Then
        q = q & " ORDER BY " & orderBy
    End If

    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim ret As Retencion

    While Not rs.EOF
        Set ret = Map(rs, fieldsIndex, DAORetenciones.TABLA_RETENCION)
        col.Add ret, CStr(ret.Id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As Retencion
    Dim T As Retencion
    Dim Id As Long
    Id = GetValue(rs, indice, tabla, DAOTareas.CAMPO_ID)
    If Id > 0 Then
        Set T = New Retencion
        T.Id = Id
        T.codigo = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_CODIGO)
        T.nombre = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_NOMBRE)
        T.Porcentaje = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_PORCENTAJE)
        T.MinimoImponible = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_MINIMO)
        T.idPadron = GetValue(rs, indice, tabla, "id_padron")
    End If

    Set Map = T
End Function


Public Function llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As Collection
    Set col = DAORetenciones.FindAll()
    Dim ret As Retencion
    cbo.Clear

    For Each ret In col
        cbo.AddItem ret.codigo & "-" & ret.nombre
        cbo.ItemData(cbo.NewIndex) = ret.Id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

End Function


Private Function OrderByRetencionesFijo() As String
    OrderByRetencionesFijo = _
        "CASE " & _
        " WHEN UPPER(ret.retencion) LIKE '%IIBB%' " & _
        "  AND (UPPER(ret.retencion) LIKE '%BS AS%' " & _
        "       OR UPPER(ret.retencion) LIKE '%BSAS%' " & _
        "       OR UPPER(ret.retencion) LIKE '%BUENOS AIRES%') THEN 10 " & _
        " WHEN UPPER(ret.retencion) LIKE '%IIBB%' " & _
        "  AND UPPER(ret.retencion) LIKE '%CABA%' THEN 20 " & _
        " WHEN UPPER(ret.retencion) LIKE '%GANANCIAS%' THEN 30 " & _
        " ELSE 999 " & _
        "END, ret.retencion"
End Function


Private Function BuscarPadronPorId(ByVal colPadron As Collection, ByVal idPadron As Long) As clsDTOPadronIIBB
    Dim c As clsDTOPadronIIBB

    For Each c In colPadron
        If c.idPadron = idPadron Then
            Set BuscarPadronPorId = c
            Exit Function
        End If
    Next c

    Set BuscarPadronPorId = Nothing
End Function
