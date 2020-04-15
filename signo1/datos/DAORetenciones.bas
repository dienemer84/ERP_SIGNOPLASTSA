Attribute VB_Name = "DAORetenciones"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_NOMBRE As String = "retencion"
Public Const CAMPO_CODIGO As String = "codigo"
Public Const CAMPO_PORCENTAJE As String = "porcentaje"
Public Const CAMPO_MINIMO As String = "minimo_imponible"
Public Const TABLA_RETENCION As String = "ret"

Public Function FindAllWithAlicuotas(Cuit As String) As Collection

    Dim d As New clsDTOPadronIIBB
    Dim col2 As New Collection
    Dim ali As New Collection
          
            Set col2 = DTOPadronIIBB.FindByCUIT2(Cuit)
          
            Dim retenciones As Collection
            Set retenciones = FindAllEsAgente
                 
            Dim rx As Retencion
            Dim c As clsDTOPadronIIBB
            Set alicuotas = New Collection
            Dim x As DTORetencionAlicuota
            For Each c In col2
        
            For Each rx In retenciones
            
            If rx.IdPadron = c.IdPadron Then
                
            Set x = New DTORetencionAlicuota
            x.alicuotaRetencion = c.alicuotaRetencion
            x.alicuotaPercepcion = c.alicuotaPercepcion
            Set x.Retencion = rx
            ali.Add x, CStr(c.IdPadron)
                
            End If
            
            Next
        
        Next
        
        Set FindAllWithAlicuotas = ali
        
End Function

Public Function FindById(id As Long) As Retencion
    Dim col As Collection: Set col = FindAll("id = " & id)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function

Public Function FindAllEsAgente() As Collection
    Set FindAllEsAgente = FindAll("1=1 and retiene=1")  'fix 28-3-2020  'reemplazar por EsAgente cuando vea el tema de los permisos de la tabla.
End Function

Public Function FindAll(Optional whereFilter As String = "1 = 1") As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount
    Dim rs As ADODB.Recordset
    Dim q As String
    Dim col As New Collection

    q = "SELECT * from retenciones ret WHERE " & whereFilter


    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim ret As Retencion

    While Not rs.EOF
        Set ret = Map(rs, fieldsIndex, DAORetenciones.TABLA_RETENCION)
        col.Add ret, CStr(ret.id)
        rs.MoveNext
    Wend

    Set FindAll = col
End Function


Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As Retencion
    Dim T As Retencion
    Dim id As Long
    id = GetValue(rs, indice, tabla, DAOTareas.CAMPO_ID)
    If id > 0 Then
        Set T = New Retencion
        T.id = id
        T.codigo = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_CODIGO)
        T.nombre = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_NOMBRE)
        T.Porcentaje = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_PORCENTAJE)
        T.MinimoImponible = GetValue(rs, indice, tabla, DAORetenciones.CAMPO_MINIMO)
        T.IdPadron = GetValue(rs, indice, tabla, "id_padron")
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
        cbo.ItemData(cbo.NewIndex) = ret.id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If

End Function
