Attribute VB_Name = "DAOComprobantesRecibidos"
Option Explicit

Public Function FindAll(Optional ByVal filter As String = vbNullString) As Collection
    On Error GoTo ErrorHandler
    
    Dim q As String
    Dim col As New Collection
    Dim rs As Recordset
    Dim fieldsIndex As New Dictionary
    Dim sin As ComprobantesRecibidos
    Dim success As Boolean
    Dim claveActual As String
    Dim contadorDuplicados As Long
    success = False
    contadorDuplicados = 0
    
'''    ' Validar conexión
'''    If conectar Is Nothing Then
'''        Err.Raise vbObjectError + 1001, "FindAll", "Objeto 'conectar' no está inicializado"
'''    End If
    
    ' Construir consulta SQL
    q = "SELECT * " _
      & " FROM sp_temporal.ComprobantesRecibidosAFIP as A" _
      & " LEFT JOIN sp_temporal.ComprobantesCargadosSP as B" _
      & " ON A.clave = B.clave" _
      & " WHERE B.clave IS NULL"
    
    ' Obtener recordset
    Set rs = conectar.RSFactory(q)
    If rs Is Nothing Then
        Err.Raise vbObjectError + 1002, "FindAll", "No se pudo obtener el Recordset"
    End If
    
    ' Verificar si hay registros
    If rs.BOF And rs.EOF Then
        Set FindAll = col  ' Retorna colección vacía
        success = True
        Exit Function
    End If
    
    ' Construir índice de campos
    BuildFieldsIndex rs, fieldsIndex
    If fieldsIndex.count = 0 Then
        Err.Raise vbObjectError + 1003, "FindAll", "No se pudieron mapear los campos"
    End If
    
    ' Procesar registros
    While Not rs.EOF
        Set sin = Map(rs, fieldsIndex, "A", "id")
        
        ' Validar objeto mapeado
        If sin Is Nothing Then
            Err.Raise vbObjectError + 1004, "FindAll", "Error al mapear el registro"
        End If
        
        ' Obtener clave y validar
        claveActual = CStr(sin.Clave_)
        If Len(claveActual) = 0 Then
            Err.Raise vbObjectError + 1005, "FindAll", "Clave inválida en registro"
        End If
        
        ' Intentar agregar a la colección con manejo de duplicados
        On Error Resume Next
        col.Add sin, claveActual
        If Err.Number = 457 Then ' Clave duplicada
            contadorDuplicados = contadorDuplicados + 1
            ' Agregar con clave modificada o simplemente omitir (depende de tu necesidad)
            col.Add sin, claveActual & "_" & contadorDuplicados
            Err.Clear
        ElseIf Err.Number <> 0 Then
            ' Otro error inesperado
            Err.Raise Err.Number, "FindAll.Add", Err.Description
        End If
        On Error GoTo ErrorHandler
        
        rs.MoveNext
    Wend
    
    ' Registrar advertencia si hubo duplicados
    If contadorDuplicados > 0 Then
        Debug.Print "Advertencia: Se encontraron " & contadorDuplicados & " claves duplicadas"
    End If
    
    Set FindAll = col
    success = True
    
ExitProcedure:
    ' Limpieza
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    If Not success Then
        Set FindAll = New Collection  ' Retorna colección vacía en caso de error
    End If
    
    Exit Function
    
ErrorHandler:
    ' Registrar error
    Debug.Print "Error " & Err.Number & " en FindAll: " & Err.Description
    Resume ExitProcedure
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional Id As String = vbNullString) As ComprobantesRecibidos

    Dim s As ComprobantesRecibidos

    Id = GetValue(rs, indice, tabla, "clave")

    'If _ > 0 Then
    Set s = New ComprobantesRecibidos

    s.Clave_ = Id

    s.Fecha_ = GetValue(rs, indice, tabla, "fecha")
    s.Tipo_ = GetValue(rs, indice, tabla, "tipo")
    s.PuntoDeVenta_ = GetValue(rs, indice, tabla, "puntodeventa")
    s.NumeroDesde_ = GetValue(rs, indice, tabla, "numerodesde")
    s.TipoDocEmisor_ = GetValue(rs, indice, tabla, "tipodocemisor")
    s.NroDocEmisor_ = GetValue(rs, indice, tabla, "nrodocemisor")
    s.DenominacionEmisor_ = GetValue(rs, indice, tabla, "denominacionemision")
    s.TipoCambio_ = GetValue(rs, indice, tabla, "tipocambio")
    s.Moneda_ = GetValue(rs, indice, tabla, "moneda")
    s.ImpNetoGravado_ = GetValue(rs, indice, tabla, "impnetogravado")
    s.ImpNetoNoGravado_ = GetValue(rs, indice, tabla, "impnetonogravado")
    s.ImpOpExentas_ = GetValue(rs, indice, tabla, "impopexentas")
    s.Iva_ = GetValue(rs, indice, tabla, "iva")
    s.ImpTotal_ = GetValue(rs, indice, tabla, "imptotal")
    s.Clave_ = GetValue(rs, indice, tabla, "clave")

'    Debug.Print (s.Clave_)
    'End If

    Set Map = s
End Function


