Attribute VB_Name = "DAOComprobantesRecibidos"
Option Explicit

Public Function FindAll(Optional ByVal filter As String = vbNullString) As Collection
    Dim q As String
    q = "SELECT * " _
    & " FROM sp_temporal.ComprobantesRecibidosAFIP as A" _
    & " LEFT JOIN sp_temporal.ComprobantesCargadosSP as B" _
    & " ON A.clave = B.clave" _
    & " WHERE B.clave IS NULL"
    
    Dim col As New Collection
    Dim rs As Recordset
    
    Set rs = conectar.RSFactory(q)
    
    Dim fieldsIndex As New Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Dim sin As ComprobantesRecibidos
   

    While Not rs.EOF
     Set sin = Map(rs, fieldsIndex, "A", "id")
        col.Add sin, CStr(sin.Clave_)
        
        rs.MoveNext
    Wend

    Set FindAll = col

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
    
    'End If

    Set Map = s
End Function


