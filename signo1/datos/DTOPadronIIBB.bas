Attribute VB_Name = "DTOPadronIIBB"
Option Explicit

Public Function FindByCUIT(Cuit As String, Tipo As TipoPadron) As clsDTOPadronIIBB
    Dim rs As Recordset
    Dim dto As clsDTOPadronIIBB
    Dim q As String

    If Tipo = TipoPadron.TipoPadronPercepcion Then
        q = "select * from sp_permisos.IIBB2_Percepcion where Cuit=" & Cuit
    ElseIf Tipo = TipoPadron.TipoPadronRetencion Then
        q = "select * from sp_permisos.IIBB2_Retencion where Cuit=" & Cuit
    End If



    Set rs = conectar.RSFactory(q)
    If Not rs.EOF And Not rs.BOF Then
        Set dto = New clsDTOPadronIIBB

        dto.Discriminador = rs!Discriminador
        dto.AltaBaja = rs!AltaBaja
        dto.Cambio = rs!Cambio
        dto.Cuit = rs!Cuit
        dto.FechaDesde = rs!FechaDesde
        dto.FechaHasta = rs!FechaHasta
        dto.FechaPublicacion = rs!FechaPublicacion
        dto.Grupo = rs!Grupo
        dto.alicuota = CDbl(rs!alicuota)
        'dto.Retencion = CDbl(rs!Retencion)
        dto.Tipo = rs!Tipo
    End If
    Set FindByCUIT = dto
End Function

Public Function FindByCUIT2(Cuit As String, Optional padron As Long = 0) As Collection
    
    Dim rs As Recordset
    Dim dto As clsDTOPadronIIBB
    Dim col As New Collection
    Dim q As String
   

        q = "select * from sp_permisos.Padron_Detalles where Cuit=" & Cuit

   
        If padron > 0 Then
         q = q & " id_padron=" & padron
        End If


    Set rs = conectar.RSFactory(q)
        While Not rs.EOF
            Set dto = New clsDTOPadronIIBB
            
            dto.id = rs!id
            dto.alicuotaRetencion = CDbl(rs!alicuotaRetencion)
            dto.alicuotaPercepcion = CDbl(rs!alicuotaPercepcion)
            dto.IdPadron = CLng(rs!padron)
    
            col.Add dto, dto.id
            rs.MoveNext
            
        Wend
    
    Set FindByCUIT2 = col
    
End Function

