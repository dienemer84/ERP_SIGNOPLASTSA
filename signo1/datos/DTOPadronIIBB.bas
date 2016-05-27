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
        dto.Alicuota = CDbl(rs!Alicuota)
        'dto.Retencion = CDbl(rs!Retencion)
        dto.Tipo = rs!Tipo
    End If
    Set FindByCUIT = dto
End Function

