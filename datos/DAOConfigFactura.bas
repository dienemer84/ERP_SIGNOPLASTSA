Attribute VB_Name = "DAOConfigFactura"
Dim rs As ADODB.recordset
Dim cn As ADODB.Connection
Dim configFactura As clsConfigFacturas
Public Function getById(Id) As clsConfigFacturas
Set rs = conectar.RSFactory("select * from AdminConfigFacturas where id=" & Id)
If Not rs.EOF And Not rs.BOF Then
Set configFactura = New clsConfigFacturas
    configFactura.Id = rs!Id
    configFactura.Discrimina = rs!discriminaIVA
    configFactura.tipoFactura = DAOTipoFactura.getById(CLng(rs!tipoFactura))
    
    configFactura.tipoIva = DAOTipoIva.getById(rs!idIVA)
    
    Set getById = configFactura
Else
    Set getById = Nothing
    
End If
End Function



Public Function getByIVA(id_iva) As clsConfigFacturas

Set rs = conectar.RSFactory("select * from AdminConfigFacturas where idIva=" & id_iva)
Set configFactura = New clsConfigFacturas


If Not rs.EOF And Not rs.BOF Then
    configFactura.Id = rs!Id
    configFactura.Discrimina = rs!discriminaIVA
    configFactura.tipoFactura = DAOTipoFactura.getById(rs!tipoFactura)
    configFactura.tipoIva = DAOTipoIva.getById(rs!idIVA)
    Set getByIVA = configFactura
Else
Set configFactura = Nothing
End If
End Function






