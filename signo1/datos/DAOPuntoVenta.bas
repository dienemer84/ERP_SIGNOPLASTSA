Attribute VB_Name = "DAOPuntoVenta"
Option Explicit


Public Function FindAll(Optional filtro As String) As Collection
    On Error GoTo err1
    Dim idx As Dictionary
    Dim rs As Recordset
    Dim strsql As String
    strsql = "Select * from AdminConfigFacturaPuntoVenta pv where 1=1"

    If LenB(filtro) > 0 Then strsql = strsql & filtro
    Dim col As New Collection
    Set rs = conectar.RSFactory(strsql)
    conectar.BuildFieldsIndex rs, idx

    While Not rs.EOF And Not rs.BOF
        col.Add Map(rs, idx, "pv")
        rs.MoveNext
    Wend

    Set FindAll = col
    Exit Function
err1:
    Set FindAll = Nothing

End Function

Public Function FindById(id As Long) As PuntoVenta

    Set FindById = FindAll(" And pv.id=" & id)(1)
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As PuntoVenta

    Dim pv As PuntoVenta
    Dim id As Long: id = GetValue(rs, indice, tabla, "id")

    If id > 0 Then
        Set pv = New PuntoVenta
        pv.id = id
        pv.descripcion = GetValue(rs, indice, tabla, "descripcion")
        pv.PuntoVenta = GetValue(rs, indice, tabla, "punto_venta")
        pv.EsElectronico = GetValue(rs, indice, tabla, "esElectronico")
        'pv.EsCredito = GetValue(rs, indice, tabla, "esCredito")
        pv.CaeManual = GetValue(rs, indice, tabla, "caeManual")
    pv.default = GetValue(rs, indice, tabla, "default")

    End If

    Set Map = pv
End Function


Public Function GetDefaultOrFirst() As PuntoVenta
Set GetDefaultOrFirst = FindAll(" And pv.default=1")(1)
End Function



Public Function llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox, Optional marcarDefault As Boolean = False)
    Dim col As Collection
    Dim pv As PuntoVenta
    Set col = FindAll

    cbo.Clear

    Dim nidx As Long
    Dim idDefault As Long



    For Each pv In col
        cbo.AddItem Format(pv.PuntoVenta, "000") & " - " & pv.descripcion
        nidx = cbo.NewIndex
        cbo.ItemData(nidx) = pv.id
        If pv.default Then idDefault = nidx
    Next
    
    If marcarDefault Then
        cbo.ListIndex = idDefault
    End If
    
    

End Function
