Attribute VB_Name = "DAOPeticionOferta"
Dim rs As ADODB.Recordset


Public Function GetById(Id As Long) As clsPeticionOferta

    On Error GoTo err1
    Dim col As New Collection
    Dim po As clsPeticionOferta
    strsql = "Select * from ComprasPeticionOferta where id=" & Id
    Set rs = conectar.RSFactory(strsql)
    If Not rs.EOF And Not rs.BOF Then
        Set po = New clsPeticionOferta
        po.idReque = rs!id_reque
        po.FechaEmision = rs!FechaEmision
        po.FechaSolicitada = rs!FechaPedido
        po.usuarioCreador = DAOUsuarios.GetById(rs!usuarioCreador)
        po.numero = rs!Id
        po.Modificado = Now
        po.estado = rs!estado
        Set po.moneda = DAOMoneda.GetById(rs!moneda_id)
        po.Proveedor = DAOProveedor.FindById(rs!id_proveedor)
        po.detalle = Nothing    'DAOPeticionOfertaDetalle.GetByIdPO(rs!id)

        po.FormaDePago = rs!FormaPago
        po.CantidadDiasPago = rs!CantDiasPago
        po.PorcentajeDescuento = rs!PorcentajeDescuento
        po.EntregaRetiramos = rs!EntregaRetiramos

        col.Add po
    End If

    Set GetById = po

    Exit Function
err1:
    Set GetById = Nothing

End Function



Public Function GetAll(Optional filter As String = vbNullString) As Collection
    On Error GoTo err1
    Dim col As New Collection
    Dim po As clsPeticionOferta
    strsql = "Select * from ComprasPeticionOferta where 1  = 1 "


    If LenB(filter) > 0 Then
        strsql = strsql & filter
    End If
    Set rs = conectar.RSFactory(strsql)
    While Not rs.EOF
        Set po = New clsPeticionOferta
        po.idReque = rs!id_reque
        po.FechaEmision = rs!FechaEmision
        po.FechaSolicitada = rs!FechaPedido
        po.usuarioCreador = DAOUsuarios.GetById(rs!usuarioCreador)
        po.numero = rs!Id
        'po.Modificado = Now
        po.estado = rs!estado
        Set po.moneda = DAOMoneda.GetById(rs!moneda_id)
        po.Proveedor = DAOProveedor.FindById(rs!id_proveedor)
        po.detalle = DAOPeticionOfertaDetalle.FindAll(rs!Id)    'Nothing 'experimental

        po.FormaDePago = rs!FormaPago
        po.CantidadDiasPago = rs!CantDiasPago
        po.PorcentajeDescuento = rs!PorcentajeDescuento
        po.EntregaRetiramos = rs!EntregaRetiramos

        col.Add po
        rs.MoveNext
    Wend
    Set GetAll = col

    Exit Function
err1:
    Set GetAll = Nothing
End Function

Public Function Update(po As clsPeticionOferta) As Boolean
    Dim q As String
    q = "UPDATE ComprasPeticionOferta SET" _
      & " moneda_id = " & GetEntityId(po.moneda) _
      & ", formaPago = " & conectar.Escape(po.FormaDePago) _
      & ", cantDiasPago = " & conectar.Escape(po.CantidadDiasPago) _
      & ", porcentajeDescuento = " & conectar.Escape(po.PorcentajeDescuento) _
      & ", entregaRetiramos = " & conectar.Escape(po.EntregaRetiramos) _
      & " WHERE id = " & po.numero
    Update = conectar.execute(q)
End Function


Public Function Nueva(T As clsRequerimiento) As Boolean
'crea todas las peticiones de oferta de este requerimiento (para cada proveedor)
    On Error GoTo err1

    antestado = T.estado
    T.estado = EnPO_

    If antestado = EstadoRequeCompra.Procesado_ Or antestado = EstadoRequeCompra.ProcesadoParcial_ Then

        conectar.BeginTransaction
        Nueva = DAORequerimiento.Guardar(T)

        If Not Nueva Then GoTo err1

        Dim Id_ As Long
        Dim colProv As Collection
        Dim colItems As Collection
        Dim colPOItems As New Collection
        Dim strsql As String
        Dim tmpProv As clsProveedor
        Set colProv = DAORequeProveedores.GetAllByReque(T)
        Dim tmpRequeDetalle As clsRequeMateriales
        Dim tmpDetallePO As clsPeticionOfertaDetalle
        Dim po As clsPeticionOferta
        'creo el detalle

        Dim entregaDetalle As EntregaPetOfDetalle
        Dim A As clsRequeEntregas

        For x = 1 To colProv.count
            Set tmpProv = colProv.item(x)
            Set po = New clsPeticionOferta
            'Debug.Print tmpProv.id & " " & tmpProv.razonFantasia
            Set colItems = DAORequeMateriales.GetByRequeByProveedor(T.Id, tmpProv.Id)
            Set colPOItems = Nothing
            For y = 1 To colItems.count
                Set tmpDetallePO = New clsPeticionOfertaDetalle
                Set tmpRequeDetalle = colItems.item(y)


                tmpDetallePO.DetalleReque = tmpRequeDetalle
                tmpDetallePO.FechaValor = Now
                tmpDetallePO.Valor = 0
                For Each A In tmpRequeDetalle.Entregas
                    Set entregaDetalle = New EntregaPetOfDetalle
                    entregaDetalle.Cantidad = A.Cantidad
                    entregaDetalle.FEcha = A.FEcha
                    tmpDetallePO.Entregas.Add entregaDetalle
                Next A

                If tmpRequeDetalle.estado = Anulado Then
                    tmpDetallePO.estado = EPOD_Anulado
                Else
                    tmpDetallePO.estado = EPOD_Activo
                End If

                tmpDetallePO.Cantidad = tmpRequeDetalle.Cantidad
                colPOItems.Add tmpDetallePO

            Next y
            po.detalle = colPOItems
            po.FechaEmision = funciones.dateFormateada(Now)
            po.FechaSolicitada = funciones.dateFormateada(Now)
            po.Proveedor = tmpProv
            po.usuarioCreador = DAOUsuarios.GetById(funciones.getUser)
            po.idReque = T.Id
            po.Modificado = Now

            strsql = "insert into ComprasPeticionOferta (FechaEmision, FechaPedido, UsuarioCreador, id_proveedor,id_reque,estado,modificado, moneda_id)  values  ('" & Format(po.FechaEmision, "yyyy-mm-yy") & "','" & Format(po.FechaSolicitada, "yyyy-mm-yy") & "'," & po.usuarioCreador.Id & "," & po.Proveedor.Id & "," & po.idReque & ",0,'" & po.Modificado & "', 0)"
            If Not conectar.execute(strsql) Then
                GoTo err1
            Else
                If conectar.UltimoId("ComprasPeticionOferta", Id_) Then po.numero = Id_ Else GoTo err1
                If Not DAOPeticionOfertaDetalle.Guardar(po) Then GoTo err1
            End If
        Next x

        conectar.CommitTransaction

        If Nueva Then DAORequeHistorial.agregar T, "REQUERIMIENTO EN PO"

    End If

    Exit Function

err1:
    Nueva = False
    T.estado = antestado
    conectar.RollBackTransaction
End Function

Public Function crearPO(Detalles As Collection) As Boolean

    Dim proveedoresYaProcesados As New Collection

    Dim pair As Collection
    Dim pair2 As Collection
    Dim petOf As clsPeticionOferta
    Dim provid As Long
    Dim prov As clsProveedor
    Dim petOfDet As clsPeticionOfertaDetalle
    Dim requeDeta As clsRequeMateriales
    Dim ent As clsRequeEntregas
    Dim q As String
    Dim Id_ As Long

    crearPO = True

    conectar.BeginTransaction

    For Each pair In Detalles
        Set requeDeta = pair(1)
        provid = pair(2)

        If Not funciones.BuscarEnColeccion(proveedoresYaProcesados, CStr(provid)) Then
            Set prov = DAOProveedor.FindById(provid)
            Set petOf = New clsPeticionOferta
            petOf.detalle = New Collection
            petOf.FechaEmision = Now
            petOf.FechaSolicitada = Now
            petOf.Proveedor = prov
            Set petOf.moneda = prov.moneda
            petOf.usuarioCreador = DAOUsuarios.GetById(funciones.getUser)
            petOf.idReque = requeDeta.RequeId
            petOf.Modificado = Now

            For Each pair2 In Detalles
                If pair2(2) = provid Then
                    Set requeDeta = pair2(1)
                    Set petOfDet = New clsPeticionOfertaDetalle
                    petOfDet.DetalleReque = requeDeta
                    petOfDet.FechaValor = Date
                    petOfDet.Valor = 0
                    petOfDet.estado = EPOD_EnEspera
                    For Each ent In requeDeta.Entregas
                        Set entregaDetalle = New EntregaPetOfDetalle
                        entregaDetalle.Cantidad = ent.Cantidad
                        entregaDetalle.FEcha = ent.FEcha
                        petOfDet.Entregas.Add entregaDetalle
                    Next ent

                    petOfDet.Cantidad = requeDeta.Cantidad

                    petOf.detalle.Add petOfDet
                End If
            Next pair2

            q = "insert into ComprasPeticionOferta (FechaEmision, FechaPedido, UsuarioCreador, id_proveedor,id_reque,estado,modificado, moneda_id)  values  ('" & Format(petOf.FechaEmision, "yyyy-mm-yy") & "','" & Format(petOf.FechaSolicitada, "yyyy-mm-yy") & "'," & petOf.usuarioCreador.Id & "," & petOf.Proveedor.Id & "," & petOf.idReque & ",0,'" & petOf.Modificado & "', 0)"
            If Not conectar.execute(q) Then
                GoTo err1
            Else
                If conectar.UltimoId("ComprasPeticionOferta", Id_) Then petOf.numero = Id_ Else GoTo err1
                If Not DAOPeticionOfertaDetalle.Guardar(petOf) Then GoTo err1
            End If

            proveedoresYaProcesados.Add provid, CStr(provid)
        End If

        requeDeta.estado = EnPOParcial_    'nopongo en po total, porque no se si estan todos los proveedores en po
        If Not DAORequeMateriales.aPO(requeDeta, requeDeta.RequeId) Then GoTo err1

    Next pair

    conectar.CommitTransaction
    DAOEvento.Publish petOf.numero, TEB_PeticionOfertaCreada

    Exit Function
err1:
    crearPO = False
    conectar.RollBackTransaction
End Function

Public Function CambiarEstado(po As clsPeticionOferta, estado As EstadoPO)
    Dim q As String
    q = "UPDATE ComprasPeticionOferta SET estado = " & estado & " WHERE id = " & po.numero
    CambiarEstado = conectar.execute(q)
    If CambiarEstado Then po.estado = estado
End Function


Public Sub ExportarExcel(POid As Long, cmd As CommonDialog)


    On Error GoTo E
    Dim filePath As String
    Dim vPeticion As clsPeticionOferta
    Set vPeticion = DAOPeticionOferta.GetById(POid)
    vPeticion.detalle = DAOPeticionOfertaDetalle.FindAll(POid)

    cmd.filter = "Excel|*.xls"
    cmd.filename = "PO " & vPeticion.numero & ".xls"
    cmd.CancelError = True
    cmd.ShowSave
    filePath = cmd.filename

    Dim xl As New Excel.Application
    Dim xlwbook As Excel.Workbook
    Dim xlsheet As Excel.Worksheet

    Set xlwbook = xl.Workbooks.Add()

    xlwbook.Worksheets(3).Delete
    xlwbook.Worksheets(2).Delete

    Set xlsheet = xlwbook.Worksheets(1)
    xlsheet.PageSetup.Orientation = xlLandscape

    Dim rowFin As Long

    Dim rowStart As Long

    rowStart = 8

    Dim ItemCount As Long

    With xlsheet
        .Range(.Cells(rowStart, 1), .Cells(rowStart + 3, 2)).rows.Font.Bold = True
        .Range(.Cells(rowStart, 2), .Cells(rowStart + 3, 3)).HorizontalAlignment = XlHAlign.xlHAlignLeft


        .Cells(rowStart, 2) = "PO Nº"
        .Cells(rowStart, 3) = vPeticion.numero
        rowStart = rowStart + 1

        .Cells(rowStart, 2) = "Proveedor"
        .Cells(rowStart, 3) = vPeticion.Proveedor.RazonSocial
        rowStart = rowStart + 1

        .Cells(rowStart, 2) = "Fecha Emisión"
        .Cells(rowStart, 3) = vPeticion.FechaEmision
        rowStart = rowStart + 1

        rowStart = 12

        Dim entrega As clsRequeEntregas
        Dim Entregas As Collection

        Dim tmpDeta As clsPeticionOfertaDetalle

        Dim rowEntrega As Long

        Dim rowInicio As Long

        .Cells(rowStart, 1) = "Item"
        .Cells(rowStart, 2) = "Cantidad"
        .Cells(rowStart, 3) = "UM"
        .Cells(rowStart, 4) = "Material"
        .Cells(rowStart, 5) = "Precio Unitario"
        .Cells(rowStart, 6) = "Precio Total"
        .Range(.Cells(rowStart, 1), .Cells(rowStart, 6)).Interior.ColorIndex = 15
        .Range(.Cells(rowStart, 1), .Cells(rowStart, 6)).rows.Font.Bold = True
        .Range(.Cells(rowStart, 1), .Cells(rowStart, 6)).BorderAround XlLineStyle.xlContinuous, xlThin, xlColorIndexAutomatic, 30


        .Cells(rowStart - 1, 7) = "Entregas"
        .Range(.Cells(rowStart - 1, 7), .Cells(rowStart - 1, 8)).MergeCells = True
        .Cells(rowStart - 1, 7).HorizontalAlignment = VtHorizontalAlignment.VtHorizontalAlignmentFill
        .Range(.Cells(rowStart - 1, 7), .Cells(rowStart - 1, 8)).Interior.ColorIndex = 15
        .Range(.Cells(rowStart - 1, 7), .Cells(rowStart - 1, 8)).rows.Font.Bold = True
        .Range(.Cells(rowStart - 1, 7), .Cells(rowStart - 1, 8)).BorderAround XlLineStyle.xlContinuous, xlThin, xlColorIndexAutomatic, 30


        'rowStart = rowStart + 1
        .Cells(rowStart, 7) = "Fecha"
        .Cells(rowStart, 7).HorizontalAlignment = VtHorizontalAlignment.VtHorizontalAlignmentFill
        .Cells(rowStart, 8) = "Cantidad"
        .Cells(rowStart, 8).HorizontalAlignment = VtHorizontalAlignment.VtHorizontalAlignmentFill
        .Range(.Cells(rowStart, 7), .Cells(rowStart, 8)).Interior.ColorIndex = 15
        .Range(.Cells(rowStart, 7), .Cells(rowStart, 8)).rows.Font.Bold = True
        .Range(.Cells(rowStart, 7), .Cells(rowStart, 8)).BorderAround XlLineStyle.xlContinuous, xlThin, xlColorIndexAutomatic, 30


        For Each tmpDeta In vPeticion.detalle
            ItemCount = ItemCount + 1

            rowInicio = rowStart

            rowStart = rowStart + 1

            .Cells(rowStart, 1) = ItemCount
            .Cells(rowStart, 2) = tmpDeta.Total
            .Cells(rowStart + 1, 2) = tmpDeta.DetalleReque.Cantidad
            .Cells(rowStart, 2).NumberFormat = "0.00"
            .Cells(rowStart + 1, 2).NumberFormat = "0.00"
            .Cells(rowStart, 3) = enums.enumUnidades(tmpDeta.DetalleReque.Material.UnidadCompra)
            .Cells(rowStart + 1, 3) = enums.enumUnidades(tmpDeta.DetalleReque.Material.UnidadPedido)

            .Cells(rowStart, 4) = tmpDeta.DetalleReque.Material.Grupo.rubros.rubro & "  " & tmpDeta.DetalleReque.Material.Grupo.Grupo & " | " & tmpDeta.DetalleReque.Material.descripcion
            .Cells(rowStart + 1, 4) = funciones.JoinCollectionValues(tmpDeta.DetalleReque.Material.Atributos, ", ")

            .Cells(rowStart, 5) = tmpDeta.Valor
            .Cells(rowStart, 5).NumberFormat = "0.00"
            '.Range(.Cells(rowStart, 5), .Cells(rowStart, 6)).Locked = False
            .Cells(rowStart, 6) = "=(E" & rowStart & "*B" & rowStart & ")"

            .Range(.Cells(rowStart, 5), .Cells(rowStart, 6)).NumberFormat = "0.00"


            rowStart = rowStart + 2

            .Cells(rowStart, 4) = tmpDeta.DetalleReque.observaciones
            .Range(.Cells(rowStart, 4), .Cells(rowStart, 6)).MergeCells = True



            rowStart = rowInicio
            rowEntrega = rowStart

            rowStart = rowStart + 1

            Set Entregas = DAORequeEntregas.GetEntregaById(tmpDeta.DetalleReque.Id, material_)
            For Each entrega In Entregas
                .Cells(rowStart, 7) = entrega.FEcha
                .Cells(rowStart, 8) = entrega.Cantidad & " " & enums.enumUnidades(tmpDeta.DetalleReque.Material.UnidadPedido)

                .Range(.Cells(rowStart, 7), .Cells(rowStart, 8)).BorderAround XlLineStyle.xlContinuous, xlThin, xlColorIndexAutomatic, 30
                .Range(.Cells(rowStart, 7), .Cells(rowStart, 8)).Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
                '.Range(.Cells(rowStart, 5), .Cells(rowStart, 6)).Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous

                rowFin = rowStart

                rowStart = rowStart + 1
            Next entrega


            '.Range(.Cells(rowEntrega, 5), .Cells(rowStart - 1, 6)).BorderAround XlLineStyle.xlContinuous, xlThin, xlColorIndexAutomatic, 30
            '.Range(.Cells(rowEntrega, 5), .Cells(rowStart - 1, 6)).Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            '.Range(.Cells(rowEntrega, 5), .Cells(rowStart - 1, 6)).Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous


            If rowStart - rowEntrega - 1 = 1 Then
                .Range(.Cells(rowEntrega + 1, 1), .Cells(rowStart, 6)).BorderAround XlLineStyle.xlContinuous, xlThin, xlColorIndexAutomatic, 30
                .Range(.Cells(rowEntrega + 1, 1), .Cells(rowStart, 6)).Interior.ColorIndex = 2
            Else

                .Range(.Cells(rowEntrega + 1, 1), .Cells(rowStart - 1, 6)).BorderAround XlLineStyle.xlContinuous, xlThin, xlColorIndexAutomatic, 30
                .Range(.Cells(rowEntrega + 1, 1), .Cells(rowStart - 1, 6)).Interior.ColorIndex = 2
            End If





            'rowStart = rowStart + 1
        Next tmpDeta

        rowStart = rowStart + 1

        .Range(.Cells(rowStart, 1), .Cells(rowStart + 4, 1)).rows.Font.Bold = True

        .Cells(rowStart, 5) = "Total"
        .Cells(rowStart, 6).NumberFormat = "0.00"
        .Range(.Cells(rowStart, 6), .Cells(rowStart, 6)).Formula = "=SUM(D6:D" & rowFin & ")"
        .Range(.Cells(rowStart, 5), .Cells(rowStart + 4, 5)).rows.Font.Bold = True

        rowStart = rowStart + 1
        .Cells(rowStart, 5) = "Moneda"
        .Cells(rowStart, 6) = vPeticion.moneda.NombreCorto
        '.Range(.Cells(rowStart, 6), .Cells(rowStart, 6)).Locked = False
        rowStart = rowStart + 1

        .Cells(rowStart, 5) = "Forma de pago"
        '.Range(.Cells(rowStart, 6), .Cells(rowStart, 6)).Locked = False
        rowStart = rowStart + 1
        .Cells(rowStart, 5) = "Descuento"
        '.Range(.Cells(rowStart, 6), .Cells(rowStart, 6)).Locked = False

        .Cells.EntireColumn.AutoFit
    End With


    Dim sBuffer As String
    Dim sTmpPicFile As String
    sTmpPicFile = App.path & "\logo.jpg"
    If LenB(Dir(sTmpPicFile)) > 0 Then Kill sTmpPicFile
    sBuffer = StrConv(LoadResData("logo", "CUSTOM"), vbUnicode)
    Open sTmpPicFile For Output As #1
    Print #1, sBuffer
    Close #1

    Dim img
    Set img = xlsheet.Pictures.insert(sTmpPicFile)
    img.Width = 219
    img.Height = 78
    img.Top = 0
    img.Left = 0

    Kill sTmpPicFile

    xlwbook.Unprotect
    xlsheet.Unprotect

    xlsheet.Protect
    xlwbook.Protect

    xlwbook.SaveAs filePath
    xlwbook.Close False
    xl.Quit

    ShellExecute -1, "open", filePath, vbNullString, vbNullString, 0

    Set xlsheet = Nothing
    Set xlwbook = Nothing
    Set xl = Nothing

    Exit Sub
E:
    If Err.Source <> "CommonDialog" And Err.Number <> 32755 Then MsgBox Err.Description, vbCritical
End Sub
