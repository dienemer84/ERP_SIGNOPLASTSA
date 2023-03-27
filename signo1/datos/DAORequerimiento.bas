Attribute VB_Name = "DAORequerimiento"
Option Explicit

Public Function FindById(id_reque As Long, Optional fetcSector As Boolean = False, _
                         Optional fetchUsuarioAprobador As Boolean = False, _
                         Optional fetchUsuarioCreador As Boolean = False, _
                         Optional fetchMateriales As Boolean = False) As clsRequerimiento
    Dim col As Collection
    Set col = FindAll("id = " & id_reque, fetcSector, fetchUsuarioAprobador, fetchUsuarioCreador, fetchMateriales)
    If col.count = 0 Then
        Set FindById = Nothing
    Else
        Set FindById = col.item(1)
    End If
End Function
Public Function FindAll(Optional ByVal filter As String = vbNullString, _
                        Optional fetcSector As Boolean = False, _
                        Optional fetchUsuarioAprobador As Boolean = False, _
                        Optional fetchUsuarioCreador As Boolean = False, _
                        Optional fetchMateriales As Boolean = False _
                      ) As Collection

    On Error GoTo err1
    Dim col As New Collection
    Dim reque As clsRequerimiento
    Dim q As String

    q = "select * from ComprasRequerimientos where 1 = 1"
    If LenB(filter) > 0 Then q = q & " and " & filter

    Dim rs As New Recordset
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF
        Set reque = New clsRequerimiento
        reque.Id = rs!Id
        reque.estado = rs!estado
        reque.fechaCreado = rs!fechaCreado
        If fetcSector Then reque.Sector = DAOSectores.GetById(rs!idSector)         '''''''''''''''''
        If fetchUsuarioAprobador Then reque.Usuario_aprobador = DAOUsuarios.GetById(rs!idUsuarioAprobador)    '''''''''''''''''
        If fetchUsuarioCreador Then reque.Usuario_creador = DAOUsuarios.GetById(rs!idUsuarioCreador)    '''''''''''''''''
        reque.DestinoOT = rs!idpedido
        reque.Tipo = rs!destino
        If fetchMateriales Then reque.Materiales = DAORequeMateriales.GetByReque(rs!Id)          '''''''''''''''''
        reque.Guardado = rs!Guardado
        '   reque.Conceptos = DAORequeConceptos.GetByIdDetalleReque(rs!id)
        col.Add reque, CStr(reque.Id)
        rs.MoveNext
    Wend

    Set FindAll = col
    Exit Function
err1:
    Set FindAll = New Collection
End Function

Public Function Save(T As clsRequerimiento) As Boolean

    Save = Guardar(T)
    If Not Save Then GoTo err1

    Exit Function
err1:
    Save = False

End Function

Public Function Guardar(T As clsRequerimiento) As Boolean
    Dim Id_ As Long
    Dim antGuardado As Date
    Dim nuevo As Boolean
    On Error GoTo err1
    Dim req As clsRequerimiento
    Guardar = True
    'conectar.BeginTransaction
    antGuardado = T.Guardado
    nuevo = True

    If T.Id > 0 Then
        nuevo = False

        Set req = FindById(T.Id, True, True, True, True)

        Dim usuApro As Long
        If req.Guardado = T.Guardado Then
            T.Guardado = Now
            If T.Usuario_aprobador Is Nothing Then usuApro = 0 Else usuApro = T.Usuario_aprobador.Id
            Guardar = conectar.execute("update ComprasRequerimientos set idSector=" & T.Sector.Id & ",idPedido=" & T.DestinoOT & ", guardado=" & conectar.Escape(T.Guardado) & ",destino=" & T.Tipo & ",estado=" & T.estado & ",idUsuarioAprobador=" & usuApro & " where id=" & T.Id)
        Else
            MsgBox "El requerimiento ya fue guardado con anterioridad en otra sesión!", vbInformation, "Información"
            Guardar = False
            'conectar.RollBackTransaction
            Exit Function
        End If

    Else

        T.Guardado = Now
        Guardar = conectar.execute("insert into ComprasRequerimientos (idSector, estado, idPedido, FechaCreado, idUsuarioCreador, idUsuarioAprobador, destino,guardado)   values  (" & T.Sector.Id & "," & T.estado & "," & T.DestinoOT & ",'" & funciones.dateFormateada(T.fechaCreado) & "'," & GetEntityId(T.Usuario_creador) & "," & GetEntityId(T.Usuario_aprobador) & "," & T.Tipo & ",'" & funciones.datetimeFormateada(T.Guardado) & "')")
        conectar.UltimoId "ComprasRequerimientos", Id_
        T.Id = Id_
    End If

    If Not DAORequeMateriales.Save(T) Then
        GoTo err1
    End If

    DAORequeHistorial.agregar T, "REQUERIMIENTO GUARDADO"
    If Not Guardar Then GoTo err1
    Set req = Nothing
    Exit Function


err1:
    If nuevo Then T.Id = 0
    Set req = Nothing
    T.Guardado = antGuardado
    Guardar = False


End Function

Public Function aprobar(T As clsRequerimiento) As Boolean
    On Error GoTo err1
    Dim antUsuario As clsUsuario
    Dim antestado As EstadoRequeCompra

    Dim deta As clsRequeMateriales
    For Each deta In T.Materiales
        If deta.estado <> EstadoRequeCompra.Finalizado_ And deta.estado <> EstadoRequeCompra.Anulado Then
            aprobar = False
            Exit Function
        End If
    Next deta

    For Each deta In T.Materiales

        If deta.estado = EstadoRequeCompra.Finalizado_ Then
            deta.estado = Aprobado_
        End If
    Next deta


    If T.estado = EstadoRequeCompra.Finalizado_ Then
        antestado = T.estado
        Set antUsuario = T.Usuario_aprobador

        T.estado = Aprobado_
        T.Usuario_aprobador = DAOUsuarios.GetById(funciones.getUser)
        aprobar = Save(T)
        If aprobar Then
            DAORequeHistorial.agregar T, "REQUERIMIENTO APROBADO"
            DAOEvento.Publish T.Id, TEB_RequerimientoCompraAprobado
        End If
    End If
    Exit Function
err1:
    aprobar = False
    T.estado = antestado
    Set T.Usuario_aprobador = antUsuario
End Function
Public Function finalizar(T As clsRequerimiento) As Boolean
    On Error GoTo err1
    Dim antestado As EstadoRequeCompra

    If T.estado = EnEdición_ Then
        antestado = T.estado
        T.estado = EstadoRequeCompra.Finalizado_
        Dim det As clsRequeMateriales
        For Each det In T.Materiales
            det.estado = EstadoRequeCompra.Finalizado_
        Next det

        finalizar = Save(T)
        If finalizar Then
            DAORequeHistorial.agregar T, "REQUERIMIENTO FINALIZADO"
            DAOEvento.Publish T.Id, TEB_RequerimientoCompraFinalizado
        End If
    End If

    Exit Function
err1:
    finalizar = False
    T.estado = antestado

    For Each det In T.Materiales
        det.estado = EnEdición_
    Next det
End Function


Public Function procesar(T As clsRequerimiento) As Boolean
    On Error GoTo err1
    Dim tmpMat As clsRequeMateriales
    Dim antestado As EstadoRequeCompra

    If T.estado = Aprobado_ Then
        antestado = T.estado

        Dim deta As clsRequeMateriales


        For Each deta In T.Materiales




            If deta.estado <> Aprobado_ And deta.estado <> Anulado Then
                procesar = False
                Exit Function
            End If



        Next deta


        For Each deta In T.Materiales

            If deta.estado = Aprobado_ Then
                deta.estado = EnProceso_
            End If
        Next deta



        T.estado = EnProceso_
        Dim i As Long
        For i = 1 To T.Materiales.count
            Set tmpMat = T.Materiales.item(i)
            T.Materiales.item(i).ListaProveedores = DAOProveedor.FindAllByRubro(tmpMat.Material.Grupo.rubros.Id)      'DAORequeProveedores.GetByRubro(T.Materiales.Item(I).id, material_)
        Next i

        procesar = Save(T)
        If procesar Then
            DAORequeHistorial.agregar T, "REQUERIMIENTO EN PROCESO"
        Else
            GoTo err1
        End If
    End If
    Exit Function
err1:
    procesar = False
    T.estado = antestado

    For Each deta In T.Materiales
        deta.estado = Aprobado_
    Next deta
End Function



Public Function FinProceso(T As clsRequerimiento) As Boolean
    On Error GoTo err1

    Dim antestado As EstadoRequeCompra
    If T.estado = EnProceso_ Then
        antestado = T.estado


        Dim deta As clsRequeMateriales
        For Each deta In T.Materiales
            If deta.estado <> EnProceso_ And deta.estado <> Anulado Then
                FinProceso = False
                Exit Function
            End If
        Next deta

        For Each deta In T.Materiales
            If deta.estado = EnProceso_ Then
                deta.estado = EstadoRequeCompra.Procesado_
            End If
        Next deta


        T.estado = EstadoRequeCompra.Procesado_

        FinProceso = Save(T)
        If FinProceso Then
            DAORequeHistorial.agregar T, "REQUERIMIENTO PROCESADO"
        Else
            GoTo err1
        End If
    End If
    Exit Function
err1:
    FinProceso = False
    T.estado = antestado
End Function


Public Function Anular(T As clsRequerimiento) As Boolean
    On Error GoTo err1
    conectar.BeginTransaction
    Dim antestado As EstadoRequeCompra
    'If T.estado = EnProceso_ Then
    antestado = T.estado


    Dim deta As clsRequeMateriales
    '        For Each deta In T.Materiales
    '            If deta.estado <> EnProceso_ Then
    '                FinProceso = False
    '                Exit Function
    '            End If
    '        Next deta
    '
    For Each deta In T.Materiales
        deta.estado = EstadoRequeCompra.Anulado
    Next deta


    T.estado = EstadoRequeCompra.Anulado

    Anular = Save(T)

    'anular tambien las PO referentes a este REQ
    Dim pos As New Collection
    Dim po As clsPeticionOferta
    Dim pod As clsPeticionOfertaDetalle
    Set pos = DAOPeticionOferta.GetAll("and id_reque=" & T.Id)
    For Each po In pos
        For Each pod In po.detalle
            pod.estado = EPOD_Anulado

            DAOPeticionOfertaDetalle.Update pod, po
        Next pod
    Next po




    If Anular Then
        DAORequeHistorial.agregar T, "REQUERIMIENTO ANULADO"
        DAOEvento.Publish T.Id, TEB_RequerimientoCompraAnulado
    Else
        GoTo err1
    End If
    'End If

    conectar.CommitTransaction
    Exit Function
err1:
    Anular = False
    T.estado = antestado
    conectar.RollBackTransaction
End Function

Public Sub ExportExcel(reqId As Long, withOCProveedoresColumn As Boolean)
    Dim req As clsRequerimiento
    Set req = DAORequerimiento.FindById(reqId, True, True, True, True)
    Dim MAT As clsRequeMateriales
    Dim ent As clsRequeEntregas
    Dim prov As clsProveedor

    Dim xlWorkbook As New Excel.Workbook
    Dim xlWorksheet As New Excel.Worksheet
    Dim xlApplication As New Excel.Application

    Set xlWorkbook = xlApplication.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets.item(1)

    xlWorksheet.Activate

    xlWorksheet.Cells(1, 1).value = "Numero"
    xlWorksheet.Cells(1, 2).value = reqId
    xlWorksheet.Cells(2, 1).value = "Sector Solicitante"
    xlWorksheet.Cells(2, 2).value = req.Sector.Sector
    xlWorksheet.Cells(3, 1).value = "Fecha Creación"
    xlWorksheet.Cells(3, 2).value = req.fechaCreado
    xlWorksheet.Cells(4, 1).value = "Destino"
    xlWorksheet.Cells(4, 2).value = req.StringDestino



    xlWorksheet.Cells(5, 1).value = "Usuario"
    xlWorksheet.Cells(5, 2).value = req.Usuario_creador.usuario
    EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(1, 1), xlWorksheet.Cells(5, 2))
    xlWorksheet.Range(xlWorksheet.Cells(1, 1), xlWorksheet.Cells(5, 1)).Interior.Color = Information.RGB(215, 215, 215)
    xlWorksheet.Range(xlWorksheet.Cells(1, 1), xlWorksheet.Cells(5, 1)).Font.Bold = True

    xlWorksheet.Range(xlWorksheet.Cells(2, 4), xlWorksheet.Cells(5, 4)).Merge
    xlWorksheet.Cells(2, 4).value = "REQUERIMIENTO INTERNO DE MATERIALES"
    Dim Ot As OrdenTrabajo
    Set Ot = DAOOrdenTrabajo.FindById(req.DestinoOT)
    xlWorksheet.Cells(6, 4).value = req.DestinoOT    ' Ot.Cliente.razon
    xlWorksheet.Cells(2, 4).HorizontalAlignment = xlCenter
    xlWorksheet.Cells(2, 4).Font.Bold = True
    xlWorksheet.Cells(2, 4).Font.Size = 15
    xlWorksheet.Cells(2, 4).VerticalAlignment = xlCenter
    xlWorksheet.Cells(6, 4).HorizontalAlignment = xlCenter
    xlWorksheet.Cells(6, 4).Font.Bold = True
    xlWorksheet.Cells(6, 4).Font.Size = 15
    xlWorksheet.Cells(6, 4).VerticalAlignment = xlCenter
    EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(2, 4), xlWorksheet.Cells(5, 4))




    Dim rowStart As Long
    rowStart = 8

    'armo cabeceras
    xlWorksheet.Cells(rowStart, 1).value = "Cod Material"
    xlWorksheet.Cells(rowStart + 1, 1).value = "Estado"

    xlWorksheet.Cells(rowStart, 2).value = "Cantidad"
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 2), xlWorksheet.Cells(rowStart + 1, 2)).Merge
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 2), xlWorksheet.Cells(rowStart, 2)).VerticalAlignment = xlCenter

    xlWorksheet.Cells(rowStart, 3).value = "UM"
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 3), xlWorksheet.Cells(rowStart + 1, 3)).Merge
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 3), xlWorksheet.Cells(rowStart, 3)).VerticalAlignment = xlCenter

    xlWorksheet.Cells(rowStart, 4).value = "Material"
    xlWorksheet.Cells(rowStart + 1, 4).value = "Atributos / Observaciones"

    ' xlWorksheet.Cells(rowStart, 5).value = "Detalle"
    ' xlWorksheet.Range(xlWorksheet.Cells(rowStart, 5), xlWorksheet.Cells(rowStart + 1, 5)).Merge
    'xlWorksheet.Range(xlWorksheet.Cells(rowStart, 5), xlWorksheet.Cells(rowStart, 5)).VerticalAlignment = xlCenter

    xlWorksheet.Cells(rowStart, 5).value = "Entregas"
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 5), xlWorksheet.Cells(rowStart, 6)).Merge
    xlWorksheet.Cells(rowStart + 1, 5).value = "Fecha"
    xlWorksheet.Cells(rowStart + 1, 6).value = "Cantidad"


    Dim colIdx As Long
    Dim proveedoresIdx As New Collection
    If withOCProveedoresColumn Then
        xlWorksheet.Cells(rowStart, 7).value = "OC"
        xlWorksheet.Range(xlWorksheet.Cells(rowStart, 7), xlWorksheet.Cells(rowStart + 1, 7)).Merge
        xlWorksheet.Range(xlWorksheet.Cells(rowStart, 7), xlWorksheet.Cells(rowStart + 1, 7)).VerticalAlignment = xlCenter

        colIdx = 8
        xlWorksheet.Cells(rowStart, colIdx).value = "Proveedores"


        Dim colFix As New Collection


        For Each MAT In req.Materiales
            For Each prov In MAT.ListaProveedores
                If Not funciones.BuscarEnColeccion(proveedoresIdx, CStr(prov.Id)) Then
                    proveedoresIdx.Add colIdx, CStr(prov.Id)
                    xlWorksheet.Cells(rowStart + 1, colIdx).value = truncar(prov.RazonSocial, 13)
                    xlWorksheet.Cells(rowStart + 1, colIdx).Font.Size = 6


                    colFix.Add colIdx



                    colIdx = colIdx + 1
                End If
            Next prov
        Next MAT
        colIdx = colIdx - 1

        xlWorksheet.Range(xlWorksheet.Cells(rowStart, 8), xlWorksheet.Cells(rowStart, colIdx)).Merge
    Else
        colIdx = 7
    End If

    EncuadrarCelda xlWorksheet.Range(xlWorksheet.Cells(rowStart, 1), xlWorksheet.Cells(rowStart + 1, colIdx))
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 1), xlWorksheet.Cells(rowStart + 1, colIdx)).Interior.Color = Information.RGB(215, 215, 215)
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 1), xlWorksheet.Cells(rowStart + 1, colIdx)).Font.Bold = True
    xlWorksheet.Range(xlWorksheet.Cells(rowStart, 1), xlWorksheet.Cells(rowStart + 1, colIdx)).HorizontalAlignment = xlCenter


    'populeo  carolina populeo
    Dim tmpValue As Long
    rowStart = rowStart + 2
    Dim poDetalle As clsPeticionOfertaDetalle
    Dim colPoDet As Collection
    Dim B As Long
    For Each MAT In req.Materiales
        xlWorksheet.Cells(rowStart, 1).value = MAT.Material.codigo
        xlWorksheet.Cells(rowStart + 1, 1).value = enums.enumEstadoRequeCompra(MAT.estado)

        xlWorksheet.Cells(rowStart, 2).value = MAT.TotalCantidad(MAT.Material.UnidadPedido)    ' MAT.Cantidad
        xlWorksheet.Cells(rowStart + 1, 2).value = MAT.TotalCantidad(MAT.Material.UnidadCompra)
        xlWorksheet.Cells(rowStart + 2, 2).value = MAT.TotalCantidad(MAT.Material.unidad)

        xlWorksheet.Cells(rowStart, 2).NumberFormat = "0.00"
        xlWorksheet.Cells(rowStart + 1, 2).NumberFormat = "0.00"
        xlWorksheet.Cells(rowStart + 2, 2).NumberFormat = "0.00"

        xlWorksheet.Cells(rowStart, 3).value = enums.enumUnidades(MAT.Material.UnidadPedido)
        xlWorksheet.Cells(rowStart + 1, 3).value = enums.enumUnidades(MAT.Material.UnidadCompra)
        xlWorksheet.Cells(rowStart + 2, 3).value = enums.enumUnidades(MAT.Material.unidad)

        xlWorksheet.Cells(rowStart, 4).value = MAT.Material.Grupo.rubros.rubro & " " & MAT.Material.Grupo.Grupo & " " & MAT.Material.descripcion
        xlWorksheet.Cells(rowStart + 1, 4).value = funciones.JoinCollectionValues(MAT.Material.Atributos, ", ")

        If Len(MAT.observaciones) > 0 Then xlWorksheet.Cells(rowStart + 2, 4).value = "Observaciones: " & MAT.observaciones

        If withOCProveedoresColumn Then
            Set colPoDet = DAOPeticionOfertaDetalle.FindAll(, "id_detalle_reque = " & MAT.Id)
            'xlWorksheet.Range(xlWorksheet.Cells(rowStart, 8), xlWorksheet.Cells(rowStart, colIdx)).Interior.Color = Information.RGB(215, 215, 215)

            If MAT.Entregas.count > 3 Then
                B = rowStart + MAT.Entregas.count
            Else
                B = rowStart + 2
            End If

            xlWorksheet.Range(xlWorksheet.Cells(rowStart, 8), xlWorksheet.Cells(B, colIdx)).Interior.Color = Information.RGB(215, 215, 215)




            If colPoDet.count > 0 Then


                For Each poDetalle In colPoDet
                    If MAT.estado <> Anulado And MAT.estado <> AnuladoParcial Then
                        xlWorksheet.Cells(rowStart, proveedoresIdx(CStr(poDetalle.ProveedorId))).value = poDetalle.moneda.NombreCorto & " " & funciones.FormatearDecimales(poDetalle.Valor)
                        xlWorksheet.Cells(rowStart, proveedoresIdx(CStr(poDetalle.ProveedorId))).HorizontalAlignment = xlRight
                        xlWorksheet.Range(xlWorksheet.Cells(rowStart, proveedoresIdx(CStr(poDetalle.ProveedorId))), xlWorksheet.Cells(rowStart + 1, proveedoresIdx(CStr(poDetalle.ProveedorId)))).Interior.Color = vbWhite
                        If MAT.Entregas.count > 3 Then
                            B = rowStart + MAT.Entregas.count
                        Else
                            B = rowStart + 3
                        End If


                        xlWorksheet.Range(xlWorksheet.Cells(rowStart + 1, proveedoresIdx(CStr(poDetalle.ProveedorId))), xlWorksheet.Cells(B - 1, proveedoresIdx(CStr(poDetalle.ProveedorId)))).Merge



                    End If
                Next poDetalle
            End If
        End If

        tmpValue = rowStart
        Dim A As Long
        A = rowStart
        For Each ent In MAT.Entregas
            xlWorksheet.Cells(tmpValue, 5).value = ent.FEcha
            xlWorksheet.Cells(tmpValue, 6).value = ent.Cantidad
            xlWorksheet.Cells(tmpValue, 6).NumberFormat = "0.00"
            tmpValue = tmpValue + 1
        Next ent


        tmpValue = rowStart

        If MAT.Entregas.count > 3 Then
            rowStart = rowStart + MAT.Entregas.count
        Else
            rowStart = rowStart + 3
        End If


        'xlWorksheet.Range(xlWorksheet.Cells(a + 2, colIdx - proveedoresIdx.count), xlWorksheet.Cells(tmpValue + 1, colIdx)).Merge


        'recuadrado item
        With xlWorksheet.Range(xlWorksheet.Cells(tmpValue, 1), xlWorksheet.Cells(rowStart - 1, colIdx))
            .Borders.item(xlEdgeBottom).Weight = xlThin
            .Borders.item(xlEdgeLeft).Weight = xlThin
            .Borders.item(xlEdgeRight).Weight = xlThin
            .Borders.item(xlInsideVertical).Weight = xlThin
            .Borders.item(xlInsideHorizontal).Weight = xlThin
        End With




        'linea gruesa
        xlWorksheet.Range(xlWorksheet.Cells(rowStart - 1, 1), xlWorksheet.Cells(rowStart - 1, colIdx)).Borders.item(xlEdgeBottom).Weight = xlMedium
    Next MAT




    'autosize
    xlApplication.ScreenUpdating = False
    Dim wkSt As String
    wkSt = xlWorksheet.Name
    xlWorksheet.Cells.EntireColumn.AutoFit
    xlWorkbook.Sheets(wkSt).Select
    xlApplication.ScreenUpdating = True

    'xlWorksheet.PageSetup.PrintTitleRows = "$1:$3" 'para que al imprimir queden las columnas fijas
    xlWorksheet.PageSetup.Orientation = xlLandscape
    xlWorksheet.PageSetup.BottomMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.TopMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.LeftMargin = xlApplication.CentimetersToPoints(1)
    xlWorksheet.PageSetup.RightMargin = xlApplication.CentimetersToPoints(1)


    Dim d
    For Each d In colFix
        xlWorksheet.Columns(d).ColumnWidth = 10
    Next


    'ajusto a 1 pagina
    xlWorksheet.PageSetup.Zoom = False
    xlWorksheet.PageSetup.FitToPagesWide = 1
    xlWorksheet.PageSetup.FitToPagesTall = 1

    Dim filename As String
    filename = funciones.GetTmpPath() & "Requerimiento Materiales Nº " & req.Id & ".xls"

    If Dir(filename) <> vbNullString Then Kill filename

    xlWorkbook.SaveAs filename

    xlWorkbook.Saved = True
    xlWorkbook.Close
    xlApplication.Quit


    funciones.ShellExecute 0, "open", filename, "", "", 0

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApplication = Nothing

End Sub

Private Sub EncuadrarCelda(ByRef rango As Range)
    rango.Borders.item(xlEdgeTop).Weight = xlThin
    rango.Borders.item(xlEdgeBottom).Weight = xlThin
    rango.Borders.item(xlEdgeLeft).Weight = xlThin
    rango.Borders.item(xlEdgeRight).Weight = xlThin
    rango.Borders.item(xlInsideVertical).Weight = xlThin
    rango.Borders.item(xlInsideHorizontal).Weight = xlThin

End Sub
