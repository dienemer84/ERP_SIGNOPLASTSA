Attribute VB_Name = "DAODocumentos"
Option Explicit

Public Enum TipoDocumentoImpresion
    TDI_Factura = 100
    TDI_Cheque = 200
    TDI_Recibo = 300
    TDI_Remito = 400


End Enum



Public Function llenarComboXtremeSuite(cbo As Xtremesuitecontrols.ComboBox)
    Dim col As Collection
    Dim Doc As documento
    Set col = DAODocumentos.FindAll
    cbo.Clear



    For Each Doc In col
        cbo.AddItem Doc.nombre
        cbo.ItemData(cbo.NewIndex) = Doc.Id
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Function

Public Function FindById(Id As Long) As documento
    Set FindById = FindAll(True, "Documentos.id=" & Id)(1)
End Function

Public Function SaveDocumento(d As documento, Optional SaveWithDetalles As Boolean) As Boolean
    On Error GoTo err1
    Dim o As New clsEventoObserver
    conectar.BeginTransaction
    SaveDocumento = True
    Dim q As String
    If d.Id = 0 Then
        q = "INSERT INTO sp.Documentos    ( Nombre,    Alto,    Ancho,    id_archivo,    Activo ,tipo_documento   )    Values" _
          & " ('Nombre',    'Alto',    'Ancho',    'id_archivo',    'Activo' ,'tipo_documento'   )"
        o.EVENTO = agregar_
    Else
        q = "Update sp.Documentos     SET     id = 'id' ,     Nombre = 'Nombre' ,     Alto = 'Alto' ,     Ancho = 'Ancho' ,     id_archivo = 'id_archivo' , " _
          & " Activo = 'Activo', tipo_documento ='tipo_documento'         Where    id = 'id' "
        o.EVENTO = modificar_
    End If


    q = Replace(q, "'id'", conectar.Escape(d.Id))
    q = Replace(q, "'Nombre'", conectar.Escape(d.nombre))
    q = Replace(q, "'Alto'", conectar.Escape(d.Alto))
    q = Replace(q, "'Ancho'", conectar.Escape(d.Ancho))
    q = Replace(q, "'id_archivo'", conectar.Escape(d.Imagen))
    q = Replace(q, "'Activo'", conectar.Escape(d.estado))
    q = Replace(q, "'id_archivo'", conectar.Escape(d.Id))
    q = Replace(q, "'tipo_documento'", conectar.Escape(d.TipoDocumento))

    If Not conectar.execute(q) Then GoTo err1

    If SaveWithDetalles Then
        Dim deta As DocumentoDetalle

        If d.Id > 0 Then
            If Not conectar.execute("delete from DocumentoDetalles where id_documento=" & d.Id) Then GoTo err1
        End If

        For Each deta In d.Detalles

            deta.Id = 0
            If Not DAODocumentos.SaveDetalles(deta) Then GoTo err1

        Next


    End If

    d.Id = conectar.UltimoId2


    Set o.Elemento = d
    o.Tipo = Documentos_
    Channel.Notificar o, Documentos_

    conectar.CommitTransaction
    Exit Function
err1:
    conectar.RollBackTransaction
    SaveDocumento = False

End Function



Private Function SaveDetalles(det As DocumentoDetalle) As Boolean
    On Error GoTo err1


    Dim q As String
    SaveDetalles = True
    If det.Id = 0 Then

        q = "INSERT INTO sp.DocumentoDetalles " _
          & " ( id_documento,   pos_x,   pos_y,   alto,   ancho,   fijo,   alineacion,   negrita,   cursiva,   tachado,   subrayado,   nombre_fuente,    tamanio,   Tag   ) " _
          & " Values " _
          & " ('id_documento',    'pos_x',    'pos_y',    'alto',    'ancho',    'fijo',    'alineacion',    'negrita',    'cursiva',     'tachado',    'subrayado',    'nombre_fuente'," _
          & " 'tamanio',     'tag'    )"


    Else

        q = "  UPDATE sp.DocumentoDetalles     SET " _
          & " id = 'id' ,     id_documento = 'id_documento' ,    pos_x = 'pos_x' ," _
          & " pos_y = 'pos_y' ,     alto = 'alto' ,    ancho = 'ancho' ,     fijo = 'fijo' , " _
          & " alineacion = 'alineacion' ,     negrita = 'negrita' ,    cursiva = 'cursiva' , " _
          & " tachado = 'tachado' ,     subrayado = 'subrayado' ,     nombre_fuente = 'nombre_fuente' , " _
          & " tamanio = 'tamanio' ,     tag = 'tag'     Where     id = 'id'  "


    End If


    q = Replace(q, "'id'", conectar.Escape(det.Id))
    q = Replace(q, "'id_documento'", conectar.Escape(det.documento.Id))
    q = Replace(q, "'pos_x'", conectar.Escape(det.PosX))
    q = Replace(q, "'pos_y'", conectar.Escape(det.PosY))
    q = Replace(q, "'alto'", conectar.Escape(det.Alto))
    q = Replace(q, "'ancho'", conectar.Escape(det.Ancho))
    q = Replace(q, "'fijo'", conectar.Escape(det.Fijo))
    q = Replace(q, "'negrita'", conectar.Escape(det.Negrita))
    q = Replace(q, "'alineacion'", conectar.Escape(det.Alineacion))
    q = Replace(q, "'cursiva'", conectar.Escape(det.Cursiva))
    q = Replace(q, "'tachado'", conectar.Escape(det.Tachado))
    q = Replace(q, "'subrayado'", conectar.Escape(det.Subrayado))
    q = Replace(q, "'nombre_fuente'", conectar.Escape(det.nombreFuente))
    q = Replace(q, "'tamanio'", conectar.Escape(det.Tamano))
    q = Replace(q, "'tag'", conectar.Escape(det.Tag))


    If Not conectar.execute(q) Then GoTo err1



    det.Id = conectar.UltimoId2

    Exit Function
err1:
    SaveDetalles = False

End Function


Public Function FindAll(Optional IncluyeDetalles As Boolean = False, Optional filtro As String) As Collection
    Dim q As String
    q = "SELECT * from Documentos left join archivos on Documentos.id_archivo=archivos.id where 1=1"
    Dim col As New Collection
    Dim d As documento
    Dim idx As Dictionary
    Dim deta As DocumentoDetalle
    Dim rs As Recordset


    If LenB(filtro) > 0 Then
        q = q & " and " & filtro
    End If
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, idx
    While Not rs.EOF And Not rs.BOF
        Set d = Map(rs, idx, "Documentos")
        If IncluyeDetalles Then Set d.Detalles = MapDetalles(d.Id, d)
        col.Add d

        rs.MoveNext
    Wend
    Set FindAll = col

End Function



Public Function Map(rs As Recordset, indice As Dictionary, tabla As String) As documento
    Dim d As documento
    Dim Id As Long: Id = GetValue(rs, indice, tabla, "id")
    If Id > 0 Then
        Set d = New documento
        d.Id = Id
        d.Alto = GetValue(rs, indice, tabla, "Alto")
        d.Ancho = GetValue(rs, indice, tabla, "Ancho")
        d.estado = GetValue(rs, indice, tabla, "Activo")
        d.TipoDocumento = GetValue(rs, indice, tabla, "tipo_documento")
        d.Imagen = GetValue(rs, indice, tabla, "id_archivo")
        d.nombre = GetValue(rs, indice, tabla, "Nombre")
    End If

    Set Map = d
End Function

Private Function MapDetalles(idDocumento As Long, d As documento) As Collection
    Dim q As String
    Dim col As New Collection
    Dim rs As Recordset
    Dim idx As Dictionary
    Dim deta As DocumentoDetalle
    q = "Select * from DocumentoDetalles where id_documento=" & idDocumento
    Set rs = conectar.RSFactory(q)
    conectar.BuildFieldsIndex rs, idx
    While Not rs.EOF And Not rs.BOF

        Set deta = New DocumentoDetalle
        deta.Id = GetValue(rs, idx, "DocumentoDetalles", "id")
        deta.Alineacion = GetValue(rs, idx, "DocumentoDetalles", "alineacion")
        deta.Ancho = GetValue(rs, idx, "DocumentoDetalles", "ancho")
        deta.Alto = GetValue(rs, idx, "DocumentoDetalles", "alto")
        deta.Cursiva = GetValue(rs, idx, "DocumentoDetalles", "cursiva")
        deta.Fijo = GetValue(rs, idx, "DocumentoDetalles", "fijo")
        deta.Negrita = GetValue(rs, idx, "DocumentoDetalles", "negrita")
        deta.nombreFuente = GetValue(rs, idx, "DocumentoDetalles", "nombre_fuente")
        deta.PosX = GetValue(rs, idx, "DocumentoDetalles", "pos_x")
        deta.PosY = GetValue(rs, idx, "DocumentoDetalles", "pos_y")
        deta.Subrayado = GetValue(rs, idx, "DocumentoDetalles", "subrayado")
        deta.Tachado = GetValue(rs, idx, "DocumentoDetalles", "tachado")
        deta.Tamano = GetValue(rs, idx, "DocumentoDetalles", "tamanio")
        deta.Tag = GetValue(rs, idx, "DocumentoDetalles", "tag")
        Set deta.documento = d
        col.Add deta

        rs.MoveNext
    Wend

    Set MapDetalles = col

End Function




Public Function GetFieldsByTipo(TipoDoc As TipoDocumentoImpresion) As Collection
    Dim dicFields As New Collection
    Dim dto As DTOCampoBD



    If TipoDoc = TDI_Cheque Then

        Set dto = New DTOCampoBD
        dto.CampoEnBD = DAOCheques.CAMPO_FECHA_VENCIMIENTO
        dto.NombreCampo = "Fecha Vencimiento"
        dicFields.Add dto

        Set dto = New DTOCampoBD
        dto.NombreCampo = "Banco"
        dto.CampoEnBD = DAOCheques.CAMPO_ID_BANCO
        dicFields.Add dto

        Set dto = New DTOCampoBD
        dto.NombreCampo = "Moneda"
        dto.CampoEnBD = DAOCheques.CAMPO_ID_MONEDA
        dicFields.Add dto

        Set dto = New DTOCampoBD
        dto.NombreCampo = "Monto"
        dto.CampoEnBD = DAOCheques.CAMPO_MONTO
        dicFields.Add dto

        Set dto = New DTOCampoBD
        dto.NombreCampo = "Número"
        dto.CampoEnBD = DAOCheques.CAMPO_NUMERO
        dicFields.Add dto

        Set dto = New DTOCampoBD
        dto.NombreCampo = "Observaciones"
        dto.CampoEnBD = DAOCheques.CAMPO_OBSERVACIONES
        dicFields.Add dto

        Set dto = New DTOCampoBD
        dto.NombreCampo = "Origen-Destino"
        dto.CampoEnBD = DAOCheques.CAMPO_ORIGEN

        Set dto = New DTOCampoBD
        dto.NombreCampo = "Monto en letras"
        dto.CampoEnBD = "null"

        dicFields.Add dto
    End If

    Set GetFieldsByTipo = dicFields
End Function






