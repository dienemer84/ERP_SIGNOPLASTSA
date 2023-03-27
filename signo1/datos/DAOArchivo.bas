Attribute VB_Name = "DAOArchivo"
Option Explicit

Public Enum OrigenArchivos
    OA_OrdenesTrabajo = 3
    OA_Presupuestos = 2
    OA_Piezas = 1
    OA_OrdenesTrabajoDetalle = 111
    OA_OrdenesTrabajoDetalleConjunto = 666
    OA_PresupuestoDetalle = 11
    OA_factura = 160
    OA_FacturaProveedor = 1600
    OA_Empleados = 225
    OA_Siniestros = 192
    OA_ArchivoDocumento = 700
    OA_FotoEmpleado = 812
    OA_Remitos = 100
    OA_Materiales = 1441
    OA_Recibos = 1442
    OA_NotaNoConformidad = 471
End Enum


Public Const CAMPO_ID As String = "id"
Public Const CAMPO_ID_REFERENCIA As String = "idPieza"
Public Const CAMPO_NOMBRE As String = "nombre"
Public Const CAMPO_TAMAÑO As String = "tamano"
Public Const CAMPO_CONTENIDO As String = "archivo"
Public Const CAMPO_COMENTARIO As String = "comentario"
Public Const CAMPO_ID_USUARIO As String = "usuario"
Public Const CAMPO_ORIGEN As String = "origen"

Public Const TABLA_ARCHIVO As String = "arch"


Public Function GetCantidadArchivosPorReferencia(ByRef Origen As OrigenArchivos, Optional ByRef idReferencias As Variant) As Dictionary
    Dim diccionarioRetorno As New Dictionary

    Dim q As String
    q = "SELECT idPieza AS idReferencia, COUNT(0) AS cant from sp_archivos.archivos WHERE origen = " & Origen

    If Not IsMissing(idReferencias) Then
        q = q & " AND idPieza IN (" & Join(idReferencias, ", ") & ")"
    End If
    q = q & " GROUP BY idPieza"

    Dim rs As New Recordset
    Set rs = conectar.RSFactory(q)
    While Not rs.EOF
        diccionarioRetorno.Add rs.Fields("idReferencia").value, rs.Fields("cant").value
        rs.MoveNext
    Wend

    Set GetCantidadArchivosPorReferencia = diccionarioRetorno
End Function



Public Function FindAll(ByVal Origen As OrigenArchivos, Optional ByRef filter As String = vbNullString) As Collection
'Dim tickStart As Double
'Dim tickEnd As Double
'tickStart = GetTickCount

    Dim rs As ADODB.Recordset
    Dim q As String

    Dim archivos As New Collection

    q = "SELECT" _
      & " arch.id,arch.idPieza,arch.nombre,arch.tamano,arch.comentario,arch.usuario,arch.origen,arch.sincro,arch.fecha,arch.de_compra," _
      & " u.*" _
      & " from sp_archivos.archivos arch" _
      & " LEFT JOIN usuarios u" _
      & " ON u.id = arch.usuario" _
      & " WHERE 1 = 1 AND arch.origen = " & Origen


    If LenB(filter) > 0 Then
        q = q & " AND " & filter

    End If


    Set rs = conectar.RSFactory(q)

    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex
    Set FindAll = New Collection

    Const usuarioTabla As String = "u"

    While Not rs.EOF
        archivos.Add DAOArchivo.Map(rs, fieldsIndex, TABLA_ARCHIVO, usuarioTabla)
        rs.MoveNext
    Wend

    'tickEnd = GetTickCount

    'Debug.Print tickEnd - tickStart, "ms elapsed"

    Set FindAll = archivos
End Function




Public Function Map(ByRef rs As Recordset, ByRef fieldsIndex As Dictionary, ByRef tableNameOrAlias As String, Optional ByRef usuarioTableNameOrAlias As String = vbNullString) As archivo
    Dim A As archivo
    Dim Id As Variant

    Id = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOArchivo.CAMPO_ID)

    If Id > 0 Then
        Set A = New archivo
        A.Id = Id
        A.Comentario = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOArchivo.CAMPO_COMENTARIO)
        'a.Contenido = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOArchivo.CAMPO_CONTENIDO)
        A.IdReferencia = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOArchivo.CAMPO_ID_REFERENCIA)
        A.nombre = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOArchivo.CAMPO_NOMBRE)
        A.Origen = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOArchivo.CAMPO_ORIGEN)
        A.Tamaño = GetValue(rs, fieldsIndex, tableNameOrAlias, DAOArchivo.CAMPO_TAMAÑO)
        A.FechaSubida = GetValue(rs, fieldsIndex, tableNameOrAlias, "fecha")
        If LenB(usuarioTableNameOrAlias) > 0 Then Set A.usuario = DAOUsuarios.Map(rs, fieldsIndex, usuarioTableNameOrAlias)
        A.DeCompra = GetValue(rs, fieldsIndex, tableNameOrAlias, "de_compra")
    End If

    Set Map = A
End Function


Public Function grabarArchivo(idPieza As Long, nombre As String, ruta As String, Comentario As String, Origen As Integer, DeCompra As Boolean, Optional value As ISuscriber) As Boolean
    On Error GoTo err22
    grabarArchivo = True
    Dim My As ADODB.Stream
    Set My = New ADODB.Stream
    Dim rss As Recordset
    My.Open
    Set rss = conectar.RSFactory("select * from sp_archivos.archivos where id=1")
    ' Me.ejecutar "select * from sp_archivos.archivos where id=1"


    rss.AddNew
    rss!idPieza = idPieza
    rss!nombre = UCase(nombre)
    rss!Comentario = UCase(Comentario)

    My.Type = adTypeBinary
    My.LoadFromFile ruta
    rss!archivo = My.Read

    rss!Tamano = My.Size
    rss!usuario = funciones.getUser
    rss!Origen = Origen
    rss!de_compra = DeCompra


    '  Dim q As String
    'q = "insert into sp_archivos.archivos    (idPieza,    nombre,    tamano,    archivo,    comentario,    usuario,    origen,    de_compra    )    Values " _
     '& "( '" & idPieza & "','" & UCase(nombre) & "',    '" & My.Size & "',  '" & My.Read & "',    '" & UCase(Comentario) & "',    '" & funciones.getUser & "' ,' " & Origen & "','" & DeCompra & "'  );"



    My.Close

    rss.Update
    rss.Close





    Dim tipoEvento As TipoEventoBroadcast: tipoEvento = -1
    Select Case Origen
    Case OrigenArchivos.OA_OrdenesTrabajo
        tipoEvento = TipoEventoBroadcast.TEB_ArchivoOrdenTrabajo
    Case OrigenArchivos.OA_OrdenesTrabajoDetalle
        tipoEvento = TipoEventoBroadcast.TEB_ArchivoDetalleOrdenTrabajo
    Case OrigenArchivos.OA_Piezas
        tipoEvento = TipoEventoBroadcast.TEB_ArchivoPieza
    End Select

    If tipoEvento <> -1 Then
        DAOEvento.Publish idPieza, tipoEvento, value
    End If

    Exit Function
err22:
    'Err.Raise Err.Number, Err.Source, Err.Description
    grabarArchivo = False

End Function
