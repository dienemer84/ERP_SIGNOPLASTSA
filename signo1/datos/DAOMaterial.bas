Attribute VB_Name = "DAOMateriales"
Dim strsql As String

Public Const TABLA_ID_ALMACEN As String = "a"
Public Const TABlA_ID_GRUPO As String = "g"
Public Const TABlA_ID_RUBRO As String = "r"
Public Const TABLA_ID_MONEDA As String = "mon"
Public Const TABLA_MATERIALES As String = "m"

Public Const CAMPO_ANCHO As String = "ancho"
Public Const CAMPO_LARGO As String = "largo"
Public Const CAMPO_CANTIDAD As String = "cantidad"
Public Const CAMPO_CODIGO As String = "codigo"
Public Const CAMPO_DESCRIPCION As String = "descripcion"
Public Const CAMPO_ESPESOR As String = "espesor"
Public Const CAMPO_ESTADO As String = "estado"
Public Const CAMPO_ID As String = "id"
Public Const CAMPO_PESO_X_UNIDAD As String = "pesoxunidad"
Public Const CAMPO_UNIDAD As String = "id_unidad"
Public Const CAMPO_VALOR_UNITARIO As String = "valor_unitario"
Public Const CAMPO_FECHA_VALOR As String = "fecha_valor"
Public Const CAMPO_TIPO As String = "tipo"
'Public Const CAMPO_VALOR As String = "valor"

Public Const CAMPO_ID_GRUPO As String = "id_grupo"
Public Const CAMPO_ID_ALMACEN As String = "idAlmacen"
Public Const CAMPO_ID_MONEDA As String = "id_moneda"


Public Function LlenarComboPorRubro(cbo As ComboBox, rubro As clsRubros)
    Dim col As Collection
    cbo.Clear
    Set col = DAOMateriales.FindAll(DAOMateriales.TABLA_MATERIALES & ".id_rubro" & "=" & rubro.Id)
    Dim m As clsMaterial

    For Each m In col
        cbo.AddItem m.descripcion
        cbo.ItemData(cbo.NewIndex) = m.Id

    Next

End Function


Public Function desaprobar(Id As Long) As Boolean
    On Error GoTo err1
    Dim strsql As String

    strsql = "update materiales set aprobado='0'  WHERE id=" & Id
    desaprobar = conectar.execute(strsql)

    Exit Function
err1:
    Err.Raise Err.Number
End Function

Public Function aprobar(Material As clsMaterial) As Boolean
    On Error GoTo err1
    If Not Material.Aprobado Then
        Material.Aprobado = True
        If Not modificar(Material) Then
            Err.Raise 101213, , "no se pudo aprobar el material"
        Else
            DaoHistorico.Save "materiales_historial", "Material aprobado", Material.Id
        End If
    End If
    Exit Function
err1:
    Err.Raise Err.Number
End Function


Public Function existeCodigo(codigo As String) As Boolean
    Set rs = conectar.RSFactory("select count(id) as c from materiales where codigo='" & codigo & "'")
    If Not rs.EOF And Not rs.BOF Then
        If rs!c > 0 Then
            existeCodigo = False
        Else
            existeCodigo = True
        End If
    Else
        existeCodigo = False
    End If

End Function
Public Function modificar(Material As clsMaterial, Optional desaprobar As Boolean = False) As Boolean
    Dim EVENTO As New clsEventoObserver
    On Error GoTo er1
    modificar = True

    Dim apro As Boolean


    With Material
        apro = .Aprobado
        If desaprobar Then apro = False
        Material.Aprobado = apro

        strsql = "update materiales set aprobado=" & conectar.Escape(apro) & ", valor_compra = " & .ValorCompra & ", unidad_compra=" & .UnidadCompra & " , unidad_pedido=" & .UnidadPedido & ", id_moneda=" & .moneda.Id & ", fecha_valor=" & Escape(.FechaValor) & ", valor_unitario=" & Escape(.Valor) & ", largo=" & .Largo & ",ancho=" & .Ancho & ",id_rubro=" & .Grupo.rubros.Id & ",id_grupo=" & .Grupo.Id & ",id_unidad=" & .unidad & ",codigo='" & .codigo & "',descripcion='" & .descripcion & "',espesor=" & .Espesor & ",pesoxunidad=" & .PesoXUnidad & ",idAlmacen=" & .almacen.Id & ",cantidad=" & .Cantidad & ", altura = " & conectar.Escape(.Altura) & ", tipo=" & .Tipo & ", puntoReposicion = " & Escape(.PuntoReposicion) & ", stockMinimo =" & Escape(.StockMinimo) & "  WHERE id=" & Material.Id
    End With
    If Not conectar.execute(strsql) Then GoTo er1
    Set EVENTO.Elemento = Material
    EVENTO.EVENTO = modificar_
    EVENTO.Tipo = Materiales_
    Channel.Notificar EVENTO, Materiales_
    Exit Function
er1:
    modificar = False
    Err.Raise 2100, , "Error al modificar"
End Function


Public Function crear(Material As clsMaterial) As Boolean
    Dim UltimoId As Long
    On Error GoTo err1

    crear = True
    conectar.BeginTransaction
    With Material
        strsql = "insert into materiales (valor_compra,unidad_pedido,unidad_compra,codigo,id_rubro,id_grupo,id_unidad,descripcion,espesor,pesoxunidad,estado,IdAlmacen,largo,ancho,valor_unitario,fecha_valor,id_moneda,tipo,altura,stockMinimo,puntoReposicion) Values (" & Escape(.ValorCompra) & " ," & Escape(.UnidadPedido) & "," & Escape(.UnidadCompra) & "," & Escape(.codigo) & "," & .Grupo.rubros.Id & "," & .Grupo.Id & "," & .unidad & ",'" & .descripcion & "'," & .Espesor & "," & .PesoXUnidad & "," & .estado & "," & .almacen.Id & "," & .Largo & "," & .Ancho & "," & Escape(.Valor) & "," & Escape(.FechaValor) & "," & .moneda.Id & ", " & .Tipo & ", " & conectar.Escape(.Altura) & ", " & Escape(.StockMinimo) & ", " & Escape(.PuntoReposicion) & ")"
    End With
    If Not conectar.execute(strsql) Then
        Err.Raise 1231, , "insert" & vbNewLine & strsql
        'GoTo err
    End If

    Material.Id = conectar.UltimoId2
    If Material.Id = 0 Then
        Err.Raise 1231, , "ultimoid2"
        'GoTo err1
    End If

    strsql = "update materiales set codigo = CONCAT(codigo, '-', id) where id = " & Material.Id
    If Not conectar.execute(strsql) Then
        Err.Raise 1231, , "update materiales" & vbNewLine & strsql
        'GoTo err1
    End If

    If Material.Valor = 0 Then MsgBox "Recuerde que el material queda pendiente de activación", vbInformation, "Información"
    If Not DAOMaterialHistorico.crear(Material) Then
        Err.Raise 1231, , "materialhistorico.crear"
        'GoTo err1
    End If
    conectar.CommitTransaction

    Dim EVENTO As New clsEventoObserver
    Set EVENTO.Elemento = Material
    EVENTO.EVENTO = agregar_
    EVENTO.Tipo = Materiales_
    Channel.Notificar EVENTO, Materiales_

    Exit Function
err1:
    crear = False
    conectar.RollBackTransaction
    MsgBox Err.Description, vbCritical
End Function


Public Function FindById(Id As Long, Optional includeHistorico As Boolean = False) As clsMaterial
    Dim col As Collection
    Set col = DAOMateriales.FindAll(DAOMateriales.TABLA_MATERIALES & "." & DAOMateriales.CAMPO_ID & " = " & Id, includeHistorico)
    If col.count > 0 Then
        Set FindById = col(1)
    Else
        Set FindById = Nothing
    End If
End Function

Public Function FindAll(Optional whereFilter As String = vbNullString, Optional includeHistorico As Boolean = False) As Collection
    Dim tickStart As Double
    Dim tickend As Double
    tickStart = GetTickCount
    Dim rs As ADODB.Recordset
    Dim q As String
    Dim Materiales As New Collection

    q = "SELECT m.*, g.*, r.*, a.*, mon.*" _
      & " FROM materiales m" _
      & " LEFT JOIN grupos g ON g.id = m.id_grupo" _
      & " LEFT JOIN rubros r ON r.id = g.id_rubro" _
      & " LEFT JOIN materialesAlmacenes a ON a.id = m.idAlmacen" _
      & " LEFT JOIN AdminConfigMonedas mon ON mon.id = m.id_moneda" _
      & " WHERE 1 = 1"

    If LenB(whereFilter) > 0 Then
        q = q & " AND " & whereFilter
    End If

    Set rs = conectar.RSFactory(q)
    Dim fieldsIndex As Dictionary
    BuildFieldsIndex rs, fieldsIndex

    Dim MAT As clsMaterial

    While Not rs.EOF
        Set MAT = DAOMateriales.Map(rs, fieldsIndex, TABLA_MATERIALES, TABLA_ID_ALMACEN, TABLA_ID_MONEDA, TABlA_ID_GRUPO, TABlA_ID_RUBRO)

        If includeHistorico Then
            MAT.historico = DAOMaterialHistorico.getAllByMaterial(MAT.Id)
        End If

        Materiales.Add MAT, CStr(MAT.Id)
        rs.MoveNext
    Wend

    Set FindAll = Materiales
End Function

Public Function Map(rs As Recordset, indice As Dictionary, tabla As String, Optional tablaAlmacen As String = vbNullString, Optional tablaMoneda As String = vbNullString, Optional tablaGrupo As String = vbNullString, Optional tablaRubro As String = vbNullString) As clsMaterial
    Dim m As clsMaterial
    Dim Id As Long

    Id = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_ID)

    If Id > 0 Then
        Set m = New clsMaterial
        m.Id = Id

        m.Ancho = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_ANCHO)
        m.Cantidad = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_CANTIDAD)
        m.codigo = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_CODIGO)
        m.descripcion = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_DESCRIPCION)
        m.Espesor = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_ESPESOR)
        m.estado = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_ESTADO)
        m.FechaValor = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_FECHA_VALOR)
        m.Largo = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_LARGO)
        m.PesoXUnidad = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_PESO_X_UNIDAD)
        m.unidad = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_UNIDAD)
        m.UnidadCompra = GetValue(rs, indice, tabla, "unidad_compra")
        m.UnidadPedido = GetValue(rs, indice, tabla, "unidad_pedido")
        m.ValorCompra = GetValue(rs, indice, tabla, "valor_compra")

        m.Valor = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_VALOR_UNITARIO)

        m.Tipo = GetValue(rs, indice, tabla, DAOMateriales.CAMPO_TIPO)
        m.StockMinimo = GetValue(rs, indice, tabla, "stockMinimo")
        m.PuntoReposicion = GetValue(rs, indice, tabla, "puntoReposicion")
        m.Aprobado = GetValue(rs, indice, tabla, "aprobado")
        'If LenB(tablaAlmacen) > 0 Then Set m.almacen = DAOAlmacenes.Map(rs, indice, tablaAlmacen)
        'If LenB(tablaGrupo) Then Set m.grupo = DAOGrupos.Map(rs, indice, tablaGrupo, tablaRubro)
        'If LenB(tablaMoneda) Then Set m.Moneda = DAOMoneda.Map(rs, indice, tablaMoneda)

        'historico?

        If LenB(tablaMoneda) > 0 Then m.moneda = DAOMoneda.Map(rs, indice, tablaMoneda)  'poner set
        If LenB(tablaAlmacen) > 0 Then m.almacen = DAOAlmacenes.Map(rs, indice, tablaAlmacen)    'poner set
        If Len(tablaGrupo) > 0 Then m.Grupo = DAOGrupos.Map(rs, indice, tablaGrupo, tablaRubro)    'poner set

    End If

    Set Map = m
End Function
