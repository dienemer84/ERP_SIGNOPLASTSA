VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function ConvertBMPtoJPG Lib "ImageUtils.dll" (ByVal InputFile _
                                                               As String, ByVal OutputFile As String, _
                                                               ByVal OverWrite As Boolean, ByVal _
                                                                                           JPGCompression As Integer, ByVal _
                                                                                                                      SaveBMP As Boolean) As Integer


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                      (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                       ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Dim cn As ADODB.Connection
Dim strsql As String
Dim rs As ADODB.Recordset
Dim classArchivos As New classArchivos
Dim vcodigoMaterial As String



Public Property Get nombre() As String
    nombre = rs!nombre
End Property

Public Property Get Tamaño() As Double
    Tamaño = rs!Tamano
End Property
Private Sub Class_Initialize()


    On Error GoTo err22
    Set cn = conectar.obternerConexion    'New ADODB.Connection

    Exit Sub
err22:
    MsgBox Err.Description
End Sub
Public Function lista(consulta As String) As Recordset
    Me.ejecutar consulta
    Set lista = rs
End Function
Public Function ejecutar(consulta As String) As Boolean
    On Error GoTo errEj
    ejecutar = True    '
    'If rs.State = 1 Then rs.Close
    'rs.Open consulta, cn, adOpenDynamic, adLockOptimistic, adCmdText
    Set rs = conectar.RSFactory(consulta)
    Exit Function
errEj:
    'MsgBox "Se produjo un error: " & Err.Description
    ejecutar = False
End Function


Public Function grabarArchivo(idPieza As Long, nombre As String, ruta As String, Comentario As String, Origen As Integer, DeCompra As Boolean) As Boolean
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
        DAOEvento.Publish idPieza, tipoEvento
    End If

    Exit Function
err22:
    MsgBox Err.Description
    grabarArchivo = False

End Function

Public Function exportarArchivo(id As Long, Optional ByVal ruta As String = vbNullString, Optional ejecutar As Boolean = False) As String

    On Error GoTo err22

    exportarArchivo = False
    Dim rsmy As New ADODB.Recordset
    Dim My As New ADODB.Stream
    exportarArchivo = True
    rsmy.CursorLocation = adUseClient
    rsmy.Open "select * from sp_archivos.archivos where id=" & id, cn, adOpenKeyset, adLockOptimistic
    c = 0
    While Not rsmy.EOF
        c = c + 1
        rsmy.MoveNext
    Wend

    If c = 1 Then
        rsmy.MoveFirst
        My.Type = adTypeBinary
        My.Open
        My.Write rsmy.Fields("archivo").value

        If LenB(ruta) = 0 Then
            ruta = Environ$("TEMP")
            If LenB(ruta) = 0 Then ruta = Environ$("TMP")
            If LenB(ruta) = 0 Then ruta = App.path
            ruta = ruta & "\" & funciones.CreateGUID() & "\"
            MkDir ruta
            ruta = ruta & rsmy.Fields("nombre").value
        End If


        My.SaveToFile ruta, adSaveCreateOverWrite
        My.Close

        rsmy.Close
        Set rsmy = Nothing
        Set My = Nothing

        If ejecutar Then
            ShellExecute -1, "open", ruta, "", "", 4
            'ShellExecute frmPrincipal.hwnd, "open", ruta, "", "", 4
        End If

        exportarArchivo = ruta
    End If

    Exit Function
err22:
    exportarArchivo = vbNullString
    Set My = Nothing
End Function

Public Function AbrirArchivo(ruta) As Boolean
    On Error GoTo er1
    Exit Function
er1:
    AbrirArchivo = False
End Function

Private Sub Class_Terminate()
    Set cn = Nothing
    Set rs = Nothing
End Sub


'Public Function cargarActualizacion(nombre As String, vruta As String, v_max As Long, v_max As Long) As Boolean
Public Function cargarActualizacion(nom, rut, vmax, vmin, vrev, Optional notas As String = Empty) As Boolean
    On Error GoTo err221
    cn.BeginTrans
    cargarActualizacion = True
    Dim My As ADODB.Stream
    Set My = New ADODB.Stream
    My.Open
    'CN.Execute "delete  "

    Me.ejecutar "select * from ActualizacionSistema limit 1"

    rs.AddNew

    rs!nombre = "abb.exe"
    rs!version_max = vmax
    rs!version_min = vmin
    rs!version_rev = vrev
    My.Type = adTypeBinary
    My.LoadFromFile rut
    rs!archivo = My.Read
    rs!Tamano = My.Size
    rs!FEcha = funciones.datetimeFormateada(Now)
    rs!notas_version = notas
    My.Close

    rs.Update


    rs.Close


    Dim id_ As Long
    MsgBox "Carga exitosa!", vbInformation, "Información"
    conectar.UltimoId "ActualizacionSistema", id_


    cn.CommitTrans


    If MsgBox("¿Desea agregar más archivos a esta actualización?", vbYesNo, "Confirmación") = vbYes Then
        frmArchivosActualizacion.idVersion = id_
        frmArchivosActualizacion.Show 1

    End If






    Exit Function
err221:
    MsgBox Err.Description
    cargarActualizacion = False
    cn.RollbackTrans
End Function




Public Function CompararConVersionActual(max, min, rev, Optional ByRef idnueva) As Boolean
    On Error GoTo err3
    CompararConVersionActual = False
    Me.ejecutar "select max(id) as id from ActualizacionSistema limit 1"
    idnueva = rs!id

    If Not IsNumeric(idnueva) Then Exit Function
    Me.ejecutar "select version_max,version_min,version_rev from ActualizacionSistema where id=" & idnueva
    If Not rs.EOF And Not rs.BOF Then
        max_ = rs!version_max
        min_ = rs!version_min
        rev_ = rs!version_rev
        If max_ > max Then
            'si la que cargo es mayor a la que esta cargada, listo!!
            CompararConVersionActual = True
            Exit Function
        ElseIf max = max_ Then    'si los mayores son iguales pregunto por el menor
            If min_ > min Then
                CompararConVersionActual = True
                Exit Function
            ElseIf min_ = min Then
                If rev_ > rev Then

                    CompararConVersionActual = True
                    Exit Function
                End If
            End If
        End If
    Else
        CompararConVersionActual = True
    End If
    Exit Function
err3:
    Err.Clear
    Exit Function
End Function


Public Function exportarArchivoSistema(idarchivo As Long) As Boolean

    On Error GoTo err22
    exportarArchivoSistema = False
    Dim rsmy As New ADODB.Recordset
    Dim My As ADODB.Stream
    exportarArchivoSistema = True
    Set My = New ADODB.Stream
    rsmy.CursorLocation = adUseClient
    rsmy.Open "select * from sp.ActualizacionSistema where id= " & idarchivo, cn, adOpenKeyset, adLockOptimistic
    c = 0
    While Not rs.EOF
        c = c + 1
        rs.MoveNext
    Wend

    If c = 1 Then
        id_ver = rsmy!id
        rs.MoveFirst
        My.Type = adTypeBinary
        My.Open
        My.Write rsmy.Fields("archivo").value

        My.SaveToFile App.path & "\abb.tmp", adSaveCreateOverWrite
        My.Close
        rsmy.Close
    End If

    rsmy.CursorLocation = adUseClient
    rsmy.Open "select * from sp.ActualizacionSistema_anexos where id_version= " & id_ver, cn, adOpenKeyset, adLockOptimistic
    c = 0
    While Not rsmy.EOF
        c = c + 1
        rsmy.MoveNext
    Wend


    If c = 0 Then
        exportarArchivoSistema = True
        Exit Function
    End If

    rsmy.MoveFirst


    While Not rsmy.EOF
        Set My = New ADODB.Stream
        My.Type = adTypeBinary
        My.Open
        My.Write rsmy.Fields("archivo").value
        My.SaveToFile App.path & "\" & rsmy.Fields("nombre"), adSaveCreateOverWrite
        rsmy.MoveNext
    Wend

    My.Close
    rsmy.Close
    Set rsmy = Nothing
    Set My = Nothing
    exportarArchivoSistema = True
    Exit Function
err22:
    exportarArchivoSistema = False
    Set My = Nothing
End Function


Public Function escanearDocumento(Origen As Integer, id As Long, Optional tipoColor As TWAIN_MDL_COLOURTYPE = RGB, Optional resol As Long = 130) As Boolean
    On Error GoTo err1
    escanearDocumento = True
    If Permisos.SistemaArchivosScannear Then
        newdoc = sTwain.PopupSelectSourceDialog

        If newdoc = 0 Then
            'newdoc = sTwain.TransferWithUI(App.Path & "\tmp.bmp")
            newdoc = sTwain.TransferWithoutUI(resol, tipoColor, 0, 0, 0, 0, App.path & "\tmp.bmp", id, Origen)
            If newdoc = 0 Then
                MsgBox "Se introdujo exitosamente el documento!", vbInformation, "Información"

            Else
                MsgBox "Se produjo algun error, no se procesan los cambios", vbCritical, "Error"
                GoTo err1
            End If
        End If
    Else
        Permisos.sinAcceso
    End If

    Exit Function
err1:
    escanearDocumento = False
End Function
Public Function grabarEscaneado(id As Long, Origen As Integer, ruta)
    Dim nombr As String
    Dim rutas As String


    If Origen = 100 Then
        Tipo = "REMITO"
    ElseIf Origen = 101 Then
        Tipo = "FACTURA"
    ElseIf Origen = 102 Then
        Tipo = "RECIBO"
    ElseIf Origen = 1 Then
        Tipo = "STOCK"
    ElseIf Origen = 2 Then
        Tipo = "PRES"
    ElseIf Origen = 3 Then
        Tipo = "OT"
    ElseIf Origen = 11 Then
        Tipo = "DETALLE_PEDIDO"
    End If
    nombr = Tipo & id & ".JPG"


    a = Me.exportarJPG(App.path & "\tmp.bmp", App.path & "\tmp.jpg")
    If a = 0 Then
        If classArchivos.grabarArchivo(id, nombr, App.path & "\tmp.jpg", "ESCANEO DIRECTO", Origen, False) Then
            Kill App.path & "\tmp.bmp"
        End If
    End If



End Function

Private Sub WriteByteArray(ByVal strPath As String, ByRef arrData() As Byte)

    Dim lngFile As Long

    ' open the file
    lngFile = FreeFile()
    Open strPath For Binary Access Write As lngFile

    ' write blob
    Put lngFile, , arrData

    ' close file
    Close lngFile

End Sub

Public Function previewImage2(a As archivo) As Picture
    On Error GoTo err22
    Dim P As String
    P = App.path & "\" & funciones.CreateGUID & ".tmp"
    WriteByteArray P, a.Contenido
    Set previewImage2 = LoadPicture(P)
    Kill P
    Exit Function
err22:
    Set previewImage2 = Nothing
End Function


Public Function previewImage(id) As Picture

    On Error GoTo err22
    Dim rsmy As New ADODB.Recordset
    Dim My As New ADODB.Stream
    rsmy.CursorLocation = adUseClient
    My.Type = adTypeBinary
    rsmy.Open "select * from sp_archivos.archivos where id=" & id, cn, adOpenKeyset, adLockOptimistic
    c = 0
    While Not rsmy.EOF
        c = c + 1
        rsmy.MoveNext
    Wend
    If c = 1 Then
        rsmy.MoveFirst
        My.Open
        My.Write rsmy.Fields("archivo").value
        Dim P As String
        P = App.path & "\" & funciones.CreateGUID & ".tmp"
        My.SaveToFile P, adSaveCreateOverWrite
        Set previewImage = LoadPicture(P)
        Kill P
        My.Close
        rsmy.Close
        Set rsmy = Nothing
        Set My = Nothing
    Else
        Set previewImage = Nothing
    End If
    Exit Function
err22:
    Set previewImage = Nothing
    Kill P
    Set My = Nothing
End Function


Public Function exportarJPG(Origen, destino)
    exportarJPG = ConvertBMPtoJPG(Origen, destino, True, 20, True)
End Function








Public Function verReporteFoto(idphoto) As Boolean    '1- enviar 2-imprimir
    On Error GoTo err2
    verReporteFoto = True



    verImagen.Sections("cuerpo").Controls("imagen1").Picture = Me.previewImage(idphoto)
    verImagen.Show


    Exit Function
err2:
    MsgBox Err.Description
    verReporteFoto = False
End Function



