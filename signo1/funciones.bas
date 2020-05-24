Attribute VB_Name = "funciones"


Public TareaAgregada As Boolean
Public DescuentoDetalleFactura As Double

Private Declare Function GetShortPathName Lib "kernel32" _
                                          Alias "GetShortPathNameA" _
                                          (ByVal lpszLongPath As String, _
                                           ByVal lpszShortPath As String, _
                                           ByVal cchBuffer As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias _
                                     "GetTempPathA" (ByVal nBufferLength As Long, ByVal _
                                                                                  lpBuffer As String) As Long


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" ( _
                                      pguid As GUID) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" ( _
                                         rguid As Any, _
                                         ByVal lpstrClsId As Long, _
                                         ByVal cbMax As Long) As Long


Dim usuario_online As clsUsuario
Dim vIdProveedorElegido As Long
Dim vCantOE As Double
Dim versionado As String
Dim vDatosCheques As Collection
Dim vCantidadOE As Double
Dim vIdEntrega As Collection
Dim vidFacturas As Long
Public serverBBDD As String
Public serverSMTPe As String
Dim vector()
Dim estados_remitos_facturas(0 To 3)
Dim estados_facturas(1 To 4)
Dim estados_recibos(1 To 3)
Public estados_pedidos(1 To 6)
Dim estados_OC(0 To 4)
Dim estados_OE(1 To 4)
Dim estados_remitos(1 To 3)
Dim vIdReciboElegido As Long
Dim estados_Reques(0 To 5)
Dim vrto As Long
Dim clssp As New classSignoplast
Dim vpieza As Long
Dim vpiezadetalle As String
Dim vpiezadetallebusqueda As String
Private usu As Long
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim vfechas As Date
Dim col As Collection
Dim vOtElegidas As Collection
Dim vValor As Double
Dim vConcValor As Double, vConcCantidad As Double, vConcConc As String
Attribute vConcCantidad.VB_VarUserMemId = 1073741858
Attribute vConcConc.VB_VarUserMemId = 1073741858
Dim vmontoDeposito As Double
Attribute vmontoDeposito.VB_VarUserMemId = 1073741859
Dim vIdCuentaDeposito As Long
Attribute vIdCuentaDeposito.VB_VarUserMemId = 1073741860
Dim vFechaDeposito As Date
Attribute vFechaDeposito.VB_VarUserMemId = 1073741861
Dim vItemRemito As Long
Attribute vItemRemito.VB_VarUserMemId = 1073741862

'para buscar archivos
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (bBrowse As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long



Dim tmpFileMetadata As FileMetadataDTO
Attribute tmpFileMetadata.VB_VarUserMemId = 1073741863

Private total_size As Double
Attribute total_size.VB_VarUserMemId = 1073741864

Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const ERROR_NO_MORE_FILES = 18&
Private Const INVALID_HANDLE_VALUE = -1
Private Const DDL_DIRECTORY = &H10

Public Enum DateRangeValue
    DRV_Today = 1
    DRV_Yesterday = 2
    DRV_WeekCurrent = 3
    DRV_WeekLast = 4
    DRV_MonthCurrent = 5
    DRV_MonthLast = 6
    DRV_YearCurrent = 7
    DRV_YearLast = 8
    DRV_MonthNext = 9
End Enum


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                                     (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
                                      ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'para medir tiempos
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private m_bInIDE As Boolean
Public Property Get InIDE() As Boolean
    Debug.Assert (IsInIDE())
    InIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
    m_bInIDE = True
    IsInIDE = m_bInIDE
End Function


Public Property Get idReciboElegido() As Long
    idReciboElegido = vIdReciboElegido
End Property

Public Property Let idReciboElegido(nIdReciboElegido As Long)
    vIdReciboElegido = nIdReciboElegido
End Property

Public Sub SalirForzado()
On Error Resume Next
Dim aa As String
aa = App.path & "\AbbKiller.exe"
        Shell aa, vbNormalFocus
End Sub

Public Function GetFileName(ByVal path As String) As String
    Dim Contador As Integer
    Contador = 1
    While Mid(path, Len(path) - Contador, 1) <> "\"
        Contador = Contador + 1
    Wend
    GetFileName = Mid(path, Len(path) - Contador + 1)
End Function

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long

    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
                               Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)

End Function

Public Property Let depositoMonto(nMonto As Double)
    vmontoDeposito = nMonto
End Property

Property Let itemRemito(nItemRemito As Long)
    vItemRemito = nItemRemito
End Property

Public Property Get itemRemito() As Long
    itemRemito = vItemRemito
End Property

Public Property Let depositoIdCuenta(nIdCuenta As Long)
    vIdCuentaDeposito = nIdCuenta
End Property
Public Property Let depositoFecha(nFechaDeposito As Date)
    vFechaDeposito = nFechaDeposito
End Property
Public Property Get depositoMonto() As Double
    depositoMonto = vmontoDeposito
End Property
Public Property Get depositoIdCuenta() As Long
    depositoIdCuenta = vIdCuentaDeposito
End Property
Public Property Get depositoFecha() As Date
    depositoFecha = vFechaDeposito
End Property


Public Property Let OTElegidas(nOtElegidas As Collection)
    Set vOtElegidas = nOtElegidas
End Property

Public Property Get OTElegidas() As Collection
    Set OTElegidas = vOtElegidas
End Property

Public Property Let datosCheques(nDatosCheques As Collection)
    Set vDatosCheques = nDatosCheques
End Property
Public Property Get datosCheques() As Collection
    Set datosCheques = vDatosCheques
End Property
Public Property Let idFactura(nIdFactura As Long)
    vidFacturas = nIdFactura
End Property
Public Property Get idFactura() As Long
    idFactura = vidFacturas
End Property
Public Property Let ConcValor(nConcValor As Double)
    vConcValor = nConcValor
End Property
Public Property Let ConcCantidad(nConcCantidad As Double)
    vConcCantidad = nConcCantidad
End Property
Public Property Let ConcConc(nConcConc As String)
    vConcConc = nConcConc
End Property
Public Property Get ConcValor() As Double
    ConcValor = vConcValor
End Property
Public Property Get ConcCantidad() As Double
    ConcCantidad = vConcCantidad
End Property
Public Property Get ConcConc() As String
    ConcConc = vConcConc
End Property




Public Function queTipoFactura(ind As Integer) As String
    If ind = 0 Then
        queTipoFactura = "Factura"
    ElseIf ind = 1 Then
        queTipoFactura = "N.Crédito"
    ElseIf ind = 2 Then
        queTipoFactura = "N.Débito"
    End If


End Function


Public Function Iva(niva As Integer) As String
    Select Case niva
        Case 0: Iva = "Exento"
        Case 1: Iva = "Resp. Monotributo"
        Case 2: Iva = "Resp. No Inscripto"
        Case 3: Iva = "Resp. Inscripto"
        Case 4: Iva = "N/d"
        Case 5: Iva = "Consumidor Final"
    End Select


End Function
Public Property Let valorOE(nValor As Double)
    vValor = nValor
End Property



Public Property Get cantOE() As Double
    cantOE = vCantOE
End Property
Public Property Let cantOE(nCantOE As Double)
    vCantOE = nCantOE
End Property



Public Property Get valorOE() As Double
    valorOE = vValor
End Property

Public Property Let CantidadOE(nCantidadoe As Double)
    vCantidadOE = nCantidadoe
End Property
Public Property Get CantidadOE() As Double
    CantidadOE = vCantidadOE
End Property


Public Property Let coleccionPiezas(nCol As Collection)
    Set col = nCol
End Property
Public Property Get coleccionPiezas() As Collection
    Set coleccionPiezas = col
End Property
Public Property Let fechas(nfechas As Date)
    vfechas = nfechas
End Property
Public Property Get fechas() As Date
    fechas = vfechas
End Property

Public Function IsValidEmail(AddressString As String) As Boolean

    Dim sTmp() As String

    ' assume failure
    IsValidEmailAddress = False

    ' sould have one "@"
    sTmp = Split(AddressString, "@")
    If UBound(sTmp) <> 1 Then Exit Function

    IsValidEmailAddress = True

End Function


Public Function EsConjunto(conj As Long) As String
    If conj = 0 Then
        EsConjunto = "Conjunto"
    ElseIf conj = -1 Then
        EsConjunto = "Unidad"
    End If
End Function

Public Property Let idEntrega(nIdEntrega As Collection)
    Set vIdEntrega = nIdEntrega

End Property
Public Property Get idEntrega() As Collection
    Set idEntrega = vIdEntrega
End Property


Public Property Let queRemitoElegido(rto As Long)
    vrto = rto
End Property
Public Property Get queRemitoElegido() As Long
    queRemitoElegido = vrto
End Property
Public Property Get quePiezaElegida() As Long
    quePiezaElegida = vpieza
End Property
Public Property Let quePiezaElegidaDetalle(idstockdetalle As String)
    vpiezadetalle = idstockdetalle
End Property
Public Property Get quePiezaElegidaDetalle() As String
    quePiezaElegidaDetalle = vpiezadetalle
End Property
Public Property Get quePiezaElegidabusqueda() As String
    quePiezaElegidabusqueda = vpiezadetallebusqueda
End Property
Public Property Let quePiezaElegidabusqueda(piezadetallebusqueda As String)
    vpiezadetallebusqueda = piezadetallebusqueda
End Property
Public Property Let quePiezaElegida(idStock As Long)
    vpieza = idStock
End Property
Public Sub setServerSMTPE(servidor)
    serverSMTPe = servidor
End Sub
Public Function getServerSMTPe() As String
    getServerSMTPe = serverSMTPe
End Function
Public Property Get itemsPorRemito() As Integer
    itemsPorRemito = 31    'cambiarrrrrrrrrrrrrrrrr
End Property
Public Property Get itemsPorFactura()
    itemsPorFactura = 34

End Property

Public Function tipoEvento(indice) As String
    Select Case indice
        Case 0: tipoEvento = "Notas"
        Case 1: tipoEvento = "Vencimientos"
        Case 2: tipoEvento = "Avisos"
        Case 3: tipoEvento = "Mensajes"
    End Select

End Function

Public Sub setUser(iduser As Long)
    usu = iduser
    Set usuario_online = DAOUsuarios.GetById(iduser)
End Sub
Public Function getUser()
    getUser = usu
End Function
Public Property Get GetUserObj() As clsUsuario
    Set GetUserObj = usuario_online
End Property
Function cuantosDias(horas) As Double
    cuantosDias = horas / 12    'despues analizar cuantas horas posibles por dia de trabajo
End Function
Function queMoneda(IdMoneda As Integer, Optional ByRef Largo As String = Empty) As String
    Dim clasea As New classAdministracion
    Dim rs As Recordset
    If IdMoneda = 0 Then
        queMoneda = "AR$"
    ElseIf IdMoneda = 1 Then
        queMoneda = "US$"
    Else
        queMoneda = "Error"
    End If

    Set rs = conectar.RSFactory("select  nombre_corto, nombre_largo from AdminConfigMonedas where id=" & IdMoneda)
    If Not rs.EOF And Not rs.BOF Then
        queMoneda = rs!Nombre_corto

        If Largo = Empty Then
            Largo = rs!nombre_largo
        End If
    Else
        queMoneda = "Error"
        If Largo <> Empty Then
            Largo = "Error"
        End If
    End If
End Function
Function queUnidad(idUnidad As Integer) As String
    If idUnidad = 1 Then
        queUnidad = "Kg"
    ElseIf idUnidad = 2 Then
        queUnidad = "M2"
    ElseIf idUnidad = 3 Then
        queUnidad = "Ml"
    ElseIf idUnidad = 4 Then
        queUnidad = "Un"
    Else
        queUnidad = "Error"
    End If
End Function

Public Sub llenar_vectores()
    estados_recibos(1) = "Pendiente"
    estados_recibos(2) = "Aprobado"
    estados_recibos(3) = "Anulado"


    estados_facturas(1) = "Pendiente"
    estados_facturas(2) = "Aprobada"
    estados_facturas(3) = "Anulada"
    estados_facturas(4) = "Cancelada NC"

    estados_pedidos(1) = "Pendiente"
    estados_pedidos(2) = "En proceso"
    estados_pedidos(3) = "Proceso completo"
    estados_pedidos(4) = "Finalizado"
    estados_pedidos(5) = "En Espera"
    estados_pedidos(6) = "Desactivado"

    estados_OE(1) = "Pendiente"
    estados_OE(2) = "Aprobado"
    estados_OE(3) = "Finalizado"
    estados_OE(4) = "Proceso Completo"

    estados_remitos(1) = "En proceso"
    estados_remitos(2) = "Procesado"
    estados_remitos(3) = "Anulado"


    estados_remitos_facturas(0) = "No Facturado"
    estados_remitos_facturas(1) = "Parcial"
    estados_remitos_facturas(2) = "Completo"
    estados_remitos_facturas(3) = "No Facturable"

    estados_OC(0) = "En proceso"
    estados_OC(1) = "Pendiente"
    estados_OC(2) = "Enviada"
    estados_OC(3) = "Rec.Parcial"
    estados_OC(4) = "Finalizada"


    versionado = App.Major & "." & App.Minor & "." & App.Revision
    mostrarInforme
End Sub
Public Property Get Version() As String
    Version = versionado
End Property
Function estado_remitos_facturas(indice As Integer) As String
    estado_remitos_facturas = estados_remitos_facturas(indice)
End Function


Function estado_factura_cobranza(indice) As String
    If indice = 0 Then
        estado_factura_cobranza = "Pendiente"
    ElseIf indice = 1 Then
        estado_factura_cobranza = "Saldada"
    ElseIf indice = 2 Then
        estado_factura_cobranza = "Parcial"
    ElseIf indice = 3 Then
        estado_factura_cobranza = "N.Crédito"
    ElseIf indice = 4 Then
        estado_factura_cobranza = "N.Débito"
    ElseIf indice = 5 Then
        estado_factura_cobranza = "Anulada"
    End If


End Function

Function estado_oc(indice As Integer) As String
    estado_oc = estados_OC(indice)
End Function
Function estado_recibo(indice As Integer) As String
    estado_recibo = estados_recibos(indice)
End Function

Function estado_factura(indice As Integer) As String
    estado_factura = estados_facturas(indice)
End Function
Function estado_pedido(indice As Integer) As String
    If indice >= LBound(estados_pedidos) And indice <= UBound(estados_pedidos) Then
        estado_pedido = estados_pedidos(indice)
    End If
End Function
Function estado_rto(indice As Integer) As String
    estado_rto = estados_remitos(indice)
End Function

Function estado_OE(indice As Integer) As String
    estado_OE = estados_OE(indice)
End Function

Function estado_reque(indice As Integer) As String
    estado_reque = estados_Reques(indice)
End Function

Function normalizaVieja(str As String)
    'normaliza un string para que la primer letra
    'sea Mayúscula y el resto minúscula.
    'y reemplaza los caracteres por los
    'caracteres de escape correspondientes
    strt = Trim(str)
    If strt <> Empty Then
        strt = Replace(strt, Chr(92), "\\")
        strt = Replace(strt, Chr(39), "\'")
        strt = Replace(strt, Chr(34), "\""")
        strt = Replace(strt, Chr(0), "\0")
        strt = LCase(strt)
        str_l = Left(strt, 1)
        str_l = UCase(str_l)
        str_r = Right(strt, Len(strt) - 1)
        strt = str_l & str_r



        normalizaVieja = strt
    End If
End Function

Function CentrarImpresion(texto As String) As Long
    Dim s As Long
    Dim usado As Long
    s = Printer.TextWidth(texto)
    usado = Printer.Width - s
    CentrarImpresion = usado / 2
End Function

Function truncar(str, Cantidad As Long) As String
    'trunca el texto a la cantidad de
    'caracteres que quiera
    If Trim(str) <> Empty Then
        If Len(Trim(str)) > Cantidad Then
            str = Left(str, Cantidad - 3)
            str = str & "..."
        End If
        truncar = str
    End If
End Function
Function FEcha(d As Date) As String
    'devuelve un string con la fecha de hoy en formato
    'Lunes, 25 de diciembre de 2005
    Dim mes(1 To 12) As String
    Dim dia(1 To 7) As String
    mes(1) = "Enero"
    mes(2) = "Febrero"
    mes(3) = "Marzo"
    mes(4) = "Abril"
    mes(5) = "Mayo"
    mes(6) = "Junio"
    mes(7) = "Julio"
    mes(8) = "Agosto"
    mes(9) = "Septiembre"
    mes(10) = "Octubre"
    mes(11) = "Noviembre"
    mes(12) = "Diciembre"
    dia(1) = "Domingo"
    dia(2) = "Lunes"
    dia(3) = "Martes"
    dia(4) = "Miercoles"
    dia(5) = "Jueves"
    dia(6) = "Viernes"
    dia(7) = "Sábado"
    FEcha = dia(Weekday(d)) & ", " & Day(d) & " de " & mes(Month(d)) & " de " & Year(d)
End Function
Function PosIndexCbo(Valor As Long, cbo As Object) As Integer
    'devuelve el indice de un combo... para poder posicionar cuando
    'se hace una modificacion
    'ej: me.cbo.listindex=posindexcbo(valor_a_modificar,cbo)
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = Valor Then
            PosIndexCbo = i
            Exit Function
        End If
    Next i
    PosIndexCbo = -1
End Function
Function PosIndexLST(Valor As Long, lst As Object) As Integer
    'devuelve el indice de un combo... para poder posicionar cuando
    'se hace una modificacion
    'ej: me.cbo.listindex=posindexcbo(valor_a_modificar,cbo)
    For i = 0 To lst.ListCount - 1
        If lst.ItemData(i) = Valor Then
            PosIndexLST = i
            Exit Function
        End If
    Next i
End Function


Public Sub foco(ByRef texto As Object)
    'selecciona para modificar cualquier textbox
    'q se pase como argumento
    texto.SelStart = 0
    texto.SelLength = Len(texto)
End Sub


Function LstOrdenar(lst As ListView, columna As Integer)
    lst.Sorted = True
    'funcion que sirve para ordenar un listview en caso
    'de hacer click en alguna columna
    lst.SortKey = columna - 1
    If lst.SortOrder = lvwAscending Then
        lst.SortOrder = lvwDescending
    Else
        lst.SortOrder = lvwAscending
    End If
End Function


Function busca_en_lista(lst As ListView, id As Integer, Optional tem As Integer) As Boolean
    On Error Resume Next
    Dim esta As Boolean
    Dim x As ListItems
    esta = False

    For nn = 1 To lst.ListItems.count
        If tem = Empty Then
            id2 = lst.ListItems(nn)
        Else
            id2 = CInt(lst.ListItems(nn).ListSubItems(tem))
        End If

        If id = id2 Then
            esta = True

        End If

    Next nn
    If esta Then
        busca_en_lista = True
    Else
        busca_en_lista = False
    End If
End Function
Function quitar_de_lista(list As ListView)
    For i = list.ListItems.count To 1 Step -1
        If list.ListItems(i).Checked = True Then
            list.ListItems.remove (i)
        End If
    Next
End Function
Function datetimeFormateada(fec As Date) As String
    A = Format(fec, "yyyy-mm-dd hh:mm:ss")
    datetimeFormateada = A
End Function
Function dateFormateada(fec As Date) As String
    A = Format(fec, "yyyy-mm-dd")
    dateFormateada = A
End Function


Function amortiza(Cantidad) As Long    'para amortizar por cantidad
    If Cantidad < 10 Then
        amortiza = 10
    Else
        amortiza = Cantidad * 2
    End If
End Function


Function amortizaV2(id, Cantidad, forma As FormaCotizar, Optional amort = 0, Optional esCon As Boolean = True) As Long    'para amortizar por cantidad
    On Error GoTo err1
    Dim rs As Recordset
    Dim canti

    canti = Cantidad

    If forma = automatica_ Then
        If canti < 10 Then
            amortizaV2 = 10
        Else
            amortizaV2 = canti * 2
        End If




    ElseIf forma = Cantidad_ Then
        amortizaV2 = canti

    ElseIf forma = fabricados_ Then
        Set rs = conectar.RSFactory("select sum(cantidad_fabricados) as a from detalles_pedidos where idPieza=" & id)
        If Not IsNumeric(rs!A) Then A = 0 Else A = rs!A



        amortizaV2 = Cantidad + A
    ElseIf forma = fijo_ Then

        If Not esCon Then canti = 1
        amortizaV2 = amort    '* canti
    End If
    Exit Function
err1:
    amortizaV2 = 0
End Function


Function valorPorPeso(Kg, men10, men15, mas15) As Double
    If Kg >= 0 Then
        If Kg < 10 Then
            valorPorPeso = (CDbl(men10) / 100) + 1
        ElseIf Kg < 15 Then
            valorPorPeso = (CDbl(men15) / 100) + 1
        ElseIf Kg >= 15 Then
            valorPorPeso = (CDbl(mas15) / 100) + 1
        End If
    Else
        valorPorPeso = 1
    End If
End Function

Function normaliza(strf As String)
    'normaliza un string para que la primer letra
    'sea Mayúscula y el resto minúscula.
    'y reemplaza los caracteres por los
    'caracteres de escape correspondientes
    strt = Trim(strf)
    If strt <> Empty Then
        strt = Replace(strt, Chr(92), "\\")
        strt = Replace(strt, Chr(39), "\'")
        strt = Replace(strt, Chr(34), "\""")
        strt = Replace(strt, Chr(0), "\0")
        str2 = UCase(Left(strt, 1))
        strt = str2 & LCase(Mid(strt, 2))

        'todas las palabras en empiezan MAYUSCULA
        tt = Empty
        cont = 1
        sipos = InStr(cont, strt, " ", 1)
        If sipos = 0 Then auxsipos = 0 Else auxsipos = 1
        While sipos <> 0
            primera = Mid$(strt, cont, sipos - cont)
            strprimera2 = UCase(Left(primera, 1))
            primera = strprimera2 & LCase(Mid(primera, 2))

            cont = sipos + 1
            sipos = InStr(cont, strt, " ", 1)
            tt = tt & " " & primera
        Wend
        If auxsipos <> 0 Then    'si no encontro ningun espacio, no hace nada
            fina = Mid(strt, cont)
            strprimera2 = UCase(Left(fina, 1))
            fina = strprimera2 & LCase(Mid(fina, 2))    'armo el final de la cadena con lo que quedo
            strt = Trim(tt & " " & fina)
        End If



        tt = Empty
        'LOS PUNTOS
        cont = 1
        sipos = InStr(cont, Trim(strt), ".", 1)
        If sipos = 0 Then auxsipos = 0 Else auxsipos = 1
        While sipos <> 0
            primera = Trim(Mid$(strt, cont, sipos - cont))
            strprimera2 = UCase(Left(primera, 1))
            primera = strprimera2 & Mid(primera, 2)



            If cont = 1 Then
                tt = primera & "."
            Else
                tt = tt & primera & "."
            End If

            cont = sipos + 1
            sipos = InStr(cont, strt, ".", 1)
        Wend
        If auxsipos <> 0 Then    'si no encontro ningun espacio, no hace nada
            fina = Trim(Mid(strt, cont))
            strprimera2 = UCase(Left(fina, 1))
            fina = strprimera2 & LCase(Mid(fina, 2))    'armo el final de la cadena con lo que quedo
            strt = tt & " " & fina

        End If
        'strt = "'" & strt & "'"
        normaliza = Trim(strt)
    End If
End Function

Function busca_en_lista2(lst As ListView, id As Integer, Optional tem As Integer) As Long
    Dim esta As Boolean
    Dim x As ListItems
    esta = False

    For nn = 1 To lst.ListItems.count
        If tem = Empty Then
            id2 = lst.ListItems(nn)
        Else
            id2 = CInt(lst.ListItems(nn).ListSubItems(tem))
        End If

        If id = id2 Then
            esta = True
            pos = nn
            Exit For
        End If

    Next nn
    If esta Then
        busca_en_lista2 = nn
    Else
        busca_en_lista2 = -1
    End If
End Function



Public Function cantxhoja(ByVal x1, ByVal y1, ByVal x2, ByVal y2) As Long
    On Error Resume Next
    If y1 = 0 Then y1 = 1
    If x1 = 0 Then x1 = 1
    If y2 = 0 Then y2 = 1
    If x2 = 0 Then x2 = 1
    cantxhoja = Fix((y2 / y1)) * Fix((x2 / x1))
End Function

Public Sub mostrarInforme()
    Set frm = frmPrincipal
    '    frm.stbar1.Panels(3).text = "Version: " & versionado
    '    frm.stbar1.Panels(4).text = funciones.FEcha(Now)
End Sub

Public Function calcularTamañoArchivo(ByVal tamOrig As Double, ByRef tamDest As Double, uni As String)
    'si pesa meno de 1kb lo muestro en bytes
    If tamOrig < 1024 Then
        tamDest = tamOrig
        uni = "By"
    Else
        'si pesa menos de 1 mb lo muestro en kb
        If tamOrig < 1048576 Then
            tamDest = Math.Round(tamOrig / 1024, 2)
            uni = "Kb"
        Else
            tamDest = tamOrig / 1024
            tamDest = Math.Round(tamDest / 1024, 2)
            uni = "Mb"
        End If
    End If

End Function

Public Function crearUsuario(nombre, Apellido) As String
    Dim A As New classSignoplast
    Dim r As Recordset
    Dim usu As String
    usu = LCase(nombre & Mid(Apellido, 1, 1) & Right(Apellido, 1))
    Set r = conectar.RSFactory("Select count(id) as c from usuarios where usuario='" & usu & "'")

    If r!c > 0 Then
        i = 1
        esta = True
        While esta
            usu = usu & i
            Set r = conectar.RSFactory("Select count(id) as c from usuarios where usuario='" & usu & "'")
            If r!c > 0 Then
                i = i + 1
            Else
                esta = False
                '    crearUsuario = usu
            End If
        Wend
    Else
        'crearUsuario = usu
    End If

    'quito los espacios en blanco.
    crearUsuario = Replace(usu, " ", "")
End Function
Public Sub ValidarPermisoss(idUsuario)
    Dim r As Recordset
    Set r = conectar.RSFactory("select * from usuariosPermisos where idUsuario=" & idUsuario)
    c = 0
    While Not r.EOF
        c = c + 1
        r.MoveNext
    Wend
    If c = 1 Then
        r.MoveFirst
        cp = r!panel
        ventas = r!ventas
        compras = r!compras
        prov = r!proveedores
        clientes = r!clientes
        seg = r!seguimientos
        des = r!desarrollo
        plan = r!planeamiento
        'grupo PANEL DE CONTROL
        'If CP = 0 Then frmPrincipal.SmartMenuXP1.MenuItems.Enabled(1) = False
        'grupo ventas
        If ventas = 0 Then
            '            With frmPrincipal.SmartMenuXP1.MenuItems
            '                .Enabled(31) = False
            '                .Enabled(34) = False
            '                .Enabled(39) = False
            '            End With
        End If
        'grupo clientes
        'If clientes = 0 Then frmPrincipal.SmartMenuXP1.MenuItems.Enabled(40) = False
        'grupo compras
        If compras = 0 Then
            '            With frmPrincipal.SmartMenuXP1.MenuItems
            '                .Enabled(42) = False
            '                .Enabled(46) = False
            '                .Enabled(51) = False
            '                .Enabled(52) = False
            '                .Enabled(53) = False
            '            End With
        End If
        'grupo proveedores
        '        If prov = 0 Then frmPrincipal.SmartMenuXP1.MenuItems.Enabled(54) = False
        'grupo desarrollo
        '        If des = 0 Then frmPrincipal.SmartMenuXP1.MenuItems.Enabled(61) = False
        'grupo seguimiento
        '        If seg = 0 Then frmPrincipal.SmartMenuXP1.MenuItems.Enabled(60) = False
        'grupo planeamiento
        If plan = 0 Then
            '            With frmPrincipal.SmartMenuXP1.MenuItems
            '                .Enabled(58) = False
            '                .Enabled(57) = False
            '                .Enabled(56) = False
            '
            '            End With

        End If
    Else
        MsgBox "Error interno de permisos. Se cierra el sistema", vbCritical, "Error"
        End
    End If

End Sub
Public Sub Imprimir_ListView(ListView As ListView, enca As String)

    Dim i As Integer, AnchoCol As Single, Espacio As Integer, x As Integer

    AnchoCol = 0
    'Recorremos desde la primer columna hasta la última para almacenar el ancho total
    For i = 1 To ListView.ColumnHeaders.count
        AnchoCol = AnchoCol + ListView.ColumnHeaders(i).Width
    Next

    Espacio = 0

    'Encabezado de ejemplo
    Printer.Print UCase(enca)
    Printer.Print
    'Imprime una línea
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)

    With ListView

        'Acá se imprimen los encabezados del ListView
        For i = 1 To .ColumnHeaders.count
            Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
            If ListView.ColumnHeaders(i).Width > 1 Then
                Printer.Print ListView.ColumnHeaders(i).text;
            End If
            Printer.CurrentX = Espacio
        Next

        Printer.Print

        'Imprime una línea
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)

        'Imprime Línea en blanco
        Printer.Print

        'Este bucle recorre los items y subitems del ListView  y los imprime
        For i = 1 To .ListItems.count
            Espacio = 0

            Set lItem = .ListItems(i)
            Printer.Print lItem.text;
            'Recorremos las columnas
            For x = 1 To .ColumnHeaders.count - 1
                Espacio = Espacio + CInt(.ColumnHeaders(x).Width * Printer.ScaleWidth / AnchoCol)
                Printer.CurrentX = Espacio
                If ListView.ColumnHeaders(x + 1).Width > 1 Then
                    Printer.Print lItem.SubItems(x);
                End If
            Next

            'Otro espacio en blanco
            Printer.Print
        Next

    End With

    Printer.Print
    'Imprime la línea de final de impresión
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    Printer.Print
    'Texto del pie
    Printer.Print " Fin de la impresión "


    'Comenzamos la impresión
    Printer.EndDoc
End Sub
Public Sub vaciarControles(frm As Form)
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            ctl.text = ""
        End If
    Next
End Sub

Public Function ConvertirAFechaAfip(entrada As String) As String
Dim FEcha As String
Dim f_anio As String, f_mes As String, f_dia As String
f_anio = Right(entrada, 4)
f_mes = Mid(entrada, 3, 2)
 f_dia = Mid(entrada, 1, 2)
ConvertirAFechaAfip = f_dia & "/" & f_mes & "/" & f_anio
End Function


Public Function FormatearDecimales(numero As Double, Optional ByVal cantDecimales As Long = 2) As String
    FormatearDecimales = Format(numero, "0." & String$(cantDecimales, "0"))
End Function

Public Function RedondearDecimales(numero As Double, Optional Decimales As Long = 2) As Double
    RedondearDecimales = Format(Math.Round(numero, Decimales), "0.00")
End Function
Public Function cuantasHoras(Inicio As Date, Fin As Date)
    segundos = DateDiff("s", Inicio, Fin)
    minutos = (segundos / 60)
    horas = Math.Round(minutos / 60, 2)
    cuantasHoras = horas
End Function
Public Function ingreso(Optional msg = Empty) As Variant
    frmSistemaIngresar.nombre = msg
    frmSistemaIngresar.Show 1
    ingreso = frmSistemaIngresar.nombre
    Unload frmSistemaIngresar

End Function



Public Function Redondear(dblntor As Double, Optional cntdecas As Integer) As Double

    Dim dblpot As Double
    Dim dblf As Double

    If dblntor < 0 Then dblf = -0.5 Else: dblf = 0.5
    dblpot = 10 ^ intcntdec
    Redondear = Fix(dblntor * dblpot * (1 + 1E-16) + dblf) / dblpot
End Function


Public Function ImprimirLista(titulo, lst As ListView, CD As CommonDialog, Optional linea2 = Empty, Optional F_1 = Empty, Optional F_2 = Empty) As Boolean
    On Error GoTo err91
    CD.ShowPrinter

    AnchoCol = 0

    For i = 1 To lst.ColumnHeaders.count
        AnchoCol = AnchoCol + lst.ColumnHeaders(i).Width
    Next
    Espacio = 0

    Printer.Font.Size = 12
    Printer.Font.Bold = True
    'Printer.Line  (8800, 1400)-(10100, 1400)
    Printer.Print UCase(titulo)
    If linea2 <> Empty Then
        Printer.Print UCase(linea2)
    End If

    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    With lst

        'Acá se imprimen los encabezados del ListView
        For i = 1 To .ColumnHeaders.count
            Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
            If lst.ColumnHeaders(i).Width > 1 Then
                'Printer.Print i
                Printer.Print lst.ColumnHeaders(i).text;
            End If
            Printer.CurrentX = Espacio
        Next
        Printer.Font.Bold = False
        Printer.Print
        'Imprime una línea
        Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
        Printer.Print

        'Este bucle recorre los items y subitems del ListView  y los imprime

        For i = 1 To .ListItems.count
            Espacio = 0

            Set lItem = .ListItems(i)
            Printer.Print lItem.text;
            'Recorremos las columnas
            For x = 1 To .ColumnHeaders.count - 1
                Espacio = Espacio + CInt(.ColumnHeaders(x).Width * Printer.ScaleWidth / AnchoCol)
                Printer.CurrentX = Espacio
                If lst.ColumnHeaders(x + 1).Width > 1 Then

                    Printer.Print lItem.SubItems(x);



                End If
            Next

            'Otro espacio en blanco
            Printer.Print


        Next

    End With

    Printer.Print
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)


    Printer.Print
    If F_1 <> Empty Then
        Printer.Print F_1
    End If

    If F_2 <> Empty Then
        Printer.Print F_2
    End If

    Printer.Print "Fecha emision: " & Format(Now, "dd-mm-yy")

    Printer.EndDoc
    ImprimirLista = True

    Exit Function
err91:
    ImprimirLista = False
    If tra Then cn.RollbackTrans

End Function
Public Sub ordenar_grilla(ByVal Column As GridEX20.JSColumn, GridEX1 As GridEX)
    Dim grTemp As JSGroup
    Dim SortOrder As jgexSortOrderConstants
    SortOrder = Column.SortOrder
    'Clear SortKeys
    GridEX1.SortKeys.Clear
    'Add this new sortkey
    If SortOrder = jgexSortAscending Then
        'if the column was sorted in ascending order, sort the column in descending order
        GridEX1.SortKeys.Add Column.Index, jgexSortDescending
    Else
        'if was sorted in descending order or not sorted, sort the column in ascending order
        GridEX1.SortKeys.Add Column.Index, jgexSortAscending
    End If
End Sub
Public Sub FillComboBox(ByRef combo As ComboBox, ByRef col As Collection, ByRef propertyForShow As String, ByRef propertyForId As String, ByRef selectFirst As Boolean)
    combo.Clear
    Dim item As Object
    For Each item In col
        combo.AddItem CallByName(item, propertyForShow, VbGet)
        combo.ItemData(combo.NewIndex) = CallByName(item, propertyForId, VbGet)
    Next item
    If (selectFirst And combo.ListCount > 0) Then combo.ListIndex = 0
End Sub
Public Function BuscarEnColeccion(col As Collection, indice As String) As Boolean
    On Error GoTo er1
    BuscarEnColeccion = True
    col (indice)
    Exit Function
er1:
    BuscarEnColeccion = False
End Function

Public Function VerificarCUIT(Cuit) As Boolean
    'Verifica si el tamaño es el correcto.
    VerificarCUIT = True

    If Len(Cuit) = 11 Then
        'Individualiza y multiplica los dígitos.
        xa = Val(Mid$(Cuit, 1, 1)) * 5
        XB = Val(Mid$(Cuit, 2, 1)) * 4
        XC = Val(Mid$(Cuit, 3, 1)) * 3
        XD = Val(Mid$(Cuit, 4, 1)) * 2
        XE = Val(Mid$(Cuit, 5, 1)) * 7
        XF = Val(Mid$(Cuit, 6, 1)) * 6
        XG = Val(Mid$(Cuit, 7, 1)) * 5
        XH = Val(Mid$(Cuit, 8, 1)) * 4
        XI = Val(Mid$(Cuit, 9, 1)) * 3
        XJ = Val(Mid$(Cuit, 10, 1)) * 2
        'xj2 = Val(Mid$(Cuit, 11, 1)) * 1


        'Suma los resultantes.
        x = xa + XB + XC + XD + XE + XF + XG + XH + XI + XJ + xj2

        'Calcula el dígito de control.
        Control = (11 - (x Mod 11)) Mod 11

        'Verifica si el dígito de control ingresado difiere con el calculado.
        If Control <> Val(Mid$(Cuit, 11, 1)) Then

            'Presenta la ventana de aviso.
            '        MsgBox "El CUIT ingresado es incorrecto. Verifíquelo e intente nuevamente." + Chr$(13) + Chr$(13) + "CUIT Ingresado: " + CUIT + Chr$(13) + "CUIT Estimativo: " + Left(CUIT, 12) + Trim$(str$(Control)), 48, "CUIT ERRONEO"
            VerificarCUIT = False

        End If

    Else
        VerificarCUIT = False
    End If

End Function

Public Function str2Array(txt1 As String) As String()


    ReDim chararray(Len(txt1) - 1) As String

    For i = 1 To Len(txt1)
        chararray(i - 1) = Mid(txt1, i, 1)
    Next
    str2Array = chararray
End Function

Public Function calcularDVFactura(numero As String) As String
    On Error GoTo err1

    ReDim pos(Len(numero)) As String

    pos = str2Array(numero)



    Dim xa As Long
    xa = 0
    Dim sumapar As Long
    Dim sumaimpaar As Long
    sumapar = 0
    sumaimpar = 0
    Dim n As Long
    Dim i As Integer
    For i = 0 To UBound(pos)
        n = CLng(pos(i))

        If (i + 1) Mod 2 = 0 Then
            sumapar = sumapar + n
        Else
            sumaimpar = sumaimpar + n
        End If

    Next i
    Dim dv As Integer
    xa = (sumaimpar * 3) + sumapar
    Dim x As Integer
    For x = 0 To 9
        If (xa + x) Mod 10 = 0 Then
            dv = x

        End If
    Next x

    calcularDVFactura = dv
    Exit Function
err1:
    calcularDVFactura = vbNullString
End Function


Public Function CreateGUID( _
       Optional strRemoveChars As String = "{}-") As String
    Dim udtGUID As GUID
    Dim strGUID As String
    Dim bytGUID() As Byte
    Dim lngLen As Long
    Dim lngRetVal As Long
    Dim lngPos As Long

    'Initialize
    lngLen = 40
    bytGUID = String(lngLen, 0)
    'Create the GUID
    CoCreateGuid udtGUID
    'Convert the structure into a displayable string
    lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
    strGUID = bytGUID
    If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
        lngRetVal = lngRetVal - 1
    End If

    'Trim the trailing characters
    strGUID = Left$(strGUID, lngRetVal)

    'Remove the unwanted characters
    For lngPos = 1 To Len(strRemoveChars)
        strGUID = Replace(strGUID, Mid(strRemoveChars, lngPos, 1), "")
    Next
    CreateGUID = strGUID
End Function
Public Function ListBoxHasCheckedItems(ByRef lst As ListBox) As Boolean
    Dim i As Long
    ListBoxHasCheckedItems = False
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            ListBoxHasCheckedItems = True
            Exit Function
        End If
    Next i
End Function
Public Function ValidarTextBox(text As TextBox, ByRef Cancel)
    Cancel = Not IsNumeric(text)
    text.BackColor = IIf(Cancel, vbRed, &H80000005)
End Function




' List all files below the directory that
' match the pattern.
Public Sub ListFiles(ByRef result As Collection, ByVal start_dir As String, ByVal pattern As String)
    Const MAXDWORD = 2 ^ 32

    Dim dir_names() As String
    Dim num_dirs As Integer
    Dim i As Integer
    Dim fname As String
    Dim attr As Integer
    Dim search_handle As Long
    Dim file_data As WIN32_FIND_DATA
    Dim file_size As Double

    ' Get the matching files in this directory.

    ' Get the first file.
    search_handle = FindFirstFile( _
                    start_dir & pattern, file_data)
    If search_handle <> INVALID_HANDLE_VALUE Then
        ' Get the rest of the files.
        Do
            fname = file_data.cFileName
            fname = Left$(fname, InStr(fname, Chr$(0)) - 1)
            file_size = (file_data.nFileSizeHigh * _
                         MAXDWORD) + file_data.nFileSizeLow
            If file_size > 0 Then
                'lst.AddItem start_dir & fname & " (" & _
                 Format$(file_size) & ")"

                'result.Add start_dir & fname '& " (" & _
                 Format$(file_size) & ")"

                Set tmpFileMetadata = New FileMetadataDTO
                tmpFileMetadata.filename = fname
                tmpFileMetadata.DirectoryName = Left$(start_dir, Len(start_dir) - 1)
                tmpFileMetadata.FileSize = file_size
                result.Add tmpFileMetadata

                total_size = total_size + file_size
            Else
                'lst.AddItem start_dir & fname
            End If

            ' Get the next file.
            If FindNextFile(search_handle, file_data) = 0 _
               Then Exit Do
        Loop

        ' Close the file search hanlde.
        FindClose search_handle
    End If

    ' Get the list of subdirectories.
    search_handle = FindFirstFile( _
                    start_dir & "*.*", file_data)
    If search_handle <> INVALID_HANDLE_VALUE Then
        ' Get the rest of the files.
        Do
            If file_data.dwFileAttributes And DDL_DIRECTORY _
               Then
                fname = file_data.cFileName
                fname = Left$(fname, InStr(fname, Chr$(0)) _
                                     - 1)
                If fname <> "." And fname <> ".." Then
                    num_dirs = num_dirs + 1
                    ReDim Preserve dir_names(1 To num_dirs)
                    dir_names(num_dirs) = fname
                End If
            End If

            ' Get the next file.
            If FindNextFile(search_handle, file_data) = 0 _
               Then Exit Do
        Loop

        ' Close the file search handle.
        FindClose search_handle
    End If

    ' Search the subdirectories.
    For i = 1 To num_dirs
        ListFiles result, start_dir & dir_names(i) & "\", pattern
    Next i
End Sub

' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
Public Function BrowseForDirectory(ByVal caption As String) As String
    Dim browse_info As BrowseInfo
    Dim item As Long
    Dim dir_name As String

    browse_info.hwndOwner = hWnd
    browse_info.pIDLRoot = 0
    browse_info.sDisplayName = Space$(260)
    browse_info.sTitle = caption
    browse_info.ulFlags = 1    ' Return directory name.
    browse_info.lpfn = 0
    browse_info.lParam = 0
    browse_info.iImage = 0

    item = SHBrowseForFolder(browse_info)
    If item Then
        dir_name = Space$(260)
        If SHGetPathFromIDList(item, dir_name) Then
            BrowseForDirectory = Left(dir_name, InStr(dir_name, Chr$(0)) - 1)
        Else
            BrowseForDirectory = vbNullString
        End If
    End If
End Function


Public Function IsSomething(Obj As Object) As Boolean
    On Error Resume Next
    IsSomething = Not Obj Is Nothing
End Function

Public Function JoinCollectionValues(col As Collection, delimiter As String, Optional objectProperty As String = vbNullString) As String
    Dim value As Variant
    Dim ret As String
    Dim cont As Long
    For Each value In col
        cont = cont + 1

        If LenB(objectProperty) = 0 Then
            ret = ret & value
        Else
            ret = ret & CallByName(value, objectProperty, VbGet)
        End If

        If cont <> col.count Then
            ret = ret & delimiter
        End If
    Next value
    JoinCollectionValues = ret
End Function

Public Function JoinDictionaryKeyValues(dic As Dictionary, delimiter As String) As String
    Dim value As Variant
    Dim ret As String
    Dim cont As Long
    For Each value In dic.Keys
        cont = cont + 1
        ret = ret & value
        If cont <> dic.count Then
            ret = ret & delimiter
        End If
    Next value
    JoinDictionaryKeyValues = ret
End Function


Public Sub FillComboBoxDateRanges(ByRef combo As Xtremesuitecontrols.ComboBox)
    combo.Clear
    combo.AddItem "Hoy"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_Today
    combo.AddItem "Ayer"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_Yesterday
    combo.AddItem "Semana Actual"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_WeekCurrent
    combo.AddItem "Semana Pasada"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_WeekLast
    combo.AddItem "Mes Próximo"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_MonthNext

    combo.AddItem "Mes Actual"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_MonthCurrent
    combo.AddItem "Mes Pasado"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_MonthLast

    combo.AddItem "Año Actual"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_YearCurrent
    combo.AddItem "Año Pasado"
    combo.ItemData(combo.NewIndex) = DateRangeValue.DRV_YearLast
    combo.ListIndex = -1
End Sub

Public Sub CalculateDateRange(ByRef combo As Xtremesuitecontrols.ComboBox, ByRef dtpDesde As Xtremesuitecontrols.DateTimePicker, ByRef dtpHasta As Xtremesuitecontrols.DateTimePicker)

    Dim desde As Date
    Dim hasta As Date

    If combo.ListIndex <> -1 Then
        Dim rangeDateId As Long

        rangeDateId = combo.ItemData(combo.ListIndex)

        Select Case rangeDateId
            Case DateRangeValue.DRV_Today
                desde = Date
                hasta = Date
            Case DateRangeValue.DRV_Yesterday
                desde = DateAdd("d", -1, Date)
                hasta = DateAdd("d", -1, Date)
            Case DateRangeValue.DRV_WeekCurrent
                desde = Date
                While (Weekday(desde, vbMonday) <> 1)
                    desde = DateAdd("d", -1, desde)
                Wend
                hasta = DateAdd("d", 6, desde)
            Case DateRangeValue.DRV_WeekLast
                'calculo semana actual y le resto 7 dias
                desde = Date
                While (Weekday(desde, vbMonday) <> 1)
                    desde = DateAdd("d", -1, desde)
                Wend
                hasta = DateAdd("d", 6, desde)

                desde = DateAdd("d", -7, desde)
                hasta = DateAdd("d", -7, hasta)

            Case DateRangeValue.DRV_MonthCurrent
                desde = DateSerial(Year(Date), Month(Date), 1)
                hasta = DateAdd("d", -1, DateAdd("m", 1, desde))

            Case DateRangeValue.DRV_MonthLast
                'calculo mest actual y resto 1 mes
                desde = DateSerial(Year(Date), Month(Date), 1)
                hasta = DateAdd("d", -1, DateAdd("m", 1, desde))

                hasta = DateAdd("d", -1, desde)
                desde = DateAdd("m", -1, desde)

            Case DateRangeValue.DRV_YearCurrent
                desde = DateSerial(Year(Date), 1, 1)
                hasta = DateSerial(Year(Date), 12, 31)
            Case DateRangeValue.DRV_YearLast
                desde = DateSerial(Year(Date) - 1, 1, 1)
                hasta = DateSerial(Year(Date) - 1, 12, 31)
            Case DateRangeValue.DRV_MonthNext
                desde = DateSerial(Year(Date), Month(Date) + 1, 1)
                hasta = DateAdd("d", -1, DateAdd("m", 1, desde))

                'h 'asta = DateAdd("d", -1, desde)
                'desde = DateAdd("m", 1, desde)


        End Select

    End If

    dtpDesde.value = desde
    dtpHasta.value = hasta

End Sub

Public Function Hours2HourMinute(hours As Double) As String
    On Error GoTo E

    Dim horas As Double
    Dim minutos As Double

    Dim tmp As Double: tmp = hours
    horas = Fix(tmp)
    minutos = tmp - Fix(tmp)
    minutos = Fix(minutos * 60)

    Hours2HourMinute = horas & ":" & Format(minutos, "00")
    Exit Function
E:
    Hours2HourMinute = "0:00"
End Function


Public Function GetTmpPath()

    Dim sFolder As String    ' Name of the folder
    Dim lRet As Long    ' Return Value

    sFolder = String(MAX_PATH, 0)
    lRet = GetTempPath(MAX_PATH, sFolder)

    If lRet <> 0 Then
        GetTmpPath = Left(sFolder, InStr(sFolder, Chr(0)) - 1)
    Else
        GetTmpPath = vbNullString
    End If

End Function

Function GetFilePath(filename As String) As String
    Dim i As Long
    For i = Len(filename) To 1 Step -1
        Select Case Mid$(filename, i, 1)
            Case ":"
                ' colons are always included in the result
                GetFilePath = Left$(filename, i)
                Exit For
            Case "\"
                ' backslash aren't included in the result
                GetFilePath = Left$(filename, i - 1)
                Exit For
        End Select
    Next
End Function

Public Function AjustarLineas(st As String) As String

    Dim A As String
    Dim b As String

    If Len(st) >= 55 Then
        Dim trozos As New Collection
        Dim i As Long
        Dim buffer As String
        Dim char As String

        For i = 1 To Len(st)
            char = Mid(st, i, 1)
            buffer = buffer & char
            If Len(buffer) = 55 Then
                trozos.Add Trim$(StrConv(buffer, vbUpperCase))
                buffer = vbNullString
            End If
        Next i
        If Len(buffer) > 0 Then trozos.Add StrConv(buffer, vbUpperCase)

        AjustarLineas = funciones.JoinCollectionValues(trozos, vbNewLine)
    Else
        AjustarLineas = st
    End If

End Function


Public Function RazonSocialFormateada(rz As String) As String
    Dim siglas As Variant
    Dim i As Long
    siglas = Array("S.A.", "S.A.C.I.F.I.A.", "S.A.I.C.", "S.R.L.", "S.A.M.", "S.H.", "S.A.C.I.", "S.A.T.I.", "S.A.C.I.F.", "S.C.S.")
    rz = StrConv(rz, vbProperCase)
    For i = 0 To UBound(siglas)
        rz = Replace$(rz, siglas(i), siglas(i), , , vbTextCompare)
    Next i
    RazonSocialFormateada = rz
End Function


Function InstrCount(StringToSearch As String, _
                    StringToFind As String) As Long

    If Len(StringToFind) Then
        InstrCount = UBound(Split(StringToSearch, StringToFind))
    End If
End Function


