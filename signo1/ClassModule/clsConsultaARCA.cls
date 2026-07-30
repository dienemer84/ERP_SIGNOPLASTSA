VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsultaARCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const URL_API_ARCA As String = _
    "http://192.168.0.2:8011"

Private mCUIT As String
Private mRazonSocial As String
Private mDomicilio As String
Private mEstado As String
Private mTipoPersona As String
Private mTipoClave As String
Private mDireccion As String
Private mLocalidad As String
Private mProvincia As String
Private mCodigoPostal As String
Private mCondicionIVA As String

Private mUltimoError As String
Private mCodigoHTTP As Long


Public Property Get cuit() As String
    cuit = mCUIT
End Property

Public Property Get RazonSocial() As String
    RazonSocial = mRazonSocial
End Property

Public Property Get Domicilio() As String
    Domicilio = mDomicilio
End Property

Public Property Get estado() As String
    estado = mEstado
End Property

Public Property Get TipoPersona() As String
    TipoPersona = mTipoPersona
End Property

Public Property Get TipoClave() As String
    TipoClave = mTipoClave
End Property

Public Property Get direccion() As String
    direccion = mDireccion
End Property

Public Property Get localidad() As String
    localidad = mLocalidad
End Property

Public Property Get provincia() As String
    provincia = mProvincia
End Property

Public Property Get CodigoPostal() As String
    CodigoPostal = mCodigoPostal
End Property

Public Property Get UltimoError() As String
    UltimoError = mUltimoError
End Property

Public Property Get CodigoHTTP() As Long
    CodigoHTTP = mCodigoHTTP
End Property

Public Property Get CondicionIVA() As String
    CondicionIVA = mCondicionIVA
End Property

Public Function Consultar(ByVal valorCUIT As String) As Boolean

    On Error GoTo ManejarError

    Dim Http As Object
    Dim url As String
    Dim respuesta As String
    Dim cuitLimpio As String
    Dim mensajeServidor As String

    Consultar = False

    LimpiarResultado

    cuitLimpio = SoloNumeros(valorCUIT)

    If Len(cuitLimpio) = 0 Then

        mUltimoError = "Debe ingresar un CUIT."
        Exit Function

    End If

    If Len(cuitLimpio) <> 11 Then

        mUltimoError = "El CUIT debe contener 11 números."
        Exit Function

    End If

    If Not CUITValido(cuitLimpio) Then

        mUltimoError = "El CUIT ingresado no es válido."
        Exit Function

    End If

    url = URL_API_ARCA & _
          "/arca/constancia/" & _
          cuitLimpio

    Set Http = CrearClienteHTTP()

    Http.setTimeouts 5000, 5000, 30000, 30000

    Http.Open "GET", url, False
    Http.setRequestHeader "Accept", "application/json"
    Http.send

    mCodigoHTTP = CLng(Http.Status)
    respuesta = CStr(Http.responseText)

    If mCodigoHTTP <> 200 Then

        mensajeServidor = ExtraerValorJSON( _
                                respuesta, _
                                "detail" _
                            )

        If Len(mensajeServidor) > 0 Then

            mUltimoError = mensajeServidor

        Else

            mUltimoError = _
                "La API respondió HTTP " & _
                CStr(mCodigoHTTP) & "."

        End If

        Set Http = Nothing
        Exit Function

    End If

    mCUIT = ExtraerValorJSON( _
                    respuesta, _
                    "cuit" _
                )

    mRazonSocial = ExtraerValorJSON( _
                            respuesta, _
                            "razon_social" _
                        )

    mDomicilio = ExtraerValorJSON( _
                            respuesta, _
                            "domicilio" _
                        )

    mEstado = ExtraerValorJSON( _
                    respuesta, _
                    "estado" _
                )

    mTipoPersona = ExtraerValorJSON( _
                            respuesta, _
                            "tipo_persona" _
                        )

    mTipoClave = ExtraerValorJSON( _
                            respuesta, _
                            "tipo_clave" _
                        )

    mDireccion = ExtraerValorJSON( _
                            respuesta, _
                            "direccion" _
                        )

    mLocalidad = ExtraerValorJSON( _
                            respuesta, _
                            "localidad" _
                        )

    mProvincia = ExtraerValorJSON( _
                            respuesta, _
                            "provincia" _
                        )

    mCodigoPostal = ExtraerValorJSON( _
                            respuesta, _
                            "codigo_postal" _
                        )

    mCondicionIVA = ExtraerValorJSON( _
                            respuesta, _
                            "condicion_iva" _
                        )

    If Len(Trim$(mCondicionIVA)) = 0 Then
        mCondicionIVA = "Sin Datos"
    End If

    If Len(mCUIT) = 0 Then
        mCUIT = cuitLimpio
    End If

    If Len(mRazonSocial) = 0 Then

        mUltimoError = _
            "ARCA no devolvió la razón social."

        Set Http = Nothing
        Exit Function

    End If

    Consultar = True

    Set Http = Nothing
    Exit Function

ManejarError:

    Consultar = False

    mUltimoError = _
        "Error " & CStr(Err.Number) & _
        ": " & Err.Description

    Set Http = Nothing

End Function


Private Sub LimpiarResultado()

    mCUIT = ""
    mRazonSocial = ""
    mDomicilio = ""
    mEstado = ""
    mTipoPersona = ""
    mTipoClave = ""
    mDireccion = ""
    mLocalidad = ""
    mProvincia = ""
    mCodigoPostal = ""
    mCondicionIVA = ""

    mUltimoError = ""
    mCodigoHTTP = 0

End Sub


Private Function CrearClienteHTTP() As Object

    Dim Http As Object

    On Error Resume Next

    Set Http = CreateObject( _
                    "MSXML2.ServerXMLHTTP.6.0" _
                )

    If Http Is Nothing Then
        Set Http = CreateObject( _
                        "MSXML2.ServerXMLHTTP.3.0" _
                    )
    End If

    On Error GoTo 0

    If Http Is Nothing Then

        Err.Raise _
            vbObjectError + 1000, _
            "clsConsultaARCA.CrearClienteHTTP", _
            "No se encontró Microsoft XML ServerHTTP."

    End If

    Set CrearClienteHTTP = Http

End Function


Private Function SoloNumeros( _
    ByVal valor As String _
) As String

    Dim i As Long
    Dim caracter As String
    Dim resultado As String

    resultado = ""

    For i = 1 To Len(valor)

        caracter = Mid$(valor, i, 1)

        If caracter >= "0" And _
           caracter <= "9" Then

            resultado = resultado & caracter

        End If

    Next i

    SoloNumeros = resultado

End Function


Private Function CUITValido( _
    ByVal cuit As String _
) As Boolean

    Dim multiplicadores As String
    Dim suma As Long
    Dim digitoCalculado As Integer
    Dim digitoIngresado As Integer
    Dim i As Integer
    Dim caracter As String

    CUITValido = False

    If Len(cuit) <> 11 Then
        Exit Function
    End If

    For i = 1 To 11

        caracter = Mid$(cuit, i, 1)

        If InStr( _
               1, _
               "0123456789", _
               caracter _
           ) = 0 Then

            Exit Function

        End If

    Next i

    multiplicadores = "5432765432"
    suma = 0

    For i = 1 To 10

        suma = suma + _
               CInt(Mid$(cuit, i, 1)) * _
               CInt(Mid$(multiplicadores, i, 1))

    Next i

    digitoCalculado = 11 - (suma Mod 11)

    Select Case digitoCalculado

        Case 11
            digitoCalculado = 0

        Case 10
            digitoCalculado = 9

    End Select

    digitoIngresado = CInt(Right$(cuit, 1))

    CUITValido = _
        (digitoCalculado = digitoIngresado)

End Function


Private Function ExtraerValorJSON( _
    ByVal json As String, _
    ByVal clave As String _
) As String

    Dim textoBuscar As String
    Dim posicion As Long
    Dim posicionValor As Long
    Dim i As Long

    Dim caracter As String
    Dim resultado As String
    Dim escapado As Boolean

    ExtraerValorJSON = ""

    textoBuscar = """" & clave & """:"

    posicion = InStr( _
                    1, _
                    json, _
                    textoBuscar, _
                    vbTextCompare _
                )

    If posicion = 0 Then
        Exit Function
    End If

    posicionValor = posicion + Len(textoBuscar)

    Do While posicionValor <= Len(json)

        caracter = Mid$(json, posicionValor, 1)

        If caracter <> " " And _
           caracter <> vbTab And _
           caracter <> vbCr And _
           caracter <> vbLf Then

            Exit Do

        End If

        posicionValor = posicionValor + 1

    Loop

    If posicionValor > Len(json) Then
        Exit Function
    End If

    caracter = Mid$(json, posicionValor, 1)

    If caracter = """" Then

        resultado = ""
        escapado = False

        For i = posicionValor + 1 To Len(json)

            caracter = Mid$(json, i, 1)

            If escapado Then

                Select Case caracter

                    Case """"
                        resultado = resultado & """"

                    Case "\"
                        resultado = resultado & "\"

                    Case "n"
                        resultado = _
                            resultado & vbCrLf

                    Case "r"
                        'No se agrega nada.

                    Case "t"
                        resultado = _
                            resultado & vbTab

                    Case Else
                        resultado = _
                            resultado & caracter

                End Select

                escapado = False

            ElseIf caracter = "\" Then

                escapado = True

            ElseIf caracter = """" Then

                Exit For

            Else

                resultado = resultado & caracter

            End If

        Next i

        ExtraerValorJSON = resultado
        Exit Function

    End If

    resultado = ""

    For i = posicionValor To Len(json)

        caracter = Mid$(json, i, 1)

        If caracter = "," Or _
           caracter = "}" Then

            Exit For

        End If

        resultado = resultado & caracter

    Next i

    resultado = Trim$(resultado)

    If LCase$(resultado) <> "null" Then
        ExtraerValorJSON = resultado
    End If

End Function

