VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form frmSistemaTests 
   Caption         =   "Tests"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17850
   Icon            =   "frmSistemasTests.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   17850
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _Version        =   786432
      _ExtentX        =   13573
      _ExtentY        =   8705
      _StockProps     =   79
      Caption         =   "Consulta a ARCA"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtCuit 
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   3375
      End
      Begin XtremeSuiteControls.PushButton cmdConsultarARCA 
         Height          =   495
         Left            =   4200
         TabIndex        =   1
         Top             =   315
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Traer Datos"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   7095
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   7095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmSistemaTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const URL_API As String = "http://192.168.0.2:8011"


Private Sub Form_Load()

    txtCuit.Text = ""
    txtCuit.MaxLength = 13

    cmdConsultarARCA.caption = "Consultar ARCA"

    Label1.caption = "Razón social:"
    Label2.caption = "Domicilio:"
    Label3.caption = "Estado: esperando ingreso"

    Label1.ForeColor = vbWindowText
    Label2.ForeColor = vbWindowText
    Label3.ForeColor = vbWindowText

End Sub


Private Sub txtCUIT_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii

        Case 8
            'Backspace

        Case 45
            'Guion

        Case 48 To 57
            'Números del 0 al 9

        Case Else
            KeyAscii = 0

    End Select

End Sub


Private Sub cmdConsultarARCA_Click()

    On Error GoTo ManejarError

    Dim consulta As New clsConsultaARCA

    Label1.caption = "Razón social:"
    Label2.caption = "Domicilio:"
    Label3.caption = "Estado: consultando ARCA..."

    Label1.ForeColor = vbWindowText
    Label2.ForeColor = vbWindowText
    Label3.ForeColor = vbWindowText

    cmdConsultarARCA.Enabled = False
    DoEvents

    If Not consulta.Consultar(txtCuit.Text) Then

        Label3.caption = _
            "Estado: " & consulta.UltimoError

        Label3.ForeColor = vbRed

        cmdConsultarARCA.Enabled = True
        Exit Sub

    End If

    Label1.caption = _
        "Razón social: " & consulta.RazonSocial

    Label2.caption = _
        "Domicilio: " & consulta.Domicilio

    Label3.caption = _
        "Estado: " & consulta.estado

    Label3.ForeColor = vbGreen

    cmdConsultarARCA.Enabled = True

    Set consulta = Nothing
    Exit Sub

ManejarError:

    cmdConsultarARCA.Enabled = True

    Label3.caption = _
        "Estado: error " & _
        CStr(Err.Number) & _
        " - " & Err.Description

    Label3.ForeColor = vbRed

    Set consulta = Nothing

End Sub


Private Sub ConsultarARCA(ByVal cuit As String)

    On Error GoTo ManejarError

    Dim Http As Object
    Dim url As String
    Dim respuesta As String

    Dim RazonSocial As String
    Dim Domicilio As String
    Dim estado As String

    url = URL_API & "/arca/constancia/" & cuit

    Set Http = CrearClienteHTTP()

    Http.setTimeouts 5000, 5000, 10000, 10000

    Http.Open "GET", url, False

    Http.setRequestHeader "Accept", "application/json"

    Http.send

    If Http.Status <> 200 Then

        Label1.caption = "Razón social:"
        Label2.caption = "Domicilio:"

        Label3.caption = "Error HTTP " & CStr(Http.Status) & _
                         ": " & ObtenerMensajeError(Http.responseText)

        Label3.ForeColor = vbRed

        Set Http = Nothing
        Exit Sub

    End If

    respuesta = CStr(Http.responseText)

    RazonSocial = ExtraerValorJSON(respuesta, "razon_social")
    Domicilio = ExtraerValorJSON(respuesta, "domicilio")
    estado = ExtraerValorJSON(respuesta, "estado")

    If Len(RazonSocial) = 0 Then
        RazonSocial = "(sin información)"
    End If

    If Len(Domicilio) = 0 Then
        Domicilio = "(sin información)"
    End If

    If Len(estado) = 0 Then
        estado = "(sin información)"
    End If

    Label1.caption = "Razón social: " & RazonSocial
    Label2.caption = "Domicilio: " & Domicilio
    Label3.caption = "Estado: " & estado

    Label1.ForeColor = vbWindowText
    Label2.ForeColor = vbWindowText
    Label3.ForeColor = vbGreen

    Set Http = Nothing

    Exit Sub

ManejarError:

    Label1.caption = "Razón social:"
    Label2.caption = "Domicilio:"

    Label3.caption = "No se pudo conectar con la API. " & _
                     "Error " & Err.Number & ": " & Err.Description

    Label3.ForeColor = vbRed

    Set Http = Nothing

End Sub


Private Function CrearClienteHTTP() As Object

    Dim Http As Object

    On Error Resume Next

    Set Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    If Http Is Nothing Then
        Set Http = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    End If

    On Error GoTo 0

    If Http Is Nothing Then

        Err.Raise vbObjectError + 1000, _
                  "CrearClienteHTTP", _
                  "No se encontró Microsoft XML ServerHTTP."

    End If

    Set CrearClienteHTTP = Http

End Function


Private Function SoloNumeros(ByVal valor As String) As String

    Dim i As Long
    Dim caracter As String
    Dim resultado As String

    resultado = ""

    For i = 1 To Len(valor)

        caracter = Mid$(valor, i, 1)

        If caracter >= "0" And caracter <= "9" Then
            resultado = resultado & caracter
        End If

    Next i

    SoloNumeros = resultado

End Function


Private Function FormatearCUIT(ByVal cuit As String) As String

    If Len(cuit) <> 11 Then

        FormatearCUIT = cuit
        Exit Function

    End If

    FormatearCUIT = Left$(cuit, 2) & "-" & _
                    Mid$(cuit, 3, 8) & "-" & _
                    Right$(cuit, 1)

End Function


Private Function CUITValido(ByVal cuit As String) As Boolean

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

        If InStr(1, "0123456789", caracter) = 0 Then
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

    CUITValido = (digitoCalculado = digitoIngresado)

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

    posicion = InStr(1, json, textoBuscar, vbTextCompare)

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
                        resultado = resultado & vbCrLf

                    Case "r"
                        'No agregamos nada.

                    Case "t"
                        resultado = resultado & vbTab

                    Case Else
                        resultado = resultado & caracter

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

        If caracter = "," Or caracter = "}" Then
            Exit For
        End If

        resultado = resultado & caracter

    Next i

    resultado = Trim$(resultado)

    If LCase$(resultado) <> "null" Then
        ExtraerValorJSON = resultado
    End If

End Function


Private Function ObtenerMensajeError(ByVal respuesta As String) As String

    Dim detalle As String

    detalle = ExtraerValorJSON(respuesta, "detail")

    If Len(detalle) > 0 Then
        ObtenerMensajeError = detalle
    Else
        ObtenerMensajeError = "Respuesta incorrecta del servidor"
    End If

End Function

